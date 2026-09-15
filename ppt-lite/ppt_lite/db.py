"""db.py：schema、连接管理、BUSY 重试、复合操作执行器（§4 全部入口）。

铁律落地：
- 每连接 PRAGMA：WAL / foreign_keys / busy_timeout（§10）
- db_transaction(busy=...)：短事务；SQLITE_BUSY 由外层复合操作整体重试
- run_locked(...)：唯一锁序执行器 = 外层 fslock → 短事务 → commit → 锁内搬 trash → 解锁
- 所有状态迁移带来源条件 + rowcount 断言
"""
from __future__ import annotations

import logging
import os
import random
import sqlite3
import threading
import time
from contextlib import contextmanager
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Iterator

from .config import Config
from .fslock import fslock, to_trash

log = logging.getLogger("pptlite.db")

# 诊断开关：死锁/卡顿时打印线程栈（PPT_LITE_DIAG=1）
_DIAG = os.environ.get("PPT_LITE_DIAG") == "1"

SCHEMA = """
create table if not exists file (
  id integer primary key,
  sha256 text unique not null,
  name text not null,
  category text not null check (category in ('material','spec','sample')),
  status text not null default 'idle'
              check (status in ('idle','queued','working','done','failed')),
  error text,
  raw_path text not null,
  derived_path text,
  doc_type text,                        -- 业务类型：洞察材料/立项材料/BP材料/…（自由值，常用项下拉）
  note text,
  created_at text default (datetime('now'))
);

create table if not exists extraction (
  id integer primary key autoincrement,
  file_id integer not null references file(id) on delete cascade,
  parser_version text, renderer_version text,
  design_json json,
  status text not null default 'queued'
              check (status in ('queued','running','done','failed')),
  started_at text, finished_at text, error text
);
create unique index if not exists uq_ex_one_active on extraction (file_id)
  where status in ('queued','running');
create unique index if not exists uq_ex_one_done on extraction (file_id) where status='done';

create table if not exists slide (
  id integer primary key,
  extraction_id integer not null references extraction(id) on delete cascade,
  file_id integer not null references file(id) on delete cascade,
  no integer not null,
  title text,
  search_text text,
  texts json, struct json,
  png text not null, thumb text not null,
  unique (extraction_id, no)
);

create table if not exists media (
  id integer primary key,
  sha256 text unique not null,
  phash text, fmt text, w integer, h integer,
  path text, preview text
);

create table if not exists media_ref (
  slide_id integer not null references slide(id) on delete cascade,
  media_id integer not null references media(id),
  role text not null default 'content',
  primary key (slide_id, media_id, role)
);

create table if not exists slide_review (
  file_id integer not null references file(id) on delete cascade,
  no integer not null,
  status text not null default 'pending'
       check (status in ('pending','approved','reference','forbidden')),
  slide_type text, quality integer,
  tags text,
  is_template integer default 0,
  template_json json,
  note text,
  updated_at text default (datetime('now')),
  primary key (file_id, no)
);

create table if not exists media_review (
  sha256 text primary key,
  status text default 'pending'
       check (status in ('pending','approved','reference','forbidden')),
  tags text, note text,
  updated_at text default (datetime('now'))
);

create table if not exists design_system (
  id integer primary key,
  version integer not null unique,
  json text,                            -- 旧版 PPT 草稿（兼容读取；新规范走 content_md）
  content_md text,                      -- 规范正文（markdown；skill 导入/人工编写）
  status text default 'draft' check (status in ('draft','confirmed')),
  confirmed_by text, confirmed_at text,
  source_file_id integer references file(id) on delete set null,
  source_extraction_id integer,
  schema_version integer default 1,
  unique (source_extraction_id)
);

create virtual table if not exists slide_fts using fts5(title, body, tokenize='trigram');

create trigger if not exists slide_fts_del after delete on slide begin
  delete from slide_fts where rowid = old.id;
end;

create view if not exists v_active_slide as
select s.* from slide s
join extraction e on e.id = s.extraction_id and e.status='done';
"""


class InvariantViolation(RuntimeError):
    """done 目录缺失等不变量破坏：不改状态、不重发，人工处理（§4.4）。"""


class InvalidState(RuntimeError):
    pass


class Conflict(Exception):
    def __init__(self, msg: str, status: int = 409):
        super().__init__(msg)
        self.status = status


@dataclass
class DeadMedia:
    """gc_media_locked 返回的具名行（编码注意事项 #4：防字段语义漂移）。"""
    path: str | None
    preview: str | None


class Database:
    def __init__(self, cfg: Config):
        self.cfg = cfg
        self._local = threading.local()
        cfg.db_path.parent.mkdir(parents=True, exist_ok=True)
        self._init_schema()

    # ---- 连接 ----
    def connect(self, busy_ms: int | None = None) -> sqlite3.Connection:
        conn = sqlite3.connect(self.cfg.db_path, timeout=(busy_ms or self.cfg.busy_timeout_ms) / 1000,
                               isolation_level=None)  # autocommit；事务显式管理
        conn.row_factory = sqlite3.Row
        conn.execute(f"pragma journal_mode=WAL")
        conn.execute("pragma foreign_keys=ON")
        conn.execute(f"pragma busy_timeout={busy_ms or self.cfg.busy_timeout_ms}")
        return conn

    @property
    def conn(self) -> sqlite3.Connection:
        c = getattr(self._local, "conn", None)
        if c is None:
            c = self.connect()
            self._local.conn = c
        return c

    def _init_schema(self) -> None:
        conn = self.connect()
        try:
            conn.executescript(SCHEMA)   # executescript 自带事务管理
            # 旧库迁移：extraction 补 design_json 列、design_system 补 content_md 列（幂等）
            cols = {r[1] for r in conn.execute("pragma table_info(extraction)")}
            if "design_json" not in cols:
                conn.execute("alter table extraction add column design_json json")
            cols = {r[1] for r in conn.execute("pragma table_info(design_system)")}
            if "content_md" not in cols:
                conn.execute("alter table design_system add column content_md text")
            cols = {r[1] for r in conn.execute("pragma table_info(file)")}
            if "doc_type" not in cols:
                conn.execute("alter table file add column doc_type text")
        finally:
            conn.close()

    # ---- 事务 ----
    @contextmanager
    def tx(self, busy: str = "normal") -> Iterator[sqlite3.Cursor]:
        """显式短事务。busy='short' 用于 fslock 内（§10）。"""
        ms = self.cfg.busy_timeout_short_ms if busy == "short" else self.cfg.busy_timeout_ms
        conn = self.connect(busy_ms=ms)
        try:
            cur = conn.cursor()
            cur.execute("begin immediate")   # 写事务直接取写锁，尽早暴露 BUSY
            yield cur
            cur.execute("commit")
        except BaseException:
            try:
                conn.execute("rollback")
            except sqlite3.Error:
                pass
            raise
        finally:
            conn.close()

    # ---- 复合操作执行器（§4.0 唯一锁序） ----
    def run_locked(self, op: Callable[[sqlite3.Cursor], list[Path]]) -> None:
        """外层取 fslock → 短 DB 事务 → commit → 仍在锁内搬 trash → 解锁。

        op(cur) 返回"待搬运清单"（事务成功路径）；BUSY 重试=重试整个复合操作（含锁）。
        """
        self._retry_locked(lambda: self._locked_once(op))

    def run_already_locked(self, op: Callable[[sqlite3.Cursor], list[Path]]) -> None:
        """调用方已持有 fslock 时使用（§4.5 _mark_failed_locked / §4.4 finalize 等）：
        不再取锁，其余语义与 run_locked 一致（BUSY 重试由调用方的锁作用域整体重试）。"""
        self._locked_once(op)

    def _locked_once(self, op) -> None:
        paths: list[Path] = []
        with self.tx(busy="short") as cur:
            paths = op(cur) or []
        for p in paths:
            to_trash(p, self.cfg.trash_dir)

    def _retry_locked(self, body: Callable[[], None]) -> None:
        last_err: Exception | None = None
        for attempt in range(self.cfg.busy_retry_max):
            paths: list[Path] = []
            try:
                with fslock:
                    body()
                return
            except sqlite3.OperationalError as e:
                if "locked" in str(e).lower() or "busy" in str(e).lower():
                    last_err = e
                    time.sleep((0.05 + random.random() * 0.1) * (attempt + 1))
                    continue
                raise
            except Exception:
                if _DIAG:
                    import traceback
                    traceback.print_stack()
                    for th in threading.enumerate():
                        print(f"  thread: {th.name} alive={th.is_alive()}")
                raise
        raise last_err if last_err else RuntimeError("run_locked 重试耗尽")

    # ---- 查询助手 ----
    def one(self, sql: str, *args) -> sqlite3.Row | None:
        return self.conn.execute(sql, args).fetchone()

    def all(self, sql: str, *args) -> list[sqlite3.Row]:
        return self.conn.execute(sql, args).fetchall()

    def close(self) -> None:
        c = getattr(self._local, "conn", None)
        if c is not None:
            c.close()
            self._local.conn = None


def norm_tags(raw: str | None) -> str:
    """标签规范化（§5）：',a,b,'；拆分→trim→丢空→去重；标签值禁含逗号。"""
    if not raw:
        return ","
    seen: list[str] = []
    for part in raw.replace("，", ",").split(","):
        t = part.strip()
        if t and t not in seen:
            seen.append(t)
    return "," + ",".join(seen) + "," if seen else ","
