"""store.py：§4 全部状态入口（领取/finalize/mark_failed/删除/启动恢复/GC）。

每个函数对应 lite-solution.md §4.x 的伪码；复合操作一律遵循唯一锁序：
外层 fslock → 短 DB 事务 → commit → 仍在锁内搬 trash → 解锁（绝不在持锁路径内再取锁）。
"""
from __future__ import annotations

import logging
import random
import shutil
import sqlite3
import time
from pathlib import Path

from . import PIPELINE_VERSION
from .config import Config
from .db import Conflict, Database, DeadMedia, InvariantViolation, InvalidState
from .fslock import fslock, to_trash, to_trash_if_exists, purge_trash

log = logging.getLogger("pptlite.store")


class Store:
    def __init__(self, db: Database, cfg: Config):
        self.db = db
        self.cfg = cfg

    # ---------- §4.1 上传登记（锁内一气呵成：文件先发布→DB 后提交） ----------
    def register_upload(self, sha: str, name: str, category: str,
                        raw_rel: str, derived_rel: str | None,
                        raw_src: Path, derived_src: Path | None,
                        doc_type: str | None = None) -> tuple[int, bool]:
        """§4.1 锁内：发布 raw/derived（文件先）→ 单事务登记 file+extraction（后）。
        返回 (file_id, created)。sha 已存在 → 复用不建任务。
        重试耗尽/非重试异常的补偿（§4.1 ※）：锁内复查 DB 无该 sha 行 → 已发布文件进 trash。
        """
        raw_path = self.cfg.data_dir / raw_rel
        dpath = (self.cfg.data_dir / derived_rel) if derived_rel else None
        result: dict = {}
        last_err: Exception | None = None

        for attempt in range(self.cfg.busy_retry_max):
            try:
                with fslock:
                    # ① 文件先发布（内容寻址幂等；已存在则保留暂存由调用方清理）
                    if not raw_path.exists() and raw_src.exists():
                        shutil.move(str(raw_src), str(raw_path))
                    if dpath is not None and not dpath.exists() \
                            and derived_src is not None and derived_src.exists():
                        shutil.move(str(derived_src), str(dpath))
                    # ② DB 后提交（短事务）
                    with self.db.tx(busy="short") as cur:
                        row = cur.execute("select id from file where sha256=?", (sha,)).fetchone()
                        if row:
                            result["fid"], result["created"] = row["id"], False
                        else:
                            cur.execute(
                                "insert into file(sha256, name, category, status, raw_path, derived_path, doc_type)"
                                " values(?,?,?,?,?,?,?)",
                                (sha, name, category, "queued", raw_rel, derived_rel, doc_type or None))
                            fid = cur.lastrowid
                            cur.execute(
                                "insert into extraction(file_id, status, parser_version) values(?,?,?)",
                                (fid, "queued", PIPELINE_VERSION))
                            result["fid"], result["created"] = fid, True
                return result["fid"], result["created"]
            except sqlite3.OperationalError as e:
                if "locked" in str(e).lower() or "busy" in str(e).lower():
                    last_err = e
                    time.sleep((0.05 + random.random() * 0.1) * (attempt + 1))
                    continue
                self._compensate_upload(sha, raw_path, dpath)
                raise
            except Exception:
                self._compensate_upload(sha, raw_path, dpath)
                raise
        self._compensate_upload(sha, raw_path, dpath)
        raise last_err if last_err else RuntimeError("register_upload 重试耗尽")

    def _compensate_upload(self, sha: str, raw_path: Path, dpath: Path | None) -> None:
        """§4.1 ※：DB 最终失败时，仅当无 file 行才把已发布文件搬 trash（防误删共享内容）。"""
        try:
            with fslock:
                if self.db.one("select 1 from file where sha256=?", sha) is None:
                    to_trash(raw_path, self.cfg.trash_dir)
                    if dpath is not None:
                        to_trash(dpath, self.cfg.trash_dir)
        except Exception:  # noqa: BLE001
            log.exception("upload 补偿失败（留待启动 GC 对账）sha=%s", sha)

    # ---------- §4.2 领取（含防饿死） ----------
    def claim(self) -> tuple[int, int] | None:
        """领取一个 queued extraction → running。返回 (ex_id, file_id) 或 None。"""
        with self.db.tx() as cur:
            row = cur.execute(
                "update extraction set status='running', started_at=datetime('now'), error=null"
                " where id = (select id from extraction"
                "             where status='queued'"
                "               and file_id not in (select file_id from extraction where status='running')"
                "             order by id limit 1)"
                " returning id, file_id").fetchone()
            if row is None:
                return None             # 空事务 commit 即可（tx 上下文自动处理）
            cur.execute("update file set status='working', error=null where id=?", (row["file_id"],))
            return row["id"], row["file_id"]

    # ---------- §4.3 写入 ----------
    def write_slide(self, ex_id: int, file_id: int, s: dict) -> None:
        """逐页短事务（无文件操作，无需锁）。"""
        with self.db.tx() as cur:
            cur.execute(
                "insert into slide(extraction_id, file_id, no, title, search_text, texts, struct, png, thumb)"
                " values(?,?,?,?,?,?,?,?,?)",
                (ex_id, file_id, s["no"], s.get("title"), s.get("search_text"),
                 _j(s.get("texts")), _j({"struct": s.get("struct"), "notes": s.get("notes"),
                                         "hidden": s.get("hidden"), "warnings": s.get("warnings")}),
                 s["png"], s["thumb"]))
            sid = cur.lastrowid
            cur.execute("insert into slide_fts(rowid, title, body) values(?,?,?)",
                        (sid, s.get("title") or "", s.get("search_text") or ""))

    # ---------- 素材目录导入（独立于 PPT 的素材入库） ----------
    def import_media_file(self, data: bytes, ext: str) -> tuple[str, bool]:
        """单个素材文件入库（内容寻址去重）。返回 (sha256, created)。

        文件先发布（锁内）→ DB 后提交；已存在则复用（created=False，不重复写盘）。
        """
        import hashlib
        sha = hashlib.sha256(data).hexdigest()
        fmt = ext.lstrip(".").lower()
        path = f"media/{sha}.{fmt}"
        canonical = Path(self.cfg.data_dir) / path

        w = h = None
        phash = None
        if fmt in {"png", "jpg", "jpeg", "gif", "bmp", "webp", "tiff", "ico"}:
            try:
                import io as _io
                import imagehash as _ih
                from PIL import Image as _Image
                img = _Image.open(_io.BytesIO(data))
                w, h = img.size
                phash = str(_ih.phash(img))
            except Exception:  # noqa: BLE001
                pass
        # svg/emf/wmf：浏览器可直接渲染 svg（preview=path）；emf/wmf 暂无预览（占位）
        preview = path if fmt == "svg" else None

        existing = self.db.one("select id from media where sha256=?", sha)
        if existing:
            return sha, False
        with fslock:
            if not canonical.exists():
                canonical.parent.mkdir(parents=True, exist_ok=True)
                canonical.write_bytes(data)
            with self.db.tx(busy="short") as cur:
                cur.execute(
                    "insert into media(sha256, fmt, w, h, phash, path, preview) values(?,?,?,?,?,?,?)"
                    " on conflict(sha256) do nothing",
                    (sha, fmt, w, h, phash, path, preview))
        return sha, True

    def upsert_media(self, sha: str, fmt: str, w, h, phash, path: str, preview, staging_file: Path) -> int:
        """媒体落盘（§4.3 铁律⑤：文件先发布→DB 后提交；锁内）。返回 media_id。"""
        canonical = Path(self.cfg.data_dir) / path
        trash: list[Path] = []
        with fslock:
            if not canonical.exists() and staging_file.exists():
                shutil.move(str(staging_file), str(canonical))
            elif staging_file.exists():
                staging_file.unlink()          # 规范文件已在（同批次前页发布过）→ 丢弃暂存
            with self.db.tx(busy="short") as cur:
                row = cur.execute(
                    "insert into media(sha256, fmt, w, h, phash, path, preview)"
                    " values(?,?,?,?,?,?,?)"
                    " on conflict(sha256) do update set phash=coalesce(media.phash, excluded.phash)"
                    " returning id",
                    (sha, fmt, w, h, phash, path, preview)).fetchone()
                mid = row["id"]
        # DB 失败补偿（确定规则 §4.3）：仅当 DB 无该 sha 的 media 行时规范文件才是孤儿
        return mid

    def add_media_ref(self, slide_row_id: int, media_id: int, role: str) -> None:
        with self.db.tx() as cur:
            cur.execute("insert or ignore into media_ref(slide_id, media_id, role) values(?,?,?)",
                        (slide_row_id, media_id, role))

    def slide_row_id(self, ex_id: int, no: int) -> int | None:
        row = self.db.one("select id from slide where extraction_id=? and no=?", ex_id, no)
        return row["id"] if row else None

    # ---------- §4.4 发布与 finalization（幂等；render_error 降级=跳过文件发布） ----------
    def finalize(self, ex_id: int, slide_count: int, has_preview: bool = True) -> None:
        """文件先发布（rename staging→previews/{ex}）→ DB 后切换。幂等。

        has_preview=False（LO 缺失/转换失败的降级）：跳过步骤 A，仅做 DB 切换；
        slide.png/thumb 保留占位路径（UI 显示"无预览"）。
        """
        tmp = self.cfg.tmp_dir / str(ex_id)
        staging = self.cfg.staging_dir / str(ex_id)
        target = self.cfg.previews_dir / str(ex_id)

        ex = self.db.one("select file_id, status from extraction where id=?", ex_id)
        if ex is None:
            to_trash(tmp, self.cfg.trash_dir)
            return
        if ex["status"] == "done":
            if has_preview and not _dir_complete(target, slide_count):
                raise InvariantViolation(f"done 目录缺失/不完整: {ex_id}（人工处理或重跑生成新 extraction）")
            return
        if ex["status"] != "running":
            raise InvalidState(f"extraction {ex_id} 状态 {ex['status']}，不可 finalize")

        with fslock:
            # A. 文件发布（rename 前失败 → raise，调用方走 mark_failed；旧 done 未动）
            if has_preview:
                to_trash(staging, self.cfg.trash_dir)
                preview_src = tmp / "preview"
                staging.mkdir(parents=True, exist_ok=True)
                shutil.copytree(preview_src, staging, dirs_exist_ok=True)
                _verify_complete(staging, slide_count)
                if target.exists():             # 上次发布后、DB 切换前崩溃的孤儿
                    to_trash(target, self.cfg.trash_dir)
                _atomic_rename(staging, target)  # 同卷原子；Windows 杀软瞬时锁 → 短退避重试

            # B. DB finalization（单事务）
            dead: list[DeadMedia] = []
            old_ex_id = None

            def op(cur: sqlite3.Cursor) -> list[Path]:
                nonlocal dead, old_ex_id
                row = cur.execute(
                    "select id from extraction where file_id=? and status='done' and id<>?",
                    (ex["file_id"], ex_id)).fetchone()
                if row:
                    old_ex_id = row["id"]
                    mids = _collect_media_ids(cur, extraction=old_ex_id)
                    cur.execute("delete from extraction where id=?", (old_ex_id,))  # 级联 slide/FTS/media_ref
                cur2 = cur.execute(
                    "update extraction set status='done', finished_at=datetime('now')"
                    " where id=? and status='running'", (ex_id,))
                assert cur2.rowcount == 1, "finalize 条件迁移失败（迟到回调）"
                cur.execute("update file set status='done', error=null where id=?", (ex["file_id"],))
                if row:
                    dead = self._gc_media(cur, mids)
                paths = [tmp]
                if old_ex_id is not None:
                    paths.append(self.cfg.previews_dir / str(old_ex_id))
                paths += [Path(self.cfg.data_dir / d.path) for d in dead if d.path]
                paths += [Path(self.cfg.data_dir / d.preview) for d in dead if d.preview]
                return paths

            try:
                self.db.run_already_locked(op)
            except Exception:
                # B 失败 → previews/{ex} 成为可 GC 孤儿；调用方走 mark_failed（旧 done 未动）
                raise

    # ---------- §4.5 失败清理（幂等；双入口） ----------
    def mark_failed(self, ex_id: int, reason: str) -> None:
        with fslock:
            self._mark_failed_locked(ex_id, reason)

    def _mark_failed_locked(self, ex_id: int, reason: str) -> None:
        row = self.db.one("select file_id from extraction where id=? and status='running'", ex_id)
        if not row:
            to_trash_if_exists(self.cfg.tmp_dir / str(ex_id), self.cfg.trash_dir)
            return

        dead: list[DeadMedia] = []

        def op(cur: sqlite3.Cursor) -> list[Path]:
            mids = _collect_media_ids(cur, extraction=ex_id)
            cur.execute("delete from slide where extraction_id=?", (ex_id,))
            cur2 = cur.execute(
                "update extraction set status='failed', error=?, finished_at=datetime('now')"
                " where id=? and status='running'", (reason, ex_id))
            assert cur2.rowcount == 1
            cur.execute("update file set status='failed', error=? where id=?", (reason, row["file_id"]))
            nonlocal dead
            dead = self._gc_media(cur, mids)
            paths = [self.cfg.tmp_dir / str(ex_id), self.cfg.previews_dir / str(ex_id)]
            paths += [Path(self.cfg.data_dir / d.path) for d in dead if d.path]
            paths += [Path(self.cfg.data_dir / d.preview) for d in dead if d.preview]
            return paths

        self.db.run_already_locked(op)

    # ---------- §4.8 删除 ----------
    def delete_file(self, fid: int) -> None:
        info_holder: dict = {}

        def op(cur: sqlite3.Cursor) -> list[Path]:
            info = cur.execute("select raw_path, derived_path from file where id=?", (fid,)).fetchone()
            if info is None:
                raise Conflict("文件不存在", status=404)
            ex_ids = [r["id"] for r in cur.execute(
                "select id from extraction where file_id=?", (fid,)).fetchall()]
            mids = _collect_media_ids(cur, file=fid)
            cur3 = cur.execute(
                "delete from file where id=?"
                " and not exists (select 1 from extraction e where e.file_id=file.id"
                "                 and e.status in ('queued','running'))", (fid,))
            if cur3.rowcount == 0:
                raise Conflict("文件有排队/处理中任务，不能删除")
            dead = self._gc_media(cur, mids)
            info_holder["info"] = dict(info)
            info_holder["ex_ids"] = ex_ids
            paths = [Path(self.cfg.data_dir) / info["raw_path"]]
            if info["derived_path"]:
                paths.append(Path(self.cfg.data_dir) / info["derived_path"])
            for ex in ex_ids:
                paths.append(self.cfg.previews_dir / str(ex))
                paths.append(self.cfg.tmp_dir / str(ex))
            paths += [Path(self.cfg.data_dir / d.path) for d in dead if d.path]
            paths += [Path(self.cfg.data_dir / d.preview) for d in dead if d.preview]
            return paths

        self.db.run_locked(op)

    # ---------- §4.10 启动恢复 + GC ----------
    def startup_recover_and_gc(self) -> None:
        with fslock:
            # 2. 启动恢复：running → failed（locked 变体，不重复取锁）
            for row in self.db.all("select id from extraction where status='running'"):
                self._mark_failed_locked(row["id"], "进程中断")
            # 3. GC（停流窗口；worker/HTTP 未启动）
            self._gc()
            # e. trash 物理删除
            purge_trash(self.cfg.trash_dir)

    def _gc(self) -> None:
        cfg = self.cfg
        # a. previews：DB 无此 extraction 或 failed → trash（done 保留）
        if cfg.previews_dir.exists():
            for child in cfg.previews_dir.iterdir():
                if child.name == ".staging":
                    continue
                if not child.name.isdigit():
                    to_trash(child, cfg.trash_dir)
                    continue
                ex = self.db.one("select status from extraction where id=?", int(child.name))
                if ex is None or ex["status"] == "failed":
                    to_trash(child, cfg.trash_dir)
        # b. staging、tmp
        if cfg.staging_dir.exists():
            for child in cfg.staging_dir.iterdir():
                to_trash(child, cfg.trash_dir)
        if cfg.tmp_dir.exists():
            for child in cfg.tmp_dir.iterdir():
                to_trash(child, cfg.trash_dir)
        # c. media 无行文件 → trash
        if cfg.media_dir.exists():
            for child in cfg.media_dir.iterdir():
                sha = child.stem
                if not self.db.one("select 1 from media where sha256=?", sha):
                    to_trash(child, cfg.trash_dir)
        # d. raw/derived 无行文件 → trash（按文件名 sha 匹配）
        for d in (cfg.raw_dir, cfg.derived_dir):
            if not d.exists():
                continue
            for child in d.iterdir():
                if child.is_dir():
                    continue
                sha = child.stem
                if not self.db.one("select 1 from file where sha256=?", sha):
                    to_trash(child, cfg.trash_dir)

    # ---------- helpers ----------
    def _gc_media(self, cur: sqlite3.Cursor, mids: list[int]) -> list[DeadMedia]:
        """删无引用 media 行；返回具名行清单（搬运由外层 commit 后执行）。"""
        dead = []
        for m in set(mids):
            if not cur.execute("select 1 from media_ref where media_id=? limit 1", (m,)).fetchone():
                dead.append(m)
        out: list[DeadMedia] = []
        if not dead:
            return out
        qmarks = ",".join("?" * len(dead))
        for r in cur.execute(f"select path, preview from media where id in ({qmarks})", dead).fetchall():
            out.append(DeadMedia(path=r["path"], preview=r["preview"]))
        cur.execute(f"delete from media where id in ({qmarks})", dead)
        return out


def _collect_media_ids(cur: sqlite3.Cursor, *, extraction: int | None = None, file: int | None = None) -> list[int]:
    if extraction is not None:
        rows = cur.execute(
            "select distinct mr.media_id from media_ref mr"
            " join slide s on s.id=mr.slide_id where s.extraction_id=?", (extraction,)).fetchall()
    else:
        rows = cur.execute(
            "select distinct mr.media_id from media_ref mr"
            " join slide s on s.id=mr.slide_id"
            " join extraction e on e.id=s.extraction_id where e.file_id=?", (file,)).fetchall()
    return [r["media_id"] for r in rows]


def _dir_complete(target: Path, slide_count: int) -> bool:
    if not target.is_dir():
        return False
    return all((target / "slides" / f"{i}.png").exists() and (target / "thumbs" / f"{i}.png").exists()
               for i in range(1, slide_count + 1))


def _verify_complete(staging: Path, slide_count: int) -> None:
    if not _dir_complete(staging, slide_count):
        raise InvalidState(f"发布校验失败：staging 不完整（期望 {slide_count} 页）")


def _atomic_rename(src: Path, dst: Path, attempts: int = 5) -> None:
    """目录原子 rename；Windows 上杀软/索引器对新写文件有瞬时锁，短退避重试。"""
    import time as _t
    for i in range(attempts):
        try:
            src.rename(dst)
            return
        except PermissionError:
            if i == attempts - 1:
                raise
            _t.sleep(0.2 * (i + 1))


def _j(obj) -> str:
    import json
    return json.dumps(obj, ensure_ascii=False)
