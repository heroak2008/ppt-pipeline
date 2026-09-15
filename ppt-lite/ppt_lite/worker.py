"""worker.py：后台单线程（领取→处理→finalize；异常→mark_failed）。"""
from __future__ import annotations

import json
import logging
import threading
from pathlib import Path

from .config import Config
from .db import Database
from .pipeline import extract_document
from .store import Store

log = logging.getLogger("pptlite.worker")


class Worker:
    def __init__(self, db: Database, cfg: Config):
        self.db = db
        self.cfg = cfg
        self.store = Store(db, cfg)
        self._wake = threading.Event()
        self._stop = threading.Event()
        self._thread: threading.Thread | None = None

    def wake(self) -> None:
        self._wake.set()

    def stop(self) -> None:
        self._stop.set()
        self._wake.set()

    def start(self) -> None:
        self._thread = threading.Thread(target=self._run, name="pptlite-worker", daemon=True)
        self._thread.start()

    # ------------------------------------------------------------------
    def _run(self) -> None:
        while not self._stop.is_set():
            try:
                job = self.store.claim()
                if job is None:
                    self._wake.wait(self.cfg.poll_interval_sec)
                    self._wake.clear()
                    continue
                ex_id, file_id = job
                self._process(ex_id, file_id)
            except Exception:  # noqa: BLE001
                log.exception("worker 循环异常")
                self._wake.wait(1.0)
                self._wake.clear()

    def _process(self, ex_id: int, file_id: int) -> None:
        f = self.db.one("select * from file where id=?", file_id)
        if f is None:
            return
        src_rel = f["derived_path"] or f["raw_path"]
        src = self.cfg.data_dir / src_rel
        try:
            doc = extract_document(self.cfg, ex_id, src, f["category"],
                                   f["name"], f["raw_path"], f["derived_path"])
            # 写 DB：逐页短事务 + 媒体（锁内复合操作）
            sha_to_mid: dict[str, int] = {}
            staging = self.cfg.tmp_dir / str(ex_id) / "media_staging"
            for m in doc.media:
                staged = staging / _staged_name(staging, m["sha256"])
                mid = self.store.upsert_media(
                    m["sha256"], m["fmt"], m["w"], m["h"], m["phash"], m["path"], m["preview"],
                    staged)
                sha_to_mid[m["sha256"]] = mid
            for s in doc.slides:
                self.store.write_slide(ex_id, file_id, s)
                for occ in [o for o in doc.occurrences if o["slide_no"] == s["no"]]:
                    mid = sha_to_mid.get(occ["media_sha256"])
                    if mid:
                        sid = self.store.slide_row_id(ex_id, s["no"])
                        if sid:
                            self.store.add_media_ref(sid, mid, occ["role"])
            # finalize（文件先发布→DB 后切换；幂等；LO 降级时跳过发布）
            self.store.finalize(ex_id, len(doc.slides), has_preview=doc.render_error is None)
            # 规范数据源入库（主题色/字体/页面尺寸；保留在 design_json 供参考）
            if doc.design:
                with self.db.tx() as cur:
                    cur.execute("update extraction set design_json=? where id=?",
                                (json.dumps(doc.design, ensure_ascii=False), ex_id))
            # 注：规范不再从 PPT 自动生成草稿——规范以 markdown 为承载，
            # 由 skill/人工经「模板与规范」页导入（见 §4.9-lite 变更）
        except Exception as e:  # noqa: BLE001 — 任何异常 → 失败清理（旧 done 不动）
            log.exception("处理失败 ex=%s", ex_id)
            self.store.mark_failed(ex_id, str(e)[:500])


def _staged_name(staging: Path, sha: str) -> str:
    for child in staging.iterdir():
        if child.stem == sha:
            return child.name
    return sha   # 不存在时 upsert_media 会跳过写盘（同批次前页已发布）
