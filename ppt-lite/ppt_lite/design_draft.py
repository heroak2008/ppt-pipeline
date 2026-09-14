"""design_draft.py：spec 文件 → design_system 自动草稿（§4.9；幂等、不覆盖人工）。"""
from __future__ import annotations

import json
import logging

from .db import Database

log = logging.getLogger("pptlite.design")


def maybe_generate(db: Database, ex_id: int) -> None:
    """finalize 步骤 D（锁外）调用。失败仅记日志，不影响已 done 的 extraction。"""
    try:
        ex = db.one(
            "select e.file_id, f.category from extraction e join file f on f.id=e.file_id"
            " where e.id=? and e.status='done'", ex_id)
        if not ex or ex["category"] != "spec":
            return
        if db.one("select 1 from design_system where source_extraction_id=?", ex_id):
            return   # 幂等：同一解析结果至多一份自动草稿
        # 聚合数据源：该文件所有生效页的 struct 统计 + extraction 时存的 design 数据不可得
        # （契约 design 字段保存在 slide.struct 之外由 pipeline 存 json 目录——此处从简：
        #   从页面统计标题字号/主题色占比，生成最小 draft）
        rows = db.all(
            "select s.struct from v_active_slide s"
            " join extraction e on e.id = s.extraction_id where e.id=?", ex_id)
        pages = [json.loads(r["struct"] or "{}") for r in rows]
        draft = {
            "schema_version": 1,
            "page_count": len(pages),
            "generated_from": {"extraction_id": ex_id},
            "common_rules": {"footer_required": _any(pages, "has_footer"),
                             "logo_required": _any(pages, "has_logo")},
            "note": "自动草稿：请人工确认颜色/字号/规则后发布版本",
        }
        with db.tx() as cur:
            cur.execute(
                "insert into design_system(version, json, status, source_file_id, source_extraction_id)"
                " values(?,?,?,?,?)",
                (_next_version(cur), json.dumps(draft, ensure_ascii=False), "draft",
                 ex["file_id"], ex_id))
    except Exception:  # noqa: BLE001 — 独立失败面
        log.exception("规范草稿生成失败 ex=%s（可重跑）", ex_id)


def _next_version(cur) -> int:
    row = cur.execute("select max(version) v from design_system").fetchone()
    return (row["v"] or 0) + 1


def _any(pages: list[dict], key: str) -> bool:
    return any(p.get(key) for p in pages)
