"""design_draft.py：spec 文件 → design_system 自动草稿（§4.9；幂等、不覆盖人工）。

草稿内容 = 主题色板 + 主题字体 + 页面尺寸 + 页面结构统计 + 待人工确认的规则骨架。
数据源：extraction.design_json（pipeline 在解析时从 theme1.xml/母版提取）。
"""
from __future__ import annotations

import json
import logging
from collections import Counter

from .db import Database

log = logging.getLogger("pptlite.design")


def maybe_generate(db: Database, ex_id: int) -> None:
    """finalize 之后调用（worker 已先写 design_json）。失败仅记日志，不影响 done。"""
    try:
        ex = db.one(
            "select e.file_id, e.design_json, f.category from extraction e join file f on f.id=e.file_id"
            " where e.id=? and e.status='done'", ex_id)
        if not ex or ex["category"] != "spec":
            return
        if db.one("select 1 from design_system where source_extraction_id=?", ex_id):
            return   # 幂等：同一解析结果至多一份自动草稿

        design = json.loads(ex["design_json"] or "{}")
        colors = [c["rgb"] for c in design.get("theme_colors", []) if c.get("rgb")]
        fonts = design.get("theme_fonts", [])
        page_size = design.get("page_size") or {}
        font_usage = design.get("font_usage", [])
        fill_usage = design.get("fill_usage", [])

        # 页面统计：标题字数分布、每页形状数（信息密度经验）
        rows = db.all(
            "select s.no, s.title, s.search_text, s.struct from v_active_slide s"
            " join extraction e on e.id=s.extraction_id where e.id=? order by s.no", ex_id)
        title_lens = [len((r["title"] or "").strip()) for r in rows if (r["title"] or "").strip()]
        shape_counts = []
        for r in rows:
            st = (json.loads(r["struct"] or "{}").get("struct")) or {}
            shape_counts.append(len(st.get("shapes") or []))

        # ── 文字规则候选条款：规范 PPT 里的文字描述 + 关联示例页码 ──
        # 规范 PPT 的典型结构：一页讲一条/几条规则（文字），旁边配样式示例。
        # 提取每页全部文字作为候选条款，页码即"示例页"证据；人工逐条勾选/改写进正式规则。
        rule_candidates = []
        for r in rows:
            text = (r["search_text"] or "").strip()
            if not text:
                continue
            rule_candidates.append({
                "page": r["no"],
                "text": text[:800],                    # 截断防超大页
                "adopted": None,                        # 人工确认：True 采纳 / False 忽略 / None 未审
            })

        draft = {
            "schema_version": 1,
            "generated_from": {"extraction_id": ex_id, "file_id": ex["file_id"]},
            "deck": {
                "width_emu": page_size.get("width_emu"),
                "height_emu": page_size.get("height_emu"),
                "aspect": _aspect(page_size),
            },
            "colors": {
                "theme_palette": colors,
                "fill_usage": fill_usage,          # 形状实际填充色直方图（真实使用 ≠ 主题声明）
                "note": "theme_palette 来自主题声明；fill_usage 是形状实采。请人工标注 primary/secondary/forbidden 及色差容忍",
            },
            "typography": {
                "theme_fonts": fonts,
                "font_usage": font_usage,          # run 级实采：字体/字号/加粗 使用频次
                "title_chars_median": _median(title_lens),
                "note": "请人工确定各级字号规则（标题/副标/正文/注释的最小-最大字号、字重）；font_usage 为实际使用分布",
            },
            "asset_rules": {
                "logo":      {"note": "使用位置、最小尺寸、留白、禁用变形/改色（人工补充）"},
                "icons":     {"note": "风格（线性/面性）、线宽、配色来源、禁用图标（人工补充）"},
                "imagery":   {"note": "照片/插图风格、版权范围、禁用图库（人工补充）"},
                "charts":    {"note": "允许的图表类型、系列配色取自 theme_palette 的顺序（人工补充）"},
                "tables":    {"note": "表头样式、斑马纹、对齐规则（人工补充）"},
            },
            # 规范 PPT 中的文字规则条款（原文 + 示例页码），人工逐条审
            "rule_candidates": rule_candidates,
            "page_stats": {
                "page_count": len(rows),
                "shapes_per_page_avg": round(sum(shape_counts) / len(shape_counts), 1) if shape_counts else 0,
            },
            "common_rules": {
                "note": "人工补充：页脚/页码/Logo 是否必须、标题结论先行、每页信息密度上限等",
            },
        }
        with db.tx() as cur:
            cur.execute(
                "insert into design_system(version, json, status, source_file_id, source_extraction_id)"
                " values(?,?,?,?,?)",
                (_next_version(cur), json.dumps(draft, ensure_ascii=False, indent=2), "draft",
                 ex["file_id"], ex_id))
    except Exception:  # noqa: BLE001 — 独立失败面
        log.exception("规范草稿生成失败 ex=%s（可重跑）", ex_id)


def _next_version(cur) -> int:
    row = cur.execute("select max(version) v from design_system").fetchone()
    return (row["v"] or 0) + 1


def _median(vals: list[int]) -> int | None:
    if not vals:
        return None
    s = sorted(vals)
    return s[len(s) // 2]


def _aspect(page_size: dict) -> str | None:
    w, h = page_size.get("width_emu"), page_size.get("height_emu")
    if not w or not h:
        return None
    ratio = w / h
    for name, target in (("16:9", 16 / 9), ("4:3", 4 / 3), ("16:10", 16 / 10)):
        if abs(ratio - target) < 0.02:
            return name
    return f"{ratio:.2f}"
