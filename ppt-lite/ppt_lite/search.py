"""search.py：FTS + LIKE 回退（§5 两条完整 SQL）。"""
from __future__ import annotations

import re

from .db import Database, norm_tags

_CJK = re.compile(r"[\u4e00-\u9fff]{3,}")


def _escape_like(q: str) -> str:
    return q.replace("\\", "\\\\").replace("%", r"\%").replace("_", r"\_")


def search_slides(db: Database, q: str, *, slide_type: str | None = None,
                  tag: str | None = None, limit: int = 100) -> list[dict]:
    q = q.strip()
    extra = ""
    args: list = []
    if slide_type:
        extra += " and sr.slide_type = ?"
        args.append(slide_type)
    if tag:
        extra += " and sr.tags like '%,'||?||',%'"
        args.append(norm_tags(tag).strip(","))
    if _CJK.fullmatch(q):
        phrase = '"' + q.replace('"', '""') + '"'
        sql = f"""
        select s.id, s.file_id, s.no, s.title, s.thumb, coalesce(sr.status,'pending') as review_status
        from slide_fts f
        join v_active_slide s on s.id = f.rowid
        left join slide_review sr on sr.file_id = s.file_id and sr.no = s.no
        where slide_fts match ?
          and coalesce(sr.status,'pending') <> 'forbidden'{extra}
        order by rank limit ?"""
        args_all = [phrase] + args + [limit]
    else:
        pattern = f"%{_escape_like(q)}%"
        sql = f"""
        select s.id, s.file_id, s.no, s.title, s.thumb, coalesce(sr.status,'pending') as review_status
        from v_active_slide s
        left join slide_review sr on sr.file_id = s.file_id and sr.no = s.no
        where s.search_text like ? escape '\\'
          and coalesce(sr.status,'pending') <> 'forbidden'{extra}
        limit ?"""
        args_all = [pattern] + args + [limit]
    return [dict(r) for r in db.all(sql, *args_all)]
