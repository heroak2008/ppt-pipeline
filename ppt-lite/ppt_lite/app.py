"""app.py：FastAPI 入口（上传/浏览/搜索/审核/删除 + 启动序列 §4.10）。"""
from __future__ import annotations

import json
import logging
import shutil
import sys
from pathlib import Path

from fastapi import BackgroundTasks, FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.concurrency import run_in_threadpool
from fastapi.responses import HTMLResponse, JSONResponse, Response
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from . import PIPELINE_VERSION
from .config import Config
from .db import Conflict, Database, norm_tags
from .fslock import fslock, to_trash
from .precheck import PrecheckError, precheck
from .search import search_slides
from .store import Store
from .worker import Worker

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(name)s %(levelname)s %(message)s")
log = logging.getLogger("pptlite")

cfg = Config()
cfg.ensure_dirs()
db = Database(cfg)
store = Store(db, cfg)
worker = Worker(db, cfg)

BASE = Path(__file__).parent
if getattr(sys, "frozen", False):              # PyInstaller：资源目录在 _MEIPASS
    BASE = Path(sys._MEIPASS) / "ppt_lite"
templates = Jinja2Templates(directory=str(BASE / "templates"))

app = FastAPI(title="ppt-lite", version=PIPELINE_VERSION)
# /static → data/（previews/media 直读；thumb/png 均为 DB 中相对路径）
app.mount("/static", StaticFiles(directory=str(cfg.data_dir)), name="static")


@app.on_event("startup")
def _startup() -> None:
    # §4.10 启动序列：取锁 → 恢复 → GC → 解锁 → worker → HTTP（FastAPI startup 即 HTTP 前）
    store.startup_recover_and_gc()
    worker.start()


@app.on_event("shutdown")
def _shutdown() -> None:
    worker.stop()


def _footer() -> str:
    lo = f'LO ✓' if cfg.soffice else 'LO ✗（渲染降级）'
    return f'{lo}<br>pipeline {PIPELINE_VERSION}'


# ---------------------------------------------------------------- 页面
@app.get("/", response_class=HTMLResponse)
def page_files(request: Request, category: str | None = None):
    files = [dict(r) for r in _file_rows(category)]
    return templates.TemplateResponse(request, "files.html",
                                      {"request": request, "files": files, "category": category,
                                       "soffice": bool(cfg.soffice), "footer_info": _footer(),
                                       "doc_types": DOC_TYPES})


def _file_rows(category: str | None = None) -> list:
    sql = ("select f.*, (select count(*) from v_active_slide s where s.file_id=f.id) as slide_count,"
           " (select max(status) from extraction e where e.file_id=f.id and e.status='done') as has_done"
           " from file f")
    args: list = []
    if category:
        sql += " where f.category=?"
        args.append(category)
    sql += " order by f.id desc"
    return db.all(sql, *args)


@app.get("/api/files/status")
def api_files_status(category: str | None = None):
    """轻量轮询端点：文件状态/错误/生效页数（前端在 queued/working 时每 2s 拉取）。"""
    return JSONResponse({"files": [dict(r) for r in _file_rows(category)]})


@app.get("/slides", response_class=HTMLResponse)
def page_slides(request: Request, q: str = "", type: str | None = None, tag: str | None = None):
    if q:
        rows = search_slides(db, q, slide_type=type or None, tag=tag or None)
    else:
        sql = ("select s.id, s.file_id, s.no, s.title, s.thumb,"
               " coalesce(sr.status,'pending') as review_status"
               " from v_active_slide s"
               " left join slide_review sr on sr.file_id=s.file_id and sr.no=s.no")
        args: list = []
        if type:
            sql += " where sr.slide_type=?"
            args.append(type)
        sql += " order by s.file_id, s.no limit 200"
        rows = [dict(r) for r in db.all(sql, *args)]
    return templates.TemplateResponse(request, "slides.html",
                                      {"request": request, "rows": rows, "q": q,
                                       "type": type or "", "tag": tag or "",
                                       "footer_info": _footer()})


@app.get("/slides/{sid}", response_class=HTMLResponse)
def page_slide_detail(request: Request, sid: int):
    s = db.one("select * from v_active_slide where id=?", sid)
    if not s:
        raise HTTPException(404)
    struct = json.loads(s["struct"] or "{}")
    texts = json.loads(s["texts"] or "[]")
    sr = db.one("select * from slide_review where file_id=? and no=?", s["file_id"], s["no"])
    refs = db.all("select m.sha256, m.fmt, m.path, mr.role from media_ref mr"
                  " join media m on m.id=mr.media_id where mr.slide_id=?", sid)
    f = db.one("select name, status, error from file where id=?", s["file_id"])
    banner = None
    if f["status"] == "failed":
        banner = f"重跑失败：{f['error']}，以下为上次成功数据"
    # 文件级形状统计（本文件全部生效页聚合）
    sibs = db.all("select texts, struct from v_active_slide where file_id=?", s["file_id"])
    file_stats = {"slides": len(sibs), "textboxes": 0, "tables": 0, "shapes": 0,
                  "smartart": 0, "charts": 0, "ole": 0}
    for r in sibs:
        st = (json.loads(r["struct"] or "{}").get("struct")) or {}
        file_stats["textboxes"] += len(json.loads(r["texts"] or "[]"))
        file_stats["tables"] += len(st.get("tables") or [])
        file_stats["shapes"] += len(st.get("shapes") or [])
        file_stats["smartart"] += st.get("has_smartart") or 0
        file_stats["charts"] += st.get("has_chart") or 0
        file_stats["ole"] += st.get("has_ole") or 0
    # 本页矢量形状（渲染布局缩略图 + 明细）
    page_shapes = (struct.get("struct") or {}).get("shapes") or []
    # 页面尺寸（EMU → 布局图比例）
    page_w = struct.get("struct", {}).get("page_w") or 12192000
    page_h = struct.get("struct", {}).get("page_h") or 6858000
    return templates.TemplateResponse(request, "slide_detail.html",
                                      {"request": request, "s": dict(s), "struct": struct,
                                       "texts": texts, "review": dict(sr) if sr else None,
                                       "refs": [dict(r) for r in refs], "fname": f["name"],
                                       "banner": banner, "file_stats": file_stats,
                                       "page_shapes": page_shapes,
                                       "page_w": page_w, "page_h": page_h,
                                       "footer_info": _footer()})


@app.get("/media", response_class=HTMLResponse)
def page_media(request: Request):
    rows = db.all(
        "select m.*, coalesce(mr.status,'pending') as review_status, mr.tags as rtags,"
        " (select count(*) from media_ref x where x.media_id=m.id) as ref_count"
        " from media m left join media_review mr on mr.sha256=m.sha256"
        " order by ref_count desc limit 300")
    return templates.TemplateResponse(request, "media.html",
                                      {"request": request, "rows": [dict(r) for r in rows],
                                       "footer_info": _footer()})


@app.get("/design", response_class=HTMLResponse)
def page_design(request: Request):
    rows = [dict(r) for r in db.all("select * from design_system order by version desc")]
    tpl_rows = db.all(
        "select s.id, s.file_id, s.no, s.title, s.thumb, sr.template_json"
        " from slide_review sr join v_active_slide s on s.file_id=sr.file_id and s.no=sr.no"
        " where sr.is_template=1 order by s.file_id, s.no")
    return templates.TemplateResponse(request, "design.html",
                                      {"request": request, "rows": rows,
                                       "templates": [dict(t) for t in tpl_rows],
                                       "footer_info": _footer()})


# ---------------------------------------------------------------- API
# 常用业务类型（自由值，下拉仅做快捷项；文件列表可改）
DOC_TYPES = ["洞察材料", "立项材料", "BP材料", "方案汇报", "总结复盘", "产品介绍", "其他"]


@app.post("/api/upload")
async def api_upload(background_tasks: BackgroundTasks,
                     file: UploadFile = File(...), category: str = Form(...),
                     doc_type: str = Form("")):
    if category not in ("material", "spec", "sample"):
        raise HTTPException(422, "category 必须是 material/spec/sample")
    doc_type = doc_type.strip() or None
    data = await file.read()
    if len(data) > cfg.max_upload_bytes:
        raise HTTPException(413, "文件过大")
    name = file.filename or "unnamed.pptx"

    # ── 锁外重活（LO 转换可能分钟级；run_in_threadpool 避免冻结事件循环）──
    def heavy() -> tuple[str, str, Path, Path | None, str | None]:
        staging = cfg.tmp_dir / f"upload-{_uuid4hex()}"
        staging.mkdir(parents=True, exist_ok=True)
        try:
            result = precheck(data, soffice_available=bool(cfg.soffice))
            raw_src = staging / "blob"
            derived_src = None
            derived_rel = None
            if result.kind == "legacy-ppt":
                src = staging / "src.ppt"
                src.write_bytes(data)
                raw_src.write_bytes(data)          # 原件按规范保留（.ppt 原样入 raw/）
                derived = _convert_ppt_to_pptx(src, staging)
                if derived is None:
                    raise PrecheckError("转换失败，请手工另存为 .pptx")
                derived_src = derived
                derived_rel = "placeholder"        # register_upload 内按 sha 组装
            else:
                raw_src.write_bytes(data)
            sha = _sha256_bytes(raw_src.read_bytes())
            return sha, result.kind, raw_src, derived_src, derived_rel
        except Exception:
            to_trash(staging, cfg.trash_dir)       # 预检失败：暂存清理（私有目录，无需锁）
            raise

    try:
        sha, kind, raw_src, derived_src, _ = await run_in_threadpool(heavy)
    except PrecheckError as e:
        raise HTTPException(422, e.reason)

    ext = Path(name).suffix.lower() or ".pptx"
    raw_rel = f"raw/{sha}{ext}"
    derived_rel = f"raw/derived/{sha}.pptx" if kind == "legacy-ppt" else None
    try:
        fid, created = await run_in_threadpool(
            store.register_upload, sha, name, category, raw_rel, derived_rel,
            raw_src, derived_src, doc_type)
    except PrecheckError as e:
        raise HTTPException(422, e.reason)
    worker.wake()
    return {"file_id": fid, "created": created,
            "message": None if created else "同内容文件已存在，复用既有记录"}


@app.post("/api/files/{fid}/attrs")
def api_file_attrs(fid: int, doc_type: str = Form(None), note: str = Form(None)):
    """改文件属性（材料类型/备注）——人工属性，重跑不覆盖。"""
    f = db.one("select id from file where id=?", fid)
    if not f:
        raise HTTPException(404)
    with db.tx() as cur:
        cur.execute("update file set doc_type=coalesce(?, doc_type), note=coalesce(?, note) where id=?",
                    (doc_type.strip() if doc_type is not None else None,
                     note if note is not None else None, fid))
    return {"ok": True}


@app.post("/api/files/{fid}/reprocess")
def api_reprocess(fid: int):
    f = db.one("select id from file where id=?", fid)
    if not f:
        raise HTTPException(404)
    try:
        with db.tx() as cur:
            cur.execute("insert into extraction(file_id, status, parser_version) values(?,?,?)",
                        (fid, "queued", PIPELINE_VERSION))
        worker.wake()
        return {"ok": True}
    except Exception as e:  # noqa: BLE001 — uq_ex_one_active 冲突 → 409
        if "unique" in str(e).lower():
            raise HTTPException(409, "已有任务在排队或处理中")
        raise


@app.post("/api/files/{fid}/delete")
def api_delete(fid: int):
    try:
        store.delete_file(fid)
        return {"ok": True}
    except Conflict as e:
        raise HTTPException(e.status, str(e))


@app.post("/api/slides/{sid}/review")
def api_review(sid: int, status: str = Form(...), slide_type: str = Form(None),
               quality: int = Form(None), tags: str = Form(None),
               is_template: bool = Form(False), template_json: str = Form(None),
               note: str = Form(None)):
    s = db.one("select file_id, no from v_active_slide where id=?", sid)
    if not s:
        raise HTTPException(404)
    if status not in ("pending", "approved", "reference", "forbidden"):
        raise HTTPException(422, "非法 status")
    with db.tx() as cur:
        cur.execute(
            "insert into slide_review(file_id, no, status, slide_type, quality, tags, is_template,"
            " template_json, note, updated_at)"
            " values(?,?,?,?,?,?,?,?,?,datetime('now'))"
            " on conflict(file_id, no) do update set"
            " status=excluded.status, slide_type=excluded.slide_type, quality=excluded.quality,"
            " tags=excluded.tags, is_template=excluded.is_template,"
            " template_json=excluded.template_json, note=excluded.note,"
            " updated_at=datetime('now')",
            (s["file_id"], s["no"], status, slide_type, quality, norm_tags(tags),
             1 if is_template else 0, template_json, note))
    return {"ok": True}


@app.post("/api/design/{did}/confirm")
def api_design_confirm(did: int, confirmed_by: str = Form("")):
    with db.tx() as cur:
        cur2 = cur.execute(
            "update design_system set status='confirmed', confirmed_by=?, confirmed_at=datetime('now')"
            " where id=? and status='draft'", (confirmed_by or "me", did))
        if cur2.rowcount == 0:
            raise HTTPException(409, "该版本不是草稿（可能已确认）")
    return {"ok": True}


# 素材目录导入允许的扩展名
MEDIA_IMPORT_EXTS = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tiff", ".webp", ".ico",
                     ".svg", ".emf", ".wmf"}


@app.post("/api/media/import")
async def api_media_import(files: list[UploadFile] = File(...)):
    """素材目录导入：浏览器选整个文件夹（webkitdirectory）批量上传，按内容寻址去重。"""
    imported = skipped_dup = skipped_type = 0
    errors: list[str] = []
    for f in files:
        name = f.filename or ""
        ext = Path(name).suffix.lower()
        if ext not in MEDIA_IMPORT_EXTS:
            skipped_type += 1
            continue
        data = await f.read()
        if len(data) > cfg.max_media_bytes:
            errors.append(f"{name}：超大小上限")
            continue
        try:
            sha, created = await run_in_threadpool(store.import_media_file, data, ext)
            if created:
                imported += 1
            else:
                skipped_dup += 1
        except Exception as e:  # noqa: BLE001
            errors.append(f"{name}：{e}")
    return {"ok": True, "imported": imported, "skipped_dup": skipped_dup,
            "skipped_type": skipped_type, "errors": errors[:10]}


@app.post("/api/design/import")
async def api_design_import(content: str = Form(""), file: UploadFile | None = File(None)):
    """导入规范（markdown）：粘贴文本 或 上传 .md 文件 → 新 draft 版本。"""
    md = content
    if file is not None and file.filename:
        raw = await file.read()
        md = raw.decode("utf-8", "replace")
    md = md.strip()
    if not md:
        raise HTTPException(422, "内容为空")
    with db.tx() as cur:
        v = (cur.execute("select max(version) v from design_system").fetchone()["v"] or 0) + 1
        cur.execute("insert into design_system(version, json, content_md, status) values(?,?,?, 'draft')",
                    (v, "{}", md))          # json 列写空对象兼容旧库 NOT NULL
    return {"ok": True, "version": v}


@app.get("/api/design/{did}/export")
def api_design_export(did: int):
    d = db.one("select version, json, content_md from design_system where id=?", did)
    if not d:
        raise HTTPException(404)
    if d["content_md"]:   # markdown 承载：导出 .md
        return Response(content=d["content_md"], media_type="text/markdown; charset=utf-8",
                        headers={"Content-Disposition": f'attachment; filename="design-spec-v{d["version"]}.md"'})
    return Response(content=d["json"], media_type="application/json",       # 旧版 PPT 草稿兼容
                    headers={"Content-Disposition": f'attachment; filename="design-system-v{d["version"]}.json"'})


@app.post("/api/design/{did}/update")
def api_design_update(did: int, payload: str = Form(None), content: str = Form(None)):
    """人工编辑草稿：markdown（content，新）或 JSON（payload，旧 PPT 草稿兼容）。仅 draft 可编辑。"""
    with db.tx() as cur:
        if content is not None:
            if not content.strip():
                raise HTTPException(422, "内容为空")
            cur2 = cur.execute("update design_system set content_md=? where id=? and status='draft'",
                               (content, did))
        else:
            try:
                parsed = json.loads(payload)
            except (json.JSONDecodeError, TypeError) as e:
                raise HTTPException(422, f"JSON 语法错误: {e}")
            cur2 = cur.execute("update design_system set json=? where id=? and status='draft'",
                               (json.dumps(parsed, ensure_ascii=False, indent=2), did))
        if cur2.rowcount == 0:
            raise HTTPException(409, "仅草稿可编辑（已确认版本不可改）")
    return {"ok": True}


@app.post("/api/media/{sha}/review")
def api_media_review(sha: str, status: str = Form(...), tags: str = Form(None), note: str = Form(None)):
    if status not in ("pending", "approved", "reference", "forbidden"):
        raise HTTPException(422)
    with db.tx() as cur:
        cur.execute(
            "insert into media_review(sha256, status, tags, note, updated_at)"
            " values(?,?,?,?,datetime('now'))"
            " on conflict(sha256) do update set status=excluded.status, tags=excluded.tags,"
            " note=excluded.note, updated_at=datetime('now')",
            (sha, status, norm_tags(tags), note))
    return {"ok": True}


@app.get("/api/search")
def api_search(q: str, type: str | None = None, tag: str | None = None):
    return JSONResponse(search_slides(db, q, slide_type=type, tag=tag))


@app.get("/api/health")
def api_health():
    queued = db.one("select count(*) c from extraction where status='queued'")["c"]
    running = db.one("select count(*) c from extraction where status='running'")["c"]
    return {"pipeline_version": PIPELINE_VERSION, "queued": queued, "running": running,
            "soffice": cfg.soffice}


# ---------------------------------------------------------------- ppt→pptx
def _convert_ppt_to_pptx(src: Path, staging: Path) -> Path | None:
    """LO 把旧 .ppt 转为 .pptx（锁外）。"""
    if not cfg.soffice:
        return None
    import subprocess, tempfile
    profile = Path(tempfile.gettempdir()) / f"pptlite-lo-{src.stem}-{_uuid4hex()}"
    out = staging / "converted"
    out.mkdir(exist_ok=True)
    cmd = [cfg.soffice, "--headless", "--norestore",
           f"-env:UserInstallation=file:///{profile.as_posix()}",
           "--convert-to", "pptx", "--outdir", str(out), str(src)]
    try:
        # .ppt 转换发生在解析前（不知页数），按文件大小估算超时
        timeout = cfg.lo_timeout(size_bytes=src.stat().st_size)
        p = subprocess.run(cmd, capture_output=True, timeout=timeout)
        files = list(out.glob("*.pptx"))
        if p.returncode == 0 and files:
            dest = staging / "src.pptx"
            shutil.move(str(files[0]), dest)
            return dest
        return None
    except Exception:  # noqa: BLE001
        return None
    finally:
        if profile.exists():
            shutil.rmtree(profile, ignore_errors=True)


def _uuid4hex() -> str:
    import uuid
    return uuid.uuid4().hex


def _sha256_bytes(data: bytes) -> str:
    import hashlib
    return hashlib.sha256(data).hexdigest()


# 内联 SVG favicon（避免 404；无外部文件）
_FAVICON = (
    '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32">'
    '<rect width="32" height="32" rx="6" fill="#1f3a5f"/>'
    '<text x="16" y="22" font-size="16" text-anchor="middle" fill="#fff" '
    'font-family="sans-serif">P</text></svg>'
)


@app.get("/favicon.ico", include_in_schema=False)
def favicon():
    return Response(content=_FAVICON, media_type="image/svg+xml")


def main() -> None:   # pyproject 入口 / PyInstaller 入口
    import uvicorn
    if getattr(sys, "frozen", False):
        # 打包模式：控制台提示 + 自动开浏览器
        import webbrowser, threading as _th
        _th.Timer(1.5, lambda: webbrowser.open("http://127.0.0.1:8765")).start()
        print("=" * 56)
        print("  ppt-lite 已启动：http://127.0.0.1:8765")
        print("  关闭本窗口即停止服务。数据目录：%LOCALAPPDATA%\\ppt-lite\\data")
        print("=" * 56)
    uvicorn.run("ppt_lite.app:app", host="127.0.0.1", port=8765, workers=1)
