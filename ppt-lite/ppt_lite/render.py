"""render.py：LO 调用（每文件独立 profile / 超时 / kill 进程树；锁外，§10）
+ pypdfium2 每页 PNG/缩略图。"""
from __future__ import annotations

import logging
import subprocess
import tempfile
import uuid
from pathlib import Path

from PIL import Image
import pypdfium2 as pdfium

log = logging.getLogger("pptlite.render")


def convert_to_pdf(pptx_path: Path, out_dir: Path, soffice: str, timeout_sec: int) -> Path | None:
    """LO 整文件转 PDF（锁外）。返回 PDF 路径；超时/失败返回 None（记 render_error 降级）。"""
    out_dir.mkdir(parents=True, exist_ok=True)
    profile = Path(tempfile.gettempdir()) / f"pptlite-lo-{uuid.uuid4().hex}"
    cmd = [
        soffice, "--headless", "--norestore", "--nolockcheck", "--nologo", "--nodefault", "--nofirststartwizard",
        f"-env:UserInstallation=file:///{profile.as_posix()}",
        "--convert-to", "pdf", "--outdir", str(out_dir), str(pptx_path),
    ]
    creationflags = subprocess.CREATE_NO_WINDOW if hasattr(subprocess, "CREATE_NO_WINDOW") else 0
    try:
        proc = subprocess.Popen(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                                stdin=subprocess.DEVNULL, creationflags=creationflags)
        try:
            proc.wait(timeout=timeout_sec)
        except subprocess.TimeoutExpired:
            _kill_tree(proc)
            log.warning("LO 转换超时（%ss）: %s", timeout_sec, pptx_path.name)
            return None
        pdf = out_dir / (pptx_path.stem + ".pdf")
        if proc.returncode != 0 or not pdf.exists():
            log.warning("LO 转换失败 rc=%s: %s", proc.returncode, pptx_path.name)
            return None
        return pdf
    except Exception:  # noqa: BLE001
        log.exception("LO 调用异常: %s", pptx_path.name)
        return None
    finally:
        try:
            if profile.exists():
                import shutil as _sh
                _sh.rmtree(profile, ignore_errors=True)
        except Exception:  # noqa: BLE001
            pass


def _kill_tree(proc: subprocess.Popen) -> None:
    try:
        subprocess.run(["taskkill", "/F", "/T", "/PID", str(proc.pid)], capture_output=True)
    except Exception:  # noqa: BLE001
        proc.kill()


def render_pdf_pages(pdf_path: Path, slides_dir: Path, thumbs_dir: Path,
                     slide_width: int = 1600, thumb_width: int = 480) -> list[tuple[str, str]]:
    """每页 PNG（≥slide_width 宽）+ 缩略图。返回 [(slide_rel, thumb_rel), ...] 按页序。

    相对名：slides/{no}.png、thumbs/{no}.png（§4.3 工作目录布局）。
    """
    doc = pdfium.PdfDocument(pdf_path)
    slides_dir.mkdir(parents=True, exist_ok=True)
    thumbs_dir.mkdir(parents=True, exist_ok=True)
    out: list[tuple[str, str]] = []
    try:
        for i in range(len(doc)):
            page = doc[i]
            scale = slide_width / page.get_width()
            bitmap = page.render(scale=scale)
            img = bitmap.to_pil()
            s_rel = f"slides/{i + 1}.png"
            t_rel = f"thumbs/{i + 1}.png"
            img.save(slides_dir / f"{i + 1}.png")
            t = img.copy()
            t.thumbnail((thumb_width, 10 ** 9))
            t.save(thumbs_dir / f"{i + 1}.png")
            out.append((s_rel, t_rel))
    finally:
        doc.close()
    return out
