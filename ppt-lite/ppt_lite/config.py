"""配置：路径解析、LibreOffice 发现（可配置/可缺省降级）。"""
from __future__ import annotations

import os
import shutil
import sys
from dataclasses import dataclass, field
from pathlib import Path


def _find_soffice() -> str | None:
    env = os.environ.get("PPT_LITE_SOFFICE")
    if env:
        return env if Path(env).exists() else None
    p = shutil.which("soffice") or shutil.which("soffice.exe")
    if p:
        return p
    candidates = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        os.path.expandvars(r"%LOCALAPPDATA%\Programs\LibreOffice\program\soffice.exe"),
        r"D:\Program Files\LibreOffice\program\soffice.exe",
        "/usr/bin/soffice",
        "/usr/local/bin/soffice",
        "/opt/libreoffice/program/soffice",
    ]
    for c in candidates:
        if Path(c).exists():
            return c
    return None


def _default_data_dir() -> Path:
    env = os.environ.get("PPT_LITE_DATA")
    if env:
        return Path(env)
    if getattr(sys, "frozen", False):           # PyInstaller 打包的 exe
        return Path(os.environ.get("LOCALAPPDATA", Path.home())) / "ppt-lite" / "data"
    return Path("data")


@dataclass
class Config:
    # 冻结模式（PyInstaller exe）：默认 data 目录放 %LOCALAPPDATA%\ppt-lite\data（用户可写）；
    # 开发模式：项目内 ./data
    data_dir: Path = field(default_factory=lambda: _default_data_dir())

    # 上限（§10）
    max_upload_bytes: int = 500 * 1024 * 1024
    max_unzip_total_bytes: int = 2 * 1024 * 1024 * 1024
    max_media_bytes: int = 512 * 1024 * 1024

    # LO
    soffice: str | None = field(default_factory=_find_soffice)
    lo_timeout_sec: int = 180

    # 渲染
    slide_png_width: int = 1600
    thumb_width: int = 480

    # DB
    busy_timeout_ms: int = 5000
    busy_timeout_short_ms: int = 500     # fslock 内短事务用（§10）
    busy_retry_max: int = 3

    # 领取轮询
    poll_interval_sec: float = 2.0

    # extraction 保留的 previews（旧版本目录延迟清理由 GC 处理）
    def __post_init__(self) -> None:
        self.data_dir = Path(self.data_dir).resolve()

    # ---- 派生目录（§2） ----
    @property
    def db_path(self) -> Path:
        return self.data_dir / "app.db"

    @property
    def raw_dir(self) -> Path:
        return self.data_dir / "raw"

    @property
    def derived_dir(self) -> Path:
        return self.raw_dir / "derived"

    @property
    def media_dir(self) -> Path:
        return self.data_dir / "media"

    @property
    def tmp_dir(self) -> Path:
        return self.data_dir / "tmp"

    @property
    def previews_dir(self) -> Path:
        return self.data_dir / "previews"

    @property
    def staging_dir(self) -> Path:
        return self.previews_dir / ".staging"

    @property
    def trash_dir(self) -> Path:
        return self.data_dir / "trash"

    def ensure_dirs(self) -> None:
        for d in (self.raw_dir, self.derived_dir, self.media_dir, self.tmp_dir,
                  self.previews_dir, self.staging_dir, self.trash_dir):
            d.mkdir(parents=True, exist_ok=True)
