"""配置：路径解析、LibreOffice 发现（可配置/可缺省降级）。"""
from __future__ import annotations

import os
import shutil
import sys
from dataclasses import dataclass, field
from pathlib import Path


def _find_soffice() -> str | None:
    """查找顺序：PPT_LITE_SOFFICE 环境变量 → 安装器写入的配置文件 → PATH → 常见路径。"""
    env = os.environ.get("PPT_LITE_SOFFICE")
    if env:
        return env if Path(env).exists() else None
    # 安装器选择页写入的配置文件（%LOCALAPPDATA%\ppt-lite\soffice.txt）
    cfg_file = Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / "ppt-lite" / "soffice.txt"
    try:
        if cfg_file.exists():
            p = cfg_file.read_text(encoding="utf-8").strip()
            if p and Path(p).exists():
                return p
    except Exception:  # noqa: BLE001
        pass
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

    # LO：超时按页数+文件规模动态计算（§4.x）
    soffice: str | None = field(default_factory=_find_soffice)
    lo_timeout_base_sec: int = 60          # 基础开销（LO 冷启动 + 字体缓存）
    lo_timeout_per_page_sec: int = 8       # 每页渲染开销
    lo_timeout_per_mb_sec: int = 10        # 每 MB 文件大小开销
    lo_timeout_max_sec: int = 1800         # 上限防失控

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

    def lo_timeout(self, *, pages: int = 0, size_bytes: int = 0) -> int:
        """LO 转换超时：base + 页数×每页开销 + 文件大小×每 MB 开销，封顶。"""
        size_mb = size_bytes / (1024 * 1024)
        t = self.lo_timeout_base_sec + pages * self.lo_timeout_per_page_sec \
            + int(size_mb * self.lo_timeout_per_mb_sec)
        return min(t, self.lo_timeout_max_sec)
