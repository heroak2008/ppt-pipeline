"""预检：魔数三级分流 + 宏/加密检测（§4.1；编码注意事项 #1/#2：魔数优先于扩展名、
olefile 按流路径规范化匹配）。"""
from __future__ import annotations

import io
import zipfile
from dataclasses import dataclass
from pathlib import Path

import olefile

ZIP_MAGIC = b"PK\x03\x04"
CFB_MAGIC = b"\xd0\xcf\x11\xe0"


class PrecheckError(Exception):
    """结构化拒收原因（HTTP 4xx）。"""

    def __init__(self, reason: str):
        super().__init__(reason)
        self.reason = reason


@dataclass
class PrecheckResult:
    kind: str                 # 'pptx' | 'legacy-ppt-converted' | 'plain-pptx'
    derived_bytes: bytes | None = None   # legacy .ppt 转换出的 .pptx 内容（由调用方落盘）
    note: str = ""


def _has_vba_zip(zf: zipfile.ZipFile) -> bool:
    return any(n == "ppt/vbaProject.bin" for n in zf.namelist())


def _cfb_stream_exists(data: bytes, streamnames: tuple[str, ...]) -> bool:
    """按流路径规范化匹配（不只查顶层展示名）。"""
    ole = olefile.OleFileIO(io.BytesIO(data))
    try:
        # olefile 列出的是 '/' 分隔的路径；目标流可能位于任意 storage 下
        targets = {s.lower() for s in streamnames}
        for entry in ole.listdir(streams=True, storages=False):
            path = "/".join(entry).lower()
            # 匹配任意层级的流名（如 'vba_project'、'encryptioninfo' 兼容下划线变体）
            leaf = entry[-1].lower().replace("_vba_project_cur", "vba_project")
            if path in targets or leaf in targets or leaf.replace("_", "") in {t.replace("_", "") for t in targets}:
                return True
        return False
    finally:
        ole.close()


def precheck(data: bytes, soffice_available: bool) -> PrecheckResult:
    """输入：文件全部字节（暂存）。输出：通过/拒收 + 派生信息。

    三级分流：ZIP → .pptx 路线；OLE CFB → .ppt 路线；其他 → 无法识别。
    """
    if len(data) == 0:
        raise PrecheckError("空文件")

    magic = data[:8]

    # ── ZIP 路线（.pptx/.pptm）──
    if magic.startswith(ZIP_MAGIC):
        try:
            zf = zipfile.ZipFile(io.BytesIO(data))
        except zipfile.BadZipFile:
            raise PrecheckError("文件损坏（ZIP 无法打开）")
        with zf:
            if _has_vba_zip(zf):
                raise PrecheckError("含宏，人工处理")
            return PrecheckResult(kind="pptx")

    # ── OLE CFB 路线（旧 .ppt / 加密的新格式）──
    if magic.startswith(CFB_MAGIC):
        if _cfb_stream_exists(data, ("encryptioninfo",)):
            raise PrecheckError("加密，人工处理")
        if _cfb_stream_exists(data, ("vba_project", "_vba_project_cur")):
            raise PrecheckError("含宏，人工处理")
        # 旧 .ppt：需要 LO 转换（转换本身由调用方在锁外执行；这里只声明类型）
        if not soffice_available:
            raise PrecheckError("旧版 .ppt 需要 LibreOffice 转换，但本机未安装/未配置（可设置 "
                                "PPT_LITE_SOFFICE 环境变量），请手工另存为 .pptx 后上传")
        return PrecheckResult(kind="legacy-ppt")

    raise PrecheckError("无法识别的格式")
