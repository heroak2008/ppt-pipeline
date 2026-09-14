"""fslock：进程级文件锁 + trash 搬运（铁律⑤：唯一锁序、uuid 隔离、幂等）。

锁序（§4.0）：外层取 fslock → 短 DB 事务 → commit → 仍在锁内把待删路径搬入 trash/{uuid}/ → 解锁。
- to_trash / to_trash_if_exists：调用方必须已持锁（共享规范路径的搬运）；
  私有 tmp 内部清理不受此限。
- 失败仅记日志（留待启动 GC），绝不抛出影响已提交的 DB 状态。
"""
from __future__ import annotations

import logging
import shutil
import threading
import uuid
from pathlib import Path

log = logging.getLogger("pptlite.fslock")

# 进程内唯一文件锁（单实例假设，§10）
fslock = threading.Lock()


def _new_trash_dir(trash_root: Path) -> Path:
    d = trash_root / uuid.uuid4().hex
    d.mkdir(parents=True, exist_ok=True)
    return d


def to_trash(path: Path, trash_root: Path) -> None:
    """把文件/目录搬入 trash/{uuid}/。目标不存在时记 debug 并返回（幂等）。"""
    path = Path(path)
    if not path.exists():
        return
    try:
        dest = _new_trash_dir(trash_root)
        shutil.move(str(path), str(dest / path.name))
    except Exception:  # noqa: BLE001 — 搬运失败仅记日志（§4.4 步骤 C 约定）
        log.exception("to_trash 失败（留待启动 GC 回收）: %s", path)


def to_trash_if_exists(path: Path | str | None, trash_root: Path) -> None:
    if path is None:
        return
    to_trash(Path(path), trash_root)


def purge_trash(trash_root: Path) -> None:
    """物理删除 trash 下全部内容（仅启动停流窗口 GC 步骤 e 调用）。"""
    if not trash_root.exists():
        return
    for child in trash_root.iterdir():
        try:
            if child.is_dir():
                shutil.rmtree(child)
            else:
                child.unlink()
        except Exception:  # noqa: BLE001
            log.exception("purge_trash 失败: %s", child)
