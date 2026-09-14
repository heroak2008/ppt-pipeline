"""PyInstaller 入口：冻结模式下 import 链与 cwd 无关。"""
import sys
import os

# 保证 ppt_lite 包可被导入（开发与打包两种情况都覆盖）
_here = os.path.dirname(os.path.abspath(__file__))
if _here not in sys.path:
    sys.path.insert(0, _here)
sys.path.insert(0, os.path.dirname(_here))

from ppt_lite.app import main

if __name__ == "__main__":
    main()
