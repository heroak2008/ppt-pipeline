# ppt-lite 启动指南

## 方式一：一键启动（推荐）

双击 `start.bat`，或命令行：

```powershell
cd E:\Code\ppt-pipeline\ppt-lite
.\start.bat            # 默认 8765
.\start.bat 9000       # 自定义端口
```

浏览器打开 http://127.0.0.1:8765

## 方式二：手动命令

```powershell
cd E:\Code\ppt-pipeline\ppt-lite      # ★ 必须在项目目录内（模块查找依赖 cwd）
uv sync                                # 首次/依赖变更时；已配置清华镜像
uv run uvicorn ppt_lite.app:app --port 8765 --workers 1
```

或使用入口命令（已作为包安装）：

```powershell
uv run ppt-lite                        # 固定 127.0.0.1:8765
```

## 常见问题

**`No module named ppt_lite`**
- 不在 `ppt-lite` 目录里执行了 uvicorn → `cd` 进项目目录
- 用了系统 Python 而不是 `uv run` → 必须 `uv run` 前缀（虚拟环境在项目 `.venv`）

**依赖安装超时**
- 已在 `uv.toml` 配置清华镜像；如仍失败可尝试 `uv sync --default-index https://mirrors.aliyun.com/pypi/simple/`

**页面预览全是"无预览"、`.ppt` 上传被拒**
- 本机未安装 LibreOffice。装好后重启服务即可（自动检测 `soffice`）；
  或设置环境变量 `PPT_LITE_SOFFICE` 指向 soffice.exe 完整路径
- 数据、检索、审核不依赖 LO，可正常使用

**数据存哪里**
- 默认 `.\data\`（与启动时工作目录相对）；可用环境变量 `PPT_LITE_DATA` 指定其他位置
