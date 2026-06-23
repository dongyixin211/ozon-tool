# Ozon 商品素材与上架工具

本仓库包含 Ozon 电商工具的源码、打包脚本和发布辅助目录：

| 目录 | 说明 |
|------|------|
| `tool/app.py` | GUI 主入口（素材生成、场景图、提示词、批量上架、商品运维） |
| `tool/batch_upload/` | Ozon Seller API、OSS 上传、库存/条码/视频/商品更新逻辑 |
| `tool/builtin_mockups/` | 本地场景图合成使用的内置模板素材 |
| `tool/exe_build/` | Windows PyInstaller 打包配置与脚本 |
| `ozon-plsj/` | 发布/打包目录（场景图 CLI、一键打包脚本） |
| `requirements.txt` | 运行 GUI 和 CLI 所需的 Python 依赖 |

## 环境

- Python 3.9+；macOS GUI 推荐 Python 3.11+ / Tk 8.6+，系统自带 Python 3.9 的 Tk 8.5 可能白屏
- macOS 运行 GUI：双击 `启动Ozon工具.command`，脚本会优先选择可正常显示 GUI 的 Python
- 命令行运行 GUI：`python3 tool/app.py`
- 场景图 CLI：`python ozon-plsj/generate_scenes.py -i 原图.png -o 输出目录`

## 配置

复制示例配置后填写本地路径与 API Key：

```text
cp config.example.json tool/config.json
cp config.example.json ozon-plsj/config.json
```

`config.json` 已加入 `.gitignore`，不会进入版本库。

## 打包

Windows 可在 `ozon-plsj` 目录执行 `一键打包.bat`，生成 `OzonTool_*.exe`。macOS 当前建议直接用源码方式运行。

## 项目整理原则

- 主源码只保留在 `tool/`，打包脚本直接引用当前源码，避免维护旧源码副本。
- 上架 Excel 模板由程序运行时生成，不在仓库里保存二进制模板文件。
- 本地配置、测试输出、打包产物和系统临时文件由 `.gitignore` 排除。
- 项目优化审计与后续路线见 `docs/enterprise_review.md`。

## GitHub

远程仓库：[dongyixin211/ozon-tool](https://github.com/dongyixin211/ozon-tool)

首次推送（需先登录 GitHub）：

```powershell
cd d:\ozon
gh auth login
.\push-to-github.ps1
```
