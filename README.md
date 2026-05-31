# Ozon 商品素材与上架工具

本仓库包含 Ozon 电商工具的两部分：

| 目录 | 说明 |
|------|------|
| `tool/` | 主程序源码（GUI、批量生图、Ozon 上架等） |
| `ozon-plsj/` | 发布/打包目录（场景图 CLI、一键打包脚本） |

## 环境

- Python 3.9+
- macOS 运行 GUI：双击 `启动Ozon工具.command`
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

## GitHub

远程仓库：[dongyixin211/ozon-tool](https://github.com/dongyixin211/ozon-tool)

首次推送（需先登录 GitHub）：

```powershell
cd d:\ozon
gh auth login
.\push-to-github.ps1
```
