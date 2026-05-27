场景图模板说明（本地合成，不消耗 API Key）
========================================

## 推荐用法

1. 准备一张「平面印花原图」（正方形图案）
2. 软件「场景图」页 → 点 **本地合成（不耗 API）**
3. 默认生成 4 种与你参考图同构图的场景：
   - headscarf_side   头巾侧戴
   - headscarf_back   头巾背面
   - bow_and_fold     蝴蝶结 + 折叠
   - size_chart       尺寸标注

## 原理

工具内置从「参考场景图」提取的模板（在 OzonTool 目录 builtin_mockups 内）：
- base.jpg      模特与背景（右侧产品区已留白）
- mask_worn.png 左侧头巾区域（只在这里替换为你的印花）
- mask_flat.png 右侧平铺区域
- template.json 右侧透视位置

因此：**同一姿势可反复换不同印花**，无需每张图都调 AI。

## 自定义模板（可选）

若内置效果不满意，可用你自己的参考图生成模板：

  python D:\ozon\tool\build_builtin_mockups.py

或按目录放置（与内置结构相同）：

  mockups\headscarf_side\
    base.jpg
    mask_worn.png
    mask_flat.png
    template.json

在软件里把「模特/场景底图目录」指到 mockups 文件夹。

## 与 AI 的区别

| 方式 | 费用 | 效果 |
|------|------|------|
| 本地模板 | 免费 | 姿势固定，印花准确，接近参考拼图 |
| AI 生图 | 耗 Key | 姿势多变，但贵且印花易变形 |

建议：日常上新用本地模板；只有新姿势才用 1 次 AI 拍底图后做成模板。
