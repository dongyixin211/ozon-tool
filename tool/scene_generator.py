"""Lifestyle / usage-scene product image generation from flat pattern images."""

from __future__ import annotations

import concurrent.futures
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Iterable

SUPPORTED_IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".webp", ".bmp"}

DEFAULT_ASPECT_RATIO = "1:1"

ASPECT_RATIO_SIZES: dict[str, str] = {
    "1:1": "1024x1024",
    "3:4": "1024x1536",
    "4:3": "1536x1024",
    "16:9": "1536x864",
    "2:3": "1024x1536",
    "9:16": "1024x1536",
}

DEFAULT_SCENE_PROMPT_TEMPLATE = """你是一名 Ozon / Wildberries 跨境电商场景图设计师。参考图为商品平面印花原图（方巾/头巾/丝巾），请生成一张「使用场景 + 产品平铺」电商附图。

【参考图 — 必须严格遵守】
1. 印花与参考图完全一致：底色、主图案、边框纹样、配色，禁止改色、重绘或新增图案
2. 面料为丝绸/缎面方巾，自然褶皱与柔光，质感真实
3. 输出比例 {aspect_ratio}（{output_size}），白底 (#FFFFFF) 影棚光，画面干净专业

【构图要求】
4. 推荐：左侧 55%–65% 真人佩戴/使用场景，右侧 35%–45% 同款方巾平铺或折叠展示
5. 人物为欧美或中性年轻女性，服装简洁（黑/白），不抢产品焦点
6. 不要中文；尺寸标注场景除外，不要大段营销文字

【本次场景】
{scene_description}

货号：{sku} | 原图：{image_name} | 场景：{scene_name}（{scene_id}）
"""

SCENE_PRESETS: list[dict[str, str]] = [
    {
        "id": "headscarf_side",
        "name": "头巾侧戴",
        "description": (
            "年轻女性侧面半身，将方巾在头顶系成头巾/头带样式，结打在颈后，巾角自然垂落肩前。"
            "右侧展示同款方巾平铺略带褶皱。明亮影棚白底。"
        ),
    },
    {
        "id": "headscarf_back",
        "name": "头巾背面",
        "description": (
            "女性背影，长发，方巾包头并在颈后打结，巾身自然垂落背部。"
            "右侧方巾平铺或中心微拧展示面料光泽与完整印花。"
        ),
    },
    {
        "id": "bow_and_fold",
        "name": "蝴蝶结展示",
        "description": (
            "左侧将方巾系成装饰蝴蝶结展示垂坠与印花。"
            "右侧方巾折叠成整齐方块，完整呈现边框纹样。"
        ),
    },
    {
        "id": "size_chart",
        "name": "尺寸标注",
        "description": (
            "左侧蝴蝶结展示，右侧折叠方块平铺。"
            "在画面合适位置用英文标注尺寸：35.83 x 35.83 inches (91 x 91 cm)，字体简洁专业。"
        ),
    },
    {
        "id": "neck_scarf",
        "name": "颈间方巾",
        "description": (
            "女性正面或微侧面，方巾作为颈巾系在白色衬衫领口，优雅日常。"
            "右侧平铺展示同款方巾。"
        ),
    },
    {
        "id": "handbag",
        "name": "包饰系带",
        "description": (
            "方巾系在浅棕色皮质手提包提手上作为装饰。"
            "右侧平铺方巾。可只拍手袋局部，白底影棚光。"
        ),
    },
    {
        "id": "ponytail",
        "name": "马尾发饰",
        "description": (
            "户外柔和阳光下，女性高马尾，方巾束在发间作为发带。"
            "右侧仍保留平铺方巾展示（白底或极浅虚化背景）。"
        ),
    },
    {
        "id": "picnic",
        "name": "野餐摆拍",
        "description": (
            "方巾铺在野餐垫上，旁边有野餐篮与草莓，夏日氛围但主体仍是方巾印花清晰。"
            "构图可偏俯拍，仍需能看清印花与边框。"
        ),
    },
    {
        "id": "table_styling",
        "name": "桌面静物",
        "description": (
            "方巾搭在木质桌面边缘，旁有茶杯与新鲜草莓，生活美学静物。"
            "印花与边框清晰可见。"
        ),
    },
    {
        "id": "sun_hat",
        "name": "草帽装饰",
        "description": (
            "宽檐草帽上系方巾作为度假配饰，可搭配沙滩感但仍以白底或浅色简洁背景为主。"
            "旁侧或一角展示平铺方巾。"
        ),
    },
]


def aspect_ratio_to_size(aspect_ratio: str) -> str:
    value = (aspect_ratio or DEFAULT_ASPECT_RATIO).strip()
    if not value:
        value = DEFAULT_ASPECT_RATIO
    normalized = value.replace("：", ":").replace(" ", "").lower()
    if normalized in ASPECT_RATIO_SIZES:
        return ASPECT_RATIO_SIZES[normalized]
    match = re.fullmatch(r"(\d{3,4})[x×](\d{3,4})", normalized)
    if match:
        return f"{match.group(1)}x{match.group(2)}"
    raise ValueError(f"不支持的比例: {aspect_ratio}，请使用 1:1、3:4、4:3、16:9 或 1024x1536 格式。")


def list_images(folder: Path) -> list[Path]:
    return sorted(
        item for item in folder.iterdir() if item.is_file() and item.suffix.lower() in SUPPORTED_IMAGE_EXTENSIONS
    )


def list_subfolders(root: Path) -> list[Path]:
    return sorted(item for item in root.iterdir() if item.is_dir())


def resolve_scene_presets(
    selected_ids: Iterable[str] | None,
    max_scenes: int,
    preset_list: list[dict[str, str]] | None = None,
) -> list[dict[str, str]]:
    catalog = preset_list or SCENE_PRESETS
    if selected_ids:
        id_set = {item.strip() for item in selected_ids if item.strip()}
        presets = [scene for scene in catalog if scene["id"] in id_set]
        if not presets:
            raise ValueError("未匹配到任何场景 ID，请检查 scene_ids 配置。")
        return presets[:max_scenes] if max_scenes > 0 else presets
    count = max(1, min(max_scenes, len(catalog)))
    return catalog[:count]


@dataclass
class SceneJobConfig:
    source_root: Path
    output_root: Path
    single_image: Path | None
    aspect_ratio: str
    scene_count: int
    scene_ids: list[str] = field(default_factory=list)
    scene_prompt_template: str = DEFAULT_SCENE_PROMPT_TEMPLATE
    quality: str = "high"
    max_workers: int = 2
    max_folders: int = 0
    sku_filter: str = ""


class SceneGenerationWorker:
    def __init__(
        self,
        config: SceneJobConfig,
        client_factory: Callable[[], object],
        logger: Callable[[str], None],
    ):
        self.config = config
        self.client_factory = client_factory
        self.logger = logger
        self._cancelled = False

    def cancel(self) -> None:
        self._cancelled = True

    def _build_prompt(self, scene: dict[str, str], sku: str, image_name: str, output_size: str) -> str:
        return self.config.scene_prompt_template.format(
            aspect_ratio=self.config.aspect_ratio,
            output_size=output_size,
            scene_description=scene["description"],
            scene_name=scene["name"],
            scene_id=scene["id"],
            sku=sku,
            image_name=image_name,
        )

    def _generate_one(
        self,
        client: object,
        source_image: Path,
        output_path: Path,
        scene: dict[str, str],
        sku: str,
        output_size: str,
    ) -> str:
        prompt = self._build_prompt(scene, sku, source_image.name, output_size)
        generated_bytes, mode_note = client.generate_image_with_fallback(
            prompt=prompt,
            output_size=output_size,
            quality=self.config.quality,
            reference_images=[source_image],
        )
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_bytes(generated_bytes)
        return mode_note

    def _process_image(self, source_image: Path, output_folder: Path, sku: str, scenes: list[dict[str, str]], output_size: str) -> None:
        worker_count = min(max(1, self.config.max_workers), len(scenes))
        self.logger(f"  原图 {source_image.name}，将生成 {len(scenes)} 张场景图，并发 {worker_count}")

        def task(scene: dict[str, str]) -> tuple[dict[str, str], Path, str]:
            index = scenes.index(scene) + 1
            output_path = output_folder / f"{source_image.stem}_scene_{index:02d}_{scene['id']}.png"
            thread_client = self.client_factory()
            mode_note = self._generate_one(thread_client, source_image, output_path, scene, sku, output_size)
            return scene, output_path, mode_note

        if worker_count <= 1:
            for scene in scenes:
                if self._cancelled:
                    return
                _, output_path, mode_note = task(scene)
                self.logger(f"    已保存 {output_path.name}" + (f" ({mode_note})" if mode_note != "edit" else ""))
            return

        with concurrent.futures.ThreadPoolExecutor(max_workers=worker_count) as executor:
            futures = {executor.submit(task, scene): scene for scene in scenes}
            for future in concurrent.futures.as_completed(futures):
                if self._cancelled:
                    executor.shutdown(wait=False, cancel_futures=True)
                    return
                scene = futures[future]
                try:
                    _, output_path, mode_note = future.result()
                    self.logger(f"    已保存 {output_path.name}" + (f" ({mode_note})" if mode_note != "edit" else ""))
                except Exception as exc:  # noqa: BLE001
                    self.logger(f"    场景 {scene['name']} 失败: {exc}")

    def run(self) -> None:
        output_size = aspect_ratio_to_size(self.config.aspect_ratio)
        scenes = resolve_scene_presets(self.config.scene_ids or None, self.config.scene_count)
        self.logger(f"图片比例 {self.config.aspect_ratio} -> API 尺寸 {output_size}")
        self.logger(f"已选 {len(scenes)} 个场景: {', '.join(s['name'] for s in scenes)}")

        if self.config.single_image and self.config.single_image.is_file():
            sku = self.config.single_image.stem
            output_folder = self.config.output_root / sku
            self.logger(f"单图模式: {self.config.single_image.name}")
            self._process_image(self.config.single_image, output_folder, sku, scenes, output_size)
            self.logger("场景图生成完成。")
            return

        if not self.config.source_root.is_dir():
            raise RuntimeError("请指定有效的源目录，或选择一张平面原图。")

        folders = list_subfolders(self.config.source_root)
        sku_filter = self.config.sku_filter.strip()
        if sku_filter:
            folders = [item for item in folders if item.name == sku_filter]
        if self.config.max_folders > 0:
            folders = folders[: self.config.max_folders]
        if not folders:
            raise RuntimeError("源目录下没有找到可处理的子文件夹。")

        total = len(folders)
        for index, sku_folder in enumerate(folders, start=1):
            if self._cancelled:
                self.logger("任务已取消。")
                return
            sku = sku_folder.name
            images = list_images(sku_folder)
            if not images:
                self.logger(f"[{index}/{total}] 跳过 {sku}：文件夹内无图片")
                continue
            self.logger(f"[{index}/{total}] 处理货号 {sku}，原图 {len(images)} 张")
            output_folder = self.config.output_root / sku
            for image_index, source_image in enumerate(images, start=1):
                if self._cancelled:
                    self.logger("任务已取消。")
                    return
                self.logger(f"  [{image_index}/{len(images)}] {source_image.name}")
                self._process_image(source_image, output_folder, sku, scenes, output_size)

        self.logger("全部场景图生成完成。")
