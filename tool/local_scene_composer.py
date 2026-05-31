"""Local (non-AI) e-commerce scene images from flat product art — no API key usage."""

from __future__ import annotations

import platform
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable

from PIL import Image, ImageChops, ImageDraw, ImageEnhance, ImageFilter, ImageFont, ImageOps

from mockup_template import compose_ozon_split_scene
from scene_generator import (
    DEFAULT_ASPECT_RATIO,
    list_images,
    list_subfolders,
    resolve_scene_presets,
)

SUPPORTED_IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".webp", ".bmp"}

CANVAS_SIZES: dict[str, tuple[int, int]] = {
    "1:1": (1200, 1200),
    "3:4": (1200, 1600),
    "4:3": (1600, 1200),
    "16:9": (1600, 900),
    "2:3": (1200, 1800),
    "9:16": (1080, 1920),
}

DEFAULT_SIZE_LABEL = "35.83 x 35.83 inches (91 x 91 cm)"

LOCAL_SCENE_PRESETS: list[dict[str, str]] = [
    {"id": "headscarf_side", "name": "头巾侧戴", "layout": "ozon"},
    {"id": "headscarf_back", "name": "头巾背面", "layout": "ozon"},
    {"id": "bow_and_fold", "name": "蝴蝶结+折叠", "layout": "ozon"},
    {"id": "size_chart", "name": "尺寸标注", "layout": "ozon_size"},
    {"id": "flat_full", "name": "整图平铺", "layout": "full"},
]


def parse_canvas_size(aspect_ratio: str) -> tuple[int, int]:
    value = (aspect_ratio or DEFAULT_ASPECT_RATIO).strip().replace("：", ":").replace(" ", "").lower()
    if value in CANVAS_SIZES:
        return CANVAS_SIZES[value]
    if "x" in value:
        parts = value.replace("×", "x").split("x", 1)
        if len(parts) == 2 and parts[0].isdigit() and parts[1].isdigit():
            return int(parts[0]), int(parts[1])
    return CANVAS_SIZES[DEFAULT_ASPECT_RATIO]


def _load_font(size: int) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    if platform.system() == "Darwin":
        candidates = [
            "/System/Library/Fonts/Supplemental/Arial.ttf",
            "/System/Library/Fonts/PingFang.ttc",
            "/System/Library/Fonts/Supplemental/Songti.ttc",
        ]
    else:
        candidates = [
            "C:/Windows/Fonts/segoeui.ttf",
            "C:/Windows/Fonts/arial.ttf",
            "C:/Windows/Fonts/msyh.ttc",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        ]
    for path in candidates:
        try:
            return ImageFont.truetype(path, size=size)
        except OSError:
            continue
    return ImageFont.load_default()


def _ensure_rgba(image: Image.Image) -> Image.Image:
    if image.mode == "RGBA":
        return image
    if image.mode == "P" and "transparency" in image.info:
        return image.convert("RGBA")
    return image.convert("RGBA")


def _white_canvas(size: tuple[int, int]) -> Image.Image:
    return Image.new("RGB", size, (255, 255, 255))


def _drop_shadow(layer: Image.Image, offset: tuple[int, int] = (10, 14), blur: int = 16, opacity: int = 70) -> Image.Image:
    rgba = _ensure_rgba(layer)
    alpha = rgba.split()[3]
    shadow = Image.new("RGBA", rgba.size, (0, 0, 0, 0))
    shadow.putalpha(alpha.point(lambda value: min(255, int(value * opacity / 100))))
    shadow = shadow.filter(ImageFilter.GaussianBlur(blur))
    canvas = Image.new("RGBA", (rgba.width + offset[0] + blur, rgba.height + offset[1] + blur), (0, 0, 0, 0))
    canvas.alpha_composite(shadow, (offset[0], offset[1]))
    canvas.alpha_composite(rgba, (0, 0))
    return canvas


def _fit_product(product: Image.Image, box: tuple[int, int], rotate_deg: float = 0.0) -> Image.Image:
    fitted = ImageOps.contain(_ensure_rgba(product), box, method=Image.Resampling.LANCZOS)
    if rotate_deg:
        fitted = fitted.rotate(rotate_deg, resample=Image.Resampling.BICUBIC, expand=True)
    return _drop_shadow(fitted)


def _paste_center(base: Image.Image, layer: Image.Image, box: tuple[int, int, int, int]) -> None:
    layer = _ensure_rgba(layer)
    x0, y0, x1, y1 = box
    max_w, max_h = x1 - x0, y1 - y0
    contained = ImageOps.contain(layer, (max_w, max_h), method=Image.Resampling.LANCZOS)
    px = x0 + (max_w - contained.width) // 2
    py = y0 + (max_h - contained.height) // 2
    base.paste(contained, (px, py), contained)


def _find_mockup(mockup_root: Path | None, scene_id: str) -> Path | None:
    if not mockup_root or not mockup_root.is_dir():
        return None
    names = (
        f"{scene_id}.jpg",
        f"{scene_id}.png",
        f"{scene_id}.jpeg",
        f"{scene_id}.webp",
    )
    for name in names:
        path = mockup_root / name
        if path.is_file():
            return path
    folder = mockup_root / scene_id
    if folder.is_dir():
        for name in names:
            path = folder / name.replace(f"{scene_id}.", "left.")
            if path.is_file():
                return path
        for name in ("left.jpg", "left.png", "left.jpeg", "mockup.jpg", "mockup.png"):
            path = folder / name
            if path.is_file():
                return path
    for name in ("_default.jpg", "_default.png", "default.jpg"):
        path = mockup_root / name
        if path.is_file():
            return path
    return None


def _render_split_mockup(
    product: Image.Image,
    canvas_size: tuple[int, int],
    mockup_path: Path | None,
    *,
    size_label: str = "",
    layout: str = "split",
) -> Image.Image:
    width, height = canvas_size
    canvas = _white_canvas(canvas_size)
    left_w = int(width * 0.58)
    right_box = (left_w + int(width * 0.03), int(height * 0.08), width - int(width * 0.04), height - int(height * 0.08))

    if mockup_path and mockup_path.is_file():
        mockup = Image.open(mockup_path).convert("RGB")
        mockup = ImageOps.fit(mockup, (left_w, height), method=Image.Resampling.LANCZOS)
        canvas.paste(mockup, (0, 0))

    product_layer = _fit_product(product, (right_box[2] - right_box[0] - 20, right_box[3] - right_box[1] - 20))
    _paste_center(canvas, product_layer, right_box)

    if layout == "size" and size_label:
        draw = ImageDraw.Draw(canvas)
        font = _load_font(max(22, width // 42))
        text_y = int(height * 0.86)
        draw.text((left_w + 24, text_y), size_label, fill=(40, 40, 40), font=font)
    return canvas


def _render_full_flat(product: Image.Image, canvas_size: tuple[int, int], enhance: bool = False) -> Image.Image:
    canvas = _white_canvas(canvas_size)
    margin = int(min(canvas_size) * 0.08)
    box = (margin, margin, canvas_size[0] - margin, canvas_size[1] - margin)
    layer = _fit_product(product, (box[2] - box[0], box[3] - box[1]))
    if enhance:
        rgb = Image.alpha_composite(_white_canvas(layer.size).convert("RGBA"), layer)
        bright = ImageEnhance.Brightness(rgb.convert("RGB")).enhance(1.03)
        layer = _drop_shadow(bright.convert("RGBA"))
    _paste_center(canvas, layer, box)
    return canvas


def _render_folded(product: Image.Image, canvas_size: tuple[int, int]) -> Image.Image:
    canvas = _white_canvas(canvas_size)
    w, h = canvas_size
    size = int(min(w, h) * 0.42)
    flat = ImageOps.contain(product.convert("RGB"), (size, size), method=Image.Resampling.LANCZOS)
    flat = flat.rotate(8, resample=Image.Resampling.BICUBIC, expand=True)
    skew = flat.resize((int(flat.width * 0.92), flat.height), Image.Resampling.BICUBIC)
    layer = _drop_shadow(skew.convert("RGBA"))
    box = (int(w * 0.48), int(h * 0.22), int(w * 0.92), int(h * 0.78))
    _paste_center(canvas, layer, box)
    if w > h * 0.9:
        side = _fit_product(product, (int(w * 0.32), int(h * 0.55)), rotate_deg=-6)
        _paste_center(canvas, side, (int(w * 0.06), int(h * 0.18), int(w * 0.4), int(h * 0.82)))
    return canvas


def _render_swirl(product: Image.Image, canvas_size: tuple[int, int]) -> Image.Image:
    canvas = _white_canvas(canvas_size)
    w, h = canvas_size
    size = int(min(w, h) * 0.62)
    base = ImageOps.contain(product.convert("RGBA"), (size, size), method=Image.Resampling.LANCZOS)
    twisted = base.rotate(18, resample=Image.Resampling.BICUBIC, expand=True)
    mask = Image.new("L", twisted.size, 0)
    draw = ImageDraw.Draw(mask)
    draw.ellipse((8, 8, twisted.width - 8, twisted.height - 8), fill=255)
    twisted.putalpha(ImageChops.multiply(twisted.split()[3], mask))
    layer = _drop_shadow(twisted)
    box = (int(w * 0.18), int(h * 0.12), int(w * 0.88), int(h * 0.9))
    _paste_center(canvas, layer, box)
    return canvas


def _render_dual(product: Image.Image, canvas_size: tuple[int, int]) -> Image.Image:
    canvas = _white_canvas(canvas_size)
    w, h = canvas_size
    left_box = (int(w * 0.05), int(h * 0.15), int(w * 0.46), int(h * 0.85))
    right_box = (int(w * 0.5), int(h * 0.08), int(w * 0.95), int(h * 0.92))
    left = _fit_product(product, (left_box[2] - left_box[0], left_box[3] - left_box[1]), rotate_deg=-12)
    right = _fit_product(product, (right_box[2] - right_box[0], right_box[3] - right_box[1]), rotate_deg=4)
    _paste_center(canvas, left, left_box)
    _paste_center(canvas, right, right_box)
    return canvas


def _render_corner(product: Image.Image, canvas_size: tuple[int, int]) -> Image.Image:
    canvas = _white_canvas(canvas_size)
    w, h = canvas_size
    src = product.convert("RGB")
    cw, ch = src.size
    crop = src.crop((int(cw * 0.55), int(ch * 0.55), cw, ch))
    crop = ImageOps.contain(crop, (int(w * 0.88), int(h * 0.88)), method=Image.Resampling.LANCZOS)
    layer = _drop_shadow(crop.convert("RGBA"))
    box = (int(w * 0.06), int(h * 0.06), int(w * 0.94), int(h * 0.94))
    _paste_center(canvas, layer, box)
    return canvas


def _render_stack(product: Image.Image, canvas_size: tuple[int, int]) -> Image.Image:
    canvas = _white_canvas(canvas_size)
    w, h = canvas_size
    back = _fit_product(product, (int(w * 0.5), int(h * 0.5)), rotate_deg=-10)
    front = _fit_product(product, (int(w * 0.52), int(h * 0.52)), rotate_deg=6)
    _paste_center(canvas, back, (int(w * 0.08), int(h * 0.2), int(w * 0.62), int(h * 0.82)))
    _paste_center(canvas, front, (int(w * 0.32), int(h * 0.12), int(w * 0.94), int(h * 0.88)))
    return canvas


def _render_bow(product: Image.Image, canvas_size: tuple[int, int]) -> Image.Image:
    canvas = _white_canvas(canvas_size)
    w, h = canvas_size
    wing = ImageOps.contain(product.convert("RGBA"), (int(w * 0.22), int(h * 0.22)), method=Image.Resampling.LANCZOS)
    left_wing = wing.transpose(Image.Transpose.FLIP_LEFT_RIGHT)
    center = ImageOps.contain(product.convert("RGBA"), (int(w * 0.12), int(h * 0.18)), method=Image.Resampling.LANCZOS)
    ribbon = _white_canvas((int(w * 0.14), int(h * 0.1)))
    draw = ImageDraw.Draw(ribbon)
    draw.rectangle((0, 0, ribbon.width - 1, ribbon.height - 1), fill=(230, 230, 230))
    bow = Image.new("RGBA", (int(w * 0.52), int(h * 0.3)), (0, 0, 0, 0))
    bow.alpha_composite(_drop_shadow(left_wing), (0, int(bow.height * 0.12)))
    bow.alpha_composite(_drop_shadow(wing), (bow.width - wing.width, int(bow.height * 0.12)))
    bow.alpha_composite(_drop_shadow(center), ((bow.width - center.width) // 2, int(bow.height * 0.38)))
    bow.alpha_composite(ribbon.convert("RGBA"), ((bow.width - ribbon.width) // 2, int(bow.height * 0.55)))
    box = (int(w * 0.06), int(h * 0.12), int(w * 0.52), int(h * 0.55))
    _paste_center(canvas, bow, box)
    flat = _fit_product(product, (int(w * 0.38), int(h * 0.38)))
    _paste_center(canvas, flat, (int(w * 0.52), int(h * 0.18), int(w * 0.94), int(h * 0.82)))
    return canvas


def compose_local_scene(
    product: Image.Image,
    scene: dict[str, str],
    canvas_size: tuple[int, int],
    mockup_root: Path | None,
    size_label: str,
) -> Image.Image:
    layout = scene.get("layout", "full")
    scene_id = scene["id"]
    if layout in {"ozon", "ozon_size"}:
        label = size_label if layout == "ozon_size" else ""
        templated = compose_ozon_split_scene(product, scene_id, canvas_size, mockup_root, size_label=label)
        if templated is not None:
            return templated
    mockup = _find_mockup(mockup_root, scene_id)

    if layout in {"split", "size"}:
        return _render_split_mockup(
            product,
            canvas_size,
            mockup,
            size_label=size_label if layout == "size" else "",
            layout=layout,
        )
    if layout == "fold":
        if mockup:
            return _render_split_mockup(product, canvas_size, mockup)
        return _render_folded(product, canvas_size)
    if layout == "swirl":
        if mockup:
            return _render_split_mockup(product, canvas_size, mockup)
        return _render_swirl(product, canvas_size)
    if layout == "dual":
        return _render_dual(product, canvas_size)
    if layout == "corner":
        return _render_corner(product, canvas_size)
    if layout == "stack":
        return _render_stack(product, canvas_size)
    if layout == "bow":
        return _render_bow(product, canvas_size)
    if layout == "soft":
        return _render_full_flat(product, canvas_size, enhance=True)
    return _render_full_flat(product, canvas_size)


@dataclass
class LocalSceneJobConfig:
    source_root: Path
    output_root: Path
    single_image: Path | None
    aspect_ratio: str
    scene_count: int
    scene_ids: list[str] = field(default_factory=list)
    mockup_root: Path | None = None
    size_label: str = DEFAULT_SIZE_LABEL
    max_folders: int = 0
    sku_filter: str = ""


class LocalSceneWorker:
    def __init__(self, config: LocalSceneJobConfig, logger: Callable[[str], None]):
        self.config = config
        self.logger = logger
        self._cancelled = False

    def cancel(self) -> None:
        self._cancelled = True

    def _process_image(self, source_image: Path, output_folder: Path, sku: str, scenes: list[dict[str, str]]) -> None:
        product = Image.open(source_image)
        canvas_size = parse_canvas_size(self.config.aspect_ratio)
        output_folder.mkdir(parents=True, exist_ok=True)
        for index, scene in enumerate(scenes, start=1):
            if self._cancelled:
                return
            output_path = output_folder / f"{source_image.stem}_local_{index:02d}_{scene['id']}.jpg"
            composed = compose_local_scene(
                product,
                scene,
                canvas_size,
                self.config.mockup_root,
                self.config.size_label,
            )
            composed.save(output_path, format="JPEG", quality=92, optimize=True)
            self.logger(f"    已保存 {output_path.name} ({scene['name']})")

    def run(self) -> None:
        scenes = resolve_scene_presets(
            self.config.scene_ids or None,
            self.config.scene_count,
            preset_list=LOCAL_SCENE_PRESETS,
        )
        canvas_size = parse_canvas_size(self.config.aspect_ratio)
        self.logger(f"本地合成模式（不调用 AI），画布 {canvas_size[0]}x{canvas_size[1]}")
        self.logger(f"将生成 {len(scenes)} 张: {', '.join(s['name'] for s in scenes)}")
        if self.config.mockup_root and self.config.mockup_root.is_dir():
            self.logger(f"自定义模板目录: {self.config.mockup_root}")
        self.logger("使用内置 Ozon 拼图模板（左佩戴区换印花 + 右侧褶皱平铺），不消耗 API。")

        if self.config.single_image and self.config.single_image.is_file():
            sku = self.config.single_image.stem
            self.logger(f"单图: {self.config.single_image.name}")
            self._process_image(self.config.single_image, self.config.output_root / sku, sku, scenes)
            self.logger("本地场景图完成。")
            return

        if not self.config.source_root.is_dir():
            raise RuntimeError("请指定平面原图或有效的批量源目录。")

        folders = list_subfolders(self.config.source_root)
        if self.config.sku_filter.strip():
            folders = [item for item in folders if item.name == self.config.sku_filter.strip()]
        if self.config.max_folders > 0:
            folders = folders[: self.config.max_folders]
        if not folders:
            raise RuntimeError("源目录下没有可处理的子文件夹。")

        total = len(folders)
        for index, sku_folder in enumerate(folders, start=1):
            if self._cancelled:
                self.logger("任务已取消。")
                return
            images = list_images(sku_folder)
            if not images:
                self.logger(f"[{index}/{total}] 跳过 {sku_folder.name}：无图片")
                continue
            self.logger(f"[{index}/{total}] {sku_folder.name}，{len(images)} 张原图")
            for source_image in images:
                if self._cancelled:
                    return
                self.logger(f"  {source_image.name}")
                self._process_image(source_image, self.config.output_root / sku_folder.name, sku_folder.name, scenes)
        self.logger("全部本地场景图完成。")
