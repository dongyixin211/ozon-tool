"""Ozon-style split mockups: replace fabric on worn + flat regions from flat pattern art."""

from __future__ import annotations

import json
import sys
from dataclasses import dataclass
from pathlib import Path

from PIL import Image, ImageChops, ImageDraw, ImageEnhance, ImageFilter, ImageFont, ImageOps


def _builtin_mockup_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys._MEIPASS) / "builtin_mockups"  # type: ignore[attr-defined]
    return Path(__file__).resolve().parent / "builtin_mockups"


@dataclass
class MockupTemplate:
    scene_id: str
    canvas_size: tuple[int, int]
    left_box: tuple[int, int, int, int]
    right_quad: tuple[tuple[int, int], tuple[int, int], tuple[int, int], tuple[int, int]]
    size_label_y_ratio: float = 0.86

    @classmethod
    def load(cls, folder: Path) -> MockupTemplate:
        meta_path = folder / "template.json"
        if not meta_path.is_file():
            raise FileNotFoundError(f"缺少 template.json: {folder}")
        data = json.loads(meta_path.read_text(encoding="utf-8"))
        canvas = tuple(data["canvas"])
        left = tuple(data["left_box"])
        quad = tuple(tuple(point) for point in data["right_quad"])
        return cls(
            scene_id=str(data.get("scene_id", folder.name)),
            canvas_size=(int(canvas[0]), int(canvas[1])),
            left_box=(int(left[0]), int(left[1]), int(left[2]), int(left[3])),
            right_quad=quad,  # type: ignore[arg-type]
            size_label_y_ratio=float(data.get("size_label_y_ratio", 0.86)),
        )


def resolve_template_dir(mockup_root: Path | None, scene_id: str) -> Path | None:
    candidates: list[Path] = []
    if mockup_root and mockup_root.is_dir():
        candidates.append(mockup_root / scene_id)
    candidates.append(_builtin_mockup_root() / scene_id)
    for folder in candidates:
        if (folder / "template.json").is_file() and (folder / "mask_worn.png").is_file():
            return folder
    return None


def _load_font(size: int) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    for path in (
        "C:/Windows/Fonts/segoeui.ttf",
        "C:/Windows/Fonts/arial.ttf",
        "C:/Windows/Fonts/msyh.ttc",
    ):
        try:
            return ImageFont.truetype(path, size=size)
        except OSError:
            continue
    return ImageFont.load_default()


def _find_perspective_coeffs(
    src: tuple[tuple[float, float], tuple[float, float], tuple[float, float], tuple[float, float]],
    dst: tuple[tuple[float, float], tuple[float, float], tuple[float, float], tuple[float, float]],
) -> tuple[float, ...]:
    """Map src quad (in source image) to axis-aligned dst rectangle."""

    def equations() -> list[list[float]]:
        matrix: list[list[float]] = []
        for (x, y), (u, v) in zip(src, dst):
            matrix.append([x, y, 1, 0, 0, 0, -u * x, -u * y, u])
            matrix.append([0, 0, 0, x, y, 1, -v * x, -v * y, v])
        return matrix

    def gauss_solve(matrix: list[list[float]]) -> list[float]:
        n = len(matrix)
        for col in range(n):
            pivot = col
            for row in range(col + 1, n):
                if abs(matrix[row][col]) > abs(matrix[pivot][col]):
                    pivot = row
            matrix[col], matrix[pivot] = matrix[pivot], matrix[col]
            pivot_val = matrix[col][col]
            if abs(pivot_val) < 1e-12:
                raise ValueError("透视变换求解失败")
            for row in range(col, n):
                factor = matrix[row][col] / pivot_val
                if row == col:
                    continue
                for j in range(col, n + 1):
                    matrix[row][j] -= factor * matrix[col][j]
        return [matrix[i][n] / matrix[i][i] for i in range(n)]

    size = (
        max(p[0] for p in dst) - min(p[0] for p in dst),
        max(p[1] for p in dst) - min(p[1] for p in dst),
    )
    rect_dst = ((0, 0), (size[0], 0), (size[0], size[1]), (0, size[1]))
    return tuple(gauss_solve(equations()))  # type: ignore[arg-type]


def _warp_pattern_to_quad(pattern: Image.Image, quad: tuple[tuple[int, int], ...], out_size: tuple[int, int]) -> Image.Image:
    xs = [p[0] for p in quad]
    ys = [p[1] for p in quad]
    min_x, max_x = min(xs), max(xs)
    min_y, max_y = min(ys), max(ys)
    bbox_w = max(32, max_x - min_x)
    bbox_h = max(32, max_y - min_y)
    fitted = ImageOps.fit(pattern.convert("RGB"), (bbox_w, bbox_h), method=Image.Resampling.LANCZOS)
    fitted = fitted.rotate(-4, resample=Image.Resampling.BICUBIC, expand=True)
    layer = Image.new("RGB", out_size, (255, 255, 255))
    px = min_x + (bbox_w - fitted.width) // 2
    py = min_y + (bbox_h - fitted.height) // 2
    layer.paste(fitted, (px, py))
    return layer


def _apply_fabric_with_shading(base: Image.Image, pattern: Image.Image, mask: Image.Image) -> Image.Image:
    base_rgb = base.convert("RGB")
    pattern = pattern.convert("RGB")
    mask_l = mask.convert("L").filter(ImageFilter.GaussianBlur(1))
    bbox = mask_l.getbbox()
    if not bbox:
        return base_rgb
    x0, y0, x1, y1 = bbox
    region_mask = mask_l.crop(bbox)
    tile = ImageOps.fit(pattern, (x1 - x0, y1 - y0), method=Image.Resampling.LANCZOS)
    shading = base_rgb.crop(bbox).convert("L")
    shaded = ImageChops.multiply(tile, Image.merge("RGB", [shading, shading, shading]))
    shaded = ImageEnhance.Contrast(shaded).enhance(1.05)
    shaded = ImageEnhance.Brightness(shaded).enhance(1.02)
    patch = Image.new("RGB", base_rgb.size, (255, 255, 255))
    patch.paste(shaded, (x0, y0), region_mask)
    return Image.composite(patch, base_rgb, mask_l)


def compose_from_template(
    pattern: Image.Image,
    template_dir: Path,
    *,
    canvas_size: tuple[int, int] | None = None,
    size_label: str = "",
) -> Image.Image:
    tpl = MockupTemplate.load(template_dir)
    target_size = canvas_size or tpl.canvas_size
    scale_x = target_size[0] / tpl.canvas_size[0]
    scale_y = target_size[1] / tpl.canvas_size[1]

    def scale_box(box: tuple[int, int, int, int]) -> tuple[int, int, int, int]:
        return (
            int(box[0] * scale_x),
            int(box[1] * scale_y),
            int(box[2] * scale_x),
            int(box[3] * scale_y),
        )

    def scale_quad(quad: tuple[tuple[int, int], ...]) -> tuple[tuple[int, int], ...]:
        return tuple((int(x * scale_x), int(y * scale_y)) for x, y in quad)

    base_path = template_dir / "base.jpg"
    if not base_path.is_file():
        base_path = template_dir / "base.png"
    base = Image.open(base_path).convert("RGB")
    base = ImageOps.fit(base, target_size, method=Image.Resampling.LANCZOS)

    worn_mask = Image.open(template_dir / "mask_worn.png").convert("L")
    worn_mask = ImageOps.fit(worn_mask, target_size, method=Image.Resampling.LANCZOS)

    flat_mask_path = template_dir / "mask_flat.png"
    flat_mask = Image.open(flat_mask_path).convert("L") if flat_mask_path.is_file() else None
    if flat_mask:
        flat_mask = ImageOps.fit(flat_mask, target_size, method=Image.Resampling.LANCZOS)

    canvas = Image.new("RGB", target_size, (255, 255, 255))
    canvas.paste(base, (0, 0))
    canvas = _apply_fabric_with_shading(canvas, pattern, worn_mask)

    if flat_mask:
        bbox = flat_mask.getbbox()
        if bbox:
            x0, y0, x1, y1 = bbox
            fitted = ImageOps.fit(
                pattern.convert("RGB"),
                (max(32, x1 - x0), max(32, y1 - y0)),
                method=Image.Resampling.LANCZOS,
            )
            fitted = fitted.rotate(-5, resample=Image.Resampling.BICUBIC, expand=True)
            patch = Image.new("RGB", target_size, (255, 255, 255))
            px = x0 + ((x1 - x0) - fitted.width) // 2
            py = y0 + ((y1 - y0) - fitted.height) // 2
            region_mask = flat_mask.crop(bbox)
            patch.paste(fitted, (px, py))
            canvas.paste(patch, (0, 0), flat_mask)
    result = canvas

    if size_label:
        draw = ImageDraw.Draw(result)
        font = _load_font(max(20, target_size[0] // 46))
        text_x = int(target_size[0] * 0.52)
        text_y = int(target_size[1] * tpl.size_label_y_ratio)
        draw.text((text_x, text_y), size_label, fill=(45, 45, 45), font=font)

    return result


def compose_ozon_split_scene(
    pattern: Image.Image,
    scene_id: str,
    canvas_size: tuple[int, int],
    mockup_root: Path | None,
    size_label: str = "",
) -> Image.Image | None:
    folder = resolve_template_dir(mockup_root, scene_id)
    if not folder:
        return None
    return compose_from_template(pattern, folder, canvas_size=canvas_size, size_label=size_label)
