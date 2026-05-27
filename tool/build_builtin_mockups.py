#!/usr/bin/env python3
"""从参考场景图生成 builtin_mockups 模板（mask + base + template.json）。"""

from __future__ import annotations

import json
import sys
from pathlib import Path

from PIL import Image, ImageChops, ImageDraw, ImageFilter, ImageOps

OUTPUT_ROOT = Path(__file__).resolve().parent / "builtin_mockups"

# 参考图路径（用户提供的 Ozon 风格场景图）
REFERENCE_SCENES: dict[str, dict] = {
    "headscarf_side": {
        "image": Path(
            r"C:\Users\23393\.cursor\projects\d-ozon-ozon-plsj\assets"
            r"\c__Users_23393_AppData_Roaming_Cursor_User_workspaceStorage_f68d0fb3ded2a28c7943df5365b6381a"
            r"_images_2_06-e0d7c0ec-2e35-4e33-acb2-f1805bf4a01a.png"
        ),
        "left_ratio": 0.54,
        "right_quad_ratio": [(0.52, 0.06), (0.97, 0.04), (0.98, 0.62), (0.50, 0.66)],
    },
    "headscarf_back": {
        "image": Path(
            r"C:\Users\23393\.cursor\projects\d-ozon-ozon-plsj\assets"
            r"\c__Users_23393_AppData_Roaming_Cursor_User_workspaceStorage_f68d0fb3ded2a28c7943df5365b6381a"
            r"_images_2_02-90f101c6-c122-42d1-a2ef-8538171f5458.png"
        ),
        "left_ratio": 0.54,
        "right_quad_ratio": [(0.52, 0.08), (0.96, 0.06), (0.97, 0.64), (0.48, 0.68)],
    },
    "bow_and_fold": {
        "image": Path(
            r"C:\Users\23393\.cursor\projects\d-ozon-ozon-plsj\assets"
            r"\c__Users_23393_AppData_Roaming_Cursor_User_workspaceStorage_f68d0fb3ded2a28c7943df5365b6381a"
            r"_images_2_01-bf4fc246-7c0b-4f67-b257-bac591d2c11a.png"
        ),
        "left_ratio": 0.52,
        "right_quad_ratio": [(0.50, 0.12), (0.96, 0.10), (0.97, 0.70), (0.46, 0.72)],
    },
    "size_chart": {
        "image": Path(
            r"C:\Users\23393\.cursor\projects\d-ozon-ozon-plsj\assets"
            r"\c__Users_23393_AppData_Roaming_Cursor_User_workspaceStorage_f68d0fb3ded2a28c7943df5365b6381a"
            r"_images_2_01-bf4fc246-7c0b-4f67-b257-bac591d2c11a.png"
        ),
        "left_ratio": 0.52,
        "right_quad_ratio": [(0.50, 0.12), (0.96, 0.10), (0.97, 0.70), (0.46, 0.72)],
        "size_label_y_ratio": 0.84,
    },
}


def _not_white_mask(region: Image.Image, threshold: int = 238) -> Image.Image:
    gray = region.convert("L")
    return gray.point(lambda value: 255 if value < threshold else 0)


def _is_skin(r: int, g: int, b: int) -> bool:
    if r < 95 or g < 60 or b < 45:
        return False
    if r > 240 and g > 220 and b > 200:
        return True
    return r > g > b and (r - b) < 95 and g > 110


def _fabric_mask(region: Image.Image) -> Image.Image:
    rgb = region.convert("RGB")
    pixels = rgb.load()
    w, h = rgb.size
    mask = Image.new("L", (w, h), 0)
    mp = mask.load()
    for y in range(h):
        for x in range(w):
            r, g, b = pixels[x, y]
            if r > 245 and g > 245 and b > 245:
                continue
            if _is_skin(r, g, b):
                continue
            if g > 145 and r < 150 and b < 150:
                mp[x, y] = 255
                continue
            if r > 135 and g < 115 and b < 115:
                mp[x, y] = 255
                continue
            if r < 120 and g > 120 and b < 120:
                mp[x, y] = 255
    return mask.filter(ImageFilter.MaxFilter(3)).filter(ImageFilter.GaussianBlur(1))


def build_scene(scene_id: str, cfg: dict) -> None:
    ref_path: Path = cfg["image"]
    if not ref_path.is_file():
        raise FileNotFoundError(ref_path)
    ref = Image.open(ref_path).convert("RGB")
    w, h = ref.size
    left_cut = int(w * float(cfg["left_ratio"]))

    left = ref.crop((0, 0, left_cut, h))
    right = ref.crop((left_cut, 0, w, h))

    worn_mask = _fabric_mask(left)
    face_cut = Image.new("L", (left_cut, h), 0)
    face_draw = ImageDraw.Draw(face_cut)
    face_draw.ellipse((int(left_cut * 0.28), int(h * 0.12), int(left_cut * 0.82), int(h * 0.72)), fill=255)
    worn_mask = ImageChops.subtract(worn_mask, face_cut)
    hair_keep = Image.new("L", (left_cut, h), 0)
    hair_draw = ImageDraw.Draw(hair_keep)
    hair_draw.rectangle((0, 0, left_cut, int(h * 0.42)), fill=255)
    worn_mask = ImageChops.multiply(worn_mask, hair_keep)
    flat_mask_full = Image.new("L", (w, h), 0)
    flat_mask_right = _fabric_mask(right)
    flat_mask_full.paste(flat_mask_right, (left_cut, 0))

    base = ref.copy()
    draw = ImageDraw.Draw(base)
    draw.rectangle((left_cut, 0, w, h), fill=(255, 255, 255))

    quad_ratio = cfg["right_quad_ratio"]
    right_quad = tuple((int(w * x), int(h * y)) for x, y in quad_ratio)

    out_dir = OUTPUT_ROOT / scene_id
    out_dir.mkdir(parents=True, exist_ok=True)
    base.save(out_dir / "base.jpg", quality=95)
    worn_mask.save(out_dir / "mask_worn.png")
    flat_mask_full.save(out_dir / "mask_flat.png")

    meta = {
        "scene_id": scene_id,
        "canvas": [w, h],
        "left_box": [0, 0, left_cut, h],
        "right_quad": [list(p) for p in right_quad],
        "size_label_y_ratio": float(cfg.get("size_label_y_ratio", 0.86)),
        "source_reference": ref_path.name,
    }
    (out_dir / "template.json").write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"OK {scene_id} -> {out_dir}")


def main() -> int:
    for scene_id, cfg in REFERENCE_SCENES.items():
        build_scene(scene_id, cfg)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
