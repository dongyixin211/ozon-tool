#!/usr/bin/env python3
"""从平面印花图批量生成电商使用场景图（命令行）。

示例:
  python generate_scenes.py -i "D:/patterns/scarf.png" -o "D:/output/scenes"
  python generate_scenes.py -i "D:/patterns/scarf.png" -o "D:/out" --count 8 --ratio 3:4
  python generate_scenes.py --source "D:/原图" --output "D:/场景图" --count 6
  python generate_scenes.py -i "D:/patterns/scarf.png" -o "D:/out" --local --count 8
  python generate_scenes.py -i "scarf.png" -o "D:/out" --local --mockup "D:/mockups"
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

TOOL_DIR = Path(__file__).resolve().parent.parent / "tool"
sys.path.insert(0, str(TOOL_DIR))

from image_providers import (  # noqa: E402
    DEFAULT_IMAGE_PROVIDER,
    create_image_client,
    list_image_provider_ids,
    migrate_provider_api_keys,
    provider_default_model,
)
from local_scene_composer import (  # noqa: E402
    DEFAULT_SIZE_LABEL,
    LocalSceneJobConfig,
    LocalSceneWorker,
)
from scene_generator import (  # noqa: E402
    DEFAULT_ASPECT_RATIO,
    DEFAULT_SCENE_PROMPT_TEMPLATE,
    SceneGenerationWorker,
    SceneJobConfig,
)

DEFAULT_IMAGE_MODEL = "gpt-image-2"
CONFIG_PATH = Path(__file__).with_name("config.json")


def humanize_api_error(raw_message: str) -> str:
    text = raw_message.strip()
    lowered = text.lower()
    if "billing hard limit" in lowered or "insufficient_quota" in lowered:
        return "账户额度不足，请检查 API 余额。"
    if "invalid_api_key" in lowered or "incorrect api key" in lowered:
        return "API Key 无效。"
    if "rate_limit" in lowered:
        return "请求过于频繁，请稍后再试。"
    return text


def load_defaults() -> dict:
    local = Path(__file__).with_name("config.json")
    if local.exists():
        return json.loads(local.read_text(encoding="utf-8"))
    return {}


def main() -> int:
    cfg = load_defaults()
    parser = argparse.ArgumentParser(description="平面印花图 -> 电商使用场景图")
    parser.add_argument("-i", "--image", help="单张平面原图路径")
    parser.add_argument("--source", help="SKU 文件夹根目录（每个子文件夹一个货号）")
    parser.add_argument("-o", "--output", required=True, help="场景图输出目录")
    parser.add_argument("--count", type=int, default=int(cfg.get("scene_count") or 8), help="生成场景数量 1-10")
    parser.add_argument("--ratio", default=str(cfg.get("scene_aspect_ratio") or DEFAULT_ASPECT_RATIO), help="比例: 1:1, 3:4, 4:3, 16:9 或 1024x1536")
    parser.add_argument("--quality", default=str(cfg.get("quality") or "high"))
    parser.add_argument("--workers", type=int, default=int(cfg.get("scene_max_workers") or cfg.get("max_workers") or 2))
    parser.add_argument("--sku", default="", help="只处理指定货号文件夹")
    provider_keys = migrate_provider_api_keys(cfg)
    default_provider = str(cfg.get("image_provider") or DEFAULT_IMAGE_PROVIDER)
    default_api_key = provider_keys.get(default_provider) or str(cfg.get("image_api_key") or "")
    default_image_url = str(cfg.get("image_base_url") or cfg.get("base_url") or "https://breakout.wenwen-ai.com/v1").rstrip("/")
    parser.add_argument("--provider", default=default_provider, choices=list_image_provider_ids())
    parser.add_argument("--api-key", default=default_api_key)
    parser.add_argument("--base-url", default=default_image_url)
    parser.add_argument(
        "--model",
        default=str(cfg.get("image_model") or provider_default_model(default_provider) or DEFAULT_IMAGE_MODEL),
    )
    parser.add_argument(
        "--scene-ids",
        default=str(cfg.get("scene_ids") or ""),
        help="逗号分隔场景 ID，留空则按顺序取前 N 个",
    )
    parser.add_argument("--local", action="store_true", help="本地 Pillow 合成，不调用 AI、不消耗 API Key")
    parser.add_argument("--mockup", default=str(cfg.get("scene_mockup_root") or ""), help="模特/场景底图目录（本地模式）")
    parser.add_argument("--size-label", default=str(cfg.get("scene_size_label") or DEFAULT_SIZE_LABEL))
    args = parser.parse_args()

    if not args.image and not args.source:
        parser.error("请指定 -i/--image 单图，或 --source 批量目录")
    if args.image and args.source:
        parser.error("-i 与 --source 不能同时使用")

    count = max(1, min(10, args.count))
    scene_ids = [item.strip() for item in args.scene_ids.split(",") if item.strip()]

    def log(msg: str) -> None:
        print(msg, flush=True)

    if args.local:
        mockup_root = Path(args.mockup).expanduser() if args.mockup.strip() else None
        if mockup_root and not mockup_root.is_dir():
            print("错误: 模特底图目录不存在。", file=sys.stderr)
            return 1
        job = LocalSceneJobConfig(
            source_root=Path(args.source).expanduser() if args.source else Path("."),
            output_root=Path(args.output).expanduser(),
            single_image=Path(args.image).expanduser() if args.image else None,
            aspect_ratio=args.ratio,
            scene_count=count,
            scene_ids=scene_ids,
            mockup_root=mockup_root,
            size_label=args.size_label,
            sku_filter=args.sku,
        )
        try:
            LocalSceneWorker(job, log).run()
        except Exception as exc:  # noqa: BLE001
            print(f"失败: {exc}", file=sys.stderr)
            return 1
        return 0

    if not args.api_key.strip():
        print("错误: 未配置 image_api_key，请使用 --local 或填写 API Key。", file=sys.stderr)
        return 1

    prompt_template = str(cfg.get("scene_prompt_template") or DEFAULT_SCENE_PROMPT_TEMPLATE)

    job = SceneJobConfig(
        source_root=Path(args.source).expanduser() if args.source else Path("."),
        output_root=Path(args.output).expanduser(),
        single_image=Path(args.image).expanduser() if args.image else None,
        aspect_ratio=args.ratio,
        scene_count=count,
        scene_ids=scene_ids,
        scene_prompt_template=prompt_template,
        quality=args.quality,
        max_workers=max(1, args.workers),
        sku_filter=args.sku,
    )

    client_holder: dict[str, object] = {}

    def client_factory() -> object:
        if "client" not in client_holder:
            client_holder["client"] = create_image_client(
                args.provider,
                args.api_key.strip(),
                args.base_url,
                args.model,
            )
        return client_holder["client"]

    worker = SceneGenerationWorker(job, client_factory, log)
    try:
        worker.run()
    except Exception as exc:  # noqa: BLE001
        print(f"失败: {humanize_api_error(str(exc))}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
