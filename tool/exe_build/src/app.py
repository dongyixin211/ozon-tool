import base64
import concurrent.futures
import json
import mimetypes
import os
import queue
import re
import subprocess
import threading
import time
import uuid
from dataclasses import dataclass
from pathlib import Path
import sys
from typing import Callable, Iterable
from urllib import error, request

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from PIL import Image, ImageFilter, ImageOps


SUPPORTED_IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".webp", ".bmp"}
APP_DIR = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
RESOURCE_DIR = Path(getattr(sys, "_MEIPASS", APP_DIR))
CONFIG_PATH = APP_DIR / "config.json"
NODE_EXECUTABLE = Path(
    r"C:\Users\23393\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe"
)
EXCEL_BUILDER_PATH = RESOURCE_DIR / "excel_export_builder.mjs"

DEFAULT_BASE_URL = "https://breakout.wenwen-ai.com/v1"
DEFAULT_IMAGE_MODEL = "gpt-image-2"
DEFAULT_TEXT_MODEL = "gpt-5-high"
DEFAULT_GENERATION_SIZE = "1024x1536"
DEFAULT_DOWNLOAD_RETRIES = 3

DEFAULT_IMAGE_PROMPT = (
    "根据参考图生成适合电商平台的产品竖版主图，保持主体一致，画面干净，"
    "白底或浅色高级背景，突出产品细节，适合俄罗斯电商平台展示。"
    "输出适合后续裁切为3:4比例。货号：{sku}。"
)

DEFAULT_TITLE_PROMPT = (
    "你是 Ozon 电商标题优化助手。请根据商品图片和基础信息，"
    "生成一个适合 Ozon 平台的中文商品标题。"
    "要求：简洁清晰，突出商品核心属性和卖点，不要堆砌词。"
    "货号：{sku}；文件夹：{folder_name}；图片数量：{image_count}；文件名：{image_names}。"
)

DEFAULT_DESCRIPTION_PROMPT = (
    "你是 Ozon 电商卖点文案助手。请根据商品图片和基础信息，"
    "生成适合 Ozon 商品详情页的中文描述和 5 条卖点。"
    "要求：表达自然，突出功能、材质、使用场景或优势；如果图片信息不足，不要编造危险参数。"
    "货号：{sku}；文件夹：{folder_name}；图片数量：{image_count}；文件名：{image_names}。"
)
DEFAULT_TEMPLATE_NAME = "默认模板"


def now_stamp() -> str:
    return time.strftime("%Y-%m-%d %H:%M:%S")


def list_subfolders(root: Path) -> list[Path]:
    return sorted([item for item in root.iterdir() if item.is_dir()])


def list_images(folder: Path) -> list[Path]:
    return sorted(
        [item for item in folder.iterdir() if item.is_file() and item.suffix.lower() in SUPPORTED_IMAGE_EXTENSIONS]
    )


def make_relative_output(root: Path, base: Path, target_root: Path) -> Path:
    return target_root / root.relative_to(base)


def image_to_data_url(path: Path) -> str:
    mime = mimetypes.guess_type(path.name)[0] or "application/octet-stream"
    encoded = base64.b64encode(path.read_bytes()).decode("ascii")
    return f"data:{mime};base64,{encoded}"


def encode_multipart(fields: list[tuple[str, str]], files: list[tuple[str, str, bytes, str]]) -> tuple[bytes, str]:
    boundary = f"----CodexBoundary{uuid.uuid4().hex}"
    chunks: list[bytes] = []

    for name, value in fields:
        chunks.extend(
            [
                f"--{boundary}\r\n".encode("utf-8"),
                f'Content-Disposition: form-data; name="{name}"\r\n\r\n'.encode("utf-8"),
                value.encode("utf-8"),
                b"\r\n",
            ]
        )

    for field_name, filename, payload, content_type in files:
        chunks.extend(
            [
                f"--{boundary}\r\n".encode("utf-8"),
                f'Content-Disposition: form-data; name="{field_name}"; filename="{filename}"\r\n'.encode("utf-8"),
                f"Content-Type: {content_type}\r\n\r\n".encode("utf-8"),
                payload,
                b"\r\n",
            ]
        )

    chunks.append(f"--{boundary}--\r\n".encode("utf-8"))
    body = b"".join(chunks)
    return body, f"multipart/form-data; boundary={boundary}"


def humanize_api_error(raw_message: str) -> str:
    text = raw_message.strip()
    lowered = text.lower()

    code_match = re.search(r'"code"\s*:\s*"([^"]+)"', text)
    error_code = code_match.group(1) if code_match else ""

    if "billing hard limit has been reached" in lowered or error_code == "billing_hard_limit_reached":
        return "账户已触发账单硬限制，请检查网关余额或上游账户。"
    if "insufficient_quota" in lowered:
        return "账户额度不足，请检查网关余额或令牌额度。"
    if "invalid_api_key" in lowered or "incorrect api key" in lowered:
        return "API Key 无效，请检查配置。"
    if "model_not_found" in lowered:
        return "模型不可用，请确认模型名或网关权限。"
    if "rate_limit" in lowered or "too many requests" in lowered:
        return "请求过于频繁，触发了限流，请稍后再试。"
    if "authentication" in lowered or "unauthorized" in lowered:
        return "鉴权失败，请检查 API Key 和接口地址。"
    if "network error" in lowered:
        return "网络连接失败，当前无法访问接口服务。"
    return text


def sanitize_json_text(raw_text: str) -> str:
    text = raw_text.strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*", "", text)
        text = re.sub(r"\s*```$", "", text)
    return text.strip()


def parse_copy_payload(raw_text: str) -> dict:
    text = sanitize_json_text(raw_text)
    if not text:
        raise RuntimeError("文本模型返回了空内容，无法生成文案。")
    try:
        data = json.loads(text)
    except json.JSONDecodeError:
        return {
            "title": "",
            "description": text,
            "bullets": [],
        }
    bullets = data.get("bullets") or data.get("selling_points") or []
    if not isinstance(bullets, list):
        bullets = [str(bullets)]
    return {
        "title": str(data.get("title", "")).strip(),
        "description": str(data.get("description", "")).strip(),
        "bullets": [str(item).strip() for item in bullets if str(item).strip()],
    }


def parse_title_payload(raw_text: str) -> str:
    text = sanitize_json_text(raw_text)
    if not text:
        raise RuntimeError("文本模型返回了空内容，无法生成标题。")
    try:
        data = json.loads(text)
        if isinstance(data, dict):
            return str(data.get("title", "")).strip()
    except json.JSONDecodeError:
        pass
    return text.strip()

def apply_image_watermark(
    base_image: Image.Image,
    watermark_path: Path,
    *,
    scale_ratio: float = 0.18,
    margin: int = 36,
) -> Image.Image:
    with Image.open(watermark_path) as watermark_source:
        watermark = watermark_source.convert("RGBA")

    max_width = max(1, int(base_image.width * scale_ratio))
    if watermark.width > max_width:
        resize_ratio = max_width / watermark.width
        watermark = watermark.resize(
            (max_width, max(1, int(watermark.height * resize_ratio))),
            Image.Resampling.LANCZOS,
        )

    layer = Image.new("RGBA", base_image.size, (0, 0, 0, 0))
    x = max(0, margin)
    y = max(0, base_image.height - watermark.height - margin)
    layer.alpha_composite(watermark, (x, y))
    return Image.alpha_composite(base_image.convert("RGBA"), layer).convert("RGB")


def create_portrait_variant(
    source_path: Path,
    destination_path: Path,
    size: tuple[int, int] = (1200, 1600),
    watermark_path: Path | None = None,
) -> None:
    with Image.open(source_path) as original:
        base = original.convert("RGB")
        canvas = ImageOps.fit(base, size, method=Image.Resampling.LANCZOS)
        canvas = canvas.filter(ImageFilter.GaussianBlur(radius=22))
        canvas = ImageOps.colorize(ImageOps.grayscale(canvas), black="#1e1e1e", white="#f8f8f8")

        foreground = ImageOps.contain(base, size, method=Image.Resampling.LANCZOS)
        composed = canvas.copy()
        offset = ((size[0] - foreground.width) // 2, (size[1] - foreground.height) // 2)
        composed.paste(foreground, offset)
        if watermark_path and watermark_path.is_file():
            composed = apply_image_watermark(composed, watermark_path)
        destination_path.parent.mkdir(parents=True, exist_ok=True)
        composed.save(destination_path, format="PNG")


def save_copy_files(destination_dir: Path, sku: str, payload: dict) -> tuple[Path, Path]:
    destination_dir.mkdir(parents=True, exist_ok=True)
    json_path = destination_dir / f"{sku}_content.json"
    txt_path = destination_dir / f"{sku}_content.txt"

    json_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    lines = [
        f"标题：{payload.get('title', '')}",
        "",
        "描述：",
        payload.get("description", ""),
        "",
        "卖点：",
    ]
    for index, item in enumerate(payload.get("bullets", []), start=1):
        lines.append(f"{index}. {item}")
    txt_path.write_text("\n".join(lines), encoding="utf-8")
    return json_path, txt_path


def choose_first_portrait_image(folder: Path) -> Path | None:
    if not folder.exists():
        return None
    return next(
        (
            item
            for item in sorted(folder.iterdir())
            if item.is_file() and item.suffix.lower() in SUPPORTED_IMAGE_EXTENSIONS
        ),
        None,
    )


def export_copy_to_excel(destination_dir: Path, rows: list[dict]) -> Path:
    destination_dir.mkdir(parents=True, exist_ok=True)
    input_path = destination_dir / "_excel_export_input.json"
    output_path = destination_dir / "ozon_content_export.xlsx"
    payload = {"rows": rows, "output_path": str(output_path)}
    input_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    completed = subprocess.run(
        [str(NODE_EXECUTABLE), str(EXCEL_BUILDER_PATH), str(input_path)],
        capture_output=True,
        text=True,
        encoding="utf-8",
        check=False,
    )
    if completed.returncode != 0:
        details = completed.stderr.strip() or completed.stdout.strip() or "unknown error"
        raise RuntimeError(f"Excel 导出失败: {details}")
    return output_path


class GatewayClient:
    def __init__(self, api_key: str, base_url: str, timeout_seconds: int = 180):
        self.api_key = api_key.strip()
        self.base_url = base_url.rstrip("/")
        self.timeout_seconds = timeout_seconds

    def _headers(self, content_type: str | None = None) -> dict[str, str]:
        headers = {"Authorization": f"Bearer {self.api_key}"}
        if content_type:
            headers["Content-Type"] = content_type
        return headers

    def _request_json(self, method: str, endpoint: str, payload: bytes | None = None, content_type: str | None = None) -> dict:
        req = request.Request(
            f"{self.base_url}{endpoint}",
            data=payload,
            headers=self._headers(content_type),
            method=method,
        )
        try:
            with request.urlopen(req, timeout=self.timeout_seconds) as response:
                return json.loads(response.read().decode("utf-8"))
        except error.HTTPError as exc:
            details = exc.read().decode("utf-8", errors="replace")
            raise RuntimeError(f"HTTP {exc.code}: {details}") from exc
        except error.URLError as exc:
            raise RuntimeError(f"Network error: {exc.reason}") from exc

    def list_models(self) -> dict:
        return self._request_json("GET", "/models")


class OpenAICompatibleClient(GatewayClient):
    def __init__(self, api_key: str, base_url: str, image_model: str, text_model: str, timeout_seconds: int = 180):
        super().__init__(api_key, base_url, timeout_seconds=timeout_seconds)
        self.image_model = image_model.strip()
        self.text_model = text_model.strip()

    def _generate_image_without_references(self, prompt: str, output_size: str, quality: str) -> dict:
        payload = json.dumps(
            {"model": self.image_model, "prompt": prompt, "size": output_size, "quality": quality}
        ).encode("utf-8")
        return self._request_json("POST", "/images/generations", payload, "application/json")

    def generate_image(self, prompt: str, output_size: str, quality: str, reference_images: Iterable[Path] | None = None) -> bytes:
        refs = list(reference_images or [])
        if refs:
            fields = [("model", self.image_model), ("prompt", prompt), ("size", output_size), ("quality", quality)]
            files = []
            for ref_path in refs[:1]:
                mime = mimetypes.guess_type(ref_path.name)[0] or "application/octet-stream"
                files.append(("image", ref_path.name, ref_path.read_bytes(), mime))
            payload, content_type = encode_multipart(fields, files)
            data = self._request_json("POST", "/images/edits", payload, content_type)
        else:
            data = self._generate_image_without_references(prompt, output_size, quality)
        return self._extract_image_bytes(data)

    def generate_image_with_fallback(
        self, prompt: str, output_size: str, quality: str, reference_images: Iterable[Path] | None = None
    ) -> tuple[bytes, str]:
        refs = list(reference_images or [])
        try:
            if refs:
                return self.generate_image(prompt, output_size, quality, reference_images=refs), "edit"
            return self.generate_image(prompt, output_size, quality, reference_images=[]), "generation"
        except Exception as exc:  # noqa: BLE001
            if not refs:
                raise
            fallback_prompt = (
                f"{prompt}\n补充要求：参考原图主体样式与构图，保持商品主体清晰，适合电商详情展示。"
            )
            note = f"参考图编辑失败，已回退为文生图。原因: {humanize_api_error(str(exc))}"
            return self.generate_image(fallback_prompt, output_size, quality, reference_images=[]), note

    def _download_image_from_url(self, image_url: str) -> bytes:
        attempts = [
            {
                "User-Agent": (
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
                ),
                "Accept": "image/avif,image/webp,image/apng,image/svg+xml,image/*,*/*;q=0.8",
            },
            {
                **self._headers(),
                "User-Agent": (
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
                ),
                "Accept": "image/avif,image/webp,image/apng,image/svg+xml,image/*,*/*;q=0.8",
            },
        ]

        last_error: Exception | None = None
        for _ in range(DEFAULT_DOWNLOAD_RETRIES):
            for headers in attempts:
                req = request.Request(image_url, headers=headers)
                try:
                    with request.urlopen(req, timeout=self.timeout_seconds) as response:
                        return response.read()
                except error.HTTPError as exc:
                    details = exc.read().decode("utf-8", errors="replace")
                    last_error = RuntimeError(f"下载图片失败 HTTP {exc.code}: {details}")
                except error.URLError as exc:
                    last_error = RuntimeError(f"下载图片失败: {exc.reason}")
            time.sleep(1.2)

        assert last_error is not None
        raise last_error

    def _extract_image_bytes(self, data: dict) -> bytes:
        items = data.get("data")
        if not isinstance(items, list) or not items:
            raise RuntimeError(f"图片接口返回缺少 data 字段: {json.dumps(data, ensure_ascii=False)[:500]}")

        first = items[0]
        if not isinstance(first, dict):
            raise RuntimeError(f"图片接口返回结构异常: {json.dumps(data, ensure_ascii=False)[:500]}")

        image_base64 = first.get("b64_json")
        if image_base64:
            return base64.b64decode(image_base64)

        image_url = first.get("url")
        if image_url:
            return self._download_image_from_url(image_url)

        raise RuntimeError(f"图片接口返回里既没有 b64_json 也没有 url: {json.dumps(first, ensure_ascii=False)[:500]}")

    def _build_multimodal_messages(self, prompt: str, reference_images: Iterable[Path] | None = None) -> list[dict]:
        content_items: list[dict] = [{"type": "text", "text": prompt}]
        for ref_path in list(reference_images or [])[:1]:
            content_items.append({"type": "image_url", "image_url": {"url": image_to_data_url(ref_path)}})
        return [{"role": "user", "content": content_items}]

    def generate_title(self, prompt: str, reference_images: Iterable[Path] | None = None) -> str:
        payload = json.dumps(
            {
                "model": self.text_model,
                "messages": [
                    {"role": "system", "content": '你是资深 Ozon 电商标题编辑。请只返回 JSON：{"title":""}'},
                    *self._build_multimodal_messages(prompt, reference_images),
                ],
                "temperature": 0.7,
            }
        ).encode("utf-8")
        data = self._request_json("POST", "/chat/completions", payload, "application/json")
        raw_text = data.get("choices", [{}])[0].get("message", {}).get("content", "")
        try:
            return parse_title_payload(raw_text)
        except Exception as exc:  # noqa: BLE001
            preview = sanitize_json_text(str(raw_text))[:200]
            raise RuntimeError(f"标题解析失败: {exc}。返回内容预览: {preview or '[empty]'}") from exc

    def generate_copy(self, prompt: str, reference_images: Iterable[Path] | None = None) -> dict:
        payload = json.dumps(
            {
                "model": self.text_model,
                "messages": [
                    {
                        "role": "system",
                        "content": '你是资深 Ozon 电商文案编辑。请只返回 JSON：{"description":"","bullets":["","","","",""]}',
                    },
                    *self._build_multimodal_messages(prompt, reference_images),
                ],
                "temperature": 0.7,
            }
        ).encode("utf-8")
        data = self._request_json("POST", "/chat/completions", payload, "application/json")
        raw_text = data.get("choices", [{}])[0].get("message", {}).get("content", "")
        try:
            return parse_copy_payload(raw_text)
        except Exception as exc:  # noqa: BLE001
            preview = sanitize_json_text(str(raw_text))[:200]
            raise RuntimeError(f"文案解析失败: {exc}。返回内容预览: {preview or '[empty]'}") from exc


class TextOnlyClient(GatewayClient):
    def __init__(self, api_key: str, base_url: str, text_model: str, timeout_seconds: int = 180):
        super().__init__(api_key, base_url, timeout_seconds=timeout_seconds)
        self.text_model = text_model.strip()

    def _build_multimodal_messages(self, prompt: str, reference_images: Iterable[Path] | None = None) -> list[dict]:
        content_items: list[dict] = [{"type": "text", "text": prompt}]
        for ref_path in list(reference_images or [])[:1]:
            content_items.append({"type": "image_url", "image_url": {"url": image_to_data_url(ref_path)}})
        return [{"role": "user", "content": content_items}]

    def generate_title(self, prompt: str, reference_images: Iterable[Path] | None = None) -> str:
        payload = json.dumps(
            {
                "model": self.text_model,
                "messages": [
                    {"role": "system", "content": '你是资深 Ozon 电商标题编辑。请只返回 JSON：{"title":""}'},
                    *self._build_multimodal_messages(prompt, reference_images),
                ],
                "temperature": 0.7,
            }
        ).encode("utf-8")
        data = self._request_json("POST", "/chat/completions", payload, "application/json")
        raw_text = data.get("choices", [{}])[0].get("message", {}).get("content", "")
        try:
            return parse_title_payload(raw_text)
        except Exception as exc:  # noqa: BLE001
            preview = sanitize_json_text(str(raw_text))[:200]
            raise RuntimeError(f"标题解析失败: {exc}。返回内容预览: {preview or '[empty]'}") from exc

    def generate_copy(self, prompt: str, reference_images: Iterable[Path] | None = None) -> dict:
        payload = json.dumps(
            {
                "model": self.text_model,
                "messages": [
                    {
                        "role": "system",
                        "content": '你是资深 Ozon 电商文案编辑。请只返回 JSON：{"description":"","bullets":["","","","",""]}',
                    },
                    *self._build_multimodal_messages(prompt, reference_images),
                ],
                "temperature": 0.7,
            }
        ).encode("utf-8")
        data = self._request_json("POST", "/chat/completions", payload, "application/json")
        raw_text = data.get("choices", [{}])[0].get("message", {}).get("content", "")
        try:
            return parse_copy_payload(raw_text)
        except Exception as exc:  # noqa: BLE001
            preview = sanitize_json_text(str(raw_text))[:200]
            raise RuntimeError(f"文案解析失败: {exc}。返回内容预览: {preview or '[empty]'}") from exc


@dataclass
class JobConfig:
    source_root: Path
    generated_root: Path
    portrait_root: Path
    content_root: Path
    image_api_key: str
    text_api_key: str
    base_url: str
    image_model: str
    text_model: str
    image_prompt_template: str
    title_prompt_template: str
    description_prompt_template: str
    watermark_path: Path | None
    quality: str
    max_folders: int
    max_workers: int
    convert_originals: bool
    generate_copy: bool
    export_excel: bool


class BatchWorker:
    def __init__(self, config: JobConfig, logger: Callable[[str], None]):
        self.config = config
        self.logger = logger
        self._cancelled = False
        self.export_rows: list[dict] = []

    def cancel(self) -> None:
        self._cancelled = True

    def _generate_single_image(
        self,
        client: OpenAICompatibleClient,
        sku: str,
        sku_folder: Path,
        images: list[Path],
        source_image: Path,
        generated_folder: Path,
        portrait_folder: Path,
    ) -> tuple[Path, Path, str]:
        self.logger(f"    开始处理图片: {source_image.name}")
        image_prompt = self.config.image_prompt_template.format(
            sku=sku,
            folder_name=sku_folder.name,
            image_count=len(images),
            image_names=source_image.name,
        )
        self.logger(f"      正在请求生图接口: {source_image.name}")
        generated_bytes, mode_note = client.generate_image_with_fallback(
            prompt=image_prompt,
            output_size=DEFAULT_GENERATION_SIZE,
            quality=self.config.quality,
            reference_images=[source_image],
        )
        generated_path = generated_folder / f"{source_image.stem}_ai_vertical.png"
        generated_path.write_bytes(generated_bytes)
        self.logger(f"      已保存竖版图: {generated_path.name}")
        portrait_path = portrait_folder / f"{source_image.stem}_ai_portrait.png"
        self.logger(f"      正在生成 3:4 图: {source_image.name}")
        create_portrait_variant(generated_path, portrait_path, watermark_path=self.config.watermark_path)
        self.logger(f"      已保存 3:4 图: {portrait_path.name}")
        return generated_path, portrait_path, mode_note

    def run(self) -> None:
        folders = list_subfolders(self.config.source_root)
        if self.config.max_folders > 0:
            folders = folders[: self.config.max_folders]
        if not folders:
            raise RuntimeError("源目录下没有找到任何子文件夹。")

        client = OpenAICompatibleClient(
            api_key=self.config.image_api_key,
            base_url=self.config.base_url,
            image_model=self.config.image_model,
            text_model=self.config.text_model,
        )
        text_client = TextOnlyClient(
            api_key=self.config.text_api_key,
            base_url=self.config.base_url,
            text_model=self.config.text_model,
        )

        total = len(folders)
        for index, sku_folder in enumerate(folders, start=1):
            if self._cancelled:
                self.logger("任务已取消。")
                return

            sku = sku_folder.name
            images = list_images(sku_folder)
            image_names = ", ".join(item.name for item in images) if images else "无"

            self.logger(f"[{index}/{total}] 处理货号 {sku}")
            generated_folder = make_relative_output(sku_folder, self.config.source_root, self.config.generated_root)
            generated_folder.mkdir(parents=True, exist_ok=True)
            portrait_folder = make_relative_output(sku_folder, self.config.source_root, self.config.portrait_root)

            if images:
                worker_count = min(max(1, self.config.max_workers), len(images))
                self.logger(f"  检测到原图 {len(images)} 张，当前并发数 {worker_count}")
                with concurrent.futures.ThreadPoolExecutor(max_workers=worker_count) as executor:
                    future_map = {}
                    for queue_index, source_image in enumerate(images, start=1):
                        self.logger(f"    [{queue_index}/{len(images)}] 已加入队列: {source_image.name}")
                        future = executor.submit(
                            self._generate_single_image,
                            client,
                            sku,
                            sku_folder,
                            images,
                            source_image,
                            generated_folder,
                            portrait_folder,
                        )
                        future_map[future] = source_image

                    completed = 0
                    for future in concurrent.futures.as_completed(future_map):
                        source_image = future_map[future]
                        completed += 1
                        if self._cancelled:
                            executor.shutdown(wait=False, cancel_futures=True)
                            self.logger("任务已取消。")
                            return
                        try:
                            generated_path, portrait_path, mode_note = future.result()
                            self.logger(
                                f"    [{completed}/{len(images)}] 已完成 {source_image.name} -> {generated_path.name}, {portrait_path.name}"
                            )
                            if mode_note != "edit":
                                self.logger(f"      {mode_note}")
                        except Exception as exc:  # noqa: BLE001
                            self.logger(f"    [{completed}/{len(images)}] 图片 {source_image.name} 失败: {humanize_api_error(str(exc))}")
            else:
                self.logger("  未发现参考图，将直接文生图生成 1 张")
                image_prompt = self.config.image_prompt_template.format(
                    sku=sku,
                    folder_name=sku_folder.name,
                    image_count=0,
                    image_names=image_names,
                )
                generated_bytes = client.generate_image(
                    prompt=image_prompt,
                    output_size=DEFAULT_GENERATION_SIZE,
                    quality=self.config.quality,
                    reference_images=[],
                )
                generated_path = generated_folder / f"{sku}_ai_vertical.png"
                generated_path.write_bytes(generated_bytes)
                portrait_path = portrait_folder / f"{sku}_ai_portrait.png"
                create_portrait_variant(generated_path, portrait_path, watermark_path=self.config.watermark_path)
                self.logger(f"  已保存竖版图: {generated_path}")
                self.logger(f"  已保存 3:4 图: {portrait_path}")

            if self.config.generate_copy:
                content_folder = make_relative_output(sku_folder, self.config.source_root, self.config.content_root)
                reference_portrait = choose_first_portrait_image(portrait_folder)
                copy_refs = [reference_portrait] if reference_portrait else []
                self.logger("  开始生成文案")
                if reference_portrait:
                    self.logger(f"    文案参考图: {reference_portrait.name}")
                else:
                    self.logger("    文案参考图: 未找到，将仅基于提示词生成")

                title_prompt = self.config.title_prompt_template.format(
                    sku=sku,
                    folder_name=sku_folder.name,
                    image_count=len(images),
                    image_names=image_names,
                )
                description_prompt = self.config.description_prompt_template.format(
                    sku=sku,
                    folder_name=sku_folder.name,
                    image_count=len(images),
                    image_names=image_names,
                )
                try:
                    self.logger("    正在生成标题")
                    title_text = text_client.generate_title(title_prompt, reference_images=copy_refs)
                    self.logger("    标题生成完成，正在生成描述和卖点")
                    copy_payload = text_client.generate_copy(description_prompt, reference_images=copy_refs)
                    copy_payload["title"] = title_text
                    json_path, txt_path = save_copy_files(content_folder, sku, copy_payload)
                    self.logger(f"  已保存文案 JSON: {json_path}")
                    self.logger(f"  已保存文案文本: {txt_path}")
                    self.export_rows.append(
                        {
                            "sku": sku,
                            "product_code": sku_folder.name,
                            "title": copy_payload.get("title", ""),
                            "summary": copy_payload.get("description", ""),
                            "rich_json": json.dumps(copy_payload, ensure_ascii=False),
                            "upload_status": "",
                        }
                    )
                except Exception as exc:  # noqa: BLE001
                    self.logger(f"  文案生成失败: {humanize_api_error(str(exc))}")

        if self.config.generate_copy and self.config.export_excel and self.export_rows:
            excel_path = export_copy_to_excel(self.config.content_root, self.export_rows)
            self.logger(f"已导出 Excel 汇总表: {excel_path}")

        self.logger("全部处理完成。")


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Ozon AI 图文素材工具 - MVP")
        self.geometry("1080x900")
        self.minsize(940, 780)

        self.log_queue: queue.Queue[str] = queue.Queue()
        self.worker: BatchWorker | None = None
        self.worker_thread: threading.Thread | None = None

        self.source_root_var = tk.StringVar()
        self.generated_root_var = tk.StringVar()
        self.portrait_root_var = tk.StringVar()
        self.content_root_var = tk.StringVar()
        self.watermark_path_var = tk.StringVar()
        self.image_api_key_var = tk.StringVar(value=os.environ.get("BREAKOUT_IMAGE_API_KEY", os.environ.get("BREAKOUT_API_KEY", "")))
        self.text_api_key_var = tk.StringVar(value=os.environ.get("BREAKOUT_TEXT_API_KEY", os.environ.get("BREAKOUT_API_KEY", "")))
        self.base_url_var = tk.StringVar(value=DEFAULT_BASE_URL)
        self.image_model_var = tk.StringVar(value=DEFAULT_IMAGE_MODEL)
        self.text_model_var = tk.StringVar(value=DEFAULT_TEXT_MODEL)
        self.quality_var = tk.StringVar(value="medium")
        self.max_folders_var = tk.StringVar(value="0")
        self.max_workers_var = tk.StringVar(value="3")
        self.convert_originals_var = tk.BooleanVar(value=False)
        self.generate_copy_var = tk.BooleanVar(value=True)
        self.export_excel_var = tk.BooleanVar(value=True)
        self.template_name_var = tk.StringVar(value=DEFAULT_TEMPLATE_NAME)
        self.prompt_templates: dict[str, dict[str, str]] = {
            DEFAULT_TEMPLATE_NAME: {
                "image_prompt_template": DEFAULT_IMAGE_PROMPT,
                "title_prompt_template": DEFAULT_TITLE_PROMPT,
                "description_prompt_template": DEFAULT_DESCRIPTION_PROMPT,
            }
        }

        self._build_ui()
        self._load_config()
        self.after(250, self._flush_logs)

    def _build_ui(self) -> None:
        frame = ttk.Frame(self, padding=16)
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(1, weight=1)

        row = 0
        self._path_row(frame, row, "源目录", self.source_root_var)
        row += 1
        self._path_row(frame, row, "AI 竖版图输出目录", self.generated_root_var)
        row += 1
        self._path_row(frame, row, "3:4 输出目录", self.portrait_root_var)
        row += 1
        self._path_row(frame, row, "文案输出目录", self.content_root_var)
        row += 1
        self._path_row(frame, row, "水印图片", self.watermark_path_var, select_file=True)
        row += 1

        ttk.Label(frame, text="接口地址").grid(row=row, column=0, sticky="w", pady=(8, 6))
        ttk.Entry(frame, textvariable=self.base_url_var).grid(row=row, column=1, columnspan=2, sticky="ew", pady=(8, 6))
        row += 1

        ttk.Label(frame, text="图片 API Key").grid(row=row, column=0, sticky="w", pady=6)
        ttk.Entry(frame, textvariable=self.image_api_key_var, show="*").grid(row=row, column=1, columnspan=2, sticky="ew", pady=6)
        row += 1

        ttk.Label(frame, text="文本 API Key").grid(row=row, column=0, sticky="w", pady=6)
        ttk.Entry(frame, textvariable=self.text_api_key_var, show="*").grid(row=row, column=1, columnspan=2, sticky="ew", pady=6)
        row += 1

        options = ttk.Frame(frame)
        options.grid(row=row, column=0, columnspan=3, sticky="ew", pady=(6, 10))
        for col in range(6):
            options.columnconfigure(col, weight=1 if col % 2 else 0)

        ttk.Label(options, text="图片模型").grid(row=0, column=0, sticky="w")
        ttk.Entry(options, textvariable=self.image_model_var, width=16).grid(row=0, column=1, sticky="ew", padx=(6, 18))
        ttk.Label(options, text="文案模型").grid(row=0, column=2, sticky="w")
        ttk.Entry(options, textvariable=self.text_model_var, width=16).grid(row=0, column=3, sticky="ew", padx=(6, 18))
        ttk.Label(options, text="质量").grid(row=0, column=4, sticky="w")
        ttk.Combobox(
            options,
            textvariable=self.quality_var,
            values=("low", "medium", "high"),
            state="readonly",
            width=10,
        ).grid(row=0, column=5, sticky="ew", padx=(6, 0))

        row += 1
        options2 = ttk.Frame(frame)
        options2.grid(row=row, column=0, columnspan=3, sticky="ew", pady=(0, 10))
        ttk.Label(options2, text="最多处理文件夹").grid(row=0, column=0, sticky="w")
        ttk.Entry(options2, textvariable=self.max_folders_var, width=10).grid(row=0, column=1, sticky="w", padx=(6, 18))
        ttk.Label(options2, text="图片并发数").grid(row=0, column=2, sticky="w")
        ttk.Entry(options2, textvariable=self.max_workers_var, width=8).grid(row=0, column=3, sticky="w", padx=(6, 18))
        ttk.Checkbutton(options2, text="同时生成标题、描述、卖点文案", variable=self.generate_copy_var).grid(
            row=0, column=4, sticky="w", padx=(0, 18)
        )
        ttk.Checkbutton(options2, text="导出 Excel 汇总表", variable=self.export_excel_var).grid(
            row=0, column=5, sticky="w"
        )

        row += 1
        template_bar = ttk.Frame(frame)
        template_bar.grid(row=row, column=0, columnspan=3, sticky="ew", pady=(4, 10))
        ttk.Label(template_bar, text="提示词模板").pack(side="left")
        self.template_combobox = ttk.Combobox(
            template_bar,
            textvariable=self.template_name_var,
            state="normal",
            width=28,
        )
        self.template_combobox.pack(side="left", padx=(8, 8))
        ttk.Button(template_bar, text="加载模板", command=self._load_selected_template).pack(side="left")
        ttk.Button(template_bar, text="保存为模板", command=self._save_prompt_template).pack(side="left", padx=8)
        ttk.Button(template_bar, text="删除模板", command=self._delete_prompt_template).pack(side="left")
        self.template_combobox.bind("<<ComboboxSelected>>", lambda _event: self._load_selected_template())
        self._refresh_template_options()

        row += 1
        ttk.Label(frame, text="图片提示词模板").grid(row=row, column=0, sticky="nw", pady=(8, 6))
        self.image_prompt_text = tk.Text(frame, height=6, wrap="word")
        self.image_prompt_text.grid(row=row, column=1, columnspan=2, sticky="nsew", pady=(8, 6))
        self.image_prompt_text.insert("1.0", DEFAULT_IMAGE_PROMPT)

        row += 1
        ttk.Label(frame, text="标题提示词模板").grid(row=row, column=0, sticky="nw", pady=(8, 6))
        self.title_prompt_text = tk.Text(frame, height=5, wrap="word")
        self.title_prompt_text.grid(row=row, column=1, columnspan=2, sticky="nsew", pady=(8, 6))
        self.title_prompt_text.insert("1.0", DEFAULT_TITLE_PROMPT)

        row += 1
        ttk.Label(frame, text="卖点文案提示词模板").grid(row=row, column=0, sticky="nw", pady=(8, 6))
        self.description_prompt_text = tk.Text(frame, height=5, wrap="word")
        self.description_prompt_text.grid(row=row, column=1, columnspan=2, sticky="nsew", pady=(8, 6))
        self.description_prompt_text.insert("1.0", DEFAULT_DESCRIPTION_PROMPT)

        row += 1
        button_bar = ttk.Frame(frame)
        button_bar.grid(row=row, column=0, columnspan=3, sticky="ew", pady=(8, 0))
        ttk.Button(button_bar, text="保存配置", command=self._save_config).pack(side="left")
        ttk.Button(button_bar, text="测试 API", command=self._test_api_connection).pack(side="left", padx=8)
        ttk.Button(button_bar, text="开始处理", command=self._start_job).pack(side="left", padx=8)
        ttk.Button(button_bar, text="取消任务", command=self._cancel_job).pack(side="left")

        row += 1
        ttk.Label(frame, text="运行日志").grid(row=row, column=0, sticky="nw", pady=(14, 6))
        self.log_text = tk.Text(frame, height=16, wrap="word", state="disabled")
        self.log_text.grid(row=row, column=1, columnspan=2, sticky="nsew", pady=(14, 6))
        frame.rowconfigure(row, weight=1)

    def _path_row(
        self,
        parent: ttk.Frame,
        row: int,
        label: str,
        variable: tk.StringVar,
        select_file: bool = False,
    ) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=6)
        ttk.Entry(parent, textvariable=variable).grid(row=row, column=1, sticky="ew", pady=6)
        chooser = (lambda var=variable: self._choose_file(var)) if select_file else (lambda var=variable: self._choose_folder(var))
        ttk.Button(parent, text="选择", command=chooser).grid(
            row=row, column=2, sticky="e", padx=(10, 0), pady=6
        )

    def _choose_folder(self, variable: tk.StringVar) -> None:
        selected = filedialog.askdirectory()
        if selected:
            variable.set(selected)

    def _choose_file(self, variable: tk.StringVar) -> None:
        selected = filedialog.askopenfilename(
            filetypes=[
                ("Image Files", "*.png;*.jpg;*.jpeg;*.webp;*.bmp"),
                ("All Files", "*.*"),
            ]
        )
        if selected:
            variable.set(selected)

    def _append_log(self, message: str) -> None:
        self.log_queue.put(f"{now_stamp()}  {message}")

    def _current_prompt_bundle(self) -> dict[str, str]:
        return {
            "image_prompt_template": self.image_prompt_text.get("1.0", "end").strip(),
            "title_prompt_template": self.title_prompt_text.get("1.0", "end").strip(),
            "description_prompt_template": self.description_prompt_text.get("1.0", "end").strip(),
        }

    def _apply_prompt_bundle(self, bundle: dict[str, str]) -> None:
        self.image_prompt_text.delete("1.0", "end")
        self.image_prompt_text.insert("1.0", bundle.get("image_prompt_template", DEFAULT_IMAGE_PROMPT))
        self.title_prompt_text.delete("1.0", "end")
        self.title_prompt_text.insert("1.0", bundle.get("title_prompt_template", DEFAULT_TITLE_PROMPT))
        self.description_prompt_text.delete("1.0", "end")
        self.description_prompt_text.insert(
            "1.0",
            bundle.get("description_prompt_template", DEFAULT_DESCRIPTION_PROMPT),
        )

    def _refresh_template_options(self) -> None:
        names = sorted(self.prompt_templates.keys())
        self.template_combobox["values"] = names
        if not self.template_name_var.get().strip():
            self.template_name_var.set(DEFAULT_TEMPLATE_NAME)

    def _load_selected_template(self) -> None:
        name = self.template_name_var.get().strip()
        if not name:
            messagebox.showerror("模板错误", "请先输入或选择模板名。")
            return
        bundle = self.prompt_templates.get(name)
        if not bundle:
            messagebox.showerror("模板不存在", f"未找到模板：{name}")
            return
        self._apply_prompt_bundle(bundle)
        self._append_log(f"已加载提示词模板：{name}")

    def _save_prompt_template(self) -> None:
        name = self.template_name_var.get().strip()
        if not name:
            messagebox.showerror("模板错误", "请先输入模板名。")
            return
        self.prompt_templates[name] = self._current_prompt_bundle()
        self._refresh_template_options()
        self.template_name_var.set(name)
        self._save_config(silent=True)
        self._append_log(f"已保存提示词模板：{name}")

    def _delete_prompt_template(self) -> None:
        name = self.template_name_var.get().strip()
        if not name:
            messagebox.showerror("模板错误", "请先选择模板名。")
            return
        if name == DEFAULT_TEMPLATE_NAME:
            messagebox.showinfo("无法删除", "默认模板不能删除。")
            return
        if name not in self.prompt_templates:
            messagebox.showerror("模板不存在", f"未找到模板：{name}")
            return
        del self.prompt_templates[name]
        self.template_name_var.set(DEFAULT_TEMPLATE_NAME)
        self._refresh_template_options()
        self._apply_prompt_bundle(self.prompt_templates[DEFAULT_TEMPLATE_NAME])
        self._save_config(silent=True)
        self._append_log(f"已删除提示词模板：{name}")

    def _flush_logs(self) -> None:
        while not self.log_queue.empty():
            line = self.log_queue.get_nowait()
            self.log_text.configure(state="normal")
            self.log_text.insert("end", line + "\n")
            self.log_text.see("end")
            self.log_text.configure(state="disabled")
        self.after(250, self._flush_logs)

    def _collect_config(self) -> JobConfig:
        source_root = Path(self.source_root_var.get()).expanduser()
        generated_root = Path(self.generated_root_var.get()).expanduser()
        portrait_root = Path(self.portrait_root_var.get()).expanduser()
        content_root = Path(self.content_root_var.get()).expanduser()
        watermark_text = self.watermark_path_var.get().strip()
        watermark_path = Path(watermark_text).expanduser() if watermark_text else None

        if not source_root.is_dir():
            raise RuntimeError("请先选择有效的源目录。")
        if not self.generated_root_var.get().strip():
            raise RuntimeError("请先选择 AI 竖版图输出目录。")
        if not self.portrait_root_var.get().strip():
            raise RuntimeError("请先选择 3:4 输出目录。")
        if self.generate_copy_var.get() and not self.content_root_var.get().strip():
            raise RuntimeError("请先选择文案输出目录。")
        if not self.image_api_key_var.get().strip():
            raise RuntimeError("请填写图片 API Key。")
        if self.generate_copy_var.get() and not self.text_api_key_var.get().strip():
            raise RuntimeError("请填写文本 API Key。")
        if not self.base_url_var.get().strip():
            raise RuntimeError("请填写接口地址。")
        if watermark_path and not watermark_path.is_file():
            raise RuntimeError("水印图片不存在，请重新选择。")

        try:
            max_folders = max(0, int(self.max_folders_var.get().strip() or "0"))
            max_workers = max(1, int(self.max_workers_var.get().strip() or "1"))
        except ValueError as exc:
            raise RuntimeError("最多处理文件夹和图片并发数必须是整数。") from exc

        return JobConfig(
            source_root=source_root,
            generated_root=generated_root,
            portrait_root=portrait_root,
            content_root=content_root,
            image_api_key=self.image_api_key_var.get().strip(),
            text_api_key=(self.text_api_key_var.get().strip() or self.image_api_key_var.get().strip()),
            base_url=self.base_url_var.get().strip().rstrip("/"),
            image_model=self.image_model_var.get().strip() or DEFAULT_IMAGE_MODEL,
            text_model=self.text_model_var.get().strip() or DEFAULT_TEXT_MODEL,
            image_prompt_template=self.image_prompt_text.get("1.0", "end").strip(),
            title_prompt_template=self.title_prompt_text.get("1.0", "end").strip(),
            description_prompt_template=self.description_prompt_text.get("1.0", "end").strip(),
            watermark_path=watermark_path,
            quality=self.quality_var.get().strip(),
            max_folders=max_folders,
            max_workers=max_workers,
            convert_originals=self.convert_originals_var.get(),
            generate_copy=self.generate_copy_var.get(),
            export_excel=self.export_excel_var.get(),
        )

    def _start_job(self) -> None:
        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showinfo("任务运行中", "当前已有任务在运行。")
            return

        try:
            config = self._collect_config()
        except RuntimeError as exc:
            messagebox.showerror("配置错误", str(exc))
            return

        self._save_config(silent=True)
        self.worker = BatchWorker(config, self._append_log)
        self.worker_thread = threading.Thread(target=self._run_worker, daemon=True)
        self.worker_thread.start()
        self._append_log("任务已启动。")

    def _run_worker(self) -> None:
        try:
            assert self.worker is not None
            self.worker.run()
        except Exception as exc:  # noqa: BLE001
            self._append_log(f"任务失败: {humanize_api_error(str(exc))}")

    def _cancel_job(self) -> None:
        if self.worker:
            self.worker.cancel()
            self._append_log("已请求取消任务。")

    def _test_api_connection(self) -> None:
        try:
            config = self._collect_config()
        except RuntimeError as exc:
            messagebox.showerror("配置错误", str(exc))
            return

        self._append_log("开始测试接口连通性...")
        threading.Thread(target=self._run_api_test, args=(config,), daemon=True).start()

    def _run_api_test(self, config: JobConfig) -> None:
        try:
            client = GatewayClient(config.image_api_key, config.base_url, timeout_seconds=90)
            data = client.list_models()
            model_count = len(data.get("data", [])) if isinstance(data, dict) else 0
            self._append_log(f"接口测试成功：已获取模型列表，数量 {model_count}")
            self.after(0, lambda: messagebox.showinfo("测试成功", f"接口连接正常。\n已获取模型数量：{model_count}"))
        except Exception as exc:  # noqa: BLE001
            friendly = humanize_api_error(str(exc))
            self._append_log(f"接口测试失败: {friendly}")
            self.after(0, lambda: messagebox.showerror("测试失败", friendly))

    def _save_config(self, silent: bool = False) -> None:
        payload = {
            "source_root": self.source_root_var.get(),
            "generated_root": self.generated_root_var.get(),
            "portrait_root": self.portrait_root_var.get(),
            "content_root": self.content_root_var.get(),
            "watermark_path": self.watermark_path_var.get(),
            "image_api_key": self.image_api_key_var.get(),
            "text_api_key": self.text_api_key_var.get(),
            "base_url": self.base_url_var.get(),
            "image_model": self.image_model_var.get(),
            "text_model": self.text_model_var.get(),
            "quality": self.quality_var.get(),
            "max_folders": self.max_folders_var.get(),
            "max_workers": self.max_workers_var.get(),
            "convert_originals": self.convert_originals_var.get(),
            "generate_copy": self.generate_copy_var.get(),
            "export_excel": self.export_excel_var.get(),
            "selected_template_name": self.template_name_var.get().strip() or DEFAULT_TEMPLATE_NAME,
            "prompt_templates": self.prompt_templates,
            **self._current_prompt_bundle(),
        }
        CONFIG_PATH.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        if not silent:
            self._append_log(f"配置已保存到 {CONFIG_PATH}")

    def _load_config(self) -> None:
        if not CONFIG_PATH.exists():
            return
        try:
            data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            self._append_log("配置文件读取失败，已忽略旧配置。")
            return

        self.source_root_var.set(data.get("source_root", ""))
        self.generated_root_var.set(data.get("generated_root", ""))
        self.portrait_root_var.set(data.get("portrait_root", ""))
        self.content_root_var.set(data.get("content_root", ""))
        self.watermark_path_var.set(data.get("watermark_path", ""))
        self.image_api_key_var.set(
            data.get("image_api_key", os.environ.get("BREAKOUT_IMAGE_API_KEY", os.environ.get("BREAKOUT_API_KEY", "")))
        )
        self.text_api_key_var.set(
            data.get("text_api_key", os.environ.get("BREAKOUT_TEXT_API_KEY", os.environ.get("BREAKOUT_API_KEY", "")))
        )
        self.base_url_var.set(data.get("base_url", DEFAULT_BASE_URL))

        saved_image_model = (data.get("image_model") or data.get("model") or "").strip()
        if saved_image_model in {"", "gpt-image-1.5", "gpt-image-1"}:
            saved_image_model = DEFAULT_IMAGE_MODEL
        self.image_model_var.set(saved_image_model)

        self.text_model_var.set((data.get("text_model") or DEFAULT_TEXT_MODEL).strip())
        self.quality_var.set(data.get("quality", "medium"))
        self.max_folders_var.set(str(data.get("max_folders", "0")))
        self.max_workers_var.set(str(data.get("max_workers", "3")))
        self.convert_originals_var.set(bool(data.get("convert_originals", False)))
        self.generate_copy_var.set(bool(data.get("generate_copy", True)))
        self.export_excel_var.set(bool(data.get("export_excel", True)))
        saved_templates = data.get("prompt_templates")
        if isinstance(saved_templates, dict):
            cleaned_templates: dict[str, dict[str, str]] = {}
            for name, bundle in saved_templates.items():
                if isinstance(name, str) and isinstance(bundle, dict):
                    cleaned_templates[name] = {
                        "image_prompt_template": str(bundle.get("image_prompt_template", DEFAULT_IMAGE_PROMPT)),
                        "title_prompt_template": str(bundle.get("title_prompt_template", DEFAULT_TITLE_PROMPT)),
                        "description_prompt_template": str(
                            bundle.get("description_prompt_template", DEFAULT_DESCRIPTION_PROMPT)
                        ),
                    }
            if cleaned_templates:
                self.prompt_templates = cleaned_templates

        if DEFAULT_TEMPLATE_NAME not in self.prompt_templates:
            self.prompt_templates[DEFAULT_TEMPLATE_NAME] = {
                "image_prompt_template": DEFAULT_IMAGE_PROMPT,
                "title_prompt_template": DEFAULT_TITLE_PROMPT,
                "description_prompt_template": DEFAULT_DESCRIPTION_PROMPT,
            }

        legacy_bundle = {
            "image_prompt_template": data.get("image_prompt_template", DEFAULT_IMAGE_PROMPT),
            "title_prompt_template": data.get("title_prompt_template", DEFAULT_TITLE_PROMPT),
            "description_prompt_template": data.get("description_prompt_template", DEFAULT_DESCRIPTION_PROMPT),
        }
        selected_name = str(data.get("selected_template_name", DEFAULT_TEMPLATE_NAME)).strip() or DEFAULT_TEMPLATE_NAME
        if selected_name not in self.prompt_templates:
            self.prompt_templates[selected_name] = {
                "image_prompt_template": str(legacy_bundle["image_prompt_template"]),
                "title_prompt_template": str(legacy_bundle["title_prompt_template"]),
                "description_prompt_template": str(legacy_bundle["description_prompt_template"]),
            }

        self.template_name_var.set(selected_name)
        self._refresh_template_options()
        self._apply_prompt_bundle(self.prompt_templates[selected_name])


def main() -> None:
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
