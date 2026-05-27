"""Image generation providers: OpenAI-compatible (Wenwen) and Apimart (async Gemini)."""

from __future__ import annotations

import base64
import json
import mimetypes
import time
from pathlib import Path
from typing import Iterable
from urllib import error, request

from image_api_client import ImageApiClient

DEFAULT_DOWNLOAD_RETRIES = 3
APIMART_BASE_URL = "https://api.apimart.ai/v1"
WENWEN_BASE_URL = "https://breakout.wenwen-ai.com/v1"
XIAOQIAN_BASE_URL = "https://xiaoqian.art/v1"
BROWSER_USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)

SIZE_TO_ASPECT: dict[str, str] = {
    "1024x1024": "1:1",
    "1024x1536": "3:4",
    "1536x1024": "4:3",
    "1536x864": "16:9",
}

IMAGE_PROVIDER_PROFILES: dict[str, dict] = {
    "wenwen": {
        "label": "Wenwen（OpenAI 兼容）",
        "base_url": WENWEN_BASE_URL,
        "models": ["gpt-image-2"],
        "default_model": "gpt-image-2",
        "supports_edits": True,
        "async_tasks": False,
    },
    "apimart_gemini": {
        "label": "Apimart Gemini 2.5 Flash",
        "base_url": APIMART_BASE_URL,
        "models": [
            "gemini-2.5-flash-image-preview",
            "gemini-2.5-flash-image-preview-official",
        ],
        "default_model": "gemini-2.5-flash-image-preview",
        "supports_edits": False,
        "async_tasks": True,
    },
    "apimart_gpt_image": {
        "label": "Apimart GPT-Image-2",
        "base_url": APIMART_BASE_URL,
        "models": [
            "gpt-image-2",
            "gpt-image-2-official",
        ],
        "default_model": "gpt-image-2",
        "supports_edits": False,
        "async_tasks": True,
    },
    "xiaoqian": {
        "label": "小钱.art（OpenAI 兼容）",
        "base_url": XIAOQIAN_BASE_URL,
        "models": [
            "gpt-image-2",
            "gpt-image-1.5",
            "gpt-image-1",
        ],
        "default_model": "gpt-image-2",
        "supports_edits": True,
        "async_tasks": False,
    },
}

DEFAULT_IMAGE_PROVIDER = "wenwen"


def list_image_provider_ids() -> list[str]:
    return list(IMAGE_PROVIDER_PROFILES.keys())


def provider_label(provider_id: str) -> str:
    profile = IMAGE_PROVIDER_PROFILES.get(provider_id) or IMAGE_PROVIDER_PROFILES[DEFAULT_IMAGE_PROVIDER]
    return str(profile.get("label") or provider_id)


def provider_models(provider_id: str) -> list[str]:
    profile = IMAGE_PROVIDER_PROFILES.get(provider_id) or IMAGE_PROVIDER_PROFILES[DEFAULT_IMAGE_PROVIDER]
    models = profile.get("models")
    if isinstance(models, list) and models:
        return [str(item) for item in models]
    return [str(profile.get("default_model") or "")]


def provider_default_model(provider_id: str) -> str:
    profile = IMAGE_PROVIDER_PROFILES.get(provider_id) or IMAGE_PROVIDER_PROFILES[DEFAULT_IMAGE_PROVIDER]
    return str(profile.get("default_model") or provider_models(provider_id)[0])


def provider_default_base_url(provider_id: str) -> str:
    profile = IMAGE_PROVIDER_PROFILES.get(provider_id) or IMAGE_PROVIDER_PROFILES[DEFAULT_IMAGE_PROVIDER]
    return str(profile.get("base_url") or WENWEN_BASE_URL).rstrip("/")


def provider_uses_async_tasks(provider_id: str) -> bool:
    profile = IMAGE_PROVIDER_PROFILES.get(provider_id) or IMAGE_PROVIDER_PROFILES[DEFAULT_IMAGE_PROVIDER]
    return bool(profile.get("async_tasks"))


def normalize_output_size_for_provider(output_size: str, provider_id: str) -> str:
    value = (output_size or "").strip().replace("：", ":").replace(" ", "").lower()
    if provider_uses_async_tasks(provider_id):
        if value in SIZE_TO_ASPECT:
            return SIZE_TO_ASPECT[value]
        if ":" in value or "x" in value:
            return value
        return "3:4"
    return output_size


def image_path_to_data_url(path: Path) -> str:
    mime = mimetypes.guess_type(path.name)[0] or "image/png"
    encoded = base64.b64encode(path.read_bytes()).decode("ascii")
    return f"data:{mime};base64,{encoded}"


class ApimartImageClient:
    """Apimart async image API (Gemini 2.5 Flash and compatible models)."""

    def __init__(
        self,
        api_key: str,
        base_url: str = APIMART_BASE_URL,
        image_model: str = "gemini-2.5-flash-image-preview",
        provider_id: str = "apimart_gemini",
        timeout_seconds: int = 180,
        poll_interval_seconds: float = 3.0,
        initial_poll_delay_seconds: float = 10.0,
    ):
        self.api_key = api_key.strip()
        self.base_url = base_url.strip().rstrip("/")
        self.image_model = image_model.strip()
        self.provider_id = provider_id if provider_id in IMAGE_PROVIDER_PROFILES else "apimart_gemini"
        self.timeout_seconds = timeout_seconds
        self.poll_interval_seconds = poll_interval_seconds
        self.initial_poll_delay_seconds = initial_poll_delay_seconds

    def _headers(self, content_type: str | None = None) -> dict[str, str]:
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "User-Agent": BROWSER_USER_AGENT,
            "Accept": "application/json",
        }
        if content_type:
            headers["Content-Type"] = content_type
        return headers

    def _request_json(self, method: str, path: str, payload: bytes | None = None, content_type: str | None = None) -> dict:
        req = request.Request(
            f"{self.base_url}{path}",
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

    def _extract_task_id(self, data: dict) -> str:
        if data.get("error"):
            raise RuntimeError(json.dumps(data.get("error"), ensure_ascii=False))
        items = data.get("data")
        if not isinstance(items, list) or not items:
            raise RuntimeError(f"Apimart 未返回任务 ID: {json.dumps(data, ensure_ascii=False)[:500]}")
        first = items[0]
        if not isinstance(first, dict):
            raise RuntimeError("Apimart 返回格式错误")
        task_id = first.get("task_id") or first.get("id")
        if not task_id:
            raise RuntimeError(f"Apimart 未返回 task_id: {json.dumps(first, ensure_ascii=False)}")
        return str(task_id)

    def _poll_task(self, task_id: str) -> bytes:
        deadline = time.time() + self.timeout_seconds
        time.sleep(self.initial_poll_delay_seconds)
        while time.time() < deadline:
            data = self._request_json("GET", f"/tasks/{task_id}")
            payload = data.get("data") if isinstance(data.get("data"), dict) else data
            if not isinstance(payload, dict):
                raise RuntimeError(f"任务查询返回异常: {json.dumps(data, ensure_ascii=False)[:500]}")
            status = str(payload.get("status") or "").lower()
            if status in {"failed", "error", "cancelled"}:
                raise RuntimeError(f"图像任务失败: {json.dumps(payload, ensure_ascii=False)[:500]}")
            if status == "completed":
                return self._extract_result_image(payload)
            time.sleep(self.poll_interval_seconds)
        raise RuntimeError(f"图像任务超时（{self.timeout_seconds}s），task_id={task_id}")

    def _extract_result_image(self, payload: dict) -> bytes:
        result = payload.get("result")
        if not isinstance(result, dict):
            raise RuntimeError(f"任务完成但无 result: {json.dumps(payload, ensure_ascii=False)[:500]}")
        images = result.get("images")
        if not isinstance(images, list) or not images:
            raise RuntimeError(f"任务完成但无 images: {json.dumps(result, ensure_ascii=False)[:500]}")
        first = images[0]
        if not isinstance(first, dict):
            raise RuntimeError("images[0] 格式错误")
        url_field = first.get("url")
        if isinstance(url_field, list) and url_field:
            image_url = str(url_field[0])
        elif isinstance(url_field, str):
            image_url = url_field
        else:
            image_url = first.get("image_url") or first.get("download_url")
        if not image_url:
            raise RuntimeError(f"未找到图片 URL: {json.dumps(first, ensure_ascii=False)[:500]}")
        return self._download_image(str(image_url))

    def _download_image(self, image_url: str) -> bytes:
        last_error: Exception | None = None
        headers = {
            "User-Agent": BROWSER_USER_AGENT,
            "Accept": "image/*,*/*;q=0.8",
        }
        for _ in range(DEFAULT_DOWNLOAD_RETRIES):
            try:
                req = request.Request(image_url, headers=headers)
                with request.urlopen(req, timeout=self.timeout_seconds) as response:
                    return response.read()
            except (error.HTTPError, error.URLError) as exc:
                last_error = exc
            time.sleep(1.2)
        raise RuntimeError(f"下载图片失败: {last_error}")

    def generate_image(
        self,
        prompt: str,
        output_size: str,
        quality: str,
        reference_images: Iterable[Path] | None = None,
    ) -> bytes:
        del quality  # Apimart async APIs use resolution; quality is ignored here.
        size = normalize_output_size_for_provider(output_size, self.provider_id)
        body: dict = {
            "model": self.image_model,
            "prompt": prompt[:1000],
            "size": size,
            "n": 1,
            "resolution": "1k",
        }
        refs = list(reference_images or [])
        if refs:
            body["image_urls"] = [image_path_to_data_url(ref) for ref in refs[:14]]
        payload = json.dumps(body).encode("utf-8")
        submit = self._request_json("POST", "/images/generations", payload, "application/json")
        task_id = self._extract_task_id(submit)
        return self._poll_task(task_id)

    def generate_image_with_fallback(
        self,
        prompt: str,
        output_size: str,
        quality: str,
        reference_images: Iterable[Path] | None = None,
    ) -> tuple[bytes, str]:
        refs = list(reference_images or [])
        try:
            if refs:
                return self.generate_image(prompt, output_size, quality, reference_images=refs), "apimart-img2img"
            return self.generate_image(prompt, output_size, quality, reference_images=[]), "apimart-generation"
        except Exception as exc:  # noqa: BLE001
            if not refs:
                raise
            fallback_prompt = (
                f"{prompt}\n\n补充要求：严格保持参考图印花、配色与边框纹样不变，"
                "仅生成符合描述的电商场景展示图，白底干净，主体清晰。"
            )
            note = f"Apimart 图生图失败，已回退文生图: {exc}"
            return self.generate_image(fallback_prompt, output_size, quality, reference_images=[]), note


def create_image_client(
    provider_id: str,
    api_key: str,
    base_url: str,
    image_model: str,
    timeout_seconds: int = 180,
) -> ImageApiClient | ApimartImageClient:
    provider = provider_id if provider_id in IMAGE_PROVIDER_PROFILES else DEFAULT_IMAGE_PROVIDER
    resolved_base = (base_url or provider_default_base_url(provider)).strip().rstrip("/")
    resolved_model = (image_model or provider_default_model(provider)).strip()
    if provider_uses_async_tasks(provider):
        return ApimartImageClient(
            api_key=api_key,
            base_url=resolved_base,
            image_model=resolved_model,
            provider_id=provider,
            timeout_seconds=timeout_seconds,
        )
    return ImageApiClient(
        api_key=api_key,
        base_url=resolved_base,
        image_model=resolved_model,
        timeout_seconds=timeout_seconds,
    )


def migrate_provider_api_keys(data: dict) -> dict[str, str]:
    keys = data.get("provider_api_keys")
    migrated: dict[str, str] = {}
    if isinstance(keys, dict):
        migrated = {str(k): str(v) for k, v in keys.items() if v}
    legacy = str(data.get("image_api_key") or "").strip()
    if legacy and not migrated:
        saved_provider = str(data.get("image_provider") or DEFAULT_IMAGE_PROVIDER).strip()
        if saved_provider in IMAGE_PROVIDER_PROFILES:
            migrated[saved_provider] = legacy
        else:
            migrated["wenwen"] = legacy
    apimart_key = migrated.get("apimart_gemini") or migrated.get("apimart_gpt_image")
    if apimart_key:
        migrated.setdefault("apimart_gemini", apimart_key)
        migrated.setdefault("apimart_gpt_image", apimart_key)
    text_keys = data.get("text_provider_api_keys")
    if isinstance(text_keys, dict):
        xiaoqian_text_key = str(text_keys.get("xiaoqian") or "").strip()
        if xiaoqian_text_key:
            migrated.setdefault("xiaoqian", xiaoqian_text_key)
    return migrated
