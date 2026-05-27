"""Minimal OpenAI-compatible image API client (no GUI / batch_upload dependencies)."""

from __future__ import annotations

import base64
import json
import mimetypes
import time
import uuid
from pathlib import Path
from typing import Iterable
from urllib import error, request

DEFAULT_DOWNLOAD_RETRIES = 3
BROWSER_USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)


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
    return b"".join(chunks), f"multipart/form-data; boundary={boundary}"


class ImageApiClient:
    def __init__(self, api_key: str, base_url: str, image_model: str, timeout_seconds: int = 180):
        self.api_key = api_key.strip()
        self.base_url = base_url.strip().rstrip("/")
        self.image_model = image_model.strip()
        self.timeout_seconds = timeout_seconds

    def _headers(self, content_type: str | None = None) -> dict[str, str]:
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "User-Agent": BROWSER_USER_AGENT,
            "Accept": "application/json",
        }
        if content_type:
            headers["Content-Type"] = content_type
        return headers

    def _request_json(self, method: str, path: str, payload: bytes, content_type: str) -> dict:
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

    def generate_image(
        self, prompt: str, output_size: str, quality: str, reference_images: Iterable[Path] | None = None
    ) -> bytes:
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
            body = json.dumps(
                {"model": self.image_model, "prompt": prompt, "size": output_size, "quality": quality}
            ).encode("utf-8")
            data = self._request_json("POST", "/images/generations", body, "application/json")
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
            fallback_prompt = f"{prompt}\n补充要求：严格保持参考图印花与配色，生成电商场景展示图。"
            note = f"参考图编辑失败，已回退文生图: {exc}"
            return self.generate_image(fallback_prompt, output_size, quality, reference_images=[]), note

    def _extract_image_bytes(self, data: dict) -> bytes:
        items = data.get("data")
        if not isinstance(items, list) or not items:
            raise RuntimeError(f"图片接口返回异常: {json.dumps(data, ensure_ascii=False)[:500]}")
        first = items[0]
        if not isinstance(first, dict):
            raise RuntimeError("图片接口返回格式错误")
        if first.get("b64_json"):
            return base64.b64decode(first["b64_json"])
        image_url = first.get("url")
        if image_url:
            return self._download_image(str(image_url))
        raise RuntimeError("图片接口未返回 b64_json 或 url")

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
