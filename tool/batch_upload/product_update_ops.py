"""Update listed Ozon products by offer_id (title, images, video, rich JSON)."""

from __future__ import annotations

import copy
import json
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

from batch_upload.core import (
    AliyunOssClient,
    OzonSellerClient,
    _extract_task_id,
    build_import_item,
    build_oss_folder,
    build_oss_object_key,
    extract_listed_image_urls,
    list_sku_images,
    read_excel_rows,
    write_results,
    write_status_to_source_excel,
)
from batch_upload.video_ops import (
    apply_video_to_import_item,
    parse_video_links_text,
    read_video_links_from_excel,
    resolve_sku_video_links,
    strip_misplaced_video_from_item,
)

Logger = Callable[[str], None]

RESULT_FILE_NAME = "product_update_results.xlsx"


@dataclass
class ListedProductUpdateConfig:
    portrait_root: Path
    excel_path: Path
    ozon_client_id: str
    ozon_api_key: str
    oss_access_key_id: str
    oss_access_key_secret: str
    oss_bucket: str
    oss_endpoint: str
    oss_public_domain: str
    max_items: int = 0
    update_title: bool = True
    update_description: bool = True
    update_images: bool = True
    update_video: bool = False
    update_rich_json: bool = True
    template_video_links: list[str] | None = None
    template_product: dict | None = None


class ListedProductUpdateWorker:
    """Update existing products via POST /v3/product/import (Ozon 创建或更新商品)."""

    def __init__(self, config: ListedProductUpdateConfig, logger: Logger):
        self.config = config
        self.logger = logger
        self._cancelled = False
        self.results: list[dict] = []

    def cancel(self) -> None:
        self._cancelled = True

    def run(self) -> Path:
        if not self.config.excel_path.is_file():
            raise RuntimeError("Excel 文件不存在。")
        if self.config.update_images and not self.config.portrait_root.is_dir():
            raise RuntimeError("3:4 输出目录不存在。")

        excel_rows = read_excel_rows(self.config.excel_path)
        rows = list(excel_rows.items())
        if self.config.max_items > 0:
            rows = rows[: self.config.max_items]

        ozon_client = OzonSellerClient(self.config.ozon_client_id, self.config.ozon_api_key)
        oss_client: AliyunOssClient | None = None
        if self.config.update_images:
            oss_client = AliyunOssClient(
                self.config.oss_access_key_id,
                self.config.oss_access_key_secret,
                self.config.oss_bucket,
                self.config.oss_endpoint,
                self.config.oss_public_domain,
            )

        per_offer_video_links = read_video_links_from_excel(self.config.excel_path)
        default_video_links = parse_video_links_text("\n".join(self.config.template_video_links or []))
        video_template = self.config.template_product

        self.logger(f"按货号更新: Excel 货号 {len(rows)} 个")
        flags = []
        if self.config.update_title:
            flags.append("标题")
        if self.config.update_description:
            flags.append("简介")
        if self.config.update_images:
            flags.append("图片")
        if self.config.update_video:
            flags.append("视频")
        if self.config.update_rich_json:
            flags.append("JSON富内容")
        self.logger(f"更新项: {', '.join(flags) or '无'}")

        for index, (offer_id, row) in enumerate(rows, start=1):
            if self._cancelled:
                self.logger("按货号更新: 任务已取消")
                break
            self.logger(f"[{index}/{len(rows)}] 更新货号 {offer_id}")
            try:
                task_id = self._update_one(
                    ozon_client,
                    oss_client,
                    offer_id,
                    row,
                    default_video_links=default_video_links,
                    per_offer_video_links=per_offer_video_links,
                    video_template=video_template,
                )
                self._add_result(offer_id, row.get("title", ""), "已提交", str(task_id or ""), "")
                self.logger(f"  已提交 Ozon，task_id: {task_id or '[未返回]'}")
            except Exception as exc:  # noqa: BLE001
                self._add_result(offer_id, row.get("title", ""), "失败", "", str(exc))
                self.logger(f"  失败: {exc}")

        result_path = self.config.portrait_root / RESULT_FILE_NAME
        write_results(result_path, self.results)
        write_status_to_source_excel(self.config.excel_path, self.results)
        self.logger(f"按货号更新: 结果已写回 Excel，汇总表 {result_path}")
        return result_path

    def _update_one(
        self,
        ozon_client: OzonSellerClient,
        oss_client: AliyunOssClient | None,
        offer_id: str,
        row: dict[str, str],
        *,
        default_video_links: list[str],
        per_offer_video_links: dict[str, list[str]],
        video_template: dict | None,
    ) -> int | str | None:
        if not ozon_client.product_exists(offer_id):
            raise RuntimeError("Ozon 上不存在该货号，请先上架或检查货号是否正确。")

        listed_product = ozon_client.get_template_product(offer_id)
        title = row.get("title", "").strip() if self.config.update_title else str(listed_product.get("name") or "").strip()
        description = (
            row.get("description", "").strip()
            if self.config.update_description
            else _listed_description(listed_product)
        )
        if self.config.update_title and not title:
            raise RuntimeError("标题为空")
        if self.config.update_description and not description:
            raise RuntimeError("简介为空")

        image_urls: list[str] = []
        if self.config.update_images:
            assert oss_client is not None
            sku_folder = self.config.portrait_root / offer_id
            if not sku_folder.is_dir():
                raise RuntimeError("未找到同名图片目录")
            images = list_sku_images(sku_folder)
            if not images:
                raise RuntimeError("图片目录为空")
            shop_id = self.config.ozon_client_id
            oss_folder = build_oss_folder(shop_id, offer_id)
            self.logger(f"  上传 {len(images)} 张图片到 OSS: {oss_folder}/")
            for image_index, image_path in enumerate(images, start=1):
                object_key = build_oss_object_key(shop_id, offer_id, image_path, image_index)
                image_url = oss_client.upload_file(image_path, object_key)
                image_urls.append(image_url)
                self.logger(f"    [{image_index}/{len(images)}] {image_url}")
        else:
            image_urls = extract_listed_image_urls(listed_product)
            if not image_urls:
                raise RuntimeError("未开启图片更新且当前商品无可用图片 URL")

        item = build_import_item(listed_product, offer_id, title, description, image_urls)

        rich_json = row.get("rich_json", "").strip()
        if self.config.update_rich_json and rich_json:
            from batch_upload.core import apply_rich_json_to_item

            if apply_rich_json_to_item(item, rich_json, image_urls=image_urls):
                self.logger("  已写入 JSON 富内容")
            else:
                self.logger("  警告: 未能匹配富内容字段，已跳过 JSON 富内容更新")

        if self.config.update_video:
            video_links = resolve_sku_video_links(
                offer_id,
                default_links=default_video_links,
                per_offer_links=per_offer_video_links,
                template_product=video_template or listed_product,
            )
            if not video_links:
                raise RuntimeError("未找到视频链接（模板区或 Excel Ozon.Video 页签）")
            apply_video_to_import_item(item, video_links, video_template or listed_product)
            strip_misplaced_video_from_item(item)
            self.logger(f"  已写入视频链接 {len(video_links)} 条")

        response = ozon_client.import_products([item])
        return _extract_task_id(response)

    def _add_result(
        self,
        sku: str,
        title: str,
        status: str,
        task_id: str,
        message: str,
    ) -> None:
        self.results.append(
            {
                "sku": sku,
                "title": title,
                "image_count": 0,
                "status": status,
                "uploaded_sku": sku if status == "已提交" else "",
                "task_id": task_id,
                "oss_folder": "",
                "error": message,
            }
        )


def _listed_description(product: dict) -> str:
    description = str(product.get("description") or "").strip()
    if description:
        return description
    attributes = product.get("attributes")
    if not isinstance(attributes, list):
        return ""
    for attribute in attributes:
        if not isinstance(attribute, dict):
            continue
        if int(attribute.get("id") or 0) == 4191:
            values = attribute.get("values")
            if isinstance(values, list) and values:
                first = values[0]
                if isinstance(first, dict):
                    return str(first.get("value") or "").strip()
                return str(first).strip()
    return ""
