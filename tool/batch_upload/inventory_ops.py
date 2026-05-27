"""Ozon inventory and barcode maintenance helpers."""
from __future__ import annotations

from dataclasses import dataclass
from typing import Callable, Iterable

from batch_upload.core import OzonSellerClient, _extract_items


Logger = Callable[[str], None]
BATCH_STOCK_LIMIT = 100
BATCH_BARCODE_LIMIT = 100


@dataclass
class WarehouseOption:
    warehouse_id: int
    name: str

    @property
    def label(self) -> str:
        return f"{self.name} ({self.warehouse_id})"


@dataclass
class ProductInventoryRow:
    product_id: int
    offer_id: str
    name: str
    barcodes: list[str]
    stock_summary: str

    @property
    def has_barcode(self) -> bool:
        return any(str(code).strip() for code in self.barcodes)


def _product_id_from_item(item: dict) -> int:
    for key in ("product_id", "id"):
        value = item.get(key)
        if value is None or value == "":
            continue
        return int(value)
    raise ValueError("商品缺少 product_id。")


def _offer_id_from_item(item: dict) -> str:
    return str(item.get("offer_id") or "").strip()


def _name_from_item(item: dict) -> str:
    return str(item.get("name") or item.get("title") or "").strip()


def _barcodes_from_item(item: dict) -> list[str]:
    raw = item.get("barcodes")
    if isinstance(raw, list):
        return [str(code).strip() for code in raw if str(code).strip()]
    barcode = item.get("barcode")
    if barcode:
        return [str(barcode).strip()]
    return []


def _stock_present(stock_entry: dict) -> int:
    for key in ("present", "free_stock", "stock", "available"):
        value = stock_entry.get(key)
        if value is None or value == "":
            continue
        return int(value)
    return 0


def _summarize_stocks(stock_entries: Iterable[dict]) -> str:
    parts: list[str] = []
    for entry in stock_entries:
        if not isinstance(entry, dict):
            continue
        warehouse_id = entry.get("warehouse_id") or entry.get("warehouseId")
        warehouse_name = str(entry.get("warehouse_name") or entry.get("warehouseName") or warehouse_id or "").strip()
        present = _stock_present(entry)
        label = warehouse_name or str(warehouse_id or "仓库")
        parts.append(f"{label}:{present}")
    return "；".join(parts) if parts else "全仓 0"


def _extract_product_list_meta(data: dict) -> tuple[list[dict], str]:
    result = data.get("result")
    if isinstance(result, dict):
        items = result.get("items") or result.get("products") or []
        last_id = str(result.get("last_id") or "")
        return [item for item in items if isinstance(item, dict)], last_id
    items = _extract_items(data)
    return items, ""


def _extract_stock_list_meta(data: dict) -> tuple[list[dict], str]:
    result = data.get("result")
    if isinstance(result, dict):
        items = result.get("items") or result.get("products") or []
        cursor = str(result.get("cursor") or "")
        return [item for item in items if isinstance(item, dict)], cursor
    items = _extract_items(data)
    cursor = str(data.get("cursor") or "")
    return items, cursor


def serialize_warehouse_options(options: list[WarehouseOption]) -> list[dict[str, int | str]]:
    return [{"warehouse_id": option.warehouse_id, "name": option.name} for option in options]


def deserialize_warehouse_options(items: list) -> list[WarehouseOption]:
    options: list[WarehouseOption] = []
    seen: set[int] = set()
    for item in items:
        if not isinstance(item, dict):
            continue
        warehouse_id = item.get("warehouse_id")
        if warehouse_id is None:
            continue
        wid = int(warehouse_id)
        if wid in seen:
            continue
        seen.add(wid)
        name = str(item.get("name") or f"仓库 {wid}").strip()
        options.append(WarehouseOption(warehouse_id=wid, name=name))
    options.sort(key=lambda option: option.name)
    return options


def load_warehouses(client: OzonSellerClient) -> list[WarehouseOption]:
    options: list[WarehouseOption] = []
    seen: set[int] = set()
    offset = 0
    while True:
        batch = client.list_warehouses(limit=200, offset=offset)
        if not batch:
            break
        for item in batch:
            warehouse_id = item.get("warehouse_id") or item.get("warehouseId") or item.get("id")
            if warehouse_id is None:
                continue
            wid = int(warehouse_id)
            if wid in seen:
                continue
            seen.add(wid)
            name = str(item.get("name") or item.get("warehouse_name") or f"仓库 {wid}").strip()
            options.append(WarehouseOption(warehouse_id=wid, name=name))
        if len(batch) < 200:
            break
        offset += len(batch)
    if not options:
        raise RuntimeError("未获取到可用仓库，请确认店铺已创建 FBS/rFBS 仓库。")
    options.sort(key=lambda option: option.name)
    return options


def iter_products_by_visibility(client: OzonSellerClient, visibility: str, *, logger: Logger | None = None) -> list[dict]:
    collected: list[dict] = []
    last_id = ""
    while True:
        response = client.list_products(last_id=last_id, limit=100, visibility=visibility)
        items, last_id = _extract_product_list_meta(response)
        collected.extend(items)
        if logger:
            logger(f"已拉取 {visibility} 商品 {len(collected)} 条...")
        if not last_id or not items:
            break
    return collected


def _build_row_from_list_item(item: dict, stock_entries: list[dict] | None = None) -> ProductInventoryRow:
    stocks = stock_entries if stock_entries is not None else item.get("stocks")
    if not isinstance(stocks, list):
        stocks = []
    stock_summary = _summarize_stocks(stocks)
    return ProductInventoryRow(
        product_id=_product_id_from_item(item),
        offer_id=_offer_id_from_item(item),
        name=_name_from_item(item),
        barcodes=_barcodes_from_item(item),
        stock_summary=stock_summary,
    )


def fetch_zero_stock_products(client: OzonSellerClient, *, logger: Logger | None = None) -> list[ProductInventoryRow]:
    items = iter_products_by_visibility(client, "EMPTY_STOCK", logger=logger)
    rows: list[ProductInventoryRow] = []
    for item in items:
        stocks = item.get("stocks")
        if isinstance(stocks, list) and stocks and any(_stock_present(entry) > 0 for entry in stocks if isinstance(entry, dict)):
            continue
        rows.append(_build_row_from_list_item(item, stocks if isinstance(stocks, list) else []))
    if logger:
        logger(f"零库存商品共 {len(rows)} 条。")
    return rows


def fetch_no_barcode_products(client: OzonSellerClient, *, logger: Logger | None = None) -> list[ProductInventoryRow]:
    items = iter_products_by_visibility(client, "EMPTY_BARCODE", logger=logger)
    rows: list[ProductInventoryRow] = []
    for item in items:
        row = _build_row_from_list_item(item)
        if row.has_barcode:
            continue
        rows.append(row)
    if logger:
        logger(f"无条形码商品共 {len(rows)} 条。")
    return rows


def update_product_stocks(
    client: OzonSellerClient,
    rows: list[ProductInventoryRow],
    *,
    warehouse_id: int,
    stock: int,
    logger: Logger | None = None,
) -> tuple[int, int]:
    if stock < 0:
        raise ValueError("库存数量不能为负数。")
    success = 0
    failed = 0
    for index in range(0, len(rows), BATCH_STOCK_LIMIT):
        chunk = rows[index : index + BATCH_STOCK_LIMIT]
        payload = [
            {
                "product_id": row.product_id,
                "offer_id": row.offer_id,
                "warehouse_id": int(warehouse_id),
                "stock": int(stock),
            }
            for row in chunk
        ]
        response = client.update_stocks(payload)
        results = response.get("result")
        if not isinstance(results, list):
            if logger:
                logger(f"库存更新批次 {index // BATCH_STOCK_LIMIT + 1} 已提交。")
            success += len(chunk)
            continue
        for result in results:
            if isinstance(result, dict) and result.get("updated"):
                success += 1
            else:
                failed += 1
                if logger:
                    errors = result.get("errors") if isinstance(result, dict) else None
                    offer_id = result.get("offer_id") if isinstance(result, dict) else ""
                    logger(f"库存更新失败 {offer_id}: {errors}")
    return success, failed


def generate_barcodes_for_products(
    client: OzonSellerClient,
    rows: list[ProductInventoryRow],
    *,
    logger: Logger | None = None,
) -> tuple[int, int]:
    success = 0
    failed = 0
    for index in range(0, len(rows), BATCH_BARCODE_LIMIT):
        chunk = rows[index : index + BATCH_BARCODE_LIMIT]
        response = client.generate_barcodes([row.product_id for row in chunk])
        errors = response.get("errors")
        error_count = len(errors) if isinstance(errors, list) else 0
        failed += error_count
        success += len(chunk) - error_count
        if logger and isinstance(errors, list):
            for error in errors:
                logger(f"条形码生成失败: {error}")
        if logger:
            logger(
                f"条形码生成批次 {index // BATCH_BARCODE_LIMIT + 1}："
                f"成功 {len(chunk) - error_count}，失败 {error_count}"
            )
    return success, failed
