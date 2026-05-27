import unittest

from batch_upload.inventory_ops import (
    ProductInventoryRow,
    WarehouseOption,
    _summarize_stocks,
    deserialize_warehouse_options,
    fetch_no_barcode_products,
    fetch_zero_stock_products,
    serialize_warehouse_options,
    update_product_stocks,
)


class InventoryOpsTests(unittest.TestCase):
    def test_warehouse_cache_roundtrip(self) -> None:
        options = [
            WarehouseOption(warehouse_id=100, name="FBS-主仓"),
            WarehouseOption(warehouse_id=200, name="FBS-备用"),
        ]
        restored = deserialize_warehouse_options(serialize_warehouse_options(options))
        self.assertEqual([option.warehouse_id for option in restored], [100, 200])
        self.assertEqual(restored[0].label, "FBS-主仓 (100)")

    def test_summarize_stocks(self) -> None:
        summary = _summarize_stocks(
            [
                {"warehouse_name": "FBS-1", "present": 0},
                {"warehouse_name": "FBS-2", "present": 3},
            ]
        )
        self.assertIn("FBS-1:0", summary)
        self.assertIn("FBS-2:3", summary)

    def test_fetch_zero_stock_products_skips_any_warehouse_with_stock(self) -> None:
        class FakeClient:
            def list_products(self, *, last_id: str = "", limit: int = 100, visibility: str = "ALL") -> dict:
                return {
                    "result": {
                        "items": [
                            {"product_id": 1, "offer_id": "A", "name": "Zero", "stocks": [{"warehouse_name": "W1", "present": 0}]},
                            {
                                "product_id": 2,
                                "offer_id": "B",
                                "name": "HasStock",
                                "stocks": [{"warehouse_name": "W1", "present": 2}],
                            },
                        ],
                        "last_id": "",
                    }
                }

        rows = fetch_zero_stock_products(FakeClient())
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0].offer_id, "A")

    def test_fetch_no_barcode_products_filters_existing_codes(self) -> None:
        class FakeClient:
            def list_products(self, *, last_id: str = "", limit: int = 100, visibility: str = "ALL") -> dict:
                return {
                    "result": {
                        "items": [
                            {"product_id": 1, "offer_id": "A", "name": "NoCode", "barcodes": []},
                            {"product_id": 2, "offer_id": "B", "name": "HasCode", "barcodes": ["123"]},
                        ],
                        "last_id": "",
                    }
                }

        rows = fetch_no_barcode_products(FakeClient())
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0].offer_id, "A")

    def test_update_product_stocks_builds_payload(self) -> None:
        captured: list[dict] = []

        class FakeClient:
            def update_stocks(self, stocks: list[dict]) -> dict:
                captured.extend(stocks)
                return {"result": [{"updated": True, "offer_id": stock["offer_id"], "errors": []} for stock in stocks]}

        rows = [ProductInventoryRow(product_id=9, offer_id="SKU9", name="Item", barcodes=[], stock_summary="")]
        success, failed = update_product_stocks(FakeClient(), rows, warehouse_id=100, stock=5)
        self.assertEqual(success, 1)
        self.assertEqual(failed, 0)
        self.assertEqual(captured[0]["warehouse_id"], 100)
        self.assertEqual(captured[0]["stock"], 5)


if __name__ == "__main__":
    unittest.main()
