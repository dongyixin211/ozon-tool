import tempfile
import unittest
from pathlib import Path

from PIL import Image

from app import (
    create_portrait_variant,
    list_images,
    list_subfolders,
    normalize_ozon_client_id,
    parse_copy_payload,
    sanitize_json_text,
)


class AppTests(unittest.TestCase):
    def test_normalize_ozon_client_id_extracts_numeric_id(self) -> None:
        self.assertEqual(normalize_ozon_client_id("你的树 -15937"), "15937")
        self.assertEqual(normalize_ozon_client_id("15937"), "15937")
        self.assertEqual(normalize_ozon_client_id("店铺 12 / 15937"), "15937")

    def test_folder_and_image_discovery(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)
            (root / "SKU001").mkdir()
            (root / "SKU002").mkdir()
            Image.new("RGB", (800, 800), "red").save(root / "SKU001" / "a.png")
            Image.new("RGB", (800, 800), "blue").save(root / "SKU001" / "b.jpg")

            folders = list_subfolders(root)
            images = list_images(root / "SKU001")

            self.assertEqual([item.name for item in folders], ["SKU001", "SKU002"])
            self.assertEqual([item.name for item in images], ["a.png", "b.jpg"])

    def test_create_portrait_variant(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)
            source = root / "source.png"
            target = root / "portrait.png"
            Image.new("RGB", (1000, 1000), "#55aaee").save(source)

            create_portrait_variant(source, target)

            self.assertTrue(target.exists())
            with Image.open(target) as result:
                self.assertEqual(result.size, (1200, 1600))

    def test_parse_copy_payload(self) -> None:
        raw = """```json
        {"title":"标题A","description":"描述B","bullets":["卖点1","卖点2"]}
        ```"""
        self.assertEqual(
            parse_copy_payload(raw),
            {
                "title": "标题A",
                "description": "描述B",
                "bullets": ["卖点1", "卖点2"],
            },
        )
        self.assertEqual(
            sanitize_json_text(raw),
            '{"title":"标题A","description":"描述B","bullets":["卖点1","卖点2"]}',
        )


if __name__ == "__main__":
    unittest.main()
