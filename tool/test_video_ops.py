import unittest

from batch_upload.core import build_import_item
from batch_upload.video_ops import (
    VIDEO_ATTR_ID_NAME,
    VIDEO_ATTR_ID_URL,
    VIDEO_COMPLEX_ID,
    apply_video_to_import_item,
    build_video_attributes,
    copy_video_attributes_from_template,
    count_video_url_attributes,
    extract_video_links_from_product,
    resolve_sku_video_links,
    strip_misplaced_video_from_item,
)


class VideoOpsTests(unittest.TestCase):
    def test_resolve_sku_video_links_prefers_excel_per_offer(self) -> None:
        links = resolve_sku_video_links(
            "SKU002",
            default_links=["https://example.com/default.mp4"],
            per_offer_links={"SKU002": ["https://example.com/sku2.mp4"]},
        )
        self.assertEqual(links, ["https://example.com/sku2.mp4"])

    def test_build_video_attributes_use_attributes_shape(self) -> None:
        attrs = build_video_attributes(["https://example.com/a.mp4"], "demo.mp4")
        self.assertEqual(len(attrs), 2)
        self.assertEqual(attrs[0]["id"], VIDEO_ATTR_ID_NAME)
        self.assertEqual(attrs[0]["complex_id"], VIDEO_COMPLEX_ID)
        self.assertEqual(attrs[1]["id"], VIDEO_ATTR_ID_URL)

    def test_extract_video_links_reads_attributes_first(self) -> None:
        product = {
            "attributes": [
                {
                    "id": VIDEO_ATTR_ID_URL,
                    "complex_id": VIDEO_COMPLEX_ID,
                    "values": [{"value": "https://example.com/from-attributes.mp4"}],
                }
            ],
            "complex_attributes": [
                {
                    "id": VIDEO_ATTR_ID_URL,
                    "values": [{"value": "https://example.com/wrong-place.mp4"}],
                }
            ],
        }
        self.assertEqual(extract_video_links_from_product(product), ["https://example.com/from-attributes.mp4"])

    def test_copy_video_attributes_from_template_reads_complex_attributes_fallback(self) -> None:
        template = {
            "offer_id": "TEMPLATE",
            "complex_attributes": [
                {
                    "id": VIDEO_ATTR_ID_NAME,
                    "complex_id": VIDEO_COMPLEX_ID,
                    "values": [{"dictionary_value_id": 0, "value": "模板视频.mp4"}],
                },
                {
                    "id": VIDEO_ATTR_ID_URL,
                    "complex_id": VIDEO_COMPLEX_ID,
                    "values": [{"dictionary_value_id": 0, "value": "https://example.com/template.mp4"}],
                },
            ],
        }
        copied = copy_video_attributes_from_template(template, "SKU001")
        self.assertEqual(len(copied), 2)
        self.assertEqual(copied[1]["values"][0]["value"], "https://example.com/template.mp4")

    def test_apply_video_writes_attributes_and_strips_complex_attributes(self) -> None:
        item = {
            "offer_id": "SKU001",
            "attributes": [{"id": 123, "values": [{"value": "keep"}]}],
            "complex_attributes": [
                {"id": 999, "values": [{"value": "keep-complex"}]},
                {
                    "id": VIDEO_ATTR_ID_URL,
                    "values": [{"value": "https://example.com/old-wrong.mp4"}],
                },
            ],
        }
        applied = apply_video_to_import_item(
            item,
            ["https://example.com/new.mp4"],
            {"offer_id": "TEMPLATE", "attributes": []},
        )
        self.assertTrue(applied)
        self.assertEqual(count_video_url_attributes(item), 1)
        url_attr = next(attr for attr in item["attributes"] if attr["id"] == VIDEO_ATTR_ID_URL)
        self.assertEqual(url_attr["values"][0]["value"], "https://example.com/new.mp4")
        self.assertEqual(url_attr["complex_id"], VIDEO_COMPLEX_ID)
        complex_ids = [int(attr.get("id") or 0) for attr in item["complex_attributes"]]
        self.assertNotIn(VIDEO_ATTR_ID_URL, complex_ids)
        self.assertIn(999, complex_ids)

    def test_build_import_item_applies_template_video_into_attributes(self) -> None:
        template = {
            "offer_id": "TEMPLATE",
            "name": "Template",
            "price": "10",
            "images": ["old"],
            "complex_attributes": [
                {
                    "id": VIDEO_ATTR_ID_NAME,
                    "complex_id": VIDEO_COMPLEX_ID,
                    "values": [{"dictionary_value_id": 0, "value": "模板视频.mp4"}],
                },
                {
                    "id": VIDEO_ATTR_ID_URL,
                    "complex_id": VIDEO_COMPLEX_ID,
                    "values": [{"dictionary_value_id": 0, "value": "https://example.com/template.mp4"}],
                },
            ],
        }
        item = build_import_item(
            template,
            "SKU001",
            "Title",
            "Description",
            ["https://oss.example.com/1.jpg"],
            use_template_video=True,
        )
        self.assertEqual(count_video_url_attributes(item), 1)
        complex_ids = [int(attr.get("id") or 0) for attr in item.get("complex_attributes") or []]
        self.assertNotIn(VIDEO_ATTR_ID_URL, complex_ids)

    def test_strip_misplaced_video_from_item(self) -> None:
        item = {
            "complex_attributes": [
                {"id": VIDEO_ATTR_ID_URL, "values": [{"value": "https://example.com/x.mp4"}]},
                {"id": 1, "values": [{"value": "ok"}]},
            ]
        }
        strip_misplaced_video_from_item(item)
        self.assertEqual([int(attr["id"]) for attr in item["complex_attributes"]], [1])


if __name__ == "__main__":
    unittest.main()
