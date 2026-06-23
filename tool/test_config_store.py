import tempfile
import unittest
from pathlib import Path

from config_store import ConfigStoreError, config_exists, read_config_payload, write_config_payload


class ConfigStoreTests(unittest.TestCase):
    def test_read_missing_config_returns_empty_payload(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "config.json"

            self.assertFalse(config_exists(path))
            self.assertEqual(read_config_payload(path), {})

    def test_write_and_read_config_payload(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "config.json"
            payload = {"source_root": "/tmp/source", "enabled": True}

            write_config_payload(payload, path)

            self.assertTrue(config_exists(path))
            self.assertEqual(read_config_payload(path), payload)

    def test_invalid_config_returns_empty_payload_by_default(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "config.json"
            path.write_text("{bad json", encoding="utf-8")

            self.assertEqual(read_config_payload(path), {})

    def test_invalid_config_can_raise_error(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "config.json"
            path.write_text("{bad json", encoding="utf-8")

            with self.assertRaises(ConfigStoreError):
                read_config_payload(path, raise_on_error=True)


if __name__ == "__main__":
    unittest.main()
