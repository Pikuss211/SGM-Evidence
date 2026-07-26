import unittest
from pathlib import Path

from PIL import Image


ROOT = Path(__file__).resolve().parents[1]
ICON = ROOT / "assets" / "sgm_jokey.ico"
SPEC = ROOT / "SGM_Evidence.spec"


class ReleaseAssetTests(unittest.TestCase):
    def test_windows_icon_contains_all_common_sizes(self):
        with Image.open(ICON) as icon:
            sizes = icon.ico.sizes()

        self.assertTrue(
            {(16, 16), (32, 32), (48, 48), (256, 256)}.issubset(sizes)
        )

    def test_build_spec_embeds_and_uses_icon(self):
        source = SPEC.read_text(encoding="utf-8")
        self.assertIn('assets" / "sgm_jokey.ico"', source)
        self.assertIn("icon=str(", source)


if __name__ == "__main__":
    unittest.main()
