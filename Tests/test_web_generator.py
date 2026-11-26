import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE


PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "Product" / "Blackbox"))

from dispatcher import parse_prompt
import dispatcher
from Web.backend.generator import create_outputs


class WebGeneratorTests(unittest.TestCase):
    def test_parser_supports_common_entities_without_icon_artwork(self):
        dispatcher._icon_entity_names.cache_clear()
        try:
            with patch.object(dispatcher, "_ICON_DIR", str(PROJECT_ROOT / "missing-icons")):
                actions = parse_prompt("the fox jumps over the dog")
        finally:
            dispatcher._icon_entity_names.cache_clear()

        self.assertEqual(len(actions), 1)
        self.assertEqual(actions[0]["shape"], "fox")
        self.assertEqual(actions[0]["related_shapes"], ["dog"])

    def test_creates_pptx_and_mp4_for_supported_prompt(self):
        actions = parse_prompt("the fox jumps over the dog")
        self.assertEqual(len(actions), 1)

        with tempfile.TemporaryDirectory() as directory:
            pptx_path = Path(directory) / "presentation.pptx"
            mp4_path = Path(directory) / "animation.mp4"
            create_outputs(actions, pptx_path, mp4_path)

            self.assertGreater(pptx_path.stat().st_size, 0)
            slides = Presentation(pptx_path).slides
            self.assertEqual(len(slides), len(actions))
            pictures = [
                shape for shape in slides[0].shapes
                if shape.shape_type == MSO_SHAPE_TYPE.PICTURE
            ]
            self.assertEqual(len(pictures), 2)
            self.assertEqual(slides[0].background.fill.fore_color.rgb, (255, 255, 255))
            self.assertGreater(mp4_path.stat().st_size, 0)
            with mp4_path.open("rb") as video:
                header = video.read(12)
            self.assertEqual(header[4:8], b"ftyp")

    def test_rejects_empty_action_list(self):
        with tempfile.TemporaryDirectory() as directory:
            with self.assertRaisesRegex(ValueError, "No supported animation actions"):
                create_outputs([], Path(directory) / "output.pptx", Path(directory) / "output.mp4")


if __name__ == "__main__":
    unittest.main()
