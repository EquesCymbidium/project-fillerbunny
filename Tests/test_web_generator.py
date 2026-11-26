import sys
import subprocess
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
from zipfile import ZipFile
from xml.etree import ElementTree

import imageio_ffmpeg
from PIL import Image
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE


PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "Product" / "Blackbox"))

from dispatcher import parse_prompt
import dispatcher
from Web.backend.app import OUTPUT_FILENAME_RE, _output_filenames
from Web.backend.generator import ENTITY_COLORS, _initial_states, create_outputs


class WebGeneratorTests(unittest.TestCase):
    def test_outputs_have_hex_timestamp_suffixes(self):
        pptx_name, mp4_name = _output_filenames(0x1ABCDEF)

        self.assertEqual(pptx_name, "presentation-1abcdef.pptx")
        self.assertEqual(mp4_name, "animation-1abcdef.mp4")
        self.assertRegex(pptx_name, OUTPUT_FILENAME_RE)
        self.assertRegex(mp4_name, OUTPUT_FILENAME_RE)
        self.assertIsNone(OUTPUT_FILENAME_RE.fullmatch("../presentation-1abcdef.pptx"))

    def test_jump_prompt_keeps_each_entities_color_and_autostarts(self):
        actions = parse_prompt("The quick orange fox jumped over the lazy brown dog")

        self.assertEqual(len(actions), 1)
        self.assertEqual(actions[0]["trigger"], "msoAnimTriggerAfterPrevious")
        self.assertEqual(actions[0]["color"], "orange")
        self.assertEqual(actions[0]["related_entity_data"]["dog"]["color"], "brown")
        states = _initial_states(actions)
        self.assertEqual(states["fox"]["color"], ENTITY_COLORS["orange"])
        self.assertEqual(states["dog"]["color"], ENTITY_COLORS["brown"])

        with tempfile.TemporaryDirectory() as directory:
            pptx_path = Path(directory) / "presentation.pptx"
            mp4_path = Path(directory) / "animation.mp4"
            create_outputs(actions, pptx_path, mp4_path)
            namespaces = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main"}
            with ZipFile(pptx_path) as package:
                slide_xml = ElementTree.fromstring(package.read("ppt/slides/slide1.xml"))
                images = sorted(
                    (name for name in package.namelist()
                     if name.startswith("ppt/media/") and name.endswith(".png")),
                    key=lambda name: int(name.rsplit("image", 1)[1].split(".", 1)[0]),
                )
                picture_colors = []
                for image_name in images:
                    image = Image.open(package.open(image_name)).convert("RGBA")
                    pixel_data = image.tobytes()
                    picture_colors.append(next(
                        tuple(pixel_data[offset:offset + 3])
                        for offset in range(0, len(pixel_data), 4)
                        if pixel_data[offset + 3] > 0
                    ))
            self.assertEqual(
                picture_colors,
                [ENTITY_COLORS["orange"], ENTITY_COLORS["brown"]],
            )
            effect = slide_xml.find(".//p:cTn[@nodeType='afterEffect']", namespaces)
            self.assertIsNotNone(effect)

    def test_parser_preserves_spin_degrees_and_direction(self):
        actions = parse_prompt("the fox spins 180 degrees counterclockwise")

        self.assertEqual(len(actions), 1)
        self.assertEqual(actions[0]["effect"], "msoAnimEffectSpin")
        self.assertEqual(actions[0]["amount"], -180)

    def test_native_spin_uses_the_requested_angle(self):
        actions = parse_prompt("the fox spins 180 degrees counterclockwise")

        with tempfile.TemporaryDirectory() as directory:
            pptx_path = Path(directory) / "presentation.pptx"
            mp4_path = Path(directory) / "animation.mp4"
            create_outputs(actions, pptx_path, mp4_path)

            namespaces = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main"}
            with ZipFile(pptx_path) as package:
                slide_xml = ElementTree.fromstring(package.read("ppt/slides/slide1.xml"))
            rotation = slide_xml.find(".//p:animRot", namespaces)

        self.assertIsNotNone(rotation)
        self.assertEqual(rotation.attrib["by"], "-10800000")

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
            self.assertEqual(len(slides), 1)
            pictures = [
                shape for shape in slides[0].shapes
                if shape.shape_type == MSO_SHAPE_TYPE.PICTURE
            ]
            self.assertEqual(len(pictures), 2)
            self.assertLess(pictures[0].left, pictures[1].left)
            self.assertEqual(slides[0].background.fill.fore_color.rgb, (255, 255, 255))
            namespaces = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main"}
            with ZipFile(pptx_path) as package:
                slide_xml = ElementTree.fromstring(package.read("ppt/slides/slide1.xml"))
            timing = slide_xml.find("p:timing", namespaces)
            self.assertIsNotNone(timing)
            motion = timing.find(".//p:animMotion", namespaces)
            self.assertIsNotNone(motion)
            self.assertIn(" C ", motion.attrib["path"])
            shape_ids = {str(shape.shape_id) for shape in slides[0].shapes}
            target_ids = {
                target.attrib["spid"]
                for target in timing.findall(".//p:spTgt", namespaces)
            }
            self.assertTrue(target_ids.issubset(shape_ids))
            self.assertEqual(len(target_ids), 1)
            self.assertGreater(mp4_path.stat().st_size, 0)
            with mp4_path.open("rb") as video:
                header = video.read(12)
            self.assertEqual(header[4:8], b"ftyp")

            decoded = subprocess.run(
                [
                    imageio_ffmpeg.get_ffmpeg_exe(), "-hide_banner", "-loglevel", "error",
                    "-i", str(mp4_path), "-f", "rawvideo", "-pix_fmt", "rgb24", "-",
                ],
                check=True,
                capture_output=True,
            ).stdout
            frame_size = 960 * 540 * 3
            first = Image.frombytes("RGB", (960, 540), decoded[:frame_size]).convert("L")
            last = Image.frombytes("RGB", (960, 540), decoded[-frame_size:]).convert("L")
            first_bounds = first.point(lambda value: 255 if value < 80 else 0).getbbox()
            last_bounds = last.point(lambda value: 255 if value < 80 else 0).getbbox()
            self.assertIsNotNone(first_bounds)
            self.assertIsNotNone(last_bounds)
            self.assertGreater(last_bounds[2], first_bounds[2] + 30)

    def test_rejects_empty_action_list(self):
        with tempfile.TemporaryDirectory() as directory:
            with self.assertRaisesRegex(ValueError, "No supported animation actions"):
                create_outputs([], Path(directory) / "output.pptx", Path(directory) / "output.mp4")


if __name__ == "__main__":
    unittest.main()
