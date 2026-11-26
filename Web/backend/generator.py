import math
import logging
import subprocess
from pathlib import Path

import imageio_ffmpeg
from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches, Pt


log = logging.getLogger(__name__)
WIDTH, HEIGHT = 960, 540
FPS = 10
BACKGROUND = (13, 19, 34)
ACCENT = (100, 220, 184)
ENTITY_SIZE = (116, 84)
ENTITY_COLORS = {
    "red": (204, 0, 0),
    "orange": (255, 140, 0),
    "yellow": (255, 204, 0),
    "green": (0, 153, 51),
    "blue": (0, 102, 204),
    "purple": (128, 0, 160),
    "pink": (255, 105, 180),
    "brown": (139, 69, 19),
    "black": (40, 45, 54),
    "white": (225, 230, 238),
    "gray": (128, 128, 128),
    "grey": (128, 128, 128),
    "teal": (0, 128, 128),
    "gold": (212, 175, 55),
}
SHAPE_TYPES = {
    "msoShapeRectangle": MSO_SHAPE.RECTANGLE,
    "msoShapeOval": MSO_SHAPE.OVAL,
    "msoShapeIsoscelesTriangle": MSO_SHAPE.ISOSCELES_TRIANGLE,
    "msoShapeDiamond": MSO_SHAPE.DIAMOND,
    "msoShape5pointStar": MSO_SHAPE.STAR_5_POINT,
    "msoShapeHeart": MSO_SHAPE.HEART,
    "msoShapeHexagon": MSO_SHAPE.HEXAGON,
    "msoShapeRegularPentagon": MSO_SHAPE.PENTAGON,
    "msoShapeCross": MSO_SHAPE.CROSS,
    "msoShapeRightArrow": MSO_SHAPE.RIGHT_ARROW,
    "msoShapeMoon": MSO_SHAPE.MOON,
    "msoShapeCloud": MSO_SHAPE.CLOUD,
}
DESTINATIONS = {
    "left": (0.18, 0.52),
    "right": (0.82, 0.52),
    "up": (0.5, 0.22),
    "down": (0.5, 0.78),
    "center": (0.5, 0.52),
    "offscreen_left": (-0.12, 0.52),
    "offscreen_right": (1.12, 0.52),
    "offscreen_up": (0.5, -0.18),
    "offscreen_down": (0.5, 1.18),
}


def _initial_states(actions: list[dict]) -> dict[str, dict]:
    states = {}
    for action in actions:
        for name in (action["shape"], *action.get("related_shapes", [])):
            if name in states:
                continue
            states[name] = {
                "color": ENTITY_COLORS.get(action.get("color"), (59, 130, 246)),
                "shape_type": action.get("shape_type", ""),
                "position": action.get("position"),
                "x": 0.28 + (len(states) % 3) * 0.22,
                "y": 0.52,
                "visible": True,
                "rotation": 0,
            }
    for name, state in states.items():
        if state["position"] in ("left", "right"):
            state["x"] = 0.18 if state["position"] == "left" else 0.82
    return states


def _font(size: int) -> ImageFont.ImageFont:
    try:
        return ImageFont.truetype("DejaVuSans.ttf", size)
    except OSError:
        return ImageFont.load_default()


def _draw_entity(
    image: Image.Image,
    name: str,
    state: dict,
    x: float,
    y: float,
    alpha: int,
    rotation: float,
    color: tuple[int, int, int] | None = None,
) -> None:
    if alpha <= 0:
        return
    overlay = Image.new("RGBA", (WIDTH, HEIGHT), (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay)
    left = int(x * WIDTH - ENTITY_SIZE[0] / 2)
    top = int(y * HEIGHT - ENTITY_SIZE[1] / 2)
    right, bottom = left + ENTITY_SIZE[0], top + ENTITY_SIZE[1]
    color = (*(color or state["color"]), alpha)
    shape_type = state.get("shape_type")
    if shape_type == "msoShapeOval":
        draw.ellipse((left, top, right, bottom), fill=color)
    elif shape_type == "msoShapeIsoscelesTriangle":
        draw.polygon(((left + ENTITY_SIZE[0] // 2, top), (right, bottom), (left, bottom)), fill=color)
    elif shape_type == "msoShapeDiamond":
        draw.polygon(((left + ENTITY_SIZE[0] // 2, top), (right, top + ENTITY_SIZE[1] // 2),
                      (left + ENTITY_SIZE[0] // 2, bottom), (left, top + ENTITY_SIZE[1] // 2)), fill=color)
    else:
        draw.rounded_rectangle((left, top, right, bottom), radius=18, fill=color)

    label = name.replace("_", " ")
    bounds = draw.textbbox((0, 0), label, font=_font(18))
    label_x = left + (ENTITY_SIZE[0] - bounds[2] + bounds[0]) // 2
    label_y = top + (ENTITY_SIZE[1] - bounds[3] + bounds[1]) // 2
    draw.text((label_x, label_y), label, font=_font(18), fill=(255, 255, 255, alpha))
    if rotation:
        overlay = overlay.rotate(rotation, resample=Image.Resampling.BICUBIC)
    image.alpha_composite(overlay)


def _render_frame(states: dict, step: int, total: int, action: dict,
                  active: dict | None = None) -> Image.Image:
    image = Image.new("RGBA", (WIDTH, HEIGHT), BACKGROUND + (255,))
    draw = ImageDraw.Draw(image)
    draw.text((40, 30), "FILLERBUNNY  /  ANIMATION PREVIEW",
              font=_font(16), fill=ACCENT)
    draw.text((40, 68), f"STEP {step:02d}  /  {total:02d}",
              font=_font(24), fill=(245, 247, 250))
    effect = action.get("effect", "").replace("msoAnimEffect", "").lower()
    detail = f"{action['shape'].replace('_', ' ')}  ·  {effect}"
    draw.text((40, 106), detail, font=_font(17), fill=(174, 188, 205))
    draw.line((40, 145, WIDTH - 40, 145), fill=(41, 54, 75), width=2)
    draw.rectangle((40, HEIGHT - 48, WIDTH - 40, HEIGHT - 46), fill=(41, 54, 75))
    progress = int((WIDTH - 80) * step / max(total, 1))
    draw.rectangle((40, HEIGHT - 48, 40 + progress, HEIGHT - 46), fill=ACCENT)

    for name, state in states.items():
        motion = active.get(name) if active else None
        x, y = (motion["x"], motion["y"]) if motion else (state["x"], state["y"])
        alpha = motion["alpha"] if motion else (255 if state["visible"] else 0)
        rotation = motion["rotation"] if motion else state["rotation"]
        color = motion.get("color") if motion else None
        _draw_entity(image, name, state, x, y, alpha, rotation, color)
    return image


def _pptx_color(rgb: tuple[int, int, int]) -> RGBColor:
    return RGBColor(*rgb)


def _add_pptx_slide(presentation: Presentation, states: dict, step: int,
                    total: int, action: dict) -> None:
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    background = slide.background.fill
    background.solid()
    background.fore_color.rgb = _pptx_color(BACKGROUND)

    title = slide.shapes.add_textbox(Inches(0.55), Inches(0.28), Inches(12.2), Inches(0.45))
    paragraph = title.text_frame.paragraphs[0]
    paragraph.text = "FILLERBUNNY  /  ANIMATION PREVIEW"
    paragraph.font.size = Pt(13)
    paragraph.font.bold = True
    paragraph.font.color.rgb = _pptx_color(ACCENT)

    heading = slide.shapes.add_textbox(Inches(0.55), Inches(0.78), Inches(12.2), Inches(0.55))
    paragraph = heading.text_frame.paragraphs[0]
    paragraph.text = f"STEP {step:02d} / {total:02d}   {action['shape'].replace('_', ' ')}"
    paragraph.font.size = Pt(24)
    paragraph.font.bold = True
    paragraph.font.color.rgb = _pptx_color((245, 247, 250))

    for name, state in states.items():
        if not state["visible"]:
            continue
        width, height = Inches(1.7), Inches(1.25)
        left = Inches(0.3 + state["x"] * 12.5) - width / 2
        top = Inches(1.45 + state["y"] * 4.7) - height / 2
        shape = slide.shapes.add_shape(
            SHAPE_TYPES.get(state["shape_type"], MSO_SHAPE.ROUNDED_RECTANGLE),
            left, top, width, height,
        )
        shape.fill.solid()
        shape.fill.fore_color.rgb = _pptx_color(state["color"])
        shape.line.fill.background()
        shape.rotation = state["rotation"]
        frame = shape.text_frame
        frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        frame.paragraphs[0].text = name.replace("_", " ")
        frame.paragraphs[0].font.size = Pt(16)
        frame.paragraphs[0].font.color.rgb = _pptx_color((255, 255, 255))

    footer = slide.shapes.add_textbox(Inches(0.55), Inches(6.95), Inches(12.2), Inches(0.25))
    footer.text_frame.paragraphs[0].text = "Static slide snapshots · animation effects are shown in the MP4 preview"
    footer.text_frame.paragraphs[0].font.size = Pt(10)
    footer.text_frame.paragraphs[0].font.color.rgb = _pptx_color((174, 188, 205))


def _apply_action(state: dict, action: dict) -> None:
    if action.get("color") in ENTITY_COLORS:
        state["color"] = ENTITY_COLORS[action["color"]]
    if action.get("shape_type"):
        state["shape_type"] = action["shape_type"]
    if action.get("exit"):
        state["visible"] = False
    else:
        state["visible"] = True
    if action.get("destination") in DESTINATIONS:
        state["x"], state["y"] = DESTINATIONS[action["destination"]]
    if action.get("destination_entity"):
        state["x"], state["y"] = 0.5, 0.52
    if action.get("effect") == "msoAnimEffectSpin":
        state["rotation"] = (state["rotation"] + float(action.get("amount", 360))) % 360
    if action.get("effect") == "msoAnimEffectChangeFillColor":
        state["color"] = ENTITY_COLORS.get(action.get("to_color"), state["color"])


def create_outputs(actions: list[dict], pptx_path: Path, mp4_path: Path) -> None:
    if not actions:
        raise ValueError("No supported animation actions were found.")

    states = _initial_states(actions)
    presentation = Presentation()
    presentation.slide_width = Inches(13.333)
    presentation.slide_height = Inches(7.5)
    width, height = 960, 540

    ffmpeg = imageio_ffmpeg.get_ffmpeg_exe()
    command = [
        ffmpeg, "-y", "-hide_banner", "-loglevel", "error",
        "-f", "rawvideo", "-pix_fmt", "rgb24", "-s", f"{width}x{height}",
        "-r", str(FPS), "-i", "-", "-an", "-c:v", "libx264",
        "-pix_fmt", "yuv420p", "-movflags", "+faststart", str(mp4_path),
    ]
    process = subprocess.Popen(command, stdin=subprocess.PIPE, stderr=subprocess.PIPE)
    try:
        for step, action in enumerate(actions, start=1):
            name = action["shape"]
            state = states[name]
            start_x, start_y = state["x"], state["y"]
            start_rotation = state["rotation"]
            target_color = ENTITY_COLORS.get(action.get("to_color"), state["color"])
            is_exit = bool(action.get("exit"))
            start_alpha = 0 if action.get("effect") == "msoAnimEffectFade" and not is_exit else 255
            end_alpha = 0 if is_exit else 255
            target_x, target_y = start_x, start_y
            if action.get("destination") in DESTINATIONS:
                target_x, target_y = DESTINATIONS[action["destination"]]
            elif action.get("destination_entity"):
                target_x, target_y = 0.5, 0.52
            rotation_delta = (
                float(action.get("amount", 360))
                if action.get("effect") == "msoAnimEffectSpin" else 0
            )
            bounce_count = int(action.get("repeat_count", 1))
            is_bounce = action.get("effect") == "msoAnimEffectBounce"
            frame_count = max(4, min(24, round(float(action.get("duration", 0.7)) * FPS)))
            if is_bounce:
                frame_count = max(frame_count, min(24, bounce_count * 6))
            for frame_index in range(frame_count):
                fraction = (frame_index + 1) / frame_count
                bounce = abs(math.sin(math.pi * bounce_count * fraction))
                active = {
                    name: {
                        "x": start_x + (target_x - start_x) * fraction,
                        "y": (
                            start_y - 0.18 * bounce
                            if is_bounce else start_y + (target_y - start_y) * fraction
                        ),
                        "alpha": round(start_alpha + (end_alpha - start_alpha) * fraction),
                        "rotation": start_rotation + rotation_delta * fraction,
                        "color": tuple(
                            round(start + (end - start) * fraction)
                            for start, end in zip(state["color"], target_color)
                        ),
                    }
                }
                image = _render_frame(states, step, len(actions), action, active)
                if process.stdin is None:
                    raise RuntimeError("The MP4 encoder input stream closed unexpectedly.")
                process.stdin.write(image.convert("RGB").tobytes())
            _apply_action(state, action)
            _add_pptx_slide(presentation, states, step, len(actions), action)

        if process.stdin is not None:
            process.stdin.close()
            process.stdin = None
        stderr = process.stderr.read() if process.stderr else b""
        return_code = process.wait()
        if return_code:
            detail = stderr.decode("utf-8", errors="replace")[-1000:]
            raise RuntimeError(f"MP4 encoding failed: {detail}")
    except Exception:
        if process.poll() is None:
            process.kill()
        process.wait()
        raise
    finally:
        if process.stderr is not None:
            process.stderr.close()

    presentation.save(pptx_path)
    if not pptx_path.is_file() or not mp4_path.is_file():
        raise RuntimeError("The generator did not create both requested output files.")
