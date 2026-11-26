import io
import math
import logging
import subprocess
from copy import deepcopy
from functools import lru_cache
from pathlib import Path

import imageio_ffmpeg
from lxml import etree
from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches, Pt


log = logging.getLogger(__name__)
WIDTH, HEIGHT = 960, 540
FPS = 10
BACKGROUND = (255, 255, 255)
ENTITY_SIZE = (116, 84)
ICON_SIZE = (144, 144)
ICON_DIR = Path(__file__).resolve().parents[1] / "assets" / "icons"
ANIMATION_TEMPLATE = Path(__file__).with_name("native_animation_templates.xml")
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
P = f"{{{P_NS}}}"
NS = {"p": P_NS}
ICON_FILES = {"fox": "Fox.png", "dog": "Dog.png"}
ENTITY_COLORS = {
    "red": (204, 0, 0),
    "orange": (255, 140, 0),
    "yellow": (255, 204, 0),
    "green": (0, 153, 51),
    "blue": (0, 102, 204),
    "purple": (128, 0, 160),
    "pink": (255, 105, 180),
    "brown": (139, 69, 19),
    "black": (0, 0, 0),
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


@lru_cache(maxsize=len(ICON_FILES))
def _load_icon(name: str) -> Image.Image | None:
    filename = ICON_FILES.get(name)
    if filename is None:
        return None
    path = ICON_DIR / filename
    if not path.is_file():
        log.warning("Icon artwork is missing: %s", path)
        return None
    with Image.open(path) as icon:
        return icon.convert("RGBA")


def _tinted_icon(name: str, color: tuple[int, int, int]) -> io.BytesIO | None:
    icon = _load_icon(name)
    if icon is None:
        return None
    tinted = Image.new("RGBA", icon.size, (*color, 255))
    tinted.putalpha(icon.getchannel("A"))
    output = io.BytesIO()
    tinted.save(output, format="PNG")
    output.seek(0)
    return output


def _initial_states(actions: list[dict]) -> dict[str, dict]:
    states = {}
    for action in actions:
        related_names = list(action.get("related_shapes", []))
        if action.get("destination_entity"):
            related_names.append(action["destination_entity"])
        for name in (action["shape"], *related_names):
            if name in states:
                continue
            attributes = (
                action if name == action["shape"]
                else action.get("related_entity_data", {}).get(name, {})
            )
            has_icon = _load_icon(name) is not None
            states[name] = {
                "color": ENTITY_COLORS.get(
                    attributes.get("color"), (0, 0, 0) if has_icon else (59, 130, 246)
                ),
                "shape_type": attributes.get("shape_type", ""),
                "has_icon": has_icon,
                "position": attributes.get("position"),
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
    icon = _load_icon(name) if state.get("has_icon") else None
    current_color = color or state["color"]
    if icon is not None:
        icon = icon.resize(ICON_SIZE, Image.Resampling.LANCZOS)
        mask = icon.getchannel("A").point(lambda value: value * alpha // 255)
        icon = Image.new("RGBA", ICON_SIZE, (*current_color, 255))
        icon.putalpha(mask)
        icon_layer = Image.new("RGBA", (WIDTH, HEIGHT), (0, 0, 0, 0))
        icon_layer.alpha_composite(
            icon,
            (round(x * WIDTH - ICON_SIZE[0] / 2), round(y * HEIGHT - ICON_SIZE[1] / 2)),
        )
        overlay.alpha_composite(icon_layer)
    else:
        draw = ImageDraw.Draw(overlay)
        left = int(x * WIDTH - ENTITY_SIZE[0] / 2)
        top = int(y * HEIGHT - ENTITY_SIZE[1] / 2)
        right, bottom = left + ENTITY_SIZE[0], top + ENTITY_SIZE[1]
        fill = (*current_color, alpha)
        shape_type = state.get("shape_type")
        if shape_type == "msoShapeOval":
            draw.ellipse((left, top, right, bottom), fill=fill)
        elif shape_type == "msoShapeIsoscelesTriangle":
            draw.polygon(((left + ENTITY_SIZE[0] // 2, top), (right, bottom), (left, bottom)), fill=fill)
        elif shape_type == "msoShapeDiamond":
            draw.polygon(((left + ENTITY_SIZE[0] // 2, top), (right, top + ENTITY_SIZE[1] // 2),
                          (left + ENTITY_SIZE[0] // 2, bottom), (left, top + ENTITY_SIZE[1] // 2)), fill=fill)
        else:
            draw.rounded_rectangle((left, top, right, bottom), radius=18, fill=fill)

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


def _action_target(action: dict, state: dict, states: dict) -> tuple[float, float]:
    if action.get("destination") in DESTINATIONS:
        return DESTINATIONS[action["destination"]]
    if action.get("destination_entity"):
        target = states.get(action["destination_entity"])
        if target is None:
            raise ValueError(f"Unknown movement target: {action['destination_entity']}")
        direction = -1 if state["x"] <= target["x"] else 1
        mover_width = ICON_SIZE[0] if state["has_icon"] else ENTITY_SIZE[0]
        target_width = ICON_SIZE[0] if target["has_icon"] else ENTITY_SIZE[0]
        target_x = target["x"] + direction * (
            target_width / 2 + mover_width / 2 + 12
        ) / WIDTH
        return target_x, target["y"]
    if action.get("motion") == "arc":
        related_shapes = action.get("related_shapes", [])
        obstacle = states.get(related_shapes[0]) if related_shapes else None
        if obstacle is None:
            raise ValueError("An arc motion requires a related obstacle shape.")
        direction = 1 if state["x"] <= obstacle["x"] else -1
        mover_width = ICON_SIZE[0] if state["has_icon"] else ENTITY_SIZE[0]
        obstacle_width = ICON_SIZE[0] if obstacle["has_icon"] else ENTITY_SIZE[0]
        target_x = obstacle["x"] + direction * (
            obstacle_width / 2 + mover_width + 20
        ) / WIDTH
        return target_x, state["y"]
    return state["x"], state["y"]


def _arc_height(state: dict, obstacle: dict) -> float:
    mover_height = ICON_SIZE[1] if state["has_icon"] else ENTITY_SIZE[1]
    obstacle_height = ICON_SIZE[1] if obstacle["has_icon"] else ENTITY_SIZE[1]
    return (mover_height / 2 + obstacle_height / 2 + 20) / HEIGHT


@lru_cache(maxsize=1)
def _animation_template() -> etree._Element:
    try:
        return etree.fromstring(ANIMATION_TEMPLATE.read_bytes())
    except (OSError, etree.XMLSyntaxError) as exc:
        raise RuntimeError(f"Could not load PowerPoint animation templates: {exc}") from exc


def _animation_template_wrappers(template: etree._Element) -> list[etree._Element]:
    path = "./p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:cTn/p:childTnLst"
    sequence = template.find(path, NS)
    if sequence is None:
        raise RuntimeError("The PowerPoint animation template has no main sequence.")
    group = sequence.find("p:par/p:cTn/p:childTnLst", NS)
    if group is None:
        raise RuntimeError("The PowerPoint animation template has no action group.")
    wrappers = group.findall("p:par", NS)
    if len(wrappers) != 6:
        raise RuntimeError("The PowerPoint animation template is incomplete.")
    return wrappers


def _animation_template_index(action: dict) -> int:
    effect = action.get("effect")
    if effect == "msoAnimEffectFade":
        return 5 if action.get("exit") else 0
    if effect == "msoAnimEffectCustom":
        return 1
    if effect == "msoAnimEffectSpin":
        return 2
    if effect == "msoAnimEffectChangeFillColor":
        return 3
    if effect == "msoAnimEffectBounce":
        return 4
    raise ValueError(f"Unsupported PowerPoint animation effect: {effect}")


def _motion_path(action: dict, state: dict, states: dict) -> str:
    start_x, start_y = state["x"], state["y"]
    target_x, target_y = _action_target(action, state, states)
    delta_x = target_x - start_x
    delta_y = target_y - start_y
    if action.get("motion") != "arc":
        return f"M 0 0 L {delta_x:.3f} {delta_y:.3f}"

    obstacle = states[action["related_shapes"][0]]
    fraction = (obstacle["x"] - start_x) / delta_x
    fraction = min(0.95, max(0.05, fraction))
    control_y = _arc_height(state, obstacle) / (3 * fraction * (1 - fraction))
    return (
        f"M 0 0 C {delta_x / 3:.3f} -{control_y:.3f} "
        f"{delta_x * 2 / 3:.3f} -{control_y:.3f} {delta_x:.3f} {delta_y:.3f}"
    )


def _configure_animation(
    wrapper: etree._Element,
    action: dict,
    shape_id: int,
    state: dict,
    states: dict,
    duration_ms: int,
    click_trigger: bool,
    start_delay_ms: int,
) -> None:
    outer_timing = wrapper.find("./p:cTn/p:stCondLst/p:cond", NS)
    if outer_timing is None:
        raise RuntimeError("An animation template is missing its start condition.")

    effect_timing = wrapper.find(".//p:cTn[@nodeType]", NS)
    if effect_timing is None:
        raise RuntimeError("An animation template is missing its effect node.")
    effect_timing.set("nodeType", "clickEffect" if click_trigger else "afterEffect")

    for node in wrapper.iter():
        if node.tag == P + "spTgt":
            node.set("spid", str(shape_id))
        elif node.tag == P + "cTn":
            duration = node.get("dur")
            if duration and duration.isdigit() and int(duration) > 1:
                node.set("dur", str(max(1, round(int(duration) * duration_ms / 1000))))
        elif node.tag == P + "cond":
            delay = node.get("delay")
            if delay and delay.isdigit() and int(delay) > 0:
                node.set("delay", str(round(int(delay) * duration_ms / 1000)))

    outer_timing.set("delay", "indefinite" if click_trigger else str(start_delay_ms))
    effect = action["effect"]
    if effect == "msoAnimEffectSpin":
        rotation = wrapper.find(".//p:animRot", NS)
        if rotation is None:
            raise RuntimeError("The spin animation template has no rotation effect.")
        amount = float(action.get("amount", 360))
        rotation.set("by", str(round(amount * 60000)))
    elif effect == "msoAnimEffectChangeFillColor":
        color = wrapper.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr", NS)
        if color is None:
            raise RuntimeError("The color animation template has no target color.")
        rgb = ENTITY_COLORS.get(action.get("to_color"), state["color"])
        color.set("val", "".join(f"{channel:02X}" for channel in rgb))
    elif effect == "msoAnimEffectCustom":
        motion = wrapper.find(".//p:animMotion", NS)
        if motion is None:
            raise RuntimeError("The motion animation template has no motion path.")
        motion.set("path", _motion_path(action, state, states))


def _add_native_animations(
    slide,
    actions: list[dict],
    states: dict,
    shape_ids: dict[str, int],
) -> None:
    timing = deepcopy(_animation_template())
    wrappers = _animation_template_wrappers(timing)
    sequence_path = "./p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:cTn/p:childTnLst"
    sequence = timing.find(sequence_path, NS)
    if sequence is None:
        raise RuntimeError("The PowerPoint animation template has no main sequence.")
    action_list = sequence.find("p:par/p:cTn/p:childTnLst", NS)
    if action_list is None:
        raise RuntimeError("The PowerPoint animation template has no action group.")
    for child in list(action_list):
        if child.tag == P + "par":
            action_list.remove(child)

    elapsed_ms = 0
    for index, action in enumerate(actions):
        name = action["shape"]
        if name not in shape_ids:
            raise ValueError(f"Animation target was not drawn on the slide: {name}")
        state = states[name]
        duration_ms = max(1, round(float(action.get("duration", 0.7)) * 1000))
        repeat_count = int(action.get("repeat_count", 1))
        template_index = _animation_template_index(action)
        wrapper = deepcopy(wrappers[template_index])
        _configure_animation(
            wrapper,
            action,
            shape_ids[name],
            state,
            states,
            duration_ms,
            action.get("trigger") == "msoAnimTriggerOnPageClick",
            elapsed_ms,
        )
        effect_timing = wrapper.find(".//p:cTn[@nodeType]", NS)
        if effect_timing is not None and action.get("effect") == "msoAnimEffectBounce":
            effect_timing.set("repeatCount", str(repeat_count * 1000))
        action_list.append(wrapper)
        elapsed_ms += duration_ms * (repeat_count if action.get("effect") == "msoAnimEffectBounce" else 1)

    for node_id, node in enumerate(timing.iter(P + "cTn"), start=1):
        node.set("id", str(node_id))
    build_list = timing.find("p:bldLst", NS)
    if build_list is not None:
        timing.remove(build_list)
    slide._element.append(timing)


def _add_pptx_slide(presentation: Presentation, states: dict):
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    background = slide.background.fill
    background.solid()
    background.fore_color.rgb = _pptx_color(BACKGROUND)

    shape_ids = {}
    for name, state in states.items():
        if not state["visible"]:
            continue
        if state.get("has_icon"):
            icon = _tinted_icon(name, state["color"])
            if icon is not None:
                width, height = Inches(1.7), Inches(1.7)
                left = Inches(0.3 + state["x"] * 12.5) - width / 2
                top = Inches(1.45 + state["y"] * 4.7) - height / 2
                picture = slide.shapes.add_picture(icon, left, top, width, height)
                shape_ids[name] = picture.shape_id
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
        shape_ids[name] = shape.shape_id
        frame = shape.text_frame
        frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        frame.paragraphs[0].text = name.replace("_", " ")
        frame.paragraphs[0].font.size = Pt(16)
        frame.paragraphs[0].font.color.rgb = _pptx_color((255, 255, 255))
    return slide, shape_ids


def _apply_action(state: dict, action: dict, states: dict) -> None:
    if action.get("color") in ENTITY_COLORS:
        state["color"] = ENTITY_COLORS[action["color"]]
    if action.get("shape_type"):
        state["shape_type"] = action["shape_type"]
    if action.get("exit"):
        state["visible"] = False
    else:
        state["visible"] = True
    if (
        action.get("destination") in DESTINATIONS
        or action.get("destination_entity")
        or action.get("motion") == "arc"
    ):
        state["x"], state["y"] = _action_target(action, state, states)
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
    slide, shape_ids = _add_pptx_slide(presentation, states)
    _add_native_animations(slide, actions, states, shape_ids)

    ffmpeg = imageio_ffmpeg.get_ffmpeg_exe()
    command = [
        ffmpeg, "-y", "-hide_banner", "-loglevel", "error",
        "-f", "rawvideo", "-pix_fmt", "rgb24", "-s", f"{WIDTH}x{HEIGHT}",
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
            target_x, target_y = _action_target(action, state, states)
            rotation_delta = (
                float(action.get("amount", 360))
                if action.get("effect") == "msoAnimEffectSpin" else 0
            )
            bounce_count = int(action.get("repeat_count", 1))
            is_bounce = action.get("effect") == "msoAnimEffectBounce"
            is_arc = action.get("motion") == "arc"
            obstacle = states[action["related_shapes"][0]] if is_arc else None
            arc_height = _arc_height(state, obstacle) if obstacle is not None else 0
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
                            if is_bounce else (
                                start_y + (target_y - start_y) * fraction
                                - arc_height * 4 * fraction * (1 - fraction)
                                if is_arc else start_y + (target_y - start_y) * fraction
                            )
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
            _apply_action(state, action, states)

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
