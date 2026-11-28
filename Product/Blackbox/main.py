"""
main.py

End-to-end pipeline: natural-language prompt -> animated PowerPoint file.

    python main.py "A fox appears, then spins 180 degrees." fox.pptx --record

Steps (see run_pipeline):
    1. dispatcher.parse_prompt      prompt -> list of animation actions
    2. vba_assembler.generate_vba   one VBA Sub per new shape and per action
    3. win32com                     inject all the VBA, run every Sub in order, save
    4. optional PowerPoint export   create an MP4 from the saved presentation
"""

import argparse
import functools
import logging
import os
import re
import sys
import tempfile
import time

import pywintypes
import win32com.client

from dispatcher import COLOR_RGB, ENTITY_ICON_ALIASES, MAX_ACTIONS, MAX_PROMPT_LENGTH
from llm_dispatcher import parse_prompt_auto
from vba_assembler import generate_vba

log = logging.getLogger("pipeline")

_PROJECT_ROOT = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", ".."))
MACRO_DIR = os.path.join(_PROJECT_ROOT, "Library", "MacroScripts")
ICON_DIR = os.path.join(_PROJECT_ROOT, "Library", "Icons")
ADD_SHAPE = os.path.join(MACRO_DIR, "AddShape.txt")
ADD_ICON = os.path.join(MACRO_DIR, "AddIcon.txt")
ADD_EMPHASIS = os.path.join(MACRO_DIR, "AddEndovital.txt")   # emphasis effects (spin, ...)
ADD_ENTRANCE_EXIT = os.path.join(MACRO_DIR, "AddExovital.txt")  # entrance / exit effects
ADD_MOTION_PATH = os.path.join(MACRO_DIR, "AddMotionPath.txt")

# Where new shapes go. Each extra shape steps right so they don't pile up.
DEFAULT_X, DEFAULT_Y, X_STEP = 300, 200, 150
LEFT_X, RIGHT_X = 120, 740
SHAPE_SIZE = 100
SHAPE_RGB = (0, 102, 204)  # default blue, used when the prompt names no colour
DEFAULT_SHAPE_TYPE = "msoShapeRectangle"
ICON_EXTENSIONS = (".svg", ".png")
EFFECT_DURATION = 1  # seconds

# Effects that belong to the emphasis template; everything else is treated as
# an entrance (or an exit when the action has "exit": True). Extend as needed.
EMPHASIS_EFFECTS = {
    "msoAnimEffectBounce",
    "msoAnimEffectChangeFillColor",
    "msoAnimEffectSpin",
    "msoAnimEffectGrowShrink",
    "msoAnimEffectTeeter",
    "msoAnimEffectWave",
    "msoAnimEffectFlicker",
    "msoAnimEffectTransparency",
}

PP_LAYOUT_BLANK = 12      # ppLayoutBlank
PP_ALERTS_NONE = 1        # ppAlertsNone
PP_SAVE_PPTX = 24         # ppSaveAsOpenXMLPresentation
PP_SAVE_PPTM = 25         # ppSaveAsOpenXMLMacroEnabled
VB_STD_MODULE = 1         # vbext_ct_StdModule
PP_ANIM_TRIGGER_ON_CLICK = 1
PP_ANIM_TRIGGER_AFTER_PREVIOUS = 3
PP_MEDIA_TASK_STATUS_IN_PROGRESS = 1
PP_MEDIA_TASK_STATUS_QUEUED = 2
PP_MEDIA_TASK_STATUS_DONE = 3
PP_MEDIA_TASK_STATUS_FAILED = 4
VIDEO_DEFAULT_SLIDE_DURATION = 3
VIDEO_VERTICAL_RESOLUTION = 1080
VIDEO_FRAMES_PER_SECOND = 30
VIDEO_QUALITY = 85
VIDEO_EXPORT_TIMEOUT_SECONDS = 600


def _identifier(text: str) -> str:
    """Make text safe to use inside a VBA Sub name."""
    return re.sub(r"[^A-Za-z0-9_]", "_", text)


def _normalize(name: str) -> str:
    """'Red_Fox' / 'red fox' -> 'redfox': lowercase letters and digits only."""
    return re.sub(r"[^a-z0-9]", "", name.lower())


@functools.lru_cache(maxsize=1)
def _icon_index() -> dict:
    """Map normalised icon name -> full path for every .svg/.png in Library/Icons."""
    index = {}
    for filename in sorted(os.listdir(ICON_DIR)):
        stem, ext = os.path.splitext(filename)
        if ext.lower() in ICON_EXTENSIONS:
            index.setdefault(_normalize(stem), os.path.join(ICON_DIR, filename))
    return index


def find_icon(shape_name: str):
    """Return the absolute path of the icon matching shape_name, or None.

    The match is on the whole file name, ignoring case and separators:
    "fox" or "Fox" finds Fox.svg, but not Fox_M.svg or FoxHole.svg.
    A simple plural ("foxes", "bears") also matches. An .svg wins over a .png.
    """
    name = _normalize(shape_name)
    candidates = [name]
    if name.endswith("es"):
        candidates.append(name[:-2])
    if name.endswith("s"):
        candidates.append(name[:-1])

    index = _icon_index()
    for candidate in candidates:
        if candidate in index:
            return index[candidate]
    return None


def tint_svg(svg_path: str, rgb: tuple) -> str:
    """Write a recoloured copy of an SVG to the temp folder and return its path.

    The stock icons have no fill of their own, so setting `fill` on the root
    <svg> element colours every path. (PNGs can't be tinted this way.)
    """
    with open(svg_path, "r", encoding="utf-8") as handle:
        text = handle.read()
    hex_color = "{:02x}{:02x}{:02x}".format(*rgb)
    text = re.sub(r"<svg\b", f'<svg fill="#{hex_color}"', text, count=1)

    out_dir = os.path.join(tempfile.gettempdir(), "fillerbunny_icons")
    os.makedirs(out_dir, exist_ok=True)
    stem = os.path.splitext(os.path.basename(svg_path))[0]
    out_path = os.path.join(out_dir, f"{stem}_{hex_color}.svg")
    with open(out_path, "w", encoding="utf-8") as handle:
        handle.write(text)
    return out_path


def _entity_vba(sub_name: str, shape: str, index: int, entity: dict) -> str:
    """Generate the VBA that creates one shape: an icon if one exists, else a basic shape.

    entity may hold "color" (a COLOR_RGB name) and "shape_type" (an AutoShape constant).
    """
    left = DEFAULT_X + index * X_STEP
    rgb = COLOR_RGB.get(entity.get("color"), SHAPE_RGB)

    # A named geometric shape ("circle", "star"...) is an explicit request for
    # that shape, so only look for an icon when none was named.
    icon_name = ENTITY_ICON_ALIASES.get(shape, shape)
    icon = None if "shape_type" in entity else find_icon(icon_name)
    if entity.get("position") == "left":
        left = LEFT_X
    elif entity.get("position") == "right":
        left = RIGHT_X
    elif entity.get("position") == "offscreen_left":
        left = -80
    elif entity.get("position") == "offscreen_right":
        left = 900

    if icon:
        log.info("Using icon %s for %r", os.path.basename(icon), shape)
        tinted = "color" in entity and icon.lower().endswith(".svg")
        if "color" in entity and not tinted:
            log.warning("Cannot tint %s; keeping its original colour.", os.path.basename(icon))
        return generate_vba(ADD_ICON, {
            "SUBNAME": sub_name,
            "SLIDEINDEX": 1,
            "ICONFILE": tint_svg(icon, rgb) if tinted else icon,  # absolute path
            "WIDTH": SHAPE_SIZE,
            "HEIGHT": SHAPE_SIZE,
            "POSLEFT": left,
            "POSTOP": DEFAULT_Y,
            "ROTATION": 0,
            "FILLR": rgb[0], "FILLG": rgb[1], "FILLB": rgb[2],
            "APPLYCOLOR": tinted,
            "SHAPENAME": shape,
        })

    log.info("No icon for %r; drawing a %s", shape, entity.get("shape_type", DEFAULT_SHAPE_TYPE))
    return generate_vba(ADD_SHAPE, {
        "SUBNAME": sub_name,
        "SLIDEINDEX": 1,
        "SHAPETYPE": entity.get("shape_type", DEFAULT_SHAPE_TYPE),
        "SIZE": SHAPE_SIZE,
        "POSLEFT": left,
        "POSTOP": DEFAULT_Y,
        "FILLR": rgb[0], "FILLG": rgb[1], "FILLB": rgb[2],
        "SHAPENAME": shape,
    })


def _action_vba(sub_name: str, action: dict, rgb: tuple) -> str:
    """Pick the right template for one action and generate its VBA."""
    common = {
        "SUBNAME": sub_name,
        "SLIDEINDEX": 1,
        "SHAPENAME": action["shape"],
        "EFFECTCONST": action["effect"],
        "TRIGGER": action["trigger"],
        "DURATION": action.get("duration", EFFECT_DURATION),
        "DELAY": 0,
        "DIRECTION": "msoAnimDirectionNone",
    }
    is_exit = bool(action.get("exit", False))

    if action.get("destination_entity"):
        return generate_vba(ADD_MOTION_PATH, {
            **common,
            "EFFECTCONST": "msoAnimEffectCustom",
            "ISCUSTOM": True,
            "TOX": 0,
            "TOY": 0,
            "REPEATCOUNT": 1,
            "OBSTACLE": "",
            "TARGET": action["destination_entity"],
        })

    if action.get("destination"):
        destination = action["destination"]
        current_x = (
            "(oShape.Left + oShape.Width / 2) / "
            "oSlide.Parent.PageSetup.SlideWidth * 100"
        )
        current_y = (
            "(oShape.Top + oShape.Height / 2) / "
            "oSlide.Parent.PageSetup.SlideWidth * 100"
        )
        destinations = {
            "right": ("85", current_y),
            "left": ("5", current_y),
            "up": (current_x, "5"),
            "down": (
                current_x,
                "(oSlide.Parent.PageSetup.SlideHeight - oShape.Height / 2) / "
                "oSlide.Parent.PageSetup.SlideWidth * 100",
            ),
            "center": (
                "50",
                "oSlide.Parent.PageSetup.SlideHeight / 2 / "
                "oSlide.Parent.PageSetup.SlideWidth * 100",
            ),
            "offscreen_right": ("110", current_y),
            "offscreen_left": ("-10", current_y),
            "offscreen_up": (current_x, "-10"),
            "offscreen_down": (
                current_x,
                "(oSlide.Parent.PageSetup.SlideHeight + oShape.Height) / "
                "oSlide.Parent.PageSetup.SlideWidth * 100",
            ),
        }
        to_x, to_y = destinations[destination]
        return generate_vba(ADD_MOTION_PATH, {
            **common,
            "EFFECTCONST": "msoAnimEffectCustom",
            "ISCUSTOM": True,
            "TOX": to_x,
            "TOY": to_y,
            "REPEATCOUNT": 1,
            "OBSTACLE": "",
            "TARGET": "",
        })

    if action.get("motion") == "arc":
        return generate_vba(ADD_MOTION_PATH, {
            **common,
            "EFFECTCONST": "msoAnimEffectCustom",
            "ISCUSTOM": True,
            "TOX": 0,
            "TOY": 0,
            "REPEATCOUNT": 1,
            "OBSTACLE": action["related_shapes"][0],
            "TARGET": "",
        })

    if is_exit or action["effect"] not in EMPHASIS_EFFECTS:
        # Entrance or exit: AddExovital has no amount setting.
        if "amount" in action:
            log.warning("Ignoring amount=%s for %s (not an emphasis effect).",
                        action["amount"], action["effect"])
        return generate_vba(ADD_ENTRANCE_EXIT, {**common, "ISEXIT": is_exit})

    # Emphasis: AddEndovital normally restyles the shape; RESTYLE=False leaves
    # the shape (or icon) looking as created. COLOR1/COLOR2 are the effect's
    # colours (only used by colour effects), set to the shape's colour. AMOUNT is optional.
    default_color = (0, 0, 0) if not action.get("color") and find_icon(action["shape"]) else rgb
    color1 = COLOR_RGB.get(action.get("from_color"), default_color)
    color2 = COLOR_RGB.get(action.get("to_color"), rgb)
    params = {
        **common,
        "RESTYLE": False,
        "COLOR1R": color1[0], "COLOR1G": color1[1], "COLOR1B": color1[2],
        "COLOR2R": color2[0], "COLOR2G": color2[1], "COLOR2B": color2[2],
    }
    if "amount" in action:
        params["AMOUNT"] = action["amount"]
    if "repeat_count" in action:
        params["REPEATCOUNT"] = action["repeat_count"]
    vba = generate_vba(ADD_EMPHASIS, params)
    if action.get("speech"):
        escaped = action["speech"].replace('"', '""')
        speech_code = (
            "\nDim oSpeech As Shape\n"
            "Set oSpeech = oSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, "
            "oShape.Left, oShape.Top - 70, 260, 60)\n"
            f'oSpeech.TextFrame.TextRange.Text = "\"{escaped}\""\n'
            "With oSpeech.TextFrame\n"
            "    .MarginLeft = 8\n"
            "    .MarginRight = 8\n"
            "    .VerticalAnchor = msoAnchorMiddle\n"
            "End With\n"
            "With oSpeech.Fill\n"
            "    .Visible = msoTrue\n"
            "    .ForeColor.RGB = RGB(255, 255, 255)\n"
            "End With\n"
            "oSpeech.Line.ForeColor.RGB = RGB(40, 40, 40)\n"
        )
        before, separator, after = vba.rpartition("End Sub")
        if separator:
            vba = before + speech_code + "\nEnd Sub" + after
    return vba


def build_vba(actions: list) -> tuple:
    """Generate all VBA for the actions.

    Returns (combined_vba, sub_names). sub_names is in run order: first one Sub
    per unique shape, then one Sub per action. Every Sub name carries a running
    number, so names are unique even when the same shape is animated twice.
    """
    parts, sub_names = [], []

    # One entry per unique shape, in order of first mention. The first mention
    # that names a colour / shape type decides it for later mentions too.
    entities = {}
    for action in actions:
        entity = entities.setdefault(action["shape"], {})
        for key in ("color", "shape_type", "position"):
            if key in action:
                entity.setdefault(key, action[key])
        for related_shape in action.get("related_shapes", []):
            related = entities.setdefault(related_shape, {})
            for key, value in action.get("related_entity_data", {}).get(related_shape, {}).items():
                related.setdefault(key, value)
        destination_entity = action.get("destination_entity")
        if destination_entity:
            target = entities.setdefault(destination_entity, {})
            for key, value in action.get("related_entity_data", {}).get(destination_entity, {}).items():
                target.setdefault(key, value)

    for i, (shape, entity) in enumerate(entities.items()):
        sub = f"AddShape{i + 1:02d}_{_identifier(shape)}"
        log.info("Creating shape %r (Sub %s)", shape, sub)
        parts.append(_entity_vba(sub, shape, i, entity))
        sub_names.append(sub)

    for n, action in enumerate(actions, start=1):
        sub = f"Anim{n:02d}_{_identifier(action['shape'])}"
        log.info("Action %d: %s %s (%s)", n, action["shape"], action["effect"], action["trigger"])
        rgb = COLOR_RGB.get(entities[action["shape"]].get("color"), SHAPE_RGB)
        parts.append(_action_vba(sub, action, rgb))
        sub_names.append(sub)

    return "\n".join(parts), sub_names


def _export_powerpoint_video(presentation, output_path: str) -> None:
    """Export the presentation's animated slides to MP4 using PowerPoint."""
    sequence = presentation.Slides(1).TimeLine.MainSequence
    for index in range(1, sequence.Count + 1):
        timing = sequence(index).Timing
        if int(timing.TriggerType) == PP_ANIM_TRIGGER_ON_CLICK:
            # Video export cannot receive slide-show clicks. Play these effects
            # automatically while preserving the original PPTX saved earlier.
            timing.TriggerType = PP_ANIM_TRIGGER_AFTER_PREVIOUS

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    log.info("Exporting MP4 with PowerPoint: %s", output_path)
    presentation.CreateVideo(
        output_path,
        True,
        VIDEO_DEFAULT_SLIDE_DURATION,
        VIDEO_VERTICAL_RESOLUTION,
        VIDEO_FRAMES_PER_SECOND,
        VIDEO_QUALITY,
    )

    deadline = time.monotonic() + VIDEO_EXPORT_TIMEOUT_SECONDS
    while True:
        status = int(presentation.CreateVideoStatus)
        if status == PP_MEDIA_TASK_STATUS_DONE:
            break
        if status == PP_MEDIA_TASK_STATUS_FAILED:
            raise RuntimeError("PowerPoint reported that video export failed")
        if status not in (PP_MEDIA_TASK_STATUS_IN_PROGRESS, PP_MEDIA_TASK_STATUS_QUEUED):
            if time.monotonic() >= deadline:
                raise RuntimeError(
                    f"PowerPoint video export did not finish (status {status})"
                )
        elif time.monotonic() >= deadline:
            raise RuntimeError(
                f"PowerPoint video export timed out after {VIDEO_EXPORT_TIMEOUT_SECONDS} seconds"
            )
        time.sleep(0.5)

    if not os.path.isfile(output_path) or os.path.getsize(output_path) == 0:
        raise RuntimeError("PowerPoint reported video export complete but no MP4 was created")


def _run_in_powerpoint(
    vba: str,
    sub_names: list,
    output_path: str,
    expected_effects: int,
    record_output_path: str | None = None,
) -> None:
    ppt = None
    presentation = None
    try:
        log.info("Starting PowerPoint...")
        ppt = win32com.client.Dispatch("PowerPoint.Application")
        ppt.Visible = True
        presentation = ppt.Presentations.Add(WithWindow=True)
        presentation.Slides.Add(1, PP_LAYOUT_BLANK)

        log.info("Injecting %d lines of VBA...", vba.count("\n") + 1)
        component = presentation.VBProject.VBComponents.Add(VB_STD_MODULE)
        component.CodeModule.AddFromString(vba)

        for i, sub in enumerate(sub_names, start=1):
            log.info("Running %d/%d: %s", i, len(sub_names), sub)
            ppt.Run(f"{presentation.Name}!{sub}")

        found = presentation.Slides(1).TimeLine.MainSequence.Count
        if found != expected_effects:
            log.warning("Expected %d animation effects but the slide has %d.", expected_effects, found)
        else:
            log.info("Slide has %d animation effects, as expected.", found)

        os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
        # .pptx can't hold macros, so silence PowerPoint's warning about that.
        ppt.DisplayAlerts = PP_ALERTS_NONE
        is_pptm = output_path.lower().endswith(".pptm")
        presentation.SaveAs(output_path, PP_SAVE_PPTM if is_pptm else PP_SAVE_PPTX)
        log.info("Saved %s", output_path)

        if record_output_path:
            _export_powerpoint_video(presentation, record_output_path)
            log.info("Saved PowerPoint video %s", record_output_path)
    finally:
        # Always clean up, even after an error.
        if presentation is not None:
            try:
                presentation.Saved = True
                presentation.Close()
            except pywintypes.com_error:
                pass
        if ppt is not None:
            try:
                # Dispatch can attach to a PowerPoint the user already has open,
                # so only quit when nothing else is open.
                if ppt.Presentations.Count == 0:
                    ppt.Quit()
                    log.info("PowerPoint closed.")
            except pywintypes.com_error:
                pass


def run_pipeline(
    user_prompt: str,
    output_path: str,
    use_llm: bool = False,
    record_output_path: str | None = None,
) -> bool:
    """Prompt in, animated presentation out. Returns True on success.

    The shared deterministic parser is used by default, matching the web API.
    Set use_llm=True to opt into Ollama fallback for prompts outside its
    vocabulary. When record_output_path is provided, PowerPoint exports the
    presentation directly to a video file.
    Errors are logged (not raised) and the function returns False.
    """
    output_path = os.path.abspath(output_path)
    if record_output_path:
        record_output_path = os.path.abspath(record_output_path)
    try:
        log.info("Prompt: %s", user_prompt)
        if len(user_prompt) > MAX_PROMPT_LENGTH:
            log.error("Prompt must be no longer than %d characters.", MAX_PROMPT_LENGTH)
            return False
        actions = parse_prompt_auto(user_prompt, use_llm=use_llm)
        if not actions:
            log.error("No animations found in the prompt; nothing to do.")
            return False
        if len(actions) > MAX_ACTIONS:
            log.error("Please limit your prompt to %d actions.", MAX_ACTIONS)
            return False
        log.info("Parsed %d action(s).", len(actions))

        vba, sub_names = build_vba(actions)
        _run_in_powerpoint(
            vba,
            sub_names,
            output_path,
            expected_effects=len(actions),
            record_output_path=record_output_path,
        )
        log.info("Done.")
        return True
    except pywintypes.com_error as exc:
        log.error("PowerPoint error: %s", exc)
        log.error("If it mentions the VBA project, enable 'Trust access to the VBA "
                  "project object model' in PowerPoint's Trust Center.")
    except (ValueError, OSError) as exc:
        log.error("%s: %s", type(exc).__name__, exc)
    except RuntimeError as exc:
        log.error("Video export failed: %s", exc)
    return False


def main() -> int:
    parser = argparse.ArgumentParser(description="Turn a text prompt into an animated PowerPoint.")
    parser.add_argument("prompt", help='e.g. "A fox appears, then spins 180 degrees."')
    parser.add_argument("output", nargs="?", default="pipeline_output.pptx",
                        help="output file (.pptx or .pptm); default: %(default)s")
    llm_group = parser.add_mutually_exclusive_group()
    llm_group.add_argument(
        "--llm", dest="use_llm", action="store_true",
        help="enable Ollama fallback for prompts the shared rule-based parser cannot handle",
    )
    llm_group.add_argument(
        "--no-llm", dest="use_llm", action="store_false",
        help="use the same deterministic parser as the web API (default)",
    )
    parser.set_defaults(use_llm=False)
    parser.add_argument("--record", action="store_true",
                        help="export the presentation as an MP4 beside the PowerPoint file")
    args = parser.parse_args()

    logging.basicConfig(level=logging.INFO, format="%(levelname)-7s %(message)s")
    logging.captureWarnings(True)  # show parser/assembler warnings in the log too
    record_output_path = (
        os.path.splitext(args.output)[0] + ".mp4" if args.record else None
    )
    return 0 if run_pipeline(
        args.prompt,
        args.output,
        use_llm=args.use_llm,
        record_output_path=record_output_path,
    ) else 1


if __name__ == "__main__":
    sys.exit(main())
