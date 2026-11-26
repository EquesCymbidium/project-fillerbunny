"""
dispatcher.py

Turns a plain-English prompt into a list of animation steps (no LLM, just
keyword matching), ready to be handed to the VBA templates later.

    >>> parse_prompt("A fox appears, then spins 180 degrees.")
    [{'shape': 'fox', 'effect': 'msoAnimEffectFade',
      'trigger': 'msoAnimTriggerAfterPrevious'},
     {'shape': 'fox', 'effect': 'msoAnimEffectSpin', 'amount': 180,
      'trigger': 'msoAnimTriggerAfterPrevious'}]
"""

import functools
import os
import re
import warnings

# Verb (as it may appear in a prompt) -> PowerPoint effect constant.
# To support a new verb, add one line here.
VERB_TO_EFFECT = {
    "appear": "msoAnimEffectFade",
    "appears": "msoAnimEffectFade",
    "fade": "msoAnimEffectFade",
    "fades": "msoAnimEffectFade",
    "emerge": "msoAnimEffectFade",
    "emerges": "msoAnimEffectFade",
    "enter": "msoAnimEffectFade",
    "enters": "msoAnimEffectFade",
    "disappear": "msoAnimEffectFade",
    "disappears": "msoAnimEffectFade",
    "disappeared": "msoAnimEffectFade",
    "disappearing": "msoAnimEffectFade",
    "vanish": "msoAnimEffectFade",
    "vanishes": "msoAnimEffectFade",
    "vanished": "msoAnimEffectFade",
    "vanishing": "msoAnimEffectFade",
    "spin": "msoAnimEffectSpin",
    "spins": "msoAnimEffectSpin",
    "rotate": "msoAnimEffectSpin",
    "rotates": "msoAnimEffectSpin",
    "twirl": "msoAnimEffectSpin",
    "twirls": "msoAnimEffectSpin",
}
EXIT_VERBS = {
    "disappear", "disappears", "disappeared", "disappearing",
    "vanish", "vanishes", "vanished", "vanishing",
}
ENTITY_MODIFIERS = {"quick", "fast", "slow", "lazy"}
SPEED_FACTORS = {"quick": 0.5, "fast": 0.5, "slow": 1.5, "lazy": 1.5}

_PROJECT_ROOT = os.path.normpath(
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..")
)
_ICON_DIR = os.path.join(_PROJECT_ROOT, "Library", "Icons")

TRIGGER_NEXT = "msoAnimTriggerAfterPrevious"
TRIGGER_FIRST = TRIGGER_NEXT
EFFECT_DURATION = 1.0
MAX_PROMPT_LENGTH = 2000
MAX_ACTIONS = 10

# Colour word -> (R, G, B). Add a line to support a new colour.
COLOR_RGB = {
    "red": (204, 0, 0),
    "orange": (255, 140, 0),
    "yellow": (255, 204, 0),
    "green": (0, 153, 51),
    "blue": (0, 102, 204),
    "purple": (128, 0, 160),
    "pink": (255, 105, 180),
    "brown": (139, 69, 19),
    "black": (0, 0, 0),
    "white": (255, 255, 255),
    "gray": (128, 128, 128),
    "grey": (128, 128, 128),
    "teal": (0, 128, 128),
    "gold": (212, 175, 55),
}

# Geometric shape word -> PowerPoint AutoShape constant.
# (AddShape draws equal-sided shapes only, so there is no "rectangle".)
SHAPE_TYPES = {
    "square": "msoShapeRectangle",
    "circle": "msoShapeOval",
    "oval": "msoShapeOval",
    "triangle": "msoShapeIsoscelesTriangle",
    "diamond": "msoShapeDiamond",
    "star": "msoShape5pointStar",
    "heart": "msoShapeHeart",
    "hexagon": "msoShapeHexagon",
    "pentagon": "msoShapeRegularPentagon",
    "cross": "msoShapeCross",
    "arrow": "msoShapeRightArrow",
    "moon": "msoShapeMoon",
    "cloud": "msoShapeCloud",
}

# Natural-language names that should resolve to available icon assets.
ENTITY_ALIASES = {
    "human": "human",
    "humans": "human",
    "person": "human",
    "persons": "human",
    "people": "group_of_people",
}
ENTITY_ICON_ALIASES = {"human": "man"}
_BUILTIN_ENTITY_NAMES = (
    "apple", "bear", "bird", "car", "cat", "dog", "fish", "fox",
    "horse", "rabbit", "robot", "tree",
)

# Words that never name a shape: articles, possessives and pronouns.
# A pronoun ("It spins") means "same shape as the previous step".
_NOT_A_SHAPE = {
    "a", "an", "the", "this", "that", "my", "our", "its", "his", "her", "their",
    "it", "he", "she", "they",
}

# Step separators: punctuation, or "then" / "and then" / "after that" / "next".
# "and then" is listed before "then" so the whole phrase is consumed.
_SPLIT_RE = re.compile(r"[.,;!?]|\b(?:and then|then|after that|next)\b", re.IGNORECASE)
_WORD_RE = re.compile(r"[A-Za-z]+")
_TOKEN_RE = re.compile(r"[A-Za-z]+")

_VERB_FORMS = {
    "appear": ("appear", "appears", "appeared", "appearing",
               "emerge", "emerges", "emerged", "emerging",
               "enter", "enters", "entered", "entering"),
    "disappear": ("disappear", "disappears", "disappeared", "disappearing",
                  "vanish", "vanishes", "vanished", "vanishing"),
    "spin": ("spin", "spins", "spun", "spinning",
             "rotate", "rotates", "rotated", "rotating",
             "twirl", "twirls", "twirled", "twirling"),
    "move": ("move", "moves", "moved", "moving",
             "run", "runs", "ran", "running",
             "walk", "walks", "walked", "walking"),
    "jump": ("jump", "jumps", "jumped", "jumping",
             "leap", "leaps", "leaped", "leaping",
             "hop", "hops", "hopped", "hopping"),
    "turn": ("turn", "turns", "turned", "turning"),
}
_ACTION_RE = re.compile(
    r"\b(" + "|".join(
        re.escape(form)
        for forms in sorted(_VERB_FORMS.values(), key=len, reverse=True)
        for form in sorted(forms, key=len, reverse=True)
    ) + r")\b",
    re.IGNORECASE,
)
_NUMBER_WORDS = {
    "one": 1, "two": 2, "three": 3, "four": 4, "five": 5,
    "six": 6, "seven": 7, "eight": 8, "nine": 9, "ten": 10,
}
_ADVERB_SPEED = {
    "quickly": 0.5, "rapidly": 0.5, "fast": 0.5,
    "slowly": 1.5, "lazily": 1.5,
}

# "180 degrees", "90 deg" or "45°"
_DEGREES_RE = re.compile(r"(\d+(?:\.\d+)?)\s*(?:°|\bdeg(?:rees?)?\b)", re.IGNORECASE)
# Spin direction: PowerPoint treats negative amounts as counterclockwise.
_COUNTERCLOCKWISE_RE = re.compile(
    r"\b(?:counter-?\s?clockwise|anti-?\s?clockwise|ccw)\b", re.IGNORECASE
)


def _split_steps(prompt: str) -> list:
    """Cut the prompt into one text chunk per action, dropping empty chunks."""
    return [chunk.strip() for chunk in _SPLIT_RE.split(prompt) if chunk.strip()]


@functools.lru_cache(maxsize=1)
def _icon_entity_names() -> tuple:
    """Return canonical icon names, excluding alternate ``_M`` artwork."""
    names = set()
    try:
        filenames = os.listdir(_ICON_DIR)
    except OSError:
        filenames = ()
    for filename in filenames:
        stem, extension = os.path.splitext(filename)
        if extension.lower() not in (".svg", ".png") or stem.lower().endswith("_m"):
            continue
        words = re.sub(r"([a-z])([A-Z])", r"\1 \2", stem).replace("_", " ")
        tokens = tuple(word.lower() for word in _TOKEN_RE.findall(words))
        if tokens:
            names.add(("_".join(tokens), tokens))
    names.update((name, (name,)) for name in _BUILTIN_ENTITY_NAMES)
    return tuple(sorted(names, key=lambda item: (-len(item[1]), item[0])))


def _entity_mentions(prompt: str) -> list:
    """Find explicit noun mentions using names from the icon library."""
    tokens = list(_TOKEN_RE.finditer(prompt))
    words = [match.group().lower() for match in tokens]
    candidates = list(_icon_entity_names())
    candidates.extend((name, (name,)) for name in SHAPE_TYPES)
    candidates.extend((canonical, (alias,)) for alias, canonical in ENTITY_ALIASES.items())
    candidates.sort(key=lambda item: (-len(item[1]), item[0]))

    mentions = []
    occupied = set()
    for canonical, phrase in candidates:
        width = len(phrase)
        for start in range(len(words) - width + 1):
            end = start + width
            if any(index in occupied for index in range(start, end)):
                continue
            if tuple(words[start:end]) != phrase:
                continue
            if width == 1 and end < len(words) and words[end] in ("s", "es"):
                end += 1
            mentions.append({
                "name": canonical,
                "start": tokens[start].start(),
                "end": tokens[end - 1].end(),
                "word_start": start,
                "word_end": end,
            })
            occupied.update(range(start, end))
    return sorted(mentions, key=lambda mention: mention["start"])


def _modifier_before(words: list, start: int) -> tuple:
    """Return explicitly named color/shape/speed modifiers before a noun."""
    prefix = words[max(0, start - 3):start]
    color = next((word for word in reversed(prefix) if word in COLOR_RGB), None)
    shape = next((word for word in reversed(prefix) if word in SHAPE_TYPES), None)
    speed = next((SPEED_FACTORS[word] for word in reversed(prefix)
                  if word in SPEED_FACTORS), None)
    return color, shape, speed


def _event_count(text: str) -> int:
    match = re.search(
        r"\b(?P<count>\d+|" + "|".join(_NUMBER_WORDS) + r")\s+times?\b",
        text.lower(),
    )
    if not match:
        return 1
    value = match.group("count")
    return int(value) if value.isdigit() else _NUMBER_WORDS[value]


def _movement_destination(text: str) -> str:
    words = set(_TOKEN_RE.findall(text.lower()))
    if "offscreen" in words or {"off", "screen"} <= words or "across" in words:
        if "left" in words:
            return "offscreen_left"
        if "up" in words:
            return "offscreen_up"
        if "down" in words:
            return "offscreen_down"
        return "offscreen_right"
    if "middle" in words or "center" in words or "centre" in words:
        return "center"
    for direction in ("right", "left", "up", "down"):
        if direction in words:
            return direction
    return "right"


def _position_for_mention(prompt: str, mention: dict, next_mention: dict | None) -> str | None:
    """Find an explicit left/right placement following a noun mention."""
    end = next_mention["start"] if next_mention else len(prompt)
    context = prompt[mention["end"]:end]
    match = re.search(
        r"\bon\s+the\s+(left|right)\b|\b(left|right)\s+side\b",
        context,
        re.IGNORECASE,
    )
    if not match:
        return None
    return (match.group(1) or match.group(2)).lower()


def _movement_target(text: str, event_end: int, tail_end: int, mentions: list) -> str | None:
    """Find an explicitly named destination in 'to/toward [entity]' phrases."""
    for mention in mentions:
        if not event_end <= mention["start"] < tail_end:
            continue
        prefix = text[event_end:mention["start"]]
        if re.search(r"\b(?:to|toward|towards)\s+(?:(?:the|a|an)\s+)?$", prefix, re.IGNORECASE):
            return mention["name"]
    return None


def _extended_parse(prompt: str) -> list:
    """Parse common multi-entity actions, movement, color changes, and modifiers."""
    tokens = list(_TOKEN_RE.finditer(prompt))
    words = [match.group().lower() for match in tokens]
    mentions = _entity_mentions(prompt)
    events = list(_ACTION_RE.finditer(prompt))
    if not mentions or not events:
        return []

    entity_data = {}
    for mention_index, mention in enumerate(mentions):
        data = entity_data.setdefault(
            mention["name"],
            {"shape": mention["name"], "adjectives": []},
        )
        if mention["name"] in SHAPE_TYPES:
            data["shape_type"] = SHAPE_TYPES[mention["name"]]
        color, shape, speed = _modifier_before(words, mention["word_start"])
        if color:
            data.setdefault("color", color)
        if shape:
            data.setdefault("shape_type", SHAPE_TYPES[shape])
        prefix = words[max(0, mention["word_start"] - 4):mention["word_start"]]
        data["adjectives"].extend(
            word for word in prefix if word in ENTITY_MODIFIERS and word not in data["adjectives"]
        )
        if speed is not None:
            data["speed"] = speed
        position = _position_for_mention(
            prompt,
            mention,
            mentions[mention_index + 1] if mention_index + 1 < len(mentions) else None,
        )
        if position:
            data.setdefault("position", position)

    actions = []
    current_actor = None
    current_colors = {
        name: data.get("color")
        for name, data in entity_data.items()
    }
    previous_event_end = 0
    for index, event in enumerate(events):
        between = [
            mention for mention in mentions
            if previous_event_end <= mention["start"] < event.start()
        ]
        subjects = [
            mention for mention in between
            if not re.search(
                r"\b(over|above|past|around|to|toward|towards|with|for|by|near)\s+$",
                prompt[max(previous_event_end, mention["start"] - 16):mention["start"]],
                re.IGNORECASE,
            )
        ]
        if subjects:
            current_actor = subjects[-1]["name"]
        if current_actor is None:
            prior = [mention for mention in mentions if mention["end"] <= event.start()]
            if prior:
                current_actor = prior[-1]["name"]
        if current_actor is None:
            warnings.warn(f"Could not identify the subject for {event.group()!r}.", stacklevel=2)
            previous_event_end = event.end()
            continue

        verb = event.group().lower()
        kind = next(name for name, forms in _VERB_FORMS.items() if verb in forms)
        tail_end = events[index + 1].start() if index + 1 < len(events) else len(prompt)
        tail = prompt[event.end():tail_end]
        action = {"shape": current_actor}
        action.update({
            key: entity_data[current_actor][key]
            for key in ("color", "shape_type", "position")
            if key in entity_data[current_actor]
        })

        if kind == "appear":
            action["effect"] = "msoAnimEffectFade"
        elif kind == "disappear":
            action["effect"] = "msoAnimEffectFade"
            action["exit"] = True
        elif kind == "spin":
            action["effect"] = "msoAnimEffectSpin"
            degrees = _DEGREES_RE.search(tail)
            if degrees:
                amount = float(degrees.group(1))
                action["amount"] = int(amount) if amount.is_integer() else amount
                if _COUNTERCLOCKWISE_RE.search(tail):
                    action["amount"] = -action["amount"]
        elif kind == "turn":
            color = next((word for word in _TOKEN_RE.findall(tail.lower())
                          if word in COLOR_RGB), None)
            if color:
                action["effect"] = "msoAnimEffectChangeFillColor"
                action["from_color"] = current_colors.get(current_actor)
                action["to_color"] = color
                current_colors[current_actor] = color
            else:
                action["effect"] = "msoAnimEffectSpin"
        elif kind == "move":
            action["effect"] = "msoAnimEffectCustom"
            target = _movement_target(prompt, event.end(), tail_end, mentions)
            if target and target != current_actor:
                action["destination_entity"] = target
            else:
                action["motion"] = _movement_destination(tail)
        else:
            over_match = re.search(r"\bover\b", tail, re.IGNORECASE)
            obstacle = None
            if over_match:
                obstacle = next((
                    mention["name"] for mention in mentions
                    if event.end() + over_match.end() <= mention["start"] < tail_end
                    and mention["name"] != current_actor
                ), None)
            if obstacle:
                action["effect"] = "msoAnimEffectCustom"
                action["motion"] = "arc"
                action["related_shapes"] = [obstacle]
            else:
                action["effect"] = "msoAnimEffectBounce"
                action["repeat_count"] = _event_count(tail)

        related_entities = list(action.get("related_shapes", []))
        if action.get("destination_entity"):
            related_entities.append(action["destination_entity"])
        if related_entities:
            action["related_entity_data"] = {
                name: {
                    key: entity_data[name][key]
                    for key in ("color", "shape_type", "position")
                    if key in entity_data[name]
                }
                for name in related_entities
                if name in entity_data
            }

        speed = entity_data[current_actor].get("speed", 1.0)
        for word, factor in _ADVERB_SPEED.items():
            if re.search(rf"\b{word}\b", prompt[max(previous_event_end, event.start() - 24):tail_end], re.IGNORECASE):
                speed = factor
                break
        action["duration"] = max(0.25, EFFECT_DURATION * speed)

        if kind == "move" and "motion" in action:
            action["destination"] = action.pop("motion")
        elif kind == "jump" and action.get("motion") != "arc":
            action.pop("motion", None)
        action["trigger"] = TRIGGER_NEXT if actions else TRIGGER_FIRST
        actions.append(action)
        previous_event_end = event.end()

    return actions


def _parse_step(chunk: str, previous_shape):
    """Parse one chunk into a dict, or return None if it has no known verb.

    Keys: shape, effect, and (only when found) amount, color, shape_type.
    """
    # Find the first word that is a known verb; everything before it is the subject.
    words = _WORD_RE.findall(chunk)
    verb_index = next(
        (i for i, word in enumerate(words) if word.lower() in VERB_TO_EFFECT), None
    )
    if verb_index is None:
        return None
    verb = words[verb_index].lower()
    effect = VERB_TO_EFFECT[verb]

    subject_words = [w.lower() for w in words[:verb_index]]

    # Description words in the subject: "a red circle" -> color red, shape circle.
    color = next((w for w in subject_words if w in COLOR_RGB), None)
    shape_type = next((SHAPE_TYPES[w] for w in subject_words if w in SHAPE_TYPES), None)

    # Shape name: "A red fox" -> "fox". Colours are not part of the name, so
    # "the fox" later in the prompt still means the same shape. No subject or
    # a pronoun -> reuse the previous shape.
    name_words = [w for w in subject_words if w not in _NOT_A_SHAPE and w not in COLOR_RGB]
    shape = "_".join(name_words) if name_words else previous_shape
    if shape is None:
        raise ValueError(f"No shape named before the verb in: {chunk!r}")

    step = {"shape": shape}
    if shape_type is not None:
        step["shape_type"] = shape_type
    if color is not None:
        step["color"] = color
    step["effect"] = effect
    if verb in EXIT_VERBS:
        step["exit"] = True

    match = _DEGREES_RE.search(chunk)
    if match:
        number = float(match.group(1))
        amount = int(number) if number.is_integer() else number
        if _COUNTERCLOCKWISE_RE.search(chunk):
            amount = -amount
        step["amount"] = amount
    return step


def parse_prompt(prompt: str) -> list:
    """Parse common animation prompts, retaining the legacy keyword fallback."""
    extended = _extended_parse(prompt)
    if extended:
        return extended

    steps = []
    shape = None
    for chunk in _split_steps(prompt):
        parsed = _parse_step(chunk, shape)
        if parsed is None:
            warnings.warn(f"No known animation verb in {chunk!r}; skipped.", stacklevel=2)
            continue
        shape = parsed["shape"]
        parsed["trigger"] = TRIGGER_NEXT if steps else TRIGGER_FIRST
        steps.append(parsed)
    return steps


if __name__ == "__main__":
    import pprint

    pprint.pprint(parse_prompt("A fox appears, then spins 180 degrees."))
