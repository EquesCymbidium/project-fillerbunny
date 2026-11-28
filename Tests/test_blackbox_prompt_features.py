import importlib
import os
import sys
import unittest
from unittest.mock import patch


_PROJECT_ROOT = os.path.normpath(os.path.join(os.path.dirname(__file__), ".."))
_BLACKBOX_DIR = os.path.join(_PROJECT_ROOT, "Product", "Blackbox")
sys.path.insert(0, _BLACKBOX_DIR)

main = importlib.import_module("main")
parse_prompt = importlib.import_module("dispatcher").parse_prompt


class PromptFeatureTests(unittest.TestCase):
    def test_cli_default_uses_the_same_parser_as_the_web_api(self):
        prompt = "the fox jumps over the dog"
        web_actions = parse_prompt(prompt)
        self.assertEqual(main.parse_prompt_auto(prompt, use_llm=False), web_actions)

        with patch.object(main, "_run_in_powerpoint") as run_in_powerpoint, \
                patch.object(main, "parse_prompt_auto", wraps=main.parse_prompt_auto) as parse:
            self.assertTrue(main.run_pipeline(prompt, "unused.pptx"))

        parse.assert_called_once_with(prompt, use_llm=False)
        self.assertEqual(run_in_powerpoint.call_args.kwargs["expected_effects"], len(web_actions))
        self.assertIn(".Path = ", run_in_powerpoint.call_args.args[0])

    def test_cli_enforces_web_prompt_and_action_limits(self):
        too_many_actions = ", ".join(["the fox appears"] * (main.MAX_ACTIONS + 1))
        with patch.object(main, "_run_in_powerpoint") as run_in_powerpoint:
            self.assertFalse(main.run_pipeline(too_many_actions, "unused.pptx"))
            self.assertFalse(main.run_pipeline("x" * (main.MAX_PROMPT_LENGTH + 1), "unused.pptx"))
        run_in_powerpoint.assert_not_called()

    def test_color_adjective_selects_icon_and_tint(self):
        actions = parse_prompt("the red fox appears")

        self.assertEqual(len(actions), 1)
        self.assertEqual(actions[0]["shape"], "fox")
        self.assertEqual(actions[0]["color"], "red")
        vba, _ = main.build_vba(actions)
        self.assertIn("Fox_", vba)
        self.assertIn("204", vba)

    def test_composite_actions_move_and_change_color(self):
        actions = parse_prompt(
            "the red fox appears, moves to the right and turns orange"
        )

        self.assertEqual(len(actions), 3)
        self.assertEqual(actions[1]["destination"], "right")
        self.assertEqual(actions[2]["effect"], "msoAnimEffectChangeFillColor")
        self.assertEqual(actions[2]["from_color"], "red")
        self.assertEqual(actions[2]["to_color"], "orange")
        vba, _ = main.build_vba(actions)
        self.assertIn(".ToX = 85", vba)
        self.assertIn("msoAnimEffectChangeFillColor", vba)
        self.assertIn("RGB(255, 140, 0)", vba)

    def test_speed_adjective_and_across_screen_motion(self):
        actions = parse_prompt("the quick fox appears and runs across the screen")

        self.assertEqual(len(actions), 2)
        self.assertEqual(actions[0]["duration"], 0.5)
        self.assertEqual(actions[1]["duration"], 0.5)
        self.assertEqual(actions[1]["destination"], "offscreen_right")
        vba, _ = main.build_vba(actions)
        self.assertIn(".ToX = 110", vba)

    def test_speed_adverb_changes_action_duration(self):
        actions = parse_prompt("the fox runs quickly across the screen")

        self.assertEqual(len(actions), 1)
        self.assertEqual(actions[0]["duration"], 0.5)
        self.assertEqual(actions[0]["destination"], "offscreen_right")

    def test_center_repeated_jumps_and_offscreen_motion(self):
        actions = parse_prompt(
            "the fox appears, runs to the middle, jumps three times "
            "before running offscreen"
        )

        self.assertEqual(len(actions), 4)
        self.assertEqual(actions[1]["destination"], "center")
        self.assertEqual(actions[2]["effect"], "msoAnimEffectBounce")
        self.assertEqual(actions[2]["repeat_count"], 3)
        self.assertEqual(actions[3]["destination"], "offscreen_right")
        vba, _ = main.build_vba(actions)
        self.assertIn(".RepeatCount = 3", vba)

    def test_jump_over_identifies_both_entities(self):
        actions = parse_prompt("the quick fox jumps over the lazy dog")

        self.assertEqual(len(actions), 1)
        self.assertEqual(actions[0]["shape"], "fox")
        self.assertEqual(actions[0]["related_shapes"], ["dog"])
        self.assertEqual(actions[0]["effect"], "msoAnimEffectCustom")
        self.assertEqual(actions[0]["motion"], "arc")
        vba, sub_names = main.build_vba(actions)
        self.assertIn("AddShape01_fox", sub_names)
        self.assertIn("AddShape02_dog", sub_names)
        self.assertIn('oSlide.Shapes("dog")', vba)
        self.assertIn(".Path = ", vba)
        self.assertIn("exitFraction", vba)

    def test_jump_over_preserves_related_entity_color_and_starts_automatically(self):
        actions = parse_prompt("The quick orange fox jumped over the lazy brown dog")

        self.assertEqual(len(actions), 1)
        self.assertEqual(actions[0]["trigger"], "msoAnimTriggerAfterPrevious")
        self.assertEqual(actions[0]["color"], "orange")
        self.assertEqual(actions[0]["related_entity_data"]["dog"]["color"], "brown")
        vba, _ = main.build_vba(actions)
        self.assertIn("COLOR_R       As Integer = 255", vba)
        self.assertIn("COLOR_G       As Integer = 140", vba)
        self.assertIn("COLOR_R       As Integer = 139", vba)
        self.assertIn("COLOR_G       As Integer = 69", vba)
        self.assertIn("COLOR_B       As Integer = 19", vba)

    def test_multiple_subjects_and_targeted_movement(self):
        actions = parse_prompt(
            "A human appears on the left, the fox appears on the right; "
            "the human walks over to the fox"
        )

        self.assertEqual(len(actions), 3)
        self.assertEqual(actions[0]["shape"], "human")
        self.assertEqual(actions[0]["position"], "left")
        self.assertEqual(actions[1]["shape"], "fox")
        self.assertEqual(actions[1]["position"], "right")
        self.assertEqual(actions[2]["shape"], "human")
        self.assertEqual(actions[2]["destination_entity"], "fox")
        self.assertNotIn("destination", actions[2])

        vba, sub_names = main.build_vba(actions)
        self.assertIn("AddShape01_human", sub_names)
        self.assertIn("AddShape02_fox", sub_names)
        self.assertIn("POS_LEFT      As Single  = 120", vba)
        self.assertIn("POS_LEFT      As Single  = 740", vba)
        self.assertIn('oSlide.Shapes("fox")', vba)

    def test_person_alias_resolves_to_human_entity(self):
        actions = parse_prompt("the person appears, then the fox appears")

        self.assertEqual([action["shape"] for action in actions], ["human", "fox"])

    def test_complex_dash_and_speech_actions(self):
        actions = parse_prompt(
            'the quick orange fox dashes in from the left screen, and leaps over '
            'the brown dog; the brown dog barks: "what?!"'
        )

        self.assertEqual(len(actions), 3)
        self.assertEqual(actions[0]["shape"], "fox")
        self.assertEqual(actions[0]["position"], "offscreen_left")
        self.assertEqual(actions[0]["destination"], "left")
        self.assertEqual(actions[1]["motion"], "arc")
        self.assertEqual(actions[2]["shape"], "dog")
        self.assertEqual(actions[2]["speech"], "what?!")
        vba, _ = main.build_vba(actions)
        self.assertIn("oSpeech.TextFrame.TextRange.Text =", vba)
        self.assertIn("what?!", vba)


if __name__ == "__main__":
    unittest.main()
