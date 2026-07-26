import json
import unittest
from pathlib import Path

import operator_manager as om


def runtime_file(name: str) -> Path:
    directory = Path(__file__).parent / "_runtime" / name
    directory.mkdir(parents=True, exist_ok=True)
    return directory / "app_settings.json"


class OperatorManagerTests(unittest.TestCase):
    def test_remembers_last_operator_and_recent_unique_names(self):
        settings_file = runtime_file("operator_recent")
        settings_file.unlink(missing_ok=True)
        settings_file.with_suffix(".json.bak").unlink(missing_ok=True)

        om.remember_operator("  Jan   Novák  ", settings_file)
        result = om.remember_operator("Eva", settings_file)
        result = om.remember_operator("jan novák", settings_file)

        self.assertEqual(result["last_operator"], "jan novák")
        self.assertEqual(result["operators"], ["jan novák", "Eva"])
        self.assertEqual(om.initial_operator(settings_file), "jan novák")

    def test_previous_valid_settings_are_kept_as_backup(self):
        settings_file = runtime_file("operator_backup")
        settings_file.unlink(missing_ok=True)
        settings_file.with_suffix(".json.bak").unlink(missing_ok=True)

        om.remember_operator("První", settings_file)
        om.remember_operator("Druhý", settings_file)

        backup = json.loads(
            settings_file.with_suffix(".json.bak").read_text(encoding="utf-8")
        )
        self.assertEqual(backup["last_operator"], "První")

    def test_corrupt_primary_falls_back_to_backup(self):
        settings_file = runtime_file("operator_corrupt")
        backup_file = settings_file.with_suffix(".json.bak")
        settings_file.write_text("{broken", encoding="utf-8")
        backup_file.write_text(
            json.dumps(
                {"last_operator": "Záloha", "operators": ["Záloha"]},
                ensure_ascii=False,
            ),
            encoding="utf-8",
        )

        self.assertEqual(om.initial_operator(settings_file), "Záloha")

    def test_environment_is_used_when_no_settings_exist(self):
        settings_file = runtime_file("operator_environment")
        settings_file.unlink(missing_ok=True)
        settings_file.with_suffix(".json.bak").unlink(missing_ok=True)

        self.assertEqual(
            om.initial_operator(settings_file, {"USERNAME": " Technik "}),
            "Technik",
        )


if __name__ == "__main__":
    unittest.main()
