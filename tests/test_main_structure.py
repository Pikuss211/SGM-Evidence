import ast
import importlib.util
import unittest
from pathlib import Path
from unittest.mock import patch


ROOT = Path(__file__).resolve().parents[1]
MAIN_PATH = ROOT / "SGM_v1.1-de_clean_fixed.py"


def load_main_module():
    spec = importlib.util.spec_from_file_location("sgm_main_for_tests", MAIN_PATH)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


class MainStructureTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.module = load_main_module()

    def test_final_grid_has_named_base_instead_of_self_inheritance(self):
        source = MAIN_PATH.read_text(encoding="utf-8-sig")
        tree = ast.parse(source)
        classes = [node for node in tree.body if isinstance(node, ast.ClassDef)]
        class_names = [node.name for node in classes]

        self.assertEqual(class_names.count("StrojeGrid"), 1)
        self.assertEqual(class_names.count("StrojeGridBase"), 1)
        final = next(node for node in classes if node.name == "StrojeGrid")
        self.assertEqual([ast.unparse(base) for base in final.bases], ["StrojeGridBase"])
        self.assertTrue(
            issubclass(self.module.StrojeGrid, self.module.StrojeGridBase)
        )

    def test_main_grid_exposes_toolbar_actions(self):
        expected = {
            "backup_zip",
            "restore_zip",
            "kontrola_dat_gui",
            "global_search_gui",
            "hromadne_uzavrit",
            "archivovat_stroj_gui",
            "obnovit_stroj_z_archivu",
            "vybrat_operatora_gui",
            "prepnout_stav_toolbar",
            "graf_top_stroje",
        }
        missing = {
            name for name in expected if not hasattr(self.module.StrojeGrid, name)
        }
        self.assertEqual(missing, set())

    def test_application_has_semantic_release_version(self):
        self.assertRegex(self.module.APP_VERSION, r"^\d+\.\d+\.\d+$")

    def test_validation_action_delegates_to_separate_ui_module(self):
        app = object()
        expected_window = object()
        with patch.object(
            self.module.validation_ui,
            "open_data_validation_dialog",
            return_value=expected_window,
        ) as open_dialog:
            result = self.module.StrojeGrid.kontrola_dat_gui(app)

        self.assertIs(result, expected_window)
        open_dialog.assert_called_once_with(app)

    def test_archived_machines_are_hidden_until_filter_is_enabled(self):
        class Value:
            def __init__(self, value):
                self.value = value

            def get(self):
                return self.value

        class FakeGrid:
            stroje = {
                "1": {"archivovan": "0"},
                "2": {"archivovan": "1"},
            }
            show_archived = Value(False)
            filter_only_problem = Value(False)
            filtr_kat = Value("vse")
            sort_mode = Value("cislo")
            poruchy = []

            def _apply_sort(self, values, *_counts):
                return sorted(values)

        fake = FakeGrid()
        method = self.module.StrojeGridBase._get_visible_machine_numbers
        self.assertEqual(method(fake, {}, {}, {}), [1])
        fake.show_archived.value = True
        self.assertEqual(method(fake, {}, {}, {}), [1, 2])


if __name__ == "__main__":
    unittest.main()
