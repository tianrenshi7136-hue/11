import copy
import tempfile
import unittest
from pathlib import Path

from els.core import ProjectStore, digest, export_project, import_project, new_project, validate
from els.templates import check_manifest, readiness


def valid_project():
    p = new_project()
    p.update(length_m=20, width_m=10, depth_m=5, embedment_m=4, analysis_reference="TEST-ONLY, revision A")
    p["members"] = [{"member_id": "S1", "kind": "支撑", "case": "stage1", "x_m": 5, "y_m": 0, "level_m": -2, "effective_length_m": 10, "n_kn": 100, "m_knm": 0, "v_kn": 0, "value_basis": "设计值", "force_basis": "单根"}]
    return p


class DataTests(unittest.TestCase):
    def test_valid_input_does_not_mean_engineering_pass(self):
        p = valid_project()
        self.assertEqual(validate(p), [])
        self.assertTrue(readiness(p))

    def test_nonfinite_and_wrong_units(self):
        p = valid_project()
        p["members"][0]["n_kn"] = float("nan")
        p["members"][0]["force_basis"] = "每延米"
        self.assertGreaterEqual(len(validate(p)), 2)
        with self.assertRaises(ValueError):
            digest(p)

    def test_duplicate_case(self):
        p = valid_project()
        p["members"].append(copy.deepcopy(p["members"][0]))
        self.assertTrue(any("重复" in e for e in validate(p)))

    def test_inconsistent_geometry(self):
        p = valid_project()
        m = copy.deepcopy(p["members"][0])
        m.update(case="stage2", effective_length_m=8)
        p["members"].append(m)
        self.assertTrue(any("不一致" in e for e in validate(p)))

    def test_bounds_and_unknown_force_basis(self):
        p = valid_project()
        p["members"][0].update(x_m=30, level_m=-7, value_basis="")
        self.assertEqual(len(validate(p)), 3)

    def test_sqlite_history_and_backup(self):
        with tempfile.TemporaryDirectory() as folder:
            store = ProjectStore(Path(folder) / "projects.db")
            p = valid_project()
            self.assertEqual(store.save(p), 1)
            self.assertEqual(store.save(p), 1)
            original = copy.deepcopy(p)
            p["members"][0]["n_kn"] = 200
            self.assertEqual(store.save(p), 2)
            self.assertEqual(store.load(p["id"], 1), original)
            target = Path(folder) / "backup.db"
            store.backup(target)
            self.assertEqual(ProjectStore(target).load(p["id"], 2), p)

    def test_roundtrip_and_invalid_import(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "project.json"
            p = valid_project()
            export_project(p, path)
            self.assertEqual(import_project(path), p)
            path.write_text('{"schema_version": 1}', encoding="utf-8")
            with self.assertRaises(ValueError):
                import_project(path)

    def test_draft_is_incomplete(self):
        self.assertTrue(validate(new_project()))

    def test_unapproved_template_is_blocked(self):
        _, problems = check_manifest(Path(__file__).resolve().parent.parent / "examples" / "template-manifest.json")
        self.assertTrue(any("审核" in p for p in problems))
        self.assertTrue(any("文件不存在" in p for p in problems))


if __name__ == "__main__":
    unittest.main()
