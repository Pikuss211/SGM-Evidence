import json
import unittest
from pathlib import Path

import audit_manager as audit


def runtime_file(name: str) -> Path:
    root = Path(__file__).parent / "_runtime" / name
    root.mkdir(parents=True, exist_ok=True)
    path = root / "audit_log.csv"
    path.unlink(missing_ok=True)
    path.with_suffix(".csv.bak").unlink(missing_ok=True)
    return path


class AuditManagerTests(unittest.TestCase):
    def test_records_and_reads_events_with_structured_details(self):
        path = runtime_file("audit_records")
        audit.record_event(
            "machine_updated",
            "machine",
            entity_id="7",
            machine_number="7",
            operator="tester",
            details={"fields": {"stav": {"old": "bezi", "new": "porucha"}}},
            audit_file=path,
        )

        events = audit.read_events(path)

        self.assertEqual(len(events), 1)
        self.assertEqual(events[0]["operator"], "tester")
        self.assertEqual(events[0]["action"], "machine_updated")
        details = json.loads(events[0]["details"])
        self.assertEqual(details["fields"]["stav"]["new"], "porucha")

    def test_second_event_keeps_previous_version_backup(self):
        path = runtime_file("audit_backup")
        audit.record_event("first", "machine", audit_file=path)
        first = path.read_bytes()
        audit.record_event("second", "machine", audit_file=path)

        self.assertEqual(path.with_suffix(".csv.bak").read_bytes(), first)
        self.assertEqual(
            [event["action"] for event in audit.read_events(path)],
            ["first", "second"],
        )

    def test_changed_fields_only_returns_differences(self):
        changes = audit.changed_fields(
            {"stav": "bezi", "typ": "A"},
            {"stav": "porucha", "typ": "A"},
        )
        self.assertEqual(
            changes, {"stav": {"old": "bezi", "new": "porucha"}}
        )


if __name__ == "__main__":
    unittest.main()
