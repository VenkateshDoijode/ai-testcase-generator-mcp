"""
TestCycleUpdater — reads test case keys and execution details from
resources/update_test_cycle.csv and updates the matching test executions
inside the Zephyr Scale test cycle specified in config.ini.

The cycle key is read from config.ini -> [testcase] -> test_cycle.

Zephyr Scale endpoints used:
    GET  /rest/atm/1.0/testrun/{cycleKey}
        -> fetch cycle and its items (executions)
    PUT  /rest/atm/1.0/testrun/{cycleKey}/testresults
        -> bulk-update execution records

Expected CSV columns (from update_test_cycle.csv):
    Key          : Zephyr test case key (e.g. PROJECT-T1234)  [required]
    Environment  : Environment string (e.g. DEV)              [optional]
    Status       : Execution status (e.g. Pass, Fail, In Progress, Not Executed) [optional]
    Assigned To  : Assignee account ID                        [optional]
    Comment      : Execution comment                          [optional]

Rows with an empty Key are skipped.

Usage:
    python -m test.update_test_cycle
    python -m test.update_test_cycle --dry-run
    python -m test.update_test_cycle --cycle PROJECT-C1234
    python -m test.update_test_cycle --file resources/update_test_cycle.csv
"""

import sys
import os
import csv
import json
import argparse

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", ".."))

from main.config import AppConfig
from main.jira_client import JiraClient
from main.report import ReportGenerator


class TestCycleUpdater:

    DEFAULT_CSV = os.path.join(
        os.path.dirname(__file__), "..", "..", "resources", "update_test_cycle.csv"
    )

    COLUMN_MAP = {
        "key"         : "key",
        "environment" : "environment",
        "status"      : "status",
        "assigned to" : "assigned_to",
        "comment"     : "comment",
    }

    def __init__(self, csv_path: str = "", cycle_key: str = "", dry_run: bool = False):
        self._cfg     = AppConfig()
        self._client  = JiraClient()
        self.csv_path = csv_path or self.DEFAULT_CSV
        self.cycle_key = (cycle_key or self._cfg.test_cycle).strip()
        self.dry_run  = dry_run

        if not self.cycle_key:
            print("[ERROR] test_cycle is not set in config.properties under [testcase].")
            print("        Add: test_cycle = <your-cycle-key>")
            sys.exit(1)

    # — Read CSV ----------------------------------------------------------

    def load_csv(self) -> list:
        """Read CSV and return list of normalised row dicts keyed by test case key."""
        if not os.path.exists(self.csv_path):
            print(f"[ERROR] CSV not found: {os.path.abspath(self.csv_path)}")
            sys.exit(1)

        rows = []
        with open(self.csv_path, encoding="utf-8-sig", errors="replace", newline="") as f:
            reader = csv.DictReader(f)
            raw_headers = reader.fieldnames or []
            header_lookup = {h.strip().lower(): h for h in raw_headers}

            if "key" not in header_lookup:
                print(f"[ERROR] 'Key' column not found. Headers: {raw_headers}")
                sys.exit(1)

            for i, row in enumerate(reader, start=2):
                record = {}
                for col_lower, field in self.COLUMN_MAP.items():
                    original = header_lookup.get(col_lower, "")
                    record[field] = row.get(original, "").strip() if original else ""
                if not record.get("key"):
                    continue
                rows.append(record)

        print(f"  Loaded {len(rows)} row(s) from: {os.path.basename(self.csv_path)}")
        return rows

    # — Fetch cycle ---------------------------------------------------------

    def fetch_cycle(self) -> dict:
        """GET /testrun/{cycleKey}. Exits on error."""
        status, data = self._client.get(f"/testrun/{self.cycle_key}")
        if status == 200:
            return data
        if status == 404:
            print(f"[ERROR] Cycle '{self.cycle_key}' not found. Check test_cycle in config.properties.")
        elif status == 401:
            print(f"[ERROR] Unauthorized — check jira_token in config.properties.")
        else:
            print(f"[ERROR] Could not fetch cycle: HTTP {status} — {data}")
        sys.exit(1)

    # — Build execution index -------------------------------------------------

    def build_execution_index(self, cycle: dict) -> dict:
        """Return {testCaseKey: execution_item} from cycle items."""
        index = {}
        for item in cycle.get("items", []):
            tc_key = item.get("testCaseKey", "")
            if tc_key:
                index[tc_key] = item
        return index

    # — Build update payload -------------------------------------------------

    def build_item_update(self, existing_item: dict, record: dict) -> dict:
        """Merge CSV fields into the existing execution item dict."""
        item = dict(existing_item)

        if record.get("environment"):
            item["environment"] = record["environment"]
        if record.get("status"):
            item["status"] = record["status"]
        if record.get("assigned_to"):
            item["assignedTo"] = record["assigned_to"]
        if record.get("comment"):
            item["comment"] = record["comment"]

        return item

    # — Update executions ------------------------------------------------------

    def update_executions(self, items: list) -> tuple:
        """POST execution items to /testrun/{cycleKey}/testresults as a JSON array.
        The 'id' field must be stripped — the API rejects it.
        Returns (success, message)."""
        payload = [{k: v for k, v in item.items() if k != "id"} for item in items]
        status, body = self._client.post_list(
            f"/testrun/{self.cycle_key}/testresults", payload
        )
        if status in (200, 201):
            return True, "ok"
        try:
            detail = body if isinstance(body, dict) else json.loads(body)
            msg = detail.get("message") or detail.get("error") or str(body)[:300]
        except Exception:
            msg = str(body)[:300]
        return False, f"{status} - {msg}"

    # — Run -------------------------------------------------------------------

    def run(self):
        print("=" * 65)
        print("  Jira Zephyr Scale — Update Test Cycle Executions")
        print("=" * 65)
        print(f"  Cycle   : {self.cycle_key}")
        print(f"  Project : {self._cfg.project_key}")
        print(f"  CSV File: {os.path.abspath(self.csv_path)}")
        if self.dry_run:
            print("  Mode    : DRY RUN (no changes will be made)")
        print("=" * 65 + "\n")

        records = self.load_csv()
        if not records:
            print("[WARN] No valid rows found in CSV.")
            sys.exit(0)

        print(f"  Fetching cycle '{self.cycle_key}'...")
        cycle      = self.fetch_cycle()
        cycle_name = cycle.get("name", self.cycle_key)
        print(f"  Cycle name : {cycle_name}")
        print(f"  Executions : {len(cycle.get('items', []))}\n")

        exec_index = self.build_execution_index(cycle)

        ok_list      = []
        fail_list    = []
        skip_list    = []
        updated_items = []
        test_cases   = []

        for i, record in enumerate(records, 1):
            tc_key = record["key"]
            print(f"  [{i:>3}/{len(records)}] {tc_key:20s} ... ", end="", flush=True)

            if tc_key not in exec_index:
                print("⚠  Not in cycle — skipped")
                skip_list.append(tc_key)
                test_cases.append({"key": tc_key, "name": tc_key})
                continue

            if self.dry_run:
                env = record.get("environment", "")
                sts = record.get("status", "")
                print(f"SKIP (dry-run) [{env or '-'} / {sts or '-'}]")
                test_cases.append({"key": tc_key, "name": tc_key})
                continue

            updated = self.build_item_update(exec_index[tc_key], record)
            updated_items.append((tc_key, updated))
            print(f"Queued")
            test_cases.append({"key": tc_key, "name": tc_key})

        if not self.dry_run and updated_items:
            tc_keys = [t[0] for t in updated_items]
            items   = [t[1] for t in updated_items]
            print(f"\n  Pushing {len(items)} update(s) to Zephyr ... ", end="", flush=True)
            success, msg = self.update_executions(items)
            if success:
                print(f"✅  Done")
                ok_list.extend(tc_keys)
            else:
                print(f"❌  Failed: {msg}")
                fail_list.extend(tc_keys)

        print()
        print("=" * 65)
        if self.dry_run:
            print(f"  DRY RUN — {len(ok_list) + len(skip_list)} row(s) processed.")
        else:
            print(f"  SUMMARY")
            print(f"  Updated : {len(ok_list)}")
            print(f"  Skipped : {len(skip_list)}  (not found in cycle)")
            print(f"  Failed  : {len(fail_list)}")
            if ok_list:
                print(f"\n  Updated keys: {', '.join(ok_list)}")
            if skip_list:
                print(f"\n  Skipped (not in cycle): {', '.join(skip_list)}")
            if fail_list:
                print(f"\n  Failed: {', '.join(fail_list)}")
        print("=" * 65)

        if not self.dry_run:
            reports_dir = os.path.join(os.path.dirname(__file__), "..", "..", "reports")
            os.makedirs(reports_dir, exist_ok=True)
            report_path = os.path.join(reports_dir, "Update_Cycle_Report.html")
            ReportGenerator(
                output_path=report_path,
                action_label="Updated",
                report_title=f"Zephyr Scale — Update Cycle Report ({self.cycle_key})",
            ).generate(
                test_cases=test_cases,
                ok_list=ok_list,
                fail_list=fail_list,
                project_key=self._cfg.project_key,
                folder_path=cycle_name,
            )


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Update test cycle executions from CSV in Jira Zephyr Scale"
    )
    parser.add_argument("--file",    default="", help="Path to CSV (default: resources/update_test_cycle.csv)")
    parser.add_argument("--cycle",   default="", help="Cycle key (default: test_cycle in config.properties)")
    parser.add_argument("--dry-run", action="store_true", help="Preview without updating")
    args = parser.parse_args()
    TestCycleUpdater(csv_path=args.file, cycle_key=args.cycle, dry_run=args.dry_run).run()
