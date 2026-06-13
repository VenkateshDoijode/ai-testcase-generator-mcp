"""
AppConfig — loads settings from resources/config.ini.
"""

import os
import configparser


class AppConfig:

    _CONFIG_PATH = os.path.join(
        os.path.dirname(__file__), "..", "..", "resources", "config.ini"
    )

    def __init__(self):
        cfg = configparser.ConfigParser()
        cfg.read(self._CONFIG_PATH, encoding="utf-8")

        if not cfg.has_section("jira"):
            raise KeyError(
                "[ERROR] 'config.ini' is missing the required [jira] section. "
                "Check resources/config.ini.example for the correct format."
            )

        jira = cfg["jira"]
        self.base_url     = jira.get("base_url", "").strip()
        self.jira_token   = jira.get("jira_token", "").strip()
        self.zephyr_path  = jira.get("zephyr_path", "").strip()
        self.project_key  = jira.get("project_key", "").strip()
        self.project_id   = jira.get("project_id", "").strip()

        tc = cfg["testcase"] if cfg.has_section("testcase") else {}
        self.owner_account_id  = tc.get("owner_account_id", "").strip()
        self.test_cycle        = tc.get("test_cycle", "").strip()
        self.test_reviewer_id  = tc.get("test_reviewer_id", "").strip()
        self.test_case_key     = tc.get("test_case_key", "").strip()

        ai = cfg["ai"] if cfg.has_section("ai") else {}
        self.hf_token    = ai.get("hf_token", "").strip()
        self.hf_endpoint = ai.get("hf_endpoint", "").strip()

        df = cfg["defaults"] if cfg.has_section("defaults") else {}
        self.default_type_of_test       = df.get("type_of_test", "").strip()
        self.default_automation_status  = df.get("automation_status", "").strip()
