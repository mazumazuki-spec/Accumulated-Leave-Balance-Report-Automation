"""Configuration loader for the Excel report automation demo."""

from __future__ import annotations

import json
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any


@dataclass(frozen=True)
class BranchConfig:
    """Configuration for a demo branch."""

    name: str
    branch_id: int
    exclude_employee_ids: list[str]
    remark_employee_ids: dict[str, str]
    report_range: str
    screenshot_range: str


@dataclass(frozen=True)
class ConsolidatedSummaryConfig:
    """Configuration for consolidated summary output."""

    template_branch: str
    summary_range: str
    output_file: str
    equal_wording: dict[str, str]
    decrease_wording: dict[str, str]
    increase_wording: dict[str, str]


@dataclass(frozen=True)
class AppConfig:
    """Application configuration."""

    project_root: Path
    input_dir: Path
    template_dir: Path
    output_dir: Path
    screenshot_dir: Path
    from_date: str
    to_date: str
    api_base_url: str | None
    summary_threshold_days: float
    delete_keywords: list[str]
    delete_columns: list[str]
    raw_file_pattern: str
    clean_file_pattern: str
    report_test_file_pattern: str
    report_file_pattern: str
    branches: list[BranchConfig]
    consolidated_summary: ConsolidatedSummaryConfig


def load_config(config_path: Path) -> AppConfig:
    """Load JSON config and convert it to a typed AppConfig object."""

    with config_path.open("r", encoding="utf-8") as file:
        raw: dict[str, Any] = json.load(file)

    project_root = Path(raw.get("project_root", ".")).resolve()
    to_date = raw.get("to_date") or date.today().strftime("%Y-%m-%d")

    branches = [
        BranchConfig(
            name=str(item["name"]),
            branch_id=int(item["branch_id"]),
            exclude_employee_ids=[str(x) for x in item.get("exclude_employee_ids", [])],
            remark_employee_ids={
                str(key): str(value)
                for key, value in item.get("remark_employee_ids", {}).items()
            },
            report_range=str(item["report_range"]),
            screenshot_range=str(item["screenshot_range"]),
        )
        for item in raw["branches"]
    ]

    summary_raw = raw["consolidated_summary"]
    consolidated_summary = ConsolidatedSummaryConfig(
        template_branch=str(summary_raw["template_branch"]),
        summary_range=str(summary_raw["summary_range"]),
        output_file=str(summary_raw["output_file"]),
        equal_wording=dict(summary_raw["equal_wording"]),
        decrease_wording=dict(summary_raw["decrease_wording"]),
        increase_wording=dict(summary_raw["increase_wording"]),
    )

    return AppConfig(
        project_root=project_root,
        input_dir=(project_root / raw.get("input_dir", "data/input")).resolve(),
        template_dir=(project_root / raw.get("template_dir", "data/templates")).resolve(),
        output_dir=(project_root / raw.get("output_dir", "data/output")).resolve(),
        screenshot_dir=(project_root / raw.get("screenshot_dir", "data/screenshots")).resolve(),
        from_date=str(raw["from_date"]),
        to_date=to_date,
        api_base_url=raw.get("api_base_url") or None,
        summary_threshold_days=float(raw.get("summary_threshold_days", 5)),
        delete_keywords=[str(x) for x in raw.get("delete_keywords", [])],
        delete_columns=[str(x) for x in raw.get("delete_columns", [])],
        raw_file_pattern=str(raw.get("raw_file_pattern", "raw_{branch}_{date}.xlsx")),
        clean_file_pattern=str(raw.get("clean_file_pattern", "clean_{branch}_{date}.xlsx")),
        report_test_file_pattern=str(
            raw.get("report_test_file_pattern", "report_test_{branch}_{date}.xlsx")
        ),
        report_file_pattern=str(raw.get("report_file_pattern", "report_{branch}_{date}.xlsx")),
        branches=branches,
        consolidated_summary=consolidated_summary,
    )


def get_branch(config: AppConfig, branch_name: str) -> BranchConfig:
    """Return branch configuration by name."""

    for branch in config.branches:
        if branch.name == branch_name:
            return branch

    raise ValueError(f"Unknown branch: {branch_name}")
