"""Run all portfolio demo phases in order."""

from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path


PHASE_FILES = [
    "src/phase1_download_clean.py",
    "src/phase2_build_branch_reports.py",
    "src/phase3_consolidate_summary.py",
    "src/phase4_finalize_export.py",
]


def run_phase(phase_file: str, config_path: str, export_screenshots: bool = False) -> None:
    """Run one phase script as a subprocess."""

    command = [sys.executable, phase_file, "--config", config_path]

    if phase_file.endswith("phase4_finalize_export.py") and export_screenshots:
        command.append("--screenshots")

    print("=" * 80)
    print(f"Running: {' '.join(command)}")
    print("=" * 80)

    result = subprocess.run(command, text=True)

    if result.returncode != 0:
        raise RuntimeError(f"Phase failed: {phase_file}")


def parse_args() -> argparse.Namespace:
    """Parse CLI arguments."""

    parser = argparse.ArgumentParser(description="Run all demo workflow phases.")
    parser.add_argument("--config", default="config/config.example.json")
    parser.add_argument("--screenshots", action="store_true")
    return parser.parse_args()


def main() -> None:
    """Run all phases in order."""

    args = parse_args()

    for phase_file in PHASE_FILES:
        if not Path(phase_file).exists():
            raise FileNotFoundError(f"Phase file not found: {phase_file}")

        run_phase(
            phase_file=phase_file,
            config_path=args.config,
            export_screenshots=args.screenshots,
        )

    print("=" * 80)
    print("All phases completed successfully.")
    print("=" * 80)


if __name__ == "__main__":
    main()
