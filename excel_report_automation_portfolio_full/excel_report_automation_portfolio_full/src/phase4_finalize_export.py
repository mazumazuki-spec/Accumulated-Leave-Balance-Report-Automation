"""Phase 4: finalize report files and optionally export screenshots."""

from __future__ import annotations

import argparse
import platform
import shutil
from pathlib import Path

from config_loader import AppConfig, BranchConfig, load_config
from excel_utils import auto_fit_columns, ensure_folder, safe_load_workbook


def resolve_report_test_file(config: AppConfig, branch: BranchConfig) -> Path:
    """Return report test file for one branch."""

    expected_file = (
        config.output_dir
        / config.to_date
        / "reports"
        / config.report_test_file_pattern.format(branch=branch.name, date=config.to_date)
    )
    if expected_file.exists():
        return expected_file

    raise FileNotFoundError(f"Report test file not found: {expected_file}")


def finalize_workbook(config: AppConfig, branch: BranchConfig) -> Path:
    """Rename report sheets and save a finalized workbook."""

    source_file = resolve_report_test_file(config, branch)
    workbook = safe_load_workbook(source_file)

    report_sheet = workbook["Report_Test"] if "Report_Test" in workbook.sheetnames else workbook.active
    original_sheet = None

    for worksheet in workbook.worksheets:
        if worksheet != report_sheet:
            original_sheet = worksheet
            break

    if original_sheet is not None:
        original_sheet.title = "Original"
        original_sheet.sheet_state = "hidden"

    report_sheet.title = "Report"
    report_sheet.sheet_state = "visible"
    auto_fit_columns(report_sheet)

    for worksheet in list(workbook.worksheets):
        if worksheet.title not in {"Original", "Report"}:
            workbook.remove(worksheet)

    output_folder = ensure_folder(config.output_dir / config.to_date / "final")
    output_file = output_folder / config.report_file_pattern.format(
        branch=branch.name,
        date=config.to_date,
    )

    workbook.save(output_file)
    workbook.close()

    print(f"[{branch.name}] finalized workbook: {output_file}")
    return output_file


def export_screenshot_with_excel(
    excel_file: Path,
    sheet_name: str,
    range_address: str,
    output_png: Path,
) -> bool:
    """Export an Excel range to PNG using Microsoft Excel COM on Windows."""

    if platform.system() != "Windows":
        print("Screenshot export skipped: Microsoft Excel COM requires Windows.")
        return False

    try:
        import time
        from PIL import ImageGrab
        import win32com.client as win32
    except ImportError:
        print("Screenshot export skipped: pywin32 or Pillow is not installed.")
        return False

    excel = None
    workbook = None

    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = True
        excel.DisplayAlerts = False

        workbook = excel.Workbooks.Open(str(excel_file))
        worksheet = workbook.Worksheets(sheet_name)
        worksheet.Activate()

        range_to_copy = worksheet.Range(range_address)
        range_to_copy.Select()
        range_to_copy.CopyPicture(Appearance=1, Format=2)
        time.sleep(1)

        image = ImageGrab.grabclipboard()
        if image is None or isinstance(image, list):
            print("Screenshot export failed: clipboard did not contain an image.")
            return False

        ensure_folder(output_png.parent)
        image.convert("RGB").save(output_png)
        print(f"Screenshot saved: {output_png}")
        return True

    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        if excel is not None:
            excel.Quit()


def finalize_all_reports(config: AppConfig, export_screenshots: bool = False) -> list[Path]:
    """Finalize all branch report files."""

    output_files: list[Path] = []

    for branch in config.branches:
        final_file = finalize_workbook(config, branch)
        output_files.append(final_file)

        if export_screenshots:
            screenshot_file = (
                ensure_folder(config.screenshot_dir / config.to_date)
                / f"report_{branch.name}_{config.to_date}.png"
            )
            export_screenshot_with_excel(
                excel_file=final_file,
                sheet_name="Report",
                range_address=branch.screenshot_range,
                output_png=screenshot_file,
            )

    summary_source = (
        config.output_dir
        / config.to_date
        / "summary"
        / config.consolidated_summary.output_file.format(date=config.to_date)
    )
    if summary_source.exists():
        final_summary = (
            ensure_folder(config.output_dir / config.to_date / "final")
            / summary_source.name
        )
        shutil.copy2(summary_source, final_summary)
        output_files.append(final_summary)
        print(f"Final summary copied: {final_summary}")

    return output_files


def run_phase(config: AppConfig, export_screenshots: bool = False) -> list[Path]:
    """Run Phase 4."""

    return finalize_all_reports(config, export_screenshots=export_screenshots)


def parse_args() -> argparse.Namespace:
    """Parse CLI arguments."""

    parser = argparse.ArgumentParser(description="Phase 4: finalize and export.")
    parser.add_argument("--config", default="config/config.example.json")
    parser.add_argument("--screenshots", action="store_true")
    return parser.parse_args()


def main() -> None:
    """CLI entry point."""

    args = parse_args()
    config = load_config(Path(args.config))
    run_phase(config, export_screenshots=args.screenshots)


if __name__ == "__main__":
    main()
