from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
import subprocess
import sys

import pandas as pd
from openpyxl import Workbook

import append_merge_excel as app


def write_book(path: Path, data: dict[str, list[object]]) -> Path:
    pd.DataFrame(data).to_excel(path, index=False)
    return path


class ExcelMergeTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_directory = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_directory.name)

    def tearDown(self) -> None:
        self.temp_directory.cleanup()

    def test_merge_preserves_file_order_and_verifies(self) -> None:
        first = write_book(self.root / "first.xlsx", {"ID": ["001", "002"], "Name": ["A", "B"]})
        second = write_book(self.root / "second.xlsx", {"ID": ["003"], "Name": ["C"]})

        result = app.merge_workbooks([first, second], output_root=self.root, verify=True)
        merged = pd.read_excel(result.output_path, dtype=str)

        self.assertEqual(
            merged.to_dict("records"),
            [
                {"ID": "001", "Name": "A"},
                {"ID": "002", "Name": "B"},
                {"ID": "003", "Name": "C"},
            ],
        )
        self.assertIsNotNone(result.verification)
        self.assertTrue(result.verification and result.verification.passed)

    def test_schema_order_mismatch_aborts_without_output(self) -> None:
        first = write_book(self.root / "first.xlsx", {"ID": [1], "Name": ["A"]})
        second = write_book(self.root / "second.xlsx", {"Name": ["B"], "ID": [2]})

        with self.assertRaisesRegex(app.SchemaMismatchError, "column order differs"):
            app.merge_workbooks([first, second], output_root=self.root)

        self.assertFalse(list((self.root / "MergedFiles").glob("*.xlsx")))

    def test_duplicate_input_is_ignored(self) -> None:
        source = write_book(self.root / "source.xlsx", {"ID": [1], "Name": ["A"]})

        result = app.merge_workbooks([source, source], output_root=self.root, verify=False)

        self.assertEqual(result.total_rows, 1)
        self.assertEqual(len(result.per_file_rows), 1)

    def test_explicit_output_path_is_respected(self) -> None:
        source = write_book(self.root / "source.xlsx", {"ID": [1]})
        output = self.root / "custom" / "result.xlsx"

        result = app.merge_workbooks([source], output_path=output, verify=False)

        self.assertEqual(result.output_path, output.resolve())
        self.assertTrue(output.exists())

    def test_numeric_merge_verifies_without_force_text(self) -> None:
        source = write_book(self.root / "numeric.xlsx", {"ID": [1, 2], "Amount": [10.5, 20.25]})

        result = app.merge_workbooks(
            [source],
            output_root=self.root,
            force_text=False,
            verify=True,
        )

        self.assertTrue(result.verification and result.verification.passed)

    def test_output_cannot_overwrite_source(self) -> None:
        source = write_book(self.root / "source.xlsx", {"ID": [1]})

        with self.assertRaisesRegex(app.MergeError, "cannot overwrite"):
            app.merge_workbooks([source], output_path=source, verify=False)

    def test_corrupt_workbook_has_safe_error(self) -> None:
        corrupt = self.root / "corrupt.xlsx"
        corrupt.write_text("not an Excel workbook", encoding="utf-8")

        with self.assertRaisesRegex(app.MergeError, "Could not inspect corrupt.xlsx"):
            app.merge_workbooks([corrupt], output_root=self.root)

    def test_duplicate_source_headers_are_rejected(self) -> None:
        source = self.root / "duplicate.xlsx"
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["ID", "ID"])
        sheet.append([1, 2])
        workbook.save(source)

        with self.assertRaisesRegex(app.MergeError, "duplicate column headings"):
            app.merge_workbooks([source], output_root=self.root)

    def test_blank_source_header_is_rejected(self) -> None:
        source = self.root / "blank-header.xlsx"
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["ID", None])
        sheet.append([1, 2])
        workbook.save(source)

        with self.assertRaisesRegex(app.MergeError, "blank column headings"):
            app.merge_workbooks([source], output_root=self.root)

    def test_archive_only_after_success(self) -> None:
        source = write_book(self.root / "source.xlsx", {"ID": [1]})
        archive = self.root / "archive"

        result = app.merge_workbooks(
            [source],
            output_root=self.root / "output",
            verify=True,
            move_after_success=True,
            archive_dir=archive,
        )

        self.assertTrue(result.verification and result.verification.passed)
        self.assertFalse(source.exists())
        self.assertTrue((archive / "source.xlsx").exists())

    def test_archive_cannot_equal_source_folder(self) -> None:
        source = write_book(self.root / "source.xlsx", {"ID": [1]})

        with self.assertRaisesRegex(app.MergeError, "same as a source folder"):
            app.merge_workbooks(
                [source],
                output_root=self.root / "output",
                move_after_success=True,
                archive_dir=self.root,
            )

    def test_excel_row_limit_checked_before_write(self) -> None:
        source = write_book(self.root / "source.xlsx", {"ID": [1, 2]})
        original_limit = app.EXCEL_MAX_ROWS
        app.EXCEL_MAX_ROWS = 2
        try:
            with self.assertRaisesRegex(app.MergeError, "exceeds Excel"):
                app.merge_workbooks([source], output_root=self.root, verify=False)
        finally:
            app.EXCEL_MAX_ROWS = original_limit

    def test_sheet_name_is_sanitized(self) -> None:
        self.assertEqual(app.sanitize_sheet_name("Report/2026"), "Report_2026")
        with self.assertRaises(app.MergeError):
            app.sanitize_sheet_name("   ")

    def test_cli_returns_schema_error_code(self) -> None:
        first = write_book(self.root / "first.xlsx", {"ID": [1]})
        second = write_book(self.root / "second.xlsx", {"Other": [2]})

        self.assertEqual(app.main([str(first), str(second), "--outdir", str(self.root)]), 3)

    def test_cli_subprocess_creates_verified_output(self) -> None:
        first = write_book(self.root / "first.xlsx", {"ID": ["001"]})
        second = write_book(self.root / "second.xlsx", {"ID": ["002"]})
        output = self.root / "result.xlsx"
        script = Path(app.__file__).resolve()

        completed = subprocess.run(
            [
                sys.executable,
                str(script),
                str(first),
                str(second),
                "--output",
                str(output),
            ],
            check=False,
            capture_output=True,
            text=True,
        )

        self.assertEqual(completed.returncode, 0, completed.stderr)
        self.assertIn("Verification: PASSED", completed.stdout)
        self.assertTrue(output.exists())


if __name__ == "__main__":
    unittest.main()
