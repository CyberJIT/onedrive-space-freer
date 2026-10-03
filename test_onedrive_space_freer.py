"""
Unit tests for onedrive_space_freer.py
"""

import os
import shutil
import tempfile
import unittest
from unittest.mock import MagicMock, patch

from onedrive_space_freer import (
    collect_files,
    format_bytes,
    matches_patterns,
    normalize_extensions,
    process_onedrive_files,
    should_include_file,
    should_include_folder,
)


class TestOneDriveSpaceFreer(unittest.TestCase):
    def setUp(self):
        self.test_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.test_dir, ignore_errors=True)

    def _create_file(self, rel_path: str, size: int = 100) -> str:
        full_path = os.path.join(self.test_dir, rel_path)
        os.makedirs(os.path.dirname(full_path), exist_ok=True)
        with open(full_path, "wb") as f:
            f.write(b"X" * size)
        return full_path

    def test_extension_normalization(self):
        res = normalize_extensions(["pdf", ".docx", "PBIX", " .xlsx "])
        self.assertEqual(res, {".pdf", ".docx", ".pbix", ".xlsx"})
        self.assertEqual(normalize_extensions(None), set())
        self.assertEqual(normalize_extensions([]), set())

    def test_pattern_matching(self):
        self.assertTrue(matches_patterns("Budget_2025_Final.xlsx", ["2025"]))
        self.assertTrue(matches_patterns("Budget_2025_Final.xlsx", ["*Budget*"]))
        self.assertFalse(matches_patterns("Budget_2025_Final.xlsx", ["2024", "Invoice"]))
        self.assertTrue(matches_patterns("Report.PBIX", ["*.pbix"]))

    def test_should_include_file(self):
        # By extension
        self.assertTrue(
            should_include_file("report.pbix", ext_include=["pbix", "xlsx"])
        )
        self.assertFalse(
            should_include_file("notes.txt", ext_include=["pbix", "xlsx"])
        )
        self.assertFalse(
            should_include_file("temp.tmp", ext_exclude=["tmp", "bak"])
        )
        # By name
        self.assertTrue(
            should_include_file("Quarterly_Sales_Report.docx", file_include=["*Sales*"])
        )
        self.assertFalse(
            should_include_file("Quarterly_Report.docx", file_include=["*Sales*"])
        )
        self.assertFalse(
            should_include_file("Sales_Draft.docx", file_include=["*Sales*"], file_exclude=["*Draft*"])
        )

    def test_should_include_folder(self):
        self.assertTrue(should_include_folder("2025_Reports", folder_include=["2025"]))
        self.assertFalse(should_include_folder("2024_Reports", folder_include=["2025"]))
        self.assertFalse(should_include_folder("Archive_Folder", folder_exclude=["Archive*"]))
        self.assertTrue(should_include_folder("Active_Folder", folder_exclude=["Archive*"]))

    def test_format_bytes(self):
        self.assertEqual(format_bytes(500), "500 B")
        self.assertEqual(format_bytes(1024), "1.00 KB")
        self.assertEqual(format_bytes(1048576), "1.00 MB")
        self.assertEqual(format_bytes(1073741824), "1.00 GB")

    def test_file_collection_flat_and_recursive(self):
        self._create_file("file1.pbix")
        self._create_file("file2.docx")
        self._create_file("sub/file3.pbix")
        self._create_file("sub/archive/file4.pbix")
        self._create_file("temp/file5.tmp")

        # Non-recursive
        flat_files = collect_files(self.test_dir, recurse=False)
        self.assertEqual(len(flat_files), 2)

        # Recursive all
        all_rec = collect_files(self.test_dir, recurse=True)
        self.assertEqual(len(all_rec), 5)

        # Recursive with ext include
        pbix_only = collect_files(self.test_dir, recurse=True, ext_include=["pbix"])
        self.assertEqual(len(pbix_only), 3)

        # Recursive with folder exclusion
        no_archive = collect_files(
            self.test_dir,
            recurse=True,
            ext_include=["pbix"],
            folder_exclude=["archive"],
        )
        self.assertEqual(len(no_archive), 2)

    @patch("onedrive_space_freer.mark_file_free_space")
    def test_process_onedrive_files_success_and_dry_run(self, mock_mark):
        mock_mark.return_value = (True, "marked")
        self._create_file("doc1.pdf", size=2048)
        self._create_file("doc2.pdf", size=4096)

        summary = process_onedrive_files(self.test_dir, recurse=True, dry_run=True)
        self.assertEqual(summary["total_discovered"], 2)
        self.assertEqual(summary["marked_for_freeing"], 2)
        self.assertEqual(summary["errors"], 0)
        self.assertEqual(summary["bytes_freed_candidate"], 6144)

    @patch("onedrive_space_freer.mark_file_free_space")
    def test_process_graceful_error_handling(self, mock_mark):
        # Return error on first file, success on second
        mock_mark.side_effect = [
            (False, "Access Denied"),
            (True, "already_freed"),
        ]
        self._create_file("fail.pdf")
        self._create_file("ok.pdf")

        summary = process_onedrive_files(self.test_dir, recurse=False)
        self.assertEqual(summary["total_discovered"], 2)
        self.assertEqual(summary["errors"], 1)
        self.assertEqual(summary["already_freed"], 1)
        self.assertEqual(summary["marked_for_freeing"], 0)
        self.assertEqual(len(summary["error_details"]), 1)


if __name__ == "__main__":
    unittest.main()
