#!/usr/bin/env python3
"""
Unit tests for onedrive_to_sharepoint_url.py
"""

import csv
import os
import tempfile
import unittest

from onedrive_to_sharepoint_url import (
    SyncMapping,
    collect_files_from_csv,
    collect_files_from_folder,
    encode_sharepoint_url_path,
    normalize_local_path,
    process_files,
    redact_path,
    resolve_file_to_urls,
)


class TestOneDriveToSharePointUrl(unittest.TestCase):

    def test_redact_path(self):
        # Test Windows user path redaction
        win_path = r"C:\Users\johndoe\OneDrive - Contoso\file.txt"
        redacted = redact_path(win_path)
        self.assertNotIn("johndoe", redacted)
        self.assertIn(r"C:\Users\<USER>\OneDrive - Contoso\file.txt", redacted)

        # Test Unix user path redaction
        unix_path = "/Users/alice/Documents/file.pdf"
        redacted_u = redact_path(unix_path)
        self.assertNotIn("alice", redacted_u)
        self.assertIn("/Users/<USER>/Documents/file.pdf", redacted_u)

    def test_encode_sharepoint_url_path(self):
        rel = r"Folder Name\Sub Folder\My Document (2026).docx"
        encoded = encode_sharepoint_url_path(rel)
        self.assertEqual(encoded, "Folder%20Name/Sub%20Folder/My%20Document%20%282026%29.docx")

    def test_resolve_file_to_urls_longest_prefix(self):
        mappings = [
            SyncMapping(
                mount_point=normalize_local_path("/mock/user/OneDrive - Tenant"),
                url_namespace="https://tenant-my.sharepoint.com/personal/user",
                scope_type="business"
            ),
            SyncMapping(
                mount_point=normalize_local_path("/mock/user/OneDrive - Tenant/Shared Library"),
                url_namespace="https://tenant.sharepoint.com/sites/Team/Shared%20Documents",
                scope_type="business"
            ),
        ]
        mappings.sort(key=lambda x: len(x.mount_point), reverse=True)

        # File in nested library
        f1 = normalize_local_path("/mock/user/OneDrive - Tenant/Shared Library/Report.xlsx")
        url, direct, mount = resolve_file_to_urls(f1, mappings)
        self.assertEqual(mount, normalize_local_path("/mock/user/OneDrive - Tenant/Shared Library"))
        self.assertEqual(url, "https://tenant.sharepoint.com/sites/Team/Shared%20Documents/Report.xlsx")

        # File in root personal onedrive
        f2 = normalize_local_path("/mock/user/OneDrive - Tenant/Documents/Notes.txt")
        url, direct, mount = resolve_file_to_urls(f2, mappings)
        self.assertEqual(mount, normalize_local_path("/mock/user/OneDrive - Tenant"))
        self.assertEqual(url, "https://tenant-my.sharepoint.com/personal/user/Documents/Notes.txt")

        # File outside
        f3 = normalize_local_path("/mock/other/file.txt")
        url, direct, mount = resolve_file_to_urls(f3, mappings)
        self.assertIsNone(url)
        self.assertIsNone(mount)

    def test_collect_files_from_csv_with_spaces_and_special_chars(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            csv_path = os.path.join(tmpdir, "input_with_spaces.csv")
            # Unquoted and quoted lines with internal spaces and brackets
            content = (
                "C:\\Users\\John Doe\\OneDrive - Contoso\\ESMO 2026 documents\\Study Protocol (Phase 3).pdf\n"
                '"C:\\Users\\John Doe\\OneDrive - Contoso\\ESMO 2026 documents\\Data Sheet [Final].xlsx"\n'
                "C:\\Users\\John Doe\\OneDrive - Contoso\\ESMO 2026 documents\\Folder #1\\Appendix A.docx\n"
            )
            with open(csv_path, "w", encoding="utf-8") as fp:
                fp.write(content)

            files = list(collect_files_from_csv(csv_path))
            self.assertEqual(len(files), 3)
            self.assertEqual(
                files[0],
                r"C:\Users\John Doe\OneDrive - Contoso\ESMO 2026 documents\Study Protocol (Phase 3).pdf"
            )
            self.assertEqual(
                files[1],
                r"C:\Users\John Doe\OneDrive - Contoso\ESMO 2026 documents\Data Sheet [Final].xlsx"
            )
            self.assertEqual(
                files[2],
                r"C:\Users\John Doe\OneDrive - Contoso\ESMO 2026 documents\Folder #1\Appendix A.docx"
            )

    def test_collect_files_from_csv_with_header_and_comma(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            csv_path = os.path.join(tmpdir, "input.csv")
            with open(csv_path, "w", newline="", encoding="utf-8") as fp:
                writer = csv.writer(fp)
                writer.writerow(["File Location", "Status"])
                writer.writerow([r"C:\Users\John Doe\OneDrive\Doc 1.docx", "Active"])
                writer.writerow([r"C:\Users\John Doe\OneDrive\Doc 2.xlsx", "Pending"])

            files = list(collect_files_from_csv(csv_path))
            self.assertEqual(files, [
                r"C:\Users\John Doe\OneDrive\Doc 1.docx",
                r"C:\Users\John Doe\OneDrive\Doc 2.xlsx"
            ])

    def test_process_files_csv_output(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            f1 = os.path.join(tmpdir, "test 1.txt")
            f2 = os.path.join(tmpdir, "test 2 (final).docx")
            f3 = os.path.join(tmpdir, "outside.pdf")
            for f in (f1, f2, f3):
                with open(f, "w") as fp:
                    fp.write("dummy")

            mapping = [
                SyncMapping(
                    mount_point=normalize_local_path(tmpdir),
                    url_namespace="https://company.sharepoint.com/sites/Docs",
                    scope_type="business"
                )
            ]

            out_csv = os.path.join(tmpdir, "output.csv")
            stats = process_files([f1, f2, "/completely/outside/file.xyz"], mapping, out_csv)

            self.assertEqual(stats["total_files"], 3)
            self.assertEqual(stats["resolved"], 2)
            self.assertEqual(stats["unmatched"], 1)

            with open(out_csv, "r", encoding="utf-8-sig") as fp:
                rows = list(csv.DictReader(fp))
                self.assertEqual(len(rows), 3)
                self.assertEqual(rows[0]["resolution_status"], "RESOLVED")
                self.assertIn("https://company.sharepoint.com/sites/Docs/test%201.txt", rows[0]["sharepoint_web_url"])
                self.assertIn("https://company.sharepoint.com/sites/Docs/test%202%20%28final%29.docx", rows[1]["sharepoint_web_url"])
                self.assertEqual(rows[2]["resolution_status"], "NOT_IN_ONEDRIVE_MOUNT")


if __name__ == "__main__":
    unittest.main()
