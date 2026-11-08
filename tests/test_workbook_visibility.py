"""Offline Windows regression tests; only disposable fixture files are changed."""
import ctypes
import os
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "PropertyManager.App" / "Backend"))
from openpyxl import Workbook as ExcelWorkbook, load_workbook
import P2P_Registry_Sync as core
import registry_worker as registry
import operations_worker as operations
from workbook_visibility import VISIBILITY


@unittest.skipUnless(os.name == "nt", "Windows file visibility")
class WorkbookVisibilityChecks(unittest.TestCase):
    @staticmethod
    def attributes(path, value):
        from ctypes import wintypes
        kernel = ctypes.WinDLL("kernel32", use_last_error=True)
        setter = kernel.SetFileAttributesW
        setter.argtypes = [wintypes.LPCWSTR, wintypes.DWORD]
        setter.restype = wintypes.BOOL
        if not setter(str(path), value):
            raise ctypes.WinError(ctypes.get_last_error())

    @staticmethod
    def fixtures(source, local):
        for path, value in ((source, "original workbook"), (local, "updated workbook")):
            book = ExcelWorkbook()
            book.active["A1"] = value
            book.save(path)
            book.close()
        return source.read_bytes(), local.read_bytes()

    def publish(self, kind, source, local):
        if kind == "registry":
            data = load_workbook(local)
            class Workbook:
                worksheets = data.worksheets
                def save(self, target):
                    Path(target).write_bytes(local.read_bytes())
            registry.CANCEL.clear()
            try:
                registry.save_safely(core, Workbook(), source, registry.fingerprint(source))
            finally:
                data.close()
        else:
            operations.publish(local, source, operations.digest(source), lambda *args: None)

    def test_both_save_paths_preserve_visibility_and_remove_inherited_hidden_state(self):
        for kind in ("registry", "operations"):
            for original_visibility in (0, 2, 4, 6):
                with self.subTest(kind=kind, original_visibility=original_visibility), tempfile.TemporaryDirectory() as temp:
                    root = Path(temp)
                    source, local = root / "original.xlsx", root / "updated.xlsx"
                    original_bytes, updated_bytes = self.fixtures(source, local)
                    self.attributes(source, 0x20 | original_visibility)
                    mkstemp = tempfile.mkstemp
                    stages = []
                    def inherited_hidden_stage(*args, **kwargs):
                        self.assertFalse(kwargs["prefix"].startswith("."))
                        fd, name = mkstemp(*args, **kwargs)
                        stages.append(Path(name))
                        # Simulate a network share assigning hidden/system to new files.
                        self.attributes(name, 0x20 | VISIBILITY)
                        return fd, name
                    with patch.object(tempfile, "mkstemp", side_effect=inherited_hidden_stage):
                        self.publish(kind, source, local)
                    self.assertEqual(source.read_bytes(), updated_bytes)
                    self.assertEqual(source.stat().st_file_attributes & VISIBILITY, original_visibility)
                    self.assertTrue(source.stat().st_file_attributes & 0x20)
                    self.assertTrue(stages)
                    self.assertTrue(all(not stage.exists() for stage in stages))
                    backup_dir = root / ("Registry Backups" if kind == "registry" else "Operations reminder backups")
                    backups = list(backup_dir.glob("*.xlsx"))
                    self.assertEqual(len(backups), 1)
                    self.assertEqual(backups[0].read_bytes(), original_bytes)

    def test_visibility_failure_preserves_original_and_backup_and_cleans_staging(self):
        for kind, module in (("registry", registry), ("operations", operations)):
            with self.subTest(kind=kind), tempfile.TemporaryDirectory() as temp:
                root = Path(temp)
                source, local = root / "original.xlsx", root / "updated.xlsx"
                original_bytes, updated_bytes = self.fixtures(source, local)
                with patch.object(module, "preserve_visibility", side_effect=PermissionError("Cannot set visibility")):
                    with self.assertRaisesRegex(PermissionError, "Cannot set visibility"):
                        self.publish(kind, source, local)
                self.assertEqual(source.read_bytes(), original_bytes)
                self.assertFalse(source.stat().st_file_attributes & VISIBILITY)
                self.assertFalse(list(root.glob("*registry-sync-*")))
                self.assertFalse(list(root.glob("*operations-*")))
                backup_dir = root / ("Registry Backups" if kind == "registry" else "Operations reminder backups")
                self.assertEqual(next(backup_dir.glob("*.xlsx")).read_bytes(), original_bytes)

    def test_native_attribute_write_failure_is_reported(self):
        import workbook_visibility
        with tempfile.TemporaryDirectory() as temp:
            source, stage = Path(temp) / "source.xlsx", Path(temp) / "stage.xlsx"
            source.write_bytes(b"original")
            stage.write_bytes(b"updated")
            self.attributes(stage, 2)
            from unittest.mock import MagicMock
            kernel = MagicMock()
            kernel.SetFileAttributesW.return_value = False
            with patch.object(ctypes, "WinDLL", return_value=kernel), patch.object(ctypes, "get_last_error", return_value=5):
                with self.assertRaises(OSError):
                    workbook_visibility.preserve_visibility(source, stage)
            self.assertEqual(source.read_bytes(), b"original")


if __name__ == "__main__":
    unittest.main()
