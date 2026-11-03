# -*- coding: utf-8 -*-
import json
import os
import tempfile
from pathlib import Path
import unittest
from unittest.mock import patch

import backup_manager
import orders


def _write(path, data):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as stream:
        json.dump(data, stream, ensure_ascii=False)


class AutoImportMissingFileTests(unittest.TestCase):
    def _run(self, last):
        import app_gui
        with tempfile.TemporaryDirectory() as directory:
            if last is not None:
                _write(os.path.join(directory, "data", "last_import.json"), last)
            with patch.object(app_gui, "APP", directory), \
                    patch.object(app_gui, "find_latest_stock_file", return_value=None), \
                    patch.object(app_gui, "latest_inbox", return_value=None), \
                    patch.object(app_gui.time, "sleep"):
                result = app_gui.auto_import_stock()
            with open(os.path.join(directory, "data", "last_import.json"), encoding="utf-8") as stream:
                saved = json.load(stream)
        return result, saved

    def test_missing_source_after_successful_import_is_only_a_notice(self):
        result, saved = self._run({"sig": "x|1|1", "sd": "2026-10-30", "fail_count": 18})
        self.assertTrue(result["ok"])
        self.assertEqual(result["reason"], "无新表")
        self.assertEqual(result["fail_count"], 0)
        self.assertEqual(saved["fail_count"], 0)

    def test_missing_source_without_any_import_still_counts_failure(self):
        result, saved = self._run(None)
        self.assertEqual(result["reason"], "未找到库存表")
        self.assertGreaterEqual(saved["fail_count"], 3)


class BackupRetentionTests(unittest.TestCase):
    def test_prune_keeps_recent_and_monthly_and_ignores_invalid(self):
        with tempfile.TemporaryDirectory() as directory:
            previous = (backup_manager.APP, backup_manager.BACKUP_DIR)
            backup_manager.APP = directory
            backup_manager.BACKUP_DIR = os.path.join(directory, "data", "backups")
            try:
                for name, data in (("stock.json", {"items": []}), ("orders.json", {"orders": []}),
                                   ("recipes.json", {}), ("mapping.json", {})):
                    _write(os.path.join(directory, name), data)
                valid = backup_manager.create_backup()["path"]
                os.makedirs(backup_manager.BACKUP_DIR, exist_ok=True)
                broken = os.path.join(backup_manager.BACKUP_DIR, "OrderSystem_20200101_000000_000000.zip")
                with open(broken, "wb") as stream:
                    stream.write(b"not a zip")
                # 复制出同一个月的多份有效备份
                import shutil
                names = []
                for i in range(5):
                    name = os.path.join(backup_manager.BACKUP_DIR,
                                        "OrderSystem_20260101_00000%d_000000.zip" % i)
                    shutil.copy(valid, name)
                    names.append(name)
                removed = backup_manager.prune_backups(keep_recent=2, keep_monthly=2)
                self.assertTrue(os.path.exists(broken))
                self.assertTrue(os.path.exists(valid))
                self.assertEqual(len(removed), 4)
                self.assertTrue(os.path.exists(names[-1]))
            finally:
                backup_manager.APP, backup_manager.BACKUP_DIR = previous


class DashboardHiddenCountTests(unittest.TestCase):
    def test_dashboard_api_reports_hidden_closed_records(self):
        import app_gui
        rows_all = [{"serial": "1", "status": "待生产"}, {"serial": "2", "status": "已作废"}]
        with patch("orders.production_dashboard",
                   side_effect=lambda include_closed=False: rows_all if include_closed else rows_all[:1]):
            result = app_gui.get_production_dashboard(False)
        self.assertTrue(result["ok"])
        self.assertEqual(result["hidden_closed"], 1)
        self.assertEqual(app_gui.get_production_dashboard(True)["hidden_closed"], 0)


class BatchReconcileTests(unittest.TestCase):
    def test_report_is_read_only_and_lists_differences(self):
        import app_gui
        stock = {"items": [
            {"name": "甲", "short": "甲", "qty": 10, "batches": [{"batch": "A", "qty": 12}]},
            {"name": "乙", "short": "乙", "qty": 5, "batches": [{"batch": "B", "qty": 5}]},
            {"name": "丙", "short": "丙", "qty": 3},
        ]}
        with tempfile.TemporaryDirectory() as directory:
            path = os.path.join(directory, "stock.json")
            _write(path, stock)
            before = Path(path).read_bytes()
            with patch.object(app_gui, "APP", directory):
                result = app_gui.batch_reconcile_report()
            self.assertEqual(Path(path).read_bytes(), before)
        self.assertEqual(result["n"], 1)
        self.assertEqual(result["items"][0]["short"], "甲")
        self.assertEqual(result["items"][0]["diff"], 2)


class UpdateOrderInfoTests(unittest.TestCase):
    def _run(self, order, **kwargs):
        data = {"orders": [order]}
        with patch.object(orders, "load", lambda: data), \
                patch.object(orders, "save", lambda d: None):
            return orders.update_order_info(order["serial"], order["product"], **kwargs), data

    def _order(self):
        return {"serial": "20261030-01", "product": "压片糖果（红色包衣）",
                "total_kg": 45, "tablet": 0.5, "wan": 9.0, "status": "待生产",
                "materials": [{"name": "甲", "kg": 30.0}, {"name": "乙", "kg": 15.0}]}

    def test_rename_and_rescale_with_history(self):
        order = self._order()
        result, data = self._run(order, new_product="虾青素压片糖果（红色包衣）", total_kg="90")
        self.assertTrue(result["changed"])
        saved = data["orders"][0]
        self.assertEqual(saved["product"], "虾青素压片糖果（红色包衣）")
        self.assertEqual(saved["total_kg"], 90)
        self.assertEqual([m["kg"] for m in saved["materials"]], [60.0, 30.0])
        self.assertEqual(saved["wan"], 18.0)
        self.assertEqual(len(saved["edit_history"]), 1)

    def test_unchanged_invalid_and_closed_are_rejected_or_noop(self):
        result, data = self._run(self._order(), new_product="压片糖果（红色包衣）", total_kg="45")
        self.assertFalse(result["changed"])
        self.assertEqual(data["orders"][0].get("edit_history", []), [])
        for bad in ("0", "-1", "abc", "nan"):
            with self.assertRaises(ValueError):
                self._run(self._order(), total_kg=bad)
        with self.assertRaises(ValueError):
            self._run(self._order(), new_product="  ")
        closed = self._order()
        closed["status"] = "已完成"
        with self.assertRaises(ValueError):
            self._run(closed, new_product="新名")

    def test_rename_collision_is_rejected(self):
        order = self._order()
        other = dict(self._order(), product="新名")
        data = {"orders": [order, other]}
        with patch.object(orders, "load", lambda: data), \
                patch.object(orders, "save", lambda d: None):
            with self.assertRaises(ValueError):
                orders.update_order_info(order["serial"], order["product"], new_product="新名")


class OpenPathSafetyTests(unittest.TestCase):
    def test_executable_and_script_files_are_refused(self):
        import app_gui
        with tempfile.TemporaryDirectory() as directory:
            for name in ("a.exe", "b.bat", "c.PS1"):
                path = os.path.join(directory, name)
                Path(path).write_text("x", encoding="utf-8")
                with patch.object(app_gui.os, "startfile", create=True) as start:
                    result = app_gui.open_path(path)
                self.assertFalse(result["ok"], name)
                start.assert_not_called()
            doc = os.path.join(directory, "d.md")
            Path(doc).write_text("x", encoding="utf-8")
            with patch.object(app_gui.os, "startfile", create=True) as start:
                self.assertTrue(app_gui.open_path(doc)["ok"])
                self.assertTrue(app_gui.open_path(directory)["ok"])
            self.assertEqual(start.call_count, 2)


if __name__ == "__main__":
    unittest.main()
