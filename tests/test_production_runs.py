# -*- coding: utf-8 -*-
"""多批次生产：批次核销/撤销、出单登记第 1 批、导入预览、批次核对表、库存快照。"""
import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import app_gui  # noqa: E402
import orders  # noqa: E402
from sales_form_import import preview_sales_form  # noqa: E402

try:
    from test_sales_form_import import make_sales_form  # noqa: E402
except ImportError:  # 直接以 tests.test_production_runs 运行时
    from tests.test_sales_form_import import make_sales_form  # noqa: E402


def _write_ledger(path, order_list):
    path.write_text(json.dumps({"orders": order_list}, ensure_ascii=False),
                    encoding="utf-8")


def _two_run_order():
    return {
        "serial": "20261030-01", "product": "压片糖果（红色包衣）",
        "status": "待生产", "order_type": "正式生产", "batch": "02",
        "total_kg": 55,
        "materials": [{"name": "碳酸钙", "kg": 1.0, "batch": "B2"}],
        "production_runs": [
            {"run_no": 1, "run_date": "20261030", "plan_kg": 45,
             "file": "第1批.xlsx", "sha256": "",
             "materials": [{"name": "一水葡萄糖", "kg": 24.4, "batch": "B1"}],
             "note": "历史导入补记"},
            {"run_no": 2, "run_date": "20261102", "plan_kg": 10,
             "file": "第2批.xlsx", "sha256": "",
             "materials": [{"name": "碳酸钙", "kg": 1.0, "batch": "B2"},
                           {"name": "麦芽糊精", "kg": 3.0, "batch": "B2"}],
             "note": "销售单导入"},
        ],
    }


class ProductionRunTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.ledger_path = Path(self.directory.name) / "orders.json"
        patcher = patch.object(orders, "ORDERS", str(self.ledger_path))
        patcher.start()
        self.addCleanup(patcher.stop)

    def _ledger(self):
        return json.loads(self.ledger_path.read_text(encoding="utf-8"))

    def test_mark_run_releases_reservation_and_advances_status(self):
        _write_ledger(self.ledger_path, [_two_run_order()])
        result = orders.mark_production_run("20261030-01", 1, True)
        self.assertTrue(result["ok"])
        ledger = self._ledger()["orders"][0]
        self.assertEqual(ledger["status"], "生产中")
        self.assertEqual(
            [(m["name"], m["kg"]) for m in ledger["materials"]],
            [("碳酸钙", 1.0), ("麦芽糊精", 3.0)])
        self.assertEqual(ledger["production_runs"][0]["status"], "已完成")
        self.assertEqual(ledger["tracking_history"][-1]["run_no"], 1)
        self.assertEqual(ledger["tracking_history"][-1]["source"], "批次核销")

    def test_mark_all_runs_done_auto_completes_order(self):
        _write_ledger(self.ledger_path, [_two_run_order()])
        orders.mark_production_run("20261030-01", 1, True)
        orders.mark_production_run("20261030-01", 2, True)
        ledger = self._ledger()["orders"][0]
        self.assertEqual(ledger["status"], "已完成")
        self.assertEqual(ledger["materials"], [])
        self.assertEqual(ledger["status_history"][-1]["reason"],
                         "全部生产批次已核销")

    def test_undo_run_reopens_and_restores_reservation(self):
        _write_ledger(self.ledger_path, [_two_run_order()])
        orders.mark_production_run("20261030-01", 1, True)
        orders.mark_production_run("20261030-01", 2, True)
        result = orders.mark_production_run("20261030-01", 2, False)
        self.assertTrue(result["ok"])
        ledger = self._ledger()["orders"][0]
        self.assertEqual(ledger["status"], "生产中")
        self.assertEqual(
            [(m["name"], m["kg"]) for m in ledger["materials"]],
            [("碳酸钙", 1.0), ("麦芽糊精", 3.0)])
        self.assertNotIn("status", ledger["production_runs"][1])

    def test_mark_run_unknown_serial_and_run_fail(self):
        _write_ledger(self.ledger_path, [_two_run_order()])
        with self.assertRaises(ValueError):
            orders.mark_production_run("20269999-01", 1, True)
        with self.assertRaises(ValueError):
            orders.mark_production_run("20261030-01", 9, True)

    def test_set_order_file_registers_first_run(self):
        _write_ledger(self.ledger_path, [{
            "serial": "20261029-01", "product": "INZENT虾青素压片糖果",
            "status": "待生产", "order_type": "测试单", "batch": "01",
            "order_date": "20261029", "materials": [],
        }])
        orders.set_order_file("20261029-01", "INZENT虾青素压片糖果",
                              r"C:\x\20261029-01_测试.xlsx", {
                                  "total_kg": 10,
                                  "materials": [{"name": "一水葡萄糖",
                                                 "kg": 5.43, "batch": "B1"}],
                              })
        ledger = self._ledger()["orders"][0]
        self.assertEqual(len(ledger["production_runs"]), 1)
        first = ledger["production_runs"][0]
        self.assertEqual(first["run_no"], 1)
        self.assertEqual(first["note"], "出单生成")
        self.assertEqual(first["file"], "20261029-01_测试.xlsx")
        self.assertEqual(first["materials"][0]["kg"], 5.43)
        self.assertEqual(ledger["materials"][0]["kg"], 5.43)
        # 幂等：再次设置文件不重复登记批次
        orders.set_order_file("20261029-01", "INZENT虾青素压片糖果",
                              r"C:\x\20261029-01_测试.xlsx", None)
        self.assertEqual(len(self._ledger()["orders"][0]["production_runs"]), 1)


class SalesPreviewTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.ledger_path = Path(self.directory.name) / "orders.json"
        patcher = patch.object(orders, "ORDERS", str(self.ledger_path))
        patcher.start()
        self.addCleanup(patcher.stop)

    def test_preview_suggests_attachable_serial_and_next(self):
        _write_ledger(self.ledger_path, [{
            "serial": "20261028-01", "product": "压片糖果（红色包衣）",
            "status": "待生产", "order_type": "正式生产", "total_kg": 45,
        }])
        source = Path(self.directory.name) / "虾青素配方单-20261028.docx"
        make_sales_form(source)
        result = preview_sales_form(source)
        self.assertTrue(result["ok"])
        self.assertEqual(result["product"], "压片糖果（红色包衣）")
        self.assertEqual(result["tablet_g"], 0.5)
        self.assertEqual(result["per_bottle"], 30)
        self.assertEqual(result["suggest_serial"], "20261028-01")
        self.assertEqual(result["next_serial"], "20261028-02")
        self.assertEqual(len(result["ledger_day_orders"]), 1)
        self.assertTrue(result["ledger_day_orders"][0]["attachable"])

    def test_preview_no_day_orders_gives_next_serial(self):
        _write_ledger(self.ledger_path, [{
            "serial": "20261030-01", "product": "压片糖果（红色包衣）",
            "status": "待生产",
        }])
        source = Path(self.directory.name) / "虾青素配方单-20261101.docx"
        make_sales_form(source)
        result = preview_sales_form(source)
        self.assertEqual(result["order_date"], "20261101")
        self.assertEqual(result["suggest_serial"], "")
        self.assertEqual(result["next_serial"], "20261101-01")


class BatchAuditTests(unittest.TestCase):
    def test_batch_audit_rows_split_missing_and_mismatch(self):
        data = {"items": [
            {"short": "碳酸钙", "cat": "原材料-填充剂", "qty": 4.5, "unit": "kg",
             "batches": []},
            {"short": "麦芽糊精", "cat": "原材料-填充剂", "qty": 15.75,
             "unit": "kg", "batches": [{"batch": "A", "qty": 10.0}]},
            {"short": "一水葡萄糖", "cat": "原材料-填充剂", "qty": 24.4,
             "unit": "kg", "batches": [{"batch": "B", "qty": 24.4}]},
            {"short": "成品瓶", "cat": "产成品-包材", "qty": 5, "unit": "个",
             "batches": []},
        ]}
        missing, mismatch = app_gui._batch_audit_rows(data)
        self.assertEqual([m["name"] for m in missing], ["碳酸钙"])
        self.assertEqual([m["name"] for m in mismatch], ["麦芽糊精"])
        self.assertEqual(mismatch[0]["diff"], -5.75)


class StockSnapshotTests(unittest.TestCase):
    def test_daily_snapshot_writes_json_and_csv_once(self):
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            stock = {"stock_date": "2026-11-02", "items": [
                {"short": "碳酸钙", "qty": 4.5, "unit": "kg",
                 "cat": "原材料-填充剂"},
                {"short": "成品瓶", "qty": 500, "unit": "个", "cat": "产成品-包材"},
            ]}
            (base / "stock.json").write_text(
                json.dumps(stock, ensure_ascii=False), encoding="utf-8")
            import importlib
            daily_backup = importlib.import_module("maintenance.daily_backup")
            if not hasattr(daily_backup, "write_stock_snapshot"):
                raise ImportError("daily_backup 未加载")
            path = daily_backup.write_stock_snapshot(base=str(base))
            self.assertTrue(Path(path).exists())
            payload = json.loads(Path(path).read_text(encoding="utf-8"))
            self.assertEqual(len(payload["items"]), 1)  # 产成品不入快照
            csv_path = base / "data" / "snapshots" / "stock_history.csv"
            rows = csv_path.read_text(encoding="utf-8").strip().splitlines()
            self.assertEqual(len(rows), 2)  # 表头 + 1 行
            # 同日重复执行不追加重复行
            daily_backup.write_stock_snapshot(base=str(base))
            rows = csv_path.read_text(encoding="utf-8").strip().splitlines()
            self.assertEqual(len(rows), 2)


if __name__ == "__main__":
    unittest.main()
