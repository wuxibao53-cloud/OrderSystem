# -*- coding: utf-8 -*-
"""柠檬云同步：名称归一 / 核对 / 安全修正决策 / 采购预测（纯函数，不碰真实文件）。"""
import importlib
import json
import sys
import unittest
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
try:  # 优先用环境里的 xlrd；没有则用项目 vendor 的 libs/xlrd
    import xlrd  # noqa: F401
except ImportError:
    sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "libs"))

lemon_sync = importlib.import_module("lemon_sync")


def _item(code, name, short=None, qty=0.0, unit="kg", cat="原材料-直接原料",
          batches=None):
    return {"code": code, "name": name, "short": short or lemon_sync.normalize_name(name),
            "cat": cat, "unit": unit, "qty": qty, "safe": 0, "batches": batches or []}


def _bal(code, name, qty, unit="kg"):
    return {"code": code, "name": name, "cat": "原材料-直接原料", "unit": unit,
            "qty": qty, "cost": None}


class NormalizeTests(unittest.TestCase):
    def test_category_prefixes_stripped(self):
        self.assertEqual(lemon_sync.normalize_name("直接原料-玉米淀粉"), "玉米淀粉")
        self.assertEqual(lemon_sync.normalize_name("添加剂-硬脂酸镁"), "硬脂酸镁")
        self.assertEqual(lemon_sync.normalize_name("辅料-维生素B12"), "维生素B12")
        self.assertEqual(lemon_sync.normalize_name("虾青素压片糖果"), "虾青素压片糖果")

    def test_units(self):
        self.assertEqual(lemon_sync.norm_unit("KG"), "kg")
        self.assertEqual(lemon_sync.norm_unit("千克"), "kg")
        self.assertEqual(lemon_sync.norm_unit("瓶"), "瓶")


class ReconcileTests(unittest.TestCase):
    def setUp(self):
        self.items = [
            _item("0201000002", "直接原料-一水葡萄糖", qty=55.19),
            _item("0201000003", "直接原料-麦芽糊精", qty=40.23),
        ]

    def test_exact_code_match_reports_diff(self):
        balance = [_bal("0201000002", "直接原料-一水葡萄糖", 52.0),
                   _bal("0201000003", "直接原料-麦芽糊精", 40.23)]
        report = lemon_sync.reconcile(self.items, balance, [])
        self.assertEqual(len(report["diffs"]), 1)
        diff = report["diffs"][0]
        self.assertEqual(diff["short"], "一水葡萄糖")
        self.assertAlmostEqual(diff["diff"], -3.19)
        self.assertTrue(diff["exact"])

    def test_name_fallback_marks_not_exact(self):
        balance = [_bal("9999", "直接原料-一水葡萄糖", 50.0)]
        report = lemon_sync.reconcile(self.items, balance, [])
        self.assertEqual(len(report["diffs"]), 1)
        self.assertFalse(report["diffs"][0]["exact"])

    def test_unit_conflict_not_a_diff(self):
        balance = [_bal("0201000002", "直接原料-一水葡萄糖", 50.0, unit="瓶")]
        report = lemon_sync.reconcile(self.items, balance, [])
        self.assertEqual(report["diffs"], [])
        self.assertEqual(len(report["unit_conflicts"]), 1)

    def test_new_and_missing_flagged(self):
        balance = [_bal("99990001", "辅料-新颖辅料", 5.0)]
        report = lemon_sync.reconcile(self.items, balance, [])
        self.assertEqual(len(report["new_in_lemon"]), 1)
        self.assertEqual(len(report["missing_in_lemon"]), 2)


class ApplyFixTests(unittest.TestCase):
    def test_fix_only_exact_match_with_backup_history(self):
        stock_data = {"items": [
            _item("0201000002", "直接原料-一水葡萄糖", qty=55.19),
            _item("88880001", "辅料-神秘物料", qty=1.0),
        ]}
        report = {"diffs": [
            {"code": "0201000002", "name": "直接原料-一水葡萄糖",
             "short": "一水葡萄糖", "sys_qty": 55.19, "lemon_qty": 52.0,
             "diff": -3.19, "exact": True},
            {"code": "88880001", "name": "辅料-神秘物料", "short": "神秘物料",
             "sys_qty": 1.0, "lemon_qty": 9.9, "diff": 8.9, "exact": False},
        ], "batch_facts": {}}
        cfg = {"auto_fix": True}
        fixes, warnings = lemon_sync._apply_fixes(stock_data, report, cfg, [])
        self.assertEqual(len(fixes), 1)
        self.assertIn("52", fixes[0])
        self.assertEqual(stock_data["items"][0]["qty"], 52.0)
        self.assertEqual(stock_data["items"][1]["qty"], 1.0)  # 模糊匹配不动
        self.assertTrue(any("未自动修改" in w for w in warnings))

    def test_fill_single_batch_with_mfg_date(self):
        stock_data = {"items": [
            _item("0201000002", "直接原料-一水葡萄糖", qty=55.19, batches=[]),
        ]}
        report = {"diffs": [
            {"code": "0201000002", "name": "直接原料-一水葡萄糖",
             "short": "一水葡萄糖", "sys_qty": 55.19, "lemon_qty": 52.0,
             "diff": -3.19, "exact": True}],
            "batch_facts": {"直接原料-一水葡萄糖": [
                {"name": "直接原料-一水葡萄糖", "batch": "20260917",
                 "mfg": "2026-09-17", "exp": "2027-09-17",
                 "doc_date": "2026-11-02", "biz": "组装出库",
                 "warehouse": "直接原料库", "qty": 4.35}]}}
        balance = [{"code": "0201000002", "name": "直接原料-一水葡萄糖"}]
        cfg = {"auto_fix": True}
        fixes, warnings = lemon_sync._apply_fixes(stock_data, report, cfg, balance)
        batches = stock_data["items"][0]["batches"]
        self.assertEqual(len(batches), 1)
        self.assertEqual(batches[0]["batch"], "20260917")
        self.assertEqual(batches[0]["qty"], 52.0)
        self.assertEqual(batches[0]["mfg"], "2026-09-17")
        self.assertTrue(any("补批次" in f for f in fixes))
        self.assertFalse([w for w in warnings if "批次明细差" in w])

    def test_report_mode_changes_nothing(self):
        stock_data = {"items": [_item("0201000002", "直接原料-一水葡萄糖", qty=55.19)]}
        cfg = {"auto_fix": False}
        report = {"diffs": [], "batch_facts": {}}
        fixes, warnings = lemon_sync._apply_fixes(stock_data, report, cfg, [])
        self.assertEqual(fixes, [])
        self.assertEqual(stock_data["items"][0]["qty"], 55.19)


class ForecastTests(unittest.TestCase):
    def test_consumption_rate_and_committed_need(self):
        items = [_item("0201000002", "直接原料-一水葡萄糖", qty=10.0),
                 _item("0201000003", "直接原料-麦芽糊精", qty=100.0)]
        orders = [
            {"status": "已完成", "order_date": "20261028",
             "production_runs": [{"materials": [
                 {"name": "一水葡萄糖", "kg": 24.4, "batch": "B1"}]}],
             "materials": []},
            {"status": "待生产", "order_date": "20261105",
             "production_runs": [],
             "materials": [{"name": "一水葡萄糖", "kg": 4.35, "batch": "B1"}]},
        ]
        result = lemon_sync.forecast(items, orders, horizon_days=14,
                                     recent_days=60)
        rows = {r["name"]: r for r in result["rows"]}
        self.assertIn("一水葡萄糖", rows)
        row = rows["一水葡萄糖"]
        self.assertAlmostEqual(row["committed"], 4.35)
        self.assertGreater(row["suggest"], 0)
        self.assertTrue(row["urgent"])  # 待产占用 4.35 > 账面 10? no; cover=10/24.4*60span <7 → check
        self.assertNotIn("麦芽糊精", rows)  # 无消耗无占用 → 不出现

    def test_products_excluded(self):
        items = [_item("0101000003", "赤藓糖醇控制能量片", qty=644, unit="瓶",
                       cat="产成品-运动营养食品")]
        result = lemon_sync.forecast(items, [])
        self.assertEqual(result["rows"], [])


if __name__ == "__main__":
    unittest.main()
