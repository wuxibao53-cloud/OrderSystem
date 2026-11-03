# -*- coding: utf-8 -*-
"""AI 提取层：确定性校验器、AI 表单 → 交接单导入、JSON 抽取容错（不调真实 AI）。"""
import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import orders  # noqa: E402
import ai_assistant  # noqa: E402
from sales_form_import import (form_from_ai, import_sales_form,  # noqa: E402
                               validate_ai_form)

try:
    from test_sales_form_import import make_template  # noqa: E402
except ImportError:  # 直接以 tests.test_ai_assistant 运行时
    from tests.test_sales_form_import import make_template  # noqa: E402


GOOD_FORM = {
    "product": "压片糖果（红色包衣）",
    "order_date": "20261105",
    "sales": "马士龙",
    "order_no": "",
    "due_date": "",
    "tablet_g": 0.5,
    "per_bottle": 30,
    "coat_type": "红色包衣",
    "materials": [
        {"name": "一水葡萄糖", "ratio": 43.5, "quote": "一水葡萄糖 43.5"},
        {"name": "麦芽糊精", "ratio": 30, "quote": "麦芽糊精 30"},
        {"name": "山梨糖醇", "ratio": 15, "quote": "山梨糖醇 15"},
        {"name": "碳酸钙", "ratio": 10, "quote": "碳酸钙 10"},
        {"name": "硬脂酸镁", "ratio": 1.5, "quote": "硬脂酸镁 1.5"},
    ],
    "coat": {"name": "复配被膜剂（红色包衣）", "ratio": 3,
             "quote": "复配被膜剂（红色包衣） 3%（按素片重量计）"},
    "packaging": [{"kind": "瓶", "detail": "茶瓶150cc"}],
    "press_process": "压片：\n1.配料：检查原辅料是否结块\n2.混料：预混10-15 min",
    "coat_process": "",
}


class AiFormValidationTests(unittest.TestCase):
    def test_good_form_passes(self):
        problems, warnings, unknown = validate_ai_form(
            GOOD_FORM, stock_names=["一水葡萄糖", "麦芽糊精", "山梨糖醇",
                                    "碳酸钙", "硬脂酸镁"])
        self.assertEqual(problems, [])
        self.assertEqual(unknown, [])

    def test_ratio_sum_off_by_row_is_problem(self):
        bad = dict(GOOD_FORM)
        bad["materials"] = [{"name": "一水葡萄糖", "ratio": 60, "quote": "x"},
                            {"name": "麦芽糊精", "ratio": 30, "quote": "y"}]
        problems, _w, _u = validate_ai_form(bad)
        self.assertTrue(any("偏离 100%" in p for p in problems))

    def test_slight_ratio_drift_is_warning_only(self):
        drifted = json.loads(json.dumps(GOOD_FORM))
        drifted["materials"][0]["ratio"] = 44.0
        problems, warnings, _u = validate_ai_form(drifted)
        self.assertEqual(problems, [])
        self.assertTrue(any("请核对" in w for w in warnings))

    def test_unknown_material_flagged(self):
        problems, warnings, unknown = validate_ai_form(
            GOOD_FORM, stock_names=["一水葡萄糖", "麦芽糊精"])
        self.assertEqual(problems, [])
        self.assertIn("山梨糖醇", unknown)
        self.assertTrue(any("找不到" in w for w in warnings))

    def test_missing_quote_is_warning(self):
        no_quote = json.loads(json.dumps(GOOD_FORM))
        no_quote["materials"][2]["quote"] = ""
        _p, warnings, _u = validate_ai_form(no_quote)
        self.assertTrue(any("山梨糖醇」没有原文出处" in w for w in warnings))

    def test_bad_tablet_and_date_are_problems(self):
        bad = json.loads(json.dumps(GOOD_FORM))
        bad["tablet_g"] = 9.9
        bad["order_date"] = "20261399"
        problems, _w, _u = validate_ai_form(bad)
        self.assertTrue(any("片重" in p for p in problems))
        self.assertTrue(any("不是真实日期" in p for p in problems))

    def test_form_from_ai_shape_matches_parser(self):
        form = form_from_ai(GOOD_FORM, source_name="AI 单")
        self.assertEqual(form["product"], "压片糖果（红色包衣）")
        self.assertEqual(form["tablet_g"], 0.5)
        self.assertEqual(form["per_bottle"], 30)
        self.assertEqual(form["coat_name"], "复配被膜剂（红色包衣）")
        self.assertEqual(form["coat_ratio"], 3)
        self.assertEqual(form["coat_type"], "红色包衣")
        self.assertAlmostEqual(form["ratio_total"], 100.0)
        self.assertEqual([m["name"] for m in form["materials"]],
                         ["一水葡萄糖", "麦芽糊精", "山梨糖醇", "碳酸钙", "硬脂酸镁"])
        self.assertTrue(form["ai_extracted"])
        self.assertIn("配料", form["press_process"])


class AiOverrideImportTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.ledger_path = Path(self.directory.name) / "orders.json"
        self.ledger_path.write_text(json.dumps({"orders": []}, ensure_ascii=False),
                                    encoding="utf-8")
        patcher = patch.object(orders, "ORDERS", str(self.ledger_path))
        patcher.start()
        self.addCleanup(patcher.stop)
        self.template = Path(self.directory.name) / "template.xlsx"
        make_template(self.template)

    def test_text_order_generates_sheet_and_ledger_without_docx(self):
        result = import_sales_form(
            "文字订单_20261105.docx", self.template,
            history_root=Path(self.directory.name) / "10月",
            options={"form_override": GOOD_FORM, "plan_qty": 10, "plan_unit": "kg",
                     "sequence": 1},
            stock_rows=[{"short": "一水葡萄糖", "batches": [
                {"batch": "20260917", "qty": 6, "in_date": "2026-09-17"}]}])
        self.assertIn("output", result)
        sheet_name = Path(result["output"]).name
        self.assertIn("20261105-01", sheet_name)
        self.assertIn("销售单导入_10kg", sheet_name)
        self.assertNotIn("第02批", sheet_name)
        ledger = json.loads(self.ledger_path.read_text(encoding="utf-8"))
        created = ledger["orders"][0]
        self.assertEqual(created["serial"], "20261105-01")
        self.assertEqual(created["sales"], "马士龙")
        self.assertEqual(created["total_kg"], 10)
        self.assertEqual(len(created["production_runs"]), 1)
        self.assertEqual(created["production_runs"][0]["note"], "销售单导入")

    def test_override_with_bad_ratios_rejected_before_any_file(self):
        bad = json.loads(json.dumps(GOOD_FORM))
        bad["materials"] = [{"name": "一水葡萄糖", "ratio": 90, "quote": "x"}]
        with self.assertRaises(ValueError):
            import_sales_form(
                "文字订单_20261105.docx", self.template,
                history_root=Path(self.directory.name) / "10月",
                options={"form_override": bad, "plan_qty": 10, "plan_unit": "kg"},
                stock_rows=[])
        ledger = json.loads(self.ledger_path.read_text(encoding="utf-8"))
        self.assertEqual(ledger["orders"], [])


class ExtractJsonTests(unittest.TestCase):
    def test_extracts_from_fenced_and_noisy_reply(self):
        text = "好的，以下是结果：\n```json\n{\"a\": 1, \"b\": {\"c\": \"}\"}}\n```\n以上。"
        self.assertEqual(ai_assistant._extract_json(text)["b"]["c"], "}")

    def test_incomplete_json_raises(self):
        with self.assertRaises(ai_assistant.AiError):
            ai_assistant._extract_json("{\"a\": 1,")

    def test_extract_json_retries_then_succeeds(self):
        calls = {"n": 0}

        def fake_chat(messages, config=None):
            calls["n"] += 1
            if calls["n"] == 1:
                return "这不是 JSON"
            return "{\"ok\": true}"

        with patch.object(ai_assistant, "chat", fake_chat):
            data, meta = ai_assistant.extract_json("s", "u", validate=lambda d: None)
        self.assertTrue(data["ok"])
        self.assertEqual(meta["attempts"], 2)


class PressSplitTests(unittest.TestCase):
    PRESS_TEXT = ("1.配料：检查原辅料是否结块 2.再加硬脂酸镁混合 2 min "
                  "3.压片：需震动下料；单片克重控制在0.48-0.52 g；"
                  "硬度要求用力刮表面或边缘无粉末")

    def test_split_separates_press_from_mixing(self):
        from sales_form_import import _split_press_sections
        mix, press = _split_press_sections(self.PRESS_TEXT)
        self.assertIn("再加硬脂酸镁混合 2 min", mix)
        self.assertNotIn("压片", mix)
        self.assertTrue(press.startswith("压片："))
        self.assertIn("震动下料", press)
        self.assertIn("0.48-0.52", press)

    def test_split_noop_without_press_marker(self):
        from sales_form_import import _split_press_sections
        mix, press = _split_press_sections("先加一水葡萄糖预混 10 min")
        self.assertIn("预混", mix)
        self.assertEqual(press, "")


class AiFormSalesFieldTests(unittest.TestCase):
    def test_form_from_ai_carries_sales_order_due(self):
        payload = json.loads(json.dumps(GOOD_FORM))
        payload["sales"] = "马士龙"
        payload["order_no"] = "SO-2026-1105"
        payload["due_date"] = "2026.11.10"
        form = form_from_ai(payload, source_name="AI 单")
        self.assertEqual(form["sales"], "马士龙")
        self.assertEqual(form["order_no"], "SO-2026-1105")
        self.assertEqual(form["due"], "20261110")


if __name__ == "__main__":
    unittest.main()
