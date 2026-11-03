# -*- coding: utf-8 -*-
"""台账问答助手：只读工具行为 + 多轮循环协议（不调真实 AI）。"""
import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import ai_agent  # noqa: E402
import ai_assistant  # noqa: E402
import orders  # noqa: E402


def _ledger_order():
    return {
        "serial": "20261030-01", "product": "压片糖果（红色包衣）",
        "status": "生产中", "order_type": "正式生产", "batch": "02",
        "total_kg": 55, "due": "2026-11-10", "sales": "马士龙",
        "materials": [{"name": "碳酸钙", "kg": 1.0, "batch": "B2"}],
        "production_runs": [
            {"run_no": 1, "run_date": "20261030", "plan_kg": 45,
             "status": "已完成", "file": "第1批.xlsx",
             "materials": [{"name": "一水葡萄糖", "kg": 24.4, "batch": "B1"}]},
            {"run_no": 2, "run_date": "20261102", "plan_kg": 10,
             "file": "第2批.xlsx",
             "materials": [{"name": "碳酸钙", "kg": 1.0, "batch": "B2"}]},
        ],
        "tracking_history": [{"at": "2026-11-03T09:00:00", "source": "批次核销",
                              "run_no": 1}],
    }


class AgentToolTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.ledger_path = Path(self.directory.name) / "orders.json"
        self.ledger_path.write_text(
            json.dumps({"orders": [_ledger_order()]}, ensure_ascii=False),
            encoding="utf-8")
        patcher = patch.object(orders, "ORDERS", str(self.ledger_path))
        patcher.start()
        self.addCleanup(patcher.stop)

    def test_query_orders_filters_and_counts_runs(self):
        rows = ai_agent.tool_query_orders({"status": "生产中"})
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["runs_done"], 1)
        self.assertEqual(rows[0]["runs_total"], 2)
        self.assertEqual(ai_agent.tool_query_orders({"keyword": "不存在的"}), [])

    def test_order_detail_full_shape(self):
        detail = ai_agent.tool_order_detail({"serial": "20261030-01"})
        self.assertEqual(detail["total_kg"], 55)
        self.assertEqual(len(detail["production_runs"]), 2)
        self.assertEqual(detail["production_runs"][0]["status"], "已完成")
        self.assertEqual(detail["pending_materials"][0]["name"], "碳酸钙")

    def test_order_detail_bad_serial(self):
        self.assertIn("error", ai_agent.tool_order_detail({"serial": "x"}))

    def test_consumption_history_reads_csv(self):
        snap = Path(ai_agent.SNAPSHOTS)
        snap.mkdir(parents=True, exist_ok=True)
        (snap / "stock_history.csv").write_text(
            "date,short,qty,unit\n20261102,碳酸钙,4.5,kg\n20261103,碳酸钙,3.5,kg\n",
            encoding="utf-8")
        result = ai_agent.tool_consumption_history({"short": "碳酸钙", "days": 7})
        self.assertEqual([r["qty"] for r in result["history"]], ["4.5", "3.5"])
        self.assertIn("error", ai_agent.tool_consumption_history({"short": ""}))


class AgentLoopTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.ledger_path = Path(self.directory.name) / "orders.json"
        self.ledger_path.write_text(
            json.dumps({"orders": [_ledger_order()]}, ensure_ascii=False),
            encoding="utf-8")
        patcher = patch.object(orders, "ORDERS", str(self.ledger_path))
        patcher.start()
        self.addCleanup(patcher.stop)
        config_patcher = patch.object(
            ai_assistant, "is_configured", return_value=True)
        config_patcher.start()
        self.addCleanup(config_patcher.stop)

    def test_two_round_tool_then_answer(self):
        replies = [
            '{"tool": "order_detail", "args": {"serial": "20261030-01"}}',
            '{"answer": "该批共 2 批，累计 55kg。"}',
        ]
        with patch.object(ai_assistant, "chat", side_effect=replies):
            result = ai_agent.ask("20261030-01 做了几批？")
        self.assertTrue(result["ok"])
        self.assertEqual(result["answer"], "该批共 2 批，累计 55kg。")
        self.assertEqual(result["rounds"], 2)
        self.assertEqual(result["tools_used"][0]["tool"], "order_detail")

    def test_unknown_tool_gets_corrected(self):
        replies = [
            '{"tool": "不存在的", "args": {}}',
            '{"tool": "query_orders", "args": {}}',
            '{"answer": "有 1 个单。"}',
        ]
        with patch.object(ai_assistant, "chat", side_effect=replies):
            result = ai_agent.ask("现在有哪些单？")
        self.assertTrue(result["ok"])
        self.assertEqual(len(result["tools_used"]), 1)

    def test_round_limit_reached(self):
        with patch.object(ai_assistant, "chat",
                          return_value='{"tool": "query_orders", "args": {}}'):
            result = ai_agent.ask("一直查")
        self.assertTrue(result["ok"])
        self.assertIn("上限", result["answer"])

    def test_invalid_json_surfaces_error(self):
        with patch.object(ai_assistant, "chat", return_value="我不会 JSON"):
            result = ai_agent.ask("问题")
        self.assertFalse(result["ok"])
        self.assertIn("格式异常", result["msg"])

    def test_unconfigured_degrades(self):
        with patch.object(ai_assistant, "is_configured", return_value=False):
            result = ai_agent.ask("问题")
        self.assertFalse(result["ok"])
        self.assertIn("未配置", result["msg"])


if __name__ == "__main__":
    unittest.main()
