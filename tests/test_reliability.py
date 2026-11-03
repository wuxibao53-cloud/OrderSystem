import json
import os
import tempfile
import unittest
import warnings
import base64
import datetime
import io
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook, load_workbook
import orders
import stock
import generator
import backup_manager
import make_package
from json_store import save_json_atomic


class Table:
    def __init__(self, rows):
        self.rows = rows
        self.nrows = len(rows)
        self.ncols = max((len(row) for row in rows), default=0)

    def cell_value(self, row, column):
        try:
            value = self.rows[row][column]
        except IndexError:
            return ""
        return "" if value is None else value


def stock_table(headers=None, product_row=None):
    headers = headers or [""] * 43
    headers[1] = "商品编码"
    headers[2] = "商品名称"
    headers[3] = "商品类别"
    headers[4] = "规格型号"
    headers[5] = "单位"
    headers[8] = "库存数量"
    if product_row is None:
        product_row = [""] * 43
        product_row[1] = "A001"
        product_row[2] = "直接原料-测试物料"
        product_row[3] = "原材料-直接原料"
        product_row[4] = "测试规格"
        product_row[5] = "KG"
        product_row[8] = 12.5
    return Table([["库存余额表"], [""], headers, product_row])


class AtomicJsonTests(unittest.TestCase):
    def test_header_detection_prefers_field_header_over_populated_subheader(self):
        table = Table([
            ["商品库存余额表"],
            [""],
            ["公司：测试公司 库存日期：2026-10-29"],
            ["序号", "商品编码", "商品名称", "商品类别", "规格型号", "单位", "库存单位", "", "全部仓库"],
            ["", "", "", "", "", "", "", "", "单位数量", "库存单位数量", "单位成本"],
        ])
        self.assertEqual(stock.find_header_row(table, ["商品名称", "物料名称"], maxr=8), 3)

    def test_serialization_failure_keeps_previous_file_and_cleans_temp(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "state.json"
            path.write_text('{"saved": true}', encoding="utf-8")

            with self.assertRaises(TypeError):
                save_json_atomic(path, {"not_json": object()})

            self.assertEqual(path.read_text(encoding="utf-8"), '{"saved": true}')
            self.assertEqual(list(Path(directory).iterdir()), [path])

    def test_stock_import_rejects_unrecognized_layout_without_replacing_data(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "stock.json"
            original = {"items": [{"name": "原库存", "qty": 9}]}
            stock.save_json(path, original)
            previous_stock_path = stock.STOCK
            stock.STOCK = str(path)
            try:
                with patch.object(stock, "open_table", return_value=Table([["不是柠檬云库存表"]])):
                    with self.assertRaisesRegex(ValueError, "未修改现有库存"):
                        stock.import_xls("invalid.xlsx")
                self.assertEqual(stock.load_json(path), original)
            finally:
                stock.STOCK = previous_stock_path

    def test_stock_import_requires_expected_positional_columns(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "stock.json"
            original = {"items": [{"name": "原库存", "qty": 9}]}
            stock.save_json(path, original)
            previous_stock_path = stock.STOCK
            stock.STOCK = str(path)
            try:
                table = stock_table()
                table.rows[2][8] = "可用量"
                with patch.object(stock, "open_table", return_value=table):
                    with self.assertRaisesRegex(ValueError, "列布局"):
                        stock.import_xls("wrong-columns.xlsx")
                self.assertEqual(stock.load_json(path), original)
            finally:
                stock.STOCK = previous_stock_path

    def test_stock_import_rejects_non_numeric_balance_without_replacing_data(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "stock.json"
            original = {"_meta": {"stock_date": "2026-10-27"},
                        "items": [{"name": "原库存", "qty": 9}]}
            stock.save_json(path, original)
            previous_stock_path = stock.STOCK
            stock.STOCK = str(path)
            table = stock_table()
            table.rows[3][8] = "未知"
            try:
                with patch.object(stock, "open_table", return_value=table):
                    with self.assertRaisesRegex(ValueError, "数量无效"):
                        stock.import_xls("invalid-quantity.xlsx")
                self.assertEqual(stock.load_json(path), original)
            finally:
                stock.STOCK = previous_stock_path

    def test_stock_import_rejects_older_snapshot_without_replacing_data(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "stock.json"
            original = {"_meta": {"stock_date": "2026-10-29"},
                        "items": [{"name": "原库存", "qty": 9}]}
            stock.save_json(path, original)
            previous_stock_path = stock.STOCK
            stock.STOCK = str(path)
            table = stock_table()
            table.rows[1][0] = "公司：测试公司 库存日期：2026-10-28"
            try:
                with patch.object(stock, "open_table", return_value=table):
                    with self.assertRaisesRegex(ValueError, "早于当前账面日期"):
                        stock.import_xls("older.xlsx")
                self.assertEqual(stock.load_json(path), original)
            finally:
                stock.STOCK = previous_stock_path

    def test_stock_comparison_rejects_invalid_order_ledger_before_import(self):
        with tempfile.TemporaryDirectory() as directory:
            stock_path = Path(directory) / "stock.json"
            original = {"_meta": {"stock_date": "2026-10-27"},
                        "items": [{"name": "原库存", "qty": 9}]}
            stock.save_json(stock_path, original)
            (Path(directory) / "orders.json").write_text("{invalid", encoding="utf-8")
            previous_stock_path = stock.STOCK
            stock.STOCK = str(stock_path)
            try:
                with patch.object(stock, "APP", directory), \
                        patch.object(stock, "open_table") as open_table:
                    with self.assertRaises(json.JSONDecodeError):
                        stock.compare_stock("new-stock.xlsx")
                    open_table.assert_not_called()
                self.assertEqual(stock.load_json(stock_path), original)
            finally:
                stock.STOCK = previous_stock_path

    def test_empty_recognized_export_does_not_clear_existing_stock(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "stock.json"
            original = {"items": [{"name": "原库存", "qty": 9}]}
            stock.save_json(path, original)
            previous_stock_path = stock.STOCK
            stock.STOCK = str(path)
            try:
                with patch.object(stock, "open_table", return_value=stock_table(product_row=[])):
                    with self.assertRaisesRegex(ValueError, "没有可识别的商品行"):
                        stock.import_xls("empty.xlsx")
                self.assertEqual(stock.load_json(path), original)
            finally:
                stock.STOCK = previous_stock_path

    def test_valid_stock_import_preserves_manual_thresholds_and_batches(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "stock.json"
            stock.save_json(path, {"items": [{
                "name": "直接原料-测试物料",
                "safe": 3,
                "batches": [{"batch": "LOT-1", "qty": 12.5}],
            }]})
            previous_stock_path = stock.STOCK
            stock.STOCK = str(path)
            try:
                with patch.object(stock, "open_table", return_value=stock_table()):
                    count, _, _ = stock.import_xls("valid.xlsx")
                self.assertEqual(count, 1)
                item = stock.load_json(path)["items"][0]
                self.assertEqual(item["qty"], 12.5)
                self.assertEqual(item["safe"], 3)
                self.assertEqual(item["batches"], [{"batch": "LOT-1", "qty": 12.5}])
            finally:
                stock.STOCK = previous_stock_path

    def test_xlsx_stock_export_is_read_and_imported(self):
        with tempfile.TemporaryDirectory() as directory:
            workbook_path = Path(directory) / "库存余额.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.append(["库存余额表"])
            sheet.append(["公司：测试公司 库存日期：2026-10-28"])
            headers = [""] * 43
            headers[1] = "商品编码"
            headers[2] = "商品名称"
            headers[3] = "商品类别"
            headers[4] = "规格型号"
            headers[5] = "单位"
            headers[8] = "库存数量"
            sheet.append(headers)
            row = [""] * 43
            row[1] = "A001"
            row[2] = "直接原料-测试物料"
            row[3] = "原材料-直接原料"
            row[4] = "测试规格"
            row[5] = "KG"
            row[8] = 12.5
            sheet.append(row)
            workbook.save(workbook_path)
            workbook.close()

            previous_stock_path = stock.STOCK
            stock.STOCK = os.path.join(directory, "stock.json")
            try:
                count, company, stock_date = stock.import_xls(workbook_path)
                self.assertEqual((count, company, stock_date), (1, "测试公司", "2026-10-28"))
                self.assertEqual(stock.load_json(stock.STOCK)["items"][0]["qty"], 12.5)
            finally:
                stock.STOCK = previous_stock_path

    def test_multirow_stock_header_uses_all_warehouse_quantity_subheader(self):
        table_rows = [
            ["商品库存余额表"],
            [""],
            ["公司：测试公司 库存日期：2026-10-29"],
        ]
        headers = [""] * 43
        headers[1:6] = ["商品编码", "商品名称", "商品类别", "规格型号", "单位"]
        headers[8] = "全部仓库"
        table_rows.append(headers)
        subheaders = [""] * 43
        subheaders[8:11] = ["单位数量", "库存单位数量", "多单位数量"]
        table_rows.append(subheaders)
        row = [""] * 43
        row[1:6] = ["A001", "直接原料-测试物料", "原材料-直接原料", "测试规格", "KG"]
        row[8] = 12.5
        row[23] = 12.5
        table_rows.append(row)

        with tempfile.TemporaryDirectory() as directory:
            previous_stock_path = stock.STOCK
            stock.STOCK = os.path.join(directory, "stock.json")
            try:
                with patch.object(stock, "open_table", return_value=Table(table_rows)):
                    count, company, stock_date = stock.import_xls("multirow.xlsx")
                item = stock.load_json(stock.STOCK)["items"][0]
                self.assertEqual((count, company, stock_date), (1, "测试公司", "2026-10-29"))
                self.assertEqual(item["qty"], 12.5)
                self.assertEqual(item["wh_raw"], 12.5)
            finally:
                stock.STOCK = previous_stock_path

    def test_batch_tracker_outflows_do_not_become_inventory(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "stock.json"
            stock.save_json(path, {"items": [{
                "name": "直接原料-测试物料",
                "short": "测试物料",
                "batches": [{"batch": "LOT-OLD", "qty": 4, "exp": "2027-01-01"}],
            }]})
            previous_stock_path = stock.STOCK
            stock.STOCK = str(path)
            try:
                headers = ["商品名称", "批次号", "数量", "业务类型", "到期日期"]
                rows = [
                    ["直接原料-测试物料", "LOT-OLD", 5, "入库", "2027-01-01"],
                    ["直接原料-测试物料", "LOT-OLD", 8, "领用", "2027-01-01"],
                    ["直接原料-测试物料", "LOT-NEW", 3, "出库", "2027-02-01"],
                ]
                with patch.object(stock, "open_table", return_value=Table([headers] + rows)):
                    items_hit, batches_hit, misses = stock.import_batch_xls("批次跟踪表.xlsx")
                item = stock.load_json(path)["items"][0]
                quantities = {batch["batch"]: batch["qty"] for batch in item["batches"]}
                self.assertEqual((items_hit, batches_hit, misses), (1, 2, []))
                self.assertEqual(quantities, {"LOT-OLD": 4, "LOT-NEW": 0})
            finally:
                stock.STOCK = previous_stock_path

    def test_batch_import_matches_changed_display_name_by_unique_item_code(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "stock.json"
            stock.save_json(path, {"items": [{
                "name": "直接原料-大豆分离蛋白",
                "short": "大豆分离蛋白",
                "code": "0201000042",
                "batches": [],
            }]})
            previous_stock_path = stock.STOCK
            stock.STOCK = str(path)
            table = Table([
                ["序号", "商品编码", "商品名称", "规格型号", "商品类别", "商品品牌",
                 "批次", "生产日期", "保质期", "保质期单位", "到期日期", "单据日期",
                 "单据编号", "业务类型", "仓库", "单位"],
                [1, "0201000042", "大豆分离蛋白粉", "25kg/袋", "原材料-直接原料",
                 "", "SOY-01", "2026-01-01", 365, "天", "2027-01-01", "2026-10-29",
                 "IN-1", "采购入库", "原料仓", "KG"],
                [2, "0201000099", "库存中不存在的原料", "25kg/袋", "原材料-直接原料",
                 "", "SOY-02", "2026-01-01", 365, "天", "2027-01-01", "2026-10-29",
                 "OUT-1", "组装出库", "原料仓", "KG"],
            ])
            try:
                with patch.object(stock, "open_table", return_value=table):
                    items_hit, batches_hit, misses = stock.import_batch_xls("batch.xls")
                self.assertEqual((items_hit, batches_hit, misses),
                                 (1, 1, ["库存中不存在的原料"]))
                data = stock.load_json(path)
                self.assertEqual(len(data["items"]), 1)
                self.assertEqual(data["items"][0]["batches"][0]["batch"], "SOY-01")
            finally:
                stock.STOCK = previous_stock_path

    def test_fifo_allocation_uses_earliest_receipt_first(self):
        allocations = stock.allocate({"batches": [
            {"batch": "RECEIVED-FIRST", "qty": 4, "in_date": "2026-01-02",
             "exp": "2027-12-01"},
            {"batch": "EXPIRES-FIRST", "qty": 8, "in_date": "2026-02-02",
             "exp": "2027-01-01"},
        ]}, 10)
        self.assertEqual(allocations, [("RECEIVED-FIRST", 4), ("EXPIRES-FIRST", 6)])

    def test_fifo_allocation_subtracts_existing_reservations_from_oldest_lots(self):
        allocations = stock.allocate({"batches": [
            {"batch": "LOT-OLD", "qty": 4, "in_date": "2026-01-01"},
            {"batch": "LOT-NEW", "qty": 6, "in_date": "2026-02-01"},
        ]}, 5, occupied=3)
        self.assertEqual(allocations, [("LOT-OLD", 1), ("LOT-NEW", 4)])

    def test_fifo_allocation_scales_overstated_batch_total_to_stock_book_balance(self):
        allocations = stock.allocate({"batches": [
            {"batch": "LOT-OLD", "qty": 8, "in_date": "2026-01-01"},
            {"batch": "LOT-NEW", "qty": 8, "in_date": "2026-02-01"},
        ]}, 10, book_qty=10)
        self.assertEqual(allocations, [("LOT-OLD", 5), ("LOT-NEW", 5)])

    def test_batch_display_scales_test_era_estimates_to_latest_book_balance(self):
        item = {"qty": 61.98, "batches": [
            {"batch": "20260225", "qty": 22.995, "mfg": "2026-02-25"},
            {"batch": "20260525", "qty": 27.594, "mfg": "2026-05-25"},
            {"batch": "20260910", "qty": 14.891, "mfg": "2026-09-10"},
        ]}
        batches, adjusted = stock.batches_at_book_balance(item)
        self.assertTrue(adjusted)
        self.assertEqual(round(sum(b["qty"] for b in batches), 3), 61.98)
        self.assertIn("按账面比例校准（批次量估算）", stock.batch_summary(item))
        self.assertNotIn("批次合计 65.48", stock.batch_summary(item))

    def test_fifo_allocation_scales_before_subtracting_reservations(self):
        allocations = stock.allocate({"batches": [
            {"batch": "LOT-OLD", "qty": 8, "in_date": "2026-01-01"},
            {"batch": "LOT-NEW", "qty": 8, "in_date": "2026-02-01"},
        ]}, 8, occupied=4, book_qty=10)
        self.assertEqual(allocations[:2], [("LOT-OLD", 1), ("LOT-NEW", 5)])
        self.assertTrue(allocations[-1][0].startswith("缺 "))
        self.assertEqual(sum(quantity for _, quantity in allocations), 6)

    def test_batch_tracker_records_earliest_inbound_document_date(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "stock.json"
            stock.save_json(path, {"items": [{
                "name": "直接原料-测试物料", "short": "测试物料", "code": "A001",
                "batches": [],
            }]})
            previous_stock_path = stock.STOCK
            stock.STOCK = str(path)
            table = Table([
                ["商品编码", "商品名称", "批号", "数量", "业务类型", "单据日期"],
                ["A001", "直接原料-测试物料", "LOT-1", 8, "采购入库", "2026-03-01"],
                ["A001", "直接原料-测试物料", "LOT-1", 2, "采购入库", "2026-02-01"],
                ["A001", "直接原料-测试物料", "LOT-1", 1, "组装出库", "2026-04-01"],
            ])
            try:
                with patch.object(stock, "open_table", return_value=table):
                    stock.import_batch_xls("batch.xls")
                item = stock.load_json(path)["items"][0]
                self.assertEqual(item["batches"][0]["qty"], 9)
                self.assertEqual(item["batches"][0]["in_date"], "2026-02-01")
            finally:
                stock.STOCK = previous_stock_path

    def test_duplicate_order_is_not_appended_twice(self):
        with tempfile.TemporaryDirectory() as directory:
            previous_orders_path = orders.ORDERS
            orders.ORDERS = os.path.join(directory, "orders.json")
            try:
                order = {"serial": "20261028-01", "product": "测试产品"}
                self.assertTrue(orders.add_order(dict(order))[0])
                self.assertFalse(orders.add_order(dict(order))[0])
                with open(orders.ORDERS, encoding="utf-8") as stream:
                    self.assertEqual(len(json.load(stream)["orders"]), 1)
            finally:
                orders.ORDERS = previous_orders_path

    def test_order_exists_finds_matching_serial_and_product(self):
        with tempfile.TemporaryDirectory() as directory:
            previous_orders_path = orders.ORDERS
            orders.ORDERS = os.path.join(directory, "orders.json")
            try:
                orders.save({"orders": [{"serial": "20261028-01", "product": "测试产品"}]})
                self.assertTrue(orders.order_exists("20261028-01", "测试产品"))
                self.assertFalse(orders.order_exists("20261028-02", "测试产品"))
                self.assertFalse(orders.order_exists("20261028-01", "其他产品"))
            finally:
                orders.ORDERS = previous_orders_path

    def test_occupied_map_surfaces_invalid_order_ledger(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "orders.json"
            path.write_text("{invalid", encoding="utf-8")
            with patch.object(stock, "APP", directory):
                with self.assertRaises(json.JSONDecodeError):
                    stock.occupied_map()

    def test_occupied_map_rejects_invalid_pending_material_quantity(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "orders.json"
            path.write_text(json.dumps({"orders": [{
                "status": "待生产",
                "materials": [{"name": "测试物料", "kg": "not-a-number"}],
            }]}), encoding="utf-8")
            with patch.object(stock, "APP", directory):
                with self.assertRaises(ValueError):
                    stock.occupied_map()

    def test_production_step_updates_are_audited(self):
        with tempfile.TemporaryDirectory() as directory:
            previous_orders_path = orders.ORDERS
            orders.ORDERS = os.path.join(directory, "orders.json")
            try:
                orders.save({"orders": [{
                    "serial": "20991028-01", "product": "测试产品", "status": "生产中",
                }]})
                result = orders.update_production_step(
                    "20991028-01", "测试产品", "压片", "进行中", "合格",
                    "操作员甲", "首件已确认")
                self.assertTrue(result["changed"])
                stored = orders.load()["orders"][0]
                self.assertEqual(stored["production_steps"]["压片"]["qc"], "合格")
                self.assertEqual(stored["production_history"][-1]["step"], "压片")
                with self.assertRaises(ValueError):
                    orders.update_production_step(
                        "20991028-01", "测试产品", "压片", qc="擅自定义结果")
            finally:
                orders.ORDERS = previous_orders_path

    def test_test_and_experiment_orders_are_classified_and_reserve_stock(self):
        with tempfile.TemporaryDirectory() as directory:
            previous_orders_path = orders.ORDERS
            orders.ORDERS = os.path.join(directory, "orders.json")
            try:
                order = {
                    "serial": "20991028-01", "product": "测试产品",
                    "status": "待生产", "order_type": "实验单",
                    "materials": [{"name": "测试物料", "kg": 2.5}],
                }
                self.assertTrue(orders.add_order(order)[0])
                with patch.object(stock, "APP", directory):
                    self.assertEqual(stock.occupied_map(), {"测试物料": 2.5})
                shown = orders.production_dashboard()
                self.assertEqual(shown[0]["order_type"], "实验单")
                self.assertFalse(shown[0]["can_delete"])
                self.assertTrue(shown[0]["can_delete_test"])
                with self.assertRaisesRegex(ValueError, "未开工"):
                    orders.delete_unstarted_order("20991028-01", "测试产品")
                orders.save({"orders": [{
                    "serial": "20991028-02", "product": "旧实验单",
                    "status": "已作废", "archive_reason": "实验",
                    "status_history": [{
                        "from": "待生产", "to": orders.VOID, "reason": "实验",
                    }],
                }]})
                legacy = orders.production_dashboard(include_closed=True)[0]
                self.assertEqual(legacy["order_type"], "实验单")
                self.assertTrue(legacy["can_delete_test"])
                deleted = orders.delete_test_order("20991028-02", "旧实验单")
                self.assertTrue(deleted["deleted"])
                self.assertEqual(orders.load()["orders"], [])
            finally:
                orders.ORDERS = previous_orders_path

    def test_unstarted_order_can_be_deleted_but_recorded_order_cannot(self):
        with tempfile.TemporaryDirectory() as directory:
            previous_orders_path = orders.ORDERS
            orders.ORDERS = os.path.join(directory, "orders.json")
            try:
                orders.add_order({"serial": "20991028-01", "product": "未开工"})
                result = orders.delete_unstarted_order("20991028-01", "未开工")
                self.assertTrue(result["deleted"])
                self.assertEqual(orders.load()["orders"], [])

                orders.add_order({"serial": "20991028-02", "product": "已有交接"})
                orders.update_production_step(
                    "20991028-02", "已有交接", "混料", "进行中", "未检",
                    "操作员甲", "开始生产")
                with self.assertRaisesRegex(ValueError, "仅可删除未开工"):
                    orders.delete_unstarted_order("20991028-02", "已有交接")
            finally:
                orders.ORDERS = previous_orders_path

    def test_completion_and_void_are_audited_and_release_reservation(self):
        with tempfile.TemporaryDirectory() as directory:
            previous_orders_path = orders.ORDERS
            orders.ORDERS = os.path.join(directory, "orders.json")
            try:
                orders.add_order({
                    "serial": "20991028-01", "product": "待完成",
                    "materials": [{"name": "测试物料", "kg": 2}],
                })
                with patch.object(stock, "APP", directory):
                    self.assertEqual(stock.occupied_map(), {"测试物料": 2.0})
                orders.update_order_status("20991028-01", "待完成", "已完成")
                completed = orders.load()["orders"][0]
                self.assertEqual(completed["status_history"][-1]["to"], "已完成")
                self.assertEqual(orders.production_dashboard(), [])
                with patch.object(stock, "APP", directory):
                    self.assertEqual(stock.occupied_map(), {})

                with self.assertRaisesRegex(ValueError, "不能重新打开"):
                    orders.update_order_status("20991028-01", "待完成", "待生产")

                orders.add_order({
                    "serial": "20991028-02", "product": "待作废",
                    "materials": [{"name": "测试物料", "kg": 1}],
                    "order_type": "测试单",
                })
                result = orders.archive_or_void_order(
                    "20991028-02", "待作废", "误建测试单")
                self.assertEqual(result["status"], orders.VOID)
                self.assertEqual(orders.production_dashboard(), [])
                all_orders = orders.production_dashboard(include_closed=True)
                self.assertEqual(all_orders[1]["archive_reason"], "误建测试单")
                with patch.object(stock, "APP", directory):
                    self.assertEqual(stock.occupied_map(), {})
            finally:
                orders.ORDERS = previous_orders_path

    def test_completed_order_archive_preserves_completed_status_and_history(self):
        with tempfile.TemporaryDirectory() as directory:
            previous_orders_path = orders.ORDERS
            orders.ORDERS = os.path.join(directory, "orders.json")
            try:
                orders.add_order({"serial": "20991028-01", "product": "已完工"})
                orders.update_order_status("20991028-01", "已完工", "已完成")
                result = orders.archive_or_void_order(
                    "20991028-01", "已完工", "归档已完成记录")
                stored = orders.load()["orders"][0]
                self.assertEqual(result["status"], "已完成")
                self.assertEqual(stored["status"], "已完成")
                self.assertTrue(stored["archived"])
                self.assertEqual(stored["archive_reason"], "归档已完成记录")
                self.assertEqual(orders.production_dashboard(), [])
            finally:
                orders.ORDERS = previous_orders_path

    def test_backup_restore_creates_safety_copy_and_restores_all_business_data(self):
        with tempfile.TemporaryDirectory() as directory:
            previous_app = backup_manager.APP
            previous_backup_dir = backup_manager.BACKUP_DIR
            backup_manager.APP = directory
            backup_manager.BACKUP_DIR = os.path.join(directory, "data", "backups")
            original = {
                "stock.json": {"items": [{"name": "原料", "qty": 7}]},
                "orders.json": {"orders": [{"serial": "old"}]},
                "recipes.json": {"旧配方": {}},
                "mapping.json": {"旧物料": "库存物料"},
            }
            try:
                for filename, value in original.items():
                    (Path(directory) / filename).write_text(
                        json.dumps(value, ensure_ascii=False), encoding="utf-8")
                backup = backup_manager.create_backup()
                for filename in backup_manager.BUSINESS_FILES:
                    (Path(directory) / filename).write_text("{}", encoding="utf-8")
                result = backup_manager.restore_backup(backup["path"])
                self.assertTrue(result["ok"])
                self.assertTrue(os.path.isfile(result["safety_backup"]))
                for filename, value in original.items():
                    self.assertEqual(json.loads((Path(directory) / filename).read_text(
                        encoding="utf-8")), value)
                self.assertEqual(len(backup_manager.list_backups()), 2)
            finally:
                backup_manager.APP = previous_app
                backup_manager.BACKUP_DIR = previous_backup_dir

    def test_restore_rejects_paths_outside_backup_folder(self):
        with tempfile.TemporaryDirectory() as directory:
            previous_backup_dir = backup_manager.BACKUP_DIR
            backup_manager.BACKUP_DIR = os.path.join(directory, "backups")
            outside = os.path.join(directory, "outside.zip")
            Path(outside).write_bytes(b"not a backup")
            try:
                with self.assertRaisesRegex(ValueError, "只能恢复"):
                    backup_manager.restore_backup(outside)
            finally:
                backup_manager.BACKUP_DIR = previous_backup_dir

    def test_backup_listing_marks_corrupt_archives_unrestorable(self):
        with tempfile.TemporaryDirectory() as directory:
            previous_app = backup_manager.APP
            previous_backup_dir = backup_manager.BACKUP_DIR
            backup_manager.APP = directory
            backup_manager.BACKUP_DIR = os.path.join(directory, "data", "backups")
            try:
                for filename, value in {
                    "stock.json": {"items": []},
                    "orders.json": {"orders": []},
                    "recipes.json": {},
                    "mapping.json": {},
                }.items():
                    (Path(directory) / filename).write_text(
                        json.dumps(value), encoding="utf-8")
                backup_manager.create_backup()
                corrupt = Path(backup_manager.BACKUP_DIR) / "OrderSystem_corrupt.zip"
                corrupt.write_bytes(b"not a zip archive")

                backups = {item["name"]: item for item in backup_manager.list_backups()}
                self.assertTrue(backups["OrderSystem_corrupt.zip"]["valid"] is False)
                self.assertIn("BadZipFile", backups["OrderSystem_corrupt.zip"]["validation_error"])
                self.assertTrue(next(item for name, item in backups.items()
                                     if name != "OrderSystem_corrupt.zip")["valid"])
            finally:
                backup_manager.APP = previous_app
                backup_manager.BACKUP_DIR = previous_backup_dir

    def test_package_build_preserves_previous_distribution(self):
        with tempfile.TemporaryDirectory() as directory:
            app = Path(directory) / "app"
            dist = app / "dist"
            root = dist / make_package.NAME
            app.mkdir()
            dist.mkdir()
            root.mkdir()
            (root / "old-release.txt").write_text("preserve", encoding="utf-8")
            for filename in make_package.PY_FILES + make_package.DATA_FILES + make_package.DOC_FILES:
                (app / filename).parent.mkdir(parents=True, exist_ok=True)
                (app / filename).write_text("{}", encoding="utf-8")
            (app / "app_gui.py").write_text(
                "def create_manual_tracking(): pass\n", encoding="utf-8")
            (app / "api_server.py").write_text(
                '"create_manual_tracking": app_gui.create_manual_tracking\n',
                encoding="utf-8")
            (app / "stock.json").write_text(
                json.dumps({"items": [{"name": "live-only", "qty": 99}]}), encoding="utf-8")
            (app / "orders.json").write_text(
                json.dumps({"orders": [{"serial": "live-only"}]}), encoding="utf-8")
            template = app / "templates" / "生产指令与交接单.xlsx"
            template.parent.mkdir(parents=True)
            template.write_bytes(b"fixed master template")
            (app / "webui").mkdir()
            (app / "webui" / "index.html").write_text(
                "<html>手动新增跟踪单</html>", encoding="utf-8")
            previous = (make_package.APP, make_package.DIST, make_package.ROOT)
            try:
                make_package.APP = str(app)
                make_package.DIST = str(dist)
                make_package.ROOT = str(root)
                with patch("make_package.os.path.expanduser",
                           return_value=os.path.join(directory, "no-home")):
                    archive_path = make_package.main()
                backups = list(dist.glob(make_package.NAME + "_backup_*"))
                self.assertEqual(len(backups), 1)
                self.assertEqual((backups[0] / "old-release.txt").read_text(
                    encoding="utf-8"), "preserve")
                self.assertTrue((root / "backup_manager.py").is_file())
                self.assertIn("def create_manual_tracking",
                              (root / "app_gui.py").read_text(encoding="utf-8"))
                self.assertIn('"create_manual_tracking": app_gui.create_manual_tracking',
                              (root / "api_server.py").read_text(encoding="utf-8"))
                self.assertIn("手动新增跟踪单",
                              (root / "webui" / "index.html").read_text(encoding="utf-8"))
                self.assertEqual((root / "templates" / "生产指令与交接单.xlsx").read_bytes(),
                                 b"fixed master template")
                self.assertEqual(json.loads((root / "stock.json").read_text(
                    encoding="utf-8")), {"_meta": {}, "items": []})
                self.assertEqual(json.loads((root / "orders.json").read_text(
                    encoding="utf-8")), {"orders": []})
                with __import__("zipfile").ZipFile(archive_path) as archive:
                    self.assertIn(make_package.NAME + "/backup_manager.py", archive.namelist())
                    self.assertIn(make_package.NAME + "/templates/生产指令与交接单.xlsx",
                                  archive.namelist())
                    self.assertNotIn("live-only", archive.read(make_package.NAME + "/stock.json").decode("utf-8"))
                    self.assertNotIn("live-only", archive.read(make_package.NAME + "/orders.json").decode("utf-8"))
            finally:
                make_package.APP, make_package.DIST, make_package.ROOT = previous


class OrderQuantityModeTests(unittest.TestCase):
    def test_order_preview_accepts_kg_bottle_and_tablet_modes(self):
        import app_gui
        name = next(iter(app_gui.load_recipes()))
        wan = app_gui.order_preview(name, "wan", 1, 0.5, None, 60)
        kg = app_gui.order_preview(name, "kg", 5, 0.5, None, 60)
        tablet = app_gui.order_preview(name, "tablet", 10000, 0.5, None, 60)
        bottle = app_gui.order_preview(name, "bottle", 1000, 0.5, None, 60)
        for result in (wan, kg, tablet, bottle):
            self.assertTrue(result["ok"], result)
        self.assertEqual((kg["total_kg"], tablet["total_kg"]), (5, 5))
        self.assertEqual(wan["total_kg"], 5)
        self.assertEqual((bottle["total_kg"], bottle["bottles"]), (30, 1000))

    def test_order_preview_tolerates_blank_packaging_counts(self):
        import app_gui

        product = "测试空包装规格"
        recipe = {
            "tablet_g": 0.5,
            "packaging": {"per_bottle": "", "per_box": ""},
            "materials": [],
        }
        with patch("app_gui.load_recipes", return_value={product: recipe}), \
                patch("app_gui.build_report", return_value=({}, [], [], [])):
            result = app_gui.order_preview(product, "tablet", 30000, 0.5)
        self.assertTrue(result["ok"], result)
        self.assertEqual(result["total_kg"], 15)
        self.assertEqual(result["bottles"], 500)

    def test_order_preview_shows_fifo_split_and_available_batch_details(self):
        import app_gui

        product = "测试批次拆分"
        recipe = {
            "tablet_g": 0.5,
            "packaging": {"per_bottle": 60, "per_box": 500},
            "materials": [{"name": "测试原料", "ratio": 100}],
        }
        rows = [{
            "short": "测试原料", "name": "直接原料-测试原料",
            "avail": 10, "qty": 10, "unit": "KG", "status": "正常",
            "batches": [
                {"batch": "LOT-2", "qty": 6, "in_date": "2026-02-01"},
                {"batch": "LOT-1", "qty": 4, "in_date": "2026-01-01"},
            ],
        }]
        with patch("app_gui.load_recipes", return_value={product: recipe}), \
                patch("app_gui.build_report", return_value=({}, rows, [], [])):
            result = app_gui.order_preview(
                product, "kg", 10, 0.5,
                [{"name": "测试原料", "ratio": 100, "stage": 0, "batch": ""}], 60)
        self.assertTrue(result["ok"], result)
        material = result["items"][0]
        self.assertEqual(material["alloc_list"], [
            {"batch": "LOT-1", "qty": 4}, {"batch": "LOT-2", "qty": 6},
        ])
        self.assertIn("入库2026-01-01", material["batch_summary"])

    def test_order_preview_does_not_guess_between_ambiguous_inventory_names(self):
        import app_gui

        product = "模糊库存匹配测试"
        recipe = {
            "tablet_g": 0.5,
            "materials": [{"name": "原料甲", "ratio": 100}],
        }
        rows = [
            {"short": "原料甲颗粒", "avail": 20, "qty": 20, "unit": "KG"},
            {"short": "预处理原料甲", "avail": 30, "qty": 30, "unit": "KG"},
        ]
        with patch.object(app_gui, "load_recipes", return_value={product: recipe}), \
                patch.object(app_gui, "build_report", return_value=({}, rows, [], [])):
            result = app_gui.order_preview(product, "kg", 10, 0.5)
        self.assertTrue(result["ok"], result)
        item = result["items"][0]
        self.assertEqual(item["matched"], "")
        self.assertEqual(item["avail"], 0)
        self.assertEqual(item["shortage"], 10)

    def test_inventory_api_exposes_batch_numbers_and_receipt_dates(self):
        import app_gui

        with tempfile.TemporaryDirectory() as directory:
            Path(directory, "stock.json").write_text(
                json.dumps({"items": []}), encoding="utf-8")
            rows = [{
                "short": "测试原料", "name": "直接原料-测试原料",
                "cat": "原材料-直接原料", "qty": 10, "unit": "KG",
                "occupied": 0, "avail": 10, "line": 0, "status": "正常",
                "batches": [
                    {"batch": "LOT-B", "qty": 6, "in_date": "2026-02-01"},
                    {"batch": "LOT-A", "qty": 4, "in_date": "2026-01-01"},
                ],
            }]
            with patch.object(app_gui, "APP", directory), \
                    patch.object(app_gui, "build_report", return_value=({}, rows, [], [])):
                result = app_gui.get_stock()
        self.assertTrue(result["ok"], result)
        self.assertIn("LOT-A(4 入库2026-01-01)", result["raw"][0]["batches"])
        self.assertLess(result["raw"][0]["batches"].index("LOT-A"),
                        result["raw"][0]["batches"].index("LOT-B"))

    def test_stock_autocomplete_returns_batch_details_in_fifo_order(self):
        import app_gui

        rows = [{
            "short": "测试原料", "name": "直接原料-测试原料",
            "cat": "原材料-直接原料", "qty": 10, "unit": "KG",
            "avail": 10, "occupied": 0, "status": "正常",
            "batches": [
                {"batch": "LOT-B", "qty": 6, "in_date": "2026-02-01"},
                {"batch": "LOT-A", "qty": 4, "in_date": "2026-01-01"},
            ],
        }]
        with patch.object(app_gui, "build_report", return_value=({}, rows, [], [])):
            result = app_gui.search_stock("测试原料")
        self.assertTrue(result["ok"], result)
        item = result["items"][0]
        self.assertEqual([x["batch"] for x in item["batch_options"]], ["LOT-A", "LOT-B"])
        self.assertIn("LOT-A(4 入库2026-01-01)", item["batches"])

    def test_self_check_includes_sales_import(self):
        import app_gui
        with warnings.catch_warnings():
            warnings.simplefilter("error", ResourceWarning)
            items = {item["name"]: item for item in app_gui.self_check()["items"]}
        self.assertEqual(items["销售单导入"]["level"], "ok")

    def test_frontend_errors_are_classified_for_self_check(self):
        import app_gui
        summary = app_gui._classify_frontend_errors([
            "2026-10-29 接口无响应/异常: Failed to fetch",
            "2026-10-29 Command plugin:dialog|confirm not allowed by ACL",
            "2026-10-29 页面脚本错误: TypeError",
        ])
        self.assertEqual(summary, {"network": 1, "dialog_permission": 1, "runtime": 1})

    def test_save_product_rejects_incomplete_ratio_without_writing(self):
        import app_gui

        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "recipes.json"
            path.write_text('{"existing": {}}', encoding="utf-8")
            with patch.object(app_gui, "RECIPE_JSON", str(path)):
                result = app_gui.save_product({
                    "name": "新产品",
                    "materials": [
                        {"name": "物料甲", "ratio": 70, "stage": 0},
                        {"name": "物料乙", "ratio": 20, "stage": 1},
                    ],
                })
            self.assertFalse(result["ok"])
            self.assertIn("未保存", result["msg"])
            self.assertEqual(json.loads(path.read_text(encoding="utf-8")), {"existing": {}})


class ManualTrackingTests(unittest.TestCase):
    def test_manual_tracking_calculates_materials_and_reserves_them(self):
        import app_gui

        product = "手动跟踪测试配方"
        recipe = {
            "tablet_g": 0.5,
            "packaging": {"per_bottle": 60, "per_box": 500},
            "materials": [
                {"name": "物料甲", "ratio": 70, "stage": 0},
                {"name": "物料乙", "ratio": 30, "stage": 1},
            ],
        }
        with tempfile.TemporaryDirectory() as directory:
            previous_orders, previous_app = orders.ORDERS, stock.APP
            orders.ORDERS = os.path.join(directory, "orders.json")
            stock.APP = directory
            try:
                with patch.object(app_gui, "load_recipes", return_value={product: recipe}), \
                        patch.object(app_gui, "build_report", return_value=({}, [], [], [])):
                    result = app_gui.create_manual_tracking({
                        "product": product, "total_kg": 10, "date": "2099-10-28",
                        "seq": 1, "order_type": "实验单",
                    })
                self.assertTrue(result["ok"], result)
                self.assertTrue(result["order"]["manual_tracking"])
                self.assertNotIn("file", result["order"])
                self.assertEqual(
                    [material["kg"] for material in result["order"]["materials"]],
                    [7, 3])
                self.assertEqual(stock.occupied_map(), {"物料甲": 7, "物料乙": 3})

                with patch.object(app_gui, "load_recipes", return_value={product: recipe}), \
                        patch.object(app_gui, "build_report", return_value=({}, [], [], [])):
                    duplicate = app_gui.create_manual_tracking({
                        "product": product, "total_kg": 10, "date": "2099-10-28",
                        "seq": 1, "order_type": "实验单",
                    })
                self.assertFalse(duplicate["ok"])
                self.assertIn("已存在", duplicate["msg"])
                self.assertEqual(len(orders.load()["orders"]), 1)
            finally:
                orders.ORDERS, stock.APP = previous_orders, previous_app

    def test_manual_tracking_rejects_bad_quantity_category_and_recipe(self):
        import app_gui

        product = "手动跟踪错误配方"
        recipe = {
            "tablet_g": 0.5,
            "materials": [{"name": "物料甲", "ratio": 90}],
        }
        with patch.object(app_gui, "load_recipes", return_value={product: recipe}):
            bad_quantity = app_gui.create_manual_tracking({
                "product": product, "total_kg": 0, "date": "2099-10-28",
                "seq": 1, "order_type": "实验单",
            })
            excessive_precision = app_gui.create_manual_tracking({
                "product": product, "total_kg": 10.0001, "date": "2099-10-28",
                "seq": 1, "order_type": "实验单",
            })
            fractional_sequence = app_gui.create_manual_tracking({
                "product": product, "total_kg": 10, "date": "2099-10-28",
                "seq": 1.5, "order_type": "实验单",
            })
            bad_category = app_gui.create_manual_tracking({
                "product": product, "total_kg": 10, "date": "2099-10-28",
                "seq": 1, "order_type": "未定义",
            })
            bad_recipe = app_gui.create_manual_tracking({
                "product": product, "total_kg": 10, "date": "2099-10-28",
                "seq": 1, "order_type": "实验单",
            })
        self.assertFalse(bad_quantity["ok"])
        self.assertFalse(excessive_precision["ok"])
        self.assertFalse(fractional_sequence["ok"])
        self.assertFalse(bad_category["ok"])
        self.assertFalse(bad_recipe["ok"])

    def test_manual_tracking_creates_product_and_versions_changed_same_name_recipe(self):
        import app_gui

        product = "虾青素压片糖果（红色包衣）"
        existing_recipe = {
            "category": "压片糖果", "tablet_g": 0.6,
            "stages": [{"name": "一次混装原料"}],
            "materials": [
                {"name": "原料甲", "ratio": 60, "stage": 0},
                {"name": "原料乙", "ratio": 40, "stage": 0},
            ],
            "packaging": {"coat": "红色包衣"},
        }
        new_materials = [
            {"name": "原料甲", "ratio": 65, "stage": 0},
            {"name": "原料乙", "ratio": 35, "stage": 1},
        ]
        with tempfile.TemporaryDirectory() as directory:
            previous_orders, previous_recipe, previous_stock = (
                orders.ORDERS, app_gui.RECIPE_JSON, stock.APP)
            orders.ORDERS = os.path.join(directory, "orders.json")
            app_gui.RECIPE_JSON = os.path.join(directory, "recipes.json")
            stock.APP = directory
            Path(app_gui.RECIPE_JSON).write_text(
                json.dumps({product: existing_recipe}, ensure_ascii=False),
                encoding="utf-8")
            try:
                with patch.object(app_gui, "build_report", return_value=({}, [], [], [])):
                    result = app_gui.create_manual_tracking({
                        "product": product, "recipe_mode": "custom", "tablet_g": 0.6,
                        "materials": new_materials, "coat": "红色包衣",
                        "total_kg": 10, "date": "2099-10-29", "seq": 7,
                        "order_type": "实验单",
                    })
                self.assertTrue(result["ok"], result)
                self.assertTrue(result["created_product"])
                self.assertEqual(result["product_name"], product + "（配方版本2）")
                self.assertEqual(result["order"]["product"], result["product_name"])
                self.assertEqual(result["order"]["recipe_source"], "新建产品配方")
                self.assertEqual([item["kg"] for item in result["order"]["materials"]],
                                 [6.5, 3.5])
                self.assertFalse(os.path.exists(os.path.join(directory, "output.xlsx")))

                with patch.object(app_gui, "build_report", return_value=({}, [], [], [])):
                    repeated = app_gui.create_manual_tracking({
                        "product": product, "recipe_mode": "custom", "tablet_g": 0.6,
                        "materials": new_materials, "coat": "红色包衣",
                        "total_kg": 10, "date": "2099-10-29", "seq": 8,
                        "order_type": "实验单",
                    })
                self.assertTrue(repeated["ok"], repeated)
                self.assertFalse(repeated["created_product"])
                self.assertEqual(repeated["product_name"], product + "（配方版本2）")
                catalog = json.loads(Path(app_gui.RECIPE_JSON).read_text(encoding="utf-8"))
                self.assertEqual(catalog[product], existing_recipe)
                stored_materials = catalog[product + "（配方版本2）"]["materials"]
                self.assertEqual(
                    [(item["name"], item["ratio"], item["stage"])
                     for item in stored_materials],
                    [("原料甲", 65, 0), ("原料乙", 35, 1)])
                self.assertEqual(len(orders.load()["orders"]), 2)
            finally:
                orders.ORDERS, app_gui.RECIPE_JSON, stock.APP = (
                    previous_orders, previous_recipe, previous_stock)

    def test_manual_tracking_rejects_duplicate_or_incomplete_temporary_recipe(self):
        import app_gui

        base = {
            "product": "临时产品", "recipe_mode": "custom", "tablet_g": 0.5,
            "materials": [{"name": "物料甲", "ratio": 50},
                          {"name": "物料甲", "ratio": 50}],
            "total_kg": 10, "date": "2099-10-29", "seq": 1,
        }
        with patch.object(app_gui, "build_report", return_value=({}, [], [], [])):
            duplicate = app_gui.create_manual_tracking(base)
            bad_total = app_gui.create_manual_tracking({
                **base, "materials": [{"name": "物料甲", "ratio": 90}],
            })
            missing_recipe = app_gui.create_manual_tracking({
                **base, "materials": [],
            })
        self.assertFalse(duplicate["ok"])
        self.assertIn("重复", duplicate["msg"])
        self.assertFalse(bad_total["ok"])
        self.assertFalse(missing_recipe["ok"])

    def test_manual_tracking_rolls_back_new_product_when_order_is_rejected(self):
        import app_gui

        with tempfile.TemporaryDirectory() as directory:
            previous_orders, previous_recipe = orders.ORDERS, app_gui.RECIPE_JSON
            orders.ORDERS = os.path.join(directory, "orders.json")
            app_gui.RECIPE_JSON = os.path.join(directory, "recipes.json")
            Path(app_gui.RECIPE_JSON).write_text('{"_说明":"保留"}', encoding="utf-8")
            try:
                with patch.object(app_gui, "build_report", return_value=({}, [], [], [])), \
                        patch("orders.add_order", return_value=(False, "重复订单")):
                    result = app_gui.create_manual_tracking({
                        "product": "订单失败时回滚产品", "recipe_mode": "custom",
                        "tablet_g": 0.5,
                        "materials": [{"name": "物料甲", "ratio": 100, "stage": 0}],
                        "total_kg": 10, "date": "2099-10-29", "seq": 9,
                    })
                self.assertFalse(result["ok"])
                self.assertEqual(
                    json.loads(Path(app_gui.RECIPE_JSON).read_text(encoding="utf-8")),
                    {"_说明": "保留"})
                self.assertFalse(os.path.exists(orders.ORDERS))
            finally:
                orders.ORDERS, app_gui.RECIPE_JSON = previous_orders, previous_recipe

    def test_manual_tracking_reserves_imported_coating_material_separately(self):
        import app_gui

        product = "带包衣用量测试产品"
        materials = [{"name": "素片原料", "ratio": 100, "stage": 0}]
        rows = [
            {"short": "素片原料", "avail": 20, "qty": 20, "unit": "KG",
             "status": "正常", "batches": []},
            {"short": "包衣粉", "avail": 2, "qty": 2, "unit": "KG",
             "status": "正常", "batches": []},
        ]
        with tempfile.TemporaryDirectory() as directory:
            previous_orders, previous_recipe, previous_stock = (
                orders.ORDERS, app_gui.RECIPE_JSON, stock.APP)
            orders.ORDERS = os.path.join(directory, "orders.json")
            app_gui.RECIPE_JSON = os.path.join(directory, "recipes.json")
            stock.APP = directory
            Path(app_gui.RECIPE_JSON).write_text("{}", encoding="utf-8")
            try:
                with patch.object(app_gui, "build_report", return_value=({}, rows, [], [])):
                    result = app_gui.create_manual_tracking({
                        "product": product, "recipe_mode": "custom", "tablet_g": 0.5,
                        "materials": materials, "coat": "红色包衣",
                        "coat_material": "包衣粉", "coat_ratio": 3,
                        "total_kg": 10, "date": "2099-10-29", "seq": 10,
                    })
                self.assertTrue(result["ok"], result)
                self.assertEqual(
                    [(item["name"], item["kg"], item.get("kind"))
                     for item in result["order"]["materials"]],
                    [("素片原料", 10, None), ("包衣粉", 0.3, "coating")])
                self.assertEqual(stock.occupied_map(), {"素片原料": 10, "包衣粉": 0.3})
            finally:
                orders.ORDERS, app_gui.RECIPE_JSON, stock.APP = (
                    previous_orders, previous_recipe, previous_stock)

    def test_manual_tracking_api_is_registered(self):
        import api_server
        import app_gui

        self.assertIs(api_server.FN["create_manual_tracking"],
                      app_gui.create_manual_tracking)


def make_production_sheet_bytes(serial="20991029-03", product="导入测试产品",
                                sheet_title=None):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = sheet_title or generator.SHEET
    sheet["B2"] = "10.29"
    sheet["D2"] = product
    sheet["G2"] = "2万片"
    sheet["B3"] = "SO-IMPORT-01"
    sheet["E3"] = serial
    sheet["H3"] = "01"
    sheet["J3"] = "2099-11-05"
    sheet["L3"] = "测试经办人"
    sheet["C38"] = "规格要求：0.5g/片"
    sheet["B6"] = "物料甲"
    sheet["E6"] = "LOT-OLD"
    sheet["H6"] = 2.5
    sheet["B7"] = "物料甲"
    sheet["E7"] = "LOT-NEW"
    sheet["H7"] = 1.5
    sheet["B14"] = "物料乙"
    sheet["E14"] = "LOT-B"
    sheet["H14"] = 6
    stream = io.BytesIO()
    workbook.save(stream)
    workbook.close()
    return stream.getvalue()


class ProductionSheetImportTests(unittest.TestCase):
    def test_imported_production_sheet_creates_confirmed_tracking_and_reserves_materials(self):
        import app_gui
        import api_server

        with tempfile.TemporaryDirectory() as directory:
            previous_orders, previous_stock = orders.ORDERS, stock.APP
            orders.ORDERS = os.path.join(directory, "orders.json")
            stock.APP = directory
            try:
                result = app_gui.import_production_sheet(
                    "C:\\uploads\\production.xlsx",
                    base64.b64encode(make_production_sheet_bytes()).decode("ascii"),
                    "正式生产")
                self.assertTrue(result["ok"], result)
                self.assertTrue(result["created"])
                self.assertEqual(api_server.FN["import_production_sheet"],
                                 app_gui.import_production_sheet)
                self.assertEqual(api_server.FN["confirm_tracking_order"],
                                 app_gui.confirm_tracking_order)
                saved = orders.load()["orders"][0]
                self.assertEqual(saved["serial"], "20991029-03")
                self.assertEqual(saved["product"], "导入测试产品")
                self.assertEqual(saved["total_kg"], 10)
                self.assertEqual(saved["wan"], 2)
                self.assertEqual(saved["order_no"], "SO-IMPORT-01")
                self.assertEqual(saved["production_sheet_name"], "production.xlsx")
                self.assertEqual(saved["tracking_confirmation_source"], "生产表导入")
                self.assertEqual(
                    [(item["name"], item["kg"], item["batch"])
                     for item in saved["materials"]],
                    [("物料甲", 4, "LOT-OLD；LOT-NEW"), ("物料乙", 6, "LOT-B")])
                self.assertEqual(stock.occupied_map(), {"物料甲": 4, "物料乙": 6})
            finally:
                orders.ORDERS, stock.APP = previous_orders, previous_stock

    def test_import_matches_existing_serial_without_overwriting_order_or_reservation(self):
        import app_gui

        with tempfile.TemporaryDirectory() as directory:
            previous_orders = orders.ORDERS
            orders.ORDERS = os.path.join(directory, "orders.json")
            try:
                original_materials = [{"name": "台账物料", "kg": 3, "batch": "原批次"}]
                orders.add_order({
                    "serial": "20991029-03", "product": "导入测试产品",
                    "status": "生产中", "order_type": "实验单", "total_kg": 7,
                    "materials": original_materials,
                })
                encoded = base64.b64encode(make_production_sheet_bytes()).decode("ascii")
                first = app_gui.import_production_sheet(
                    "production.xlsx", encoded, "正式生产")
                second = app_gui.import_production_sheet(
                    "production.xlsx", encoded, "正式生产")
                self.assertTrue(first["ok"], first)
                self.assertFalse(first["created"])
                self.assertTrue(first["warnings"])
                self.assertTrue(second["ok"], second)
                self.assertFalse(second["changed"])
                self.assertTrue(second["repeated_import"])
                self.assertEqual(len(orders.load()["orders"]), 1)
                saved = orders.load()["orders"][0]
                self.assertEqual(saved["status"], "生产中")
                self.assertEqual(saved["order_type"], "实验单")
                self.assertEqual(saved["total_kg"], 7)
                self.assertEqual(saved["materials"], original_materials)
                self.assertEqual(saved["tracking_confirmation_source"], "生产表导入")
                self.assertEqual(len(saved["production_sheet_imports"]), 1)
                mismatch = app_gui.import_production_sheet(
                    "production.xlsx",
                    base64.b64encode(make_production_sheet_bytes(
                        product="其他产品")).decode("ascii"),
                    "正式生产")
                self.assertFalse(mismatch["ok"])
                self.assertIn("产品名称不同", mismatch["msg"])
                self.assertEqual(len(orders.load()["orders"]), 1)
            finally:
                orders.ORDERS = previous_orders

    def test_import_rejects_noncanonical_workbook_and_missing_category(self):
        import app_gui

        self.assertFalse(app_gui.import_production_sheet(
            "production.xls", "eA==", "正式生产")["ok"])
        self.assertFalse(app_gui.import_production_sheet(
            "production.xlsx", base64.b64encode(b"not an xlsx").decode("ascii"),
            "正式生产")["ok"])
        self.assertFalse(app_gui.import_production_sheet(
            "production.xlsx",
            base64.b64encode(make_production_sheet_bytes(
                sheet_title="Wrong sheet")).decode("ascii"),
            "正式生产")["ok"])
        self.assertFalse(app_gui.import_production_sheet(
            "production.xlsx",
            base64.b64encode(make_production_sheet_bytes()).decode("ascii"),
            "")["ok"])

    def test_import_accepts_numeric_or_fullwidth_serial_cell(self):
        import app_gui

        parsed = app_gui._parse_production_sheet(
            make_production_sheet_bytes(serial=2026102901), "实验单")
        self.assertEqual(parsed["serial"], "20261029-01")
        parsed = app_gui._parse_production_sheet(
            make_production_sheet_bytes(serial=" 20991029－03 "), "实验单")
        self.assertEqual(parsed["serial"], "20991029-03")
        bad = app_gui.import_production_sheet(
            "production.xlsx",
            base64.b64encode(make_production_sheet_bytes(serial="abc")).decode("ascii"),
            "实验单")
        self.assertFalse(bad["ok"])
        self.assertIn("生产单号格式无效", bad["msg"])
        self.assertIn("20261029-01", bad["msg"])

    def test_import_reads_kg_glued_to_bottle_count(self):
        import app_gui

        for text in ("45kg3000瓶", "45kg（3000瓶）", "45 KG / 3000瓶", "45公斤"):
            workbook = load_workbook(io.BytesIO(make_production_sheet_bytes()))
            workbook[generator.SHEET]["G2"] = text
            buffer = io.BytesIO()
            workbook.save(buffer)
            parsed = app_gui._parse_production_sheet(buffer.getvalue(), "实验单")
            self.assertEqual(parsed["total_kg"], 45, text)

    def test_import_rejects_nonexistent_serial_date_in_chinese(self):
        import app_gui

        with self.assertRaises(ValueError) as caught:
            app_gui._parse_production_sheet(
                make_production_sheet_bytes(serial="20261399-01"), "实验单")
        self.assertIn("日期不存在", str(caught.exception))

    def test_tracking_reminder_appears_after_three_days_and_manual_confirmation_clears_it(self):
        with tempfile.TemporaryDirectory() as directory:
            previous_orders = orders.ORDERS
            orders.ORDERS = os.path.join(directory, "orders.json")
            try:
                orders.add_order({
                    "serial": "20991029-01", "product": "超时待核对",
                    "status": "待生产", "order_type": "正式生产",
                })
                data = orders.load()
                now = datetime.datetime.now()
                data["orders"][0]["tracking_created_at"] = (
                    now - datetime.timedelta(days=2, hours=23, minutes=59)).isoformat(
                        timespec="microseconds")
                data["orders"].append({
                    "serial": "20991029-02", "product": "旧台账",
                    "status": "待生产", "order_type": "正式生产",
                    "tracking_created_at": (
                        now - datetime.timedelta(days=4)).isoformat(timespec="microseconds"),
                    "production_steps": {"混料": {"progress": "进行中", "qc": "未检"}},
                    "production_history": [{"step": "混料"}],
                })
                orders.save(data)
                dashboard = orders.production_dashboard()
                stale = next(item for item in dashboard
                             if item["serial"] == "20991029-01")
                legacy = next(item for item in dashboard
                              if item["serial"] == "20991029-02")
                self.assertFalse(stale["tracking_stale"])
                self.assertEqual(stale["tracking_age_days"], 2)
                self.assertTrue(legacy["tracking_stale"])
                self.assertFalse(legacy["can_delete"])
                self.assertEqual(legacy["tracking_age_days"], 4)

                data = orders.load()
                data["orders"][0]["tracking_created_at"] = (
                    datetime.datetime.now() - datetime.timedelta(days=3)).isoformat(
                        timespec="microseconds")
                orders.save(data)
                stale = next(item for item in orders.production_dashboard()
                             if item["serial"] == "20991029-01")
                self.assertTrue(stale["tracking_stale"])
                self.assertEqual(stale["tracking_age_days"], 3)
                result = orders.confirm_tracking_order("20991029-01", "超时待核对")
                self.assertTrue(result["changed"])
                repeated = orders.confirm_tracking_order("20991029-01", "超时待核对")
                self.assertFalse(repeated["changed"])
                confirmed = next(item for item in orders.production_dashboard()
                                 if item["serial"] == "20991029-01")
                self.assertFalse(confirmed["tracking_stale"])
                saved = orders.load()["orders"][0]
                self.assertEqual(saved["tracking_confirmation_source"], "手动确认")
                self.assertEqual(saved["tracking_history"][-1]["source"], "手动确认")
                self.assertEqual(len(saved["tracking_history"]), 1)
            finally:
                orders.ORDERS = previous_orders


class CanonicalTemplateTests(unittest.TestCase):
    def test_template_path_ignores_desktop_and_config_overrides(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "data").mkdir()
            (root / "data" / "config.json").write_text(
                json.dumps({"template": str(root / "old-template.xlsx")}),
                encoding="utf-8")
            with patch.object(generator, "APP", directory):
                self.assertEqual(
                    generator._template_path(),
                    os.path.join(directory, "templates", generator.TPL_NAME))

    def test_order_generation_rejects_packaging_rows_over_template_capacity(self):
        recipe = {
            "tablet_g": 0.5,
            "materials": [{"name": "测试原料", "ratio": 100, "stage": 0}],
            "packaging": {"per_bottle": 30, "per_box": 500, "materials": []},
        }
        with tempfile.TemporaryDirectory() as directory:
            with patch.object(generator, "load_recipes", return_value={"测试产品": recipe}):
                result = generator.generate({
                    "product": "测试产品", "tablet_g": 0.5, "wan": 1,
                    "date": "20991028", "seq": 8,
                    "pack_materials": [
                        {"name": "包材%d" % index, "qty": 1} for index in range(7)
                    ],
                }, out_root=directory)
            self.assertFalse(result["ok"])
            self.assertIn("6 行上限", result["msg"])
            self.assertFalse(list(Path(directory).rglob("*.xlsx")))

    def test_order_generation_uses_master_and_fills_all_four_section_headers(self):
        recipe = {
            "tablet_g": 0.5,
            "materials": [{"name": "测试原料", "ratio": 100, "stage": 0}],
            "stages": [{"name": "一次混装原料", "prefix": "先加", "action": "混合", "time": "5"}],
            "packaging": {"per_bottle": 30, "per_box": 500,
                          "materials": [{"name": "测试瓶", "qty": 1}]},
        }
        with tempfile.TemporaryDirectory() as directory:
            with patch.object(generator, "load_recipes", return_value={"测试产品": recipe}), \
                    patch.object(generator, "stock_rows", return_value=[]):
                result = generator.generate({
                    "product": "测试产品", "tablet_g": 0.5, "total_kg": 10,
                    "date": "20991028", "seq": 7, "order_no": "SO-7",
                    "manager": "销售经理", "maker": "制表人", "batch": "02",
                    "coat": "红色包衣", "shape": "圆形",
                    "coating_steps": {
                        "包衣液配制": "按工艺配制包衣液",
                        "素片预热": "预热至工艺要求",
                        "包衣液雾化": "按工艺参数喷涂",
                        "包衣设备清洗": "按清洗要求处理",
                    },
                }, out_root=directory)
            self.assertTrue(result["ok"], result)
            self.assertEqual(
                Path(result["file"]).name,
                "20991028-07_正式生产_测试产品_10kg.xlsx")
            self.assertEqual(
                Path(result["file"]).parent.name,
                "20991028-07_正式生产_测试产品_10kg")
            workbook = __import__("openpyxl").load_workbook(result["file"], data_only=True)
            sheet = workbook[generator.SHEET]
            for date_cell, product_cell, qty_cell, maker_cell, order_cell, serial_cell, batch_cell, manager_cell in (
                    ("B2", "D2", "G2", "L2", "B3", "E3", "H3", "L3"),
                    ("B33", "D33", "G33", "L33", "B34", "E34", "H34", "L34"),
                    ("B43", "D43", "G43", "L43", "B44", "E44", "H44", "L44"),
                    ("B61", "D61", "G61", "L61", "B62", "E62", "H62", "L62")):
                self.assertEqual(sheet[date_cell].value, "10.28")
                self.assertEqual(sheet[product_cell].value, "测试产品")
                self.assertEqual(sheet[maker_cell].value, "制表人")
                self.assertEqual(sheet[order_cell].value, "SO-7")
                self.assertEqual(sheet[serial_cell].value, "20991028-07")
                self.assertEqual(sheet[batch_cell].value, "02")
                self.assertEqual(sheet[manager_cell].value, "销售经理")
                self.assertTrue(sheet[qty_cell].value)
            self.assertIn("红色包衣", sheet["C38"].value)
            self.assertEqual(sheet["E54"].value, "按工艺配制包衣液")
            self.assertEqual(sheet["E57"].value, "按清洗要求处理")
            self.assertEqual(sheet["D69"].value, "测试瓶")
            self.assertEqual(sheet["G69"].value, "667 套")
            workbook.close()


class OrderGenerationTests(unittest.TestCase):
    def test_generate_splits_ten_kg_across_oldest_batches_in_receipt_order(self):
        recipe = {
            "tablet_g": 0.5,
            "materials": [{"name": "测试原料", "ratio": 100, "stage": 0}],
            "stages": [{"name": "一次混装原料", "prefix": "先加",
                        "action": "混合", "time": "5"}],
            "packaging": {"per_bottle": 60, "per_box": 500, "materials": []},
        }
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            template_path = root / "template.xlsx"
            workbook = Workbook()
            workbook.active.title = generator.SHEET
            workbook.save(template_path)
            workbook.close()
            rows = [{
                "short": "测试原料", "name": "直接原料-测试原料",
                "qty": 10, "avail": 10, "occupied": 0, "unit": "KG",
                "batches": [
                    {"batch": "LOT-NEWER", "qty": 6, "in_date": "2026-02-01"},
                    {"batch": "LOT-OLDER", "qty": 4, "in_date": "2026-01-01"},
                ],
            }]
            with patch.object(generator, "load_recipes", return_value={"测试产品": recipe}), \
                    patch.object(generator, "TEMPLATE", str(template_path)), \
                    patch.object(generator, "stock_rows", return_value=rows):
                result = generator.generate({
                    "product": "测试产品", "tablet_g": 0.5, "total_kg": 10,
                    "materials": [{"name": "测试原料", "ratio": 100,
                                   "stage": 0, "batch": ""}],
                    "date": "20991028", "seq": 7,
                }, out_root=str(root / "output"))
            self.assertTrue(result["ok"], result)
            workbook = load_workbook(result["file"], data_only=True)
            sheet = workbook[generator.SHEET]
            self.assertEqual(
                [(sheet["E6"].value, sheet["H6"].value),
                 (sheet["E7"].value, sheet["H7"].value)],
                [("LOT-OLDER", 4), ("LOT-NEWER", 6)],
            )
            workbook.close()

    def test_save_product_accepts_valid_recipe_and_keeps_stage_indexes(self):
        import app_gui

        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "recipes.json"
            path.write_text('{"existing": {}}', encoding="utf-8")
            with patch.object(app_gui, "RECIPE_JSON", str(path)):
                result = app_gui.save_product({
                    "name": "新产品",
                    "materials": [
                        {"name": "物料甲", "ratio": 60, "stage": 0},
                        {"name": "物料乙", "ratio": 40, "stage": 1},
                    ],
                })
            self.assertTrue(result["ok"], result)
            saved = json.loads(path.read_text(encoding="utf-8"))["新产品"]
            self.assertEqual([m["stage"] for m in saved["materials"]], [0, 1])
            self.assertTrue(all("batch" not in m for m in saved["materials"]))
            self.assertEqual(saved["stages"][1]["name"], "二次混装原料")

    def test_generate_order_checks_ledger_before_creating_a_file(self):
        import app_gui

        with patch("orders.order_exists", return_value=True), \
                patch.object(app_gui, "generate") as generate_mock:
            result = app_gui.generate_order({
                "product": "测试产品", "date": "20991028", "seq": 1,
            })
        self.assertFalse(result["ok"])
        self.assertIn("已存在", result["msg"])
        generate_mock.assert_not_called()

    def test_test_order_delete_api_is_registered(self):
        import api_server
        import app_gui

        self.assertIs(api_server.FN["delete_test_order"], app_gui.delete_test_order)

    def test_generate_order_does_not_create_file_when_ledger_write_fails(self):
        import app_gui

        with patch("orders.order_exists", return_value=False), \
                patch.object(app_gui, "generate") as generate_mock, \
                patch("orders.add_order", return_value=(False, "台账中已有")):
            result = app_gui.generate_order({
                "product": "测试产品", "date": "20991028", "seq": 1,
            })
        self.assertFalse(result["ok"])
        self.assertIn("未生成文件", result["msg"])
        generate_mock.assert_not_called()

    def test_generate_order_rolls_back_ledger_when_document_generation_fails(self):
        import app_gui

        with tempfile.TemporaryDirectory() as directory:
            previous_orders_path = orders.ORDERS
            orders.ORDERS = os.path.join(directory, "orders.json")
            try:
                with patch.object(app_gui, "generate", return_value={
                    "ok": False, "msg": "模板错误",
                }):
                    result = app_gui.generate_order({
                        "product": "测试产品", "date": "20991028", "seq": 1,
                        "order_type": "测试单",
                    })
                self.assertFalse(result["ok"])
                self.assertEqual(result["msg"], "模板错误")
                self.assertEqual(orders.load()["orders"], [])
            finally:
                orders.ORDERS = previous_orders_path

    def test_generate_order_records_selected_category(self):
        import app_gui

        with patch("orders.order_exists", return_value=False), \
                patch.object(app_gui, "generate", return_value={
                    "ok": True, "serial": "20991028-01", "product": "测试产品",
                    "plan_wan": 1, "tablet_g": 0.5, "materials": [], "packaging": [],
                    "file": "result.xlsx", "total_kg": 10,
                }), \
                patch("orders.add_order", return_value=(True, "已加入台账")) as add_order, \
                patch("orders.set_order_file") as set_file:
            result = app_gui.generate_order({
                "product": "测试产品", "date": "20991028", "seq": 1,
                "order_type": "实验单",
            })
        self.assertTrue(result["ok"])
        self.assertEqual(add_order.call_args.args[0]["order_type"], "实验单")
        set_file.assert_called_once()

    def test_generate_order_rejects_unknown_category_before_writing_workbook(self):
        import app_gui

        with patch.object(app_gui, "generate") as generate_mock:
            result = app_gui.generate_order({
                "product": "测试产品", "date": "20991028", "seq": 1,
                "order_type": "其他",
            })
        self.assertFalse(result["ok"])
        generate_mock.assert_not_called()

    def test_generate_rejects_existing_output_without_overwriting(self):
        recipe = {
            "tablet_g": 0.5,
            "materials": [{"name": "物料甲", "ratio": 100, "stage": 0}],
            "stages": [{"name": "一次混装原料", "prefix": "先加", "action": "混合", "time": "5"}],
            "packaging": {"per_bottle": 60, "per_box": 500, "materials": []},
        }
        with tempfile.TemporaryDirectory() as directory:
            out_root = Path(directory)
            folder = out_root / "10月" / "20991028" / "20991028-01_正式生产_测试产品_1万片"
            folder.mkdir(parents=True)
            target = folder / "20991028-01_正式生产_测试产品_1万片.xlsx"
            target.write_bytes(b"existing-user-file")
            workbook_path = out_root / "template.xlsx"
            workbook = Workbook()
            workbook.active.title = generator.SHEET
            workbook.save(workbook_path)
            workbook.close()
            with patch.object(generator, "load_recipes", return_value={"测试产品": recipe}), \
                    patch.object(generator, "TEMPLATE", str(workbook_path)), \
                    patch.object(generator, "stock_rows", return_value=[]):
                result = generator.generate({
                    "product": "测试产品", "tablet_g": 0.5, "wan": 1,
                    "materials": [{"name": "物料甲", "ratio": 100, "stage": 0}],
                    "date": "20991028", "seq": 1,
                }, out_root=str(out_root))
            self.assertFalse(result["ok"])
            self.assertIn("未覆盖", result["msg"])
            self.assertEqual(target.read_bytes(), b"existing-user-file")

    def test_generate_uses_zero_based_stage_and_rejects_overflow(self):
        recipe = {
            "tablet_g": 0.5,
            "materials": [{"name": "物料甲", "ratio": 100, "stage": 0}],
            "stages": [
                {"name": "一次混装原料", "prefix": "先加", "action": "混合", "time": "5"},
                {"name": "二次混装原料", "prefix": "再加", "action": "混合", "time": "5"},
            ],
            "packaging": {"per_bottle": 60, "per_box": 500, "materials": []},
        }
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            workbook_path = root / "template.xlsx"
            workbook = Workbook()
            workbook.active.title = generator.SHEET
            workbook.save(workbook_path)
            workbook.close()
            with patch.object(generator, "load_recipes", return_value={"测试产品": recipe}), \
                    patch.object(generator, "TEMPLATE", str(workbook_path)), \
                    patch.object(generator, "stock_rows", return_value=[]):
                stage_root = root / "stage"
                result = generator.generate({
                    "product": "测试产品", "tablet_g": 0.5, "wan": 1,
                    "materials": [{"name": "物料乙", "ratio": 100, "stage": 1}],
                    "date": "20991028", "seq": 2,
                }, out_root=str(stage_root))
                self.assertTrue(result["ok"], result)
                workbook = __import__("openpyxl").load_workbook(result["file"], data_only=True)
                sheet = workbook[generator.SHEET]
                self.assertEqual(sheet["B14"].value, "物料乙")
                self.assertEqual(sheet["A14"].value, "二次混装原料")
                workbook.close()

                overflow_root = root / "overflow"
                too_many = [{"name": "物料%02d" % i, "ratio": 100 / 9, "stage": 0}
                            for i in range(9)]
                overflow = generator.generate({
                    "product": "测试产品", "tablet_g": 0.5, "wan": 1,
                    "materials": too_many, "date": "20991028", "seq": 3,
                }, out_root=str(overflow_root))
                self.assertFalse(overflow["ok"])
                self.assertEqual(overflow["overflow"], ["物料08"])
                self.assertFalse((overflow_root / "10月" / "20991028" /
                                  "20991028-03_正式生产_测试产品_1万片" /
                                  "20991028-03_正式生产_测试产品_1万片.xlsx").exists())


if __name__ == "__main__":
    unittest.main()
