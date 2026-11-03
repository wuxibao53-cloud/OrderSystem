import json
import tempfile
import unittest
import warnings
from pathlib import Path
from unittest.mock import patch

from docx import Document
from openpyxl import Workbook, load_workbook

import app_gui
import orders
from sales_form_import import (PlanMismatchError, fill_production_template,
                               import_sales_form, parse_sales_form, plan_to_kg)


def make_sales_form(path):
    document = Document()
    table = document.add_table(rows=11, cols=2)
    values = [
        ("产品名称：压片糖果（红色包衣）", "产品名称：压片糖果（红色包衣）"),
        ("原料名称", "配比（%）"),
        ("一水葡萄糖", "54.3"),
        ("麦芽糊精", "35"),
        ("碳酸钙", "10"),
        ("硬脂酸镁", "0.7"),
        ("复配被膜剂（红色包衣）", "3%（按素片重量计）"),
        ("规格：0.5 g/片；30片/瓶", "规格：0.5 g/片；30片/瓶"),
        ("包装和标签：瓶（茶瓶150cc）；盖（银盖）；标签（）；纸盒（）；防伪码（）",) * 2,
        ("压片：\n1.配料：检查原辅料是否结块，如有结块过筛后使用\n"
         "2.混料：先加一水葡萄糖、麦芽糊精、碳酸钙预混10-15 min；"
         "再加硬脂酸镁混合2 min\n3.压片：单片克重控制在0.48-0.52 g/片；"
         "硬度要求用力刮表面或边缘无粉末",) * 2,
        ("包衣：\n1.包衣液配制：称取素片总量3%的包衣粉；"
         "加入88%（以包衣粉重量计）的纯化水充分溶解后待用\n"
         "2.素片预热：将素片倒入包衣锅内，打开加热与进风，"
         "预热几分钟后，直至素片温度稳定\n"
         "3.包衣液雾化：喷枪垂直与片床，相距20-30 cm；雾化压力0.40-0.60 Mpa；"
         "包衣锅转速2-3 r/min\n"
         "4.包衣设备清洗：先用纯化水清洗，再用75%乙醇擦拭消毒，晾干备用",) * 2,
    ]
    for row, pair in zip(table.rows, values):
        for cell, value in zip(row.cells, pair):
            cell.text = value
    document.save(path)


def make_template(path):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "批生产指令与交接单"
    sheet["A1"] = "生产指令与交接单（混料）"
    sheet["A32"] = "生产指令与交接单（压片）"
    sheet["A42"] = "生产指令与交接单（包衣）"
    sheet["A60"] = "生产指令与交接单（包装）"
    sheet["A6"] = "一次\n混装\n原料"
    for row, label in zip(range(54, 58), ("包衣液配制", "素片预热", "包衣液雾化", "包衣设备清洗")):
        sheet.cell(row, 3).value = label
    sheet["C63"] = "关键信息"
    sheet["E25"] = "设备外部： □ 干净 □ 不干净"
    sheet["G69"] = "_____套"
    for start, end in ((6, 13), (14, 16), (17, 19)):
        sheet.merge_cells(start_row=start, start_column=7, end_row=end, end_column=7)
    sheet.merge_cells("C38:G38")
    sheet.merge_cells("C39:G39")
    for row in range(54, 58):
        sheet.merge_cells(start_row=row, start_column=5, end_row=row, end_column=10)
    sheet.merge_cells("C63:L63")
    sheet.merge_cells("D75:J75")
    workbook.save(path)
    workbook.close()


class SalesFormImportTests(unittest.TestCase):
    def setUp(self):
        ledger = Path(tempfile.mkdtemp(prefix="ledger-")) / "orders.json"
        ledger.write_text('{"orders": []}', encoding="utf-8")
        patcher = patch.object(orders, "ORDERS", str(ledger))
        patcher.start()
        self.addCleanup(patcher.stop)
        self.addCleanup(lambda: ledger.unlink(missing_ok=True))

    def test_manual_recipe_upload_parses_product_formula_without_creating_order(self):
        import base64
        import api_server

        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "虾青素配方单-20261028.docx"
            make_sales_form(source)
            result = app_gui.preview_manual_recipe_upload(
                source.name, base64.b64encode(source.read_bytes()).decode("ascii"))
        self.assertTrue(result["ok"], result)
        recipe = result["recipe"]
        self.assertEqual(recipe["product"], "压片糖果（红色包衣）")
        self.assertEqual(recipe["tablet_g"], 0.5)
        self.assertEqual(recipe["coat"], "红色包衣")
        self.assertEqual(recipe["coat_material"], "复配被膜剂（红色包衣）")
        self.assertEqual(recipe["coat_ratio"], 3.0)
        self.assertEqual(
            [(material["name"], material["stage"]) for material in recipe["materials"]],
            [("一水葡萄糖", 0), ("麦芽糊精", 0), ("碳酸钙", 0), ("硬脂酸镁", 1)])
        self.assertIs(api_server.FN["preview_manual_recipe_upload"],
                      app_gui.preview_manual_recipe_upload)
        self.assertEqual(result["warnings"], [])

    def test_manual_recipe_upload_rejects_bad_file_input(self):
        self.assertFalse(app_gui.preview_manual_recipe_upload("recipe.pdf", "eA==")["ok"])
        self.assertFalse(app_gui.preview_manual_recipe_upload("recipe.docx", "bad")["ok"])

    def test_imported_recipe_creates_versioned_product_and_tracking_order(self):
        import base64
        import json
        import os
        import orders
        import stock

        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "虾青素配方单-20991029.docx"
            make_sales_form(source)
            parsed = app_gui.preview_manual_recipe_upload(
                source.name, base64.b64encode(source.read_bytes()).decode("ascii"))
            self.assertTrue(parsed["ok"], parsed)
            recipe = parsed["recipe"]
            previous_orders = (orders.ORDERS, app_gui.RECIPE_JSON, stock.APP)
            orders.ORDERS = os.path.join(directory, "orders.json")
            app_gui.RECIPE_JSON = os.path.join(directory, "recipes.json")
            stock.APP = directory
            Path(app_gui.RECIPE_JSON).write_text("{}", encoding="utf-8")
            try:
                with patch.object(app_gui, "build_report", return_value=({}, [], [], [])):
                    result = app_gui.create_manual_tracking({
                        "product": recipe["product"], "recipe_mode": "custom",
                        "tablet_g": recipe["tablet_g"], "materials": recipe["materials"],
                        "coat": recipe["coat"], "coat_material": recipe["coat_material"],
                        "coat_ratio": recipe["coat_ratio"],
                        "per_bottle": recipe["per_bottle"],
                        "total_kg": 10, "date": "2099-10-29", "seq": 1,
                    })
                self.assertTrue(result["ok"], result)
                self.assertTrue(result["created_product"])
                self.assertEqual(result["product_name"], recipe["product"])
                ledger_order = orders.load()["orders"][0]
                self.assertEqual(ledger_order["product"], recipe["product"])
                coat = next(item for item in ledger_order["materials"]
                            if item.get("kind") == "coating")
                self.assertEqual((coat["name"], coat["kg"]), ("复配被膜剂（红色包衣）", 0.3))
                catalog = json.loads(Path(app_gui.RECIPE_JSON).read_text(encoding="utf-8"))
                self.assertEqual(
                    catalog[recipe["product"]]["packaging"]["coat_material"],
                    "复配被膜剂（红色包衣）")
            finally:
                orders.ORDERS, app_gui.RECIPE_JSON, stock.APP = previous_orders

    def test_browser_upload_validates_docx_and_passes_original_filename(self):
        import base64

        payload = b"docx upload payload"
        received = {}
        def capture_upload(source, template, history_root=None, options=None,
                           stock_rows=None):
            received["name"] = Path(source).name
            received["content"] = Path(source).read_bytes()
            received["template"] = template
            received["options"] = options
            received["stock_rows"] = stock_rows
            return {"output": "C:/orders/result.xlsx", "serial": 2026102801,
                    "plan_kg": 10, "warnings": []}

        with patch("app_gui.build_report", return_value=({}, [], [], [])), \
                patch("sales_form_import.import_sales_form") as importer:
            importer.side_effect = capture_upload
            result = app_gui.import_sales_form_upload(
                "虾青素配方单-20261028.docx", base64.b64encode(payload).decode("ascii"), "")
        self.assertTrue(result["ok"])
        self.assertEqual(result["folder"], "C:/orders")
        self.assertEqual(received["name"], "虾青素配方单-20261028.docx")
        self.assertEqual(received["content"], payload)
        self.assertEqual(received["options"], {"serial": ""})
        self.assertEqual(received["stock_rows"], [])
        self.assertEqual(Path(received["template"]).name, "生产指令与交接单.xlsx")
        self.assertEqual(Path(received["template"]).parent.name, "templates")

    def test_browser_import_accepts_tablets_when_sales_form_omits_bottles(self):
        import base64

        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "维生素配方单-20261028.docx"
            make_sales_form(source)
            document = Document(source)
            document.tables[0].rows[7].cells[0].text = "规格：0.5 g/片；每瓶数量未提供"
            document.tables[0].rows[7].cells[1].text = "规格：0.5 g/片；每瓶数量未提供"
            document.save(source)
            with patch("sales_form_import._default_history_root",
                       return_value=Path(directory) / "history"):
                result = app_gui.import_sales_form_upload(
                    source.name,
                    base64.b64encode(source.read_bytes()).decode("ascii"),
                    "30000",
                    "tablet",
                )

        self.assertTrue(result["ok"], result)
        self.assertEqual(result["plan_kg"], 15)
        self.assertIsNone(result["bottles"])

    def test_incomplete_process_and_missing_coating_still_generate_with_warnings(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "维生素配方单-20261028.docx"
            template = Path(directory) / "template.xlsx"
            output = Path(directory) / "out.xlsx"
            make_sales_form(source)
            document = Document(source)
            table = document.tables[0]
            table.rows[9].cells[0].text = (
                "配料顺序：先将一水葡萄糖、麦芽糊精、碳酸钙预混5min；"
                "再加硬脂酸镁混合10-15min"
            )
            table.rows[6].cells[0].text = "包衣剂名称未提供"
            table.rows[6].cells[1].text = ""
            table.rows[10].cells[0].text = ""
            table.rows[10].cells[1].text = ""
            document.save(source)
            make_template(template)

            form = parse_sales_form(source)
            self.assertIn("先将", form["press_process"])
            self.assertIsNone(form["coat_ratio"])
            order = {
                "date": "20261028", "serial": 2026102801,
                "product": "维生素耐力片", "plan_kg": 15,
                "plan_text": "15kg", "maker": "", "manager": "",
                "batch": "01", "material_batches": {},
            }
            with warnings.catch_warnings():
                warnings.simplefilter("error", DeprecationWarning)
                result = fill_production_template(form, template, output, order)
            sheet = load_workbook(output, data_only=True).active
            self.assertEqual(sheet["B6"].value, "一水葡萄糖")
            self.assertEqual(sheet["B14"].value, "硬脂酸镁")
            self.assertEqual(sheet["E26"].value, "5 min")
            self.assertEqual(sheet["E27"].value, "10-15 min")
            self.assertIsNone(sheet["D49"].value)
            self.assertIsNone(sheet["G49"].value)
            self.assertIsNone(sheet["E54"].value)
            self.assertTrue(any("包衣剂" in warning for warning in result["warnings"]))
            self.assertTrue(output.is_file())

    def test_browser_upload_rejects_non_docx_and_invalid_base64(self):
        self.assertFalse(app_gui.import_sales_form_upload(
            "sales.xlsx", "eA==")["ok"])
        self.assertFalse(app_gui.import_sales_form_upload(
            "sales.docx", "not-base64")["ok"])

    def test_parse_sales_form_extracts_formula_specs_and_coating(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "虾青素配方单-20261028.docx"
            make_sales_form(source)
            form = parse_sales_form(source)
            self.assertEqual(form["source_date"], "20261028")
            self.assertEqual(form["tablet_g"], 0.5)
            self.assertEqual(form["per_bottle"], 30)
            self.assertEqual(form["ratio_total"], 100)
            self.assertEqual([m["ratio"] for m in form["materials"]],
                             [54.3, 35, 10, 0.7])
            self.assertEqual(form["coat_ratio"], 3)
            self.assertEqual(form["coat_type"], "红色包衣")
            self.assertIn("检查原辅料是否结块", form["prep_instruction"])
            self.assertIn("充分溶解后待用", form["coating_steps"]["包衣液配制"])
            self.assertIn("75%乙醇", form["coating_steps"]["包衣设备清洗"])
            self.assertEqual([p["kind"] for p in form["packaging"]],
                             ["瓶", "盖", "标签", "纸盒", "防伪码"])

    def test_fill_four_section_template_calculates_only_supported_quantities(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "虾青素配方单-20261028.docx"
            template = Path(directory) / "template.xlsx"
            output = Path(directory) / "out.xlsx"
            make_sales_form(source)
            make_template(template)
            form = parse_sales_form(source)
            order = {
                "date": "20261028",
                "serial": 2026102801,
                "product": "INZENT虾青素压片糖果",
                "plan_kg": 10,
                "plan_text": "10kg",
                "maker": "制表人",
                "manager": "实验",
                "batch": "01",
                "material_batches": {"一水葡萄糖": "B1", "麦芽糊精": "B2",
                                     "碳酸钙": "B3", "硬脂酸镁": "B4"},
            }
            result = fill_production_template(form, template, output, order)
            workbook = load_workbook(output, data_only=True)
            sheet = workbook.active
            self.assertEqual(sheet["D2"].value, "INZENT虾青素压片糖果")
            self.assertEqual(sheet["E3"].value, 2026102801)
            self.assertEqual(sheet["G2"].value, "10kg（667瓶）")
            self.assertEqual(sheet["G33"].value, "10kg（667瓶）")
            self.assertEqual(sheet["G43"].value, "10kg（667瓶）")
            self.assertEqual(sheet["G61"].value, "10kg（667瓶）")
            self.assertEqual(sheet["D61"].value, "INZENT虾青素压片糖果")
            self.assertEqual(sheet["E62"].value, 2026102801)
            self.assertEqual(sheet["L62"].value, "实验")
            self.assertEqual(sheet["B6"].value, "一水葡萄糖")
            self.assertEqual(sheet["H6"].value, 5.43)
            self.assertEqual(sheet["J6"].value, 5430)
            self.assertEqual(sheet["E14"].value, "B4")
            self.assertEqual(sheet["J14"].value, 70)
            self.assertAlmostEqual(sheet["G49"].value, 0.3)
            self.assertAlmostEqual(sheet["G50"].value, 0.264)
            self.assertIn("红色包衣", sheet["C38"].value)
            self.assertIn("检查原辅料是否结块", sheet["C23"].value)
            self.assertEqual(sheet["D69"].value, "瓶（茶瓶150cc）")
            self.assertEqual(sheet["G69"].value, "_____套")
            self.assertEqual(sheet["D71"].value, "标签（）")
            self.assertIsNone(sheet["G72"].value)
            self.assertEqual(sheet["C63"].value, "关键信息")
            self.assertEqual(sheet["A6"].value, "一次\n混装\n原料")
            self.assertEqual(sheet["E25"].value, "设备外部： □ 干净 □ 不干净")
            self.assertEqual(sheet["E26"].value, "10-15 min")
            self.assertEqual(sheet["E27"].value, "2 min")
            self.assertIsNone(sheet["E28"].value)
            self.assertIn("0.48-0.52g/片", sheet["C39"].value)
            self.assertIn("用力刮表面或边缘无粉末", sheet["C39"].value)
            self.assertNotIn("见销售单", sheet["C39"].value)
            self.assertNotIn(" 2.", sheet["E54"].value)
            self.assertNotIn(" 3.", sheet["E55"].value)
            self.assertNotIn(" 4.", sheet["E56"].value)
            workbook.close()
            self.assertEqual(result["bottles"], 667)
            self.assertTrue(output.is_file())

    def test_sales_import_assigns_and_splits_fifo_batches(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "虾青素配方单-20261028.docx"
            template = Path(directory) / "template.xlsx"
            output = Path(directory) / "out.xlsx"
            make_sales_form(source)
            make_template(template)
            form = parse_sales_form(source)
            form["materials"] = [{"name": "一水葡萄糖", "ratio": 100}]
            order = {
                "date": "20261028", "serial": 2026102801,
                "product": "测试产品", "plan_kg": 10, "plan_text": "10kg",
                "maker": "", "manager": "", "batch": "01",
                "material_batches": {},
            }
            rows = [{
                "short": "一水葡萄糖", "batches": [
                    {"batch": "LOT-NEW", "qty": 6, "in_date": "2026-02-01"},
                    {"batch": "LOT-OLD", "qty": 4, "in_date": "2026-01-01"},
                ],
            }]
            result = fill_production_template(
                form, template, output, order, stock_rows=rows)
            workbook = load_workbook(output, data_only=True)
            sheet = workbook.active
            self.assertEqual(
                [(sheet["E6"].value, sheet["H6"].value),
                 (sheet["E7"].value, sheet["H7"].value)],
                [("LOT-OLD", 4), ("LOT-NEW", 6)],
            )
            self.assertFalse(any("批次可用数量不足" in note for note in result["warnings"]))
            workbook.close()

    def test_batch_cell_keeps_codes_only_and_names_follow_stock_ledger(self):
        from sales_form_import import _canonical_material_name

        stock_reference = [{"short": "Vc", "name": "维生素C"},
                           {"short": "一水葡萄糖", "name": "直接原料-一水葡萄糖"}]
        self.assertEqual(_canonical_material_name("vc", stock_reference), "Vc")
        self.assertEqual(_canonical_material_name("一水 葡萄糖", stock_reference),
                         "一水葡萄糖")
        self.assertEqual(_canonical_material_name("麦芽 糊精", stock_reference),
                         "麦芽糊精")

        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "虾青素配方单-20261028.docx"
            template = Path(directory) / "template.xlsx"
            output = Path(directory) / "out.xlsx"
            make_sales_form(source)
            make_template(template)
            form = parse_sales_form(source)
            form["materials"] = [{"name": "一水 葡萄糖", "ratio": 60},
                                 {"name": "麦芽 糊精", "ratio": 40}]
            order = {
                "date": "20261028", "serial": 2026102801,
                "product": "测试产品", "plan_kg": 10, "plan_text": "10kg",
                "maker": "", "manager": "", "batch": "01",
                "material_batches": {"一水葡萄糖": "LOT-A；LOT-B"},
            }
            rows = [{
                "short": "一水葡萄糖", "batches": [
                    {"batch": "LOT-A", "qty": 3, "in_date": "2026-02-01"},
                    {"batch": "LOT-B", "qty": 2, "in_date": "2026-01-01"},
                ],
            }, {
                "short": "麦芽糊精", "batches": [
                    {"batch": "LOT-C", "qty": 5, "in_date": "2026-01-05"},
                ],
            }]
            result = fill_production_template(
                form, template, output, order, stock_rows=rows)
            workbook = load_workbook(output, data_only=True)
            sheet = workbook.active
            # 名称按库存档案写法统一（去空格、用档案大小写）
            self.assertEqual(sheet["B6"].value, "一水葡萄糖")
            self.assertEqual(sheet["B7"].value, "麦芽糊精")
            # 批次栏只写批号编码；「取多少」只体现在 H（kg）/ J（g）列
            self.assertEqual(sheet["E6"].value, "LOT-A；LOT-B")
            self.assertNotIn("取", sheet["E6"].value)
            self.assertNotIn("不足", sheet["E6"].value)
            self.assertEqual(sheet["H6"].value, 6)
            self.assertEqual(sheet["J6"].value, 6000)
            self.assertEqual(sheet["E7"].value, "LOT-C")
            self.assertEqual(sheet["J7"].value, 4000)
            self.assertFalse(any("批次可用数量不足" in note
                                 for note in result["warnings"]))
            workbook.close()

    def test_sales_import_defaults_maker_and_header_shows_weight_and_bottles(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory) / "10月"
            source = Path(directory) / "虾青素配方单-20261028.docx"
            template = Path(directory) / "template.xlsx"
            make_sales_form(source)
            make_template(template)
            result = import_sales_form(source, template, root, {"plan_kg": 10})
            self.assertTrue(Path(result["output"]).is_file())
            sheet = load_workbook(result["output"], data_only=True).active
            self.assertEqual(sheet["L2"].value, "李浩冉")
            self.assertEqual(sheet["G2"].value, "10kg（667瓶）")
            self.assertEqual(result["plan_display"], "10kg（667瓶）")

    def test_plan_mismatch_on_auto_match_stops_without_new_serial(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory) / "10月"
            source = Path(directory) / "虾青素配方单-20261028.docx"
            template = Path(directory) / "template.xlsx"
            make_sales_form(source)
            make_template(template)
            day = root / "20261028" / "20261028-01_既有产品"
            day.mkdir(parents=True)
            form = parse_sales_form(source)
            order = {
                "date": "20261028", "serial": 2026102801,
                "product": "INZENT虾青素压片糖果", "plan_kg": 15,
                "plan_text": "15kg", "maker": "", "manager": "",
                "batch": "01", "material_batches": {},
            }
            fill_production_template(form, template, day / "20261028-01_既有产品.xlsx",
                                     order)
            with self.assertRaises(PlanMismatchError):
                import_sales_form(source, template, root, {"plan_kg": 10})
            self.assertFalse((root / "20261028" / "20261028-02_压片糖果（红色包衣）").exists())

    def test_explicit_serial_attaches_to_pending_ledger_batch(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory) / "10月"
            source = Path(directory) / "虾青素配方单-20261028.docx"
            template = Path(directory) / "template.xlsx"
            make_sales_form(source)
            make_template(template)
            ledger_path = Path(directory) / "orders.json"
            ledger_path.write_text(json.dumps({"orders": [{
                "serial": "20261028-01", "product": "INZENT虾青素压片糖果",
                "total_kg": 45, "status": "待生产", "order_type": "正式生产",
                "batch": "01", "sales": "马士龙",
                "materials": [{"name": "一水葡萄糖", "kg": 24.4, "batch": "OLD"}],
                "production_sheet_imports": [
                    {"at": "2026-10-30T16:24:42", "file": "旧表.xlsx", "sha256": ""}],
            }]}, ensure_ascii=False), encoding="utf-8")
            batch_folder = root / "20261028" / "20261028-01_旧文件夹名"
            batch_folder.mkdir(parents=True)
            stock_rows = [{
                "short": "一水葡萄糖", "batches": [
                    {"batch": "20260917", "qty": 6, "in_date": "2026-09-17"}],
            }, {
                "short": "麦芽糊精", "batches": [
                    {"batch": "20260225", "qty": 5, "in_date": "2026-02-25"}],
            }, {
                "short": "碳酸钙", "batches": [
                    {"batch": "20260302", "qty": 2, "in_date": "2026-03-02"}],
            }, {
                "short": "硬脂酸镁", "batches": [
                    {"batch": "20260402", "qty": 1, "in_date": "2026-04-02"}],
            }]
            with patch.object(orders, "ORDERS", str(ledger_path)):
                result = import_sales_form(
                    source, template, root,
                    {"plan_qty": 10, "plan_unit": "kg", "serial": "20261028-01",
                     "run_date": "20261101"},
                    stock_rows=stock_rows)
            self.assertTrue(result.get("attached"), result)
            self.assertEqual(result.get("run_no"), 2)
            self.assertTrue(Path(result["output"]).name.endswith("_第02批.xlsx"))
            sheet = load_workbook(result["output"], data_only=True).active
            self.assertEqual(sheet["E3"].value, 2026102801)
            self.assertEqual(sheet["H3"].value, "02")
            self.assertEqual(sheet["G2"].value, "10kg（667瓶）")
            # 第 2 批起：生产日期栏显示实际生产日，表尾注明批次
            self.assertEqual(sheet["B2"].value, "11.01")
            self.assertIn("第 2 批", sheet["D75"].value or "")
            self.assertIn("实际生产 11.01", sheet["D75"].value or "")
            self.assertEqual(Path(result["output"]).parent, batch_folder)
            ledger = json.loads(ledger_path.read_text(encoding="utf-8"))
            updated = ledger["orders"][0]
            self.assertEqual(updated["total_kg"], 55)
            self.assertEqual(updated["batch"], "02")
            self.assertEqual(updated["status"], "待生产")
            self.assertEqual(len(updated["production_runs"]), 2)
            self.assertEqual(updated["production_runs"][0]["note"], "历史导入补记")
            self.assertEqual(updated["production_runs"][1]["plan_kg"], 10)
            # 第 1 批还没核销：预占 = 两批物料合计（旧葡萄糖 + 新配方）
            pending_names = sorted(m["name"] for m in updated["materials"])
            self.assertEqual(pending_names.count("一水葡萄糖"), 2)
            self.assertEqual(updated["tracking_history"][-1]["source"], "销售单导入")
            self.assertEqual(updated["tracking_history"][-1]["run_no"], 2)
            # 车间做完第 1 批 → 核销后预占只剩第 2 批新配方
            with patch.object(orders, "ORDERS", str(ledger_path)):
                orders.mark_production_run("20261028-01", 1, True)
            ledger = json.loads(ledger_path.read_text(encoding="utf-8"))
            updated = ledger["orders"][0]
            self.assertEqual(updated["status"], "生产中")
            self.assertEqual([m["name"] for m in updated["materials"]],
                             ["一水葡萄糖", "麦芽糊精", "碳酸钙", "硬脂酸镁"])
            self.assertEqual([m["kg"] for m in updated["materials"]],
                             [5.43, 3.5, 1.0, 0.07])
            self.assertTrue(all(m["batch"] for m in updated["materials"]))
            self.assertEqual(updated["tracking_history"][-1]["source"], "批次核销")
            self.assertEqual(updated["tracking_history"][-1]["run_no"], 1)

    def test_explicit_serial_first_run_on_manual_order(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory) / "10月"
            source = Path(directory) / "虾青素配方单-20261028.docx"
            template = Path(directory) / "template.xlsx"
            make_sales_form(source)
            make_template(template)
            ledger_path = Path(directory) / "orders.json"
            ledger_path.write_text(json.dumps({"orders": [{
                "serial": "20261028-01", "product": "INZENT虾青素压片糖果",
                "status": "待生产", "order_type": "正式生产", "batch": "01",
                "materials": [],
            }]}, ensure_ascii=False), encoding="utf-8")
            with patch.object(orders, "ORDERS", str(ledger_path)):
                result = import_sales_form(
                    source, template, root,
                    {"plan_qty": 10, "plan_unit": "kg", "serial": "20261028-01"})
            self.assertTrue(result.get("attached"), result)
            self.assertEqual(result.get("run_no"), 1)
            self.assertTrue(Path(result["output"]).name.endswith("_销售单导入_10kg.xlsx"))
            sheet = load_workbook(result["output"], data_only=True).active
            self.assertEqual(sheet["H3"].value, "01")
            ledger = json.loads(ledger_path.read_text(encoding="utf-8"))
            updated = ledger["orders"][0]
            self.assertEqual(updated["total_kg"], 10)
            self.assertEqual(len(updated["production_runs"]), 1)

    def test_explicit_serial_refuses_to_duplicate_existing_number(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory) / "10月"
            source = Path(directory) / "虾青素配方单-20261028.docx"
            template = Path(directory) / "template.xlsx"
            make_sales_form(source)
            make_template(template)
            day = root / "20261028" / "20261028-01_既有产品"
            day.mkdir(parents=True)
            stub = Workbook()
            stub.active["E3"] = 2026102801
            stub.save(day / "20261028-01_既有产品.xlsx")
            with self.assertRaises(ValueError):
                import_sales_form(source, template, root,
                                  {"plan_kg": 10, "serial": "20261028-01"})

    def test_next_sequence_reserves_ledger_serials(self):
        from sales_form_import import _next_sequence
        with tempfile.TemporaryDirectory() as directory:
            ledger_path = Path(directory) / "orders.json"
            ledger_path.write_text(json.dumps({"orders": [
                {"serial": "20261028-01", "status": "待生产"},
                {"serial": "20261028-05", "status": "已作废"},
            ]}, ensure_ascii=False), encoding="utf-8")
            with patch.object(orders, "ORDERS", str(ledger_path)):
                self.assertEqual(_next_sequence(Path(directory), "20261028"), 6)

    def test_existing_output_is_never_overwritten(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "虾青素配方单-20261028.docx"
            template = Path(directory) / "template.xlsx"
            output = Path(directory) / "out.xlsx"
            make_sales_form(source)
            make_template(template)
            output.write_bytes(b"existing")
            with self.assertRaises(FileExistsError):
                fill_production_template(
                    parse_sales_form(source), template, output, {
                        "date": "20261028", "serial": 2026102801,
                        "product": "产品", "plan_kg": 10, "plan_text": "10kg",
                        "maker": "", "manager": "", "batch": "01",
                        "material_batches": {},
                    })
            self.assertEqual(output.read_bytes(), b"existing")

    def test_plan_quantity_units_convert_to_kg(self):
        form = {"tablet_g": 0.5, "per_bottle": 30}
        self.assertEqual(plan_to_kg(form, 10, "kg"), 10)
        self.assertEqual(plan_to_kg(form, 20000, "tablet"), 10)
        self.assertEqual(plan_to_kg(form, 667, "bottle"), 10.005)
        with self.assertRaises(ValueError):
            plan_to_kg({"tablet_g": 0.5, "per_bottle": None}, 10, "bottle")
        with self.assertRaises(ValueError):
            plan_to_kg(form, 0, "kg")

    def test_missing_weight_range_and_hardness_stay_blank(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "虾青素配方单-20261028.docx"
            template = Path(directory) / "template.xlsx"
            output = Path(directory) / "out.xlsx"
            make_sales_form(source)
            make_template(template)
            form = parse_sales_form(source)
            form["press_process"] = "压片：先加一水葡萄糖、麦芽糊精、碳酸钙混合10 min；再加硬脂酸镁混合2 min"
            result = fill_production_template(form, template, output, {
                "date": "20261028", "serial": 2026102801, "product": "产品",
                "plan_kg": 0, "plan_text": "", "maker": "", "manager": "",
                "batch": "01", "material_batches": {}})
            sheet = load_workbook(output).active
            self.assertNotIn("见销售单", sheet["C39"].value)
            self.assertIn("________g/片", sheet["C39"].value)
            self.assertIsNone(sheet["H6"].value)
            self.assertIsNone(sheet["G49"].value)
            self.assertIsNone(sheet["G2"].value)
            self.assertTrue(any("片重" in w for w in result["warnings"]))

    def test_new_order_gets_next_daily_sequence_and_plan_is_optional(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory) / "10月"
            day_dir = root / "20261028"
            old_order = day_dir / "20261028-02其他产品"
            old_order.mkdir(parents=True)
            old_book = Workbook()
            old_book.active["E3"] = 2026102802
            old_book.save(old_order / "existing.xlsx")
            old_book.close()
            source = Path(directory) / "虾青素配方单-20261028.docx"
            template = Path(directory) / "template.xlsx"
            make_sales_form(source)
            make_template(template)
            blank = import_sales_form(source, template, root)
            self.assertEqual(blank["serial"], 2026102803)
            self.assertEqual(blank["plan_kg"], 0)
            result = import_sales_form(source, template, root, {"plan_kg": 10})
            self.assertEqual(result["serial"], 2026102804)
            self.assertTrue(Path(result["output"]).is_file())
            self.assertEqual(
                Path(result["output"]).name,
                "20261028-04_压片糖果（红色包衣）_销售单导入_10kg.xlsx")


if __name__ == "__main__":
    unittest.main()
