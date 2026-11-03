# -*- coding: utf-8 -*-
"""Import a Word sales/formula sheet into the production handover workbook."""
import datetime
import hashlib
import os
import re
import tempfile
from copy import copy
from pathlib import Path

from docx import Document
from openpyxl import load_workbook

COATING_STEP_LABELS = ("包衣液配制", "素片预热", "包衣液雾化", "包衣设备清洗")


class PlanMismatchError(Exception):
    """自动匹配到既有生产记录，但计划产量与本次填写不一致（不允许静默改序号）。"""


def _text(cell):
    return " ".join((cell.text or "").split()).strip()


def _safe_path_segment(value):
    value = str(value or "").strip()
    value = "".join("_" if ord(char) < 32 or char in '<>:"/\\|?*' else char
                    for char in value)
    return value.rstrip(" .")[:100] or "未命名"


def _unique_row_text(row):
    result = []
    for cell in row.cells:
        value = _text(cell)
        if value and (not result or result[-1] != value):
            result.append(value)
    return result


def _first_match(pattern, text, cast=str):
    match = re.search(pattern, text, re.IGNORECASE)
    return cast(match.group(1)) if match else None


def _coating_steps(text):
    step_pattern = re.compile(
        r"(?:\d+\s*[.．、]\s*)?(" + "|".join(map(re.escape, COATING_STEP_LABELS)) +
        r")\s*[:：]?"
    )
    matches = list(step_pattern.finditer(text))
    sections = {label: "" for label in COATING_STEP_LABELS}
    for index, match in enumerate(matches):
        end = matches[index + 1].start() if index + 1 < len(matches) else len(text)
        value = text[match.end():end].strip(" \t\r\n /；;。")
        if value:
            sections[match.group(1)] = value
    return sections


def _prep_instruction(text):
    match = re.search(
        r"(?:\d+\s*[.．、]\s*)?配料\s*[:：]\s*(.*?)"
        r"(?=(?:\d+\s*[.．、]\s*)?(?:混料|压片)\s*[:：]|$)",
        text,
    )
    return match.group(1).strip(" \t\r\n /；;。") if match else ""


def parse_sales_form(path):
    """Extract product, BOM, specifications, packaging and process from a DOCX."""
    path = Path(path)
    document = Document(str(path))
    rows = []
    for table in document.tables:
        rows.extend(_unique_row_text(row) for row in table.rows)
    rows = [row for row in rows if row]
    all_text = "\n".join(" ".join(row) for row in rows)

    product = ""
    materials = []
    coat_name = ""
    coat_ratio = None
    for row in rows:
        first = row[0]
        if "产品名称" in first and not product:
            product = first.split("：", 1)[-1].split(":", 1)[-1].strip()
            continue
        if any("原料名称" in cell for cell in row):
            continue
        if len(row) < 2:
            continue
        ratio_match = re.fullmatch(r"\s*(\d+(?:\.\d+)?)\s*%?\s*", row[1])
        if ratio_match and first not in ("配比（%）", "配比(%)"):
            materials.append({"name": first, "ratio": float(ratio_match.group(1))})
        elif any(token in first for token in ("被膜剂", "包衣粉")):
            coat_name = first
            coat_ratio = _first_match(r"(\d+(?:\.\d+)?)\s*%", row[1], float)

    if not product:
        raise ValueError("销售单中未找到「产品名称」")
    if not materials:
        raise ValueError("销售单中未找到带数字配比的原料行")
    ratio_total = sum(item["ratio"] for item in materials)
    if abs(ratio_total - 100.0) > 0.05:
        raise ValueError("基础配方比例合计为 %.3f%%，不是 100%%；已停止导入" % ratio_total)

    spec_text = next((text for text in all_text.splitlines() if "规格" in text), "")
    tablet_g = _first_match(r"(\d+(?:\.\d+)?)\s*g\s*/\s*片", spec_text, float)
    per_bottle = _first_match(r"(\d+)\s*片\s*/\s*瓶", spec_text, int)
    if tablet_g is None:
        raise ValueError("销售单中未识别片重（例如 0.5 g/片）")

    packaging_text = next((text for text in all_text.splitlines()
                           if "包装和标签" in text or "包装及标签" in text), "")
    packaging = []
    for label, detail in re.findall(r"(瓶|盖|标签|纸盒|防伪码)\s*[（(]([^）)]*)[）)]",
                                    packaging_text):
        packaging.append({"name": "%s（%s）" % (label, detail.strip()),
                          "kind": label})

    press_text = ""
    coat_text = ""
    for row in rows:
        if row[0].startswith("压片：") and not press_text:
            press_text = " ".join(row)
        elif row[0].startswith("包衣：") and not coat_text:
            coat_text = " ".join(row)
        elif any(label in row[0] for label in ("工艺顺序", "配料顺序")) and not press_text:
            inline_process = re.split(
                r"(?:工艺顺序|配料顺序)\s*[:：]?", row[0], maxsplit=1)[-1].strip()
            press_text = " ".join([inline_process] + row[1:]).strip()
    coat_type = next((name for name in
                      ("红色包衣", "胃溶包衣", "肠溶包衣", "薄膜包衣", "糖衣")
                      if name in all_text), "")
    coating_steps = _coating_steps(coat_text)
    date_match = re.search(r"(20\d{6})", path.stem)
    order_date = date_match.group(1) if date_match else ""
    sales = ""
    order_no = ""
    due = ""
    for row in rows:
        line = " ".join(row)
        if not sales:
            sales = _first_match(r"(?:下单销售|销售经理|业务员|销售员)\s*[:：]\s*"
                                 r"([\u4e00-\u9fa5A-Za-z·]{2,4})(?=\s|$|｜|｜|，|,|；)",
                                 line) or ""
        if not order_no:
            order_no = _first_match(r"订单号\s*[:：]\s*([A-Za-z0-9\-_]+)", line) or ""
        if not due:
            due_match = re.search(r"交货日期\s*[:：]\s*(20\d{2}[.年/-]?\d{1,2}[.月/-]?\d{1,2})",
                                  line)
            if due_match:
                due = re.sub(r"\D", "", due_match.group(1))[:8]

    return {
        "source": str(path),
        "source_date": order_date,
        "order_no": order_no,
        "due": due,
        "sales": sales,
        "product": product,
        "materials": materials,
        "ratio_total": round(ratio_total, 3),
        "coat_name": coat_name,
        "coat_ratio": coat_ratio,
        "coat_type": coat_type,
        "tablet_g": tablet_g,
        "per_bottle": per_bottle,
        "packaging": packaging,
        "press_process": press_text,
        "prep_instruction": _prep_instruction(press_text),
        "coat_process": coat_text,
        "coating_steps": coating_steps,
    }


def _clean_product(value):
    return re.sub(r"\s+", "", str(value or "")).strip()


def _parse_kg(value):
    if isinstance(value, (int, float)):
        return float(value)
    match = re.search(r"(\d+(?:\.\d+)?)\s*kg\b", str(value), re.IGNORECASE)
    return float(match.group(1)) if match else None


def _clean_material_key(name):
    text = re.sub(r"[\s\u3000]+", "", str(name or ""))
    return "".join(chr(ord(char) - 0xFEE0) if "！" <= char <= "～" else char
                   for char in text)


def _canonical_material_name(name, stock_rows=None):
    """统一物料名称空格/全半角/大小写；能唯一精确匹配库存档案时用档案写法。"""
    text = _clean_material_key(name)
    if stock_rows and text:
        for row in stock_rows:
            for key in ("short", "name"):
                candidate = _clean_material_key(row.get(key))
                if candidate and candidate.lower() == text.lower():
                    return str(row.get(key))
    return text


def _grams(kg):
    """计划称重（g）栏：kg → g，保留 1 位小数。"""
    return round(kg * 1000, 1) if kg else None


def _find_existing_order(history_root, order_date, source, requested_serial=None):
    if not order_date:
        raise ValueError("销售单文件名没有 YYYYMMDD 日期，无法自动归档")
    day_dir = Path(history_root) / order_date
    if not day_dir.is_dir():
        raise FileNotFoundError("未找到日期归档目录：%s" % day_dir)
    source_stem = Path(source).stem
    source_stem = re.sub(
        r"^[0-9a-f]{8}-(?:[0-9a-f]{4}-){3}[0-9a-f]{12}-",
        "",
        source_stem,
        flags=re.IGNORECASE,
    )
    source_keyword = re.sub(r"配方单.*$", "", source_stem)
    candidates = []
    for path in day_dir.rglob("*.xlsx"):
        if path.name.startswith("~$") or "销售单导入" in path.stem:
            continue
        try:
            book = load_workbook(path, read_only=True, data_only=True)
            sheet = book.active
            serial = sheet["E3"].value
            title = sheet["D2"].value
            if serial and str(int(serial)).startswith(order_date) and \
                    source_keyword and source_keyword in _clean_product(title):
                plan = _parse_kg(sheet["G2"].value)
                if plan is not None:
                    candidates.append({
                        "path": path,
                        "serial": int(serial),
                        "product": _clean_product(title),
                        "plan_kg": plan,
                        "plan_text": str(sheet["G2"].value),
                        "maker": sheet["L2"].value or "",
                        "manager": sheet["L3"].value or "",
                        "batch": str(sheet["H3"].value or "01"),
                        "material_batches": {
                            _clean_product(sheet.cell(row, 2).value):
                            str(sheet.cell(row, 5).value or "")
                            for row in list(range(6, 20))
                            if sheet.cell(row, 2).value
                        },
                    })
            book.close()
        except (OSError, ValueError, TypeError):
            continue
    if requested_serial is not None:
        candidates = [candidate for candidate in candidates
                      if candidate["serial"] == int(requested_serial)]
    if len(candidates) != 1:
        raise ValueError("日期 %s 下按产品关键词找到 %d 条既有生产记录；需先明确订单数量和生产序号"
                         % (order_date, len(candidates)))
    return candidates[0]


def _stage_for(material_name, press_process):
    stages = []
    marker_pattern = re.compile(r"最后(?:加|加入|将)|再(?:加|将)|先(?:加|将)")
    markers = list(marker_pattern.finditer(press_process))
    for index, marker in enumerate(markers):
        end = markers[index + 1].start() if index + 1 < len(markers) else len(press_process)
        if material_name in press_process[marker.end():end]:
            prefix = marker.group(0)
            stage = 2 if prefix.startswith("最后") else (1 if prefix.startswith("再") else 0)
            return stage
    raise ValueError("无法从工艺说明确定「%s」的混料阶段" % material_name)


def _stage_text(text, stage):
    marker_pattern = re.compile(r"最后(?:加|加入|将)|再(?:加|将)|先(?:加|将)")
    markers = list(marker_pattern.finditer(text))
    for index, marker in enumerate(markers):
        prefix = marker.group(0)
        current = 2 if prefix.startswith("最后") else (1 if prefix.startswith("再") else 0)
        if current == stage:
            end = markers[index + 1].start() if index + 1 < len(markers) else len(text)
            return text[marker.start():end].strip(" \t；;。")
    return ""


def _set_all_headers(sheet, context, due=""):
    order_date = datetime.datetime.strptime(context["date"], "%Y%m%d")
    date_text = "%d.%02d" % (order_date.month, order_date.day)
    due_text = ""
    if due:
        due_text = "%d.%02d" % (int(due[4:6]), int(due[6:8])) \
            if len(due) == 8 and due.isdigit() else due
    for order_cell, date_cell, product_cell, qty_cell, maker_cell, serial_cell, batch_cell, due_cell, manager_cell in (
            ("B3", "B2", "D2", "G2", "L2", "E3", "H3", "J3", "L3"),
            ("B34", "B33", "D33", "G33", "L33", "E34", "H34", "J34", "L34"),
            ("B44", "B43", "D43", "G43", "L43", "E44", "H44", "J44", "L44"),
            ("B62", "B61", "D61", "G61", "L61", "E62", "H62", "J62", "L62")):
        sheet[order_cell] = context.get("order_no") or None
        sheet[date_cell] = date_text
        sheet[product_cell] = context["product"]
        sheet[qty_cell] = context.get("plan_display") or context.get("plan_text") or None
        sheet[maker_cell] = context["maker"]
        sheet[serial_cell] = context["serial"]
        sheet[batch_cell] = context["batch"]
        sheet[due_cell] = due_text or None
        sheet[manager_cell] = context["manager"]


def _mix_minutes(text, marker):
    stage = {"先加": 0, "再加": 1, "最后加": 2}[marker]
    section = _stage_text(text, stage)
    if not section:
        return None
    match = re.search(r"(\d+(?:\s*[-–]\s*\d+)?)\s*min", section, re.IGNORECASE)
    return "%s min" % re.sub(r"\s+", "", match.group(1)) if match else None


def _missing_notes(order, plan_kg, tablet_range, hardness, form, unassigned,
                   batch_warnings=None):
    notes = []
    if not plan_kg:
        notes.append("计划产量未填写：计划产量及各物料用量留空")
    if not order.get("order_no"):
        notes.append("订单号未提供，留空")
    if not order.get("due"):
        notes.append("交货日期未提供，留空")
    if not tablet_range:
        notes.append("销售单未写片重范围，留空待手填")
    if not hardness:
        notes.append("销售单未写硬度，留空待手填")
    if unassigned:
        notes.append("未识别到投料阶段，原料已列出但投料顺序待人工确认")
    if not form["press_process"]:
        notes.append("销售单未提供压片/混料工艺，工艺栏留空")
    if not form["coat_name"] or form["coat_ratio"] is None:
        notes.append("销售单未提供完整包衣剂名称/比例，包衣用量留空")
    missing_coating_steps = [label for label, value in form["coating_steps"].items() if not value]
    if missing_coating_steps:
        notes.append("销售单未提供完整包衣工艺步骤：%s；对应位置留空"
                     % "、".join(missing_coating_steps))
    notes.append("每箱瓶数及包材数量销售单未提供，留空")
    notes.extend(batch_warnings or [])
    return notes


def fill_production_template(form, template_path, output_path, order, stock_rows=None):
    """Fill the provided four-section production template without altering its source."""
    if os.path.exists(output_path):
        raise FileExistsError("目标文件已存在，为避免覆盖而停止：%s（如需重新生成，请先把旧文件改名或移走）" % output_path)
    plan_kg = order["plan_kg"] or 0
    if plan_kg < 0:
        raise ValueError("计划投料量不能为负数")
    if not form["materials"]:
        raise ValueError("销售单没有可识别的基础原料配比，无法计算物料用量")
    if form["coat_ratio"] is not None and form["coat_ratio"] < 0:
        raise ValueError("包衣剂比例不能为负数")

    materials_by_stage = {0: [], 1: [], 2: []}
    unassigned = False
    batch_warnings = []
    mix_text, press_text_only = _split_press_sections(form["press_process"])
    for material in form["materials"]:
        try:
            material["stage"] = _stage_for(material["name"], mix_text)
        except ValueError:
            material["stage"] = 0
            unassigned = True
        material["qty"] = round(plan_kg * material["ratio"] / 100, 3) if plan_kg else None
        material["display_name"] = _canonical_material_name(
            material["name"], stock_rows)
        material["batch"] = (order["material_batches"].get(material["display_name"])
                             or order["material_batches"].get(material["name"]) or "")
        material["batch_plan"] = []
        material["batch_shortage"] = 0.0
        if material["qty"] and stock_rows is not None and not material["batch"]:
            row = (_lookup_stock_row(stock_rows, material["display_name"])
                   or _lookup_stock_row(stock_rows, material["name"]))
            if row and row.get("batches"):
                from stock import allocate
                plan = allocate({"batches": row["batches"]}, material["qty"],
                                occupied=row.get("occupied", 0),
                                book_qty=row.get("qty"))
                material["batch_plan"] = [
                    {"batch": batch, "qty": amount}
                    for batch, amount in plan if batch and amount > 0
                ]
                assigned = sum(item["qty"] for item in material["batch_plan"])
                if assigned + 0.0005 < material["qty"]:
                    material["batch_shortage"] = round(material["qty"] - assigned, 3)
                    batch_warnings.append(
                        "%s 批次可用数量不足：需求 %g kg，仅能分配 %g kg"
                        % (material["name"], material["qty"], assigned))
            elif row:
                batch_warnings.append("%s 库存未维护批号，生产指令的批号栏留空"
                                      % material["name"])
            else:
                batch_warnings.append("%s 未匹配到库存档案，批号栏留空"
                                      % material["name"])
        materials_by_stage[material["stage"]].append(material)

    book = load_workbook(template_path)
    try:
        if not book.sheetnames:
            raise ValueError("目标模板没有工作表")
        sheet = book[book.sheetnames[0]]
        if (sheet.max_row < 75 or
                not str(sheet["A42"].value or "").startswith("生产指令与交接单") or
                not str(sheet["A60"].value or "").startswith("生产指令与交接单") or
                tuple(str(sheet["C%d" % row].value or "") for row in range(54, 58)) !=
                COATING_STEP_LABELS):
            raise ValueError("目标模板缺少新版包衣/包装区或四项包衣工艺栏")

        header_order = dict(order)
        header_order.setdefault(
            "plan_display", _plan_display(plan_kg, form) if plan_kg else "")
        try:
            run_no = int(order.get("run_no") or 0)
        except (TypeError, ValueError):
            run_no = 0
        run_date = str(order.get("run_date") or "")
        if run_no > 1 and re.fullmatch(r"\d{8}", run_date):
            # 同一批次的第 2 批起：生产日期栏显示实际生产日
            header_order["date"] = run_date
        _set_all_headers(sheet, header_order, order.get("due", ""))
        stage_rows = {0: (6, 13), 1: (14, 16), 2: (17, 19)}
        stage_notes = {}
        for stage, items in materials_by_stage.items():
            if not items:
                continue
            first, last = stage_rows[stage]
            capacity = last - first + 1
            split_count = sum(max(1, len(item["batch_plan"])) for item in items)
            split_batches = split_count <= capacity
            if len(items) > capacity:
                raise ValueError("第 %d 混料阶段物料超过模板行数上限" % (stage + 1))
            row = first
            for material in items:
                allocations = material["batch_plan"]
                name = material["display_name"]
                if (allocations and split_batches and len(allocations) > 1
                        and not material["batch_shortage"]):
                    for allocation in allocations:
                        # 每批一行：E 列只写批号编码；用量只体现在 H（kg）/ J（g）
                        sheet.cell(row, 2).value = name
                        sheet.cell(row, 5).value = allocation["batch"]
                        sheet.cell(row, 8).value = allocation["qty"]
                        sheet.cell(row, 10).value = _grams(allocation["qty"])
                        row += 1
                else:
                    batch = material["batch"]
                    if len(allocations) == 1:
                        batch = allocations[0]["batch"]
                    elif len(allocations) > 1:
                        batch = "；".join(a["batch"] for a in allocations)
                    if material["batch_shortage"]:
                        batch += "%s批次可用不足 %g kg" % (
                            "；" if batch else "", material["batch_shortage"])
                    sheet.cell(row, 2).value = name
                    sheet.cell(row, 5).value = batch or None
                    sheet.cell(row, 8).value = material["qty"]
                    sheet.cell(row, 10).value = _grams(material["qty"])
                    row += 1
            if stage == 0:
                stage_notes[stage] = _stage_text(mix_text, stage) or (
                    "投料阶段未提供，顺序待确认" if unassigned else
                    "先加%s" % "、".join(material["name"] for material in items))
            elif stage == 1:
                stage_notes[stage] = _stage_text(mix_text, stage) or (
                    "再加%s" % "、".join(material["name"] for material in items))
            else:
                stage_notes[stage] = _stage_text(mix_text, stage) or (
                    "最后加入%s" % "、".join(material["name"] for material in items))
        sheet["G6"] = stage_notes.get(0)
        sheet["G14"] = stage_notes.get(1)
        sheet["G17"] = stage_notes.get(2)
        if form["prep_instruction"]:
            sheet["C23"] = "配料要求：" + form["prep_instruction"]
        for stage, cell, marker in ((0, "E26", "先加"), (1, "E27", "再加"),
                                    (2, "E28", "最后加")):
            sheet[cell] = _mix_minutes(mix_text, marker) \
                if materials_by_stage[stage] else None

        sheet["C38"] = "规格要求：%sg/片\n包衣要求：%s     形状：  " % (
            form["tablet_g"], form["coat_type"] or "________")
        tablet_range = _first_match(
            r"单片克重控制在\s*(\d+(?:\.\d+)?\s*[-–~～]\s*\d+(?:\.\d+)?)\s*g",
            press_text_only or mix_text)
        hardness = _first_match(r"硬度要求\s*[:：]?\s*([^；;。\n]+)",
                                press_text_only or mix_text)
        press_note = ""
        if press_text_only:
            press_note = "\n压片要求：" + re.sub(r"\s+", " ", press_text_only)[:120]
        sheet["C39"] = " 片重范围：%s\u3000\u3000\n\n硬度：%s\u3000 厚度 ______mm%s" % (
            ("%sg/片" % re.sub(r"\s+", "", tablet_range)) if tablet_range else "________g/片",
            hardness.strip() if hardness else "_______ N",
            press_note)

        coat_qty = (round(plan_kg * form["coat_ratio"] / 100, 3)
                    if plan_kg and form["coat_ratio"] is not None else None)
        water_qty = round(coat_qty * 0.88, 3) if coat_qty else None
        sheet["D49"] = form["coat_name"] or None
        sheet["G49"] = coat_qty
        sheet["D50"] = "纯化水（包衣粉重量的88%）" if coat_qty is not None else None
        sheet["G50"] = water_qty

        for row, label in zip((54, 55, 56, 57), COATING_STEP_LABELS):
            sheet.cell(row, 5).value = form["coating_steps"][label] or None
            alignment = copy(sheet.cell(row, 5).alignment)
            alignment.wrap_text = True
            sheet.cell(row, 5).alignment = alignment

        if plan_kg and form["per_bottle"]:
            tablets = round(plan_kg * 1000 / form["tablet_g"])
            bottle_count = (tablets + form["per_bottle"] - 1) // form["per_bottle"]
        else:
            bottle_count = None
        for index, package in enumerate(form["packaging"][:6]):
            sheet.cell(69 + index, 4).value = package["name"]

        bottle_text = "%d" % form["per_bottle"] if form["per_bottle"] else "  "
        sheet["D75"] = "重量：%sg/ 片\u3000  每瓶 %s 片\u3000每箱   瓶\u3000" % (
            form["tablet_g"], bottle_text)
        if run_no > 1:
            note = "　第 %d 批" % run_no
            if run_date and re.fullmatch(r"\d{8}", run_date):
                note += "（实际生产 %d.%02d）" % (
                    int(run_date[4:6]), int(run_date[6:8]))
            sheet["D75"] = "%s%s" % (sheet["D75"].value or "", note)

        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        fd, temporary = tempfile.mkstemp(
            prefix=".%s." % output_path.stem, suffix=".xlsx", dir=str(output_path.parent))
        os.close(fd)
        try:
            book.save(temporary)
            os.replace(temporary, output_path)
        except BaseException:
            try:
                os.unlink(temporary)
            except FileNotFoundError:
                pass
            raise
        return {
            "output": str(output_path),
            "serial": order["serial"],
            "product": order["product"],
            "plan_kg": order["plan_kg"],
            "plan_display": order.get("plan_display") or ("%gkg" % plan_kg if plan_kg else ""),
            "materials": materials_by_stage,
            "coating_powder_kg": coat_qty,
            "coating_water_kg": water_qty,
            "bottles": bottle_count,
            "warnings": _missing_notes(
                order, plan_kg, tablet_range, hardness, form, unassigned,
                batch_warnings),
        }
    finally:
        book.close()


def _next_sequence(history_root, order_date):
    day_dir = Path(history_root) / order_date
    used = _ledger_serial_sequences(order_date)
    if day_dir.is_dir():
        for path in day_dir.rglob("*.xlsx"):
            if path.name.startswith("~$"):
                continue
            try:
                book = load_workbook(path, read_only=True, data_only=True)
                serial = book.active["E3"].value
                if serial and str(int(serial)).startswith(order_date):
                    used.add(int(serial) % 100)
                book.close()
            except (OSError, ValueError, TypeError):
                continue
    sequence = max(used, default=0) + 1
    if sequence > 99:
        raise ValueError("日期 %s 的生产序号已达到 99，无法自动分配" % order_date)
    return sequence


def _serial_exists(history_root, order_date, sequence):
    if sequence in _ledger_serial_sequences(order_date):
        return True
    day_dir = Path(history_root) / order_date
    expected = int(order_date + "%02d" % sequence)
    if not day_dir.is_dir():
        return False
    for path in day_dir.rglob("*.xlsx"):
        if path.name.startswith("~$"):
            continue
        try:
            book = load_workbook(path, read_only=True, data_only=True)
            serial = book.active["E3"].value
            book.close()
            if serial and int(serial) == expected:
                return True
        except (OSError, ValueError, TypeError):
            continue
    return False


def _parse_explicit_serial(value, default_date):
    """生产编号支持 20261030-01 / 2026103001 / 01 三种写法 → (日期, 序号)。"""
    text = str(value or "").strip()
    text = "".join("-" if char in "－—–" else char for char in text)
    if not text:
        return None
    match = re.fullmatch(r"(\d{8})-(\d{1,2})", text)
    if match:
        return match.group(1), int(match.group(2))
    if re.fullmatch(r"\d{10}", text):
        return text[:8], int(text[8:])
    if re.fullmatch(r"\d{1,2}", text):
        if not default_date:
            raise ValueError("销售单文件名缺少日期，请填写完整生产编号（如 20261030-01）")
        return default_date, int(text)
    raise ValueError("生产编号格式无效：%s（示例 20261030-01）" % value)


def _ledger_serial_sequences(order_date):
    """台账中已占用的序号（编号与桌面归档共用一个序号空间）。"""
    try:
        from orders import load as load_orders
        data = load_orders()
    except Exception:
        return set()
    used = set()
    for existing in data.get("orders", []):
        serial = str(existing.get("serial") or "")
        if serial.startswith(order_date + "-"):
            tail = serial[len(order_date) + 1:]
            if tail.isdigit():
                used.add(int(tail))
    return used


def _ledger_pending_order(serial_text):
    try:
        from orders import load as load_orders
        data = load_orders()
    except Exception:
        return None
    for existing in data.get("orders", []):
        if str(existing.get("serial") or "") == serial_text:
            return existing
    return None


def _existing_batch_folder(history_root, order_date, sequence):
    day_dir = Path(history_root) / order_date
    prefix = "%s-%02d_" % (order_date, sequence)
    if day_dir.is_dir():
        for path in sorted(day_dir.iterdir()):
            if path.is_dir() and path.name.startswith(prefix):
                return path
    return None


def _default_history_root(order_date):
    desktop = Path.home() / "Desktop"
    month_folder = "%d月" % int(order_date[4:6])
    return desktop / month_folder


def plan_to_kg(form, quantity, unit):
    """Convert a plan quantity in kg / bottles / tablets to kg using the sales-form specs."""
    quantity = float(quantity)
    if quantity <= 0:
        raise ValueError("计划数量必须大于 0")
    if unit == "kg":
        return round(quantity, 3)
    if unit == "tablet":
        return round(quantity * form["tablet_g"] / 1000, 3)
    if unit == "bottle":
        if not form["per_bottle"]:
            raise ValueError("销售单未写每瓶片数，无法按瓶数换算；请改用 kg 或片数")
        return round(quantity * form["per_bottle"] * form["tablet_g"] / 1000, 3)
    raise ValueError("未知计划单位：%s" % unit)


def _lookup_stock_row(rows, name):
    exact = [row for row in rows if row.get("short") == name or row.get("name") == name]
    if exact:
        return exact[0]
    lowered = name.lower()
    matches = [row for row in rows
               if lowered in (row.get("short") or "").lower()
               or (row.get("short") or "").lower() in lowered]
    return matches[0] if len(matches) == 1 else None


def _plan_display(plan_kg, form):
    """计划产量栏：总重 + 瓶数；每瓶片数未知时不编造瓶数。"""
    text = "%gkg" % plan_kg
    per_bottle = form.get("per_bottle")
    tablet_g = form.get("tablet_g")
    if per_bottle and tablet_g:
        tablets = round(plan_kg * 1000 / tablet_g)
        bottles = (tablets + per_bottle - 1) // per_bottle
        if bottles > 0:
            text += "（%d瓶）" % bottles
    return text


def import_sales_form(sales_path, template_path, history_root=None, options=None,
                      stock_rows=None):
    options = dict(options or {})
    override = options.get("form_override")
    if override:
        # AI 提取（或人工确认过的）结构化配方：跳过 Word 解析，仍走同一校验器
        form = form_from_ai(
            override.get("form") if isinstance(override, dict) and override.get("form")
            else override,
            source_name=str(sales_path or "AI 提取单据"))
        if not form["source_date"] and options.get("date"):
            form["source_date"] = re.sub(r"\D", "", str(options["date"]))[:8]
        if not form["source_date"]:
            raise ValueError("缺少生产日期：请在页面填写生产日期后重新导入")
    else:
        form = parse_sales_form(sales_path)
    if options.get("plan_qty") not in (None, ""):
        options["plan_kg"] = plan_to_kg(form, options["plan_qty"], options.get("plan_unit") or "kg")
    # AI/解析表单里带的销售、订单号、交期自动补入（页面显式填写优先）
    if not options.get("manager") and form.get("sales"):
        options["manager"] = form["sales"]
    if not options.get("order_no") and form.get("order_no"):
        options["order_no"] = form["order_no"]
    if not options.get("due") and form.get("due"):
        options["due"] = form["due"]
    explicit = _parse_explicit_serial(options.get("serial"), form["source_date"])
    if explicit:
        order_date, sequence = explicit
    else:
        order_date = form["source_date"]
        sequence = None
    history_root = Path(history_root) if history_root else _default_history_root(order_date)
    requested_serial = None
    if sequence is not None:
        requested_serial = int(order_date + "%02d" % sequence)
    elif options.get("sequence"):
        requested_serial = int(form["source_date"] + "%02d" % int(options["sequence"]))
    serial_text = ("%s-%02d" % (order_date, requested_serial % 100)
                   if requested_serial else "")

    ledger_order = _ledger_pending_order(serial_text) if requested_serial else None
    if ledger_order is not None and str(ledger_order.get("status") or "待生产") != "待生产":
        raise ValueError("台账中 %s 状态为「%s」，仅待生产订单可按新销售单更新"
                         % (serial_text, ledger_order.get("status")))
    if ledger_order is not None and options.get("plan_kg") is None:
        raise ValueError("挂接台账批次 %s 需要填写计划产量" % serial_text)

    if ledger_order is not None:
        # 挂接台账既有批次：生产编号沿用台账，计划/配方/批号预占按本次销售单更新；
        # 同一编号多次生产时，生产批次（H3）自动递增（01→02→…）
        runs = ledger_order.get("production_runs")
        if not isinstance(runs, list):
            has_history = bool(ledger_order.get("production_sheet_imports")
                               or ledger_order.get("total_kg")
                               or ledger_order.get("materials"))
            runs = [{"run_no": 1}] if has_history else []
        run_no = len(runs) + 1
        order = {
            "date": order_date,
            "serial": requested_serial,
            "product": _clean_product(options.get("product") or form["product"]),
            "plan_kg": float(options["plan_kg"]),
            "plan_text": "%gkg" % float(options["plan_kg"]),
            "plan_display": _plan_display(float(options["plan_kg"]), form),
            "maker": options.get("maker") or "李浩冉",
            "manager": options.get("manager") or str(ledger_order.get("sales") or ""),
            "batch": str(options.get("batch") or "%02d" % run_no),
            "order_no": options.get("order_no") or str(ledger_order.get("order_no") or ""),
            "due": options.get("due") or str(ledger_order.get("due") or ""),
            "material_batches": {},
            "run_no": run_no,
            "run_date": str(options.get("run_date") or form["source_date"]),
        }
        order_folder = _existing_batch_folder(history_root, order_date, sequence) or (
            history_root / order_date / (
                "%s-%02d_%s" % (order_date, sequence,
                                _safe_path_segment(order["product"]))))
        attach_target = serial_text
    else:
        attach_target = None
        run_no = None
        try:
            if options.get("plan_kg") is None:
                raise LookupError("未提供计划产量")
            order = _find_existing_order(history_root, order_date, sales_path,
                                         requested_serial)
            if (options.get("plan_kg") is not None
                    and float(options["plan_kg"]) != float(order["plan_kg"])):
                raise PlanMismatchError(
                    "匹配到既有生产记录 %s-%02d（计划 %g kg）与本次计划 %g kg 不一致；"
                    "如确属新一批，请填写新的生产编号后重新导入"
                    % (order_date, int(order["serial"]) % 100,
                       float(order["plan_kg"]), float(options["plan_kg"])))
        except (FileNotFoundError, ValueError, LookupError):
            if sequence is None:
                sequence = int(options.get("sequence")
                               or _next_sequence(history_root, order_date))
            if not 1 <= sequence <= 99:
                raise ValueError("生产序号必须在 01-99 之间")
            if _serial_exists(history_root, order_date, sequence):
                raise ValueError(
                    "生产编号 %s-%02d 已存在，不能重复占用；请改用其他序号或留空自动分配"
                    % (order_date, sequence))
            order = {
                "date": order_date,
                "serial": int(order_date + "%02d" % sequence),
                "product": _clean_product(options.get("product") or form["product"]),
                "plan_kg": float(options["plan_kg"]) if options.get("plan_kg") is not None else 0,
                "plan_text": ("%gkg" % float(options["plan_kg"]))
                if options.get("plan_kg") is not None else "",
                "plan_display": _plan_display(float(options["plan_kg"]), form)
                if options.get("plan_kg") is not None else "",
                "maker": options.get("maker") or "李浩冉",
                "manager": options.get("manager", ""),
                "batch": str(options.get("batch") or ("%02d" % sequence)),
                "order_no": options.get("order_no", ""),
                "due": options.get("due", ""),
                "material_batches": {},
            }
            order_folder = history_root / order_date / (
                "%s-%02d_%s" % (
                    order_date, sequence,
                    _safe_path_segment(order["product"])))
        else:
            order["date"] = order_date
            order["order_no"] = options.get("order_no", "")
            order["due"] = options.get("due", "")
            order["maker"] = str(order.get("maker") or "").strip() or "李浩冉"
            order["plan_display"] = _plan_display(float(order.get("plan_kg") or 0), form)
            order_folder = order["path"].parent
    sequence = order["serial"] % 100
    output_name = "%s-%02d_%s_销售单导入_%gkg.xlsx" % (
        order_date, sequence, _safe_path_segment(order["product"]),
        order["plan_kg"])
    if run_no and run_no > 1:
        output_name = "%s_第%02d批.xlsx" % (output_name[:-5], run_no)
    result = fill_production_template(
        form, template_path, order_folder / output_name, order, stock_rows=stock_rows)
    result["source"] = str(sales_path)
    result["history_reference"] = str(order.get("path", order_folder))
    merged = {}
    for stage_items in result["materials"].values():
        for material in stage_items:
            if not material.get("qty"):
                continue
            batches = "；".join(a["batch"] for a in material["batch_plan"]
                               if a.get("batch")) or material["batch"] or ""
            item = {"name": material["display_name"],
                    "kg": round(float(material["qty"]), 3), "batch": batches}
            if item["name"] in merged:
                merged[item["name"]]["kg"] = round(
                    merged[item["name"]]["kg"] + item["kg"], 3)
            else:
                merged[item["name"]] = item
    if attach_target:
        try:
            from orders import attach_sales_import
            ledger_result = attach_sales_import(
                attach_target, order["product"], order["plan_kg"],
                list(merged.values()), Path(result["output"]).name,
                hashlib.sha256(Path(result["output"]).read_bytes()).hexdigest(),
                manager=order.get("manager", ""),
                run_date=str(options.get("run_date") or form["source_date"]))
        except Exception as exc:
            result["warnings"] = list(result.get("warnings") or []) + [
                "交接单已生成，但台账更新失败：%s: %s" % (type(exc).__name__, exc)]
        else:
            result["attached"] = True
            result["run_no"] = ledger_result.get("run_no")
            result["ledger"] = ledger_result
            result["warnings"] = list(result.get("warnings") or []) + \
                ledger_result.get("warnings", [])
    else:
        # 新编号：直接登记台账并记第 1 批（此前需要再手动「导入生产表」才能入账）
        try:
            from orders import add_order, order_exists, set_order_file
            serial_full = "%s-%02d" % (order_date, sequence)
            if not order_exists(serial_full, order["product"]):
                added, message = add_order({
                    "serial": serial_full,
                    "product": order["product"],
                    "wan": 0,
                    "tablet": form["tablet_g"] or 0.5,
                    "total_kg": order.get("plan_kg") or None,
                    "customer": "",
                    "due": order.get("due", ""),
                    "sales": order.get("manager", ""),
                    "status": "待生产",
                    "order_date": order_date,
                    "batch": order.get("batch", "01"),
                    "order_type": options.get("order_type") or "正式生产",
                    "materials": [], "packaging": [],
                })
                if added:
                    set_order_file(serial_full, order["product"], result["output"], {
                        "total_kg": order.get("plan_kg") or None,
                        "materials": list(merged.values()),
                    }, run_note="销售单导入")
                    result["ledger_serial"] = serial_full
                else:
                    result["warnings"] = list(result.get("warnings") or []) + [
                        "台账登记未完成：%s" % message]
        except Exception as exc:
            result["warnings"] = list(result.get("warnings") or []) + [
                "交接单已生成，但台账登记失败：%s: %s" % (type(exc).__name__, exc)]
    return result


# ================= AI 提取与校验（AI 只负责理解，事实由校验器与人工确认把关） =================

AI_FORM_SCHEMA = """{
  "product": "产品全名（按单原文）",
  "order_date": "生产日期 8 位数字 YYYYMMDD，单上没有就填空字符串",
  "order_no": "订单号（按单原文，没有填空字符串）",
  "due_date": "交货日期 8 位数字 YYYYMMDD，没有就填空字符串",
  "sales": "下单销售/销售经理/业务员姓名（按单原文，没有填空字符串）",
  "tablet_g": 片重数字（g/片，没有填 null）,
  "per_bottle": 每瓶片数整数（没有填 null）,
  "coat_type": "红色包衣/胃溶包衣/肠溶包衣/薄膜包衣/糖衣 之一，没有填空字符串",
  "materials": [{"name": "原料名（按单原文）", "ratio": 配比数字（%），"quote": "该行原文"}],
  "coat": {"name": "包衣剂名（按单原文）", "ratio": 包衣比例数字（%，按素片重量计）,
            "quote": "该行原文"} 或 null,
  "packaging": [{"kind": "瓶/盖/标签/纸盒/防伪码", "detail": "括号内原文"}],
  "press_process": "工艺原文整段（配料/混料/压片步骤都要，含各自的 min 时间、震动下料、克重范围、硬度要求；没有填空字符串）",
  "coat_process": "包衣工艺原文（整段，没有填空字符串）"
}"""


def _normalize_ai_name(value):
    return re.sub(r"[\s\u3000]+", "", str(value or "")).strip()


def _ai_number(value, field, minimum=None, maximum=None):
    if isinstance(value, bool) or value in (None, ""):
        return None
    try:
        number = float(value)
    except (TypeError, ValueError):
        raise ValueError("字段 %s 不是数字：%r" % (field, value))
    if minimum is not None and number < minimum:
        raise ValueError("字段 %s 低于合理下限 %s：%s" % (field, minimum, number))
    if maximum is not None and number > maximum:
        raise ValueError("字段 %s 超出合理上限 %s：%s" % (field, maximum, number))
    return number


def validate_ai_form(payload, stock_names=None):
    """确定性校验 AI 提取结果。返回 (problems, warnings, unknown_materials)。
    problems 非空 = 不可导入；warnings 需要人工在预览里确认。"""
    problems, warnings, unknown = [], [], []
    if not isinstance(payload, dict):
        raise ValueError("AI 提取结果不是对象")
    product = _normalize_ai_name(payload.get("product"))
    if not product:
        problems.append("产品名为空")
    materials = payload.get("materials")
    if not isinstance(materials, list) or len(materials) < 1:
        problems.append("未提取到任何原料配比行")
        materials = []
    total = 0.0
    for index, item in enumerate(materials, 1):
        if not isinstance(item, dict) or not _normalize_ai_name(item.get("name")):
            problems.append("第 %d 个原料缺名称" % index)
            continue
        try:
            ratio = _ai_number(item.get("ratio"), "第 %d 个原料配比" % index, 0, 100)
        except ValueError as exc:
            problems.append(str(exc))
            continue
        if ratio is None:
            problems.append("第 %d 个原料（%s）配比缺失或不是数字"
                            % (index, _normalize_ai_name(item.get("name"))))
            continue
        total += ratio
        if not str(item.get("quote") or "").strip():
            warnings.append("原料「%s」没有原文出处引用，请对照原单核对"
                            % _normalize_ai_name(item.get("name")))
    coat = payload.get("coat") if isinstance(payload.get("coat"), dict) else None
    if coat:
        try:
            coat_ratio = _ai_number(coat.get("ratio"), "包衣比例", 0, 30)
        except ValueError as exc:
            problems.append(str(exc))
            coat_ratio = None
        if not str(coat.get("quote") or "").strip():
            warnings.append("包衣剂「%s」没有原文出处引用" %
                            _normalize_ai_name(coat.get("name")))
    if materials and abs(total - 100.0) > 2.0:
        problems.append("原料配比合计 %.3f%%，偏离 100%% 超过 2%%，疑似漏行/多行" % total)
    elif materials and abs(total - 100.0) > 0.05:
        warnings.append("原料配比合计 %.3f%%（应为 100%%），请核对" % total)
    try:
        tablet_g = _ai_number(payload.get("tablet_g"), "片重", 0.05, 3.0)
    except ValueError as exc:
        problems.append(str(exc))
        tablet_g = None
    if tablet_g is None:
        problems.append("片重（g/片）未提取到；可在导入页手工填写计划量规避换算，"
                        "但交接单规格栏会留空")
    per_bottle = payload.get("per_bottle")
    if per_bottle not in (None, ""):
        try:
            per_bottle = _ai_number(per_bottle, "每瓶片数", 1, 10000)
        except ValueError as exc:
            problems.append(str(exc))
    order_date = str(payload.get("order_date") or "").strip()
    if order_date:
        if not re.fullmatch(r"20\d{6}", order_date):
            problems.append("生产日期 %r 不是 8 位 YYYYMMDD" % order_date)
        else:
            try:
                datetime.datetime.strptime(order_date, "%Y%m%d")
            except ValueError:
                problems.append("生产日期 %s 不是真实日期" % order_date)
    packaging = payload.get("packaging")
    if packaging and not isinstance(packaging, list):
        problems.append("包装信息格式异常")
    if stock_names is not None:
        normalized = {_normalize_ai_name(name): name for name in stock_names}
        for item in materials:
            name = _normalize_ai_name(item.get("name") if isinstance(item, dict) else "")
            if not name:
                continue
            if name in normalized:
                continue
            fuzzy = [original for key, original in normalized.items()
                     if name in key or key in name]
            if len(fuzzy) == 1:
                warnings.append("原料「%s」将按库存档案「%s」对齐，请确认"
                                % (name, fuzzy[0]))
            else:
                unknown.append(name)
        if unknown:
            warnings.append("以下原料在库存档案中找不到，批号无法自动带出：%s"
                            % "、".join(unknown))
    return problems, warnings, unknown


_PRESS_SPLIT = re.compile(r"(?:\d+\s*[.．、]\s*)?压片\s*[:：]")


def _split_press_sections(text):
    """把工艺原文拆成（混料部分, 压片部分）：取最后一个「压片：」标记（步骤而非栏头），
    压片及其后的内容不再混进混料说明；整段以「压片：」栏头开头且无后续压片步骤时不拆。"""
    text = str(text or "")
    strip_chars = " \t\r\n /；;。"
    markers = list(_PRESS_SPLIT.finditer(text))
    if not markers:
        return (text.strip(strip_chars), "")
    match = markers[-1]
    if match.start() == 0:
        return (text.strip(strip_chars), "")
    press = re.sub(r"^(?:\d+\s*[.．、]\s*)", "",
                   text[match.start():].strip(strip_chars))
    return (text[:match.start()].strip(strip_chars), press)


def form_from_ai(payload, source_name):
    """把通过校验的 AI JSON 规范化为 parse_sales_form 同构的 form 字典。"""
    problems, _warnings, _unknown = validate_ai_form(payload, stock_names=None)
    if problems:
        raise ValueError("AI 提取结果未通过校验：%s" % "；".join(problems))
    product = _normalize_ai_name(payload.get("product"))
    materials = []
    coat_name = ""
    coat_ratio = None
    for item in payload.get("materials") or []:
        ratio = _ai_number(item.get("ratio"), "原料配比", 0, 100)
        if ratio is None:
            continue
        materials.append({"name": str(item.get("name")).strip(), "ratio": ratio})
    coat = payload.get("coat") if isinstance(payload.get("coat"), dict) else None
    if coat and _normalize_ai_name(coat.get("name")):
        coat_name = str(coat.get("name")).strip()
        coat_ratio = _ai_number(coat.get("ratio"), "包衣比例", 0, 30)
    tablet_g = _ai_number(payload.get("tablet_g"), "片重", 0.05, 3.0)
    per_bottle = payload.get("per_bottle")
    per_bottle = int(per_bottle) if per_bottle not in (None, "") else None
    order_date = str(payload.get("order_date") or "").strip()
    press = str(payload.get("press_process") or "").strip()
    coat_process = str(payload.get("coat_process") or "").strip()
    packaging = []
    for item in payload.get("packaging") or []:
        if not isinstance(item, dict):
            continue
        kind = str(item.get("kind") or "").strip()
        detail = str(item.get("detail") or "").strip()
        if kind:
            packaging.append({"name": "%s（%s）" % (kind, detail), "kind": kind})
    coat_type = str(payload.get("coat_type") or "").strip()
    if coat_type not in ("红色包衣", "胃溶包衣", "肠溶包衣", "薄膜包衣", "糖衣"):
        coat_type = next((name for name in
                          ("红色包衣", "胃溶包衣", "肠溶包衣", "薄膜包衣", "糖衣")
                          if name in (coat_name + coat_process)), "")
    coating_steps = _coating_steps(coat_process)
    due_digits = re.sub(r"\D", "", str(payload.get("due_date") or ""))[:8]
    return {
        "source": source_name,
        "source_date": order_date,
        "order_no": str(payload.get("order_no") or "").strip(),
        "due": due_digits,
        "sales": str(payload.get("sales") or "").strip(),
        "product": product,
        "materials": materials,
        "ratio_total": round(sum(m["ratio"] for m in materials), 3),
        "coat_name": coat_name,
        "coat_ratio": coat_ratio,
        "coat_type": coat_type,
        "tablet_g": tablet_g,
        "per_bottle": per_bottle,
        "packaging": packaging,
        "press_process": press,
        "prep_instruction": _prep_instruction(press),
        "coat_process": coat_process,
        "coating_steps": coating_steps,
        "ai_extracted": True,
    }


def ai_extract_form(docx_path=None, raw_text=None, stock_names=None):
    """AI 提取单据（Word 文件、纯文字，或两者合并）→ 校验后的 form 字典 + 溯源与校验结果。
    纯提取，不写任何文件、不动台账。两者都给时合并为一张订单：文字里明确给出的字段优先。"""
    import ai_assistant
    if not ai_assistant.is_configured():
        raise RuntimeError("AI 未配置（data\\ai_config.json）")
    sections = []
    if docx_path:
        import ai_assistant as _ai
        _ai.set_progress("读取 Word 单据", "解析表格与段落…")
        document = Document(str(docx_path))
        lines = []
        for table in document.tables:
            for row in table.rows:
                line = " | ".join(cell.text.strip() for cell in row.cells
                                  if cell.text.strip())
                if line:
                    lines.append(line)
        for paragraph in document.paragraphs:
            text = paragraph.text.strip()
            if text:
                lines.append(text)
        docx_content = "\n".join(lines)
        if len(docx_content.strip()) >= 10:
            sections.append("【Word 单据内容】\n" + docx_content)
    extra_text = str(raw_text or "").strip()[:8000]
    if extra_text and len(extra_text) >= 5:
        sections.append("【客户补充文字订单（微信/电话记录）】\n" + extra_text)
    if not sections:
        raise ValueError("单据内容太短，无法提取")
    if len(sections) == 2:
        source_kind = "Word 销售单 + 补充文字订单"
    elif docx_path:
        source_kind = "Word 销售单"
    else:
        source_kind = "文字订单"
    content = "\n\n".join(sections)
    if len(content) > 24000:
        content = content[:24000]
    import ai_assistant as _ai
    _ai.set_progress("整理单据内容", "共 %d 字，准备提取" % len(content))
    merge_rule = ""
    if len(sections) == 2:
        merge_rule = ("5) 两份材料合并为一张订单：补充文字里明确给出的信息（产品、日期、数量、"
                      "包装、配方数字）优先采用；Word 独有的内容照常保留；两者矛盾时以文字为准；"
                      "并在 JSON 顶层增加 \"merge_notes\" 数组（字符串列表），逐条说明关键取舍，"
                      "例如「生产日期取自补充文字 20261105」「配比采用文字订单的 43.5/30/15/10/1.5」。")
    system_prompt = (
        "你是食品厂单据提取助手。从用户提供的%s里抽取结构化信息，"
        "只输出一个 JSON 对象，字段结构如下（不要输出任何解释或代码块标记）：\n%s\n"
        "规则：1) 一切以单据原文为准，禁止猜测或编造；单上没有的字段填 null 或空字符串；"
        "2) 每个原料的 quote 必须逐字复制单据里对应那一行的原文；"
        "3) 包衣剂（被膜剂/包衣粉）按比例外加时不计入 materials，放进 coat 字段；"
        "4) 配比是百分比数字，不带 %% 符号。%s" % (source_kind, AI_FORM_SCHEMA, merge_rule))
    data, meta = ai_assistant.extract_json(system_prompt, content)
    ai_assistant.set_progress("校验提取结果", "配比合计/物料对齐/日期检查…")
    problems, warnings, unknown = validate_ai_form(data, stock_names=stock_names)
    merge_notes = data.get("merge_notes") if isinstance(data.get("merge_notes"), list) else []
    warnings = list(warnings) + [str(n) for n in merge_notes if str(n).strip()]
    form = form_from_ai(data, source_name=source_kind)
    return {"form": form, "ai": meta, "problems": problems,
            "warnings": warnings, "unknown_materials": unknown,
            "raw": data}


def preview_sales_form(sales_path, history_root=None):
    """导入前解析预览：配方/规格/日期 + 台账同日批次建议（不写任何文件）。"""
    form = parse_sales_form(sales_path)
    order_date = form["source_date"]
    history_root = Path(history_root) if history_root else _default_history_root(order_date)
    try:
        from orders import load as load_orders
        day_orders = [
            {"serial": o.get("serial"), "product": o.get("product"),
             "status": o.get("status"), "total_kg": o.get("total_kg"),
             "batch": o.get("batch"),
             "attachable": o.get("status") == "待生产"}
            for o in load_orders().get("orders", [])
            if str(o.get("serial") or "").startswith(order_date + "-")]
    except Exception:
        day_orders = []
    product_key = _clean_product(form["product"])
    attachable = [o for o in day_orders if o["attachable"]]
    matching = [o for o in attachable
                if product_key and (
                    product_key in _clean_product(o.get("product") or "")
                    or _clean_product(o.get("product") or "") in product_key)]
    suggest_serial = matching[0]["serial"] if len(matching) == 1 else ""
    try:
        used = set(_ledger_serial_sequences(order_date))
        seq = _next_sequence(history_root, order_date)
        while seq in used:
            seq += 1
        next_serial = "%s-%02d" % (order_date, seq)
    except Exception:
        next_serial = ""
    hints = []
    if not form.get("per_bottle"):
        hints.append("销售单未写每瓶片数，计划瓶数无法折算")
    if not form.get("coat_type"):
        hints.append("未识别到包衣剂（素片工艺）")
    if not form.get("packaging"):
        hints.append("包装及标签未提供，包装段留空")
    return {
        "ok": True,
        "product": form["product"],
        "tablet_g": form["tablet_g"],
        "tablet_range": form.get("tablet_range") or "",
        "hardness": form.get("hardness") or "",
        "per_bottle": form["per_bottle"],
        "coat_type": form.get("coat_type") or "",
        "materials": [{"name": m["name"], "ratio": m["ratio"]}
                      for m in form["materials"]],
        "order_date": order_date,
        "next_serial": next_serial,
        "ledger_day_orders": day_orders,
        "suggest_serial": suggest_serial,
        "hints": hints,
    }


if __name__ == "__main__":
    import argparse
    import json

    parser = argparse.ArgumentParser(description="将 Word 销售配方单导入生产指令单")
    parser.add_argument("sales_docx", help="销售/配方 Word 文件")
    parser.add_argument("template_xlsx", nargs="?",
                        default=str(Path(__file__).parent / "templates" / "生产指令与交接单.xlsx"),
                        help="四段式生产指令单母版（默认使用程序模板目录）")
    parser.add_argument("history_root", nargs="?", default=None,
                        help="按月份分组的生产记录根目录；默认自动查找桌面月份目录")
    parser.add_argument("--plan-kg", type=float, help="计划产量（kg）；不填则留空")
    parser.add_argument("--plan-bottles", type=float, help="计划产量（瓶），按销售单每瓶片数和片重换算为 kg")
    parser.add_argument("--plan-tablets", type=float, help="计划产量（片），按销售单片重换算为 kg")
    parser.add_argument("--product", help="明确指定生产产品名；默认沿用唯一匹配历史单或来源名称")
    parser.add_argument("--sequence", type=int, help="指定生产序号 01-99；默认沿用历史单或自动取下一个")
    parser.add_argument("--maker", default="李浩冉", help="制表人（新建记录时，默认李浩冉）")
    parser.add_argument("--manager", default="", help="销售经理（新建记录时）")
    parser.add_argument("--batch", default="01", help="生产批次（默认 01）")
    parser.add_argument("--order-no", default="", help="销售订单号（可选）")
    parser.add_argument("--due", default="", help="交期 YYYYMMDD（可选）")
    args = parser.parse_args()
    try:
        result = import_sales_form(args.sales_docx, args.template_xlsx, args.history_root, {
            "plan_kg": args.plan_kg,
            **({"plan_qty": args.plan_bottles, "plan_unit": "bottle"} if args.plan_bottles else {}),
            **({"plan_qty": args.plan_tablets, "plan_unit": "tablet"} if args.plan_tablets else {}),
            "product": args.product,
            "sequence": args.sequence, "maker": args.maker,
            "manager": args.manager, "batch": args.batch,
            "order_no": args.order_no, "due": args.due,
        })
    except Exception as exc:
        parser.exit(1, "导入失败：%s: %s\n" % (type(exc).__name__, exc))
    print(json.dumps(result, ensure_ascii=False, indent=2))
