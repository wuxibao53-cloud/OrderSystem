# -*- coding: utf-8 -*-
"""库存管理：批次管理、库存播报、采购建议、日报、网页看板"""
import os
import sys
import json
import math
import re
import ctypes
import datetime
from json_store import save_json_atomic

APP = os.path.dirname(os.path.abspath(__file__))   # 便携：系统放哪个目录都能跑
STOCK = os.path.join(APP, "stock.json")
MAPPING = os.path.join(APP, "mapping.json")
RECIPE = os.path.join(APP, "recipes.json")
DAILY_DIR = os.path.join(APP, "data", "daily")
DASH_DIR = os.path.join(APP, "data", "dashboard")

sys.path.insert(0, APP)

try:
    _k = ctypes.windll.kernel32
    _h = _k.GetStdHandle(-11)
    _m = ctypes.c_uint32()
    _k.GetConsoleMode(_h, ctypes.byref(_m))
    _k.SetConsoleMode(_h, _m.value | 0x0004)   # 启用 ANSI
    _k.SetConsoleOutputCP(65001)
    _k.SetConsoleCP(65001)
except Exception:
    pass
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

PREFIXES = ("直接原料-", "添加剂-", "辅料-", "包装袋-", "包装瓶-",
            "包装膜-", "标签-", "瓶盖-", "干燥剂-")
BATCH_KG = 10.0   # 标准批次：2 万片 × 0.5g

C = {
    "b": "\033[1m", "dim": "\033[2m", "r": "\033[0m",
    "red": "\033[91m", "yel": "\033[93m", "grn": "\033[92m",
    "cyn": "\033[96m", "blu": "\033[94m", "gry": "\033[90m",
}


def short_name(full):
    for p in PREFIXES:
        if full.startswith(p):
            return full[len(p):]
    return full


def load_json(path, default=None):
    if not os.path.exists(path):
        return default if default is not None else {}
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def save_json(path, data):
    save_json_atomic(path, data)


def clear():
    os.system("cls" if os.name == "nt" else "clear")


def bar(title, width=66):
    print("%s╔%s╗%s" % (C["blu"], "═" * width, C["r"]))
    pad = width - len(title) * 2
    left = max(0, pad // 2)
    print("%s║%s%s%s║%s" % (C["blu"], " " * left, C["b"] + title + C["r"],
                            " " * max(0, pad - left), C["r"]))
    print("%s╚%s╝%s" % (C["blu"], "═" * width, C["r"]))


# ================= 导入 =================
class _Table(object):
    """统一 .xls / .xlsx 读取，接口与 xlrd sheet 一致（nrows / ncols / cell_value）"""

    def __init__(self, rows):
        self._rows = rows
        self.nrows = len(rows)
        self.ncols = max((len(r) for r in rows), default=0)

    def cell_value(self, r, c):
        try:
            v = self._rows[r][c]
        except Exception:
            return ""
        return "" if v is None else v


def open_table(path):
    ext = os.path.splitext(str(path))[1].lower()
    if ext == ".xls":
        import xlrd
        book = xlrd.open_workbook(path)
        sh = book.sheet_by_index(0)
        rows = [[sh.cell_value(r, c) for c in range(sh.ncols)] for r in range(sh.nrows)]
    elif ext == ".xlsx":
        import openpyxl
        wb = openpyxl.load_workbook(path, data_only=True)
        try:
            ws = wb[wb.sheetnames[0]]
            rows = []
            for r in range(ws.max_row or 1):
                rows.append([ws.cell(row=r + 1, column=c + 1).value
                             for c in range(ws.max_column or 1)])
        finally:
            wb.close()
    else:
        raise ValueError("只支持 .xls 或 .xlsx 库存表")
    return _Table(rows)


def _stock_date_key(value):
    text = str(value or "").strip().replace("/", "-")
    if len(text) >= 10:
        try:
            return datetime.datetime.strptime(text[:10], "%Y-%m-%d").date()
        except ValueError:
            return None
    return None


def import_xls(path):
    path = os.fspath(path)
    sh = open_table(path)
    header_row = find_header_row(sh, ["商品名称", "物料名称"], maxr=8)
    if header_row < 0:
        raise ValueError("库存表前 8 行找不到「商品名称」表头，未修改现有库存")
    head = [str(sh.cell_value(header_row, c)).strip() for c in range(sh.ncols)]
    required_headers = {
        1: ("编码",),
        2: ("商品名称", "物料名称"),
    }
    for column, candidates in required_headers.items():
        actual = head[column] if column < len(head) else ""
        if not any(candidate in actual for candidate in candidates):
            raise ValueError(
                "库存表列布局与系统要求不符（第 %d 列应为 %s，实际为「%s」），未修改现有库存"
                % (column + 1, " / ".join(candidates), actual)
            )
    quantity_header = head[8] if len(head) > 8 else ""
    quantity_subheader = (
        str(sh.cell_value(header_row + 1, 8)).strip()
        if header_row + 1 < sh.nrows else ""
    )
    has_quantity_header = any(k in quantity_header for k in ("数量", "余额"))
    has_total_warehouse_quantity = (
        "全部仓库" in quantity_header
        and quantity_subheader in ("单位数量", "库存单位数量", "多单位数量")
    )
    if not (has_quantity_header or has_total_warehouse_quantity):
        raise ValueError(
            "库存表列布局与系统要求不符（第 9 列应为数量 / 余额，"
            "或「全部仓库」下有数量子表头，实际为「%s / %s」），未修改现有库存"
            % (quantity_header, quantity_subheader)
        )
    company, stock_date = "", ""
    for r in range(min(4, sh.nrows)):
        for c in range(min(6, sh.ncols)):
            v = str(sh.cell_value(r, c))
            if "库存日期" in v:
                stock_date = v.split("库存日期")[-1].strip(" :：")
                head = v.split("库存日期")[0]
                if "公司" in head:
                    company = head.partition("公司")[2].strip("：: ")
            elif "公司" in v and not company:
                company = v.partition("公司")[2].strip("：: ")

    old = {i["name"]: i for i in load_json(STOCK, {}).get("items", [])}
    # 库存表里如果自带「批号 / 批次」列，就顺带把批号解析出来（无需另外导批次表）
    bmap = extract_batches_from_stock(sh)

    items = []
    seen = {}
    for r in range(max(header_row + 1, 3), sh.nrows):
        def cell(c):
            return sh.cell_value(r, c)

        name = str(cell(2)).strip()
        code = str(cell(1)).strip()
        if not name or not code or name in ("商品名称", "物料名称"):
            continue

        def num(c):
            try:
                return float(cell(c) or 0)
            except Exception:
                return 0.0

        quantity_raw = cell(8)
        quantity_text = str(quantity_raw).strip() if quantity_raw is not None else ""
        if "," in quantity_text:
            if not re.fullmatch(r"[+-]?\d{1,3}(?:,\d{3})+(?:\.\d+)?", quantity_text):
                raise ValueError("库存表第 %d 行数量格式无效（%r），未修改现有库存" %
                                 (r + 1, quantity_raw))
            quantity_text = quantity_text.replace(",", "")
        try:
            quantity = float(quantity_text or 0)
        except (TypeError, ValueError):
            raise ValueError("库存表第 %d 行数量无效（%r），未修改现有库存" %
                             (r + 1, quantity_raw))
        if not math.isfinite(quantity):
            raise ValueError("库存表第 %d 行数量不是有限数值，未修改现有库存" % (r + 1))

        cat = str(cell(3)).strip()
        prev = old.get(name, {})
        safe = prev.get("safe")
        if not safe:
            safe = (20.0 if str(cell(5)).strip() == "KG" else 500.0) \
                if cat.startswith("原材料-包装材料") else 0.0
        row = {
            "name": name, "short": short_name(name), "code": code, "cat": cat,
            "spec": str(cell(4)).strip(), "unit": str(cell(5)).strip(),
            "qty": round(quantity, 3),
            "wh_fg": round(num(18), 3), "wh_raw": round(num(23), 3),
            "wh_pk": round(num(28), 3), "wh_add": round(num(33), 3),
            "wh_aux": round(num(38), 3),
            "safe": safe,
            "batches": prev.get("batches", []),
        }
        # 库存表自带批次列时：同商品多批次会占多行，合并成一条；批号以表内为准
        if bmap.get(name):
            bs = bmap[name]
            row["batches"] = [{"batch": b["batch"], "qty": b["qty"],
                               "mfg": b.get("mfg", ""), "exp": b.get("exp", "")}
                              for b in sorted(bs.values(),
                                              key=lambda x: (x.get("exp") or "9999", x.get("batch", "")))]
            if name in seen:                      # 同一商品的多批次行：数量累加
                seen[name]["qty"] = round(seen[name]["qty"] + row["qty"], 3)
                for k in ("wh_fg", "wh_raw", "wh_pk", "wh_add", "wh_aux"):
                    seen[name][k] = round(seen[name][k] + row[k], 3)
                continue
        elif name in seen:
            continue                              # 无批次列时保持一条一商品
        seen[name] = row
        items.append(row)

    if not items:
        raise ValueError("库存表中没有可识别的商品行，未修改现有库存")

    previous_date = _stock_date_key(load_json(STOCK, {}).get("_meta", {}).get("stock_date"))
    incoming_date = _stock_date_key(stock_date)
    if previous_date and incoming_date and incoming_date < previous_date:
        raise ValueError("库存表日期 %s 早于当前账面日期 %s，已保留现有库存" %
                         (stock_date, load_json(STOCK, {}).get("_meta", {}).get("stock_date")))

    meta = {"source": path, "company": company, "stock_date": stock_date,
            "imported_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M")}
    if bmap:
        meta["inline_batches"] = len(bmap)
    save_json(STOCK, {"_meta": meta, "items": items})
    return len(items), company, stock_date


def extract_batches_from_stock(sh):
    """库存表里若自带「批号 / 批次」列，直接解析出 {商品名: {批号: {...}}}。
    没有批次列就返回空字典（不影响原有导入逻辑）。"""
    hr = find_header_row(sh, ["商品名称", "物料名称"], maxr=8)
    if hr < 0:
        return {}
    head = [str(sh.cell_value(hr, c)).strip() for c in range(sh.ncols)]

    def col(cands, start=0):
        for i in range(start, len(head)):
            for k in cands:
                if k in head[i]:
                    return i
        return -1

    c_name = col(["商品名称", "物料名称", "存货名称"])
    c_batch = col(["批号", "批次号", "批次"])
    if c_name < 0 or c_batch < 0:
        return {}
    c_qty = col(["结存数量", "库存数量", "数量", "余额"], start=c_batch)
    if c_qty < 0:
        c_qty = col(["结存数量", "库存数量", "数量", "余额"])
    c_mfg = col(["生产日期", "制造日期"])
    c_exp = col(["到期日期", "有效期至", "失效日期", "保质期至"])
    c_in_date = col(["入库日期", "收货日期", "入库时间", "单据日期"])

    out = {}
    for r in range(hr + 1, sh.nrows):
        name = str(sh.cell_value(r, c_name)).strip()
        bno = str(sh.cell_value(r, c_batch)).strip()
        if not name or not bno or name in ("商品名称", "物料名称", "合计", "小计"):
            continue
        if bno.endswith(".0"):
            bno = bno[:-2]
        try:
            qty = round(float(sh.cell_value(r, c_qty) or 0), 3) if c_qty >= 0 else 0.0
        except Exception:
            qty = 0.0
        d = out.setdefault(name, {})
        if bno in d:
            d[bno]["qty"] = round(d[bno]["qty"] + qty, 3)
        else:
            d[bno] = {"batch": bno, "qty": qty,
                      "mfg": _cell_date(sh.cell_value(r, c_mfg)) if c_mfg >= 0 else "",
                      "exp": _cell_date(sh.cell_value(r, c_exp)) if c_exp >= 0 else "",
                      "in_date": _cell_date(sh.cell_value(r, c_in_date))
                      if c_in_date >= 0 else ""}
    return out


# ================= 批次表导入（柠檬云「批次状态表 / 批次跟踪表」）=================
def _cell_date(v):
    """Excel 单元格转日期字符串：支持 datetime / 序列号 / 文本"""
    if v in (None, "", 0):
        return ""
    if isinstance(v, datetime.datetime):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, datetime.date):
        return v.strftime("%Y-%m-%d")
    try:
        f = float(v)
        if 20000 < f < 60000:      # Excel 日期序列号
            return (datetime.datetime(1899, 12, 30) + datetime.timedelta(days=f)).strftime("%Y-%m-%d")
    except Exception:
        pass
    s = str(v).strip().replace("/", "-").replace(".", "-")
    return s


def find_header_row(sh, keys, maxr=12):
    """在前 maxr 行里找表头行。
    优先选择命中关键词最多的行，再用非空单元格数打破平手；
    避免多层表头的子标题行因填充单元格更多而盖过真正的字段名行。"""
    best, best_score = -1, 0
    for r in range(min(maxr, sh.nrows)):
        row = [str(sh.cell_value(r, c)).strip() for c in range(sh.ncols)]
        filled = sum(1 for v in row if v)
        hit_cells = sum(1 for v in row if v and any(k in v for k in keys))
        score = hit_cells * 1000 + filled
        if score > best_score:
            best, best_score = r, score
    return best


def import_batch_xls(path):
    """导入柠檬云批次状态表 / 批次跟踪表（带「批号」列的任意表都尽量兼容）。
    返回 (匹配物料数, 批次数, 未匹配清单)"""
    sh = open_table(path)
    hr = find_header_row(sh, ["批号", "批次"])
    if hr < 0:
        raise ValueError("这张表里找不到「批号 / 批次」列，请确认导出的是柠檬云的批次状态表")
    head = [str(sh.cell_value(hr, c)).strip() for c in range(sh.ncols)]

    def find_col(cands, exclude=None):
        for i, h in enumerate(head):
            if exclude and i in exclude:
                continue
            for k in cands:
                if k in h:
                    return i
        return -1

    c_name = find_col(["商品名称", "物料名称", "存货名称", "名称"])
    c_code = find_col(["商品编码", "物料编码", "存货编码", "编码"])
    c_batch = find_col(["批号", "批次号", "批次"])
    used = {c_name, c_batch}
    c_qty = find_col(["结存数量", "库存数量", "结存", "数量", "余额"], exclude=used)
    c_type = find_col(["业务类型", "单据类型", "类型"])
    c_mfg = find_col(["生产日期", "制造日期", "生产日"])
    c_exp = find_col(["到期日期", "有效期至", "失效日期", "保质期至", "到期日"])
    c_doc_date = find_col(["单据日期"])
    c_in_date = find_col(["入库日期", "收货日期", "入库时间"])
    if c_name < 0 or c_batch < 0:
        raise ValueError("表头缺少「商品名称」或「批号」列（找到的表头：%s）" % " | ".join([h for h in head if h][:12]))

    # 汇总：物料 → 批次 → 数量
    agg = {}

    def num(v):
        try:
            return round(float(v or 0), 3)
        except Exception:
            return 0.0

    for r in range(hr + 1, sh.nrows):
        name = str(sh.cell_value(r, c_name)).strip() if c_name >= 0 else ""
        code = str(sh.cell_value(r, c_code)).strip() if c_code >= 0 else ""
        bno = str(sh.cell_value(r, c_batch)).strip() if c_batch >= 0 else ""
        if not name or not bno or name in ("商品名称", "物料名称", "合计", "小计"):
            continue
        if bno in ("批号", "批次号", "批次"):
            continue
        qty = num(sh.cell_value(r, c_qty)) if c_qty >= 0 else 0.0
        # 批次「跟踪表」是流水账：靠业务类型判定进出方向，否则出库会被加进去
        if c_type >= 0:
            btype = str(sh.cell_value(r, c_type)).strip()
            if any(k in btype for k in ("出库", "领用", "领料", "盘亏", "退货")):
                qty = -abs(qty)
            elif any(k in btype for k in ("入库", "盘盈", "退库")):
                qty = abs(qty)
            else:
                btype = ""
        else:
            btype = ""
        mfg = _cell_date(sh.cell_value(r, c_mfg)) if c_mfg >= 0 else ""
        exp = _cell_date(sh.cell_value(r, c_exp)) if c_exp >= 0 else ""
        in_date = ""
        if btype and qty > 0:
            date_column = c_in_date if c_in_date >= 0 else c_doc_date
            if date_column >= 0:
                in_date = _cell_date(sh.cell_value(r, date_column))
        # 批号可能是数字（如 20260225.0）
        if bno.endswith(".0"):
            bno = bno[:-2]
        d = agg.setdefault(name, {})
        if bno in d:
            d[bno]["qty"] += qty
            if mfg and not d[bno].get("mfg"):
                d[bno]["mfg"] = mfg
            if exp and not d[bno].get("exp"):
                d[bno]["exp"] = exp
            if in_date and (not d[bno].get("in_date") or in_date < d[bno]["in_date"]):
                d[bno]["in_date"] = in_date
        else:
            d[bno] = {"batch": bno, "qty": qty, "mfg": mfg, "exp": exp,
                      "in_date": in_date, "code": code}

    data = load_json(STOCK, {})
    items = data.get("items", [])
    by_name = {i["name"]: i for i in items}
    by_short = {}
    by_code = {}
    duplicate_codes = set()
    for i in items:
        by_short.setdefault(i.get("short", i["name"]), i)
        by_short.setdefault(i["name"], i)
        code = str(i.get("code", "")).strip()
        if code:
            if code in by_code:
                duplicate_codes.add(code)
            else:
                by_code[code] = i
    for code in duplicate_codes:
        by_code.pop(code, None)

    hit, nb, miss = 0, 0, []
    for name, bmap in agg.items():
        target = by_name.get(name)
        if target is None:
            codes = {str(batch.get("code", "")).strip() for batch in bmap.values()}
            codes.discard("")
            if len(codes) == 1:
                target = by_code.get(next(iter(codes)))
        if target is None:
            sn = short_name(name)
            target = by_short.get(sn) or by_short.get(name)
            if target is None:
                for i in items:
                    s = i.get("short", i["name"])
                    if s and (s in name or name in s):
                        target = i
                        break
        if target is None:
            miss.append(name)
            continue
        # 跟踪表只有本期流水：原材料常常「只有出库没有入库」，净额为负不等于没库存。
        # 所以批号一律收下，数量不可信时记 0，后面由 fix_batch_qty 按账面校正。
        fresh = [{"batch": b["batch"], "qty": round(b["qty"], 3) if b.get("qty", 0) > 0 else 0.0,
                  "mfg": b.get("mfg", ""), "exp": b.get("exp", ""),
                  "in_date": b.get("in_date", "")} for b in bmap.values()]
        # 与已有批次合并（同批号以新数据为准），绝不整批清空
        merged = {}
        for b in (target.get("batches") or []):
            merged[b.get("batch", "")] = b
        for b in fresh:
            if not b["batch"]:
                continue
            oldb = merged.get(b["batch"])
            if oldb:
                # 新数据数量未知（跟踪表只有出库）时，沿用已有的数量与日期，别把数抹掉
                if b["qty"] <= 0 and float(oldb.get("qty", 0)) > 0:
                    b["qty"] = float(oldb.get("qty", 0))
                if not b.get("mfg") and oldb.get("mfg"):
                    b["mfg"] = oldb["mfg"]
                if not b.get("exp") and oldb.get("exp"):
                    b["exp"] = oldb["exp"]
                if not b.get("in_date") and oldb.get("in_date"):
                    b["in_date"] = oldb["in_date"]
            merged[b["batch"]] = b
        blist = sorted(merged.values(), key=batch_fifo_key)
        target["batches"] = blist
        hit += 1
        nb += len(blist)

    meta = data.get("_meta", {})
    meta["batch_source"] = path
    meta["batch_imported_at"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    meta["batch_items"] = hit
    meta["batch_count"] = nb
    data["_meta"] = meta
    save_json(STOCK, data)
    return hit, nb, miss


def default_download_dirs():
    """常见的下载目录（换电脑也能找对）：各盘根目录的 Downloads + 当前用户的 Downloads"""
    dirs = [os.path.join(os.path.expanduser("~"), "Downloads")]
    for drv in ("D:\\", "E:\\", "C:\\"):
        d = os.path.join(drv, "Downloads")
        if os.path.isdir(d):
            dirs.append(d)
    # 去重保持顺序
    out = []
    for d in dirs:
        if d not in out:
            out.append(d)
    return out


def find_latest_batch_file():
    """自动寻找最新的批次表：inbox + 常见下载目录，文件名含「批次」"""
    dirs = [os.path.join(APP, "inbox")] + default_download_dirs()
    cands = []
    for d in dirs:
        if not os.path.isdir(d):
            continue
        try:
            for f in os.listdir(d):
                if not f.lower().endswith((".xls", ".xlsx")):
                    continue
                if "批次" not in f:
                    continue
                cands.append(os.path.join(d, f))
        except Exception:
            pass
    return max(cands, key=os.path.getmtime) if cands else None


# ================= 批次 =================
def batch_total(item):
    return round(sum(b.get("qty", 0) for b in item.get("batches", [])), 3)


def has_batches(item):
    return bool(item.get("batches"))


def batch_text(item, maxn=3):
    bs = item.get("batches", [])
    if not bs:
        return ""
    parts = [format_batch(b) for b in sorted(bs, key=batch_fifo_key)[:maxn]]
    if len(bs) > maxn:
        parts.append("…共%d批" % len(bs))
    return " + ".join(parts)


def batch_fifo_key(batch):
    """Oldest recorded receipt first; use manufacture/expiry only for legacy lots."""
    for field in ("in_date", "mfg", "exp"):
        value = str(batch.get(field, "") or "").strip()
        if value:
            return (value[:10], batch.get("batch", ""))
    return ("9999-12-31", batch.get("batch", ""))


def format_batch(batch):
    if batch.get("in_date"):
        date_note = " 入库%s" % batch["in_date"]
    elif batch.get("mfg"):
        date_note = " 生产%s" % batch["mfg"]
    else:
        date_note = ""
    return "%s(%g%s)" % (batch.get("batch", "?"), batch.get("qty", 0), date_note)


def batches_at_book_balance(item):
    """Return estimated lot quantities scaled to the authoritative stock balance."""
    batches = sorted(item.get("batches", []), key=batch_fifo_key)
    if not batches or item.get("qty") is None:
        return batches, False
    book_qty = max(0.0, float(item.get("qty", 0) or 0))
    quantities = [max(0.0, float(batch.get("qty", 0) or 0)) for batch in batches]
    total = sum(quantities)
    if total <= 0 or abs(total - book_qty) <= max(0.01, abs(book_qty) * 0.001):
        return batches, False
    ratio = book_qty / total
    adjusted = []
    for batch, quantity in zip(batches, quantities):
        scaled = dict(batch)
        scaled["qty"] = round(quantity * ratio, 3)
        adjusted.append(scaled)
    residual = round(book_qty - sum(batch["qty"] for batch in adjusted), 3)
    if adjusted and abs(residual) > 0.0005:
        adjusted[-1]["qty"] = round(adjusted[-1]["qty"] + residual, 3)
    return adjusted, True


def batch_summary(item):
    batches, adjusted = batches_at_book_balance(item)
    summary = " + ".join(format_batch(b) for b in batches)
    if adjusted:
        summary = "按账面比例校准（批次量估算）；" + summary
    return summary


def allocate(item, need, occupied=0, book_qty=None):
    """扣除预占后按账面校准批次量及 FIFO 顺序分配投料。"""
    bs = item.get("batches", [])
    if not bs:
        return [("", round(need, 3))] if need > 0 else []
    ordered = sorted(bs, key=batch_fifo_key)
    out, left = [], need
    reserved = max(0.0, float(occupied or 0))
    if book_qty is not None:
        ordered, _ = batches_at_book_balance(
            {"batches": ordered, "qty": book_qty})
    quantities = [max(0.0, float(b.get("qty", 0))) for b in ordered]
    for b, quantity in zip(ordered, quantities):
        if left <= 0 and reserved <= 0:
            break
        reserve_from_batch = min(quantity, reserved)
        reserved -= reserve_from_batch
        available = quantity - reserve_from_batch
        take = min(left, available)
        if take > 0:
            out.append((b.get("batch", ""), round(take, 3)))
            left -= take
    if left > 0:
        out.append(("缺 %.3f" % left, 0))
    return out


def batch_menu():
    data = load_json(STOCK, {})
    items = [i for i in data.get("items", []) if not i["cat"].startswith("产成品")]
    clear()
    bar("批次管理")
    print()
    for idx, it in enumerate(items, 1):
        tag = ""
        if has_batches(it):
            gap = it["qty"] - batch_total(it)
            note = "" if abs(gap) < 0.01 else (" %s(与账面差%+.2f)%s" % (C["yel"], gap, C["r"]))
            tag = "%s%d批%s%s" % (C["grn"], len(it["batches"]), C["r"], note)
        else:
            tag = "%s无批次%s" % (C["gry"], C["r"])
        print("  %s%2d%s  %-18s %10s %-3s  %s" %
              (C["dim"], idx, C["r"], it["short"], it["qty"], it["unit"], tag))
    print()
    print("  %s输入序号录入/修改批次，直接回车返回%s" % (C["dim"], C["r"]))
    s = input("  > ").strip()
    if not s:
        return
    try:
        it = items[int(s) - 1]
    except Exception:
        return
    edit_batches(it)


def edit_batches(item):
    while True:
        clear()
        bar("%s 批次明细" % item["short"])
        print()
        print("  账面库存：%s%s %s%s" % (C["b"], item["qty"], item["unit"], C["r"]))
        bs = item.get("batches", [])
        if bs:
            print()
            print("  %s%-4s %-16s %10s  %-12s %-12s%s" %
                  (C["dim"], "#", "批号", "数量", "生产日期", "有效期至", C["r"]))
            for i, b in enumerate(bs, 1):
                print("  %-4d %-16s %10s  %-12s %-12s" %
                      (i, b.get("batch", ""), b.get("qty", 0),
                       b.get("mfg", ""), b.get("exp", "")))
            print("  %s合计：%.3f %s%s" % (C["dim"], batch_total(item), item["unit"], C["r"]))
        else:
            print("  %s（尚未录入批次）%s" % (C["gry"], C["r"]))
        print()
        print("  1 添加批次   2 清空重录   0 返回")
        c = input("  请选择：").strip()
        if c == "1":
            no = input("  批号：").strip()
            if not no:
                continue
            try:
                q = float(input("  数量（%s）：" % item["unit"]).strip())
            except Exception:
                continue
            mfg = input("  生产日期（如 2026-05-25，可空）：").strip()
            exp = input("  有效期至（可空，用于先进先出）：").strip()
            item.setdefault("batches", []).append(
                {"batch": no, "qty": q, "mfg": mfg, "exp": exp})
            data = load_json(STOCK, {})
            for x in data.get("items", []):
                if x["name"] == item["name"]:
                    x["batches"] = item["batches"]
            save_json(STOCK, data)
            print("  %s已保存%s" % (C["grn"], C["r"]))
        elif c == "2":
            if input("  确认清空该物料全部批次？y/n：").strip().lower() == "y":
                item["batches"] = []
                data = load_json(STOCK, {})
                for x in data.get("items", []):
                    if x["name"] == item["name"]:
                        x["batches"] = []
                save_json(STOCK, data)
        elif c == "0":
            return


# ================= 库存比对（以柠檬云为准） =================
def latest_inbox():
    """从 inbox 目录取最新的库存余额表"""
    inbox = os.path.join(APP, "inbox")
    if not os.path.isdir(inbox):
        return None
    cands = [os.path.join(inbox, f) for f in os.listdir(inbox)
             if f.lower().endswith((".xls", ".xlsx"))]
    return max(cands, key=os.path.getmtime) if cands else None


def find_latest_stock_file():
    """自动寻找最新的库存表：inbox + 常见下载目录，
    文件名含「库存」或「余额」，取修改时间最新的"""
    dirs = [os.path.join(APP, "inbox")] + default_download_dirs()
    cands = []
    for d in dirs:
        if not os.path.isdir(d):
            continue
        try:
            for f in os.listdir(d):
                if not f.lower().endswith((".xls", ".xlsx")):
                    continue
                if ("库存" not in f and "余额" not in f):
                    continue
                cands.append(os.path.join(d, f))
        except Exception:
            pass
    return max(cands, key=os.path.getmtime) if cands else None


def compare_stock(new_path, save_report=True):
    old_data = load_json(STOCK, {})
    old_items = {i["name"]: i for i in old_data.get("items", [])}
    # Validate the ledger before import_xls writes the new authoritative balance.
    exp = {}
    ofile = os.path.join(APP, "orders.json")
    if os.path.exists(ofile):
        with open(ofile, encoding="utf-8") as stream:
            od = json.load(stream)
        if isinstance(od, list):
            ledger_orders = od
        elif isinstance(od, dict):
            ledger_orders = od.get("orders", [])
        else:
            raise ValueError("订单台账格式无效，未导入库存")
        if not isinstance(ledger_orders, list):
            raise ValueError("订单台账的 orders 必须是列表，未导入库存")
        for order in ledger_orders:
            if not isinstance(order, dict):
                raise ValueError("订单台账含无效记录，未导入库存")
            if order.get("status") not in ("已完成", "已发货"):
                continue
            materials = order.get("materials", [])
            if not isinstance(materials, list):
                raise ValueError("已完成订单物料清单格式无效，未导入库存")
            for material in materials:
                if not isinstance(material, dict) or not material.get("name"):
                    raise ValueError("已完成订单含无效物料记录，未导入库存")
                try:
                    quantity = float(material.get("kg", 0) or 0)
                except (TypeError, ValueError):
                    raise ValueError("订单 %s 的物料应耗量无效，未导入库存" %
                                     order.get("serial", ""))
                if not math.isfinite(quantity) or quantity < 0:
                    raise ValueError("订单 %s 的物料应耗量无效，未导入库存" %
                                     order.get("serial", ""))
                name = material["name"]
                exp[name] = exp.get(name, 0.0) + quantity

    n, company, sd = import_xls(new_path)
    new_data = load_json(STOCK, {})
    new_items = {i["name"]: i for i in new_data.get("items", [])}

    rows = []
    for name in sorted(set(old_items) | set(new_items)):
        o, nw = old_items.get(name), new_items.get(name)
        base = nw or o
        old_q = o["qty"] if o else 0.0
        new_q = nw["qty"] if nw else 0.0
        rows.append({"name": name, "short": short_name(name),
                     "unit": base.get("unit", ""), "cat": base.get("cat", ""),
                     "old": old_q, "new": new_q, "delta": round(old_q - new_q, 3),
                     "is_new": o is None, "is_gone": nw is None})

    rev = {}
    for k, v in load_mapping().items():
        for vv in (v if isinstance(v, list) else [v]):
            if vv:
                rev[vv] = k

    diffs = []
    for r in rows:
        if not r["cat"].startswith("原材料"):
            continue
        e = exp.get(rev.get(r["name"], r["short"]), 0.0)
        r["expected"] = round(e, 3)
        r["diff"] = round(r["delta"] - e, 3)
        if abs(r["diff"]) > 0.05:
            diffs.append(r)
    diffs.sort(key=lambda x: -abs(x["diff"]))

    report = None
    if save_report:
        cdir = os.path.join(APP, "data", "compare")
        os.makedirs(cdir, exist_ok=True)
        today = datetime.datetime.now().strftime("%Y%m%d_%H%M")
        report = os.path.join(cdir, "库存比对_%s.md" % today)
        L = ["# 库存比对报告 %s" % datetime.datetime.now().strftime("%Y-%m-%d %H:%M"), ""]
        L.append("> 新表：%s　库存日期 %s　%s" % (os.path.basename(new_path), sd, company))
        L.append("> 原则：以柠檬云为准，系统账面已更新为本次导入值")
        L.append("")
        L.append("## 一、变动汇总（消耗为正 / 入库为负）")
        L.append("")
        moved = [r for r in rows if abs(r["delta"]) > 0.005 or r["is_new"] or r["is_gone"]]
        L.append("| 物料 | 上次 | 本次 | 变化 | 单位 |")
        L.append("|---|---|---|---|---|")
        for r in moved:
            note = "（新档案）" if r["is_new"] else ("（库存表已无此项）" if r["is_gone"] else "")
            L.append("| %s | %s | %s | %s | %s %s |" %
                     (r["short"], r["old"], r["new"], r["delta"], r["unit"], note))
        if not moved:
            L.append("无变动")
        L.append("")
        L.append("## 二、误差分析（已完成订单应消耗 vs 实际账面变化）")
        L.append("")
        if diffs:
            L.append("以下物料的账面消耗与订单应消耗差异超过 0.05，请核对：")
            L.append("")
            L.append("| 物料 | 账面变化 | 订单应消耗 | 误差 |")
            L.append("|---|---|---|---|")
            for r in diffs:
                L.append("| %s | %s | %s | %s |" % (r["short"], r["delta"], r["expected"], r["diff"]))
        else:
            L.append("各项原料的账面变化与订单应消耗一致，无异常误差。")
        L.append("")
        L.append("> 误差为参考值：试验投料、退料、盘盈盘亏都会造成差异，")
        L.append("> 超过 0.05kg 的建议在柠檬云里查该物料的出入库明细确认原因。")
        with open(report, "w", encoding="utf-8") as f:
            f.write("\n".join(L))
    return rows, diffs, report, n, company, sd


# ================= 状态与统计 =================
def batch_usage():
    usage = {}
    for k, r in load_json(RECIPE, {}).items():
        if k.startswith("_") or not isinstance(r, dict):
            continue
        for m in r.get("materials", []):
            u = BATCH_KG * m["ratio"] / 100.0
            usage[m["name"]] = max(usage.get(m["name"], 0.0), u)
    return usage


def threshold_map():
    usage = batch_usage()
    rev = {}
    for k, v in load_mapping().items():
        for vv in (v if isinstance(v, list) else [v]):
            if vv:
                rev[vv] = k
    th = {}
    for it in load_json(STOCK, {}).get("items", []):
        if it["cat"].startswith("产成品"):
            continue
        u = usage.get(rev.get(it["name"], it["short"]), 0.0)
        th[it["name"]] = max(it.get("safe", 0) or 0, u * 2)
    return th


def status_of(it, th):
    """it.qty = 扣除在途订单预占后的可用量；it.book = 柠檬云账面。
    断料 = 平台真没货；超订 = 库里有货但全被待生产订单预订（可能还超订）"""
    if it["cat"].startswith("产成品"):
        return "成品"
    book = it.get("book", it["qty"])
    qty = it["qty"]
    if book <= 0:
        return "断料"
    if qty <= 0:
        return "超订"
    if th > 0 and qty < th:
        return "不足"
    if th > 0 and qty < th * 1.5:
        return "偏低"
    return "正常"


def color_of(st):
    return {"断料": C["red"], "不足": C["red"], "偏低": C["yel"],
            "正常": C["grn"], "成品": C["cyn"]}.get(st, "")


def occupied_map():
    """「待生产」订单预占的物料 kg，按配方名汇总。
    「生产中」= 原料已领出（柠檬云账面已扣减），不再重复占用。"""
    ofile = os.path.join(APP, "orders.json")
    occ = {}
    if not os.path.exists(ofile):
        return occ
    with open(ofile, encoding="utf-8") as stream:
        od = json.load(stream)
    if isinstance(od, list):
        orders = od
    elif isinstance(od, dict):
        orders = od.get("orders", [])
    else:
        raise ValueError("订单台账格式无效，不能安全计算预占库存")
    if not isinstance(orders, list):
        raise ValueError("订单台账格式无效，不能安全计算预占库存")
    for order in orders:
        if not isinstance(order, dict):
            raise ValueError("订单台账含无效订单记录，不能安全计算预占库存")
        if order.get("status") != "待生产" or order.get("archived"):
            continue
        materials = order.get("materials", [])
        if not isinstance(materials, list):
            raise ValueError("待生产订单的物料清单格式无效")
        for material in materials:
            if not isinstance(material, dict) or not material.get("name"):
                raise ValueError("待生产订单含无效物料记录，不能安全计算预占库存")
            quantity = float(material.get("kg", 0) or 0)
            if not math.isfinite(quantity) or quantity < 0:
                raise ValueError("待生产订单物料「%s」的预占量无效" % material["name"])
            occ[material["name"]] = occ.get(material["name"], 0.0) + quantity
    return occ


def build_report():
    data = load_json(STOCK, {})
    th = threshold_map()
    occ = occupied_map()
    rev = {}
    for k, v in load_mapping().items():
        for vv in (v if isinstance(v, list) else [v]):
            if vv:
                rev[vv] = k
    items = [i for i in data.get("items", []) if not i["cat"].startswith("产成品")]
    items.sort(key=lambda x: (x["cat"], x["name"]))

    # 按"配方名（或自身）"分组：多映射的库存项合并为一个可用池
    groups, order = {}, []
    for it in items:
        key = rev.get(it["name"], it["short"])
        if key not in groups:
            groups[key] = []
            order.append(key)
        groups[key].append(it)

    rows = []
    for key in order:
        pool = groups[key]
        oc = occ.get(key, 0.0)
        qty = sum(i["qty"] for i in pool)
        avail = qty - oc
        t = max(th.get(i["name"], 0) for i in pool)
        base = dict(pool[0])
        base["short"] = key
        base["qty"] = round(qty, 3)
        base["batches"] = [b for i in pool for b in i.get("batches", [])]
        base["status"] = status_of({"qty": avail, "book": qty, "cat": base["cat"]}, t)
        base["line"] = round(t, 3)
        base["occupied"] = round(oc, 3)
        base["avail"] = round(avail, 3)
        base["parts"] = [i["name"] for i in pool]
        rows.append(base)
    alert = [r for r in rows if r["status"] in ("断料", "不足")]
    warn = [r for r in rows if r["status"] == "偏低"]
    return data.get("_meta", {}), rows, alert, warn


def load_mapping():
    return {k: v for k, v in load_json(MAPPING, {}).items() if not k.startswith("_")}


# ================= 控制台播报 =================
def print_report():
    meta, rows, alert, warn = build_report()
    clear()
    bar("库存播报  %s" % datetime.datetime.now().strftime("%Y-%m-%d %H:%M"))
    print("  %s数据基准 %s　%s　共 %d 项原料%s" %
          (C["dim"], meta.get("stock_date", "-"), meta.get("company", ""),
           len(rows), C["r"]))
    print()
    print("  %s● 需立即采购（%d）%s" % (C["red"] + C["b"], len(alert), C["r"]))
    if alert:
        print("  %s%-16s %9s %-3s %9s %8s%s" %
              (C["dim"], "物料", "现有", "单位", "警戒", "可做批", C["r"]))
        for r in alert:
            b = ("%.1f" % (r["qty"] / (r["line"] / 2))) if r["line"] > 0 else "-"
            extra = ""
            if has_batches(r):
                extra = "  %s%s%s" % (C["gry"], batch_text(r), C["r"])
            print("  %s%-16s %9s %-3s %9s %8s%s%s" %
                  (C["red"], r["short"], r["qty"], r["unit"], r["line"], b, C["r"], extra))
    else:
        print("  %s无%s" % (C["gry"], C["r"]))
    print()
    print("  %s● 库存偏低（%d）%s" % (C["yel"] + C["b"], len(warn), C["r"]))
    if warn:
        for r in warn:
            b = ("%.1f" % (r["qty"] / (r["line"] / 2))) if r["line"] > 0 else "-"
            print("  %s%-16s %9s %-3s %9s %8s%s" %
                  (C["yel"], r["short"], r["qty"], r["unit"], r["line"], b, C["r"]))
    else:
        print("  %s无%s" % (C["gry"], C["r"]))
    print()
    print("  %s● 全部原料（%d）　占用 = 在途订单已预扣%s" % (C["grn"] + C["b"], len(rows), C["r"]))
    cur = None
    for r in rows:
        cat = r["cat"].replace("原材料-", "")
        if cat != cur:
            print("  %s┌ %s%s" % (C["blu"], cat, C["r"]))
            cur = cat
        bt = batch_text(r)
        occ = (" %s占用%.2f→可用%.2f%s" % (C["yel"], r["occupied"], r["avail"], C["r"])) \
            if r["occupied"] > 0 else ""
        print("  %s│%s %-16s %9s %-3s %s%s%s%s%s" %
              (C["blu"], C["r"], r["short"], r["qty"], r["unit"],
               color_of(r["status"]), r["status"].ljust(4), C["r"],
               ("  " + C["gry"] + bt + C["r"]) if bt else "", occ))
    print()


# ================= 采购建议 =================
def advice_flow():
    recipes = load_json(RECIPE, {})
    names = [k for k in recipes if not k.startswith("_")]
    clear()
    bar("按生产计划算采购建议")
    print()
    for i, k in enumerate(names, 1):
        print("  %s%2d%s  %s" % (C["dim"], i, C["r"], k))
    print()
    plans = []
    while True:
        s = input("  产品序号（回车结束）：").strip()
        if not s:
            break
        try:
            product = names[int(s) - 1]
        except Exception:
            print("  %s序号无效%s" % (C["red"], C["r"]))
            continue
        wan = input("  计划产量（万片）：").strip()
        try:
            wan = float(wan)
        except Exception:
            print("  %s产量无效%s" % (C["red"], C["r"]))
            continue
        plans.append({"product": product, "wan": wan})
        print("  %s已加入：%s %g 万片%s" % (C["grn"], product, wan, C["r"]))

    if not plans:
        return

    result, unknown = purchase_advice(plans)
    print()
    bar("采购建议")
    print()
    print("  %s%-14s %9s %9s %9s %9s%s" %
          (C["dim"], "物料", "本批需求", "现有", "缺口", "建议采购", C["r"]))
    print("  " + "─" * 62)
    for r in result:
        col = C["red"] if not r["enough"] else C["grn"]
        print("  %s%-14s %9s %9s %9s %9s%s" %
              (col, r["name"], r["need"], r["stock"],
               r["gap"] if not r["enough"] else 0,
               r["gap"] if not r["enough"] else 0, C["r"]))
    print("  " + "─" * 62)
    if unknown:
        print()
        print("  %s注意：以下物料库存表无建档，请先建档：%s" % (C["yel"], C["r"]))
        for u in unknown:
            print("    · %s" % u)
    print()

    # 批次分配明细（合并映射池，先进先出）
    batch_lines = []
    for r in result:
        allb = []
        for it in r.get("pool", []):
            allb.extend(it.get("batches", []))
        if allb and r["need"] > 0:
            alloc = allocate({"batches": allb}, r["need"])
            if alloc:
                batch_lines.append((r["name"], alloc))
    if batch_lines:
        print("  %s● 投料批次分配（先进先出）%s" % (C["cyn"] + C["b"], C["r"]))
        for name, alloc in batch_lines:
            seg = "".join("%s %s" % (b, q) for b, q in alloc)
            print("    %-14s %s" % (name, seg))
        print()

    if input("  是否保存这份采购建议？y/n：").strip().lower() == "y":
        os.makedirs(os.path.join(APP, "data", "purchase"), exist_ok=True)
        p = os.path.join(APP, "data", "purchase",
                         "采购建议_%s.md" % datetime.datetime.now().strftime("%Y%m%d_%H%M"))
        with open(p, "w", encoding="utf-8") as f:
            f.write("# 采购建议 %s\n\n" % datetime.datetime.now().strftime("%Y-%m-%d %H:%M"))
            f.write("## 生产计划\n\n")
            for pl in plans:
                f.write("- %s  %g 万片\n" % (pl["product"], pl["wan"]))
            f.write("\n## 需求与缺口\n\n| 物料 | 本批需求 | 现有 | 缺口 | 单位 |\n|---|---|---|---|---|\n")
            for r in result:
                f.write("| %s | %s | %s | %s | %s |\n" %
                        (r["name"], r["need"], r["stock"],
                         r["gap"] if not r["enough"] else 0, r["unit"]))
            if batch_lines:
                f.write("\n## 投料批次分配\n\n")
                for name, alloc in batch_lines:
                    f.write("- %s：%s\n" % (name, "，".join("%s 取 %s" % (b, q) for b, q in alloc)))
        print("  %s已保存：%s%s" % (C["grn"], p, C["r"]))
        os.startfile(p)


def purchase_advice(plans):
    recipes = load_json(RECIPE, {})
    data = load_json(STOCK, {})
    mapping = load_mapping()
    by_name = {i["name"]: i for i in data.get("items", [])}
    by_short = {}
    for i in data.get("items", []):
        by_short.setdefault(i["short"], i)

    need = {}
    for p in plans:
        r = recipes.get(p["product"])
        if not r:
            continue
        total_kg = float(p["wan"]) * 10000 * float(r.get("tablet_g", 0.5)) / 1000
        for m in r["materials"]:
            need[m["name"]] = need.get(m["name"], 0) + total_kg * m["ratio"] / 100

    result, unknown = [], set()
    for name, qty in need.items():
        m = mapping.get(name)
        pool = []
        if m is None:
            it = by_short.get(name)
            if it:
                pool = [it]
        else:
            for f in (m if isinstance(m, list) else [m]):
                it = by_name.get(f)
                if it:
                    pool.append(it)
        if not pool:
            unknown.add(name)
            result.append({"name": name, "stock_name": "", "pool": [], "unit": "KG",
                           "need": round(qty, 3), "stock": 0.0, "safe": 0.0,
                           "gap": round(qty, 3), "enough": False})
            continue
        total = sum(i["qty"] for i in pool)
        safe = max(i.get("safe", 0) or 0 for i in pool)
        gap = qty + safe - total
        result.append({"name": name,
                       "stock_name": "、".join(i["name"] for i in pool),
                       "pool": pool, "unit": pool[0].get("unit", "KG"),
                       "need": round(qty, 3), "stock": round(total, 3),
                       "gap": round(max(0.0, gap), 3), "enough": gap <= 0})
    result.sort(key=lambda x: (x["enough"], -x["gap"]))
    return result, sorted(unknown)


# ================= 搜索 =================
def search_flow():
    while True:
        kw = input("  关键字（物料名 / 规格 / 类别，回车返回）：").strip()
        if not kw:
            return
        _run_search(kw)


def _run_search(kw):
    k = kw.lower()
    meta, rows, alert, warn = build_report()
    hits = [r for r in rows
            if k in r["short"].lower() or k in r["name"].lower()
            or k in (r.get("spec") or "").lower() or k in r["cat"].lower()]
    fg = [i for i in load_json(STOCK, {}).get("items", [])
          if i["cat"].startswith("产成品")
          and k in (i["name"] + " " + i.get("spec", "")).lower()]
    print()
    if hits:
        print("  %s原料命中 %d 项：%s" % (C["cyn"] + C["b"], len(hits), C["r"]))
        print("  %s%-18s %10s %-3s %9s %-4s%s" %
              (C["dim"], "物料", "现有", "单位", "警戒", "状态", C["r"]))
        for r in hits:
            extra = ""
            if has_batches(r):
                extra = "  " + C["gry"] + batch_text(r) + C["r"]
            elif r.get("occupied"):
                extra = "  " + C["gry"] + "占用%.2f→可用%.2f" % (r["occupied"], r["avail"]) + C["r"]
            print("  %s%-18s %10s %-3s %9s %s%s%s%s" %
                  (color_of(r["status"]), r["short"], r["qty"], r["unit"],
                   r["line"], color_of(r["status"]), r["status"], C["r"], extra))
    else:
        print("  原料无匹配")
    if fg:
        print("  %s成品命中 %d 项：%s" % (C["cyn"] + C["b"], len(fg), C["r"]))
        for i in fg:
            print("  %-26s %10s %s  [%s]" %
                  (i["name"], i["qty"], i["unit"], i.get("spec", "")[:30]))
    if not hits and not fg:
        print("  %s没有任何匹配%s" % (C["gry"], C["r"]))
    print()


# ================= 日报 =================
def generate_daily(date8=None):
    date8 = date8 or datetime.datetime.now().strftime("%Y%m%d")
    os.makedirs(DAILY_DIR, exist_ok=True)
    meta, rows, alert, warn = build_report()

    orders = []
    day_dir = os.path.join(os.path.expanduser("~"), "Desktop",
                           "%d月" % int(date8[4:6]), date8)
    if os.path.isdir(day_dir):
        orders = sorted(os.listdir(day_dir))

    L = []
    L.append("# %s 跟单日报" % date8)
    L.append("")
    L.append("> 生成时间 %s　库存基准 %s　%s" %
             (datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
              meta.get("stock_date", "-"), meta.get("company", "")))
    L.append("")
    L.append("## 一、当日生产指令单（%d 单）" % len(orders))
    L.append("")
    if orders:
        L.append("| 序号 | 单号+产品 |")
        L.append("|---|---|")
        for i, o in enumerate(orders, 1):
            L.append("| %d | %s |" % (i, o))
    else:
        L.append("（今日暂无）")
    L.append("")

    # 订单交期（台账联动）
    late, soon = [], []
    ofile = os.path.join(APP, "orders.json")
    if os.path.exists(ofile):
        try:
            od = json.load(open(ofile, encoding="utf-8"))
            today = datetime.date.today()
            for o in od.get("orders", []):
                if o.get("status") in ("已完成", "已发货"):
                    continue
                due = (o.get("due") or "").strip()
                if not due:
                    continue
                try:
                    dd = (datetime.datetime.strptime(due, "%Y%m%d").date() - today).days
                except Exception:
                    continue
                if dd < 0:
                    late.append((o, dd))
                elif dd <= 3:
                    soon.append((o, dd))
        except Exception:
            pass
    L.append("## 二、订单交期（逾期 %d，临期 %d）" % (len(late), len(soon)))
    L.append("")
    if late or soon:
        if late:
            L.append("**逾期（必须当日处理）：**")
            L.append("")
            for o, dd in late:
                L.append("- %s %s　逾期 %d 天　%s" %
                         (o.get("serial"), o.get("product"), -dd, o.get("status")))
        if soon:
            L.append("**3 天内到期：**")
            L.append("")
            for o, dd in soon:
                L.append("- %s %s　剩 %d 天　%s" %
                         (o.get("serial"), o.get("product"), dd, o.get("status")))
    else:
        L.append("无逾期、无临期订单")
    L.append("")
    L.append("## 三、库存警报（%d 项）" % len(alert))
    L.append("")
    if alert:
        L.append("| 物料 | 现有 | 警戒线 | 可做批次 | 单位 | 状态 | 批次构成 |")
        L.append("|---|---|---|---|---|---|---|")
        for r in alert:
            b = ("%.1f" % (r["qty"] / (r["line"] / 2))) if r["line"] > 0 else "-"
            L.append("| %s | %s | %s | %s | %s | **%s** | %s |" %
                     (r["short"], r["qty"], r["line"], b, r["unit"], r["status"],
                      batch_text(r) or "-"))
    else:
        L.append("无")
    L.append("")
    L.append("## 四、库存偏低（%d 项）" % len(warn))
    L.append("")
    if warn:
        L.append("| 物料 | 现有 | 警戒线 | 单位 | 批次构成 |")
        L.append("|---|---|---|---|---|")
        for r in warn:
            L.append("| %s | %s | %s | %s | %s |" %
                     (r["short"], r["qty"], r["line"], r["unit"], batch_text(r) or "-"))
    else:
        L.append("无")
    L.append("")
    L.append("## 五、原料库存全表")
    L.append("")
    L.append("| 物料 | 规格 | 类别 | 现有 | 警戒线 | 单位 | 状态 | 批次 |")
    L.append("|---|---|---|---|---|---|---|---|")
    for r in rows:
        L.append("| %s | %s | %s | %s | %s | %s | %s | %s |" %
                 (r["short"], (r.get("spec") or "")[:22],
                  r["cat"].replace("原材料-", ""), r["qty"], r["line"],
                  r["unit"], r["status"], batch_text(r) or "-"))
    L.append("")
    L.append("## 六、备注")
    L.append("")
    L.append("- 数据来源：柠檬云业财《商品库存余额表》")
    L.append("- 警戒线 = max(手动安全线, 单批用量 × 2)，单批按 2 万片 × 0.5g = 10kg")
    L.append("- 批次按有效期先进先出（FIFO）分配")
    L.append("")

    path = os.path.join(DAILY_DIR, "%s.md" % date8)
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(L))
    return path, alert, warn, len(orders)


# ================= 网页看板 =================
def build_dashboard():
    meta, rows, alert, warn = build_report()
    os.makedirs(DASH_DIR, exist_ok=True)
    total = len(rows)
    ok = len([r for r in rows if r["status"] == "正常"])
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")

    def card(cls, r):
        b = ("%.1f 批" % (r["qty"] / (r["line"] / 2))) if r["line"] > 0 else "-"
        bt = batch_text(r)
        bt_html = ('<div class="bt">批次：%s</div>' % bt) if bt else ""
        return ('<div class="item %s"><div class="nm">%s</div>'
                '<div class="qty">%s <span>%s</span></div>'
                '<div class="meta">警戒 %s　可做 %s</div>%s</div>'
                % (cls, r["short"], r["qty"], r["unit"], r["line"], b, bt_html))

    alert_html = "".join(card("red", r) for r in alert) or '<div class="empty">当前无缺料</div>'
    warn_html = "".join(card("yel", r) for r in warn) or '<div class="empty">当前无偏低物料</div>'

    cats = {}
    for r in rows:
        cats.setdefault(r["cat"].replace("原材料-", ""), []).append(r)
    tab_html = ""
    for cat, lst in cats.items():
        tab_html += '<h3>%s</h3><table><tr><th>物料</th><th>规格</th><th>现有</th>' \
                    '<th>警戒线</th><th>单位</th><th>状态</th><th>批次构成</th></tr>' % cat
        for r in lst:
            cls = {"断料": "s-red", "不足": "s-red", "偏低": "s-yel",
                   "正常": "s-grn"}.get(r["status"], "")
            tab_html += ('<tr><td>%s</td><td class="dim">%s</td><td class="num">%s</td>'
                         '<td class="num">%s</td><td>%s</td><td class="%s">%s</td>'
                         '<td class="dim">%s</td></tr>'
                         % (r["short"], (r.get("spec") or "")[:26], r["qty"],
                            r["line"], r["unit"], cls, r["status"],
                            batch_text(r) or "-"))
        tab_html += "</table>"

    html = """<!DOCTYPE html><html><head><meta charset="utf-8">
<title>库存看板 %s</title><style>
*{box-sizing:border-box}
body{margin:0;padding:26px 30px;background:#f5f7fa;font-family:"Microsoft YaHei",sans-serif;color:#1f2937}
h1{font-size:21px;margin:0 0 4px;color:#12355b}
.sub{color:#6b7280;font-size:13px;margin-bottom:20px}
.cards{display:flex;gap:12px;margin-bottom:22px}
.kpi{flex:1;background:#fff;border-radius:10px;padding:14px 18px;border:1px solid #e5e9f0}
.kpi .n{font-size:26px;font-weight:600;line-height:1.2}
.kpi .l{font-size:12px;color:#6b7280;margin-top:2px}
.kpi.red .n{color:#c0392b}.kpi.yel .n{color:#b8860b}.kpi.grn .n{color:#1e7d43}
h2{font-size:15px;margin:22px 0 10px;padding-left:10px;border-left:4px solid #12355b}
.grid{display:flex;flex-wrap:wrap;gap:10px}
.item{background:#fff;border-radius:8px;padding:11px 14px;min-width:190px;flex:1 1 190px;border:1px solid #e5e9f0}
.item.red{border-left:4px solid #c0392b}
.item.yel{border-left:4px solid #d4a017}
.nm{font-size:13px;color:#374151;font-weight:600}
.qty{font-size:20px;font-weight:600;margin:2px 0}
.qty span{font-size:12px;color:#9ca3af;font-weight:400}
.meta{font-size:12px;color:#8b95a5}
.bt{font-size:12px;color:#7a8699;margin-top:4px;border-top:1px dashed #e5e9f0;padding-top:4px}
.empty{color:#9ca3af;font-size:13px;padding:8px 2px}
table{border-collapse:collapse;width:100%%;background:#fff;border-radius:8px;overflow:hidden;font-size:13px;margin-bottom:16px}
th{background:#eef3f9;color:#12355b;font-weight:600;text-align:left;padding:7px 10px;font-size:12px}
td{padding:6px 10px;border-top:1px solid #eff2f6}
tr:hover td{background:#fafbfd}
.num{text-align:right}
.dim{color:#8b95a5}
.s-red{color:#c0392b;font-weight:600}.s-yel{color:#b8860b;font-weight:600}.s-grn{color:#1e7d43}
.foot{margin-top:26px;padding-top:12px;border-top:1px solid #e5e9f0;color:#9ca3af;font-size:12px}
</style></head><body>
<h1>库存看板</h1>
<div class="sub">数据基准 %s　%s　生成于 %s</div>
<div class="cards">
  <div class="kpi"><div class="n">%d</div><div class="l">原料总数</div></div>
  <div class="kpi red"><div class="n">%d</div><div class="l">需立即采购</div></div>
  <div class="kpi yel"><div class="n">%d</div><div class="l">库存偏低</div></div>
  <div class="kpi grn"><div class="n">%d</div><div class="l">库存正常</div></div>
</div>
<h2>需立即采购</h2><div class="grid">%s</div>
<h2>库存偏低</h2><div class="grid">%s</div>
<h2>全部原料明细</h2>%s
<div class="foot">警戒线 = max(手动安全线, 单批用量×2)，单批按 2 万片×0.5g = 10kg；批次按有效期先进先出。</div>
</body></html>""" % (datetime.datetime.now().strftime("%Y%m%d_%H%M"),
                     meta.get("stock_date", "-"), meta.get("company", ""), now,
                     total, len(alert), len(warn), ok,
                     alert_html, warn_html, tab_html)

    path = os.path.join(DASH_DIR, "库存看板_%s.html" % datetime.datetime.now().strftime("%Y%m%d"))
    with open(path, "w", encoding="utf-8") as f:
        f.write(html)
    return path


# ================= 主入口 =================
def main():
    if len(sys.argv) > 1:
        cmd = sys.argv[1]
        if cmd == "--import" and len(sys.argv) > 2:
            n, company, sd = import_xls(sys.argv[2])
            print("导入完成：%d 项　%s　库存日期 %s" % (n, company, sd))
            return
        if cmd == "--report":
            print_report()
            return
        if cmd == "--daily":
            path, alert, warn, n = generate_daily(sys.argv[2] if len(sys.argv) > 2 else None)
            print("日报：%s（单据 %d，警报 %d，偏低 %d）" % (path, n, len(alert), len(warn)))
            return
        if cmd == "--dashboard":
            print("看板：%s" % build_dashboard())
            return
        if cmd == "--compare":
            p = sys.argv[2] if len(sys.argv) > 2 else latest_inbox()
            if not p or not os.path.exists(p):
                print("未找到库存表（inbox 为空）")
                return
            rows, diffs, report, n, company, sd = compare_stock(p)
            print("导入 %d 项　库存日期 %s　报告：%s" % (n, sd, report))
            print("误差超阈值：%d 项" % len(diffs))
            return

    while True:
        clear()
        bar("库存管理中心  %s" % datetime.datetime.now().strftime("%Y-%m-%d"))
        print()
        print("  %s1%s  库存播报（控制台）" % (C["cyn"], C["r"]))
        print("  %s2%s  采购建议（按生产计划算缺口）" % (C["cyn"], C["r"]))
        print("  %s3%s  库存看板（网页版，更好看）" % (C["cyn"], C["r"]))
        print("  %s4%s  批次管理（录入 / 查看批次）" % (C["cyn"], C["r"]))
        print("  %s5%s  导入库存余额表 + 误差比对（每天用这个）" % (C["cyn"], C["r"]))
        print("  %s6%s  生成今日日报" % (C["cyn"], C["r"]))
        print("  %s7%s  打开收件箱（柠檬云导出的表放这里）" % (C["cyn"], C["r"]))
        print("  %s8%s  搜索库存（按名称 / 规格）" % (C["cyn"], C["r"]))
        print("  %s0%s  退出" % (C["gry"], C["r"]))
        print()
        c = input("  请选择：").strip()
        if c == "1":
            print_report()
            input("\n  按回车返回")
        elif c == "2":
            advice_flow()
            input("\n  按回车返回")
        elif c == "3":
            p = build_dashboard()
            print("  %s看板已生成：%s%s" % (C["grn"], p, C["r"]))
            os.startfile(p)
            input("\n  按回车返回")
        elif c == "4":
            batch_menu()
        elif c == "5":
            p = input("  库存表路径（回车 = 自动取 inbox 最新，可直接拖入）：").strip().strip('"')
            if not p:
                p = latest_inbox()
            if p and os.path.exists(p):
                rows, diffs, report, n, company, sd = compare_stock(p)
                print("  %s导入 %d 项　库存日期 %s%s" % (C["grn"], n, sd, C["r"]))
                print("  比对报告：%s" % report)
                if diffs:
                    print("  %s发现 %d 项误差超阈值：%s" % (C["red"], len(diffs), C["r"]))
                    for r in diffs[:6]:
                        print("    · %-14s 账面变化 %s  应耗 %s  误差 %s" %
                              (r["short"], r["delta"], r["expected"], r["diff"]))
                else:
                    print("  %s账面变化与订单应消耗一致，无异常误差%s" % (C["grn"], C["r"]))
            else:
                print("  %s未找到库存表。可把柠檬云导出的文件放进：%s" % (C["red"], os.path.join(APP, "inbox")))
            input("\n  按回车返回")
        elif c == "6":
            path, alert, warn, n = generate_daily()
            print("  %s日报：%s%s" % (C["grn"], path, C["r"]))
            input("\n  按回车返回")
        elif c == "7":
            inbox = os.path.join(APP, "inbox")
            os.makedirs(inbox, exist_ok=True)
            os.startfile(inbox)
            input("\n  按回车返回")
        elif c == "8":
            search_flow()
            input("\n  按回车返回")
        elif c == "0":
            break


if __name__ == "__main__":
    main()
