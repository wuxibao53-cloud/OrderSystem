# -*- coding: utf-8 -*-
"""跟单助手 HTTP API 服务器（纯 HTTP，不依赖 WebSocket，绕开代理拦截）"""
import os
import sys
import json
import datetime
import threading
import traceback
import time
import hashlib
import ctypes
from ctypes import wintypes

APP = os.path.dirname(os.path.abspath(__file__))   # 便携：系统放哪个目录都能跑
sys.path.insert(0, APP)

from bottle import Bottle, request, response, static_file  # noqa: E402
import app_gui  # noqa: E402  复用全部业务函数（__main__ 已保护，import 不会启动服务）
FN = {
    "ping": app_gui.ping,
    "get_products": app_gui.get_products,
    "suggest_batch": app_gui.suggest_batch,
    "order_preview": app_gui.order_preview,
    "generate_order": app_gui.generate_order,
    "get_stock": app_gui.get_stock,
    "search_stock": app_gui.search_stock,
    "auto_import_stock": app_gui.auto_import_stock,
    "import_batches": app_gui.import_batches,
    "fix_batch_qty": app_gui.fix_batch_qty,
    "batch_reconcile_report": app_gui.batch_reconcile_report,
    "self_check": app_gui.self_check,
    "clear_logs": app_gui.clear_logs,
    "open_doc": app_gui.open_doc,
    "get_replenish": app_gui.get_replenish,
    "import_stock": app_gui.import_stock,
    "save_product": app_gui.save_product,
    "open_inbox": app_gui.open_inbox,
    "open_report": app_gui.open_report,
    "open_path": app_gui.open_path,
    "import_sales_form_upload": app_gui.import_sales_form_upload,
    "ai_status": app_gui.ai_status,
    "ai_preview_sales_form_upload": app_gui.ai_preview_sales_form_upload,
    "ai_extract_order_text": app_gui.ai_extract_order_text,
    "ai_ask": app_gui.ai_ask,
    "ai_progress": app_gui.ai_progress,
    "lemon_sync_run": app_gui.lemon_sync_run,
    "lemon_status": app_gui.lemon_status,
    "purchase_forecast_view": app_gui.purchase_forecast_view,
    "ai_import_text_order": app_gui.ai_import_text_order,
    "ai_daily_briefing": app_gui.ai_daily_briefing,
    "preview_sales_form_upload": app_gui.preview_sales_form_upload,
    "mark_production_run": app_gui.mark_production_run,
    "export_batch_audit": app_gui.export_batch_audit,
    "preview_manual_recipe_upload": app_gui.preview_manual_recipe_upload,
    "report_js_error": app_gui.report_js_error,
    "get_production_dashboard": app_gui.get_production_dashboard,
    "import_production_sheet": app_gui.import_production_sheet,
    "confirm_tracking_order": app_gui.confirm_tracking_order,
    "create_manual_tracking": app_gui.create_manual_tracking,
    "update_production_step": app_gui.update_production_step,
    "update_order_status": app_gui.update_order_status,
    "update_order_info": app_gui.update_order_info,
    "archive_or_void_order": app_gui.archive_or_void_order,
    "delete_unstarted_order": app_gui.delete_unstarted_order,
    "delete_test_order": app_gui.delete_test_order,
    "create_backup": app_gui.create_backup,
    "list_backups": app_gui.list_backups,
    "restore_backup": app_gui.restore_backup,
}

from bottle import BaseRequest  # noqa: E402
BaseRequest.MEMFILE_MAX = 32 * 1024 * 1024   # DOCX/XLSX 以 base64 上传（上限 20MB）
app = Bottle()

# 崩溃原因记录：主线程与工作线程的异常全部落盘
import traceback
CRASH_LOG = os.path.join(APP, "data", "server_crash.log")
_INSTANCE_MUTEX = None


def acquire_single_instance():
    """Keep one OS mutex for the lifetime of the service process."""
    global _INSTANCE_MUTEX
    identity = hashlib.sha256(os.path.normcase(APP).encode("utf-8")).hexdigest()[:20]
    name = "Local\\OrderSystem-%s" % identity
    if os.name == "nt":
        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        kernel32.CreateMutexW.argtypes = [wintypes.LPVOID, wintypes.BOOL, wintypes.LPCWSTR]
        kernel32.CreateMutexW.restype = ctypes.c_void_p
        ctypes.set_last_error(0)
        handle = kernel32.CreateMutexW(None, True, name)
        if not handle:
            raise OSError(ctypes.get_last_error(), "无法创建单实例互斥锁")
        if ctypes.get_last_error() == 183:
            kernel32.CloseHandle(handle)
            return False
        _INSTANCE_MUTEX = handle
        return True
    import fcntl
    lock_path = os.path.join(APP, "data", "server.lock")
    os.makedirs(os.path.dirname(lock_path), exist_ok=True)
    stream = open(lock_path, "a+")
    try:
        fcntl.flock(stream.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
    except OSError:
        stream.close()
        return False
    _INSTANCE_MUTEX = stream
    return True


def _write_crash(text):
    try:
        with open(CRASH_LOG, "a", encoding="utf-8") as f:
            f.write("%s\n%s\n" % (datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), text))
    except Exception:
        pass


def _sys_hook(tp, val, tb):
    _write_crash("".join(traceback.format_exception(tp, val, tb)))
    sys.__excepthook__(tp, val, tb)


def _thread_hook(args):
    _write_crash("线程异常: " + "".join(traceback.format_exception(args.exc_type, args.exc_value, args.exc_traceback)))
    sys.__excepthook__(args.exc_type, args.exc_value, args.exc_traceback)


sys.excepthook = _sys_hook
threading.excepthook = _thread_hook

LOG = os.path.join(APP, "data", "http_log.txt")


def log_req(line):
    try:
        with open(LOG, "a", encoding="utf-8") as f:
            f.write("%s  %s\n" % (datetime.datetime.now().strftime("%H:%M:%S"), line))
    except Exception:
        pass


@app.hook("before_request")
def _log_before():
    log_req("→ %s %s（开始处理）" % (request.method, request.path))


@app.hook("after_request")
def _log_after():
    log_req("← %s %s（完成）" % (request.method, request.path))


@app.hook("after_request")
def no_cache():
    response.set_header("Cache-Control", "no-store, must-revalidate")
    response.set_header("Pragma", "no-cache")


@app.post("/api/<fn>")
def call(fn):
    if fn not in FN:
        response.status = 404
        response.content_type = "application/json"
        return json.dumps({"ok": False, "msg": "未知接口：%s" % fn}, ensure_ascii=False)
    try:
        args = request.json
        if args is None:
            args = []
        if not isinstance(args, list):
            args = [args]
        result = FN[fn](*args)
    except Exception as exc:
        response.status = 500
        _write_crash("API %s 调用失败：\n%s" % (fn, traceback.format_exc()))
        result = {"ok": False, "msg": "%s: %s" % (type(exc).__name__, exc)}
    # 统一 JSON 输出（函数可能返回字符串/dict，浏览器端一律按 JSON 解析）
    response.content_type = "application/json"
    return json.dumps(result, ensure_ascii=False)


@app.get("/")
def index():
    return static_file("index.html", root=os.path.join(APP, "webui"), mimetype="text/html")


@app.get("/<path:path>")
def statics(path):
    if path == "eel.js":
        response.content_type = "application/javascript"
        return ""          # 兼容旧引用：返回空脚本
    return static_file(path, root=os.path.join(APP, "webui"))


AUTO_LOG = os.path.join(APP, "data", "auto_import.log")
# 每天固定跑的时间点（开机那次由服务启动自动覆盖；每晚 17:30 收工前再跑一次）
SCHED_TIMES = ["11:30", "17:30"]


def _auto_once(tag=""):
    r = app_gui.auto_import_stock(False)
    if r.get("updated"):
        line = "[%s] 库存表已更新：%s 项，库存日期 %s，来源 %s" % (
            tag, r.get("n"), r.get("sd"), r.get("source"))
    else:
        line = "[%s] 无新表（%s），上次导入 %s" % (tag, r.get("reason", ""), r.get("last_time", "无"))
    if r.get("fail_count"):
        line += "　失败计数 %s" % r["fail_count"]
    try:
        b = app_gui.import_batches(None, False)
        if b.get("ok") and b.get("updated"):
            line += "　｜　批次表已更新：%s 物料 / %s 批号" % (b.get("items"), b.get("batches"))
    except Exception as exc:
        line += "　｜　批次表导入异常：%s: %s" % (type(exc).__name__, exc)
    with open(AUTO_LOG, "a", encoding="utf-8") as f:
        f.write("%s  %s\n" % (datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), line))
    return r


def auto_stock_worker():
    """后台定时自动导入：开机跑一次 + 每天 11:30 / 17:30 + 每 5 分钟兜底检查"""
    today, done = "", set()
    try:
        time.sleep(20)
        _auto_once("开机")            # 开机即更新一次
    except Exception:
        _write_crash("开机自动导入异常: " + traceback.format_exc())
    while True:
        try:
            now = datetime.datetime.now()
            key = now.strftime("%Y-%m-%d")
            if key != today:
                today, done = key, set()
            hm = now.strftime("%H:%M")
            for t in SCHED_TIMES:
                if hm >= t and t not in done:
                    done.add(t)
                    _auto_once(t)
        except Exception:
            _write_crash("自动导入线程异常: " + traceback.format_exc())
        time.sleep(300)


if __name__ == "__main__":
    if not acquire_single_instance():
        log_req("检测到已有服务实例，当前启动请求已忽略")
        sys.exit(0)
    threading.Thread(target=auto_stock_worker, daemon=True).start()

    def _open():
        try:
            os.startfile("http://127.0.0.1:8765/")
        except Exception:
            import webbrowser
            webbrowser.open("http://127.0.0.1:8765/")
    threading.Timer(1.2, _open).start()
    # waitress：多线程 WSGI 服务器，单个请求卡住不影响其他请求
    from waitress import serve
    log_req("服务启动（waitress 多线程, 127.0.0.1:8765；自动更新：开机 + 11:30 + 17:30）")
    serve(app, host="127.0.0.1", port=8765, threads=12)
