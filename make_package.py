# -*- coding: utf-8 -*-
"""Build a green distribution without destroying the previous one."""
import datetime
import json
import os
import shutil
import sys
import tempfile
import zipfile

APP = os.path.dirname(os.path.abspath(__file__))
DIST = os.path.join(APP, "dist")
NAME = "生产指令单系统_绿色版"
ROOT = os.path.join(DIST, NAME)

PY_FILES = ["generator.py", "app_gui.py", "api_server.py", "stock.py", "orders.py",
            "auto_update.py", "daily_report.py", "add_product.py", "interactive.py",
            "json_store.py", "sales_form_import.py", "backup_manager.py"]
DATA_FILES = ["recipes.json", "mapping.json"]
EMPTY_RUNTIME_DATA = {
    "stock.json": {"_meta": {}, "items": []},
    "orders.json": {"orders": []},
}
DOC_FILES = ["使用与维护手册.md", "换电脑安装说明.md", "requirements.txt",
             "安装依赖.bat", "启动系统.bat", "run_gui.bat", "daily.bat",
             "stock.bat", "orders.bat"]


def _copy_tree_files(source, destination):
    count = 0
    for root, dirs, files in os.walk(source):
        dirs.sort()
        for filename in files:
            src = os.path.join(root, filename)
            rel = os.path.relpath(src, source)
            dst = os.path.join(destination, rel)
            os.makedirs(os.path.dirname(dst), exist_ok=True)
            shutil.copy2(src, dst)
            count += 1
    return count


def main():
    os.makedirs(DIST, exist_ok=True)
    stage_dir = tempfile.mkdtemp(prefix=".package-stage-", dir=DIST)
    staged_root = os.path.join(stage_dir, NAME)
    os.makedirs(staged_root)
    for subdir in ("images", "inbox", "templates", "data", os.path.join("data", "backups")):
        os.makedirs(os.path.join(staged_root, subdir), exist_ok=True)

    count = 0
    for filename in PY_FILES + DATA_FILES + DOC_FILES:
        src = os.path.join(APP, filename)
        if os.path.isfile(src):
            shutil.copy2(src, os.path.join(staged_root, filename))
            count += 1
    for filename, value in EMPTY_RUNTIME_DATA.items():
        with open(os.path.join(staged_root, filename), "w", encoding="utf-8") as stream:
            json.dump(value, stream, ensure_ascii=False, indent=2)
            stream.write("\n")
        count += 1

    webui_src = os.path.join(APP, "webui")
    webui_dst = os.path.join(staged_root, "webui")
    os.makedirs(webui_dst, exist_ok=True)
    for filename in os.listdir(webui_src):
        if filename.endswith((".html", ".js", ".css", ".ico", ".png")):
            shutil.copy2(os.path.join(webui_src, filename), os.path.join(webui_dst, filename))
            count += 1

    template = os.path.join(APP, "templates", "生产指令与交接单.xlsx")
    if not os.path.isfile(template):
        raise FileNotFoundError("缺少固定生产母版：%s" % template)
    shutil.copy2(template, os.path.join(staged_root, "templates", "生产指令与交接单.xlsx"))
    count += 1

    image_source = os.path.join(APP, "images")
    if os.path.isdir(image_source):
        image_count = _copy_tree_files(image_source, os.path.join(staged_root, "images"))
        print("  已附带图片：%d 张" % image_count)

    info = {
        "打包时间": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "文件数": count,
        "用法": "1) 双击 安装依赖.bat（首次）  2) 双击 启动系统.bat  3) 浏览器会自动打开",
        "注意": "换电脑后先看 换电脑安装说明.md",
    }
    with open(os.path.join(staged_root, "版本说明.txt"), "w", encoding="utf-8") as stream:
        stream.write(json.dumps(info, ensure_ascii=False, indent=2))

    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    suffix = 0
    while True:
        unique_stamp = stamp if suffix == 0 else "%s_%02d" % (stamp, suffix)
        backup_root = os.path.join(DIST, "%s_backup_%s" % (NAME, unique_stamp))
        zip_path = os.path.join(DIST, "%s_%s.zip" % (NAME, unique_stamp))
        if not os.path.exists(backup_root) and not os.path.exists(zip_path):
            break
        suffix += 1
    staged_zip = os.path.join(stage_dir, "%s_%s.zip" % (NAME, stamp))
    with zipfile.ZipFile(staged_zip, "w", zipfile.ZIP_DEFLATED) as archive:
        for root, dirs, files in os.walk(staged_root):
            dirs.sort()
            for filename in sorted(files):
                filepath = os.path.join(root, filename)
                archive.write(filepath, os.path.join(
                    NAME, os.path.relpath(filepath, staged_root)))

    moved_old = False
    promoted_root = False
    promoted_zip = False
    try:
        if os.path.exists(ROOT):
            os.replace(ROOT, backup_root)
            moved_old = True
        os.replace(staged_root, ROOT)
        promoted_root = True
        os.replace(staged_zip, zip_path)
        promoted_zip = True
    except Exception:
        if promoted_zip and os.path.exists(zip_path):
            os.remove(zip_path)
        if promoted_root and os.path.exists(ROOT):
            os.replace(ROOT, staged_root)
        if moved_old and os.path.exists(backup_root):
            os.replace(backup_root, ROOT)
        raise
    finally:
        if os.path.isdir(stage_dir):
            shutil.rmtree(stage_dir)

    size_kb = os.path.getsize(zip_path) / 1024.0
    print("\n打包完成：%s（%.0f KB，%d 个文件）" % (zip_path, size_kb, count))
    if moved_old:
        print("旧发行目录已保留：%s" % backup_root)
    return zip_path


if __name__ == "__main__":
    sys.exit(0 if main() else 1)
