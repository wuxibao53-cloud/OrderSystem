# -*- coding: utf-8 -*-
"""Validated snapshots and restoration for the system's business JSON files."""
import datetime
import hashlib
import json
import os
import tempfile
import zipfile

from json_store import save_json_atomic

APP = os.path.dirname(os.path.abspath(__file__))
BACKUP_DIR = os.path.join(APP, "data", "backups")
BUSINESS_FILES = ("stock.json", "orders.json", "recipes.json", "mapping.json")


def _read_business_files():
    values = {}
    for name in BUSINESS_FILES:
        path = os.path.join(APP, name)
        if not os.path.isfile(path):
            raise FileNotFoundError("缺少业务数据文件：%s" % name)
        with open(path, "r", encoding="utf-8") as stream:
            values[name] = json.load(stream)
    return values


def create_backup():
    os.makedirs(BACKUP_DIR, exist_ok=True)
    values = _read_business_files()
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    target = os.path.join(BACKUP_DIR, "OrderSystem_%s.zip" % timestamp)
    fd, temp_path = tempfile.mkstemp(prefix=".backup-", suffix=".zip", dir=BACKUP_DIR)
    os.close(fd)
    manifest = {
        "created_at": datetime.datetime.now().isoformat(timespec="seconds"),
        "files": {},
    }
    try:
        with zipfile.ZipFile(temp_path, "w", zipfile.ZIP_DEFLATED) as archive:
            for name in BUSINESS_FILES:
                content = json.dumps(values[name], ensure_ascii=False, indent=2).encode("utf-8")
                archive.writestr(name, content)
                manifest["files"][name] = hashlib.sha256(content).hexdigest()
            archive.writestr("manifest.json", json.dumps(
                manifest, ensure_ascii=False, indent=2).encode("utf-8"))
        os.replace(temp_path, target)
    finally:
        if os.path.exists(temp_path):
            os.remove(temp_path)
    prune_backups()
    return {"ok": True, "path": target, "created_at": manifest["created_at"],
            "files": list(BUSINESS_FILES)}


def prune_backups(keep_recent=30, keep_monthly=12):
    """保留最近 N 份有效备份，另保留近 M 个月各月最新一份；无效备份不删除，留待人工检查。"""
    removed = []
    try:
        valid = [b for b in list_backups() if b.get("valid")]
        keep = {b["name"] for b in valid[:keep_recent]}
        months = []
        for b in valid:
            month = b["name"][len("OrderSystem_"):][:6]
            if month not in months:
                months.append(month)
                if len(months) <= keep_monthly:
                    keep.add(b["name"])
        for b in valid:
            if b["name"] not in keep:
                os.remove(b["path"])
                removed.append(b["name"])
    except Exception:
        pass
    return removed


def list_backups():
    if not os.path.isdir(BACKUP_DIR):
        return []
    backups = []
    for name in os.listdir(BACKUP_DIR):
        if not name.lower().endswith(".zip") or not name.startswith("OrderSystem_"):
            continue
        path = os.path.join(BACKUP_DIR, name)
        if os.path.isfile(path):
            item = {"name": name, "path": path,
                    "modified": datetime.datetime.fromtimestamp(
                        os.path.getmtime(path)).isoformat(timespec="minutes"),
                    "size": os.path.getsize(path)}
            try:
                _validated_backup(path)
                item["valid"] = True
            except Exception as exc:
                item["valid"] = False
                item["validation_error"] = "%s: %s" % (type(exc).__name__, exc)
            backups.append(item)
    backups.sort(key=lambda item: item["name"], reverse=True)
    return backups


def _validated_backup(path):
    root = os.path.realpath(BACKUP_DIR)
    resolved = os.path.realpath(path or "")
    if os.path.commonpath([root, resolved]) != root:
        raise ValueError("只能恢复系统备份目录中的备份文件")
    if not os.path.isfile(resolved) or not resolved.lower().endswith(".zip"):
        raise ValueError("备份文件不存在或格式不正确")

    with zipfile.ZipFile(resolved, "r") as archive:
        names = set(archive.namelist())
        manifest_name = "manifest.json"
        if manifest_name not in names:
            raise ValueError("备份缺少校验清单")
        manifest = json.loads(archive.read(manifest_name).decode("utf-8"))
        restored = {}
        for name in BUSINESS_FILES:
            if name not in names or manifest.get("files", {}).get(name) is None:
                raise ValueError("备份缺少必要文件：%s" % name)
            content = archive.read(name)
            if hashlib.sha256(content).hexdigest() != manifest["files"][name]:
                raise ValueError("备份文件校验失败：%s" % name)
            restored[name] = json.loads(content.decode("utf-8"))
    if not isinstance(restored["stock.json"], dict) or not isinstance(
            restored["stock.json"].get("items"), list):
        raise ValueError("备份中的库存结构无效")
    orders = restored["orders.json"]
    if isinstance(orders, dict):
        orders = orders.get("orders", [])
    if not isinstance(orders, list):
        raise ValueError("备份中的订单台账结构无效")
    if not isinstance(restored["recipes.json"], dict):
        raise ValueError("备份中的配方结构无效")
    if not isinstance(restored["mapping.json"], dict):
        raise ValueError("备份中的映射结构无效")
    return restored


def restore_backup(path):
    restored = _validated_backup(path)
    before = _read_business_files()
    created = create_backup()
    try:
        for name in BUSINESS_FILES:
            save_json_atomic(os.path.join(APP, name), restored[name])
    except Exception:
        for name in BUSINESS_FILES:
            save_json_atomic(os.path.join(APP, name), before[name])
        raise
    return {"ok": True, "restored": os.path.basename(path),
            "safety_backup": created["path"], "files": list(BUSINESS_FILES)}
