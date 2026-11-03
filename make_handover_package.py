# -*- coding: utf-8 -*-
"""构建全量交接备份包：完整项目快照 + 对话归档 + SHA-256 清单。
用法：python make_handover_package.py
输出：D:\\QMDownload\\生产指令单系统_全量交接备份_<时间戳>\\（ZIP + SHA256.txt）"""
import datetime
import hashlib
import json
import os
import shutil
import sys
import tempfile
import zipfile

APP = os.path.dirname(os.path.abspath(__file__))
DOWNLOAD = r"D:\QMDownload"
NAME = "生产指令单系统_全量交接备份"
PREVIOUS_PKG = os.path.join(
    DOWNLOAD, "生产指令单系统_全量交接备份_20261030_164236")

EXCLUDE_DIRS = {"__pycache__", ".venv", ".git", "node_modules", "记忆层"}
EXCLUDE_DIR_SUFFIXES = ("_backup_",)
# 含密钥/本机隐私的文件不进交接包
EXCLUDE_FILES = {os.path.join("data", "ai_config.json")}


def iter_project_files():
    for root, dirs, files in os.walk(APP):
        dirs[:] = sorted(d for d in dirs
                         if d not in EXCLUDE_DIRS
                         and not any(s in d for s in EXCLUDE_DIR_SUFFIXES))
        for filename in sorted(files):
            if filename.endswith(".pyc"):
                continue
            full = os.path.join(root, filename)
            rel = os.path.relpath(full, APP)
            if rel.replace("\\", "/") in {p.replace("\\", "/") for p in EXCLUDE_FILES}:
                continue
            yield full


def sha256_of(path):
    digest = hashlib.sha256()
    with open(path, "rb") as fp:
        for chunk in iter(lambda: fp.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def main():
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    pkg_name = "%s_%s" % (NAME, stamp)
    stage = tempfile.mkdtemp(prefix="handover-")
    pkg_root = os.path.join(stage, pkg_name)
    os.makedirs(os.path.join(pkg_root, "软件"))

    # 1) 完整项目快照
    manifest = []
    dest_base = os.path.join(pkg_root, "软件", "OrderSystem")
    for src in iter_project_files():
        rel = os.path.relpath(src, APP)
        dst = os.path.join(dest_base, rel)
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        shutil.copy2(src, dst)
        manifest.append({"path": "软件/OrderSystem/" + rel.replace("\\", "/"),
                         "bytes": os.path.getsize(src),
                         "sha256": sha256_of(src)})

    # 2) 对话归档：沿用最近一个仍存在的旧包（本次会话未重新导出）
    archive_src = None
    prev_dirs = sorted(d for d in os.listdir(DOWNLOAD)
                       if d.startswith(NAME + "_") and os.path.isdir(
                           os.path.join(DOWNLOAD, d)))
    for prev in reversed(prev_dirs):
        candidate = os.path.join(DOWNLOAD, prev, "对话归档")
        if os.path.isdir(candidate):
            archive_src = candidate
            break
    if archive_src:
        archive_dst = os.path.join(pkg_root, "对话归档")
        shutil.copytree(archive_src, archive_dst)
        for root, _dirs, files in os.walk(archive_dst):
            for filename in sorted(files):
                path = os.path.join(root, filename)
                rel = os.path.relpath(path, pkg_root).replace("\\", "/")
                manifest.append({"path": rel, "bytes": os.path.getsize(path),
                                 "sha256": sha256_of(path)})
    else:
        print("!! 警告：找不到旧包的「对话归档」，本包不含对话历史。"
              "请从 Harness 界面导出后放入包目录再重打包。")

    # 3) README（随项目维护，每包自带）
    readme_src = os.path.join(APP, "README_交接与安装.md")
    if os.path.isfile(readme_src):
        shutil.copy2(readme_src, os.path.join(pkg_root, "README_交接与安装.md"))
        manifest.append({"path": "README_交接与安装.md",
                         "bytes": os.path.getsize(readme_src),
                         "sha256": sha256_of(readme_src)})

    with open(os.path.join(pkg_root, "SHA256_文件清单.json"), "w",
              encoding="utf-8") as fp:
        json.dump({"generated": datetime.datetime.now().isoformat(timespec="seconds"),
                   "files": manifest}, fp, ensure_ascii=False, indent=1)

    # 4) ZIP + SHA256.txt
    out_dir = os.path.join(DOWNLOAD, pkg_name)
    os.makedirs(out_dir, exist_ok=True)
    zip_path = os.path.join(out_dir, pkg_name + ".zip")
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as archive:
        for root, dirs, files in os.walk(pkg_root):
            dirs.sort()
            for filename in sorted(files):
                filepath = os.path.join(root, filename)
                archive.write(filepath, os.path.join(
                    pkg_name, os.path.relpath(filepath, pkg_root)))
    zip_sha = sha256_of(zip_path)
    with open(os.path.join(out_dir, "SHA256.txt"), "w", encoding="utf-8") as fp:
        fp.write("%s  %s.zip\n" % (zip_sha, pkg_name))

    shutil.rmtree(stage, ignore_errors=True)
    size_mb = os.path.getsize(zip_path) / (1024 * 1024)
    print("打包完成：%s" % zip_path)
    print("  ZIP %.1f MB，包内 %d 个文件（清单见 SHA256_文件清单.json）" % (
        size_mb, len(manifest)))
    print("  ZIP SHA-256：%s" % zip_sha)
    return 0


if __name__ == "__main__":
    sys.exit(main())
