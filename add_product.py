# -*- coding: utf-8 -*-
"""添加新产品到配方库（交互式，自动校验比例合计）"""
import json
import os
import sys
import ctypes
from json_store import save_json_atomic

APP = os.path.dirname(os.path.abspath(__file__))
RECIPE = os.path.join(APP, "recipes.json")

try:
    ctypes.windll.kernel32.SetConsoleOutputCP(65001)
    ctypes.windll.kernel32.SetConsoleCP(65001)
except Exception:
    pass
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

STAGE_DEF = [
    {"name": "一次混装原料", "prefix": "先加", "action": "预混", "time": "10-15"},
    {"name": "二次混装原料", "prefix": "再加", "action": "混合", "time": "5"},
    {"name": "三次混装原料", "prefix": "最后加", "action": "混合", "time": "5"},
]


def ask(prompt, default=None):
    tip = prompt
    if default is not None:
        tip += "（回车用 %s）" % default
    v = input(tip + "：").strip()
    if v == "" and default is not None:
        return default
    return v


def main():
    if not os.path.exists(RECIPE):
        print("找不到配方库：%s" % RECIPE)
        input("按回车退出")
        return

    with open(RECIPE, "r", encoding="utf-8") as f:
        data = json.load(f)

    existing = [k for k in data if not k.startswith("_")]
    print("=" * 56)
    print(" 添加新产品")
    print("=" * 56)
    print("现有产品（%d 个）：" % len(existing))
    for i, k in enumerate(existing, start=1):
        print("   %d  %s" % (i, k))
    print()

    name = ask("新产品名称")
    if not name:
        print("产品名称不能为空")
        input("按回车退出")
        return
    if name in data:
        print("该产品已存在，请用别的名称，或直接修改配方库文件")
        input("按回车退出")
        return

    tablet = ask("片重 g/片", "0.5")
    cat = ask("产品类别（压片糖果/运动营养食品/固体饮料…）", "压片糖果")

    print()
    print("-" * 56)
    print(" 录入物料（一次录一种，物料名称直接回车 = 结束）")
    print(" 混料阶段：1=一次混装（主料）  2=二次混装  3=三次混装")
    print("-" * 56)

    mats = []
    while True:
        mn = ask("物料名称")
        if not mn:
            break
        try:
            ratio = float(ask("  比例 %"))
        except Exception:
            print("  比例必须是数字，这一项已跳过")
            continue
        st = ask("  混料阶段 1/2/3", "1")
        try:
            st = int(st)
        except Exception:
            st = 1
        if st not in (1, 2, 3):
            st = 1
        batch = ask("  物料批号（可空）", "")
        mats.append({"name": mn, "ratio": ratio, "stage": st - 1, "batch": batch})
        print("  已录入：%s  %s%%" % (mn, ratio))
        print()

    if not mats:
        print("没有录入任何物料，已取消")
        input("按回车退出")
        return

    total = sum(m["ratio"] for m in mats)
    print("-" * 56)
    print("比例合计：%s%%（应为 100%%）" % round(total, 2))
    if abs(total - 100) > 0.01:
        print("注意：合计不等于 100%，请确认研发给的配比是否录全。")
        go = ask("仍然保存吗？y/n", "n")
        if go.lower() != "y":
            print("已取消，未保存")
            input("按回车退出")
            return

    stages = STAGE_DEF[:max(m["stage"] for m in mats) + 1]
    data[name] = {"category": cat, "tablet_g": float(tablet), "stages": stages, "materials": mats}

    save_json_atomic(RECIPE, data, indent=2)

    print()
    print("已保存：《%s》" % name)
    print("以后打开「生产指令单生成器」，产品列表里就能选到它了。")
    print()
    input("按回车退出")


if __name__ == "__main__":
    main()
