# -*- coding: utf-8 -*-
"""把整个会话导出成一份 Markdown，供其他 AI 阅读"""
import json
import re
import io
import os

BASE = os.path.dirname(os.path.abspath(__file__))
SRC = os.environ.get("CHAT_SOURCE", "")
OUT = os.path.join(BASE, "项目全档案_对话与更改记录.md")


def text_of(content):
    if isinstance(content, str):
        return content
    if isinstance(content, list):
        out = []
        for p in content:
            if isinstance(p, dict) and p.get("type") in ("input_text", "output_text"):
                out.append(p.get("text", ""))
        return "\n".join(out)
    return ""


def clean(t):
    t = re.sub(r"<system-reminder[\s\S]*?</system-reminder>", "", t)
    t = re.sub(r"<user_context>[\s\S]*?</user_context>", "", t)
    t = re.sub(r"<user_info>[\s\S]*?</user_info>", "", t)
    t = re.sub(r"<task-notification>[\s\S]*?</task-notification>", "[后台任务通知，已省略]", t)
    t = re.sub(r"<image_local_path>[\s\S]*?</image_local_path>", "", t)
    t = re.sub(r"<user_query>\s*", "", t)
    t = re.sub(r"</user_query>", "", t)
    t = re.sub(r"\n{3,}", "\n\n", t)
    return t.strip()


def short(t, limit=2600):
    """过长的助手回复保留头尾，中间摘要"""
    if len(t) <= limit:
        return t
    head = t[: limit - 700]
    tail = t[-600:]
    return head + "\n\n……（此处省略约 %d 字的工具输出/代码细节）……\n\n" % (len(t) - limit + 100) + tail


def main():
    if not SRC:
        raise SystemExit("请设置 CHAT_SOURCE 环境变量指向会话 JSONL")
    msgs = []
    with open(SRC, encoding="utf-8") as fp:
        for line in fp:
            try:
                d = json.loads(line)
            except Exception:
                continue
            if d.get("type") != "message":
                continue
            t = clean(text_of(d.get("content")))
            if not t:
                continue
            msgs.append((d.get("timestamp", ""), d.get("role", ""), t))

    users = [m for m in msgs if m[1] == "user"]
    assis = [m for m in msgs if m[1] == "assistant"]

    out = io.StringIO()
    w = out.write

    w("# 生产指令单系统 · 项目全档案（对话与更改记录）\n\n")
    w("> 本文档由系统自动生成，包含本次项目从零到交付的**全部用户对话**与**全部代码更改**。\n")
    w("> 阅读对象：其他 AI 助手。文件位置：`D:\\QMDownload\\OrderSystem\\`\n\n")
    w("---\n\n")

    # ============ 项目概览 ============
    w("## 一、项目概览\n\n")
    w("**用户**：李浩冉，河南佰雀归健康科技有限公司（南阳）跟单/计划岗，负责对接销售、对接生产、管理库存。\n\n")
    w("**核心诉求**：把「生产指令单」从手工填表变成一套可实际使用的系统——产品配方、库存管理、批次管理、订单台账、生产指令单 Excel 生成。\n\n")
    w("**技术栈**：Python（Bottle + Waitress 多线程 HTTP）+ 纯 JS 前端（无框架，自绘下拉联想）+ openpyxl 填充 Excel 模板。\n")
    w("> 早期用 Eel 框架，因系统代理拦截 WebSocket（502）而废弃，改为纯 HTTP + JS Proxy 兼容层。\n\n")
    w("**系统位置**：`D:\\QMDownload\\OrderSystem\\`\n")
    w("**界面**：http://127.0.0.1:8765 （`启动系统.bat` 拉起，守护循环自愈）\n\n")
    w("**数据来源**：柠檬云进销存导出的两张表：\n")
    w("1. `商品库存余额表`（数量、可用量；**无批号列**）\n")
    w("2. `批次跟踪表`（**流水账**：批次号、生产/到期日期、业务类型、数量——需要按业务类型判进出方向）\n\n")
    w("---\n\n")

    # ============ 时间线 ============
    w("## 二、需求演进时间线（64 轮对话摘要）\n\n")
    w("按主题分组，完整原文见第四章。\n\n")
    timeline = [
        ("阶段 1｜早期桌面工具（对话 1-7）",
         "锁屏快捷键与密码解锁、桌面加锁屏小程序、桌面图标位置固定（DesktopOK 便携版）、"
         "办公软件选型建议（填表/效率）。"),
        ("阶段 2｜生产指令单系统立项（8-15）",
         "从用户给的《生产流程.docx》《公司组织架构图》和 BOM 表切入，"
         "明确生产编号规则（当天第几单流水号）、销售经理名单（曾帅/付豪/房浩绅/马士龙）、"
         "实际产量需入表、片重可变、配方比例不变。"),
        ("阶段 3｜系统落地与库存体系（16-25）",
         "整套系统搬到 D:\\QMDownload\\OrderSystem；"
         "接入柠檬云库存表建立库存账（以柠檬云为准）；批次管理（一个产品分多批一起做）；"
         "库存警报与补料建议；每日库存比对（账面 vs 柠檬云 vs 订单应耗）；订单台账。"),
        ("阶段 4｜界面可用性反复打磨（26-47）",
         "MANAGERS 常量丢失→修复；片重选项 0.50-0.80+其他；销售经理/制表人/批次/当天第几单"
         "全部下拉化（含「其他」自填）；深色模式；库存按紧急度排序；批次 FIFO（优先用先到期的批）；"
         "物料增删改、比例可改、可导入新比例；包材一行行手动输入；每瓶片数；"
         "Eel WebSocket 被系统代理拦截（502）→ 改纯 HTTP + waitress + 守护 bat；"
         "batchOverrides is not defined、pb is not defined 等 JS 崩溃修复。"),
        ("阶段 5｜严重 UI bug 与批量修复（48-57）",
         "物料表列错位根因：recalcNeeds() 重绘成 6 列但表头 9 列 → 改为"
         "「renderOrderMats() 唯一渲染 + recalcNeeds() 只更新单元格」架构；"
         "名称联想检索（打字即搜库存）、自动带默认批号、多批次自动拆分；"
         "交货日期改为 10.30 格式。"),
        ("阶段 6｜批次与自动更新（58-61）",
         "批次号缺失根因：库存表无批号列，需导「批次状态表」；"
         "自动库存更新改为 开机+11:30+17:30 三时点 + 三次失败报警；"
         "系统自检面板（9 项检查 + 每项「怎么办」）；维护手册（换模板四种情况）。"),
        ("阶段 7｜批次跟踪表导入与打包（61-64）",
         "识别「批次跟踪表」是流水账（数量全正数，靠业务类型判进出）；"
         "原材料只有出库行导致净额为负 → 批号照收、数量按账面校正；"
         "路径硬编码全部改为动态定位；绿色版打包（换电脑可用）。"),
    ]
    for t, desc in timeline:
        w("**%s**\n\n%s\n\n" % (t, desc))

    w("---\n\n")

    # ============ 更改全记录 ============
    w("## 三、全部代码更改清单\n\n")
    w("### 3.1 文件清单（当前系统）\n\n")
    w("| 文件 | 作用 |\n|---|---|\n")
    for f, desc in [
        ("generator.py", "核心：读 Excel 模板，填入 BOM/表头/图片，输出生产指令单 xlsx"),
        ("app_gui.py", "业务函数层（被 api_server 复用）：预览计算、库存、批次、自检、导入"),
        ("api_server.py", "HTTP API 服务（Bottle + Waitress），崩溃钩子、自动更新线程"),
        ("stock.py", "库存：解析柠檬云两张表、批次 FIFO 分配、库存报告、警报"),
        ("orders.py", "订单台账（待生产/生产中/已完成/已发货），预占库存逻辑"),
        ("auto_update.py", "命令行自动更新入口（供定时任务/后台调用）"),
        ("recipes.json", "产品配方（6 个产品）：片重、物料配比、包装配置"),
        ("mapping.json", "物料名 → 库存名的映射（一水葡萄糖等）"),
        ("stock.json", "库存账面（53 项 + 批次信息）"),
        ("orders.json", "订单台账数据"),
        ("webui/index.html", "前端单页：出单/库存/新增产品/系统自检"),
        ("使用与维护手册.md", "出问题怎么处理、模板怎么换"),
        ("换电脑安装说明.md", "新电脑 3 步部署"),
        ("requirements.txt / 安装依赖.bat / 启动系统.bat", "部署与启动"),
        ("make_package.py", "打包脚本（生成绿色版 zip）"),
    ]:
        w("| `%s` | %s |\n" % (f, desc))

    w("\n### 3.2 关键改动点（按文件）\n\n")
    w("#### `generator.py` —— 指令单生成\n\n")
    w("- **模板填充**：三段表头（混料 row2 / 压片 row33 / 包装 row43）+ BOM 区（一次 6-13、二次 14-16、三次/其他 17-19，行区有保护不越到 row20 领料核对区）\n")
    w("- **日期格式**：生产日期 `10.27`，交货日期 `10.30`（只体现月.日）\n")
    w("- **物料可编辑**：前端传 materials（stage 1-4 → 内部 0-3），支持增删改比例、改阶段、改批号\n")
    w("- **多批次拆分**：`plan_alloc()` 按有效期先进先出；单批不够时拆成多行（每批一行：名称/批号/用量）；行不够时合并写成「A 取 25；B 取 14」\n")
    w("- **批号以库存为准**：`stock_rows()` + `lookup_row()` 现算，不用配方库死值\n")
    w("- **模板路径**：`_template_path()` → config.json > 桌面 > 包内 templates 目录\n")
    w("- **包材**：手动录入时数量=总套数（`pack_direct` 标志），否则按瓶数乘\n")
    w("- **图片**：`images/<产品>/压片图|物料1-6|成品图.*`，`find_img`/`put_img`\n\n")

    w("#### `stock.py` —— 库存与批次\n\n")
    w("- **`open_table(path)`**：统一读 .xls(xlrd) / .xlsx(openpyxl)，接口与 xlrd sheet 一致\n")
    w("- **`import_xls`**：解析商品库存余额表；若表内自带「批号」列则顺带解析批次（`extract_batches_from_stock`）\n")
    w("- **`import_batch_xls`**：解析批次状态表/批次跟踪表；\n")
    w("  - 自动找表头行（`find_header_row` 按「命中单元格×10+非空单元格数」评分，避免标题行抢占）\n")
    w("  - 识别商品名称/批号/数量/生产日期/到期日期/业务类型列\n")
    w("  - **按业务类型判方向**：出库/领用/领料/盘亏/退货 → 负；入库/盘盈/退库 → 正\n")
    w("  - **合并策略**：净额≤0 的批次 qty 记 0 但不删除；与旧批次合并，新数量为 0 时沿用旧数量\n")
    w("- **`allocate(item, need)`**：批次 FIFO 分配，返回 [(批号, 用量)]\n")
    w("- **`build_report()`**：合并池（一水+无水葡萄糖）、预占（待生产订单）、可用量、状态（断料/超订/不足/偏低/正常）\n")
    w("- **`default_download_dirs()`**：用户 Downloads + C/D/E 盘的 Downloads（换电脑自动找对）\n\n")

    w("#### `app_gui.py` —— 业务逻辑\n\n")
    w("- `order_preview(product, mode, qty, tablet, materials, per_bottle)`：反算产量、模糊匹配库存、批次 FIFO 分配，返回 default_batch/alloc_list/multi\n")
    w("- `search_stock(kw, limit)`：联想检索（前缀>包含>规格类别），返回可用量/批号/状态\n")
    w("- `auto_import_stock(force)`：自动导入（文件指纹去重 + 3 次重试 + fail_count 报警）\n")
    w("- `import_batches(path, force)`：批次表导入（独立指纹）\n")
    w("- `fix_batch_qty()`：按账面等比修正批次数量；数量全未知时归到最早到期批次\n")
    w("- `self_check()`：9 项自检（服务/库存新鲜度/批次覆盖/批次一致性/模板/配方库/台账/图片/报错/自动更新），每项带「怎么办」\n")
    w("- `open_doc(kind)` / `clear_logs()`：打开手册/图片目录、清日志\n\n")

    w("#### `api_server.py` —— 服务\n\n")
    w("- Bottle + Waitress（threads=12），端口 8765\n")
    w("- 崩溃钩子（sys.excepthook / threading.excepthook）→ data/server_crash.log\n")
    w("- `auto_stock_worker` 线程：开机跑一次 + 每天 11:30/17:30 + 每 5 分钟兜底\n")
    w("- FN 注册表：ping/get_products/suggest_batch/order_preview/generate_order/get_stock/search_stock/auto_import_stock/import_batches/fix_batch_qty/get_replenish/import_stock/save_product/open_inbox/open_report/open_path/open_doc/clear_logs/self_check/report_js_error\n\n")

    w("#### `webui/index.html` —— 前端\n\n")
    w("- **eel 兼容层**：`new Proxy` 双层函数，内部 fetch POST /api/<fn>，AbortController 20 秒超时\n")
    w("- **物料表（9 列，铁律）**：只由 `renderOrderMats()` 渲染；`recalcNeeds()` 只更新 need-i/avail-i/st-i 三个单元格，**绝不重绘**（此前 bug：重绘成 6 列导致列错位）\n")
    w("- **联想检索**：`acQuery/acRender/acKey/acPick`，打字即搜库存，↑↓ 选择，选中自动带名称+默认批号+可用量\n")
    w("- **多批次显示**：批号列 placeholder 显示「自动拆批：20260225 取 25；20260525 取 14」\n")
    w("- **深色模式**、库存按紧急度排序、包材行 4 列（名称/批号/数量/删）\n")
    w("- **自检面板** `#checkBox`、错误横幅、自动更新状态提示\n")
    w("- 全部选项化：片重 0.50-0.80+其他、销售经理、制表人、批次 01-99、生产日期、混装阶段 1-4（含其他）\n\n")

    w("### 3.3 修复过的典型问题\n\n")
    w("| 问题 | 根因 | 修法 |\n|---|---|---|\n")
    w("| 页面一直「未连接」 | Eel WebSocket 被系统代理拦截（CONNECT 502） | 弃 Eel，改纯 HTTP + waitress |\n")
    w("| `batchOverrides is not defined` | 重写前端漏掉变量声明 | 补 `let batchOverrides = {}` |\n")
    w("| `pb is not defined` | recalcNeeds 中漏定义 | 就地补 `const pb` |\n")
    w("| 物料行/比例列消失、数据错行 | recalcNeeds 重绘成 6 列但表头 9 列 | 唯一渲染 + 单元格更新 |\n")
    w("| 麦芽糊精显示缺料 | 4 张测试单预占 71kg，账面只有 65kg | 清空测试单 |\n")
    w("| 批号全部为空 | 库存表本身无批号列 | 支持批次表导入 + 表内自解析 |\n")
    w("| 批次数量虚高/被清空 | 跟踪表是流水账，原材料只有出库行 | 按业务类型判方向 + 合并保留旧数据 |\n")
    w("| 手动包材数量被再乘瓶数 | pack_direct 未区分 | 加 pack_direct 标志 |\n")
    w("| 同一物料多映射重复扣库存 | 一水/无水各扣一次 | 合并池分组 |\n")
    w("| `MANAGERS` 导入错误 | 重写 generator 时漏常量 | 补回常量 |\n\n")

    w("---\n\n")

    # ============ 完整对话 ============
    w("## 四、完整对话记录\n\n")
    w("> 按时间顺序。用户消息为**原文**；助手回复保留核心内容（超长部分省略中间细节）。\n")
    w("> 图片附件、系统提示词、后台任务通知已省略。\n\n")

    for i, (ts, role, t) in enumerate(msgs, 1):
        if role == "user":
            w("\n### 🧑 用户 #%d\n\n%s\n\n" % (i, t))
        else:
            w("\n### 🤖 助手 #%d\n\n%s\n\n" % (i, short(t)))
            w("---\n")

    w("\n---\n\n## 五、当前系统状态（导出时刻）\n\n")
    w("- 自检结果：0 个问题，2 个提醒（批次覆盖 36%、暂无产品图片）\n")
    w("- 库存：53 项，库存日期 2026-10-28；原材料 17/47 有批号（批次合计=账面一致）\n")
    w("- 订单台账：0 张（测试单已清空）\n")
    w("- 配方库：6 个产品，配比合计均 100%\n")
    w("- 绿色版打包：`dist\\生产指令单系统_绿色版_20261028_1033.zip`（含自带模板）\n")
    w("- 自动更新：开机 + 11:30 + 17:30 三时点，三次失败报警\n\n")
    w("## 六、待办事项\n\n")
    w("1. 大豆分离蛋白：等仓库在柠檬云建档后补 mapping\n")
    w("2. 产品图片：放 `images\\<产品名>\\` 下（压片图/物料1-6/成品图）\n")
    w("3. 批次真实余额：如需准确批次数量，导「批次状态表」（当前跟踪表只有流水，数量为按账面推算值）\n")
    w("4. 其余 5 个产品的包材清单（现仅示例）\n")

    with open(OUT, "w", encoding="utf-8") as f:
        f.write(out.getvalue())
    print("导出完成：%s（%.0f KB）" % (OUT, os.path.getsize(OUT) / 1024.0))
    print("消息条数：用户 %d / 助手 %d" % (len(users), len(assis)))


if __name__ == "__main__":
    main()
