# 生产指令单系统 · 交接与安装速览

> 详细版本历史与设计说明见 `全量备份与安装交接说明.md`（随包携带）。本文件是快速上手卡。

## 这是什么
食品厂生产指令单/跟单系统：销售单（Word/文字/微信记录）→ AI 提取+人工核对 → 四段式生产交接单 Excel → 台账批次跟踪 → 库存与采购预测 → 每日备份与恢复演练。

## 安装（新电脑）
1. 解压本包到 `D:\QMDownload\OrderSystem`（路径变动则同步修改 `maintenance\install_tasks.ps1` 里的绝对路径）。
2. 安装 WorkBuddy Python（版本见交接说明），确认 `openpyxl / python-docx / xlrd` 可导入（`libs\xlrd` 为内置备份）。
3. 管理员 PowerShell 执行 `maintenance\install_tasks.ps1`，注册 5 个定时任务：
   AutoStart（登录启动）/ HealthCheck（5 分钟心跳自愈）/ DailyBackup（18:00 备份）/ RestoreVerify（季度恢复演练）/ LemonSync（开机+12:00+17:35 柠檬云库存同步）。
4. 双击 `启动系统.bat`，浏览器开 `http://127.0.0.1:8765`。
5. AI 功能（可选）：手动创建 `data\ai_config.json`（含 key，**不进包**，从旧电脑拷贝或重新填写）。柠檬云导出目录默认 `D:\Downloads`，可在 `data\lemon\lemon_config.json` 修改。

## 每天怎么用（一分钟版）
- 生产跟踪页看今天盯哪些单；销售导入页：Word+微信文字合并解析 → 逐项核对 → 生成交接单（批次自动入账）。
- 柠檬云导出两个表到 D:\Downloads 后无需任何操作；差异自动核对修正（备份+历史留痕）。
- 首页「AI 今日简报」看当日盯单/采购/自检摘要；「AI 助手」页可直接用中文问台账。

## 安全红线
- AI 只读+提取，生成任何文件/改台账必须人工勾选确认。
- 库存自动修正仅限编码精确匹配；所有修正可从 `data\lemon\stock_backup_*.json` 回滚。
- 业务数据每天 18:00 自动备份（校验+快照），季度恢复演练确保备份真的能用。
