# 生产指令单生成器 - 图形版（出单 / 库存 / 新增产品 三合一）
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$APP = Split-Path -Parent $MyInvocation.MyCommand.Path
$PYW = $null
foreach ($candidate in @(
    (Join-Path $APP ".venv\Scripts\pythonw.exe"),
    (Join-Path $env:USERPROFILE ".workbuddy\binaries\python\envs\default\Scripts\pythonw.exe")
)) {
    if (Test-Path $candidate) { $PYW = $candidate; break }
}
if (-not $PYW) {
    $cmd = Get-Command pythonw.exe -ErrorAction SilentlyContinue
    if ($cmd) { $PYW = $cmd.Source }
}
if (-not $PYW) {
    [System.Windows.Forms.MessageBox]::Show("Python not found. Run the install script first.", "OrderSystem")
    exit 1
}
$GEN = Join-Path $APP "generator.py"
$RECIPE_JSON = Join-Path $APP "recipes.json"
$STOCK_JSON = Join-Path $APP "stock.json"
$INBOX = Join-Path $APP "inbox"

function Read-Json($path) {
    if (-not (Test-Path $path)) { return $null }
    try {
        return Get-Content $path -Raw -Encoding UTF8 | ConvertFrom-Json
    } catch { return $null }
}

function Save-JsonUtf8($obj, $path) {
    $json = $obj | ConvertTo-Json -Depth 8
    [System.IO.File]::WriteAllText($path, $json, (New-Object System.Text.UTF8Encoding $false))
}

function Load-Recipes {
    $r = Read-Json $RECIPE_JSON
    $out = @()
    if ($r) {
        foreach ($p in $r.PSObject.Properties) {
            if ($p.Name.StartsWith("_")) { continue }
            $out += $p
        }
    }
    return $out
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "生产指令单 · 跟单助手"
$form.Size = New-Object System.Drawing.Size(760, 640)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Microsoft YaHei", 9)

$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Location = New-Object System.Drawing.Point(8, 8)
$tabs.Size = New-Object System.Drawing.Size(728, 585)
$form.Controls.Add($tabs)

# ============ 页1：出单 ============
$tab1 = New-Object System.Windows.Forms.TabPage
$tab1.Text = "出单"
$tabs.TabPages.Add($tab1)

function New-Field($parent, $label, $x, $y, $w) {
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $label; $l.Location = New-Object System.Drawing.Point($x, $y)
    $l.Size = New-Object System.Drawing.Size(96, 20)
    $parent.Controls.Add($l)
    $t = New-Object System.Windows.Forms.TextBox
    $t.Location = New-Object System.Drawing.Point(($x + 100), ($y - 2))
    $t.Size = New-Object System.Drawing.Size($w, 22)
    $parent.Controls.Add($t)
    return $t
}

$y = 14
$cbProduct = New-Object System.Windows.Forms.ComboBox
$cbProduct.Location = New-Object System.Drawing.Point(108, $y)
$cbProduct.Size = New-Object System.Drawing.Size(280, 22)
$cbProduct.DropDownStyle = "DropDownList"
$tab1.Controls.Add($cbProduct)
$btnNewProd = New-Object System.Windows.Forms.Button
$btnNewProd.Text = "新增产品…"
$btnNewProd.Location = New-Object System.Drawing.Point(400, ($y - 2))
$btnNewProd.Size = New-Object System.Drawing.Size(96, 26)
$tab1.Controls.Add($btnNewProd)
$btnReload = New-Object System.Windows.Forms.Button
$btnReload.Text = "刷新"
$btnReload.Location = New-Object System.Drawing.Point(504, ($y - 2))
$btnReload.Size = New-Object System.Drawing.Size(60, 26)
$tab1.Controls.Add($btnReload)
$y += 34

$txtTablet = New-Field $tab1 "片重 g/片" 8 $y 90 | Select-Object -Last 1
$y += 32

$rbWan = New-Object System.Windows.Forms.RadioButton
$rbWan.Text = "按片数（万片）"; $rbWan.Location = New-Object System.Drawing.Point(8, $y)
$rbWan.Size = New-Object System.Drawing.Size(120, 20); $rbWan.Checked = $true
$tab1.Controls.Add($rbWan)
$rbTotal = New-Object System.Windows.Forms.RadioButton
$rbTotal.Text = "按总量（kg）"; $rbTotal.Location = New-Object System.Drawing.Point(130, $y)
$rbTotal.Size = New-Object System.Drawing.Size(110, 20)
$tab1.Controls.Add($rbTotal)
$txtQty = New-Object System.Windows.Forms.TextBox
$txtQty.Location = New-Object System.Drawing.Point(250, ($y - 2))
$txtQty.Size = New-Object System.Drawing.Size(110, 22)
$tab1.Controls.Add($txtQty)
$y += 32

$txtBottles = New-Field $tab1 "瓶数（可空）" 8 $y 90 | Select-Object -Last 1
$txtDue = New-Field $tab1 "交货日期" 320 $y 110 | Select-Object -Last 1
$y += 32

$l = New-Object System.Windows.Forms.Label
$l.Text = "销售经理"; $l.Location = New-Object System.Drawing.Point(8, $y)
$l.Size = New-Object System.Drawing.Size(96, 20)
$tab1.Controls.Add($l)
$cbManager = New-Object System.Windows.Forms.ComboBox
$cbManager.Location = New-Object System.Drawing.Point(108, $y)
$cbManager.Size = New-Object System.Drawing.Size(140, 22)
$cbManager.DropDownStyle = "DropDownList"
$tab1.Controls.Add($cbManager)
$y += 32

$txtDate = New-Field $tab1 "生产日期" 8 $y 100 | Select-Object -Last 1
$txtSeq = New-Field $tab1 "当天第几单" 320 $y 60 | Select-Object -Last 1
$y += 32
$txtMaker = New-Field $tab1 "制表人" 8 $y 100 | Select-Object -Last 1
$txtBatch = New-Field $tab1 "生产批次" 320 $y 60 | Select-Object -Last 1
$y += 34

$l = New-Object System.Windows.Forms.Label
$l.Text = "物料批号（本次进货不同可直接改）"
$l.Location = New-Object System.Drawing.Point(8, $y)
$l.Size = New-Object System.Drawing.Size(300, 18)
$tab1.Controls.Add($l)
$y += 20

$gridBatch = New-Object System.Windows.Forms.DataGridView
$gridBatch.Location = New-Object System.Drawing.Point(8, $y)
$gridBatch.Size = New-Object System.Drawing.Size(420, 96)
$gridBatch.AllowUserToAddRows = $false
$gridBatch.RowHeadersVisible = $false
$tab1.Controls.Add($gridBatch)
$gridBatch.Columns.Add("c1", "物料") | Out-Null
$gridBatch.Columns.Add("c2", "默认批号") | Out-Null
$gridBatch.Columns.Add("c3", "本次批号（空=默认）") | Out-Null
$gridBatch.Columns[0].ReadOnly = $true
$gridBatch.Columns[1].ReadOnly = $true
$gridBatch.Columns[0].Width = 160
$gridBatch.Columns[1].Width = 120
$gridBatch.Columns[2].Width = 132
$y += 104

$txtCoat = New-Field $tab1 "包衣要求" 8 $y 90 | Select-Object -Last 1
$txtShape = New-Field $tab1 "形状" 320 $y 90 | Select-Object -Last 1
$y += 32

$l = New-Object System.Windows.Forms.Label
$l.Text = "包材（本次实际用的瓶/盖/标签…，留空用默认）"
$l.Location = New-Object System.Drawing.Point(8, $y)
$l.Size = New-Object System.Drawing.Size(380, 18)
$tab1.Controls.Add($l)
$btnPickPack = New-Object System.Windows.Forms.Button
$btnPickPack.Text = "从库存选"
$btnPickPack.Location = New-Object System.Drawing.Point(420, ($y - 3))
$btnPickPack.Size = New-Object System.Drawing.Size(84, 24)
$tab1.Controls.Add($btnPickPack)
$y += 20

$gridPack = New-Object System.Windows.Forms.DataGridView
$gridPack.Location = New-Object System.Drawing.Point(8, $y)
$gridPack.Size = New-Object System.Drawing.Size(420, 88)
$gridPack.AllowUserToAddRows = $true
$gridPack.RowHeadersVisible = $false
$tab1.Controls.Add($gridPack)
$gridPack.Columns.Add("p1", "包材名称") | Out-Null
$gridPack.Columns.Add("p2", "数量（套）") | Out-Null
$gridPack.Columns[0].Width = 270
$gridPack.Columns[1].Width = 140
$y += 94

$btnGen = New-Object System.Windows.Forms.Button
$btnGen.Text = "生成生产指令单（自动进台账 + 预扣库存）"
$btnGen.Location = New-Object System.Drawing.Point(8, $y)
$btnGen.Size = New-Object System.Drawing.Size(330, 34)
$btnGen.BackColor = [System.Drawing.Color]::FromArgb(26, 77, 143)
$btnGen.ForeColor = [System.Drawing.Color]::White
$tab1.Controls.Add($btnGen)
$y += 40

$txtResult = New-Object System.Windows.Forms.TextBox
$txtResult.Location = New-Object System.Drawing.Point(8, $y)
$txtResult.Size = New-Object System.Drawing.Size(706, 160)
$txtResult.Multiline = $true
$txtResult.ScrollBars = "Vertical"
$txtResult.ReadOnly = $true
$tab1.Controls.Add($txtResult)

# ============ 页2：库存 ============
$tab2 = New-Object System.Windows.Forms.TabPage
$tab2.Text = "库存"
$tabs.TabPages.Add($tab2)

$txtSearch = New-Object System.Windows.Forms.TextBox
$txtSearch.Location = New-Object System.Drawing.Point(10, 12)
$txtSearch.Size = New-Object System.Drawing.Size(300, 22)
$tab2.Controls.Add($txtSearch)
$btnSearch = New-Object System.Windows.Forms.Button
$btnSearch.Text = "搜索"
$btnSearch.Location = New-Object System.Drawing.Point(318, 10)
$btnSearch.Size = New-Object System.Drawing.Size(70, 25)
$tab2.Controls.Add($btnSearch)
$btnAll = New-Object System.Windows.Forms.Button
$btnAll.Text = "显示全部"
$btnAll.Location = New-Object System.Drawing.Point(394, 10)
$btnAll.Size = New-Object System.Drawing.Size(80, 25)
$tab2.Controls.Add($btnAll)
$btnOpenInbox = New-Object System.Windows.Forms.Button
$btnOpenInbox.Text = "打开收件箱"
$btnOpenInbox.Location = New-Object System.Drawing.Point(480, 10)
$btnOpenInbox.Size = New-Object System.Drawing.Size(96, 25)
$tab2.Controls.Add($btnOpenInbox)
$btnImportStock = New-Object System.Windows.Forms.Button
$btnImportStock.Text = "导入最新库存表"
$btnImportStock.Location = New-Object System.Drawing.Point(582, 10)
$btnImportStock.Size = New-Object System.Drawing.Size(132, 25)
$tab2.Controls.Add($btnImportStock)

$gridStock = New-Object System.Windows.Forms.DataGridView
$gridStock.Location = New-Object System.Drawing.Point(10, 44)
$gridStock.Size = New-Object System.Drawing.Size(702, 458)
$gridStock.ReadOnly = $true
$gridStock.AllowUserToAddRows = $false
$gridStock.RowHeadersVisible = $false
$gridStock.AutoSizeColumnsMode = "Fill"
$tab2.Controls.Add($gridStock)
foreach ($h in @("物料", "类别", "账面", "占用", "可用", "警戒线", "单位", "状态", "批次")) {
    $gridStock.Columns.Add($h, $h) | Out-Null
}

$l = New-Object System.Windows.Forms.Label
$l.Text = "数据来源：柠檬云《商品库存余额表》（以柠檬云为准）。占用 = 在途订单已预扣。"
$l.Location = New-Object System.Drawing.Point(10, 512)
$l.Size = New-Object System.Drawing.Size(690, 20)
$tab2.Controls.Add($l)

# ============ 页3：新增产品 ============
$tab3 = New-Object System.Windows.Forms.TabPage
$tab3.Text = "新增产品"
$tabs.TabPages.Add($tab3)

$txtName = New-Field $tab3 "产品名称" 8 16 240 | Select-Object -Last 1
$txtCat = New-Field $tab3 "类别" 8 46 180 | Select-Object -Last 1
$txtCat.Text = "压片糖果"
$txtTabletN = New-Field $tab3 "片重 g/片" 8 76 80 | Select-Object -Last 1
$txtTabletN.Text = "0.5"
$l = New-Object System.Windows.Forms.Label
$l.Text = "物料配比（比例合计必须 100%）"
$l.Location = New-Object System.Drawing.Point(8, 108)
$l.Size = New-Object System.Drawing.Size(300, 18)
$tab3.Controls.Add($l)

$gridMat = New-Object System.Windows.Forms.DataGridView
$gridMat.Location = New-Object System.Drawing.Point(8, 128)
$gridMat.Size = New-Object System.Drawing.Size(480, 180)
$gridMat.AllowUserToAddRows = $true
$gridMat.RowHeadersVisible = $false
$tab3.Controls.Add($gridMat)
$gridMat.Columns.Add("m1", "物料名称") | Out-Null
$gridMat.Columns.Add("m2", "比例 %") | Out-Null
$gridMat.Columns.Add("m3", "阶段(1/2/3)") | Out-Null
$gridMat.Columns.Add("m4", "批号") | Out-Null
$gridMat.Columns[0].Width = 180
$gridMat.Columns[1].Width = 90
$gridMat.Columns[2].Width = 100
$gridMat.Columns[3].Width = 100

$btnAddMat = New-Object System.Windows.Forms.Button
$btnAddMat.Text = "加一行"
$btnAddMat.Location = New-Object System.Drawing.Point(500, 128)
$btnAddMat.Size = New-Object System.Drawing.Size(80, 26)
$tab3.Controls.Add($btnAddMat)
$btnSaveProd = New-Object System.Windows.Forms.Button
$btnSaveProd.Text = "保存新产品（记住）"
$btnSaveProd.Location = New-Object System.Drawing.Point(8, 320)
$btnSaveProd.Size = New-Object System.Drawing.Size(200, 34)
$btnSaveProd.BackColor = [System.Drawing.Color]::FromArgb(30, 125, 67)
$btnSaveProd.ForeColor = [System.Drawing.Color]::White
$tab3.Controls.Add($btnSaveProd)

$txtProdMsg = New-Object System.Windows.Forms.TextBox
$txtProdMsg.Location = New-Object System.Drawing.Point(8, 364)
$txtProdMsg.Size = New-Object System.Drawing.Size(560, 40)
$txtProdMsg.Multiline = $true
$txtProdMsg.ReadOnly = $true
$tab3.Controls.Add($txtProdMsg)

# ============ 逻辑 ============
$script:recipes = Load-Recipes

function Fill-Products {
    $cbProduct.Items.Clear()
    foreach ($p in $script:recipes) { $cbProduct.Items.Add($p.Name) | Out-Null }
    if ($cbProduct.Items.Count -gt 0) { $cbProduct.SelectedIndex = 0 }
    $cbManager.Items.Clear()
    foreach ($m in @("曾帅", "付豪", "房浩绅", "马士龙")) { $cbManager.Items.Add($m) | Out-Null }
    if ($cbManager.Items.Count -gt 0) { $cbManager.SelectedIndex = 0 }
}

function Get-Recipe($name) {
    foreach ($p in $script:recipes) { if ($p.Name -eq $name) { return $p.Value } }
    return $null
}

function Refresh-BatchGrid {
    $gridBatch.Rows.Clear()
    $r = Get-Recipe $cbProduct.Text
    if (-not $r) { return }
    foreach ($m in $r.materials) {
        $gridBatch.Rows.Add($m.name, $m.batch, "") | Out-Null
    }
    $txtTablet.Text = $r.tablet_g
}

$cbProduct.Add_SelectedIndexChanged({ Refresh-BatchGrid })
$btnReload.Add_Click({
    $script:recipes = Load-Recipes
    Fill-Products
    Refresh-BatchGrid
})

$txtDate.Text = (Get-Date).ToString("yyyyMMdd")
$txtSeq.Text = "01"
$txtBatch.Text = "01"
$txtMaker.Text = "李浩冉"
Fill-Products
Refresh-BatchGrid

$btnPickPack.Add_Click({
    $stock = Read-Json $STOCK_JSON
    if (-not $stock) { [System.Windows.Forms.MessageBox]::Show("请先导入库存数据"); return }
    $packs = @($stock.items | Where-Object { $_.cat -like "原材料-包装材料*" })
    if ($packs.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("库存中没有包材"); return }
    $f2 = New-Object System.Windows.Forms.Form
    $f2.Text = "选择包材"; $f2.Size = New-Object System.Drawing.Size(420, 420)
    $f2.StartPosition = "CenterParent"
    $lb = New-Object System.Windows.Forms.ListBox
    $lb.Location = New-Object System.Drawing.Point(10, 10)
    $lb.Size = New-Object System.Drawing.Size(380, 300)
    foreach ($p in $packs) {
        $lb.Items.Add(("{0}  （库存 {1} {2}）" -f $p.name, $p.qty, $p.unit)) | Out-Null
    }
    $f2.Controls.Add($lb)
    $ok2 = New-Object System.Windows.Forms.Button
    $ok2.Text = "选定"; $ok2.Location = New-Object System.Drawing.Point(10, 320)
    $ok2.Size = New-Object System.Drawing.Size(90, 30)
    $f2.Controls.Add($ok2)
    $ok2.Add_Click({ $f2.Tag = $lb.SelectedIndex; $f2.Close() })
    $lb.Add_DoubleClick({ $f2.Tag = $lb.SelectedIndex; $f2.Close() })
    $f2.ShowDialog() | Out-Null
    if ($f2.Tag -ne $null -and $f2.Tag -ge 0) {
        $gridPack.Rows.Add($packs[$f2.Tag].name, 1) | Out-Null
    }
})

$btnSearch.Add_Click({ Refresh-StockGrid $txtSearch.Text })
$btnAll.Add_Click({ Refresh-StockGrid "" })

function Refresh-StockGrid($kw) {
    $stock = Read-Json $STOCK_JSON
    $gridStock.Rows.Clear()
    if (-not $stock) { return }
    $mapping = Read-Json (Join-Path $APP "mapping.json")
    $rev = @{}
    if ($mapping) {
        foreach ($p in $mapping.PSObject.Properties) {
            if ($p.Name.StartsWith("_")) { continue }
            if ($p.Value -is [System.Array]) {
                foreach ($v in $p.Value) { $rev[$v] = $p.Name }
            } elseif ($p.Value) { $rev[$p.Value] = $p.Name }
        }
    }
    $occ = @{}
    $ofile = Join-Path $APP "orders.json"
    if (Test-Path $ofile) {
        $od = Read-Json $ofile
        if ($od) {
            foreach ($o in $od.orders) {
                if ($o.status -in @("已完成", "已发货")) { continue }
                foreach ($m in $o.materials) {
                    if (-not $occ.ContainsKey($m.name)) { $occ[$m.name] = 0.0 }
                    $occ[$m.name] += [double]$m.kg
                }
            }
        }
    }
    foreach ($it in $stock.items) {
        if ($it.cat -like "产成品*") { continue }
        $key = $it.short
        if ($rev.ContainsKey($it.name)) { $key = $rev[$it.name] }
        $oc = 0.0
        if ($occ.ContainsKey($key)) { $oc = $occ[$key] }
        $avail = [double]$it.qty - $oc
        $st = "正常"
        if ($avail -le 0) { $st = "断料" } elseif ($it.safe -gt 0 -and $avail -lt $it.safe) { $st = "不足" }
        $bt = ""
        if ($it.batches) {
            $bt = ($it.batches | ForEach-Object { "{0}({1})" -f $_.batch, $_.qty }) -join "+"
        }
        $kw2 = $kw.ToLower()
        if ($kw2 -and -not ($it.short.ToLower().Contains($kw2) -or $it.name.ToLower().Contains($kw2) -or $it.cat.ToLower().Contains($kw2))) { continue }
        $gridStock.Rows.Add($it.short, $it.cat.Replace("原材料-", ""), $it.qty,
            [math]::Round($oc, 2), [math]::Round($avail, 2), $it.safe, $it.unit, $st, $bt) | Out-Null
    }
}
Refresh-StockGrid ""

$btnImportStock.Add_Click({
    $latest = $null
    if (Test-Path $INBOX) {
        $c = @(Get-ChildItem $INBOX -Include *.xls, *.xlsx -Recurse | Sort-Object LastWriteTime -Descending)
        if ($c.Count -gt 0) { $latest = $c[0].FullName }
    }
    if (-not $latest) {
        [System.Windows.Forms.MessageBox]::Show("收件箱里没有库存表。请把柠檬云导出的文件放进 inbox 文件夹。")
        return
    }
    & $PYW (Join-Path $APP "stock.py") "--compare" $latest | Out-Null
    Refresh-StockGrid ""
    $cmpDir = Join-Path $APP "data\compare"
    $newest = Get-ChildItem $cmpDir -Filter *.md | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    $msg = "已导入并更新账面（以柠檬云为准）。"
    if ($newest) { $msg += "`r`n比对报告：`r`n" + $newest.FullName }
    [System.Windows.Forms.MessageBox]::Show($msg, "导入完成")
})

$btnAddMat.Add_Click({ $gridMat.Rows.Add("", "", "1", "") | Out-Null })
1..3 | ForEach-Object { $gridMat.Rows.Add("", "", "1", "") | Out-Null }

$btnSaveProd.Add_Click({
    $name = $txtName.Text.Trim()
    if (-not $name) { $txtProdMsg.Text = "产品名称不能为空"; return }
    $mats = @(); $total = 0.0; $okAll = $true
    foreach ($row in $gridMat.Rows) {
        $mn = $row.Cells[0].Value
        if (-not $mn) { continue }
        try {
            $rv = [double]$row.Cells[1].Value; $total += $rv
            $st = 1; try { $st = [int]$row.Cells[2].Value } catch { }
            if ($st -lt 1 -or $st -gt 3) { $st = 1 }
            $mats += @{ name = ("$mn").Trim(); ratio = $rv; stage = ($st - 1); batch = ("$($row.Cells[3].Value)").Trim() }
        } catch {
            $txtProdMsg.Text = "比例必须是数字：$mn"; $okAll = $false; break
        }
    }
    if (-not $okAll) { return }
    if ($mats.Count -eq 0) { $txtProdMsg.Text = "至少录入一种物料"; return }
    if ([math]::Abs($total - 100) -gt 0.01) {
        $r2 = [System.Windows.Forms.MessageBox]::Show(("比例合计 {0}% 不是 100%，仍要保存吗？" -f $total), "确认", "YesNo")
        if ($r2 -ne "Yes") { return }
    }
    $rec = Read-Json $RECIPE_JSON
    if ($rec.PSObject.Properties[$name]) { $txtProdMsg.Text = "该产品已存在"; return }
    $maxStage = 0
    foreach ($m in $mats) { if ($m.stage -gt $maxStage) { $maxStage = $m.stage } }
    $stagesDef = @(
        @{ name = "一次混装原料"; prefix = "先加"; action = "预混"; time = "10-15" },
        @{ name = "二次混装原料"; prefix = "再加"; action = "混合"; time = "5" },
        @{ name = "三次混装原料"; prefix = "最后加"; action = "混合"; time = "5" }
    )
    $newP = @{
        category = $txtCat.Text
        tablet_g = [double]$txtTabletN.Text
        packaging = @{ per_bottle = 60; per_box = 500; coat = "无"; shape = "圆形" }
        stages = @($stagesDef[0..$maxStage])
        materials = @($mats)
    }
    $rec | Add-Member -MemberType NoteProperty -Name $name -Value $newP
    Save-JsonUtf8 $rec $RECIPE_JSON
    $script:recipes = Load-Recipes
    Fill-Products
    Refresh-BatchGrid
    $txtProdMsg.Text = "已保存《$name》，出单页的产品列表里已经可以选到（永久记住）。"
})

$btnGen.Add_Click({
    $r = Get-Recipe $cbProduct.Text
    if (-not $r) { $txtResult.Text = "请先选择产品"; return }
    $cfg = [ordered]@{ product = $cbProduct.Text; tablet_g = [double]$txtTablet.Text }
    if ($rbTotal.Checked) {
        try { $cfg.total_kg = [double]$txtQty.Text } catch { $txtResult.Text = "总量必须是数字"; return }
        if ($cfg.total_kg -le 0) { $txtResult.Text = "总量必须大于 0"; return }
    } else {
        try { $cfg.wan = [double]$txtQty.Text } catch { $txtResult.Text = "产量必须是数字"; return }
        if ($cfg.wan -le 0) { $txtResult.Text = "产量必须大于 0"; return }
    }
    $cfg.bottles = $txtBottles.Text
    $cfg.due = $txtDue.Text
    $cfg.date = $txtDate.Text
    $cfg.seq = [int]$txtSeq.Text
    $cfg.manager = $cbManager.Text
    $cfg.maker = $txtMaker.Text
    $cfg.batch = $txtBatch.Text
    $cfg.stock = @{}
    $cfg.batch_override = @{}
    foreach ($row in $gridBatch.Rows) {
        $bn = $row.Cells[2].Value
        if ($bn -and ("$bn").Trim() -ne "") {
            $cfg.batch_override["$($row.Cells[0].Value)"] = ("$bn").Trim()
        }
    }
    if ($txtCoat.Text.Trim() -ne "") { $cfg.coat = $txtCoat.Text }
    if ($txtShape.Text.Trim() -ne "") { $cfg.shape = $txtShape.Text }
    $pack = @()
    foreach ($row in $gridPack.Rows) {
        $pn = $row.Cells[0].Value
        if ($pn) {
            $q = 1
            try { $q = [int]$row.Cells[1].Value } catch { }
            $pack += @{ name = ("$pn").Trim(); qty = $q }
        }
    }
    $cfg.pack_materials = @($pack)

    $inPath = Join-Path $env:TEMP "produce_input.json"
    Save-JsonUtf8 $cfg $inPath
    & $PYW $GEN $inPath | Out-Null
    $resPath = Join-Path $env:TEMP "produce_result.txt"
    if (-not (Test-Path $resPath)) { $txtResult.Text = "生成失败：无返回"; return }
    $res = @{}
    foreach ($line in (Get-Content $resPath -Encoding UTF8)) {
        if ($line -match '^(.+?)=(.*)$') { $res[$Matches[1]] = $Matches[2] }
    }
    if ($res['OK'] -eq '1') {
        $out = "生成成功`r`n单号：" + $res['SERIAL'] + "  产品：" + $res['PRODUCT'] + "`r`n文件：" + $res['FILE'] + "`r`n`r`n投料量：`r`n"
        foreach ($pair in ($res['MATS'] -split '\|')) {
            $pp = $pair -split ':'
            if ($pp.Count -ge 2) { $out += ("  " + $pp[0] + " : " + $pp[1] + " kg`r`n") }
        }
        $out += "`r`n理论总重 " + $res['TOTAL_KG'] + " kg　实际投料 " + $res['ACTUAL_KG'] + " kg"
        $out += "`r`n包材：" + $res['BOTTLES'] + " 瓶（" + $res['BOXES'] + " 箱）`r`n"
        foreach ($pair in ($res['PACKS'] -split '\|')) { $out += ("  " + $pair + "`r`n") }
        if ($res['LIMITED'] -ne "") {
            $out += "`r`n注意：受限物料 " + $res['LIMITED'] + "，最多可做 " + $res['ACTUAL_WAN'] + " 万片（计划 " + $res['PLAN_WAN'] + " 万片）"
        }
        $txtResult.Text = $out
    } else {
        $txtResult.Text = "生成失败：" + $res['MSG']
    }
})

[System.Windows.Forms.Application]::Run($form)
