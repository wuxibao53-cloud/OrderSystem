using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace StudentArchive
{
    public class Student
    {
        public string Id = "", Name = "", Gender = "", ClassName = "", Phone = "", Origin = "",
            Family = "", Dorm = "", DormLeader = "", Tags = "", Photo = "";
    }

    public class Rec
    {
        public string StudentId = "", Cat = "", Date = "", Title = "", Detail = "";
        public string Status = "", Owner = "", Closed = "";
    }

    public class Db
    {
        public List<Student> Students = new List<Student>();
        public List<Rec> Recs = new List<Rec>();
        public List<string> Log = new List<string>();
        public string Salt = "", Hash = "";
    }

    static class Program
    {
        public static readonly string[] Cats = { "成绩", "荣誉奖项", "任职实践", "宿舍生活", "问题关注", "辅导员记录" };
        public static readonly string[] Sensitive = { "问题关注", "辅导员记录" };
        public static string DataFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "data.json");
        public static string PhotoDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "photos");
        public static Db Data = new Db();

        public static Image LoadPhoto(Student s)
        {
            if (s == null || s.Photo == "") return null;
            string p = Path.Combine(PhotoDir, s.Photo);
            if (!File.Exists(p)) return null;
            try { using (var ms = new MemoryStream(File.ReadAllBytes(p))) return new Bitmap(Image.FromStream(ms)); }
            catch { return null; }
        }

        public static bool SetPhoto(Student s, string src)
        {
            try
            {
                Directory.CreateDirectory(PhotoDir);
                string name = s.Id + Path.GetExtension(src).ToLower();
                File.Copy(src, Path.Combine(PhotoDir, name), true);
                s.Photo = name; return true;
            }
            catch { return false; }
        }

        static double? ParseScore(string t)
        {
            var m = Regex.Match(t ?? "", @"\d+(\.\d+)?");
            double v;
            if (m.Success && double.TryParse(m.Value, out v)) return v;
            return null;
        }

        public static double? ParseScorePub(string t) { return ParseScore(t); }

        public static List<KeyValuePair<string, double>> ScoreSeries(string id)
        {
            return Data.Recs.Where(r => r.StudentId == id && r.Cat == "成绩" && r.Date != "")
                .Select(r => new { r.Date, V = ParseScore(r.Detail) }).Where(x => x.V.HasValue)
                .GroupBy(x => x.Date).OrderBy(g => g.Key)
                .Select(g => new KeyValuePair<string, double>(g.Key, Math.Round(g.Average(x => x.V.Value), 1))).ToList();
        }

        public static DateTime? LastRecDate(string id)
        {
            string last = Data.Recs.Where(r => r.StudentId == id && r.Date != "")
                .Select(r => r.Date).OrderBy(x => x).LastOrDefault();
            if (last == "") return null;
            DateTime d;
            string t = last.Length > 10 ? last.Substring(0, 10) : last;
            if (DateTime.TryParse(t, out d)) return d;
            return null;
        }

        public static List<string> AlertReasons(string id)
        {
            var reasons = new List<string>();
            var recs = Data.Recs.Where(r => r.StudentId == id).ToList();
            var scores = recs.Where(r => r.Cat == "成绩").Select(r => ParseScore(r.Detail)).Where(v => v.HasValue).Select(v => v.Value).ToList();
            int fail = scores.Count(v => v < 60);
            if (fail >= 2) reasons.Add("挂科 " + fail + " 门");
            var ser = ScoreSeries(id);
            if (ser.Count >= 2)
            {
                double drop = ser[ser.Count - 2].Value - ser[ser.Count - 1].Value;
                if (drop > 10) reasons.Add("平均分骤降 " + drop.ToString("0.#") + " 分");
            }
            if (!recs.Any(r => r.Date != "")) reasons.Add("档案无任何记录");
            else
            {
                var d = LastRecDate(id);
                if (d.HasValue && (DateTime.Now - d.Value).TotalDays > 60)
                    reasons.Add("已 " + (int)(DateTime.Now - d.Value).TotalDays + " 天无新记录");
            }
            return reasons;
        }

        public static string ReadText(string path)
        {
            var b = File.ReadAllBytes(path);
            try { return new UTF8Encoding(false, true).GetString(b); }
            catch { return Encoding.GetEncoding(936).GetString(b); }
        }
        public static bool Unlocked = false;
        static JavaScriptSerializer js = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };

        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Load();
            DailyAutoBackup();
            Application.Run(new MainForm());
        }

        public static void DailyAutoBackup()
        {
            try
            {
                string dir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "自动备份");
                Directory.CreateDirectory(dir);
                string today = Path.Combine(dir, "data_" + DateTime.Now.ToString("yyyyMMdd") + ".json");
                if (!File.Exists(today) && File.Exists(DataFile))
                {
                    File.Copy(DataFile, today, true);
                    // 只保留最近 14 份
                    var old = new DirectoryInfo(dir).GetFiles("data_*.json").OrderByDescending(f => f.Name).Skip(14);
                    foreach (var f in old) { try { f.Delete(); } catch { } }
                }
            }
            catch { }
        }

        public static void Load()
        {
            if (File.Exists(DataFile)) Data = js.Deserialize<Db>(File.ReadAllText(DataFile, Encoding.UTF8));
        }

        public static void Save()
        {
            if (File.Exists(DataFile)) File.Copy(DataFile, DataFile + ".bak", true);
            File.WriteAllText(DataFile, js.Serialize(Data), Encoding.UTF8);
        }

        public static string HashPwd(string pwd, string salt)
        {
            using (var d = new Rfc2898DeriveBytes(pwd, Convert.FromBase64String(salt), 20000))
                return Convert.ToBase64String(d.GetBytes(32));
        }

        public static void AddLog(string msg)
        {
            Data.Log.Add(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " + Environment.UserName + " " + msg);
            Save();
        }
    }

    class InputDlg : Form
    {
        public TextBox Box;
        public InputDlg(string title, string prompt, bool pwd)
        {
            Text = title; Width = 360; Height = 150; StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog; MaximizeBox = false; MinimizeBox = false;
            var l = new Label { Text = prompt, Left = 12, Top = 12, Width = 320 };
            Box = new TextBox { Left = 12, Top = 36, Width = 320, UseSystemPasswordChar = pwd };
            var ok = new Button { Text = "确定", Left = 170, Top = 70, DialogResult = DialogResult.OK };
            var cancel = new Button { Text = "取消", Left = 257, Top = 70, DialogResult = DialogResult.Cancel };
            Controls.AddRange(new Control[] { l, Box, ok, cancel });
            AcceptButton = ok; CancelButton = cancel;
        }
        public static string Ask(string title, string prompt, bool pwd)
        {
            using (var d = new InputDlg(title, prompt, pwd))
                return d.ShowDialog() == DialogResult.OK ? d.Box.Text : null;
        }
    }

    class RecDlg : Form
    {
        public TextBox Date, Title, Detail, RespOwner, ClosedDate;
        public ComboBox StatusBox;
        public RecDlg(string cat, Rec r)
        {
            Text = cat + (r.Title == "" ? " - 新增" : " - 编辑");
            bool isIssue = cat == "问题关注";
            Width = 460; Height = isIssue ? 500 : 380; StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog; MaximizeBox = false; MinimizeBox = false;
            string t1 = "日期/学期", t2 = "标题", t3 = "详情";
            if (cat == "成绩") { t1 = "学期(如 2023-2024-1)"; t2 = "课程"; t3 = "成绩/绩点/排名"; }
            if (cat == "荣誉奖项") { t1 = "获奖日期"; t2 = "奖项名称"; t3 = "级别、颁发单位等"; }
            if (cat == "问题关注") { t1 = "发生日期"; t2 = "类型(处分/预警/心理/经济/家庭…)"; t3 = "经过与处理"; }
            if (cat == "辅导员记录") { t1 = "日期"; t2 = "类型(谈话/家访/帮扶…)"; t3 = "内容"; }
            if (cat == "辅导员记录" && r.Detail == "")
                r.Detail = "时间/地点: \r\n谈话事由: \r\n学生状态: \r\n后续措施: \r\n";
            int dlgY = 126;
            Date = new TextBox { Left = 12, Top = 30, Width = 420, Text = r.Date == "" ? DateTime.Now.ToString("yyyy-MM-dd") : r.Date };
            Title = new TextBox { Left = 12, Top = 78, Width = 420, Text = r.Title };
            Detail = new TextBox { Left = 12, Top = dlgY, Width = 420, Height = isIssue ? 150 : 150, Multiline = true, ScrollBars = ScrollBars.Vertical, Text = r.Detail };
            var ctrl = new List<Control> {
                new Label { Text = t1, Left = 12, Top = 10, Width = 420 }, Date,
                new Label { Text = t2, Left = 12, Top = 58, Width = 420 }, Title,
                new Label { Text = t3, Left = 12, Top = dlgY - 20, Width = 420 }, Detail };
            int by = dlgY + 158;
            if (isIssue)
            {
                StatusBox = new ComboBox { Left = 12, Top = by, Width = 120, DropDownStyle = ComboBoxStyle.DropDownList };
                StatusBox.Items.AddRange(new object[] { "发现", "处理中", "已销号" });
                StatusBox.SelectedItem = string.IsNullOrEmpty(r.Status) ? "发现" : r.Status;
                RespOwner = new TextBox { Left = 175, Top = by, Width = 120, Text = r.Owner };
                ClosedDate = new TextBox { Left = 335, Top = by, Width = 97, Text = r.Closed };
                ctrl.Add(new Label { Text = "状态", Left = 12, Top = by - 18 });
                ctrl.Add(new Label { Text = "责任人", Left = 175, Top = by - 18 });
                ctrl.Add(new Label { Text = "销号日期", Left = 335, Top = by - 18 });
                ctrl.Add(StatusBox); ctrl.Add(RespOwner); ctrl.Add(ClosedDate);
                by += 34;
            }
            var ok = new Button { Text = "保存", Left = 270, Top = by + 4, DialogResult = DialogResult.OK };
            var cancel = new Button { Text = "取消", Left = 357, Top = by + 4, DialogResult = DialogResult.Cancel };
            ctrl.Add(ok); ctrl.Add(cancel);
            Controls.AddRange(ctrl.ToArray());
            AcceptButton = null; CancelButton = cancel;
        }
    }

    class MainForm : Form
    {
        TextBox search = new TextBox();
        ListView list = new ListView();
        TabControl tabs = new TabControl();
        Button lockBtn = new Button();
        Label status = new Label();
        Dictionary<string, ListView> catViews = new Dictionary<string, ListView>();
        ListView timeline = new ListView();
        Dictionary<string, TextBox> fields = new Dictionary<string, TextBox>();
        string[] fieldKeys = { "Id", "Name", "Gender", "ClassName", "Phone", "Origin", "Family", "Dorm", "DormLeader", "Tags" };
        string[] fieldNames = { "学号(唯一)", "姓名", "性别", "专业班级", "联系方式", "籍贯", "家庭情况", "宿舍(楼-房号)", "宿舍长", "标签(逗号分隔)" };
        Student cur;

        public MainForm()
        {
            Text = "学生档案系统";
            Width = 1180; Height = 720; StartPosition = FormStartPosition.CenterScreen;
            Font = new Font("Microsoft YaHei UI", 9f);

            var left = new Panel { Dock = DockStyle.Left, Width = 360 };
            var top = new Panel { Dock = DockStyle.Top, Height = 34 };
            search.SetBounds(4, 6, 350, 24);
            search.TextChanged += delegate { Refill(); };
            top.Controls.Add(search);
            var btns = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 36 };
            btns.Controls.Add(MkBtn("新增学生", delegate { NewStudent(); }));
            btns.Controls.Add(MkBtn("删除学生", delegate { DelStudent(); }));
            var imp = MkBtn("导入资料 ▾", null);
            var menu = new ContextMenuStrip();
            menu.Items.Add("导入学生名单(CSV)", null, delegate { ImportCsv(0); });
            menu.Items.Add("导入成绩(CSV)", null, delegate { ImportCsv(1); });
            menu.Items.Add("导入其他记录:奖项/任职/宿舍/问题等(CSV)", null, delegate { ImportCsv(2); });
            menu.Items.Add("批量导入照片(文件夹,文件名=学号)", null, delegate { ImportPhotos(); });
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add("生成CSV导入模板", null, delegate { MakeTemplates(); });
            imp.Click += delegate { menu.Show(imp, new Point(0, imp.Height)); };
            btns.Controls.Add(imp);
            list.Dock = DockStyle.Fill; list.View = View.Details; list.FullRowSelect = true; list.HideSelection = false;
            list.Columns.Add("学号", 80); list.Columns.Add("姓名", 60); list.Columns.Add("班级", 85); list.Columns.Add("标签", 65);
            list.Columns.Add("最后记录", 75); list.Columns.Add("预警", 90);
            list.SelectedIndexChanged += delegate { Pick(); };
            left.Controls.Add(list); left.Controls.Add(top); left.Controls.Add(btns);

            var bar = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 36 };
            lockBtn.Text = "🔒 解锁敏感信息"; lockBtn.AutoSize = true;
            lockBtn.Click += delegate { ToggleLock(); };
            bar.Controls.Add(lockBtn);
            bar.Controls.Add(MkBtn("导出画像卡(HTML)", delegate { ExportCard(false); }));
            bar.Controls.Add(MkBtn("打印画像卡", delegate { ExportCard(true); }));
            bar.Controls.Add(MkBtn("同宿舍成员", delegate { ShowDormMates(); }));
            bar.Controls.Add(MkBtn("班级汇总报告", delegate { ExportClassReport(); }));
            bar.Controls.Add(MkBtn("备份全部数据", delegate { BackupAll(); }));
            bar.Controls.Add(MkBtn("查看访问日志", delegate { ShowLog(); }));
            bar.Controls.Add(MkBtn("修改密码", delegate { SetPwd(); }));
            status.AutoSize = true; status.Padding = new Padding(8, 8, 0, 0);
            bar.Controls.Add(status);

            tabs.Dock = DockStyle.Fill;
            tabs.TabPages.Add(BasicTab());
            tabs.TabPages.Add(AlertTab());
            tabs.TabPages.Add(TimelineTab());
            tabs.TabPages.Add(ChartTab());
            tabs.TabPages.Add(StatsTab());
            foreach (var c in Program.Cats) tabs.TabPages.Add(CatTab(c));

            Controls.Add(tabs); Controls.Add(bar); Controls.Add(left);
            Refill(); RefillAlerts(); Pick(); UpdateLock();
        }

        Button MkBtn(string t, EventHandler h)
        {
            var b = new Button { Text = t, AutoSize = true };
            b.Click += h; return b;
        }

        TabPage BasicTab()
        {
            var p = new TabPage("基础信息") { AutoScroll = true };
            int y = 14;
            for (int i = 0; i < fieldKeys.Length; i++)
            {
                p.Controls.Add(new Label { Text = fieldNames[i], Left = 14, Top = y + 3, Width = 130 });
                var tb = new TextBox { Left = 150, Top = y, Width = 420 };
                fields[fieldKeys[i]] = tb; p.Controls.Add(tb); y += 34;
            }
            var save = MkBtn("保存基础信息", delegate { SaveBasic(); });
            save.Left = 150; save.Top = y + 6; p.Controls.Add(save);
            pic.SetBounds(620, 14, 150, 200); pic.SizeMode = PictureBoxSizeMode.Zoom; pic.BorderStyle = BorderStyle.FixedSingle;
            var up = MkBtn("上传照片", delegate { UploadPhoto(); }); up.Left = 620; up.Top = 222;
            var rm = MkBtn("移除照片", delegate { if (cur != null) { cur.Photo = ""; Program.Save(); ShowPhoto(); } }); rm.Left = 700; rm.Top = 222;
            p.Controls.AddRange(new Control[] { pic, up, rm });
            return p;
        }

        PictureBox pic = new PictureBox();
        Panel chart = new Panel();
        Panel stats = new Panel();

        void ShowPhoto()
        {
            if (pic.Image != null) { pic.Image.Dispose(); pic.Image = null; }
            pic.Image = Program.LoadPhoto(cur);
        }

        void UploadPhoto()
        {
            if (cur == null) return;
            var d = new OpenFileDialog { Filter = "图片|*.jpg;*.jpeg;*.png;*.bmp" };
            if (d.ShowDialog() != DialogResult.OK) return;
            if (Program.SetPhoto(cur, d.FileName)) { Program.Save(); ShowPhoto(); } else MessageBox.Show("照片保存失败");
        }

        TabPage ChartTab()
        {
            var p = new TabPage("成绩曲线");
            chart.Dock = DockStyle.Fill; chart.BackColor = Color.White;
            typeof(Panel).GetProperty("DoubleBuffered", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).SetValue(chart, true, null);
            var top = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 34 };
            courseFilter = new ComboBox { Width = 220, DropDownStyle = ComboBoxStyle.DropDownList };
            courseFilter.SelectedIndexChanged += delegate { chart.Invalidate(); };
            top.Controls.Add(new Label { Text = "课程:", Left = 4, Top = 8, AutoSize = true });
            top.Controls.Add(courseFilter);
            chart.Paint += DrawChart; chart.Resize += delegate { chart.Invalidate(); };
            p.Controls.Add(chart); p.Controls.Add(top); return p;
        }

        ComboBox courseFilter;

        void RefillCourses()
        {
            if (courseFilter == null) return;
            string sel = courseFilter.Text;
            courseFilter.BeginUpdate();
            var items = new List<string> { "(全部课程-平均分)" };
            if (cur != null)
                foreach (var c in Program.Data.Recs.Where(r => r.StudentId == cur.Id && r.Cat == "成绩").Select(r => r.Title).Distinct().OrderBy(x => x)) items.Add(c);
            courseFilter.Items.Clear();
            foreach (var it in items) courseFilter.Items.Add(it);
            courseFilter.SelectedIndex = items.Contains(sel) ? items.IndexOf(sel) : 0;
            courseFilter.EndUpdate();
        }

        TabPage StatsTab()
        {
            var p = new TabPage("统计看板");
            stats.Dock = DockStyle.Fill; stats.BackColor = Color.White; stats.AutoScroll = true;
            typeof(Panel).GetProperty("DoubleBuffered", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).SetValue(stats, true, null);
            stats.Paint += DrawStats; stats.Resize += delegate { stats.Invalidate(); };
            p.Controls.Add(stats); return p;
        }

        void Bar(Graphics g, string label, int count, int maxCount, int x, int y, int maxW, Color color)
        {
            g.DrawString(label, Font, Brushes.Black, x, y - 2);
            int w = maxCount <= 0 ? 0 : (int)((double)count / maxCount * maxW);
            var r = new Rectangle(x + 110, y, w, 18);
            using (var b = new SolidBrush(color)) g.FillRectangle(b, r);
            g.DrawRectangle(Pens.Gray, x + 110, y, maxW, 18);
            g.DrawString(count.ToString(), Font, Brushes.Black, x + 118 + maxW, y + 1);
        }

        void DrawStats(object o, PaintEventArgs e)
        {
            var g = e.Graphics; g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            g.DrawString("整体统计看板", new Font(Font.FontFamily, 13, FontStyle.Bold), Brushes.Black, 20, 12);
            int L = 30, W = Math.Max(200, stats.ClientSize.Width - 220);
            int y = 50;
            var studs = Program.Data.Students;
            g.DrawString("学生总数: " + studs.Count, Font, Brushes.DimGray, L, y); y += 30;
            if (studs.Count == 0)
            {
                g.DrawString("暂无学生数据。请先在左侧“新增学生”或“导入资料”里导入学生名单。", Font, Brushes.Gray, L, y);
                return;
            }

            // 按班级统计
            g.DrawString("一、各班级人数", new Font(Font.FontFamily, 11, FontStyle.Bold), Brushes.Black, L, y); y += 24;
            var byClass = studs.GroupBy(s => string.IsNullOrEmpty(s.ClassName) ? "(未填班级)" : s.ClassName)
                .OrderByDescending(gp => gp.Count()).ToList();
            int maxC = byClass.Max(gp => gp.Count());
            foreach (var gp in byClass)
            {
                if (y > stats.ClientSize.Height - 20) { stats.AutoScrollMinSize = new Size(0, y + 60); break; }
                Bar(g, gp.Key + " (" + gp.Count() + "人)", gp.Count(), maxC, L, y, W, Color.SteelBlue); y += 26;
            }
            y += 16;

            // 按标签统计
            g.DrawString("二、标签分布(一个学生可带多个标签)", new Font(Font.FontFamily, 11, FontStyle.Bold), Brushes.Black, L, y); y += 24;
            var tagGroups = studs.SelectMany(s => s.Tags.Split(new[] { ',', ',', ';', ';', ' ' }, StringSplitOptions.RemoveEmptyEntries))
                .GroupBy(t => t).OrderByDescending(gp => gp.Count()).ToList();
            if (tagGroups.Count == 0) { g.DrawString("(暂无标签)", Font, Brushes.Gray, L + 20, y); y += 26; }
            else
            {
                int maxT = tagGroups.Max(gp => gp.Count());
                foreach (var gp in tagGroups)
                {
                    if (y > stats.ClientSize.Height - 20) { stats.AutoScrollMinSize = new Size(0, y + 60); break; }
                    Color c = gp.Key.Contains("重点") ? Color.Firebrick : Color.MediumSeaGreen;
                    Bar(g, gp.Key, gp.Count(), maxT, L, y, W, c); y += 26;
                }
            }
            y += 16;

            // 成绩区间分布(以最新学期平均分算)
            g.DrawString("三、学生成绩水平分布(按每人所有成绩平均分)", new Font(Font.FontFamily, 11, FontStyle.Bold), Brushes.Black, L, y); y += 24;
            var avg = studs.Select(s => new {
                S = Program.Data.Recs.Where(r => r.StudentId == s.Id && r.Cat == "成绩")
                    .Select(r => Regex.Match(r.Detail ?? "", @"\d+(\.\d+)?")).Where(m => m.Success)
                    .Select(m => { double v; return double.TryParse(m.Value, out v) ? v : (double?)null; })
                    .Where(v => v.HasValue).Select(v => v.Value).DefaultIfEmpty(-1).Average()
            }).Where(x => x.S >= 0).ToList();
            string[] bands = { "90-100(优秀)", "80-89(良好)", "70-79(中等)", "60-69(及格)", "<60(不及格)" };
            Func<double, int> bandOf = v => v >= 90 ? 0 : v >= 80 ? 1 : v >= 70 ? 2 : v >= 60 ? 3 : 4;
            var bandCount = new int[5];
            foreach (var x in avg) bandCount[bandOf(x.S)]++;
            Color[] bandColors = { Color.SeaGreen, Color.MediumSeaGreen, Color.Goldenrod, Color.DarkOrange, Color.Firebrick };
            int maxB = Math.Max(1, bandCount.Max());
            for (int i = 0; i < 5; i++)
            {
                Bar(g, bands[i], bandCount[i], maxB, L, y, W, bandColors[i]); y += 26;
            }
            y += 16;

            // 重点关注
            g.DrawString("四、问题关注情况", new Font(Font.FontFamily, 11, FontStyle.Bold), Brushes.Black, L, y); y += 24;
            var withIssue = Program.Data.Recs.Where(r => r.Cat == "问题关注").Select(r => r.StudentId).Distinct().Count();
            g.DrawString("有“问题关注”记录的学生: " + withIssue + " 人 / 共 " + studs.Count + " 人", Font, Brushes.DimGray, L, y); y += 26;
            g.DrawString("(该数字不区分是否已解锁;具体名单请在解锁后到各学生“问题关注”页查看)", Font, Brushes.Gray, L, y);
        }

        void DrawChart(object o, PaintEventArgs e)
        {
            var g = e.Graphics; g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            bool byCourse = cur != null && courseFilter != null && courseFilter.SelectedIndex > 0;
            string course = byCourse ? courseFilter.Text : "";
            List<KeyValuePair<string, double>> data;
            if (cur == null) data = new List<KeyValuePair<string, double>>();
            else if (byCourse)
                data = Program.Data.Recs.Where(r => r.StudentId == cur.Id && r.Cat == "成绩" && r.Title == course && r.Date != "")
                    .Select(r => new { r.Date, V = Program.ParseScorePub(r.Detail) }).Where(x => x.V.HasValue)
                    .GroupBy(x => x.Date).OrderBy(gp => gp.Key)
                    .Select(gp => new KeyValuePair<string, double>(gp.Key, gp.First().V.Value)).ToList();
            else data = Program.ScoreSeries(cur.Id);
            if (data.Count == 0)
            {
                g.DrawString("暂无成绩数据。请在“成绩”标签添加(学期如 2023-2024-1,详情里填分数),或导入成绩CSV。", Font, Brushes.Gray, 20, 20); return;
            }
            int L = 60, T = 50, R = 30, B = 60;
            int w = chart.Width - L - R, h = chart.Height - T - B;
            if (w < 50 || h < 50) return;
            g.DrawString(cur.Name + (byCourse ? " 「" + course + "」" : " 各学期平均成绩"), new Font(Font.FontFamily, 12, FontStyle.Bold), Brushes.Black, L, 12);
            double min = Math.Max(0, Math.Floor(data.Min(x => x.Value) / 10) * 10 - 10), max = Math.Min(100, Math.Ceiling(data.Max(x => x.Value) / 10) * 10 + 5);
            if (max <= min) max = min + 10;
            for (double v = min; v <= max; v += 10)
            {
                float y = T + h - (float)((v - min) / (max - min) * h);
                g.DrawLine(Pens.LightGray, L, y, L + w, y);
                g.DrawString(v.ToString("0"), Font, Brushes.Gray, 20, y - 8);
            }
            var pts = new List<PointF>();
            for (int i = 0; i < data.Count; i++)
            {
                float x = data.Count == 1 ? L + w / 2 : L + (float)i / (data.Count - 1) * w;
                float y = T + h - (float)((data[i].Value - min) / (max - min) * h);
                pts.Add(new PointF(x, y));
            }
            if (pts.Count > 1) g.DrawLines(new Pen(Color.SteelBlue, 2.5f), pts.ToArray());
            for (int i = 0; i < pts.Count; i++)
            {
                g.FillEllipse(Brushes.SteelBlue, pts[i].X - 5, pts[i].Y - 5, 10, 10);
                g.DrawString(data[i].Value.ToString("0.#"), Font, Brushes.Firebrick, pts[i].X - 12, pts[i].Y - 24);
                g.DrawString(data[i].Key, Font, Brushes.Black, pts[i].X - 35, T + h + 8);
            }
        }

        void ImportPhotos()
        {
            var fb = new FolderBrowserDialog { Description = "选择照片文件夹(照片文件名为学号,如 2023001.jpg)" };
            if (fb.ShowDialog() != DialogResult.OK) return;
            int ok = 0, miss = 0;
            foreach (var f in Directory.GetFiles(fb.SelectedPath))
            {
                string ext = Path.GetExtension(f).ToLower();
                if (ext != ".jpg" && ext != ".jpeg" && ext != ".png" && ext != ".bmp") continue;
                string id = Path.GetFileNameWithoutExtension(f);
                var s = Program.Data.Students.FirstOrDefault(x => x.Id == id);
                if (s != null && Program.SetPhoto(s, f)) ok++; else miss++;
            }
            Program.Save(); Refill(); RefillAlerts(); Pick();
            MessageBox.Show("成功 " + ok + " 张,未匹配/失败 " + miss + " 张");
        }

        void MakeTemplates()
        {
            string dir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "导入模板");
            Directory.CreateDirectory(dir);
            var enc = new UTF8Encoding(true);
            File.WriteAllText(Path.Combine(dir, "学生名单模板.csv"), "学号,姓名,性别,专业班级,联系方式,籍贯,家庭情况,宿舍,宿舍长,标签\r\n2023001,张三,男,软件工程2301,13800000000,河南南阳,普通,3号楼301,李四,班干部\r\n", enc);
            File.WriteAllText(Path.Combine(dir, "成绩模板.csv"), "学号,学期,课程,成绩\r\n2023001,2023-2024-1,高等数学,86\r\n2023001,2023-2024-1,大学英语,78\r\n", enc);
            File.WriteAllText(Path.Combine(dir, "其他记录模板.csv"), "学号,类别,日期,标题,详情\r\n2023001,荣誉奖项,2024-10-01,国家励志奖学金,国家级\r\n2023001,任职实践,2024-09-01,班长,2023级软件1班\r\n", enc);
            System.Diagnostics.Process.Start(dir);
        }

        ListView alertList = new ListView();

        TabPage AlertTab()
        {
            var p = new TabPage("⚠ 预警看板");
            alertList.Dock = DockStyle.Fill; alertList.View = View.Details; alertList.FullRowSelect = true; alertList.GridLines = true;
            alertList.Columns.Add("学号", 85); alertList.Columns.Add("姓名", 70); alertList.Columns.Add("班级", 100);
            alertList.Columns.Add("预警原因", 260); alertList.Columns.Add("最后记录", 85);
            alertList.DoubleClick += delegate
            {
                if (alertList.SelectedItems.Count == 0) return;
                var s = (Student)alertList.SelectedItems[0].Tag;
                search.Text = s.Id; Refill();
                foreach (ListViewItem it in list.Items) if ((Student)it.Tag == s) { it.Selected = true; break; }
                tabs.SelectedIndex = 0;
            };
            p.Controls.Add(alertList); return p;
        }

        void RefillAlerts()
        {
            alertList.BeginUpdate(); alertList.Items.Clear();
            foreach (var s in Program.Data.Students.OrderBy(x => x.ClassName).ThenBy(x => x.Id))
            {
                var reasons = Program.AlertReasons(s.Id);
                if (reasons.Count == 0) continue;
                string last = "";
                var ld = Program.LastRecDate(s.Id);
                if (ld.HasValue) last = ld.Value.ToString("yyyy-MM-dd");
                var it = new ListViewItem(new[] { s.Id, s.Name, s.ClassName, string.Join("; ", reasons.ToArray()), last });
                it.Tag = s; it.ForeColor = Color.Firebrick;
                alertList.Items.Add(it);
            }
            alertList.EndUpdate();
        }

        TabPage TimelineTab()
        {
            var p = new TabPage("时间线(总览)");
            timeline.Dock = DockStyle.Fill; timeline.View = View.Details; timeline.FullRowSelect = true; timeline.GridLines = true;
            timeline.Columns.Add("日期", 100); timeline.Columns.Add("类别", 80); timeline.Columns.Add("标题", 220); timeline.Columns.Add("详情", 420);
            p.Controls.Add(timeline); return p;
        }

        TabPage CatTab(string cat)
        {
            bool sens = Program.Sensitive.Contains(cat);
            var p = new TabPage(sens ? cat + " 🔒" : cat);
            var lv = new ListView { Dock = DockStyle.Fill, View = View.Details, FullRowSelect = true, GridLines = true };
            if (cat == "问题关注") { lv.Columns.Add("状态", 60); lv.Columns.Add("责任人", 70); lv.Columns.Add("销号日期", 80); }
            lv.Columns.Add("日期/学期", 110); lv.Columns.Add("标题", 220); lv.Columns.Add("详情", 520);
            lv.DoubleClick += delegate { EditRec(cat); };
            var f = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 36 };
            f.Controls.Add(MkBtn("新增", delegate { AddRec(cat); }));
            f.Controls.Add(MkBtn("编辑", delegate { EditRec(cat); }));
            f.Controls.Add(MkBtn("删除", delegate { DelRec(cat); }));
            p.Controls.Add(lv); p.Controls.Add(f);
            catViews[cat] = lv; p.Tag = cat; return p;
        }

        void Refill()
        {
            string k = search.Text.Trim();
            list.BeginUpdate(); list.Items.Clear();
            foreach (var s in Program.Data.Students.OrderBy(x => x.ClassName).ThenBy(x => x.Id))
            {
                if (k != "" && !(s.Id + s.Name + s.ClassName + s.Tags + s.Dorm).Contains(k)) continue;
                string last = "";
                var ld = Program.LastRecDate(s.Id);
                if (ld.HasValue) last = ld.Value.ToString("yyyy-MM-dd");
                string alerts = string.Join("; ", Program.AlertReasons(s.Id).ToArray());
                var it = new ListViewItem(new[] { s.Id, s.Name, s.ClassName, s.Tags, last, alerts });
                it.Tag = s;
                if (alerts != "") it.ForeColor = Color.Firebrick;
                else if (last == "") it.ForeColor = Color.Gray;
                list.Items.Add(it);
            }
            list.EndUpdate();
        }

        void Pick()
        {
            cur = list.SelectedItems.Count > 0 ? (Student)list.SelectedItems[0].Tag : null;
            foreach (var key in fieldKeys)
                fields[key].Text = cur == null ? "" : (string)typeof(Student).GetField(key).GetValue(cur);
            Reload();
        }

        void Reload()
        {
            timeline.Items.Clear();
            ShowPhoto(); chart.Invalidate(); RefillCourses();
            foreach (var lv in catViews.Values) lv.Items.Clear();
            if (cur == null) return;
            var recs = Program.Data.Recs.Where(r => r.StudentId == cur.Id).OrderBy(r => r.Date).ToList();
            foreach (var r in recs)
            {
                bool sens = Program.Sensitive.Contains(r.Cat);
                if (sens && !Program.Unlocked) continue;
                var it = new ListViewItem(new[] { r.Date, r.Title, r.Detail }); it.Tag = r;
                if (catViews.ContainsKey(r.Cat))
                {
                    ListView lvCat = catViews[r.Cat];
                    if (r.Cat == "问题关注")
                    {
                        string st = string.IsNullOrEmpty(r.Status) ? "发现" : r.Status;
                        it = new ListViewItem(new[] { st, r.Owner, r.Closed, r.Date, r.Title, r.Detail }); it.Tag = r;
                        if (st != "已销号") it.ForeColor = Color.Firebrick;
                        else it.ForeColor = Color.Green;
                    }
                    lvCat.Items.Add(it);
                }
                var t = new ListViewItem(new[] { r.Date, r.Cat, r.Title, r.Detail });
                if (sens) t.ForeColor = Color.Firebrick;
                timeline.Items.Add(t);
            }
            if (Program.Unlocked && recs.Any(r => Program.Sensitive.Contains(r.Cat)))
                Program.AddLog("查看敏感信息: " + cur.Id + " " + cur.Name);
        }

        void SaveBasic()
        {
            if (cur == null) return;
            string id = fields["Id"].Text.Trim();
            if (id == "") { MessageBox.Show("学号不能为空"); return; }
            if (id != cur.Id && Program.Data.Students.Any(s => s.Id == id)) { MessageBox.Show("学号已存在"); return; }
            if (id != cur.Id) foreach (var r in Program.Data.Recs.Where(r => r.StudentId == cur.Id)) r.StudentId = id;
            foreach (var key in fieldKeys)
                typeof(Student).GetField(key).SetValue(cur, fields[key].Text.Trim());
            Program.Save(); Refill(); RefillAlerts(); status.Text = "已保存";
        }

        void NewStudent()
        {
            string id = InputDlg.Ask("新增学生", "请输入学号:", false);
            if (string.IsNullOrWhiteSpace(id)) return;
            id = id.Trim();
            if (Program.Data.Students.Any(s => s.Id == id)) { MessageBox.Show("学号已存在"); return; }
            string name = InputDlg.Ask("新增学生", "请输入姓名:", false);
            Program.Data.Students.Add(new Student { Id = id, Name = (name ?? "").Trim() });
            Program.Save(); search.Text = id; Refill();
            if (list.Items.Count > 0) list.Items[0].Selected = true;
        }

        void DelStudent()
        {
            if (cur == null) return;
            if (MessageBox.Show("确定删除 " + cur.Name + " 及其全部记录?(会先自动备份 data.json.bak)", "确认", MessageBoxButtons.YesNo) != DialogResult.Yes) return;
            Program.Data.Recs.RemoveAll(r => r.StudentId == cur.Id);
            Program.Data.Students.Remove(cur);
            Program.AddLog("删除学生 " + cur.Id + " " + cur.Name);
            Refill(); RefillAlerts(); Pick();
        }

        bool CheckAccess(string cat)
        {
            if (cur == null) { MessageBox.Show("请先选择学生"); return false; }
            if (Program.Sensitive.Contains(cat) && !Program.Unlocked) { MessageBox.Show("请先解锁敏感信息"); return false; }
            return true;
        }

        void AddRec(string cat)
        {
            if (!CheckAccess(cat)) return;
            var r = new Rec { StudentId = cur.Id, Cat = cat };
            using (var d = new RecDlg(cat, r))
                if (d.ShowDialog() == DialogResult.OK)
                {
                    r.Date = d.Date.Text.Trim(); r.Title = d.Title.Text.Trim(); r.Detail = d.Detail.Text.Trim();
                    if (cat == "问题关注") { r.Status = (string)d.StatusBox.SelectedItem; r.Owner = d.RespOwner.Text.Trim(); r.Closed = d.ClosedDate.Text.Trim(); }
                    Program.Data.Recs.Add(r); Program.Save();
                    if (Program.Sensitive.Contains(cat)) Program.AddLog("新增" + cat + ": " + cur.Id);
                    Reload();
                }
        }

        Rec Selected(string cat)
        {
            var lv = catViews[cat];
            return lv.SelectedItems.Count > 0 ? (Rec)lv.SelectedItems[0].Tag : null;
        }

        void EditRec(string cat)
        {
            if (!CheckAccess(cat)) return;
            var r = Selected(cat); if (r == null) return;
            using (var d = new RecDlg(cat, r))
                if (d.ShowDialog() == DialogResult.OK)
                {
                    r.Date = d.Date.Text.Trim(); r.Title = d.Title.Text.Trim(); r.Detail = d.Detail.Text.Trim();
                    if (cat == "问题关注") { r.Status = (string)d.StatusBox.SelectedItem; r.Owner = d.RespOwner.Text.Trim(); r.Closed = d.ClosedDate.Text.Trim(); }
                    Program.Save();
                    if (Program.Sensitive.Contains(cat)) Program.AddLog("修改" + cat + ": " + cur.Id);
                    Reload();
                }
        }

        void DelRec(string cat)
        {
            if (!CheckAccess(cat)) return;
            var r = Selected(cat); if (r == null) return;
            if (MessageBox.Show("删除这条记录?", "确认", MessageBoxButtons.YesNo) != DialogResult.Yes) return;
            Program.Data.Recs.Remove(r); Program.Save();
            if (Program.Sensitive.Contains(cat)) Program.AddLog("删除" + cat + ": " + cur.Id);
            Reload();
        }

        void UpdateLock()
        {
            lockBtn.Text = Program.Unlocked ? "🔓 已解锁(点击锁定)" : "🔒 解锁敏感信息";
            status.Text = Program.Unlocked ? "敏感信息可见,所有查看均有日志" : "敏感信息已隐藏";
        }

        void ToggleLock()
        {
            if (Program.Unlocked) { Program.Unlocked = false; UpdateLock(); Reload(); return; }
            if (Program.Data.Hash == "") { SetPwd(); if (Program.Data.Hash == "") return; }
            string p = InputDlg.Ask("解锁", "请输入敏感信息密码:", true);
            if (p == null) return;
            if (Program.HashPwd(p, Program.Data.Salt) != Program.Data.Hash)
            {
                Program.AddLog("解锁失败"); MessageBox.Show("密码错误"); return;
            }
            Program.Unlocked = true; Program.AddLog("解锁敏感信息"); UpdateLock(); Reload();
        }

        void SetPwd()
        {
            if (Program.Data.Hash != "")
            {
                string old = InputDlg.Ask("修改密码", "请输入当前密码:", true);
                if (old == null || Program.HashPwd(old, Program.Data.Salt) != Program.Data.Hash) { MessageBox.Show("当前密码错误"); return; }
            }
            string p1 = InputDlg.Ask("设置密码", "请输入新密码(至少6位):", true);
            if (p1 == null) return;
            if (p1.Length < 6) { MessageBox.Show("密码太短"); return; }
            string p2 = InputDlg.Ask("设置密码", "再输入一次:", true);
            if (p1 != p2) { MessageBox.Show("两次不一致"); return; }
            var salt = new byte[16];
            using (var rng = new RNGCryptoServiceProvider()) rng.GetBytes(salt);
            Program.Data.Salt = Convert.ToBase64String(salt);
            Program.Data.Hash = Program.HashPwd(p1, Program.Data.Salt);
            Program.AddLog("设置/修改密码");
            MessageBox.Show("密码已设置。请牢记,忘记无法找回。");
        }

        void ShowLog()
        {
            var f = new Form { Text = "访问日志", Width = 720, Height = 480, StartPosition = FormStartPosition.CenterParent };
            var tb = new TextBox { Dock = DockStyle.Fill, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Both, WordWrap = false };
            tb.Text = string.Join("\r\n", Enumerable.Reverse(Program.Data.Log).Take(500).ToArray());
            f.Controls.Add(tb); f.ShowDialog();
        }

        static string[] ParseCsv(string line)
        {
            var res = new List<string>(); var sb = new StringBuilder(); bool q = false;
            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];
                if (q) { if (c == '"') { if (i + 1 < line.Length && line[i + 1] == '"') { sb.Append('"'); i++; } else q = false; } else sb.Append(c); }
                else if (c == '"') q = true;
                else if (c == ',') { res.Add(sb.ToString()); sb.Length = 0; }
                else sb.Append(c);
            }
            res.Add(sb.ToString()); return res.ToArray();
        }

        void ImportCsv(int kind)
        {
            string[] tip = {
                "列顺序:学号,姓名,性别,专业班级,联系方式,籍贯,家庭情况,宿舍,宿舍长,标签",
                "列顺序:学号,学期,课程,成绩",
                "列顺序:学号,类别,日期,标题,详情\r\n类别可填:荣誉奖项/任职实践/宿舍生活/问题关注/辅导员记录(后两类需先解锁)" };
            if (MessageBox.Show(tip[kind] + "\r\n\r\n第一行为表头;Excel 另存为 CSV 即可(UTF-8 或默认格式均可)。\r\n可先用“生成CSV导入模板”。", "导入说明", MessageBoxButtons.OKCancel) != DialogResult.OK) return;
            var dlg = new OpenFileDialog { Filter = "CSV|*.csv" };
            if (dlg.ShowDialog() != DialogResult.OK) return;
            int added = 0, skipped = 0;
            var lines = Program.ReadText(dlg.FileName).Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
            for (int i = 1; i < lines.Length; i++)
            {
                if (lines[i].Trim() == "") continue;
                var c = ParseCsv(lines[i]);
                Func<int, string> g = n => n < c.Length ? c[n].Trim() : "";
                if (kind == 0)
                {
                    if (g(0) == "" || Program.Data.Students.Any(s => s.Id == g(0))) { skipped++; continue; }
                    Program.Data.Students.Add(new Student { Id = g(0), Name = g(1), Gender = g(2), ClassName = g(3), Phone = g(4), Origin = g(5), Family = g(6), Dorm = g(7), DormLeader = g(8), Tags = g(9) });
                }
                else if (kind == 2)
                {
                    if (!Program.Data.Students.Any(s => s.Id == g(0)) || !Program.Cats.Contains(g(1))
                        || (Program.Sensitive.Contains(g(1)) && !Program.Unlocked)) { skipped++; continue; }
                    Program.Data.Recs.Add(new Rec { StudentId = g(0), Cat = g(1), Date = g(2), Title = g(3), Detail = g(4) });
                    if (Program.Sensitive.Contains(g(1))) Program.AddLog("导入" + g(1) + ": " + g(0));
                }
                else
                {
                    if (!Program.Data.Students.Any(s => s.Id == g(0))) { skipped++; continue; }
                    Program.Data.Recs.Add(new Rec { StudentId = g(0), Cat = "成绩", Date = g(1), Title = g(2), Detail = g(3) });
                }
                added++;
            }
            Program.Save(); Refill(); RefillAlerts(); Pick();
            MessageBox.Show("导入 " + added + " 条,跳过 " + skipped + " 条(学号为空/重复/不存在)");
        }

        void ExportCard(bool print)
        {
            if (cur == null) return;
            bool withSens = Program.Unlocked && MessageBox.Show("画像卡是否包含敏感信息(问题关注、辅导员记录)?\r\n给领导/同事看建议选“否”。", "导出", MessageBoxButtons.YesNo) == DialogResult.Yes;
            var sb = new StringBuilder();
            sb.Append("<html><head><meta charset='utf-8'><style>@media print{.noprint{display:none}}body{font-family:Microsoft YaHei;margin:30px;}table{border-collapse:collapse;width:100%}td,th{border:1px solid #ccc;padding:4px 8px;text-align:left}h2{border-left:5px solid #2a6;padding-left:8px}</style></head><body>");
            sb.Append("<div class='noprint' style='margin-bottom:10px'><button onclick='window.print()' style='font-size:14px;padding:6px 16px'>🖨 打印</button></div>");
            var ph = Program.LoadPhoto(cur);
            if (ph != null)
                using (var ms = new MemoryStream())
                {
                    ph.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                    sb.Append("<img src='data:image/jpeg;base64," + Convert.ToBase64String(ms.ToArray()) + "' style='float:right;height:160px;border:1px solid #ccc'>");
                }
            sb.Append("<h1>" + H(cur.Name) + " <small>" + H(cur.Id) + "</small></h1>");
            var ser = Program.ScoreSeries(cur.Id);
            if (ser.Count > 0)
            {
                var svg = new StringBuilder("<h2>成绩曲线(各学期平均分)</h2><svg width='640' height='220' style='border:1px solid #eee'>");
                var pl = new List<string>();
                for (int i = 0; i < ser.Count; i++)
                {
                    double x = ser.Count == 1 ? 320 : 40 + i * 560.0 / (ser.Count - 1), y = 190 - ser[i].Value / 100 * 160;
                    pl.Add(x.ToString("0.#") + "," + y.ToString("0.#"));
                    svg.Append("<circle cx='" + x.ToString("0.#") + "' cy='" + y.ToString("0.#") + "' r='4' fill='steelblue'/><text x='" + (x - 12).ToString("0.#") + "' y='" + (y - 10).ToString("0.#") + "' font-size='12'>" + ser[i].Value + "</text><text x='" + (x - 30).ToString("0.#") + "' y='212' font-size='11'>" + H(ser[i].Key) + "</text>");
                }
                svg.Insert(svg.ToString().IndexOf("<circle"), "<polyline fill='none' stroke='steelblue' stroke-width='2' points='" + string.Join(" ", pl.ToArray()) + "'/>");
                sb.Append(svg.ToString() + "</svg>");
            }
            sb.Append("<p>班级:" + H(cur.ClassName) + " | 性别:" + H(cur.Gender) + " | 籍贯:" + H(cur.Origin) + " | 宿舍:" + H(cur.Dorm) + " | 宿舍长:" + H(cur.DormLeader) + "</p>");
            sb.Append("<p>标签:" + H(cur.Tags) + "</p>");
            foreach (var cat in Program.Cats)
            {
                if (Program.Sensitive.Contains(cat) && !withSens) continue;
                var rs = Program.Data.Recs.Where(r => r.StudentId == cur.Id && r.Cat == cat).OrderBy(r => r.Date).ToList();
                if (rs.Count == 0) continue;
                sb.Append("<h2>" + cat + "</h2><table><tr><th>日期/学期</th><th>标题</th><th>详情</th></tr>");
                foreach (var r in rs) sb.Append("<tr><td>" + H(r.Date) + "</td><td>" + H(r.Title) + "</td><td>" + H(r.Detail) + "</td></tr>");
                sb.Append("</table>");
            }
            sb.Append("<p style='color:#888'>导出时间:" + DateTime.Now.ToString("yyyy-MM-dd HH:mm") + (withSens ? " (含敏感信息,请妥善保管)" : "") + "</p></body></html>");
            string dir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "导出");
            Directory.CreateDirectory(dir);
            string path = Path.Combine(dir, cur.Id + "_" + cur.Name + ".html");
            File.WriteAllText(path, sb.ToString(), Encoding.UTF8);
            Program.AddLog("导出画像卡: " + cur.Id + (withSens ? " (含敏感)" : "") + (print ? " (打印)" : ""));
            System.Diagnostics.Process.Start(path);
        }

        void ShowDormMates()
        {
            if (cur == null || string.IsNullOrEmpty(cur.Dorm)) { MessageBox.Show("请先选择学生且该生已填宿舍信息"); return; }
            var mates = Program.Data.Students.Where(s => s.Dorm == cur.Dorm && s.Id != cur.Id).OrderBy(s => s.Id).ToList();
            var f = new Form { Text = "宿舍 " + cur.Dorm + " 成员(" + (mates.Count + 1) + " 人,含宿舍长 " + (string.IsNullOrEmpty(cur.DormLeader) ? "-" : cur.DormLeader) + ")", Width = 560, Height = 420, StartPosition = FormStartPosition.CenterParent };
            var lv = new ListView { Dock = DockStyle.Fill, View = View.Details, FullRowSelect = true, GridLines = true };
            lv.Columns.Add("学号", 90); lv.Columns.Add("姓名", 80); lv.Columns.Add("班级", 130); lv.Columns.Add("标签", 160);
            lv.Items.Add(MakeRow(cur));
            foreach (var m in mates) lv.Items.Add(MakeRow(m));
            var tip = new Label { Dock = DockStyle.Bottom, Height = 28, Text = "双击某成员可切换到该学生档案", Padding = new Padding(6, 4, 0, 0) };
            lv.DoubleClick += delegate
            {
                if (lv.SelectedItems.Count == 0) return;
                var s = (Student)lv.SelectedItems[0].Tag;
                f.Close();
                search.Text = s.Id; Refill();
                foreach (ListViewItem it in list.Items) if ((Student)it.Tag == s) { it.Selected = true; break; }
            };
            f.Controls.Add(lv); f.Controls.Add(tip);
            f.ShowDialog();
        }

        ListViewItem MakeRow(Student s)
        {
            var it = new ListViewItem(new[] { s.Id, s.Name, s.ClassName, s.Tags }); it.Tag = s;
            if (cur != null && s.Id == cur.Id) it.ForeColor = Color.SteelBlue;
            return it;
        }

        void ExportClassReport()
        {
            var classes = Program.Data.Students.Select(s => string.IsNullOrEmpty(s.ClassName) ? "(未填班级)" : s.ClassName).Distinct().OrderBy(x => x).ToList();
            if (classes.Count == 0) { MessageBox.Show("暂无学生数据"); return; }
            string cls = InputDlg.Ask("班级汇总报告", "请输入班级名称(留空 = 全部学生汇总):", false);
            if (cls == null) return;
            var studs = cls == "" ? Program.Data.Students : Program.Data.Students.Where(s => s.ClassName == cls).ToList();
            if (studs.Count == 0) { MessageBox.Show("没有找到该班级的学生"); return; }
            var sb = new StringBuilder();
            sb.Append("<html><head><meta charset='utf-8'><style>body{font-family:Microsoft YaHei;margin:30px}table{border-collapse:collapse;width:100%}td,th{border:1px solid #ccc;padding:4px 8px;text-align:left}h2{border-left:5px solid #2a6;padding-left:8px}.warn{color:#b22}</style></head><body>");
            sb.Append("<h1>" + H(cls == "" ? "全部学生" : cls) + " 班级汇总报告</h1>");
            sb.Append("<p>生成时间:" + DateTime.Now.ToString("yyyy-MM-dd HH:mm") + " | 学生数:" + studs.Count + "</p>");
            // 成绩分布
            var avg = studs.Select(s => new {
                S = Program.Data.Recs.Where(r => r.StudentId == s.Id && r.Cat == "成绩")
                    .Select(r => Regex.Match(r.Detail ?? "", @"\d+(\.\d+)?")).Where(m => m.Success)
                    .Select(m => { double v; return double.TryParse(m.Value, out v) ? v : (double?)null; })
                    .Where(v => v.HasValue).Select(v => v.Value).DefaultIfEmpty(-1).Average()
            }).Where(x => x.S >= 0).Select(x => x.S).ToList();
            sb.Append("<h2>一、成绩水平分布(按每人所有成绩平均分,共 " + avg.Count + " 人有成绩)</h2><table><tr><th>分数段</th><th>人数</th></tr>");
            string[] bands = { "90-100(优秀)", "80-89(良好)", "70-79(中等)", "60-69(及格)", "<60(不及格)" };
            Func<double, int> bandOf = v => v >= 90 ? 0 : v >= 80 ? 1 : v >= 70 ? 2 : v >= 60 ? 3 : 4;
            var bc = new int[5];
            foreach (var v in avg) bc[bandOf(v)]++;
            for (int i = 0; i < 5; i++) sb.Append("<tr><td>" + bands[i] + "</td><td>" + bc[i] + "</td></tr>");
            sb.Append("</table>");
            // 获奖统计
            sb.Append("<h2>二、荣誉奖项(共 " + Program.Data.Recs.Count(r => r.Cat == "荣誉奖项" && studs.Any(s => s.Id == r.StudentId)) + " 条)</h2><table><tr><th>学号</th><th>姓名</th><th>日期</th><th>奖项</th><th>详情</th></tr>");
            foreach (var r in Program.Data.Recs.Where(r => r.Cat == "荣誉奖项" && studs.Any(s => s.Id == r.StudentId)).OrderBy(r => r.StudentId).ThenBy(r => r.Date))
            {
                var st = studs.First(s => s.Id == r.StudentId);
                sb.Append("<tr><td>" + H(st.Id) + "</td><td>" + H(st.Name) + "</td><td>" + H(r.Date) + "</td><td>" + H(r.Title) + "</td><td>" + H(r.Detail) + "</td></tr>");
            }
            sb.Append("</table>");
            // 预警名单
            var alerts = studs.Select(s => new { S = s, R = Program.AlertReasons(s.Id) }).Where(x => x.R.Count > 0).ToList();
            sb.Append("<h2>三、预警名单(共 " + alerts.Count + " 人,不含敏感记录内容)</h2>");
            if (alerts.Count == 0) sb.Append("<p>无预警学生。</p>");
            else
            {
                sb.Append("<table><tr><th>学号</th><th>姓名</th><th>班级</th><th>预警原因</th></tr>");
                foreach (var x in alerts)
                    sb.Append("<tr class='warn'><td>" + H(x.S.Id) + "</td><td>" + H(x.S.Name) + "</td><td>" + H(x.S.ClassName) + "</td><td>" + H(string.Join("; ", x.R.ToArray())) + "</td></tr>");
                sb.Append("</table><p style='color:#888'>具体处理情况请在本机系统内解锁查看,本报告不含问题关注与辅导员记录内容。</p>");
            }
            sb.Append("<h2>四、学生名册</h2><table><tr><th>学号</th><th>姓名</th><th>性别</th><th>宿舍</th><th>宿舍长</th><th>标签</th></tr>");
            foreach (var s in studs.OrderBy(x => x.Id))
                sb.Append("<tr><td>" + H(s.Id) + "</td><td>" + H(s.Name) + "</td><td>" + H(s.Gender) + "</td><td>" + H(s.Dorm) + "</td><td>" + H(s.DormLeader) + "</td><td>" + H(s.Tags) + "</td></tr>");
            sb.Append("</table></body></html>");
            string dir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "导出");
            Directory.CreateDirectory(dir);
            string path = Path.Combine(dir, "班级报告_" + (cls == "" ? "全部学生" : cls) + "_" + DateTime.Now.ToString("yyyyMMdd_HHmm") + ".html");
            File.WriteAllText(path, sb.ToString(), Encoding.UTF8);
            Program.AddLog("导出班级报告: " + (cls == "" ? "全部" : cls));
            System.Diagnostics.Process.Start(path);
        }

        void BackupAll()
        {
            string backupRoot;
            using (var fb = new FolderBrowserDialog { Description = "选择备份目标文件夹(将生成 JSON 数据 + CSV 学生/成绩/其他记录 + 照片)" })
            {
                if (fb.ShowDialog() != DialogResult.OK) return;
                backupRoot = fb.SelectedPath;
            }
            try
            {
                string stamp = DateTime.Now.ToString("yyyyMMdd_HHmm");
                string dir = Path.Combine(backupRoot, "档案备份_" + stamp);
                Directory.CreateDirectory(dir);
                File.Copy(Program.DataFile, Path.Combine(dir, "data.json"), true);
                var enc = new UTF8Encoding(true);
                File.WriteAllText(Path.Combine(dir, "学生.csv"), "学号,姓名,性别,专业班级,联系方式,籍贯,家庭情况,宿舍,宿舍长,标签\r\n" +
                    string.Join("", Program.Data.Students.Select(s => CsvLine(s.Id) + "," + CsvLine(s.Name) + "," + CsvLine(s.Gender) + "," + CsvLine(s.ClassName) + "," + CsvLine(s.Phone) + "," + CsvLine(s.Origin) + "," + CsvLine(s.Family) + "," + CsvLine(s.Dorm) + "," + CsvLine(s.DormLeader) + "," + CsvLine(s.Tags) + "\r\n").ToArray()), enc);
                File.WriteAllText(Path.Combine(dir, "记录.csv"), "学号,类别,日期,标题,详情,状态,责任人,销号日期\r\n" +
                    string.Join("", Program.Data.Recs.Select(r => CsvLine(r.StudentId) + "," + CsvLine(r.Cat) + "," + CsvLine(r.Date) + "," + CsvLine(r.Title) + "," + CsvLine(r.Detail) + "," + CsvLine(r.Status) + "," + CsvLine(r.Owner) + "," + CsvLine(r.Closed) + "\r\n").ToArray()), enc);
                if (Directory.Exists(Program.PhotoDir))
                    foreach (var f in Directory.GetFiles(Program.PhotoDir))
                        File.Copy(f, Path.Combine(dir, "照片_" + Path.GetFileName(f)), true);
                Program.AddLog("备份全部数据到 " + dir);
                MessageBox.Show("备份完成:" + dir);
            }
            catch (Exception ex) { MessageBox.Show("备份失败:" + ex.Message); }
        }

        static string CsvLine(string s)
        {
            s = s ?? "";
            if (s.Contains(",") || s.Contains("\"") || s.Contains("\n") || s.Contains("\r"))
                return "\"" + s.Replace("\"", "\"\"") + "\"";
            return s;
        }

        static string H(string s) { return System.Net.WebUtility.HtmlEncode(s ?? ""); }
    }
}
