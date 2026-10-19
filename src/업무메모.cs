// 업무메모 - 업무지시 입력 / 결재 미마감 찾기·완료 프로그램
// 빌드: .NET Framework 4.8 (Windows 10/11 기본 탑재) WinForms
//
// 데이터 파일 (모두 UTF-8 BOM, 업무지시 폴더 안)
//   업무지시.txt          메모 업무. 한 줄 = 한 건. [ ] 진행중 / [v] 완료
//   완료기록.txt          결재(엑셀) 항목 완료 기록. 시각<TAB>키<TAB>내용
//   입력창_오류기록.txt   오류 기록
//   결재_미마감_업무지시_*.xlsx  15:48 자동 작업이 만드는 파일. 가장 최근 파일의 첫 시트를 읽기만 함

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using System.Xml;

[assembly: AssemblyTitle("업무메모")]
[assembly: AssemblyProduct("업무메모")]
[assembly: AssemblyDescription("업무지시 입력 및 결재 미마감 완료 처리")]
[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion("1.0.0.0")]

namespace FmjMemo
{
    public class Item
    {
        public string Source;   // "결재" 또는 "메모"
        public string Key;
        public string Text;
        public string Display;
        public string Seq = "";      // 결재 엑셀의 줄 순번 (찾기 칸에 번호를 넣으면 이 값으로 찾음)
        public string Detail = "";   // 찾기 칸에 번호를 넣었을 때 보여줄 전체 내용 (여러 줄)
    }

    public class ExcelResult
    {
        public string FileName;
        public List<Item> Items = new List<Item>();
        public int Hidden;
    }

    // 파일 읽기/쓰기 (화면과 분리)
    public static class Store
    {
        public static string BaseDir, DataFile, DoneFile, LogFile, SettingFile;
        public const string ExcelMask = "결재_미마감_업무지시_*.xlsx";
        static readonly Encoding Utf8Bom = new UTF8Encoding(true);
        public static readonly string[] Header = {
            "# 업무지시 목록 - 입력창이 자동으로 관리합니다.",
            "# [ ] = 진행중, [v] = 완료.  한 줄이 한 건입니다. 메모장에서 직접 고쳐도 됩니다.",
            ""
        };

        public static void Init(string dir)
        {
            BaseDir = dir;
            DataFile = Path.Combine(dir, "업무지시.txt");
            DoneFile = Path.Combine(dir, "완료기록.txt");
            LogFile = Path.Combine(dir, "입력창_오류기록.txt");
            SettingFile = Path.Combine(dir, "업무메모_설정.txt");
        }

        // exe가 있는 폴더에 업무 파일이 있으면 거기를, 없으면 바탕화면\업무지시를 씀
        public static string PickDir(string exeDir, string desktopDir)
        {
            var candidates = new List<string> { exeDir };
            if (!string.IsNullOrEmpty(desktopDir)) candidates.Add(Path.Combine(desktopDir, "업무지시"));
            foreach (var d in candidates)
            {
                if (!Directory.Exists(d)) continue;
                if (File.Exists(Path.Combine(d, "업무지시.txt")) || File.Exists(Path.Combine(d, "task_input.ps1"))
                    || Directory.GetFiles(d, ExcelMask).Length > 0)
                    return d;
            }
            return exeDir;
        }

        public static void Log(string msg)
        {
            try { File.AppendAllText(LogFile, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "  " + msg + "\r\n", Utf8Bom); }
            catch { }
        }

        public static string Clean(string s)
        {
            if (s == null) return "";
            s = Regex.Replace(s, @"\s*\r?\n\s*", " / ");
            return s.Replace('\t', ' ').Trim();
        }

        static void Retry(Action a)
        {
            Exception last = null;
            for (int i = 0; i < 5; i++)
            {
                try { a(); return; }
                catch (IOException e) { last = e; }
                catch (UnauthorizedAccessException e) { last = e; }
                Thread.Sleep(100);
            }
            throw last;
        }

        static string[] ReadAllLinesShared(string path)
        {
            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            using (var r = new StreamReader(fs, Utf8Bom, true))
            {
                var list = new List<string>();
                string line;
                while ((line = r.ReadLine()) != null) list.Add(line);
                return list.ToArray();
            }
        }

        static string[] ReadLines()
        {
            if (!File.Exists(DataFile)) return (string[])Header.Clone();
            return ReadAllLinesShared(DataFile);
        }

        public static void SaveLines(string[] lines)
        {
            string tmp = DataFile + ".tmp";
            File.WriteAllLines(tmp, lines, Utf8Bom);
            Retry(() => { File.Copy(tmp, DataFile, true); File.Delete(tmp); });
        }

        static void AppendText(string path, string text)
        {
            if (File.Exists(path))
            {
                using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
                {
                    if (fs.Length > 0)
                    {
                        fs.Seek(-1, SeekOrigin.End);
                        if (fs.ReadByte() != 10) text = "\r\n" + text;
                    }
                }
            }
            Retry(() => File.AppendAllText(path, text, Utf8Bom));
        }

        public static string[] OpenTasks()
        {
            return ReadLines().Where(l => l.StartsWith("[ ] ")).ToArray();
        }

        public static string NewTaskLine(string content, string owner, DateTime? due, bool urgent)
        {
            string text = Clean(content);
            if (urgent) text = "[긴급] " + text;
            string o = (owner ?? "").Trim(); if (o.Length == 0) o = "-";
            string d = due.HasValue ? due.Value.ToString("yyyy-MM-dd") : "-";
            return string.Format("[ ] 등록 {0} | 마감 {1} | 담당 {2} | {3}", DateTime.Now.ToString("yyyy-MM-dd HH:mm"), d, o, text);
        }

        public static void AddTask(string line)
        {
            string text = line + "\r\n";
            if (!File.Exists(DataFile)) text = string.Join("\r\n", Header) + "\r\n" + text;
            AppendText(DataFile, text);
        }

        public static bool CompleteTask(string line)
        {
            var lines = ReadLines();
            int idx = Array.IndexOf(lines, line);
            if (idx < 0) return false;   // 그 사이 메모장 등에서 바뀐 경우
            lines[idx] = "[v]" + line.Substring(3) + " | 완료 " + DateTime.Now.ToString("yyyy-MM-dd HH:mm");
            SaveLines(lines);
            return true;
        }

        public static void EnsureDataFile()
        {
            if (!File.Exists(DataFile)) SaveLines((string[])Header.Clone());
        }

        // ---------- 결재 엑셀 (읽기 전용) ----------

        public static string LatestExcel()
        {
            if (!Directory.Exists(BaseDir)) return null;
            return Directory.GetFiles(BaseDir, ExcelMask)
                .Where(f => !Path.GetFileName(f).StartsWith("~$") && f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                .OrderByDescending(f => File.GetLastWriteTime(f))
                .FirstOrDefault();
        }

        const string L = "*[local-name()='{0}']";
        static string P(params string[] names) { return string.Join("/", names.Select(n => string.Format(L, n))); }

        static XmlDocument LoadXml(ZipArchive zip, string name)
        {
            var e = zip.GetEntry(name);
            if (e == null) return null;
            var doc = new XmlDocument();
            using (var s = e.Open()) doc.Load(s);
            return doc;
        }

        public static List<string[]> ReadXlsxRows(string path)
        {
            var rows = new List<string[]>();
            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            using (var zip = new ZipArchive(fs, ZipArchiveMode.Read))
            {
                var strings = new List<string>();
                var ss = LoadXml(zip, "xl/sharedStrings.xml");
                if (ss != null)
                {
                    foreach (XmlNode si in ss.DocumentElement.SelectNodes(P("si")))
                    {
                        var sb = new StringBuilder();
                        foreach (XmlNode t in si.SelectNodes(P("t") + " | " + P("r", "t"))) sb.Append(t.InnerText);
                        strings.Add(sb.ToString());
                    }
                }

                // 날짜 서식이 걸린 칸 (엑셀은 날짜를 숫자로 저장함)
                var isDate = new List<bool>();
                var st = LoadXml(zip, "xl/styles.xml");
                if (st != null)
                {
                    var custom = new Dictionary<int, string>();
                    foreach (XmlElement nf in st.DocumentElement.SelectNodes(P("numFmts", "numFmt")))
                    {
                        int id;
                        if (int.TryParse(nf.GetAttribute("numFmtId"), out id)) custom[id] = nf.GetAttribute("formatCode");
                    }
                    foreach (XmlElement xf in st.DocumentElement.SelectNodes(P("cellXfs", "xf")))
                    {
                        int id; int.TryParse(xf.GetAttribute("numFmtId"), out id);
                        bool d = (id >= 14 && id <= 22) || (id >= 45 && id <= 47);
                        string code;
                        if (!d && custom.TryGetValue(id, out code))
                        {
                            code = Regex.Replace(code, "\"[^\"]*\"", "");
                            code = Regex.Replace(code, @"\[[^\]]*\]", "");
                            d = Regex.IsMatch(code, "[ymdhs]", RegexOptions.IgnoreCase);
                        }
                        isDate.Add(d);
                    }
                }

                var wb = LoadXml(zip, "xl/workbook.xml");
                var first = (XmlElement)wb.DocumentElement.SelectSingleNode(P("sheets", "sheet"));
                string rid = first.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                var rels = LoadXml(zip, "xl/_rels/workbook.xml.rels");
                string target = null;
                foreach (XmlElement rel in rels.DocumentElement.SelectNodes(P("Relationship")))
                    if (rel.GetAttribute("Id") == rid) { target = rel.GetAttribute("Target"); break; }
                string sheetPath = target.StartsWith("/") ? target.Substring(1) : "xl/" + target;
                var sh = LoadXml(zip, sheetPath);

                foreach (XmlElement row in sh.DocumentElement.SelectNodes(P("sheetData", "row")))
                {
                    var cells = new Dictionary<int, string>();
                    int max = -1, auto = 0;
                    foreach (XmlElement c in row.SelectNodes(P("c")))
                    {
                        int col = auto;
                        var m = Regex.Match(c.GetAttribute("r"), "^([A-Z]+)");
                        if (m.Success)
                        {
                            col = 0;
                            foreach (char ch in m.Groups[1].Value) col = col * 26 + (ch - 'A' + 1);
                            col--;
                        }
                        auto = col + 1;
                        var vNode = c.SelectSingleNode(P("v"));
                        string v = vNode != null ? vNode.InnerText : null;
                        string t = c.GetAttribute("t");
                        string val = "";
                        int si;
                        if (t == "s") { if (int.TryParse(v, out si) && si >= 0 && si < strings.Count) val = strings[si]; }
                        else if (t == "inlineStr") { var n = c.SelectSingleNode(P("is")); if (n != null) val = n.InnerText; }
                        else if (t == "b") val = v == "1" ? "TRUE" : "FALSE";
                        else if (!string.IsNullOrEmpty(v))
                        {
                            val = v;
                            int s; double num;
                            if (int.TryParse(c.GetAttribute("s"), out s) && s >= 0 && s < isDate.Count && isDate[s]
                                && double.TryParse(v, NumberStyles.Float, CultureInfo.InvariantCulture, out num))
                            {
                                try
                                {
                                    var dt = DateTime.FromOADate(num);
                                    val = dt.TimeOfDay.TotalSeconds >= 1 ? dt.ToString("yyyy-MM-dd HH:mm") : dt.ToString("yyyy-MM-dd");
                                }
                                catch { }
                            }
                        }
                        cells[col] = Clean(val);
                        if (col > max) max = col;
                    }
                    if (max < 0) continue;
                    var arr = new string[max + 1];
                    for (int k = 0; k <= max; k++) { string x; arr[k] = cells.TryGetValue(k, out x) ? x : ""; }
                    if (arr.Any(x => x.Length > 0)) rows.Add(arr);
                }
            }
            return rows;
        }

        public static HashSet<string> DoneKeys()
        {
            var set = new HashSet<string>();
            if (File.Exists(DoneFile))
                foreach (var l in ReadAllLinesShared(DoneFile))
                {
                    var p = l.Split('\t');
                    if (p.Length >= 2 && p[1].Length > 0) set.Add(p[1]);
                }
            return set;
        }

        public static ExcelResult ExcelItems()
        {
            var res = new ExcelResult();
            string file = LatestExcel();
            if (file == null) return res;
            res.FileName = Path.GetFileName(file);
            var rows = ReadXlsxRows(file);
            if (rows.Count < 2) return res;
            var head = rows[0];
            // '번호/ID/문서/코드'가 들어간 칸이 있으면 그 칸 값으로 같은 건인지 구분, 없으면 줄 전체로 구분
            // 단, 1,2,3…처럼 줄 순번인 칸은 매일 바뀔 수 있으므로 구분 기준에서 뺌
            int seqCol = SequenceColumn(rows);
            int keyCol = -1;
            for (int k = 0; k < head.Length; k++)
                if (k != seqCol && Regex.IsMatch(head[k], "번호|ID|문서|코드")) { keyCol = k; break; }
            var done = DoneKeys();
            for (int r = 1; r < rows.Count; r++)
            {
                var row = rows[r];
                string text = string.Join(" | ", row.Where(x => x.Length > 0));
                string key = string.Join(" | ", row.Where((x, i) => i != seqCol && x.Length > 0));
                if (keyCol >= 0 && keyCol < row.Length && row[keyCol].Length > 0) key = head[keyCol] + "=" + row[keyCol];
                if (done.Contains(key)) { res.Hidden++; continue; }
                var detail = new StringBuilder();
                for (int k = 0; k < row.Length; k++)
                {
                    if (row[k].Length == 0) continue;
                    string name = k < head.Length && head[k].Length > 0 ? head[k] : (k + 1) + "열";
                    detail.Append(name).Append(" : ").Append(row[k].Replace(" / ", "\r\n    ")).Append("\r\n");
                }
                // 순번 칸이 없으면 엑셀 줄 순서(1,2,3…)를 번호로 씀. 목록 앞에도 번호를 붙여 찾기 칸에 넣을 번호를 알 수 있게 함
                string seq = seqCol >= 0 && seqCol < row.Length ? row[seqCol] : r.ToString();
                string shownText = seqCol >= 0 ? text : seq + " | " + text;
                res.Items.Add(new Item { Source = "결재", Key = key, Text = text, Display = "[결재] " + shownText,
                    Seq = seq, Detail = detail.ToString() });
            }
            return res;
        }

        // 모든 값이 정수이고 위에서부터 계속 커지는 칸 = 줄 순번 칸. 없으면 -1
        public static int SequenceColumn(List<string[]> rows)
        {
            if (rows.Count < 2) return -1;
            int cols = rows.Max(r => r.Length);
            for (int k = 0; k < cols; k++)
            {
                int prev = int.MinValue; bool ok = true;
                for (int r = 1; r < rows.Count && ok; r++)
                {
                    int v;
                    if (k >= rows[r].Length || !int.TryParse(rows[r][k], out v) || v <= prev) ok = false;
                    else prev = v;
                }
                if (ok) return k;
            }
            return -1;
        }

        public static void CompleteExcel(Item it)
        {
            AppendText(DoneFile, DateTime.Now.ToString("yyyy-MM-dd HH:mm") + "\t" + it.Key + "\t" + it.Text + "\r\n");
        }

        public static Item MemoItem(string line)
        {
            return new Item { Source = "메모", Key = line, Text = line, Display = "[메모] " + line.Substring(4),
                Detail = line.Substring(4).Replace(" | ", "\r\n").Replace(" / ", "\r\n    ") };
        }
    }

    public class MainForm : Form
    {
        TextBox txtContent, txtOwner, txtSearch, txtDetail;
        DateTimePicker dtDue;
        CheckBox chkUrgent;
        ListBox lstOpen;
        Label lblList, lblStatus;
        Panel pnlList;
        Button btnView;
        bool expanded;
        int expandedH;
        const int Dy = 300;  // 업무 내용 칸을 기본(44)보다 키운 만큼 아래 칸들이 내려감 (창 크기 = 스티커 메모 정도)
        const int TopH = 134 + Dy, StatusH = 26, CollapsedH = TopH + StatusH;
        Item detailItem;     // 찾기 칸에 번호를 넣어 내용을 보고 있는 항목 (없으면 null)
        List<Item> allItems = new List<Item>();
        List<Item> shown = new List<Item>();

        Label AddLabel(string text, int x, int y)
        {
            var lb = new Label { Text = text, AutoSize = true, Location = new Point(x, y) };
            Controls.Add(lb);
            return lb;
        }

        Button AddButton(string text, int x, int y, int w, int h, EventHandler onClick)
        {
            var b = new Button { Text = text, Location = new Point(x, y), Size = new Size(w, h) };
            b.Click += onClick;
            Controls.Add(b);
            return b;
        }

        public MainForm()
        {
            // 평소에는 입력칸 + 찾기 칸만 보이는 창. 찾기 칸에 번호를 넣으면 그 업무 내용이, [리스트]를 누르면 전체 목록이 펼쳐짐
            const int W = 420;
            expandedH = 330 + Dy;
            var small = new Font("맑은 고딕", 8.5F);
            Text = "업무메모";
            Font = new Font("맑은 고딕", 9F);
            ClientSize = new Size(W, expandedH);
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.Sizable;
            MaximizeBox = false;
            KeyPreview = true;
            try { Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch { }
            var tip = new ToolTip { AutoPopDelay = 20000, InitialDelay = 300 };
            var TL = AnchorStyles.Top | AnchorStyles.Left;
            var TLR = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            var TR = AnchorStyles.Top | AnchorStyles.Right;
            var BL = AnchorStyles.Bottom | AnchorStyles.Left;

            var lbHint = AddLabel("업무 내용  (Ctrl+Enter = 저장)", 6, 6);
            lbHint.Font = small; lbHint.ForeColor = Color.DimGray;
            txtContent = new TextBox { Multiline = true, ScrollBars = ScrollBars.Vertical, AcceptsReturn = true,
                Location = new Point(8, 24), Size = new Size(W - 16, 44 + Dy), Anchor = TLR };
            Controls.Add(txtContent);

            AddLabel("담당", 6, 78 + Dy);
            txtOwner = new TextBox { Location = new Point(38, 75 + Dy), Size = new Size(62, 23), Anchor = TL };
            Controls.Add(txtOwner);
            dtDue = new DateTimePicker { Format = DateTimePickerFormat.Custom, CustomFormat = "MM-dd (ddd)",
                ShowCheckBox = true, Checked = false, Location = new Point(104, 75 + Dy), Size = new Size(112, 23), Anchor = TL };
            tip.SetToolTip(dtDue, "마감일 (체크해야 저장됨)");
            Controls.Add(dtDue);
            chkUrgent = new CheckBox { Text = "긴급", AutoSize = true, Location = new Point(222, 77 + Dy), Anchor = TL };
            Controls.Add(chkUrgent);
            AddButton("저장", W - 68, 74 + Dy, 60, 25, (s, e) => DoSave()).Anchor = TR;

            AddLabel("찾기", 6, 110 + Dy);
            txtSearch = new TextBox { Location = new Point(38, 107 + Dy), Size = new Size(W - 46 - 80, 23), Anchor = TLR };
            tip.SetToolTip(txtSearch, "번호를 넣으면 그 업무 내용이 나옵니다. 글자를 넣으면 맞는 업무만 목록에 나옵니다 (Ctrl+F, Esc = 지우기)");
            // 찾기 칸이 비어 있으면 목록을 접고, 글자를 넣으면 펼침
            txtSearch.TextChanged += (s, e) =>
            {
                ApplyFilter();
                bool has = txtSearch.Text.Trim().Length > 0;
                if (has != expanded) SetExpanded(has);
            };
            Controls.Add(txtSearch);
            btnView = AddButton("리스트", W - 82, 106 + Dy, 74, 25, (s, e) => SetExpanded(!expanded));
            btnView.Anchor = TR;

            // 목록 영역 (접었다 폈다 함)
            pnlList = new Panel { Location = new Point(0, TopH), Size = new Size(W, expandedH - TopH - StatusH), Anchor = TLR, Visible = false };
            Controls.Add(pnlList);
            int ph = pnlList.Height;
            lblList = new Label { AutoSize = true, Location = new Point(6, 0), Font = small, ForeColor = Color.DimGray };
            pnlList.Controls.Add(lblList);
            lstOpen = new ListBox { HorizontalScrollbar = true, IntegralHeight = false,
                Location = new Point(8, 18), Size = new Size(W - 16, ph - 18 - 34),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right };
            lstOpen.DoubleClick += (s, e) => DoComplete();
            pnlList.Controls.Add(lstOpen);
            // 찾기 칸에 순번을 넣으면 목록 자리에 그 업무의 전체 내용을 보여줌
            txtDetail = new TextBox { Multiline = true, ReadOnly = true, WordWrap = true, ScrollBars = ScrollBars.Vertical,
                BackColor = Color.White, Location = lstOpen.Location, Size = lstOpen.Size, Anchor = lstOpen.Anchor, Visible = false };
            pnlList.Controls.Add(txtDetail);
            Func<string, int, int, EventHandler, AnchorStyles, Button> panelButton = (text, x, w, onClick, anchor) =>
            {
                var b = new Button { Text = text, Location = new Point(x, ph - 30), Size = new Size(w, 26), Anchor = anchor };
                b.Click += onClick;
                pnlList.Controls.Add(b);
                return b;
            };
            panelButton("완료", 8, 60, (s, e) => DoComplete(), BL);
            panelButton("새로고침", 72, 70, (s, e) => RefreshList(), BL);
            panelButton("폴더", 146, 50, (s, e) => OpenFolder(), BL);
            panelButton("메모장", W - 72, 64, (s, e) => OpenNotepad(), AnchorStyles.Bottom | AnchorStyles.Right);

            lblStatus = new Label { AutoSize = false, AutoEllipsis = true, Font = small, ForeColor = Color.DimGray,
                Location = new Point(6, expandedH - 22), Size = new Size(W - 12, 20),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right };
            lblStatus.TextChanged += (s, e) => tip.SetToolTip(lblStatus, lblStatus.Text);
            Controls.Add(lblStatus);

            expanded = true;   // 처음엔 펼친 크기로 만든 뒤 접어서 시작
            SetExpanded(false);

            KeyDown += OnKey;
            Shown += (s, e) =>
            {
                RefreshList();
                txtContent.Focus();
                Activate();
                BeginInvoke(new Action(Program.OfferShortcut));
            };
        }

        // 목록 펼치기/접기. 접으면 세로 크기를 고정하고, 펴면 마지막으로 쓰던 높이로 돌아감
        void SetExpanded(bool open)
        {
            if (open == expanded && pnlList.Visible == open) return;
            SuspendLayout();
            if (open)
            {
                MaximumSize = Size.Empty;
                MinimumSize = Size.Empty;
                ClientSize = new Size(ClientSize.Width, expandedH);
                pnlList.Size = new Size(ClientSize.Width, expandedH - TopH - StatusH);
                pnlList.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
                pnlList.Visible = true;
                MinimumSize = new Size(Width, Height - ClientSize.Height + CollapsedH + 140);
                btnView.Text = "리스트닫기";
            }
            else
            {
                if (expanded) expandedH = ClientSize.Height;
                pnlList.Visible = false;
                pnlList.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
                MaximumSize = Size.Empty;
                MinimumSize = Size.Empty;
                ClientSize = new Size(ClientSize.Width, CollapsedH);
                MinimumSize = new Size(Width, Height);
                // 가로는 제한 없음. 0을 넣으면 Windows가 창 폭을 0(최소 폭)으로 줄여 버리므로 화면 최대 폭을 넣음
                MaximumSize = new Size(SystemInformation.MaxWindowTrackSize.Width, Height);
                btnView.Text = "리스트";
            }
            expanded = open;
            ResumeLayout(true);
        }

        void SetStatus(string msg, bool isError)
        {
            lblStatus.Text = msg;
            lblStatus.ForeColor = isError ? Color.Firebrick : Color.SeaGreen;
        }

        void ShowFailure(string what, Exception ex)
        {
            Store.Log(what + ": " + ex.Message);
            SetStatus(what + " 실패 - 오류기록 파일을 확인하세요.", true);
            MessageBox.Show(this, what + "에 실패했습니다.\n\n" + ex.Message, "업무메모", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        // 메모 항목은 순번이 없으므로 결재 번호 다음 번호부터 차례로 붙임
        Item Numbered(Item it)
        {
            int max = 0, n;
            foreach (var x in allItems) if (int.TryParse(x.Seq, out n) && n > max) max = n;
            it.Seq = (max + 1).ToString();
            it.Display = "[메모] " + it.Seq + " | " + it.Text.Substring(4);
            return it;
        }

        bool NumberQuery() { return Regex.IsMatch(txtSearch.Text.Trim(), "^[0-9]+$"); }

        void ApplyFilter()
        {
            // 찾기 칸에 숫자만 있고 그 순번의 결재 항목이 있으면 목록 대신 내용을 보여줌
            string q = txtSearch.Text.Trim();
            bool byNumber = NumberQuery();
            detailItem = byNumber ? allItems.FirstOrDefault(it => it.Seq == q) : null;
            lstOpen.Visible = !byNumber;
            txtDetail.Visible = byNumber;
            if (byNumber)
            {
                txtDetail.Text = detailItem != null ? "[" + detailItem.Source + "] " + q + "번\r\n" + detailItem.Detail : q + "번 업무가 없습니다.";
                lblList.Text = detailItem != null ? q + "번 업무 내용  ·  [완료] = 완료 처리" : "번호를 확인하세요";
                return;
            }
            var words = txtSearch.Text.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
            shown = allItems.Where(it => words.All(w => it.Display.IndexOf(w, StringComparison.OrdinalIgnoreCase) >= 0)).ToList();
            lstOpen.BeginUpdate();
            lstOpen.Items.Clear();
            foreach (var it in shown) lstOpen.Items.Add(it.Display);
            lstOpen.EndUpdate();
            lblList.Text = words.Length > 0
                ? string.Format("찾은 결과 {0} / 전체 {1}건  ·  더블클릭 = 완료", shown.Count, allItems.Count)
                : string.Format("진행중 {0}건  ·  더블클릭 = 완료", allItems.Count);
        }

        void RefreshList()
        {
            allItems = new List<Item>();
            var msgs = new List<string>();
            bool isErr = false;
            try
            {
                var ex = Store.ExcelItems();
                allItems.AddRange(ex.Items);
                if (ex.FileName != null) msgs.Add(string.Format("결재 {1}건 (완료 숨김 {2}) - {0}", ex.FileName, ex.Items.Count, ex.Hidden));
                else { msgs.Add("결재 엑셀 파일 없음"); isErr = true; }
            }
            catch (Exception e)
            {
                Store.Log("엑셀 읽기: " + e.Message);
                msgs.Add("결재 엑셀 읽기 실패 (오류기록 확인)"); isErr = true;
            }
            try
            {
                var memo = Store.OpenTasks();
                foreach (var t in memo) allItems.Add(Numbered(Store.MemoItem(t)));
                msgs.Add(string.Format("메모 {0}건", memo.Length));
            }
            catch (Exception e)
            {
                Store.Log("메모 읽기: " + e.Message);
                msgs.Add("메모 읽기 실패"); isErr = true;
            }
            ApplyFilter();
            SetStatus(string.Join("  ·  ", msgs), isErr);
        }

        void DoSave()
        {
            string content = txtContent.Text.Trim();
            if (content.Length == 0) { SetStatus("업무 내용을 입력하세요.", true); txtContent.Focus(); return; }
            try
            {
                var sw = Stopwatch.StartNew();
                string line = Store.NewTaskLine(content, txtOwner.Text, dtDue.Checked ? (DateTime?)dtDue.Value : null, chkUrgent.Checked);
                Store.AddTask(line);
                sw.Stop();
                allItems.Add(Numbered(Store.MemoItem(line)));
                ApplyFilter();
                txtContent.Clear();
                chkUrgent.Checked = false;
                dtDue.Checked = false;
                SetStatus(string.Format("저장했습니다. ({0}, {1:N2}초)", DateTime.Now.ToString("HH:mm:ss"), sw.Elapsed.TotalSeconds), false);
                txtContent.Focus();
            }
            catch (Exception e) { ShowFailure("저장", e); }
        }

        void DoComplete()
        {
            int idx = lstOpen.SelectedIndex;
            Item it = detailItem;
            bool fromDetail = it != null;
            if (it == null && NumberQuery()) { SetStatus("그 번호의 업무가 없습니다.", true); return; }
            if (it == null)
            {
                if (idx < 0 || idx >= shown.Count) { SetStatus("완료할 항목을 먼저 고르세요.", true); return; }
                it = shown[idx];
            }
            try
            {
                if (it.Source == "결재")
                {
                    Store.CompleteExcel(it);
                    allItems.Remove(it);
                    ApplyFilter();
                    SetStatus("완료 처리했습니다: " + it.Text, false);
                }
                else if (Store.CompleteTask(it.Key))
                {
                    allItems.Remove(it);
                    ApplyFilter();
                    SetStatus("완료 처리했습니다: " + it.Display, false);
                }
                else
                {
                    RefreshList();
                    SetStatus("파일 내용이 바뀌어 항목을 찾지 못했습니다. 목록을 새로 불러왔습니다.", true);
                }
                if (fromDetail) { txtSearch.Clear(); txtSearch.Focus(); }   // 번호로 찾아 완료했으면 찾기 칸을 비워 목록을 접음
                else if (lstOpen.Items.Count > 0) lstOpen.SelectedIndex = Math.Min(Math.Max(idx, 0), lstOpen.Items.Count - 1);
            }
            catch (Exception e) { ShowFailure("완료 처리", e); }
        }

        void OpenNotepad()
        {
            try
            {
                Store.EnsureDataFile();
                Process.Start("notepad.exe", "\"" + Store.DataFile + "\"");
            }
            catch (Exception e) { ShowFailure("메모장 열기", e); }
        }

        void OpenFolder()
        {
            try { Process.Start("explorer.exe", "\"" + Store.BaseDir + "\""); }
            catch (Exception e) { ShowFailure("폴더 열기", e); }
        }

        void OnKey(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; DoSave(); }
            else if (e.Control && e.KeyCode == Keys.F) { e.SuppressKeyPress = true; txtSearch.Focus(); txtSearch.SelectAll(); }
            else if (e.KeyCode == Keys.Enter && lstOpen.Focused) { e.SuppressKeyPress = true; DoComplete(); }
            else if (e.KeyCode == Keys.Escape)
            {
                if (txtSearch.Text.Length > 0) txtSearch.Clear();
                else if (expanded) SetExpanded(false);
                else Close();
            }
        }
    }

    static class Program
    {
        // 바탕화면 '업무메모' 아이콘이 없거나 예전 입력창을 가리키면, 이 프로그램으로 연결할지 한 번 물어봄
        public static void OfferShortcut()
        {
            try
            {
                if (File.Exists(Store.SettingFile) && File.ReadAllText(Store.SettingFile).Contains("바로가기 묻지 않음")) return;
                string desk = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                string lnkPath = Path.Combine(desk, "업무메모.lnk");
                string exe = Application.ExecutablePath;

                Type shellType = Type.GetTypeFromProgID("WScript.Shell");
                object shell = Activator.CreateInstance(shellType);
                object lnk = shellType.InvokeMember("CreateShortcut", BindingFlags.InvokeMethod, null, shell, new object[] { lnkPath });
                Type lt = lnk.GetType();
                string current = File.Exists(lnkPath) ? (string)lt.InvokeMember("TargetPath", BindingFlags.GetProperty, null, lnk, null) : null;
                if (string.Equals(current, exe, StringComparison.OrdinalIgnoreCase)) return;

                string q = current == null
                    ? "바탕화면에 '업무메모' 아이콘을 만들까요?"
                    : "바탕화면 '업무메모' 아이콘이 예전 입력창에 연결되어 있습니다.\n이 프로그램으로 바꿀까요?";
                var ans = MessageBox.Show(q + "\n\n('아니요'를 누르면 다시 묻지 않습니다)", "업무메모", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (ans != DialogResult.Yes)
                {
                    File.AppendAllText(Store.SettingFile, "바로가기 묻지 않음\r\n", new UTF8Encoding(true));
                    return;
                }
                lt.InvokeMember("TargetPath", BindingFlags.SetProperty, null, lnk, new object[] { exe });
                lt.InvokeMember("WorkingDirectory", BindingFlags.SetProperty, null, lnk, new object[] { Path.GetDirectoryName(exe) });
                lt.InvokeMember("IconLocation", BindingFlags.SetProperty, null, lnk, new object[] { exe + ",0" });
                lt.InvokeMember("Description", BindingFlags.SetProperty, null, lnk, new object[] { "업무지시 입력 / 결재 미마감 완료" });
                lt.InvokeMember("Save", BindingFlags.InvokeMethod, null, lnk, null);
                MessageBox.Show("바탕화면에 '업무메모' 아이콘을 만들었습니다.", "업무메모");
            }
            catch (Exception e)
            {
                Store.Log("바로가기 만들기: " + e.Message);
            }
        }

        [STAThread]
        static void Main()
        {
            string exeDir = Path.GetDirectoryName(Application.ExecutablePath);
            Store.Init(Store.PickDir(exeDir, Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)));

            bool createdNew;
            using (var mutex = new Mutex(true, @"Local\FMJ_TaskInput", out createdNew))
            {
                if (!createdNew)
                {
                    MessageBox.Show("업무메모(입력창)가 이미 열려 있습니다.", "업무메모");
                    return;
                }
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.ThreadException += (s, e) =>
                {
                    Store.Log("실행 오류: " + e.Exception);
                    MessageBox.Show("오류가 났습니다.\n\n" + e.Exception.Message + "\n\n기록: " + Store.LogFile, "업무메모", MessageBoxButtons.OK, MessageBoxIcon.Error);
                };
                try { Application.Run(new MainForm()); }
                catch (Exception e)
                {
                    Store.Log("실행 오류: " + e);
                    MessageBox.Show("프로그램 실행 중 오류가 났습니다.\n\n" + e.Message + "\n\n기록: " + Store.LogFile, "업무메모", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally { mutex.ReleaseMutex(); }
            }
        }
    }
}
