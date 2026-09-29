using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

// Excel -> PDF. C# / WinForms, Excel через COM (late binding, любая версия Office).
// Сборка: build.bat (нужен только .NET Framework 4, есть в любой Windows 10/11).

static class Program
{
    [STAThread]
    static void Main(string[] args)
    {
        // PdfSharp.dll вшит в exe как ресурс — отдельный файл не нужен.
        AppDomain.CurrentDomain.AssemblyResolve += delegate(object s, ResolveEventArgs a)
        {
            if (new AssemblyName(a.Name).Name != "PdfSharp") return null;
            using (Stream st = Assembly.GetExecutingAssembly().GetManifestResourceStream("PdfSharp.dll"))
            {
                if (st == null) return null;
                MemoryStream ms = new MemoryStream();
                st.CopyTo(ms);
                return Assembly.Load(ms.ToArray());
            }
        };
        Run(args);
    }

    static void Run(string[] args)
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new MainForm(args));
    }
}

class Settings
{
    public bool StampFirst, PerSheet, Fit, IgnoreAreas, Merge, Stamp, Top, Bot, TopReset, BotReset;
    public int Quality, Paper, Orient, PosS, PosTop, PosBot;
    public string FmtTop, FmtBot, StampText, StampStart, MergeName, OutDir;
    public int StartTop, StartBot;
}

class Item
{
    public string Tmp, Final;
    public PdfDocument Doc;
}

class MainForm : Form
{
    const string Letters = "АБВГДЕЖИКЛМНПРСТУФХЦШЩЭЮЯ"; // буквы приложений по ГОСТ 2.105
    static readonly string[] Pos = { "Справа", "По центру", "Слева" };
    static readonly string[] Fmts = { "{P}", "Стр. {P}", "Стр. {P} из {T}", "Лист {P}", "Лист {P} из {T}" };

    readonly List<string> queue = new List<string>();
    bool busy;
    readonly string cfgDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ExcelToPDF");
    string cfgOut { get { return Path.Combine(cfgDir, "outdir.txt"); } }

    Label drop, hintN, hint;
    ListBox lbFiles, log;
    Button btnAdd, btnDel, btnClear, btnGo, btnOut, btnClr;
    RadioButton rbBook, rbSheet;
    ComboBox cbQuality, cbPaper, cbOrient, cbPosS, cbFmtTop, cbPosTop, cbFmtBot, cbPosBot;
    CheckBox chStampFirst, chFit, chIgnore, chMerge, chStamp, chTop, chBot, chTopReset, chBotReset;
    TextBox tbMerge, tbStamp, tbStart, tbOut;
    NumericUpDown nuTop, nuBot;

    public MainForm(string[] args)
    {
        Text = "Excel → PDF";
        ClientSize = new Size(566, 756);
        StartPosition = FormStartPosition.CenterScreen;
        Font = new Font("Segoe UI", 10);
        AllowDrop = true;
        try { Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch { }

        drop = new Label { TextAlign = ContentAlignment.MiddleCenter, BorderStyle = BorderStyle.FixedSingle, BackColor = Color.FromArgb(238, 245, 255), AllowDrop = true };
        Put(drop, 15, 12, 535, 50);
        lbFiles = new ListBox { SelectionMode = SelectionMode.MultiExtended, HorizontalScrollbar = true, AllowDrop = true };
        Put(lbFiles, 15, 67, 535, 110);
        btnAdd = Btn("Добавить…", 15, 182, 120, 32);
        btnDel = Btn("Убрать", 140, 182, 90, 32);
        btnClear = Btn("Очистить", 235, 182, 100, 32);
        btnGo = Btn("Конвертировать", 380, 180, 170, 36);
        btnGo.Font = new Font("Segoe UI", 10, FontStyle.Bold);

        rbBook = new RadioButton { Text = "Один PDF на файл", Checked = true, AutoSize = true };
        Put(rbBook, 15, 228, 0, 0);
        rbSheet = new RadioButton { Text = "Каждый лист отдельным PDF", AutoSize = true };
        Put(rbSheet, 290, 228, 0, 0);
        Lbl("Качество:", 15, 262);
        cbQuality = Combo(new[] { "Высокое (печать)", "Сжатое (меньше размер)" }, 110, 259, 170, 0);
        Lbl("Бумага:", 290, 262);
        cbPaper = Combo(new[] { "Как в файле", "A4", "A3" }, 400, 259, 150, 0);
        Lbl("Ориентация:", 15, 296);
        cbOrient = Combo(new[] { "Как в файле", "Книжная", "Альбомная" }, 110, 293, 170, 0);
        chFit = Chk("Вписать по ширине", 290, 296);
        chIgnore = Chk("Игнорировать области печати", 15, 328);

        Label sep = new Label { BorderStyle = BorderStyle.Fixed3D };
        Put(sep, 15, 358, 535, 2);
        chMerge = Chk("Склеить всё в один PDF:", 15, 370);
        tbMerge = Txt("Приложения.pdf", 230, 368, 320);

        chStamp = Chk("Штамп:", 15, 405);
        tbStamp = Txt("Приложение {N}", 110, 403, 190);
        Lbl("с", 308, 406);
        tbStart = Txt("А", 328, 403, 55);
        cbPosS = Combo(Pos, 390, 402, 160, 0);
        chStampFirst = Chk("только на первой странице", 110, 429);
        chStampFirst.Checked = true;
        hintN = Lbl("{N} — буква или число, как в поле «с»", 330, 433);
        hintN.ForeColor = Color.Gray; hintN.Font = new Font("Segoe UI", 8);

        chTop = Chk("Номер сверху:", 15, 458);
        cbFmtTop = Combo(Fmts, 140, 456, 160, 0);
        Lbl("с", 308, 459);
        nuTop = Num(328, 456);
        cbPosTop = Combo(Pos, 390, 455, 160, 1);
        chTopReset = Chk("заново в каждом приложении", 140, 486);

        chBot = Chk("Номер снизу:", 15, 518);
        cbFmtBot = Combo(Fmts, 140, 516, 160, 0);
        Lbl("с", 308, 519);
        nuBot = Num(328, 516);
        cbPosBot = Combo(Pos, 390, 515, 160, 1);
        chBotReset = Chk("заново в каждом приложении", 140, 546);

        Label sep2 = new Label { BorderStyle = BorderStyle.Fixed3D };
        Put(sep2, 15, 576, 535, 2);
        Lbl("Сохранять в:", 15, 591);
        tbOut = Txt("", 110, 588, 290);
        try { if (File.Exists(cfgOut)) tbOut.Text = File.ReadAllText(cfgOut).Trim(); } catch { }
        btnOut = Btn("Обзор…", 405, 586, 80, 30);
        btnClr = Btn("Сброс", 490, 586, 60, 30);
        hint = Lbl("пусто = подпапка PDF рядом с файлами Excel", 110, 617);
        hint.ForeColor = Color.Gray; hint.Font = new Font("Segoe UI", 8);

        log = new ListBox { HorizontalScrollbar = true };
        Put(log, 15, 640, 535, 100);

        DragEventHandler onEnter = delegate(object s, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy;
        };
        DragEventHandler onDrop = delegate(object s, DragEventArgs e)
        {
            AddFiles((string[])e.Data.GetData(DataFormats.FileDrop));
        };
        foreach (Control c in new Control[] { this, drop, lbFiles })
        {
            c.DragEnter += onEnter;
            c.DragDrop += onDrop;
        }

        btnAdd.Click += delegate
        {
            using (OpenFileDialog dlg = new OpenFileDialog())
            {
                dlg.Filter = "Excel|*.xls;*.xlsx;*.xlsm";
                dlg.Multiselect = true;
                if (dlg.ShowDialog() == DialogResult.OK) AddFiles(dlg.FileNames);
            }
        };
        btnDel.Click += delegate { RemoveSelected(); };
        lbFiles.KeyDown += delegate(object s, KeyEventArgs e) { if (e.KeyCode == Keys.Delete) RemoveSelected(); };
        btnClear.Click += delegate { queue.Clear(); RefreshQueue(); };
        btnGo.Click += delegate { StartConvert(); };
        btnOut.Click += delegate
        {
            using (FolderBrowserDialog d = new FolderBrowserDialog())
            {
                d.Description = "Папка для PDF";
                if (tbOut.Text.Length > 0 && Directory.Exists(tbOut.Text)) d.SelectedPath = tbOut.Text;
                if (d.ShowDialog() == DialogResult.OK) { tbOut.Text = d.SelectedPath; SaveOutCfg(); }
            }
        };
        btnClr.Click += delegate { tbOut.Text = ""; SaveOutCfg(); };

        RefreshQueue();
        if (args.Length > 0) AddFiles(args);
    }

    // ---------- helpers UI ----------
    void Put(Control c, int x, int y, int w, int h)
    {
        c.Location = new Point(x, y);
        if (w > 0) c.Size = new Size(w, h); else c.AutoSize = true;
        Controls.Add(c);
    }
    Label Lbl(string t, int x, int y) { Label l = new Label { Text = t }; Put(l, x, y, 0, 0); return l; }
    CheckBox Chk(string t, int x, int y) { CheckBox c = new CheckBox { Text = t }; Put(c, x, y, 0, 0); return c; }
    TextBox Txt(string t, int x, int y, int w) { TextBox c = new TextBox { Text = t }; Put(c, x, y, w, 26); return c; }
    Button Btn(string t, int x, int y, int w, int h) { Button b = new Button { Text = t }; Put(b, x, y, w, h); return b; }
    NumericUpDown Num(int x, int y)
    {
        NumericUpDown c = new NumericUpDown { Minimum = 0, Maximum = 99999, Value = 1 };
        Put(c, x, y, 55, 26); return c;
    }
    ComboBox Combo(string[] items, int x, int y, int w, int sel)
    {
        ComboBox c = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
        c.Items.AddRange(items);
        c.SelectedIndex = sel;
        Put(c, x, y, w, 28); return c;
    }

    void Log(string m)
    {
        if (InvokeRequired) { Invoke(new Action<string>(Log), m); return; }
        log.Items.Add(m);
        log.TopIndex = log.Items.Count - 1;
    }

    void SaveOutCfg()
    {
        try { Directory.CreateDirectory(cfgDir); File.WriteAllText(cfgOut, tbOut.Text); } catch { }
    }

    // ---------- очередь ----------
    static string NatKey(string s)
    {
        return Regex.Replace(Path.GetFileName(s), @"\d+", delegate(Match m) { return m.Value.PadLeft(10, '0'); });
    }

    void RefreshQueue()
    {
        queue.Sort(delegate(string a, string b) { return string.Compare(NatKey(a), NatKey(b), StringComparison.OrdinalIgnoreCase); });
        lbFiles.Items.Clear();
        foreach (string q in queue)
            lbFiles.Items.Add(Path.GetFileName(q) + "   —   " + Path.GetDirectoryName(q));
        drop.Text = queue.Count > 0
            ? "Файлов в очереди: " + queue.Count + ". Можно добавить ещё — перетащите сюда"
            : "Перетащите сюда файлы или папки Excel (.xls, .xlsx, .xlsm)";
        btnGo.Enabled = queue.Count > 0 && !busy;
    }

    static readonly Regex XlsRe = new Regex(@"\.xls[xm]?$", RegexOptions.IgnoreCase);

    void AddFiles(string[] paths)
    {
        foreach (string p in paths)
        {
            IEnumerable<string> list;
            if (Directory.Exists(p))
                list = Directory.GetFiles(p).Where(f => XlsRe.IsMatch(f));
            else if (XlsRe.IsMatch(p))
                list = new[] { p };
            else
                continue;
            foreach (string f in list)
                if (!Path.GetFileName(f).StartsWith("~$") && !queue.Contains(f)) queue.Add(f);
        }
        RefreshQueue();
    }

    void RemoveSelected()
    {
        foreach (int i in lbFiles.SelectedIndices.Cast<int>().OrderByDescending(x => x).ToList())
            queue.RemoveAt(i);
        RefreshQueue();
    }

    // ---------- запуск ----------
    void StartConvert()
    {
        if (busy || queue.Count == 0) return;
        SaveOutCfg();
        Settings s = new Settings
        {
            StampFirst = chStampFirst.Checked, PerSheet = rbSheet.Checked, Fit = chFit.Checked, IgnoreAreas = chIgnore.Checked,
            Merge = chMerge.Checked, Stamp = chStamp.Checked, Top = chTop.Checked, Bot = chBot.Checked,
            TopReset = chTopReset.Checked, BotReset = chBotReset.Checked,
            Quality = cbQuality.SelectedIndex, Paper = cbPaper.SelectedIndex, Orient = cbOrient.SelectedIndex,
            PosS = cbPosS.SelectedIndex, PosTop = cbPosTop.SelectedIndex, PosBot = cbPosBot.SelectedIndex,
            FmtTop = cbFmtTop.Text, FmtBot = cbFmtBot.Text, StampText = tbStamp.Text, StampStart = tbStart.Text,
            MergeName = tbMerge.Text.Trim(), OutDir = tbOut.Text.Trim(),
            StartTop = (int)nuTop.Value, StartBot = (int)nuBot.Value
        };
        List<string> files = new List<string>(queue);
        busy = true; Cursor = Cursors.WaitCursor; btnGo.Enabled = false;

        // Вся работа с COM — в отдельном STA-потоке, без DoEvents (он и вызывал 0x80010114).
        Thread t = new Thread(delegate()
        {
            List<string> made = new List<string>();
            int err = 0;
            try { made = new Converter(this, s).Run(files, out err); }
            catch (Exception ex) { err++; Log("ОШИБКА: " + ex.Message); }
            Log("Итого PDF: " + made.Count + ", ошибок: " + err);
            BeginInvoke(new Action(delegate
            {
                busy = false; Cursor = Cursors.Default; RefreshQueue();
                if (made.Count > 0)
                    try { Process.Start("explorer.exe", "/select,\"" + made[made.Count - 1] + "\""); } catch { }
            }));
        });
        t.SetApartmentState(ApartmentState.STA);
        t.IsBackground = true;
        t.Start();
    }

    public void LogFromWorker(string m) { Log(m); }

    // ---------- конвертер ----------
    class Converter
    {
        readonly MainForm f;
        readonly Settings s;
        dynamic excel;
        XFont font;

        public Converter(MainForm form, Settings st) { f = form; s = st; }

        void Log(string m) { f.LogFromWorker(m); }

        static bool IsComDead(Exception ex)
        {
            COMException ce = ex as COMException;
            if (ce == null) return false;
            uint h = (uint)ce.ErrorCode;
            // 0x80010114 объект отсутствует, 0x800706BA/0x800706BE RPC недоступен, 0x80010108/0x80010105 отключён
            return h == 0x80010114 || h == 0x800706BA || h == 0x800706BE || h == 0x80010108 || h == 0x80010105 || h == 0x800706BF;
        }

        void StartExcel()
        {
            Type t = Type.GetTypeFromProgID("Excel.Application");
            if (t == null) throw new InvalidOperationException("Не найден Microsoft Excel. Для работы нужен установленный Excel.");
            excel = Activator.CreateInstance(t);
            excel.Visible = false;
            excel.DisplayAlerts = false;
            excel.ScreenUpdating = false;
            excel.AskToUpdateLinks = false;
        }

        void StopExcel()
        {
            if (excel == null) return;
            try { excel.Quit(); } catch { }
            try { Marshal.FinalReleaseComObject(excel); } catch { }
            excel = null;
            GC.Collect(); GC.WaitForPendingFinalizers();
        }

        public List<string> Run(List<string> files, out int err)
        {
            err = 0;
            List<Item> items = new List<Item>();
            string tmp = Path.Combine(Path.GetTempPath(), "x2p_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tmp);
            int k = 0;
            List<string> made = new List<string>();
            try
            {
                try { StartExcel(); }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Excel → PDF");
                    err++;
                    return made;
                }

                foreach (string file in files)
                {
                    Log("→ " + Path.GetFileName(file));
                    int before = items.Count;
                    for (int attempt = 0; attempt < 2; attempt++)
                    {
                        try
                        {
                            ConvertOne(file, tmp, items, ref k);
                            Log("   готово");
                            break;
                        }
                        catch (Exception ex)
                        {
                            items.RemoveRange(before, items.Count - before);
                            if (attempt == 0 && IsComDead(ex))
                            {
                                Log("   Excel не отвечает, перезапускаю…");
                                StopExcel();
                                try { StartExcel(); } catch (Exception ex2) { err++; Log("   ОШИБКА: " + ex2.Message); break; }
                                continue;
                            }
                            err++;
                            Log("   ОШИБКА: " + ex.Message);
                            break;
                        }
                    }
                }
                StopExcel();

                if (items.Count > 0)
                {
                    try { made = Build(items); }
                    catch (Exception ex) { err++; Log("ОШИБКА сборки PDF: " + ex.Message); }
                }
            }
            finally
            {
                StopExcel();
                try { Directory.Delete(tmp, true); } catch { }
            }
            return made;
        }

        void ConvertOne(string file, string tmp, List<Item> items, ref int k)
        {
            string outDir = s.OutDir.Length > 0 ? s.OutDir : Path.Combine(Path.GetDirectoryName(file), "PDF");
            Directory.CreateDirectory(outDir);
            string baseName = Path.GetFileNameWithoutExtension(file);
            dynamic wb = null;
            try
            {
                wb = excel.Workbooks.Open(file, 0, true);
                SetPages(wb);
                if (s.PerSheet)
                {
                    foreach (dynamic ws in wb.Sheets)
                    {
                        if ((int)ws.Visible != -1) continue;
                        string name = Regex.Replace((string)ws.Name, @"[\\/:*?""<>|]", "_");
                        string t = Path.Combine(tmp, (++k).ToString("D4") + ".pdf");
                        try
                        {
                            Export(ws, t);
                            items.Add(new Item { Tmp = t, Final = Path.Combine(outDir, baseName + " - " + name + ".pdf") });
                        }
                        catch (COMException ex)
                        {
                            if (IsComDead(ex)) throw;
                            Log("   лист «" + (string)ws.Name + "» пропущен (пустой?)");
                        }
                    }
                }
                else
                {
                    string t = Path.Combine(tmp, (++k).ToString("D4") + ".pdf");
                    Export(wb, t);
                    items.Add(new Item { Tmp = t, Final = Path.Combine(outDir, baseName + ".pdf") });
                }
            }
            finally
            {
                if (wb != null)
                {
                    try { wb.Close(false); } catch { }
                    try { Marshal.FinalReleaseComObject(wb); } catch { }
                }
            }
        }

        void SetPages(dynamic wb)
        {
            if (s.Paper == 0 && s.Orient == 0 && !s.Fit) return;
            try { excel.PrintCommunication = false; } catch { }
            try
            {
                foreach (dynamic ws in wb.Worksheets)
                {
                    try
                    {
                        if ((int)ws.Visible != -1) continue;
                        dynamic ps = ws.PageSetup;
                        if (s.Paper == 1) ps.PaperSize = 9;   // A4
                        if (s.Paper == 2) ps.PaperSize = 8;   // A3
                        if (s.Orient == 1) ps.Orientation = 1;
                        if (s.Orient == 2) ps.Orientation = 2;
                        if (s.Fit) { ps.Zoom = false; ps.FitToPagesWide = 1; ps.FitToPagesTall = false; }
                    }
                    catch (COMException ex) { if (IsComDead(ex)) throw; }
                }
            }
            finally
            {
                try { excel.PrintCommunication = true; } catch { }
            }
        }

        void Export(dynamic obj, string path)
        {
            // xlTypePDF=0; Quality: 0 = standard (печать), 1 = minimum (сжатое)
            obj.ExportAsFixedFormat(0, path, s.Quality, true, s.IgnoreAreas);
        }

        // ---------- PdfSharp ----------
        static string AppLabel(string start, int i)
        {
            string st = start.Trim();
            int n;
            if (Regex.IsMatch(st, @"^\d+$") && int.TryParse(st, out n)) return (n + i).ToString();
            int idx = st.Length > 0 ? Letters.IndexOf(char.ToUpper(st[0])) : 0;
            if (idx < 0) idx = 0;
            int kk = idx + i;
            string l = Letters[kk % Letters.Length].ToString();
            if (kk >= Letters.Length) l += (kk / Letters.Length + 1).ToString();
            return l;
        }

        static string Fmt(string f, int p, int t)
        {
            return f.Replace("{P}", p.ToString()).Replace("{T}", t.ToString());
        }

        static XStringFormat Align(int pos)
        {
            if (pos == 0) return XStringFormats.TopRight;
            if (pos == 1) return XStringFormats.TopCenter;
            return XStringFormats.TopLeft;
        }

        void StampPage(PdfPage page, List<KeyValuePair<string, int>> top, List<KeyValuePair<string, int>> bottom)
        {
            if (top.Count == 0 && bottom.Count == 0) return;
            using (XGraphics g = XGraphics.FromPdfPage(page))
            {
                double w = page.Width.Point, h = page.Height.Point;
                Dictionary<int, int> used = new Dictionary<int, int>();
                foreach (KeyValuePair<string, int> e in top)
                {
                    int n; used.TryGetValue(e.Value, out n); used[e.Value] = n + 1;
                    g.DrawString(e.Key, font, XBrushes.Black, new XRect(30, 16 + 16 * n, w - 60, 16), Align(e.Value));
                }
                used.Clear();
                foreach (KeyValuePair<string, int> e in bottom)
                {
                    int n; used.TryGetValue(e.Value, out n); used[e.Value] = n + 1;
                    g.DrawString(e.Key, font, XBrushes.Black, new XRect(30, h - 32 - 16 * n, w - 60, 16), Align(e.Value));
                }
            }
        }

        List<string> Build(List<Item> items)
        {
            List<string> made = new List<string>();
            bool post = s.Merge || s.Stamp || s.Top || s.Bot;
            if (!post)
            {
                foreach (Item it in items)
                {
                    if (File.Exists(it.Final)) File.Delete(it.Final);
                    File.Move(it.Tmp, it.Final);
                    made.Add(it.Final);
                }
                return made;
            }

            font = new XFont("Times New Roman", 12, XFontStyle.Regular, new XPdfFontOptions(PdfFontEncoding.Unicode));
            try
            {
                foreach (Item it in items) it.Doc = PdfReader.Open(it.Tmp, PdfDocumentOpenMode.Import);
                int total = items.Sum(x => x.Doc.PageCount);
                int pTop = s.StartTop, pBot = s.StartBot, tTop, tBot, ai = 0;
                PdfDocument outDoc = s.Merge ? new PdfDocument() : null;

                foreach (Item it in items)
                {
                    if (!s.Merge) outDoc = new PdfDocument();
                    int cnt = it.Doc.PageCount;
                    if (s.TopReset) { pTop = s.StartTop; tTop = s.StartTop + cnt - 1; } else tTop = s.StartTop + total - 1;
                    if (s.BotReset) { pBot = s.StartBot; tBot = s.StartBot + cnt - 1; } else tBot = s.StartBot + total - 1;
                    string stampText = s.Stamp ? s.StampText.Replace("{N}", AppLabel(s.StampStart, ai)) : "";
                    for (int i = 0; i < cnt; i++)
                    {
                        PdfPage pg = outDoc.AddPage(it.Doc.Pages[i]);
                        List<KeyValuePair<string, int>> top = new List<KeyValuePair<string, int>>();
                        List<KeyValuePair<string, int>> bottom = new List<KeyValuePair<string, int>>();
                        if (stampText.Length > 0 && (i == 0 || !s.StampFirst)) top.Add(new KeyValuePair<string, int>(stampText, s.PosS));
                        if (s.Top) top.Add(new KeyValuePair<string, int>(Fmt(s.FmtTop, pTop, tTop), s.PosTop));
                        if (s.Bot) bottom.Add(new KeyValuePair<string, int>(Fmt(s.FmtBot, pBot, tBot), s.PosBot));
                        StampPage(pg, top, bottom);
                        pTop++; pBot++;
                    }
                    ai++;
                    if (!s.Merge) { outDoc.Save(it.Final); made.Add(it.Final); }
                }
                if (s.Merge)
                {
                    string name = s.MergeName.Length > 0 ? s.MergeName : "Приложения";
                    if (!name.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase)) name += ".pdf";
                    string path = Path.Combine(Path.GetDirectoryName(items[0].Final), name);
                    outDoc.Save(path);
                    made.Add(path);
                }
                Log("Страниц всего: " + total);
            }
            finally
            {
                foreach (Item it in items) if (it.Doc != null) it.Doc.Close();
            }
            return made;
        }
    }
}
