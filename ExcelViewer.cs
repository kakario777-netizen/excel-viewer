using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using System.Xml;

[assembly: System.Reflection.AssemblyTitle("엑셀 뷰어")]
[assembly: System.Reflection.AssemblyProduct("엑셀 뷰어")]
[assembly: System.Reflection.AssemblyVersion("1.2.0.0")]
[assembly: System.Reflection.AssemblyFileVersion("1.2.0.0")]

namespace ExcelViewer
{
    static class AppInfo
    {
        public const string Version = "1.2.0";
        public const string Repo = "kakario777-netizen/excel-viewer";
        public static readonly Color Accent = Color.FromArgb(33, 115, 70);
    }

    // ───────────────────────── 데이터 모델 ─────────────────────────
    class XBorder { public string Style; public Color Color = Color.Black; }

    class CellStyle
    {
        public string FontName = "맑은 고딕";
        public float FontSize = 11;
        public bool Bold, Italic, Underline, Strike;
        public Color FontColor = Color.Black;
        public bool HasFill;
        public Color Fill = Color.White;
        public XBorder L, R, T, B;
        public string HAlign = "general", VAlign = "bottom";
        public bool Wrap;
        public int Indent;
        public int NumFmtId;
        public string NumFmt;
        public bool HasBorder { get { return L != null || R != null || T != null || B != null; } }
    }

    class Cell { public string Text; public int Style; public bool IsNum, IsBool; }
    class Merge { public int R1, C1, R2, C2; }

    class Sheet
    {
        public string Name;
        public List<CellStyle> Styles;
        public Dictionary<long, Cell> Cells = new Dictionary<long, Cell>();
        public List<Merge> Merges = new List<Merge>();
        public Dictionary<long, Merge> MergeOf = new Dictionary<long, Merge>();
        public Dictionary<int, double> ColW = new Dictionary<int, double>();   // 엑셀 문자 단위 폭
        public Dictionary<int, double> RowHt = new Dictionary<int, double>();  // pt
        public HashSet<int> CustomHt = new HashSet<int>();
        public HashSet<int> HiddenRows = new HashSet<int>(), HiddenCols = new HashSet<int>();
        public double DefColWidthChars = -1, DefRowHtPt = 15;
        public bool ShowGrid = true;
        public int Rows, Cols;
        // 레이아웃 (96dpi 픽셀)
        public float[] X, Y, CW, RH;
        public List<long> TextOrder;
        public string EntryName;
        public int Mdw = 8;
        public Dictionary<int, float> EditCols = new Dictionary<int, float>(), EditRows = new Dictionary<int, float>();
        public Dictionary<long, string> EditCells = new Dictionary<long, string>();
        public bool Dirty { get { return EditCols.Count + EditRows.Count + EditCells.Count > 0; } }
        public void Recalc()
        {
            X = new float[Cols + 1]; Y = new float[Rows + 1];
            for (int c = 0; c < Cols; c++) X[c + 1] = X[c] + CW[c];
            for (int r = 0; r < Rows; r++) Y[r + 1] = Y[r] + RH[r];
        }

        public static long Key(int r, int c) { return ((long)r << 16) | (uint)c; }
        public Cell Get(int r, int c) { Cell x; return Cells.TryGetValue(Key(r, c), out x) ? x : null; }
        public Merge MergeAt(int r, int c) { Merge m; return MergeOf.TryGetValue(Key(r, c), out m) ? m : null; }
        public CellStyle StyleOf(Cell c) { return c != null && c.Style >= 0 && c.Style < Styles.Count ? Styles[c.Style] : Styles[0]; }
    }

    // ───────────────────────── 글꼴 캐시 ─────────────────────────
    static class Fonts
    {
        static Dictionary<string, Font> cache = new Dictionary<string, Font>();
        static HashSet<string> installed;
        public static Font For(CellStyle st, float z)
        {
            string key = st.FontName + "|" + st.FontSize + "|" + st.Bold + st.Italic + st.Underline + st.Strike + "|" + z.ToString("0.00");
            Font f;
            if (cache.TryGetValue(key, out f)) return f;
            if (installed == null)
            {
                installed = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                try { foreach (var fam in new InstalledFontCollection().Families) installed.Add(fam.Name); } catch { }
            }
            string name = installed.Contains(st.FontName) ? st.FontName : "맑은 고딕";
            var fs = FontStyle.Regular;
            if (st.Bold) fs |= FontStyle.Bold;
            if (st.Italic) fs |= FontStyle.Italic;
            if (st.Underline) fs |= FontStyle.Underline;
            if (st.Strike) fs |= FontStyle.Strikeout;
            float px = Math.Max(1f, st.FontSize * 96f / 72f * z);
            try { f = new Font(name, px, fs, GraphicsUnit.Pixel); }
            catch { f = new Font(FontFamily.GenericSansSerif, px, fs, GraphicsUnit.Pixel); }
            cache[key] = f;
            return f;
        }
    }

    // ───────────────────────── xlsx 읽기 ─────────────────────────
    static class XlsxReader
    {
        static readonly int[] Indexed = {
            0x000000,0xFFFFFF,0xFF0000,0x00FF00,0x0000FF,0xFFFF00,0xFF00FF,0x00FFFF,
            0x000000,0xFFFFFF,0xFF0000,0x00FF00,0x0000FF,0xFFFF00,0xFF00FF,0x00FFFF,
            0x800000,0x008000,0x000080,0x808000,0x800080,0x008080,0xC0C0C0,0x808080,
            0x9999FF,0x993366,0xFFFFCC,0xCCFFFF,0x660066,0xFF8080,0x0066CC,0xCCCCFF,
            0x000080,0xFF00FF,0xFFFF00,0x00FFFF,0x800080,0x800000,0x008080,0x0000FF,
            0x00CCFF,0xCCFFFF,0xCCFFCC,0xFFFF99,0x99CCFF,0xFF99CC,0xCC99FF,0xFFCC99,
            0x3366FF,0x33CCCC,0x99CC00,0xFFCC00,0xFF9900,0xFF6600,0x666699,0x969696,
            0x003366,0x339966,0x003300,0x333300,0x993300,0x993366,0x333399,0x333333 };

        public static List<Sheet> Read(string path)
        {
            var result = new List<Sheet>();
            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var zip = new ZipArchive(fs, ZipArchiveMode.Read))
            {
                var theme = ReadTheme(zip);
                var shared = ReadShared(zip);
                var styles = ReadStyles(zip, theme);
                var rels = ReadRels(zip);
                var wb = Load(zip, "xl/workbook.xml");
                if (wb == null) throw new Exception("엑셀 파일 구조를 읽을 수 없습니다.");
                foreach (XmlElement s in wb.GetElementsByTagName("sheet", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"))
                {
                    string state = s.GetAttribute("state");
                    if (state == "hidden" || state == "veryHidden") continue;
                    string rid = null;
                    foreach (XmlAttribute a in s.Attributes)
                        if (a.LocalName == "id" && a.NamespaceURI.Contains("relationships")) rid = a.Value;
                    string target;
                    if (rid == null || !rels.TryGetValue(rid, out target)) continue;
                    target = target.StartsWith("/") ? target.Substring(1) : "xl/" + target;
                    var entry = FindEntry(zip, target);
                    if (entry == null) continue;
                    var sh = new Sheet { Name = s.GetAttribute("name"), Styles = styles, EntryName = entry.FullName };
                    ReadSheet(entry, sh, shared);
                    result.Add(sh);
                }
            }
            return result;
        }

        static ZipArchiveEntry FindEntry(ZipArchive zip, string name)
        {
            foreach (var e in zip.Entries)
                if (string.Equals(e.FullName.Replace('\\', '/'), name, StringComparison.OrdinalIgnoreCase)) return e;
            return null;
        }

        static XmlDocument Load(ZipArchive zip, string name)
        {
            var e = FindEntry(zip, name);
            if (e == null) return null;
            var doc = new XmlDocument();
            using (var st = e.Open()) doc.Load(st);
            return doc;
        }

        static Dictionary<string, string> ReadRels(ZipArchive zip)
        {
            var d = new Dictionary<string, string>();
            var doc = Load(zip, "xl/_rels/workbook.xml.rels");
            if (doc == null) return d;
            foreach (XmlElement r in doc.GetElementsByTagName("Relationship"))
                d[r.GetAttribute("Id")] = r.GetAttribute("Target");
            return d;
        }

        static List<Color> ReadTheme(ZipArchive zip)
        {
            var list = new List<Color>();
            var doc = Load(zip, "xl/theme/theme1.xml");
            if (doc != null)
            {
                foreach (XmlNode n in doc.GetElementsByTagName("clrScheme", "http://schemas.openxmlformats.org/drawingml/2006/main"))
                {
                    foreach (XmlNode c in n.ChildNodes)
                    {
                        var el = c as XmlElement;
                        if (el == null) continue;
                        Color col = Color.Black;
                        foreach (XmlNode cc in el.ChildNodes)
                        {
                            var e2 = cc as XmlElement;
                            if (e2 == null) continue;
                            string v = e2.LocalName == "sysClr" ? e2.GetAttribute("lastClr") : e2.GetAttribute("val");
                            col = Hex(v, Color.Black);
                        }
                        list.Add(col);
                    }
                    break;
                }
            }
            if (list.Count < 12)
            {
                list = new List<Color>();
                foreach (var h in new[] { "000000", "FFFFFF", "44546A", "E7E6E6", "4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47", "0563C1", "954F72" })
                    list.Add(Hex(h, Color.Black));
            }
            // 엑셀 테마 인덱스: 0=lt1 1=dk1 2=lt2 3=dk2
            var t0 = list[0]; list[0] = list[1]; list[1] = t0;
            var t2 = list[2]; list[2] = list[3]; list[3] = t2;
            return list;
        }

        static Color Hex(string v, Color def)
        {
            if (string.IsNullOrEmpty(v)) return def;
            if (v.Length == 8) v = v.Substring(2);
            int x;
            if (v.Length == 6 && int.TryParse(v, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out x))
                return Color.FromArgb(255, (x >> 16) & 255, (x >> 8) & 255, x & 255);
            return def;
        }

        static bool ParseColor(XmlElement el, List<Color> theme, out Color col)
        {
            col = Color.Black;
            if (el == null) return false;
            if (el.GetAttribute("auto") == "1") return false;
            bool ok = false;
            if (el.HasAttribute("rgb")) { col = Hex(el.GetAttribute("rgb"), Color.Black); ok = true; }
            else if (el.HasAttribute("theme"))
            {
                int t;
                if (int.TryParse(el.GetAttribute("theme"), out t) && t >= 0 && t < theme.Count) { col = theme[t]; ok = true; }
            }
            else if (el.HasAttribute("indexed"))
            {
                int i;
                if (int.TryParse(el.GetAttribute("indexed"), out i) && i >= 0 && i < Indexed.Length)
                { int x = Indexed[i]; col = Color.FromArgb((x >> 16) & 255, (x >> 8) & 255, x & 255); ok = true; }
            }
            if (ok && el.HasAttribute("tint"))
            {
                double tint;
                if (double.TryParse(el.GetAttribute("tint"), NumberStyles.Float, CultureInfo.InvariantCulture, out tint) && tint != 0)
                    col = ApplyTint(col, tint);
            }
            return ok;
        }

        static Color ApplyTint(Color c, double tint)
        {
            double r = c.R / 255.0, g = c.G / 255.0, b = c.B / 255.0;
            double max = Math.Max(r, Math.Max(g, b)), min = Math.Min(r, Math.Min(g, b));
            double h = 0, s = 0, l = (max + min) / 2;
            if (max != min)
            {
                double d = max - min;
                s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
                if (max == r) h = (g - b) / d + (g < b ? 6 : 0);
                else if (max == g) h = (b - r) / d + 2;
                else h = (r - g) / d + 4;
                h /= 6;
            }
            l = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
            Func<double, double, double, double> hue = (p, q, t) =>
            {
                if (t < 0) t += 1; if (t > 1) t -= 1;
                if (t < 1.0 / 6) return p + (q - p) * 6 * t;
                if (t < 0.5) return q;
                if (t < 2.0 / 3) return p + (q - p) * (2.0 / 3 - t) * 6;
                return p;
            };
            if (s == 0) r = g = b = l;
            else
            {
                double q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q;
                r = hue(p, q, h + 1.0 / 3); g = hue(p, q, h); b = hue(p, q, h - 1.0 / 3);
            }
            Func<double, int> cl = v => Math.Max(0, Math.Min(255, (int)Math.Round(v * 255)));
            return Color.FromArgb(cl(r), cl(g), cl(b));
        }

        static List<string> ReadShared(ZipArchive zip)
        {
            var list = new List<string>();
            var e = FindEntry(zip, "xl/sharedStrings.xml");
            if (e == null) return list;
            using (var st = e.Open())
            using (var xr = XmlReader.Create(st))
                while (xr.Read())
                    if (xr.NodeType == XmlNodeType.Element && xr.LocalName == "si") list.Add(ReadRichText(xr));
            return list;
        }

        static string ReadRichText(XmlReader xr)
        {
            if (xr.IsEmptyElement) return "";
            var sb = new StringBuilder();
            int depth = xr.Depth, phDepth = -1;
            bool adv = false;
            while (adv || xr.Read())
            {
                adv = false;
                if (xr.NodeType == XmlNodeType.EndElement && xr.Depth == depth) break;
                if (xr.NodeType == XmlNodeType.Element)
                {
                    if (xr.LocalName == "rPh" && !xr.IsEmptyElement) phDepth = xr.Depth;
                    else if (xr.LocalName == "t" && phDepth < 0 && !xr.IsEmptyElement) { sb.Append(xr.ReadElementContentAsString()); adv = true; }
                }
                else if (xr.NodeType == XmlNodeType.EndElement && xr.LocalName == "rPh" && xr.Depth == phDepth) phDepth = -1;
            }
            return sb.ToString().Replace("\r\n", "\n").Replace("\r", "\n");
        }

        static XmlElement Child(XmlElement e, string name)
        {
            if (e == null) return null;
            foreach (XmlNode n in e.ChildNodes) { var x = n as XmlElement; if (x != null && x.LocalName == name) return x; }
            return null;
        }

        static List<CellStyle> ReadStyles(ZipArchive zip, List<Color> theme)
        {
            var styles = new List<CellStyle>();
            var doc = Load(zip, "xl/styles.xml");
            if (doc == null) { styles.Add(new CellStyle()); return styles; }
            var fmts = new Dictionary<int, string>();
            foreach (XmlElement f in doc.GetElementsByTagName("numFmt"))
            {
                int id;
                if (int.TryParse(f.GetAttribute("numFmtId"), out id)) fmts[id] = f.GetAttribute("formatCode");
            }
            var fonts = new List<CellStyle>();
            var fontsNode = doc.GetElementsByTagName("fonts");
            if (fontsNode.Count > 0)
                foreach (XmlNode n in fontsNode[0].ChildNodes)
                {
                    var f = n as XmlElement;
                    if (f == null || f.LocalName != "font") continue;
                    var fs = new CellStyle();
                    foreach (XmlNode cn in f.ChildNodes)
                    {
                        var c = cn as XmlElement;
                        if (c == null) continue;
                        string val = c.GetAttribute("val");
                        switch (c.LocalName)
                        {
                            case "b": fs.Bold = val != "0" && val != "false"; break;
                            case "i": fs.Italic = val != "0" && val != "false"; break;
                            case "u": fs.Underline = val != "none"; break;
                            case "strike": fs.Strike = val != "0" && val != "false"; break;
                            case "sz": float sz; if (float.TryParse(val, NumberStyles.Float, CultureInfo.InvariantCulture, out sz)) fs.FontSize = sz; break;
                            case "name": if (val.Length > 0) fs.FontName = val; break;
                            case "color": Color col; if (ParseColor(c, theme, out col)) fs.FontColor = col; break;
                        }
                    }
                    fonts.Add(fs);
                }
            var fills = new List<Color?>();
            var fillsNode = doc.GetElementsByTagName("fills");
            if (fillsNode.Count > 0)
                foreach (XmlNode n in fillsNode[0].ChildNodes)
                {
                    var f = n as XmlElement;
                    if (f == null || f.LocalName != "fill") continue;
                    Color? fc = null;
                    var pf = Child(f, "patternFill");
                    if (pf != null)
                    {
                        string pt = pf.GetAttribute("patternType");
                        if (pt != "" && pt != "none" && pt != "gray125")
                        {
                            Color col;
                            if (ParseColor(Child(pf, "fgColor"), theme, out col)) fc = col;
                            else if (pt == "solid") fc = Color.Black;
                        }
                    }
                    else
                    {
                        var gf = Child(f, "gradientFill");
                        if (gf != null)
                            foreach (XmlNode s in gf.ChildNodes)
                            {
                                var se = s as XmlElement;
                                Color col;
                                if (se != null && ParseColor(Child(se, "color"), theme, out col)) { fc = col; break; }
                            }
                    }
                    fills.Add(fc);
                }
            var borders = new List<XBorder[]>();
            var bNode = doc.GetElementsByTagName("borders");
            if (bNode.Count > 0)
                foreach (XmlNode n in bNode[0].ChildNodes)
                {
                    var b = n as XmlElement;
                    if (b == null || b.LocalName != "border") continue;
                    var arr = new XBorder[4];
                    string[] names = { "left", "right", "top", "bottom" };
                    for (int i = 0; i < 4; i++)
                    {
                        var e = Child(b, names[i]);
                        if (e == null && i == 0) e = Child(b, "start");
                        if (e == null && i == 1) e = Child(b, "end");
                        if (e == null) continue;
                        string st = e.GetAttribute("style");
                        if (st == "" || st == "none") continue;
                        var xb = new XBorder { Style = st };
                        Color col;
                        if (ParseColor(Child(e, "color"), theme, out col)) xb.Color = col;
                        arr[i] = xb;
                    }
                    borders.Add(arr);
                }
            var xfs = doc.GetElementsByTagName("cellXfs");
            if (xfs.Count > 0)
                foreach (XmlNode n in xfs[0].ChildNodes)
                {
                    var x = n as XmlElement;
                    if (x == null || x.LocalName != "xf") continue;
                    var cs = new CellStyle();
                    int fid, fillId, bid, nid;
                    if (int.TryParse(x.GetAttribute("fontId"), out fid) && fid >= 0 && fid < fonts.Count)
                    {
                        var f = fonts[fid];
                        cs.FontName = f.FontName; cs.FontSize = f.FontSize; cs.Bold = f.Bold; cs.Italic = f.Italic;
                        cs.Underline = f.Underline; cs.Strike = f.Strike; cs.FontColor = f.FontColor;
                    }
                    if (int.TryParse(x.GetAttribute("fillId"), out fillId) && fillId >= 0 && fillId < fills.Count && fills[fillId].HasValue)
                    { cs.HasFill = true; cs.Fill = fills[fillId].Value; }
                    if (int.TryParse(x.GetAttribute("borderId"), out bid) && bid >= 0 && bid < borders.Count)
                    { var b = borders[bid]; cs.L = b[0]; cs.R = b[1]; cs.T = b[2]; cs.B = b[3]; }
                    if (int.TryParse(x.GetAttribute("numFmtId"), out nid)) { cs.NumFmtId = nid; string code; if (fmts.TryGetValue(nid, out code)) cs.NumFmt = code; }
                    var al = Child(x, "alignment");
                    if (al != null)
                    {
                        if (al.HasAttribute("horizontal")) cs.HAlign = al.GetAttribute("horizontal");
                        if (al.HasAttribute("vertical")) cs.VAlign = al.GetAttribute("vertical");
                        string w = al.GetAttribute("wrapText");
                        cs.Wrap = w == "1" || w == "true";
                        int ind; if (int.TryParse(al.GetAttribute("indent"), out ind)) cs.Indent = ind;
                    }
                    styles.Add(cs);
                }
            if (styles.Count == 0) styles.Add(fonts.Count > 0 ? fonts[0] : new CellStyle());
            return styles;
        }

        static void ParseRef(string r, out int row, out int col)
        {
            col = 0; int i = 0;
            while (i < r.Length && char.IsLetter(r[i])) { col = col * 26 + (char.ToUpperInvariant(r[i]) - 'A' + 1); i++; }
            col -= 1;
            int rn; row = int.TryParse(r.Substring(i), out rn) ? rn - 1 : -1;
        }

        static double D(string s, double def)
        {
            double d;
            return double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out d) ? d : def;
        }

        static void ReadSheet(ZipArchiveEntry entry, Sheet sh, List<string> shared)
        {
            int maxR = -1, maxC = -1;
            using (var st = entry.Open())
            using (var xr = XmlReader.Create(st))
            {
                int curRow = -1, nextCol = 0;
                bool adv = false;
                while (adv || xr.Read())
                {
                    adv = false;
                    if (xr.NodeType != XmlNodeType.Element) continue;
                    switch (xr.LocalName)
                    {
                        case "sheetView":
                            if (xr.GetAttribute("showGridLines") == "0" || xr.GetAttribute("showGridLines") == "false") sh.ShowGrid = false;
                            break;
                        case "sheetFormatPr":
                            sh.DefRowHtPt = D(xr.GetAttribute("defaultRowHeight"), 15);
                            sh.DefColWidthChars = D(xr.GetAttribute("defaultColWidth"), -1);
                            if (sh.DefColWidthChars < 0)
                            {
                                double bw = D(xr.GetAttribute("baseColWidth"), -1);
                                if (bw > 0) sh.DefColWidthChars = -100 - bw; // 표시: baseColWidth
                            }
                            break;
                        case "col":
                            {
                                int mn = (int)D(xr.GetAttribute("min"), 1), mx = (int)D(xr.GetAttribute("max"), mn);
                                double w = D(xr.GetAttribute("width"), -1);
                                bool hid = xr.GetAttribute("hidden") == "1" || xr.GetAttribute("hidden") == "true";
                                mx = Math.Min(mx, 16384);
                                if (mx - mn < 2000)
                                    for (int c = mn - 1; c < mx; c++) { if (w >= 0) sh.ColW[c] = w; if (hid) sh.HiddenCols.Add(c); }
                                break;
                            }
                        case "row":
                            {
                                string r = xr.GetAttribute("r");
                                int rn;
                                curRow = (r != null && int.TryParse(r, out rn)) ? rn - 1 : curRow + 1;
                                nextCol = 0;
                                double ht = D(xr.GetAttribute("ht"), -1);
                                if (ht >= 0) sh.RowHt[curRow] = ht;
                                string ch = xr.GetAttribute("customHeight");
                                if (ch == "1" || ch == "true") sh.CustomHt.Add(curRow);
                                string hd = xr.GetAttribute("hidden");
                                if (hd == "1" || hd == "true") sh.HiddenRows.Add(curRow);
                                break;
                            }
                        case "mergeCell":
                            {
                                string rf = xr.GetAttribute("ref");
                                if (rf == null || !rf.Contains(":")) break;
                                var p = rf.Split(':');
                                int r1, c1, r2, c2;
                                ParseRef(p[0], out r1, out c1); ParseRef(p[1], out r2, out c2);
                                if (r1 < 0 || c1 < 0 || r2 < r1 || c2 < c1) break;
                                if ((long)(r2 - r1 + 1) * (c2 - c1 + 1) > 200000) break;
                                var m = new Merge { R1 = r1, C1 = c1, R2 = r2, C2 = c2 };
                                sh.Merges.Add(m);
                                for (int rr = r1; rr <= r2; rr++)
                                    for (int cc = c1; cc <= c2; cc++) sh.MergeOf[Sheet.Key(rr, cc)] = m;
                                maxR = Math.Max(maxR, r2); maxC = Math.Max(maxC, c2);
                                break;
                            }
                        case "c":
                            {
                                string rref = xr.GetAttribute("r"), t = xr.GetAttribute("t"), s = xr.GetAttribute("s");
                                int row = curRow, col = nextCol;
                                if (rref != null) ParseRef(rref, out row, out col);
                                nextCol = col + 1;
                                string v = null, inl = null;
                                if (!xr.IsEmptyElement)
                                {
                                    int depth = xr.Depth;
                                    bool a2 = false;
                                    while (a2 || xr.Read())
                                    {
                                        a2 = false;
                                        if (xr.NodeType == XmlNodeType.EndElement && xr.Depth == depth) break;
                                        if (xr.NodeType == XmlNodeType.Element)
                                        {
                                            if (xr.LocalName == "v") { v = xr.ReadElementContentAsString(); a2 = true; }
                                            else if (xr.LocalName == "is") inl = ReadRichText(xr);
                                        }
                                    }
                                }
                                if (row < 0 || col < 0 || col > 16383) break;
                                int si = 0;
                                if (s != null) int.TryParse(s, out si);
                                if (si < 0 || si >= sh.Styles.Count) si = 0;
                                var cell = new Cell { Style = si };
                                cell.Text = Format(t, v, inl, shared, sh.Styles[si], cell);
                                var stl = sh.Styles[si];
                                bool visible = !string.IsNullOrEmpty(cell.Text) || stl.HasFill || stl.HasBorder;
                                if (!visible) break;
                                sh.Cells[Sheet.Key(row, col)] = cell;
                                maxR = Math.Max(maxR, row); maxC = Math.Max(maxC, col);
                                break;
                            }
                    }
                }
            }
            sh.Rows = Math.Max(maxR + 1 + 5, 30);
            sh.Cols = Math.Max(maxC + 1 + 3, 12);
            sh.Cols = Math.Min(sh.Cols, 16384);
        }

        static string Format(string t, string v, string inl, List<string> shared, CellStyle st, Cell cell)
        {
            if (t == "inlineStr") return inl ?? "";
            if (v == null) return inl;
            if (t == "s")
            {
                int idx;
                return int.TryParse(v, out idx) && idx >= 0 && idx < shared.Count ? shared[idx] : "";
            }
            if (t == "b") { cell.IsBool = true; return v == "1" ? "TRUE" : "FALSE"; }
            if (t == "str" || t == "e") return v;
            double d;
            if (!double.TryParse(v, NumberStyles.Float, CultureInfo.InvariantCulture, out d)) return v;
            cell.IsNum = true;
            return NumFormat.Apply(d, st.NumFmtId, st.NumFmt);
        }
    }

    static class NumFormat
    {
        static string Clean(string code)
        {
            if (code == null) return "";
            int semi = code.IndexOf(';');
            if (semi >= 0) code = code.Substring(0, semi);
            var sb = new StringBuilder();
            bool q = false, br = false;
            foreach (char c in code)
            {
                if (c == '"') { q = !q; continue; }
                if (q) continue;
                if (c == '[') { br = true; continue; }
                if (c == ']') { br = false; continue; }
                if (br || c == '\\' || c == '_' || c == '*') continue;
                sb.Append(c);
            }
            return sb.ToString();
        }

        static string Literal(string code, bool prefix)
        {
            // "원" 같은 접미/접두 문자열 추출
            if (code == null) return "";
            int semi = code.IndexOf(';');
            if (semi >= 0) code = code.Substring(0, semi);
            var m = Regex.Matches(code, "\"([^\"]*)\"");
            if (m.Count == 0) return "";
            int firstDigit = code.IndexOfAny("0#?".ToCharArray());
            string res = "";
            foreach (Match x in m)
                if (prefix ? x.Index < firstDigit : x.Index > firstDigit) res += x.Groups[1].Value;
            return res;
        }

        public static string Apply(double d, int id, string code)
        {
            string c = Clean(code);
            string lc = c.ToLowerInvariant();
            if (lc == "@") return d.ToString("G15", CultureInfo.InvariantCulture);
            bool isDate = (id >= 14 && id <= 22) || (id >= 27 && id <= 36) || (id >= 45 && id <= 47) || (id >= 50 && id <= 58)
                || (code != null && id >= 164 && !lc.Contains("0") && !lc.Contains("#") && (lc.Contains("y") || lc.Contains("d") || lc.Contains("h") || lc.Contains("s") || lc.Contains("m")));
            if (isDate && d >= 0 && d < 2958466)
            {
                try
                {
                    var dt = DateTime.FromOADate(d);
                    bool hasTime = (id >= 18 && id <= 22) || (id >= 45 && id <= 47) || lc.Contains("h") || lc.Contains("s");
                    bool hasDate = !(id >= 18 && id <= 21) && !(id >= 45 && id <= 47) && (lc.Contains("y") || lc.Contains("d") || id < 164 || !hasTime);
                    string sec = lc.Contains("s") ? ":ss" : "";
                    if (hasDate && hasTime) return dt.ToString("yyyy-MM-dd HH:mm" + sec);
                    if (hasTime) return dt.ToString("HH:mm" + sec);
                    if (lc.Contains("y") && !lc.Contains("d")) return dt.ToString("yyyy-MM");
                    if (!lc.Contains("y") && lc.Contains("d") && id >= 164) return dt.ToString("MM-dd");
                    return dt.ToString("yyyy-MM-dd");
                }
                catch { }
            }
            bool pct = id == 9 || id == 10 || c.Contains("%");
            if (pct) d *= 100;
            bool comma = id == 3 || id == 4 || (id >= 37 && id <= 40) || c.Contains(",");
            int dec = -1;
            if (id == 1 || id == 3 || id == 9 || id == 37 || id == 38) dec = 0;
            else if (id == 2 || id == 4 || id == 10 || id == 39 || id == 40) dec = 2;
            else if (c.Contains("0") || c.Contains("#"))
            {
                int dot = c.IndexOf('.');
                dec = 0;
                if (dot >= 0) for (int i = dot + 1; i < c.Length && (c[i] == '0' || c[i] == '#'); i++) dec++;
            }
            string res;
            if (dec < 0)
            {
                res = d.ToString("G15", CultureInfo.InvariantCulture);
                if (res.Contains("E") && Math.Abs(d) >= 1e-5 && Math.Abs(d) < 1e15)
                    res = ((decimal)d).ToString(CultureInfo.InvariantCulture);
                if (comma) res = d.ToString("#,##0.##########", CultureInfo.InvariantCulture);
            }
            else res = d.ToString((comma ? "N" : "F") + dec, CultureInfo.InvariantCulture);
            if (pct) res += "%";
            if (code != null && id >= 164) res = Literal(code, true) + res + Literal(code, false);
            return res;
        }
    }

    static class CsvReader
    {
        public static List<Sheet> Read(string path)
        {
            byte[] bytes = File.ReadAllBytes(path);
            string text;
            try { text = new UTF8Encoding(false, true).GetString(bytes); }
            catch { text = Encoding.GetEncoding(949).GetString(bytes); }
            if (text.Length > 0 && text[0] == '﻿') text = text.Substring(1);
            char sep = path.EndsWith(".tsv", StringComparison.OrdinalIgnoreCase) ? '\t' : ',';
            var sh = new Sheet { Name = Path.GetFileNameWithoutExtension(path), Styles = new List<CellStyle> { new CellStyle() } };
            int r = 0, c = 0, maxC = 0;
            var cell = new StringBuilder();
            bool q = false;
            Action put = () =>
            {
                string v = cell.ToString();
                if (v.Length > 0)
                {
                    double d;
                    sh.Cells[Sheet.Key(r, c)] = new Cell { Text = v, IsNum = double.TryParse(v.Replace(",", ""), NumberStyles.Float, CultureInfo.InvariantCulture, out d) };
                    maxC = Math.Max(maxC, c);
                }
                cell.Clear();
            };
            for (int i = 0; i < text.Length; i++)
            {
                char ch = text[i];
                if (q)
                {
                    if (ch == '"') { if (i + 1 < text.Length && text[i + 1] == '"') { cell.Append('"'); i++; } else q = false; }
                    else cell.Append(ch);
                }
                else if (ch == '"') q = true;
                else if (ch == sep) { put(); c++; }
                else if (ch == '\r') { }
                else if (ch == '\n') { put(); r++; c = 0; }
                else cell.Append(ch);
            }
            put();
            sh.Rows = Math.Max(r + 6, 30);
            sh.Cols = Math.Max(maxC + 4, 12);
            return new List<Sheet> { sh };
        }
    }

    // ───────────────────────── 레이아웃 + 그리기 ─────────────────────────
    static class Render
    {
        public static void Layout(Sheet s)
        {
            var defFont = Fonts.For(s.Styles[0], 1f);
            string dn = s.Styles[0].FontName ?? "";
            float dsz = s.Styles[0].FontSize;
            int mdw;
            if ((dn == "맑은 고딕" || dn.Equals("Malgun Gothic", StringComparison.OrdinalIgnoreCase)) && Math.Abs(dsz - 11) < 0.1) mdw = 8;
            else if ((dn == "Calibri" || dn == "Arial") && Math.Abs(dsz - 11) < 0.1) mdw = 7;
            else if (dn == "돋움" || dn == "굴림" || dn.Equals("Dotum", StringComparison.OrdinalIgnoreCase) || dn.Equals("Gulim", StringComparison.OrdinalIgnoreCase)) mdw = dsz <= 10 ? 7 : 8;
            else
            {
                mdw = 7;
                try { using (var bmp = new Bitmap(1, 1)) using (var g = Graphics.FromImage(bmp)) mdw = Math.Max(5, (int)Math.Round(g.MeasureString("0000000000", defFont, 1000, StringFormat.GenericTypographic).Width / 10)); } catch { }
            }
            s.Mdw = mdw;
            Func<double, float> chars2px = w => (float)Math.Truncate(((256 * w + Math.Truncate(128.0 / mdw)) / 256) * mdw);
            float defCol;
            if (s.DefColWidthChars > 0) defCol = chars2px(s.DefColWidthChars);
            else
            {
                double bw = s.DefColWidthChars <= -100 ? -100 - s.DefColWidthChars : 8;
                defCol = (float)(bw * mdw + 5);
            }
            float defRow = (float)Math.Round(s.DefRowHtPt * 96 / 72);
            s.CW = new float[s.Cols];
            s.RH = new float[s.Rows];
            for (int c = 0; c < s.Cols; c++)
            {
                double w;
                s.CW[c] = s.HiddenCols.Contains(c) ? 0 : (s.ColW.TryGetValue(c, out w) ? chars2px(w) : defCol);
            }
            for (int r = 0; r < s.Rows; r++)
            {
                double h;
                s.RH[r] = s.HiddenRows.Contains(r) ? 0 : (s.RowHt.TryGetValue(r, out h) ? (float)Math.Round(h * 96 / 72) : defRow);
            }
            // 내용이 다 보이도록 행 높이 늘리기
            var flags = TextFormatFlags.WordBreak | TextFormatFlags.NoPrefix | TextFormatFlags.TextBoxControl | TextFormatFlags.NoPadding;
            foreach (var kv in s.Cells)
            {
                var cell = kv.Value;
                if (string.IsNullOrEmpty(cell.Text)) continue;
                int r = (int)(kv.Key >> 16), c = (int)(kv.Key & 0xFFFF);
                if (r >= s.Rows || c >= s.Cols || s.RH[r] == 0) continue;
                var st = s.StyleOf(cell);
                var m = s.MergeAt(r, c);
                if (m != null && (m.R1 != r || m.C1 != c)) continue;
                var font = Fonts.For(st, 1f);
                if (m == null)
                {
                    float need;
                    if (st.Wrap)
                    {
                        int w = Math.Max(4, (int)s.CW[c] - 6 - st.Indent * 9);
                        need = TextRenderer.MeasureText(cell.Text, font, new Size(w, int.MaxValue), flags).Height + 3;
                    }
                    else need = font.Height + 2;
                    if (need > s.RH[r] && !s.RowHt.ContainsKey(r)) s.RH[r] = need;
                }
                else
                {
                    float w = 0, h = 0;
                    for (int cc = m.C1; cc <= m.C2 && cc < s.Cols; cc++) w += s.CW[cc];
                    int last = -1;
                    bool stored = false;
                    for (int rr = m.R1; rr <= m.R2 && rr < s.Rows; rr++) { h += s.RH[rr]; if (s.RH[rr] > 0) last = rr; if (s.RowHt.ContainsKey(rr)) stored = true; }
                    if (last < 0 || stored) continue;
                    float need;
                    if (st.Wrap) need = TextRenderer.MeasureText(cell.Text, font, new Size(Math.Max(4, (int)w - 6), int.MaxValue), flags).Height + 3;
                    else need = font.Height + 2;
                    if (need > h) s.RH[last] += need - h;
                }
            }
            foreach (var kv in s.EditCols) if (kv.Key < s.Cols) s.CW[kv.Key] = kv.Value;
            foreach (var kv in s.EditRows) if (kv.Key < s.Rows) s.RH[kv.Key] = kv.Value;
            s.Recalc();
            var order = new List<long>();
            foreach (var kv in s.Cells) if (!string.IsNullOrEmpty(kv.Value.Text)) order.Add(kv.Key);
            order.Sort();
            s.TextOrder = order;
        }

        public static int Find(float[] arr, float v)
        {
            int lo = 0, hi = arr.Length - 2;
            if (hi < 0) return 0;
            while (lo < hi)
            {
                int mid = (lo + hi + 1) / 2;
                if (arr[mid] <= v) lo = mid; else hi = mid - 1;
            }
            return lo;
        }

        static TextFormatFlags Align(CellStyle st, Cell cell, out bool left)
        {
            var f = TextFormatFlags.NoPrefix | TextFormatFlags.NoPadding;
            string h = st.HAlign;
            if (h == "general") h = cell.IsBool ? "center" : cell.IsNum ? "right" : "left";
            left = false;
            if (h == "center" || h == "centerContinuous" || h == "distributed") f |= TextFormatFlags.HorizontalCenter;
            else if (h == "right") f |= TextFormatFlags.Right;
            else left = true;
            if (st.VAlign == "top") f |= TextFormatFlags.Top;
            else if (st.VAlign == "center" || st.VAlign == "justify" || st.VAlign == "distributed") f |= TextFormatFlags.VerticalCenter;
            else f |= TextFormatFlags.Bottom;
            return f;
        }

        // 화면 x = ox + X[c]*z
        public static void Draw(Graphics g, Sheet s, int r1, int r2, int c1, int c2, float ox, float oy, float z, bool grid)
        {
            Func<int, int> CX = c => (int)Math.Round(ox + s.X[c] * z);
            Func<int, int> CY = r => (int)Math.Round(oy + s.Y[r] * z);
            int left = CX(c1), top = CY(r1), right = CX(c2 + 1), bottom = CY(r2 + 1);
            g.FillRectangle(Brushes.White, left, top, right - left + 1, bottom - top + 1);

            if (grid)
            {
                using (var p = new Pen(Color.FromArgb(222, 222, 222)))
                {
                    for (int c = c1; c <= c2 + 1; c++) g.DrawLine(p, CX(c), top, CX(c), bottom);
                    for (int r = r1; r <= r2 + 1; r++) g.DrawLine(p, left, CY(r), right, CY(r));
                }
            }
            // 배경색
            for (int r = r1; r <= r2; r++)
            {
                if (s.RH[r] == 0) continue;
                for (int c = c1; c <= c2; c++)
                {
                    var cell = s.Get(r, c);
                    if (cell == null || s.MergeAt(r, c) != null) continue;
                    var st = s.StyleOf(cell);
                    if (!st.HasFill || s.CW[c] == 0) continue;
                    using (var b = new SolidBrush(st.Fill)) g.FillRectangle(b, CX(c), CY(r), CX(c + 1) - CX(c) + 1, CY(r + 1) - CY(r) + 1);
                }
            }
            var merges = new List<Merge>();
            foreach (var m in s.Merges)
                if (m.R2 >= r1 && m.R1 <= r2 && m.C2 >= c1 && m.C1 <= c2 && m.R1 < s.Rows && m.C1 < s.Cols) merges.Add(m);
            foreach (var m in merges)
            {
                int x1 = CX(m.C1), y1 = CY(m.R1), x2 = CX(Math.Min(m.C2 + 1, s.Cols)), y2 = CY(Math.Min(m.R2 + 1, s.Rows));
                var st = s.StyleOf(s.Get(m.R1, m.C1));
                if (st.HasFill) using (var b = new SolidBrush(st.Fill)) g.FillRectangle(b, x1, y1, x2 - x1 + 1, y2 - y1 + 1);
                else g.FillRectangle(Brushes.White, x1 + 1, y1 + 1, x2 - x1 - 1, y2 - y1 - 1);
            }
            // 글자
            int cStart = Math.Max(0, c1 - 20);
            for (int r = r1; r <= r2; r++)
            {
                if (s.RH[r] == 0) continue;
                for (int c = cStart; c <= c2; c++)
                {
                    var cell = s.Get(r, c);
                    if (cell == null || string.IsNullOrEmpty(cell.Text) || s.CW[c] == 0) continue;
                    if (s.MergeAt(r, c) != null) continue;
                    DrawText(g, s, cell, r, c, CX(c), CY(r), CX(c + 1), CY(r + 1), z, true, c2);
                }
            }
            foreach (var m in merges)
            {
                var cell = s.Get(m.R1, m.C1);
                if (cell == null || string.IsNullOrEmpty(cell.Text)) continue;
                DrawText(g, s, cell, m.R1, m.C1, CX(m.C1), CY(m.R1), CX(Math.Min(m.C2 + 1, s.Cols)), CY(Math.Min(m.R2 + 1, s.Rows)), z, false, c2);
            }
            // 테두리
            for (int r = r1; r <= r2; r++)
            {
                if (s.RH[r] == 0) continue;
                for (int c = c1; c <= c2; c++)
                {
                    var cell = s.Get(r, c);
                    if (cell == null || s.CW[c] == 0) continue;
                    var st = s.StyleOf(cell);
                    if (!st.HasBorder) continue;
                    var m = s.MergeAt(r, c);
                    int x1 = CX(c), y1 = CY(r), x2 = CX(c + 1), y2 = CY(r + 1);
                    if (st.L != null && (m == null || c == m.C1)) Edge(g, st.L, x1, y1, x1, y2, z);
                    if (st.R != null && (m == null || c == m.C2)) Edge(g, st.R, x2, y1, x2, y2, z);
                    if (st.T != null && (m == null || r == m.R1)) Edge(g, st.T, x1, y1, x2, y1, z);
                    if (st.B != null && (m == null || r == m.R2)) Edge(g, st.B, x1, y2, x2, y2, z);
                }
            }
        }

        static void DrawText(Graphics g, Sheet s, Cell cell, int r, int c, int x1, int y1, int x2, int y2, float z, bool allowOverflow, int c2)
        {
            var st = s.StyleOf(cell);
            var font = Fonts.For(st, z);
            bool isLeft;
            var flags = Align(st, cell, out isLeft);
            string text = cell.Text;
            if (st.Wrap) flags |= TextFormatFlags.WordBreak | TextFormatFlags.TextBoxControl;
            else { text = text.Replace("\n", " "); flags |= TextFormatFlags.SingleLine; }
            int pad = Math.Max(2, (int)Math.Round(3 * z));
            int ind = (int)Math.Round(st.Indent * 9 * z);
            var rect = new Rectangle(x1 + pad + (isLeft ? ind : 0), y1 + 1, Math.Max(1, x2 - x1 - pad * 2 - ind), Math.Max(1, y2 - y1 - 2));
            if (allowOverflow && !st.Wrap && isLeft && !cell.IsNum)
            {
                int need = TextRenderer.MeasureText(g, text, font, Size.Empty, TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix | TextFormatFlags.SingleLine).Width;
                int cc = c + 1;
                while (need > rect.Width && cc < s.Cols && cc <= c2 + 30)
                {
                    var n = s.Get(r, cc);
                    if ((n != null && !string.IsNullOrEmpty(n.Text)) || s.MergeAt(r, cc) != null) break;
                    rect.Width += (int)Math.Round(s.CW[cc] * z);
                    cc++;
                }
            }
            if (!st.Wrap && (flags & TextFormatFlags.Bottom) != 0) rect.Height = Math.Max(1, rect.Height);
            TextRenderer.DrawText(g, text, font, rect, st.FontColor, flags);
        }

        static void Edge(Graphics g, XBorder b, int x1, int y1, int x2, int y2, float z)
        {
            bool vert = x1 == x2;
            int w = 1;
            string s = b.Style;
            if (s.StartsWith("medium")) w = 2;
            else if (s == "thick") w = 3;
            w = Math.Max(1, (int)Math.Round(w * Math.Min(z, 2f)));
            if (s == "double")
            {
                using (var br = new SolidBrush(b.Color))
                {
                    if (vert) { g.FillRectangle(br, x1 - 1, y1 - 1, 1, y2 - y1 + 3); g.FillRectangle(br, x1 + 1, y1 - 1, 1, y2 - y1 + 3); }
                    else { g.FillRectangle(br, x1 - 1, y1 - 1, x2 - x1 + 3, 1); g.FillRectangle(br, x1 - 1, y1 + 1, x2 - x1 + 3, 1); }
                }
                return;
            }
            bool dashed = s.Contains("ash") || s == "dotted" || s == "hair";
            if (!dashed)
            {
                using (var br = new SolidBrush(b.Color))
                {
                    int o = (w - 1) / 2;
                    if (vert) g.FillRectangle(br, x1 - o, y1 - o, w, y2 - y1 + w);
                    else g.FillRectangle(br, x1 - o, y1 - o, x2 - x1 + w, w);
                }
                return;
            }
            using (var p = new Pen(b.Color, w))
            {
                if (s == "dotted" || s == "hair") p.DashPattern = new float[] { 1, 1 };
                else if (s.Contains("DashDotDot")) p.DashPattern = new float[] { 3, 1, 1, 1, 1, 1 };
                else if (s.Contains("DashDot")) p.DashPattern = new float[] { 3, 1, 1, 1 };
                else p.DashPattern = new float[] { 3, 1 };
                g.DrawLine(p, x1, y1, x2, y2);
            }
        }

        public static string ColName(int i)
        {
            string s = "";
            i++;
            while (i > 0) { int m = (i - 1) % 26; s = (char)('A' + m) + s; i = (i - 1) / 26; }
            return s;
        }
    }

    // ───────────────────────── 시트 보기 컨트롤 ─────────────────────────
    class SheetView : Control
    {
        public Sheet Sheet;
        public float Zoom = 1f;
        float dpi = 1f;
        VScrollBar vs = new VScrollBar();
        HScrollBar hs = new HScrollBar();
        int ar = -1, ac = -1, cr = -1, cc = -1;   // 선택: 기준(anchor) / 현재
        bool dragging;
        int dragMode; // 0 셀, 1 열, 2 행
        public event EventHandler SelectionChanged;
        public event EventHandler ZoomChanged;
        public event EventHandler Edited;
        Font headerFont;
        bool editMode;
        int resizeCol = -1, resizeRow = -1, resizeStart; float resizeOrig;
        TextBox editor;
        int edR = -1, edC = -1;

        public bool EditMode
        {
            get { return editMode; }
            set { CommitEdit(); editMode = value; Cursor = Cursors.Default; Invalidate(); }
        }

        void RaiseEdited() { if (Edited != null) Edited(this, EventArgs.Empty); }

        public SheetView()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint | ControlStyles.ResizeRedraw | ControlStyles.Selectable, true);
            TabStop = true;
            BackColor = Color.White;
            Controls.Add(vs); Controls.Add(hs);
            vs.Scroll += (s, e) => { CommitEdit(); Invalidate(); };
            hs.Scroll += (s, e) => { CommitEdit(); Invalidate(); };
            vs.ValueChanged += (s, e) => Invalidate();
            hs.ValueChanged += (s, e) => Invalidate();
            dpi = Program.Dpi;
            headerFont = new Font("맑은 고딕", 12f * dpi, GraphicsUnit.Pixel);
        }

        float Z { get { return Zoom * dpi; } }
        int HW { get { return (int)(Math.Max(42, 14 + 8 * (Sheet == null ? 2 : Sheet.Rows.ToString().Length)) * dpi); } }
        int HH { get { return (int)(22 * dpi); } }
        int ViewW { get { return Math.Max(1, ClientSize.Width - vs.Width - HW); } }
        int ViewH { get { return Math.Max(1, ClientSize.Height - hs.Height - HH); } }

        public void SetSheet(Sheet s)
        {
            CancelEdit();
            Sheet = s;
            ar = ac = cr = cc = -1;
            vs.Value = 0; hs.Value = 0;
            UpdateBars();
            Invalidate();
            if (SelectionChanged != null) SelectionChanged(this, EventArgs.Empty);
        }

        public void SetZoom(float z)
        {
            z = Math.Max(0.5f, Math.Min(3f, z));
            if (Math.Abs(z - Zoom) < 0.001f) return;
            float fx = hs.Value / Z, fy = vs.Value / Z;
            Zoom = z;
            UpdateBars();
            SetVal(hs, (int)(fx * Z)); SetVal(vs, (int)(fy * Z));
            Invalidate();
            if (ZoomChanged != null) ZoomChanged(this, EventArgs.Empty);
        }

        protected override void OnLayout(LayoutEventArgs e)
        {
            base.OnLayout(e);
            vs.SetBounds(ClientSize.Width - vs.Width, 0, vs.Width, Math.Max(0, ClientSize.Height - hs.Height));
            hs.SetBounds(0, ClientSize.Height - hs.Height, Math.Max(0, ClientSize.Width - vs.Width), hs.Height);
            UpdateBars();
        }

        void UpdateBars()
        {
            if (Sheet == null || Sheet.X == null) { vs.Enabled = hs.Enabled = false; return; }
            int tw = (int)(Sheet.X[Sheet.Cols] * Z) + 40, th = (int)(Sheet.Y[Sheet.Rows] * Z) + 40;
            Bar(hs, tw, ViewW, (int)(60 * dpi));
            Bar(vs, th, ViewH, (int)(20 * dpi));
        }

        static void Bar(ScrollBar b, int total, int view, int small)
        {
            if (total <= view) { b.Value = 0; b.Enabled = false; b.Maximum = 0; return; }
            b.Enabled = true;
            b.Minimum = 0;
            b.LargeChange = view;
            b.SmallChange = Math.Max(1, small);
            b.Maximum = total - 1;
            if (b.Value > total - view) b.Value = total - view;
        }

        static void SetVal(ScrollBar b, int v)
        {
            if (!b.Enabled) { b.Value = 0; return; }
            v = Math.Max(0, Math.Min(v, b.Maximum - b.LargeChange + 1));
            b.Value = v;
        }

        public bool HasSelection { get { return ar >= 0; } }
        public bool IsEditing { get { return editor != null; } }

        public void GetSel(out int r1, out int c1, out int r2, out int c2)
        {
            r1 = Math.Min(ar, cr); r2 = Math.Max(ar, cr); c1 = Math.Min(ac, cc); c2 = Math.Max(ac, cc);
            if (Sheet == null) return;
            bool changed = true;
            while (changed)
            {
                changed = false;
                foreach (var m in Sheet.Merges)
                {
                    if (m.R2 < r1 || m.R1 > r2 || m.C2 < c1 || m.C1 > c2) continue;
                    if (m.R1 < r1) { r1 = m.R1; changed = true; }
                    int nr = Math.Min(m.R2, Sheet.Rows - 1); if (nr > r2) { r2 = nr; changed = true; }
                    if (m.C1 < c1) { c1 = m.C1; changed = true; }
                    int nc = Math.Min(m.C2, Sheet.Cols - 1); if (nc > c2) { c2 = nc; changed = true; }
                }
            }
        }

        public void SelectAll()
        {
            if (Sheet == null) return;
            int maxR = 0, maxC = 0;
            foreach (var k in Sheet.Cells.Keys) { maxR = Math.Max(maxR, (int)(k >> 16)); maxC = Math.Max(maxC, (int)(k & 0xFFFF)); }
            foreach (var m in Sheet.Merges) { maxR = Math.Max(maxR, m.R2); maxC = Math.Max(maxC, m.C2); }
            ar = 0; ac = 0; cr = Math.Min(maxR, Sheet.Rows - 1); cc = Math.Min(maxC, Sheet.Cols - 1);
            Invalidate();
            if (SelectionChanged != null) SelectionChanged(this, EventArgs.Empty);
        }

        public void SelectCell(int r, int c)
        {
            ar = cr = r; ac = cc = c;
            EnsureVisible(r, c);
            Invalidate();
            if (SelectionChanged != null) SelectionChanged(this, EventArgs.Empty);
        }

        void EnsureVisible(int r, int c)
        {
            if (Sheet == null) return;
            int x1 = (int)(Sheet.X[c] * Z), x2 = (int)(Sheet.X[c + 1] * Z), y1 = (int)(Sheet.Y[r] * Z), y2 = (int)(Sheet.Y[r + 1] * Z);
            if (x1 < hs.Value) SetVal(hs, x1); else if (x2 > hs.Value + ViewW) SetVal(hs, x2 - ViewW);
            if (y1 < vs.Value) SetVal(vs, y1); else if (y2 > vs.Value + ViewH) SetVal(vs, y2 - ViewH);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            try { PaintAll(e); }
            catch (Exception ex)
            {
                e.Graphics.ResetClip();
                TextRenderer.DrawText(e.Graphics, "화면 그리기 오류: " + ex.Message, Font, new Point(10, 40), Color.Red);
                Console.Error.WriteLine(ex);
            }
        }

        void PaintAll(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.Clear(Color.White);
            if (Sheet == null || Sheet.X == null) return;
            float z = Z;
            int hw = HW, hh = HH;
            float ox = hw - hs.Value, oy = hh - vs.Value;
            int cA = Render.Find(Sheet.X, hs.Value / z), cB = Render.Find(Sheet.X, (hs.Value + ViewW) / z);
            int rA = Render.Find(Sheet.Y, vs.Value / z), rB = Render.Find(Sheet.Y, (vs.Value + ViewH) / z);
            cB = Math.Min(cB, Sheet.Cols - 1); rB = Math.Min(rB, Sheet.Rows - 1);
            var area = new Rectangle(hw, hh, ViewW, ViewH);
            g.SetClip(area);
            Render.Draw(g, Sheet, rA, rB, cA, cB, ox, oy, z, Sheet.ShowGrid || true);
            int sr1 = -1, sc1 = -1, sr2 = -1, sc2 = -1;
            if (HasSelection)
            {
                GetSel(out sr1, out sc1, out sr2, out sc2);
                int x1 = (int)Math.Round(ox + Sheet.X[sc1] * z), x2 = (int)Math.Round(ox + Sheet.X[sc2 + 1] * z);
                int y1 = (int)Math.Round(oy + Sheet.Y[sr1] * z), y2 = (int)Math.Round(oy + Sheet.Y[sr2 + 1] * z);
                using (var b = new SolidBrush(Color.FromArgb(45, 33, 115, 70))) g.FillRectangle(b, x1, y1, x2 - x1, y2 - y1);
                using (var p = new Pen(AppInfo.Accent, Math.Max(2, 2 * dpi))) g.DrawRectangle(p, x1, y1, x2 - x1, y2 - y1);
            }
            g.ResetClip();
            // 머리글
            var hb = editMode ? Color.FromArgb(255, 243, 214) : Color.FromArgb(243, 244, 246);
            var hsel = Color.FromArgb(211, 230, 218);
            using (var bg = new SolidBrush(hb)) { g.FillRectangle(bg, 0, 0, ClientSize.Width - vs.Width, hh); g.FillRectangle(bg, 0, 0, hw, ClientSize.Height - hs.Height); }
            using (var line = new Pen(Color.FromArgb(210, 212, 216)))
            using (var selB = new SolidBrush(hsel))
            {
                g.SetClip(new Rectangle(hw, 0, ViewW, hh));
                for (int c = cA; c <= cB; c++)
                {
                    int x1 = (int)Math.Round(ox + Sheet.X[c] * z), x2 = (int)Math.Round(ox + Sheet.X[c + 1] * z);
                    if (x2 <= x1) continue;
                    if (HasSelection && c >= sc1 && c <= sc2) g.FillRectangle(selB, x1, 0, x2 - x1, hh);
                    TextRenderer.DrawText(g, Render.ColName(c), headerFont, new Rectangle(x1, 0, x2 - x1, hh), Color.FromArgb(80, 80, 80), TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding);
                    g.DrawLine(line, x2, 0, x2, hh);
                }
                g.SetClip(new Rectangle(0, hh, hw, ViewH));
                for (int r = rA; r <= rB; r++)
                {
                    int y1 = (int)Math.Round(oy + Sheet.Y[r] * z), y2 = (int)Math.Round(oy + Sheet.Y[r + 1] * z);
                    if (y2 <= y1) continue;
                    if (HasSelection && r >= sr1 && r <= sr2) g.FillRectangle(selB, 0, y1, hw, y2 - y1);
                    TextRenderer.DrawText(g, (r + 1).ToString(), headerFont, new Rectangle(0, y1, hw - 4, y2 - y1), Color.FromArgb(80, 80, 80), TextFormatFlags.Right | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding);
                    g.DrawLine(line, 0, y2, hw, y2);
                }
                g.ResetClip();
                g.DrawLine(line, 0, hh, ClientSize.Width, hh);
                g.DrawLine(line, hw, 0, hw, ClientSize.Height);
            }
            using (var bg = new SolidBrush(SystemColors.Control)) g.FillRectangle(bg, ClientSize.Width - vs.Width, ClientSize.Height - hs.Height, vs.Width, hs.Height);
        }

        bool Hit(Point p, out int r, out int c, out int zone)
        {
            r = c = -1; zone = 0;
            if (Sheet == null || Sheet.X == null) return false;
            float z = Z;
            int hw = HW, hh = HH;
            if (p.X >= ClientSize.Width - vs.Width || p.Y >= ClientSize.Height - hs.Height) return false;
            float ex = (p.X - hw + hs.Value) / z, ey = (p.Y - hh + vs.Value) / z;
            c = Math.Min(Sheet.Cols - 1, Math.Max(0, Render.Find(Sheet.X, Math.Max(0, ex))));
            r = Math.Min(Sheet.Rows - 1, Math.Max(0, Render.Find(Sheet.Y, Math.Max(0, ey))));
            if (p.X < hw && p.Y < hh) zone = 3;
            else if (p.Y < hh) zone = 1;
            else if (p.X < hw) zone = 2;
            return true;
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            CommitEdit();
            Focus();
            if (e.Button != MouseButtons.Left) return;
            if (editMode && Sheet != null)
            {
                int bc = ColBorderAt(e.Location), br = RowBorderAt(e.Location);
                if (bc >= 0) { resizeCol = bc; resizeStart = e.X; resizeOrig = Sheet.CW[bc]; Capture = true; return; }
                if (br >= 0) { resizeRow = br; resizeStart = e.Y; resizeOrig = Sheet.RH[br]; Capture = true; return; }
            }
            int r, c, zone;
            if (!Hit(e.Location, out r, out c, out zone)) return;
            if (editMode && zone == 0 && e.Clicks == 2) { BeginEdit(r, c); return; }
            if (zone == 3) { SelectAll(); return; }
            bool shift = (ModifierKeys & Keys.Shift) != 0 && HasSelection;
            dragMode = zone;
            if (zone == 1) { if (!shift) { ac = c; ar = 0; } cc = c; cr = Sheet.Rows - 1; }
            else if (zone == 2) { if (!shift) { ar = r; ac = 0; } cr = r; cc = Sheet.Cols - 1; }
            else { if (!shift) { ar = r; ac = c; } cr = r; cc = c; }
            dragging = true;
            Capture = true;
            Invalidate();
            if (SelectionChanged != null) SelectionChanged(this, EventArgs.Empty);
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            if (resizeCol >= 0 || resizeRow >= 0)
            {
                if (resizeCol >= 0) Sheet.CW[resizeCol] = Math.Max(6, resizeOrig + (e.X - resizeStart) / Z);
                else Sheet.RH[resizeRow] = Math.Max(6, resizeOrig + (e.Y - resizeStart) / Z);
                Sheet.Recalc(); UpdateBars(); Invalidate();
                return;
            }
            if (editMode && !dragging && Sheet != null)
                Cursor = ColBorderAt(e.Location) >= 0 ? Cursors.SizeWE : RowBorderAt(e.Location) >= 0 ? Cursors.SizeNS : Cursors.Default;
            if (!dragging) return;
            var p = e.Location;
            int step = (int)(25 * dpi);
            if (p.Y > ClientSize.Height - hs.Height) SetVal(vs, vs.Value + step);
            else if (p.Y < HH && dragMode != 1) SetVal(vs, vs.Value - step);
            if (p.X > ClientSize.Width - vs.Width) SetVal(hs, hs.Value + step);
            else if (p.X < HW && dragMode != 2) SetVal(hs, hs.Value - step);
            p.X = Math.Max(HW, Math.Min(p.X, ClientSize.Width - vs.Width - 1));
            p.Y = Math.Max(HH, Math.Min(p.Y, ClientSize.Height - hs.Height - 1));
            int r, c, zone;
            if (!Hit(p, out r, out c, out zone)) return;
            if (dragMode == 1) cc = c;
            else if (dragMode == 2) cr = r;
            else { cr = r; cc = c; }
            Invalidate();
            if (SelectionChanged != null) SelectionChanged(this, EventArgs.Empty);
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            base.OnMouseUp(e);
            if (resizeCol >= 0) { Sheet.EditCols[resizeCol] = (float)Math.Round(Sheet.CW[resizeCol]); Sheet.CW[resizeCol] = Sheet.EditCols[resizeCol]; Sheet.Recalc(); RaiseEdited(); }
            if (resizeRow >= 0) { Sheet.EditRows[resizeRow] = (float)Math.Round(Sheet.RH[resizeRow]); Sheet.RH[resizeRow] = Sheet.EditRows[resizeRow]; Sheet.Recalc(); RaiseEdited(); }
            resizeCol = resizeRow = -1;
            dragging = false; Capture = false;
            Invalidate();
        }

        int ColBorderAt(Point p)
        {
            if (p.Y >= HH || p.X < HW) return -1;
            float z = Z; int tol = (int)(4 * dpi);
            for (int c = Render.Find(Sheet.X, Math.Max(0, (p.X - HW + hs.Value) / z) ) + 1; c >= 1 && c <= Sheet.Cols; c--)
            {
                int x = (int)Math.Round(HW - hs.Value + Sheet.X[c] * z);
                if (Math.Abs(p.X - x) <= tol) { int k = c - 1; while (k > 0 && Sheet.CW[k] == 0) k--; return k; }
                if (x < p.X - tol) break;
            }
            return -1;
        }

        int RowBorderAt(Point p)
        {
            if (p.X >= HW || p.Y < HH) return -1;
            float z = Z; int tol = (int)(3 * dpi);
            for (int r = Render.Find(Sheet.Y, Math.Max(0, (p.Y - HH + vs.Value) / z)) + 1; r >= 1 && r <= Sheet.Rows; r--)
            {
                int y = (int)Math.Round(HH - vs.Value + Sheet.Y[r] * z);
                if (Math.Abs(p.Y - y) <= tol) { int k = r - 1; while (k > 0 && Sheet.RH[k] == 0) k--; return k; }
                if (y < p.Y - tol) break;
            }
            return -1;
        }

        // ── 글자 수정 ──
        public void BeginEdit(int r, int c)
        {
            if (Sheet == null) return;
            CommitEdit();
            var m = Sheet.MergeAt(r, c);
            if (m != null) { r = m.R1; c = m.C1; }
            ar = cr = r; ac = cc = c;
            EnsureVisible(r, c);
            int r2 = m != null ? Math.Min(m.R2, Sheet.Rows - 1) : r, c2 = m != null ? Math.Min(m.C2, Sheet.Cols - 1) : c;
            float z = Z;
            int x1 = (int)Math.Round(HW - hs.Value + Sheet.X[c] * z), x2 = (int)Math.Round(HW - hs.Value + Sheet.X[c2 + 1] * z);
            int y1 = (int)Math.Round(HH - vs.Value + Sheet.Y[r] * z), y2 = (int)Math.Round(HH - vs.Value + Sheet.Y[r2 + 1] * z);
            var cell = Sheet.Get(r, c);
            var st = Sheet.StyleOf(cell);
            editor = new TextBox
            {
                Multiline = true, AcceptsTab = false, BorderStyle = BorderStyle.FixedSingle, WordWrap = true,
                Font = Fonts.For(st, z), Text = cell == null ? "" : (cell.Text ?? "").Replace("\n", "\r\n"),
                Bounds = new Rectangle(x1, y1, Math.Max(x2 - x1 + 1, (int)(80 * dpi)), Math.Max(y2 - y1 + 1, (int)(26 * dpi)))
            };
            edR = r; edC = c;
            editor.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Enter && e.Alt) { int p = editor.SelectionStart; editor.Text = editor.Text.Insert(p, "\r\n"); editor.SelectionStart = p + 2; e.SuppressKeyPress = true; }
                else if (e.KeyCode == Keys.Enter && !e.Shift) { CommitEdit(); Focus(); e.SuppressKeyPress = true; }
                else if (e.KeyCode == Keys.Escape) { CancelEdit(); Focus(); e.SuppressKeyPress = true; }
            };
            editor.TextChanged += (s, e) =>
            {
                int need = TextRenderer.MeasureText(editor.Text + "\n", editor.Font, new Size(editor.Width - 6, int.MaxValue), TextFormatFlags.WordBreak | TextFormatFlags.TextBoxControl).Height + 6;
                if (need > editor.Height) editor.Height = Math.Min(need, ClientSize.Height - editor.Top - hs.Height);
            };
            editor.LostFocus += (s, e) => CommitEdit();
            Controls.Add(editor);
            editor.BringToFront();
            editor.Focus();
            editor.SelectionStart = editor.Text.Length;
            Invalidate();
        }

        public void CommitEdit()
        {
            if (editor == null) return;
            var ed = editor; editor = null;
            string text = ed.Text.Replace("\r\n", "\n").Replace("\r", "\n");
            Controls.Remove(ed); ed.Dispose();
            SetCellText(edR, edC, text);
        }

        void CancelEdit()
        {
            if (editor == null) return;
            var ed = editor; editor = null;
            Controls.Remove(ed); ed.Dispose();
        }

        void SetCellText(int r, int c, string text)
        {
            var cell = Sheet.Get(r, c);
            string old = cell == null ? "" : cell.Text ?? "";
            if (old == text) return;
            if (cell == null) { cell = new Cell(); Sheet.Cells[Sheet.Key(r, c)] = cell; }
            cell.Text = text;
            double d;
            cell.IsNum = double.TryParse(text.Replace(",", ""), NumberStyles.Float, CultureInfo.InvariantCulture, out d);
            cell.IsBool = false;
            Sheet.EditCells[Sheet.Key(r, c)] = text;
            Render.Layout(Sheet);
            UpdateBars();
            Invalidate();
            RaiseEdited();
        }

        public void ClearSelectionText()
        {
            if (!HasSelection) return;
            int r1, c1, r2, c2; GetSel(out r1, out c1, out r2, out c2);
            for (int r = r1; r <= r2; r++)
                for (int c = c1; c <= c2; c++)
                {
                    var cell = Sheet.Get(r, c);
                    if (cell != null && !string.IsNullOrEmpty(cell.Text)) { cell.Text = ""; Sheet.EditCells[Sheet.Key(r, c)] = ""; }
                }
            Render.Layout(Sheet); UpdateBars(); Invalidate(); RaiseEdited();
        }

        protected override void OnMouseWheel(MouseEventArgs e)
        {
            base.OnMouseWheel(e);
            if ((ModifierKeys & Keys.Control) != 0) { SetZoom((float)Math.Round(Zoom + (e.Delta > 0 ? 0.1f : -0.1f), 1)); return; }
            int amt = (int)(60 * dpi) * (e.Delta > 0 ? -1 : 1);
            if ((ModifierKeys & Keys.Shift) != 0) SetVal(hs, hs.Value + amt * 2);
            else SetVal(vs, vs.Value + amt);
        }

        protected override bool IsInputKey(Keys k)
        {
            var kc = k & Keys.KeyCode;
            if (kc == Keys.Up || kc == Keys.Down || kc == Keys.Left || kc == Keys.Right || kc == Keys.Enter) return true;
            return base.IsInputKey(k);
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            base.OnKeyDown(e);
            if (Sheet == null) return;
            int dr = 0, dc = 0;
            switch (e.KeyCode)
            {
                case Keys.Up: dr = -1; break;
                case Keys.Down: dr = 1; break;
                case Keys.Left: dc = -1; break;
                case Keys.Right: dc = 1; break;
                case Keys.PageDown: SetVal(vs, vs.Value + ViewH); return;
                case Keys.PageUp: SetVal(vs, vs.Value - ViewH); return;
                case Keys.Home: if (e.Control) { SetVal(vs, 0); SetVal(hs, 0); } return;
                case Keys.F2:
                case Keys.Enter: if (editMode && HasSelection) { BeginEdit(Math.Min(ar, cr), Math.Min(ac, cc)); e.Handled = true; } return;
                case Keys.Delete: if (editMode) ClearSelectionText(); return;
                default: return;
            }
            if (!HasSelection) { SelectCell(0, 0); return; }
            int nr = Math.Max(0, Math.Min(Sheet.Rows - 1, cr + dr)), nc = Math.Max(0, Math.Min(Sheet.Cols - 1, cc + dc));
            if (e.Shift) { cr = nr; cc = nc; EnsureVisible(nr, nc); Invalidate(); if (SelectionChanged != null) SelectionChanged(this, EventArgs.Empty); }
            else SelectCell(nr, nc);
            e.Handled = true;
        }

        // ── 복사 ──
        public string SelectionText()
        {
            int r1, c1, r2, c2; GetSel(out r1, out c1, out r2, out c2);
            var sb = new StringBuilder();
            for (int r = r1; r <= r2; r++)
            {
                if (Sheet.RH[r] == 0) continue;
                bool first = true;
                for (int c = c1; c <= c2; c++)
                {
                    if (Sheet.CW[c] == 0) continue;
                    if (!first) sb.Append('\t');
                    first = false;
                    var m = Sheet.MergeAt(r, c);
                    var cell = (m == null || (m.R1 == r && m.C1 == c)) ? Sheet.Get(r, c) : null;
                    string v = cell == null ? "" : cell.Text ?? "";
                    if (v.Contains("\t") || v.Contains("\n") || v.Contains("\"")) v = "\"" + v.Replace("\"", "\"\"") + "\"";
                    sb.Append(v);
                }
                sb.Append("\r\n");
            }
            return sb.ToString();
        }

        static string Css(Color c) { return string.Format("#{0:x2}{1:x2}{2:x2}", c.R, c.G, c.B); }

        static string BorderCss(XBorder b)
        {
            if (b == null) return "none";
            string s = b.Style, w = "1px", t = "solid";
            if (s.StartsWith("medium")) w = "2px";
            if (s == "thick") w = "3px";
            if (s == "double") { w = "3px"; t = "double"; }
            else if (s == "dotted" || s == "hair") t = "dotted";
            else if (s.Contains("ash")) t = "dashed";
            return w + " " + t + " " + Css(b.Color);
        }

        public string SelectionHtml()
        {
            int r1, c1, r2, c2; GetSel(out r1, out c1, out r2, out c2);
            var s = Sheet;
            var sb = new StringBuilder();
            float tw = 0;
            for (int c = c1; c <= c2; c++) tw += s.CW[c];
            sb.AppendFormat("<table cellspacing=\"0\" cellpadding=\"0\" style=\"border-collapse:collapse;table-layout:fixed;width:{0}px;font-family:'맑은 고딕',sans-serif;\">", (int)tw);
            sb.Append("<colgroup>");
            for (int c = c1; c <= c2; c++) if (s.CW[c] > 0) sb.AppendFormat("<col style=\"width:{0}px\">", (int)s.CW[c]);
            sb.Append("</colgroup>");
            for (int r = r1; r <= r2; r++)
            {
                if (s.RH[r] == 0) continue;
                sb.AppendFormat("<tr style=\"height:{0}px\">", (int)s.RH[r]);
                for (int c = c1; c <= c2; c++)
                {
                    if (s.CW[c] == 0) continue;
                    var m = s.MergeAt(r, c);
                    if (m != null && (m.R1 != r || m.C1 != c)) continue;
                    var cell = s.Get(r, c);
                    var st = s.StyleOf(cell);
                    XBorder bl = st.L, bt = st.T, br = st.R, bb = st.B;
                    string span = "";
                    if (m != null)
                    {
                        int cs = 0, rs = 0;
                        for (int x = m.C1; x <= m.C2 && x < s.Cols; x++) if (s.CW[x] > 0) cs++;
                        for (int y = m.R1; y <= m.R2 && y < s.Rows; y++) if (s.RH[y] > 0) rs++;
                        if (cs > 1) span += " colspan=\"" + cs + "\"";
                        if (rs > 1) span += " rowspan=\"" + rs + "\"";
                        var rc = s.Get(m.R1, Math.Min(m.C2, s.Cols - 1)); if (rc != null) br = s.StyleOf(rc).R;
                        var bc = s.Get(Math.Min(m.R2, s.Rows - 1), m.C1); if (bc != null) bb = s.StyleOf(bc).B;
                    }
                    else
                    {
                        // 이웃 칸이 가진 테두리도 반영
                        if (br == null && c + 1 < s.Cols) { var n = s.Get(r, c + 1); if (n != null) br = s.StyleOf(n).L; }
                        if (bb == null && r + 1 < s.Rows) { var n = s.Get(r + 1, c); if (n != null) bb = s.StyleOf(n).T; }
                    }
                    var css = new StringBuilder();
                    css.AppendFormat("border-left:{0};border-top:{1};border-right:{2};border-bottom:{3};", BorderCss(bl), BorderCss(bt), BorderCss(br), BorderCss(bb));
                    if (st.HasFill) css.Append("background:" + Css(st.Fill) + ";");
                    css.Append("color:" + Css(st.FontColor) + ";");
                    css.AppendFormat(CultureInfo.InvariantCulture, "font-size:{0}pt;", st.FontSize);
                    if (st.Bold) css.Append("font-weight:bold;");
                    if (st.Italic) css.Append("font-style:italic;");
                    if (st.Underline) css.Append("text-decoration:underline;");
                    string h = st.HAlign;
                    if (h == "general") h = cell != null && cell.IsNum ? "right" : cell != null && cell.IsBool ? "center" : "left";
                    if (h == "centerContinuous" || h == "distributed") h = "center";
                    if (h == "fill" || h == "justify") h = "left";
                    css.Append("text-align:" + h + ";");
                    css.Append("vertical-align:" + (st.VAlign == "top" ? "top" : st.VAlign == "bottom" ? "bottom" : "middle") + ";");
                    css.Append(st.Wrap ? "white-space:pre-wrap;word-break:break-all;" : "white-space:nowrap;");
                    css.Append("padding:2px 4px;");
                    string text = cell == null || string.IsNullOrEmpty(cell.Text) ? "&nbsp;" : WebUtility.HtmlEncode(cell.Text).Replace("\n", "<br>");
                    sb.AppendFormat("<td{0} style=\"{1}\">{2}</td>", span, css, text);
                }
                sb.Append("</tr>");
            }
            sb.Append("</table>");
            return sb.ToString();
        }

        public Bitmap SelectionImage()
        {
            int r1, c1, r2, c2; GetSel(out r1, out c1, out r2, out c2);
            float z = Math.Max(1f, dpi);
            int w = (int)Math.Ceiling((Sheet.X[c2 + 1] - Sheet.X[c1]) * z) + 2, h = (int)Math.Ceiling((Sheet.Y[r2 + 1] - Sheet.Y[r1]) * z) + 2;
            w = Math.Max(1, Math.Min(w, 12000)); h = Math.Max(1, Math.Min(h, 12000));
            var bmp = new Bitmap(w, h);
            using (var g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.White);
                Render.Draw(g, Sheet, r1, r2, c1, c2, 1 - Sheet.X[c1] * z, 1 - Sheet.Y[r1] * z, z, false);
            }
            return bmp;
        }
    }

    // ───────────────────────── 수정 내용 저장 ─────────────────────────
    static class Saver
    {
        const string NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        public static void Save(string path, List<Sheet> sheets)
        {
            string ext = Path.GetExtension(path).ToLowerInvariant();
            if (ext == ".csv" || ext == ".tsv") { SaveCsv(path, sheets[0], ext == ".tsv" ? '\t' : ','); return; }
            string tmp = path + ".saving";
            File.Copy(path, tmp, true);
            try
            {
                bool cellsEdited = false;
                using (var fs = new FileStream(tmp, FileMode.Open, FileAccess.ReadWrite))
                using (var zip = new ZipArchive(fs, ZipArchiveMode.Update))
                {
                    foreach (var sh in sheets)
                    {
                        if (!sh.Dirty || sh.EntryName == null) continue;
                        if (sh.EditCells.Count > 0) cellsEdited = true;
                        var entry = zip.GetEntry(sh.EntryName);
                        if (entry == null) continue;
                        var doc = new XmlDocument();
                        doc.PreserveWhitespace = true;
                        using (var st = entry.Open()) doc.Load(st);
                        Apply(doc, sh);
                        entry.Delete();
                        var ne = zip.CreateEntry(sh.EntryName, CompressionLevel.Optimal);
                        using (var st = ne.Open())
                        using (var xw = XmlWriter.Create(st, new XmlWriterSettings { Encoding = new UTF8Encoding(false) })) doc.Save(xw);
                    }
                    if (cellsEdited) RemoveCalcChain(zip);
                }
                File.Copy(tmp, path, true);
            }
            finally { try { File.Delete(tmp); } catch { } }
        }

        static void SaveCsv(string path, Sheet sh, char sep)
        {
            int maxR = -1, maxC = -1;
            foreach (var k in sh.Cells.Keys) { if (!string.IsNullOrEmpty(sh.Cells[k].Text)) { maxR = Math.Max(maxR, (int)(k >> 16)); maxC = Math.Max(maxC, (int)(k & 0xFFFF)); } }
            var sb = new StringBuilder();
            for (int r = 0; r <= maxR; r++)
            {
                for (int c = 0; c <= maxC; c++)
                {
                    if (c > 0) sb.Append(sep);
                    var cell = sh.Get(r, c);
                    string v = cell == null ? "" : cell.Text ?? "";
                    if (v.IndexOf(sep) >= 0 || v.Contains("\n") || v.Contains("\"")) v = "\"" + v.Replace("\"", "\"\"") + "\"";
                    sb.Append(v);
                }
                sb.Append("\r\n");
            }
            File.WriteAllText(path, sb.ToString(), new UTF8Encoding(true));
        }

        static void RemoveCalcChain(ZipArchive zip)
        {
            var cc = zip.GetEntry("xl/calcChain.xml");
            if (cc == null) return;
            cc.Delete();
            foreach (var name in new[] { "[Content_Types].xml", "xl/_rels/workbook.xml.rels" })
            {
                var e = zip.GetEntry(name);
                if (e == null) continue;
                var doc = new XmlDocument();
                using (var st = e.Open()) doc.Load(st);
                var rm = new List<XmlNode>();
                foreach (XmlNode n in doc.DocumentElement.ChildNodes)
                {
                    var el = n as XmlElement;
                    if (el == null) continue;
                    string v = el.GetAttribute("PartName") + el.GetAttribute("Target");
                    if (v.EndsWith("calcChain.xml", StringComparison.OrdinalIgnoreCase)) rm.Add(el);
                }
                foreach (var n in rm) doc.DocumentElement.RemoveChild(n);
                e.Delete();
                var ne = zip.CreateEntry(name, CompressionLevel.Optimal);
                using (var st = ne.Open())
                using (var xw = XmlWriter.Create(st, new XmlWriterSettings { Encoding = new UTF8Encoding(false) })) doc.Save(xw);
            }
        }

        static XmlElement Child(XmlElement p, string name)
        {
            foreach (XmlNode n in p.ChildNodes) { var e = n as XmlElement; if (e != null && e.LocalName == name) return e; }
            return null;
        }

        static string Ref(int r, int c) { return Render.ColName(c) + (r + 1); }

        static int ColOfRef(string r)
        {
            int col = 0, i = 0;
            while (i < r.Length && char.IsLetter(r[i])) { col = col * 26 + (char.ToUpperInvariant(r[i]) - 'A' + 1); i++; }
            return col - 1;
        }

        static string Inv(double d) { return d.ToString("0.##", CultureInfo.InvariantCulture); }

        static void Apply(XmlDocument doc, Sheet sh)
        {
            var root = doc.DocumentElement;
            var sheetData = Child(root, "sheetData");
            if (sheetData == null) return;
            string ns = root.NamespaceURI;

            // 열 너비
            if (sh.EditCols.Count > 0)
            {
                var cols = Child(root, "cols");
                if (cols == null) { cols = doc.CreateElement("cols", ns); root.InsertBefore(cols, sheetData); }
                var list = new List<XmlElement>();
                foreach (XmlNode n in cols.ChildNodes) { var e = n as XmlElement; if (e != null && e.LocalName == "col") list.Add(e); }
                foreach (var kv in sh.EditCols)
                {
                    int k = kv.Key + 1;
                    XmlElement target = null;
                    foreach (var e in list.ToArray())
                    {
                        int mn = int.Parse(e.GetAttribute("min")), mx = int.Parse(e.GetAttribute("max"));
                        if (k < mn || k > mx) continue;
                        if (mn < k) { var a = (XmlElement)e.Clone(); a.SetAttribute("max", (k - 1).ToString()); list.Add(a); }
                        if (mx > k) { var b = (XmlElement)e.Clone(); b.SetAttribute("min", (k + 1).ToString()); list.Add(b); }
                        e.SetAttribute("min", k.ToString()); e.SetAttribute("max", k.ToString());
                        target = e;
                        break;
                    }
                    if (target == null)
                    {
                        target = doc.CreateElement("col", ns);
                        target.SetAttribute("min", k.ToString()); target.SetAttribute("max", k.ToString());
                        list.Add(target);
                    }
                    double w = Math.Round((kv.Value / sh.Mdw) * 256) / 256.0;
                    target.SetAttribute("width", w.ToString("0.########", CultureInfo.InvariantCulture));
                    target.SetAttribute("customWidth", "1");
                    target.RemoveAttribute("bestFit");
                    if (kv.Value > 0) target.RemoveAttribute("hidden");
                }
                list.Sort((a, b) => int.Parse(a.GetAttribute("min")).CompareTo(int.Parse(b.GetAttribute("min"))));
                while (cols.HasChildNodes) cols.RemoveChild(cols.FirstChild);
                foreach (var e in list) cols.AppendChild(e);
            }

            // 행 찾기/만들기
            var rows = new SortedDictionary<int, XmlElement>();
            foreach (XmlNode n in sheetData.ChildNodes)
            {
                var e = n as XmlElement;
                int rn;
                if (e != null && e.LocalName == "row" && int.TryParse(e.GetAttribute("r"), out rn)) rows[rn - 1] = e;
            }
            Func<int, XmlElement> getRow = r =>
            {
                XmlElement e;
                if (rows.TryGetValue(r, out e)) return e;
                e = doc.CreateElement("row", ns);
                e.SetAttribute("r", (r + 1).ToString());
                XmlElement after = null;
                foreach (var kv in rows) { if (kv.Key > r) { after = kv.Value; break; } }
                if (after != null) sheetData.InsertBefore(e, after); else sheetData.AppendChild(e);
                rows[r] = e;
                return e;
            };

            foreach (var kv in sh.EditRows)
            {
                var row = getRow(kv.Key);
                row.SetAttribute("ht", Inv(kv.Value * 72.0 / 96.0));
                row.SetAttribute("customHeight", "1");
                if (kv.Value > 0) row.RemoveAttribute("hidden");
            }

            foreach (var kv in sh.EditCells)
            {
                int r = (int)(kv.Key >> 16), c = (int)(kv.Key & 0xFFFF);
                var row = getRow(r);
                string rf = Ref(r, c);
                XmlElement cel = null, after = null;
                foreach (XmlNode n in row.ChildNodes)
                {
                    var e = n as XmlElement;
                    if (e == null || e.LocalName != "c") continue;
                    string er = e.GetAttribute("r");
                    if (er == rf) { cel = e; break; }
                    if (er.Length > 0 && ColOfRef(er) > c) { after = e; break; }
                }
                if (cel == null)
                {
                    cel = doc.CreateElement("c", ns);
                    cel.SetAttribute("r", rf);
                    if (after != null) row.InsertBefore(cel, after); else row.AppendChild(cel);
                }
                string ot = cel.GetAttribute("t");
                bool wasNum = (ot == "" || ot == "n") && Child(cel, "v") != null;
                foreach (var nm in new[] { "f", "v", "is" }) { var x = Child(cel, nm); while (x != null) { cel.RemoveChild(x); x = Child(cel, nm); } }
                cel.RemoveAttribute("t");
                string text = kv.Value ?? "";
                if (text.Length == 0) continue;
                double d;
                if (wasNum && double.TryParse(text.Replace(",", ""), NumberStyles.Float, CultureInfo.InvariantCulture, out d))
                {
                    var v = doc.CreateElement("v", ns); v.InnerText = d.ToString("R", CultureInfo.InvariantCulture); cel.AppendChild(v);
                }
                else
                {
                    cel.SetAttribute("t", "inlineStr");
                    var isEl = doc.CreateElement("is", ns);
                    var t = doc.CreateElement("t", ns);
                    t.SetAttribute("xml:space", "preserve");
                    t.InnerText = text;
                    isEl.AppendChild(t); cel.AppendChild(isEl);
                }
            }
        }
    }

    // ───────────────────────── 보관함 + 설정 ─────────────────────────
    class Store
    {
        public string Dir, FilesDir, SettingsPath;
        public Dictionary<string, string> Settings = new Dictionary<string, string>();

        public Store()
        {
            string exeDir = Path.GetDirectoryName(Application.ExecutablePath);
            Dir = Path.Combine(exeDir, "ExcelViewer.data");
            if (!TryInit())
            {
                Dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ExcelViewer");
                TryInit();
            }
            LoadSettings();
        }

        bool TryInit()
        {
            try
            {
                FilesDir = Path.Combine(Dir, "files");
                Directory.CreateDirectory(FilesDir);
                try { var di = new DirectoryInfo(Dir); di.Attributes |= FileAttributes.Hidden; } catch { }
                string test = Path.Combine(Dir, ".w");
                File.WriteAllText(test, "1");
                File.Delete(test);
                SettingsPath = Path.Combine(Dir, "settings.ini");
                return true;
            }
            catch { return false; }
        }

        void LoadSettings()
        {
            try
            {
                if (!File.Exists(SettingsPath)) return;
                foreach (var line in File.ReadAllLines(SettingsPath, Encoding.UTF8))
                {
                    int i = line.IndexOf('=');
                    if (i > 0) Settings[line.Substring(0, i)] = line.Substring(i + 1);
                }
            }
            catch { }
        }

        public void SaveSettings()
        {
            try
            {
                var sb = new StringBuilder();
                foreach (var kv in Settings) sb.Append(kv.Key).Append('=').Append(kv.Value).Append("\r\n");
                File.WriteAllText(SettingsPath, sb.ToString(), Encoding.UTF8);
            }
            catch { }
        }

        public string Get(string k, string def) { string v; return Settings.TryGetValue(k, out v) ? v : def; }
        public int GetInt(string k, int def) { int v; return int.TryParse(Get(k, ""), out v) ? v : def; }

        public static bool Supported(string p)
        {
            string e = Path.GetExtension(p).ToLowerInvariant();
            return e == ".xlsx" || e == ".xlsm" || e == ".csv" || e == ".tsv";
        }

        public List<string> Files()
        {
            var list = new List<string>();
            try { foreach (var f in Directory.GetFiles(FilesDir)) if (Supported(f)) list.Add(f); } catch { }
            var order = new List<string>(Get("order", "").Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries));
            list.Sort((a, b) =>
            {
                int ia = order.IndexOf(Path.GetFileName(a)), ib = order.IndexOf(Path.GetFileName(b));
                if (ia < 0) ia = int.MaxValue; if (ib < 0) ib = int.MaxValue;
                if (ia != ib) return ia.CompareTo(ib);
                return File.GetCreationTime(a).CompareTo(File.GetCreationTime(b));
            });
            return list;
        }

        public void SaveOrder(List<string> files)
        {
            var names = new List<string>();
            foreach (var f in files) names.Add(Path.GetFileName(f));
            Settings["order"] = string.Join("|", names);
            SaveSettings();
        }
    }

    // ───────────────────────── 업데이트 ─────────────────────────
    static class Updater
    {
        public static string LatestTag, ExeUrl;

        public static bool Check()
        {
            ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072;
            var req = (HttpWebRequest)WebRequest.Create("https://api.github.com/repos/" + AppInfo.Repo + "/releases/latest");
            req.UserAgent = "ExcelViewer";
            req.Timeout = 10000;
            string json;
            using (var resp = req.GetResponse())
            using (var sr = new StreamReader(resp.GetResponseStream())) json = sr.ReadToEnd();
            var t = Regex.Match(json, "\"tag_name\"\\s*:\\s*\"([^\"]+)\"");
            if (!t.Success) return false;
            LatestTag = t.Groups[1].Value.TrimStart('v', 'V');
            ExeUrl = null;
            foreach (Match m in Regex.Matches(json, "\"browser_download_url\"\\s*:\\s*\"([^\"]+\\.exe)\""))
            { ExeUrl = m.Groups[1].Value; break; }
            return ExeUrl != null && IsNewer(LatestTag, AppInfo.Version);
        }

        public static bool IsNewer(string a, string b)
        {
            Version va, vb;
            return Version.TryParse(a, out va) && Version.TryParse(b, out vb) && va > vb;
        }

        public static void Apply()
        {
            string exe = Application.ExecutablePath;
            string tmp = exe + ".new", old = exe + ".old";
            ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072;
            using (var wc = new WebClient())
            {
                wc.Headers.Add("User-Agent", "ExcelViewer");
                wc.DownloadFile(ExeUrl, tmp);
            }
            if (new FileInfo(tmp).Length < 10000) throw new Exception("내려받은 파일이 올바르지 않아요.");
            if (File.Exists(old)) File.Delete(old);
            File.Move(exe, old);
            File.Move(tmp, exe);
            Process.Start(exe);
        }

        public static void CleanOld()
        {
            string old = Application.ExecutablePath + ".old";
            for (int i = 0; i < 10 && File.Exists(old); i++)
            {
                try { File.Delete(old); } catch { Thread.Sleep(500); }
            }
        }
    }

    // ───────────────────────── 메인 창 ─────────────────────────
    class MainForm : Form
    {
        Store store = new Store();
        SheetView view = new SheetView();
        FlowLayoutPanel fileList = new FlowLayoutPanel();
        FlowLayoutPanel sheetTabs = new FlowLayoutPanel();
        Label home = new Label();
        Label status = new Label();
        Label zoomLbl = new Label();
        TextBox search = new TextBox();
        Button updBtn, editBtn, saveBtn;
        int curSheet;
        Panel sidebar = new Panel();
        Panel center = new Panel();
        List<Sheet> sheets = new List<Sheet>();
        string openPath;
        float dpi = 1f;
        System.Windows.Forms.Timer flashTimer;

        public MainForm(string file)
        {
            dpi = Program.Dpi;
            Text = "엑셀 뷰어";
            Font = new Font("맑은 고딕", 9.5f);
            StartPosition = FormStartPosition.Manual;
            AllowDrop = true;
            KeyPreview = true;
            try { Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch { }

            // 상단 도구줄
            var top = new Panel { Dock = DockStyle.Top, BackColor = Color.FromArgb(248, 249, 250) };
            var flow = new FlowLayoutPanel { Dock = DockStyle.Fill, WrapContents = false, Padding = new Padding(6, 6, 6, 4) };
            flow.Controls.Add(Btn("파일 추가", (s, e) => AddDialog()));
            flow.Controls.Add(Sep());
            flow.Controls.Add(Btn("복사 (Ctrl+C)", (s, e) => CopySel(false)));
            flow.Controls.Add(Btn("그림으로 복사", (s, e) => CopySel(true)));
            flow.Controls.Add(Btn("전체 선택", (s, e) => view.SelectAll()));
            flow.Controls.Add(Sep());
            editBtn = Btn("내용 수정", (s, e) => ToggleEdit());
            flow.Controls.Add(editBtn);
            saveBtn = Btn("저장 (Ctrl+S)", (s, e) => SaveEdits());
            saveBtn.Visible = false; saveBtn.Enabled = false;
            flow.Controls.Add(saveBtn);
            flow.Controls.Add(Sep());
            flow.Controls.Add(new Label { Text = "찾기", AutoSize = true, Margin = new Padding(4, 8, 4, 0) });
            search.Width = (int)(150 * dpi); search.Margin = new Padding(0, 5, 4, 0);
            search.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) { FindNext(); e.SuppressKeyPress = true; } };
            flow.Controls.Add(search);
            flow.Controls.Add(Btn("다음", (s, e) => FindNext()));
            flow.Controls.Add(Sep());
            flow.Controls.Add(Btn("－", (s, e) => view.SetZoom((float)Math.Round(view.Zoom - 0.1f, 1))));
            zoomLbl.AutoSize = true; zoomLbl.Margin = new Padding(2, 8, 2, 0); zoomLbl.Text = "100%";
            flow.Controls.Add(zoomLbl);
            flow.Controls.Add(Btn("＋", (s, e) => view.SetZoom((float)Math.Round(view.Zoom + 0.1f, 1))));
            var right = new FlowLayoutPanel { Dock = DockStyle.Right, AutoSize = true, WrapContents = false, Padding = new Padding(0, 6, 6, 4) };
            updBtn = Btn("v" + AppInfo.Version + " · 업데이트 확인", (s, e) => UpdateClick());
            right.Controls.Add(updBtn);
            top.Controls.Add(flow); top.Controls.Add(right);
            top.Height = flow.PreferredSize.Height + 2;

            // 왼쪽 보관함
            sidebar.Dock = DockStyle.Left;
            sidebar.Width = (int)(store.GetInt("sidebar", 210) * dpi);
            sidebar.BackColor = Color.FromArgb(246, 247, 249);
            var title = new Label { Text = "보관함", Dock = DockStyle.Top, Height = (int)(30 * dpi), Font = new Font("맑은 고딕", 10.5f, FontStyle.Bold), Padding = new Padding(8, 7, 0, 0), ForeColor = Color.FromArgb(40, 40, 40) };
            var drop = new Label
            {
                Dock = DockStyle.Bottom, Height = (int)(84 * dpi), TextAlign = ContentAlignment.MiddleCenter,
                Text = "엑셀 파일을\n여기로 끌어다 놓으면\n보관함에 저장돼요", ForeColor = Color.Gray
            };
            drop.Paint += (s, e) => DashedBox(e.Graphics, ((Control)s).ClientRectangle);
            fileList.Dock = DockStyle.Fill;
            fileList.FlowDirection = FlowDirection.TopDown;
            fileList.WrapContents = false;
            fileList.AutoScroll = true;
            fileList.Padding = new Padding(6, 2, 6, 6);
            fileList.Resize += (s, e) => FitButtons();
            sidebar.Controls.Add(fileList); sidebar.Controls.Add(drop); sidebar.Controls.Add(title);
            var split = new Splitter { Dock = DockStyle.Left, Width = 4, BackColor = Color.FromArgb(230, 232, 235) };

            // 가운데
            center.Dock = DockStyle.Fill;
            view.Dock = DockStyle.Fill;
            view.Visible = false;
            view.SelectionChanged += (s, e) => UpdateStatus();
            view.Edited += (s, e) => { saveBtn.Enabled = true; if (!Text.StartsWith("* ")) Text = "* " + Text; };
            view.ZoomChanged += (s, e) => { zoomLbl.Text = (int)Math.Round(view.Zoom * 100) + "%"; };
            view.SetZoom(store.GetInt("zoom", 100) / 100f);
            zoomLbl.Text = (int)Math.Round(view.Zoom * 100) + "%";
            home.Dock = DockStyle.Fill;
            home.TextAlign = ContentAlignment.MiddleCenter;
            home.Font = new Font("맑은 고딕", 12f);
            home.ForeColor = Color.Gray;
            home.Text = "왼쪽 보관함에서 파일을 누르면 열려요.\n\n새 엑셀 파일은 이 빈 곳으로 끌어다 놓으세요.\n프로그램 보관함에 저장되고 버튼이 생겨요.";
            home.Paint += (s, e) => { var r = ((Control)s).ClientRectangle; r.Inflate(-(int)(30 * dpi), -(int)(30 * dpi)); DashedBox(e.Graphics, r); };
            sheetTabs.Dock = DockStyle.Bottom;
            sheetTabs.Height = (int)(32 * dpi);
            sheetTabs.WrapContents = false;
            sheetTabs.AutoScroll = true;
            sheetTabs.BackColor = Color.FromArgb(243, 244, 246);
            sheetTabs.Padding = new Padding(4, 3, 4, 0);
            sheetTabs.Visible = false;
            center.Controls.Add(view); center.Controls.Add(home); center.Controls.Add(sheetTabs);

            status.Dock = DockStyle.Bottom;
            status.Height = (int)(24 * dpi);
            status.Padding = new Padding(8, 4, 0, 0);
            status.BackColor = Color.FromArgb(248, 249, 250);

            Controls.Add(center);
            Controls.Add(split);
            Controls.Add(sidebar);
            Controls.Add(status);
            Controls.Add(top);

            EnableDrop(this);
            KeyDown += OnKeys;
            FormClosing += (s, e) => { if (!ConfirmDiscard()) { e.Cancel = true; return; } SaveState(); };

            // 창 위치/크기 복원
            var bounds = new Rectangle(store.GetInt("x", -1), store.GetInt("y", -1), store.GetInt("w", (int)(1150 * dpi)), store.GetInt("h", (int)(760 * dpi)));
            bool onScreen = false;
            foreach (var sc in Screen.AllScreens) if (sc.WorkingArea.IntersectsWith(new Rectangle(bounds.X, bounds.Y, 200, 100))) onScreen = true;
            if (bounds.X == -1 || !onScreen)
            {
                var wa = Screen.PrimaryScreen.WorkingArea;
                bounds.Width = Math.Min(bounds.Width, wa.Width); bounds.Height = Math.Min(bounds.Height, wa.Height);
                bounds.X = wa.X + (wa.Width - bounds.Width) / 2; bounds.Y = wa.Y + (wa.Height - bounds.Height) / 2;
            }
            Bounds = bounds;
            if (store.Get("max", "0") == "1") WindowState = FormWindowState.Maximized;

            RefreshFiles();
            UpdateStatus();
            Shown += (s, e) =>
            {
                if (!string.IsNullOrEmpty(file) && File.Exists(file)) OpenFile(file);
                ThreadPool.QueueUserWorkItem(_ =>
                {
                    Updater.CleanOld();
                    try
                    {
                        if (Updater.Check())
                            BeginInvoke((Action)(() =>
                            {
                                updBtn.Text = "새 버전 v" + Updater.LatestTag + " 받기";
                                updBtn.BackColor = Color.FromArgb(255, 236, 179);
                            }));
                    }
                    catch { }
                });
            };
        }

        static void DashedBox(Graphics g, Rectangle r)
        {
            r.Inflate(-6, -6);
            using (var p = new Pen(Color.FromArgb(190, 195, 200), 1.5f) { DashStyle = DashStyle.Dash }) g.DrawRectangle(p, r);
        }

        Control Sep() { return new Label { Width = 1, Height = (int)(22 * dpi), BackColor = Color.FromArgb(215, 218, 222), Margin = new Padding(6, 4, 6, 0) }; }

        Button Btn(string text, EventHandler h)
        {
            var b = new Button { Text = text, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Padding = new Padding(4, 1, 4, 1), Margin = new Padding(0, 2, 4, 0), FlatStyle = FlatStyle.System };
            b.Click += h;
            return b;
        }

        void EnableDrop(Control c)
        {
            c.AllowDrop = true;
            c.DragEnter += (s, e) => { if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy; };
            c.DragDrop += (s, e) => { var f = e.Data.GetData(DataFormats.FileDrop) as string[]; if (f != null) AddFiles(f); };
            foreach (Control ch in c.Controls) EnableDrop(ch);
        }

        void OnKeys(object s, KeyEventArgs e)
        {
            bool typing = search.Focused || view.IsEditing;
            if (e.Control && e.KeyCode == Keys.S) { if (view.EditMode) SaveEdits(); e.SuppressKeyPress = true; }
            else if (typing) return;
            else if (e.Control && e.KeyCode == Keys.C) { CopySel(e.Shift); e.SuppressKeyPress = true; }
            else if (e.Control && e.KeyCode == Keys.A) { view.SelectAll(); e.SuppressKeyPress = true; }
            else if (e.Control && e.KeyCode == Keys.F) { search.Focus(); search.SelectAll(); e.SuppressKeyPress = true; }
            else if (e.Control && e.KeyCode == Keys.O) { AddDialog(); e.SuppressKeyPress = true; }
            else if (e.KeyCode == Keys.F3) { FindNext(); e.SuppressKeyPress = true; }
            else if (e.Control && (e.KeyCode == Keys.Oemplus || e.KeyCode == Keys.Add)) { view.SetZoom((float)Math.Round(view.Zoom + 0.1f, 1)); e.SuppressKeyPress = true; }
            else if (e.Control && (e.KeyCode == Keys.OemMinus || e.KeyCode == Keys.Subtract)) { view.SetZoom((float)Math.Round(view.Zoom - 0.1f, 1)); e.SuppressKeyPress = true; }
            else if (e.Control && e.KeyCode == Keys.D0) { view.SetZoom(1f); e.SuppressKeyPress = true; }
        }

        // ── 보관함 ──
        void RefreshFiles()
        {
            fileList.SuspendLayout();
            foreach (Control c in fileList.Controls) c.Dispose();
            fileList.Controls.Clear();
            var files = store.Files();
            if (files.Count == 0)
                fileList.Controls.Add(new Label { Text = "아직 보관된 파일이 없어요.", AutoSize = true, ForeColor = Color.Gray, Margin = new Padding(4, 6, 0, 0) });
            foreach (var f in files)
            {
                string path = f;
                var b = new Button
                {
                    Text = Path.GetFileNameWithoutExtension(f),
                    TextAlign = ContentAlignment.MiddleLeft,
                    FlatStyle = FlatStyle.Flat,
                    Height = (int)(40 * dpi),
                    Margin = new Padding(0, 0, 0, 4),
                    Padding = new Padding(6, 0, 4, 0),
                    BackColor = Color.White,
                    AutoEllipsis = true,
                    Tag = f,
                    Cursor = Cursors.Hand
                };
                b.FlatAppearance.BorderColor = Color.FromArgb(215, 218, 222);
                b.FlatAppearance.MouseOverBackColor = Color.FromArgb(232, 242, 236);
                new ToolTip().SetToolTip(b, Path.GetFileName(f));
                b.Click += (s, e) => OpenFile(path);
                var menu = new ContextMenuStrip();
                menu.Items.Add("열기", null, (s, e) => OpenFile(path));
                menu.Items.Add("이름 바꾸기", null, (s, e) => RenameFile(path));
                menu.Items.Add("위로 올리기", null, (s, e) => MoveFile(path, -1));
                menu.Items.Add("아래로 내리기", null, (s, e) => MoveFile(path, 1));
                menu.Items.Add("컴퓨터에 꺼내 저장...", null, (s, e) => ExportFile(path));
                menu.Items.Add(new ToolStripSeparator());
                menu.Items.Add("보관함에서 삭제", null, (s, e) => DeleteFile(path));
                b.ContextMenuStrip = menu;
                fileList.Controls.Add(b);
            }
            fileList.ResumeLayout();
            FitButtons();
            HighlightOpen();
            EnableDrop(fileList);
        }

        void FitButtons()
        {
            int w = fileList.ClientSize.Width - fileList.Padding.Horizontal - (fileList.VerticalScroll.Visible ? 0 : 0);
            foreach (Control c in fileList.Controls) if (c is Button) c.Width = Math.Max(40, w);
        }

        void HighlightOpen()
        {
            foreach (Control c in fileList.Controls)
            {
                var b = c as Button;
                if (b == null) continue;
                bool on = openPath != null && string.Equals((string)b.Tag, openPath, StringComparison.OrdinalIgnoreCase);
                b.BackColor = on ? Color.FromArgb(220, 238, 227) : Color.White;
                b.FlatAppearance.BorderColor = on ? AppInfo.Accent : Color.FromArgb(215, 218, 222);
                b.Font = on ? new Font(Font, FontStyle.Bold) : Font;
            }
        }

        void AddDialog()
        {
            using (var d = new OpenFileDialog { Filter = "엑셀/CSV 파일|*.xlsx;*.xlsm;*.csv;*.tsv", Multiselect = true })
                if (d.ShowDialog(this) == DialogResult.OK) AddFiles(d.FileNames);
        }

        void AddFiles(string[] paths)
        {
            string firstSaved = null;
            foreach (var p in paths)
            {
                if (Directory.Exists(p)) continue;
                string ext = Path.GetExtension(p).ToLowerInvariant();
                if (ext == ".xls") { MessageBox.Show(this, "예전 형식(.xls)은 열 수 없어요.\n엑셀에서 '다른 이름으로 저장 → .xlsx'로 바꾼 뒤 넣어주세요.", "엑셀 뷰어"); continue; }
                if (!Store.Supported(p)) { MessageBox.Show(this, Path.GetFileName(p) + "\n엑셀(.xlsx) 또는 CSV 파일만 넣을 수 있어요.", "엑셀 뷰어"); continue; }
                try
                {
                    string dest = Path.Combine(store.FilesDir, Path.GetFileName(p));
                    if (string.Equals(Path.GetFullPath(p), Path.GetFullPath(dest), StringComparison.OrdinalIgnoreCase)) { firstSaved = firstSaved ?? dest; continue; }
                    if (File.Exists(dest))
                    {
                        var ans = MessageBox.Show(this, "'" + Path.GetFileName(p) + "' 파일이 이미 보관함에 있어요.\n\n[예] 새 파일로 바꾸기\n[아니요] 다른 이름으로 하나 더 보관", "엑셀 뷰어", MessageBoxButtons.YesNoCancel);
                        if (ans == DialogResult.Cancel) continue;
                        if (ans == DialogResult.No)
                        {
                            string bn = Path.GetFileNameWithoutExtension(p), ex = Path.GetExtension(p);
                            int i = 2;
                            do dest = Path.Combine(store.FilesDir, bn + " (" + i++ + ")" + ex); while (File.Exists(dest));
                        }
                        else if (dest == openPath) CloseFile();
                    }
                    using (var src = new FileStream(p, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    using (var dst = new FileStream(dest, FileMode.Create, FileAccess.Write)) src.CopyTo(dst);
                    try { File.SetCreationTime(dest, DateTime.Now); } catch { }
                    if (firstSaved == null) firstSaved = dest;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "보관하지 못했어요.\n" + ex.Message + "\n\n(엑셀에서 파일이 열려 있으면 닫고 다시 해보세요.)", "엑셀 뷰어");
                }
            }
            RefreshFiles();
            if (firstSaved != null) { OpenFile(firstSaved); Flash("보관함에 저장했어요."); }
        }

        void RenameFile(string path)
        {
            string name = Prompt("새 이름", Path.GetFileNameWithoutExtension(path));
            if (string.IsNullOrWhiteSpace(name)) return;
            foreach (var ch in Path.GetInvalidFileNameChars()) name = name.Replace(ch.ToString(), "");
            string dest = Path.Combine(store.FilesDir, name + Path.GetExtension(path));
            if (File.Exists(dest)) { MessageBox.Show(this, "같은 이름의 파일이 이미 있어요.", "엑셀 뷰어"); return; }
            try
            {
                var files = store.Files();
                int idx = files.FindIndex(f => string.Equals(f, path, StringComparison.OrdinalIgnoreCase));
                File.Move(path, dest);
                if (idx >= 0) { files[idx] = dest; store.SaveOrder(files); }
                if (openPath == path) { openPath = dest; Text = Path.GetFileName(dest) + " - 엑셀 뷰어"; }
                RefreshFiles();
            }
            catch (Exception ex) { MessageBox.Show(this, ex.Message, "엑셀 뷰어"); }
        }

        void MoveFile(string path, int dir)
        {
            var files = store.Files();
            int i = files.FindIndex(f => string.Equals(f, path, StringComparison.OrdinalIgnoreCase));
            int j = i + dir;
            if (i < 0 || j < 0 || j >= files.Count) return;
            var t = files[i]; files[i] = files[j]; files[j] = t;
            store.SaveOrder(files);
            RefreshFiles();
        }

        void ExportFile(string path)
        {
            using (var d = new SaveFileDialog { FileName = Path.GetFileName(path), Filter = "엑셀/CSV 파일|*" + Path.GetExtension(path) })
                if (d.ShowDialog(this) == DialogResult.OK)
                    try { File.Copy(path, d.FileName, true); Flash("꺼내 저장했어요."); }
                    catch (Exception ex) { MessageBox.Show(this, ex.Message, "엑셀 뷰어"); }
        }

        void DeleteFile(string path)
        {
            if (MessageBox.Show(this, "'" + Path.GetFileName(path) + "'을(를) 보관함에서 지울까요?\n(컴퓨터에 있던 원래 파일은 그대로예요.)", "엑셀 뷰어", MessageBoxButtons.YesNo) != DialogResult.Yes) return;
            if (openPath == path) CloseFile();
            try { File.Delete(path); } catch (Exception ex) { MessageBox.Show(this, ex.Message, "엑셀 뷰어"); }
            RefreshFiles();
        }

        string Prompt(string title, string value)
        {
            using (var f = new Form { Text = title, FormBorderStyle = FormBorderStyle.FixedDialog, StartPosition = FormStartPosition.CenterParent, MinimizeBox = false, MaximizeBox = false, ClientSize = new Size((int)(340 * dpi), (int)(90 * dpi)), Font = Font })
            {
                var tb = new TextBox { Text = value, Left = (int)(12 * dpi), Top = (int)(14 * dpi), Width = (int)(316 * dpi) };
                var ok = new Button { Text = "확인", DialogResult = DialogResult.OK, Left = (int)(168 * dpi), Top = (int)(50 * dpi), Width = (int)(76 * dpi), Height = (int)(28 * dpi) };
                var cn = new Button { Text = "취소", DialogResult = DialogResult.Cancel, Left = (int)(252 * dpi), Top = (int)(50 * dpi), Width = (int)(76 * dpi), Height = (int)(28 * dpi) };
                f.Controls.AddRange(new Control[] { tb, ok, cn });
                f.AcceptButton = ok; f.CancelButton = cn;
                return f.ShowDialog(this) == DialogResult.OK ? tb.Text.Trim() : null;
            }
        }

        // ── 파일 보기 ──
        void CloseFile()
        {
            foreach (var s in sheets) { s.EditCells.Clear(); s.EditCols.Clear(); s.EditRows.Clear(); }
            if (view.EditMode) ToggleEdit();
            openPath = null;
            sheets = new List<Sheet>();
            view.SetSheet(null);
            view.Visible = false; sheetTabs.Visible = false; home.Visible = true;
            Text = "엑셀 뷰어";
            HighlightOpen();
            UpdateStatus();
        }

        bool AnyDirty()
        {
            view.CommitEdit();
            foreach (var s in sheets) if (s.Dirty) return true;
            return false;
        }

        bool ConfirmDiscard()
        {
            if (!AnyDirty()) return true;
            var a = MessageBox.Show(this, "수정한 내용을 저장할까요?\n\n[예] 저장   [아니요] 저장 안 함", "엑셀 뷰어", MessageBoxButtons.YesNoCancel);
            if (a == DialogResult.Cancel) return false;
            if (a == DialogResult.Yes) return SaveEdits();
            foreach (var s in sheets) { s.EditCells.Clear(); s.EditCols.Clear(); s.EditRows.Clear(); }
            return true;
        }

        void ToggleEdit()
        {
            if (view.Sheet == null) { Flash("먼저 보관함에서 파일을 열어주세요."); return; }
            if (!view.EditMode)
            {
                view.EditMode = true;
                editBtn.Text = "수정 끝내기";
                editBtn.BackColor = Color.FromArgb(255, 236, 179);
                saveBtn.Visible = true;
                UpdateStatus();
                return;
            }
            bool dirty = AnyDirty();
            if (dirty)
            {
                if (!ConfirmDiscard()) return;
                if (openPath != null && sheets.Count > 0 && sheets[0].Dirty == false) { }
            }
            view.EditMode = false;
            editBtn.Text = "내용 수정";
            editBtn.BackColor = SystemColors.Control;
            editBtn.UseVisualStyleBackColor = true;
            saveBtn.Visible = false;
            if (dirty && openPath != null) ReloadKeep();
            UpdateStatus();
        }

        void ReloadKeep()
        {
            int idx = curSheet;
            OpenFile(openPath, true);
            ShowSheet(Math.Min(idx, sheets.Count - 1));
        }

        bool SaveEdits()
        {
            if (openPath == null) return true;
            if (!AnyDirty()) { Flash("바뀐 내용이 없어요."); return true; }
            try
            {
                Cursor = Cursors.WaitCursor;
                Saver.Save(openPath, sheets);
                foreach (var s in sheets) { s.EditCells.Clear(); s.EditCols.Clear(); s.EditRows.Clear(); }
                ReloadKeep();
                saveBtn.Enabled = false;
                Flash("저장했어요.");
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "저장하지 못했어요.\n" + ex.Message + "\n\n(엑셀에서 이 파일이 열려 있으면 닫고 다시 해보세요.)", "엑셀 뷰어");
                return false;
            }
            finally { Cursor = Cursors.Default; }
        }

        void OpenFile(string path) { OpenFile(path, false); }

        void OpenFile(string path, bool reload)
        {
            if (!reload && !ConfirmDiscard()) return;
            Cursor = Cursors.WaitCursor;
            try
            {
                string ext = Path.GetExtension(path).ToLowerInvariant();
                var list = (ext == ".csv" || ext == ".tsv") ? CsvReader.Read(path) : XlsxReader.Read(path);
                if (list.Count == 0) throw new Exception("보이는 시트가 없어요.");
                foreach (var s in list) Render.Layout(s);
                sheets = list;
                openPath = path;
                Text = Path.GetFileName(path) + " - 엑셀 뷰어";
                BuildTabs();
                home.Visible = false; view.Visible = true; sheetTabs.Visible = true;
                saveBtn.Enabled = false;
                ShowSheet(0);
                HighlightOpen();
                view.Focus();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "파일을 열 수 없어요.\n" + ex.Message, "엑셀 뷰어");
            }
            finally { Cursor = Cursors.Default; }
        }

        void BuildTabs()
        {
            sheetTabs.SuspendLayout();
            foreach (Control c in sheetTabs.Controls) c.Dispose();
            sheetTabs.Controls.Clear();
            for (int i = 0; i < sheets.Count; i++)
            {
                int idx = i;
                var b = new Button
                {
                    Text = sheets[i].Name, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, FlatStyle = FlatStyle.Flat,
                    Padding = new Padding(8, 0, 8, 0), Margin = new Padding(0, 0, 2, 0), Height = (int)(26 * dpi), BackColor = Color.White
                };
                b.FlatAppearance.BorderColor = Color.FromArgb(210, 212, 216);
                b.Click += (s, e) => ShowSheet(idx);
                sheetTabs.Controls.Add(b);
            }
            sheetTabs.ResumeLayout();
        }

        void ShowSheet(int i)
        {
            if (i < 0 || i >= sheets.Count) return;
            view.CommitEdit();
            curSheet = i;
            view.SetSheet(sheets[i]);
            for (int k = 0; k < sheetTabs.Controls.Count; k++)
            {
                var b = (Button)sheetTabs.Controls[k];
                bool on = k == i;
                b.BackColor = on ? Color.White : Color.FromArgb(243, 244, 246);
                b.ForeColor = on ? AppInfo.Accent : Color.FromArgb(70, 70, 70);
                b.Font = on ? new Font(Font, FontStyle.Bold) : Font;
                b.FlatAppearance.BorderColor = on ? AppInfo.Accent : Color.FromArgb(210, 212, 216);
            }
            UpdateStatus();
        }

        void UpdateStatus()
        {
            if (view.Sheet == null) { status.Text = "파일을 끌어다 놓으면 보관함에 저장돼요.  보관함 버튼을 오른쪽 클릭하면 이름 바꾸기·삭제를 할 수 있어요."; return; }
            if (view.EditMode) { status.Text = "수정 모드 · 칸 더블클릭(또는 Enter)으로 글자 수정, Alt+Enter 줄바꿈 · 위/왼쪽 머리글 경계선을 끌어 칸 크기 조절 · Delete 지우기 · Ctrl+S 저장"; return; }
            if (!view.HasSelection) { status.Text = "마우스로 칸을 드래그해서 고른 뒤 Ctrl+C → 다른 곳에 붙여넣기.  (Ctrl+A 전체 선택, Ctrl+휠 확대)"; return; }
            int r1, c1, r2, c2; view.GetSel(out r1, out c1, out r2, out c2);
            string a = Render.ColName(c1) + (r1 + 1), b = Render.ColName(c2) + (r2 + 1);
            status.Text = (a == b ? a : a + " : " + b) + "   ·   Ctrl+C 표로 복사   ·   Ctrl+Shift+C 그림으로 복사";
        }

        void CopySel(bool image)
        {
            if (view.Sheet == null) return;
            if (!view.HasSelection) { Flash("먼저 복사할 칸을 마우스로 드래그해서 골라주세요."); return; }
            try
            {
                if (image)
                {
                    using (var bmp = view.SelectionImage()) Clipboard.SetImage(bmp);
                    Flash("그림으로 복사했어요. 붙여넣으면 이미지로 들어가요.");
                }
                else
                {
                    var obj = new DataObject();
                    obj.SetText(view.SelectionText(), TextDataFormat.UnicodeText);
                    obj.SetData(DataFormats.Html, HtmlClip(view.SelectionHtml()));
                    Clipboard.SetDataObject(obj, true);
                    Flash("표 모양 그대로 복사했어요.");
                }
            }
            catch (Exception ex) { Flash("복사 실패: " + ex.Message); }
        }

        static string HtmlClip(string fragment)
        {
            const string header = "Version:0.9\r\nStartHTML:{0:D10}\r\nEndHTML:{1:D10}\r\nStartFragment:{2:D10}\r\nEndFragment:{3:D10}\r\n";
            string pre = "<html><head><meta charset=\"utf-8\"></head><body><!--StartFragment-->";
            string post = "<!--EndFragment--></body></html>";
            int headerLen = string.Format(header, 0, 0, 0, 0).Length;
            var enc = Encoding.UTF8;
            int startHtml = headerLen;
            int startFrag = startHtml + enc.GetByteCount(pre);
            int endFrag = startFrag + enc.GetByteCount(fragment);
            int endHtml = endFrag + enc.GetByteCount(post);
            return string.Format(header, startHtml, endHtml, startFrag, endFrag) + pre + fragment + post;
        }

        void FindNext()
        {
            var s = view.Sheet;
            string q = search.Text.Trim();
            if (s == null || q.Length == 0 || s.TextOrder.Count == 0) return;
            long cur = -1;
            if (view.HasSelection) { int r1, c1, r2, c2; view.GetSel(out r1, out c1, out r2, out c2); cur = Sheet.Key(r1, c1); }
            int start = 0;
            while (start < s.TextOrder.Count && s.TextOrder[start] <= cur) start++;
            for (int k = 0; k < s.TextOrder.Count; k++)
            {
                long key = s.TextOrder[(start + k) % s.TextOrder.Count];
                var cell = s.Cells[key];
                if (cell.Text.IndexOf(q, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    view.SelectCell((int)(key >> 16), (int)(key & 0xFFFF));
                    return;
                }
            }
            Flash("'" + q + "'을(를) 찾지 못했어요.");
        }

        void Flash(string msg)
        {
            status.Text = msg;
            if (flashTimer == null) { flashTimer = new System.Windows.Forms.Timer { Interval = 3000 }; flashTimer.Tick += (s, e) => { flashTimer.Stop(); UpdateStatus(); }; }
            flashTimer.Stop(); flashTimer.Start();
        }

        void SaveState()
        {
            var b = WindowState == FormWindowState.Normal ? Bounds : RestoreBounds;
            store.Settings["x"] = b.X.ToString(); store.Settings["y"] = b.Y.ToString();
            store.Settings["w"] = b.Width.ToString(); store.Settings["h"] = b.Height.ToString();
            store.Settings["max"] = WindowState == FormWindowState.Maximized ? "1" : "0";
            store.Settings["zoom"] = ((int)Math.Round(view.Zoom * 100)).ToString();
            store.Settings["sidebar"] = ((int)(sidebar.Width / dpi)).ToString();
            store.SaveSettings();
        }

        void UpdateClick()
        {
            Cursor = Cursors.WaitCursor;
            bool newer;
            try { newer = Updater.Check(); }
            catch (WebException ex)
            {
                Cursor = Cursors.Default;
                var r = ex.Response as HttpWebResponse;
                if (r != null && r.StatusCode == HttpStatusCode.NotFound) MessageBox.Show(this, "아직 올라온 업데이트가 없어요. (현재 v" + AppInfo.Version + ")", "엑셀 뷰어");
                else MessageBox.Show(this, "업데이트 확인을 못 했어요. 인터넷 연결을 확인해 주세요.\n" + ex.Message, "엑셀 뷰어");
                return;
            }
            catch (Exception ex) { Cursor = Cursors.Default; MessageBox.Show(this, "업데이트 확인 실패: " + ex.Message, "엑셀 뷰어"); return; }
            Cursor = Cursors.Default;
            if (!newer) { MessageBox.Show(this, "최신 버전이에요. (v" + AppInfo.Version + ")", "엑셀 뷰어"); return; }
            if (MessageBox.Show(this, "새 버전 v" + Updater.LatestTag + "이 있어요.\n바로 업데이트하고 다시 켤까요?\n(보관함 파일과 설정은 그대로예요)", "엑셀 뷰어", MessageBoxButtons.YesNo) != DialogResult.Yes) return;
            try
            {
                Cursor = Cursors.WaitCursor;
                SaveState();
                Updater.Apply();
                Application.Exit();
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                MessageBox.Show(this, "업데이트 실패: " + ex.Message, "엑셀 뷰어");
            }
        }
    }

    static class Program
    {
        [DllImport("user32.dll")] static extern bool SetProcessDPIAware();
        static float dpi = -1;
        public static float Dpi
        {
            get
            {
                if (dpi < 0)
                {
                    dpi = 1f;
                    try { using (var g = Graphics.FromHwnd(IntPtr.Zero)) dpi = Math.Max(1f, g.DpiX / 96f); } catch { }
                }
                return dpi;
            }
        }

        [STAThread]
        static void Main(string[] args)
        {
            try { if (Environment.OSVersion.Version.Major >= 6) SetProcessDPIAware(); } catch { }
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm(args.Length > 0 ? args[0] : null));
        }
    }
}
