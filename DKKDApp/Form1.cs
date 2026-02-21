
using Xceed.Words.NET;
using Xceed.Document.NET;
using System.Text.RegularExpressions;
using System.Text;

namespace DKKDApp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();


            GetProcess();
        }
        private int GetProcess()
        {
            try
            {
                string baseDir = AppContext.BaseDirectory;

                string notesPath = Path.Combine(baseDir, "Notes_Placeholder.txt");
                if (!File.Exists(notesPath))
                {
                    Console.Error.WriteLine("Không tìm thấy Notes_Placeholder.txt tại: " + notesPath);
                    return 1;
                }

                // 1) Parse Notes: key-value + block ngành nghề
                var parsed = ParseNotesWithBusinessBlock(notesPath);
                var map = parsed.KeyValues;
                var businessLines = parsed.BusinessLines;

                Console.WriteLine($"Đã đọc Notes: {map.Count} key.");
                Console.WriteLine($"Đã đọc ngành nghề: {businessLines.Count} dòng.");

                // 2) Danh sách 4 template nằm ở thư mục gốc
                var templates = new[]
                {
                    "GIẤY ĐỀ NGHỊ_Mẫu số 2.docx",
                    "Điều lệ công ty.docx",
                    "Danh Sách Chủ Sở Hữu Hưởng Lợi_Mẫu số 10.docx",
                    "Ủy Quyền.docx"
                };

                int ok = 0;
                foreach (var name in templates)
                {
                    string inputPath = Path.Combine(baseDir, name);
                    if (!File.Exists(inputPath))
                    {
                        Console.Error.WriteLine("Không tìm thấy template: " + inputPath);
                        continue;
                    }

                    string outPath = Path.Combine(baseDir, Path.GetFileNameWithoutExtension(name) + "_DA_DIEN.docx");
                    FillOneDocx(inputPath, outPath, map, businessLines);

                    Console.WriteLine("OK: " + outPath);
                    ok++;
                }

                Console.WriteLine($"Hoàn tất. Thành công: {ok}/{templates.Length}");
                return ok == templates.Length ? 0 : 2;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("Lỗi: " + ex);
                return 99;
            }
        }
        static void FillOneDocx(string inputDocx, string outputDocx,
            Dictionary<string, string> map,
            List<BusinessLine> businessLines)
        {
            using var doc = DocX.Load(inputDocx);

            // (A) Replace các placeholder bình thường {KEY} (trừ DANH_SACH_NGANH_NGHE)
            ReplaceEverywhereExceptBusinessPlaceholder(doc, map);

            // (B) Chèn bảng ngành nghề tại {DANH_SACH_NGANH_NGHE} (nếu file có placeholder này)
            InsertBusinessLineTableAtPlaceholder(doc, businessLines, "{DANH_SACH_NGANH_NGHE}");

            doc.SaveAs(outputDocx);
        }

        private class NotesParsed
        {
            public Dictionary<string, string> KeyValues { get; } = new(StringComparer.OrdinalIgnoreCase);
            public List<BusinessLine> BusinessLines { get; } = new();
        }

        static NotesParsed ParseNotesWithBusinessBlock(string notesPath)
        {
            var parsed = new NotesParsed();

            bool inBusiness = false;
            var rawBusinessLines = new List<string>();

            foreach (var raw in File.ReadAllLines(notesPath, Encoding.UTF8))
            {
                var line = raw.Trim();
                if (string.IsNullOrWhiteSpace(line)) continue;

                // Begin / End markers
                if (line.StartsWith("DANH_SACH_NGANH_NGHE_BEGIN", StringComparison.OrdinalIgnoreCase))
                {
                    inBusiness = true;
                    continue;
                }
                if (line.StartsWith("DANH_SACH_NGANH_NGHE_END", StringComparison.OrdinalIgnoreCase))
                {
                    inBusiness = false;
                    continue;
                }

                if (inBusiness)
                {
                    rawBusinessLines.Add(line);
                    continue;
                }

                // Parse normal KEY={VALUE}
                var m = Regex.Match(line, @"^([A-Za-z0-9_]+)\s*=\s*\{(.*)\}\s*$");
                if (!m.Success) continue;

                string key = m.Groups[1].Value.Trim();
                string value = m.Groups[2].Value; // inside {...}, may be empty

                parsed.KeyValues[key] = value;
            }

            parsed.BusinessLines.AddRange(ParseBusinessLines(rawBusinessLines));
            return parsed;
        }

        // Input line format: "4933;Tên ngành;Ngành chính" or "5229;Tên ngành;"
        static List<BusinessLine> ParseBusinessLines(IEnumerable<string> rawLines)
        {
            var list = new List<BusinessLine>();
            int stt = 1;

            foreach (var raw in rawLines)
            {
                var line = (raw ?? "").Trim();
                if (string.IsNullOrWhiteSpace(line)) continue;

                // Split by ';'
                var parts = line.Split(';').Select(p => p.Trim()).ToList();
                if (parts.Count == 0) continue;

                string code = parts.Count >= 1 ? parts[0] : "";
                string name = parts.Count >= 2 ? parts[1] : "";

                // main flag if any part contains "ngành chính"
                bool isMain = parts.Any(p => p.IndexOf("ngành chính", StringComparison.OrdinalIgnoreCase) >= 0)
                              || line.IndexOf("ngành chính", StringComparison.OrdinalIgnoreCase) >= 0;

                if (string.IsNullOrWhiteSpace(code) && string.IsNullOrWhiteSpace(name))
                    continue;

                list.Add(new BusinessLine
                {
                    Stt = stt++,
                    Code = code,
                    Name = name,
                    IsMain = isMain
                });
            }

            return list;
        }

        // =========================
        // Replace placeholders (except business placeholder)
        // =========================
        static void ReplaceEverywhereExceptBusinessPlaceholder(DocX doc, Dictionary<string, string> map)
        {
            // Body
            ReplaceInParagraphsExcept(doc.Paragraphs, map, "DANH_SACH_NGANH_NGHE");
            foreach (var t in doc.Tables) ReplaceInTableExcept(t, map, "DANH_SACH_NGANH_NGHE");

            // Headers/Footers
            foreach (var s in doc.Sections)
            {
                if (s.Headers != null)
                {
                    if (s.Headers.First != null) ReplaceInHeaderFooterExcept(s.Headers.First, map, "DANH_SACH_NGANH_NGHE");
                    if (s.Headers.Odd != null) ReplaceInHeaderFooterExcept(s.Headers.Odd, map, "DANH_SACH_NGANH_NGHE");
                    if (s.Headers.Even != null) ReplaceInHeaderFooterExcept(s.Headers.Even, map, "DANH_SACH_NGANH_NGHE");
                }
                if (s.Footers != null)
                {
                    if (s.Footers.First != null) ReplaceInHeaderFooterExcept(s.Footers.First, map, "DANH_SACH_NGANH_NGHE");
                    if (s.Footers.Odd != null) ReplaceInHeaderFooterExcept(s.Footers.Odd, map, "DANH_SACH_NGANH_NGHE");
                    if (s.Footers.Even != null) ReplaceInHeaderFooterExcept(s.Footers.Even, map, "DANH_SACH_NGANH_NGHE");
                }
            }
        }

        static void ReplaceInHeaderFooterExcept(dynamic hf, Dictionary<string, string> map, string exceptKey)
        {
            ReplaceInParagraphsExcept(hf.Paragraphs, map, exceptKey);
            foreach (var t in hf.Tables) ReplaceInTableExcept(t, map, exceptKey);
        }

        static void ReplaceInTableExcept(Table table, Dictionary<string, string> map, string exceptKey)
        {
            foreach (var row in table.Rows)
                foreach (var cell in row.Cells)
                {
                    ReplaceInParagraphsExcept(cell.Paragraphs, map, exceptKey);

                    foreach (var nested in cell.Tables)
                        ReplaceInTableExcept(nested, map, exceptKey);
                }
        }

        static void ReplaceInParagraphsExcept(IEnumerable<Paragraph> paragraphs, Dictionary<string, string> map, string exceptKey)
        {
            foreach (var p in paragraphs)
            {
                foreach (var kv in map)
                {
                    if (kv.Key.Equals(exceptKey, StringComparison.OrdinalIgnoreCase))
                        continue;

                    string token = "{" + kv.Key + "}";
                    string value = kv.Value ?? "";
                    p.ReplaceText(token, value, false, RegexOptions.None);
                }
            }
        }

        // =========================
        // Insert business table at placeholder
        // =========================
        static void InsertBusinessLineTableAtPlaceholder(DocX doc, List<BusinessLine> lines, string placeholderToken)
        {
            // Tìm paragraph trong body
            var p = doc.Paragraphs.FirstOrDefault(x => (x.Text ?? "").Contains(placeholderToken));
            if (p != null)
            {
                InsertTableAfterParagraph(doc, p, lines, placeholderToken);
                return;
            }

            // Nếu placeholder nằm trong bảng/cell, tìm trong tables
            foreach (var t in doc.Tables)
            {
                if (TryInsertInTable(doc, t, lines, placeholderToken))
                    return;
            }

            // Nếu cần, bạn có thể mở rộng thêm: header/footer
        }

        static bool TryInsertInTable(DocX doc, Table table, List<BusinessLine> lines, string placeholderToken)
        {
            foreach (var row in table.Rows)
                foreach (var cell in row.Cells)
                {
                    var p = cell.Paragraphs.FirstOrDefault(x => (x.Text ?? "").Contains(placeholderToken));
                    if (p != null)
                    {
                        InsertTableAfterParagraph(doc, p, lines, placeholderToken);
                        return true;
                    }

                    foreach (var nested in cell.Tables)
                    {
                        if (TryInsertInTable(doc, nested, lines, placeholderToken))
                            return true;
                    }
                }
            return false;
        }

        static void InsertTableAfterParagraph(DocX doc, Paragraph p, List<BusinessLine> lines, string placeholderToken)
        {
            try
            {
                // Xóa token trong paragraph
                p.ReplaceText(placeholderToken, "", false, RegexOptions.None);

                // Tạo bảng: header + dữ liệu
                int rows = 1 + lines.Count;
                int cols = 4;

                var table = doc.AddTable(rows, cols);
                table.Design = TableDesign.TableGrid;

                // Header
                table.Rows[0].Cells[0].Paragraphs[0].Append("STT").Bold();
                table.Rows[0].Cells[1].Paragraphs[0].Append("Mã ngành").Bold();
                table.Rows[0].Cells[2].Paragraphs[0].Append("Tên Ngành").Bold();
                table.Rows[0].Cells[3].Paragraphs[0].Append("Ngành Chính").Bold();

                // Rows
                for (int i = 0; i < lines.Count; i++)
                {
                    int r = i + 1;
                    var item = lines[i];

                    table.Rows[r].Cells[0].Paragraphs[0].Append(item.Stt.ToString());
                    table.Rows[r].Cells[1].Paragraphs[0].Append(item.Code);
                    table.Rows[r].Cells[2].Paragraphs[0].Append(item.Name);
                    table.Rows[r].Cells[3].Paragraphs[0].Append(item.IsMain ? "X" : "");
                }

                // Chèn bảng ngay sau paragraph chứa placeholder
                p.InsertTableAfterSelf(table);
            }
            catch (Exception ex) { 
            
            }
        }
    }

}