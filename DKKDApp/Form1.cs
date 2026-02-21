
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

                // 2) Danh sách 4 template cùng thư mục gốc
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

            // (A) Điền bảng ngành nghề theo dòng mẫu {BL_*} (nếu file có bảng này)
            FillBusinessTableByTemplateRow(doc, businessLines);

            // (B) Replace các placeholder bình thường {KEY}
            //     KHÔNG replace các token {BL_STT},{BL_CODE},{BL_NAME},{BL_MAIN} nữa (vì đã xử lý ở bước A)
            ReplaceEverywhereExcept(doc, map, new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "BL_STT","BL_CODE","BL_NAME","BL_MAIN"
            });

            doc.SaveAs(outputDocx);
        }

        // =========================================================
        // 1) Parse Notes (KEY={VALUE} + block ngành nghề BEGIN/END)
        // =========================================================
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

                // KEY={VALUE}  (non-greedy để an toàn)
                var m = Regex.Match(line, @"^([A-Za-z0-9_]+)\s*=\s*\{(.*?)\}\s*$");
                if (!m.Success) continue;

                string key = m.Groups[1].Value.Trim();
                string value = m.Groups[2].Value;

                parsed.KeyValues[key] = value;
            }

            parsed.BusinessLines.AddRange(ParseBusinessLines(rawBusinessLines));
            return parsed;
        }

        // Mỗi dòng: "4933;Tên ngành;Ngành chính" hoặc "5229;Tên ngành;"
        static List<BusinessLine> ParseBusinessLines(IEnumerable<string> rawLines)
        {
            var list = new List<BusinessLine>();
            int stt = 1;

            foreach (var raw in rawLines)
            {
                var line = (raw ?? "").Trim();
                if (string.IsNullOrWhiteSpace(line)) continue;

                var parts = line.Split(';').Select(p => p.Trim()).ToList();
                if (parts.Count == 0) continue;

                string code = parts.Count >= 1 ? parts[0] : "";
                string name = parts.Count >= 2 ? parts[1] : "";

                // Ngành chính: nếu có "ngành chính" ở bất kỳ phần nào sau dấu ';'
                bool isMain = parts.Skip(2).Any(p => p.IndexOf("ngành chính", StringComparison.OrdinalIgnoreCase) >= 0)
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

        // =========================================================
        // 2) Fill bảng ngành nghề theo "dòng mẫu" {BL_*}
        // =========================================================
        static void FillBusinessTableByTemplateRow(DocX doc, List<BusinessLine> lines)
        {
            if (lines == null || lines.Count == 0) return;

            // Tìm table có row chứa {BL_CODE} (điểm nhận dạng chắc nhất)
            var table = doc.Tables.FirstOrDefault(t =>
                t.Rows.Any(r => RowContainsToken(r, "{BL_CODE}")));

            if (table == null) return;

            // Tìm index dòng mẫu
            int templateRowIndex = -1;
            for (int i = 0; i < table.RowCount; i++)
            {
                if (RowContainsToken(table.Rows[i], "{BL_CODE}"))
                {
                    templateRowIndex = i;
                    break;
                }
            }
            if (templateRowIndex < 0) return;

            var templateRow = table.Rows[templateRowIndex];

            // Xoá các dòng dữ liệu cũ (tuỳ bạn). Ở đây: giữ header (row 0) + template row, xoá các row khác.
            for (int i = table.RowCount - 1; i >= 0; i--)
            {
                if (i != 0 && i != templateRowIndex)
                    table.RemoveRow(i);
            }

            // Sau khi remove, templateRowIndex có thể thay đổi nếu templateRow không ở cuối.
            // Ta tìm lại lần nữa:
            templateRowIndex = -1;
            for (int i = 0; i < table.RowCount; i++)
            {
                if (RowContainsToken(table.Rows[i], "{BL_CODE}"))
                {
                    templateRowIndex = i;
                    break;
                }
            }
            if (templateRowIndex < 0) return;

            templateRow = table.Rows[templateRowIndex];

            // Chèn N dòng dựa trên template row (insert trước template row)
            for (int i = 0; i < lines.Count; i++)
            {
                var item = lines[i];

                // Clone row format
                var newRow = table.InsertRow(templateRow, templateRowIndex);
                ReplaceInRow(newRow, item);

                templateRowIndex++; // vì template row bị đẩy xuống dưới
            }

            // Xoá template row còn lại
            table.RemoveRow(templateRowIndex);
        }

        static bool RowContainsToken(Row row, string token)
        {
            return row.Cells.Any(c => c.Paragraphs.Any(p => (p.Text ?? "").Contains(token)));
        }

        static void ReplaceInRow(Row row, BusinessLine item)
        {
            foreach (var cell in row.Cells)
            {
                foreach (var p in cell.Paragraphs)
                {
                    p.ReplaceText("{BL_STT}", item.Stt.ToString(), false, RegexOptions.None);
                    p.ReplaceText("{BL_CODE}", item.Code ?? "", false, RegexOptions.None);
                    p.ReplaceText("{BL_NAME}", item.Name ?? "", false, RegexOptions.None);
                    p.ReplaceText("{BL_MAIN}", item.IsMain ? "X" : "", false, RegexOptions.None);
                }
            }
        }

        // =========================================================
        // 3) Replace placeholders {KEY} (trừ danh sách loại trừ)
        // =========================================================
        static void ReplaceEverywhereExcept(DocX doc, Dictionary<string, string> map, HashSet<string> excludeKeys)
        {
            // Body
            ReplaceInParagraphsExcept(doc.Paragraphs, map, excludeKeys);
            foreach (var t in doc.Tables) ReplaceInTableExcept(t, map, excludeKeys);

            // Header/Footer
            foreach (var s in doc.Sections)
            {
                if (s.Headers != null)
                {
                    if (s.Headers.First != null) ReplaceInHeaderFooterExcept(s.Headers.First, map, excludeKeys);
                    if (s.Headers.Odd != null) ReplaceInHeaderFooterExcept(s.Headers.Odd, map, excludeKeys);
                    if (s.Headers.Even != null) ReplaceInHeaderFooterExcept(s.Headers.Even, map, excludeKeys);
                }
                if (s.Footers != null)
                {
                    if (s.Footers.First != null) ReplaceInHeaderFooterExcept(s.Footers.First, map, excludeKeys);
                    if (s.Footers.Odd != null) ReplaceInHeaderFooterExcept(s.Footers.Odd, map, excludeKeys);
                    if (s.Footers.Even != null) ReplaceInHeaderFooterExcept(s.Footers.Even, map, excludeKeys);
                }
            }
        }

        static void ReplaceInHeaderFooterExcept(dynamic hf, Dictionary<string, string> map, HashSet<string> excludeKeys)
        {
            ReplaceInParagraphsExcept(hf.Paragraphs, map, excludeKeys);
            foreach (var t in hf.Tables) ReplaceInTableExcept(t, map, excludeKeys);
        }

        static void ReplaceInTableExcept(Table table, Dictionary<string, string> map, HashSet<string> excludeKeys)
        {
            foreach (var row in table.Rows)
                foreach (var cell in row.Cells)
                {
                    ReplaceInParagraphsExcept(cell.Paragraphs, map, excludeKeys);
                    foreach (var nested in cell.Tables) ReplaceInTableExcept(nested, map, excludeKeys);
                }
        }

        static void ReplaceInParagraphsExcept(IEnumerable<Paragraph> paragraphs, Dictionary<string, string> map, HashSet<string> excludeKeys)
        {
            foreach (var p in paragraphs)
            {
                foreach (var kv in map)
                {
                    if (excludeKeys.Contains(kv.Key))
                        continue;

                    string token = "{" + kv.Key + "}";
                    string value = kv.Value ?? "";
                    p.ReplaceText(token, value, false, RegexOptions.None);
                }
            }
        }
    }

}