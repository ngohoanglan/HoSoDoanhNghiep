
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
            // 2 file nằm ở thư mục chạy chương trình
            string baseDir = AppContext.BaseDirectory;
            string notesPath = Path.Combine(baseDir, "Notes_Placeholder.txt");
            string inputDocx = Path.Combine(baseDir, "Ủy Quyền.docx");
            string outputDocx = Path.Combine(baseDir, "Ủy Quyền_DA_DIEN.docx");

            if (!File.Exists(notesPath))
            {
                Console.Error.WriteLine("Không tìm thấy Notes_Placeholder.txt tại: " + notesPath);
                return 1;
            }
            if (!File.Exists(inputDocx))
            {
                Console.Error.WriteLine("Không tìm thấy Ủy Quyền.docx tại: " + inputDocx);
                return 1;
            }

            var map = LoadNotesKeyBraceValue(notesPath);

            using var doc = DocX.Load(inputDocx);

            // Replace trong toàn bộ document (paragraphs, tables, headers/footers)
            ReplaceEverywhere(doc, map);

            doc.SaveAs(outputDocx);
            Console.WriteLine("OK: " + outputDocx);
            return 0;
        }
        /// <summary>
        /// Parse Notes format: KEY={VALUE}
        /// - Giữ cả key có value trống: KEY={}
        /// - Bỏ qua dòng trống
        /// - Bỏ qua các dòng không đúng format
        /// </summary>
        static Dictionary<string, string> LoadNotesKeyBraceValue(string path)
        {
            var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            foreach (var raw in File.ReadAllLines(path, Encoding.UTF8))
            {
                var line = raw.Trim();
                if (string.IsNullOrWhiteSpace(line)) continue;

                // Match KEY={VALUE}
                var m = Regex.Match(line, @"^([A-Za-z0-9_]+)\s*=\s*\{(.*)\}\s*$");
                if (!m.Success) continue;

                string key = m.Groups[1].Value.Trim();
                string value = m.Groups[2].Value; // value bên trong {...} (có thể rỗng)

                dict[key] = value;
            }

            return dict;
        }

        static void ReplaceEverywhere(DocX doc, Dictionary<string, string> map)
        {
            ReplaceInParagraphs(doc.Paragraphs, map);

            foreach (var t in doc.Tables)
                ReplaceInTable(t, map);

            // Header/Footer theo sections
            foreach (var s in doc.Sections)
            {
                if (s.Headers != null)
                {
                    if (s.Headers.First != null) ReplaceInHeaderFooter(s.Headers.First, map);
                    if (s.Headers.Odd != null) ReplaceInHeaderFooter(s.Headers.Odd, map);
                    if (s.Headers.Even != null) ReplaceInHeaderFooter(s.Headers.Even, map);
                }
                if (s.Footers != null)
                {
                    if (s.Footers.First != null) ReplaceInHeaderFooter(s.Footers.First, map);
                    if (s.Footers.Odd != null) ReplaceInHeaderFooter(s.Footers.Odd, map);
                    if (s.Footers.Even != null) ReplaceInHeaderFooter(s.Footers.Even, map);
                }
            }
        }

        static void ReplaceInHeaderFooter(dynamic hf, Dictionary<string, string> map)
        {
            ReplaceInParagraphs(hf.Paragraphs, map);
            foreach (var t in hf.Tables) ReplaceInTable(t, map);
        }

        static void ReplaceInTable(Table table, Dictionary<string, string> map)
        {
            foreach (var row in table.Rows)
                foreach (var cell in row.Cells)
                {
                    ReplaceInParagraphs(cell.Paragraphs, map);

                    // nested tables
                    foreach (var nested in cell.Tables)
                        ReplaceInTable(nested, map);
                }
        }

        /// <summary>
        /// Replace đúng token {KEY} bằng VALUE. (Không thay KEY trần)
        /// Dùng ReplaceText dạng literal để tránh regex-run issues.
        /// </summary>
        static void ReplaceInParagraphs(IEnumerable<Paragraph> paragraphs, Dictionary<string, string> map)
        {
            foreach (var p in paragraphs)
            {
                foreach (var kv in map)
                {
                    string token = "{" + kv.Key + "}";
                    string value = kv.Value ?? "";

                    // replace literal token
                    // matchCase=false để không phân biệt hoa/thường (tuỳ bạn, nếu muốn strict thì để true)
                    p.ReplaceText(token, value, false, RegexOptions.None);
                }
            }
        }
    }
}
