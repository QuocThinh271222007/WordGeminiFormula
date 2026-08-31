using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using WordGeminiFormula.AddIn.Models;

namespace WordGeminiFormula.AddIn.Word
{
    /// <summary>
    /// Safe renderer for OCR documents. Pending formulas are tracked with ordinary Word
    /// Bookmarks, never Content Controls. Bookmarks follow edits without creating a
    /// restricted/limited editing zone, so Word can continue TypeText/TypeParagraph after
    /// every inline formula. OMath conversion is performed only in a separate pass.
    /// </summary>
    public sealed class WordDocumentServiceV3
    {
        private readonly dynamic _wordApplication;
        private readonly LatexToUnicodeMathConverter _converter = new LatexToUnicodeMathConverter();
        private readonly MathRepairService _repair = new MathRepairService();

        private const string MathStart = "[[MATH]]";
        private const string MathEnd = "[[/MATH]]";
        private const string PendingMathBookmarkPrefix = "WGF_MATH_";
        private const int WdCollapseEnd = 0;
        private const int WdFindStop = 0;
        private const int WdYellow = 7;

        public int LastNormalizationFailureCount { get; private set; }

        public WordDocumentServiceV3(object wordApplication)
        {
            _wordApplication = wordApplication ?? throw new ArgumentNullException(nameof(wordApplication));
        }

        public void InsertOcrBlocks(
            OcrDocument document,
            bool autoNormalize,
            bool autoBeautify,
            string sourceImagePath,
            bool preserveDifficultRegions)
        {
            if (document?.blocks == null) return;

            dynamic selection = _wordApplication.Selection;
            selection.Collapse(WdCollapseEnd);
            LastNormalizationFailureCount = 0;

            if (autoBeautify)
                ApplyPageDefaults(_wordApplication.ActiveDocument);

            for (int i = 0; i < document.blocks.Count; i++)
            {
                OcrBlock block = document.blocks[i];
                if (block == null) continue;
                string type = (block.type ?? "text").Trim().ToLowerInvariant();

                if (autoBeautify && type == "header_left" && i + 1 < document.blocks.Count &&
                    string.Equals(document.blocks[i + 1]?.type, "header_right", StringComparison.OrdinalIgnoreCase))
                {
                    InsertHeaderPair(block, document.blocks[++i]);
                    continue;
                }

                switch (type)
                {
                    case "header_left":
                    case "header_right":
                    case "title":
                        InsertStyledParagraph(block.text, 1, true, type == "title" ? 13f : 11f, 3f, 0f);
                        break;
                    case "subtitle":
                    case "meta":
                        InsertStyledParagraph(block.text, 1, type == "subtitle", 11f, 2f, 0f);
                        break;
                    case "candidate_field":
                        InsertCandidateField(block);
                        break;
                    case "code_box":
                        InsertCodeBox(block);
                        break;
                    case "section":
                        InsertStyledParagraph(block.text, 0, true, 11.5f, 4f, 6f);
                        break;
                    case "question":
                        InsertQuestion(block, autoBeautify);
                        break;
                    case "formula":
                        InsertStandaloneFormula(block);
                        break;
                    case "figure":
                    case "table_image":
                    case "unresolved":
                        InsertDifficultRegion(block, sourceImagePath, preserveDifficultRegions);
                        break;
                    case "footer":
                        InsertStyledParagraph(block.text, 2, false, 9.5f, 0f, 6f, true);
                        break;
                    default:
                        InsertStyledParagraph(block.text, 0, false, 11.5f, 3f, 0f);
                        break;
                }
            }

            // Deliberately do not create OMath here. Connect performs normalization only
            // after this render method returns and the entire page is complete.
        }

        public int BeautifyActiveDocument()
        {
            dynamic document = _wordApplication.ActiveDocument;
            ApplyPageDefaults(document);
            int formatted = 0;
            int paragraphCount = (int)document.Paragraphs.Count;

            for (int i = 1; i <= paragraphCount; i++)
            {
                dynamic paragraph = document.Paragraphs.Item(i);
                dynamic range = paragraph.Range;
                string raw = ((string)range.Text ?? string.Empty).Replace("\r", string.Empty).Replace("\a", string.Empty);
                string text = raw.Trim();
                if (string.IsNullOrWhiteSpace(text)) continue;

                range.Font.Name = "Times New Roman";
                range.Font.Size = 11.5f;
                paragraph.Format.Alignment = 0;
                paragraph.Format.SpaceAfter = 3f;
                paragraph.Format.SpaceBefore = 0f;
                paragraph.Format.LeftIndent = 0f;
                paragraph.Format.FirstLineIndent = 0f;

                if (text.StartsWith("BỘ GIÁO DỤC", StringComparison.OrdinalIgnoreCase) ||
                    text.StartsWith("KỲ THI", StringComparison.OrdinalIgnoreCase))
                {
                    paragraph.Format.Alignment = 1;
                    range.Font.Bold = 1;
                    range.Font.Size = 12f;
                }
                else if (text.StartsWith("ĐỀ THI", StringComparison.OrdinalIgnoreCase) ||
                         text.StartsWith("Môn thi:", StringComparison.OrdinalIgnoreCase) ||
                         text.StartsWith("Thời gian làm bài:", StringComparison.OrdinalIgnoreCase))
                {
                    paragraph.Format.Alignment = 1;
                    range.Font.Bold = text.StartsWith("ĐỀ THI", StringComparison.OrdinalIgnoreCase) ? 1 : 0;
                    range.Font.Size = 10.5f;
                }
                else if (text.StartsWith("PHẦN ", StringComparison.OrdinalIgnoreCase))
                {
                    range.Font.Bold = 1;
                    paragraph.Format.SpaceBefore = 6f;
                    paragraph.Format.SpaceAfter = 4f;
                }
                else if (Regex.IsMatch(text, @"^Câu\s+\d+\.", RegexOptions.IgnoreCase))
                {
                    Match m = Regex.Match(text, @"^Câu\s+\d+\.", RegexOptions.IgnoreCase);
                    dynamic prefix = document.Range((int)range.Start, (int)range.Start + m.Length);
                    prefix.Font.Bold = 1;
                    paragraph.Format.KeepTogether = -1;
                }
                else if (Regex.IsMatch(text, @"^[A-D]\.", RegexOptions.IgnoreCase))
                {
                    dynamic prefix = document.Range((int)range.Start, Math.Min((int)range.Start + 2, (int)range.End));
                    prefix.Font.Bold = 1;
                    paragraph.Format.LeftIndent = 18f;
                }
                else if (text.StartsWith("Trang ", StringComparison.OrdinalIgnoreCase))
                {
                    paragraph.Format.Alignment = 2;
                    range.Font.Size = 9.5f;
                    range.Font.Italic = 1;
                }

                formatted++;
            }

            return formatted;
        }

        public int NormalizeAllMarkedFormulas()
        {
            dynamic document = _wordApplication.ActiveDocument;
            LastNormalizationFailureCount = 0;
            int converted = 0;

            // Bookmark ranges are owned and tracked by Word. Snapshot names/starts first,
            // then process from the end of the document backwards so equation expansion
            // cannot disturb any earlier pending formula.
            var pending = new List<PendingBookmark>();
            dynamic bookmarks = document.Bookmarks;
            int count = (int)bookmarks.Count;
            for (int i = 1; i <= count; i++)
            {
                dynamic bookmark = bookmarks.Item(i);
                string name = Convert.ToString(bookmark.Name) ?? string.Empty;
                if (!name.StartsWith(PendingMathBookmarkPrefix, StringComparison.OrdinalIgnoreCase)) continue;
                pending.Add(new PendingBookmark { Name = name, Start = (int)bookmark.Range.Start });
            }

            foreach (PendingBookmark info in pending.OrderByDescending(x => x.Start))
            {
                try
                {
                    if (!(bool)document.Bookmarks.Exists(info.Name)) continue;
                    dynamic bookmark = document.Bookmarks.Item(info.Name);
                    if (TryNormalizePendingBookmark(bookmark))
                        converted++;
                    else
                        LastNormalizationFailureCount++;
                }
                catch
                {
                    LastNormalizationFailureCount++;
                }
            }

            // Compatibility for pasted formulas and documents made by pre-bookmark builds.
            converted += NormalizeLegacyMarkersWithWordFind();
            return converted;
        }

        public void NormalizeSelection()
        {
            dynamic selection = _wordApplication.Selection;
            if ((int)selection.Start == (int)selection.End)
                throw new InvalidOperationException("Hãy bôi đen công thức LaTeX/linear cần chuẩn hóa trước.");

            string raw = ExtractMathPayload(((string)selection.Text ?? string.Empty).Trim());
            if (string.IsNullOrWhiteSpace(raw))
                throw new InvalidOperationException("Vùng chọn không có công thức.");

            string linear = ConvertToWordLinear(raw, null);
            dynamic range = selection.Range;
            int end = ReplaceRangeWithEquation(range, linear);
            selection.SetRange(end, end);
            selection.Collapse(WdCollapseEnd);
        }

        private void InsertHeaderPair(OcrBlock left, OcrBlock right)
        {
            dynamic document = _wordApplication.ActiveDocument;
            dynamic selection = _wordApplication.Selection;
            dynamic table = document.Tables.Add(selection.Range, 1, 2);
            table.Borders.Enable = 0;
            table.AllowAutoFit = -1;
            SetCellText(table.Cell(1, 1), left?.text, true, 10.5f, 1);
            SetCellText(table.Cell(1, 2), right?.text, true, 10.5f, 1);
            MoveSelectionAfterTable(table);
        }

        private static void SetCellText(dynamic cell, string text, bool bold, float size, int alignment)
        {
            dynamic range = cell.Range;
            range.Text = text ?? string.Empty;
            range.Font.Name = "Times New Roman";
            range.Font.Size = size;
            range.Font.Bold = bold ? 1 : 0;
            range.ParagraphFormat.Alignment = alignment;
            range.ParagraphFormat.SpaceAfter = 0f;
        }

        private void InsertCandidateField(OcrBlock block)
        {
            dynamic document = _wordApplication.ActiveDocument;
            dynamic selection = _wordApplication.Selection;
            ResetSelectionStyle(selection);
            int start = (int)selection.Start;
            string label = string.IsNullOrWhiteSpace(block.label) ? "Thông tin" : block.label.Trim();
            string value = string.IsNullOrWhiteSpace(block.text) ? new string('.', 62) : block.text.Trim();
            selection.TypeText(label + ": " + value);
            dynamic prefix = document.Range(start, start + label.Length + 1);
            prefix.Font.Bold = 1;
            selection.TypeParagraph();
        }

        private void InsertCodeBox(OcrBlock block)
        {
            dynamic document = _wordApplication.ActiveDocument;
            dynamic selection = _wordApplication.Selection;
            dynamic table = document.Tables.Add(selection.Range, 1, 1);
            table.Borders.Enable = 1;
            try { table.Rows.Alignment = 2; } catch { }
            string text = (string.IsNullOrWhiteSpace(block.label) ? "Mã đề" : block.label.Trim()) + ": " + (block.text ?? string.Empty).Trim();
            SetCellText(table.Cell(1, 1), text, true, 10.5f, 1);
            try { table.Columns.Item(1).PreferredWidth = 110f; } catch { }
            MoveSelectionAfterTable(table);
        }

        private void InsertQuestion(OcrBlock block, bool beautify)
        {
            dynamic document = _wordApplication.ActiveDocument;
            dynamic selection = _wordApplication.Selection;
            ResetSelectionStyle(selection);
            int paragraphStart = (int)selection.Start;
            string prefix = string.IsNullOrWhiteSpace(block.number) ? string.Empty : "Câu " + block.number.Trim() + ". ";
            selection.TypeText(prefix);

            if (block.content != null && block.content.Count > 0)
                InsertInlineParts(block.content);
            else
                selection.TypeText(block.text ?? string.Empty);

            int paragraphEnd = (int)selection.End;
            if (!string.IsNullOrEmpty(prefix))
            {
                dynamic prefixRange = document.Range(paragraphStart, paragraphStart + prefix.Length);
                prefixRange.Font.Bold = 1;
            }
            dynamic questionRange = document.Range(paragraphStart, paragraphEnd);
            questionRange.ParagraphFormat.SpaceAfter = 3f;
            questionRange.ParagraphFormat.KeepTogether = -1;

            // Selection is guaranteed to be ordinary editable text because math fragments
            // are only bookmarks at this stage, never Content Controls or OMath zones.
            selection.SetRange(paragraphEnd, paragraphEnd);
            selection.Collapse(WdCollapseEnd);
            selection.TypeParagraph();

            if (block.choices == null || block.choices.Count == 0) return;
            if (beautify && block.choices.Count >= 2)
                InsertChoiceTable(block.choices);
            else
                InsertChoiceParagraphs(block.choices);
        }

        private void InsertChoiceTable(List<OcrChoice> choices)
        {
            dynamic document = _wordApplication.ActiveDocument;
            dynamic selection = _wordApplication.Selection;
            const int cols = 2;
            int rows = (int)Math.Ceiling(choices.Count / 2.0);
            dynamic table = document.Tables.Add(selection.Range, rows, cols);
            table.Borders.Enable = 0;
            table.AllowAutoFit = -1;

            for (int i = 0; i < choices.Count; i++)
            {
                int row = (i / cols) + 1;
                int col = (i % cols) + 1;
                OcrChoice choice = choices[i];
                string label = string.IsNullOrWhiteSpace(choice?.label) ? ((char)('A' + i)).ToString() : choice.label.Trim();
                dynamic cell = table.Cell(row, col);
                int cellStart = (int)cell.Range.Start;
                int cellEnd = Math.Max(cellStart, (int)cell.Range.End - 1);

                selection.SetRange(cellStart, cellEnd);
                selection.Text = string.Empty;
                selection.SetRange(cellStart, cellStart);
                ResetSelectionStyle(selection);
                selection.ParagraphFormat.SpaceAfter = 1f;
                selection.TypeText(label + ". ");
                InsertInlineParts(choice?.content);

                dynamic labelRange = document.Range(cellStart, cellStart + label.Length + 1);
                labelRange.Font.Bold = 1;
            }

            MoveSelectionAfterTable(table);
        }

        private void InsertChoiceParagraphs(List<OcrChoice> choices)
        {
            dynamic document = _wordApplication.ActiveDocument;
            dynamic selection = _wordApplication.Selection;
            for (int i = 0; i < choices.Count; i++)
            {
                OcrChoice choice = choices[i];
                string label = string.IsNullOrWhiteSpace(choice?.label) ? ((char)('A' + i)).ToString() : choice.label.Trim();
                ResetSelectionStyle(selection);
                int start = (int)selection.Start;
                selection.TypeText(label + ". ");
                InsertInlineParts(choice?.content);
                dynamic prefix = document.Range(start, start + label.Length + 1);
                prefix.Font.Bold = 1;
                selection.TypeParagraph();
            }
        }

        private void InsertStandaloneFormula(OcrBlock block)
        {
            dynamic selection = _wordApplication.Selection;
            ResetSelectionStyle(selection);
            selection.ParagraphFormat.Alignment = block.display ? 1 : 0;
            InsertMathFragment(block.latex, block.word_linear);
            selection.TypeParagraph();
        }

        private void InsertInlineParts(IEnumerable<OcrInline> parts)
        {
            if (parts == null) return;
            dynamic selection = _wordApplication.Selection;
            bool previousWasMath = false;

            foreach (OcrInline part in parts)
            {
                if (part == null) continue;
                bool isMath = string.Equals(part.type, "math", StringComparison.OrdinalIgnoreCase) ||
                              string.Equals(part.type, "formula", StringComparison.OrdinalIgnoreCase);

                if (isMath)
                {
                    InsertMathFragment(part.latex, part.word_linear);
                    previousWasMath = true;
                }
                else
                {
                    string text = part.text ?? string.Empty;
                    if (previousWasMath && text.Length > 0 && !char.IsWhiteSpace(text[0]) && !IsNoLeadingSpacePunctuation(text[0]))
                        selection.TypeText(" ");
                    selection.TypeText(text);
                    previousWasMath = false;
                }
            }
        }

        private void InsertMathFragment(string latex, string wordLinear)
        {
            string payload = _repair.RepairLatex(!string.IsNullOrWhiteSpace(latex) ? latex : wordLinear ?? string.Empty);
            if (string.IsNullOrWhiteSpace(payload))
            {
                payload = "[OCR_MATH_EMPTY]";
                LastNormalizationFailureCount++;
            }
            InsertPendingMathBookmark(payload, string.Equals(payload, "[OCR_MATH_EMPTY]", StringComparison.Ordinal));
        }

        private void InsertPendingMathBookmark(string latex, bool highlight)
        {
            dynamic document = _wordApplication.ActiveDocument;
            dynamic selection = _wordApplication.Selection;
            string payload = _repair.RepairLatex(latex ?? string.Empty);
            string marker = MathStart + payload + MathEnd;

            int start = (int)selection.Start;
            selection.TypeText(marker);
            int end = (int)selection.End;
            dynamic range = document.Range(start, end);

            string bookmarkName = PendingMathBookmarkPrefix + Guid.NewGuid().ToString("N").Substring(0, 24);
            document.Bookmarks.Add(bookmarkName, range);
            if (highlight) range.HighlightColorIndex = WdYellow;

            // A bookmark is non-locking. Explicitly collapse to its end to guarantee the
            // next TypeText/TypeParagraph is outside the tracked formula range.
            selection.SetRange(end, end);
            selection.Collapse(WdCollapseEnd);
        }

        private bool TryNormalizePendingBookmark(dynamic bookmark)
        {
            dynamic document = _wordApplication.ActiveDocument;
            dynamic range = bookmark.Range;
            string rawMarker = (string)range.Text ?? string.Empty;
            string raw = ExtractMathPayload(rawMarker);
            int start = (int)range.Start;

            try
            {
                string linear = ConvertToWordLinear(raw, null);
                string name = Convert.ToString(bookmark.Name) ?? string.Empty;
                bookmark.Delete();
                dynamic exact = document.Range(start, start + rawMarker.Length);
                ReplaceRangeWithEquation(exact, linear);
                return true;
            }
            catch
            {
                try
                {
                    dynamic failed = document.Range(start, Math.Min((int)document.Content.End, start + rawMarker.Length));
                    failed.HighlightColorIndex = WdYellow;
                }
                catch { }
                return false;
            }
        }

        private int NormalizeLegacyMarkersWithWordFind()
        {
            dynamic document = _wordApplication.ActiveDocument;
            int converted = 0;
            int cursor = 0;

            while (cursor < (int)document.Content.End)
            {
                dynamic startMarker = FindTextRange(document, MathStart, cursor);
                if (startMarker == null) break;
                dynamic endMarker = FindTextRange(document, MathEnd, (int)startMarker.End);
                if (endMarker == null)
                {
                    try { startMarker.HighlightColorIndex = WdYellow; } catch { }
                    LastNormalizationFailureCount++;
                    break;
                }

                int fullStart = (int)startMarker.Start;
                int fullEnd = (int)endMarker.End;
                dynamic payloadRange = document.Range((int)startMarker.End, (int)endMarker.Start);
                string raw = ((string)payloadRange.Text ?? string.Empty).Trim();
                dynamic fullRange = document.Range(fullStart, fullEnd);

                try
                {
                    string linear = ConvertToWordLinear(raw, null);
                    cursor = ReplaceRangeWithEquation(fullRange, linear);
                    converted++;
                }
                catch
                {
                    try { fullRange.HighlightColorIndex = WdYellow; } catch { }
                    LastNormalizationFailureCount++;
                    cursor = Math.Max(fullStart + 1, fullEnd);
                }
            }

            return converted;
        }

        private static dynamic FindTextRange(dynamic document, string needle, int start)
        {
            int documentEnd = (int)document.Content.End;
            if (start >= documentEnd) return null;
            dynamic range = document.Range(Math.Max(0, start), documentEnd);
            dynamic find = range.Find;
            find.ClearFormatting();
            find.Text = needle;
            find.Forward = true;
            find.Wrap = WdFindStop;
            find.Format = false;
            bool found = find.Execute();
            return found ? range : null;
        }

        private string ConvertToWordLinear(string latex, string preferredWordLinear)
        {
            if (!string.IsNullOrWhiteSpace(preferredWordLinear))
            {
                string preferred = _repair.RepairWordLinear(preferredWordLinear);
                if (!string.IsNullOrWhiteSpace(preferred)) return preferred;
            }

            string repaired = _repair.RepairLatex(latex ?? string.Empty);
            string linear = _converter.Convert(repaired);
            if (string.IsNullOrWhiteSpace(linear))
                throw new InvalidOperationException("Biểu thức rỗng sau khi chuyển đổi.");
            return linear;
        }

        private static string ExtractMathPayload(string value)
        {
            string raw = (value ?? string.Empty).Trim();
            if (raw.StartsWith(MathStart, StringComparison.Ordinal) && raw.EndsWith(MathEnd, StringComparison.Ordinal))
                return raw.Substring(MathStart.Length, raw.Length - MathStart.Length - MathEnd.Length).Trim();
            return raw.Replace(MathStart, string.Empty).Replace(MathEnd, string.Empty).Trim();
        }

        private static bool IsNoLeadingSpacePunctuation(char c)
        {
            return c == '.' || c == ',' || c == ';' || c == ':' || c == '?' || c == '!' || c == ')' || c == ']' || c == '}';
        }

        private void InsertDifficultRegion(OcrBlock block, string sourceImagePath, bool preserve)
        {
            dynamic selection = _wordApplication.Selection;
            if (preserve && TryCropSourceRegion(sourceImagePath, block?.bbox, out string cropPath))
            {
                try
                {
                    dynamic shape = selection.InlineShapes.AddPicture(cropPath, false, true);
                    try
                    {
                        dynamic document = _wordApplication.ActiveDocument;
                        float maxWidth = (float)document.PageSetup.PageWidth - (float)document.PageSetup.LeftMargin - (float)document.PageSetup.RightMargin;
                        if ((float)shape.Width > maxWidth) shape.Width = maxWidth;
                    }
                    catch { }
                    selection.TypeParagraph();
                    return;
                }
                finally
                {
                    try { File.Delete(cropPath); } catch { }
                }
            }

            ResetSelectionStyle(selection);
            int start = (int)selection.Start;
            string reason = string.IsNullOrWhiteSpace(block?.text) ? "Vùng hình/bảng chưa thể chuyển đổi chắc chắn" : block.text.Trim();
            selection.TypeText("[CẦN KIỂM TRA: " + reason + "]");
            int end = (int)selection.End;
            dynamic range = _wordApplication.ActiveDocument.Range(start, end);
            range.HighlightColorIndex = WdYellow;
            range.Font.Italic = 1;
            selection.TypeParagraph();
        }

        private static bool TryCropSourceRegion(string sourceImagePath, OcrBoundingBox bbox, out string cropPath)
        {
            cropPath = null;
            if (string.IsNullOrWhiteSpace(sourceImagePath) || !File.Exists(sourceImagePath) || bbox == null)
                return false;

            double x = bbox.x;
            double y = bbox.y;
            double w = bbox.width;
            double h = bbox.height;
            if (x > 1 || y > 1 || w > 1 || h > 1)
            {
                x /= 1000.0;
                y /= 1000.0;
                w /= 1000.0;
                h /= 1000.0;
            }

            x = Clamp01(x); y = Clamp01(y); w = Clamp01(w); h = Clamp01(h);
            if (w < 0.01 || h < 0.01) return false;

            using (Image source = Image.FromFile(sourceImagePath))
            {
                int left = Math.Max(0, Math.Min(source.Width - 1, (int)Math.Round(x * source.Width)));
                int top = Math.Max(0, Math.Min(source.Height - 1, (int)Math.Round(y * source.Height)));
                int width = Math.Max(1, Math.Min(source.Width - left, (int)Math.Round(w * source.Width)));
                int height = Math.Max(1, Math.Min(source.Height - top, (int)Math.Round(h * source.Height)));
                if (width < 2 || height < 2) return false;

                using (var bitmap = new Bitmap(width, height))
                using (Graphics graphics = Graphics.FromImage(bitmap))
                {
                    graphics.DrawImage(source, new Rectangle(0, 0, width, height), new Rectangle(left, top, width, height), GraphicsUnit.Pixel);
                    cropPath = Path.Combine(Path.GetTempPath(), "WordGeminiFormula-crop-" + Guid.NewGuid().ToString("N") + ".png");
                    bitmap.Save(cropPath, ImageFormat.Png);
                    return true;
                }
            }
        }

        private static double Clamp01(double value)
        {
            if (value < 0) return 0;
            if (value > 1) return 1;
            return value;
        }

        private void InsertStyledParagraph(string text, int alignment, bool bold, float size, float after, float before, bool italic = false)
        {
            dynamic selection = _wordApplication.Selection;
            ResetSelectionStyle(selection);
            selection.ParagraphFormat.Alignment = alignment;
            selection.ParagraphFormat.SpaceAfter = after;
            selection.ParagraphFormat.SpaceBefore = before;
            selection.Font.Bold = bold ? 1 : 0;
            selection.Font.Italic = italic ? 1 : 0;
            selection.Font.Size = size;
            selection.TypeText(text ?? string.Empty);
            selection.TypeParagraph();
        }

        private void MoveSelectionAfterTable(dynamic table)
        {
            dynamic selection = _wordApplication.Selection;
            int end = (int)table.Range.End;
            selection.SetRange(end, end);
            selection.Collapse(WdCollapseEnd);
            selection.TypeParagraph();
        }

        private static void ResetSelectionStyle(dynamic selection)
        {
            selection.Font.Name = "Times New Roman";
            selection.Font.Size = 11.5f;
            selection.Font.Bold = 0;
            selection.Font.Italic = 0;
            selection.ParagraphFormat.Alignment = 0;
            selection.ParagraphFormat.SpaceAfter = 3f;
            selection.ParagraphFormat.SpaceBefore = 0f;
            selection.ParagraphFormat.LeftIndent = 0f;
            selection.ParagraphFormat.FirstLineIndent = 0f;
        }

        private static void ApplyPageDefaults(dynamic document)
        {
            try
            {
                document.PageSetup.TopMargin = 42.5f;
                document.PageSetup.BottomMargin = 42.5f;
                document.PageSetup.LeftMargin = 48f;
                document.PageSetup.RightMargin = 48f;
            }
            catch { }

            try
            {
                dynamic normal = document.Styles.Item("Normal");
                normal.Font.Name = "Times New Roman";
                normal.Font.Size = 11.5f;
                normal.ParagraphFormat.SpaceAfter = 3f;
            }
            catch { }
        }

        private int ReplaceRangeWithEquation(dynamic range, string wordLinear)
        {
            int start = (int)range.Start;
            string linear = wordLinear ?? string.Empty;
            range.Text = linear;
            dynamic exactRange = _wordApplication.ActiveDocument.Range(start, start + linear.Length);
            return ConvertExistingRangeToEquation(exactRange);
        }

        private int ConvertExistingRangeToEquation(dynamic range)
        {
            dynamic document = _wordApplication.ActiveDocument;
            dynamic eqRange = document.OMaths.Add(range);
            if ((int)eqRange.OMaths.Count <= 0)
                throw new InvalidOperationException("Word không tạo được OMath từ biểu thức.");

            dynamic equation = eqRange.OMaths.Item(1);
            equation.BuildUp();
            return (int)eqRange.End;
        }

        private sealed class PendingBookmark
        {
            public string Name { get; set; }
            public int Start { get; set; }
        }
    }
}
