using System;
using System.Collections.Generic;
using System.Linq;
using WordGeminiFormula.AddIn.Models;

namespace WordGeminiFormula.AddIn.Word
{
    /// <summary>
    /// V4 presentation layer on top of the bookmark-safe V3 renderer.
    /// Adds Word-safe math canonicalization, visual anchoring for exam questions,
    /// compact one-page exam formatting, punctuation hygiene, and real footer placement.
    /// </summary>
    public sealed class WordDocumentServiceV4
    {
        private readonly dynamic _wordApplication;
        private readonly WordDocumentServiceV3 _inner;
        private readonly WordMathCanonicalizer _canonicalizer = new WordMathCanonicalizer();

        public int LastNormalizationFailureCount => _inner.LastNormalizationFailureCount;

        public WordDocumentServiceV4(object wordApplication)
        {
            _wordApplication = wordApplication ?? throw new ArgumentNullException(nameof(wordApplication));
            _inner = new WordDocumentServiceV3(wordApplication);
        }

        public void InsertOcrBlocks(
            OcrDocument document,
            bool autoNormalize,
            bool autoBeautify,
            string sourceImagePath,
            bool preserveDifficultRegions)
        {
            if (document == null) return;

            string footerText;
            OcrDocument prepared = PrepareDocument(document, out footerText);

            // V3 remains responsible for the safe bookmark-based render. Never create
            // OMath during this pass; Connect performs normalization after render returns.
            _inner.InsertOcrBlocks(
                prepared,
                false,
                autoBeautify,
                sourceImagePath,
                preserveDifficultRegions);

            if (IsExam(prepared))
                ApplyCompactExamLayout();

            CenterPreservedVisuals();
            if (!string.IsNullOrWhiteSpace(footerText))
                SetPrimaryFooter(footerText);
        }

        public int BeautifyActiveDocument()
        {
            int count = _inner.BeautifyActiveDocument();
            ApplyCompactExamLayout();
            CenterPreservedVisuals();
            return count;
        }

        public int NormalizeAllMarkedFormulas() => _inner.NormalizeAllMarkedFormulas();
        public void NormalizeSelection() => _inner.NormalizeSelection();

        private OcrDocument PrepareDocument(OcrDocument source, out string footerText)
        {
            footerText = null;
            var result = new OcrDocument
            {
                document_type = source.document_type ?? "general",
                warnings = source.warnings != null ? new List<string>(source.warnings) : new List<string>(),
                blocks = new List<OcrBlock>()
            };

            var blocks = source.blocks ?? new List<OcrBlock>();
            for (int i = 0; i < blocks.Count; i++)
            {
                OcrBlock block = blocks[i];
                if (block == null) continue;
                string type = (block.type ?? "text").Trim().ToLowerInvariant();

                if (type == "footer")
                {
                    if (!string.IsNullOrWhiteSpace(block.text)) footerText = block.text.Trim();
                    continue;
                }

                // Gemini commonly emits a visual block immediately after the whole question.
                // Re-anchor that visual between the question stem and choices. If the stem
                // contains a newline (e.g. Câu 10), text after that newline is rendered after
                // the image before choices.
                if (type == "question" && i + 1 < blocks.Count && IsVisualBlock(blocks[i + 1]))
                {
                    OcrBlock visual = blocks[i + 1];
                    List<OcrInline> before;
                    List<OcrInline> after;
                    SplitAtFirstNewline(block.content, out before, out after);

                    var stem = CloneQuestion(block);
                    stem.content = CanonicalizeParts(before.Count > 0 ? before : block.content);
                    stem.choices = new List<OcrChoice>();
                    result.blocks.Add(stem);
                    result.blocks.Add(CloneBlock(visual));

                    if (after.Count > 0 || (block.choices != null && block.choices.Count > 0))
                    {
                        var continuation = new OcrBlock
                        {
                            type = "question",
                            number = string.Empty,
                            content = CanonicalizeParts(after),
                            choices = CanonicalizeChoices(block.choices),
                            text = string.Empty
                        };
                        result.blocks.Add(continuation);
                    }

                    i++; // consume paired visual
                    continue;
                }

                result.blocks.Add(CanonicalizeBlock(block));
            }

            return result;
        }

        private OcrBlock CanonicalizeBlock(OcrBlock source)
        {
            OcrBlock block = CloneBlock(source);
            block.content = CanonicalizeParts(source.content);
            block.choices = CanonicalizeChoices(source.choices);

            if (string.Equals(block.type, "formula", StringComparison.OrdinalIgnoreCase))
            {
                string input = !string.IsNullOrWhiteSpace(block.latex) ? block.latex : block.word_linear;
                block.latex = _canonicalizer.NormalizeForWord(input);
                block.word_linear = null;
            }
            return block;
        }

        private List<OcrChoice> CanonicalizeChoices(List<OcrChoice> choices)
        {
            var result = new List<OcrChoice>();
            if (choices == null) return result;
            foreach (OcrChoice choice in choices)
            {
                if (choice == null) continue;
                result.Add(new OcrChoice
                {
                    label = choice.label,
                    content = CanonicalizeParts(choice.content)
                });
            }
            return result;
        }

        private List<OcrInline> CanonicalizeParts(IEnumerable<OcrInline> parts)
        {
            var result = new List<OcrInline>();
            if (parts == null) return result;

            foreach (OcrInline part in parts)
            {
                if (part == null) continue;
                bool isMath = string.Equals(part.type, "math", StringComparison.OrdinalIgnoreCase) ||
                              string.Equals(part.type, "formula", StringComparison.OrdinalIgnoreCase);
                if (!isMath)
                {
                    result.Add(CloneInline(part));
                    continue;
                }

                string input = !string.IsNullOrWhiteSpace(part.latex) ? part.latex : part.word_linear;
                string suffix;
                string core = SplitTrailingPunctuation(input, out suffix);
                string canonical = _canonicalizer.NormalizeForWord(core);

                result.Add(new OcrInline
                {
                    type = "math",
                    latex = canonical,
                    word_linear = null,
                    confidence = part.confidence
                });

                if (!string.IsNullOrEmpty(suffix))
                    result.Add(new OcrInline { type = "text", text = suffix, confidence = part.confidence });
            }
            return result;
        }

        private static string SplitTrailingPunctuation(string input, out string suffix)
        {
            suffix = string.Empty;
            string s = (input ?? string.Empty).Trim();
            while (s.Length > 0)
            {
                char c = s[s.Length - 1];
                if (c != '.' && c != ',' && c != ';' && c != ':' && c != '?' && c != '!') break;
                suffix = c + suffix;
                s = s.Substring(0, s.Length - 1).TrimEnd();
            }
            return s;
        }

        private static void SplitAtFirstNewline(List<OcrInline> content, out List<OcrInline> before, out List<OcrInline> after)
        {
            before = new List<OcrInline>();
            after = new List<OcrInline>();
            if (content == null) return;

            bool split = false;
            foreach (OcrInline part in content)
            {
                if (part == null) continue;
                if (!split && string.Equals(part.type, "text", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrEmpty(part.text))
                {
                    int nl = part.text.IndexOf('\n');
                    if (nl >= 0)
                    {
                        string left = part.text.Substring(0, nl).TrimEnd();
                        string right = part.text.Substring(nl + 1).TrimStart();
                        if (left.Length > 0) before.Add(new OcrInline { type = "text", text = left, confidence = part.confidence });
                        if (right.Length > 0) after.Add(new OcrInline { type = "text", text = right, confidence = part.confidence });
                        split = true;
                        continue;
                    }
                }

                (split ? after : before).Add(CloneInline(part));
            }
        }

        private void ApplyCompactExamLayout()
        {
            dynamic document = _wordApplication.ActiveDocument;
            try
            {
                document.PageSetup.TopMargin = 27f;
                document.PageSetup.BottomMargin = 27f;
                document.PageSetup.LeftMargin = 34f;
                document.PageSetup.RightMargin = 34f;
                document.PageSetup.HeaderDistance = 14f;
                document.PageSetup.FooterDistance = 14f;
            }
            catch { }

            try
            {
                dynamic normal = document.Styles.Item("Normal");
                normal.Font.Name = "Times New Roman";
                normal.Font.Size = 10f;
                normal.ParagraphFormat.SpaceAfter = 0f;
                normal.ParagraphFormat.SpaceBefore = 0f;
            }
            catch { }

            try
            {
                int paragraphCount = (int)document.Paragraphs.Count;
                for (int i = 1; i <= paragraphCount; i++)
                {
                    dynamic paragraph = document.Paragraphs.Item(i);
                    dynamic range = paragraph.Range;
                    string text = (((string)range.Text) ?? string.Empty).Replace("\r", string.Empty).Replace("\a", string.Empty).Trim();
                    paragraph.Format.SpaceAfter = 0f;
                    paragraph.Format.SpaceBefore = 0f;
                    paragraph.Format.LeftIndent = 0f;
                    paragraph.Format.FirstLineIndent = 0f;

                    if (string.IsNullOrWhiteSpace(text))
                    {
                        range.Font.Size = 1f;
                        try
                        {
                            paragraph.Format.LineSpacingRule = 4; // wdLineSpaceExactly
                            paragraph.Format.LineSpacing = 1f;
                        }
                        catch { }
                        continue;
                    }

                    range.Font.Name = "Times New Roman";
                    if (text.StartsWith("PHẦN ", StringComparison.OrdinalIgnoreCase))
                    {
                        range.Font.Size = 10f;
                        range.Font.Bold = 1;
                        paragraph.Format.SpaceBefore = 2f;
                        paragraph.Format.SpaceAfter = 1f;
                    }
                    else if (text.StartsWith("Câu ", StringComparison.OrdinalIgnoreCase))
                    {
                        range.Font.Size = 10f;
                        paragraph.Format.SpaceBefore = 1f;
                        paragraph.Format.SpaceAfter = 0f;
                    }
                    else if (text.StartsWith("Họ, tên thí sinh", StringComparison.OrdinalIgnoreCase) ||
                             text.StartsWith("Số báo danh", StringComparison.OrdinalIgnoreCase))
                    {
                        range.Font.Size = 9.5f;
                    }
                    else
                    {
                        if ((float)range.Font.Size > 10.5f || (float)range.Font.Size <= 0f)
                            range.Font.Size = 10f;
                    }
                }
            }
            catch { }

            try
            {
                int tableCount = (int)document.Tables.Count;
                for (int i = 1; i <= tableCount; i++)
                {
                    dynamic table = document.Tables.Item(i);
                    string text = (((string)table.Range.Text) ?? string.Empty).Replace("\r", " ").Replace("\a", " ");
                    table.Range.Font.Name = "Times New Roman";
                    table.Range.ParagraphFormat.SpaceAfter = 0f;
                    table.Range.ParagraphFormat.SpaceBefore = 0f;

                    if (text.IndexOf("BỘ GIÁO DỤC", StringComparison.OrdinalIgnoreCase) >= 0)
                        table.Range.Font.Size = 9.5f;
                    else if (text.IndexOf("Mã đề", StringComparison.OrdinalIgnoreCase) >= 0)
                        table.Range.Font.Size = 9.5f;
                    else
                        table.Range.Font.Size = 10f;

                    try { table.Rows.AllowBreakAcrossPages = 0; } catch { }
                }
            }
            catch { }
        }

        private void CenterPreservedVisuals()
        {
            try
            {
                dynamic document = _wordApplication.ActiveDocument;
                int count = (int)document.InlineShapes.Count;
                for (int i = 1; i <= count; i++)
                {
                    dynamic shape = document.InlineShapes.Item(i);
                    shape.Range.ParagraphFormat.Alignment = 1;
                }
            }
            catch { }
        }

        private void SetPrimaryFooter(string text)
        {
            try
            {
                dynamic document = _wordApplication.ActiveDocument;
                int sections = (int)document.Sections.Count;
                for (int i = 1; i <= sections; i++)
                {
                    dynamic section = document.Sections.Item(i);
                    dynamic footer = section.Footers.Item(1); // wdHeaderFooterPrimary
                    footer.Range.Text = text;
                    footer.Range.Font.Name = "Times New Roman";
                    footer.Range.Font.Size = 8.5f;
                    footer.Range.Font.Italic = 1;
                    footer.Range.ParagraphFormat.Alignment = 2;
                }
            }
            catch { }
        }

        private static bool IsExam(OcrDocument document)
        {
            return string.Equals(document?.document_type, "exam", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsVisualBlock(OcrBlock block)
        {
            string type = (block?.type ?? string.Empty).Trim().ToLowerInvariant();
            return type == "figure" || type == "table_image" || type == "unresolved";
        }

        private static OcrBlock CloneQuestion(OcrBlock source)
        {
            OcrBlock b = CloneBlock(source);
            b.content = source.content != null ? source.content.Select(CloneInline).ToList() : new List<OcrInline>();
            b.choices = source.choices != null
                ? source.choices.Where(x => x != null).Select(x => new OcrChoice
                {
                    label = x.label,
                    content = x.content != null ? x.content.Select(CloneInline).ToList() : new List<OcrInline>()
                }).ToList()
                : new List<OcrChoice>();
            return b;
        }

        private static OcrBlock CloneBlock(OcrBlock source)
        {
            if (source == null) return null;
            return new OcrBlock
            {
                type = source.type,
                text = source.text,
                label = source.label,
                number = source.number,
                latex = source.latex,
                word_linear = source.word_linear,
                confidence = source.confidence,
                display = source.display,
                bbox = source.bbox == null ? null : new OcrBoundingBox
                {
                    x = source.bbox.x,
                    y = source.bbox.y,
                    width = source.bbox.width,
                    height = source.bbox.height
                },
                content = source.content != null ? source.content.Select(CloneInline).ToList() : new List<OcrInline>(),
                choices = source.choices != null
                    ? source.choices.Where(x => x != null).Select(x => new OcrChoice
                    {
                        label = x.label,
                        content = x.content != null ? x.content.Select(CloneInline).ToList() : new List<OcrInline>()
                    }).ToList()
                    : new List<OcrChoice>()
            };
        }

        private static OcrInline CloneInline(OcrInline source)
        {
            if (source == null) return null;
            return new OcrInline
            {
                type = source.type,
                text = source.text,
                latex = source.latex,
                word_linear = source.word_linear,
                confidence = source.confidence
            };
        }
    }
}
