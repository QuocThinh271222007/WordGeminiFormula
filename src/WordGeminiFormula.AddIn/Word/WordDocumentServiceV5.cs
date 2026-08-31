using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Xsl;
using WordGeminiFormula.AddIn.Models;

namespace WordGeminiFormula.AddIn.Word
{
    /// <summary>
    /// V5 math-structure layer.
    ///
    /// V4 remains responsible for document layout, images, tables and compact exam styling.
    /// V5 changes only the equation pipeline: when Gemini provides Presentation MathML,
    /// the formula is rendered through Word's own MML2OMML.XSL transform instead of
    /// asking OMath.BuildUp to guess structure from linear text.
    ///
    /// This fixes classes of errors that cannot be made reliable with regex rules alone,
    /// such as decimal commas, P(A), cases/systems, matrices, nested radicals/fractions,
    /// scripts, accents and other structured mathematical constructs.
    /// </summary>
    public sealed class WordDocumentServiceV5
    {
        private const string SentinelPrefix = "WGFML_";
        private const string BookmarkPrefix = "WGF_MATH_";
        private const string MathStart = "[[MATH]]";
        private const string MathEnd = "[[/MATH]]";
        private const string VariableMathMlPrefix = "WGF_MML_";
        private const string VariableLatexPrefix = "WGF_TEX_";
        private const int WdYellow = 7;

        private readonly dynamic _wordApplication;
        private readonly WordDocumentServiceV4 _inner;
        private readonly MathMlOfficeBridge _mathMlBridge;

        // New service instances are created for Ribbon callbacks. Keep payloads for the
        // current Word process, and also persist them into Document.Variables best-effort.
        private static readonly ConcurrentDictionary<string, PendingMathMl> SessionPayloads =
            new ConcurrentDictionary<string, PendingMathMl>(StringComparer.OrdinalIgnoreCase);

        private int _v5FailureCount;

        public int LastNormalizationFailureCount => _v5FailureCount + _inner.LastNormalizationFailureCount;

        public WordDocumentServiceV5(object wordApplication)
        {
            _wordApplication = wordApplication ?? throw new ArgumentNullException(nameof(wordApplication));
            _inner = new WordDocumentServiceV4(wordApplication);
            _mathMlBridge = new MathMlOfficeBridge(wordApplication);
        }

        public void InsertOcrBlocks(
            OcrDocument document,
            bool autoNormalize,
            bool autoBeautify,
            string sourceImagePath,
            bool preserveDifficultRegions)
        {
            if (document == null) return;

            _v5FailureCount = 0;
            OcrDocument prepared = PrepareMathMlSentinels(document);

            // V4/V3 still perform the safe bookmark-only first pass. MathML is never
            // inserted while Word is still typing the page.
            _inner.InsertOcrBlocks(
                prepared,
                false,
                autoBeautify,
                sourceImagePath,
                preserveDifficultRegions);
        }

        public int BeautifyActiveDocument() => _inner.BeautifyActiveDocument();

        public void NormalizeSelection()
        {
            // Manual selection is intentionally kept on the legacy LaTeX/UnicodeMath path,
            // because a selected string has no associated Gemini MathML payload.
            _inner.NormalizeSelection();
        }

        public int NormalizeAllMarkedFormulas()
        {
            dynamic document = _wordApplication.ActiveDocument;
            _v5FailureCount = 0;
            int converted = 0;

            var pending = SnapshotPendingBookmarks(document);
            foreach (PendingBookmark info in pending.OrderByDescending(x => x.Start))
            {
                try
                {
                    if (!(bool)document.Bookmarks.Exists(info.Name)) continue;
                    dynamic bookmark = document.Bookmarks.Item(info.Name);
                    string marker = Convert.ToString(bookmark.Range.Text) ?? string.Empty;
                    string token = ExtractMathPayload(marker);
                    if (!token.StartsWith(SentinelPrefix, StringComparison.OrdinalIgnoreCase))
                        continue;

                    PendingMathMl payload;
                    if (!TryLoadPayload(document, token, out payload))
                    {
                        RestoreFallbackBookmark(document, bookmark, token, token);
                        _v5FailureCount++;
                        continue;
                    }

                    if (TryReplaceBookmarkWithMathMl(document, bookmark, payload.MathMl))
                    {
                        converted++;
                        DeletePayload(document, token);
                    }
                    else
                    {
                        // If native MathML insertion is unavailable on a particular Office
                        // installation, restore original LaTeX and let V4/V3 use its old
                        // converter rather than losing the formula.
                        RestoreFallbackBookmark(document, bookmark, token, payload.Latex);
                        DeletePayload(document, token);
                    }
                }
                catch
                {
                    _v5FailureCount++;
                }
            }

            // Process restored LaTeX fallbacks and old documents made before V5.
            converted += _inner.NormalizeAllMarkedFormulas();
            return converted;
        }

        private OcrDocument PrepareMathMlSentinels(OcrDocument source)
        {
            var result = new OcrDocument
            {
                document_type = source.document_type,
                warnings = source.warnings != null ? new List<string>(source.warnings) : new List<string>(),
                blocks = new List<OcrBlock>()
            };

            foreach (OcrBlock sourceBlock in source.blocks ?? new List<OcrBlock>())
            {
                if (sourceBlock == null) continue;
                OcrBlock block = CloneBlock(sourceBlock);

                if (IsMathType(block.type) && TryCreateSentinel(sourceBlock.mathml, sourceBlock.latex, sourceBlock.word_linear, out string blockToken))
                {
                    block.latex = blockToken;
                    block.word_linear = null;
                    block.mathml = null;
                }

                block.content = PrepareParts(sourceBlock.content);
                block.choices = new List<OcrChoice>();
                foreach (OcrChoice sourceChoice in sourceBlock.choices ?? new List<OcrChoice>())
                {
                    if (sourceChoice == null) continue;
                    block.choices.Add(new OcrChoice
                    {
                        label = sourceChoice.label,
                        content = PrepareParts(sourceChoice.content)
                    });
                }

                result.blocks.Add(block);
            }

            return result;
        }

        private List<OcrInline> PrepareParts(IEnumerable<OcrInline> parts)
        {
            var result = new List<OcrInline>();
            if (parts == null) return result;

            foreach (OcrInline source in parts)
            {
                if (source == null) continue;
                OcrInline part = CloneInline(source);
                if (IsMathType(source.type) && TryCreateSentinel(source.mathml, source.latex, source.word_linear, out string token))
                {
                    part.latex = token;
                    part.word_linear = null;
                    part.mathml = null;
                }
                result.Add(part);
            }
            return result;
        }

        private bool TryCreateSentinel(string mathml, string latex, string wordLinear, out string token)
        {
            token = null;
            string normalizedMathMl;
            if (!_mathMlBridge.TryNormalizeMathMl(mathml, out normalizedMathMl))
                return false;

            token = SentinelPrefix + Guid.NewGuid().ToString("N").Substring(0, 24);
            string fallback = !string.IsNullOrWhiteSpace(latex) ? latex.Trim() : (wordLinear ?? string.Empty).Trim();
            var payload = new PendingMathMl { MathMl = normalizedMathMl, Latex = fallback };
            SessionPayloads[token] = payload;
            PersistPayloadBestEffort(_wordApplication.ActiveDocument, token, payload);
            return true;
        }

        private bool TryReplaceBookmarkWithMathMl(dynamic document, dynamic bookmark, string mathml)
        {
            string marker = Convert.ToString(bookmark.Range.Text) ?? string.Empty;
            int start = (int)bookmark.Range.Start;
            int end = (int)bookmark.Range.End;
            string name = Convert.ToString(bookmark.Name) ?? string.Empty;

            try
            {
                bookmark.Delete();
                dynamic exact = document.Range(start, end);
                _mathMlBridge.InsertMathMl(exact, mathml);
                return true;
            }
            catch
            {
                // Recreate a bookmark over the original marker when possible so fallback
                // handling can safely replace it with LaTeX.
                try
                {
                    dynamic original = document.Range(start, Math.Min((int)document.Content.End, start + marker.Length));
                    if (!string.IsNullOrWhiteSpace(name)) document.Bookmarks.Add(name, original);
                }
                catch { }
                return false;
            }
        }

        private static void RestoreFallbackBookmark(dynamic document, dynamic bookmark, string token, string latex)
        {
            string marker = Convert.ToString(bookmark.Range.Text) ?? string.Empty;
            int start = (int)bookmark.Range.Start;
            string name = Convert.ToString(bookmark.Name) ?? string.Empty;
            string fallback = string.IsNullOrWhiteSpace(latex) ? token : latex.Trim();
            string replacement = MathStart + fallback + MathEnd;

            try { bookmark.Delete(); } catch { }
            dynamic exact = document.Range(start, Math.Min((int)document.Content.End, start + marker.Length));
            exact.Text = replacement;
            dynamic restored = document.Range(start, start + replacement.Length);
            if (!string.IsNullOrWhiteSpace(name)) document.Bookmarks.Add(name, restored);
            try { restored.HighlightColorIndex = WdYellow; } catch { }
        }

        private static List<PendingBookmark> SnapshotPendingBookmarks(dynamic document)
        {
            var result = new List<PendingBookmark>();
            dynamic bookmarks = document.Bookmarks;
            int count = (int)bookmarks.Count;
            for (int i = 1; i <= count; i++)
            {
                dynamic bookmark = bookmarks.Item(i);
                string name = Convert.ToString(bookmark.Name) ?? string.Empty;
                if (!name.StartsWith(BookmarkPrefix, StringComparison.OrdinalIgnoreCase)) continue;
                result.Add(new PendingBookmark { Name = name, Start = (int)bookmark.Range.Start });
            }
            return result;
        }

        private static string ExtractMathPayload(string value)
        {
            string raw = (value ?? string.Empty).Trim();
            if (raw.StartsWith(MathStart, StringComparison.Ordinal) && raw.EndsWith(MathEnd, StringComparison.Ordinal))
                return raw.Substring(MathStart.Length, raw.Length - MathStart.Length - MathEnd.Length).Trim();
            return raw.Replace(MathStart, string.Empty).Replace(MathEnd, string.Empty).Trim();
        }

        private static bool IsMathType(string type)
        {
            return string.Equals(type, "math", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(type, "formula", StringComparison.OrdinalIgnoreCase);
        }

        private static void PersistPayloadBestEffort(dynamic document, string token, PendingMathMl payload)
        {
            try
            {
                UpsertVariable(document, VariableMathMlPrefix + token, Encode(payload.MathMl));
                UpsertVariable(document, VariableLatexPrefix + token, Encode(payload.Latex));
            }
            catch
            {
                // SessionPayloads still keeps the current Word session functional.
            }
        }

        private static bool TryLoadPayload(dynamic document, string token, out PendingMathMl payload)
        {
            if (SessionPayloads.TryGetValue(token, out payload)) return true;

            try
            {
                string mathml = Decode(Convert.ToString(document.Variables.Item(VariableMathMlPrefix + token).Value));
                string latex = Decode(Convert.ToString(document.Variables.Item(VariableLatexPrefix + token).Value));
                if (string.IsNullOrWhiteSpace(mathml)) return false;
                payload = new PendingMathMl { MathMl = mathml, Latex = latex };
                SessionPayloads[token] = payload;
                return true;
            }
            catch
            {
                payload = null;
                return false;
            }
        }

        private static void DeletePayload(dynamic document, string token)
        {
            SessionPayloads.TryRemove(token, out _);
            TryDeleteVariable(document, VariableMathMlPrefix + token);
            TryDeleteVariable(document, VariableLatexPrefix + token);
        }

        private static void UpsertVariable(dynamic document, string name, string value)
        {
            TryDeleteVariable(document, name);
            document.Variables.Add(name, value ?? string.Empty);
        }

        private static void TryDeleteVariable(dynamic document, string name)
        {
            try { document.Variables.Item(name).Delete(); } catch { }
        }

        private static string Encode(string value)
        {
            return Convert.ToBase64String(Encoding.UTF8.GetBytes(value ?? string.Empty));
        }

        private static string Decode(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return string.Empty;
            return Encoding.UTF8.GetString(Convert.FromBase64String(value));
        }

        private static OcrBlock CloneBlock(OcrBlock source)
        {
            return new OcrBlock
            {
                type = source.type,
                text = source.text,
                label = source.label,
                number = source.number,
                latex = source.latex,
                word_linear = source.word_linear,
                mathml = source.mathml,
                confidence = source.confidence,
                display = source.display,
                bbox = source.bbox == null ? null : new OcrBoundingBox
                {
                    x = source.bbox.x,
                    y = source.bbox.y,
                    width = source.bbox.width,
                    height = source.bbox.height
                },
                content = new List<OcrInline>(),
                choices = new List<OcrChoice>()
            };
        }

        private static OcrInline CloneInline(OcrInline source)
        {
            return new OcrInline
            {
                type = source.type,
                text = source.text,
                latex = source.latex,
                word_linear = source.word_linear,
                mathml = source.mathml,
                confidence = source.confidence
            };
        }

        private sealed class PendingBookmark
        {
            public string Name { get; set; }
            public int Start { get; set; }
        }

        private sealed class PendingMathMl
        {
            public string MathMl { get; set; }
            public string Latex { get; set; }
        }
    }

    /// <summary>
    /// Imports Presentation MathML into native Office Math using Microsoft's own
    /// MML2OMML.XSL stylesheet shipped with desktop Word. Microsoft 365 supports
    /// Presentation MathML import; using the Office transform preserves mathematical
    /// structure instead of relying on ambiguous linear-parser heuristics.
    /// </summary>
    internal sealed class MathMlOfficeBridge
    {
        private const string MathMlNamespace = "http://www.w3.org/1998/Math/MathML";
        private readonly dynamic _wordApplication;
        private string _stylesheetText;
        private string _stylesheetPath;

        public MathMlOfficeBridge(object wordApplication)
        {
            _wordApplication = wordApplication ?? throw new ArgumentNullException(nameof(wordApplication));
        }

        public bool TryNormalizeMathMl(string value, out string normalized)
        {
            normalized = null;
            if (string.IsNullOrWhiteSpace(value)) return false;

            try
            {
                var settings = new XmlReaderSettings
                {
                    DtdProcessing = DtdProcessing.Prohibit,
                    XmlResolver = null
                };
                var doc = new XmlDocument { XmlResolver = null, PreserveWhitespace = true };
                using (var sr = new StringReader(value.Trim()))
                using (var reader = XmlReader.Create(sr, settings))
                    doc.Load(reader);

                XmlElement root = doc.DocumentElement;
                if (root == null || !string.Equals(root.LocalName, "math", StringComparison.OrdinalIgnoreCase))
                    return false;
                if (!string.Equals(root.NamespaceURI, MathMlNamespace, StringComparison.Ordinal))
                    return false;

                normalized = doc.OuterXml;
                return true;
            }
            catch
            {
                return false;
            }
        }

        public void InsertMathMl(dynamic range, string mathml)
        {
            string normalized;
            if (!TryNormalizeMathMl(mathml, out normalized))
                throw new InvalidOperationException("MathML OCR không hợp lệ.");

            string xsl = GetStylesheetText();

            // Preferred path: let Word apply the same transform it uses for MathML import.
            try
            {
                range.InsertXML(normalized, xsl);
                return;
            }
            catch
            {
                // Some Word builds are stricter about InsertXML's Transform argument.
                // Transform with .NET first, then insert the resulting OMML fragment.
            }

            string omml = TransformToOmml(normalized, GetStylesheetPath());
            range.InsertXML(omml);
        }

        private string GetStylesheetText()
        {
            if (!string.IsNullOrWhiteSpace(_stylesheetText)) return _stylesheetText;
            _stylesheetText = File.ReadAllText(GetStylesheetPath(), Encoding.UTF8);
            return _stylesheetText;
        }

        private string GetStylesheetPath()
        {
            if (!string.IsNullOrWhiteSpace(_stylesheetPath) && File.Exists(_stylesheetPath))
                return _stylesheetPath;

            var candidates = new List<string>();
            try
            {
                string wordPath = Convert.ToString(_wordApplication.Path);
                if (!string.IsNullOrWhiteSpace(wordPath))
                    candidates.Add(Path.Combine(wordPath, "MML2OMML.XSL"));
            }
            catch { }

            string programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            string programFilesX86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);
            foreach (string root in new[] { programFiles, programFilesX86 }.Where(x => !string.IsNullOrWhiteSpace(x)))
            {
                candidates.Add(Path.Combine(root, "Microsoft Office", "root", "Office16", "MML2OMML.XSL"));
                candidates.Add(Path.Combine(root, "Microsoft Office", "Office16", "MML2OMML.XSL"));
            }

            _stylesheetPath = candidates.FirstOrDefault(File.Exists);
            if (string.IsNullOrWhiteSpace(_stylesheetPath))
                throw new FileNotFoundException(
                    "Không tìm thấy MML2OMML.XSL của Microsoft Word. Tool sẽ fallback sang bộ chuyển LaTeX cũ.");
            return _stylesheetPath;
        }

        private static string TransformToOmml(string mathml, string stylesheetPath)
        {
            var transform = new XslCompiledTransform();
            var xsltSettings = new XsltSettings(false, false);
            var resolver = new XmlUrlResolver();
            transform.Load(stylesheetPath, xsltSettings, resolver);

            var inputSettings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null };
            using (var inputText = new StringReader(mathml))
            using (var input = XmlReader.Create(inputText, inputSettings))
            using (var output = new StringWriter())
            using (var writer = XmlWriter.Create(output, transform.OutputSettings))
            {
                transform.Transform(input, writer);
                writer.Flush();
                return output.ToString();
            }
        }
    }
}
