using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Xsl;
using WordGeminiFormula.AddIn.Models;

namespace WordGeminiFormula.AddIn.Word
{
    /// <summary>
    /// V5 math-structure layer. V4 keeps responsibility for page/layout rendering;
    /// V5 replaces ambiguous Word-linear parsing with Presentation MathML -> OMML when
    /// Gemini supplies MathML, while retaining the old converter as a safe fallback.
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

        public void InsertOcrBlocks(OcrDocument document, bool autoNormalize, bool autoBeautify, string sourceImagePath, bool preserveDifficultRegions)
        {
            if (document == null) return;
            _v5FailureCount = 0;
            OcrDocument prepared = PrepareMathMlSentinels(document);
            _inner.InsertOcrBlocks(prepared, false, autoBeautify, sourceImagePath, preserveDifficultRegions);
        }

        public int BeautifyActiveDocument() => _inner.BeautifyActiveDocument();
        public void NormalizeSelection() => _inner.NormalizeSelection();

        public int NormalizeAllMarkedFormulas()
        {
            dynamic document = _wordApplication.ActiveDocument;
            _v5FailureCount = 0;
            int converted = 0;

            List<PendingBookmark> pending = SnapshotPendingBookmarks((object)document);
            pending.Sort(delegate(PendingBookmark a, PendingBookmark b) { return b.Start.CompareTo(a.Start); });

            foreach (PendingBookmark info in pending)
            {
                try
                {
                    if (!(bool)document.Bookmarks.Exists(info.Name)) continue;
                    dynamic bookmark = document.Bookmarks.Item(info.Name);
                    string marker = Convert.ToString(bookmark.Range.Text) ?? string.Empty;
                    string token = ExtractMathPayload(marker);
                    if (!token.StartsWith(SentinelPrefix, StringComparison.OrdinalIgnoreCase)) continue;

                    PendingMathMl payload;
                    if (!TryLoadPayload((object)document, token, out payload))
                    {
                        RestoreFallbackBookmark(document, bookmark, token, token);
                        _v5FailureCount++;
                        continue;
                    }

                    if (TryReplaceBookmarkWithMathMl(document, bookmark, payload.MathMl))
                    {
                        converted++;
                        DeletePayload((object)document, token);
                    }
                    else
                    {
                        RestoreFallbackBookmark(document, bookmark, token, payload.Latex);
                        DeletePayload((object)document, token);
                    }
                }
                catch
                {
                    _v5FailureCount++;
                }
            }

            // Restored LaTeX fallbacks and documents created before V5 still work.
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

                string blockToken;
                if (IsMathType(block.type) && TryCreateSentinel(sourceBlock.mathml, sourceBlock.latex, sourceBlock.word_linear, out blockToken))
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
                    block.choices.Add(new OcrChoice { label = sourceChoice.label, content = PrepareParts(sourceChoice.content) });
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
                string token;
                if (IsMathType(source.type) && TryCreateSentinel(source.mathml, source.latex, source.word_linear, out token))
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
            if (!_mathMlBridge.TryNormalizeMathMl(mathml, out normalizedMathMl)) return false;

            token = SentinelPrefix + Guid.NewGuid().ToString("N").Substring(0, 24);
            string fallback = !string.IsNullOrWhiteSpace(latex) ? latex.Trim() : (wordLinear ?? string.Empty).Trim();
            var payload = new PendingMathMl { MathMl = normalizedMathMl, Latex = fallback };
            SessionPayloads[token] = payload;
            PersistPayloadBestEffort((object)_wordApplication.ActiveDocument, token, payload);
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

        private static List<PendingBookmark> SnapshotPendingBookmarks(object documentObject)
        {
            dynamic document = documentObject;
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

        private static void PersistPayloadBestEffort(object documentObject, string token, PendingMathMl payload)
        {
            dynamic document = documentObject;
            try
            {
                UpsertVariable(document, VariableMathMlPrefix + token, Encode(payload.MathMl));
                UpsertVariable(document, VariableLatexPrefix + token, Encode(payload.Latex));
            }
            catch { }
        }

        private static bool TryLoadPayload(object documentObject, string token, out PendingMathMl payload)
        {
            if (SessionPayloads.TryGetValue(token, out payload)) return true;
            dynamic document = documentObject;
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

        private static void DeletePayload(object documentObject, string token)
        {
            PendingMathMl removed;
            SessionPayloads.TryRemove(token, out removed);
            dynamic document = documentObject;
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

        private static string Encode(string value) => Convert.ToBase64String(Encoding.UTF8.GetBytes(value ?? string.Empty));
        private static string Decode(string value) => string.IsNullOrWhiteSpace(value) ? string.Empty : Encoding.UTF8.GetString(Convert.FromBase64String(value));

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
                bbox = source.bbox == null ? null : new OcrBoundingBox { x = source.bbox.x, y = source.bbox.y, width = source.bbox.width, height = source.bbox.height },
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
                var settings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null };
                var doc = new XmlDocument { XmlResolver = null, PreserveWhitespace = true };
                using (var sr = new StringReader(value.Trim()))
                using (var reader = XmlReader.Create(sr, settings)) doc.Load(reader);

                XmlElement root = doc.DocumentElement;
                if (root == null || !string.Equals(root.LocalName, "math", StringComparison.OrdinalIgnoreCase)) return false;
                if (!string.Equals(root.NamespaceURI, MathMlNamespace, StringComparison.Ordinal)) return false;
                normalized = doc.OuterXml;
                return true;
            }
            catch { return false; }
        }

        public void InsertMathMl(dynamic range, string mathml)
        {
            string normalized;
            if (!TryNormalizeMathMl(mathml, out normalized)) throw new InvalidOperationException("MathML OCR không hợp lệ.");
            string xsl = GetStylesheetText();

            try
            {
                range.InsertXML(normalized, xsl);
                return;
            }
            catch { }

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
            if (!string.IsNullOrWhiteSpace(_stylesheetPath) && File.Exists(_stylesheetPath)) return _stylesheetPath;

            var candidates = new List<string>();
            try
            {
                string wordPath = Convert.ToString(_wordApplication.Path);
                if (!string.IsNullOrWhiteSpace(wordPath)) candidates.Add(Path.Combine(wordPath, "MML2OMML.XSL"));
            }
            catch { }

            string pf = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            string pfx86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);
            AddOfficeCandidates(candidates, pf);
            AddOfficeCandidates(candidates, pfx86);

            foreach (string candidate in candidates)
            {
                if (File.Exists(candidate))
                {
                    _stylesheetPath = candidate;
                    return _stylesheetPath;
                }
            }

            throw new FileNotFoundException("Không tìm thấy MML2OMML.XSL của Microsoft Word. Tool sẽ fallback sang bộ chuyển LaTeX cũ.");
        }

        private static void AddOfficeCandidates(List<string> candidates, string root)
        {
            if (string.IsNullOrWhiteSpace(root)) return;
            candidates.Add(Path.Combine(root, "Microsoft Office", "root", "Office16", "MML2OMML.XSL"));
            candidates.Add(Path.Combine(root, "Microsoft Office", "Office16", "MML2OMML.XSL"));
        }

        private static string TransformToOmml(string mathml, string stylesheetPath)
        {
            var transform = new XslCompiledTransform();
            transform.Load(stylesheetPath, new XsltSettings(false, false), new XmlUrlResolver());
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
