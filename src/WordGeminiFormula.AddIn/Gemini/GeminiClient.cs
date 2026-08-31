using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using WordGeminiFormula.AddIn.Models;

namespace WordGeminiFormula.AddIn.Gemini
{
    public sealed class GeminiClient
    {
        private const int MaxInlineImageBytes = 18 * 1024 * 1024;
        private static readonly TimeSpan OcrAttemptTimeout = TimeSpan.FromMinutes(5);
        private static readonly TimeSpan TestAttemptTimeout = TimeSpan.FromSeconds(30);
        private const int OcrMaxAttempts = 2;

        private static readonly HttpClient Http = CreateHttpClient();
        private readonly JavaScriptSerializer _json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };

        public string LastRawOcrJson { get; private set; }
        public string LastTransportDiagnostic { get; private set; }

        private static HttpClient CreateHttpClient()
        {
            // Per-request CancellationTokenSource instances below own timeout policy.
            // A global HttpClient timeout used to surface only the unhelpful
            // "A task was canceled" message after 120 seconds.
            return new HttpClient { Timeout = Timeout.InfiniteTimeSpan };
        }

        public OcrDocument OcrImage(string apiKey, string model, string imagePath, string documentPreset = null)
        {
            ValidateConfiguration(apiKey, model);
            if (string.IsNullOrWhiteSpace(imagePath) || !File.Exists(imagePath))
                throw new FileNotFoundException("Không tìm thấy ảnh cần OCR.", imagePath);

            byte[] bytes = File.ReadAllBytes(imagePath);
            if (bytes.Length > MaxInlineImageBytes)
                throw new InvalidOperationException("Ảnh quá lớn cho chế độ inline. Giới hạn hiện tại là 18 MB.");

            string mimeType = GetMimeType(imagePath);
            string prompt = BuildPrompt(documentPreset);
            var request = new
            {
                contents = new[]
                {
                    new
                    {
                        role = "user",
                        parts = new object[]
                        {
                            new { text = prompt },
                            new
                            {
                                inlineData = new
                                {
                                    mimeType,
                                    data = Convert.ToBase64String(bytes)
                                }
                            }
                        }
                    }
                },
                generationConfig = new
                {
                    responseMimeType = "application/json",
                    // OCR/layout extraction is latency-sensitive and does not need the
                    // default medium reasoning effort of Gemini 3.7 Flash.
                    thinkingConfig = new { thinkingLevel = "low" },
                    maxOutputTokens = 16384
                }
            };

            string output = StripJsonFence(Generate(
                apiKey,
                model,
                request,
                OcrAttemptTimeout,
                OcrMaxAttempts,
                "Gemini OCR"));

            LastRawOcrJson = output;
            OcrDocument doc;
            try
            {
                doc = _json.Deserialize<OcrDocument>(output);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(
                    "Gemini đã phản hồi nhưng JSON OCR không hợp lệ. Raw response đã được giữ để chẩn đoán.",
                    ex);
            }

            if (doc?.blocks == null)
                throw new InvalidOperationException("Gemini trả về JSON không đúng schema OCR V0.2.");

            NormalizeDocument(doc);
            return doc;
        }

        public bool TestConnection(string apiKey, string model, out string message)
        {
            try
            {
                ValidateConfiguration(apiKey, model);
                var request = new
                {
                    contents = new[]
                    {
                        new
                        {
                            role = "user",
                            parts = new[] { new { text = "Reply with exactly: OK" } }
                        }
                    },
                    generationConfig = new
                    {
                        thinkingConfig = new { thinkingLevel = "low" },
                        maxOutputTokens = 32
                    }
                };

                string output = Generate(
                    apiKey,
                    model,
                    request,
                    TestAttemptTimeout,
                    1,
                    "Kiểm tra kết nối Gemini").Trim();

                message = string.IsNullOrWhiteSpace(output) ? "API phản hồi rỗng." : "Kết nối thành công: " + output;
                return !string.IsNullOrWhiteSpace(output);
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return false;
            }
        }

        private static string BuildPrompt(string documentPreset)
        {
            string preset = string.IsNullOrWhiteSpace(documentPreset)
                ? "auto"
                : documentPreset.Trim().ToLowerInvariant();
            return OcrPrompt + "\n\nUSER DOCUMENT PRESET: " + preset +
                   "\nIf preset is exam, prioritize the exam-specific block rules. If preset is general, preserve layout but do not force exam semantics. If preset is auto, infer the document type from the image.";
        }

        private static void NormalizeDocument(OcrDocument doc)
        {
            doc.document_type = (doc.document_type ?? "general").Trim().ToLowerInvariant();
            if (doc.warnings == null) doc.warnings = new List<string>();

            foreach (var block in doc.blocks.Where(b => b != null))
            {
                block.type = (block.type ?? "text").Trim().ToLowerInvariant();
                block.text = block.text?.Trim();
                block.label = block.label?.Trim();
                block.number = block.number?.Trim();
                block.latex = block.latex?.Trim();
                block.word_linear = block.word_linear?.Trim();
                if (block.content == null) block.content = new List<OcrInline>();
                if (block.choices == null) block.choices = new List<OcrChoice>();

                foreach (var part in block.content.Where(p => p != null))
                {
                    part.type = (part.type ?? "text").Trim().ToLowerInvariant();
                    part.text = part.text ?? string.Empty;
                    part.latex = part.latex?.Trim();
                    part.word_linear = part.word_linear?.Trim();
                }

                foreach (var choice in block.choices.Where(c => c != null))
                {
                    choice.label = choice.label?.Trim();
                    if (choice.content == null) choice.content = new List<OcrInline>();
                    foreach (var part in choice.content.Where(p => p != null))
                    {
                        part.type = (part.type ?? "text").Trim().ToLowerInvariant();
                        part.text = part.text ?? string.Empty;
                        part.latex = part.latex?.Trim();
                        part.word_linear = part.word_linear?.Trim();
                    }
                }
            }
        }

        private string Generate(
            string apiKey,
            string model,
            object request,
            TimeSpan attemptTimeout,
            int maxAttempts,
            string operationName)
        {
            string url = "https://generativelanguage.googleapis.com/v1beta/models/" + Uri.EscapeDataString(model) + ":generateContent";
            string body = _json.Serialize(request);
            Exception lastTransportError = null;
            maxAttempts = Math.Max(1, maxAttempts);

            for (int attempt = 1; attempt <= maxAttempts; attempt++)
            {
                var stopwatch = Stopwatch.StartNew();
                try
                {
                    using (var cts = new CancellationTokenSource(attemptTimeout))
                    using (var message = new HttpRequestMessage(HttpMethod.Post, url))
                    {
                        message.Headers.Add("x-goog-api-key", apiKey.Trim());
                        message.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                        message.Headers.ExpectContinue = false;
                        message.Content = new StringContent(body, Encoding.UTF8, "application/json");

                        using (HttpResponseMessage response = Http.SendAsync(
                            message,
                            HttpCompletionOption.ResponseContentRead,
                            cts.Token).GetAwaiter().GetResult())
                        {
                            stopwatch.Stop();
                            LastTransportDiagnostic = string.Format(
                                "{0}: attempt {1}/{2}, HTTP {3}, elapsed {4:0.0}s",
                                operationName,
                                attempt,
                                maxAttempts,
                                (int)response.StatusCode,
                                stopwatch.Elapsed.TotalSeconds);

                            string jsonText = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                            GeminiGenerateResponse parsed = null;
                            try { parsed = _json.Deserialize<GeminiGenerateResponse>(jsonText); } catch { }

                            if (!response.IsSuccessStatusCode)
                            {
                                string detail = parsed?.error?.message;
                                if (string.IsNullOrWhiteSpace(detail)) detail = jsonText;
                                throw new InvalidOperationException(
                                    string.Format("Gemini API lỗi {0}: {1}", (int)response.StatusCode, detail));
                            }

                            if (parsed?.error != null)
                                throw new InvalidOperationException("Gemini API: " + parsed.error.message);

                            string text = parsed?.candidates?
                                .SelectMany(c => c.content?.parts ?? new List<GeminiPart>())
                                .Select(p => p.text)
                                .FirstOrDefault(t => !string.IsNullOrWhiteSpace(t));

                            if (string.IsNullOrWhiteSpace(text))
                                throw new InvalidOperationException("Gemini không trả về nội dung văn bản.");

                            return text;
                        }
                    }
                }
                catch (TaskCanceledException ex)
                {
                    stopwatch.Stop();
                    lastTransportError = ex;
                    LastTransportDiagnostic = string.Format(
                        "{0}: timeout attempt {1}/{2} after {3:0.0}s",
                        operationName,
                        attempt,
                        maxAttempts,
                        stopwatch.Elapsed.TotalSeconds);

                    if (attempt < maxAttempts)
                    {
                        Thread.Sleep(1500 * attempt);
                        continue;
                    }

                    throw new TimeoutException(
                        string.Format(
                            "{0} quá thời gian chờ ({1:0} giây/lần, {2} lần thử). Hãy kiểm tra mạng hoặc thử lại; model hiện dùng thinking=low để giảm độ trễ.",
                            operationName,
                            attemptTimeout.TotalSeconds,
                            maxAttempts),
                        ex);
                }
                catch (HttpRequestException ex)
                {
                    stopwatch.Stop();
                    lastTransportError = ex;
                    LastTransportDiagnostic = string.Format(
                        "{0}: network failure attempt {1}/{2} after {3:0.0}s: {4}",
                        operationName,
                        attempt,
                        maxAttempts,
                        stopwatch.Elapsed.TotalSeconds,
                        ex.Message);

                    if (attempt < maxAttempts)
                    {
                        Thread.Sleep(1500 * attempt);
                        continue;
                    }

                    throw new InvalidOperationException(
                        operationName + " không kết nối được tới Gemini API sau " + maxAttempts + " lần thử: " + ex.Message,
                        ex);
                }
            }

            throw new InvalidOperationException(
                operationName + " thất bại do lỗi truyền tải không xác định.",
                lastTransportError);
        }

        private static void ValidateConfiguration(string apiKey, string model)
        {
            if (string.IsNullOrWhiteSpace(apiKey))
                throw new InvalidOperationException("Chưa cấu hình Gemini API key. Mở Settings trên Ribbon để nhập key.");
            if (string.IsNullOrWhiteSpace(model))
                throw new InvalidOperationException("Chưa chọn Gemini model.");
        }

        private static string GetMimeType(string path)
        {
            switch ((Path.GetExtension(path) ?? string.Empty).ToLowerInvariant())
            {
                case ".png": return "image/png";
                case ".webp": return "image/webp";
                case ".gif": return "image/gif";
                case ".bmp": return "image/bmp";
                case ".jpg":
                case ".jpeg":
                default: return "image/jpeg";
            }
        }

        private static string StripJsonFence(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return value;
            string trimmed = value.Trim();
            if (!trimmed.StartsWith("```", StringComparison.Ordinal)) return trimmed;

            int firstNewline = trimmed.IndexOf('\n');
            int lastFence = trimmed.LastIndexOf("```", StringComparison.Ordinal);
            if (firstNewline >= 0 && lastFence > firstNewline)
                return trimmed.Substring(firstNewline + 1, lastFence - firstNewline - 1).Trim();
            return trimmed;
        }

        private const string OcrPrompt = @"
You are a high-precision OCR + document-layout engine for Vietnamese academic documents, especially mathematics exams.
Your output will be rendered into an editable Microsoft Word document.

CORE REQUIREMENTS
- Transcribe faithfully. Do NOT solve questions, simplify expressions, change answers, or invent missing content.
- Preserve reading order, punctuation and visible document hierarchy.
- Recognize layout semantics instead of emitting every visual line as an unrelated paragraph.
- Keep mathematical expressions as math fragments, not flattened prose.
- Preserve spaces at the boundaries of inline text fragments so prose and formulas do not run together.
- Use standard mathematical notation when the image is visually unambiguous: f(x), g(x), u_n, u_1, P(A), vectors, coordinates, intervals, derivatives, integrals, matrices, cases, logarithm bases, etc.
- Every math fragment must be complete and self-contained. Do not split one visible formula into separate math fragments.
- Do not emit square-bracket MATH control markers in any text field; those markers belong to the Word renderer.
- If a symbol is genuinely unclear, keep the most literal reading and lower confidence; do not guess a different problem.

DOCUMENT TYPE
Set document_type to ""exam"" for an exam/test page, otherwise ""general"".

RETURN JSON ONLY. Use this schema:
{
  ""document_type"": ""exam"",
  ""warnings"": [""optional warning""],
  ""blocks"": [
    { ""type"": ""header_left"", ""text"": ""multi-line text"" },
    { ""type"": ""header_right"", ""text"": ""multi-line text"" },
    { ""type"": ""title"", ""text"": ""..."" },
    { ""type"": ""subtitle"", ""text"": ""..."" },
    { ""type"": ""meta"", ""text"": ""..."" },
    { ""type"": ""candidate_field"", ""label"": ""Họ, tên thí sinh"", ""text"": """" },
    { ""type"": ""code_box"", ""label"": ""Mã đề"", ""text"": ""0101"" },
    { ""type"": ""section"", ""text"": ""PHẦN I: ..."" },
    {
      ""type"": ""question"",
      ""number"": ""1"",
      ""content"": [
        { ""type"": ""text"", ""text"": ""Cho cấp số cộng "" },
        { ""type"": ""math"", ""latex"": ""u_n"", ""word_linear"": ""u_n"", ""confidence"": 1.0 },
        { ""type"": ""text"", ""text"": "" có "" }
      ],
      ""choices"": [
        { ""label"": ""A"", ""content"": [{ ""type"": ""text"", ""text"": ""4."" }] }
      ]
    },
    { ""type"": ""formula"", ""latex"": ""..."", ""word_linear"": ""..."", ""display"": true },
    {
      ""type"": ""figure"",
      ""text"": ""short description only"",
      ""bbox"": { ""x"": 0.1, ""y"": 0.2, ""width"": 0.4, ""height"": 0.3 },
      ""confidence"": 0.7
    },
    {
      ""type"": ""table_image"",
      ""text"": ""variation table / diagram that should be preserved visually"",
      ""bbox"": { ""x"": 0.1, ""y"": 0.2, ""width"": 0.7, ""height"": 0.2 },
      ""confidence"": 0.8
    },
    { ""type"": ""footer"", ""text"": ""..."" }
  ]
}

BLOCK RULES
1. For an exam header with two visual columns, emit one header_left and one header_right block. Do not split each line into separate generic text blocks.
2. Emit each multiple-choice problem as ONE question block with inline content plus its A/B/C/D choices.
3. Inline formulas inside sentences belong in question.content/choice.content as type=math.
4. A standalone displayed formula may use type=formula.
5. candidate_field is for dotted candidate fields such as name/student number.
6. code_box is for boxed exam code fields.
7. section is for headings such as PHẦN I.
8. footer is for page number / exam-code footer.

MATH RULES
- latex must be syntactically valid LaTeX for the exact visible expression.
- word_linear must be equivalent Microsoft Word UnicodeMath/Math AutoCorrect linear input.
- Preserve superscripts, subscripts, fractions, radicals, integrals, sums, limits, vectors, matrices, systems/cases, Greek symbols, intervals, sets and accents.
- Prefer explicit grouping. Examples: f(x), g'(x)=x^2, \\mathbb{R}, \\overrightarrow{AB}, A(1;5;1), \\log_3(x-1), \\begin{cases}...\\end{cases}.
- Do not remove parentheses from function notation or coordinates.

DIFFICULT VISUAL REGIONS
- Do NOT flatten geometry diagrams, graphs, variation tables, or visually structured tables into unreliable prose.
- Emit figure/table_image with a tight bbox using normalized x/y/width/height in [0,1], so Word can preserve that region as an image.
- Use type=unresolved with bbox when a region cannot be transcribed reliably. Include a short text reason.
";
    }
}