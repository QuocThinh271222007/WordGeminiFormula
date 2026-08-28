using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Web.Script.Serialization;
using WordGeminiFormula.AddIn.Models;

namespace WordGeminiFormula.AddIn.Gemini
{
    public sealed class GeminiClient
    {
        private const int MaxInlineImageBytes = 18 * 1024 * 1024;
        private static readonly HttpClient Http = CreateHttpClient();
        private readonly JavaScriptSerializer _json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };

        private static HttpClient CreateHttpClient()
        {
            return new HttpClient { Timeout = TimeSpan.FromSeconds(90) };
        }

        public OcrDocument OcrImage(string apiKey, string model, string imagePath)
        {
            ValidateConfiguration(apiKey, model);
            if (string.IsNullOrWhiteSpace(imagePath) || !File.Exists(imagePath))
                throw new FileNotFoundException("Không tìm thấy ảnh cần OCR.", imagePath);

            byte[] bytes = File.ReadAllBytes(imagePath);
            if (bytes.Length > MaxInlineImageBytes)
                throw new InvalidOperationException("Ảnh quá lớn cho chế độ inline. V1 giới hạn 18 MB để chừa dung lượng cho prompt và JSON request.");

            string mimeType = GetMimeType(imagePath);
            string prompt = OcrPrompt;

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
                    responseMimeType = "application/json"
                }
            };

            string output = Generate(apiKey, model, request);
            output = StripJsonFence(output);
            var doc = _json.Deserialize<OcrDocument>(output);
            if (doc?.blocks == null)
                throw new InvalidOperationException("Gemini trả về JSON không đúng schema OCR.");

            foreach (var block in doc.blocks)
            {
                block.type = (block.type ?? string.Empty).Trim().ToLowerInvariant();
                if (block.type == "formula")
                {
                    block.latex = (block.latex ?? string.Empty).Trim();
                    block.word_linear = (block.word_linear ?? string.Empty).Trim();
                }
            }

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
                    }
                };

                string output = Generate(apiKey, model, request).Trim();
                message = string.IsNullOrWhiteSpace(output) ? "API phản hồi rỗng." : "Kết nối thành công: " + output;
                return !string.IsNullOrWhiteSpace(output);
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return false;
            }
        }

        private string Generate(string apiKey, string model, object request)
        {
            string url = "https://generativelanguage.googleapis.com/v1beta/models/" + Uri.EscapeDataString(model) + ":generateContent";
            string body = _json.Serialize(request);

            using (var message = new HttpRequestMessage(HttpMethod.Post, url))
            {
                message.Headers.Add("x-goog-api-key", apiKey.Trim());
                message.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                message.Content = new StringContent(body, Encoding.UTF8, "application/json");

                using (HttpResponseMessage response = Http.SendAsync(message).GetAwaiter().GetResult())
                {
                    string jsonText = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                    GeminiGenerateResponse parsed = null;
                    try { parsed = _json.Deserialize<GeminiGenerateResponse>(jsonText); } catch { }

                    if (!response.IsSuccessStatusCode)
                    {
                        string detail = parsed?.error?.message;
                        if (string.IsNullOrWhiteSpace(detail)) detail = jsonText;
                        throw new InvalidOperationException($"Gemini API lỗi {(int)response.StatusCode}: {detail}");
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
You are an OCR engine for Vietnamese academic and mathematics documents.
Transcribe the image faithfully. Do NOT solve, simplify, infer missing values, or change mathematical meaning.
Preserve paragraph order and punctuation.

Return JSON only with this exact shape:
{
  ""blocks"": [
    { ""type"": ""text"", ""text"": ""..."" },
    { ""type"": ""formula"", ""latex"": ""..."", ""word_linear"": ""..."" }
  ]
}

Rules for formula blocks:
1. latex: valid LaTeX for the exact visible expression.
2. word_linear: equivalent Microsoft Word UnicodeMath linear input suitable for an equation region and Professional conversion.
3. Preserve superscripts, subscripts, fractions, radicals, integrals, sums, limits, vectors, matrices, cases, Greek symbols, intervals, set notation, and accents.
4. For Word UnicodeMath use forms such as a/(b+c), \sqrt(x), \sqrt(n&x), \matrix(a&b@c&d), and standard Math AutoCorrect commands where applicable.
5. Never merge normal prose into a formula block unless it is visibly part of the mathematical expression.
6. If a symbol is unclear, transcribe the most literal visible symbol; do not guess a different problem statement.
";
    }
}
