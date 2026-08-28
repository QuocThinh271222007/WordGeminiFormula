using System.Collections.Generic;

namespace WordGeminiFormula.AddIn.Models
{
    public sealed class OcrDocument
    {
        public List<OcrBlock> blocks { get; set; } = new List<OcrBlock>();
    }

    public sealed class OcrBlock
    {
        public string type { get; set; }
        public string text { get; set; }
        public string latex { get; set; }
        public string word_linear { get; set; }
    }

    internal sealed class GeminiGenerateResponse
    {
        public List<GeminiCandidate> candidates { get; set; }
        public GeminiError error { get; set; }
    }

    internal sealed class GeminiCandidate
    {
        public GeminiContent content { get; set; }
    }

    internal sealed class GeminiContent
    {
        public List<GeminiPart> parts { get; set; }
    }

    internal sealed class GeminiPart
    {
        public string text { get; set; }
    }

    internal sealed class GeminiError
    {
        public int code { get; set; }
        public string message { get; set; }
        public string status { get; set; }
    }
}
