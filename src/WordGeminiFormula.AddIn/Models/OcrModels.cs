using System.Collections.Generic;

namespace WordGeminiFormula.AddIn.Models
{
    public sealed class OcrDocument
    {
        public string document_type { get; set; } = "general";
        public List<OcrBlock> blocks { get; set; } = new List<OcrBlock>();
        public List<string> warnings { get; set; } = new List<string>();
    }

    public sealed class OcrBlock
    {
        // Supported V0.2 types include:
        // header_left, header_right, title, subtitle, meta, candidate_field,
        // code_box, section, question, text, formula, figure, table_image,
        // unresolved and footer.
        public string type { get; set; }
        public string text { get; set; }
        public string label { get; set; }
        public string number { get; set; }
        public string latex { get; set; }
        public string word_linear { get; set; }
        public double confidence { get; set; } = 1.0;
        public bool display { get; set; }
        public List<OcrInline> content { get; set; } = new List<OcrInline>();
        public List<OcrChoice> choices { get; set; } = new List<OcrChoice>();
        public OcrBoundingBox bbox { get; set; }
    }

    public sealed class OcrInline
    {
        public string type { get; set; }
        public string text { get; set; }
        public string latex { get; set; }
        public string word_linear { get; set; }
        public double confidence { get; set; } = 1.0;
    }

    public sealed class OcrChoice
    {
        public string label { get; set; }
        public List<OcrInline> content { get; set; } = new List<OcrInline>();
    }

    public sealed class OcrBoundingBox
    {
        // Normalized image coordinates in [0, 1].
        public double x { get; set; }
        public double y { get; set; }
        public double width { get; set; }
        public double height { get; set; }
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
