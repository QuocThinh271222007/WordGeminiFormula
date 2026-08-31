using System;
using System.Text.RegularExpressions;

namespace WordGeminiFormula.AddIn.Word
{
    /// <summary>
    /// Repairs only high-confidence OCR/LaTeX formatting defects. This class must
    /// never solve or algebraically transform the mathematical expression.
    /// </summary>
    public sealed class MathRepairService
    {
        public string RepairLatex(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return string.Empty;
            string s = value.Trim();

            s = Regex.Replace(s, @"\\mathbb\s*R\b", @"\\mathbb{R}");
            s = Regex.Replace(s, @"\\mathbb\s*N\b", @"\\mathbb{N}");
            s = Regex.Replace(s, @"\\mathbb\s*Z\b", @"\\mathbb{Z}");
            s = Regex.Replace(s, @"\\mathbb\s*Q\b", @"\\mathbb{Q}");
            s = Regex.Replace(s, @"\\mathbb\s*C\b", @"\\mathbb{C}");

            s = RepairUnaryCommandBody(s, "overrightarrow");
            s = RepairUnaryCommandBody(s, "overleftarrow");
            s = RepairUnaryCommandBody(s, "vec");
            s = RepairUnaryCommandBody(s, "overline");

            s = s.Replace("\\dfrac", "\\frac")
                 .Replace("\\tfrac", "\\frac")
                 .Replace("\\operatorname {", "\\operatorname{")
                 .Replace("\\mathrm {", "\\mathrm{");

            return s.Trim();
        }

        public string RepairWordLinear(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return string.Empty;
            string s = value.Trim();
            s = s.Replace("\\mathbb{R}", "ℝ")
                 .Replace("\\mathbb{N}", "ℕ")
                 .Replace("\\mathbb{Z}", "ℤ")
                 .Replace("\\mathbb{Q}", "ℚ")
                 .Replace("\\mathbb{C}", "ℂ");
            return Regex.Replace(s, @"\s+", " ").Trim();
        }

        private static string RepairUnaryCommandBody(string input, string command)
        {
            // Repair the common OCR form \overrightarrowAB -> \overrightarrow{AB}.
            // Only touch a short immediately-adjacent alphanumeric token.
            string pattern = @"\\" + Regex.Escape(command) + @"(?!\s*\{)([A-Za-z][A-Za-z0-9']{0,5})\b";
            return Regex.Replace(input, pattern, m => "\\" + command + "{" + m.Groups[1].Value + "}");
        }
    }
}
