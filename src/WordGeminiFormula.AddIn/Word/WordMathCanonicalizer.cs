using System;
using System.Text.RegularExpressions;

namespace WordGeminiFormula.AddIn.Word
{
    /// <summary>
    /// Converts OCR LaTeX into a conservative Word/UnicodeMath-friendly representation
    /// before the existing converter and OMath.BuildUp pass. The goal is to remove
    /// Math-AutoCorrect-only tokens that OMath.BuildUp may leave as literal text.
    /// This class never solves or algebraically changes an expression.
    /// </summary>
    public sealed class WordMathCanonicalizer
    {
        public string NormalizeForWord(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return string.Empty;
            string s = value.Trim();

            s = s.Replace("\\left", string.Empty)
                 .Replace("\\right", string.Empty)
                 .Replace("\\displaystyle", string.Empty)
                 .Replace("\\,", " ")
                 .Replace("\\;", " ")
                 .Replace("\\!", string.Empty);

            // Sets and common relation/special symbols: use actual Unicode characters so
            // OMath.BuildUp does not depend on Math AutoCorrect expansion.
            s = s.Replace("\\mathbb{R}", "ℝ")
                 .Replace("\\mathbb{N}", "ℕ")
                 .Replace("\\mathbb{Z}", "ℤ")
                 .Replace("\\mathbb{Q}", "ℚ")
                 .Replace("\\mathbb{C}", "ℂ")
                 .Replace("\\doubleR", "ℝ")
                 .Replace("\\neq", "≠")
                 .Replace("\\ne", "≠")
                 .Replace("\\leq", "≤")
                 .Replace("\\geq", "≥")
                 .Replace("\\le", "≤")
                 .Replace("\\ge", "≥")
                 .Replace("\\infty", "∞")
                 .Replace("\\rightarrow", "→")
                 .Replace("\\to", "→");

            // N-ary and named functions. Word builds actual Unicode n-ary symbols much
            // more reliably through OMath.BuildUp than backslash AutoCorrect tokens.
            s = s.Replace("\\iiint", "∭")
                 .Replace("\\iint", "∬")
                 .Replace("\\int", "∫")
                 .Replace("\\sum", "∑")
                 .Replace("\\prod", "∏")
                 .Replace("\\sin", "sin")
                 .Replace("\\cos", "cos")
                 .Replace("\\tan", "tan")
                 .Replace("\\cot", "cot")
                 .Replace("\\ln", "ln")
                 .Replace("\\log", "log");

            // Roman differential and operator names.
            s = Regex.Replace(s, @"\\mathrm\s*\{([^{}]*)\}", "$1");
            s = Regex.Replace(s, @"\\operatorname\s*\{([^{}]*)\}", "$1");

            // Vector/accent commands: use a combining accent on a grouped operand rather
            // than leaving a literal \\vec token for Math AutoCorrect to resolve.
            s = ReplaceUnaryBraced(s, "overrightarrow", body => "(" + body + ")⃗");
            s = ReplaceUnaryBraced(s, "vec", body => "(" + body + ")⃗");
            s = ReplaceUnaryBraced(s, "hat", body => "(" + body + ")̂");

            // Word's UnicodeMath matrix object is represented by U+25A0 (black square).
            // This avoids relying on Math AutoCorrect to turn the literal \\matrix token
            // into the matrix object during a programmatic BuildUp call.
            s = ConvertEnvironment(s, "cases", body => "{■(" + NormalizeRows(body) + ")");
            s = ConvertEnvironment(s, "matrix", body => "■(" + NormalizeRows(body) + ")");
            s = ConvertEnvironment(s, "pmatrix", body => "(■(" + NormalizeRows(body) + "))");
            s = ConvertEnvironment(s, "bmatrix", body => "[■(" + NormalizeRows(body) + ")]");

            // Make one-character subscripts/superscripts explicit before a following
            // parenthesized function argument: f_2(x) -> f_{2}(x), log_3(x) -> log_{3}(x).
            s = Regex.Replace(s, @"_([A-Za-z0-9])(?=\s*\()", "_{$1}");
            s = Regex.Replace(s, @"\^([A-Za-z0-9])(?=\s*\()", "^{$1}");

            // Normalize OCR spacing without changing mathematical tokens.
            s = Regex.Replace(s, @"[ \t]+", " ").Trim();
            return s;
        }

        private static string NormalizeRows(string body)
        {
            string s = body ?? string.Empty;
            s = s.Replace("\\\\", "@");
            return Regex.Replace(s, @"\s*@\s*", " @ ").Trim();
        }

        private static string ReplaceUnaryBraced(string input, string command, Func<string, string> projector)
        {
            string token = "\\" + command;
            string s = input;
            int guard = 0;
            while (guard++ < 128)
            {
                int pos = s.IndexOf(token, StringComparison.Ordinal);
                if (pos < 0) break;
                int brace = pos + token.Length;
                while (brace < s.Length && char.IsWhiteSpace(s[brace])) brace++;
                if (brace >= s.Length || s[brace] != '{') break;
                int end = FindMatchingBrace(s, brace);
                if (end < 0) break;
                string body = s.Substring(brace + 1, end - brace - 1);
                string replacement = projector(body);
                s = s.Substring(0, pos) + replacement + s.Substring(end + 1);
            }
            return s;
        }

        private static string ConvertEnvironment(string input, string env, Func<string, string> projector)
        {
            string begin = "\\begin{" + env + "}";
            string endToken = "\\end{" + env + "}";
            string s = input;
            int guard = 0;
            while (guard++ < 64)
            {
                int a = s.IndexOf(begin, StringComparison.Ordinal);
                if (a < 0) break;
                int b = s.IndexOf(endToken, a + begin.Length, StringComparison.Ordinal);
                if (b < 0) break;
                string body = s.Substring(a + begin.Length, b - a - begin.Length);
                s = s.Substring(0, a) + projector(body) + s.Substring(b + endToken.Length);
            }
            return s;
        }

        private static int FindMatchingBrace(string s, int start)
        {
            int depth = 0;
            for (int i = start; i < s.Length; i++)
            {
                if (s[i] == '{') depth++;
                else if (s[i] == '}')
                {
                    depth--;
                    if (depth == 0) return i;
                }
            }
            return -1;
        }
    }
}
