using System;
using System.Text.RegularExpressions;

namespace WordGeminiFormula.AddIn.Word
{
    /// <summary>
    /// Conservative LaTeX -> Word UnicodeMath converter for common school/university formulas.
    /// Unknown commands are intentionally left intact so Word Math AutoCorrect can still handle them.
    /// </summary>
    public sealed class LatexToUnicodeMathConverter
    {
        private readonly MathRepairService _repair = new MathRepairService();

        public string Convert(string latex)
        {
            if (string.IsNullOrWhiteSpace(latex)) return string.Empty;

            string s = _repair.RepairLatex(latex);
            s = Regex.Replace(s, @"^\$\$?|\$\$?$", string.Empty).Trim();
            s = s.Replace("\\displaystyle", string.Empty)
                 .Replace("\\left", string.Empty)
                 .Replace("\\right", string.Empty)
                 .Replace("\\,", " ")
                 .Replace("\\;", " ")
                 .Replace("\\!", string.Empty)
                 .Replace("~", " ");

            s = ConvertEnvironments(s);
            s = ConvertFunctions(s);
            s = ConvertFractions(s);
            s = ConvertRoots(s);
            s = ConvertUnaryBracedCommand(s, "overrightarrow", body => "(" + body + ")\\vec");
            s = ConvertUnaryBracedCommand(s, "overleftarrow", body => "(" + body + ")\\vec");
            s = ConvertUnaryBracedCommand(s, "vec", body => "(" + body + ")\\vec");
            s = ConvertUnaryBracedCommand(s, "hat", body => "(" + body + ")\\hat");
            s = ConvertUnaryBracedCommand(s, "bar", body => "\\overbar(" + body + ")");
            s = ConvertUnaryBracedCommand(s, "overline", body => "\\overbar(" + body + ")");
            s = ConvertUnaryBracedCommand(s, "underline", body => "\\underbar(" + body + ")");
            s = ConvertText(s);
            s = NormalizeBracedScripts(s);

            s = s.Replace("\\mathbb{R}", "ℝ")
                 .Replace("\\mathbb{N}", "ℕ")
                 .Replace("\\mathbb{Z}", "ℤ")
                 .Replace("\\mathbb{Q}", "ℚ")
                 .Replace("\\mathbb{C}", "ℂ")
                 .Replace("\\neq", "≠")
                 .Replace("\\ne", "≠")
                 .Replace("\\rightarrow", "->");
            s = Regex.Replace(s, @"\\to(?![A-Za-z])", "->");
            s = Regex.Replace(s, @"\\le(?![A-Za-z])", "\\leq");
            s = Regex.Replace(s, @"\\ge(?![A-Za-z])", "\\geq");

            s = Regex.Replace(s, @"\s+", " ").Trim();
            return _repair.RepairWordLinear(s);
        }

        private string ConvertFractions(string input)
        {
            string s = input;
            while (TryFindCommandWithTwoBracedArgs(s, "\\frac", out int start, out int length, out string a, out string b))
            {
                string replacement = "(" + Convert(a) + ")/(" + Convert(b) + ")";
                s = s.Substring(0, start) + replacement + s.Substring(start + length);
            }
            return s;
        }

        private string ConvertRoots(string input)
        {
            string s = input;
            int guard = 0;
            while (guard++ < 256)
            {
                int pos = s.IndexOf("\\sqrt", StringComparison.Ordinal);
                if (pos < 0) break;

                int i = pos + 5;
                string degree = null;
                SkipSpaces(s, ref i);
                if (i < s.Length && s[i] == '[')
                {
                    int close = FindMatching(s, i, '[', ']');
                    if (close < 0) break;
                    degree = s.Substring(i + 1, close - i - 1);
                    i = close + 1;
                }

                SkipSpaces(s, ref i);
                if (i >= s.Length || s[i] != '{') break;
                int bodyEnd = FindMatching(s, i, '{', '}');
                if (bodyEnd < 0) break;

                string body = s.Substring(i + 1, bodyEnd - i - 1);
                string replacement = degree == null
                    ? "\\sqrt(" + Convert(body) + ")"
                    : "\\sqrt(" + Convert(degree) + "&" + Convert(body) + ")";

                s = s.Substring(0, pos) + replacement + s.Substring(bodyEnd + 1);
            }
            return s;
        }

        private string ConvertUnaryBracedCommand(string input, string command, Func<string, string> projector)
        {
            string token = "\\" + command;
            string s = input;
            int guard = 0;
            while (guard++ < 256)
            {
                int pos = s.IndexOf(token, StringComparison.Ordinal);
                if (pos < 0) break;
                int i = pos + token.Length;
                SkipSpaces(s, ref i);
                if (i >= s.Length || s[i] != '{') break;
                int end = FindMatching(s, i, '{', '}');
                if (end < 0) break;
                string body = s.Substring(i + 1, end - i - 1);
                string replacement = projector(Convert(body));
                s = s.Substring(0, pos) + replacement + s.Substring(end + 1);
            }
            return s;
        }

        private static string ConvertFunctions(string input)
        {
            return input
                .Replace("\\operatorname{sin}", "sin")
                .Replace("\\operatorname{cos}", "cos")
                .Replace("\\operatorname{tan}", "tan")
                .Replace("\\operatorname{ln}", "ln")
                .Replace("\\operatorname{log}", "log");
        }

        private static string ConvertText(string input)
        {
            string s = input;
            int guard = 0;
            while (guard++ < 256)
            {
                int pos = s.IndexOf("\\text", StringComparison.Ordinal);
                if (pos < 0) break;
                int i = pos + 5;
                SkipSpaces(s, ref i);
                if (i >= s.Length || s[i] != '{') break;
                int end = FindMatching(s, i, '{', '}');
                if (end < 0) break;
                string body = s.Substring(i + 1, end - i - 1);
                string replacement = "\\text(" + body + ")";
                s = s.Substring(0, pos) + replacement + s.Substring(end + 1);
            }
            return s;
        }

        private static string NormalizeBracedScripts(string input)
        {
            string s = input;
            s = Regex.Replace(s, @"\^\{([^{}]+)\}", "^($1)");
            s = Regex.Replace(s, @"_\{([^{}]+)\}", "_($1)");
            return s;
        }

        private string ConvertEnvironments(string input)
        {
            string s = input;
            s = ConvertArrayEnvironment(s);
            s = ConvertMatrixEnvironment(s, "matrix");
            s = ConvertMatrixEnvironment(s, "pmatrix");
            s = ConvertMatrixEnvironment(s, "bmatrix");
            s = ConvertMatrixEnvironment(s, "vmatrix");
            s = ConvertMatrixEnvironment(s, "Vmatrix");
            s = ConvertCasesEnvironment(s);
            return s;
        }

        private string ConvertArrayEnvironment(string input)
        {
            string s = input;
            const string beginToken = "\\begin{array}";
            const string endToken = "\\end{array}";
            int guard = 0;
            while (guard++ < 64)
            {
                int a = s.IndexOf(beginToken, StringComparison.Ordinal);
                if (a < 0) break;
                int bodyStart = a + beginToken.Length;
                SkipSpaces(s, ref bodyStart);
                if (bodyStart < s.Length && s[bodyStart] == '{')
                {
                    int specEnd = FindMatching(s, bodyStart, '{', '}');
                    if (specEnd > bodyStart) bodyStart = specEnd + 1;
                }
                int b = s.IndexOf(endToken, bodyStart, StringComparison.Ordinal);
                if (b < 0) break;
                string body = s.Substring(bodyStart, b - bodyStart)
                    .Replace("\\hline", string.Empty)
                    .Replace("\\\\", "@");
                string replacement = "\\matrix(" + body + ")";
                s = s.Substring(0, a) + replacement + s.Substring(b + endToken.Length);
            }
            return s;
        }

        private string ConvertMatrixEnvironment(string input, string env)
        {
            string begin = "\\begin{" + env + "}";
            string end = "\\end{" + env + "}";
            string s = input;
            int guard = 0;
            while (guard++ < 64)
            {
                int a = s.IndexOf(begin, StringComparison.Ordinal);
                if (a < 0) break;
                int b = s.IndexOf(end, a + begin.Length, StringComparison.Ordinal);
                if (b < 0) break;
                string body = s.Substring(a + begin.Length, b - a - begin.Length);
                body = body.Replace("\\\\", "@");
                string matrix = "\\matrix(" + body + ")";
                if (env == "pmatrix") matrix = "(" + matrix + ")";
                else if (env == "bmatrix") matrix = "[" + matrix + "]";
                else if (env == "vmatrix") matrix = "|" + matrix + "|";
                else if (env == "Vmatrix") matrix = "‖" + matrix + "‖";
                s = s.Substring(0, a) + matrix + s.Substring(b + end.Length);
            }
            return s;
        }

        private static string ConvertCasesEnvironment(string input)
        {
            const string begin = "\\begin{cases}";
            const string end = "\\end{cases}";
            string s = input;
            int guard = 0;
            while (guard++ < 64)
            {
                int a = s.IndexOf(begin, StringComparison.Ordinal);
                if (a < 0) break;
                int b = s.IndexOf(end, a + begin.Length, StringComparison.Ordinal);
                if (b < 0) break;
                string body = s.Substring(a + begin.Length, b - a - begin.Length).Replace("\\\\", "@");
                string replacement = "{\\matrix(" + body + ")";
                s = s.Substring(0, a) + replacement + s.Substring(b + end.Length);
            }
            return s;
        }

        private static bool TryFindCommandWithTwoBracedArgs(string s, string command, out int start, out int length, out string a, out string b)
        {
            start = s.IndexOf(command, StringComparison.Ordinal);
            length = 0;
            a = b = null;
            if (start < 0) return false;

            int i = start + command.Length;
            SkipSpaces(s, ref i);
            if (i >= s.Length || s[i] != '{') return false;
            int aEnd = FindMatching(s, i, '{', '}');
            if (aEnd < 0) return false;
            a = s.Substring(i + 1, aEnd - i - 1);

            i = aEnd + 1;
            SkipSpaces(s, ref i);
            if (i >= s.Length || s[i] != '{') return false;
            int bEnd = FindMatching(s, i, '{', '}');
            if (bEnd < 0) return false;
            b = s.Substring(i + 1, bEnd - i - 1);

            length = bEnd + 1 - start;
            return true;
        }

        private static int FindMatching(string s, int start, char open, char close)
        {
            int depth = 0;
            for (int i = start; i < s.Length; i++)
            {
                if (s[i] == open) depth++;
                else if (s[i] == close)
                {
                    depth--;
                    if (depth == 0) return i;
                }
            }
            return -1;
        }

        private static void SkipSpaces(string s, ref int i)
        {
            while (i < s.Length && char.IsWhiteSpace(s[i])) i++;
        }
    }
}
