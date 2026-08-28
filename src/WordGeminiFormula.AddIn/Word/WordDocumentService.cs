using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using WordGeminiFormula.AddIn.Models;

namespace WordGeminiFormula.AddIn.Word
{
    public sealed class WordDocumentService
    {
        private readonly dynamic _wordApplication;
        private readonly LatexToUnicodeMathConverter _converter = new LatexToUnicodeMathConverter();
        private const string MathStart = "[[MATH]]";
        private const string MathEnd = "[[/MATH]]";

        public WordDocumentService(object wordApplication)
        {
            _wordApplication = wordApplication ?? throw new ArgumentNullException(nameof(wordApplication));
        }

        public void InsertOcrBlocks(OcrDocument document, bool autoNormalize)
        {
            if (document?.blocks == null) return;
            dynamic selection = _wordApplication.Selection;
            selection.Collapse(0); // wdCollapseEnd

            foreach (OcrBlock block in document.blocks)
            {
                if (block == null) continue;
                if (string.Equals(block.type, "formula", StringComparison.OrdinalIgnoreCase))
                {
                    string source = !string.IsNullOrWhiteSpace(block.latex) ? block.latex : block.word_linear;
                    if (autoNormalize)
                    {
                        string linear = !string.IsNullOrWhiteSpace(block.word_linear)
                            ? block.word_linear
                            : _converter.Convert(source);
                        InsertEquationAtSelection(linear);
                    }
                    else
                    {
                        selection.TypeText(MathStart + source + MathEnd);
                    }
                }
                else
                {
                    string text = block.text ?? string.Empty;
                    selection.TypeText(text);
                }

                selection.TypeParagraph();
            }
        }

        public int NormalizeAllMarkedFormulas()
        {
            dynamic document = _wordApplication.ActiveDocument;
            string fullText = (string)document.Content.Text;
            var regex = new Regex(@"\[\[MATH\]\](.*?)\[\[/MATH\]\]", RegexOptions.Singleline);
            MatchCollection matches = regex.Matches(fullText);

            int converted = 0;
            for (int i = matches.Count - 1; i >= 0; i--)
            {
                Match match = matches[i];
                int start = match.Index;
                int end = match.Index + match.Length;
                dynamic range = document.Range(start, end);
                string latex = match.Groups[1].Value.Trim();
                string linear = _converter.Convert(latex);
                ReplaceRangeWithEquation(range, linear);
                converted++;
            }

            return converted;
        }

        public void NormalizeSelection()
        {
            dynamic selection = _wordApplication.Selection;
            if ((int)selection.Start == (int)selection.End)
                throw new InvalidOperationException("Hãy bôi đen công thức LaTeX/linear cần chuẩn hóa trước.");

            string raw = ((string)selection.Text ?? string.Empty).Trim();
            raw = raw.Replace(MathStart, string.Empty).Replace(MathEnd, string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(raw))
                throw new InvalidOperationException("Vùng chọn không có công thức.");

            string linear = _converter.Convert(raw);
            dynamic range = selection.Range;
            ReplaceRangeWithEquation(range, linear);
        }

        private void InsertEquationAtSelection(string wordLinear)
        {
            dynamic selection = _wordApplication.Selection;
            dynamic document = _wordApplication.ActiveDocument;
            int start = (int)selection.Start;
            selection.TypeText(wordLinear ?? string.Empty);
            int end = (int)selection.End;
            dynamic range = document.Range(start, end);
            dynamic eqRange = document.OMaths.Add(range);
            if ((int)eqRange.OMaths.Count > 0)
            {
                dynamic equation = eqRange.OMaths.Item(1);
                equation.BuildUp();
            }
            selection.SetRange((int)eqRange.End, (int)eqRange.End);
        }

        private void ReplaceRangeWithEquation(dynamic range, string wordLinear)
        {
            dynamic document = _wordApplication.ActiveDocument;
            int start = (int)range.Start;
            range.Text = wordLinear ?? string.Empty;
            int end = (int)range.End;
            dynamic eqInput = document.Range(start, end);
            dynamic eqRange = document.OMaths.Add(eqInput);
            if ((int)eqRange.OMaths.Count > 0)
            {
                dynamic equation = eqRange.OMaths.Item(1);
                equation.BuildUp();
            }
        }
    }
}
