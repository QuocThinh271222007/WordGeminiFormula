using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace WordGeminiFormula.AddIn.Word
{
    public sealed class WordImageService
    {
        private readonly dynamic _wordApplication;

        public WordImageService(object wordApplication)
        {
            _wordApplication = wordApplication ?? throw new ArgumentNullException(nameof(wordApplication));
        }

        /// <summary>
        /// Exports the currently selected Word picture to a temporary PNG via the Windows clipboard.
        /// Returns null when the current selection does not contain/select a picture.
        /// Caller owns the returned temporary file and should delete it.
        /// </summary>
        public string TryExportSelectedPicture()
        {
            dynamic selection = _wordApplication.Selection;
            bool copied = false;

            try
            {
                if ((int)selection.InlineShapes.Count > 0)
                {
                    selection.InlineShapes.Item(1).Range.CopyAsPicture();
                    copied = true;
                }
            }
            catch { }

            if (!copied)
            {
                try
                {
                    if ((int)selection.ShapeRange.Count > 0)
                    {
                        // Copy the selected floating shape as a picture.
                        selection.CopyAsPicture();
                        copied = true;
                    }
                }
                catch { }
            }

            if (!copied) return null;

            // Word can populate the clipboard a few milliseconds after CopyAsPicture returns.
            for (int i = 0; i < 8 && !Clipboard.ContainsImage(); i++)
            {
                Application.DoEvents();
                Thread.Sleep(40);
            }
            if (!Clipboard.ContainsImage()) return null;

            using (Image image = Clipboard.GetImage())
            {
                if (image == null) return null;
                string path = Path.Combine(Path.GetTempPath(), "WordGeminiFormula-" + Guid.NewGuid().ToString("N") + ".png");
                image.Save(path, ImageFormat.Png);
                return path;
            }
        }
    }
}
