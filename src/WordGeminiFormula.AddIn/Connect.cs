using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using WordGeminiFormula.AddIn.Gemini;
using WordGeminiFormula.AddIn.Interop;
using WordGeminiFormula.AddIn.Ribbon;
using WordGeminiFormula.AddIn.Settings;
using WordGeminiFormula.AddIn.Word;

namespace WordGeminiFormula.AddIn
{
    [ComVisible(true)]
    [Guid("7BA1B881-3DA4-4FBA-A25D-5F92141658EE")]
    [ProgId("WordGeminiFormula.AddIn")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public sealed class Connect : IDTExtensibility2, IRibbonExtensibility
    {
        private object _wordApplication;
        private readonly SettingsStore _settingsStore;
        private readonly GeminiClient _geminiClient;

        public Connect()
        {
            StartupTrace.Write("Connect.ctor begin");
            try
            {
                _settingsStore = new SettingsStore();
                _geminiClient = new GeminiClient();
                StartupTrace.Write("Connect.ctor success");
            }
            catch (Exception ex)
            {
                StartupTrace.Write("Connect.ctor FAILED: " + ex);
                throw;
            }
        }

        public string GetCustomUI(string ribbonId)
        {
            StartupTrace.Write("GetCustomUI: " + (ribbonId ?? "<null>"));
            return RibbonXml.Value;
        }

        public void OnConnection(object application, ExtConnectMode connectMode, object addInInst, ref Array custom)
        {
            StartupTrace.Write("OnConnection begin; mode=" + connectMode);
            _wordApplication = application;
            StartupTrace.Write("OnConnection success");
        }

        public void OnDisconnection(ExtDisconnectMode removeMode, ref Array custom)
        {
            StartupTrace.Write("OnDisconnection; mode=" + removeMode);
            _wordApplication = null;
        }

        public void OnAddInsUpdate(ref Array custom) { StartupTrace.Write("OnAddInsUpdate"); }
        public void OnStartupComplete(ref Array custom) { StartupTrace.Write("OnStartupComplete"); }
        public void OnBeginShutdown(ref Array custom) { StartupTrace.Write("OnBeginShutdown"); }

        public void OnOpenSettings(object control)
        {
            try
            {
                using (var form = new SettingsForm(_settingsStore, _geminiClient))
                    form.ShowDialog();
            }
            catch (Exception ex)
            {
                ShowError(ex);
            }
        }

        public void OnOcrImage(object control)
        {
            try
            {
                EnsureConnected();
                AppSettings settings = _settingsStore.Load();
                string apiKey = _settingsStore.GetApiKey(settings);
                if (string.IsNullOrWhiteSpace(apiKey))
                {
                    using (var form = new SettingsForm(_settingsStore, _geminiClient))
                    {
                        if (form.ShowDialog() != DialogResult.OK) return;
                    }
                    settings = _settingsStore.Load();
                    apiKey = _settingsStore.GetApiKey(settings);
                }

                string tempSelectedImage = null;
                string imagePath = null;
                try
                {
                    tempSelectedImage = new WordImageService(_wordApplication).TryExportSelectedPicture();
                    imagePath = tempSelectedImage;

                    if (string.IsNullOrWhiteSpace(imagePath))
                    {
                        using (var dialog = new OpenFileDialog())
                        {
                            dialog.Title = "Chọn ảnh cần OCR";
                            dialog.Filter = "Image files|*.png;*.jpg;*.jpeg;*.webp;*.bmp;*.gif|All files|*.*";
                            dialog.Multiselect = false;
                            if (dialog.ShowDialog() != DialogResult.OK) return;
                            imagePath = dialog.FileName;
                        }
                    }

                    var document = _geminiClient.OcrImage(apiKey, settings.model, imagePath, settings.documentPreset);
                    string rawSnapshotPath = SaveRawOcrSnapshot(_geminiClient.LastRawOcrJson);

                    var word = new WordDocumentServiceV5(_wordApplication);

                    // V5 uses Presentation MathML + Microsoft's MML2OMML transform as the
                    // preferred equation path while retaining V4 layout and V3 bookmark
                    // safety. Native Office Math insertion remains deferred until the page
                    // has completely rendered.
                    word.InsertOcrBlocks(
                        document,
                        false,
                        settings.autoBeautifyAfterOcr,
                        imagePath,
                        settings.preserveDifficultRegionsAsImage);

                    int normalizedCount = 0;
                    if (settings.autoNormalizeAfterOcr)
                    {
                        StartupTrace.Write("OCR render complete; starting deferred V5 MathML/OMML normalization");
                        normalizedCount = word.NormalizeAllMarkedFormulas();
                        StartupTrace.Write("Deferred V5 math normalization complete; converted=" + normalizedCount + "; failed=" + word.LastNormalizationFailureCount);
                    }

                    string warningText = document.warnings != null && document.warnings.Count > 0
                        ? "\n\nGemini cảnh báo:\n- " + string.Join("\n- ", document.warnings)
                        : string.Empty;
                    string formulaText = settings.autoNormalizeAfterOcr && word.LastNormalizationFailureCount > 0
                        ? $"\n\nĐã chuẩn hóa {normalizedCount} công thức. Có {word.LastNormalizationFailureCount} công thức chưa chuyển đổi được và đã được giữ nguyên/tô vàng để kiểm tra."
                        : settings.autoNormalizeAfterOcr
                            ? $"\n\nĐã chuẩn hóa {normalizedCount} công thức bằng MathML/Office Math sau khi hoàn tất render."
                            : string.Empty;
                    string snapshotText = string.IsNullOrWhiteSpace(rawSnapshotPath)
                        ? string.Empty
                        : "\n\nRaw OCR: " + rawSnapshotPath;

                    MessageBox.Show(
                        "Đã OCR và chèn nội dung vào Word." + warningText + formulaText + snapshotText,
                        "Word Gemini Formula",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
                finally
                {
                    if (!string.IsNullOrWhiteSpace(tempSelectedImage))
                    {
                        try { File.Delete(tempSelectedImage); } catch { }
                    }
                }
            }
            catch (Exception ex)
            {
                StartupTrace.Write("OnOcrImage FAILED: " + ex);
                ShowError(ex);
            }
        }

        public void OnBeautifyFormat(object control)
        {
            try
            {
                EnsureConnected();
                var word = new WordDocumentServiceV5(_wordApplication);
                int count = word.BeautifyActiveDocument();
                MessageBox.Show(
                    count == 0 ? "Không tìm thấy đoạn văn nào để format." : $"Đã làm đẹp format cho {count} đoạn văn.",
                    "Word Gemini Formula",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                ShowError(ex);
            }
        }

        public void OnNormalizeAll(object control)
        {
            try
            {
                EnsureConnected();
                var word = new WordDocumentServiceV5(_wordApplication);
                int count = word.NormalizeAllMarkedFormulas();
                string failed = word.LastNormalizationFailureCount > 0
                    ? $" Có {word.LastNormalizationFailureCount} công thức chưa chuyển đổi được; chúng được giữ nguyên và tô vàng."
                    : string.Empty;
                MessageBox.Show(
                    count == 0 && word.LastNormalizationFailureCount == 0
                        ? "Không tìm thấy công thức chờ chuẩn hóa."
                        : $"Đã chuẩn hóa {count} công thức thành Word Equation.{failed}",
                    "Word Gemini Formula",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                ShowError(ex);
            }
        }

        public void OnNormalizeSelection(object control)
        {
            try
            {
                EnsureConnected();
                new WordDocumentServiceV5(_wordApplication).NormalizeSelection();
            }
            catch (Exception ex)
            {
                ShowError(ex);
            }
        }

        private void EnsureConnected()
        {
            if (_wordApplication == null)
                throw new InvalidOperationException("Add-in chưa kết nối được với Microsoft Word.");
        }

        private static string SaveRawOcrSnapshot(string rawJson)
        {
            if (string.IsNullOrWhiteSpace(rawJson)) return null;
            try
            {
                string dir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "WordGeminiFormula");
                Directory.CreateDirectory(dir);
                string path = Path.Combine(dir, "last-ocr.json");
                File.WriteAllText(path, rawJson, new UTF8Encoding(false));
                return path;
            }
            catch (Exception ex)
            {
                StartupTrace.Write("SaveRawOcrSnapshot FAILED: " + ex.Message);
                return null;
            }
        }

        private static void ShowError(Exception ex)
        {
            MessageBox.Show(ex.Message, "Word Gemini Formula", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private static class StartupTrace
        {
            private static readonly object Sync = new object();

            internal static void Write(string message)
            {
                try
                {
                    string dir = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                        "WordGeminiFormula");
                    Directory.CreateDirectory(dir);
                    string path = Path.Combine(dir, "addin-startup.log");
                    lock (Sync)
                    {
                        File.AppendAllText(
                            path,
                            DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + " | " + message + Environment.NewLine);
                    }
                }
                catch
                {
                    // Diagnostics must never prevent Word from loading the add-in.
                }
            }
        }
    }
}
