using System;
using System.IO;
using System.Runtime.InteropServices;
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
        private readonly SettingsStore _settingsStore = new SettingsStore();
        private readonly GeminiClient _geminiClient = new GeminiClient();

        public string GetCustomUI(string ribbonId) => RibbonXml.Value;

        public void OnConnection(object application, ExtConnectMode connectMode, object addInInst, ref Array custom)
        {
            _wordApplication = application;
        }

        public void OnDisconnection(ExtDisconnectMode removeMode, ref Array custom)
        {
            _wordApplication = null;
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }

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

                    var document = _geminiClient.OcrImage(apiKey, settings.model, imagePath);
                    var word = new WordDocumentService(_wordApplication);
                    word.InsertOcrBlocks(document, settings.autoNormalizeAfterOcr);
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
                ShowError(ex);
            }
        }

        public void OnNormalizeAll(object control)
        {
            try
            {
                EnsureConnected();
                var word = new WordDocumentService(_wordApplication);
                int count = word.NormalizeAllMarkedFormulas();
                MessageBox.Show(count == 0
                        ? "Không tìm thấy khối [[MATH]] nào để chuẩn hóa."
                        : $"Đã chuẩn hóa {count} công thức thành Word Equation.",
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
                new WordDocumentService(_wordApplication).NormalizeSelection();
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

        private static void ShowError(Exception ex)
        {
            MessageBox.Show(ex.Message, "Word Gemini Formula", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
