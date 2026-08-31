using System;
using System.Drawing;
using System.Windows.Forms;
using WordGeminiFormula.AddIn.Gemini;

namespace WordGeminiFormula.AddIn.Settings
{
    public sealed class SettingsForm : Form
    {
        private readonly SettingsStore _store;
        private readonly GeminiClient _client;
        private readonly AppSettings _settings;

        private TextBox _apiKey;
        private ComboBox _model;
        private ComboBox _preset;
        private CheckBox _showKey;
        private CheckBox _autoNormalize;
        private CheckBox _autoBeautify;
        private CheckBox _preserveDifficult;
        private Button _test;
        private Button _save;
        private Button _cancel;
        private Label _status;

        public SettingsForm(SettingsStore store, GeminiClient client)
        {
            _store = store ?? throw new ArgumentNullException(nameof(store));
            _client = client ?? throw new ArgumentNullException(nameof(client));
            _settings = _store.Load();
            InitializeUi();
        }

        private void InitializeUi()
        {
            Text = "Word Gemini Formula - Settings";
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ClientSize = new Size(590, 405);
            Font = new Font("Segoe UI", 9F);

            var apiLabel = new Label { Left = 20, Top = 24, Width = 120, Text = "Gemini API key" };
            _apiKey = new TextBox { Left = 150, Top = 20, Width = 400, UseSystemPasswordChar = true };
            _apiKey.Text = _store.GetApiKey(_settings);

            _showKey = new CheckBox { Left = 150, Top = 52, Width = 150, Text = "Hiện API key" };
            _showKey.CheckedChanged += (_, __) => _apiKey.UseSystemPasswordChar = !_showKey.Checked;

            var modelLabel = new Label { Left = 20, Top = 91, Width = 120, Text = "Gemini model" };
            _model = new ComboBox { Left = 150, Top = 86, Width = 400, DropDownStyle = ComboBoxStyle.DropDown };
            _model.Items.AddRange(new object[] { "gemini-3.7-flash", "gemini-3.6-flash", "gemini-3.5-flash" });
            _model.Text = string.IsNullOrWhiteSpace(_settings.model) ? "gemini-3.7-flash" : _settings.model;

            var presetLabel = new Label { Left = 20, Top = 131, Width = 120, Text = "Kiểu tài liệu" };
            _preset = new ComboBox { Left = 150, Top = 126, Width = 240, DropDownStyle = ComboBoxStyle.DropDownList };
            _preset.Items.AddRange(new object[] { "exam", "general" });
            _preset.SelectedItem = string.IsNullOrWhiteSpace(_settings.documentPreset) ? "exam" : _settings.documentPreset;
            if (_preset.SelectedIndex < 0) _preset.SelectedIndex = 0;

            _autoBeautify = new CheckBox
            {
                Left = 150,
                Top = 166,
                Width = 380,
                Text = "Tự làm đẹp format sau OCR",
                Checked = _settings.autoBeautifyAfterOcr
            };

            _autoNormalize = new CheckBox
            {
                Left = 150,
                Top = 196,
                Width = 400,
                Text = "Tự chuẩn hóa công thức ngay sau OCR",
                Checked = _settings.autoNormalizeAfterOcr
            };

            _preserveDifficult = new CheckBox
            {
                Left = 150,
                Top = 226,
                Width = 420,
                Text = "Giữ hình/bảng khó OCR bằng ảnh crop gốc",
                Checked = _settings.preserveDifficultRegionsAsImage
            };

            var hint = new Label
            {
                Left = 150,
                Top = 254,
                Width = 400,
                Height = 38,
                ForeColor = Color.DimGray,
                Text = "Khuyến nghị cho đề Toán: bật làm đẹp + giữ vùng khó; để tự chuẩn hóa công thức tắt khi cần kiểm tra OCR trước."
            };

            _test = new Button { Left = 150, Top = 302, Width = 120, Height = 30, Text = "Test API" };
            _test.Click += TestClicked;

            _status = new Label { Left = 285, Top = 307, Width = 265, Height = 42, AutoEllipsis = true, Text = "" };

            _save = new Button { Left = 370, Top = 353, Width = 85, Height = 32, Text = "Lưu", DialogResult = DialogResult.None };
            _save.Click += SaveClicked;
            _cancel = new Button { Left = 465, Top = 353, Width = 85, Height = 32, Text = "Hủy", DialogResult = DialogResult.Cancel };

            AcceptButton = _save;
            CancelButton = _cancel;
            Controls.AddRange(new Control[]
            {
                apiLabel, _apiKey, _showKey, modelLabel, _model, presetLabel, _preset,
                _autoBeautify, _autoNormalize, _preserveDifficult, hint,
                _test, _status, _save, _cancel
            });
        }

        private void TestClicked(object sender, EventArgs e)
        {
            _test.Enabled = false;
            _status.Text = "Đang kiểm tra...";
            try
            {
                bool ok = _client.TestConnection(_apiKey.Text.Trim(), _model.Text.Trim(), out string message);
                _status.ForeColor = ok ? Color.DarkGreen : Color.DarkRed;
                _status.Text = message;
            }
            finally
            {
                _test.Enabled = true;
            }
        }

        private void SaveClicked(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_apiKey.Text))
            {
                MessageBox.Show(this, "Hãy nhập Gemini API key.", "Word Gemini Formula", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (string.IsNullOrWhiteSpace(_model.Text))
            {
                MessageBox.Show(this, "Hãy nhập/chọn Gemini model.", "Word Gemini Formula", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            _settings.model = _model.Text.Trim();
            _settings.documentPreset = _preset.SelectedItem?.ToString() ?? "exam";
            _settings.autoBeautifyAfterOcr = _autoBeautify.Checked;
            _settings.autoNormalizeAfterOcr = _autoNormalize.Checked;
            _settings.preserveDifficultRegionsAsImage = _preserveDifficult.Checked;
            _store.Save(_settings, _apiKey.Text.Trim());
            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
