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
        private CheckBox _showKey;
        private CheckBox _autoNormalize;
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
            ClientSize = new Size(560, 285);
            Font = new Font("Segoe UI", 9F);

            var apiLabel = new Label { Left = 20, Top = 24, Width = 110, Text = "Gemini API key" };
            _apiKey = new TextBox { Left = 140, Top = 20, Width = 380, UseSystemPasswordChar = true };
            _apiKey.Text = _store.GetApiKey(_settings);

            _showKey = new CheckBox { Left = 140, Top = 52, Width = 150, Text = "Hiện API key" };
            _showKey.CheckedChanged += (_, __) => _apiKey.UseSystemPasswordChar = !_showKey.Checked;

            var modelLabel = new Label { Left = 20, Top = 91, Width = 110, Text = "Gemini model" };
            _model = new ComboBox { Left = 140, Top = 86, Width = 380, DropDownStyle = ComboBoxStyle.DropDown };
            _model.Items.AddRange(new object[] { "gemini-3.7-flash", "gemini-3.6-flash", "gemini-3.5-flash" });
            _model.Text = string.IsNullOrWhiteSpace(_settings.model) ? "gemini-3.7-flash" : _settings.model;

            _autoNormalize = new CheckBox
            {
                Left = 140,
                Top = 126,
                Width = 360,
                Text = "Tự chuẩn hóa công thức ngay sau OCR",
                Checked = _settings.autoNormalizeAfterOcr
            };

            _test = new Button { Left = 140, Top = 166, Width = 120, Height = 30, Text = "Test API" };
            _test.Click += TestClicked;

            _status = new Label { Left = 275, Top = 171, Width = 245, Height = 42, AutoEllipsis = true, Text = "" };

            _save = new Button { Left = 340, Top = 232, Width = 85, Height = 32, Text = "Lưu", DialogResult = DialogResult.None };
            _save.Click += SaveClicked;
            _cancel = new Button { Left = 435, Top = 232, Width = 85, Height = 32, Text = "Hủy", DialogResult = DialogResult.Cancel };

            AcceptButton = _save;
            CancelButton = _cancel;
            Controls.AddRange(new Control[] { apiLabel, _apiKey, _showKey, modelLabel, _model, _autoNormalize, _test, _status, _save, _cancel });
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
            _settings.autoNormalizeAfterOcr = _autoNormalize.Checked;
            _store.Save(_settings, _apiKey.Text.Trim());
            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
