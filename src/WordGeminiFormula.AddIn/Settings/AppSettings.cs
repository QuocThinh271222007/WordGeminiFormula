namespace WordGeminiFormula.AddIn.Settings
{
    public sealed class AppSettings
    {
        public string encryptedApiKey { get; set; }
        public string model { get; set; } = "gemini-3.7-flash";
        public bool autoNormalizeAfterOcr { get; set; } = false;
        public bool autoBeautifyAfterOcr { get; set; } = true;
        public bool preserveDifficultRegionsAsImage { get; set; } = true;
        public string documentPreset { get; set; } = "exam";
    }
}
