namespace WordGeminiFormula.AddIn.Settings
{
    public sealed class AppSettings
    {
        public string encryptedApiKey { get; set; }
        public string model { get; set; } = "gemini-3.7-flash";
        public bool autoNormalizeAfterOcr { get; set; } = false;
    }
}
