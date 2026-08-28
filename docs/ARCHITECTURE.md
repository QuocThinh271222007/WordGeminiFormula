# Architecture

WordGeminiFormula is intentionally independent from any other project.

## Runtime flow

1. Word loads `WordGeminiFormula.AddIn.Connect` as a per-user COM add-in.
2. `IRibbonExtensibility.GetCustomUI` injects the **AI Formula** Ribbon tab.
3. **Ảnh → Word** first uses a picture selected inside Word (exported temporarily to PNG); otherwise it opens a local image file. The image is sent directly to Gemini `generateContent` using the API key stored by the user in Settings.
4. Gemini returns ordered blocks: normal text + formula blocks (`latex`, `word_linear`).
5. By default formulas are staged as `[[MATH]]...[[/MATH]]` so the user can inspect the OCR before conversion.
6. **Chuẩn hóa tất cả** converts marked LaTeX to Word UnicodeMath, creates an `OMath`, then invokes `BuildUp()` to produce native editable Word equations.
7. **Chuẩn hóa vùng chọn** performs the same conversion only for the selected expression.

## Secret storage

The API key is entered in the Word UI. It is encrypted using Windows DPAPI with `DataProtectionScope.CurrentUser` and stored under:

`%LOCALAPPDATA%\WordGeminiFormula\settings.json`

No API key is checked into source control and no backend configuration file is required.

## V1 constraints

- Windows desktop Word only.
- .NET Framework 4.8.
- Images are sent inline and capped at 18 MB by the add-in.
- The local LaTeX converter targets common school/university notation. OCR output also contains Gemini-generated `word_linear` for optional auto-normalization.
- The add-in is synchronous during the API request in V1, so Word may be temporarily unresponsive while Gemini processes an image.
