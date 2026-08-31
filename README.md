# Word Gemini Formula

A standalone Microsoft Word desktop add-in that converts images of Vietnamese/math documents into editable Word text and native Word equations using the Gemini API.

## V0.2 features

- **Ảnh → Word đẹp**: OCR the picture currently selected in Word; if no picture is selected, choose a PNG/JPG/WEBP/BMP/GIF file.
- **Layout-aware OCR**: Gemini now returns structured document blocks instead of only flat text/formula lines.
- **Exam reconstruction**: recognizes exam headers, candidate fields, exam-code boxes, sections, questions, inline math, A/B/C/D choices and footers.
- **Làm đẹp format**: beautifies the active Word document using a school/exam-oriented Times New Roman layout, spacing, headings and question formatting.
- **Multiple-choice layout**: structured choices are rebuilt into a clean 2-column borderless table when beautification is enabled.
- **Difficult-region fallback**: geometry diagrams, graphs, variation tables and unresolved regions can be cropped from the original image and embedded instead of being flattened into incorrect prose.
- **Math repair layer**: conservative repairs for common OCR defects such as `\mathbbR`, vector commands, array/cases structures and several Word UnicodeMath conversions.
- **Review-first formulas**: OCR formulas are kept as `[[MATH]]...[[/MATH]]` by default so you can verify them before native equation conversion.
- **Chuẩn hóa tất cả**: converts marked formulas into native editable Word Equation objects and calls Word's Professional `BuildUp()` conversion. Failed conversions remain editable and are highlighted yellow instead of being silently destroyed.
- **Chuẩn hóa vùng chọn**: converts only the selected LaTeX/linear formula.
- **Settings inside Word**: enter the Gemini API key and model without editing backend/source files.
- **Protected API key**: stored locally with Windows DPAPI for the current Windows user.

## Recommended settings for mathematics exams

- **Tự làm đẹp format sau OCR**: ON
- **Giữ hình/bảng khó OCR bằng ảnh crop gốc**: ON
- **Tự chuẩn hóa công thức ngay sau OCR**: OFF for review-first workflows
- **Kiểu tài liệu**: `exam`

## Requirements

- Windows 10/11
- Microsoft Word desktop (Microsoft 365 / Word 2019+ recommended)
- .NET Framework 4.8
- A Gemini Developer API key
- Visual Studio/MSBuild only if building from source

The default model is `gemini-3.7-flash`, and the model field is editable in Settings.

## Fast install from GitHub Actions artifact

1. Open the latest successful **build** workflow under GitHub Actions.
2. Download the `WordGeminiFormula` artifact.
3. Extract the complete ZIP into a permanent folder, for example:

```text
C:\Tools\WordGeminiFormula\
```

The folder should contain:

```text
WordGeminiFormula.AddIn.dll
install.ps1
uninstall.ps1
README.md
```

4. Close all Word windows.
5. Run PowerShell **as Administrator** and execute:

```powershell
cd C:\Tools\WordGeminiFormula
powershell -NoProfile -ExecutionPolicy Bypass -File .\install.ps1
```

The installer uses the .NET Framework `RegAsm.exe /codebase` path and registers both 32-bit and 64-bit COM views when applicable. A RegAsm warning about an unsigned `/codebase` assembly is expected for this development build; `Types registered successfully` is the important success message.

6. Reopen Word. The Ribbon tab **AI Formula** should appear.

## Build from source

Open `WordGeminiFormula.sln` in Visual Studio or run:

```powershell
msbuild WordGeminiFormula.sln /t:Restore,Build /p:Configuration=Release /p:Platform="Any CPU"
```

Output:

```text
src\WordGeminiFormula.AddIn\bin\Release\net48\WordGeminiFormula.AddIn.dll
```

Then run `scripts\install.ps1` as Administrator.

## First use

1. Open Word → **AI Formula** → **Settings**.
2. Paste your Gemini API key.
3. Choose a model and click **Test API**.
4. Keep **Tự làm đẹp format sau OCR** enabled.
5. Keep **Giữ hình/bảng khó OCR bằng ảnh crop gốc** enabled for exam scans.
6. For safer math review, leave **Tự chuẩn hóa công thức ngay sau OCR** disabled.
7. Click **Lưu**.
8. Select an image already pasted in Word and click **Ảnh → Word đẹp**, or click it with no image selected to choose an image file.
9. Review any `[[MATH]]...[[/MATH]]` blocks and any yellow-highlighted unresolved formulas.
10. Click **Chuẩn hóa tất cả** when the OCR content is correct.

## Làm đẹp format

The **Làm đẹp format** button can also be used on an existing Word document. It applies conservative document styling such as:

- Times New Roman base font
- page margins suitable for printed academic documents
- centered/bold official headers and exam titles
- section heading emphasis
- bold `Câu N.` prefixes
- answer indentation
- candidate-field emphasis
- compact footer/page-number styling

It does not rewrite the mathematical meaning of the document.

## Difficult diagrams and tables

For `figure`, `table_image` or `unresolved` OCR blocks, Gemini returns a normalized bounding box. When preservation is enabled, the add-in crops that region from the original image and embeds it in Word. This is intended for content such as geometry diagrams and variation tables where converting the visual structure to plain text would be less reliable.

If the crop cannot be created, Word inserts a yellow `[CẦN KIỂM TRA: ...]` placeholder instead of silently losing the content.

## API-key behavior

The key is **not** stored in source code, `.env`, registry plaintext, or a backend config file. The Settings form encrypts it with Windows DPAPI (`CurrentUser`) before writing the local settings file under:

```text
%LOCALAPPDATA%\WordGeminiFormula\settings.json
```

## Diagnostics

If Word reports that the add-in failed to load, inspect:

```powershell
Get-Content "$env:LOCALAPPDATA\WordGeminiFormula\addin-startup.log"
```

## Uninstall

Close Word, run PowerShell as Administrator, then:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\uninstall.ps1
```

## Security notes

- The Gemini API key is sent only to Google's Gemini API endpoint over HTTPS.
- The add-in makes no requests to a custom backend.
- Do not commit `%LOCALAPPDATA%\WordGeminiFormula\settings.json`.
- Treat OCR results as untrusted content and review mathematical expressions before relying on them.

## Status

V0.2 development implementation. GitHub Actions validates build, COM registration and the Office interface IIDs used by Word. Actual document rendering still requires runtime testing in Microsoft Word on Windows.
