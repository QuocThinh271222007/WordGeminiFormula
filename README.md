# Word Gemini Formula

A standalone Microsoft Word add-in that converts images of Vietnamese/math documents into editable Word text and native Word equations using the Gemini API.

## Features

- **Ảnh → Word**: OCR the picture currently selected in Word; if no picture is selected, choose a PNG/JPG/WEBP/BMP/GIF file.
- **Settings inside Word**: enter the Gemini API key and model without editing backend/source files.
- **Protected API key**: stored locally with Windows DPAPI for the current Windows user.
- **Review-first formula workflow**: OCR formulas are inserted as `[[MATH]]...[[/MATH]]` by default so you can verify them before conversion.
- **Chuẩn hóa tất cả**: converts all marked formulas into native editable Word Equation objects and calls Word's Professional `BuildUp()` conversion.
- **Chuẩn hóa vùng chọn**: converts only the selected LaTeX/linear formula.
- Optional **auto-normalize after OCR** setting.

## Requirements

- Windows 10/11
- Microsoft Word desktop (Microsoft 365 / Word 2019+ recommended)
- .NET Framework 4.8
- Visual Studio 2022 or Build Tools with MSBuild for building from source
- A Gemini Developer API key

The default model is `gemini-3.7-flash`, and the model field is editable in Settings.

## Build

Open `WordGeminiFormula.sln` in Visual Studio 2022 and build **Release**, or from Developer PowerShell:

```powershell
msbuild WordGeminiFormula.sln /t:Restore,Build /p:Configuration=Release /p:Platform="Any CPU"
```

Output:

```text
src\WordGeminiFormula.AddIn\bin\Release\net48\WordGeminiFormula.AddIn.dll
```

## Install for the current Windows user

Close all Word windows, then run:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\install.ps1
```

Open Word again. A new Ribbon tab named **AI Formula** should appear.

No administrator rights are intended for the default per-user registration path.

## First use

1. Open Word → **AI Formula** → **Settings**.
2. Paste your Gemini API key.
3. Choose a model. The default is `gemini-3.7-flash`.
4. Click **Test API**.
5. Click **Lưu**.
6. Select an image already pasted in Word and click **Ảnh → Word**, or click it with no image selected to choose an image file.
7. Review any `[[MATH]]...[[/MATH]]` blocks.
8. Click **Chuẩn hóa tất cả** to turn them into native Word equations.

## API-key behavior

The key is **not** stored in source code, `.env`, registry plaintext, or a backend config file. The Settings form encrypts it with Windows DPAPI (`CurrentUser`) before writing the local settings file.

## Uninstall

Close Word and run:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\uninstall.ps1
```

## Why COM instead of a web Office Add-in?

This V1 needs direct access to Word's native `OMath` model and `BuildUp()` behavior. A desktop COM add-in gives direct control of those Word equation objects while keeping the Gemini request client-side.

## Security notes

- The Gemini API key is sent only to Google's Gemini API endpoint over HTTPS.
- The add-in makes no requests to a custom backend.
- Do not commit `%LOCALAPPDATA%\WordGeminiFormula\settings.json`.
- Treat OCR results as untrusted content and review mathematical expressions before relying on them.

## Status

V0.1 source implementation. The repository includes a Windows GitHub Actions build workflow, but runtime testing must be performed with Microsoft Word on Windows.
