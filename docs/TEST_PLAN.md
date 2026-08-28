# V0.1 Windows / Word test plan

## Build and registration

- Build `Release` on Windows with .NET Framework 4.8 reference assemblies.
- Run `scripts/install.ps1` as a normal user.
- Restart Word and verify the **AI Formula** tab loads without Word disabling the add-in.
- Run `scripts/uninstall.ps1`, restart Word, and verify the tab is removed.

## Settings

- Open Settings with no existing settings file.
- Verify API key is masked by default and the show-key checkbox works.
- Test a valid key and an invalid key.
- Save settings and restart Word; verify model/auto-normalize persist.
- Verify `%LOCALAPPDATA%\\WordGeminiFormula\\settings.json` does not contain the plaintext API key.

## OCR

- OCR a local JPEG and PNG containing Vietnamese prose and math.
- Select an inline image already pasted into Word and OCR it.
- Select a floating image and OCR it.
- Verify an image above the inline size limit gives a clear error.
- Verify API/network failures do not delete or corrupt existing document content.

## Formula normalization

Verify at minimum:

- `x^2 + y_1`
- `\\frac{x^2+1}{x-1}`
- `\\sqrt{x+1}`
- `\\sqrt[3]{x}`
- `\\sum_{i=1}^{n} i`
- `\\int_0^1 x^2 dx`
- `\\lim_{x\\to0} \\frac{\\sin x}{x}`
- `\\vec{AB}`
- `\\begin{pmatrix}a&b\\\\c&d\\end{pmatrix}`

For each, verify the result is a native editable Word equation and `BuildUp()` renders Professional format.

## Regression / safety

- A document with no `[[MATH]]` blocks must remain unchanged when Normalize All is clicked.
- Normalize Selection must refuse an empty selection.
- Normal prose containing square brackets but not exact math markers must remain unchanged.
- Formula conversion should proceed from the end of the document toward the beginning so range offsets do not invalidate later replacements.
