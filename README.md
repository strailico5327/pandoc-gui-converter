# Pandoc Converter

Pandoc Converter is a small desktop Python GUI for converting Markdown files to
Word DOCX files and DOCX files back to Markdown with Pandoc.

## About Pandoc

This tool is a simple GUI wrapper for Pandoc, a universal document converter. Pandoc must be installed separately before using this program.

## Features

- Convert Markdown-like files (`.md`, `.markdown`, `.mdown`, `.mkd`, `.txt`) to
  `.docx`.
- Convert `.docx` files to `.md`.
- Auto-detect conversion direction from the input extension.
- Choose a custom `pandoc.exe` path from the GUI when Pandoc is not in `PATH`.
- Choose input and output files with file picker buttons.
- Show the generated Pandoc command and conversion log in the window.
- Check the installed Pandoc version from the GUI.
- Optionally remove single line breaks inside ordinary Markdown paragraphs
  after DOCX-to-Markdown conversion.
- Open the output folder after conversion.

## Repository Layout

```text
Pandoc Converter/
  .gitignore
  LICENSE
  README.md
  pandoc_gui_converter.py
  pandoc_gui_converter.spec
  requirements.txt
  template_0.docx
```

`pandoc_gui_converter.py` is the application entry point. The existing
`pandoc_gui_converter.exe` file is a local packaged build artefact and is
ignored by Git.

## Dependencies

- Python 3.9 or newer
- Pandoc installed and available in `PATH`, or selected manually in the GUI
- PySide6

The converter uses Qt controls through `PySide6`.

## Run

Run from the repository root:

```powershell
python -m pip install -r requirements.txt
python pandoc_gui_converter.py
```

## Checks

Check that the script parses correctly:

```powershell
python -m py_compile pandoc_gui_converter.py
```

Open the GUI and use the `Check Pandoc` button to confirm Pandoc is available.

## Git Ignore

The `.gitignore` excludes Python caches, virtual environments, package build
outputs, coverage artefacts, editor folders, OS metadata, local `.exe` builds,
and logs.

## Licence

This project is licensed under the GNU General Public License v3.0.

Copyright (C) 2026 strailico5327.

## Notes

This project was developed with assistance from OpenAI Codex. The code has been reviewed and tested before release.
