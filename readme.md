# MD to PDF — Obsidian Plugin

Convert Markdown files to styled PDF or DOCX documents using Pandoc with Typst or LibreOffice. Supports MLA and Chicago/Turabian formatting via YAML frontmatter. The export pipeline matches [Folio](https://github.com/charleskcisco/folio-software)'s.

## Features

- **MLA & Chicago/Turabian styles**: Automatically generates the correct header or cover page from YAML frontmatter, including an optional `subtitle:`
- **Two PDF engines**: Typst (bundled MLA/Turabian templates and fonts, no LibreOffice needed) or LibreOffice via `.docx` reference documents
- **PDF or DOCX export**: Separate commands for each
- **Insert frontmatter**: One command adds every export field the note is missing
- **Custom reference documents**: Control fonts, margins, and line spacing using `.docx` templates
- **Bibliographies**: Enables Pandoc's `citeproc` when a `bibliography:` field is found, resolving the `.bib` against the vault (falling back to `sources/*.bib`) and passing only the cited entries to citeproc for speed. Without `csl:`, `style: chicago` uses Chicago notes-bibliography and `style: mla` uses MLA (both bundled)
- **Hand-typed bibliographies**: A Bibliography/References/Works Cited section typed by hand gets a page break and a hanging indent
- **Tables**: Column widths fit their contents rather than the dash counts in the markdown
- **Configurable output locations**: Save PDFs and intermediate `.docx` files to custom folders
- **Cross-platform**: Works on macOS, Linux, and Windows
- **Auto-detection**: Finds Pandoc, Typst and LibreOffice automatically on first load
- **After export**: Offers to open the file or show it in its folder; failures write the full command and output to `.md2pdf/export-error.log` in the vault

## Prerequisites

### Pandoc

- **macOS**: `brew install pandoc`
- **Linux**: `sudo apt install pandoc`
- **Windows**: Download from [pandoc.org/installing.html](https://pandoc.org/installing.html)

### Typst (recommended for PDF)

- **macOS**: `brew install typst`
- **Linux**: `cargo install typst-cli`, or your distribution's package
- **Windows**: `winget install --id Typst.Typst`

### LibreOffice (optional with Typst; needed for custom reference documents)

- **macOS/Windows**: Download from [libreoffice.org](https://www.libreoffice.org/download/)
- **Linux**: `sudo apt install libreoffice`

## Installation

1. Download the plugin files (`main.js`, `manifest.json`, `styles.css`, the `.docx` reference documents, and the `templates/`, `csl/` and `fonts/` folders)
2. Copy them to `.obsidian/plugins/obsidian-md2pdf/` inside your vault
3. Enable the plugin in **Settings → Community plugins**

## Usage

### Command Palette

1. Open a Markdown file
2. Press `Cmd+P` (macOS) or `Ctrl+P` (Windows/Linux)
3. Run **Export current note to PDF** or **Export current note to DOCX**

Run **Insert export frontmatter** to add `title`, `subtitle`, `author`, `instructor`, `course`, `date`, `spacing` and `style` to the note (only the ones it lacks). Assign hotkeys to any of these under **Settings → Hotkeys**.

### Ribbon Icon

Click the document icon in the left sidebar ribbon.

## YAML Frontmatter

Control the document style and template via frontmatter fields.

### `style` — document structure

| Value | Description |
|-------|-------------|
| `chicago` / `turabian` | Turabian-style cover page: title centered ~1/3 down, author/course/instructor/date in the bottom half |
| `mla` | MLA-style info block: author, instructor, course, date left-aligned at top, then centered title |
| *(omitted)* | Plain document — no special header or cover page |

### `spacing` — typography template

Place `.docx` reference files in the plugin folder (`.obsidian/plugins/obsidian-md2pdf/`) and reference them by filename without the extension.

```yaml
spacing: double   # uses double.docx
spacing: single   # uses single.docx
```

`single` and `double` work with both engines. Any other value names a reference document that only LibreOffice can honour, so those notes always export through LibreOffice.

The reference document controls fonts, margins, line spacing, and running headers/footers.

### Full YAML Examples

**Chicago/Turabian cover page:**
```yaml
---
title: My Thesis Title
subtitle: An Optional Second Line
author: Jane Smith
course: English 401
instructor: Dr. Johnson
date: 2026-03-07
style: chicago
spacing: double
bibliography: references.bib
---
```

**MLA header:**
```yaml
---
title: Literary Analysis of The Great Gatsby
author: Jane Smith
instructor: Mrs. Williams
course: 12th Grade Literature
date: 2026-03-07
style: mla
spacing: double
---
```

**Plain document:**
```yaml
---
title: Quick Notes
spacing: single
---
```

## Configuration

Go to **Settings → md2pdf** to configure:

| Setting | Description |
|---------|-------------|
| **PDF Engine** | Typst or LibreOffice. Installs from before 1.3.0 keep LibreOffice; Typst falls back to LibreOffice when it is not installed |
| **Default Reference Document** | Template to use when no `spacing:` is set in YAML |
| **Refresh Reference Documents** | Rescan the plugin folder for `.docx` files |
| **Pandoc Path** | Auto-detected, or set manually |
| **Typst Path** | Auto-detected, or set manually |
| **LibreOffice Path** | Auto-detected, or set manually |
| **Auto-detect Paths** | Re-run auto-detection (clears current paths first) |
| **Custom Pandoc Arguments** | Additional arguments to pass to pandoc (space-separated) |
| **Auto-detect Bibliography** | Enable citeproc automatically when `bibliography:` is in frontmatter |
| **Offer to Open Export** | After exporting, offer to open the file or show it in its folder |
| **Delete Intermediate .docx** | Remove the temporary `.docx` file after a LibreOffice PDF is created |
| **PDF Output Location** | Where to save PDFs — leave empty for same folder as source, or set a vault-relative or absolute path |
| **DOCX Output Location** | Where to save `.docx` exports, and intermediate `.docx` files when not deleted |

## Adding Reference Documents

Place any `.docx` file into `.obsidian/plugins/obsidian-md2pdf/`. The filename (without extension) is used as the `spacing` value in your frontmatter. Running headers/footers in the reference doc can use the `{{LASTNAME}}` placeholder, which is replaced automatically from the `author` field.

Click **Refresh Reference Documents** in settings after adding new files.

## Troubleshooting

### "Pandoc not found" / "LibreOffice not found"

Use **Auto-detect** in settings, or find the path manually:

- macOS/Linux: `which pandoc`, `which libreoffice`
- Windows: `where pandoc`, `where soffice`

Then paste the path into the appropriate field in plugin settings.

### Conversion fails

Read `.md2pdf/export-error.log` in your vault (it is hidden from the file explorer; open it from your file manager), or open the developer console (`Ctrl+Shift+I` / `Cmd+Option+I`) and check for error messages in the Console tab.
