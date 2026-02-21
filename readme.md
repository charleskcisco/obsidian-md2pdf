# MD to PDF - Obsidian Plugin

Convert Markdown files to beautifully styled PDF documents using Pandoc and LibreOffice.

## Features

- **Multiple Class Formats**: Different formatting styles for different classes (Chicago/Turabian, MLA, etc.)
- **Reference Document Styling**: Use custom DOCX templates for styling
- **Title Page**: Automatically creates a title page with author, date, and instructor
- **MLA Header**: Generates proper MLA-style headers for class assignments
- **Automatic Bibliography Detection**: Enables Pandoc's citeproc when bibliography is detected
- **Page Breaks**: Automatically inserts page breaks before Bibliography/References sections
- **Cross-Platform**: Works on Linux, macOS, and Windows
- **Auto-Detection**: Automatically finds Pandoc and LibreOffice installations

## Prerequisites

### 1. Pandoc

- **Windows**: Download from [pandoc.org/installing.html](https://pandoc.org/installing.html)
- **macOS**: `brew install pandoc`
- **Linux**: `sudo apt install pandoc`

### 2. LibreOffice

- **Windows/macOS**: Download from [libreoffice.org](https://www.libreoffice.org/download/)
- **Linux**: `sudo apt install libreoffice`

## Installation

1. Download the plugin ZIP file
2. Extract to `.obsidian/plugins/md2pdf/` in your vault
3. Enable the plugin in Obsidian Settings → Community plugins

## Class Formats

The plugin supports multiple formatting styles. Set the format using `class:` in your YAML frontmatter.

### 12th Grade—Thesis (`class: 12-thesis`)

Chicago/Turabian format. Creates a title page with:
- Title centered at top
- Author and date at bottom
- "Submitted to" before instructor name
- Page break before content

```yaml
---
title: The Case for Taco Tuesday as a Weekly National Holiday
author: Your Name
date: 2025-12-04
teacher: Dr. Cisco
class: 12-thesis
bibliography: references.bib
---

Your content starts here...
```

### 12th Grade—Modern Literature (`class: 12-literature`)

MLA format. Creates a double-spaced header on the first page:
- Your name
- Date (formatted as "Month DD, YYYY")
- Class name
- Teacher name
- Centered title

```yaml
---
title: Analysis of The Great Gatsby
author: Your Name
date: 2025-12-04
teacher: Mrs. Greb
class: 12-literature
bibliography: references.bib
---

Your content starts here...
```

### 12th Grade—Modern History (`class: 12-history`)

Same as Literature but with "12th Grade Modern History" as the class name.

```yaml
---
title: The Causes of World War I
author: Your Name
date: 2025-12-04
teacher: Mr. Johnson
class: 12-history
bibliography: references.bib
---

Your content starts here...
```

## Configuration

Go to **Settings → MD to PDF** to configure:

| Setting | Description |
|---------|-------------|
| **Default Class Format** | Format to use when no `class:` is specified |
| **Pandoc Path** | Auto-detected, or set manually |
| **LibreOffice Path** | Auto-detected, or set manually |
| **Custom Pandoc Arguments** | Additional arguments to pass to pandoc |
| **Override reference.docx** | Override all formats with a custom template |
| **Auto-detect Bibliography** | Enable citeproc when `bibliography:` is found |
| **Delete Intermediate .docx** | Clean up temp files after conversion |

## Usage

### Method 1: Command Palette
1. Open a Markdown file
2. Press `Ctrl+P` (Windows/Linux) or `Cmd+P` (macOS)
3. Type "MD to PDF" and select:
   - **Convert current Markdown to PDF** — uses class from YAML
   - **Convert to PDF (select class format)** — choose format from dropdown

### Method 2: Ribbon Icon
Click the document icon in the left ribbon.

## Troubleshooting

### "Pandoc not found"
- Windows: `where pandoc`
- macOS/Linux: `which pandoc`
- Enter the path manually in settings

### "LibreOffice not found"
- Windows: Usually `C:\Program Files\LibreOffice\program\soffice.exe`
- macOS: `/Applications/LibreOffice.app/Contents/MacOS/soffice`
- Linux: `/usr/bin/libreoffice`

### Conversion fails
- Press `Ctrl+Shift+I` to open developer console
- Check for error messages

