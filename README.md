# 📄 PinkCat Autoheader

Windows desktop application that automates bulk conversion of `.docx`/`.docm` documents to PDF, inserting custom headers (logo and course code) and footers (author and page numbering).

![PinkCat Autoheader interface](img/002.png)

---

## What is this?

At the start of each course, every document in the course materials needs updated headers — logo, course code, author, and page numbers. PinkCat Autoheader processes entire folder structures in one pass, handles renaming with folder codes, filters out draft or non-processable files, and preserves the original folder structure in the output.

---

## Features

- Bulk processing of `.docx` and `.docm` files
- Automatic header insertion: logo, course code, decorative lines
- Automatic footer insertion: author name, page numbering
- Banned-word filtering (non-processable and non-copyable files)
- Automatic renaming with folder code (pattern detection + manual fallback)
- Optional: preserve original folder structure in output
- Drag & drop folder selection
- Persistent configuration via `config.ini`

---

## Requirements

- Windows 10/11
- Microsoft Word (Office 365 desktop) — **must be closed before processing**
- Python 3.7+

```bash
pip install pywin32 psutil Pillow tkinterdnd2
```

---

## Usage

1. Configure format options via checkboxes (logo, decorative lines, author, code, numbering, folder structure, PDF conversion, renaming).
2. Configure filters: banned words for non-processable files, and words for files/folders to ignore completely.
3. Select source folders via drag & drop or the folder picker button.
4. Set the destination folder.
5. Press **EMPEZAR**.

---

## Filtering System

### Non-processable files
Files matching these words are **copied but not modified**:
- No headers, footers, or format changes
- No PDF conversion

### Non-copyable / ignored files and folders
Files and folders matching these words are **completely ignored**:
- Not copied, not processed, not renamed

**Priority rule:** a file not in the "non-copyable" list will be either processed (if `.docx`/`.docm` and not non-processable), copied as an annex, or renamed — depending on options.

---

## Automatic Renaming System

When "Rename with folder code" is active, the app prefixes each file with the folder's code.

### Pattern detection (two or more dashes)

The system extracts the part of the name to keep:

```
Folder: CAL-05-Geometría

DOC-13-Teoremas.docx     →  CAL-05-Teoremas.docx
MAT-10 - Ejercicios.pdf  →  CAL-05 - Ejercicios.pdf
```

### Manual fallback (fewer than two dashes)

A dialog asks the user to define the root of the name:

```
Folder: CAL-05-Geometría

Tareas.pdf       →  asks user  →  CAL-05-Tareas.pdf
```

### Special cases

- File already has the correct code → not renamed
- File has a different code → replaced with the folder code
- Renaming can be cancelled per file during processing

---

## Technical Documentation

For architecture, module responsibilities, processing flow, and AI instructions, see the **[Technical README](./README_TECH.md)**.
