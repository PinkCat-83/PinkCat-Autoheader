# 🔧 Technical README — PinkCat Autoheader

> Internal reference for development, debugging, and AI-assisted work.  
> → [Presentation README](./README.md)

---

## 🤖 AI Instructions

- Wait for the author to specify what needs to be done before proceeding.
- Ask for the relevant files before making any modifications.
- Do not refactor across multiple modules without explicit instruction — `word_processor.py`, `gui.py`, and `controller.py` are known candidates for modularization but have not been split yet.
- The app requires Word to be closed before processing — any feature touching `word_processor.py` must preserve this check.
- Configuration persistence is handled by `config_manager.py` via `config.ini` — do not add persistent state elsewhere.

---

## 1. Project Structure

```
autopdf/
├── main.py                 # Entry point
├── config.ini              # Persistent configuration
└── src/
    ├── config.py           # Constants and styles
    ├── config_manager.py   # config.ini read/write
    ├── utils.py            # General utilities
    ├── file_manager.py     # File operations
    ├── word_processor.py   # Word/PDF processing (complex — refactor pending)
    ├── gui.py              # Tkinter GUI (complex — refactor pending)
    └── controller.py       # Business logic / MVC controller (complex — refactor pending)
```

---

## 2. Processing Flow

1. User configures logo, author, banned words, and folders.
2. On EMPEZAR, app verifies Word is closed.
3. **Renaming phase** (if enabled): renames files using folder code pattern detection.
4. **Processing phase**: processes Word documents according to selected options.
5. Outputs results to destination folder.

---

## 3. Filtering System

Two independent filter lists:

| List | Effect |
|---|---|
| **Non-processable** | File is copied but not modified and not converted to PDF |
| **Non-copyable / ignored** | File or folder is completely ignored — not copied, not processed, not renamed |

**Priority:** non-copyable takes precedence. A file not in that list is either processed, copied as annex, or renamed depending on active options.

---

## 4. Renaming System

Triggered only when "Rename with folder code" is active. Only applies to files not in the non-copyable list.

### Pattern detection (≥ 2 dashes in filename)

Extracts the suffix to keep and replaces the prefix with the folder code:

```
Folder: CAL-05-Geometría
DOC-13-Teoremas.docx     →  CAL-05-Teoremas.docx
MAT-10 - Ejercicios.pdf  →  CAL-05 - Ejercicios.pdf
01-02-Resumen.docx       →  CAL-05-Resumen.docx
```

### Manual fallback (< 2 dashes)

Dialog prompts the user to define the root name:

```
Folder: CAL-05-Geometría
Tareas.pdf      →  asks user  →  CAL-05-Tareas.pdf
```

### Special cases

- File already has the correct folder code → not renamed
- File has a different code → prefix replaced with folder code
- User can cancel renaming per individual file during processing

---

## 5. MVC Pattern

| Layer | File | Responsibility |
|---|---|---|
| Model | `config.py`, `config_manager.py`, `file_manager.py`, `word_processor.py` | Data, configuration, file operations, Word automation |
| View | `gui.py` | Tkinter GUI, drag & drop, checkboxes, progress display |
| Controller | `controller.py` | Business logic, orchestrates model and view |

---

## 6. Dependencies

| Package | Purpose |
|---|---|
| `pywin32` | Word COM automation (required for header/footer insertion and PDF export) |
| `psutil` | Detect if Word is running before processing starts |
| `Pillow` | Logo image handling |
| `tkinterdnd2` | Drag & drop support in the GUI |

---

## 7. Pending Tasks

- [ ] **Refactor** `word_processor.py`, `gui.py`, and `controller.py` into smaller modules
- [ ] **Export log** — generate an external `.txt` with the full processing history
- [ ] **Drag & drop for destination folder** — currently only available for source folders
