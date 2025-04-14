## Overview

This script reads a `.docx` (Word) file, applies a series of hard‑coded text corrections, and writes the result to a new `.docx`. It does **not** use `python-docx`, but instead automates Word via COM (`pywin32`).

### Corrections performed

1. **Normalization**  
   - Converts curly quotes (`“”`) to straight quotes (`"`).  
   - Removes non‑printable characters.  
   - Collapses extra spaces (but preserves line breaks).

2. **Spelling**  
   - In _unquoted_ text only, changes `organize` → `organise`.  
   - In _unquoted_ text only, changes `eg` → `for example`.

3. **Name disambiguation**  
   - Ensures middle initials have a trailing period (e.g. `D` → `D.`).  
   - Tracks occurrences of titles (Dr/Mr/Mrs/Ms) + names:  
     - First mention: full name (`Dr John D. Smith`).  
     - Subsequent mentions: `Title Last` (`Dr Smith`), unless that last name is shared by multiple full names (in which case always use full).

4. **Output prefix**  
   - Prepends the entire corrected text with `Corrected:\n`.

## Requirements

- Windows with Microsoft Word installed.
- Python 3.7+
- `pywin32` package

### Installation

```bash
pip install pywin32
```

## Usage

```bash
python main.py <input.docx> <output.docx>
```

- `<input.docx>`: Path to your source Word document.  
- `<output.docx>`: Path where the corrected document will be saved (will overwrite if exists).

### Example

```bash
python main.py sample.docx corrected_sample.docx
```

