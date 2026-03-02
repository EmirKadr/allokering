# CLAUDE.md — AI Assistant Guide for `allokering`

## Project Overview

**Buffertpallar → Order-allokering (GUI)** is a Swedish warehouse management desktop application that allocates customer orders to buffer (inventory) pallets. It calculates optimal pallet space usage, refill requirements, and generates reports for warehouse operators.

**Current version:** 9.2 (file: `allokera9.2.py`)
**Language:** Swedish (all UI text, comments, and docstrings)
**Runtime:** Python 3.9+ desktop application

---

## Repository Structure

```
allokering/
├── allokera9.2.py                      # Entire application — single file (~2,500 lines)
└── .github/
    └── workflows/
        └── auto-merge-claude.yml       # Auto-merges claude/** branches into main
```

This is a **single-file monolith**. All logic, GUI, schemas, and constants live in `allokera9.2.py`. There are no separate `src/`, `tests/`, or `config/` directories.

---

## Technology Stack

| Component        | Technology                              |
|------------------|-----------------------------------------|
| GUI              | `tkinter` / `ttk` (stdlib)              |
| Data Processing  | `pandas`, `numpy`                       |
| Excel I/O        | `openpyxl` (read), `openpyxl`/`xlsxwriter` (write, optional) |
| Drag-and-drop    | `tkinterdnd2` (optional)                |
| Language         | Python 3.9+                             |
| CI/CD            | GitHub Actions                          |

Optional dependencies are detected at import time; the app degrades gracefully if they're missing.

---

## Running the Application

```bash
# Install dependencies
pip install pandas numpy openpyxl tkinterdnd2

# Launch
python3 allokera9.2.py
```

No build step, no tests, no configuration files — just run the Python file.

---

## Key Constants (top of `allokera9.2.py`)

```python
APP_TITLE = "Buffertpallar → Order-allokering (GUI) — 8.5"
DEFAULT_OUTPUT = "allocated_orders.csv"
ALLOC_BUFFER_STATUSES = {29, 30, 32}   # Valid statuses for allocation
REFILL_BUFFER_STATUSES = {29, 30}      # Valid statuses for refill
NEAR_MISS_PCT = 0.30                   # 30% threshold for near-miss pallets
INVALID_LOC_PREFIXES = ("AA",)         # Warehouse location prefixes to exclude
INVALID_LOC_EXACT = {"TRANSIT", "TRANSIT_ERROR", "MISSING", "UT2"}
```

These are the primary configuration points. There are no external config files or `.env` files.

---

## Code Architecture

### Input Schemas (flexible column mapping)

Each schema defines candidate column names for fuzzy, case-insensitive matching via `find_col()`:

| Schema            | Purpose                                 |
|-------------------|-----------------------------------------|
| `ORDER_SCHEMA`    | Customer order lines (article, qty, status, ordernr, radnr) |
| `BUFFER_SCHEMA`   | Buffer/pallet inventory (artikel, qty, location, timestamp, id, status) |
| `NOT_PUTAWAY_SCHEMA` | Articles not yet put away           |
| `SALDO_SCHEMA`    | Current inventory balance              |
| `ITEM_SCHEMA`     | Item master data (stackability flags)  |

`find_col(df, candidates)` matches columns case-insensitively, stripping non-alphanumeric characters. This tolerates varied input file formats.

### Core Processing Functions

| Function                          | Description                                          |
|-----------------------------------|------------------------------------------------------|
| `read_prognos_xlsx(path)`         | Parse forecast XLSX (strips header rows/columns)     |
| `read_campaign_xlsx(path)`        | Parse campaign volume XLSX                           |
| `normalize_not_putaway(path)`     | Process articles not yet put away                    |
| `normalize_saldo(path)`           | Normalize inventory balance CSV                      |
| `normalize_pick_log(path)`        | Process pick log and compute sales metrics           |
| `normalize_items(path)`           | Load item master data (stackability)                 |
| `compute_sales_metrics(df)`       | ABC classification and sales analytics               |
| `allocate(orders_df, buffer_df)`  | Core FIFO allocation algorithm                       |
| `calculate_refill(buffer_df, ...)`| Calculate replenishment requirements                 |
| `compute_pallet_spaces(...)`      | Required pallet spaces per customer                  |
| `build_prognos_vs_autoplock_report(...)` | Forecast vs. autopick comparison report      |

### GUI Class

`App(ttk.Frame)` — single main window with:
- File picker buttons (`pick_*()` methods) for all input files
- `run_allocation()` — main entry point that orchestrates all processing
- `open_*_in_excel()` — exports result DataFrames to temp files, opens in system Excel
- `reset_cache()` — clears cached DataFrames

### Utility Helpers

```python
find_col(df, candidates, required=True, default=None)  # Fuzzy column finder
to_num(x) -> float                                     # Safe numeric conversion
_safe_str_series(s: pd.Series) -> pd.Series            # Safe string conversion
```

---

## Allocation Logic

### Three-Stage Priority

1. **HELPALL** — Manual picking zone pallets (highest priority)
2. **AUTOSTORE** — Automated bin storage pallets
3. **HAUPTPLOCK** — Main picking area pallets (fallback)

### Zone Classification

| Code | Zone Name      | Description                    |
|------|----------------|--------------------------------|
| `H`  | HELPALL        | Manual warehouse picking       |
| `A`  | AUTOSTORE      | Automated bin storage          |
| `R`  | REFILL         | Replenishment zone             |
| `F`  | HIB            | Special handling               |
| `S`  | SKRYMMANDE     | Bulky/overflow items           |
| `E`  | EHANDEL        | E-commerce zone                |
| `Q`, `O` | —         | Miscellaneous special zones    |
| `D`  | DISPLAY        | Display picking                |

### FIFO Allocation

- Pallets sorted by received timestamp
- Cumulative sum used to determine how many pallets cover an order
- **Near-miss detection**: a pallet is flagged as near-miss if it covers 70–100% of remaining demand (within `NEAR_MISS_PCT = 0.30`)

---

## Data Flow

```
Input CSVs/XLSXs
  (orders, buffer, saldo, items, forecast, campaign)
         ↓
  normalize_* functions
  (fuzzy column matching, type coercion, filtering)
         ↓
  allocate() — FIFO allocation per article
         ↓
  calculate_refill(), compute_pallet_spaces(),
  build_prognos_vs_autoplock_report()
         ↓
  Temporary CSV/XLSX files → opened in system Excel
```

No database is used. All processing is in-memory with pandas DataFrames.

---

## Git Workflow

### Branch Convention

- **Main branch:** `master` (auto-merge target)
- **AI branches:** `claude/<description>-<session-id>` (e.g., `claude/fix-bug-abc123`)

### Auto-merge CI

Pushing to any `claude/**` branch triggers `.github/workflows/auto-merge-claude.yml`, which:
1. Merges the branch into `main` with `--no-ff`
2. Pushes the result automatically

**Always push Claude-generated changes to a `claude/` prefixed branch.** The auto-merge workflow handles integration into `main`.

```bash
git push -u origin claude/<your-branch-name>
```

---

## Versioning Convention

The version is stored in `APP_TITLE` as a string constant — there is no separate version file or `pyproject.toml`. When incrementing versions, update `APP_TITLE` and rename the file (e.g., `allokera9.2.py` → `allokera9.3.py`).

---

## Language Note

All user-facing text, comments, docstrings, and error messages are written in **Swedish**. When adding new UI strings, error messages, or comments, use Swedish to match the existing codebase style.

---

## What Does Not Exist (don't create unless requested)

- No test suite — no pytest, unittest, or test files
- No `requirements.txt` or `pyproject.toml`
- No database or migrations
- No CLI interface
- No API or server component
- No `.env` or external configuration files
- No separate modules or packages — keep everything in the single `.py` file
