# AGENTS.md — Data Dictionary Skill

## Purpose

This file defines the **agent skill** for converting a `ydata_profiling` JSON output
into the exact styled Excel data dictionary established in this project.

When invoked, an agent following this spec must:

1. Accept a `ydata_profiling` JSON blob (string or parsed dict) and the source `pd.DataFrame`.
2. Parse every variable's stats into a flat, tidy table (`dict_df`).
3. Write a fully styled `.xlsx` workbook matching the visual spec below — pixel-perfect colours,
   two-level header, frozen panes, type badges, heat maps, and gradients.

The canonical implementation lives in **`data_dictionary.py`**.
Use that file as ground truth if any detail below is ambiguous.

---

## Inputs

| Name | Type | Description |
|---|---|---|
| `profile_json` | `str` or `dict` | Output of `ProfileReport(df).to_json()` — parse with `json.loads()` if a string |
| `df` | `pd.DataFrame` | The original DataFrame (used only to look up `dtype` per column) |
| `out_path` | `str` / `Path` | Destination `.xlsx` file |

---

## Pipeline overview

```
profile_json
     │
     ▼
[Step 1] Extract  profile_json["variables"]
               → one stats-dict per variable name
     │
     ▼
[Step 2] Parse each stats-dict
               → one flat row per variable (43 columns)
               → collect into dict_df  (index = Variable name)
     │
     ▼
[Step 3] Build openpyxl Workbook
               → Row 1:  group-label header  (merged cells)
               → Row 2:  column-name header
               → Row 3+: one styled data row per variable
     │
     ▼
[Step 4] Apply dimensions + freeze → save
```

---

## Step 1 — JSON structure reference

`profile_json["variables"]` is a `dict` keyed by **column name**.
Each value is a stats-dict whose keys depend on the profiling type.

### Keys present for ALL types

```
n                   total row count (including missing)
n_missing           count of null values
p_missing           fraction [0, 1] of null values
n_distinct          count of distinct values
p_distinct          fraction [0, 1]
n_unique            count of values that appear exactly once
p_unique            fraction [0, 1]
is_unique           bool — True if every non-null value appears once
memory_size         bytes consumed by the column
type                string: "Numeric" | "Text" | "Categorical" | "DateTime" | "Boolean" | "Unsupported"
value_counts_without_nan   ordered dict {value: count} — first key is the most frequent
```

### Additional keys for `type == "Numeric"`

```
mean  std  variance  cv  skewness  kurtosis
min   max  range     sum mad
5%    25%  50%       75%  95%   iqr
n_zeros   p_zeros
n_negative  p_negative
n_infinite  p_infinite
monotonic   (float: 0=none, 1=increase, 2=decrease, …)
```

### Additional keys for `type == "Text"`

```
min_length   max_length   mean_length   median_length
n_characters          total character count
n_characters_distinct count of distinct characters
```

---

## Step 2 — Parsing rules (JSON → dict_df)

Apply the rules below for every variable in `profile_json["variables"]`.
The result is one row in `dict_df`.

### Safe-access helpers (implement or import from `data_dictionary.py`)

```python
def _get(d, key, default=np.nan):
    v = d.get(key)
    return default if v is None else v

def _top(vc_dict):
    # returns (most_frequent_value, its_count) or (np.nan, np.nan)
    if isinstance(vc_dict, dict) and vc_dict:
        k, v = next(iter(vc_dict.items()))
        return k, int(v)
    return np.nan, np.nan

def _pct(v):
    # converts a [0,1] fraction to a % rounded to 4 dp
    return round(float(v) * 100, 4) if not (isinstance(v, float) and np.isnan(v)) else np.nan

def _r(v, dp=4):
    # rounds to dp decimal places; NaN on failure
    try:
        f = float(v)
        return round(f, dp) if not np.isnan(f) else np.nan
    except Exception:
        return np.nan

def _safe_int(stats_dict, key):
    v = stats_dict.get(key)
    return int(v) if v is not None and not isinstance(v, float) else np.nan
```

### Column-by-column extraction rules

Every output column and its extraction rule, in order:

| Output column | JSON key(s) | Transformation |
|---|---|---|
| `Type` | `type` | raw string |
| `Pandas dtype` | — | `str(df[col].dtype)` from source DataFrame |
| `Memory (bytes)` | `memory_size` | `_safe_int` |
| `Total Rows` | `n` | `int(n)` |
| `Count (non-null)` | `n`, `n_missing` | `int(n) - int(n_missing)` |
| `# Missing` | `n_missing` | `_safe_int` |
| `% Missing` | `p_missing` | `_pct` |
| `# Distinct` | `n_distinct` | `_safe_int` |
| `% Distinct` | `p_distinct` | `_pct` |
| `# Unique (exact)` | `n_unique` | `_safe_int` |
| `% Unique (exact)` | `p_unique` | `_pct` |
| `Is Unique` | `is_unique` | `bool(v)` if not float, else `np.nan` |
| `Mean` | `mean` | `_r` |
| `Median (50%)` | `50%` | `_r` |
| `Std Dev` | `std` | `_r` |
| `Variance` | `variance` | `_r` |
| `MAD` | `mad` | `_r` |
| `CV` | `cv` | `_r` |
| `IQR` | `iqr` | `_r` |
| `Skewness` | `skewness` | `_r` |
| `Kurtosis` | `kurtosis` | `_r` |
| `Min` | `min` | raw (preserve original type) |
| `5th Pct` | `5%` | `_r` |
| `25th Pct` | `25%` | `_r` |
| `75th Pct` | `75%` | `_r` |
| `95th Pct` | `95%` | `_r` |
| `Max` | `max` | raw (preserve original type) |
| `Range` | `range` | `_r` |
| `Sum` | `sum` | `_r(v, dp=2)` |
| `# Zeros` | `n_zeros` | `_safe_int` |
| `% Zeros` | `p_zeros` | `_pct` |
| `# Negative` | `n_negative` | `_safe_int` |
| `% Negative` | `p_negative` | `_pct` |
| `# Infinite` | `n_infinite` | `_safe_int` |
| `Monotonic` | `monotonic` | raw float |
| `Min Length` | `min_length` | `_safe_int` |
| `Max Length` | `max_length` | `_safe_int` |
| `Mean Length` | `mean_length` | `_r` |
| `Median Length` | `median_length` | `_r` |
| `# Characters` | `n_characters` | `_safe_int` |
| `# Distinct Chars` | `n_characters_distinct` | `_safe_int` |
| `Top Value` | `value_counts_without_nan` | `_top()[0]` — first key of the dict |
| `Top Freq` | `value_counts_without_nan` | `_top()[1]` — first value of the dict |

**NaN convention:** Keys absent from a stats-dict (e.g., `mean` on a Text variable) produce
`np.nan`.  In the Excel output these cells are blank (no value written, no fill override).

**dtype coercion before writing to openpyxl:** Convert numpy scalars with `.item()`.
Write `None` (not `np.nan`) for missing values — openpyxl renders `None` as an empty cell.

```python
val = None if pd.isna(raw) else (raw.item() if hasattr(raw, "item") else raw)
```

---

## Step 3 — Excel workbook specification

### Sheet settings

| Property | Value |
|---|---|
| Sheet name | `"Data Dictionary"` |
| Tab colour | `1d4ed8` (blue) |
| Freeze panes | `"B3"` — locks column A and rows 1–2 simultaneously |

### Row layout

```
Row 1 : Group-label header  (merged cells, one per group)
Row 2 : Column-name header
Row 3 : Variable row 1
Row 4 : Variable row 2
  …
```

Column A always holds the variable name (`index` of `dict_df`).
Data columns start at column B (Excel index 2).

---

### Colour palette (all hex, no leading `#`)

```python
PALETTE = {
    "hdr_bg":  "0f172a",   # deep navy  — both header rows
    "hdr_fg":  "cbd5e1",   # slate-300  — column-name text
    "grp_fg":  "93c5fd",   # blue-300   — group-label text
    "idx_bg":  "1e293b",   # dark slate — Variable column fill
    "idx_fg":  "e2e8f0",   # slate-200  — Variable column text
    "odd_bg":  "f8fafc",   # data row odd (default fill)
    "even_bg": "ffffff",   # data row even (default fill)
    # type badge fills
    "num_bg":  "1d4ed8",   # Numeric  → blue
    "txt_bg":  "0f766e",   # Text     → teal
    "dt_bg":   "7c3aed",   # DateTime → violet
    "bool_bg": "b45309",   # Boolean  → amber
    "cat_bg":  "15803d",   # Categorical → green
    "unk_bg":  "64748b",   # all other types → slate
}
```

### Group-row background colours (cycling)

```python
GROUP_BG = [
    "1e3a5f", "164e3f", "3b1f63", "4a1c1c", "1c3a4a",
    "2a2a10", "1a3a2a", "1f2a40", "2f1f40", "1a3030", "10302a",
]
# gi = group index (0-based); background = GROUP_BG[gi % len(GROUP_BG)]
```

---

### Column groups (ordered — determines column order in the output)

```python
COLUMN_GROUPS = [
    ("Identity",          ["Type", "Pandas dtype", "Memory (bytes)"]),
    ("Completeness",      ["Total Rows", "Count (non-null)", "# Missing", "% Missing"]),
    ("Uniqueness",        ["# Distinct", "% Distinct", "# Unique (exact)", "% Unique (exact)", "Is Unique"]),
    ("Central Tendency",  ["Mean", "Median (50%)"]),
    ("Spread",            ["Std Dev", "Variance", "MAD", "CV", "IQR"]),
    ("Shape",             ["Skewness", "Kurtosis"]),
    ("Range / Quantiles", ["Min", "5th Pct", "25th Pct", "75th Pct", "95th Pct", "Max", "Range"]),
    ("Aggregate",         ["Sum"]),
    ("Special Counts",    ["# Zeros", "% Zeros", "# Negative", "% Negative", "# Infinite", "Monotonic"]),
    ("Text Metrics",      ["Min Length", "Max Length", "Mean Length", "Median Length",
                           "# Characters", "# Distinct Chars"]),
    ("Most Frequent",     ["Top Value", "Top Freq"]),
]
```

**Absent columns are silently skipped** — if a column named in a group does not exist in
`dict_df`, omit it from the span count and do not write it.

---

### Row 1 — Group-label header

- **Column A, rows 1–2**: merge with `ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=1)`.
  Write `"VARIABLE"` (uppercase) to the top-left cell only.

- For each group (left to right): merge the span of present columns in row 1.
  Write the group name (uppercase) **only to the top-left cell of the merged range** —
  openpyxl merged cells are read-only except at their origin.

- After merging, never write to non-origin cells within a merged range.

| Property | Value |
|---|---|
| Fill | `GROUP_BG[gi % 11]` (cycling) |
| Font | Calibri 9 pt, colour `grp_fg`, bold |
| Alignment | center / center |
| Border | thin, colour `d0d7de` on all four sides |
| Row height | 20 pt |

"VARIABLE" cell (A1:A2):

| Property | Value |
|---|---|
| Fill | `hdr_bg` |
| Font | Calibri 9 pt, colour `grp_fg`, bold |
| Alignment | center / center |
| Border | thin `d0d7de` |

---

### Row 2 — Column-name header

One cell per column in `dict_df.columns` (starting at column B).

| Property | Value |
|---|---|
| Fill | `hdr_bg` |
| Font | Calibri 9 pt, colour `hdr_fg`, bold |
| Alignment | center / center, **wrap_text = True** |
| Border | thin `d0d7de` |
| Row height | 34 pt |

---

### Column A — Variable names (data rows)

| Property | Value |
|---|---|
| Value | `dict_df.index[ri]` |
| Fill | `idx_bg` |
| Font | Calibri 10 pt, colour `idx_fg`, bold |
| Alignment | **left** / center |
| Border | thin `d0d7de` |
| Column width | 22 characters |

---

### Data cells — priority order for fill and font

For each data cell, evaluate these conditions **in order** and apply the first match:

```
1. TYPE BADGE
   Condition : column is "Type" AND cell value is a non-null string
   Fill      : TYPE_BADGE_FILL.get(value, PALETTE["unk_bg"])
                 Numeric     → num_bg  (1d4ed8)
                 Text        → txt_bg  (0f766e)
                 DateTime    → dt_bg   (7c3aed)
                 Boolean     → bool_bg (b45309)
                 Categorical → cat_bg  (15803d)
                 anything else → unk_bg (64748b)
   Font      : Calibri 10 pt, colour ffffff, bold

2. % MISSING HEAT MAP
   Condition : column is "% Missing" AND value is not None
   Fill      : white → red interpolation
                 t   = min(value / 100, 1.0) × 0.85
                 g   = int(255 × (1 − t))
                 b   = int(255 × (1 − t))
                 hex = "FF" + f"{g:02X}" + f"{b:02X}"
   Font      : colour ffffff if value > 55, else 1a1a1a

3. BLUE GRADIENT  (spread / shape / aggregate columns)
   Condition : column is in GRADIENT_COLS AND value is not None
               AND the column has a valid (min, max) range pre-computed
   GRADIENT_COLS = [
       "Mean", "Median (50%)", "Std Dev", "Variance", "MAD",
       "CV", "IQR", "Skewness", "Kurtosis", "Sum", "Range"
   ]
   Per-column range: scan dict_df[col].dropna() for (min, max).
   Skip gradient if min == max (all values identical) or column is all-NaN.
   Fill : lerp between lo=(239,246,255) and hi=(29,78,216):
             t = (value − col_min) / (col_max − col_min)
             r = int(239 + (29  − 239) × t)
             g = int(246 + (78  − 246) × t)
             b = int(255 + (216 − 255) × t)
             hex = f"{r:02X}{g:02X}{b:02X}"
   Font : colour ffffff if t > 0.55, else 1a1a2e

4. DEFAULT (alternating row fill)
   Fill : odd_bg (f8fafc) for ri % 2 == 0, even_bg (ffffff) for ri % 2 == 1
   Font : Calibri 10 pt, colour 1a1a2e
```

All data cells (regardless of priority branch):
- Alignment: center / center
- Border: thin `d0d7de` on all four sides
- Row height: 17 pt

---

### Number formats (Excel `number_format` strings)

Apply to every data cell in the named column where `value is not None`:

```python
NUM_FMT = {
    "% Missing":      '0.00"%"',  "% Distinct":    '0.00"%"',
    "% Unique (exact)":'0.00"%"', "% Zeros":       '0.00"%"',
    "% Negative":     '0.00"%"',
    "Mean":           "#,##0.0000", "Median (50%)": "#,##0.0000",
    "Std Dev":        "#,##0.0000", "Variance":     "#,##0.0000",
    "MAD":            "#,##0.0000", "CV":           "#,##0.0000",
    "IQR":            "#,##0.0000", "Skewness":     "#,##0.0000",
    "Kurtosis":       "#,##0.0000", "5th Pct":      "#,##0.0000",
    "25th Pct":       "#,##0.0000", "75th Pct":     "#,##0.0000",
    "95th Pct":       "#,##0.0000", "Range":        "#,##0.0000",
    "Sum":            "#,##0.00",   "Mean Length":  "0.00",
    "Median Length":  "0.00",       "Total Rows":   "#,##0",
    "Count (non-null)":"#,##0",     "# Missing":    "#,##0",
    "# Distinct":     "#,##0",      "# Unique (exact)":"#,##0",
    "# Zeros":        "#,##0",      "# Negative":   "#,##0",
    "# Infinite":     "#,##0",      "Memory (bytes)":"#,##0",
    "# Characters":   "#,##0",      "# Distinct Chars":"#,##0",
    "Top Freq":       "#,##0",
}
```

---

### Column widths

```python
# Column A (Variable names) — fixed
ws.column_dimensions["A"].width = 22

# All other columns — auto-sized to content, capped
for ci, col_name in enumerate(all_cols):
    letter  = get_column_letter(ci + 2)
    lengths = [len(str(col_name))] + [len(str(v)) for v in dict_df[col_name].dropna()]
    ws.column_dimensions[letter].width = min(max(max(lengths) * 1.05 + 2, 9), 26)
```

---

## Step 4 — Finalise and save

```python
ws.freeze_panes = "B3"           # locks rows 1-2 AND column A
wb.save(out_path)
```

---

## Validation checklist

Before marking the task complete, verify:

- [ ] `dict_df` has exactly **43 columns** (excluding the Variable index).
- [ ] Row count equals the number of variables in `profile_json["variables"]`.
- [ ] Column A is frozen and displays variable names in dark navy with white bold text.
- [ ] Row 1 shows **uppercase group labels** in coloured merged cells.
- [ ] Row 2 shows **column names** in deep-navy cells with slate text.
- [ ] The `Type` column cells display coloured badges (no default fill visible behind them).
- [ ] The `% Missing` column shifts from white (0 %) toward red (100 %).
- [ ] Gradient columns (`Mean`, `Std Dev`, etc.) transition from pale-blue (low) to dark-blue (high).
- [ ] NaN values in the source data produce **blank cells** (`None` written to openpyxl).
- [ ] Sheet tab is blue (`1d4ed8`).
- [ ] Freeze pane is set to `"B3"`.
- [ ] File is saved as `.xlsx`.

---

## Edge cases

| Situation | Correct behaviour |
|---|---|
| A group's columns are all absent from `dict_df` | Skip the group entirely — do not write a group header cell for it |
| A gradient column is entirely NaN or all values are equal | Skip gradient; apply default alternating fill |
| `value_counts_without_nan` is missing or empty | `Top Value` and `Top Freq` → `np.nan` → blank cell |
| `is_unique` is stored as a Python `bool` in JSON | Cast with `bool(v)` — openpyxl writes it as `TRUE`/`FALSE` |
| `Min` / `Max` values are strings (e.g., for date columns stored as text) | Write the raw string value; do not apply a number format |
| A merged-cell range has only one column (`s == e`) | Do not call `merge_cells` (a single cell cannot be merged); just style and write to it normally |
| `memory_size` is an integer but stored as float in JSON | Safe-cast: `int(v)` after confirming `v` is not `float('nan')` |

---

## Reference files

| File | Role |
|---|---|
| `data_dictionary.py` | Canonical implementation — import or run directly |
| `northwind_data_dictionary.ipynb` | Interactive walkthrough with live outputs |
| `README.md` | User-facing quick-start, CLI reference, and customisation guide |
| `northwind_data_dictionary.xlsx` | Reference output — the gold-standard styled workbook |
