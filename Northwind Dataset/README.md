# Data Dictionary — Style Guide & Usage

Generate a comprehensive, styled Excel data dictionary from any pandas DataFrame.
Every variable becomes a row; every statistical attribute becomes a column.

---

## What the output looks like

| Feature | Detail |
|---|---|
| **Two-level header** | Row 1 = colour-coded group labels (merged); Row 2 = column names |
| **Type badge** | Solid fill per profiling type — blue (Numeric), teal (Text), violet (DateTime) |
| **% Missing heat map** | White → red cell background |
| **Spread gradient** | Light-blue → dark-blue intensity per column (Mean, Std Dev, IQR, …) |
| **Frozen panes** | Column A + rows 1–2 always visible while scrolling |
| **Variable column** | Dark navy background, white bold text |
| **Alternating rows** | Pale-blue / white |
| **Number formats** | `#,##0.0000` for stats · `0.00%` for percentages · `#,##0` for counts |
| **Auto column widths** | Sized to content, capped at 26 characters |
| **Sheet tab** | Blue (`#1d4ed8`) |

---

## Column groups

| Group | Attributes |
|---|---|
| **Identity** | Type, Pandas dtype, Memory (bytes) |
| **Completeness** | Total Rows, Count (non-null), # Missing, % Missing |
| **Uniqueness** | # Distinct, % Distinct, # Unique (exact), % Unique (exact), Is Unique |
| **Central Tendency** | Mean, Median (50%) |
| **Spread** | Std Dev, Variance, MAD, CV, IQR |
| **Shape** | Skewness, Kurtosis |
| **Range / Quantiles** | Min, 5th Pct, 25th Pct, 75th Pct, 95th Pct, Max, Range |
| **Aggregate** | Sum |
| **Special Counts** | # Zeros, % Zeros, # Negative, % Negative, # Infinite, Monotonic |
| **Text Metrics** | Min/Max/Mean/Median Length, # Characters, # Distinct Chars |
| **Most Frequent** | Top Value, Top Freq |

Cells that are N/A for a variable's type (e.g. *Mean* on a Text column) display `—`.

---

## Requirements

```
pandas
numpy
ydata-profiling
openpyxl
```

Install everything at once:

```bash
pip install pandas numpy ydata-profiling openpyxl
```

---

## Quick start

### Option A — From a CSV file

```bash
python data_dictionary.py --csv mydata.csv --output mydata_dict.xlsx
```

### Option B — From a SQLite database

```bash
python data_dictionary.py \
  --sqlite mydb.db \
  --query "SELECT * FROM sales" \
  --output sales_dict.xlsx
```

### Option C — Reproduce the Northwind enriched dataset exactly

```bash
python data_dictionary.py --northwind northwind.db --output northwind_data_dictionary.xlsx
```

Add `--minimal` to any command to skip expensive correlation steps (much faster, same per-variable stats):

```bash
python data_dictionary.py --csv mydata.csv --minimal
```

---

## Use as a Python library

```python
import pandas as pd
from data_dictionary import build_data_dictionary

df = pd.read_csv("sales.csv")

# One call — profiles, builds, and saves the styled Excel file
dict_df = build_data_dictionary(df, "sales_dict.xlsx")

# dict_df is a plain DataFrame — inspect or do further work on it
print(dict_df[["Type", "% Missing", "Mean", "Std Dev", "Top Value"]])
```

### Step-by-step (lower-level)

```python
from data_dictionary import profile_to_dict_df, build_excel

# 1. Profile — returns a tidy DataFrame of stats
dict_df = profile_to_dict_df(df, minimal=True)

# 2. Inspect / filter before writing
print(dict_df.shape)              # (n_variables, 43)
high_missing = dict_df[dict_df["% Missing"] > 5]
print(high_missing[["Type", "% Missing", "# Distinct"]])

# 3. Write styled Excel
build_excel(dict_df, "my_dict.xlsx")
```

---

## CLI reference

```
usage: data_dictionary [-h]
                       (--csv FILE | --sqlite FILE | --northwind FILE)
                       [--query SQL]
                       [--output FILE]
                       [--minimal]
                       [--title TEXT]

options:
  --csv FILE          Path to a CSV file
  --sqlite FILE       Path to a SQLite database file
  --northwind FILE    Reproduce the Northwind enriched dataset
  --query SQL         SQL query (required with --sqlite)
  --output FILE       Output Excel path  (default: data_dictionary.xlsx)
  --minimal           Skip expensive correlations for faster profiling
  --title TEXT        Report title (cosmetic)
```

---

## Notebook version

Open `northwind_data_dictionary.ipynb` for an interactive walkthrough:

| Cell | What it does |
|---|---|
| **1 · Data Preparation** | Builds the enriched Northwind line-item DataFrame |
| **2 · Generate Profile → JSON** | Runs `ProfileReport` and exports stats as JSON |
| **3 · Parse JSON** | Extracts all per-variable attributes into `dict_df` |
| **4 · Styled Table** | Renders the data dictionary inline (scrollable HTML) |
| **5 · Export to Excel** | Writes `northwind_data_dictionary.xlsx` with full styling |

---

## Customising the style

All styling constants live at the top of `data_dictionary.py`.

| Constant | What it controls |
|---|---|
| `PALETTE` | All hex colours (header, row index, type badges, …) |
| `GROUP_BG` | Rotating background colours for the group-label row |
| `TYPE_BADGE_FILL` | Per-type badge fill colour (`Numeric`, `Text`, `DateTime`, …) |
| `COLUMN_GROUPS` | Group names and which columns belong to each group |
| `GRADIENT_COLS` | Which attribute columns get the blue intensity gradient |
| `NUM_FMT` | Excel number format strings per column |

**Example — change the Numeric badge to green:**

```python
from data_dictionary import TYPE_BADGE_FILL, build_data_dictionary

TYPE_BADGE_FILL["Numeric"] = "15803d"   # green hex, no leading #
build_data_dictionary(df, "out.xlsx")
```

**Example — add a new column group:**

```python
from data_dictionary import COLUMN_GROUPS

COLUMN_GROUPS.append(("My Group", ["My Col A", "My Col B"]))
```
