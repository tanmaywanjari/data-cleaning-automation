# Data Cleaning & Reporting Automation

A one-command Python pipeline that takes a messy sales file and automatically **cleans it, logs every fix, and produces an Excel report, an HTML report and charts**. Re-run it on any new file and the whole report refreshes.

> **Data note:** no dataset was provided with the task, so `generate_messy_data.py` creates a deliberately messy file (`data/raw_sales.csv`, 1,555 rows). To use real data, keep the same column names and run the pipeline on your file.

## Project structure
| File | Purpose |
|------|---------|
| `generate_messy_data.py` | Creates the messy sample data |
| `clean_and_report.py` | The automated pipeline: clean -> analyse -> report |
| `data/raw_sales.csv` | Raw input |
| `outputs/` | `clean_sales.csv`, `sales_report.xlsx`, `sales_report.html`, 5 charts, `run_log.txt` |

## How to run
```bash
pip install -r requirements.txt
python generate_messy_data.py                       # optional, sample data is included
python clean_and_report.py                          # default input
python clean_and_report.py --input data/my_file.csv --out outputs
```

## Problems found and fixed (see the Cleaning Log sheet in the Excel report)
| Problem | How it was handled |
|---------|--------------------|
| 289 values with extra spaces | Trimmed |
| Region, Payment Method and Status written many ways (`north`, `N.`, `CC`, `Done`...) | Mapped to one standard label (761 fixes) |
| Dates in 3 different formats | Parsed into one date type |
| 601 prices stored as text (`$1,200.50`) | Converted to numbers |
| 40 exact duplicate rows | Removed |
| 15 repeated Order IDs (updated records) | Kept the latest version |
| 12 negative quantities, 6 invalid emails | Set to missing, then imputed or blanked |
| Missing quantity (71) and price (60) | Filled with the median for the same product |
| Missing customer name and region (45 each) | Filled with `Unknown` |
| 3 rows with no order date | Dropped (cannot be placed in time) |

Result: 1,555 raw rows became **1,497 clean rows**, with no duplicates, one spelling for every category, and no missing values except Email, which is left blank on purpose because it isn't needed for revenue.

## Reports generated automatically
- **Excel report:** Summary KPIs, Charts, Cleaning Log, Monthly Revenue, By Region, By Category, By Product, By Payment and the full Clean Data, with formatted headers.
- **HTML report:** a single shareable page with KPIs, the cleaning log and all charts.
- **Charts:** data quality before vs after, monthly revenue, revenue by region, revenue by product, category share.

## Headline numbers (completed orders)
Total revenue about **924.7K** from 1,221 orders, average order value about **757**, return rate 9.2%, cancel rate 9.3%.

## Scheduling it (true automation)
- **Windows:** Task Scheduler -> run `python clean_and_report.py` daily or weekly.
- **Mac / Linux:** add a cron line, for example `0 8 * * 1 cd /path/to/project && python clean_and_report.py` (every Monday at 8am).
Each run appends a line to `outputs/run_log.txt`.

## Tech stack
Python, pandas, NumPy, matplotlib, seaborn, openpyxl.

## Learning outcomes
Data preprocessing, handling missing values, duplicates and inconsistent data, logging and validation, and automated Excel/HTML reporting.
