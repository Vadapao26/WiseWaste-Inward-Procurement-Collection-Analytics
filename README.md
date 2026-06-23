# 📥 Inward Analytics

Analyzes material procurement data across vendors, sources, locations, drivers, and vehicles — producing a structured multi-sheet Excel workbook per facility with 18 KPIs, rejection flags, and full SVL (Source → Vendor → Location) drilldowns.

---

## The Problem It Solves

Understanding procurement performance across DWCCs, ULBs, and aggregators used to require manually building pivot tables in Excel for every facility every month — grouping by vendor, filtering by material, computing rejection rates, and cross-checking driver trip counts. With multiple facilities and dozens of vendors, this took **3–5 hours per reporting cycle**.

This notebook processes all of that in **under 2 minutes**, producing a clean, formatted Excel workbook that's ready to share with field teams, program managers, and ULB stakeholders.

---

## Architecture

```
Cleaned Inward CSV(s) Upload
           │
           ▼
┌─────────────────────────────────┐
│  1. Column Standardization      │  ← rename map + duplicate merge
└──────────────┬──────────────────┘
               │
               ▼
┌─────────────────────────────────┐
│  2. Date Parsing (6 formats)    │  ← dayfirst + known format fallbacks
└──────────────┬──────────────────┘
               │
               ▼
┌─────────────────────────────────┐
│  3. Derived Column Computation  │  ← Accepted Qty, Rejection %, Value cols
└──────────────┬──────────────────┘
               │
               ▼
┌─────────────────────────────────┐
│  4. Driver Fuzzy Name Grouping  │  ← SequenceMatcher at 0.82 threshold
└──────────────┬──────────────────┘
               │
               ▼
┌─────────────────────────────────┐
│  5. Material Category Mapping   │  ← keyword-based → Paper/Plastics/etc.
└──────────────┬──────────────────┘
               │
               ▼
┌─────────────────────────────────┐
│  6. Multi-Dimensional Analysis  │
│   ├── KPI Summary (18 metrics)  │
│   ├── Material Analysis         │
│   ├── SVL Analysis              │
│   ├── Driver Analysis           │
│   └── Vehicle Analysis          │
└──────────────┬──────────────────┘
               │
               ▼
┌─────────────────────────────────┐
│  7. Excel Export + Formatting   │  ← freeze panes, autofit, header style
└──────────────┬──────────────────┘
               │
               ▼
  Inward_Analytics_Output.xlsx
```

---

## Languages & Libraries

| Layer | Tool | Purpose |
|---|---|---|
| Language | Python 3 | Core logic |
| Data processing | Pandas | Groupby, aggregation, pivots |
| Numeric ops | NumPy | Conditional calculations |
| Fuzzy matching | `difflib.SequenceMatcher` | Driver name deduplication |
| Text processing | `re` (regex) | Name normalization, repeated char removal |
| Excel output | openpyxl | Sheet creation, formatting, autofit |
| Date parsing | `datetime` | Multi-format date handling |
| Environment | Google Colab | Upload / download |

---

## Output Sheets (per uploaded facility file)

Each file gets its own set of sheets, prefixed by filename:

| Sheet | Contents |
|---|---|
| `{file}_KPI` | 18 headline metrics — total qty, rejection %, avg rate, unique vendors/drivers/vehicles, data quality flags |
| `{file}_Raw` | Full enriched dataset — 37 columns including parsed dates, derived values, exception flags, doc availability |
| `{file}_Material` | 3 blocks: Major Category summary, Material-level summary, Valuable vs Non-Valuable split |
| `{file}_SVL` | 3 blocks: Source summary, Source→Vendor→Location drilldown, Hierarchy material/category mix |
| `{file}_Driver` | Driver-level totals, rejection, valuable/non-valuable split, vehicles used, original name variants |
| `{file}_Vehicle` | Vehicle-level totals, rejection, drivers used, valuable/non-valuable split |

---

## KPI Summary — 18 Metrics Computed

1. Total Quantity Received
2. Total Rejection
3. Average Rejection %
4. Average Rate (₹/kg)
5. Average Quantity per Day
6. Average Quantity per Inward
7. Total Value of Received Material
8. Total Value of Accepted Material
9. Total Net Procurement Cost
10. Total Valuable Quantity
11. Total Non-Valuable Quantity
12. Total Distinct Sources
13. Total Distinct Vendors
14. Total Distinct Locations
15. Total Distinct Standardised Drivers
16. Total Distinct Vehicles
17. Rows with Negative Net Procurement Cost
18. Rows with Unparsed Date

---

## Features in Detail

### Driver Fuzzy Name Grouping
Driver names in the MIS are entered manually and contain typos, repeated characters, and spelling variations (`"Ravi Kumar"`, `"ravi kumarrr"`, `"Ravi Kmr"`). The notebook uses `SequenceMatcher` at a 0.82 similarity threshold to group these variants under one standardized name — preventing the same driver from appearing as 3–4 separate entries in the analysis.

### SVL (Source → Vendor → Location) Drilldown
Three-level procurement hierarchy analysis: which source type (DWCC, ULB, Aggregator) is contributing the most, which specific vendors within each source, and which locations. Each level shows quantity, rejection %, valuable share %, net procurement cost, and material category mix.

### Exception Flags
Every row is automatically flagged for:
- Unparsed dates
- Negative Net Procurement Cost
- Rejection % above 20%

These surface in the Raw sheet and can be filtered to find data quality issues or anomalous records immediately.

### Document Availability Tracking
Tracks whether key compliance documents (Weighment Slip, Vendor Invoice, GPS Photos, Inward Record) are present per trip — outputs a Yes/No column for each, making document audit-readiness checks instant.

---

## Use Cases

- **Monthly procurement reporting** — share KPI sheet directly with program managers and ULB stakeholders
- **Underperforming source identification** — sort SVL sheet by Rejection % to find which wards or buildings need field intervention
- **Driver performance review** — Driver sheet shows which drivers are bringing the most material and at what rejection rate
- **Data quality audit** — Exception Flags column immediately surfaces negative costs and unparsed dates
- **Multi-facility comparison** — upload 3–4 facility files at once; one workbook with all facilities' sheets for side-by-side review

---

## Time Saved

| Task | Manual (Excel) | This Notebook |
|---|---|---|
| Build vendor-level pivot | 20–30 min | < 10 sec |
| Compute rejection % per material per vendor | 15–20 min | < 10 sec |
| Driver deduplication + grouping | 30–45 min | < 10 sec |
| SVL three-level drilldown | 30–45 min | < 10 sec |
| Format + share-ready Excel output | 20–30 min | < 10 sec |
| Process 3 facility files | 3–5 hours | ~90 sec |
| **Total per reporting cycle** | **~4 hours** | **~2 min** |

**Estimated time saving: ~97% reduction per reporting cycle.**

---

## Input / Output

| | Detail |
|---|---|
| **Input** | One or more `*_cleaned.csv` files from the [Data Cleaning Pipeline](../data-cleaning-pipeline/) |
| **Output** | `Inward_Analytics_Output.xlsx` — multi-sheet workbook, auto-downloads from Colab |

---

## How to Use

### Google Colab
1. Click **Open in Colab** above
2. Run all cells — you'll be prompted to upload files
3. Upload one or more cleaned inward CSVs (select multiple with Ctrl/Cmd+click)
4. Each file is analyzed separately; one Excel workbook with all sheets downloads automatically


---
