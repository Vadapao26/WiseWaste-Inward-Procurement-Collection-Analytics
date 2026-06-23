2. 📥 Inward Analytics

inward-analytics/

Analyzes material procurement data across vendors, ULBs, DWCCs, drivers, and sources. Upload one or more cleaned inward CSVs (one per facility) and the notebook analyzes each separately, outputting a single Excel workbook with dedicated sheets per file.

What it does:


Multi-file support — upload multiple facility CSVs at once; each analyzed separately with its own Excel sheets
Vendor-level analysis — total quantity received, rejection rate, net procurement cost, and inward count per vendor
ULB / DWCC / Source breakdowns — quantity and rejection trends by collection source type and location
Driver-level analysis — trips per driver, material type split, total quantity delivered
Material category mapping — materials auto-mapped to Paper, Glass, Plastics, E-Waste, C&D Waste, Wet Waste, Others
Fuzzy name matching — groups similar vendor and driver name variants using sequence similarity scoring
Rejection flagging — rejection % per vendor, material, and source with threshold-based highlighting
Flexible date parsing — handles DD/MM/YYYY, DD-MM-YYYY, DD Mon YYYY, and other formats


Input: One or more *_cleaned.csv files from the Data Cleaning Pipeline
Output: Excel workbook with per-facility analysis sheets — auto-downloads from Colab
