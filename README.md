# Footwear Production Defect & Performance Analysis

An end-to-end Power BI project that analyzes one month of production at a footwear factory (Cutting → Gluing → Stitching → Assembly) and turns it into a 5-page dashboard for the **Production General Manager**. It combines quality metrics (defect rate, FPY, RTY, COPQ) with Industrial Engineering (IE) metrics (SAM, efficiency, utilization, downtime, changeover loss).

> **Note:** The dataset is synthetic and was created for this portfolio project. It does not come from a real company.

---

## Business problem

The overall defect rate looks low (2.35%), but is that the full story? The Production GM needs to know:

1. **Where** do defects happen (process, line, style, shift) and which ones exceed their targets?
2. **Why** do they happen (6M root causes, humidity, changeovers, new staff)?
3. How are **efficiency, plan attainment and downtime** performing?
4. How much do defects **cost** (rework + scrap), and where is that cost concentrated?

## Key findings

| # | Finding | Evidence |
|---|---------|----------|
| 1 | **Assembly exceeds its threshold**; Stitching and Gluing are at warning level; only Cutting is on target | Assembly 3.08% (target 2.5%), Stitching 2.89% (2.6%), Gluing 1.95% (1.8%), Cutting 1.50% (1.6%) |
| 2 | **Cutting drives over half of the cost of poor quality**, even though it meets its defect target | ~7.1K of 13.26K USD (53.55%); 1,299 of 1,755 scrapped pairs (74%) |
| 3 | **Method and Man cause 54% of defects**; high humidity makes Assembly spike | Method 29.0%, Man 25.3%; Assembly 3.86% when humidity ≥ 90% vs ~2.9% otherwise |
| 4 | **Efficiency misses its target**; machine breakdown is the largest source of downtime | Efficiency 82% vs 85% target; plan attainment 93%; ~85 downtime hours from breakdowns |
| 5 | **Shift 2 and complex styles have higher defect rates**; Line 06 is the weakest line | Shift 2 2.54% vs Shift 1 2.20% (+15%); Court High-Top 2.87%; Line 06 Stitching 3.96%, Assembly 4.28% |

Headline KPIs: **139,759 pairs** produced vs 150,300 planned · **RTY 90.9%** · **COPQ 13.26K USD** (~0.095 USD per pair) · downtime **367.5 hours**.

## Dashboard

| Page | Purpose |
|------|---------|
| **1. Overview** | Output vs plan, efficiency, defect rate, RTY, COPQ, process status vs target, weekly trend, defect rate by line, key findings |
| **2. Quality** | Defect Pareto, line × process heatmap, defect rate by shift and style |
| **3. Root Cause** | 6M analysis, severity, defect rate by humidity band, method-related defects by style |
| **4. IE & Downtime** | Efficiency, plan attainment, utilization, downtime by reason, changeover loss, performance by line |
| **5. COPQ** | Cost of poor quality by process, defect code, line and shift |

![Overview](images/01_dashboard_overview.png)
![Quality](images/02_dashboard_quality.png)
![Root Cause](images/03_dashboard_root_cause.png)
![IE & Downtime](images/04_dashboard_ie_downtime.png)
![COPQ](images/05_dashboard_copq.png)

## Dataset

One month of simulated production (October 2026): 27 working days (Mon–Sat), 6 lines, 6 styles, 2 shifts, 4 processes and 25 defect codes.

| Table | Type | Grain | Rows |
|-------|------|-------|------|
| `Fact_Production` | Fact | Day × shift × line × process | 1,176 |
| `Fact_Defects` | Fact | Production row × defect code | 5,530 |
| `Fact_Downtime` | Fact | One stoppage event | 1,189 |
| `Dim_Calendar` | Dimension | Date (with humidity and weather) | 31 |
| `Dim_Line`, `Dim_Style`, `Dim_WorkOrder`, `Dim_SAM`, `Dim_Process`, `Dim_Defect`, `Dim_DowntimeReason` | Dimension | — | — |

The workbook also contains `Summary_Check` (KPIs computed with Excel formulas, used to validate the Power BI numbers), `Data_Dictionary`, `Assumptions` and `Scenario_Key` (the anomalies planted in the data).

### Data model (star schema)

All relationships are one-to-many, single direction, from dimension (or parent fact) to fact:

```
Dim_Calendar[Date]            → Fact_Production[Date]
Dim_Line[Line_ID]             → Fact_Production[Line_ID]
Dim_Style[Style_ID]           → Fact_Production[Style_ID]
Dim_Process[Process_ID]       → Fact_Production[Process_ID]
Dim_WorkOrder[WO_ID]          → Fact_Production[WO_ID]
Fact_Production[Prod_ID]      → Fact_Defects[Prod_ID]
Fact_Production[Prod_ID]      → Fact_Downtime[Prod_ID]
Dim_Defect[Defect_Code]       → Fact_Defects[Defect_Code]
Dim_DowntimeReason[Reason_Code] → Fact_Downtime[Reason_Code]
```

> **Lesson learned:** the duplicate `Date`, `Line_ID`, `Style_ID` and `Process_ID` columns in `Fact_Defects` and `Fact_Downtime` must not be used as chart axes or joined to dimensions. Doing so filters only the numerator of a rate and produces wrong results (for example an RTY of 97.7% instead of 90.9%). Always filter through the dimensions and `Fact_Production`.

## KPIs and IE concepts

| KPI | Definition |
|-----|------------|
| Defect rate | Defective pairs ÷ inspected pairs (DHU = rate × 100) |
| FPY | First pass yield of a process = 1 − defect rate |
| RTY | Rolled throughput yield = product of the four process FPYs |
| SAM | Standard allowed minutes per pair, by style and process |
| Efficiency | Earned minutes (output × SAM) ÷ available minutes (operators × shift minutes) |
| Utilization | (Shift minutes − downtime) ÷ shift minutes |
| Plan attainment | Assembly output ÷ planned pairs |
| COPQ | Rework cost (rework minutes × labor rate) + scrap cost (scrapped pairs × cumulative material value) |
| 6M | Man, Machine, Material, Method, Measurement, Environment |

## Selected DAX

```dax
Defect Rate =
DIVIDE ( SUM ( Fact_Defects[Defect_Qty] ), SUM ( Fact_Production[Output_Pairs] ) )

FPY = 1 - [Defect Rate]

FPY CUT = CALCULATE ( [FPY], Dim_Process[Process_ID] = "CUT" )
FPY GLU = CALCULATE ( [FPY], Dim_Process[Process_ID] = "GLU" )
FPY STI = CALCULATE ( [FPY], Dim_Process[Process_ID] = "STI" )
FPY ASM = CALCULATE ( [FPY], Dim_Process[Process_ID] = "ASM" )

RTY = [FPY CUT] * [FPY GLU] * [FPY STI] * [FPY ASM]

Efficiency =
DIVIDE ( SUM ( Fact_Production[Earned_Min] ), SUM ( Fact_Production[Available_Min] ) )

COPQ = SUM ( Fact_Defects[Total_Quality_Cost_USD] )

Target Defect Rate (Process) =
IF (
    HASONEVALUE ( Dim_Process[Process_ID] ),
    MAX ( Dim_Process[Target_Defect_Rate] ),
    [Target Defect Rate]
)

Defect Status (Process) =
VAR _r = [Defect Rate]
VAR _t = [Target Defect Rate (Process)]
RETURN
    SWITCH (
        TRUE (),
        _r > _t * 1.2, "Over threshold",
        _r > _t, "Warning",
        "On target"
    )
```

The full set of 50+ measures (grouped into 7 KPI folders), the date table and the process-alert measures are in the [`dax/`](dax/) folder.

## Repository structure

```
.
├── README.md
├── data/
│   └── Shoe_Factory_Defect_Simulation_Oct2026.xlsx
├── dax/
│   ├── PowerBI_DAX_Measures.txt
│   ├── PowerBI_Dim_Date.txt
│   └── PowerBI_DAX_Process_Alert.txt
├── dashboard/
│   └── Production_Dashboard.pbix          # add your Power BI file here
├── docs/
│   └── Portfolio_Production_Performance_Analysis_EN.pptx
└── images/
    ├── 01_dashboard_overview.png
    ├── 02_dashboard_quality.png
    ├── 03_dashboard_root_cause.png
    ├── 04_dashboard_ie_downtime.png
    └── 05_dashboard_copq.png
```

## How to use

1. Download the repository and open the `.pbix` file in Power BI Desktop.
2. If the data source path is broken, go to **Transform data → Data source settings** and point it to `data/Shoe_Factory_Defect_Simulation_Oct2026.xlsx`.
3. Check the model relationships against the list above before building new visuals.
4. Validate your numbers against the `Summary_Check` sheet: 139,759 final pairs, defect rate 2.35%, RTY 90.9%, efficiency 82.0%.

## Limitations

- **Synthetic data:** generated for this project, with planted anomalies; it does not reflect a real company.
- **One main defect per pair:** pairs with several simultaneous defects cannot be analyzed.
- **Cost assumptions:** the labor rate (0.055 USD/minute) and cumulative material value are assumptions, so COPQ is indicative.
- **One month only:** not enough to conclude on trends or seasonality; weeks W40 and W44 are only partly covered.

## Tools

Microsoft Excel · Power BI Desktop · DAX · Industrial Engineering concepts (SAM, efficiency, line balancing, learning curve, 6M)

## Author

**[Your Name]** · Data Analyst
[LinkedIn](https://www.linkedin.com/in/your-profile) · [Email](mailto:your.email@example.com)
