# Motor Insurance Portfolio Analytics

## Excel Data Cleaning, Insurance KPI Analysis, Risk Segmentation and Management Dashboard

> **Project type:** Practical Excel Analytics Project  
> **Domain:** Motor Insurance Portfolio Analysis  
> **Observation period:** 2022–2024  
> **Primary tool:** Microsoft Excel  

---

## Contents

1. [Executive Summary](#1-executive-summary)
2. [Project Objectives](#2-project-objectives)
3. [Dataset and Source](#3-dataset-and-source)
4. [Workbook Structure](#4-workbook-structure)
5. [Project Workflow: Seven Modules](#5-project-workflow-seven-modules)
6. [Data Audit and Quality Controls](#6-data-audit-and-quality-controls)
7. [Core Insurance KPIs and Definitions](#7-core-insurance-kpis-and-definitions)
8. [Portfolio-Level Results](#8-portfolio-level-results)
9. [Annual Portfolio Analysis](#9-annual-portfolio-analysis)
10. [Policy-Type Analysis](#10-policy-type-analysis)
11. [Driver-Age Risk Segmentation](#11-driver-age-risk-segmentation)
12. [Vehicle-Age Risk Segmentation](#12-vehicle-age-risk-segmentation)
13. [Management Dashboard](#13-management-dashboard)
14. [Key Findings and Recommended Follow-Up](#14-key-findings-and-recommended-follow-up)
15. [Excel Skills Demonstrated](#15-excel-skills-demonstrated)
16. [Scope Decisions](#16-scope-decisions)
17. [Limitations and Interpretation Notes](#17-limitations-and-interpretation-notes)
18. [Portfolio and Summary](#18-portfolio-and-summary)
19. [Final Project Status](#19-final-project-status)
20. [Author](#20-author)

---

## 1. Executive Summary

This project analyses an anonymised motor-insurance portfolio in Microsoft Excel. It was designed as a practical, job-focused project for developing skills relevant to an insurance analyst role. The work covers data auditing and cleaning, insurance KPI calculations, policy and insured lookups, annual portfolio trends, driver- and vehicle-age segmentation, interactive reporting, reconciliation and documentation.

The working dataset contains **354,140 policy-year observations** across **2022, 2023 and 2024**, with **185,678 unique insureds**. The portfolio has approximately **€105.08 million in total premium**, **€73.86 million in total incurred**, **77,099 claims** and **250,325.90 exposure-years**. The overall loss ratio is **70.29%**.

Annual analysis shows that the loss ratio rose from **65.62% in 2022** to **74.66% in 2024**. Claim frequency rose modestly over the same period, while claim severity decreased in 2023 before returning close to its 2022 level in 2024. These findings indicate areas for further investigation, particularly the 2024 loss-ratio increase, but descriptive results alone do not establish why the changes occurred.

The completed workbook includes separate raw and cleaned data worksheets, analysis and validation sheets, lookup examples, risk-segmentation summaries and an interactive management dashboard. A policy-type loss-ratio comparison is documented in this README because the dashboard did not have sufficient space for another chart.

## 2. Project Objectives

The project was deliberately kept focused on useful, transferable Excel skills rather than becoming an exhaustive catalogue of disconnected exercises.

The main objectives were to:

- Inspect and prepare a large insurance dataset while retaining the raw source data.
- Check data types, missing values, duplicates, categorical consistency and unusual values.
- Use Excel Tables and structured references to make formulas easier to read and maintain.
- Aggregate premium, exposure, claims and incurred amounts by relevant dimensions.
- Calculate claim frequency, claim severity, loss ratio, premium per exposure and year-over-year change.
- Use **XLOOKUP as the primary lookup function** for practical insured and policy-year questions.
- Analyse trends by year and compare outcomes across policy types, driver-age bands and vehicle-age bands.
- Build charts, PivotCharts, slicers and a management dashboard.
- Reconcile analysis totals to the overall portfolio totals.
- Translate quantitative findings into cautious, evidence-based investigation priorities.
- Prepare professional documentation that explains the workbook and its principal findings.

## 3. Dataset and Source

The project uses a published, anonymised motor-insurance dataset containing coverage-specific financial information from a Spanish insurance company over January 2022 to December 2024.

- **Dataset landing page:** <https://data.mendeley.com/datasets/sw4jmdb2sm/1>
- **Related research article:** <https://www.sciencedirect.com/science/article/pii/S235234092600435X>
- **Research/institutional record:** <https://portalcientifico.uah.es/documentos/6993491d7a4cc00ebaa84b87?lang=es>
- **Licence reported for the dataset:** Creative Commons Attribution 4.0 International (CC BY 4.0). Review the dataset page and licence terms when redistributing the data.

### Dataset profile

| Attribute | Description |
|---|---|
| Observation period | 2022–2024 |
| Records | 354,140 policy-year observations |
| Variables | 47 |
| Unique insureds | 185,678 |
| Geography of source insurer | Spain |
| Data type | Anonymised historical insurance portfolio data |
| Currency display in workbook | Euro (€), as configured in the finished workbook |

Each row represents a policy-year observation. An insured can therefore appear in more than one year. Repeated `insured_id` values are not automatically duplicate records and should not be removed simply because the insured ID repeats.

The financial cells in the final workbook are formatted with the **euro (€) symbol**, as confirmed during the final dashboard review. Display formatting does not change the underlying numeric values.

## 4. Workbook Structure

The final workbook separates source data, analysis, validation and management reporting into distinct worksheets.

| Worksheet | Purpose |
|---|---|
| `Dashboard` | Portfolio KPI cards, charts, slicers and annual frequency/severity/loss-ratio table. |
| `Raw_Data` | Preserves the original source data for traceability. |
| `Clean_Data` | Holds the cleaned working dataset used for analysis. |
| `Data_Analysis` | Contains portfolio summaries, KPI calculations, annual analysis and related PivotTables/calculations. |
| `Risk_Segment_Analysis` | Contains driver-age and vehicle-age segment summaries. |
| `Lookups` | Contains practical lookup examples using insured and policy-year criteria. |
| `Checks` | Holds data-quality checks, reconciliations and validation work. |

The names above reflect the worksheet names used in the finished workbook. Keeping these functions separate makes it easier to trace results from the source data through to final reporting.

## 5. Project Workflow: Seven Modules

### Module 1 — Data Audit and Cleaning

Activities included reviewing the table structure and field types; checking missing values and duplicates; inspecting numeric ranges and categorical values; standardising the treatment of missing markers; retaining raw data separately; and creating data-quality checks. Excel Tables and structured references were used to support repeatable formulas.

### Module 2 — Insurance Calculations and KPIs

Built portfolio and segment summaries for premium, exposure, claims and incurred. Calculated claim frequency, claim severity, loss ratio, premium per exposure and premium share. PivotTables were used for aggregation, and ratio measures were calculated from aggregated totals rather than by averaging row-level ratios.

### Module 3 — Advanced Lookups and Business Questions

Practised XLOOKUP with a helper key and multiple criteria, not-found handling, and dynamic-array functions such as FILTER, UNIQUE and SORT. INDEX/MATCH and XMATCH were reviewed to understand alternative approaches, but **XLOOKUP remained the main lookup method used in the project**.

### Module 4 — Date and Portfolio Analysis

The available time dimension is `year`; the dataset does not provide the detailed policy start/end, claim or cancellation dates needed for more granular date analysis. The module therefore used the available annual information to audit record distribution, analyse premium/exposure/claims/incurred by year, calculate year-over-year changes and compare annual insurance ratios. No artificial dates were created.

### Module 5 — Risk Segmentation and Statistical Analysis

Created driver-age and vehicle-age bands, summarised premium, exposure, claims and incurred by segment, and calculated frequency, severity and loss ratio. The segment results were interpreted as descriptive evidence for further investigation, not as proof of causation. Vehicle value banding was considered but excluded from the core project to keep the scope manageable.

### Module 6 — Data Visualization and Management Dashboard

Built KPI cards, charts, PivotCharts and interactive slicers. The dashboard shows portfolio totals and comparisons across years, policy types and age segments. Slicer connections and chart behaviour were checked, and the currency display was changed to euros. A policy-type loss-ratio chart was not added because dashboard space was limited; its results are included in this README.

### Module 7 — Final Workbook Review and Documentation

The final stage covered the workbook audit, reconciliation of key outputs, final dashboard review and documentation. Power Query was deliberately excluded because the necessary data-cleaning and analysis workflow had already been completed and adding another preparation workflow would have duplicated work rather than improved this project.

## 6. Data Audit and Quality Controls

### Missing values identified in the initial audit

The dataset initially used the literal text marker `NA` in some numeric or categorical fields. These entries were identified and converted to blanks in the working data so Excel could treat them consistently as missing values.

| Field | Missing entries initially identified |
|---|---:|
| `vehicle_age` | 2 |
| `age_driving_licence` | 2 |
| `fuel_type` | 1,287 |
| `vehicle_value` | 513 |

The missing-value review did not reveal an obvious concentration pattern in the fields examined. Missing values were not automatically replaced with zero, because a missing value and a real zero can have different meanings. The vehicle-age analysis retains a `Missing` band for the two records without a vehicle age.

### Duplicate and identifier checks

- The duplicate review found no duplicate records under the checks performed.
- Repeated insured IDs were treated as expected policy-year behaviour rather than as proof of duplication.
- `policy_status` uses `A` for active and `C` for cancelled, but status is not a unique record identifier.
- The record-level identifier and relevant year/insured fields must be considered together when examining repeated insureds.

### Selected categorical counts

| Policy type | Records |
|---|---:|
| CC | 204,636 |
| COMP_E | 91,199 |
| COMP_N | 16,181 |
| TP | 5,757 |
| TPG | 36,367 |
| **Total** | **354,140** |

| Policy status | Records |
|---|---:|
| A | 315,812 |
| C | 38,328 |
| **Total** | **354,140** |

### Additional quality observations

- The reviewed numeric fields had these observed ranges: `driver_age` 18–90; `vehicle_age` 0–68; `age_driving_licence` 0–80; `seats` 2–9; `vehicle_value` 1,630.125–374,034.2; `total_premium` 0–5,107.64; `total_claims` 0–18; `total_incurred` 0–571,128.3; and `total_exposure` 0–1.
- Unusual values were reviewed instead of automatically deleted. In particular, exposure can be a fractional policy-year measure and should retain its precision in calculations; round it for display only.
- A data-quality flag/check count of 8,543 was recorded in the project checks. The exposure-premium check returned zero exceptions under the implemented check.
- The final workbook audit confirmed that the missing-value and duplicate checks were documented and that KPI totals reconciled to portfolio totals.

These are findings from the project audit and should not be read as proof that every possible data-quality issue has been eliminated.

## 7. Core Insurance KPIs and Definitions

### Portfolio KPI summary

| KPI | Result |
|---|---:|
| Total records / policy-year observations | 354,140 |
| Unique insureds | 185,678 |
| Total premium | €105,079,669.43 |
| Average premium per record | €296.72 |
| Total claims | 77,099 |
| Total incurred | €73,864,810.69 |
| Total exposure | 250,325.8986 |
| Claim frequency | 0.3080 |
| Claim severity | €958.05 |
| Loss ratio | 70.29% |
| Premium per exposure | €419.77 |

Figures are rounded for presentation. Underlying calculations should retain full precision where appropriate.

### Claim frequency

**Formula:**

```excel
=SUM(tblInsurance[total_claims])/SUM(tblInsurance[total_exposure])
```

**Definition:** Total claims divided by total exposure.

Overall: 77,099 ÷ 250,325.8986 ≈ **0.3080 claims per exposure-year**, equivalent to about **30.8 claims per 100 exposure-years**. It is not the percentage of customers who had at least one claim.

### Claim severity

**Formula:**

```excel
=SUM(tblInsurance[total_incurred])/SUM(tblInsurance[total_claims])
```

**Definition:** Total incurred divided by total claims.

Overall: €73,864,810.69 ÷ 77,099 ≈ **€958.05 per claim**. This is not the average incurred amount per policy-year record.

### Loss ratio

**Formula:**

```excel
=SUM(tblInsurance[total_incurred])/SUM(tblInsurance[total_premium])
```

**Definition:** Total incurred divided by total premium.

Overall: €73,864,810.69 ÷ €105,079,669.43 ≈ **70.29%**. Use incurred divided by premium, not the inverse.

### Premium per exposure

**Formula:**

```excel
=SUM(tblInsurance[total_premium])/SUM(tblInsurance[total_exposure])
```

Overall: approximately **€419.77 per exposure-year**.

These are descriptive portfolio metrics. Loss ratio is not a complete measure of underwriting profit because it may not include expenses, commissions, capital costs or every other financial consideration.

## 8. Portfolio-Level Results

The main totals reconcile across the portfolio and the segment summary tables.

| Measure | Portfolio total |
|---|---:|
| Records | 354,140 |
| Unique insureds | 185,678 |
| Premium | €105,079,669.43 |
| Exposure | 250,325.8986 |
| Claims | 77,099 |
| Incurred | €73,864,810.69 |

The sum of premium, claims, incurred and exposure across the policy-type and age-band tables matches the portfolio-level totals at the displayed precision.

## 9. Annual Portfolio Analysis

### Annual totals

| Year | Records | Premium | Exposure | Claims | Incurred |
|---|---:|---:|---:|---:|---:|
| 2022 | 67,172 | €18,697,928.34 | 41,912.50 | 12,665 | €12,270,105.06 |
| 2023 | 118,835 | €35,344,308.08 | 82,399.04 | 25,158 | €23,488,354.35 |
| 2024 | 168,133 | €51,037,433.00 | 126,014.36 | 39,276 | €38,106,351.28 |
| **Total** | **354,140** | **€105,079,669.43** | **250,325.90** | **77,099** | **€73,864,810.69** |

### Portfolio share by year

| Year | Records | Share of all records |
|---|---:|---:|
| 2022 | 67,172 | 18.97% |
| 2023 | 118,835 | 33.56% |
| 2024 | 168,133 | 47.47% |
| **Total** | **354,140** | **100.00%** |

### Year-over-year change

| Metric | 2023 vs 2022 | 2024 vs 2023 |
|---|---:|---:|
| Premium | +89.03% | +44.40% |
| Exposure | +96.60% | +52.93% |
| Claims | +98.64% | +56.12% |
| Incurred | +91.43% | +62.24% |

**Interpretation:** The portfolio grew substantially across the period. Premium, claims and incurred all increased, but raw totals are strongly influenced by the growth in exposure and record volume. Normalised measures are needed to assess whether portfolio experience changed as well as portfolio size.

### Annual frequency, severity and loss ratio

| Year | Claim frequency | Claim severity | Loss ratio |
|---|---:|---:|---:|
| 2022 | 0.3022 | €968.82 | 65.62% |
| 2023 | 0.3053 | €933.63 | 66.46% |
| 2024 | 0.3117 | €970.22 | 74.66% |

**Interpretation:** Claim frequency increased modestly from 0.3022 to 0.3117. Severity declined in 2023 and then returned close to the 2022 level in 2024. Loss ratio rose notably in 2024, meaning incurred claims consumed a larger share of premium. The pattern warrants investigation of pricing adequacy, coverage mix, claims costs and other portfolio changes; this analysis alone cannot identify the cause.

## 10. Policy-Type Analysis

### Premium, incurred, claims and exposure by policy type

| Policy type | Premium | Incurred | Claims | Exposure |
|---|---:|---:|---:|---:|
| CC | €49,801,769.35 | €31,563,558.55 | 31,842 | 144,235.19 |
| COMP_E | €35,712,391.85 | €26,367,075.59 | 26,687 | 65,040.34 |
| COMP_N | €10,344,550.09 | €9,709,280.83 | 12,907 | 12,342.35 |
| TP | €1,019,549.10 | €553,434.48 | 507 | 3,802.67 |
| TPG | €8,201,409.03 | €5,671,461.24 | 5,156 | 24,905.35 |
| **Total** | **€105,079,669.43** | **€73,864,810.69** | **77,099** | **250,325.90** |

### Claim frequency, claim severity and loss ratio by policy type

| Policy type | Claim frequency | Claim severity | Loss ratio |
|---|---:|---:|---:|
| CC | 0.2208 | €991.26 | 63.38% |
| COMP_E | 0.4103 | €988.01 | 73.83% |
| COMP_N | 1.0457 | €752.25 | 93.86% |
| TP | 0.1333 | €1,091.59 | 54.28% |
| TPG | 0.2070 | €1,099.97 | 69.15% |
| **Portfolio** | **0.3080** | **€958.05** | **70.29%** |

**Interpretation:**

- **CC** contributes the largest premium volume.
- **COMP_N** has the highest claim frequency and highest loss ratio among the policy types. Its 93.86% loss ratio means incurred claims are close to premium before other expenses are considered. It deserves further investigation, not an automatic pricing or underwriting change.
- **TP** and **TPG** have relatively high claim severity but lower loss ratios than the portfolio overall. Severity alone therefore does not determine segment performance.
- Policy-type differences should be interpreted in the context of coverage definitions, exposure, customer mix and policy terms. The policy-type codes are retained as provided by the source and are not redefined in this project.

This comparison is documented here rather than displayed as an additional dashboard chart because dashboard space was limited.

## 11. Driver-Age Risk Segmentation

Driver age was grouped into five bands:

| Band | Driver age |
|---|---|
| Young | 18–25 |
| Early Adult | 26–35 |
| Mature | 36–50 |
| Older | 51–65 |
| Senior | 66–90 |

An Excel `IFS` formula using structured references assigned each record to a band.

### Segment totals

| Driver-age band | Premium | Claims | Incurred | Exposure |
|---|---:|---:|---:|---:|
| Young | €1,022,740.68 | 563 | €639,394.77 | 1,153.00 |
| Early Adult | €15,022,984.86 | 12,435 | €11,005,455.48 | 33,858.07 |
| Mature | €42,747,762.04 | 32,789 | €30,491,019.38 | 104,236.35 |
| Older | €35,443,687.89 | 24,308 | €23,833,985.98 | 84,090.96 |
| Senior | €10,842,493.96 | 7,004 | €7,894,955.08 | 26,987.52 |
| **Total** | **€105,079,669.43** | **77,099** | **€73,864,810.69** | **250,325.90** |

### Metrics by driver-age band

| Driver-age band | Claim frequency | Claim severity | Loss ratio |
|---|---:|---:|---:|
| Young | 0.4883 | €1,135.69 | 62.52% |
| Early Adult | 0.3673 | €885.04 | 73.26% |
| Mature | 0.3146 | €929.92 | 71.33% |
| Older | 0.2891 | €980.50 | 67.24% |
| Senior | 0.2595 | €1,127.21 | 72.81% |
| **Portfolio** | **0.3080** | **€958.05** | **70.29%** |

**Interpretation:**

- Young drivers show the highest frequency and severity among driver-age bands, while having the lowest loss ratio. That combination shows why frequency and severity should be considered alongside premium and exposure.
- Early Adult drivers have the highest driver-age loss ratio at 73.26%.
- Seniors have relatively low frequency but high severity.
- Young drivers represent a small share of total exposure, so their segment results may be less stable than those for larger groups.

The results should prompt further investigation of claims, coverage, exposure, pricing and risk mix, rather than being treated as proof that age alone causes the observed differences.

## 12. Vehicle-Age Risk Segmentation

Vehicle age was grouped into the following bands:

| Band | Vehicle age |
|---|---|
| Newer | 0–3 years |
| Mid-age | 4–7 years |
| Older | 8–12 years |
| Very old | 13+ years |
| Missing | Vehicle age is blank/missing |

The `Missing` group was retained to make the data-quality exception visible rather than classifying unknown ages as newer vehicles.

### Segment totals

| Vehicle-age band | Premium | Claims | Incurred | Exposure |
|---|---:|---:|---:|---:|
| Newer | €1,453,607.34 | 873 | €1,029,667.40 | 1,666.21 |
| Mid-age | €4,260,567.61 | 3,225 | €3,153,556.40 | 7,970.76 |
| Older | €9,608,812.29 | 7,817 | €7,171,116.02 | 22,264.78 |
| Very old | €89,755,611.44 | 65,182 | €62,509,091.12 | 218,423.14 |
| Missing | €1,070.75 | 2 | €1,379.75 | 1.00 |
| **Total** | **€105,079,669.43** | **77,099** | **€73,864,810.69** | **250,325.90** |

### Metrics by vehicle-age band

| Vehicle-age band | Claim frequency | Claim severity | Loss ratio |
|---|---:|---:|---:|
| Newer | 0.5239 | €1,179.46 | 70.84% |
| Mid-age | 0.4046 | €977.85 | 74.02% |
| Older | 0.3511 | €917.37 | 74.63% |
| Very old | 0.2984 | €958.99 | 69.64% |
| Missing | 2.0000 | €689.88 | 128.86% |
| **Portfolio** | **0.3080** | **€958.05** | **70.29%** |

**Interpretation:**

- Newer vehicles have the highest frequency and severity among the vehicle-age bands.
- Older vehicles have the highest loss ratio among the meaningful vehicle-age bands at 74.63%.
- Very old vehicles account for most of the portfolio's exposure and incurred amount, but their loss ratio is 69.64%. A large absolute incurred total does not necessarily mean worse performance relative to premium.
- The Missing band contains only two records and exposure of 1.00. Its ratios are unstable and should be treated as a data-quality exception rather than reliable evidence about vehicle risk.

Vehicle-value banding was intentionally excluded from the core project to preserve a manageable, job-focused scope. It remains a reasonable extension if future work needs to examine relationships between vehicle value, claim severity and premium adequacy.

## 13. Management Dashboard

The `Dashboard` worksheet combines portfolio KPIs, visuals and slicers in one management-facing view.

### KPI cards

The dashboard displays:

- Total records
- Unique insureds
- Total premium
- Average premium
- Total claims
- Total incurred
- Total exposure

### Visuals and performance table

The dashboard includes:

- Premium vs. incurred by year
- Policy-type premium by year
- Policy-type claims by year
- Policy-type incurred by year
- Driver-age band premium and incurred
- Annual claim frequency, claim severity and loss ratio table

### Slicers

The dashboard includes slicers for year, vehicle-age band, policy status, policy type and business type. The final workbook audit confirmed that the dashboard slicers and charts respond to their intended filters.

The overall KPI cards should be understood as whole-portfolio figures if they remain unchanged when slicers are applied. Any chart or KPI that is intended to respond to a filter should be connected to the relevant PivotTable/slicer connections.

### Dashboard design decisions

The dashboard was kept as a single-page overview rather than adding every possible chart. The loss-ratio-by-policy-type comparison is documented in Section 10 instead of adding another chart to an already populated canvas. This preserves readability while keeping the finding accessible to a reviewer.

<img width="1765" height="822" alt="image" src="https://github.com/user-attachments/assets/14b45dd1-95f4-4cdc-8eb7-a524e8f3a0d4" />


## 14. Key Findings and Recommended Follow-Up

1. **Investigate the 2024 loss-ratio increase.** Loss ratio rose from 66.46% in 2023 to 74.66% in 2024. Review the possible roles of claims costs, pricing, portfolio mix, coverage and policy changes.
2. **Review COMP_N performance.** Its loss ratio is 93.86% and frequency is 1.0457 claims per exposure-year. Check the coverage definition, exposure calculation, claim patterns and premium adequacy before taking action.
3. **Treat driver-age signals separately.** Young drivers show high frequency and severity, while Early Adult drivers have the highest loss ratio among the driver-age groups. Different signals may call for different investigative questions.
4. **Investigate vehicle-age differences.** Older vehicles have the highest meaningful vehicle-age loss ratio, while newer vehicles have the highest frequency and severity. Further work should consider coverage, usage and segment mix.
5. **Keep missing data visible.** Investigate the two missing vehicle-age values, and avoid silent imputation or assigning unknown records to an arbitrary category.
6. **Use totals and ratios together.** Premium and incurred totals describe portfolio scale; frequency, severity and loss ratio provide complementary information about experience.
7. **Avoid unsupported causal conclusions.** These are descriptive comparisons. They identify possible investigation priorities but do not prove that a driver or vehicle characteristic caused an outcome.

## 15. Excel Skills Demonstrated

This project demonstrates practical capabilities in:

- Excel Tables and structured references
- Sorting, filtering and data-type validation
- Missing-value and duplicate checks
- Data-quality flags and reconciliation
- Text cleaning and standardisation
- `SUM`, `COUNT`, `COUNTA` and aggregation logic
- `IF`, `IFS`, `AND` and `OR` concepts
- `XLOOKUP` with helper keys and multiple criteria
- Familiarity with `INDEX`/`MATCH` and `XMATCH`
- Dynamic-array functions: `FILTER`, `UNIQUE` and `SORT`
- PivotTables and grouped summaries
- Year-over-year calculations and percentage formatting
- Claim frequency, claim severity, loss ratio and premium-per-exposure calculations
- Formula-based risk-segmentation categories
- Column, bar and line charts
- PivotCharts and slicers
- Dashboard composition, number formatting and visual presentation
- Cross-checking group totals against portfolio totals
- Business interpretation and technical documentation

## 16. Scope Decisions

The project made several deliberate choices to keep the work relevant and achievable.

- **XLOOKUP is the primary lookup function.** INDEX/MATCH and XMATCH were explored for awareness, not forced into the main workflow.
- **Vehicle value banding was excluded.** Vehicle value is relevant to motor-insurance analysis, but it was not essential to complete the core project objectives and would add another segmentation layer. It can be added later as an extension.
- **Power Query was excluded from the final scope.** The project's required cleaning and analysis had already been completed. Repeating the process through another preparation tool was not necessary for the agreed deliverable.
- **No artificial dates were created.** The annual `year` field was used as provided because more detailed dates were unavailable.
- **The dashboard was not overcrowded.** Policy-type loss ratios are documented in this README rather than displayed as an additional chart.

## 17. Limitations and Interpretation Notes

- The dataset is anonymised historical data, not a live operational portfolio.
- The time dimension available for the project is annual `year`; monthly seasonality or precise policy/claim-date analysis was not attempted.
- Repeated insured IDs can represent legitimate observations across years and are not duplicates by themselves.
- Exposure is a fractional measure. Retain precision for formulas; round only for presentation.
- Claim frequency is claims divided by exposure and is not the percentage of customers with claims.
- Claim severity is incurred divided by claims and is not average incurred per record.
- Loss ratio is incurred divided by premium and is not a complete profitability measure.
- Segment metrics can be unstable when exposure or claim counts are small, particularly the Missing vehicle-age group.
- Segment differences may reflect several factors at once; this work is descriptive and does not establish causality.
- Currency symbol formatting in Excel changes the display, not the underlying numeric values. Retain the dataset source and metadata when reusing the workbook.

## 18. Portfolio and Summary

### Project description

**Motor Insurance Portfolio Analytics — Microsoft Excel**

Developed an Excel-based analysis of 354,140 motor-insurance policy-year observations covering 2022–2024. Audited and cleaned data, reconciled portfolio KPIs, calculated claim frequency, claim severity and loss ratio, and analysed trends by year, policy type, driver age and vehicle age. Used PivotTables, XLOOKUP, structured references, charts and interactive slicers to build a management dashboard. Identified an increase in annual loss ratio to 74.66% in 2024 and a policy-type loss ratio of 93.86% for COMP_N, documenting these as priorities for further investigation rather than unsupported causal conclusions.

### Skills line for portfolio

**Excel:** PivotTables, PivotCharts, slicers, XLOOKUP, structured references, dynamic-array formulas, data cleaning, data validation, KPI calculations, trend analysis, risk segmentation, reconciliation and dashboard development.

## 19. Final Project Status

| Module | Topic | Status |
|---|---|---|
| 1 | Data Audit and Cleaning | Complete |
| 2 | Insurance Calculations and KPIs | Complete |
| 3 | Advanced Lookups and Business Questions | Complete |
| 4 | Date and Portfolio Analysis | Complete |
| 5 | Risk Segmentation and Statistical Analysis | Complete |
| 6 | Data Visualization and Management Dashboard | Complete |
| 7 | Final Workbook Review and Documentation | Complete |

---

**Project outcome:** A structured Excel insurance analytics workbook with audited data, reconciled portfolio summaries, practical lookup examples, annual performance analysis, risk-segmentation results, an interactive management dashboard and professional project documentation.

## 20. Author

**Author:** Stefan Frimpong

**Role:** Data Scientist, Data Analyst & Business Analyst

**LinkedIn:** www.linkedin.com/in/stefan-frimpong-370a9b2b1

**GitHub:** https://github.com/StefanJnr

**Email:** stefanfrimpongjnr@gmail.com
