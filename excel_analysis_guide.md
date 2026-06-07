# Excel Analysis Guide – Sickle Cell Home Pain Management Study

**Analyst:** Ashima Singh, MPH  
**Tool:** Microsoft Excel (Data Collection, Tracking, and Analysis)  
**Purpose:** Documents the Excel-based data management and analysis workflow used for this study.

---

## Overview

All participant data were collected via online survey and tracked in Microsoft Excel. Excel served as the primary tool for:
- Participant enrollment tracking
- Survey response entry and validation
- Descriptive analysis (frequencies, means, cross-tabulations)
- Chart and visualization creation for study findings

---

## Workbook Structure

The master Excel workbook (`scd_study_master.xlsx`) was organized across the following sheets:

| Sheet Name | Contents |
|---|---|
| `Participant_Tracker` | Enrollment log, recruitment source, consent status, survey completion flag |
| `Survey_Data` | Full survey response entry (one row per participant) |
| `Descriptive_Stats` | Summary statistics calculated using Excel formulas |
| `Strategy_Analysis` | Pain management strategy frequency counts and percentages |
| `Utilization_Analysis` | ED visits, hospitalizations, and missed work days analysis |
| `Charts` | All visualizations generated for reporting |
| `Data_Dictionary` | Variable definitions, coding, and notes |

---

## Key Excel Formulas Used

### Frequency and Count
```excel
=COUNTIF(D2:D50,"Severe")           → Count of severe disease severity
=COUNTIF(L2:L50,1)/COUNTA(L2:L50)  → Proportion using a specific strategy
=AVERAGEIF(F2:F50,"Female",G2:G50) → Average pain frequency among females
```

### Summary Statistics
```excel
=AVERAGE(H2:H50)    → Mean pain intensity (VAS)
=MEDIAN(R2:R50)     → Median self-efficacy score
=STDEV(S2:S50)      → Standard deviation of ED visits
=MIN(T2:T50)        → Minimum hospitalizations
=MAX(T2:T50)        → Maximum hospitalizations
```

### Strategy Count (Sum across binary columns)
```excel
=SUM(L2:R2)         → Total strategies used by participant (row-wise)
```

### Pivot Tables
Used to cross-tabulate:
- Strategy use by disease severity
- ED visits by insurance type
- Self-efficacy score ranges by recruitment source

---

## Recruitment Source Tracking

Participant recruitment source was tracked across all entries:

| Source | Description |
|---|---|
| Instagram | Study flyer posted on Instagram (static image + story) |
| Facebook | Study flyer posted in SCD support groups on Facebook |
| YouTube Video | Short study explainer video posted on YouTube |
| Flyer – Clinic Waiting Room | Physical flyers placed in healthcare waiting areas |
| Flyer – Community Center | Physical flyers distributed at SCD community events |

Recruitment source was recorded at enrollment and used to calculate cost-effectiveness of each channel.

---

## Charts Created in Excel

All charts created using Excel's built-in chart tools and formatted for reporting:

- **Clustered bar chart:** Pain management strategy prevalence (% of participants)
- **Pie chart:** Disease severity distribution
- **Bar chart:** Recruitment source breakdown
- **Column chart:** Average ED visits by disease severity
- **Scatter plot:** Self-efficacy score vs. number of ED visits
- **Stacked bar:** Strategy use by severity group

---

## Data Quality Checks Performed

- Duplicate participant IDs checked using `=COUNTIF($A$2:$A$50,A2)>1` (flagged in red)
- Out-of-range values identified using conditional formatting (e.g., VAS > 10 flagged)
- Consent column verified: all included participants must have `consent_obtained = 1`
- Survey completion column verified before inclusion in analysis

---

*Note: Raw participant data is not included in this repository to protect participant confidentiality. This file documents the analytical workflow only.*
