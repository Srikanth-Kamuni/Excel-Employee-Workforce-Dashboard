# Employee Workforce Performance Dashboard

Interactive Excel workforce reporting project based on **785 employee records** and 19 fields.

## Business problem
HR stakeholders need a consistent way to monitor employee status, departures, workforce distribution and tenure across organizational segments. This dashboard consolidates descriptive workforce measures in Excel with PivotTables, slicers and charts.

## Workbook inventory
- `Employee_Data`: source records (785 rows; 19 columns).
- `Pivot`: summary PivotTables and KPI calculations.
- `Dashboard`: presentation and interactive reporting.
- The `.xlsm` file contains seven chart definitions, five slicer caches and an embedded VBA project.

## Headline KPIs
| KPI | Result | Definition |
|---|---:|---|
| Employee records | 785 | Count of employee records |
| Active | 561 | Employment Status = Active |
| Attrition | 224 | Employment Status = Attrition |
| Attrition share | 28.54% | 224 / 785 |
| Average tenure | 5.09 years | Average tenure from workbook PivotTable |

**Interpretation:** Attrition share is a share of records, not necessarily an annual turnover rate. Data is descriptive and does not establish causal drivers.

## Dataset terminology
Employee ID, Employee Name, Gender, Age, Department, Job Level, Location, Employment Type, Hire Date, Exit Date, Employment Status, Annual Salary (INR), Performance Rating, Training Hours, Attendance %, Promotion in Last 12 Months, Year, Month, Tenure. Full definitions are in the project documentation.

## Core calculations
```excel
=GETPIVOTDATA("Employee ID",Pivot!$D$5)
=GETPIVOTDATA("Employee ID",Pivot!$D$5,"Employment Status","Active")
=GETPIVOTDATA("Employee ID",Pivot!$D$5,"Employment Status","Attrition")
=IFERROR(Pivot!B8/Pivot!B6,0)
=GETPIVOTDATA("Tenure",Pivot!$D$17)
```

## Rebuild instructions
See **Employee_Workforce_Rebuild_Guide.pdf** for detailed steps covering source preparation, PivotTables, KPI formulas, charts, slicers, report connections, example VBA fullscreen/refresh macros, testing and GitHub publication.

## Requirements
Microsoft Excel desktop with PivotTable, slicer and VBA support. Save the workbook as `.xlsm`. Enable macros only in trusted files.

## Data quality and privacy
The source contains employee names and other potentially identifying information. **Do not publish the original employee-level workbook publicly** without appropriate authorization and anonymization. Validate reporting dates, tenure calculations, status definitions and slicer connections before interpreting metrics. The VBA snippets in the guide are illustrative replacements; the original embedded macro source has not been verified.

## Repository structure
```
Excel-Employee-Workforce-Dashboard/
├── README.md
├── docs/
│   ├── Employee_Workforce_Project_Documentation.pdf
│   └── Employee_Workforce_Rebuild_Guide.pdf
├── workbook/
│   └── Employee_Work_Force_Dashboard_SANITIZED.xlsm
└── images/
    └── dashboard.png
```
The sanitized workbook and screenshot are suggested future assets, not included in this documentation package.

# Dashboard Preview
1. Summary Dashboard
##[Employee Work Force Dashboard] <img width="1838" height="845" alt="Screenshot 2026-10-31 195653" src="https://github.com/user-attachments/assets/e8d2cc51-cf6f-4041-a312-3e3813ecf744" />


## Portfolio description
Built an interactive Excel workforce dashboard using 785 employee records to report employee status, attrition share, tenure and workforce composition. Applied PivotTables, GETPIVOTDATA, charts and slicers, with a documented workflow for rebuilding the analysis and adding VBA-powered refresh and fullscreen controls.
