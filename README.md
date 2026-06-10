<div align="center">

# Excel Roadmap — Zero to Advanced

**The complete, structured guide to mastering Microsoft Excel.**

[![Stars](https://img.shields.io/github/stars/riya1o1/Excel-Roadmap?style=social)](https://github.com/riya1o1/Excel-Roadmap)
[![PRs Welcome](https://img.shields.io/badge/PRs-welcome-brightgreen.svg)](CONTRIBUTING.md)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

*Formulas · Charts · Pivot Tables · Power Query · VBA · Dashboards · Real Projects*

> ⭐ Star this repo if it helps you — it keeps this updated!

</div>

---

## The Learning Roadmap

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                         EXCEL LEARNING ROADMAP                              │
├───────────────┬─────────────────┬──────────────────┬────────────────────────┤
│    BEGINNER   │   INTERMEDIATE  │     ADVANCED     │        EXPERT          │
├───────────────┼─────────────────┼──────────────────┼────────────────────────┤
│ Interface     │ IF / IFS / AND  │ Pivot Tables     │ Power Query            │
│ Navigation    │ VLOOKUP         │ Advanced Charts  │ VBA & Macros           │
│ Basic Math    │ XLOOKUP         │ Dashboards       │ Power Pivot + DAX      │
│ Cell Refs     │ INDEX/MATCH     │ What-If Analysis │ LAMBDA / LET           │
│ SUM/AVG/COUNT │ Text Functions  │ Array Formulas   │ Python in Excel        │
│ Formatting    │ Date Functions  │ Power Query      │ Automation             │
│ Basic Charts  │ COUNTIF/SUMIF   │ Dynamic Arrays   │ Enterprise Models      │
│ Tables        │ Cond. Formatting│ Scenario Mgr     │ External Data Conn.    │
└───────────────┴─────────────────┴──────────────────┴────────────────────────┘
         ↓               ↓                ↓                    ↓
      2-4 weeks       4-8 weeks        2-3 months          6+ months
```

---

## Table of Contents

- [Stage 1 — Beginner](#-stage-1--beginner)
- [Stage 2 — Intermediate](#-stage-2--intermediate)
- [Stage 3 — Advanced](#-stage-3--advanced)
- [Stage 4 — Expert](#-stage-4--expert)
- [Formula Cheat Sheet](#-formula-cheat-sheet)
- [Keyboard Shortcuts](#️-keyboard-shortcuts)
- [Charts Guide](#-charts-guide)
- [Pivot Tables Deep Dive](#-pivot-tables-deep-dive)
- [Dashboard Building](#️-dashboard-building)
- [Real-World Projects](#-real-world-projects)
- [Resources](#-resources) (go to resources if prefer watching over reading)

---

## Stage 1 — Beginner

> **Goal:** Get comfortable with the interface and perform basic calculations confidently.

### 1.1 The Excel Interface

```
┌──────────────────────────────────────────────────────────────────────────┐
│   File   Home   Insert   Page Layout   Formulas   Data   Review   View   │  ← Ribbon
├──────────────────────────────────────────────────────────────────────────┤
│  Name Box │ fx │  Formula Bar: =SUM(B2:B10)                              │
├───────────┴────┴─────────────────────────────────────────────────────────┤
│     │   A          B          C          D          E                    │
│  1  │  Name      Score      Grade      City      Status                  │  ← Row 1 (header)
│  2  │  Arjun      85          B        Mumbai    Active                  │  ← Row 2
│  3  │  Priya      92          A        Delhi     Active                  │
│  4  │  Rohan      67          C        Pune      Inactive                │
│     │   ↑                                                                │
│     │ Column A                                                           │
├─────┴─────────────────────────────────────────────────────────────────── ┤
│  + Sheet1    Sheet2    Sheet3                                            │  ← Sheet tabs
└───────────────────────────────────────────────────────────────────────── ┘
         ↑
    Cell A2 = "Arjun"   |   Range B2:B4 = {85, 92, 67}
```

**Key Vocabulary:**
| Term | What it means |
|---|---|
| Cell | One box, identified by column + row (e.g. `B3`) |
| Range | A group of cells (e.g. `B2:B100`) |
| Workbook | The entire `.xlsx` file |
| Worksheet | One tab inside the workbook |
| Formula Bar | Where you see/edit the formula of selected cell |
| Name Box | Shows current cell address (e.g. `A2`) |

---

### 1.2 Cell References — The Most Important Concept

```
┌─────────────────────────────────────────────────────────┐
│                 CELL REFERENCE TYPES                    │
├────────────────┬────────────────────────────────────────┤
│    A2          │  Relative  — both shift when copied    │
│    $A2         │  Mixed     — column locked, row shifts │
│    A$2         │  Mixed     — row locked, col shifts    │
│    $A$2        │  Absolute  — nothing shifts (locked)   │
└────────────────┴────────────────────────────────────────┘

When to lock with $:
  Copying formula DOWN a column?  → lock column:   $A2
  Copying formula ACROSS a row?   → lock row:      A$2
  Referencing a fixed table?      → lock both:    $A$2
  Normal shifting calculation?    → no lock:       A2

Quick tip: Press F4 while in formula to cycle through all 4 types
```

---

### 1.3 Basic Functions

```excel
=SUM(A2:A100)        → Add all values
=AVERAGE(A2:A100)    → Mean
=MIN(A2:A100)        → Smallest value
=MAX(A2:A100)        → Largest value
=COUNT(A2:A100)      → Count numbers only
=COUNTA(A2:A100)     → Count non-empty cells (any type)
=ROUND(A2, 2)        → Round to 2 decimal places
=ABS(A2)             → Remove negative sign
=SQRT(A2)            → Square root
=MOD(A2, 3)          → Remainder after dividing by 3
```

---

### 1.4 Number Formatting

```
┌────────────────┬──────────────────┬────────────────────┐
│ Format Code    │ Raw Value        │ Displayed As       │
├────────────────┼──────────────────┼────────────────────┤
│ #,##0          │ 150000           │ 1,50,000           │
│ ₹#,##0         │ 50000            │ ₹50,000            │
│ 0.00%          │ 0.756            │ 75.60%             │
│ DD-MMM-YYYY    │ 45505            │ 01-Sep-2025        │
│ 0.0            │ 6                │ 6.0                │
│ "Yes";"No"     │ 1 / 0            │ Yes / No           │
└────────────────┴──────────────────┴────────────────────┘
Access via: Ctrl+1 → Number tab → Custom
```

---

## Stage 2 — Intermediate

> **Goal:** Write real-world formulas, work with logic, look up data across sheets.

### 2.1 Logical Functions

```
Decision tree for choosing the right logical function:

One condition?
  └── =IF(condition, true_result, false_result)
       =IF(A2>=60, "Pass", "Fail")

Multiple conditions, all must be true?
  └── =IF(AND(A2>18, B2="Active"), "Eligible", "No")

Any one condition true is enough?
  └── =IF(OR(A2="Delhi", A2="Mumbai"), "Metro", "Other")

More than 3 outcomes?
  └── =IFS(A2>=90,"A", A2>=75,"B", A2>=60,"C", TRUE,"F")

Formula might error?
  └── =IFERROR(your_formula, 0)   ← always wrap risky formulas
```

**IFERROR — wrap every risky formula:**
```excel
=IFERROR(A2/B2, 0)                     → 0 instead of #DIV/0!
=IFERROR(VLOOKUP(...), "Not Found")    → clean message instead of #N/A
=IFERROR(INDEX(MATCH(...)), 0)         → 0 if lookup fails
```

---

### 2.2 Lookup Functions — Comparison

```
┌──────────────────────────────────────────────────────────────────────┐
│                   LOOKUP FUNCTION COMPARISON                         │
├─────────────────┬────────────┬──────────────┬────────────────────────┤
│ Feature         │ VLOOKUP    │ XLOOKUP      │ INDEX/MATCH            │
├─────────────────┼────────────┼──────────────┼────────────────────────┤
│ Look direction  │ Right only │ Any direction│ Any direction          │
│ Column insert   │ Breaks     │ Safe         │ Safe                   │
│ Error handling  │ Needs IFERR│ Built-in     │ Needs IFERROR          │
│ Return multiple │ No         │ Yes          │ One at a time          │
│ Excel version   │ All        │ 2019+ only   │ All versions           │
│ Speed (large)   │ Slow       │ Fast         │ Fast                   │
│ Syntax ease     │ Medium     │ Easy         │ Harder                 │
└─────────────────┴────────────┴──────────────┴────────────────────────┘
Recommendation:
  Excel 365 / 2021  →  Use XLOOKUP
  Excel 2016/2019   →  Use INDEX/MATCH
  Legacy files      →  VLOOKUP is fine for simple cases
```

**VLOOKUP:**
```excel
=VLOOKUP(lookup_value, table_array, col_index, FALSE)
           ↑                ↑            ↑         ↑
        find this      search here   return col  exact match

=VLOOKUP(A2, $F$2:$H$100, 2, FALSE)
→ Find A2 in column F, return value from column G (2nd col of range)
```

**XLOOKUP:**
```excel
=XLOOKUP(lookup, lookup_array, return_array, [if_not_found])

=XLOOKUP(A2, F:F, G:G, "Not Found")          → basic
=XLOOKUP(A2, F:F, G:H, "Not Found")          → return 2 columns at once
=XLOOKUP(1,(C:C="IT")*(D:D>50000),B:B)       → multiple conditions
```

**INDEX / MATCH:**
```excel
=INDEX(return_range, MATCH(lookup_value, lookup_range, 0))

How it works step by step:
  Step 1: MATCH("Arjun", A2:A100, 0)  →  finds position 5
  Step 2: INDEX(C2:C100, 5)           →  returns value at row 5

=IFERROR(INDEX(C2:C100, MATCH(A2, A2:A100, 0)), "Not Found")
```

---

### 2.3 COUNTIF / SUMIF Family

```excel
── Count ──────────────────────────────────────────────────────────────
=COUNTIF(range, criteria)
  =COUNTIF(D:D, "Completed")          → count "Completed" in col D
  =COUNTIF(B:B, ">50000")             → count values over 50k
  =COUNTIF(A:A, "Arjun*")            → wildcard: starts with "Arjun"

=COUNTIFS(range1, crit1, range2, crit2)
  =COUNTIFS(D:D,"IT", E:E,">40000")  → IT dept AND salary > 40k

── Sum ────────────────────────────────────────────────────────────────
=SUMIF(range, criteria, sum_range)
  =SUMIF(C:C, "Mumbai", E:E)          → sum salaries where city=Mumbai

=SUMIFS(sum_range, range1, crit1, range2, crit2)
  =SUMIFS(E:E, C:C,"Delhi", D:D,"IT")→ IT dept salaries in Delhi

── Average ────────────────────────────────────────────────────────────
=AVERAGEIF(range, criteria, avg_range)
  =AVERAGEIF(D:D, "Finance", E:E)     → avg salary in Finance dept
```

---

### 2.4 Text Functions

```excel
── Extract ────────────────────────────────────────────────────────────
=LEFT(A2, 3)              → first 3 characters
=RIGHT(A2, 4)             → last 4 characters
=MID(A2, 3, 5)            → 5 chars starting from position 3
=LEN(A2)                  → total number of characters

── Clean ──────────────────────────────────────────────────────────────
=TRIM(A2)                 → remove extra spaces
=UPPER(A2)                → ALL CAPS
=LOWER(A2)                → all lowercase
=PROPER(A2)               → Title Case

── Combine ────────────────────────────────────────────────────────────
=A2 & " " & B2            → join: "Arjun Sharma"
=TEXTJOIN(", ",TRUE,A2:A5)→ join range: "Arjun, Priya, Rohan"

── Find & Replace ─────────────────────────────────────────────────────
=FIND("@", A2)             → position of @ in email (case-sensitive)
=SUBSTITUTE(A2,"old","new")→ replace text inside string
=SUBSTITUTE(A2,"old","char(160)")→ unbreakable space
        
── Real example: extract domain from email ────────────────────────────
Email in A2: arjun@gmail.com
=MID(A2, FIND("@",A2)+1, LEN(A2))  →  "gmail.com"
```

---

### 2.5 Date Functions

```excel
=TODAY()                         → today's date (live, updates daily)
=NOW()                           → current date + time
=YEAR(A2) / MONTH(A2) / DAY(A2) → extract parts of a date
=DATEDIF(start, end, "D")        → days between two dates
=DATEDIF(start, end, "M")        → months between two dates
=DATEDIF(start, end, "Y")        → years between two dates
=NETWORKDAYS(A2, B2)             → working days between dates
=WORKDAY(A2, 10)                 → date 10 working days later
=EOMONTH(A2, 0)                  → last day of the same month
=TEXT(A2, "DD-MMM-YYYY")         → format date as readable text

Practical:
=A2 - TODAY()                    → days remaining until deadline
```

---

### 2.6 Conditional Formatting

```
Types and when to use them:

┌────────────────────┬──────────────────────────────────────────────────┐
│ Type               │ Best For                                         │
├────────────────────┼──────────────────────────────────────────────────┤
│ Color Scale        │ Gradient across a range (e.g. % used: green→red) │
│ Data Bars          │ Mini bar chart inside cell (progress)            │
│ Icon Sets          │ Traffic lights / arrows for KPI status           │
│ Highlight Rules    │ Flag specific values ("Overdue" → red)           │
│ Formula-Based      │ Highlight ENTIRE ROW based on one column's value │
└────────────────────┴──────────────────────────────────────────────────┘

Most powerful — Formula-Based (highlights entire row):
  Apply to: $A$2:$J$100
  Formula:  =$G2="Overdue"
  Effect:   entire row turns red when column G = "Overdue"

3-color scale for % used (green → yellow → red):
  Min:  0   = Green  (#63BE7B)
  Mid: 70%  = Yellow (#FFEB84)
  Max: 100% = Red    (#F8696B)
```

---

### 2.7 Data Validation (Dropdowns)

```
Why it matters:
  Without validation:  User types "compelted" → COUNTIF returns 0 → KPI wrong
  With validation:     User picks from list   → data is always clean

Setup:
  Data → Data Validation → Settings → List
  Source: "Paid,Pending,Partial,Cancelled"
  Or reference a range: =$H$1:$H$4 (easier to update later)

Common validations:
  Status dropdowns    → "Not Started,In Progress,Completed,Overdue"
  Priority dropdowns  → "High,Medium,Low"
  Date range          → only allow future dates: =A2 >= TODAY()
  No duplicates       → Custom formula: =COUNTIF($A$2:$A$100,A2)=1
  Whole numbers only  → Whole Number between 1 and 100
```

---

## Stage 3 — Advanced

> **Goal:** Build real dashboards, automate reporting, analyse complex data.

### 3.1 Pivot Tables — Complete Guide

```
Raw Data → Pivot Table → Instant Summary

Example:
┌─────────┬──────────┬─────────┬────────┐       ┌───────────┬────────────┐
│ Name    │ Region   │ Product │ Sales  │       │ Region    │ Total Sales│
├─────────┼──────────┼─────────┼────────┤  →    ├───────────┼────────────┤
│ Arjun   │ North    │ Laptop  │ 45000  │       │ North     │ 83,000     │
│ Priya   │ South    │ Phone   │ 25000  │       │ South     │ 25,000     │
│ Rohan   │ North    │ Laptop  │ 38000  │       │ Grand Tot │ 1,08,000   │
└─────────┴──────────┴─────────┴────────┘       └───────────┴────────────┘

The 4 Drop Zones:
┌────────────────────────────────────────────┐
│  FILTERS  →  top-level filter (e.g. Year)  │
├────────────────┬───────────────────────────┤
│  ROWS          │  COLUMNS                  │
│  (group by)    │  (second grouping)        │
├────────────────┴───────────────────────────┤
│  VALUES  →  what to calculate              │
│  (Sum, Count, Average, %, Running Total)   │
└────────────────────────────────────────────┘
```

**Power user tricks:**
```
Right-click → Show Values As → % of Grand Total   → instant % breakdown
Right-click → Group          → group dates by Month/Quarter/Year
Insert       → Slicer         → visual click-to-filter buttons
Insert       → Timeline       → drag-to-filter date range slider

Calculated Field (add a formula inside the pivot):
  PivotTable Analyze → Fields, Items & Sets → Calculated Field
  Name: Profit Margin
  Formula: = Profit / Revenue
```

**Pivot Table rules (must follow):**
```
Format source data as Table first (Ctrl+T)
No blank rows or merged cells in source
Unique column headers in row 1
Refresh after data changes (right-click → Refresh)
```

---

### 3.2 Advanced Charts

```
Which chart for which question?

┌────────────────────────────────────────────────────────────────────┐
│  What are you trying to show?          Best Chart                  │
├────────────────────────────────────────────────────────────────────┤
│  Compare values (A vs B vs C)        → Column / Bar                │
│  Change over time                    → Line Chart                  │
│  Part of a whole (proportions)       → Pie / Doughnut(max 6 slices)│
│  Relationship / correlation          → Scatter Plot                │
│  Distribution / spread               → Histogram                   │
│  Progress toward a target            → Stacked Bar                 │
│  Two metrics on different scales     → Combo (Bar + Line)          │
│  Data in geography                   → Map Chart                   │
│  Trend within a row                  → Sparklines (in-cell mini)   │
└────────────────────────────────────────────────────────────────────┘
```

**Combo Chart (Bar + Line with dual axis) — very common in dashboards:**
```
Use case: Revenue (₹) and Growth % on the same chart

Steps:
  1. Select data with both metrics
  2. Insert → Recommended Charts → All Charts → Combo
  3. Revenue  → Clustered Bar  → Primary Axis
  4. Growth % → Line           → Secondary Axis (tick this box)
```

**Chart formatting checklist:**
```
Clear, descriptive title (not just "Chart 1")
Axis labels with units (₹, %, count)
No 3D effects — they distort proportions
Minimal gridlines — less clutter
Consistent colour scheme throughout dashboard
Dynamic title: click title → type = → click a cell
```

---

### 3.3 Dynamic Array Functions (Excel 365)

```excel
── Spill formulas (auto-expand to fill results) ──────────────────────
=UNIQUE(A2:A100)                → list all unique values
=SORT(A2:A100)                  → sorted list ascending
=SORT(A2:A100, 1, -1)           → sorted descending
=FILTER(A2:C100, B2:B100="IT")  → rows where dept = IT
=SEQUENCE(10)                   → generate 1, 2, 3 … 10

── Combine them ──────────────────────────────────────────────────────
=SORT(                          → sort the result of filter
  FILTER(A2:C100,
    C2:C100>50000),             → filter: salary > 50k
  3, -1)                        → sort by col 3 descending

── XLOOKUP with multiple conditions ──────────────────────────────────
=XLOOKUP(1,
  (C2:C100="IT") * (D2:D100>50000),
  B2:B100)
→ Find first person in IT with salary > 50k
```

---

### 3.4 What-If Analysis

```
Three tools for "what if" scenarios:

┌──────────────────┬─────────────────────────────────────────────────┐
│ Tool             │ Question it answers                             │
├──────────────────┼─────────────────────────────────────────────────┤
│ Goal Seek        │ "What input gives me this output?"              │
│                  │ e.g. "What sales do I need to hit ₹1L profit?"  │
├──────────────────┼─────────────────────────────────────────────────┤
│ Data Table       │ "How does changing 1-2 inputs affect output?"   │
│                  │ e.g. Sensitivity table: price × volume → profit │
├──────────────────┼─────────────────────────────────────────────────┤
│ Scenario Manager │ "Compare multiple named scenarios side-by-side" │
│                  │ e.g. Base case vs Best case vs Worst case       │
└──────────────────┴─────────────────────────────────────────────────┘

Goal Seek:
  Data → What-If Analysis → Goal Seek
  Set cell:    B10  (result formula)
  To value:    100000
  By changing: A2   (input cell)
```

---

### 3.5 Power Query — Data Cleaning

```
The best tool for messy data. Applies steps non-destructively.

Data → Get Data → From File / From Web / From Database

Common transformations:
┌───────────────────────────────┬──────────────────────────────────────┐
│ Problem                       │ Power Query Solution                 │
├───────────────────────────────┼──────────────────────────────────────┤
│ Duplicate rows                │ Home → Remove Rows → Remove Dups     │
│ Extra spaces / weird chars    │ Transform → Clean / Trim             │
│ Column split by comma         │ Transform → Split Column             │
│ Wide → tall (unpivot)         │ Transform → Unpivot Columns          │
│ Two tables joined             │ Home → Merge Queries (like SQL JOIN) │
│ Blanks filled from above      │ Transform → Fill → Down              │
│ Calculated column             │ Add Column → Custom Column           │
└───────────────────────────────┴──────────────────────────────────────┘
```

---

## Stage 4 — Expert

> **Goal:** Automate everything, build enterprise models, connect to external systems.

### 4.1 VBA — Automate Repetitive Tasks

```vba
' Open with: Alt + F11

Sub FormatOverdueTasks()
    Dim i As Integer
    For i = 2 To 100
        If Cells(i, 7).Value = "Overdue" Then
            Rows(i).Interior.Color = RGB(255, 200, 200)  ' light red
            Cells(i, 7).Font.Bold = True
        End If
    Next i
    MsgBox "Formatting complete!"
End Sub

' Common operations:
Range("B2").Value = 100              ' write to cell
Cells(2, 3).Value = "Mumbai"         ' row 2, col 3
Sheets("Data").Copy                  ' copy a sheet
Cells.EntireColumn.AutoFit           ' autofit all columns
ActiveWorkbook.Save                  ' save
```

---

### 4.2 Power Pivot + DAX

```
For datasets with millions of rows or multiple related tables.

DAX examples:
  Total Sales     = SUM(Sales[Amount])
  Sales % Total   = DIVIDE([Total Sales], CALCULATE([Total Sales], ALL(Sales)))
  YTD Sales       = TOTALYTD([Total Sales], Dates[Date])
  Running Total   = CALCULATE([Total Sales],
                     FILTER(ALL(Dates), Dates[Date] <= MAX(Dates[Date])))
```

---

### 4.3 Advanced Formulas (Excel 365)

```excel
── LET: define variables inside a formula ────────────────────────────
=LET(
  sales,  B2:B100,
  target, C2:C100,
  ratio,  sales/target,
  IF(ratio >= 1, "Hit", "Missed")
)

── LAMBDA: create your own reusable function ─────────────────────────
=LAMBDA(price, tax_rate, price * (1 + tax_rate))
→ name it "AddTax" in Name Manager, then use =AddTax(A2, 0.18)
```

---

## Formula Cheat Sheet

### Lookup & Reference
| Formula | Purpose |
|---|---|
| `=XLOOKUP(val, look, return, "NA")` | Modern lookup — use this first (2019+) |
| `=INDEX(range, MATCH(val, range, 0))` | Classic — works in all Excel versions |
| `=VLOOKUP(val, table, col, FALSE)` | Simple right-only lookup |
| `=OFFSET(cell, rows, cols, h, w)` | Dynamic range reference |
| `=INDIRECT("Sheet1!A"&ROW())` | Build cell reference from text |

### Logic & Conditions
| Formula | Purpose |
|---|---|
| `=IF(cond, true, false)` | Basic condition |
| `=IFS(c1,v1, c2,v2, TRUE,default)` | Multiple conditions — cleaner than nested IF |
| `=AND(c1, c2)` | All conditions must be true |
| `=OR(c1, c2)` | Any condition true is enough |
| `=IFERROR(formula, 0)` | Handle any error gracefully |
| `=IFNA(formula, "NA")` | Handle only #N/A errors |
| `=SWITCH(val, c1,r1, c2,r2, default)` | Match exact value to result |

### Math & Aggregation
| Formula | Purpose |
|---|---|
| `=SUMIF(r, crit, sum_r)` | Sum with one condition |
| `=SUMIFS(sum_r, r1,c1, r2,c2)` | Sum with multiple conditions |
| `=COUNTIF(r, crit)` | Count matching cells |
| `=COUNTIFS(r1,c1, r2,c2)` | Count with multiple conditions |
| `=AVERAGEIF(r, crit, avg_r)` | Average with condition |
| `=ROUND(val, n)` | Round to N decimals |
| `=MAX(a, 0)` | Floor at zero — never negative |
| `=LARGE(range, 3)` | 3rd largest value |
| `=RANK(val, range, 0)` | Rank (0 = descending) |

### Text
| Formula | Purpose |
|---|---|
| `=A1 & " " & B1` | Join text |
| `=TEXTJOIN(",", TRUE, range)` | Join range with separator |
| `=LEFT/RIGHT/MID` | Extract substring |
| `=FIND/SEARCH` | Find position (SEARCH is case-insensitive) |
| `=SUBSTITUTE(txt, old, new)` | Find and replace within text |
| `=TRIM(text)` | Remove extra spaces |
| `=TEXT(value, "format")` | Format number as display text |
| `=VALUE(text)` | Convert text string to number |

### Date & Time
| Formula | Purpose |
|---|---|
| `=TODAY()` | Today's date (live) |
| `=DATEDIF(start, end, "D")` | Days between dates |
| `=NETWORKDAYS(start, end)` | Working days between dates |
| `=EOMONTH(date, 0)` | Last day of the month |
| `=TEXT(date, "DD-MMM-YYYY")` | Format date as readable text |
| `=TIMEVALUE("14:30")*24` | Convert time text to decimal hours |

---

## Keyboard Shortcuts

### Essential (memorise these first)
| Shortcut | Action |
|---|---|
| `Ctrl + C / V / X` | Copy / Paste / Cut |
| `Ctrl + Z / Y` | Undo / Redo |
| `Ctrl + S` | Save |
| `Ctrl + F / H` | Find / Find & Replace |
| `Ctrl + Arrow` | Jump to edge of data block |
| `Ctrl + Shift + Arrow` | Select to edge of data block |
| `Ctrl + Home / End` | Go to A1 / last used cell |
| `Ctrl + Page Up/Down` | Switch between sheets |
| `F4` | Toggle `$` in formula (cycles all 4 types) |
| `Alt + =` | AutoSum selected range |

### Formatting
| Shortcut | Action |
|---|---|
| `Ctrl + 1` | Format Cells dialog |
| `Ctrl + B / I / U` | Bold / Italic / Underline |
| `Ctrl + Shift + $` | Currency format |
| `Ctrl + Shift + %` | Percentage format |
| `Alt + Enter` | New line within a cell |
| `Ctrl + T` | Create Table from selection |

### Data & Formulas
| Shortcut | Action |
|---|---|
| `F2` | Edit selected cell |
| `Ctrl + D / R` | Fill Down / Fill Right |
| `Ctrl + Shift + L` | Toggle AutoFilter |
| `Ctrl + ;` | Insert today's date |
| `Ctrl + '` | Copy formula from cell above |
| `Alt + F11` | Open VBA Editor |
| `Ctrl + G` | Go To (use Special for blanks, formulas etc.) |

---

## Charts Guide

### Choosing the Right Chart

```
What are you showing?
│
├─ Comparing values (A vs B)
│    └── Column Chart (vertical bars) — good for categories
│    └── Bar Chart (horizontal) — better when labels are long
│
├─ Trend over time
│    └── Line Chart — connect the dots, show direction
│
├─ Part of a whole
│    └── Pie Chart — use when ≤5 slices, all add to 100%
│    └── Doughnut — same as pie, with a hole (looks cleaner)
│
├─ Correlation / relationship
│    └── Scatter Plot — one variable per axis
│    └── Bubble Chart — 3rd dimension as bubble size
│
├─ Distribution
│    └── Histogram — frequency of values in ranges
│
├─ Two metrics with different scales
│    └── Combo Chart (Bar + Line, dual Y-axis)
│         e.g. Revenue (₹) vs Growth Rate (%)
│
└─ Trend in a single cell
     └── Sparklines — tiny chart inside a cell (Insert → Sparklines)
```

---

## Pivot Tables Deep Dive

### The 4 Drop Zones Explained

```
┌──────────────────────────────────────────────────────────────────┐
│  FILTERS  →  Slice the entire table (e.g. show only Year=2025)   │
├─────────────────────┬────────────────────────────────────────────┤
│  ROWS               │  COLUMNS                                   │
│  Department         │  Q1      Q2      Q3      Q4                │
├─────────────────────┼────────────────────────────────────────────┤
│  Finance            │ 45,000  52,000  48,000  61,000             │
│  IT                 │ 38,000  41,000  55,000  49,000             │
│  Marketing          │ 22,000  28,000  31,000  35,000             │
│  Grand Total        │  ...     ...     ...     ...               │
└─────────────────────┴────────────────────────────────────────────┘
                VALUES = Sum of Sales (₹)

Other VALUE calculations: Count, Average, Min, Max,
                          % of Grand Total, Running Total, Rank
```

### Slicers & Timelines

```
Slicer (any field):
  Insert → Slicer → pick field
  Click buttons to filter — works across multiple pivot tables

Timeline (date fields only):
  Insert → Timeline → pick date field
  Drag to select date range — intuitive and visual

Connect one slicer to multiple pivots:
  Right-click slicer → Report Connections → tick all pivot tables
  Now one click filters everything on the dashboard
```

---

## Dashboard Building

### Layout Blueprint

```
┌─────────────────────────────────────────────────────────────────────┐
│     EVENT DASHBOARD   |  Annual Tech Fest 2025  |  Updated: Today   │  ← Title
├────────────┬────────────┬────────────┬───────────────────────────── ┤
│ Budget     │Utilised    │ Vols       │  Tasks Done                  │  ← KPI Row
│ ₹3,70,000  │  72.5%     │  8/10      │  60%                         │
├────────────┴────────────┼────────────┴───────────────────────────── ┤
│                         │                                           │
│   Bar Chart             │   Doughnut Chart                          │  ← Visual Row
│   Allocated vs Spent    │   Task Status Split                       │
│                         │                                           │
├─────────────────────────┼───────────────────────────────────────────┤
│      Summary Table      │     Alerts                                │  ← Detail Row
│  (key rows, top 10)     │   AV Tech understaffed                    │
│                         │   Decoration budget at 96%                │
│                         │  Registration fully staffed               │
└─────────────────────────┴───────────────────────────────────────────┘

Rules:
  - Everything visible without scrolling (1 screen)
  - KPIs always at top
  - Charts in the middle
  - Detail / drill-down at bottom
  - Alerts panel for exceptions only
```

### 5 Techniques That Make Dashboards Professional

```
1. Named Ranges
   Instead of: =SUM($B$4:$B$9)
   Use:        =SUM(BudgetSpent)       ← readable, won't break if rows move
   Define via: Formulas → Name Manager

2. Camera Tool (live picture of a range)
   Select a range → Copy → Paste Special → Linked Picture
   Paste a "live photo" of any range onto the dashboard sheet

3. Form Controls (interactive without VBA)
   Developer → Insert → Dropdown/Scrollbar/Checkbox
   Link to a cell → use that cell in CHOOSE/INDEX to drive charts

4. Freeze + Hide
   View → Freeze Panes → keeps headers visible while scrolling
   Right-click column header → Hide → hides helper/calculation columns

5. Consistent colour system
   Pick 2-3 brand colours and stick to them everywhere
   Use one colour for headers, one for positive, one for negative
```

---

## Real-World Projects

### Beginner Projects
| Project | Key Skills |
|---|---|
| Personal Budget Tracker | SUM, IF, Pie Chart, Data Validation |
| Student Grade Calculator | AVERAGE, IF, RANK, Conditional Formatting |
| Attendance Register | COUNTIF, Percentage, Bar Chart |

### Intermediate Projects
| Project | Key Skills |
|---|---|
| Sales Performance Report | XLOOKUP, SUMIFS, Pivot Table, Slicers |
| HR Headcount Dashboard | COUNTIFS, DATEDIF, Combo Chart, Conditional Fmt |
| Inventory Tracker | VLOOKUP, MIN/MAX, Alerts, Data Validation |

### Advanced Projects
| Project | Key Skills |
|---|---|
| **Event Management Dashboard** *(this repo)* | INDEX/MATCH, 9 Charts, Pivot Summary, Live KPIs |
| Financial Model (P&L) | Scenario Manager, Data Tables, Goal Seek, Power Query |
| HR Analytics Dashboard | Power Query, Power Pivot, DAX, Slicers, Timeline |

---

### Featured Project — Event Management Dashboard

```
┌─────────────────────────────────────────────────────────────────┐
│              EVENT MANAGEMENT DASHBOARD STRUCTURE               │
├─────────────────────────────────────────────────────────────────┤
│                                                                 │
│  Dashboard  ←── reads from all sheets, shows KPIs + charts      │
│       ↑                                                         │
│       ├── BudgetTracker    (allocated, spent, invoices)         │
│       ├── VolunteerAlloc   (roles, shifts, hours)               │
│       ├── TaskTracker      (owner, status, priority, % done)    │
│       ├── ResourcePlan    (needed, available, suppliers)        │
│       └── Pivot Summary    (aggregated analytics)               │
│                                                                 │
│  Formulas used:                                                 │
│  • INDEX/MATCH   — cross-sheet lookups                          │
│  • COUNTIF       — live KPI counts                              │
│  • IFERROR       — error-proof calculations                     │
│  • MAX(x,0)      — shortfall without negatives                  │
│  • TIMEVALUE     — auto-calculate volunteer hours               │
│  • Conditional Formatting — colour scales + row highlighting    │
│  • Data Validation — dropdowns on every status column           │
│  • 9 Embedded Charts — bar, pie, doughnut across 6 sheets       │
└─────────────────────────────────────────────────────────────────┘
```

---

## Resources
Just watch this one shot video if in hurry - (https://www.youtube.com/watch?v=XbyiTh-6k9Q&t=16696s)
### Practice Datasets
- [Kaggle Datasets](https://kaggle.com/datasets) — real-world data for practice
- [data.gov.in](https://data.gov.in) — Indian government open datasets
- [Superstore Dataset](https://www.kaggle.com/datasets/vivek468/superstore-dataset-final) — classic Excel practice data

### Official Docs
- [Microsoft Excel Function Reference](https://support.microsoft.com/excel) — every function explained
- [Excel Tech Community](https://techcommunity.microsoft.com/excel) — forums and tips

### YouTube Channels
| Channel | Best For |
|---|---|
| ExcelIsFun | Deep formula tutorials |
| Leila Gharani | Dashboards and advanced techniques |
| MyOnlineTrainingHub | Power Query and Power Pivot |
| MrExcel | Quick tips and tricks |

### Books
- *Excel: The Missing Manual* — Matthew MacDonald
- *Excel Power Programming with VBA* — John Walkenbach
- *M Is for Data Monkey* — Power Query deep dive

---

## Contributing

Contributions are welcome! If you find an error, want to add a topic, or have a better example:

1. Fork this repo
2. Create a branch: `git checkout -b improve/pivot-tables-section`
3. Make your changes
4. Submit a Pull Request

---

## License

MIT License — free to use, share, and adapt with attribution.

---

<div align="center">

**If this helped you, please ⭐ star the repo — it helps others find it!**

*Built with ❤️ for the Excel community*

</div>
