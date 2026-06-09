# Smart Event Management & Resource Allocation System

> An Excel-based operations dashboard to manage event budgets, volunteer allocation, resource planning, and task tracking for large-scale college events.

---

## Features

- **Interactive Dashboard** — Live KPI cards, charts, and alert panels that update automatically
- **Budget Tracker** — Line-item budget management with allocated vs spent tracking, invoice numbers, and payment status
- **Volunteer Allocation** — Role assignment, shift scheduling, auto-calculated hours, and status tracking
- **Task Tracker** — Full task lifecycle with owner, priority, due dates, dependencies, and % completion
- **Resource Planning** — Equipment and supply tracking with shortfall detection and supplier status
- **Pivot Summary** — Aggregated analytics with cross-sheet insights

---

## Tech Stack

| Tool | Usage |
|---|---|
| Microsoft Excel / Google Sheets | Core platform |
| Pivot Tables | Aggregated reporting |
| INDEX / MATCH | Cross-sheet dynamic lookups |
| Conditional Formatting | Visual alerts and progress indicators |
| Data Validation | Dropdown controls for status, category, priority |
| Embedded Charts | Bar, Pie, Doughnut — 9 charts across 6 sheets |
| COUNTIF / IFERROR | Live KPI calculations |

---

## File Structure

```
event-management-dashboard/
│
├── Event_Management_Dashboard_v2.xlsx   # Main dashboard file
├── README.md                            # This file
└── screenshots/                         # Sheet screenshots
    ├── dashboard.png
    ├── budget_tracker.png
    ├── volunteer_allocation.png
    ├── task_tracker.png
    ├── resource_planning.png
    └── pivot_summary.png
```

---

## How to Use

1. **Download** `Event_Management_Dashboard_v2.xlsx` and open in Microsoft Excel (2016+ recommended) or Google Sheets
2. **Start with Budget Tracker** — enter your event categories, allocated amounts, and track spending
3. **Fill Volunteer Allocation** — add volunteer names, assign roles and shifts
4. **Update Task Tracker** — add tasks, assign owners, set due dates, and update status as work progresses
5. **Log Resources** in the Resource Planning sheet with quantities and suppliers
6. **View Dashboard** — all charts and KPI cards update automatically from the data you enter
7. **Check Pivot Summary** for high-level analytics and cross-sheet insights

---

## Key Formulas Used

```excel
# Cross-sheet dynamic lookup (Budget → Dashboard)
=IFERROR(INDEX(BudgetTracker!F$4:F$9, MATCH(A11, BudgetTracker!C$4:C$9, 0)), 0)

# Task completion rate
=IFERROR(COUNTIF(TaskTracker!G:G,"Completed") / COUNTA(TaskTracker!G:G), 0)

# Volunteer hours (auto-calculated)
=IFERROR((TIMEVALUE(H4) - TIMEVALUE(G4)) * 24, 0)

# Budget utilisation
=IFERROR(SUM(F4:F9) / B2, 0)

# Resource shortfall detection
=MAX(D4 - E4, 0)
```

---

## Project Context

Developed as part of college event management operations to streamline planning for large-scale events (500+ attendees). The dashboard reduces manual reporting effort and provides real-time visibility into budget health, volunteer coverage, and task progress.

---

## Contact

Feel free to open an issue or reach out if you'd like to adapt this for your own event!

---

*Built using Microsoft Excel*
