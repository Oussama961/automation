# Project Automation & Dashboard Suite

This workspace provides advanced tools for automating project management, reporting, and dashboard generation from Excel files. It covers both calendar/event automation and project Gantt chart visualization, with robust reporting and export features.

---

## Calendare Folder

**Purpose:** Automate event management and reporting in Excel calendar files.

### Key Tools
- `Calendar.py`: Add, update, remove, and batch-manage events in Excel calendars. Generates event summary sheets and supports persistent event parsing.
- `generate_dashboard.py`: Consolidate and analyze calendar/event data from multiple Excel files, generate summary dashboards (Excel, PDF, charts).

### Features
- Command-line automation for event management
- Batch import from CSV/text
- Event summary sheet with hyperlinks
- Logging and verbose/debug mode
- Dashboard generation with pivot tables, charts, and conditional formatting
- PDF export for reports

---

## Plans_tasks Folder

**Purpose:** Automate project plan visualization, reporting, and dashboarding from Gantt chart Excel files.

### Key Tools
- `GrantChartManager.py`: Visualize project plans as interactive Gantt charts (HTML, PNG), with support for task dependencies and robust data validation/cleaning.
- `generate_dashboard.py`: Consolidate and analyze project data from multiple Excel files, generate summary dashboards (Excel, PDF, charts).

### Features
- Gantt chart visualization with dependency arrows
- Data entry and cleaning automation
- Batch reporting and dashboard generation
- Pivot tables, summary charts, and conditional formatting
- PDF export for reports
- Exporting & sharing: automate PDF generation, link Excel to Word for live-updating reports

---

## Advanced & Balanced Automation
- Both folders support batch processing, dashboard generation, and PDF export.
- Both provide robust data cleaning and validation.
- Both are ready for integration with reporting workflows (including Excel-to-Word linking and automated sharing).
- Folder structure and tool capabilities are now balanced for advanced project and event management.

---

## How to Use
- See each folder's main script for command-line usage and options.
- For dashboard/report generation, use the `generate_dashboard.py` in each folder.
- For advanced exporting/sharing, see the Exporting & Sharing section in the main README.

---

## Requirements
- Python 3.8+
- pandas, openpyxl, plotly, matplotlib
- (Optional for PDF export) win32com.client (Windows)

Install dependencies:
```sh
pip install pandas openpyxl plotly matplotlib
```

---

## Command Line Usage

Below are example command lines for the main automation and dashboard scripts in this workspace.

### Calendar/Event Automation (`Calendare` folder)

**Add an event to a calendar:**
```sh
python Calendare/Calendar.py --load path/to/calendar.xlsx --add-event 2025-07-01 "Project Kickoff"
```

**Batch import events from CSV or text:**
```sh
python Calendare/Calendar.py --load path/to/calendar.xlsx --batch-import events.csv
```

**Generate a dashboard from all calendar files in a folder:**
```sh
python Calendare/generate_dashboard.py --folder Calendare/
```

---

### Project Gantt Chart Automation (`Plans_tasks` folder)

**Visualize a Gantt chart from a project Excel file:**
```sh
python Plans_tasks/GrantChartManager.py --load path/to/Project-Gantt-Chart.xlsx --gantt --output Plans_tasks/output/
```

**Generate a dashboard from all project files in a folder:**
```sh
python Plans_tasks/generate_dashboard.py --folder Plans_tasks/
```

---

### General Notes

- Use `--help` with any script to see all available options.
- Output files (Excel, PDF, charts) are saved in the specified output folder or alongside the master dashboard file.
- For custom sheet names, pivot columns, or value columns, use the corresponding script arguments (see script help).

*This suite is designed for teams who want to automate and visualize project timelines, events, and KPIs directly from Excel files, with advanced reporting and export options.*
