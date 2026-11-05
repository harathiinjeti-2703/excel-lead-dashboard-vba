# Excel Lead Dashboard with VBA

An Excel dashboard for reviewing lead activity, tracking KPIs, and exploring lead data with PivotTables. A VBA date picker lets users filter the dashboard by a start and end date.

## Features

- KPI summary for total leads, interested leads, demos completed, joined students, pending follow-ups, conversion rate, and overdue follow-ups
- PivotTables for summarizing lead activity
- Start and end date selection on the Summary sheet
- VBA macros that apply the selected date range to the KPIs and PivotTables
- Lead data stored in an Excel Table so it can expand as records are added

## Files

- `*.xlsm` — macro-enabled Excel workbook
- `README.md` — project description and usage notes

## How to use

1. Download the `.xlsm` workbook.
2. Open it in desktop Microsoft Excel.
3. Enable macros if Excel prompts you.
4. Add or edit records inside the `Leads_Data` table.
5. On the `Summary` sheet, select the start and end dates with the calendar buttons.
6. Review the updated KPI values and PivotTables.

The date cells display dates as `dd-mm-yyyy`.

## Requirements

- Desktop Microsoft Excel
- Macros enabled to use the VBA calendar and automatic filtering

## Privacy

The workbook should contain sample or anonymized data only. Do not publish real names, phone numbers, email addresses, or other private lead information.

## Built with

- Microsoft Excel
- Excel Tables and PivotTables
- VBA
