# Truck Excel Assistant

A Windows desktop application that reduces repetitive data entry across truck bookkeeping and two reusable customer invoice layouts.

## Current phase

The current phase provides a responsive native WinForms shell, a universal haul-entry screen, and portable local storage. It includes:

- enter each haul once for truck bookkeeping and customer invoicing;
- selectable bookkeeping, compact invoice, and complete invoice layouts;
- editable customer dropdown ready to learn from saved customer names;
- a SQLite database stored beside the executable;
- working saved and draft records;
- a searchable Data Angkutan screen;
- manual entry for journey, weight, rate, and expense information;
- live gross, adjustment, and final calculations;
- licence-plate normalization;
- validation for required invoice fields; and
- an Excel-row preview.

The database is created automatically as `truck_excel_assistant.db`. Excel generation is the next implementation phase.

## Technology

- C#
- Windows Forms
- .NET 10 LTS
- Microsoft.Data.Sqlite

## Run locally

Open `TruckExcelAssistant.slnx` in Visual Studio with the **.NET desktop development** workload, then press `F5`.
