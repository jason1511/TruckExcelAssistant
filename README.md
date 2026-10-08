# Truck Excel Assistant

A Windows desktop application that reduces repetitive data entry across truck bookkeeping and reusable customer-like invoice layouts.

## Current phase

The current phase provides a responsive native WinForms shell, a universal haul-entry screen, and portable local storage. It includes:

- enter each haul once for truck bookkeeping and customer invoicing;
- one universal entry form with Pembukuan, Mirip Miguno, Mirip Agrico, and Mirip Sumber Pangan modes;
- separate Dari and Ke fields without route management;
- Agrico-style claim quantity, claim rate, and calculated claim total;
- editable customer dropdown ready to learn from saved customer names;
- a SQLite database stored beside the executable;
- working saved and draft records;
- a searchable Data Angkutan screen with saved, draft, and trash filters;
- edit and resume-draft workflows that update the original record;
- recoverable deletion with restore from Sampah;
- manual entry for journey, weight, rate, and expense information;
- live gross, adjustment, and final calculations;
- licence-plate normalization;
- validation for required invoice fields; and
- an Excel-row preview; and
- working Excel generation for the current Mirip Miguno, Mirip Sumber Pangan, and Pembukuan layouts.

The database is created automatically as `truck_excel_assistant.db`. The next phase is exact template-based Excel generation, including Mirip Agrico and its separate claim sheet.

## Technology

- C#
- Windows Forms
- .NET 10 LTS
- Microsoft.Data.Sqlite

## Run locally

Open `TruckExcelAssistant.slnx` in Visual Studio with the **.NET desktop development** workload, then press `F5`.
