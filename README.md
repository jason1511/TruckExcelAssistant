# Truck Excel Assistant

A Windows desktop application that reduces repetitive data entry across truck bookkeeping and customer invoices for Miguno and Agrico.

## Current phase

The first phase establishes a responsive native WinForms shell and a universal haul-entry screen. It includes:

- enter each haul once for both truck bookkeeping and customer invoicing;
- customer-specific invoice fields for Miguno and Agrico;
- manual entry for journey, weight, rate, and expense information;
- live gross, adjustment, and final calculations;
- licence-plate normalization;
- validation for required invoice fields; and
- an Excel-row preview.

Database persistence and Excel generation are intentionally reserved for the next phase.

## Technology

- C#
- Windows Forms
- .NET 10 LTS

## Run locally

Open `TruckExcelAssistant.slnx` in Visual Studio with the **.NET desktop development** workload, then press `F5`.
