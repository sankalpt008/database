# UAB Access Assignments Automation

This repository delivers three Microsoft Access automation assignments built for Access 2019/365. Each assignment can be generated from source artifacts (SQL, CSV, VBA) using the included VBScript automation without any manual object creation.

## Project Structure

- `common/` – shared automation scripts and Access modules.
- `module3_queries/` – Maintaining & Querying assignment assets.
- `chapter4_forms/` – Creating & Customizing Forms assignment assets.
- `capstone1/` – Capstone Project 1 integrating chapters 1–4.
- `DEVLOG/` – timestamped change history for all generated assets.

## Prerequisites

- Windows environment with Microsoft Access 2019 or Microsoft 365 installed.
- Permissions to execute Windows Script Host (`cscript.exe`).

## Build Instructions

1. Open a Command Prompt in the `UAB-Access-Assignments\common` folder.
2. Run the VBScript with the assignment name (`module3`, `chapter4`, or `capstone`). For example:
   ```cmd
   cscript //nologo build_access.vbs module3
   ```
3. When the script completes, open the generated `.accdb` file in Microsoft Access and review the saved queries, forms, and reports as outlined below.

## Verification Checklists

### Module 3 – Maintaining & Querying a Database
- Run `cscript //nologo build_access.vbs module3`.
- Open `module3_queries/module3.accdb`.
- Confirm the `Customers` table exists with imported data.
- Locate and run the saved queries:
  - `q_HoustonCustomers`
  - `q_LogicalOperators_DallasOrHouston`
  - `q_NotHouston`
  - `q_Calculated_FullName`
  - `q_SortedByLastThenFirst`
  - `q_ByCityParam` (verify the parameter prompt).

### Chapter 4 – Creating & Customizing Forms
- Run `cscript //nologo build_access.vbs chapter4`.
- Open `chapter4_forms/chapter4.accdb`.
- Open the `CustomerEntry` form.
- Use the navigation buttons to browse records and the combo box to filter by city.
- Add or edit customer records and confirm the `Orders_Subform` reflects related order data.

### Capstone Project 1 – Integrated Database
- Run `cscript //nologo build_access.vbs capstone`.
- Open `capstone1/capstone.accdb`.
- From `MainMenu`, open `CustomerForm`, `OrderEntry`, and `ReportsMenu`.
- Execute parameter queries such as `q_OrdersByDateRangeParam` and `q_CityFilterParam`.
- Generate the `SalesByCustomer` and `TopProducts` reports.

## Troubleshooting

- Ensure Access is not already running with the target database open before executing the script.
- Review `DEVLOG/development_log.md` for a chronological list of build steps and logged automation messages.
- Use the Immediate Window (`Ctrl+G` in Access) to review `Debug.Print` output when running modules manually.

