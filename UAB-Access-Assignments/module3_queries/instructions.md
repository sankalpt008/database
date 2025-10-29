# Module 3 – Maintaining & Querying a Database

Follow these steps to generate and review the Module 3 assignment database.

## Build the database

1. Open a Command Prompt in `UAB-Access-Assignments\common`.
2. Run the builder:
   ```cmd
   cscript //nologo build_access.vbs module3
   ```
3. Wait for Microsoft Access to open, import the data, and close automatically.
4. Open `UAB-Access-Assignments\module3_queries\module3.accdb` in Access.

## What the automation performs

- Creates the `Customers` table when missing.
- Imports `data/customers.csv` so the table is populated with sample rows.
- Saves the instructional queries listed below for quick review.

## Review checklist

1. Locate the **Queries** section of the Navigation Pane.
2. Run each saved query and observe the expected output:
   - `q_HoustonCustomers` (filters Houston customers).
   - `q_LogicalOperators_DallasOrHouston` (Dallas or Houston customers).
   - `q_NotHouston` (excludes Houston customers).
   - `q_Calculated_FullName` (shows the computed full name field).
   - `q_SortedByLastThenFirst` (sorted alphabetically by last/first name).
   - `q_ByCityParam` (prompts for a city and filters accordingly).
3. Try entering `Houston` when prompted by `q_ByCityParam` and verify the result set.

## Troubleshooting

- Confirm Access was closed before launching the script to avoid file lock issues.
- Review the Immediate Window (`Ctrl+G`) or `DEVLOG/development_log.md` for step-by-step logs.
