# Chapter 4 – Creating & Customizing Forms

Use the automation to build the Chapter 4 practice database and review the generated forms.

## Build the database

1. Open a Command Prompt in `UAB-Access-Assignments\common`.
2. Run the builder:
   ```cmd
   cscript //nologo build_access.vbs chapter4
   ```
3. Allow Access to create objects and close on its own.
4. Open `UAB-Access-Assignments\chapter4_forms\chapter4.accdb` in Access.

## What the automation performs

- Ensures `Customers` and `Orders` tables exist.
- Imports sample CSV data for both tables and establishes a relationship.
- Generates the `CustomerEntry` form with:
  - Bound text boxes for FirstName, LastName, City, and Email.
  - An unbound city filter combo box tied to `FilterCustomersByCity`.
  - Navigation and data-entry command buttons.
  - An `Orders_Subform` showing related order history.
  - A header label describing usage.

## Review checklist

1. Open the `CustomerEntry` form.
2. Use the navigation buttons to browse through the records.
3. Pick a city in the **Filter by City** combo box and verify the form filters accordingly.
4. Observe the `Orders_Subform` updates as you select different customers.
5. Add a new customer record and choose **Save** to commit the changes.

## Troubleshooting

- The city filter uses the shared `FilterCustomersByCity` helper; verify the combo box is populated.
- Access prompts for saving design changes when closing forms; choose **Yes** to keep adjustments.
- Consult `DEVLOG/development_log.md` or the Immediate Window for runtime logs.
