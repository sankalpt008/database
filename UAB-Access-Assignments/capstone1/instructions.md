# Capstone Project 1 – Integrated Database

Automate the end-to-end database build, then explore the generated queries, forms, and reports.

## Build the database

1. Open a Command Prompt in `UAB-Access-Assignments\common`.
2. Run the capstone builder:
   ```cmd
   cscript //nologo build_access.vbs capstone
   ```
3. Wait for Access to create objects, calculate totals, and close automatically.
4. Launch `UAB-Access-Assignments\capstone1\capstone.accdb` in Access.

## What the automation performs

- Creates four core tables and loads sample customers, products, orders, and order details.
- Recomputes `LineTotal` values from pricing data and ensures `Active` flags are set.
- Establishes relationships among all tables with referential integrity.
- Builds reusable parameter and summary queries.
- Generates navigation forms (`MainMenu`, `CustomerForm`, `OrderEntry`, `ReportsMenu`).
- Generates preview-ready reports: `rptSalesByCustomer` and `rptTopProducts`.

## Review checklist

1. From `MainMenu`, open `CustomerForm` and review customer data.
2. Open `OrderEntry`, choose a customer, and inspect the related `OrderDetails_Subform`.
3. Confirm the `Order Total` textbox updates as line totals change.
4. Run the parameter queries (e.g., `q_OrdersByDateRangeParam`, `q_CityFilterParam`) and provide sample values.
5. From `ReportsMenu`, preview both reports to confirm totals and formatting.

## Troubleshooting

- Ensure Access is closed before launching the script to avoid locking the `.accdb` file.
- If `LineTotal` values look incorrect, rerun the builder to recalculate from `UnitPrice` and `Quantity`.
- Review `DEVLOG/development_log.md` or Access's Immediate Window for detailed build logs.
