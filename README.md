# SAP ME5A Purchase Requisition Tracker (VBA + PowerShell)

Excel VBA and PowerShell automation that pulls open purchase requisitions (PRs) from SAP transaction **ME5A** through **SAP GUI Scripting**, enriches them with material **lead time** and **days until delivery**, and prepares a ready-to-send **Outlook** email with the results. It can be triggered on demand from a worksheet button or run weekly by **Windows Task Scheduler**.

## Problem

The entire process was manual. It started with accessing transaction **ME5A** in **SAP** and checking the purchase requisitions (PRs) of each requestioner one by one. For every PR, we had to calculate how many days were left until the delivery date and compare that with the material's lead time. If the days remaining were fewer than the lead time, the PR was already late: it should have been converted in a purchase order earlier. We then confirmed the information and follow up with the procurement department.

This was repettive and time-consuming, and the raw **SAP** list gave no clear view of which PRs were the most urgent.

## How It Works

```bash
1. **PowerShell script** opens `ME5A_REPORT.xlsm` in the background (Excel hidden) through COM.
2. **`SapUpdate` module** attaches to the open SAP GUI session, runs ME5A for each requester in the configuration list, copies the ALV Grid output into the report sheet, and formats it (filters, borders, alignment, sorting).
3. **Two calculated columns** are added to the report: **days until delivery** (delivery date from SAP minus today) and the material's **lead time**, fetched with `VLOOKUP` on the material code. Comparing them shows how urgent each PR is. For example, a PR with 28 days until delivery and a 30-day lead time is **2 days late**, meaning it should already have become a purchase order.
4. **`Email` module** builds a corporate-style email in Outlook: recipients in *To* and *CC*, a short message, and the report table pasted in the body. The email window opens ready for review; **it is not sent automatically**.
5. A confirmation message tells the user that the weekly email is ready to send.
```

The same `SAP_ME5A_UPDATE` macro can also be run manually from a button in the workbook, without the scheduler.

## Solution 

The workbook automates the whole routine, from extracting the data in **SAP** to preparing the report email, and adds the place that was missing: a clear view of urgency.

After pasting the ME5A data, two columns are added for every PR:

- **Days until delivery:** the delivery date from SAP minus today
- **Lead time:** the material's lead time, fetched with `VLOOKUP` on the material code

Comparing the two show immediately which PRs are late. For example, a PR with **28 day** until delivery and a material lead time of **30 days** means that is **2 days late**: it should have been converted into a purchase order.

## Results 

- Reduced the process from **up to 20 minutes to less than 1 minute** when executed on demand.
- Automated the weekly monitoring of **50+ PRs across 3 requestioners** with no manual interaction during execution. 
- Improved visibility of critical PRs by automatically comparing **delivery dates against material's lead time**. 
- Automatically filters and prioritizes PRs with fewer days until delivery date, making urgent cases easier to check.
- Eliminated manual calculations and the need to cross-check lead-time information across different sources.
- Standardized the weekly report and follow-up process with Procurement.

## Technologies Used

- VBA - Data processing, calculations, report formatting, filtering, and Outlook email generation.
- SAP GUI Scripting - Automated interaction with SAP ME5A transaction and extraction of purchase requesitions data.
- PowerShell - Automated workbook execution through COM and Windows Task Scheduler integration.
- Microsoft Excel - Report generation, data processing, days until delivery, and material lead-time `VLOOKUP`.
- Microsoft Outlook - Automated preparation of the weekly follow up email.
- Windows Task Scheduler - Weekly unattended execution of the automation.

## Author

**Nicolas Bittencourt** 

[GitHub](https://github.com/nicolasbitt)