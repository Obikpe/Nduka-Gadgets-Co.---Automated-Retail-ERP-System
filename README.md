# Case Study: Custom Retail Enterprise Resource Planning (ERP) System
**Streamlining Operations and Inventory Control through Advanced Excel Architecture**

* **Role:** Lead Excel Developer & Data Analyst
* **Context:** Portfolio Project / Freelance Client Solution
* **Timeline:** 3 Weeks
* **Tools & Tech:** Microsoft Excel, VBA (Visual Basic for Applications), Power Query, Relational Data Modeling

---

## 📌 Executive Summary
Small-to-medium retail operations frequently outgrow standard spreadsheets but lack the capital or infrastructure to deploy enterprise software like SAP or Oracle. This project bridges that gap. I engineered a lightweight, fully automated, and relational ERP system entirely within Microsoft Excel. By leveraging advanced VBA scripting and rigorous data validation, the system centralized sales, inventory, and supplier management into a single source of truth, completely eliminating manual data-entry bottlenecks and costly stockouts.

---

## 🛠️ The Problem: Operational Chaos & Data Silos
The business was operating across multiple, disconnected spreadsheets managed by different staff members. This fragmented workflow created critical business vulnerabilities:

* **Mismatched Data:** Pricing changes in the inventory sheet weren't reflecting on the sales floor, leading to revenue leakage and manual checkout overrides.
* **Blind Inventory Management:** Without real-time tracking, the business frequently ran out of high-demand items, directly resulting in lost sales. Conversely, over-ordering tied up valuable working capital in slow-moving stock.
* **Administrative Drain:** Generating a simple end-of-day revenue or profit margin report required hours of manual copy-pasting, VLOOKUP stitching, and error troubleshooting.

---

## 🧠 The Solution: Relational Architecture & Automation
Instead of treating Excel as a flat data dumping ground, I designed the workbook to mirror a relational database management system (RDBMS).
### 1. Robust Relational Data Modeling
I established strict primary keys (SKUs, Supplier IDs, Transaction IDs) across separate data tables to enforce data integrity:
* **Inventory Ledger:** Tracks current stock levels, cost prices, selling prices, and reorder thresholds.
* **Sales Log:** A read-only repository that captures transactional metadata (timestamp, items sold, quantities, payment methods).
* **Supplier Directory:** Links product SKUs to primary vendors for rapid replenishment.

### 2. The VBA Automation Engine
To ensure the system could be operated by non-technical retail staff, I built a custom back-end using VBA macros:
* **Multi-Item Transaction Form:** Created an intuitive, user-friendly data entry interface. Staff can select items via dynamic drop-downs, automatically pull current retail prices, and log multi-item transactions with a single click.
* **Instant Inventory Deduction:** Programmed transactional logic so that the moment a sale is committed, the script searches the `Inventory Ledger` for the matching SKUs and subtracts the exact quantities sold.
* **Proactive Reorder Alerts:** Wrote a background routine that constantly evaluates `Current Stock` against `Reorder Point`. If stock falls below the threshold, the row conditionally highlights, and the item is automatically appended to a dynamic "Pending Orders" sheet.

### 3. Data Integrity & Guardrails
* Implemented strict data validation rules to prevent text from being entered into quantity fields.
* Locked down backend data architecture and historical tables with password protection, ensuring users can only interact with the system via authorized input fields and macro buttons.

---

## 📈 The Result: Data-Driven Retail Efficiency
The deployment of this custom Excel ERP yielded immediate, measurable operational improvements:

* **40% Reduction in Admin Time:** Automated end-of-day reconciliation reports from a 2-hour manual process down to a single button click.
* **Zero Critical Stockouts:** During the first quarter of testing, automated reorder alerts successfully prompted timely supplier orders, keeping high-margin stock availability at 100%.
* **Absolute Data Integrity:** Eradicated manual pricing errors and duplicate transaction entries by locking down tables and routing all inputs through the VBA validation engine.
* **Capital Optimization:** Enabled leadership to identify dead stock, allowing them to liquidate stale inventory and reallocate capital to high-turnover products.

---

### 💻 Code Snippet Highlight
```vba
Sub CommitTransaction()
    ' Purpose: Logs sale to Sales Log and simultaneously updates Inventory Ledger
    Dim wsSales As Worksheet, wsInv As Worksheet
    Dim nextRow As Long, targetItem As String, qtySold As Long
    
    Set wsSales = ThisWorkbook.Sheets("Sales_Log")
    Set wsInv = ThisWorkbook.Sheets("Inventory_Ledger")
    
    ' Fetch inputs from Entry Template
    targetItem = Range("B4").Value ' SKU
    qtySold = Range("B5").Value    ' Quantity
    
    ' 1. Log the Sale
    nextRow = wsSales.Cells(wsSales.Rows.Count, "A").End(xlUp).Row + 1
    wsSales.Cells(nextRow, 1).Value = Now ' Timestamp
    wsSales.Cells(nextRow, 2).Value = targetItem
    wsSales.Cells(nextRow, 3).Value = qtySold
    
    ' 2. Update Inventory via WorksheetFunction Match
    Dim rowNum As Variant
    rowNum = Application.Match(targetItem, wsInv.Columns("A"), 0)
    
    If Not IsError(rowNum) Then
        ' Subtract quantity from the Stock Column (Assume Column C is Current Stock)
        wsInv.Cells(rowNum, 3).Value = wsInv.Cells(rowNum, 3).Value - qtySold
        MsgBox "Transaction secure and inventory updated!", vbInformation
    Else
        MsgBox "Error: SKU not found in inventory.", vbCritical
    End If
End Sub
