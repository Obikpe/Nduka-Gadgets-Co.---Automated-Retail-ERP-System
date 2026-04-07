# Nduka Gadgets & Co. | Automated Retail ERP & Analytics System
## Executive Summary
In fast-paced retail environments, manual sales tracking often leads to data fragmentation, calculation errors, and delayed business insights. This project solves these challenges by transforming a standard spreadsheet into a robust, "app-like" ERP (Enterprise Resource Planning) tool. By centralizing sales entry through an automated interface and piping that data into a relational database, the system ensures 100% data accuracy and provides real-time visibility into business performance for stakeholder decision-making.
## Key Technical Features 🛠️
* Relational Data Mapping: Utilizes a custom VBA-driven looping script to transfer multi-line sales from a front-end "Sales Order Form" into a centralized, structured database.
* Automated Data Integrity: Features dynamic sheet protection toggles within the VBA backend. This prevents accidental deletion of complex formulas (like XLOOKUPs for product pricing) while allowing authorized macros to update records seamlessly.
* Dynamic Business Intelligence: An interactive Executive Dashboard provides real-time KPI updates. The interface uses advanced Pivot Table logic and Slicers to allow users to filter performance by category, date, or payment method.
* Synthetic Data Engineering: Built using a custom dataset of 100+ electronic products across multiple categories (MacBooks, iPhones, Laptops, etc.) to simulate a high-volume retail environment.
## Technical Stack 💻
* Platform: Microsoft Excel
* Language: VBA (Visual Basic for Applications)
* Analysis Tools: Power Pivot, Slicers, Data Validation
* Design: UI/UX optimized for 16:9 widescreen displays
## How to Use 📖
* Enable Macros: Upon opening the .xlsm file, click "Enable Content" to allow the VBA automation to run.
* Fill the Entry Form: Navigate to the Sales Records tab. Enter customer details and select products from the dropdown menus.
* Submit Order: Click the "Submit" button. The system will validate the data, generate a unique Sales ID, move the record to the database, and reset the form.
* Analyze Performance: Navigate to the Overview tab. Use the Slicers to filter the data and view updated revenue trends and best-selling products.
## Project Structure 📂
* Nduka_Gadgets_ERP.xlsm: The core macro-enabled application.
* vba_source_code.txt: Raw VBA scripts for SubmitMultiSale and CancelSale for easy code review.
* Dashboard_Preview.png: Screenshot of the interactive data visualization interface.
## Contact & Portfolio
Designed and Developed by a Professional Data Analyst focused on delivering actionable business solutions. Check out https://peter-portfolio11.vercel.app for more projects involving Python, SQL, and Power BI.
Designed and Developed by a Professional Data Analyst focused on delivering actionable business solutions. Check out my [React Portfolio Website] for more projects involving Python, SQL, and Power BI.
