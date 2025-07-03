

 📋 PDF-generator-tool

Overview:
The Dynamic Export Report Tool is a comprehensive ASP.NET Core MVC application designed to connect to any Microsoft SQL Server database using a provided connection string. It enables users to dynamically select tables and fields, preview data, and export reports in **PDF**, **Excel**, **Word**, and **ZIP** formats — with options for multi-table export, custom ZIP packaging, and email delivery.

---

Key Features:

Dynamic SQL Integration:

  * Accepts a live SQL Server connection string from the user.
  * Fetches tables and their respective columns dynamically.

Table & Field Selection:

  * Supports multi-table selection.
  * Allows users to choose specific columns from each table.
  * Provides “Select All Fields” functionality per table.

Live Data Preview:

  * Real-time preview of selected data (first 5–10 rows per table).
  * Rendered using a responsive HTML table (Kendo-ready structure).

Export Options:

  * Export selected data into:

    * 📄 PDF (using iText 7)
    * 📊 Excel (using ClosedXML)
    * 📝 Word (using OpenXML SDK)
  * Supports combined export (multi-table content in one document).
  * Includes per-table export as ZIP**.

ZIP Export Functionality:

  * Choose export format(s): PDF, Excel, Word, or All.
  * Input optional custom ZIP file name.
  * Downloads compressed ZIP containing exported files per table.

Email Delivery:

  * Send the generated PDF file via Gmail SMTP.
  * Option to also download a local copy.
  * Visual success/failure alerts shown after email action.

Dark Mode Toggle:

  * Full UI toggle between light and dark themes.

Tech Stack:

Frontend: Razor Views (HTML, Bootstrap, JavaScript)
Backend: ASP.NET Core MVC, C#
Database: SQL Server (connected via dynamic string)
Export Libraries:

  * iText 7 (PDF)
  * ClosedXML (Excel)
  * OpenXML SDK (Word)
* **Other Tools:** Kendo UI integration-ready, SMTP for Email



Use Case:

This tool is ideal for database administrators, analysts, or reporting teams who need a quick and flexible interface to extract and package data into clean reports without manual SQL scripting or hardcoded export logic.



