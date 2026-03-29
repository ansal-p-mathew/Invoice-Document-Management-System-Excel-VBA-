![License: MIT](https://img.shields.io)
# Invoice Document Management System (Excel VBA)

## Overview

This project is an Excel-based Invoice Document Management System developed using VBA to automate the organization, storage, and tracking of invoice-related documents and payment receipts.

The solution eliminates manual file handling by dynamically creating structured folders, enabling document uploads through a user-friendly interface, and maintaining real-time file counts within Excel.

---

## Key Features

* **Automated Folder Creation**
  Dynamically creates invoice-specific folders and subfolders based on user input.

* **Document & Receipt Upload System**
  Upload multiple files using button-driven controls:

  * Invoice Documents
  * Payment Receipts

* **Real-Time File Tracking**
  Automatically updates document and receipt counts in Excel.

* **Cross-Platform Compatibility**
  Designed to work on both Windows and macOS environments.

* **User-Friendly Interface**
  Replaces unreliable event-based triggers with a stable button-based workflow.

* **Protected System Fields**
  Prevents manual editing of document and receipt counts to maintain data integrity.

---

## How It Works

1. Enter:

   * Customer PO
   * Customer Invoice Number

2. The system automatically:

   * Creates an invoice folder
   * Creates a "Payment Receipts" subfolder

3. Select any cell in the desired row.

4. Use the buttons:

   * **Upload Docs** → Upload invoice-related documents
   * **Upload Receipts** → Upload payment receipts

5. The system updates:

   * Number of documents
   * Number of receipts

---

## File Structure

Base Folder (defined in Excel cell `K2`):

Invoices/
├── INV001/
│   ├── document1.pdf
│   ├── document2.xlsx
│   └── Payment Receipts/
│       ├── receipt1.pdf
│
├── INV002/
│   ├── ...
│   └── Payment Receipts/

---

## Setup Instructions

1. Download or clone this repository.
2. Open the `.xlsm` file in Microsoft Excel.
3. Enable macros when prompted.
4. Set the **Base Folder Path** in cell `K2`.

### Example Paths

**Windows:**
C:\Users\YourName\Documents\Invoices

**Mac:**
/Users/yourname/Documents/Invoices

---

## Technologies Used

* Microsoft Excel
* VBA (Visual Basic for Applications)

---

## Business Value

* Reduces manual file handling and organization effort
* Ensures structured storage of invoice-related documents
* Improves traceability and document tracking
* Enhances workflow efficiency for finance and administrative processes

---

## Technical Highlights

* VBA-based automation for file handling and workflow control
* Dynamic folder creation using file system operations
* Event-driven validation to prevent manual data manipulation
* Cross-platform path handling using `Application.PathSeparator`

---

## Limitations

* Requires macros to be enabled
* File path must be configured per user environment
* Basic duplicate file handling (can be enhanced)

---

## Future Improvements

* Prevent duplicate file uploads
* Add “Open Folder” quick access buttons
* Dashboard for tracking invoices and documents
* Integration with Power BI for reporting and analytics

---

## Author

Developed as part of a practical automation project to streamline invoice document workflows using Excel VBA.
