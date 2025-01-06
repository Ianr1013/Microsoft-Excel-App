# ExcelApp GUI Application

This project is a Python-based GUI application built with `tkinter` and `openpyxl` that allows users to manage data stored in an Excel file. The app provides a user-friendly interface for inserting and viewing data.

---

## Features

- **Light/Dark Mode:** Toggle between `forest-light` and `forest-dark` themes.
- **Data Management:** 
  - Insert data (Name, Age, Subscription Status, Employment Status) into an Excel file.
  - Automatically update a `Treeview` table with the new entries.
- **Excel Integration:** Read and write data directly from an Excel file (`people.xlsx`).
- **Custom Widgets:** Use styled `ttk` widgets for better appearance and functionality.

---

## Prerequisites

1. **Python:** Make sure Python 3.x is installed on your system.
2. **Dependencies:**
   - `tkinter` (Included with Python standard library).
   - `openpyxl` (Install via pip).

   ```bash
   pip install openpyxl
