# Excel Worksheet Unprotector

An application to remove worksheet password protection from Microsoft Excel files (`.xlsx`). This tool does **not** recover the password; it directly removes the protection flag from the file's underlying XML structure.

---

## How It Works 

`.xlsx` files are essentially ZIP archives containing various XML files and other assets that define the workbook. Worksheet protection is specified by a `<sheetProtection ... />` tag within the XML file for each specific sheet (e.g., `xl/worksheets/sheet1.xml`).

This tool automates the following process:
1.  **Unzips** the input `.xlsx` file in memory.
2.  Iterates through each file in the archive.
3.  For each worksheet XML file, it uses a regular expression to find and **remove the `<sheetProtection.../>` tag**.
4.  **Re-zips** the modified contents into a new `.xlsx` file, leaving the original file untouched.

---

## Features

* **Simple GUI:** An intuitive interface built with Tkinter.
* **No Password Needed:** Removes protection without needing the original password.
* **Safe:** Creates a new, unprotected copy of your file. Your original file is never modified.
* **Cross-Platform:** The Python script works on Windows, macOS, and Linux. Standalone executables can be created for Windows and macOS.

---

## Installation and Usage

You can either run the application directly from the Python source code or package it as a standalone executable.

### Option 1: Running from Source Code

You don't need any external libraries to run this script, as it only uses Python's standard libraries.

1.  **Prerequisites:** Make sure you have **Python 3** installed on your system.
2.  **Save the Code:** Save the provided code as a Python file (e.g., `main.py`).
3.  **Run the App:** Open your terminal or command prompt, navigate to the directory where you saved the file, and run:
    ```bash
    python main.py
    ```
4.  **you must see the GUI**

5.  **Follow GUI steps:**
    * Click "**1. Select Excel File (.xlsx)**" to choose the protected file.
    * Click "**2. Remove Protection and Save**" to choose a location to save the new, unprotected file.

### Option 2: Creating a Standalone Application

This process bundles the script into a single executable file that can run on other computers without needing Python installed.

1.  **Install PyInstaller:** First, install the packaging tool from your terminal:
    ```bash
    pip install pyinstaller
    ```
2.  **Create the Executable:** Navigate to your script's directory in the terminal and run the command for your operating system.

    * **For Windows (.exe)**
        ```bash
        pyinstaller --onefile --windowed --name ExcelUnprotector main.py
        ```

    * **For macOS (.app)**
        This command creates the `.app` bundle. You can optionally add an icon file (`.icns` format) for a more professional look.
        ```bash
        pyinstaller --onefile --windowed --name ExcelUnprotector --icon=your_icon.icns main.py
        ```

3.  **Find Your App:** A `dist` folder will be created. Your standalone application (`ExcelUnprotector.exe` or `ExcelUnprotector.app`) will be inside it.

---

## Disclaimer

This tool is intended for educational purposes and for recovering your own files where you have forgotten the password. Use it responsibly and only on files you have the right to modify.