# Excel Worksheet Unprotector

An application to remove worksheet password protection from Microsoft Excel files (`.xlsx`). This tool does **not** recover the password; it directly removes the protection flag from the file's underlying XML structure.

**Publisher:** Noshadi-sec

![Python](https://img.shields.io/badge/python-3.8+-blue.svg)
![License](https://img.shields.io/badge/License-MIT-yellow.svg)

---

## Getting Started

There are two ways to use this application. The easiest method is to download the pre-built application.

### Method 1: Download the Application (Easiest)

You can download the ready-to-use `.exe` application for Windows from the official Releases page. No installation is needed.

**[Download the latest release here](https://github.com/Noshadi-sec/Excel_Worksheet_Unlocker/releases)**

### Method 2: Run from Source Code (For Developers)

If you want to run the Python script directly, you can download the source code.

1.  **Prerequisites**: Make sure you have **Python 3** installed on your system.
2.  **Download the Code**: You can either clone the repository with Git or download it as a ZIP file.
    * **Using Git (recommended):**
        ```bash
        git clone https://github.com/Noshadi-sec/Excel_Worksheet_Unlocker.git
        ```
    * **As a ZIP file:** On the main repository page, click the green `< > Code` button and select "Download ZIP".
3.  **Run the App**: Open your terminal, navigate into the project directory, and run the script:
    ```bash
    python main.py
    ```

---

## How It Works

An `.xlsx` file is essentially a ZIP archive containing various XML files. Worksheet protection is specified by a `<sheetProtection ... />` tag within the XML for that sheet. This tool automates the process of unzipping the file, finding and removing that tag, and re-zipping the contents into a new, unprotected file.

---

## Features

* **Simple GUI**: An intuitive interface built with Tkinter.
* **No Password Needed**: Removes protection without needing the original password.
* **Safe**: Creates a new, unprotected copy of your file, leaving the original untouched.
* **Cross-Platform**: The Python script works on Windows, macOS, and Linux.

---

## Building executable file from Source (Optional)

If you want to create your own standalone executable from the source code, you can use PyInstaller.

1.  **Install PyInstaller**:
    ```bash
    pip install pyinstaller
    ```
2.  **Create the Executable**: Navigate to the script's directory and run the command for your OS.
    * **For Windows (.exe)**
        ```bash
        pyinstaller --onefile --windowed --name ExcelUnprotector main.py
        ```
    * **For macOS (.app)**
        ```bash
        pyinstaller --onefile --windowed --name ExcelUnprotector --icon=your_icon.icns main.py
        ```
3.  **Find Your App**: The new application will be inside the `dist` folder.

---

## Acknowledgements

This project was inspired by the work done on the [Excel-Worksheet-Password-Remover](https://github.com/aziascreations/Excel-Worksheet-Password-Remover) repository.

---

## Disclaimer

This tool is intended for educational purposes and for recovering access to your own files. Use it responsibly and only on files you have the right to modify.
