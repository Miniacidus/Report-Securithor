# Securithor Report Automator

### Description
A data automation tool developed in Python to optimize reporting workflows for security monitoring centers using **Securithor** software. This application processes raw CSV signal logs and generates professional Excel reports for attendance and technical failure tracking.

### Problem Solved
Manual report generation used to take several hours of administrative work and was highly prone to human error. This tool reduces processing time to **seconds**, ensuring 100% accuracy in client mapping and data filtering.

### Key Features
- **Automatic Name Mapping:** Links account numbers to client names using an external dictionary.
- **Termination Filtering:** Automatically excludes inactive or cancelled accounts from the final report.
- **Failure Detection Logic:** Identifies clients without active test signals (Signal 88) and separates results into color-coded spreadsheets.
- **Print-Ready Formatting:** Generates Excel files pre-configured for printing (50 rows per page, landscape orientation).

### Tech Stack
- **Language:** Python 3.x
- **Libraries:** Pandas (Data Processing), Tkinter (GUI), OpenPyxl (Excel Formatting).
- **Deployment:** PyInstaller (Standalone Executable).

### How to use
1. Run the `RepotesAlarmas.exe`.
2. Configure your client list using the **"Nombres"** button.
3. Load your Securithor CSV signals.
4. Click **"GENERAR REPORTES"** to get your Excel files.
