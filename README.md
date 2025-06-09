# ICTA Data Quality Check Tool

<img src="./Icon.ico" alt="ICTA Icon" width="64"/>


## Overview

The ICTA Data Quality Check Tool is an automated system developed by the Statistics unit of the Information Communication Technologies Agency (ICTA) of Azerbaijan. Its primary purpose is to verify and validate the quality of data submitted by Internet Service Providers (ISPs) and Public Switched Telephone Networks (PSTNs). As the telecom regulator of Azerbaijan, ICTA uses this tool to ensure the accuracy and integrity of both Quality of Service (QoS) and economic data provided by operators.

## Purpose

- 🤖 **Automate** quality verification of data submitted by ISPs and PSTN operators
- ✓ **Validate** both QoS and economic data against regulatory formulas
- 📄 **Generate** comprehensive reports highlighting discrepancies and issues
- 🏢 **Support** ICTA's mission of ensuring regulatory compliance in Azerbaijan's telecom sector

## Features

- 📊 Interactive interface for selecting reporting periods (quarters)
- 🔄 Database connectivity for retrieving historical data
- 📈 Comparison between current and previous reporting periods
- ✅ Automated data validation using predefined formulas
- 📋 Report generation for ISP and PSTN data
- 🇦🇿 Azerbaijani language interface

## System Components

### Core Files

- 🐍 **economics.py**: Main Python script that handles economic data validation for both ISPs and PSTNs
- 📊 **ISP.xlsx**: Excel template with formulas for ISP data validation
- 📊 **PSTN.xlsx**: Excel template with formulas for PSTN data validation
- 📊 **QOS DB Model.xlsx**: Database model for Quality of Service metrics
- 🗺️ **Economic reports mapping.xlsx**: Mapping file for economic data reports

### Data Files

- 💾 **data ISP.xlsx**: Working file for ISP data processing
- 💾 **data PSTN.xlsx**: Working file for PSTN data processing
- 📡 **ikta_ookla_data.xlsx**: Additional data source for analysis

### Documentation

- 📝 **Economics.docx**: Documentation for economic data analysis
- 📝 **Qos.docx**: Documentation for Quality of Service analysis

### Visual Resources

- 🖼️ **Background.jpg**: Background image for the application
- 🖼️ **Background_reports.jpg**: Background image for reports
- 🔍 **Icon.ico**: Application icon
- 🔤 **DejaVuSans.ttf** and **DejaVuSans-Bold.ttf**: Font files for the application

## Workflow

<img src="./Background.jpg" alt="Application Background" width="400"/>

1. The user initiates the tool and is prompted to enter a reporting period (quarter/year)
2. The system connects to the database and retrieves both current and previous period data
3. Data is exported to working Excel files for processing
4. The tool applies validation formulas to identify inconsistencies and errors
5. Reports are generated highlighting any data quality issues
6. Results are presented to ICTA statistics specialists for review

<img src="./Background_reports.jpg" alt="Reports Background" width="400"/>

## Technical Requirements

### For End Users (Executable Version)
- 🖥️ Windows operating system
- 🗄️ MySQL/MariaDB database connection
- 📊 Microsoft Excel (for viewing reports)

### For Developers
- 🐍 Python 3.x with required libraries (pymysql, pandas, tkinter, openpyxl)
- 🗄️ MySQL/MariaDB database
- 🖥️ Windows operating system
- 📊 Microsoft Excel (for viewing reports)
- 📦 PyInstaller (for creating executable)

## Usage

### Using the Executable

The tool can be run directly using the executable file:

```
economics.exe
```

The executable includes all necessary dependencies and assets, making it easy to distribute to users without requiring Python installation.

### Development Mode

For development purposes, you can run the main Python script:

```bash
python economics.py
```

Both methods provide a graphical interface for selecting the reporting period. The system then automatically handles the data retrieval, validation, and report generation processes.

### Data Flow Diagram

```
+----------------+     +----------------+     +----------------+
|                |     |                |     |                |
|  ISP/PSTN      |---->|  ICTA Database |---->|  DCS Tool     |
|  Operators     |     |                |     |                |
|                |     |                |     |                |
+----------------+     +----------------+     +-------+--------+
                                                     |
                                                     |
                                                     v
+----------------+     +----------------+     +----------------+
|                |     |                |     |                |
|  Report        |<----|  Validation   |<----|  Data          |
|  Generation    |     |  Engine       |     |  Processing    |
|                |     |                |     |                |
+----------------+     +----------------+     +----------------+
```

---

## Project Structure

```
DCS_Tool/
├── economics.exe         # Executable application file
├── Background.jpg        # Main application background image
├── Background_reports.jpg # Reports background image
├── DejaVuSans.ttf        # Standard font
├── DejaVuSans-Bold.ttf   # Bold font
├── Icon.ico              # Application icon
├── ISP.xlsx              # ISP template
├── PSTN.xlsx             # PSTN template
├── data ISP.xlsx         # ISP data working file
├── data PSTN.xlsx        # PSTN data working file
├── QOS DB Model.xlsx     # Database structure
├── Economic reports mapping.xlsx  # Report mapping
├── ikta_ookla_data.xlsx  # Additional data
└── README.md             # This documentation
```

*This tool was developed by the Statistics unit of the Information Communication Technologies Agency (ICTA) of Azerbaijan.*