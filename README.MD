# Report Automation Application

- Automated Excel Reports is a VBA application developed to streamline the process of creating and updating reports from raw data. 
- The application consists of different modules, each serving a specific purpose, to import, clean, analyze data, create dynamic charts, generate formatted reports, and add user interactivity.

# Modules:

## ImportModule

- Purpose: Which allows users to import raw data from external files, such as CSV files or other Excel sheets, into the active Excel workbook.
- How to Use: Run the procedure to open the file selection dialog and choose the data file to import. The data will be copied to the active sheet.
## FormatCleanerModule

- Purpose: Which is responsible for cleaning the imported data by removing duplicates and applying necessary formatting.
- How to Use: Execute the procedure to clean and format the imported data.
## AnalysisModule

- Purpose: The module houses various procedures to perform specific calculations and analyses on the imported data. Functions like summing values, calculating averages, and finding min/max values are implemented here.
- How to Use: Call the relevant procedures (e.g., "CalculateSum," "CalculateAverage") to perform specific data analyses.
## GraphicsModule

- Purpose: Which automatically generates dynamic charts that update as the raw data is updated. This facilitates better data visualization.
- How to Use: Execute the "CreateDynamicCharts" procedure to create and insert dynamic charts into the report.
## ReportModule

- Purpose: Responsible for automatically generating a formatted report with the results of data analyses and the charts produced.
- How to Use: Run the "GenerateReport" procedure to generate the report with the relevant data and charts.

## MainModule
- The Automated Excel Reports application aims to automate the process of creating and updating reports in Excel. 
- Users can import raw data, clean and format it, perform data analysis, generate dynamic charts, and automatically create formatted reports. 
- The application enhances data visualization, simplifies reporting tasks, and saves valuable time for users working with repetitive reporting tasks.

# Usage

- Open the Excel file containing the VBA code.
- Enable macros in Excel to allow VBA code execution.
- Press "Alt + F11" to access the VBA editor.
- Review and execute the modules in the desired sequence based on your reporting needs.
- Use the control buttons on the worksheet to perform specific tasks.