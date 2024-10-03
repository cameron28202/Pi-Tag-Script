# Tag Mapping Script

This PowerShell script reads from an Excel file to create attributes on the template level in OSIsoft PI AF (Asset Framework), then maps a tag and label to each attribute.

## Prerequisites

- PowerShell
- OSIsoft PI AF SDK
- Microsoft Excel

## Usage

1. Ensure that the Excel file containing the tag information is closed before running the script.

2. Run the script from PowerShell, providing the path to your Excel file as an argument:

   ```powershell
   .\script.ps1 -excelPath "C:\Path\To\Your\Excel\File.xlsx"
   ```

   Replace `"C:\Path\To\Your\Excel\File.xlsx"` with the actual path to your Excel file.

## Important Note

**Always close the Excel file before running the script.** Keeping the file open may cause issues with file access and data reading.
