# Monthly Excel File Generator

This PowerShell script generates monthly Excel files for a given year, with each file containing a worksheet for each day of the month. Each worksheet includes headers provided by the user and populates cells with serial numbers and dates.

## Requirements

- PowerShell
- ImportExcel module

## How to Use

1. Clone or download this repository to your local machine.
2. Open PowerShell.
3. Navigate to the directory where the script is located.
4. Run the script by executing the following command:
   ```powershell
   .\Create-Records.ps1
   ```
5. Follow the prompts:
 - Enter the year for which you want to generate Excel files.
 - Provide headers for the Excel file (separated by commas).

The script will then generate Excel files for each month of the specified year in the current directory.

## Notes
Leap years are considered when determining the number of days in February.
Each Excel file will have a worksheet for each day of the month, with headers provided by the user.
Serial numbers are included in the first column of each worksheet.
The script uses the ImportExcel module to work with Excel files.

# License
This project is licensed under the MIT License.
Feel free to adjust the content according to your preferences or additional information about the project!
