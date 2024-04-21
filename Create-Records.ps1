# Import the ImportExcel module to work with Excel files
Import-Module ImportExcel

# Prompt the user to input the year
$year = Read-Host "Just name the year"

# Prompt the user to provide headers separated by commas, then split them into an array
$hearders = (Read-Host "Provide headers (separated by commas only)").Split(",")

# Get the names of the months
$month = (New-Object System.Globalization.DateTimeFormatInfo).MonthNames

# Convert the year input to an integer
[int]$year_int = $year

# Define the number of days in each month based on the year (considering leap years)
$days = 31, $(if ($year_int % 4 -eq 0) {29} else {28}), 31, 30, 31, 30, 31, 31, 30, 31, 30, 31

# Create a directory for the given year
New-Item -ItemType dir $year | Out-Null

# Change the location to the directory of the given year
Set-Location $year

# Loop through each month
for ($i = 0; $i -lt $month.Count - 1; $i++) {
    # Inform the user about the file generation process for the current month
    Write-Output "File generation $($month[$i]).xlsx, Please wait..."

    # Create a new Excel application object
    $exel = New-Object -ComObject Excel.Application
    
    # Add a new workbook
    $workbook = $exel.Workbooks.Add()
    
    # Loop through each day in the current month
    for( $j = 0; $j -lt $days[$i]; $j++ ) {
        # If it's the first day, use the first worksheet; otherwise, add a new worksheet
        if( $j -eq 0 ) {
            $sheet = $workbook.Worksheets.Item(1)
        } else {
            $sheet = $workbook.Worksheets.Add()
        }

        # Set the name of the worksheet to the date
        $sheet.name = "$( if( $j -lt 9 ) { "0$( $j + 1 )" } else { $j + 1 } ).$( if( $i -lt 9 ) { "0$( $i + 1 )" } else { $i + 1 } ).$year"

        # Populate the cells with date, serial numbers, and headers
        $sheet.Cells.Item( 1, 2 ) = "$( if( $j -lt 9 ) { "0$( $j + 1 )" } else { $j + 1 } ).$( if( $i -lt 9 ) { "0$( $i + 1 )" } else { $i + 1 } ).$year"
        $sheet.Cells.Item( 2, 1 ) = "LP"
        for( $k = 0; $k -lt $hearders.Count; $k++ ) {
            $sheet.Cells.Item( 2, $k + 2 ) = $hearders[$k]
        }

        # Populate the first column with serial numbers
        for( $k = 0; $k -lt 40; $k++ ) {
            $sheet.Cells.Item( $k + 3, 1 ) = $k + 1
        }
    }

    # Inform the user about saving the file for the current month
    Write-Output "I'm saving the file $($month[$i]).xlsx, Please wait..."

    # Save the workbook
    $workbook.SaveAs( "$(Get-Location)\$($month[$i]).xlsx" )

    # Close Excel application
    $exel.Quit()

    # Release COM objects to avoid memory leaks
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($exel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

# Move back to the parent directory
Set-Location ..
