# Load the Outlook Application
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")

# Get the Inbox folder
$inbox = $namespace.GetDefaultFolder(6)

# Create or get the "Processed" folder
$processedFolder = $inbox.Folders | Where-Object { $_.Name -eq "Processed" }
if (-not $processedFolder) {
    $processedFolder = $inbox.Folders.Add("Processed")
}

# Search for emails with the subject "Daily Dollars Raw"
$emails = $inbox.Items | Where-Object { $_.Subject -eq "Daily Dollars Raw" }

# Process each email
foreach ($email in $emails) {
    # Check if the email is already in the "Processed" folder
    if ($email.Parent -eq $processedFolder) {
        continue
    }

    # Extract the date
    $dateReceived = $email.ReceivedTime
    $previousDay = $dateReceived.AddDays(-1).ToString("yyyy-MM-dd")

    # Extract the attachment
    foreach ($attachment in $email.Attachments) {
        if ($attachment.FileName -match "\.csv$") {
            $filename = "C:\Users\ben.brannen\OneDrive - RIBBIT\DailyTransactionAutomation\$($attachment.FileName)"
            $attachment.SaveAsFile($filename)

            # Read CSV content and handle "LLC ," issue
            $csvContent = Get-Content -Path $filename -Raw
            $csvContent = $csvContent -replace ', LLC', '"LLC"'
            $csvData = $csvContent | ConvertFrom-Csv

            # Replace the placeholder back to ", LLC" in the loaded data
            foreach ($row in $csvData) {
                foreach ($property in $row.PSObject.Properties) {
                    $property.Value = $property.Value -replace '"LLC"', ', LLC'
                    $property.Value = ($property.Value -replace '\s+', ' ').Trim()
                }
            }

            # Convert CSV to Excel
            $excel = New-Object -ComObject Excel.Application
            $workbook = $excel.Workbooks.Add()
            $worksheet = $workbook.ActiveSheet
            $colIndex = 1
            foreach ($header in $csvData[0].PSObject.Properties.Name) {
                $worksheet.Cells.Item(1, $colIndex).Value2 = $header
                $colIndex++
            }
            $worksheet.Cells.Item(1, $colIndex).Value2 = "Date"
            $worksheet.Cells.Item(1, $colIndex + 1).Value2 = "BU"
            $rowIndex = 2
            foreach ($row in $csvData) {
                $colIndex = 1
                foreach ($value in $row.PSObject.Properties.Value) {
                    $worksheet.Cells.Item($rowIndex, $colIndex).Value2 = $value
                    $colIndex++
                }
                $worksheet.Cells.Item($rowIndex, $colIndex).Value2 = $previousDay
                $worksheet.Cells.Item($rowIndex, $colIndex + 1).Value2 = "ValidiFI"
                $rowIndex++
            }
            $outputPath = "C:\Users\ben.brannen\OneDrive - RIBBIT\DailyTransactionAutomation\output_$previousDay.xlsx"
            try {
                if (Test-Path $outputPath) {
                    Remove-Item $outputPath
                }
            } catch {
                Write-Error "Failed to delete file: $_"
                exit
            }
            try {
                $workbook.SaveAs($outputPath)
            } catch {
                Write-Error "Failed to save workbook: $_"
                exit
            }
            $workbook.Close()
            $excel.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet)
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)

            # Remove the CSV file
            Remove-Item $filename -Force
        }
    }

    # Move the email to the "Processed" folder
    $email.Move($processedFolder)
}
