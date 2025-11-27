# DataImporter.psm1 - Import data from CSV and Excel files

function Import-CSVData {
    param(
        [Parameter(Mandatory=$true)]
        [string]$FilePath
    )
    
    Write-Host "Importing CSV file: $FilePath"
    
    try {
        $csvData = Import-Csv -Path $FilePath -Encoding UTF8
        
        # Convert to application objects
        $applications = @()
        $rowNumber = 1
        
        foreach ($row in $csvData) {
            # Try to find application name column (flexible column naming)
            $appName = $null
            
            # Check common column names
            $possibleColumns = @(
                'ApplicationName', 'Application Name', 'App Name', 'AppName',
                'Name', 'Software', 'SoftwareName', 'Software Name',
                'Product', 'ProductName', 'Product Name', 'Title'
            )
            
            foreach ($colName in $possibleColumns) {
                if ($row.PSObject.Properties.Name -contains $colName) {
                    $appName = $row.$colName
                    break
                }
            }
            
            # If still not found, use first column
            if ([string]::IsNullOrWhiteSpace($appName)) {
                $firstProp = $row.PSObject.Properties | Select-Object -First 1
                $appName = $firstProp.Value
            }
            
            # Skip empty rows
            if ([string]::IsNullOrWhiteSpace($appName)) {
                continue
            }
            
            $applications += [PSCustomObject]@{
                RowNumber = $rowNumber
                ApplicationName = $appName.Trim()
                WingetID = ""
                MatchedName = ""
                Version = ""
                Status = "Pending"
                Confidence = "N/A"
                SearchStrategy = "None"
                MappingType = ""
            }
            
            $rowNumber++
        }
        
        Write-Host "Imported $($applications.Count) applications from CSV"
        return $applications
    }
    catch {
        Write-Host "Error importing CSV: $_"
        throw "Failed to import CSV file: $_"
    }
}

function Import-ExcelData {
    param(
        [Parameter(Mandatory=$true)]
        [string]$FilePath
    )
    
    Write-Host "Importing Excel file: $FilePath"
    
    try {
        # Create Excel COM object
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        
        # Open workbook
        $workbook = $excel.Workbooks.Open($FilePath)
        $worksheet = $workbook.Worksheets.Item(1)
        
        # Get used range
        $usedRange = $worksheet.UsedRange
        $rowCount = $usedRange.Rows.Count
        $columnCount = $usedRange.Columns.Count
        
        Write-Host "Excel file has $rowCount rows and $columnCount columns"
        
        # Read header row to find application name column
        $headers = @()
        for ($col = 1; $col -le $columnCount; $col++) {
            $headerValue = $worksheet.Cells.Item(1, $col).Text
            $headers += $headerValue
        }
        
        Write-Host "Headers: $($headers -join ', ')"
        
        # Find application name column
        $appNameColumnIndex = 0
        $possibleColumns = @(
            'ApplicationName', 'Application Name', 'App Name', 'AppName',
            'Name', 'Software', 'SoftwareName', 'Software Name',
            'Product', 'ProductName', 'Product Name', 'Title'
        )
        
        for ($i = 0; $i -lt $headers.Count; $i++) {
            foreach ($colName in $possibleColumns) {
                if ($headers[$i] -like "*$colName*") {
                    $appNameColumnIndex = $i + 1
                    break
                }
            }
            if ($appNameColumnIndex -gt 0) { break }
        }
        
        # If not found, use first column
        if ($appNameColumnIndex -eq 0) {
            $appNameColumnIndex = 1
        }
        
        Write-Host "Using column $appNameColumnIndex for application names"
        
        # Read data rows
        $applications = @()
        $rowNumber = 1
        
        for ($row = 2; $row -le $rowCount; $row++) {
            $appName = $worksheet.Cells.Item($row, $appNameColumnIndex).Text
            
            # Skip empty rows
            if ([string]::IsNullOrWhiteSpace($appName)) {
                continue
            }
            
            $applications += [PSCustomObject]@{
                RowNumber = $rowNumber
                ApplicationName = $appName.Trim()
                WingetID = ""
                MatchedName = ""
                Version = ""
                Status = "Pending"
                Confidence = "N/A"
                SearchStrategy = "None"
                MappingType = ""
            }
            
            $rowNumber++
        }
        
        # Close and cleanup
        $workbook.Close($false)
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        
        Write-Host "Imported $($applications.Count) applications from Excel"
        return $applications
    }
    catch {
        Write-Host "Error importing Excel: $_"
        
        # Cleanup on error
        if ($workbook) {
            $workbook.Close($false)
        }
        if ($excel) {
            $excel.Quit()
        }
        
        throw "Failed to import Excel file: $_"
    }
}

Export-ModuleMember -Function Import-CSVData, Import-ExcelData
