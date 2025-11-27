# ExportManager.psm1 - Export results to CSV and Excel formats

function Export-ResultsToCSV {
    param(
        [Parameter(Mandatory=$true)]
        $Data,
        
        [Parameter(Mandatory=$true)]
        [string]$FilePath
    )
    
    Write-Host "Exporting to CSV: $FilePath"
    
    try {
        # Convert data to export format
        $exportData = @()
        
        foreach ($item in $Data) {
            $exportData += [PSCustomObject]@{
                'Row Number' = $item.RowNumber
                'Application Name' = $item.ApplicationName
                'Winget ID' = $item.WingetID
                'Matched Package Name' = $item.MatchedName
                'Latest Version' = $item.Version
                'Status' = $item.Status
                'Confidence' = $item.Confidence
                'Search Strategy' = $item.SearchStrategy
            }
        }
        
        # Export to CSV
        $exportData | Export-Csv -Path $FilePath -NoTypeInformation -Encoding UTF8
        
        Write-Host "Successfully exported $($exportData.Count) records to CSV"
        return $true
    }
    catch {
        Write-Host "Error exporting to CSV: $_"
        throw "Failed to export to CSV: $_"
    }
}

function Export-ResultsToExcel {
    param(
        [Parameter(Mandatory=$true)]
        $Data,
        
        [Parameter(Mandatory=$true)]
        [string]$FilePath
    )
    
    Write-Host "Exporting to Excel: $FilePath"
    
    try {
        # Create Excel COM object
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        
        # Create new workbook
        $workbook = $excel.Workbooks.Add()
        $worksheet = $workbook.Worksheets.Item(1)
        $worksheet.Name = "Winget Discovery Results"
        
        # Define headers
        $headers = @(
            'Row Number',
            'Application Name',
            'Winget ID',
            'Matched Package Name',
            'Latest Version',
            'Status',
            'Confidence',
            'Search Strategy'
        )
        
        # Write headers
        for ($col = 1; $col -le $headers.Count; $col++) {
            $worksheet.Cells.Item(1, $col) = $headers[$col - 1]
            $worksheet.Cells.Item(1, $col).Font.Bold = $true
            $worksheet.Cells.Item(1, $col).Interior.Color = 15773696  # Light blue
            $worksheet.Cells.Item(1, $col).Font.Color = 16777215  # White
        }
        
        # Write data
        $row = 2
        foreach ($item in $Data) {
            $worksheet.Cells.Item($row, 1) = $item.RowNumber
            $worksheet.Cells.Item($row, 2) = $item.ApplicationName
            $worksheet.Cells.Item($row, 3) = $item.WingetID
            $worksheet.Cells.Item($row, 4) = $item.MatchedName
            $worksheet.Cells.Item($row, 5) = $item.Version
            $worksheet.Cells.Item($row, 6) = $item.Status
            $worksheet.Cells.Item($row, 7) = $item.Confidence
            $worksheet.Cells.Item($row, 8) = $item.SearchStrategy
            
            # Color code based on status
            switch ($item.Status) {
                "Found" {
                    $worksheet.Cells.Item($row, 6).Interior.Color = 13561798  # Light green
                    $worksheet.Cells.Item($row, 6).Font.Color = 0  # Black
                }
                "Not Found" {
                    $worksheet.Cells.Item($row, 6).Interior.Color = 13421823  # Light red
                    $worksheet.Cells.Item($row, 6).Font.Color = 0  # Black
                }
                "Pending" {
                    $worksheet.Cells.Item($row, 6).Interior.Color = 16777164  # Light yellow
                    $worksheet.Cells.Item($row, 6).Font.Color = 0  # Black
                }
            }
            
            $row++
        }
        
        # Auto-fit columns
        $usedRange = $worksheet.UsedRange
        $usedRange.EntireColumn.AutoFit() | Out-Null
        
        # Add filters
        $usedRange.AutoFilter() | Out-Null
        
        # Freeze header row
        $worksheet.Application.ActiveWindow.SplitRow = 1
        $worksheet.Application.ActiveWindow.FreezePanes = $true
        
        # Add summary statistics
        $summaryRow = $row + 2
        
        $totalCount = $Data.Count
        $foundCount = ($Data | Where-Object { $_.Status -eq "Found" }).Count
        $notFoundCount = ($Data | Where-Object { $_.Status -eq "Not Found" }).Count
        $successRate = if ($totalCount -gt 0) { [math]::Round(($foundCount / $totalCount) * 100, 2) } else { 0 }
        
        $worksheet.Cells.Item($summaryRow, 1) = "Summary Statistics:"
        $worksheet.Cells.Item($summaryRow, 1).Font.Bold = $true
        
        $worksheet.Cells.Item($summaryRow + 1, 1) = "Total Applications:"
        $worksheet.Cells.Item($summaryRow + 1, 2) = $totalCount
        
        $worksheet.Cells.Item($summaryRow + 2, 1) = "Packages Found:"
        $worksheet.Cells.Item($summaryRow + 2, 2) = $foundCount
        $worksheet.Cells.Item($summaryRow + 2, 2).Interior.Color = 13561798
        
        $worksheet.Cells.Item($summaryRow + 3, 1) = "Not Found:"
        $worksheet.Cells.Item($summaryRow + 3, 2) = $notFoundCount
        $worksheet.Cells.Item($summaryRow + 3, 2).Interior.Color = 13421823
        
        $worksheet.Cells.Item($summaryRow + 4, 1) = "Success Rate:"
        $worksheet.Cells.Item($summaryRow + 4, 2) = "$successRate%"
        $worksheet.Cells.Item($summaryRow + 4, 2).Font.Bold = $true
        
        # Save workbook
        $workbook.SaveAs($FilePath)
        $workbook.Close($false)
        $excel.Quit()
        
        # Cleanup
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        
        Write-Host "Successfully exported $totalCount records to Excel"
        return $true
    }
    catch {
        Write-Host "Error exporting to Excel: $_"
        
        # Cleanup on error
        if ($workbook) {
            $workbook.Close($false)
        }
        if ($excel) {
            $excel.Quit()
        }
        
        throw "Failed to export to Excel: $_"
    }
}

function Export-ResultsToJSON {
    param(
        [Parameter(Mandatory=$true)]
        $Data,
        
        [Parameter(Mandatory=$true)]
        [string]$FilePath
    )
    
    Write-Host "Exporting to JSON: $FilePath"
    
    try {
        $exportData = @()
        
        foreach ($item in $Data) {
            $exportData += [PSCustomObject]@{
                'RowNumber' = $item.RowNumber
                'ApplicationName' = $item.ApplicationName
                'WingetID' = $item.WingetID
                'MatchedPackageName' = $item.MatchedName
                'LatestVersion' = $item.Version
                'Status' = $item.Status
                'Confidence' = $item.Confidence
                'SearchStrategy' = $item.SearchStrategy
                'MappingType' = $item.MappingType
            }
        }
        
        $exportData | ConvertTo-Json | Out-File -FilePath $FilePath -Encoding UTF8
        
        Write-Host "Successfully exported $($exportData.Count) records to JSON"
        return $true
    }
    catch {
        Write-Host "Error exporting to JSON: $_"
        throw "Failed to export to JSON: $_"
    }
}

Export-ModuleMember -Function Export-ResultsToCSV, Export-ResultsToExcel, Export-ResultsToJSON
