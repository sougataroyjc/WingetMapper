# WingetDiscovery.ps1 - Intelligent Winget Package Discovery Tool
# Entry point for the application

# Resolve base paths
$ScriptBasePath = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
$LogDir = Join-Path $ScriptBasePath "Logs"
$OutputDir = Join-Path $ScriptBasePath "Output"
$CacheDir = Join-Path $ScriptBasePath "Cache"

# Create necessary folders if not existing
foreach ($folder in @($LogDir, $OutputDir, $CacheDir)) {
    if (-not (Test-Path $folder)) {
        New-Item -Path $folder -ItemType Directory | Out-Null
    }
}

# Import modules with Global scope
$modulesPath = Join-Path $ScriptBasePath "Modules"

Import-Module (Join-Path $modulesPath "UILoader\UILoader.psm1") -Force -Global
Import-Module (Join-Path $modulesPath "WingetSearch\WingetSearch.psm1") -Force -Global
Import-Module (Join-Path $modulesPath "DataImporter\DataImporter.psm1") -Force -Global
Import-Module (Join-Path $modulesPath "ExportManager\ExportManager.psm1") -Force -Global
Import-Module (Join-Path $modulesPath "ManualMapping\ManualMapping.psm1") -Force -Global

Write-Host "Winget Package Discovery Tool - Starting..."
Write-Host "Base Path: $ScriptBasePath"

Add-Type -AssemblyName System.Windows.Forms

# Initialize global collections
$Global:ApplicationList = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
$Global:SearchCache = @{}
$Global:AutoMappedCount = 0
$Global:KnownApps = @{}
$script:CancelRequested = $false
$script:MappingFilePath = Join-Path $ScriptBasePath "MappedApplications.json"
$script:KnownAppsPath = Join-Path $ScriptBasePath "KnownApplications.json"

# Load known applications database
if (Test-Path $script:KnownAppsPath) {
    try {
        $knownAppsData = Get-Content $script:KnownAppsPath | ConvertFrom-Json
        foreach ($app in $knownAppsData) {
            $Global:KnownApps[$app.ApplicationName.ToLower()] = $app.WingetID
            foreach ($alias in $app.Aliases) {
                $Global:KnownApps[$alias.ToLower()] = $app.WingetID
            }
        }
        Write-Host "Loaded $($knownAppsData.Count) known applications"
    } catch {
        Write-Host "Failed to load known applications: $_"
    }
}

# Show main window
$window = Show-DiscoveryWindow

# Get UI controls
$importBtn = $window.FindName('ImportBtn')
$searchAllBtn = $window.FindName('SearchAllBtn')
$exportBtn = $window.FindName('ExportBtn')
$clearBtn = $window.FindName('ClearBtn')
$updateDatabaseBtn = $window.FindName('UpdateDatabaseBtn')
$cancelBtn = $window.FindName('CancelBtn')
$dataGrid = $window.FindName('ApplicationDataGrid')
$statusText = $window.FindName('StatusText')
$progressBar = $window.FindName('ProgressBar')
$totalAppsText = $window.FindName('TotalAppsText')
$foundText = $window.FindName('FoundText')
$notFoundText = $window.FindName('NotFoundText')
$progressText = $window.FindName('ProgressText')

# Set data binding
$dataGrid.ItemsSource = $Global:ApplicationList

# Helper function to update status
function Update-Status {
    param([string]$Text, [string]$Color = "#27AE60")
    $statusText.Dispatcher.Invoke([action]{
        $statusText.Text = $Text
        $statusText.Foreground = $Color
    })
}

# Helper function to update statistics
function Update-Statistics {
    $total = $Global:ApplicationList.Count
    $found = ($Global:ApplicationList | Where-Object { $_.Status -eq "Found" }).Count
    $notFound = ($Global:ApplicationList | Where-Object { $_.Status -eq "Not Found" }).Count
    $successRate = if ($total -gt 0) { [math]::Round(($found / $total) * 100, 1) } else { 0 }
    
    $totalAppsText.Dispatcher.Invoke([action]{ $totalAppsText.Text = $total })
    $foundText.Dispatcher.Invoke([action]{ $foundText.Text = $found })
    $notFoundText.Dispatcher.Invoke([action]{ $notFoundText.Text = $notFound })
    
    $successRateText = $window.FindName('SuccessRateText')
    if ($successRateText) {
        $successRateText.Dispatcher.Invoke([action]{ $successRateText.Text = "$successRate%" })
    }
    
    $autoMappedText = $window.FindName('AutoMappedText')
    if ($autoMappedText) {
        $autoMappedText.Dispatcher.Invoke([action]{ $autoMappedText.Text = $Global:AutoMappedCount })
    }
}

# Function to update known apps database
function Update-KnownAppsDatabase {
    try {
        Write-Host "Updating known apps database..."
        
        # Load existing known apps
        $existingApps = @()
        if (Test-Path $script:KnownAppsPath) {
            $existingApps = Get-Content $script:KnownAppsPath | ConvertFrom-Json
        }
        
        # Create a hashtable for quick lookup by WingetID
        $knownAppsHash = @{}
        foreach ($app in $existingApps) {
            $knownAppsHash[$app.WingetID] = $app
        }
        
        # Get all successfully mapped applications
        $mappedApps = $Global:ApplicationList | Where-Object { 
            $_.Status -eq "Found" -and 
            -not [string]::IsNullOrWhiteSpace($_.WingetID) 
        }
        
        $newAppsAdded = 0
        $aliasesAdded = 0
        
        foreach ($app in $mappedApps) {
            $wingetID = $app.WingetID
            $appName = $app.ApplicationName
            
            # Clean the app name
            $cleanName = $appName -replace '\d+\.\d+[\d\.]*', ''
            $cleanName = $cleanName -replace '\(.*?\)', ''
            $cleanName = $cleanName -replace '\[.*?\]', ''
            $cleanName = $cleanName.Trim()
            
            # Skip if empty after cleaning
            if ([string]::IsNullOrWhiteSpace($cleanName)) {
                continue
            }
            
            if ($knownAppsHash.ContainsKey($wingetID)) {
                # App exists - check if we should add this as an alias
                $existingApp = $knownAppsHash[$wingetID]
                $mainName = $existingApp.ApplicationName.ToLower()
                $currentAliases = @($existingApp.Aliases)
                
                # Check if this name is different and not already in aliases
                if ($cleanName.ToLower() -ne $mainName -and 
                    $cleanName -notin $currentAliases -and
                    $cleanName.ToLower() -notin ($currentAliases | ForEach-Object { $_.ToLower() })) {
                    
                    # Add as alias
                    $existingApp.Aliases = @($existingApp.Aliases) + @($cleanName)
                    $aliasesAdded++
                    Write-Host "  Added alias '$cleanName' to $wingetID"
                }
            }
            else {
                # New app - add it
                $newApp = [PSCustomObject]@{
                    ApplicationName = $cleanName
                    WingetID = $wingetID
                    Aliases = @()
                }
                $knownAppsHash[$wingetID] = $newApp
                $newAppsAdded++
                Write-Host "  Added new app: $cleanName -> $wingetID"
            }
        }
        
        # Convert back to array and sort
        $updatedApps = $knownAppsHash.Values | Sort-Object -Property ApplicationName
        
        # Save to file with pretty formatting
        $json = $updatedApps | ConvertTo-Json -Depth 10
        Set-Content -Path $script:KnownAppsPath -Value $json -Encoding UTF8
        
        # Reload the global known apps
        $Global:KnownApps.Clear()
        foreach ($app in $updatedApps) {
            $Global:KnownApps[$app.ApplicationName.ToLower()] = $app.WingetID
            foreach ($alias in $app.Aliases) {
                $Global:KnownApps[$alias.ToLower()] = $app.WingetID
            }
        }
        
        Write-Host "Known apps database updated successfully"
        Write-Host "  New apps added: $newAppsAdded"
        Write-Host "  Aliases added: $aliasesAdded"
        
        return @{
            Success = $true
            NewApps = $newAppsAdded
            Aliases = $aliasesAdded
        }
    }
    catch {
        Write-Host "Error updating known apps database: $_"
        return @{
            Success = $false
            Error = $_
        }
    }
}

# Import Button Click
$importBtn.Add_Click({
    Write-Host "Import button clicked"
    Update-Status "Importing file..." "#3498DB"
    
    try {
        Add-Type -AssemblyName System.Windows.Forms
        
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.Filter = "Excel/CSV files (*.xlsx;*.csv)|*.xlsx;*.csv|All files (*.*)|*.*"
        $openFileDialog.Title = "Select Application List File"
        $openFileDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
        
        if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $filePath = $openFileDialog.FileName
            Write-Host "Selected file: $filePath"
            
            $Global:ApplicationList.Clear()
            
            $extension = [System.IO.Path]::GetExtension($filePath).ToLower()
            
            Update-Status "Loading file..." "#E67E22"
            
            if ($extension -eq ".csv") {
                $data = Import-CSVData -FilePath $filePath
            } elseif ($extension -eq ".xlsx") {
                $data = Import-ExcelData -FilePath $filePath
            } else {
                throw "Unsupported file format: $extension"
            }
            
            Write-Host "Adding $($data.Count) applications to collection"
            
            foreach ($item in $data) {
                $Global:ApplicationList.Add($item) | Out-Null
            }
            
            Update-Statistics
            Update-Status "Imported $($data.Count) applications" "#27AE60"
            
            $searchAllBtn.IsEnabled = $true
            $exportBtn.IsEnabled = $true
            $updateDatabaseBtn.IsEnabled = $true
        }
    }
    catch {
        $errorMsg = "Import failed: $_"
        Write-Host $errorMsg
        Update-Status "Import failed" "#E74C3C"
        [System.Windows.MessageBox]::Show($errorMsg, "Import Error", "OK", "Error")
    }
})

# Search All Button Click
$searchAllBtn.Add_Click({
    Write-Host "Search All button clicked"
    
    $script:CancelRequested = $false
    $Global:AutoMappedCount = 0
    $searchAllBtn.IsEnabled = $false
    $importBtn.IsEnabled = $false
    $cancelBtn.Visibility = "Visible"
    $progressBar.Visibility = "Visible"
    $progressText.Visibility = "Visible"
    
    Update-Status "Searching Winget repository..." "#E67E22"
    
    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.ApartmentState = "STA"
    $runspace.ThreadOptions = "ReuseThread"
    $runspace.Open()
    $runspace.SessionStateProxy.SetVariable("ApplicationList", $Global:ApplicationList)
    $runspace.SessionStateProxy.SetVariable("SearchCache", $Global:SearchCache)
    $runspace.SessionStateProxy.SetVariable("KnownApps", $Global:KnownApps)
    $runspace.SessionStateProxy.SetVariable("CancelFlag", $script:CancelRequested)
    $runspace.SessionStateProxy.SetVariable("dataGrid", $dataGrid)
    $runspace.SessionStateProxy.SetVariable("progressBar", $progressBar)
    $runspace.SessionStateProxy.SetVariable("progressText", $progressText)
    $runspace.SessionStateProxy.SetVariable("AutoMappedCount", $Global:AutoMappedCount)
    $runspace.SessionStateProxy.SetVariable("MappingFile", $script:MappingFilePath)
    $runspace.SessionStateProxy.SetVariable("ModulePath", (Join-Path $ScriptBasePath "Modules\WingetSearch\WingetSearch.psm1"))
    
    $powershell = [powershell]::Create().AddScript({
        param($appList, $cache, $knownApps, $dg, $pb, $pt, $mapFile, $modPath)
        
        Import-Module $modPath -Force
        
        $total = $appList.Count
        $current = 0
        $foundCount = 0
        
        foreach ($app in $appList) {
            $current++
            $percentComplete = [math]::Round(($current / $total) * 100)
            
            $pb.Dispatcher.Invoke([action]{ 
                $pb.Value = $percentComplete 
            })
            
            $pt.Dispatcher.Invoke([action]{
                $pt.Text = "Processing $current of $total..."
            })
            
            if ($app.Status -ne "Pending") {
                continue
            }
            
            $result = Search-WingetIntelligent -ApplicationName $app.ApplicationName -Cache $cache -KnownApps $knownApps
            
            $app.WingetID = $result.WingetID
            $app.MatchedName = $result.MatchedName
            $app.Version = $result.Version
            $app.Status = $result.Status
            $app.Confidence = $result.Confidence
            $app.SearchStrategy = $result.SearchStrategy
            $app.MappingType = "Auto"
            
            if ($result.Status -eq "Found") {
                $foundCount++
            }
            
            if ($current % 5 -eq 0) {
                $dg.Dispatcher.Invoke([action]{ $dg.Items.Refresh() }, [System.Windows.Threading.DispatcherPriority]::Background)
            }
            
            Start-Sleep -Milliseconds 50
        }
        
        $dg.Dispatcher.Invoke([action]{ $dg.Items.Refresh() })
        
        return $foundCount
    }).AddArgument($Global:ApplicationList).AddArgument($Global:SearchCache).AddArgument($Global:KnownApps).AddArgument($dataGrid).AddArgument($progressBar).AddArgument($progressText).AddArgument($script:MappingFilePath).AddArgument((Join-Path $ScriptBasePath "Modules\WingetSearch\WingetSearch.psm1"))
    
    $powershell.Runspace = $runspace
    $handle = $powershell.BeginInvoke()
    
    $timer = New-Object System.Windows.Threading.DispatcherTimer
    $timer.Interval = [TimeSpan]::FromMilliseconds(500)
    $timer.Add_Tick({
        if ($handle.IsCompleted) {
            $timer.Stop()
            
            try {
                $foundCount = $powershell.EndInvoke($handle)
                $Global:AutoMappedCount = $foundCount
            } catch {
                Write-Host "Search error: $_"
            }
            
            $powershell.Dispose()
            $runspace.Close()
            
            $dataGrid.Items.Refresh()
            Update-Statistics
            
            Save-MappingsToFile -FilePath $script:MappingFilePath -Applications $Global:ApplicationList
            
            $progressBar.Value = 0
            $progressBar.Visibility = "Collapsed"
            $progressText.Visibility = "Collapsed"
            $cancelBtn.Visibility = "Collapsed"
            
            $searchAllBtn.IsEnabled = $true
            $importBtn.IsEnabled = $true
            
            $found = ($Global:ApplicationList | Where-Object { $_.Status -eq "Found" }).Count
            Update-Status "Search completed: $found packages found" "#27AE60"
            
            # Prompt to update known apps database
            if ($found -gt 0) {
                $promptMessage = "Search completed successfully!`n`n"
                $promptMessage += "Would you like to learn from these results?`n`n"
                $promptMessage += "This will help improve future searches by adding new apps and aliases."
                
                $result = [System.Windows.MessageBox]::Show(
                    $promptMessage,
                    "Learn from Results",
                    "YesNo",
                    "Question"
                )
                
                if ($result -eq "Yes") {
                    $updateResult = Update-KnownAppsDatabase
                    
                    if ($updateResult.Success) {
                        $successMessage = "Database updated successfully!`n`n"
                        $successMessage += "New apps added: $($updateResult.NewApps)`n"
                        $successMessage += "Aliases added: $($updateResult.Aliases)"
                        
                        [System.Windows.MessageBox]::Show(
                            $successMessage,
                            "Learning Complete",
                            "OK",
                            "Information"
                        )
                    }
                    else {
                        [System.Windows.MessageBox]::Show(
                            "Failed to update database.`n`nError: $($updateResult.Error)",
                            "Update Failed",
                            "OK",
                            "Error"
                        )
                    }
                }
            }
        } else {
            Update-Statistics
        }
    })
    $timer.Start()
})

# Cancel Button Click
$cancelBtn.Add_Click({
    $script:CancelRequested = $true
    Update-Status "Cancelling..." "#E74C3C"
})

# Update Database Button Click
$updateDatabaseBtn.Add_Click({
    Write-Host "Learn from Results button clicked"
    
    $found = ($Global:ApplicationList | Where-Object { $_.Status -eq "Found" }).Count
    
    if ($found -eq 0) {
        [System.Windows.MessageBox]::Show(
            "No mapped applications found.`n`nPlease run 'Search All' first to find Winget packages.",
            "No Data",
            "OK",
            "Warning"
        )
        return
    }
    
    $message = "Learn from $found mapped applications?`n`n"
    $message += "This will:`n"
    $message += "- Add new applications to the database`n"
    $message += "- Add aliases for existing applications`n"
    $message += "- Improve future search accuracy`n`n"
    $message += "Do you want to continue?"
    
    $result = [System.Windows.MessageBox]::Show(
        $message,
        "Learn from Results",
        "YesNo",
        "Question"
    )
    
    if ($result -eq "Yes") {
        Update-Status "Learning from results..." "#3498DB"
        
        $updateResult = Update-KnownAppsDatabase
        
        if ($updateResult.Success) {
            $successMessage = "Database updated successfully!`n`n"
            $successMessage += "New apps added: $($updateResult.NewApps)`n"
            $successMessage += "Aliases added: $($updateResult.Aliases)"
            
            Update-Status "Database updated successfully" "#27AE60"
            
            [System.Windows.MessageBox]::Show(
                $successMessage,
                "Learning Complete",
                "OK",
                "Information"
            )
        }
        else {
            Update-Status "Database update failed" "#E74C3C"
            
            [System.Windows.MessageBox]::Show(
                "Failed to update database.`n`nError: $($updateResult.Error)",
                "Update Failed",
                "OK",
                "Error"
            )
        }
    }
})

# Export Button Click
$exportBtn.Add_Click({
    Write-Host "Export button clicked"
    
    try {
        Add-Type -AssemblyName System.Windows.Forms
        
        $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|JSON files (*.json)|*.json|CSV files (*.csv)|*.csv"
        $saveFileDialog.Title = "Export Results"
        $saveFileDialog.FileName = "WingetDiscoveryResults_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
        
        if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $exportPath = $saveFileDialog.FileName
            $extension = [System.IO.Path]::GetExtension($exportPath).ToLower()
            
            Update-Status "Exporting results..." "#3498DB"
            
            if ($extension -eq ".csv") {
                Export-ResultsToCSV -Data $Global:ApplicationList -FilePath $exportPath
            } elseif ($extension -eq ".xlsx") {
                Export-ResultsToExcel -Data $Global:ApplicationList -FilePath $exportPath
            } elseif ($extension -eq ".json") {
                Export-ResultsToJSON -Data $Global:ApplicationList -FilePath $exportPath
            }
            
            Update-Status "Export completed" "#27AE60"
            [System.Windows.MessageBox]::Show("Results exported successfully", "Export Complete", "OK", "Information")
            
            Start-Process "explorer.exe" "/select,$exportPath"
        }
    }
    catch {
        $errorMsg = "Export failed: $_"
        Write-Host $errorMsg
        Update-Status "Export failed" "#E74C3C"
        [System.Windows.MessageBox]::Show($errorMsg, "Export Error", "OK", "Error")
    }
})

# Clear Button Click
$clearBtn.Add_Click({
    $result = [System.Windows.MessageBox]::Show(
        "Are you sure you want to clear all data?",
        "Clear Data",
        "YesNo",
        "Question"
    )
    
    if ($result -eq "Yes") {
        $Global:ApplicationList.Clear()
        Update-Statistics
        Update-Status "Data cleared" "#95A5A6"
        $searchAllBtn.IsEnabled = $false
        $exportBtn.IsEnabled = $false
        $updateDatabaseBtn.IsEnabled = $false
    }
})

# Initialize status
Update-Status "Ready" "#27AE60"
Update-Statistics

# Handle Map Manually button clicks
$dataGrid.AddHandler(
    [System.Windows.Controls.Button]::ClickEvent,
    [System.Windows.RoutedEventHandler]{
        param($sender, $e)
        
        $button = $e.OriginalSource -as [System.Windows.Controls.Button]
        
        if ($button -ne $null -and $button.Name -eq 'MapManuallyBtn') {
            $app = $button.DataContext
            if ($app) {
                Show-ManualMappingWindow -App $app -DataGrid $dataGrid -MappingFile $script:MappingFilePath -OnMappingSaved {
                    Update-Statistics
                }
            }
        }
    }
)

# Handle Undo Mapping context menu
$undoMenuItem = $window.FindName('UndoMappingMenuItem')
if ($undoMenuItem) {
    $undoMenuItem.Add_Click({
        $selectedApp = $dataGrid.SelectedItem
        if ($selectedApp -and $selectedApp.Status -eq "Found") {
            $result = [System.Windows.MessageBox]::Show(
                "Mark '$($selectedApp.ApplicationName)' as Not Found?`n`nThis will clear: $($selectedApp.WingetID)",
                "Undo Mapping",
                "YesNo",
                "Question"
            )
            
            if ($result -eq "Yes") {
                $selectedApp.WingetID = ""
                $selectedApp.MatchedName = ""
                $selectedApp.Version = ""
                $selectedApp.Status = "Not Found"
                $selectedApp.Confidence = "N/A"
                $selectedApp.SearchStrategy = "Undone"
                $selectedApp.MappingType = ""
                
                if ($Global:AutoMappedCount -gt 0) {
                    $Global:AutoMappedCount--
                }
                
                $dataGrid.Items.Refresh()
                Update-Statistics
                
                Save-MappingsToFile -FilePath $script:MappingFilePath -Applications $Global:ApplicationList
            }
        }
    })
}

# Helper function to save mappings
function Save-MappingsToFile {
    param($FilePath, $Applications)
    
    $mappings = @()
    foreach ($app in $Applications) {
        if ($app.Status -eq "Found") {
            $mappings += [PSCustomObject]@{
                ApplicationName = $app.ApplicationName
                WingetID = $app.WingetID
                MatchedName = $app.MatchedName
                Version = $app.Version
                Confidence = $app.Confidence
                SearchStrategy = $app.SearchStrategy
                MappingType = $app.MappingType
                LastUpdated = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
            }
        }
    }
    
    $mappings | ConvertTo-Json | Out-File $FilePath -Encoding UTF8
}

# Display the main window
$window.ShowDialog() | Out-Null
