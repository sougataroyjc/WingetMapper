# ManualMapping.psm1 - Manual Winget mapping functionality

function Show-ManualMappingWindow {
    param(
        $App,
        $DataGrid,
        [string]$MappingFile,
        [scriptblock]$OnMappingSaved
    )
    
    Add-Type -AssemblyName PresentationFramework
    
    $window = New-Object System.Windows.Window
    $window.Title = "Manual Mapping - $($App.ApplicationName)"
    $window.Width = 900
    $window.Height = 600
    $window.WindowStartupLocation = "CenterScreen"
    $window.Background = [System.Windows.Media.Brushes]::WhiteSmoke
    
    $mainGrid = New-Object System.Windows.Controls.Grid
    $mainGrid.Margin = New-Object System.Windows.Thickness(20)
    
    $row1 = New-Object System.Windows.Controls.RowDefinition
    $row1.Height = [System.Windows.GridLength]::Auto
    $row2 = New-Object System.Windows.Controls.RowDefinition
    $row2.Height = [System.Windows.GridLength]::Auto
    $row3 = New-Object System.Windows.Controls.RowDefinition
    $row3.Height = New-Object System.Windows.GridLength(1, [System.Windows.GridUnitType]::Star)
    $row4 = New-Object System.Windows.Controls.RowDefinition
    $row4.Height = [System.Windows.GridLength]::Auto
    
    $mainGrid.RowDefinitions.Add($row1)
    $mainGrid.RowDefinitions.Add($row2)
    $mainGrid.RowDefinitions.Add($row3)
    $mainGrid.RowDefinitions.Add($row4)
    
    $headerBorder = New-Object System.Windows.Controls.Border
    $headerBorder.Background = [System.Windows.Media.Brushes]::White
    $headerBorder.CornerRadius = 8
    $headerBorder.Padding = New-Object System.Windows.Thickness(20)
    $headerBorder.Margin = New-Object System.Windows.Thickness(0, 0, 0, 15)
    [System.Windows.Controls.Grid]::SetRow($headerBorder, 0)
    
    $headerStack = New-Object System.Windows.Controls.StackPanel
    
    $headerText = New-Object System.Windows.Controls.TextBlock
    $headerText.Text = "Manual Winget Mapping"
    $headerText.FontSize = 22
    $headerText.FontWeight = "Bold"
    $headerText.Foreground = [System.Windows.Media.Brushes]::DarkSlateGray
    $headerStack.Children.Add($headerText)
    
    $appNameText = New-Object System.Windows.Controls.TextBlock
    $appNameText.Text = $App.ApplicationName
    $appNameText.FontSize = 16
    $appNameText.Foreground = [System.Windows.Media.Brushes]::Gray
    $appNameText.Margin = New-Object System.Windows.Thickness(0, 5, 0, 0)
    $headerStack.Children.Add($appNameText)
    
    $headerBorder.Child = $headerStack
    $mainGrid.Children.Add($headerBorder)
    
    $searchBorder = New-Object System.Windows.Controls.Border
    $searchBorder.Background = [System.Windows.Media.Brushes]::White
    $searchBorder.CornerRadius = 8
    $searchBorder.Padding = New-Object System.Windows.Thickness(15)
    $searchBorder.Margin = New-Object System.Windows.Thickness(0, 0, 0, 15)
    [System.Windows.Controls.Grid]::SetRow($searchBorder, 1)
    
    $searchGrid = New-Object System.Windows.Controls.Grid
    $scol1 = New-Object System.Windows.Controls.ColumnDefinition
    $scol1.Width = New-Object System.Windows.GridLength(1, [System.Windows.GridUnitType]::Star)
    $scol2 = New-Object System.Windows.Controls.ColumnDefinition
    $scol2.Width = [System.Windows.GridLength]::Auto
    $searchGrid.ColumnDefinitions.Add($scol1)
    $searchGrid.ColumnDefinitions.Add($scol2)
    
    $searchBox = New-Object System.Windows.Controls.TextBox
    $searchBox.Text = $App.ApplicationName
    $searchBox.FontSize = 14
    $searchBox.Padding = New-Object System.Windows.Thickness(10)
    $searchBox.Margin = New-Object System.Windows.Thickness(0, 0, 10, 0)
    [System.Windows.Controls.Grid]::SetColumn($searchBox, 0)
    $searchGrid.Children.Add($searchBox)
    
    $searchButton = New-Object System.Windows.Controls.Button
    $searchButton.Content = "Search Winget"
    $searchButton.Width = 150
    $searchButton.Height = 38
    $searchButton.Background = [System.Windows.Media.Brushes]::DodgerBlue
    $searchButton.Foreground = [System.Windows.Media.Brushes]::White
    $searchButton.FontSize = 14
    $searchButton.FontWeight = "SemiBold"
    $searchButton.Cursor = "Hand"
    [System.Windows.Controls.Grid]::SetColumn($searchButton, 1)
    $searchGrid.Children.Add($searchButton)
    
    $searchBorder.Child = $searchGrid
    $mainGrid.Children.Add($searchBorder)
    
    $resultsBorder = New-Object System.Windows.Controls.Border
    $resultsBorder.Background = [System.Windows.Media.Brushes]::White
    $resultsBorder.CornerRadius = 8
    $resultsBorder.Padding = New-Object System.Windows.Thickness(15)
    $resultsBorder.Margin = New-Object System.Windows.Thickness(0, 0, 0, 15)
    [System.Windows.Controls.Grid]::SetRow($resultsBorder, 2)
    
    $resultsDataGrid = New-Object System.Windows.Controls.DataGrid
    $resultsDataGrid.AutoGenerateColumns = $false
    $resultsDataGrid.IsReadOnly = $true
    $resultsDataGrid.CanUserAddRows = $false
    $resultsDataGrid.Background = [System.Windows.Media.Brushes]::White
    $resultsDataGrid.SelectionMode = "Single"
    $resultsDataGrid.FontSize = 13
    
    $col1 = New-Object System.Windows.Controls.DataGridTextColumn
    $col1.Header = "Package ID"
    $col1.Binding = New-Object System.Windows.Data.Binding("Id")
    $col1.Width = 250
    $resultsDataGrid.Columns.Add($col1)
    
    $col2 = New-Object System.Windows.Controls.DataGridTextColumn
    $col2.Header = "Name"
    $col2.Binding = New-Object System.Windows.Data.Binding("Name")
    $col2.Width = 300
    $resultsDataGrid.Columns.Add($col2)
    
    $col3 = New-Object System.Windows.Controls.DataGridTextColumn
    $col3.Header = "Version"
    $col3.Binding = New-Object System.Windows.Data.Binding("Version")
    $col3.Width = 150
    $resultsDataGrid.Columns.Add($col3)
    
    $resultsBorder.Child = $resultsDataGrid
    $mainGrid.Children.Add($resultsBorder)
    
    $buttonPanel = New-Object System.Windows.Controls.StackPanel
    $buttonPanel.Orientation = "Horizontal"
    $buttonPanel.HorizontalAlignment = "Right"
    [System.Windows.Controls.Grid]::SetRow($buttonPanel, 3)
    
    $saveButton = New-Object System.Windows.Controls.Button
    $saveButton.Content = "Save Mapping"
    $saveButton.Width = 120
    $saveButton.Height = 38
    $saveButton.Background = [System.Windows.Media.Brushes]::Green
    $saveButton.Foreground = [System.Windows.Media.Brushes]::White
    $saveButton.FontSize = 14
    $saveButton.FontWeight = "SemiBold"
    $saveButton.Cursor = "Hand"
    $saveButton.Margin = New-Object System.Windows.Thickness(0, 0, 10, 0)
    $saveButton.IsEnabled = $false
    $buttonPanel.Children.Add($saveButton)
    
    $cancelButton = New-Object System.Windows.Controls.Button
    $cancelButton.Content = "Cancel"
    $cancelButton.Width = 100
    $cancelButton.Height = 38
    $cancelButton.Background = [System.Windows.Media.Brushes]::Gray
    $cancelButton.Foreground = [System.Windows.Media.Brushes]::White
    $cancelButton.FontSize = 14
    $cancelButton.FontWeight = "SemiBold"
    $cancelButton.Cursor = "Hand"
    $cancelButton.Add_Click({ $window.Close() })
    $buttonPanel.Children.Add($cancelButton)
    
    $mainGrid.Children.Add($buttonPanel)
    
    $searchButton.Add_Click({
        $searchButton.IsEnabled = $false
        $searchButton.Content = "Searching..."
        $resultsDataGrid.ItemsSource = $null
        
        $searchQuery = $searchBox.Text
        
        try {
            $packages = @()
            $searchResults = winget search $searchQuery --accept-source-agreements 2>&1 | Out-String
            
            $lines = $searchResults -split "`n"
            $foundHeader = $false
            
            foreach ($line in $lines) {
                if ($line -match "^Name\s+Id\s+Version") {
                    $foundHeader = $true
                    continue
                }
                
                if ($foundHeader -and $line.Trim() -ne "" -and $line -notmatch "^-+" -and $packages.Count -lt 50) {
                    if ($line -match "^(.+?)\s{2,}([\w\.-]+\.\w+[\w\.-]*)\s+(.+)$") {
                        $newPackage = [PSCustomObject]@{
                            Name = $matches[1].Trim()
                            Id = $matches[2].Trim()
                            Version = $matches[3].Trim()
                        }
                        if (-not ($packages | Where-Object { $_.Id -eq $newPackage.Id })) {
                            $packages += $newPackage
                        }
                    }
                }
            }
            
            if ($packages.Count -eq 0) {
                [System.Windows.MessageBox]::Show("No packages found", "No Results", "OK", "Information")
            } else {
                $resultsDataGrid.ItemsSource = $packages
            }
        }
        catch {
            [System.Windows.MessageBox]::Show("Error searching Winget: $_", "Search Error", "OK", "Error")
        }
        finally {
            $searchButton.IsEnabled = $true
            $searchButton.Content = "Search Winget"
        }
    })
    
    $resultsDataGrid.Add_SelectionChanged({
        $selectedPackage = $resultsDataGrid.SelectedItem
        if ($selectedPackage) {
            $saveButton.IsEnabled = $true
        }
    })
    
    $saveButton.Add_Click({
        $selectedPackage = $resultsDataGrid.SelectedItem
        if ($selectedPackage) {
            try {
                $detailOutput = winget show --id $selectedPackage.Id --exact 2>&1 | Out-String
                $latestVersion = "Unknown"
                if ($detailOutput -match "Version:\s+(.+)") {
                    $latestVersion = $matches[1].Trim()
                }
                
                $App.WingetID = $selectedPackage.Id
                $App.MatchedName = $selectedPackage.Name
                $App.Version = $latestVersion
                $App.Status = "Found"
                $App.Confidence = "Manual"
                $App.SearchStrategy = "Manual Mapping"
                $App.MappingType = "Manual"
                
                if ($DataGrid) {
                    $DataGrid.Dispatcher.Invoke([action]{ $DataGrid.Items.Refresh() })
                }
                
                # Call the callback to update statistics
                if ($OnMappingSaved) {
                    & $OnMappingSaved
                }
                
                [System.Windows.MessageBox]::Show("Mapping saved successfully", "Success", "OK", "Information")
                
                $window.Close()
            }
            catch {
                [System.Windows.MessageBox]::Show("Error saving mapping: $_", "Error", "OK", "Error")
            }
        }
    })
    
    $window.Add_Loaded({
        $searchButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
    })
    
    $window.Content = $mainGrid
    $window.ShowDialog() | Out-Null
}

Export-ModuleMember -Function Show-ManualMappingWindow
