# WingetSearch.psm1 - Intelligent Winget search with multiple strategies

function Write-SearchLog {
    param ([string]$Message)
    $logDir = Join-Path $PSScriptRoot "..\..\Logs"
    if (-not (Test-Path $logDir)) {
        New-Item -Path $logDir -ItemType Directory | Out-Null
    }
    $logFile = Join-Path $logDir "WingetSearch.log"
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $Message" | Out-File -FilePath $logFile -Append -Encoding utf8
}

function Search-WingetIntelligent {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ApplicationName,
        
        [hashtable]$Cache = @{},
        [hashtable]$KnownApps = @{}
    )
    
    Write-SearchLog "Starting intelligent search for: $ApplicationName"
    
    if ($Cache.ContainsKey($ApplicationName)) {
        Write-SearchLog "Cache hit for: $ApplicationName"
        return $Cache[$ApplicationName]
    }
    
    $result = [PSCustomObject]@{
        WingetID = ""
        MatchedName = ""
        Version = ""
        Status = "Not Found"
        Confidence = "N/A"
        SearchStrategy = "None"
    }
    
    # Check known applications database first
    $cleanAppName = $ApplicationName -replace '\d+\.\d+[\d\.]*', ''
    $cleanAppName = $cleanAppName -replace '\(.*?\)', ''
    $cleanAppName = $cleanAppName.Trim().ToLower()
    
    # Direct match
    if ($KnownApps.ContainsKey($cleanAppName)) {
        $wingetID = $KnownApps[$cleanAppName]
        Write-SearchLog "Found in known apps database (direct): $wingetID"
        
        try {
            $detailOutput = winget show --id $wingetID --exact 2>&1 | Out-String
            $latestVersion = "Unknown"
            $packageName = $ApplicationName
            
            if ($detailOutput -match "Version:\s+(.+)") {
                $latestVersion = $matches[1].Trim()
            }
            if ($detailOutput -match "^Name:\s+(.+)" -and $detailOutput -match "Name:\s+(.+)") {
                $packageName = $matches[1].Trim()
            }
            
            $result.WingetID = $wingetID
            $result.MatchedName = $packageName
            $result.Version = $latestVersion
            $result.Status = "Found"
            $result.Confidence = "High"
            $result.SearchStrategy = "Known Database"
            
            $Cache[$ApplicationName] = $result
            return $result
        } catch {
            Write-SearchLog "Error verifying known app: $_"
        }
    }
    
    # Split camelCase/PascalCase words (MozillaFirefox -> Mozilla Firefox)
    $splitName = $ApplicationName -creplace '([a-z])([A-Z])', '$1 $2'
    
    # Remove version numbers and clean
    $cleanedName = $ApplicationName -replace '\d+\.\d+[\d\.]*', ''
    $cleanedName = $cleanedName -replace '\(.*?\)', ''
    $cleanedName = $cleanedName -replace '\[.*?\]', ''
    $cleanedName = $cleanedName.Trim()
    
    # High confidence strategies - use --exact flag
    $highConfidenceStrategies = @(
        @{ Name = "Exact Match"; Query = $ApplicationName }
        @{ Name = "Cleaned Name"; Query = $cleanedName }
        @{ Name = "Split CamelCase"; Query = $splitName }
        @{ Name = "Without Version"; Query = $cleanedName }
        @{ Name = "Hyphen to Space"; Query = ($ApplicationName -replace '-', ' ') }
        @{ Name = "No Spaces"; Query = ($ApplicationName -replace '\s+', '') }
    )
    
    # Add first two words strategy for multi-word apps (minimum 2 words required)
    $words = $cleanedName -split '\s+' | Where-Object { $_.Length -ge 2 }
    if ($words.Count -ge 2) {
        $firstTwo = "$($words[0]) $($words[1])"
        $highConfidenceStrategies += @{ Name = "First Two Words"; Query = $firstTwo }
    }
    
    # REMOVED: First Word Only - causes too many wrong matches
    # Single word searches produce generic results like "Microsoft" matching "Codium"
    
    foreach ($strategy in $highConfidenceStrategies) {
        $query = $strategy.Query
        
        if ([string]::IsNullOrWhiteSpace($query) -or $query.Length -lt 2) {
            continue
        }
        
        # For First Two Words, use non-exact search to allow fuzzy matching
        $useExact = $true
        if ($strategy.Name -eq "First Two Words") {
            $useExact = $false
        }
        
        Write-SearchLog "Trying strategy: $($strategy.Name) with query: $query (exact=$useExact)"
        $searchResult = Invoke-WingetSearch -Query $query -UseExact $useExact -CheckNameOnly $true
        
        if ($searchResult.Found) {
            $result.WingetID = $searchResult.WingetID
            $result.MatchedName = $searchResult.MatchedName
            $result.Version = $searchResult.Version
            $result.Status = "Found"
            $result.Confidence = "High"
            $result.SearchStrategy = $strategy.Name
            
            $Cache[$ApplicationName] = $result
            Write-SearchLog "Found via $($strategy.Name): $($result.WingetID)"
            return $result
        }
    }
    
    # Multi-word validation strategy - only accept if multiple searches agree
    $words = $cleanedName -split '\s+' | Where-Object { $_.Length -gt 2 }
    
    if ($words.Count -ge 2) {
        Write-SearchLog "Attempting multi-word validation with $($words.Count) words"
        
        $foundIds = @{}
        
        # Search each word individually
        foreach ($word in $words) {
            $wordResult = Invoke-WingetSearch -Query $word -UseExact $false -CheckNameOnly $true
            if ($wordResult.Found) {
                $id = $wordResult.WingetID
                if ($foundIds.ContainsKey($id)) {
                    $foundIds[$id]++
                } else {
                    $foundIds[$id] = 1
                }
            }
        }
        
        # Search combinations
        if ($words.Count -ge 2) {
            $firstTwo = $words[0..1] -join ' '
            $comboResult = Invoke-WingetSearch -Query $firstTwo -UseExact $false -CheckNameOnly $true
            if ($comboResult.Found) {
                $id = $comboResult.WingetID
                if ($foundIds.ContainsKey($id)) {
                    $foundIds[$id] += 2
                } else {
                    $foundIds[$id] = 2
                }
            }
        }
        
        # Only accept if found by at least 2 different searches
        $validIds = $foundIds.GetEnumerator() | Where-Object { $_.Value -ge 2 } | Sort-Object -Property Value -Descending
        
        if ($validIds.Count -gt 0) {
            $wingetID = $validIds[0].Name
            Write-SearchLog "Multi-word validation found: $wingetID (confidence: $($validIds[0].Value))"
            
            $verifyResult = Invoke-WingetSearch -Query $wingetID -UseExact $true -CheckNameOnly $true
            if ($verifyResult.Found) {
                $result.WingetID = $verifyResult.WingetID
                $result.MatchedName = $verifyResult.MatchedName
                $result.Version = $verifyResult.Version
                $result.Status = "Found"
                $result.Confidence = "Medium"
                $result.SearchStrategy = "Multi-word Validated"
                
                $Cache[$ApplicationName] = $result
                return $result
            }
        }
    }
    
    Write-SearchLog "No package found for: $ApplicationName"
    $Cache[$ApplicationName] = $result
    return $result
}

function Invoke-WingetSearch {
    param(
        [string]$Query,
        [bool]$UseExact = $true,
        [bool]$CheckNameOnly = $false
    )
    
    try {
        $exactFlag = if ($UseExact) { "--exact" } else { "" }
        $output = winget search $Query $exactFlag --accept-source-agreements 2>&1 | Out-String
        
        $lines = $output -split "`n"
        $foundHeader = $false
        
        foreach ($line in $lines) {
            if ($line -match "^Name\s+Id\s+Version") {
                $foundHeader = $true
                continue
            }
            
            if ($foundHeader -and $line.Trim() -ne "" -and $line -notmatch "^-+") {
                # Parse the line - Name and Id are separated by 2+ spaces
                # Winget ID must contain at least one dot and have Publisher.Product pattern
                if ($line -match "^(.+?)\s{2,}([\w-]+\.[\w\.-]+)\s+(.+)$") {
                    $packageName = $matches[1].Trim()
                    $packageId = $matches[2].Trim()
                    $packageVersion = $matches[3].Trim()
                    
                    # Validate Winget ID format - must have Publisher.Product pattern
                    # Reject if it looks like a version number (starts with digit or has only numbers/dots)
                    if ($packageId -match '^\d' -or $packageId -match '^[\d\.]+$') {
                        Write-SearchLog "Skipping invalid ID (looks like version): $packageId"
                        continue
                    }
                    
                    # Winget ID should have alphabetic characters before the first dot
                    if ($packageId -notmatch '^[a-zA-Z]') {
                        Write-SearchLog "Skipping invalid ID format: $packageId"
                        continue
                    }
                    
                    # If CheckNameOnly, verify the match is in the Name field, not just Tag/Moniker
                    if ($CheckNameOnly) {
                        $showOutput = winget show --id $packageId --exact 2>&1 | Out-String
                        
                        # Extract actual package name from show output
                        $actualName = ""
                        if ($showOutput -match "(?m)^Name:\s+(.+)$") {
                            $actualName = $matches[1].Trim()
                        }
                        
                        # Check if query matches the actual name (case insensitive, partial match OK)
                        $queryLower = $Query.ToLower()
                        $nameLower = $actualName.ToLower()
                        
                        if (-not ($nameLower -match [regex]::Escape($queryLower) -or $queryLower -match [regex]::Escape($nameLower))) {
                            Write-SearchLog "Skipping $packageId - matched only in Tag/Moniker, not in Name field"
                            continue
                        }
                    }
                    
                    return @{
                        Found = $true
                        WingetID = $packageId
                        MatchedName = $packageName
                        Version = $packageVersion
                    }
                }
            }
        }
        
        return @{ Found = $false }
    }
    catch {
        Write-SearchLog "Error in winget search: $_"
        return @{ Found = $false }
    }
}

function Search-WingetFuzzy {
    param([string]$ApplicationName)
    
    try {
        # Try a broader search without --exact
        $output = winget search $ApplicationName --accept-source-agreements 2>&1 | Out-String
        
        # Parse and get the first result
        $lines = $output -split "`n"
        $foundHeader = $false
        
        foreach ($line in $lines) {
            if ($line -match "^Name\s+Id\s+Version") {
                $foundHeader = $true
                continue
            }
            
            if ($foundHeader -and $line.Trim() -ne "" -and $line -notmatch "^-+") {
                # Parse the line
                if ($line -match "^(.+?)\s{2,}([\w\.-]+\.\w+[\w\.-]*)\s+(.+)$") {
                    $matchedName = $matches[1].Trim()
                    
                    # Calculate similarity score
                    $similarity = Get-StringSimilarity -String1 $ApplicationName.ToLower() -String2 $matchedName.ToLower()
                    
                    # Only return if similarity is above threshold (50%)
                    if ($similarity -ge 50) {
                        return @{
                            Found = $true
                            WingetID = $matches[2].Trim()
                            MatchedName = $matchedName
                            Version = $matches[3].Trim()
                        }
                    }
                }
            }
        }
        
        return @{ Found = $false }
    }
    catch {
        Write-SearchLog "Error in fuzzy search: $_"
        return @{ Found = $false }
    }
}

function Get-CommonVariations {
    param([string]$ApplicationName)
    
    $variations = @()
    
    # Remove common suffixes
    $suffixes = @(
        " \(x64\)", " \(x86\)", " \(64-bit\)", " \(32-bit\)",
        " x64", " x86", " 64-bit", " 32-bit",
        " Desktop", " Professional", " Enterprise", " Standard",
        " Free", " Trial", " Lite"
    )
    
    foreach ($suffix in $suffixes) {
        if ($ApplicationName -match [regex]::Escape($suffix)) {
            $variations += $ApplicationName -replace [regex]::Escape($suffix), ''
        }
    }
    
    # Try with publisher name (common patterns)
    if ($ApplicationName -match "^(Microsoft|Adobe|Google|Oracle|Mozilla|Apple|Autodesk)\s+(.+)") {
        $variations += $matches[2]  # Without publisher
        $variations += "$($matches[1]).$($matches[2])"  # With dot separator
    }
    
    # Try reversing order for Publisher Product format
    if ($ApplicationName -match "^(\w+)\s+(\w+)$") {
        $variations += "$($matches[2]) $($matches[1])"
    }
    
    # Try with hyphens instead of spaces
    $variations += $ApplicationName -replace '\s+', '-'
    
    # Remove version numbers
    $variations += $ApplicationName -replace '\d+\.\d+.*', ''
    
    # Remove everything after first digit
    if ($ApplicationName -match '^([^\d]+)') {
        $variations += $matches[1].Trim()
    }
    
    return $variations | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique
}

function Get-StringSimilarity {
    param(
        [string]$String1,
        [string]$String2
    )
    
    if ($String1 -eq $String2) {
        return 100
    }
    
    # Levenshtein distance calculation - simple implementation
    $len1 = $String1.Length
    $len2 = $String2.Length
    
    if ($len1 -eq 0) { return 0 }
    if ($len2 -eq 0) { return 0 }
    
    # Create array using alternative syntax
    $d = @{}
    
    for ($i = 0; $i -le $len1; $i++) {
        $d["$i,0"] = $i
    }
    
    for ($j = 0; $j -le $len2; $j++) {
        $d["0,$j"] = $j
    }
    
    for ($i = 1; $i -le $len1; $i++) {
        for ($j = 1; $j -le $len2; $j++) {
            $cost = if ($String1[$i - 1] -eq $String2[$j - 1]) { 0 } else { 1 }
            
            $deletion = $d["$($i-1),$j"] + 1
            $insertion = $d["$i,$($j-1)"] + 1
            $substitution = $d["$($i-1),$($j-1)"] + $cost
            
            $d["$i,$j"] = [Math]::Min([Math]::Min($deletion, $insertion), $substitution)
        }
    }
    
    $distance = $d["$len1,$len2"]
    $maxLen = [Math]::Max($len1, $len2)
    
    $similarity = [Math]::Round(((1 - ($distance / $maxLen)) * 100), 2)
    
    return $similarity
}

Export-ModuleMember -Function Search-WingetIntelligent, Invoke-WingetSearch, Search-WingetFuzzy, Get-CommonVariations, Get-StringSimilarity, Write-SearchLog
