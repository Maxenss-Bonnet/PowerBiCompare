# Power BI Report Analysis and Comparison Script (PBIR format)

#==========================================================================================================================================
# FUNCTIONS
#==========================================================================================================================================

Function LoadReportProjectVersions {
    param (
        [string] $newVersionProjectDirectory,
        [string] $oldVersionProjectDirectory
    )

    Write-Host "=== Chargement des projets ===" -ForegroundColor Green
    Write-Host "Nouveau projet: $newVersionProjectDirectory" -ForegroundColor Cyan
    Write-Host "Ancien projet: $oldVersionProjectDirectory" -ForegroundColor Cyan

    $returnArray = @()
    $projectArray = @($newVersionProjectDirectory, $oldVersionProjectDirectory)

    foreach ($projectPath in $projectArray) {
        Write-Host "Traitement de: $projectPath" -ForegroundColor Yellow
        
        if (-not (Test-Path $projectPath)) {
            throw "Le dossier '$projectPath' n'existe pas"
        }

        $pbipFiles = Get-ChildItem -Path $projectPath -Filter "*.pbip"
        
        if ($pbipFiles.Count -eq 0) {
            throw "Aucun fichier .pbip trouve dans '$projectPath'"
        }
        
        if ($pbipFiles.Count -gt 1) {
            throw "Plusieurs fichiers .pbip trouves dans '$projectPath'"
        }

        $pbipFile = $pbipFiles[0]
        $reportProject = [ReportProject]::new()
        
        $reportProject.name = $pbipFile.BaseName
        $reportProject.nameAndTimestamp = $pbipFile.BaseName + '__' + $pbipFile.LastWriteTimeUtc.ToString("yyyy-MM-dd_HH-mm-ss")
        $reportProject.path = $projectPath
        $reportProject.reportPath = Join-Path $projectPath "$($pbipFile.BaseName).Report\definition"

        Write-Host "  - Nom du projet: $($reportProject.name)" -ForegroundColor White
        Write-Host "  - Chemin Report: $($reportProject.reportPath)" -ForegroundColor White

        if (-not (Test-Path $reportProject.reportPath)) {
            Write-Warning "Le dossier Report n'existe pas dans '$projectPath'. Creation d'un projet vide."
            $reportProject.reportData = $null
            $reportProject.pagesData = @{}
            $reportProject.bookmarksData = @{}
            $reportProject.bookmarkMetadata = $null
        }
        else {
            try {
                $reportProject.reportData = LoadReportData -reportPath $reportProject.reportPath
                $reportProject.pagesData = LoadPagesData -reportPath $reportProject.reportPath
                $bookmarkLoadResult = LoadBookmarksData -reportPath $reportProject.reportPath
                $reportProject.bookmarksData = if ($bookmarkLoadResult -and $bookmarkLoadResult.PSObject.Properties['Bookmarks']) {
                    $bookmarkLoadResult.Bookmarks
                } else { @{} }
                $reportProject.bookmarkMetadata = if ($bookmarkLoadResult -and $bookmarkLoadResult.PSObject.Properties['Metadata']) {
                    $bookmarkLoadResult.Metadata
                } else { $null }
                
                Write-Host "  - Pages chargees: $($reportProject.pagesData.Count)" -ForegroundColor Green
                Write-Host "  - Signets charges: $($reportProject.bookmarksData.Count)" -ForegroundColor Green
                if ($reportProject.bookmarkMetadata -and $reportProject.bookmarkMetadata.PSObject.Properties['items']) {
                    $groupCount = ($reportProject.bookmarkMetadata.items | Measure-Object).Count
                    Write-Host "  - Groupes de signets: $groupCount" -ForegroundColor Green
                }
            }
            catch {
                Write-Warning "Erreur lors du chargement des donnees: $($_.Exception.Message)"
                $reportProject.reportData = $null
                $reportProject.pagesData = @{}
                $reportProject.bookmarksData = @{}
                $reportProject.bookmarkMetadata = $null
            }
        }

        $returnArray += $reportProject
    }

    return $returnArray
}

# Secure loading function for main report data
Function LoadReportData {
    param([string] $reportPath)
    
    $reportJsonPath = Join-Path $reportPath "report.json"
    
    if (Test-Path $reportJsonPath) {
        try {
            $content = Get-Content $reportJsonPath -Raw -Encoding UTF8
            if ([string]::IsNullOrWhiteSpace($content)) {
                Write-Warning "Fichier report.json vide"
                return $null
            }
            return $content | ConvertFrom-Json
        }
        catch {
            Write-Warning "Erreur lors de la lecture de report.json: $($_.Exception.Message)"
            return $null
        }
    }
    else {
        Write-Warning "Fichier report.json non trouve dans $reportPath"
        return $null
    }
}

# Enhanced secure loading function for page data
Function LoadPagesData {
    param([string] $reportPath)
    
    $pagesPath = Join-Path $reportPath "pages"
    $pagesData = @{}
    
    if (-not (Test-Path $pagesPath)) {
        Write-Host "    Aucun dossier pages trouve" -ForegroundColor Gray
        return $pagesData
    }

    # Loading data for each page
    $pageDirectories = Get-ChildItem -Path $pagesPath -Directory -ErrorAction SilentlyContinue
    
    foreach ($pageDir in $pageDirectories) {
        $pageName = $pageDir.Name
        $pageJsonPath = Join-Path $pageDir.FullName "page.json"
        
        if (Test-Path $pageJsonPath) {
            try {
                $content = Get-Content $pageJsonPath -Raw -Encoding UTF8
                if (-not [string]::IsNullOrWhiteSpace(
                    $content)) {
                    $pageData = $content | ConvertFrom-Json
                    
                    # Extract readable information
                    $displayName = if ($pageData.displayName) { $pageData.displayName } else { "(No name)" }
                    Write-Host "    Chargement: '$displayName' ($pageName)" -ForegroundColor Gray
                    
                    
                    # Load visuals for this page
                    $visualsPath = Join-Path $pageDir.FullName "visuals"
                    $pageData | Add-Member -NotePropertyName "visuals" -NotePropertyValue @{} -Force
                    
                    if (Test-Path $visualsPath) {
                        $visualDirectories = Get-ChildItem -Path $visualsPath -Directory -ErrorAction SilentlyContinue
                        
                        foreach ($visualDir in $visualDirectories) {
                            $visualName = $visualDir.Name
                            $visualJsonPath = Join-Path $visualDir.FullName "visual.json"
                            
                            if (Test-Path $visualJsonPath) {
                                try {
                                    $visualContent = Get-Content $visualJsonPath -Raw -Encoding UTF8
                                    if (-not [string]::IsNullOrWhiteSpace($visualContent)) {
                                        $visualData = $visualContent | ConvertFrom-Json
                                        
                                        # Extract readable visual information
                                        $visualType = if ($visualData.visual -and $visualData.visual.visualType) { 
                                            $visualData.visual.visualType 
                                        } else { "unknown" }
                                        
                                        # Extract main fields used (rows/columns/measures...)
                                        $mainFields = New-Object System.Collections.Generic.List[string]
                                        if ($visualData.visual -and $visualData.visual.query -and $visualData.visual.query.queryState) {
                                            $queryState = $visualData.visual.query.queryState

                                            $addField = {
                                                param($projection)

                                                if ($null -eq $projection) { return }

                                                $candidate = $null

                                                if ($projection.PSObject.Properties['displayName'] -and $projection.displayName) {
                                                    $candidate = [string]$projection.displayName
                                                }
                                                elseif ($projection.field -and $projection.field.Column -and $projection.field.Column.Property) {
                                                    $candidate = [string]$projection.field.Column.Property
                                                }
                                                elseif ($projection.field -and $projection.field.Measure -and $projection.field.Measure.Property) {
                                                    $candidate = [string]$projection.field.Measure.Property
                                                }
                                                elseif ($projection.field -and $projection.field.Aggregation -and $projection.field.Aggregation.Expression -and $projection.field.Aggregation.Expression.SourceRef -and $projection.field.Aggregation.Expression.SourceRef.Property) {
                                                    $candidate = [string]$projection.field.Aggregation.Expression.SourceRef.Property
                                                }
                                                elseif ($projection.PSObject.Properties['queryRef'] -and $projection.queryRef) {
                                                    $candidate = [string]$projection.queryRef
                                                }
                                                elseif ($projection.PSObject.Properties['nativeQueryRef'] -and $projection.nativeQueryRef) {
                                                    $candidate = [string]$projection.nativeQueryRef
                                                }

                                                if (-not [string]::IsNullOrWhiteSpace($candidate)) {
                                                    $normalized = $candidate.Trim()
                                                    if (-not $mainFields.Contains($normalized)) {
                                                        $mainFields.Add($normalized) | Out-Null
                                                    }
                                                }
                                            }

                                            $addFromContainer = {
                                                param($container)

                                                if ($null -eq $container) {
                                                    return
                                                }

                                                $projections = $null

                                                if ($container.PSObject.Properties['projections']) {
                                                    $projections = $container.projections
                                                }
                                                elseif ($container -is [System.Collections.IEnumerable] -and -not ($container -is [string]) -and -not ($container -is [System.Collections.IDictionary])) {
                                                    $projections = $container
                                                }

                                                if ($projections) {
                                                    foreach ($proj in $projections) {
                                                        & $addField $proj
                                                    }
                                                }
                                            }

                                            $projectionKeys = @(
                                                'Values', 'Rows', 'Columns', 'Measures', 'PrimaryValues', 'SecondaryValues',
                                                'Series', 'Categories', 'Category', 'Legends', 'Legend', 'XAxis', 'YAxis',
                                                'Tooltips', 'Breakdowns', 'Data'
                                            )

                                            foreach ($key in $projectionKeys) {
                                                if ($queryState.PSObject.Properties[$key]) {
                                                    $containerValue = $queryState.$key
                                                    & $addFromContainer $containerValue
                                                }
                                            }
                                        }
                                        $mainFields = $mainFields.ToArray()
                                        
                                        # ENHANCED: Extract special button information
                                        $buttonDetails = @()
                                        if ($visualType -eq "actionButton") {
                                            $buttonInfo = ExtractButtonDetails -visual $visualData
                                            $buttonDetails = $buttonInfo.ButtonDetails
                                        }
                                        
                                        $visualData | Add-Member -NotePropertyName "_extracted_type" -NotePropertyValue $visualType -Force
                                        $visualData | Add-Member -NotePropertyName "_extracted_fields" -NotePropertyValue $mainFields -Force
                                        $visualData | Add-Member -NotePropertyName "_extracted_button_details" -NotePropertyValue $buttonDetails -Force
                                        
                                        $fieldsText = if ($mainFields.Count -gt 0) { " [" + ($mainFields -join ", ") + "]" } else { "" }
                                        $buttonText = if ($visualType -eq "actionButton" -and $buttonDetails.Count -gt 0) { " - BOUTON: " + ($buttonDetails -join " | ") } else { "" }
                                        
                                        if ($visualType -eq "actionButton") {
                                            Write-Host "      Visuel: $visualType$fieldsText ($visualName)$buttonText" -ForegroundColor Cyan
                                        } else {
                                            Write-Host "      Visuel: $visualType$fieldsText ($visualName)" -ForegroundColor DarkGray
                                        }
                                        
                                        $pageData.visuals[$visualName] = $visualData
                                    }
                                }
                                catch {
                                    Write-Warning "      Erreur lors du chargement du visuel $visualName : $($_.Exception.Message)"
                                }
                            }
                        }
                    }
                    
                    $pagesData[$pageName] = $pageData
                    Write-Host "      Page '$displayName' chargee avec $($pageData.visuals.Count) visuels" -ForegroundColor Green
                }
            }
            catch {
                Write-Warning "    Erreur lors du chargement de la page $pageName : $($_.Exception.Message)"
            }
        }
    }
    
    return $pagesData
}

# Secure loading function for bookmark data
Function LoadBookmarksData {
    param([string] $reportPath)
    
    $bookmarksPath = Join-Path $reportPath "bookmarks"
    $bookmarksData = @{}
    $metadata = $null
    
    if (-not (Test-Path $bookmarksPath)) {
        Write-Host "    Aucun dossier bookmarks trouve" -ForegroundColor Gray
        return [pscustomobject]@{
            Bookmarks = $bookmarksData
            Metadata  = $null
        }
    }

    # Loading individual bookmark files
    $bookmarkFiles = Get-ChildItem -Path $bookmarksPath -Filter "*.bookmark.json" -ErrorAction SilentlyContinue
    
    foreach ($bookmarkFile in $bookmarkFiles) {
        $bookmarkName = $bookmarkFile.BaseName -replace '\.bookmark$', ''
        Write-Host "    Chargement du signet: $bookmarkName" -ForegroundColor Gray
        
        try {
            $content = Get-Content $bookmarkFile.FullName -Raw -Encoding UTF8
            if (-not [string]::IsNullOrWhiteSpace($content)) {
                $bookmarkData = $content | ConvertFrom-Json
                $bookmarksData[$bookmarkName] = $bookmarkData
                Write-Host "      Signet $bookmarkName charge" -ForegroundColor Green
            }
        }
        catch {
            Write-Warning "    Erreur lors du chargement du signet $bookmarkName : $($_.Exception.Message)"
        }
    }
    
    $metadataPath = Join-Path $bookmarksPath "bookmarks.json"
    if (Test-Path $metadataPath) {
        try {
            $metadataContent = Get-Content $metadataPath -Raw -Encoding UTF8
            if (-not [string]::IsNullOrWhiteSpace($metadataContent)) {
                $metadata = $metadataContent | ConvertFrom-Json
                if ($metadata.PSObject.Properties['items']) {
                    $itemCount = ($metadata.items | Measure-Object).Count
                    Write-Host "      Metadonnees de signets chargees ($itemCount groupes)" -ForegroundColor Green
                }
            }
        }
        catch {
            Write-Warning "    Erreur lors du chargement de bookmarks.json : $($_.Exception.Message)"
            $metadata = $null
        }
    }
    
    return [pscustomobject]@{
        Bookmarks = $bookmarksData
        Metadata  = $metadata
    }
}


Function Get-PageDisplayIndex {
    param([hashtable] $pagesData)

    $index = @{}

    if ($null -eq $pagesData) {
        return $index
    }

    foreach ($pageKey in $pagesData.Keys) {
        $pageData = $pagesData[$pageKey]
        if ($null -eq $pageData) {
            continue
        }

        $displayName = if ($pageData.displayName) { [string]$pageData.displayName } else { [string]$pageKey }

        if (-not $index.ContainsKey($displayName)) {
            $index[$displayName] = @()
        }

        $index[$displayName] += $pageKey
    }

    return $index
}

Function Test-CollectionHasKey {
    param(
        $collection,
        [string] $key
    )

    if ($null -eq $collection) {
        return $false
    }

    if ($collection -is [hashtable]) {
        return $collection.ContainsKey($key)
    }

    if ($collection -is [System.Collections.Specialized.OrderedDictionary]) {
        return $collection.Contains($key)
    }

    return $collection.PSObject.Properties.Name -contains $key
}

Function Test-VisualIsVisible {
    param(
        [PSCustomObject] $pageData,
        [PSCustomObject] $visualData
    )

    if ($null -eq $visualData) {
        return $false
    }

    if ($visualData.PSObject.Properties['isHidden'] -and $visualData.isHidden -eq $true) {
        return $false
    }

    $groupId = if ($visualData.PSObject.Properties['parentGroupName']) { [string]$visualData.parentGroupName } else { $null }
    $visited = [System.Collections.Generic.HashSet[string]]::new()
    # Removed hardcoded allowedHiddenGroups - now using pattern matching for groups ending with _filter, _period, _navigation

    while ($groupId) {
        if (-not $visited.Add($groupId)) {
            break
        }

        if (-not (Test-CollectionHasKey -collection $pageData.visuals -key $groupId)) {
            break
        }

        $groupData = $pageData.visuals[$groupId]

        if ($null -eq $groupData) {
            break
        }

        $groupDisplay = $null
        if ($groupData.PSObject.Properties['visualGroup'] -and $groupData.visualGroup.PSObject.Properties['displayName']) {
            $groupDisplay = [string]$groupData.visualGroup.displayName
        }

        $normalizedGroup = $null
        if ($groupDisplay) {
            $normalizedGroup = $groupDisplay.Trim().Trim("'")
            if ($normalizedGroup) {
                $normalizedGroup = $normalizedGroup.ToLowerInvariant()
            }
        }

        # Check if group matches patterns for allowed hidden groups (suffix-based matching)
        $isAllowedHidden = $false
        if ($normalizedGroup) {
            if ($normalizedGroup -like '*_filter' -or $normalizedGroup -like '*_period' -or $normalizedGroup -like '*_navigation') {
                $isAllowedHidden = $true
            }
        }

        if ($groupData.PSObject.Properties['isHidden'] -and $groupData.isHidden -eq $true -and -not $isAllowedHidden) {
            return $false
        }

        if ($groupDisplay) {
            if ($groupDisplay -match '(?i)filter\s+pane') {
                return $false
            }
        }

        if ($groupData.PSObject.Properties['parentGroupName']) {
            $groupId = [string]$groupData.parentGroupName
        }
        else {
            $groupId = $null
        }
    }

    return $true
}

Function Get-VisibleSlicerQualityResult {
    param(
        [PSCustomObject[]] $checkResults,
        [ReportProject] $newProject
    )

    if ($null -eq $checkResults -or $checkResults.Count -eq 0) {
        return $checkResults
    }

    if ($null -eq $newProject -or $null -eq $newProject.pagesData) {
        return $checkResults
    }

    $pageIndex = Get-PageDisplayIndex -pagesData $newProject.pagesData
    $visibleResults = @()

    foreach ($result in $checkResults) {
        if ($null -eq $result -or [string]::IsNullOrWhiteSpace($result.VisualId)) {
            continue
        }

        $pageCandidates = $pageIndex[$result.PageName]

        if (-not $pageCandidates) {
            $visibleResults += $result
            continue
        }

        $isVisible = $false

        foreach ($pageId in $pageCandidates) {
            $pageData = $newProject.pagesData[$pageId]

            if ($null -eq $pageData -or $null -eq $pageData.visuals) {
                continue
            }

            if (-not (Test-CollectionHasKey -collection $pageData.visuals -key $result.VisualId)) {
                continue
            }

            $visualData = $pageData.visuals[$result.VisualId]
            if ($null -eq $visualData) {
                continue
            }

            $visualType = $null
            if ($visualData.PSObject.Properties['visual']) {
                $visualNode = $visualData.visual
                if ($visualNode -and $visualNode.PSObject.Properties['visualType']) {
                    $visualType = [string]$visualNode.visualType
                }
            }

            if ($visualType -and $visualType -ne 'slicer') {
                continue
            }

            if (Test-VisualIsVisible -pageData $pageData -visualData $visualData) {
                $isVisible = $true
                break
            }
        }

        if ($isVisible) {
            $visibleResults += $result
        }
    }

    return $visibleResults
}

Function Get-VisualBooleanValue {
    param($expression)

    if ($null -eq $expression) {
        return $null
    }

    try {
        $rawValue = $null

        if (Get-Command -Name 'Get-LiteralStringValue' -CommandType Function -ErrorAction SilentlyContinue) {
            $rawValue = Get-LiteralStringValue -value $expression
        }

        if ($null -eq $rawValue -or $rawValue -eq '') {
            if ($expression -is [bool]) {
                return [bool]$expression
            }

            if ($expression.PSObject.Properties['Value']) {
                $rawValue = $expression.Value
            }
            elseif ($expression.PSObject.Properties['Literal']) {
                return Get-VisualBooleanValue -expression $expression.Literal
            }
            elseif ($expression.PSObject.Properties['expr']) {
                return Get-VisualBooleanValue -expression $expression.expr
            }
        }

        if ($null -ne $rawValue -and $rawValue -ne '') {
            $normalized = $rawValue.ToString().Trim("'").Trim()
            if ($normalized -match '^(?i:true|1)$') {
                return $true
            }
            if ($normalized -match '^(?i:false|0)$') {
                return $false
            }
        }
    }
    catch {
        Write-Debug "Echec conversion booleenne: $($_.Exception.Message)"
    }

    return $null
}

Function Test-VisualSingleSelectMode {
    param([PSCustomObject] $visual)

    try {
        if ($null -eq $visual -or -not ($visual.PSObject.Properties['visual'])) {
            return $false
        }

        $objects = $visual.visual.objects
        if ($null -eq $objects) {
            return $false
        }

        $collections = @()

        if ($objects.PSObject.Properties['selection']) {
            $collections += @($objects.selection)
        }

        if ($objects.PSObject.Properties['general']) {
            $collections += @($objects.general)
        }

        foreach ($collection in $collections) {
            if ($null -eq $collection) {
                continue
            }

            foreach ($item in @($collection)) {
                if ($null -eq $item) {
                    continue
                }

                $propertyBag = if ($item.PSObject.Properties['properties']) { $item.properties } else { $item }
                if ($null -eq $propertyBag) {
                    continue
                }

                foreach ($propName in @('strictSingleSelect', 'singleSelect')) {
                    if ($propertyBag.PSObject.Properties[$propName]) {
                        $value = Get-VisualBooleanValue -expression $propertyBag.$propName
                        if ($value -eq $true) {
                            return $true
                        }
                    }
                }
            }
        }

        return $false
    }
    catch {
        Write-Warning "Erreur lors de la detection du mode single select: $($_.Exception.Message)"
        return $false
    }
}

Function Resolve-SlicerDisplayName {
    param(
        [PSCustomObject] $visual,
        [string] $fallbackField,
        [string] $fallbackGroup,
        [string] $visualId
    )

    $candidates = New-Object System.Collections.Generic.List[string]

    try {
        if (Get-Command -Name 'Get-SlicerDisplayName' -CommandType Function -ErrorAction SilentlyContinue) {
            $baseName = Get-SlicerDisplayName -visual $visual -fallbackField $fallbackField -fallbackGroup $fallbackGroup
            if (-not [string]::IsNullOrWhiteSpace($baseName)) {
                $candidates.Add($baseName)
            }
        }

        if ($visual.visual -and $visual.visual.query -and $visual.visual.query.queryState) {
            $queryState = $visual.visual.query.queryState
            if ($queryState.Values -and $queryState.Values.projections) {
                foreach ($projection in $queryState.Values.projections) {
                    if ($projection.PSObject.Properties['displayName']) {
                        $candidates.Add([string]$projection.displayName)
                    }
                    elseif ($projection.field -and $projection.field.Column -and $projection.field.Column.Property) {
                        $candidates.Add([string]$projection.field.Column.Property)
                    }
                }
            }
        }

        if ($visual.visual -and $visual.visual.objects -and $visual.visual.objects.header) {
            foreach ($headerItem in $visual.visual.objects.header) {
                if ($headerItem -and $headerItem.properties -and $headerItem.properties.text) {
                    $headerText = $null
                    if (Get-Command -Name 'Get-LiteralStringValue' -CommandType Function -ErrorAction SilentlyContinue) {
                        $headerText = Get-LiteralStringValue -value $headerItem.properties.text
                    }

                    if ($headerText) {
                        $candidates.Add($headerText)
                    }
                }
            }
        }
    }
    catch {
        Write-Warning "Erreur lors de la resolution du nom du slicer: $($_.Exception.Message)"
    }

    $candidates = $candidates | Where-Object { $_ -and $_.Trim("'").Trim().Length -gt 0 } | Select-Object -Unique

    $selectedName = $null

    $preferred = $candidates | Where-Object {
        $valueString = ([string]$_)
        $cleanCandidate = $valueString.Trim("'").Trim()
        -not (Test-SlicerNameIsTechnical -value $cleanCandidate -fallbackField $fallbackField -fallbackGroup $fallbackGroup -visualId $visualId)
    }

    if ($preferred.Count -gt 0) {
        $selectedName = ([string]$preferred[0]).Trim("'").Trim()
    }

    if (-not $selectedName) {
        foreach ($candidate in $candidates) {
            $clean = ([string]$candidate).Trim("'").Trim()
            if (-not [string]::IsNullOrWhiteSpace($clean)) {
                $selectedName = $clean
                break
            }
        }
    }

    if ([string]::IsNullOrWhiteSpace($selectedName)) {
        $selectedName = $fallbackGroup
    }

    if ([string]::IsNullOrWhiteSpace($selectedName)) {
        $selectedName = $fallbackField
    }

    if ([string]::IsNullOrWhiteSpace($selectedName)) {
        $selectedName = $visualId
    }

    if ([string]::IsNullOrWhiteSpace($selectedName)) {
        $selectedName = 'Slicer'
    }

    $selectedName = ([string]$selectedName).Trim()

    if (Test-SlicerNameIsTechnical -value $selectedName -fallbackField $fallbackField -fallbackGroup $fallbackGroup -visualId $visualId) {
        if (-not [string]::IsNullOrWhiteSpace($fallbackField)) {
            return ([string]$fallbackField).Trim()
        }

        if (-not [string]::IsNullOrWhiteSpace($fallbackGroup)) {
            return ([string]$fallbackGroup).Trim()
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($fallbackField)) {
        if ($selectedName.Equals($fallbackField, [System.StringComparison]::InvariantCultureIgnoreCase) -and $selectedName -cne $fallbackField) {
            return ([string]$fallbackField).Trim()
        }
    }

    return $selectedName
}

Function Test-SlicerNameIsTechnical {
    param(
        [string] $value,
        [string] $fallbackField,
        [string] $fallbackGroup,
        [string] $visualId
    )

    if ([string]::IsNullOrWhiteSpace($value)) {
        return $true
    }

    $normalized = ([string]$value).Trim()

    if ($normalized.Length -eq 0) {
        return $true
    }

    if ($normalized.Length -le 1) {
        return $true
    }

    if ($fallbackField -and $normalized -eq $fallbackField) {
        return $false
    }

    if ($normalized -match '(?i)^slc\s+') {
        return $true
    }

    if ($normalized -match '(?i)^[0-9a-f]{16,}$') {
        return $true
    }

    if ($visualId -and $normalized -eq $visualId) {
        return $true
    }

    if ($normalized -match '^Slicer\s*$') {
        return $true
    }

    if ($fallbackGroup -and $normalized -eq $fallbackGroup) {
        return $false
    }

    return $false
}

Function Invoke-VisibleSlicerQualityCheck {
    param(
        [ReportProject] $project,
        [string] $projectRoot
    )

    Write-Host "=== Analyse qualite des slicers visibles ===" -ForegroundColor Green
    $results = @()
    $totalCandidates = 0
    $hiddenCount = 0

    try {
        if ($null -eq $project) {
            throw "Aucun projet charge pour l'analyse des slicers."
        }

        $reportPath = $project.reportPath
        if ([string]::IsNullOrWhiteSpace($reportPath) -or -not (Test-Path $reportPath)) {
            throw "Chemin du rapport introuvable pour le projet '$($project.name)'."
        }

        if (Get-Command -Name 'Invoke-FilterPaneQualityCheck' -CommandType Function -ErrorAction SilentlyContinue) {
            $paneResults = Invoke-FilterPaneQualityCheck -reportPath $reportPath
            $results += $paneResults
        }

        $pagesData = $project.pagesData
        if ($null -eq $pagesData -or $pagesData.Count -eq 0) {
            Write-Host "Pages non prechargees, chargement depuis le disque..." -ForegroundColor Yellow
            $pagesData = LoadPagesDataForCheck -reportPath $reportPath
        }

        Write-Host "Pages chargees: $($pagesData.Count)" -ForegroundColor Green

        foreach ($pageKey in $pagesData.Keys) {
            $pageData = $pagesData[$pageKey]
            if ($null -eq $pageData) {
                continue
            }

            $pageDisplayName = if ($pageData.displayName) { [string]$pageData.displayName } else { [string]$pageKey }
            Write-Host "  Analyse de la page: '$pageDisplayName' ($pageKey)" -ForegroundColor Yellow

            if (-not $pageData.PSObject.Properties['visuals'] -or $null -eq $pageData.visuals) {
                continue
            }

            foreach ($visualKey in $pageData.visuals.Keys) {
                $visualData = $pageData.visuals[$visualKey]
                if ($null -eq $visualData) {
                    continue
                }

                $visualType = $null
                if ($visualData.PSObject.Properties['_extracted_type']) {
                    $visualType = [string]$visualData._extracted_type
                }
                elseif ($visualData.PSObject.Properties['visual'] -and $visualData.visual.PSObject.Properties['visualType']) {
                    $visualType = [string]$visualData.visual.visualType
                }

                if ($visualType -ne 'slicer') {
                    continue
                }

                $totalCandidates++

                if (-not (Test-VisualIsVisible -pageData $pageData -visualData $visualData)) {
                    $hiddenCount++
                    continue
                }

                $groupName = if (Get-Command -Name 'Get-SlicerGroupName' -CommandType Function -ErrorAction SilentlyContinue) {
                    Get-SlicerGroupName -visual $visualData -pageData $pageData
                } else { $null }

                $fieldName = if (Get-Command -Name 'Get-SlicerFieldName' -CommandType Function -ErrorAction SilentlyContinue) {
                    Get-SlicerFieldName -visual $visualData
                } else { $null }

                $displayName = Resolve-SlicerDisplayName -visual $visualData -fallbackField $fieldName -fallbackGroup $groupName -visualId $visualKey

                $searchInfo = if (Get-Command -Name 'Get-SlicerSearchInfo' -CommandType Function -ErrorAction SilentlyContinue) {
                    Get-SlicerSearchInfo -visual $visualData
                } else { [PSCustomObject]@{ HasSearch = $false; SearchText = $null } }

                Write-Host "    Slicer visible: $displayName ($visualKey)" -ForegroundColor Gray

                $hasSelected = if (Get-Command -Name 'Test-SlicerHasSelection' -CommandType Function -ErrorAction SilentlyContinue) {
                    Test-SlicerHasSelection -visual $visualData
                } else { $false }

                $isSingleSelect = Test-VisualSingleSelectMode -visual $visualData

                $hasSearch = if ($searchInfo -and $searchInfo.PSObject.Properties['HasSearch']) { [bool]$searchInfo.HasSearch } else { $false }
                $searchText = if ($searchInfo -and $searchInfo.PSObject.Properties['SearchText']) { [string]$searchInfo.SearchText } else { $null }
                $searchDetail = if (-not [string]::IsNullOrWhiteSpace($searchText)) { " ('{0}')" -f $searchText } else { '' }

                $status = 'OK'
                $messageKey = 'ok'
                $message = 'Le filtre est correctement configure.'

                $normalizedGroup = if ($groupName) { $groupName.ToString().Trim().ToLowerInvariant() } else { '' }
                $isMenuFilter = $normalizedGroup -like '*_filter'
                $isMenuPeriod = $normalizedGroup -like '*_period'

                if ($hasSearch) {
                    if ($isSingleSelect) {
                        $status = 'ALERTE'
                        $messageKey = 'radio_search'
                        $message = "Le filtre en mode radio contient du texte dans la loupe$searchDetail (non autorise)."
                    }
                    elseif ($hasSelected) {
                        $status = 'ALERTE'
                        $messageKey = 'selection_search'
                        $message = "Le filtre contient une selection et du texte dans la loupe$searchDetail."
                    }
                    else {
                        $status = 'ALERTE'
                        $messageKey = 'search'
                        $message = "Le filtre contient du texte dans la loupe$searchDetail."
                    }
                }
                elseif ($isMenuFilter) {
                    if ($hasSelected) {
                        $status = 'ALERTE'
                        $messageKey = 'menu_selected'
                        $message = "Ce filtre du groupe $groupName (terminant par _filter) ne doit pas contenir de selection."
                    }
                    else {
                        $status = 'OK'
                        $messageKey = 'menu_ok'
                        $message = "Ce filtre du groupe $groupName (terminant par _filter) est vide comme attendu."
                    }
                }
                elseif ($isMenuPeriod) {
                    # Groups ending with _period: ONLY "Current Year" allowed
                    $isCurrentYearOnly = if (Get-Command -Name 'Test-SlicerIsCurrentYearOnly' -ErrorAction SilentlyContinue) {
                        Test-SlicerIsCurrentYearOnly -visual $visualData
                    } else { $false }

                    if ($isCurrentYearOnly) {
                        $status = 'OK'
                        $messageKey = 'menu_period_ok'
                        $message = "Ce filtre du groupe $groupName (terminant par _period) a uniquement 'Current Year' comme attendu."
                    }
                    else {
                        # Extract current values for detailed error message
                        $currentValues = if (Get-Command -Name 'Get-SlicerSelectedValues' -ErrorAction SilentlyContinue) {
                            Get-SlicerSelectedValues -visual $visualData
                        } else { @() }

                        if ($currentValues.Count -eq 0) {
                            $status = 'ALERTE'
                            $messageKey = 'menu_period_empty'
                            $message = "Ce filtre du groupe $groupName (terminant par _period) doit contenir uniquement 'Current Year' (actuellement vide)."
                        }
                        elseif ($currentValues.Count -gt 1) {
                            $valuesList = $currentValues -join ', '
                            $status = 'ALERTE'
                            $messageKey = 'menu_period_multiple'
                            $message = "Ce filtre du groupe $groupName (terminant par _period) doit contenir uniquement 'Current Year' (actuellement: $valuesList)."
                        }
                        else {
                            $status = 'ALERTE'
                            $messageKey = 'menu_period_invalid'
                            $message = "Ce filtre du groupe $groupName (terminant par _period) doit contenir uniquement 'Current Year' (actuellement: $($currentValues[0]))."
                        }
                    }
                }
                elseif ($isSingleSelect) {
                    if ($hasSelected) {
                        $status = 'OK'
                        $messageKey = 'radio_ok'
                        $message = 'Le filtre en mode radio a une selection (autorise) et pas de recherche.'
                    }
                }
                else {
                    if ($hasSelected) {
                        $status = 'OK'
                        $messageKey = 'selection_allowed'
                        $message = 'Le filtre contient une selection (autorise).'
                    }
                }

                $result = [PSCustomObject]@{
                    PageName    = $pageDisplayName
                    DisplayName = if ($displayName) { $displayName } else { $visualKey }
                    VisualId    = $visualKey
                    FieldName   = if ($fieldName) { $fieldName } else { '' }
                    GroupName   = if ($groupName) { $groupName } else { '' }
                    HasSelected = $hasSelected
                    HasSearch   = $hasSearch
                    SearchText  = $searchText
                    IsRadio     = $isSingleSelect
                    Status      = $status
                    Message     = $message
                    MessageKey  = $messageKey
                    MessageDetail = $searchDetail
                }

                $results += $result

                $statusColor = if ($status -eq 'OK') { 'Green' } else { 'Red' }
                Write-Host "      Status: $status - $message" -ForegroundColor $statusColor
            }
        }

        Write-Host "=== Resultats (slicers visibles) ===" -ForegroundColor Green
        Write-Host "Nombre total de slicers visibles: $($results.Count)" -ForegroundColor White

        if ($hiddenCount -gt 0) {
            Write-Host "  - Filtres caches ignores: $hiddenCount" -ForegroundColor DarkGray
        }

        if ($results.Count -gt 0) {
            $alertes = $results | Where-Object { $_.Status -eq 'ALERTE' }
            $ok = $results | Where-Object { $_.Status -eq 'OK' }
            Write-Host "  - OK: $($ok.Count)" -ForegroundColor Green
            Write-Host "  - ALERTES: $($alertes.Count)" -ForegroundColor Red

            if ($alertes.Count -gt 0) {
                Write-Host "`nDetail des alertes visibles:" -ForegroundColor Yellow
                foreach ($alerte in $alertes) {
                    Write-Host "  - Page '$($alerte.PageName)' / Visuel '$($alerte.VisualId)': $($alerte.Message)" -ForegroundColor Red
                }
            }
        }
        else {
            Write-Host "  Aucun slicer visible trouve dans le projet." -ForegroundColor Yellow
        }
    }
    catch {
        Write-Error "Erreur lors de l'analyse des slicers visibles: $($_.Exception.Message)"
        $errorResult = [PSCustomObject]@{
            PageName    = 'ERREUR'
            DisplayName = 'N/A'
            VisualId    = 'N/A'
            FieldName   = ''
            GroupName   = ''
            HasSelected = $false
            HasSearch   = $false
            SearchText  = $null
            IsRadio     = $false
            Status      = 'ERREUR'
            Message     = "Erreur lors de l'analyse: $($_.Exception.Message)"
        }
        $results = ,$errorResult
    }

    return [PSCustomObject]@{
        Results        = $results
        HiddenCount    = $hiddenCount
        TotalCandidates = $totalCandidates
        ProjectPath    = $projectRoot
    }
}

Function Convert-ReportDifferencesForHtml {
    param([object[]] $differences)

    if ($null -eq $differences -or $differences.Count -eq 0) {
        return @()
    }

    # Use only the Orange HTML builder
    $command = Get-Command -Name 'BuildReportHTMLReport_Orange' -CommandType Function -ErrorAction SilentlyContinue
    if ($null -eq $command) { return $differences }

    $parameter = $command.Parameters['differences']
    if ($null -eq $parameter) {
        return $differences
    }

    $targetArrayType = $parameter.ParameterType
    if ($targetArrayType -and $targetArrayType.IsInstanceOfType($differences)) {
        return $differences
    }

    $targetType = $targetArrayType.GetElementType()
    if ($null -eq $targetType) {
        return $differences
    }

    $currentType = $differences[0].GetType()
    if ($currentType -eq $targetType) {
        return $differences
    }

    $properties = $targetType.GetProperties()
    $converted = New-Object System.Collections.Generic.List[object]

    foreach ($diff in $differences) {
        if ($null -eq $diff) {
            continue
        }

        $newDiff = [System.Activator]::CreateInstance($targetType)
        foreach ($prop in $properties) {
            $name = $prop.Name
            if ($diff.PSObject.Properties[$name]) {
                $newDiff.$name = $diff.$name
            }
        }
        $converted.Add($newDiff) | Out-Null
    }

    if ($converted.Count -eq 0) {
        return @()
    }

    $typedArray = [Array]::CreateInstance($targetType, $converted.Count)
    for ($i = 0; $i -lt $converted.Count; $i++) {
        $typedArray.SetValue($converted[$i], $i)
    }

    return $typedArray
}

# Main report comparison function WITH PARALLELIZATION
Function CompareReportProjects {
    param (
        [ReportProject] $newProject,
        [ReportProject] $oldProject
    )

    $differences = @()
    
    Write-Host "=== Comparaison des rapports Power BI ===" -ForegroundColor Green
    Write-Host "Nouveau: $($newProject.nameAndTimestamp)" -ForegroundColor Cyan
    Write-Host "Ancien: $($oldProject.nameAndTimestamp)" -ForegroundColor Cyan
    
    # Sequential comparisons for reliability
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    
    # Page comparison
    Write-Host "Comparaison des pages..." -ForegroundColor Yellow
    $pageDifferences = ComparePagesData -newPages $newProject.pagesData -oldPages $oldProject.pagesData
    $differences += $pageDifferences
    Write-Host "  $($pageDifferences.Count) differences de pages trouvees" -ForegroundColor White
    
    # Bookmark comparison
    Write-Host "Comparaison des signets..." -ForegroundColor Yellow
    $bookmarkDifferences = CompareBookmarksData `
        -newBookmarks $newProject.bookmarksData `
        -oldBookmarks $oldProject.bookmarksData `
        -newMetadata $newProject.bookmarkMetadata `
        -oldMetadata $oldProject.bookmarkMetadata `
        -newPages $newProject.pagesData `
        -oldPages $oldProject.pagesData
    $differences += $bookmarkDifferences
    Write-Host "  $($bookmarkDifferences.Count) differences de signets trouvees" -ForegroundColor White
    
    # General report configuration comparison
    Write-Host "Comparaison de la configuration..." -ForegroundColor Yellow
    $configDifferences = CompareReportData -newReport $newProject.reportData -oldReport $oldProject.reportData
    $differences += $configDifferences
    Write-Host "  $($configDifferences.Count) differences de configuration trouvees" -ForegroundColor White
    
    # System configuration files comparison (.platform, definition.pbir, etc.)
    Write-Host "Comparaison des fichiers système..." -ForegroundColor Yellow
    $systemConfigDifferences = CompareSystemConfigurationFiles -newReportPath $newProject.reportPath -oldReportPath $oldProject.reportPath
    $differences += $systemConfigDifferences
    Write-Host "  $($systemConfigDifferences.Count) differences de fichiers système trouvees" -ForegroundColor White
    
    # Theme comparison
    Write-Host "Comparaison des themes..." -ForegroundColor Yellow
    $themeDifferences = CompareReportThemesAndStyles -newReport $newProject.reportData -oldReport $oldProject.reportData
    $differences += $themeDifferences
    Write-Host "  $($themeDifferences.Count) differences de themes trouvees" -ForegroundColor White
    
    
    $stopwatch.Stop()
    Write-Host "Comparaisons terminees en $($stopwatch.ElapsedMilliseconds)ms" -ForegroundColor Green
    
    return $differences
}

# Secure page comparison function
Function ComparePagesData {
    param (
        [hashtable] $newPages,
        [hashtable] $oldPages
    )
    
    $differences = @()
    
    if ($null -eq $newPages) { $newPages = @{} }
    if ($null -eq $oldPages) { $oldPages = @{} }
    
    # Get all pages from both versions
    $allPageNames = @()
    $allPageNames += $newPages.Keys
    $allPageNames += $oldPages.Keys
    $allPageNames = $allPageNames | Select-Object -Unique
    
    foreach ($pageName in $allPageNames) {
        $newPage = $newPages[$pageName]
        $oldPage = $oldPages[$pageName]
        
        # Deleted page
        if ($oldPage -and -not $newPage) {
            $oldDisplayName = if ($oldPage.displayName) { $oldPage.displayName } else { "(Pas de nom)" }
            
            $diff = [ReportDifference]::new()
            $diff.ElementType = "Page"
            $diff.ElementName = $pageName
            $diff.ElementDisplayName = "'$oldDisplayName' ($pageName)"
            $diff.ElementPath = "pages/$pageName"
            $diff.PropertyName = "existence"
            $diff.DifferenceType = "Supprime"
            $diff.OldValue = "Present"
            $diff.NewValue = "Absent"
            $diff.HierarchyLevel = "pages"
            $diff.AdditionalInfo = "Taille: $($oldPage.width)x$($oldPage.height), Visibilite: $($oldPage.visibility)"
            $differences += $diff
            Write-Host "    Page supprimee: '$oldDisplayName' ($pageName)" -ForegroundColor Red
            continue
        }
        
        # Added page
        if ($newPage -and -not $oldPage) {
            $newDisplayName = if ($newPage.displayName) { $newPage.displayName } else { "(Pas de nom)" }
            
            $diff = [ReportDifference]::new()
            $diff.ElementType = "Page"
            $diff.ElementName = $pageName
            $diff.ElementDisplayName = "'$newDisplayName' ($pageName)"
            $diff.ElementPath = "pages/$pageName"
            $diff.PropertyName = "existence"
            $diff.DifferenceType = "Ajoute"
            $diff.OldValue = "Absent"
            $diff.NewValue = "Present"
            $diff.HierarchyLevel = "pages"
            $diff.AdditionalInfo = "Taille: $($newPage.width)x$($newPage.height), Visibilite: $($newPage.visibility)"
            $differences += $diff
            Write-Host "    Page ajoutee: '$newDisplayName' ($pageName)" -ForegroundColor Green
            continue
        }
        
        # Modified page - compare properties
        if ($newPage -and $oldPage) {
            $pageDisplayName = if ($newPage.displayName) { $newPage.displayName } else { 
                if ($oldPage.displayName) { $oldPage.displayName } else { "(Pas de nom)" }
            }
            Write-Host "    Analyse de la page: '$pageDisplayName' ($pageName)" -ForegroundColor Gray
            
            # Page filter comparison
            $filterDiffs = ComparePageFilters -pageName $pageName -newPage $newPage -oldPage $oldPage
            $differences += $filterDiffs

            # Page canvas settings comparison (Canvas Settings & Page View)
            $canvasDiffs = ComparePageCanvasSettings -pageName $pageName -newPage $newPage -oldPage $oldPage
            $differences += $canvasDiffs

            # Visual interactions comparison
            $interactionDiffs = CompareVisualInteractions -pageName $pageName -newPage $newPage -oldPage $oldPage
            $differences += $interactionDiffs

            # Page visual comparison
            $visualDiffs = ComparePageVisuals -pageName $pageName -newPage $newPage -oldPage $oldPage -allNewPages $newPages -allOldPages $oldPages
            $differences += $visualDiffs
        }
    }
    
    return $differences
}

# Enhanced secure page filter comparison function
Function ComparePageFilters {
    param (
        [string] $pageName,
        [PSCustomObject] $newPage,
        [PSCustomObject] $oldPage
    )
    
    $differences = @()
    
    try {
        $pageDisplayName = if ($newPage -and $newPage.displayName) { $newPage.displayName } else {
            if ($oldPage -and $oldPage.displayName) { $oldPage.displayName } else { "(Pas de nom)" }
        }
        $pageDisplayInfo = "'$pageDisplayName' ($pageName)"
        
        # Enhanced page filter extraction
        $newPageFilters = ExtractPageFilterDetails -page $newPage
        $oldPageFilters = ExtractPageFilterDetails -page $oldPage
        
        if ($newPageFilters.HasFilters -or $oldPageFilters.HasFilters) {
            if ($newPageFilters.FilterSummary -ne $oldPageFilters.FilterSummary) {
                $diff = [ReportDifference]::new()
                $diff.ElementType = "Page"
                $diff.ElementName = $pageName
                $diff.ElementDisplayName = $pageDisplayInfo
                $diff.ElementPath = "pages/$pageName"
                $diff.PropertyName = "filters"
                $diff.DifferenceType = "Modifie"
                $diff.OldValue = if ($oldPageFilters.FilterSummary) { $oldPageFilters.FilterSummary } else { "Aucun filtre" }
                $diff.NewValue = if ($newPageFilters.FilterSummary) { $newPageFilters.FilterSummary } else { "Aucun filtre" }
                $diff.HierarchyLevel = "pages.filters"
                $diff.AdditionalInfo = "Ancien: $($oldPageFilters.FilterDetails) | Nouveau: $($newPageFilters.FilterDetails)"
                $differences += $diff
                Write-Host "      Filtres de page '$pageDisplayName' modifies" -ForegroundColor Yellow
                Write-Host "        Ancien: $($oldPageFilters.FilterSummary)" -ForegroundColor DarkYellow
                Write-Host "        Nouveau: $($newPageFilters.FilterSummary)" -ForegroundColor DarkYellow
            }
        }
    }
    catch {
        Write-Warning "Erreur lors de la comparaison des filtres de page $pageName : $($_.Exception.Message)"
    }

    return $differences
}

# Helper function to determine canvas type from dimensions
Function Get-CanvasTypeFromDimensions {
    param(
        [int]$width,
        [int]$height,
        [string]$type
    )

    # Check for Tooltip type
    if ($type -eq "Tooltip") {
        return "Info-bulle"
    }

    # Check standard aspect ratios
    if ($width -eq 1280 -and $height -eq 720) {
        return "16:9"
    }
    elseif ($width -eq 960 -and $height -eq 720) {
        return "4:3"
    }
    elseif ($width -eq 816 -and $height -eq 1056) {
        return "Lettre"
    }
    elseif ($width -eq 320 -and $height -eq 240) {
        return "Info-bulle"
    }
    else {
        return "Personnalise ($width x $height)"
    }
}

# Helper function to translate technical property values to French
Function Translate-PropertyValueToFrench {
    param(
        [string]$propertyName,
        [string]$value
    )
    
    if ([string]::IsNullOrWhiteSpace($value)) {
        return $value
    }
    
    # Translation based on property type
    switch ($propertyName) {
        "displayOption" {
            switch ($value) {
                "FitToPage" { return "Ajuster à la page" }
                "FitToWidth" { return "Ajuster à la largeur" }
                "ActualSize" { return "Taille réelle" }
                default { return $value }
            }
        }
        "verticalAlignment" {
            switch ($value) {
                "Top" { return "Haut" }
                "Middle" { return "Centre" }
                "Bottom" { return "Bas" }
                default { return $value }
            }
        }
        "interactionType" {
            switch ($value) {
                "DataFilter" { return "Filtrer" }
                "HighlightFilter" { return "Mettre en surbrillance" }
                "NoFilter" { return "Aucun" }
                default { return $value }
            }
        }
        default {
            return $value
        }
    }
}

# Page canvas settings comparison function
Function ComparePageCanvasSettings {
    param (
        [string] $pageName,
        [PSCustomObject] $newPage,
        [PSCustomObject] $oldPage
    )

    $differences = @()

    try {
        $pageDisplayName = if ($newPage -and $newPage.displayName) { $newPage.displayName } else {
            if ($oldPage -and $oldPage.displayName) { $oldPage.displayName } else { "(Pas de nom)" }
        }
        $pageDisplayInfo = "'$pageDisplayName' ($pageName)"

        # Compare displayOption (Page View)
        $newDisplayOption = if ($newPage -and $newPage.PSObject.Properties['displayOption']) { $newPage.displayOption } else { "FitToPage" }
        $oldDisplayOption = if ($oldPage -and $oldPage.PSObject.Properties['displayOption']) { $oldPage.displayOption } else { "FitToPage" }

        if ($newDisplayOption -ne $oldDisplayOption) {
            # Translate values to French for display
            $newDisplayOptionFr = Translate-PropertyValueToFrench -propertyName "displayOption" -value $newDisplayOption
            $oldDisplayOptionFr = Translate-PropertyValueToFrench -propertyName "displayOption" -value $oldDisplayOption
            
            $diff = [ReportDifference]::new()
            $diff.ElementType = "Configuration"
            $diff.ElementName = $pageName
            $diff.ElementDisplayName = $pageDisplayInfo
            $diff.ElementPath = "pages/$pageName"
            $diff.PropertyName = "displayOption"
            $diff.DifferenceType = "Modifie"
            $diff.OldValue = $oldDisplayOptionFr
            $diff.NewValue = $newDisplayOptionFr
            $diff.HierarchyLevel = "pages.canvasSettings"
            $diff.AdditionalInfo = "Page: $pageDisplayName"
            $differences += $diff
            Write-Host "      Page View de '$pageDisplayName' modifie: $oldDisplayOptionFr -> $newDisplayOptionFr" -ForegroundColor Yellow
        }

        # Compare canvas dimensions and type
        $newWidth = if ($newPage -and $newPage.PSObject.Properties['width']) { [int]$newPage.width } else { 1280 }
        $oldWidth = if ($oldPage -and $oldPage.PSObject.Properties['width']) { [int]$oldPage.width } else { 1280 }
        $newHeight = if ($newPage -and $newPage.PSObject.Properties['height']) { [int]$newPage.height } else { 720 }
        $oldHeight = if ($oldPage -and $oldPage.PSObject.Properties['height']) { [int]$oldPage.height } else { 720 }
        $newType = if ($newPage -and $newPage.PSObject.Properties['type']) { $newPage.type } else { "" }
        $oldType = if ($oldPage -and $oldPage.PSObject.Properties['type']) { $oldPage.type } else { "" }

        # Determine canvas types
        $newCanvasType = Get-CanvasTypeFromDimensions -width $newWidth -height $newHeight -type $newType
        $oldCanvasType = Get-CanvasTypeFromDimensions -width $oldWidth -height $oldHeight -type $oldType

        if ($newCanvasType -ne $oldCanvasType) {
            $diff = [ReportDifference]::new()
            $diff.ElementType = "Configuration"
            $diff.ElementName = $pageName
            $diff.ElementDisplayName = $pageDisplayInfo
            $diff.ElementPath = "pages/$pageName"
            $diff.PropertyName = "canvasType"
            $diff.DifferenceType = "Modifie"
            $diff.OldValue = $oldCanvasType
            $diff.NewValue = $newCanvasType
            $diff.HierarchyLevel = "pages.canvasSettings"
            $diff.AdditionalInfo = "Page: $pageDisplayName | Dimensions: $oldWidth x $oldHeight -> $newWidth x $newHeight"
            $differences += $diff
            Write-Host "      Type de canevas de '$pageDisplayName' modifie: $oldCanvasType -> $newCanvasType" -ForegroundColor Yellow
        }

        # Compare vertical alignment
        $newAlignment = ""
        if ($newPage -and $newPage.PSObject.Properties['objects'] -and
            $newPage.objects.PSObject.Properties['displayArea'] -and
            $newPage.objects.displayArea.Count -gt 0) {
            $displayArea = $newPage.objects.displayArea[0]
            if ($displayArea.PSObject.Properties['properties'] -and
                $displayArea.properties.PSObject.Properties['verticalAlignment']) {
                $alignmentExpr = $displayArea.properties.verticalAlignment
                if ($alignmentExpr.PSObject.Properties['expr'] -and
                    $alignmentExpr.expr.PSObject.Properties['Literal'] -and
                    $alignmentExpr.expr.Literal.PSObject.Properties['Value']) {
                    $newAlignment = $alignmentExpr.expr.Literal.Value -replace "'", ""
                }
            }
        }

        $oldAlignment = ""
        if ($oldPage -and $oldPage.PSObject.Properties['objects'] -and
            $oldPage.objects.PSObject.Properties['displayArea'] -and
            $oldPage.objects.displayArea.Count -gt 0) {
            $displayArea = $oldPage.objects.displayArea[0]
            if ($displayArea.PSObject.Properties['properties'] -and
                $displayArea.properties.PSObject.Properties['verticalAlignment']) {
                $alignmentExpr = $displayArea.properties.verticalAlignment
                if ($alignmentExpr.PSObject.Properties['expr'] -and
                    $alignmentExpr.expr.PSObject.Properties['Literal'] -and
                    $alignmentExpr.expr.Literal.PSObject.Properties['Value']) {
                    $oldAlignment = $alignmentExpr.expr.Literal.Value -replace "'", ""
                }
            }
        }

        # Default alignment is Top if not specified
        if ([string]::IsNullOrEmpty($newAlignment)) { $newAlignment = "Top" }
        if ([string]::IsNullOrEmpty($oldAlignment)) { $oldAlignment = "Top" }

        if ($newAlignment -ne $oldAlignment) {
            # Translate values to French for display
            $newAlignmentFr = Translate-PropertyValueToFrench -propertyName "verticalAlignment" -value $newAlignment
            $oldAlignmentFr = Translate-PropertyValueToFrench -propertyName "verticalAlignment" -value $oldAlignment
            
            $diff = [ReportDifference]::new()
            $diff.ElementType = "Configuration"
            $diff.ElementName = $pageName
            $diff.ElementDisplayName = $pageDisplayInfo
            $diff.ElementPath = "pages/$pageName"
            $diff.PropertyName = "verticalAlignment"
            $diff.DifferenceType = "Modifie"
            $diff.OldValue = $oldAlignmentFr
            $diff.NewValue = $newAlignmentFr
            $diff.HierarchyLevel = "pages.canvasSettings"
            $diff.AdditionalInfo = "Page: $pageDisplayName"
            $differences += $diff
            Write-Host "      Alignement vertical de '$pageDisplayName' modifie: $oldAlignmentFr -> $newAlignmentFr" -ForegroundColor Yellow
        }
    }
    catch {
        Write-Warning "Erreur lors de la comparaison des parametres de canevas de page $pageName : $($_.Exception.Message)"
    }

    return $differences
}

# Visual interactions comparison function
Function CompareVisualInteractions {
    param (
        [string] $pageName,
        [PSCustomObject] $newPage,
        [PSCustomObject] $oldPage
    )

    $differences = @()

    try {
        $pageDisplayName = if ($newPage -and $newPage.displayName) { $newPage.displayName } else {
            if ($oldPage -and $oldPage.displayName) { $oldPage.displayName } else { "(Pas de nom)" }
        }
        $pageDisplayInfo = "'$pageDisplayName' ($pageName)"

        # Extract visual interactions from both pages
        $newInteractions = @{}
        $oldInteractions = @{}

        if ($newPage -and $newPage.PSObject.Properties['visualInteractions'] -and $newPage.visualInteractions) {
            foreach ($interaction in $newPage.visualInteractions) {
                $key = "$($interaction.source)|$($interaction.target)"
                $newInteractions[$key] = $interaction.type
            }
        }

        if ($oldPage -and $oldPage.PSObject.Properties['visualInteractions'] -and $oldPage.visualInteractions) {
            foreach ($interaction in $oldPage.visualInteractions) {
                $key = "$($interaction.source)|$($interaction.target)"
                $oldInteractions[$key] = $interaction.type
            }
        }

        # Get all interaction keys
        $allKeys = @()
        $allKeys += $newInteractions.Keys
        $allKeys += $oldInteractions.Keys
        $allKeys = $allKeys | Select-Object -Unique

        # Load visuals for display name resolution
        $newVisuals = @{}
        $oldVisuals = @{}

        if ($newPage -and $newPage.PSObject.Properties['visuals']) {
            $newVisuals = $newPage.visuals
        }

        if ($oldPage -and $oldPage.PSObject.Properties['visuals']) {
            $oldVisuals = $oldPage.visuals
        }

        foreach ($key in $allKeys) {
            $parts = $key -split '\|'
            if ($parts.Count -ne 2) { continue }

            $sourceId = $parts[0]
            $targetId = $parts[1]

            $newType = $newInteractions[$key]
            $oldType = $oldInteractions[$key]

            # Resolve display names for source and target
            $sourceDisplayName = $sourceId
            $targetDisplayName = $targetId

            # Try to resolve source display name
            $sourceVisual = $null
            if ($newVisuals -and $newVisuals.ContainsKey($sourceId)) {
                $sourceVisual = $newVisuals[$sourceId]
            }
            elseif ($oldVisuals -and $oldVisuals.ContainsKey($sourceId)) {
                $sourceVisual = $oldVisuals[$sourceId]
            }

            if ($sourceVisual) {
                # Extract visual type and fields for better fallback
                $visualType = if ($sourceVisual._extracted_type) { $sourceVisual._extracted_type } else { "visual" }
                $fields = if ($sourceVisual._extracted_fields -and $sourceVisual._extracted_fields.Count -gt 0) { 
                    $sourceVisual._extracted_fields[0] 
                } else { "" }
                
                $resolvedName = Resolve-SlicerDisplayName -visual $sourceVisual -fallbackField $fields -fallbackGroup $visualType -visualId $sourceId
                if (-not [string]::IsNullOrWhiteSpace($resolvedName)) {
                    $sourceDisplayName = $resolvedName
                }
            }

            # Try to resolve target display name
            $targetVisual = $null
            if ($newVisuals -and $newVisuals.ContainsKey($targetId)) {
                $targetVisual = $newVisuals[$targetId]
            }
            elseif ($oldVisuals -and $oldVisuals.ContainsKey($targetId)) {
                $targetVisual = $oldVisuals[$targetId]
            }

            if ($targetVisual) {
                # Extract visual type and fields for better fallback
                $visualType = if ($targetVisual._extracted_type) { $targetVisual._extracted_type } else { "visual" }
                $fields = if ($targetVisual._extracted_fields -and $targetVisual._extracted_fields.Count -gt 0) { 
                    $targetVisual._extracted_fields[0] 
                } else { "" }
                
                $resolvedName = Resolve-SlicerDisplayName -visual $targetVisual -fallbackField $fields -fallbackGroup $visualType -visualId $targetId
                if (-not [string]::IsNullOrWhiteSpace($resolvedName)) {
                    $targetDisplayName = $resolvedName
                }
            }

            # Map interaction types to French names
            $oldTypeDisplay = switch ($oldType) {
                "DataFilter" { "Filtrer" }
                "HighlightFilter" { "Mettre en surbrillance" }
                "NoFilter" { "Aucun" }
                default { if ($oldType) { $oldType } else { "Non definie" } }
            }

            $newTypeDisplay = switch ($newType) {
                "DataFilter" { "Filtrer" }
                "HighlightFilter" { "Mettre en surbrillance" }
                "NoFilter" { "Aucun" }
                default { if ($newType) { $newType } else { "Supprimee" } }
            }

            # Determine difference type
            $diffType = ""
            if (-not $oldType -and $newType) {
                $diffType = "Ajoute"
            }
            elseif ($oldType -and -not $newType) {
                $diffType = "Supprime"
            }
            elseif ($oldType -and $newType -and $oldType -ne $newType) {
                $diffType = "Modifie"
            }

            if ($diffType) {
                $diff = [ReportDifference]::new()
                $diff.ElementType = "VisualInteraction"
                $diff.ElementName = $key
                $diff.ElementDisplayName = "$sourceDisplayName → $targetDisplayName"
                $diff.ElementPath = "pages/$pageName/visualInteractions"
                $diff.PropertyName = "interactionType"
                $diff.DifferenceType = $diffType
                $diff.OldValue = $oldTypeDisplay
                $diff.NewValue = $newTypeDisplay
                $diff.HierarchyLevel = "pages.visualInteractions"
                $diff.ParentElementName = $pageName
                $diff.ParentDisplayName = $pageDisplayName
                $diff.AdditionalInfo = "Source: $sourceDisplayName ($sourceId) | Cible: $targetDisplayName ($targetId)"
                $differences += $diff

                $color = switch ($diffType) {
                    "Ajoute" { "Green" }
                    "Supprime" { "Red" }
                    "Modifie" { "Yellow" }
                    default { "Gray" }
                }

                Write-Host "      Interaction $diffType sur '$pageDisplayName': $sourceDisplayName → $targetDisplayName ($oldTypeDisplay → $newTypeDisplay)" -ForegroundColor $color
            }
        }
    }
    catch {
        Write-Warning "Erreur lors de la comparaison des interactions visuelles de page $pageName : $($_.Exception.Message)"
    }

    return $differences
}

# Page-level filter extraction function
Function ExtractPageFilterDetails {
    param ([PSCustomObject] $page)
    
    $result = @{
        HasFilters = $false
        FilterSummary = ""
        FilterDetails = ""
        Filters = @()
    }
    
    try {
        if (-not $page) {
            return $result
        }
        
        $filterInfos = @()
        
        # Check direct page filters
        if ($page.filters) {
            $filterInfos += "Filtres de page detectes"
        }
        
        if ($page.filterConfig) {
            if ($page.filterConfig.filterSortOrder) {
                $filterInfos += "Tri des filtres: $($page.filterConfig.filterSortOrder)"
            }

            if ($page.filterConfig.filters) {
                foreach ($cfgFilter in $page.filterConfig.filters) {
                    if ($cfgFilter) {
                        $fieldLabel = "filtre"
                        if ($cfgFilter.field -and $cfgFilter.field.Column -and $cfgFilter.field.Column.Property) {
                            $fieldLabel = $cfgFilter.field.Column.Property
                        }
                        elseif ($cfgFilter.field -and $cfgFilter.field.Measure -and $cfgFilter.field.Measure.Property) {
                            $fieldLabel = $cfgFilter.field.Measure.Property
                        }

                        $filterDetail = $null
                        if ($cfgFilter.filter) {
                            $filterDetail = ParseAdvancedFilterCondition -filter $cfgFilter.filter
                        }

                        if (-not $filterDetail) {
                            $filterDetail = "Configuration de filtre détectée"
                        }

                        $filterInfos += "${fieldLabel}: $filterDetail"
                    }
                }
            }
        }
        
        # Check other filter-related properties
        $filterProperties = $page.PSObject.Properties | Where-Object { $_.Name -like "*filter*" -or $_.Name -like "*Filter*" }
        foreach ($prop in $filterProperties) {
            if ($prop.Value) {
                $filterInfos += "Propriete filtre: $($prop.Name)"
            }
        }
        
        if ($filterInfos.Count -gt 0) {
            $result.HasFilters = $true
            $result.FilterSummary = ($filterInfos -join " | ")
            $result.FilterDetails = ($filterInfos -join "; ")
            $result.Filters = $filterInfos
        }
    }
    catch {
        Write-Warning "Erreur lors de l'extraction des filtres de page: $($_.Exception.Message)"
    }
    
    return $result
}

# NEW SIMPLIFIED FUNCTION: Extract simple synchronization state for a slicer
Function Get-SimpleSlicerSyncState {
    param(
        [PSCustomObject] $visual,
        [hashtable] $allPages,
        [string] $currentPageName
    )
    
    if ($null -eq $visual) { return $null }
    
    # Vérifier si le visual a un syncGroup
    $syncGroup = $null
    if ($visual.PSObject.Properties['visual'] -and 
        $visual.visual.PSObject.Properties['syncGroup']) {
        $syncGroup = $visual.visual.syncGroup
    }
    
    if ($null -eq $syncGroup) { return $null }
    
    $groupName = if ($syncGroup.groupName) { $syncGroup.groupName } else { "Sans nom" }
    
    # Trouver toutes les pages qui ont un slicer avec le même groupName
    $syncedPages = @()
    
    foreach ($pageKey in ($allPages.Keys | Sort-Object)) {
        $pageData = $allPages[$pageKey]
        if (-not $pageData) { continue }
        
        $pageDisplayName = if ($pageData.displayName) { $pageData.displayName } else { $pageKey }
        
        # Chercher dans les visuels de cette page
        if ($pageData.PSObject.Properties['visuals'] -and $pageData.visuals) {
            foreach ($visualKey in $pageData.visuals.Keys) {
                $pageVisual = $pageData.visuals[$visualKey]
                if (-not $pageVisual) { continue }
                
                # Vérifier si ce visual a le même syncGroup
                if ($pageVisual.PSObject.Properties['visual'] -and 
                    $pageVisual.visual.PSObject.Properties['syncGroup']) {
                    $pageSyncGroup = $pageVisual.visual.syncGroup
                    
                    if ($pageSyncGroup.groupName -eq $groupName) {
                        $syncedPages += $pageDisplayName
                        break
                    }
                }
            }
        }
    }
    
    if ($syncedPages.Count -eq 0) {
        return "Groupe: $groupName (aucune page synchronisée)"
    }
    
    return "Groupe: $groupName | Pages: $($syncedPages -join ', ')"
}

# OLD COMPLEX FUNCTION - KEPT FOR REFERENCE BUT NOT USED
# Extract complete synchronization state for a slicer
Function Extract-SlicerSynchronizationState {
    param (
        [PSCustomObject] $visual,
        [hashtable] $allPages,
        [string] $currentPageName
    )

    $syncState = @{
        GroupName = $null
        CurrentPage = $currentPageName
        SynchronizedPages = @{}
        FormattedString = ""
    }

    try {
        # Extract syncGroup from visual
        $syncGroup = $null
        if ($visual -and $visual.PSObject.Properties['visual'] -and
            $visual.visual.PSObject.Properties['syncGroup']) {
            $syncGroup = $visual.visual.syncGroup
        }

        if ($null -eq $syncGroup) {
            return $syncState
        }

        $syncState.GroupName = if ($syncGroup.groupName) { $syncGroup.groupName } else { "Sans nom" }

        # Check synchronization state for all pages
        foreach ($pageKey in $allPages.Keys) {
            $pageData = $allPages[$pageKey]
            $pageDisplayName = if ($pageData.displayName) { $pageData.displayName } else { $pageKey }

            # Default state: not synchronized
            $pageSync = @{
                PageId = $pageKey
                DisplayName = $pageDisplayName
                IsSynchronized = $false
                IsVisible = $false
            }

            # Check if page is in syncGroup
            # In PBIR format, syncGroup defines which pages are synchronized
            # fieldChanges = true means values are synchronized (left checkbox)
            # filterChanges = true means slicer is visible on that page (right checkbox)

            # Since we're looking at a single slicer's syncGroup, we need to check
            # if the current page being examined is synchronized with this slicer
            # The syncGroup on the slicer indicates it's part of a synchronization group

            # For now, we'll check if the page has the same slicer with the same syncGroup
            if ($pageData.PSObject.Properties['visuals'] -and $pageData.visuals) {
                foreach ($visualKey in $pageData.visuals.Keys) {
                    $pageVisual = $pageData.visuals[$visualKey]
                    if ($pageVisual -and $pageVisual.PSObject.Properties['visual'] -and
                        $pageVisual.visual -and $pageVisual.visual.PSObject.Properties['syncGroup']) {
                        $pageSyncGroup = $pageVisual.visual.syncGroup

                        # If the groups match, this page is synchronized
                        if ($pageSyncGroup.groupName -eq $syncGroup.groupName) {
                            $pageSync.IsSynchronized = $true

                            # Check the specific sync settings
                            if ($pageSyncGroup.PSObject.Properties['fieldChanges']) {
                                $pageSync.IsSynchronized = [bool]$pageSyncGroup.fieldChanges
                            }
                            if ($pageSyncGroup.PSObject.Properties['filterChanges']) {
                                $pageSync.IsVisible = [bool]$pageSyncGroup.filterChanges
                            }
                            break
                        }
                    }
                }
            }

            # Special handling for the current page where the slicer is located
            if ($pageKey -eq $currentPageName) {
                # The slicer is on this page, so it has the syncGroup
                $pageSync.IsSynchronized = $true
                if ($syncGroup.PSObject.Properties['fieldChanges']) {
                    $pageSync.IsSynchronized = [bool]$syncGroup.fieldChanges
                }
                if ($syncGroup.PSObject.Properties['filterChanges']) {
                    $pageSync.IsVisible = [bool]$syncGroup.filterChanges
                }
            }

            $syncState.SynchronizedPages[$pageKey] = $pageSync
        }

        # Format the state as a readable string
        $syncedPages = @()
        foreach ($pageKey in ($syncState.SynchronizedPages.Keys | Sort-Object)) {
            $pageInfo = $syncState.SynchronizedPages[$pageKey]
            if ($pageInfo.IsSynchronized -or $pageInfo.IsVisible) {
                $syncIcon = if ($pageInfo.IsSynchronized) { "✓" } else { "✗" }
                $visibleIcon = if ($pageInfo.IsVisible) { "✓" } else { "✗" }
                $syncedPages += "$($pageInfo.DisplayName) (sync: $syncIcon, visible: $visibleIcon)"
            }
        }

        if ($syncedPages.Count -gt 0) {
            $syncState.FormattedString = "Groupe: $($syncState.GroupName), Pages: " + ($syncedPages -join ", ")
        } else {
            $syncState.FormattedString = "Aucune synchronisation"
        }

    }
    catch {
        Write-Warning "Erreur lors de l'extraction de l'état de synchronisation: $($_.Exception.Message)"
    }

    return $syncState
}

# Get all sync group members for a specific field across all pages
Function Get-AllSyncGroups {
    param (
        [hashtable] $allPages,
        [string] $fieldName  # Ex: "CHANNEL_CD"
    )

    $groupMembers = @{}

    try {
        foreach ($pageKey in $allPages.Keys) {
            $pageData = $allPages[$pageKey]
            if (-not $pageData) { continue }

            $pageDisplayName = if ($pageData.displayName) { $pageData.displayName } else { $pageKey }

            if ($pageData.PSObject.Properties['visuals'] -and $pageData.visuals) {
                foreach ($visualKey in $pageData.visuals.Keys) {
                    $visual = $pageData.visuals[$visualKey]
                    if (-not $visual) { continue }

                    # Check if it's a slicer
                    $visualType = $null
                    if ($visual.PSObject.Properties['_extracted_type']) {
                        $visualType = $visual._extracted_type
                    } elseif ($visual.PSObject.Properties['visual'] -and
                              $visual.visual.PSObject.Properties['visualType']) {
                        $visualType = $visual.visual.visualType
                    }

                    if ($visualType -ne 'slicer') { continue }

                    # Extract field name from slicer
                    $slicerFieldName = ""
                    if ($visual.visual.PSObject.Properties['query'] -and
                        $visual.visual.query.PSObject.Properties['queryState'] -and
                        $visual.visual.query.queryState.PSObject.Properties['Values'] -and
                        $visual.visual.query.queryState.Values.PSObject.Properties['projections']) {
                        $projections = $visual.visual.query.queryState.Values.projections
                        if ($projections -is [Array] -and $projections.Count -gt 0 -and
                            $projections[0].PSObject.Properties['field'] -and
                            $projections[0].field.PSObject.Properties['Column'] -and
                            $projections[0].field.Column.PSObject.Properties['Property']) {
                            $slicerFieldName = $projections[0].field.Column.Property
                        }
                    }

                    # Only process if this slicer uses the specified field
                    if ($slicerFieldName -ne $fieldName) { continue }

                    # Check if has syncGroup
                    $syncGroup = $null
                    $hasExplicitSyncGroup = $false
                    if ($visual.visual.PSObject.Properties['syncGroup']) {
                        $syncGroup = $visual.visual.syncGroup
                        $hasExplicitSyncGroup = $true
                    }

                    # Extract sync group properties
                    $groupName = if ($syncGroup -and $syncGroup.PSObject.Properties['groupName']) {
                        $syncGroup.groupName
                    } else {
                        $fieldName  # Use field name as fallback group name
                    }
                    $fieldChanges = if ($syncGroup -and $syncGroup.PSObject.Properties['fieldChanges']) {
                        [bool]$syncGroup.fieldChanges
                    } else {
                        $false
                    }
                    $filterChanges = if ($syncGroup -and $syncGroup.PSObject.Properties['filterChanges']) {
                        [bool]$syncGroup.filterChanges
                    } else {
                        $false
                    }

                    # Store this page as a member of the group
                    $groupMembers[$pageKey] = @{
                        PageId = $pageKey
                        PageDisplayName = $pageDisplayName
                        VisualId = $visualKey
                        HasExplicitSyncGroup = $hasExplicitSyncGroup
                        FieldChanges = $fieldChanges
                        FilterChanges = $filterChanges
                        GroupName = $groupName
                    }
                }
            }
        }
    }
    catch {
        Write-Warning "Erreur lors de l'extraction des groupes de synchronisation: $($_.Exception.Message)"
    }

    return $groupMembers
}

# Compare sync group members between old and new versions
Function Compare-SyncGroupMembers {
    param (
        [hashtable] $oldMembers,  # Result from Get-AllSyncGroups (old version)
        [hashtable] $newMembers   # Result from Get-AllSyncGroups (new version)
    )

    try {
        $oldPageIds = @($oldMembers.Keys)
        $newPageIds = @($newMembers.Keys)

        # Pages added to the group (slicer appears on these pages)
        $addedPageIds = $newPageIds | Where-Object { $_ -notin $oldPageIds }

        # Pages removed from the group (slicer disappears from these pages)
        $removedPageIds = $oldPageIds | Where-Object { $_ -notin $newPageIds }

        # Pages with property modifications (fieldChanges/filterChanges)
        $modifiedPages = @()
        foreach ($pageId in $newPageIds) {
            if ($pageId -in $oldPageIds) {
                $oldMember = $oldMembers[$pageId]
                $newMember = $newMembers[$pageId]

                if ($oldMember.FieldChanges -ne $newMember.FieldChanges -or
                    $oldMember.FilterChanges -ne $newMember.FilterChanges) {
                    $modifiedPages += @{
                        PageId = $pageId
                        PageDisplayName = $newMember.PageDisplayName
                        OldFieldChanges = $oldMember.FieldChanges
                        NewFieldChanges = $newMember.FieldChanges
                        OldFilterChanges = $oldMember.FilterChanges
                        NewFilterChanges = $newMember.FilterChanges
                    }
                }
            }
        }

        # Build page info arrays for added/removed
        $addedPages = @()
        foreach ($pageId in $addedPageIds) {
            $addedPages += $newMembers[$pageId]
        }

        $removedPages = @()
        foreach ($pageId in $removedPageIds) {
            $removedPages += $oldMembers[$pageId]
        }

        return @{
            AddedPages = $addedPages
            RemovedPages = $removedPages
            ModifiedPages = $modifiedPages
            HasChanges = ($addedPages.Count -gt 0 -or $removedPages.Count -gt 0 -or $modifiedPages.Count -gt 0)
            OldMemberCount = $oldPageIds.Count
            NewMemberCount = $newPageIds.Count
        }
    }
    catch {
        Write-Warning "Erreur lors de la comparaison des membres du groupe de synchronisation: $($_.Exception.Message)"
        return @{
            AddedPages = @()
            RemovedPages = @()
            ModifiedPages = @()
            HasChanges = $false
            OldMemberCount = 0
            NewMemberCount = 0
        }
    }
}

# Compare two synchronization states and return structured differences
Function Compare-SyncStates {
    param (
        [hashtable] $oldState,
        [hashtable] $newState
    )

    try {
        # Collect all page IDs from both states
        $allPageIds = @()
        if ($oldState -and $oldState.SynchronizedPages) {
            $allPageIds += $oldState.SynchronizedPages.Keys
        }
        if ($newState -and $newState.SynchronizedPages) {
            $allPageIds += $newState.SynchronizedPages.Keys
        }
        $allPageIds = $allPageIds | Select-Object -Unique

        $addedPages = @()
        $removedPages = @()
        $modifiedPages = @()

        foreach ($pageId in $allPageIds) {
            $oldPage = if ($oldState -and $oldState.SynchronizedPages) { $oldState.SynchronizedPages[$pageId] } else { $null }
            $newPage = if ($newState -and $newState.SynchronizedPages) { $newState.SynchronizedPages[$pageId] } else { $null }

            # Check if page was synchronized in old state
            $oldWasSynced = $oldPage -and ($oldPage.IsSynchronized -or $oldPage.IsVisible)
            # Check if page is synchronized in new state
            $newIsSynced = $newPage -and ($newPage.IsSynchronized -or $newPage.IsVisible)

            # Page added to synchronization
            if (-not $oldWasSynced -and $newIsSynced) {
                $addedPages += $newPage
            }
            # Page removed from synchronization
            elseif ($oldWasSynced -and -not $newIsSynced) {
                $removedPages += $oldPage
            }
            # Page exists in both but sync/visible flags changed
            elseif ($oldWasSynced -and $newIsSynced) {
                if ($oldPage.IsSynchronized -ne $newPage.IsSynchronized -or
                    $oldPage.IsVisible -ne $newPage.IsVisible) {
                    $modifiedPages += @{
                        PageId = $pageId
                        OldState = $oldPage
                        NewState = $newPage
                    }
                }
            }
        }

        return @{
            AddedPages = $addedPages
            RemovedPages = $removedPages
            ModifiedPages = $modifiedPages
            HasChanges = ($addedPages.Count -gt 0 -or $removedPages.Count -gt 0 -or $modifiedPages.Count -gt 0)
        }
    }
    catch {
        Write-Warning "Erreur lors de la comparaison des états de synchronisation: $($_.Exception.Message)"
        return @{
            AddedPages = @()
            RemovedPages = @()
            ModifiedPages = @()
            HasChanges = $false
        }
    }
}

# Build a description of sync state with optional comparison highlighting
Function Build-SyncDescription {
    param (
        [hashtable] $syncState,
        [hashtable] $comparison,
        [switch] $HighlightChanges
    )

    try {
        if (-not $syncState -or -not $syncState.SynchronizedPages) {
            return "Aucune synchronisation"
        }

        $syncedPages = @()
        $groupName = if ($syncState.GroupName) { $syncState.GroupName } else { "Sans nom" }

        foreach ($pageKey in ($syncState.SynchronizedPages.Keys | Sort-Object)) {
            $pageInfo = $syncState.SynchronizedPages[$pageKey]
            if ($pageInfo.IsSynchronized -or $pageInfo.IsVisible) {
                $syncIcon = if ($pageInfo.IsSynchronized) { "✓" } else { "✗" }
                $visibleIcon = if ($pageInfo.IsVisible) { "✓" } else { "✗" }

                $pageDescription = "$($pageInfo.DisplayName) (sync: $syncIcon, visible: $visibleIcon)"

                # Add highlighting if requested and comparison provided
                if ($HighlightChanges -and $comparison) {
                    $isAdded = $comparison.AddedPages | Where-Object { $_.PageId -eq $pageKey }
                    $isRemoved = $comparison.RemovedPages | Where-Object { $_.PageId -eq $pageKey }

                    if ($isAdded) {
                        $pageDescription = "➕ $pageDescription"
                    }
                    if ($isRemoved) {
                        $pageDescription = "➖ $pageDescription"
                    }
                }

                $syncedPages += $pageDescription
            }
        }

        if ($syncedPages.Count -gt 0) {
            return "Groupe: $groupName, Pages: " + ($syncedPages -join ", ")
        } else {
            return "Aucune synchronisation"
        }
    }
    catch {
        Write-Warning "Erreur lors de la construction de la description de synchronisation: $($_.Exception.Message)"
        return "Erreur de formatage"
    }
}

# Calculate consequences of sync changes (slicer appearances/disappearances)
Function Calculate-SlicerConsequences {
    param (
        [hashtable] $comparison,
        [string] $slicerField,
        [string] $currentPage
    )

    try {
        $consequences = @()

        if (-not $comparison -or -not $comparison.HasChanges) {
            return $consequences
        }

        foreach ($addedPage in $comparison.AddedPages) {
            if ($addedPage.IsVisible) {
                $consequences += @{
                    Type = "Apparition"
                    Icon = "✨"
                    Message = "Le slicer '$slicerField' apparaît sur la page '$($addedPage.DisplayName)'"
                    Page = $addedPage.DisplayName
                }
            }
        }

        foreach ($removedPage in $comparison.RemovedPages) {
            if ($removedPage.IsVisible) {
                $consequences += @{
                    Type = "Disparition"
                    Icon = "🗑️"
                    Message = "Le slicer '$slicerField' disparaît de la page '$($removedPage.DisplayName)'"
                    Page = $removedPage.DisplayName
                }
            }
        }

        foreach ($modifiedPageInfo in $comparison.ModifiedPages) {
            $oldPage = $modifiedPageInfo.OldState
            $newPage = $modifiedPageInfo.NewState

            if ($oldPage.IsVisible -ne $newPage.IsVisible) {
                if ($newPage.IsVisible) {
                    $consequences += @{
                        Type = "Apparition"
                        Icon = "✨"
                        Message = "Le slicer '$slicerField' devient visible sur la page '$($newPage.DisplayName)'"
                        Page = $newPage.DisplayName
                    }
                } else {
                    $consequences += @{
                        Type = "Disparition"
                        Icon = "🗑️"
                        Message = "Le slicer '$slicerField' devient invisible sur la page '$($oldPage.DisplayName)'"
                        Page = $oldPage.DisplayName
                    }
                }
            }
        }

        return $consequences
    }
    catch {
        Write-Warning "Erreur lors du calcul des conséquences: $($_.Exception.Message)"
        return @()
    }
}

# Function to identify primary slicer in a synchronization group
Function Get-PrimarySlicer {
    param (
        [System.Collections.ArrayList] $syncGroupChanges
    )

    try {
        if (-not $syncGroupChanges -or $syncGroupChanges.Count -eq 0) {
            return $null
        }

        # Strategy to identify primary slicer:
        # 1. The slicer that exists in both old and new versions (not newly added)
        # 2. The slicer with actual syncGroup property modifications
        # 3. If multiple candidates, use the first one alphabetically by page name

        $primaryCandidate = $null
        $highestScore = -1

        foreach ($change in $syncGroupChanges) {
            $score = 0

            # Score based on change characteristics
            if ($change.OldValue -ne "Aucune synchronisation" -and $change.OldValue -ne "") {
                $score += 10 # Existed before
            }

            if ($change.NewValue -match "Pages:.*\+") {
                $score += 5 # Has additions
            }

            if ($change.DifferenceType -eq "Modifie") {
                $score += 3 # Is a modification
            }

            # Additional scoring based on the description
            if ($change.AdditionalInfo -match "Pages ajoutées|Pages retirées") {
                $score += 2
            }

            if ($score -gt $highestScore) {
                $highestScore = $score
                $primaryCandidate = $change
            } elseif ($score -eq $highestScore -and $primaryCandidate) {
                # Tie-breaker: alphabetical by parent page name
                if ($change.ParentDisplayName -lt $primaryCandidate.ParentDisplayName) {
                    $primaryCandidate = $change
                }
            }
        }

        return $primaryCandidate
    }
    catch {
        Write-Warning "Error identifying primary slicer: $($_.Exception.Message)"
        return $null
    }
}

Function Find-SyncOriginForSlicerAddition {
    param (
        [string] $fieldName,
        [string] $groupName,
        [string] $currentPageName,
        [string] $currentVisualName,
        [hashtable] $allNewPages,
        [hashtable] $allOldPages
    )

    try {
        if ([string]::IsNullOrWhiteSpace($fieldName)) {
            return $null
        }

        foreach ($pageKey in $allNewPages.Keys) {
            if ($pageKey -eq $currentPageName) {
                continue
            }

            $pageData = $allNewPages[$pageKey]
            if (-not $pageData -or -not $pageData.PSObject.Properties['visuals'] -or -not $pageData.visuals) {
                continue
            }

            foreach ($visualKey in $pageData.visuals.Keys) {
                if ($visualKey -eq $currentVisualName) {
                    continue
                }

                $candidateVisualWrapper = $pageData.visuals[$visualKey]
                if (-not $candidateVisualWrapper -or -not $candidateVisualWrapper.PSObject.Properties['visual']) {
                    continue
                }

                $candidateVisual = $candidateVisualWrapper.visual
                if (-not $candidateVisual -or $candidateVisual.visualType -ne "slicer") {
                    continue
                }

                $candidateField = Get-SlicerFieldName -visual $candidateVisualWrapper
                if ($candidateField -ne $fieldName) {
                    continue
                }

                # CRITICAL: Also check if the syncGroup.groupName matches to avoid confusing different sync groups with same field name
                if (-not [string]::IsNullOrWhiteSpace($groupName)) {
                    $candidateGroupName = ""
                    if ($candidateVisual.PSObject.Properties['syncGroup'] -and
                        $candidateVisual.syncGroup.PSObject.Properties['groupName']) {
                        $candidateGroupName = $candidateVisual.syncGroup.groupName
                    }
                    if ($candidateGroupName -ne $groupName) {
                        continue
                    }
                }

                $oldPage = if ($allOldPages.ContainsKey($pageKey)) { $allOldPages[$pageKey] } else { $null }
                if (-not $oldPage -or -not $oldPage.PSObject.Properties['visuals'] -or -not $oldPage.visuals -or -not $oldPage.visuals.ContainsKey($visualKey)) {
                    continue
                }

                $pageDisplayName = if ($pageData.displayName) { $pageData.displayName } else { $pageKey }

                return [PSCustomObject]@{
                    PageName        = $pageKey
                    PageDisplayName = $pageDisplayName
                    VisualId        = $visualKey
                    NewVisual       = $candidateVisualWrapper
                    OldVisual       = $oldPage.visuals[$visualKey]
                }
            }
        }
    }
    catch {
        Write-Warning "Error finding sync origin for slicer addition: $($_.Exception.Message)"
    }

    return $null
}

# Function to group synchronization changes by sync group
Function Group-SynchronizationChanges {
    param (
        [ReportDifference[]] $differences
    )

    try {
        $syncDifferences = $differences | Where-Object { $_.ElementType -eq "Synchronisation" }
        if ($syncDifferences.Count -eq 0) {
            return $differences
        }

        # Group by sync group name
        $syncGroups = @{}
        foreach ($diff in $syncDifferences) {
            if (-not $diff.SyncGroupName) {
                # Extract group name from NewValue or OldValue
                if ($diff.NewValue -match "Groupe:\s*([^,]+)") {
                    $diff.SyncGroupName = $matches[1].Trim()
                } elseif ($diff.OldValue -match "Groupe:\s*([^,]+)") {
                    $diff.SyncGroupName = $matches[1].Trim()
                }
            }

            if ($diff.SyncGroupName) {
                if (-not $syncGroups.ContainsKey($diff.SyncGroupName)) {
                    $syncGroups[$diff.SyncGroupName] = [System.Collections.ArrayList]::new()
                }
                [void]$syncGroups[$diff.SyncGroupName].Add($diff)
            }
        }

        # Identify primary slicer for each group and mark relationships
        $syncDifferencesToRemove = [System.Collections.ArrayList]::new()

        foreach ($groupName in $syncGroups.Keys) {
            $groupChanges = $syncGroups[$groupName]

            if ($groupChanges.Count -gt 1) {
                # Multiple slicers in same group - identify primary
                $primarySlicer = Get-PrimarySlicer -syncGroupChanges $groupChanges

                if ($primarySlicer) {
                    $primarySlicer.IsPrimarySyncChange = $true
                    $primarySlicer.SyncGroupId = [System.Guid]::NewGuid().ToString()

                    # Mark all others as related to the primary
                    foreach ($change in $groupChanges) {
                        if ($change -ne $primarySlicer) {
                            # Add to primary's related changes
                            [void]$primarySlicer.RelatedSyncChanges.Add($change)
                            # Mark for removal from main list
                            [void]$syncDifferencesToRemove.Add($change)
                        }
                    }
                }
            }
        }

        # Build final list: non-sync differences + primary sync changes only
        $finalDifferences = [System.Collections.ArrayList]::new()

        # Add all non-synchronization differences
        foreach ($diff in $differences) {
            if ($diff.ElementType -ne "Synchronisation") {
                [void]$finalDifferences.Add($diff)
            }
        }

        # Add only primary synchronization changes
        foreach ($diff in $syncDifferences) {
            if (-not $syncDifferencesToRemove.Contains($diff)) {
                [void]$finalDifferences.Add($diff)
            }
        }

        return $finalDifferences.ToArray()
    }
    catch {
        Write-Warning "Error grouping synchronization changes: $($_.Exception.Message)"
        return $differences
    }
}

# Build hierarchical index of slicer synchronization groups
Function Analyze-SyncGroupCauses {
    param (
        [ReportDifference[]] $differences
    )

    $causesIndex = @{
        Causes = [ordered]@{}
        Stats = @{
            TotalCauses = 0
            TotalConsequences = 0
            AddedSlicers = 0
            ModifiedSlicers = 0
            RemovedSlicers = 0
        }
    }

    try {
        Write-Host "=== Analyse CAUSE → CONSEQUENCES des synchronisations ===" -ForegroundColor Green

        # Filter synchronization differences
        $syncDiffs = $differences | Where-Object { $_.ElementType -eq "Synchronisation" }
        Write-Host "  Nombre de différences de synchronisation détectées: $($syncDiffs.Count)" -ForegroundColor Cyan

        if ($syncDiffs.Count -eq 0) {
            Write-Host "  Aucune différence de synchronisation à analyser" -ForegroundColor Yellow
            return $causesIndex
        }

        # Separate isHidden differences from syncGroup differences
        $isHiddenDiffs = $syncDiffs | Where-Object { $_.PropertyName -eq "isHidden" }
        $syncGroupDiffs = $syncDiffs | Where-Object { $_.PropertyName -ne "isHidden" }

        Write-Host "  Différences de visibilité (isHidden): $($isHiddenDiffs.Count)" -ForegroundColor Cyan
        Write-Host "  Différences de syncGroup: $($syncGroupDiffs.Count)" -ForegroundColor Cyan

        # Process isHidden differences separately as individual causes
        foreach ($hiddenDiff in $isHiddenDiffs) {
            # Parse AdditionalInfo to get SOURCE and TARGET page info
            $additionalInfo = $null
            try {
                $additionalInfo = $hiddenDiff.AdditionalInfo | ConvertFrom-Json
            } catch {
                Write-Warning "Impossible de parser AdditionalInfo pour isHidden: $_"
                continue
            }

            # Extract info from parsed JSON
            $sourcePageId = $additionalInfo.SourcePageId
            $sourcePageDisplayName = $additionalInfo.SourcePageDisplayName
            $targetPageId = $additionalInfo.TargetPageId
            $targetPageDisplayName = $additionalInfo.TargetPageDisplayName
            $slicerDisplayName = $additionalInfo.SlicerDisplayName
            $isHiding = $additionalInfo.IsHiding

            # Create unique key for this visual visibility change based on SOURCE page
            $causeKey = "$sourcePageId|$($additionalInfo.SourceVisualId)|visibility"

            # Determine if this is showing or hiding
            $causeType = if ($isHiding) { "Removed" } else { "Added" }

            # CAUSE: The slicer on the SOURCE page had its visibility setting changed
            $causesIndex.Causes[$causeKey] = @{
                CauseType = $causeType
                SyncGroupName = "Visibilité"
                CauseSlicer = @{
                    VisualId = $additionalInfo.SourceVisualId
                    PageId = $sourcePageId
                    PageDisplayName = $sourcePageDisplayName
                    FieldName = $slicerDisplayName -replace "Slicer \[|\]", ""
                    SyncGroupName = "Visibilité"
                }
                OldConfig = @{ Visibility = $hiddenDiff.OldValue }
                NewConfig = @{ Visibility = $hiddenDiff.NewValue }
                AdditionalInfo = "Paramétrage de visibilité modifié pour la page '$targetPageDisplayName'"
                Consequences = [System.Collections.ArrayList]::new()
            }

            # CONSEQUENCE: The slicer on the TARGET page is now hidden/visible
            $visualFieldName = $slicerDisplayName -replace "Slicer \[|\]", ""
            $descriptionText = if ($isHiding) {
                "Le visuel '$visualFieldName' est maintenant masqué sur cette page"
            } else {
                "Le visuel '$visualFieldName' est maintenant visible sur cette page"
            }

            $consequence = @{
                ActionType = if ($isHiding) { "Hidden" } else { "Shown" }
                PageId = $targetPageId
                PageDisplayName = $targetPageDisplayName
                VisualId = $additionalInfo.TargetVisualId
                VisualDisplayName = $visualFieldName
                FieldName = $visualFieldName
                Description = $descriptionText
            }
            [void]$causesIndex.Causes[$causeKey].Consequences.Add($consequence)

            # Update stats
            if ($isHiding) {
                $causesIndex.Stats.RemovedSlicers++
            } else {
                $causesIndex.Stats.AddedSlicers++
            }

            Write-Host "  ✓ Cause de visibilité créée: $causeType sur page SOURCE '$sourcePageDisplayName' → CONSÉQUENCE sur page TARGET '$targetPageDisplayName'" -ForegroundColor $(if ($isHiding) { "Red" } else { "Green" })
        }

        # Group by sync group and page
        $bySyncGroup = @{}
        foreach ($diff in $syncGroupDiffs) {
            $syncGroupName = $diff.SyncGroupName
            $pageId = $diff.ParentElementName
            $pageDisplayName = $diff.ParentDisplayName
            $visualId = $diff.ElementName

            if (-not $bySyncGroup.ContainsKey($syncGroupName)) {
                $bySyncGroup[$syncGroupName] = @{
                    GroupName = $syncGroupName
                    Changes = [System.Collections.ArrayList]::new()
                }
            }

            [void]$bySyncGroup[$syncGroupName].Changes.Add(@{
                Diff = $diff
                PageId = $pageId
                PageDisplayName = $pageDisplayName
                VisualId = $visualId
                ChangeType = $diff.DifferenceType
            })
        }

        Write-Host "  Groupes de synchronisation affectés: $($bySyncGroup.Count)" -ForegroundColor Cyan

        # Build CAUSE index - each changed slicer is a cause
        foreach ($groupName in $bySyncGroup.Keys) {
            $groupData = $bySyncGroup[$groupName]

            foreach ($change in $groupData.Changes) {
                $causeKey = "$($change.PageId)|$($change.VisualId)|$groupName"

                # Map French change types to English for HTML display
                $causeTypeEnglish = switch ($change.ChangeType) {
                    "Ajoute" { "Added" }
                    "Modifie" { "Modified" }
                    "Supprime" { "Removed" }
                    default { $change.ChangeType }
                }

                # Parse OldValue and NewValue to extract configuration details
                $oldConfig = $null
                $newConfig = $null

                # OldValue format: "Groupe: NAME, Page: PAGE (sync: ✓, visible: ✓)" or similar
                if ($change.Diff.OldValue -and $change.Diff.OldValue -ne "Absent") {
                    $oldConfig = @{
                        SyncGroupName = $groupName
                        FieldChanges = if ($change.Diff.OldValue -match "sync:\s*✓") { "True" } else { "False" }
                        FilterChanges = if ($change.Diff.OldValue -match "visible:\s*✓") { "True" } else { "False" }
                    }
                }

                # NewValue format: same as OldValue
                if ($change.Diff.NewValue -and $change.Diff.NewValue -ne "Absent") {
                    $newConfig = @{
                        SyncGroupName = $groupName
                        FieldChanges = if ($change.Diff.NewValue -match "sync:\s*✓") { "True" } else { "False" }
                        FilterChanges = if ($change.Diff.NewValue -match "visible:\s*✓") { "True" } else { "False" }
                    }
                }

                # Extract slicer name without "sur page" suffix
                $slicerNameOnly = $change.Diff.ElementDisplayName -replace " sur page '.*'$", ""

                $causesIndex.Causes[$causeKey] = @{
                    CauseType = $causeTypeEnglish
                    SyncGroupName = $groupName
                    CauseSlicer = @{
                        VisualId = $change.VisualId
                        PageId = $change.PageId
                        PageDisplayName = $change.PageDisplayName
                        FieldName = $slicerNameOnly
                        SyncGroupName = $groupName
                    }
                    OldConfig = $oldConfig
                    NewConfig = $newConfig
                    AdditionalInfo = $change.Diff.AdditionalInfo
                    Consequences = [System.Collections.ArrayList]::new()
                }

                # Add consequences based on real sync changes (if available) or fallback to group members
                if ($change.Diff._syncComparison) {
                    # Use the structured comparison to get real consequences
                    $syncComp = $change.Diff._syncComparison
                    $slicerFieldName = $change.Diff.ElementDisplayName -replace " sur page '.*'$", ""

                    # Added pages: slicer now appears on these pages
                    foreach ($addedPage in $syncComp.AddedPages) {
                        [void]$causesIndex.Causes[$causeKey].Consequences.Add(@{
                            PageId = $addedPage.PageId
                            PageDisplayName = $addedPage.PageDisplayName
                            FieldName = $slicerFieldName
                            ActionType = "Synchronized"
                            Description = "Le slicer '$slicerFieldName' apparaît sur la page '$($addedPage.PageDisplayName)'"
                        })
                    }

                    # Removed pages: slicer no longer appears on these pages
                    foreach ($removedPage in $syncComp.RemovedPages) {
                        [void]$causesIndex.Causes[$causeKey].Consequences.Add(@{
                            PageId = $removedPage.PageId
                            PageDisplayName = $removedPage.PageDisplayName
                            FieldName = $slicerFieldName
                            ActionType = "Desynchronized"
                            Description = "Le slicer '$slicerFieldName' disparaît de la page '$($removedPage.PageDisplayName)'"
                        })
                    }

                    # Modified pages: changes in sync/visible state
                    foreach ($modifiedPage in $syncComp.ModifiedPages) {
                        $modDesc = @()
                        if ($modifiedPage.OldFieldChanges -ne $modifiedPage.NewFieldChanges) {
                            $modDesc += if ($modifiedPage.NewFieldChanges) { "synchronisation activée" } else { "synchronisation désactivée" }
                        }
                        if ($modifiedPage.OldFilterChanges -ne $modifiedPage.NewFilterChanges) {
                            $modDesc += if ($modifiedPage.NewFilterChanges) { "visibilité activée" } else { "visibilité désactivée" }
                        }
                        if ($modDesc.Count -gt 0) {
                            [void]$causesIndex.Causes[$causeKey].Consequences.Add(@{
                                PageId = $modifiedPage.PageId
                                PageDisplayName = $modifiedPage.PageDisplayName
                                FieldName = $slicerFieldName
                                ActionType = "Modified"
                                Description = "Le slicer '$slicerFieldName' sur la page '$($modifiedPage.PageDisplayName)': $($modDesc -join ', ')"
                            })
                        }
                    }
                }
                else {
                    # Fallback: Add other pages in same group as consequences
                    # Determine ActionType based on cause type
                    $actionType = switch ($causeTypeEnglish) {
                        "Added" { "Synchronized" }
                        "Removed" { "Desynchronized" }
                        "Modified" { "Synchronized" }
                        default { "Synchronized" }
                    }

                    foreach ($otherChange in $groupData.Changes) {
                        if ($otherChange.PageId -ne $change.PageId -or $otherChange.VisualId -ne $change.VisualId) {
                            [void]$causesIndex.Causes[$causeKey].Consequences.Add(@{
                                PageId = $otherChange.PageId
                                PageDisplayName = $otherChange.PageDisplayName
                                VisualId = $otherChange.VisualId
                                FieldName = $otherChange.Diff.ElementDisplayName
                                ActionType = $actionType
                            })
                        }
                    }
                }

                # Update stats (will be recalculated after refinement)
                if ($change.ChangeType -eq "Ajoute") {
                    $causesIndex.Stats.AddedSlicers++
                } elseif ($change.ChangeType -eq "Modifie") {
                    $causesIndex.Stats.ModifiedSlicers++
                } elseif ($change.ChangeType -eq "Supprime") {
                    $causesIndex.Stats.RemovedSlicers++
                }
            }
        }

        # Refine CauseType based on consequence analysis
        Write-Host "  Raffinement des types de causes basé sur les conséquences..." -ForegroundColor Yellow
        $refinedCount = 0

        foreach ($causeKey in $causesIndex.Causes.Keys) {
            $cause = $causesIndex.Causes[$causeKey]

            # Only refine "Modified" causes that have consequences
            if ($cause.CauseType -eq "Modified" -and $cause.Consequences.Count -gt 0) {
                # Analyze consequence types
                $syncedCount = ($cause.Consequences | Where-Object { $_.ActionType -eq "Synchronized" }).Count
                $desyncedCount = ($cause.Consequences | Where-Object { $_.ActionType -eq "Desynchronized" }).Count
                $modifiedCount = ($cause.Consequences | Where-Object { $_.ActionType -eq "Modified" }).Count

                $originalType = $cause.CauseType

                # If all consequences are Synchronized → it's an addition of pages to sync group
                if ($syncedCount -gt 0 -and $desyncedCount -eq 0 -and $modifiedCount -eq 0) {
                    $cause.CauseType = "Added"
                    Write-Host "    ✓ Raffiné: Modified → Added ($syncedCount page(s) synchronisée(s))" -ForegroundColor Green
                    $refinedCount++
                }
                # If all consequences are Desynchronized → it's a removal of pages from sync group
                elseif ($desyncedCount -gt 0 -and $syncedCount -eq 0 -and $modifiedCount -eq 0) {
                    $cause.CauseType = "Removed"
                    Write-Host "    ✓ Raffiné: Modified → Removed ($desyncedCount page(s) désynchronisée(s))" -ForegroundColor Red
                    $refinedCount++
                }
                # Otherwise keep as Modified (mixed changes)
            }
        }

        Write-Host "  Nombre de causes raffinées: $refinedCount" -ForegroundColor Cyan

        # Recalculate stats after refinement
        $causesIndex.Stats.AddedSlicers = 0
        $causesIndex.Stats.ModifiedSlicers = 0
        $causesIndex.Stats.RemovedSlicers = 0

        foreach ($cause in $causesIndex.Causes.Values) {
            if ($cause.CauseType -eq "Added") {
                $causesIndex.Stats.AddedSlicers++
            } elseif ($cause.CauseType -eq "Modified") {
                $causesIndex.Stats.ModifiedSlicers++
            } elseif ($cause.CauseType -eq "Removed") {
                $causesIndex.Stats.RemovedSlicers++
            }
        }

        # Calculate total stats
        $causesIndex.Stats.TotalCauses = $causesIndex.Causes.Count
        $causesIndex.Stats.TotalConsequences = ($causesIndex.Causes.Values | ForEach-Object { $_.Consequences.Count } | Measure-Object -Sum).Sum

        Write-Host "  CAUSES détectées: $($causesIndex.Stats.TotalCauses)" -ForegroundColor Green
        Write-Host "    - Ajouts: $($causesIndex.Stats.AddedSlicers)" -ForegroundColor Cyan
        Write-Host "    - Modifications: $($causesIndex.Stats.ModifiedSlicers)" -ForegroundColor Cyan
        Write-Host "    - Suppressions: $($causesIndex.Stats.RemovedSlicers)" -ForegroundColor Cyan
        Write-Host "  CONSEQUENCES identifiées: $($causesIndex.Stats.TotalConsequences)" -ForegroundColor Green

        return $causesIndex
    }
    catch {
        Write-Warning "Erreur lors de l'analyse CAUSE → CONSEQUENCES : $($_.Exception.Message)"
        Write-Warning "Stack Trace: $($_.ScriptStackTrace)"
        return $causesIndex
    }
}

# Secure page visual comparison function
Function ComparePageVisuals {
    param (
        [string] $pageName,
        [string] $pageDisplayName,
        [PSCustomObject] $newPage,
        [PSCustomObject] $oldPage,
        [hashtable] $allNewPages = @{},
        [hashtable] $allOldPages = @{}
    )
    
    $differences = @()
    
    try {
        if ([string]::IsNullOrWhiteSpace($pageDisplayName)) {
            if ($newPage -and $newPage.displayName) {
                $pageDisplayName = $newPage.displayName
            }
            elseif ($oldPage -and $oldPage.displayName) {
                $pageDisplayName = $oldPage.displayName
            }
            else {
                $pageDisplayName = "(Pas de nom)"
            }
        }
        # Get visuals securely
        $newVisuals = @{}
        $oldVisuals = @{}
        
        if ($null -ne $newPage -and $null -ne (Get-Member -InputObject $newPage -Name "visuals" -MemberType Properties) -and $null -ne $newPage.visuals) {
            $newVisuals = $newPage.visuals
        }
        
        if ($null -ne $oldPage -and $null -ne (Get-Member -InputObject $oldPage -Name "visuals" -MemberType Properties) -and $null -ne $oldPage.visuals) {
            $oldVisuals = $oldPage.visuals
        }
        
        # Get all visual names
        $allVisualNames = @()
        $allVisualNames += $newVisuals.Keys
        $allVisualNames += $oldVisuals.Keys
        $allVisualNames = $allVisualNames | Select-Object -Unique
        
        foreach ($visualName in $allVisualNames) {
            $newVisual = $newVisuals[$visualName]
            $oldVisual = $oldVisuals[$visualName]
            
            # Deleted visual
            if ($oldVisual -and -not $newVisual) {
                $oldVisualType = if ($oldVisual._extracted_type) { $oldVisual._extracted_type } else { "unknown" }
                $oldFields = if ($oldVisual._extracted_fields -and $oldVisual._extracted_fields.Count -gt 0) {
                    " [" + ($oldVisual._extracted_fields -join ", ") + "]"
                } else { "" }
                
                # ENHANCED: Special handling for buttons
                $elementType = if ($oldVisualType -eq "actionButton") { "Bouton" } else { "Visuel" }
                $buttonDetails = if ($oldVisual._extracted_button_details -and $oldVisual._extracted_button_details.Count -gt 0) {
                    " - " + ($oldVisual._extracted_button_details -join " | ")
                } else { "" }
                
                $diff = [ReportDifference]::new()
                $diff.ElementType = $elementType
                $diff.ElementName = $visualName
                $diff.ElementDisplayName = "$oldVisualType$oldFields sur page '$pageDisplayName' ($visualName)"
                $diff.ElementPath = "pages/$pageName/visuals/$visualName"
                $diff.PropertyName = "existence"
                $diff.DifferenceType = "Supprime"
                $diff.OldValue = "Present"
                $diff.NewValue = "Absent"
                $diff.HierarchyLevel = "pages.visuals"
                $diff.AdditionalInfo = "Type: $oldVisualType, Champs: " + ($oldVisual._extracted_fields -join ", ") + $buttonDetails
                $diff.ParentElementName = $pageName
                $diff.ParentDisplayName = $pageDisplayName
                $differences += $diff
                
                # NEW SIMPLIFIED SYNC FOR DELETED SLICERS
                if ($oldVisualType -eq "slicer") {
                    try {
                        $oldSyncState = Get-SimpleSlicerSyncState -visual $oldVisual -allPages $allOldPages -currentPageName $pageName
                        
                        if ($oldSyncState) {
                            # Get slicer display name
                            $slicerDisplayName = Resolve-SlicerDisplayName -visual $oldVisual -visualId $visualName -fallbackField "" -fallbackGroup ""
                            if ([string]::IsNullOrWhiteSpace($slicerDisplayName)) {
                                $slicerDisplayName = "Slicer ($visualName)"
                            }
                            
                            $syncDiff = [ReportDifference]::new()
                            $syncDiff.ElementType = "Synchronisation"
                            $syncDiff.ElementName = $visualName
                            $syncDiff.ElementDisplayName = "$slicerDisplayName sur page '$pageDisplayName'"
                            $syncDiff.ElementPath = "pages/$pageName/visuals/$visualName/syncGroup"
                            $syncDiff.PropertyName = "Groupe de synchronisation"
                            $syncDiff.DifferenceType = "Supprime"
                            $syncDiff.OldValue = $oldSyncState
                            $syncDiff.NewValue = "Aucune synchronisation"
                            $syncDiff.HierarchyLevel = "pages.visuals.synchronisation"
                            $syncDiff.ParentElementName = $pageName
                            $syncDiff.ParentDisplayName = $pageDisplayName
                            $syncDiff.AdditionalInfo = "Suppression de synchronisation pour $slicerDisplayName sur page '$pageDisplayName'"
                            
                            $differences += $syncDiff
                        }
                    }
                    catch {
                        Write-Warning "Erreur lors de la creation de la difference de synchronisation pour le slicer supprime $visualName : $($_.Exception.Message)"
                    }
                }
                
                if ($oldVisualType -eq "actionButton") {
                    Write-Host "      BOUTON supprime: $oldVisualType$oldFields sur page '$pageDisplayName' ($visualName)$buttonDetails" -ForegroundColor Magenta
                } else {
                    Write-Host "      Visuel supprime: $oldVisualType$oldFields sur page '$pageDisplayName' ($visualName)" -ForegroundColor Red
                }
                continue
            }
            
            # Added visual
            if ($newVisual -and -not $oldVisual) {
                $newVisualType = if ($newVisual._extracted_type) { $newVisual._extracted_type } else { "unknown" }
                $newFields = if ($newVisual._extracted_fields -and $newVisual._extracted_fields.Count -gt 0) {
                    " [" + ($newVisual._extracted_fields -join ", ") + "]"
                } else { "" }
                
                # ENHANCED: Special handling for buttons
                $elementType = if ($newVisualType -eq "actionButton") { "Bouton" } else { "Visuel" }
                $buttonDetails = if ($newVisual._extracted_button_details -and $newVisual._extracted_button_details.Count -gt 0) {
                    " - " + ($newVisual._extracted_button_details -join " | ")
                } else { "" }
                
                $diff = [ReportDifference]::new()
                $diff.ElementType = $elementType
                $diff.ElementName = $visualName
                $diff.ElementDisplayName = "$newVisualType$newFields sur page '$pageDisplayName' ($visualName)"
                $diff.ElementPath = "pages/$pageName/visuals/$visualName"
                $diff.PropertyName = "existence"
                $diff.DifferenceType = "Ajoute"
                $diff.OldValue = "Absent"
                $diff.NewValue = "Present"
                $diff.HierarchyLevel = "pages.visuals"
                $diff.AdditionalInfo = "Type: $newVisualType, Champs: " + ($newVisual._extracted_fields -join ", ") + $buttonDetails
                $diff.ParentElementName = $pageName
                $diff.ParentDisplayName = $pageDisplayName
                $differences += $diff
                
                # NEW SIMPLIFIED SYNC FOR ADDED SLICERS
                if ($newVisualType -eq "slicer") {
                    try {
                        $newSyncState = Get-SimpleSlicerSyncState -visual $newVisual -allPages $allNewPages -currentPageName $pageName
                        
                        if ($newSyncState) {
                            # Get slicer display name
                            $slicerDisplayName = Resolve-SlicerDisplayName -visual $newVisual -visualId $visualName -fallbackField "" -fallbackGroup ""
                            if ([string]::IsNullOrWhiteSpace($slicerDisplayName)) {
                                $slicerDisplayName = "Slicer ($visualName)"
                            }
                            
                            $syncDiff = [ReportDifference]::new()
                            $syncDiff.ElementType = "Synchronisation"
                            $syncDiff.ElementName = $visualName
                            $syncDiff.ElementDisplayName = "$slicerDisplayName sur page '$pageDisplayName'"
                            $syncDiff.ElementPath = "pages/$pageName/visuals/$visualName/syncGroup"
                            $syncDiff.PropertyName = "Groupe de synchronisation"
                            $syncDiff.DifferenceType = "Ajoute"
                            $syncDiff.OldValue = "Aucune synchronisation"
                            $syncDiff.NewValue = $newSyncState
                            $syncDiff.HierarchyLevel = "pages.visuals.synchronisation"
                            $syncDiff.ParentElementName = $pageName
                            $syncDiff.ParentDisplayName = $pageDisplayName
                            $syncDiff.AdditionalInfo = "Ajout de synchronisation pour $slicerDisplayName sur page '$pageDisplayName'"
                            
                            $differences += $syncDiff
                        }
                    }
                    catch {
                        Write-Warning "Erreur lors de la creation de la difference de synchronisation pour le slicer ajoute $visualName : $($_.Exception.Message)"
                    }
                }

                if ($newVisualType -eq "actionButton") {
                    Write-Host "      BOUTON ajoute: $newVisualType$newFields sur page '$pageDisplayName' ($visualName)$buttonDetails" -ForegroundColor Cyan
                } else {
                    Write-Host "      Visuel ajoute: $newVisualType$newFields sur page '$pageDisplayName' ($visualName)" -ForegroundColor Green
                }
                continue
            }
            
            # Modified visual
            if ($newVisual -and $oldVisual) {
                $visualType = if ($newVisual._extracted_type) { $newVisual._extracted_type } else {
                    if ($oldVisual._extracted_type) { $oldVisual._extracted_type } else { "unknown" }
                }
                Write-Host "        Analyse du visuel: $visualType sur page '$pageDisplayName' ($visualName)" -ForegroundColor DarkGray
                $visualDiffs = CompareVisualProperties -pageName $pageName -pageDisplayName $pageDisplayName -visualName $visualName -newVisual $newVisual -oldVisual $oldVisual -allNewPages $allNewPages -allOldPages $allOldPages
                $differences += $visualDiffs
                
                # NEW SIMPLIFIED SYNC COMPARISON: Compare synchronization state for slicers
                if ($visualType -eq 'slicer') {
                    try {
                        $newSyncState = Get-SimpleSlicerSyncState -visual $newVisual -allPages $allNewPages -currentPageName $pageName
                        $oldSyncState = Get-SimpleSlicerSyncState -visual $oldVisual -allPages $allOldPages -currentPageName $pageName
                        
                        # Compare sync states
                        if ($newSyncState -ne $oldSyncState) {
                            # Get slicer display name
                            $slicerDisplayName = Resolve-SlicerDisplayName -visual $newVisual -visualId $visualName -fallbackField "" -fallbackGroup ""
                            if ([string]::IsNullOrWhiteSpace($slicerDisplayName)) {
                                $slicerDisplayName = "Slicer ($visualName)"
                            }
                            
                            $syncDiff = [ReportDifference]::new()
                            $syncDiff.ElementType = "Synchronisation"
                            $syncDiff.ElementName = $visualName
                            $syncDiff.ElementDisplayName = "$slicerDisplayName sur page '$pageDisplayName'"
                            $syncDiff.ElementPath = "pages/$pageName/visuals/$visualName/syncGroup"
                            $syncDiff.PropertyName = "Groupe de synchronisation"
                            $syncDiff.OldValue = if ($oldSyncState) { $oldSyncState } else { "Aucune synchronisation" }
                            $syncDiff.NewValue = if ($newSyncState) { $newSyncState } else { "Aucune synchronisation" }
                            $syncDiff.DifferenceType = if ($oldSyncState -and $newSyncState) { "Modifie" } 
                                                       elseif ($newSyncState) { "Ajoute" } 
                                                       else { "Supprime" }
                            $syncDiff.HierarchyLevel = "pages.visuals.synchronisation"
                            $syncDiff.ParentElementName = $pageName
                            $syncDiff.ParentDisplayName = $pageDisplayName
                            $syncDiff.AdditionalInfo = "Modification de synchronisation pour $slicerDisplayName sur page '$pageDisplayName'"
                            
                            $differences += $syncDiff
                            Write-Host "          Synchronisation modifiee: $slicerDisplayName" -ForegroundColor Yellow
                        }
                    }
                    catch {
                        Write-Warning "Erreur lors de la comparaison de synchronisation pour le slicer $visualName : $($_.Exception.Message)"
                    }
                }
            }
        }
    }
    catch {
        Write-Warning "Erreur lors de la comparaison des visuels de la page $pageName : $($_.Exception.Message)"
    }
    
    return $differences
}

# Function to extract the display name of a visual from its title properties
Function Get-VisualDisplayName {
    param (
        [PSCustomObject] $visual,
        [string] $visualName
    )

    try {
        # Priority 1: Check visual.objects.general[].properties.title
        if ($visual.PSObject.Properties['visual'] -and
            $visual.visual.PSObject.Properties['objects'] -and
            $visual.visual.objects.PSObject.Properties['general']) {

            $generalObjects = $visual.visual.objects.general
            if ($generalObjects -is [Array]) {
                foreach ($generalItem in $generalObjects) {
                    if ($generalItem.PSObject.Properties['properties'] -and
                        $generalItem.properties.PSObject.Properties['title']) {
                        $titleProp = $generalItem.properties.title
                        if ($titleProp.PSObject.Properties['expr'] -and
                            $titleProp.expr.PSObject.Properties['Literal'] -and
                            $titleProp.expr.Literal.PSObject.Properties['Value']) {
                            $titleValue = $titleProp.expr.Literal.Value
                            if ($titleValue -and $titleValue.ToString().Trim() -ne "") {
                                return $titleValue.ToString().Trim().Trim("'")
                            }
                        }
                    }
                }
            }
        }

        # Priority 2: Check visual.objects.title[].properties.text
        if ($visual.PSObject.Properties['visual'] -and
            $visual.visual.PSObject.Properties['objects'] -and
            $visual.visual.objects.PSObject.Properties['title']) {

            $titleObjects = $visual.visual.objects.title
            if ($titleObjects -is [Array]) {
                foreach ($titleItem in $titleObjects) {
                    if ($titleItem.PSObject.Properties['properties'] -and
                        $titleItem.properties.PSObject.Properties['text']) {
                        $textProp = $titleItem.properties.text
                        if ($textProp.PSObject.Properties['expr'] -and
                            $textProp.expr.PSObject.Properties['Literal'] -and
                            $textProp.expr.Literal.PSObject.Properties['Value']) {
                            $textValue = $textProp.expr.Literal.Value
                            if ($textValue -and $textValue.ToString().Trim() -ne "") {
                                return $textValue.ToString().Trim().Trim("'")
                            }
                        }
                    }
                }
            }
        }

        # Fallback: Return technical type with visual ID
        $visualType = if ($visual._extracted_type) { $visual._extracted_type } else { "unknown" }
        return "$visualType ($visualName)"
    }
    catch {
        # In case of error, return technical type
        $visualType = if ($visual._extracted_type) { $visual._extracted_type } else { "unknown" }
        return "$visualType ($visualName)"
    }
}

# Enhanced visual properties comparison function
Function CompareVisualProperties {
    param (
        [string] $pageName,
        [string] $pageDisplayName,
        [string] $visualName,
        [PSCustomObject] $newVisual,
        [PSCustomObject] $oldVisual,
        [hashtable] $allNewPages = @{},
        [hashtable] $allOldPages = @{}
    )
    
    $differences = @()
    
    try {
        # Get the user-friendly display name of the visual
        $visualDisplayName = Get-VisualDisplayName -visual $newVisual -visualName $visualName
        $visualType = if ($newVisual._extracted_type) { $newVisual._extracted_type } else { "unknown" }

        # Format: "Matrix rating (pivotTable)" if has title, or "pivotTable (id)" if not
        if ($visualDisplayName -notlike "*$visualName*") {
            $visualDisplayInfo = "$visualDisplayName ($visualType)"
        } else {
            $visualDisplayInfo = $visualDisplayName
        }

        $fullDisplayInfo = "$visualDisplayInfo sur page '$pageDisplayName'"
        
        # Enhanced visual filter comparison
        $newFiltersInfo = ExtractFilterDetails -visual $newVisual
        $oldFiltersInfo = ExtractFilterDetails -visual $oldVisual
        
        if ($newFiltersInfo.HasFilters -or $oldFiltersInfo.HasFilters) {
            if ($newFiltersInfo.FilterSummary -ne $oldFiltersInfo.FilterSummary) {
                $diff = [ReportDifference]::new()
                $diff.ElementType = "Visuel"
                $diff.ElementName = $visualName
                $diff.ElementDisplayName = $fullDisplayInfo
                $diff.ElementPath = "pages/$pageName/visuals/$visualName"
                $diff.PropertyName = "filters"
                $diff.DifferenceType = "Modifie"
                $diff.OldValue = if ($oldFiltersInfo.FilterSummary) { $oldFiltersInfo.FilterSummary } else { "Aucun filtre" }
                $diff.NewValue = if ($newFiltersInfo.FilterSummary) { $newFiltersInfo.FilterSummary } else { "Aucun filtre" }
                $diff.HierarchyLevel = "pages.visuals.filters"
                $diff.AdditionalInfo = "Ancien: $($oldFiltersInfo.FilterDetails) | Nouveau: $($newFiltersInfo.FilterDetails)"
                $diff.ParentElementName = $pageName
                $diff.ParentDisplayName = $pageDisplayName
                $differences += $diff
                Write-Host "        Filtres du visuel $fullDisplayInfo modifies" -ForegroundColor Yellow
                Write-Host "          Ancien: $($oldFiltersInfo.FilterSummary)" -ForegroundColor DarkYellow
                Write-Host "          Nouveau: $($newFiltersInfo.FilterSummary)" -ForegroundColor DarkYellow
            }
        }
        
        # Visual type comparison
        $newType = if ($newVisual._extracted_type) { $newVisual._extracted_type } else { "unknown" }
        $oldType = if ($oldVisual._extracted_type) { $oldVisual._extracted_type } else { "unknown" }
        
        if ($newType -ne $oldType) {
            $diff = [ReportDifference]::new()
            $diff.ElementType = "Visuel"
            $diff.ElementName = $visualName
            $diff.ElementDisplayName = $fullDisplayInfo
            $diff.ElementPath = "pages/$pageName/visuals/$visualName"
            $diff.PropertyName = "visualType"
            $diff.DifferenceType = "Modifie"
            $diff.OldValue = $oldType
            $diff.NewValue = $newType
            $diff.HierarchyLevel = "pages.visuals.visualType"
            $diff.AdditionalInfo = "Changement de type de visuel"
            $diff.ParentElementName = $pageName
            $diff.ParentDisplayName = $pageDisplayName
            $differences += $diff
            Write-Host "        Type du visuel modifie: $oldType -> $newType sur page '$pageDisplayName' ($visualName)" -ForegroundColor Yellow
        }
        
        # Used fields comparison - Create separate differences for each field
        $newFields = @()
        if ($newVisual._extracted_fields) { $newFields = @($newVisual._extracted_fields) }

        $oldFields = @()
        if ($oldVisual._extracted_fields) { $oldFields = @($oldVisual._extracted_fields) }

        $newFieldsText = if ($newFields.Count -gt 0) { $newFields -join ", " } else { "" }
        $oldFieldsText = if ($oldFields.Count -gt 0) { $oldFields -join ", " } else { "" }

        if ($newFieldsText -ne $oldFieldsText) {
            $addedFields = $newFields | Where-Object { $_ -notin $oldFields }
            $removedFields = $oldFields | Where-Object { $_ -notin $newFields }

            # Create a summary "Modifie" difference for the visual (for "Visuels" table)
            if (($removedFields.Count -gt 0) -or ($addedFields.Count -gt 0)) {
                $summaryParts = @()
                if ($removedFields.Count -gt 0) {
                    $summaryParts += "$($removedFields.Count) champ(s) supprime(s)"
                }
                if ($addedFields.Count -gt 0) {
                    $summaryParts += "$($addedFields.Count) champ(s) ajoute(s)"
                }
                $summary = $summaryParts -join ", "

                $diff = [ReportDifference]::new()
                $diff.ElementType = "Visuel"
                $diff.ElementName = $visualName
                $diff.ElementDisplayName = $visualDisplayName
                $diff.ElementPath = "pages/$pageName/visuals/$visualName"
                $diff.PropertyName = "fields_summary"
                $diff.DifferenceType = "Modifie"
                $diff.OldValue = $summary
                $diff.NewValue = $summary
                $diff.HierarchyLevel = "pages.visuals"
                $diff.AdditionalInfo = $summary
                $diff.ParentElementName = $pageName
                $diff.ParentDisplayName = $pageDisplayName
                $differences += $diff
                Write-Host "        Visuel $visualDisplayName modifie: $summary" -ForegroundColor Yellow
            }

            # Create individual differences for each removed field (for "Champs du visuel" table)
            # IMPORTANT: ElementType MUST be "Visuel" and PropertyName MUST be "fields" for HTML filter to work
            foreach ($removedField in $removedFields) {
                $diff = [ReportDifference]::new()
                $diff.ElementType = "Visuel"
                $diff.ElementName = $visualName
                $diff.ElementDisplayName = $visualDisplayName
                $diff.ElementPath = "pages/$pageName/visuals/$visualName/fields/$removedField"
                $diff.PropertyName = "fields"
                $diff.DifferenceType = "Supprime"
                $diff.OldValue = $removedField
                $diff.NewValue = "Absent"
                $diff.HierarchyLevel = "pages.visuals.fields"
                $diff.AdditionalInfo = "Champ '$removedField' supprime"
                $diff.ParentElementName = $pageName
                $diff.ParentDisplayName = $pageDisplayName
                $differences += $diff
                Write-Host "        Champ supprime: '$removedField' du visuel $visualDisplayName" -ForegroundColor Red
            }

            # Create individual differences for each added field (for "Champs du visuel" table)
            foreach ($addedField in $addedFields) {
                $diff = [ReportDifference]::new()
                $diff.ElementType = "Visuel"
                $diff.ElementName = $visualName
                $diff.ElementDisplayName = $visualDisplayName
                $diff.ElementPath = "pages/$pageName/visuals/$visualName/fields/$addedField"
                $diff.PropertyName = "fields"
                $diff.DifferenceType = "Ajoute"
                $diff.OldValue = "Absent"
                $diff.NewValue = $addedField
                $diff.HierarchyLevel = "pages.visuals.fields"
                $diff.AdditionalInfo = "Champ '$addedField' ajoute"
                $diff.ParentElementName = $pageName
                $diff.ParentDisplayName = $pageDisplayName
                $differences += $diff
                Write-Host "        Champ ajoute: '$addedField' au visuel $visualDisplayName" -ForegroundColor Green
            }

            # If fields changed but no adds/removes (reordering only), create a "Visuel" difference
            if (($addedFields.Count -eq 0) -and ($removedFields.Count -eq 0) -and ($newFields.Count -gt 0) -and ($oldFields.Count -gt 0)) {
                $diff = [ReportDifference]::new()
                $diff.ElementType = "Visuel"
                $diff.ElementName = $visualName
                $diff.ElementDisplayName = $visualDisplayName
                $diff.ElementPath = "pages/$pageName/visuals/$visualName"
                $diff.PropertyName = "fieldsOrder"
                $diff.DifferenceType = "Modifie"
                $diff.OldValue = $oldFieldsText
                $diff.NewValue = $newFieldsText
                $diff.HierarchyLevel = "pages.visuals.fields"
                $diff.AdditionalInfo = "Ordre ou configuration des champs mise a jour"
                $diff.ParentElementName = $pageName
                $diff.ParentDisplayName = $pageDisplayName
                $differences += $diff
                Write-Host "        Ordre des champs modifie pour le visuel $visualDisplayInfo" -ForegroundColor Yellow
            }
        }

        # ===== COMPARAISON DE LA VISIBILITÉ (isHidden) =====
        # Détecte les changements de visibilité des visuels (principalement pour les slicers synchronisés)
        $oldIsHidden = if ($oldVisual.PSObject.Properties['isHidden']) { [bool]$oldVisual.isHidden } else { $false }
        $newIsHidden = if ($newVisual.PSObject.Properties['isHidden']) { [bool]$newVisual.isHidden } else { $false }

        if ($oldIsHidden -ne $newIsHidden) {
            # Extract field name for proper display name resolution
            $fieldName = ""
            if ($newVisual.visual.PSObject.Properties['query'] -and
                $newVisual.visual.query.PSObject.Properties['queryState'] -and
                $newVisual.visual.query.queryState.PSObject.Properties['Values'] -and
                $newVisual.visual.query.queryState.Values.PSObject.Properties['projections']) {
                $projections = $newVisual.visual.query.queryState.Values.projections
                if ($projections -is [Array] -and $projections.Count -gt 0 -and
                    $projections[0].PSObject.Properties['field'] -and
                    $projections[0].field.PSObject.Properties['Column'] -and
                    $projections[0].field.Column.PSObject.Properties['Property']) {
                    $fieldName = $projections[0].field.Column.Property
                }
            }

            # Use Resolve-SlicerDisplayName to get the proper display name (like other sync differences)
            $actualDisplayName = $null
            if ($newVisual) {
                $actualDisplayName = Resolve-SlicerDisplayName -visual $newVisual -fallbackField $fieldName -fallbackGroup "" -visualId $visualName
            } elseif ($oldVisual) {
                $actualDisplayName = Resolve-SlicerDisplayName -visual $oldVisual -fallbackField $fieldName -fallbackGroup "" -visualId $visualName
            }

            # Use the actual display name if found, otherwise fallback to field name
            $slicerDisplayName = $visualDisplayName
            if ($actualDisplayName -and $actualDisplayName -ne $fieldName) {
                $slicerDisplayName = "Slicer [$actualDisplayName]"
            } elseif ($fieldName -ne "") {
                $slicerDisplayName = "Slicer [$fieldName]"
            }

            # Extract syncGroup to find the source page
            $syncGroup = $null
            $syncGroupName = ""
            if ($newVisual -and $newVisual.PSObject.Properties['visual'] -and
                $newVisual.visual.PSObject.Properties['syncGroup']) {
                $syncGroup = $newVisual.visual.syncGroup
                $syncGroupName = if ($syncGroup.PSObject.Properties['groupName']) { $syncGroup.groupName } else { $fieldName }
            }

            # Find the SOURCE page (the page that has the slicer without isHidden=true)
            # This is the page where the visibility setting is controlled
            $sourcePageId = $null
            $sourcePageDisplayName = $null
            $sourceVisualId = $null

            if ($syncGroupName -and $allNewPages) {
                foreach ($pageKey in $allNewPages.Keys) {
                    $pageData = $allNewPages[$pageKey]
                    if ($pageData.PSObject.Properties['visuals']) {
                        foreach ($visKey in $pageData.visuals.Keys) {
                            $vis = $pageData.visuals[$visKey]
                            # Check if this visual has the same syncGroup
                            if ($vis -and $vis.PSObject.Properties['visual'] -and
                                $vis.visual.PSObject.Properties['syncGroup']) {
                                $visSyncGroup = $vis.visual.syncGroup
                                $visSyncGroupName = if ($visSyncGroup.PSObject.Properties['groupName']) { $visSyncGroup.groupName } else { "" }

                                # If same syncGroup and NOT hidden (or no isHidden property)
                                if ($visSyncGroupName -eq $syncGroupName) {
                                    $visIsHidden = if ($vis.PSObject.Properties['isHidden']) { [bool]$vis.isHidden } else { $false }
                                    if (-not $visIsHidden) {
                                        # This is a candidate for SOURCE page
                                        $sourcePageId = $pageKey
                                        $sourcePageDisplayName = if ($pageData.PSObject.Properties['displayName']) { $pageData.displayName } else { $pageKey }
                                        $sourceVisualId = $visKey
                                        break
                                    }
                                }
                            }
                        }
                        if ($sourcePageId) { break }
                    }
                }
            }

            # If no source page found, fallback to current page
            if (-not $sourcePageId) {
                $sourcePageId = $pageName
                $sourcePageDisplayName = $pageDisplayName
                $sourceVisualId = $visualName
            }

            # Store this difference with syncGroupName for later processing in Analyze-SyncGroupCauses
            $diff = [ReportDifference]::new()
            $diff.ElementType = "Synchronisation"
            $diff.ElementName = $visualName
            $diff.ElementDisplayName = "$slicerDisplayName sur page '$pageDisplayName'"
            $diff.ElementPath = "pages/$pageName/visuals/$visualName"
            $diff.PropertyName = "isHidden"
            $diff.DifferenceType = "Modifie"
            $diff.OldValue = if ($oldIsHidden) { "Caché" } else { "Visible" }
            $diff.NewValue = if ($newIsHidden) { "Caché" } else { "Visible" }
            $diff.HierarchyLevel = "pages.visuals.visibility"
            $diff.ParentElementName = $pageName
            $diff.ParentDisplayName = $pageDisplayName
            $diff.SyncGroupName = $syncGroupName

            # Store additional info for CAUSE → CONSEQUENCE analysis
            $diff.AdditionalInfo = @{
                FieldName = $fieldName
                SyncGroupName = $syncGroupName
                SourcePageId = $sourcePageId
                SourcePageDisplayName = $sourcePageDisplayName
                SourceVisualId = $sourceVisualId
                TargetPageId = $pageName
                TargetPageDisplayName = $pageDisplayName
                TargetVisualId = $visualName
                SlicerDisplayName = $slicerDisplayName
                IsHiding = $newIsHidden
            } | ConvertTo-Json -Compress

            $differences += $diff

            $visibilityChange = if ($newIsHidden) { "masqué" } else { "affiché" }
            Write-Host "        Visibilité modifiée: $slicerDisplayName est maintenant $visibilityChange sur page '$pageDisplayName' (source: '$sourcePageDisplayName')" -ForegroundColor $(if ($newIsHidden) { "Red" } else { "Green" })
        }

        # ===== COMPARAISON DE LA SYNCHRONISATION DES SLICERS (SYNC SLICER) =====
        # Nouvelle implémentation basée sur la détection au niveau du groupe

        # Ne comparer la synchronisation que pour les slicers
        if ($visualType -eq "slicer") {
            # Extraire le nom du champ pour un affichage plus clair
            $fieldName = ""
            if ($newVisual.visual.PSObject.Properties['query'] -and
                $newVisual.visual.query.PSObject.Properties['queryState'] -and
                $newVisual.visual.query.queryState.PSObject.Properties['Values'] -and
                $newVisual.visual.query.queryState.Values.PSObject.Properties['projections']) {
                $projections = $newVisual.visual.query.queryState.Values.projections
                if ($projections -is [Array] -and $projections.Count -gt 0 -and
                    $projections[0].PSObject.Properties['field'] -and
                    $projections[0].field.PSObject.Properties['Column'] -and
                    $projections[0].field.Column.PSObject.Properties['Property']) {
                    $fieldName = $projections[0].field.Column.Property
                }
            }

            # Fallback to old visual if new doesn't have field name
            if ([string]::IsNullOrWhiteSpace($fieldName) -and $oldVisual.visual.PSObject.Properties['query']) {
                if ($oldVisual.visual.query.PSObject.Properties['queryState'] -and
                    $oldVisual.visual.query.queryState.PSObject.Properties['Values'] -and
                    $oldVisual.visual.query.queryState.Values.PSObject.Properties['projections']) {
                    $projections = $oldVisual.visual.query.queryState.Values.projections
                    if ($projections -is [Array] -and $projections.Count -gt 0 -and
                        $projections[0].PSObject.Properties['field'] -and
                        $projections[0].field.PSObject.Properties['Column'] -and
                        $projections[0].field.Column.PSObject.Properties['Property']) {
                        $fieldName = $projections[0].field.Column.Property
                    }
                }
            }

            # Améliorer le nom d'affichage du slicer en utilisant le display name du header
            $slicerDisplayName = $visualDisplayName

            # Try to extract the actual display name (header text) from the visual
            $actualDisplayName = $null
            if ($newVisual) {
                $actualDisplayName = Resolve-SlicerDisplayName -visual $newVisual -fallbackField $fieldName -fallbackGroup "" -visualId $visualName
            } elseif ($oldVisual) {
                $actualDisplayName = Resolve-SlicerDisplayName -visual $oldVisual -fallbackField $fieldName -fallbackGroup "" -visualId $visualName
            }

            # Use the actual display name if found, otherwise fallback to field name
            if ($actualDisplayName -and $actualDisplayName -ne $fieldName) {
                $slicerDisplayName = "Slicer [$actualDisplayName]"
            } elseif ($slicerDisplayName -like "slicer (*" -and $fieldName -ne "") {
                $slicerDisplayName = "Slicer [$fieldName]"
            }

            # Compare synchronization using group-based approach
            # Get all sync group members for this field in both versions
            if (-not [string]::IsNullOrWhiteSpace($fieldName)) {
                # CRITICAL FIX: Get current visual's groupName to filter by group
                $currentGroupName = $fieldName  # Default fallback
                if ($newVisual.visual.PSObject.Properties['syncGroup'] -and
                    $newVisual.visual.syncGroup.PSObject.Properties['groupName']) {
                    $currentGroupName = $newVisual.visual.syncGroup.groupName
                } elseif ($oldVisual.visual.PSObject.Properties['syncGroup'] -and
                          $oldVisual.visual.syncGroup.PSObject.Properties['groupName']) {
                    $currentGroupName = $oldVisual.visual.syncGroup.groupName
                }

                $allOldMembers = Get-AllSyncGroups -allPages $allOldPages -fieldName $fieldName
                $allNewMembers = Get-AllSyncGroups -allPages $allNewPages -fieldName $fieldName

                # CRITICAL FIX: Filter by groupName to avoid mixing different sync groups with same field name
                $oldGroupMembers = @{}
                foreach ($pageId in $allOldMembers.Keys) {
                    $member = $allOldMembers[$pageId]
                    if ($member.GroupName -eq $currentGroupName) {
                        $oldGroupMembers[$pageId] = $member
                    }
                }

                $newGroupMembers = @{}
                foreach ($pageId in $allNewMembers.Keys) {
                    $member = $allNewMembers[$pageId]
                    if ($member.GroupName -eq $currentGroupName) {
                        $newGroupMembers[$pageId] = $member
                    }
                }

                # Compare the groups (now filtered to only matching groupName)
                $syncComparison = Compare-SyncGroupMembers -oldMembers $oldGroupMembers -newMembers $newGroupMembers

                if ($syncComparison.HasChanges) {
                    # Build descriptions
                    $oldDescription = if ($syncComparison.OldMemberCount -gt 0) {
                        "$($syncComparison.OldMemberCount) page(s) synchronisée(s)"
                    } else {
                        "Aucune synchronisation"
                    }

                    $newDescription = if ($syncComparison.NewMemberCount -gt 0) {
                        "$($syncComparison.NewMemberCount) page(s) synchronisée(s)"
                    } else {
                        "Aucune synchronisation"
                    }

                    # Build changed pages summary
                    $changedPages = @()
                    foreach ($addedPage in $syncComparison.AddedPages) {
                        $syncInfo = ""
                        if ($addedPage.FieldChanges) { $syncInfo += "sync: ✓" } else { $syncInfo += "sync: ✗" }
                        if ($addedPage.FilterChanges) { $syncInfo += ", visible: ✓" } else { $syncInfo += ", visible: ✗" }
                        $changedPages += "'$($addedPage.PageDisplayName)' ajoutée ($syncInfo)"
                    }
                    foreach ($removedPage in $syncComparison.RemovedPages) {
                        $changedPages += "'$($removedPage.PageDisplayName)' retirée"
                    }
                    foreach ($modifiedPage in $syncComparison.ModifiedPages) {
                        $changes = @()
                        if ($modifiedPage.OldFieldChanges -ne $modifiedPage.NewFieldChanges) {
                            $changes += if ($modifiedPage.NewFieldChanges) { "sync activée" } else { "sync désactivée" }
                        }
                        if ($modifiedPage.OldFilterChanges -ne $modifiedPage.NewFilterChanges) {
                            $changes += if ($modifiedPage.NewFilterChanges) { "visible activée" } else { "visible désactivée" }
                        }
                        if ($changes.Count -gt 0) {
                            $changedPages += "'$($modifiedPage.PageDisplayName)' (" + ($changes -join ", ") + ")"
                        }
                    }

                    # Use the currentGroupName already determined above (filtered groupName)
                    $groupName = $currentGroupName

                    # Create main synchronization difference
                    $diff = [ReportDifference]::new()
                    $diff.ElementType = "Synchronisation"
                    $diff.ElementName = $visualName
                    $diff.ElementDisplayName = "$slicerDisplayName sur page '$pageDisplayName'"
                    $diff.ElementPath = "pages/$pageName/visuals/$visualName/syncGroup"
                    $diff.PropertyName = "syncGroup"
                    $diff.DifferenceType = "Modifie"
                    $diff.OldValue = $oldDescription
                    $diff.NewValue = $newDescription
                    $diff.HierarchyLevel = "pages.visuals.synchronisation"
                    $diff.SyncGroupName = $groupName

                    # Detailed message about changes
                    if ($changedPages.Count -gt 0) {
                        $changesDetail = $changedPages -join "; "
                        $diff.AdditionalInfo = "Slicer [$fieldName] sur page '$pageDisplayName' - Modifications de synchronisation : $changesDetail"
                    } else {
                        $diff.AdditionalInfo = "Slicer [$fieldName] sur page '$pageDisplayName' - Synchronisation modifiée"
                    }

                    $diff.ParentElementName = $pageName
                    $diff.ParentDisplayName = $pageDisplayName

                    # Store the structured comparison for later use by Analyze-SyncGroupCauses
                    $diff | Add-Member -MemberType NoteProperty -Name '_syncComparison' -Value $syncComparison -Force

                    $differences += $diff

                    Write-Host "        Synchronisation modifiée pour $slicerDisplayName sur page '$pageDisplayName'" -ForegroundColor Yellow
                    if ($changedPages.Count -gt 0) {
                        Write-Host "          Changements: $($changedPages -join "; ")" -ForegroundColor DarkYellow
                    }
                }
            }
        }

        # ===== FIN COMPARAISON SYNCHRONISATION DES SLICERS =====

        # ENHANCED: Button-specific properties comparison
        $newButtonDetails = if ($newVisual._extracted_button_details) { $newVisual._extracted_button_details -join " | " } else { "" }
        $oldButtonDetails = if ($oldVisual._extracted_button_details) { $oldVisual._extracted_button_details -join " | " } else { "" }
        
        if ($newButtonDetails -ne $oldButtonDetails) {
            $diff = [ReportDifference]::new()
            $diff.ElementType = "Bouton"
            $diff.ElementName = $visualName
            $diff.ElementDisplayName = $fullDisplayInfo
            $diff.ElementPath = "pages/$pageName/visuals/$visualName"
            $diff.PropertyName = "buttonProperties"
            $diff.DifferenceType = "Modifie"
            $diff.OldValue = if ($oldButtonDetails) { $oldButtonDetails } else { "Aucune propriété bouton" }
            $diff.NewValue = if ($newButtonDetails) { $newButtonDetails } else { "Aucune propriété bouton" }
            $diff.HierarchyLevel = "pages.visuals.buttonProperties"
            $diff.AdditionalInfo = "Modification des propriétés spécifiques du bouton actionButton"
            $diff.ParentElementName = $pageName
            $diff.ParentDisplayName = $pageDisplayName
            $differences += $diff
            Write-Host "        Propriétés du bouton modifiées: $fullDisplayInfo" -ForegroundColor Magenta
            Write-Host "          Ancien: $oldButtonDetails" -ForegroundColor DarkYellow
            Write-Host "          Nouveau: $newButtonDetails" -ForegroundColor DarkYellow
        }
    }
    catch {
        Write-Warning "Erreur lors de la comparaison des proprietes du visuel $visualName : $($_.Exception.Message)"
    }
    
    return $differences
}

# Enhanced filter details extraction function V2
Function ExtractFilterDetails {
    param ([PSCustomObject] $visual)
    
    $result = @{
        HasFilters = $false
        FilterSummary = ""
        FilterDetails = ""
        Filters = @()
    }
    
    try {
        if (-not $visual -or -not $visual.visual) {
            return $result
        }
        
        $filterInfos = @()
        
        # 1. NEW: Search in visual.query.queryState.Filters
        if ($visual.visual.query -and $visual.visual.query.queryState -and $visual.visual.query.queryState.Filters) {
            foreach ($filter in $visual.visual.query.queryState.Filters) {
                if ($filter.Filter) {
                    $filterDetail = ParseAdvancedFilterCondition -filter $filter.Filter
                    if ($filterDetail) {
                        $filterInfos += "QueryState: $filterDetail"
                    }
                }
            }
        }
        
        # 2. ENHANCED: Search in visual.objects with extended detection
        if ($visual.visual.objects) {
            $objects = $visual.visual.objects
            
            # Search in general[].properties.filter.filter (existing)
            if ($objects.general) {
                foreach ($generalItem in $objects.general) {
                    if ($generalItem.properties -and $generalItem.properties.filter -and $generalItem.properties.filter.filter) {
                        $filterObj = $generalItem.properties.filter.filter
                        
                        if ($filterObj.Where) {
                            foreach ($whereClause in $filterObj.Where) {
                                if ($whereClause.Condition) {
                                    $filterDetail = ParseFilterCondition -condition $whereClause.Condition
                                    if ($filterDetail) {
                                        $filterInfos += "General: $filterDetail"
                                    }
                                }
                            }
                        }
                    }
                }
            }
            
            # NEW: Search in data[].properties with extended detection
            if ($objects.data) {
                foreach ($dataItem in $objects.data) {
                    if ($dataItem.properties) {
                        foreach ($prop in $dataItem.properties.PSObject.Properties) {
                            if ($prop.Name -like "*filter*" -or $prop.Name -like "*Filter*") {
                                if ($prop.Value -and $prop.Value -ne $null) {
                                    try {
                                        $propValue = if ($prop.Value.GetType().Name -eq "PSCustomObject") {
                                            ($prop.Value | ConvertTo-Json -Depth 2 -Compress).Substring(0, [Math]::Min(100, ($prop.Value | ConvertTo-Json -Depth 2 -Compress).Length))
                                        } else {
                                            $prop.Value.ToString().Substring(0, [Math]::Min(100, $prop.Value.ToString().Length))
                                        }
                                        $filterInfos += "Data.$($prop.Name): $propValue"
                                    } catch {
                                        $filterInfos += "Data.$($prop.Name): (valeur complexe)"
                                    }
                                }
                            }
                        }
                    }
                }
            }
            
            # NEW: Search in all other objects for filter properties
            foreach ($objectType in $objects.PSObject.Properties) {
                if ($objectType.Name -notin @("general", "data")) {
                    if ($objectType.Value -and $objectType.Value.GetType().Name -eq "Object[]") {
                        foreach ($item in $objectType.Value) {
                            if ($item.properties) {
                                foreach ($prop in $item.properties.PSObject.Properties) {
                                    if ($prop.Name -like "*filter*" -or $prop.Name -like "*Filter*") {
                                        if ($prop.Value -and $prop.Value -ne $null) {
                                            $filterInfos += "$($objectType.Name).$($prop.Name): (detecte)"
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
        
        # 3. NEW: Search in visual.drillFilterOtherVisuals
        if ($visual.visual.drillFilterOtherVisuals) {
            $filterInfos += "DrillFilter: Active"
        }
        
        # 4. NEW: Search in visual.vcObjects for visual filters
        if ($visual.visual.vcObjects) {
            foreach ($vcObj in $visual.visual.vcObjects.PSObject.Properties) {
                if ($vcObj.Name -like "*filter*" -or $vcObj.Name -like "*Filter*") {
                    $filterInfos += "VC.$($vcObj.Name): (detecte)"
                }
            }
        }
        
        if ($filterInfos.Count -gt 0) {
            $result.HasFilters = $true
            $result.FilterSummary = ($filterInfos -join " | ")
            $result.FilterDetails = ($filterInfos -join "; ")
            $result.Filters = $filterInfos
        }
    }
    catch {
        Write-Warning "Erreur lors de l'extraction des filtres: $($_.Exception.Message)"
    }
    
    return $result
}

# NEW FUNCTION: Extract button-specific details
Function ExtractButtonDetails {
    param ([PSCustomObject] $visual)
    
    $result = @{
        HasButtonDetails = $false
        ButtonDetails = @()
        ButtonSummary = ""
        IconType = ""
        NavigationType = ""
        NavigationTarget = ""
        WebUrl = ""
    }
    
    try {
        if (-not $visual -or -not $visual.visual -or $visual.visual.visualType -ne "actionButton") {
            return $result
        }
        
        $details = @()
        
        # Extract icon information
        if ($visual.visual.objects -and $visual.visual.objects.icon) {
            foreach ($iconItem in $visual.visual.objects.icon) {
                if ($iconItem.properties -and $iconItem.properties.shapeType -and $iconItem.properties.shapeType.expr -and $iconItem.properties.shapeType.expr.Literal) {
                    $iconType = $iconItem.properties.shapeType.expr.Literal.Value -replace "'", ""
                    $result.IconType = $iconType
                    $details += "Icône: $iconType"
                }
            }
        }
        
        # Extract navigation/link information
        if ($visual.visual.visualContainerObjects -and $visual.visual.visualContainerObjects.visualLink) {
            foreach ($linkItem in $visual.visual.visualContainerObjects.visualLink) {
                if ($linkItem.properties) {
                    # Check if link is enabled
                    if ($linkItem.properties.show -and $linkItem.properties.show.expr -and $linkItem.properties.show.expr.Literal -and $linkItem.properties.show.expr.Literal.Value -eq "true") {
                        $details += "Lien activé"
                        
                        # Extract link type
                        if ($linkItem.properties.type -and $linkItem.properties.type.expr -and $linkItem.properties.type.expr.Literal) {
                            $navType = $linkItem.properties.type.expr.Literal.Value -replace "'", ""
                            $result.NavigationType = $navType
                            $details += "Type: $navType"
                        }
                        
                        # Extract web URL
                        if ($linkItem.properties.webUrl -and $linkItem.properties.webUrl.expr -and $linkItem.properties.webUrl.expr.Literal) {
                            $webUrl = $linkItem.properties.webUrl.expr.Literal.Value -replace "'", ""
                            $result.WebUrl = $webUrl
                            $details += "URL: $webUrl"
                        }
                        
                        # Extract navigation section/page
                        if ($linkItem.properties.navigationSection -and $linkItem.properties.navigationSection.expr -and $linkItem.properties.navigationSection.expr.Literal) {
                            $navTarget = $linkItem.properties.navigationSection.expr.Literal.Value -replace "'", ""
                            $result.NavigationTarget = $navTarget
                            $details += "Cible navigation: $navTarget"
                        }
                    }
                }
            }
        }
        
        # Extract position and size
        if ($visual.position) {
            $position = "Position: ($($visual.position.x), $($visual.position.y))"
            $size = "Taille: $($visual.position.width) x $($visual.position.height)"
            $details += $position
            $details += $size
        }
        
        # Extract creation method
        if ($visual.howCreated) {
            $details += "Méthode création: $($visual.howCreated)"
        }
        
        if ($details.Count -gt 0) {
            $result.HasButtonDetails = $true
            $result.ButtonDetails = $details
            $result.ButtonSummary = ($details -join " | ")
        }
    }
    catch {
        Write-Warning "Erreur lors de l'extraction des détails du bouton: $($_.Exception.Message)"
    }
    
    return $result
}

# Enhanced filter condition analysis function
Function Get-ComparisonFieldName {
    param([PSCustomObject] $expression)

    if ($null -eq $expression) {
        return 'Champ'
    }

    if ($expression.Measure -and $expression.Measure.Property) {
        return [string]$expression.Measure.Property
    }

    if ($expression.Column -and $expression.Column.Property) {
        return [string]$expression.Column.Property
    }

    if ($expression.SourceRef -and $expression.SourceRef.Entity) {
        return [string]$expression.SourceRef.Entity
    }

    return 'Champ'
}

Function Get-ComparisonValueText {
    param([PSCustomObject] $value)

    if ($null -eq $value) {
        return 'null'
    }

    if ($value.Literal -and $value.Literal.Value) {
        $raw = [string]$value.Literal.Value
        if ($raw -eq 'null') {
            return 'null'
        }
        return $raw.Trim("'")
    }

    if ($value.PSObject.Properties['Value']) {
        return [string]$value.Value
    }

    return 'valeur'
}

Function Get-ComparisonOperator {
    param([object] $kind)

    switch ([int]$kind) {
        0 { return '=' }
        1 { return '<>' }
        2 { return '>' }
        3 { return '>=' }
        4 { return '<' }
        5 { return '<=' }
        default { return '=' }
    }
}

Function Get-NegatedOperator {
    param([string] $operator)

    switch ($operator) {
        '=' { return '<>' }
        '<>' { return '=' }
        '>' { return '<=' }
        '>=' { return '<' }
        '<' { return '>=' }
        '<=' { return '>' }
        default { return "NOT $operator" }
    }
}

Function Format-ComparisonExpression {
    param(
        [PSCustomObject] $comparison,
        [bool] $negate
    )

    if (-not $comparison) {
        return $null
    }

    $field = Get-ComparisonFieldName $comparison.Left
    $value = Get-ComparisonValueText $comparison.Right
    $operator = Get-ComparisonOperator $comparison.ComparisonKind

    if ($negate) {
        if ($operator -eq '=' -and $value -eq 'null') {
            return "$field IS NOT NULL"
        }
        $operator = Get-NegatedOperator $operator
    }

    if ($operator -eq '=' -and $value -eq 'null') {
        return "$field IS NULL"
    }

    if ($value -match '^[A-Za-z0-9 _-]+$') {
        return "$field $operator $value"
    }

    return "$field $operator '$value'"
}

Function ParseFilterCondition {
    param ([PSCustomObject] $condition)
    
    try {
        if ($condition.In) {
            # "In" condition (contains)
            $field = ""
            $values = @()
            
            if ($condition.In.Expressions -and $condition.In.Expressions.Count -gt 0) {
                $expr = $condition.In.Expressions[0]
                if ($expr.Column -and $expr.Column.Property) {
                    $field = $expr.Column.Property
                }
            }
            
            if ($condition.In.Values -and $condition.In.Values.Count -gt 0) {
                foreach ($valueArray in $condition.In.Values) {
                    if ($valueArray -and $valueArray.Count -gt 0) {
                        $value = $valueArray[0]
                        if ($value.Literal -and $value.Literal.Value) {
                            $values += $value.Literal.Value
                        }
                    }
                }
            }
            
            $valuesText = if ($values.Count -gt 0) { ($values -join ", ") } else { "(valeurs non extraites)" }
            return "$field IN [$valuesText]"
        }
        
        if ($condition.Not) {
            $inner = $condition.Not
            if ($inner.Expression) {
                $inner = $inner.Expression
            }

            if ($inner.Comparison) {
                $formatted = Format-ComparisonExpression -comparison $inner.Comparison -negate $true
                if ($formatted) {
                    return $formatted
                }
            }

            $innerCondition = ParseFilterCondition -condition $inner
            return "NOT ($innerCondition)"
        }

        if ($condition.And) {
            $left = ParseFilterCondition -condition $condition.And.Left
            $right = ParseFilterCondition -condition $condition.And.Right
            return "$left AND $right"
        }

        if ($condition.Or) {
            $left = ParseFilterCondition -condition $condition.Or.Left
            $right = ParseFilterCondition -condition $condition.Or.Right
            return "$left OR $right"
        }

        if ($condition.Expression) {
            return ParseFilterCondition -condition $condition.Expression
        }

        if ($condition.Compare) {
            $formatted = Format-ComparisonExpression -comparison $condition.Compare -negate $false
            if ($formatted) {
                return $formatted
            }
        }

        if ($condition.Comparison) {
            $formatted = Format-ComparisonExpression -comparison $condition.Comparison -negate $false
            if ($formatted) {
                return $formatted
            }
        }
        
        # NEW: Support for Between
        if ($condition.Between) {
            $field = ""
            $lowerValue = ""
            $upperValue = ""
            
            if ($condition.Between.Expression -and $condition.Between.Expression.Column -and $condition.Between.Expression.Column.Property) {
                $field = $condition.Between.Expression.Column.Property
            }
            
            if ($condition.Between.LowerBound -and $condition.Between.LowerBound.Literal) {
                $lowerValue = $condition.Between.LowerBound.Literal.Value
            }
            
            if ($condition.Between.UpperBound -and $condition.Between.UpperBound.Literal) {
                $upperValue = $condition.Between.UpperBound.Literal.Value
            }
            
            return "$field BETWEEN $lowerValue AND $upperValue"
        }
        
        # NEW: Support for IsEmpty/IsNotEmpty
        if ($condition.IsEmpty) {
            $field = ""
            if ($condition.IsEmpty.Expression -and $condition.IsEmpty.Expression.Column -and $condition.IsEmpty.Expression.Column.Property) {
                $field = $condition.IsEmpty.Expression.Column.Property
            }
            return "$field IS EMPTY"
        }
        
        if ($condition.IsNotEmpty) {
            $field = ""
            if ($condition.IsNotEmpty.Expression -and $condition.IsNotEmpty.Expression.Column -and $condition.IsNotEmpty.Expression.Column.Property) {
                $field = $condition.IsNotEmpty.Expression.Column.Property
            }
            return "$field IS NOT EMPTY"
        }
        
        # Other condition types
        return "Condition complexe (type: $($condition.PSObject.Properties.Name -join ', '))"
    }
    catch {
        return "Condition non analysable"
    }
}

# NEW function to analyze QueryState filters
Function ParseAdvancedFilterCondition {
    param ([PSCustomObject] $filter)
    
    try {
        if (-not $filter) {
            return $null
        }
        
        # Extract target field
        $field = ""
        if ($filter.Expression -and $filter.Expression.SourceRef -and $filter.Expression.SourceRef.Source) {
            $field = $filter.Expression.SourceRef.Source
        } elseif ($filter.Expression -and $filter.Expression.Column -and $filter.Expression.Column.Property) {
            $field = $filter.Expression.Column.Property
        }
        
        # Analyze filter type
        $effectiveFilter = if ($filter.Filter) { $filter.Filter } else { $filter }

        if ($effectiveFilter -and $effectiveFilter.Where) {
            $conditions = @()
            foreach ($whereClause in $effectiveFilter.Where) {
                if ($whereClause.Condition) {
                    $conditionDetail = ParseFilterCondition -condition $whereClause.Condition
                    if ($conditionDetail) {
                        $conditions += $conditionDetail
                    }
                }
            }
            
            if ($conditions.Count -gt 0) {
                return ($conditions -join " AND ")
            }
        }
        
        # Fallback: try to analyze structure directly
        if ($effectiveFilter.Values -and $effectiveFilter.Values.Count -gt 0) {
            $values = @()
            foreach ($value in $effectiveFilter.Values) {
                if ($value -and $value.Value) {
                    $values += $value.Value
                }
            }
            
            if ($values.Count -gt 0) {
                return "$field IN [" + ($values -join ", ") + "]"
            }
        }
        
        return "$field (filtre complexe)"
    }
    catch {
        return "Filtre QueryState non analysable"
    }
}

# Bookmark comparison function
Function Get-PageDisplayName {
    param(
        [string] $pageId,
        [hashtable[]] $pageMaps
    )

    if ([string]::IsNullOrWhiteSpace($pageId)) {
        return ''
    }

    foreach ($map in $pageMaps) {
        if ($map -and $map.ContainsKey($pageId)) {
            $page = $map[$pageId]
            if ($page -and $page.displayName) {
                return [string]$page.displayName
            }
        }
    }

    return $pageId
}

Function Get-BookmarkSummary {
    param(
        [PSCustomObject] $bookmark,
        [hashtable[]] $pageMaps
    )

    if (-not $bookmark) {
        return ''
    }

    $summaryParts = New-Object System.Collections.Generic.List[string]

    $options = $bookmark.options
    if ($options) {
        $targetCount = 0
        if ($options.targetVisualNames) {
            $targetCount = (@($options.targetVisualNames | Where-Object { $_ -and $_.ToString().Trim().Length -gt 0 })).Count
        }
        if ($targetCount -gt 0) {
            $targetLabel = if ($options.applyOnlyToTargetVisuals) { "Ciblage isolé" } else { "Ciblage" }
            $summaryParts.Add("${targetLabel}: $targetCount visuel(s)")
        }
        if ($options.suppressData) {
            $summaryParts.Add("Données figées")
        }
    }

    $exploration = $bookmark.explorationState
    if ($exploration) {
        $activeSectionKey = if ($exploration.activeSection) { [string]$exploration.activeSection } else { '' }
        if ($activeSectionKey) {
            $activeName = Get-PageDisplayName -pageId $activeSectionKey -pageMaps $pageMaps
            $summaryParts.Add("Section active: $activeName")
        }

        $groupHidden = 0
        $groupVisible = 0
        if ($exploration.sections) {
            foreach ($sectionProp in $exploration.sections.PSObject.Properties) {
                $sectionValue = $sectionProp.Value
                if ($sectionValue -and $sectionValue.visualContainerGroups) {
                    foreach ($groupProp in $sectionValue.visualContainerGroups.PSObject.Properties) {
                        $groupState = $groupProp.Value
                        $isHidden = $false
                        if ($groupState -and $groupState.PSObject.Properties['isHidden']) {
                            $isHidden = [bool]$groupState.isHidden
                        }
                        if ($isHidden) {
                            $groupHidden++
                        }
                        else {
                            $groupVisible++
                        }
                    }
                }
            }
        }
        if ($groupVisible -gt 0 -or $groupHidden -gt 0) {
            $summaryParts.Add("Groupes visibles: $groupVisible")
            if ($groupHidden -gt 0) {
                $summaryParts.Add("Groupes masqués: $groupHidden")
            }
        }

        if ($exploration.objects -and $exploration.objects.merge -and $exploration.objects.merge.outspacePane) {
            $expandedStates = @()
            foreach ($pane in $exploration.objects.merge.outspacePane) {
                if ($pane -and $pane.properties) {
                    if ($pane.properties.expanded) {
                        $expandedStates += 'ouvert'
                    }
                    elseif ($pane.properties.visible) {
                        $expandedStates += 'visible'
                    }
                }
            }
            if ($expandedStates.Count -gt 0) {
                $summaryParts.Add("Volets: " + ($expandedStates -join ', '))
            }
        }
    }

    if ($summaryParts.Count -eq 0) {
        return 'Configuration inchangée'
    }

    return ($summaryParts -join ' | ')
}

Function CompareBookmarksData {
    param (
        [hashtable] $newBookmarks,
        [hashtable] $oldBookmarks,
        [PSCustomObject] $newMetadata,
        [PSCustomObject] $oldMetadata,
        [hashtable] $newPages,
        [hashtable] $oldPages
    )
    
    $differences = @()
    
    if ($null -eq $newBookmarks) { $newBookmarks = @{} }
    if ($null -eq $oldBookmarks) { $oldBookmarks = @{} }

    $pageMaps = @()
    if ($newPages) { $pageMaps += ,$newPages }
    if ($oldPages) { $pageMaps += ,$oldPages }
    
    # Get all bookmarks from both versions
    $allBookmarkNames = @()
    $allBookmarkNames += $newBookmarks.Keys
    $allBookmarkNames += $oldBookmarks.Keys
    $allBookmarkNames = $allBookmarkNames | Select-Object -Unique
    
    foreach ($bookmarkName in $allBookmarkNames) {
        $newBookmark = $newBookmarks[$bookmarkName]
        $oldBookmark = $oldBookmarks[$bookmarkName]
        
        # Deleted bookmark
        if ($oldBookmark -and -not $newBookmark) {
            $diff = [ReportDifference]::new()
            $diff.ElementType = "Signet"
            $diff.ElementName = $bookmarkName
            $diff.ElementDisplayName = if ($oldBookmark.displayName) { [string]$oldBookmark.displayName } else { $bookmarkName }
            $diff.ElementPath = "bookmarks/$bookmarkName"
            $diff.PropertyName = "existence"
            $diff.DifferenceType = "Supprime"
            $diff.OldValue = "Present"
            $diff.NewValue = "Absent"
            $diff.HierarchyLevel = "bookmarks"
            $differences += $diff
            Write-Host "    Signet supprime: $bookmarkName" -ForegroundColor Red
        }
        
        # Added bookmark
        if ($newBookmark -and -not $oldBookmark) {
            $diff = [ReportDifference]::new()
            $diff.ElementType = "Signet"
            $diff.ElementName = $bookmarkName
            $diff.ElementPath = "bookmarks/$bookmarkName"
            $diff.PropertyName = "existence"
            $diff.DifferenceType = "Ajoute"
            $diff.OldValue = "Absent"
            $diff.NewValue = "Present"
            $diff.HierarchyLevel = "bookmarks"
            $diff.ElementDisplayName = if ($newBookmark.displayName) { [string]$newBookmark.displayName } else { $bookmarkName }
            $differences += $diff
            Write-Host "    Signet ajoute: $bookmarkName" -ForegroundColor Green
        }

        if ($newBookmark -and $oldBookmark) {
            $displayNameNew = if ($newBookmark.displayName) { [string]$newBookmark.displayName } else { $bookmarkName }
            $displayNameOld = if ($oldBookmark.displayName) { [string]$oldBookmark.displayName } else { $bookmarkName }

            if ($displayNameNew -ne $displayNameOld) {
                $diff = [ReportDifference]::new()
                $diff.ElementType = "Signet"
                $diff.ElementName = $bookmarkName
                $diff.ElementDisplayName = $displayNameOld
                $diff.ElementPath = "bookmarks/$bookmarkName"
                $diff.PropertyName = "displayName"
                $diff.DifferenceType = "Modifie"
                $diff.OldValue = $displayNameOld
                $diff.NewValue = $displayNameNew
                $diff.HierarchyLevel = "bookmarks"
                $diff.AdditionalInfo = "Nom du signet mis à jour"
                $differences += $diff
                Write-Host "    Signet renomme: $displayNameOld -> $displayNameNew" -ForegroundColor Yellow
            }

            $oldSummary = Get-BookmarkSummary -bookmark $oldBookmark -pageMaps $pageMaps
            $newSummary = Get-BookmarkSummary -bookmark $newBookmark -pageMaps $pageMaps

            if ($newSummary -ne $oldSummary) {
                $diff = [ReportDifference]::new()
                $diff.ElementType = "Signet"
                $diff.ElementName = $bookmarkName
                $diff.ElementDisplayName = $displayNameNew
                $diff.ElementPath = "bookmarks/$bookmarkName"
                $diff.PropertyName = "configuration"
                $diff.DifferenceType = "Modifie"
                $diff.OldValue = $oldSummary
                $diff.NewValue = $newSummary
                $diff.HierarchyLevel = "bookmarks"
                $diff.AdditionalInfo = "Ancien: $oldSummary | Nouveau: $newSummary"
                $differences += $diff
                Write-Host "    Signet modifie: $displayNameNew" -ForegroundColor Yellow
            }
        }
    }
    
    $metadataDiffs = CompareBookmarkMetadata -newMetadata $newMetadata -oldMetadata $oldMetadata -newBookmarks $newBookmarks -oldBookmarks $oldBookmarks
    if ($metadataDiffs -and $metadataDiffs.Count -gt 0) {
        $differences += $metadataDiffs
    }

    return $differences
}

Function CompareBookmarkMetadata {
    param (
        [PSCustomObject] $newMetadata,
        [PSCustomObject] $oldMetadata,
        [hashtable] $newBookmarks,
        [hashtable] $oldBookmarks
    )

    $differences = @()

    $newItems = @()
    $oldItems = @()

    if ($newMetadata -and $newMetadata.PSObject.Properties['items']) {
        $newItems = @($newMetadata.items)
    }

    if ($oldMetadata -and $oldMetadata.PSObject.Properties['items']) {
        $oldItems = @($oldMetadata.items)
    }

    if (($newItems.Count -eq 0) -and ($oldItems.Count -eq 0)) {
        return $differences
    }

    $newIndex = @{}
    foreach ($item in $newItems) {
        if ($item -and $item.name) {
            $newIndex[$item.name] = $item
        }
    }

    $oldIndex = @{}
    foreach ($item in $oldItems) {
        if ($item -and $item.name) {
            $oldIndex[$item.name] = $item
        }
    }

    $allGroupNames = @()
    $allGroupNames += $newIndex.Keys
    $allGroupNames += $oldIndex.Keys
    $allGroupNames = $allGroupNames | Select-Object -Unique

    foreach ($groupName in $allGroupNames) {
        $newGroup = $newIndex[$groupName]
        $oldGroup = $oldIndex[$groupName]

        $newDisplay = if ($newGroup -and $newGroup.displayName) { [string]$newGroup.displayName } elseif ($oldGroup -and $oldGroup.displayName) { [string]$oldGroup.displayName } else { $groupName }
        $oldDisplay = if ($oldGroup -and $oldGroup.displayName) { [string]$oldGroup.displayName } elseif ($newGroup -and $newGroup.displayName) { [string]$newGroup.displayName } else { $groupName }
        $displayLabel = if ($newDisplay) { $newDisplay } else { $oldDisplay }
        $displayLabel = if ($displayLabel) { $displayLabel } else { $groupName }

        $groupPath = "bookmarks/groups/$groupName"

        if ($oldGroup -and -not $newGroup) {
            $diff = [ReportDifference]::new()
            $diff.ElementType = "Signet"
            $diff.ElementName = $groupName
            $diff.ElementDisplayName = "$oldDisplay (groupe)"
            $diff.ElementPath = $groupPath
            $diff.PropertyName = "existence"
            $diff.DifferenceType = "Supprime"
            $diff.OldValue = "Present"
            $diff.NewValue = "Absent"
            $diff.HierarchyLevel = "bookmarks.groups"
            $diff.AdditionalInfo = "Groupe supprime"
            $differences += $diff
            Write-Host "    Groupe de signets supprime: $oldDisplay" -ForegroundColor Red
            continue
        }

        if ($newGroup -and -not $oldGroup) {
            $childIds = @()
            if ($newGroup.children) { $childIds = @($newGroup.children) }
            $childNames = Get-BookmarkChildNames -childIds $childIds -bookmarkLookup $newBookmarks
            $childSummary = if ($childNames.Count -gt 0) { $childNames -join ', ' } else { 'Aucun enfant' }

            $diff = [ReportDifference]::new()
            $diff.ElementType = "Signet"
            $diff.ElementName = $groupName
            $diff.ElementDisplayName = "$newDisplay (groupe)"
            $diff.ElementPath = $groupPath
            $diff.PropertyName = "existence"
            $diff.DifferenceType = "Ajoute"
            $diff.OldValue = "Absent"
            $diff.NewValue = "Present"
            $diff.HierarchyLevel = "bookmarks.groups"
            $diff.AdditionalInfo = "Enfants: $childSummary"
            $differences += $diff
            Write-Host "    Groupe de signets ajoute: $newDisplay" -ForegroundColor Green
            continue
        }

        if ($newGroup -and $oldGroup) {
            if ($newDisplay -ne $oldDisplay) {
                $diff = [ReportDifference]::new()
                $diff.ElementType = "Signet"
                $diff.ElementName = $groupName
                $diff.ElementDisplayName = "$oldDisplay (groupe)"
                $diff.ElementPath = $groupPath
                $diff.PropertyName = "displayName"
                $diff.DifferenceType = "Modifie"
                $diff.OldValue = $oldDisplay
                $diff.NewValue = $newDisplay
                $diff.HierarchyLevel = "bookmarks.groups"
                $diff.AdditionalInfo = "Nom du groupe mis à jour"
                $differences += $diff
                Write-Host "    Groupe de signets renomme: $oldDisplay -> $newDisplay" -ForegroundColor Yellow
            }

            $oldChildren = @()
            if ($oldGroup.children) { $oldChildren = @($oldGroup.children) }
            $newChildren = @()
            if ($newGroup.children) { $newChildren = @($newGroup.children) }

            $oldSequence = ($oldChildren -join '|')
            $newSequence = ($newChildren -join '|')

            if ($oldSequence -ne $newSequence) {
                $childNamesOld = Get-BookmarkChildNames -childIds $oldChildren -bookmarkLookup $oldBookmarks
                $childNamesNew = Get-BookmarkChildNames -childIds $newChildren -bookmarkLookup $newBookmarks

                $removedIds = $oldChildren | Where-Object { $_ -notin $newChildren }
                $addedIds = $newChildren | Where-Object { $_ -notin $oldChildren }

                $removedNames = if ($removedIds.Count -gt 0) { Get-BookmarkChildNames -childIds $removedIds -bookmarkLookup $oldBookmarks } else { @() }
                $addedNames = if ($addedIds.Count -gt 0) { Get-BookmarkChildNames -childIds $addedIds -bookmarkLookup $newBookmarks } else { @() }

                $detailsParts = New-Object System.Collections.Generic.List[string]
                $detailsParts.Add("Ancien: $($childNamesOld -join ', ')") | Out-Null
                $detailsParts.Add("Nouveau: $($childNamesNew -join ', ')") | Out-Null

                if ($addedNames.Count -gt 0) {
                    $detailsParts.Add("Ajoutes: " + ($addedNames -join ', ')) | Out-Null
                }
                if ($removedNames.Count -gt 0) {
                    $detailsParts.Add("Supprimes: " + ($removedNames -join ', ')) | Out-Null
                }
                if (($addedNames.Count -eq 0) -and ($removedNames.Count -eq 0)) {
                    $detailsParts.Add("Ordre mis à jour") | Out-Null
                }

                $diff = [ReportDifference]::new()
                $diff.ElementType = "Signet"
                $diff.ElementName = $groupName
                $diff.ElementDisplayName = "$displayLabel (groupe)"
                $diff.ElementPath = $groupPath
                $diff.PropertyName = "children"
                $diff.DifferenceType = "Modifie"
                $diff.OldValue = $childNamesOld -join ', '
                $diff.NewValue = $childNamesNew -join ', '
                $diff.HierarchyLevel = "bookmarks.groups"
                $diff.AdditionalInfo = ($detailsParts -join ' | ')
                $differences += $diff
                Write-Host "    Groupe de signets modifie: $displayLabel" -ForegroundColor Yellow
            }
        }
    }

    return $differences
}

Function Get-BookmarkChildNames {
    param (
        [string[]] $childIds,
        [hashtable] $bookmarkLookup
    )

    $names = New-Object System.Collections.Generic.List[string]

    if (-not $childIds) {
        return @()
    }

    foreach ($childId in $childIds) {
        if ([string]::IsNullOrWhiteSpace($childId)) {
            continue
        }

        $displayName = $null
        if ($bookmarkLookup -and $bookmarkLookup.ContainsKey($childId)) {
            $childBookmark = $bookmarkLookup[$childId]
            if ($childBookmark -and $childBookmark.displayName) {
                $displayName = [string]$childBookmark.displayName
            }
        }

        if (-not $displayName) {
            $displayName = $childId
        }

        if (-not [string]::IsNullOrWhiteSpace($displayName)) {
            $names.Add($displayName.Trim()) | Out-Null
        }
    }

    return $names.ToArray()
}

# General report configuration comparison function
Function CompareReportData {
    param (
        [PSCustomObject] $newReport,
        [PSCustomObject] $oldReport
    )
    
    $differences = @()
    
    # If both are null, no differences
    if ($null -eq $newReport -and $null -eq $oldReport) {
        return $differences
    }
    
    # If only one is null, it's a difference
    if ($null -eq $newReport -or $null -eq $oldReport) {
        $diff = [ReportDifference]::new()
        $diff.ElementType = "Configuration"
        $diff.ElementName = "report.json"
        $diff.ElementPath = "report"
        $diff.PropertyName = "existence"
        $diff.DifferenceType = "Modifie"
        $diff.OldValue = if ($null -eq $oldReport) { "Absent" } else { "Present" }
        $diff.NewValue = if ($null -eq $newReport) { "Absent" } else { "Present" }
        $diff.HierarchyLevel = "report"
        $differences += $diff
        return $differences
    }
    
    # Comparison of important report properties
    $propertiesToCompare = @("version", "theme", "config", "settings")
    
    foreach ($property in $propertiesToCompare) {
        try {
            $newHasProp = $null -ne (Get-Member -InputObject $newReport -Name $property -MemberType Properties)
            $oldHasProp = $null -ne (Get-Member -InputObject $oldReport -Name $property -MemberType Properties)
            
            if ($newHasProp -or $oldHasProp) {
                $newValue = if ($newHasProp -and $null -ne $newReport.$property) { 
                    try { $newReport.$property | ConvertTo-Json -Depth 5 -Compress } catch { "" }
                } else { "" }
                
                $oldValue = if ($oldHasProp -and $null -ne $oldReport.$property) { 
                    try { $oldReport.$property | ConvertTo-Json -Depth 5 -Compress } catch { "" }
                } else { "" }
                
                if ($newValue -ne $oldValue) {
                    $diff = [ReportDifference]::new()
                    $diff.ElementType = "Configuration"
                    $diff.ElementName = $property
                    $diff.ElementPath = "report"
                    $diff.PropertyName = $property
                    $diff.DifferenceType = "Modifie"
                    $diff.OldValue = if ($oldValue.Length -gt 100) { $oldValue.Substring(0,100) + "..." } else { $oldValue }
                    $diff.NewValue = if ($newValue.Length -gt 100) { $newValue.Substring(0,100) + "..." } else { $newValue }
                    $diff.HierarchyLevel = "report.$property"
                    $differences += $diff
                    Write-Host "    Configuration $property modifiee" -ForegroundColor Yellow
                }
            }
        }
        catch {
            Write-Warning "Erreur lors de la comparaison de la propriete $property : $($_.Exception.Message)"
        }
    }
    
    return $differences
}

# ===== NEW FUNCTION SYSTEM CONFIGURATION ANALYSIS =====
# System configuration files comparison function (.platform, definition.pbir, version.json)
Function CompareSystemConfigurationFiles {
    param (
        [string] $newReportPath,
        [string] $oldReportPath
    )
    
    $differences = @()
    
    try {
        Write-Host "    Analyse des fichiers système de configuration..." -ForegroundColor Gray
        
        # System files are in the parent folder of $newReportPath and $oldReportPath
        $newReportParent = Split-Path $newReportPath -Parent
        $oldReportParent = Split-Path $oldReportPath -Parent
        
        # 1. .platform file comparison (project metadata)
        $platformDiffs = CompareSystemFile -newReportPath $newReportParent -oldReportPath $oldReportParent -fileName ".platform" -fileDescription "Métadonnées projet"
        $differences += $platformDiffs
        
        # 2. definition.pbir file comparison (dataset references)
        $definitionDiffs = CompareSystemFile -newReportPath $newReportParent -oldReportPath $oldReportParent -fileName "definition.pbir" -fileDescription "Références dataset"
        $differences += $definitionDiffs
        
        # 3. version.json file comparison (version information)
        $versionPath = "definition/version.json"
        $versionDiffs = CompareSystemFile -newReportPath $newReportPath -oldReportPath $oldReportPath -fileName $versionPath -fileDescription "Version du rapport"
        $differences += $versionDiffs
        
        # 4. Search for other configuration files in definition/ folder
        $additionalConfigFiles = @("report.json", "config.json", "metadata.json") # Fichiers supplémentaires potentiels
        foreach ($configFile in $additionalConfigFiles) {
            $configPath = "definition/$configFile"
            if ((Test-Path (Join-Path $newReportPath $configPath)) -or (Test-Path (Join-Path $oldReportPath $configPath))) {
                $configDiffs = CompareSystemFile -newReportPath $newReportPath -oldReportPath $oldReportPath -fileName $configPath -fileDescription "Configuration $configFile"
                $differences += $configDiffs
            }
        }
        
    }
    catch {
        Write-Warning "Erreur lors de la comparaison des fichiers de configuration système : $($_.Exception.Message)"
    }
    
    return $differences
}

# Utility function to compare a specific system file
Function CompareSystemFile {
    param (
        [string] $newReportPath,
        [string] $oldReportPath,
        [string] $fileName,
        [string] $fileDescription
    )
    
    $differences = @()
    
    try {
        $newFilePath = Join-Path $newReportPath $fileName
        $oldFilePath = Join-Path $oldReportPath $fileName
        
        $newExists = Test-Path $newFilePath
        $oldExists = Test-Path $oldFilePath
        
        # Case 1: File exists in both versions
        if ($newExists -and $oldExists) {
            $newContent = Get-Content $newFilePath -Raw -Encoding UTF8 | ConvertFrom-Json -ErrorAction SilentlyContinue
            $oldContent = Get-Content $oldFilePath -Raw -Encoding UTF8 | ConvertFrom-Json -ErrorAction SilentlyContinue
            
            if ($newContent -and $oldContent) {
                # Compare important properties based on file type
                $propertiesToCompare = Get-SystemFilePropertiesToCompare -fileName $fileName
                
                foreach ($property in $propertiesToCompare) {
                    $newValue = Get-NestedProperty -object $newContent -propertyPath $property
                    $oldValue = Get-NestedProperty -object $oldContent -propertyPath $property
                    
                    $newValueStr = if ($null -eq $newValue) { "" } else { $newValue.ToString() }
                    $oldValueStr = if ($null -eq $oldValue) { "" } else { $oldValue.ToString() }
                    
                    if ($newValueStr -ne $oldValueStr) {
                        $diff = [ReportDifference]::new()
                        $diff.ElementType = "Configuration"
                        $diff.ElementName = $fileName
                        $diff.ElementDisplayName = "$fileDescription ($fileName)"
                        $diff.ElementPath = $fileName
                        $diff.PropertyName = $property
                        $diff.DifferenceType = "Modifie"
                        $diff.OldValue = if ($oldValueStr.Length -gt 100) { $oldValueStr.Substring(0,100) + "..." } else { $oldValueStr }
                        $diff.NewValue = if ($newValueStr.Length -gt 100) { $newValueStr.Substring(0,100) + "..." } else { $newValueStr }
                        $diff.HierarchyLevel = "system.$fileName.$property"
                        $diff.AdditionalInfo = "Fichier système: $fileDescription"
                        $differences += $diff
                        
                        Write-Host "      $fileDescription ($property): '$oldValueStr' → '$newValueStr'" -ForegroundColor Yellow
                    }
                }
            }
        }
        # Case 2: File exists only in one version
        elseif ($newExists -or $oldExists) {
            $diff = [ReportDifference]::new()
            $diff.ElementType = "Configuration"
            $diff.ElementName = $fileName
            $diff.ElementDisplayName = "$fileDescription ($fileName)"
            $diff.ElementPath = $fileName
            $diff.PropertyName = "existence"
            $diff.DifferenceType = if ($newExists) { "Ajoute" } else { "Supprime" }
            $diff.OldValue = if ($oldExists) { "Present" } else { "Absent" }
            $diff.NewValue = if ($newExists) { "Present" } else { "Absent" }
            $diff.HierarchyLevel = "system.$fileName"
            $diff.AdditionalInfo = "Fichier système: $fileDescription"
            $differences += $diff
            
            $statusMessage = if ($newExists) { 'ajouté' } else { 'supprimé' }
            $statusColor = if ($newExists) { 'Green' } else { 'Red' }
            Write-Host "      $fileDescription : fichier $statusMessage" -ForegroundColor $statusColor
        }
    }
    catch {
        Write-Warning "Erreur lors de la comparaison du fichier $fileName : $($_.Exception.Message)"
    }
    
    return $differences
}

# Function to determine which properties to compare based on file type
Function Get-SystemFilePropertiesToCompare {
    param ([string] $fileName)
    
    switch -Regex ($fileName) {
        "\.platform$" {
            return @("metadata.displayName", "metadata.type", "config.version", "config.workspaceId")
        }
        "definition\.pbir$" {
            return @("version", "datasetReference.byPath.path", "datasetReference.byConnection.connectionString")
        }
        "version\.json$" {
            return @("version", "minVersion", "gitCommitSha")
        }
        "report\.json$" {
            return @("version", "config.theme", "config.layoutOptimization", "settings.locale")
        }
        default {
            # For unknown files, compare all first-level properties
            return @("version", "name", "displayName", "type", "config")
        }
    }
}

# Utility function to access nested properties with path (ex: "metadata.displayName")
Function Get-NestedProperty {
    param (
        [PSCustomObject] $object,
        [string] $propertyPath
    )
    
    if (-not $object -or -not $propertyPath) {
        return $null
    }
    
    $parts = $propertyPath.Split('.')
    $current = $object
    
    foreach ($part in $parts) {
        if ($current -and (Get-Member -InputObject $current -Name $part -MemberType Properties)) {
            $current = $current.$part
        } else {
            return $null
        }
    }
    
    return $current
}

# NEW FUNCTION: Theme and style extraction
Function ExtractReportThemesAndStyles {
    param ([PSCustomObject] $reportData)
    
    $result = @{
        HasThemes = $false
        Theme = $null
        VisualStyles = @{}
        LayoutOptimization = $null
        ColorScheme = @{}
        ThemeSummary = ""
    }
    
    try {
        if (-not $reportData) {
            return $result
        }
        
        $themeInfos = @()
        
        # 1. Search in config.theme
        if ($reportData.config -and $reportData.config.theme) {
            $result.Theme = $reportData.config.theme
            $result.HasThemes = $true
            
            if ($reportData.config.theme.name) {
                $themeInfos += "Theme: $($reportData.config.theme.name)"
            }
            
            # Analyze theme colors
            if ($reportData.config.theme.dataColors) {
                $colorCount = $reportData.config.theme.dataColors.Count
                $themeInfos += "Couleurs de donnees: $colorCount"
                $result.ColorScheme["dataColors"] = $reportData.config.theme.dataColors
            }
            
            if ($reportData.config.theme.background) {
                $themeInfos += "Arriere-plan: $($reportData.config.theme.background)"
                $result.ColorScheme["background"] = $reportData.config.theme.background
            }
            
            if ($reportData.config.theme.foreground) {
                $themeInfos += "Premier-plan: $($reportData.config.theme.foreground)"
                $result.ColorScheme["foreground"] = $reportData.config.theme.foreground
            }
            
            # Analyze text styles
            if ($reportData.config.theme.textClasses) {
                $textClassCount = $reportData.config.theme.textClasses.Count
                $themeInfos += "Classes de texte: $textClassCount"
                $result.ColorScheme["textClasses"] = $reportData.config.theme.textClasses
            }
        }
        
        # 2. Direct search in theme
        if ($reportData.theme) {
            $result.Theme = $reportData.theme
            $result.HasThemes = $true
            $themeInfos += "Theme direct: Present"
        }
        
        # 3. Search in config.visualStyles
        if ($reportData.config -and $reportData.config.visualStyles) {
            $result.VisualStyles = $reportData.config.visualStyles
            $styleCount = $reportData.config.visualStyles.PSObject.Properties.Count
            $themeInfos += "Styles visuels: $styleCount types"
        }
        
        # 4. Search in config.layoutOptimization
        if ($reportData.config -and $reportData.config.layoutOptimization) {
            $result.LayoutOptimization = $reportData.config.layoutOptimization
            $themeInfos += "Optimisation de mise en page: Active"
        }
        
        # 5. Search in settings for global styles
        if ($reportData.settings) {
            $styleProperties = $reportData.settings.PSObject.Properties | Where-Object { $_.Name -like "*style*" -or $_.Name -like "*Style*" -or $_.Name -like "*theme*" -or $_.Name -like "*Theme*" }
            foreach ($styleProp in $styleProperties) {
                $themeInfos += "Setting.$($styleProp.Name): Active"
            }
        }
        
        if ($themeInfos.Count -gt 0) {
            $result.ThemeSummary = ($themeInfos -join " | ")
            if (-not $result.HasThemes) {
                $result.HasThemes = $true
            }
        }
        
    }
    catch {
        Write-Warning "Erreur lors de l'extraction des themes et styles: $($_.Exception.Message)"
    }
    
    return $result
}

# NEW FUNCTION: Theme and style comparison
Function CompareReportThemesAndStyles {
    param (
        [PSCustomObject] $newReport,
        [PSCustomObject] $oldReport
    )
    
    $differences = @()
    
    try {
        Write-Host "Comparaison des themes et styles..." -ForegroundColor Yellow
        
        $newThemes = ExtractReportThemesAndStyles -reportData $newReport
        $oldThemes = ExtractReportThemesAndStyles -reportData $oldReport
        
        # Compare theme presence
        if ($newThemes.HasThemes -ne $oldThemes.HasThemes) {
            $diff = [ReportDifference]::new()
            $diff.ElementType = "Theme"
            $diff.ElementName = "presence"
            $diff.ElementDisplayName = "Presence de themes"
            $diff.ElementPath = "report.theme"
            $diff.PropertyName = "hasThemes"
            $diff.DifferenceType = "Modifie"
            $diff.OldValue = if ($oldThemes.HasThemes) { "Present" } else { "Absent" }
            $diff.NewValue = if ($newThemes.HasThemes) { "Present" } else { "Absent" }
            $diff.HierarchyLevel = "report.theme"
            $diff.AdditionalInfo = "Ancien: $($oldThemes.ThemeSummary) | Nouveau: $($newThemes.ThemeSummary)"
            $differences += $diff
            Write-Host "    Presence de themes modifiee" -ForegroundColor Yellow
        }
        
        # Compare theme summary
        if ($newThemes.ThemeSummary -ne $oldThemes.ThemeSummary) {
            $diff = [ReportDifference]::new()
            $diff.ElementType = "Theme"
            $diff.ElementName = "configuration"
            $diff.ElementDisplayName = "Configuration des themes"
            $diff.ElementPath = "report.theme.config"
            $diff.PropertyName = "themeSummary"
            $diff.DifferenceType = "Modifie"
            $diff.OldValue = if ($oldThemes.ThemeSummary) { $oldThemes.ThemeSummary } else { "Aucun theme" }
            $diff.NewValue = if ($newThemes.ThemeSummary) { $newThemes.ThemeSummary } else { "Aucun theme" }
            $diff.HierarchyLevel = "report.theme"
            $diff.AdditionalInfo = "Configuration generale des themes modifiee"
            $differences += $diff
            Write-Host "    Configuration des themes modifiee" -ForegroundColor Yellow
        }
        
        # Compare color schemes
        if ($newThemes.ColorScheme.Count -gt 0 -or $oldThemes.ColorScheme.Count -gt 0) {
            $allColorKeys = @()
            $allColorKeys += $newThemes.ColorScheme.Keys
            $allColorKeys += $oldThemes.ColorScheme.Keys
            $allColorKeys = $allColorKeys | Select-Object -Unique
            
            foreach ($colorKey in $allColorKeys) {
                $newColors = $newThemes.ColorScheme[$colorKey]
                $oldColors = $oldThemes.ColorScheme[$colorKey]
                
                $newColorsJson = if ($newColors) { $newColors | ConvertTo-Json -Compress -Depth 3 } else { "" }
                $oldColorsJson = if ($oldColors) { $oldColors | ConvertTo-Json -Compress -Depth 3 } else { "" }
                
                if ($newColorsJson -ne $oldColorsJson) {
                    $diff = [ReportDifference]::new()
                    $diff.ElementType = "Theme"
                    $diff.ElementName = $colorKey
                    $diff.ElementDisplayName = "Schema de couleurs '$colorKey'"
                    $diff.ElementPath = "report.theme.$colorKey"
                    $diff.PropertyName = "colors"
                    $diff.DifferenceType = "Modifie"
                    $diff.OldValue = if ($oldColorsJson.Length -gt 150) { $oldColorsJson.Substring(0,150) + "..." } else { $oldColorsJson }
                    $diff.NewValue = if ($newColorsJson.Length -gt 150) { $newColorsJson.Substring(0,150) + "..." } else { $newColorsJson }
                    $diff.HierarchyLevel = "report.theme.colors"
                    $diff.AdditionalInfo = "Schema de couleurs $colorKey modifie"
                    $differences += $diff
                    Write-Host "    Schema de couleurs $colorKey modifie" -ForegroundColor Yellow
                }
            }
        }
        
        # Compare visual styles
        $newStylesJson = if ($newThemes.VisualStyles.Count -gt 0) { $newThemes.VisualStyles | ConvertTo-Json -Compress -Depth 3 } else { "" }
        $oldStylesJson = if ($oldThemes.VisualStyles.Count -gt 0) { $oldThemes.VisualStyles | ConvertTo-Json -Compress -Depth 3 } else { "" }
        
        if ($newStylesJson -ne $oldStylesJson) {
            $diff = [ReportDifference]::new()
            $diff.ElementType = "Theme"
            $diff.ElementName = "visualStyles"
            $diff.ElementDisplayName = "Styles visuels globaux"
            $diff.ElementPath = "report.theme.visualStyles"
            $diff.PropertyName = "visualStyles"
            $diff.DifferenceType = "Modifie"
            $diff.OldValue = if ($oldStylesJson.Length -gt 150) { $oldStylesJson.Substring(0,150) + "..." } else { $oldStylesJson }
            $diff.NewValue = if ($newStylesJson.Length -gt 150) { $newStylesJson.Substring(0,150) + "..." } else { $newStylesJson }
            $diff.HierarchyLevel = "report.theme.visualStyles"
            $diff.AdditionalInfo = "Styles visuels globaux modifies"
            $differences += $diff
            Write-Host "    Styles visuels globaux modifies" -ForegroundColor Yellow
        }
        
        Write-Host "    $($differences.Count) differences de themes/styles trouvees" -ForegroundColor White
        
    }
    catch {
        Write-Warning "Erreur lors de la comparaison des themes et styles: $($_.Exception.Message)"
    }
    
    return $differences
}




# EXECUTION AUTOMATIQUE - VERSION SIMPLIFIEE
# Garde d'execution: ne s'exécute que lorsque ce fichier est lancé directement (pas lorsqu'il est dot-sourcé)
$__thisScriptPath = $MyInvocation.MyCommand.Path
$__callerPath = $PSCommandPath
$__invocationName = $MyInvocation.InvocationName
$__isDotSourced = ($__invocationName -eq '.')
$shouldRunHarness = $false
if (-not $__isDotSourced) {
    if ($__thisScriptPath -and $__callerPath) {
        $shouldRunHarness = ($__thisScriptPath -eq $__callerPath)
    } else {
        $shouldRunHarness = $true
    }
}

# Autoriser la désactivation explicite du harnais via une variable d'environnement
if ($env:PBI_DISABLE_COMPARE_HARNESS -eq '1') {
    $shouldRunHarness = $false
}
if ($shouldRunHarness) {
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "COMPARAISON AUTOMATIQUE DES RAPPORTS POWER BI" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan

try {
    # Verification des chemins
    Write-Host "Verification des chemins..." -ForegroundColor Yellow
    
    if (-not (Test-Path $NOUVEAU_PROJET)) {
        throw "Le chemin du nouveau projet n'existe pas: $NOUVEAU_PROJET"
    }
    
    if (-not (Test-Path $ANCIEN_PROJET)) {
        throw "Le chemin de l'ancien projet n'existe pas: $ANCIEN_PROJET"
    }
    
    if (-not (Test-Path $DOSSIER_SORTIE)) {
        throw "Le dossier de sortie n'existe pas: $DOSSIER_SORTIE"
    }
    
    Write-Host "Tous les chemins sont valides" -ForegroundColor Green
    
    # Chargement des projets
    $projects = LoadReportProjectVersions -newVersionProjectDirectory $NOUVEAU_PROJET -oldVersionProjectDirectory $ANCIEN_PROJET
    
    if ($projects.Count -ne 2) {
        throw "Erreur lors du chargement des projets"
    }
    
    $newProject = $projects[0]
    $oldProject = $projects[1]
    
    Write-Host "Projets charges avec succes" -ForegroundColor Green
    
    # Comparison
    $differences = CompareReportProjects -newProject $newProject -oldProject $oldProject
    
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "RESULTATS DE LA COMPARAISON" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Nombre total de differences: $($differences.Count)" -ForegroundColor White
    
    if ($differences.Count -gt 0) {
        $grouped = $differences | Group-Object ElementType
        foreach ($group in $grouped) {
            Write-Host "  - $($group.Name): $($group.Count) differences" -ForegroundColor White
        }
    }
    else {
        Write-Host "  Aucune difference detectee dans les rapports!" -ForegroundColor Green
    }
    
    # Verification qualite des slicers (nouveau projet uniquement)
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "VERIFICATION QUALITE DES SLICERS" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    
    $visibleQuality = Invoke-VisibleSlicerQualityCheck -project $newProject -projectRoot $NOUVEAU_PROJET
    $checkResults = @()

    if ($visibleQuality -and $visibleQuality.PSObject.Properties['Results']) {
        $checkResults = $visibleQuality.Results
    }

    if ($checkResults.Count -gt 0) {
        $checkResults = Get-VisibleSlicerQualityResult -checkResults $checkResults -newProject $newProject
    }

    $removedHidden = if ($visibleQuality -and $visibleQuality.PSObject.Properties['HiddenCount']) { [int]$visibleQuality.HiddenCount } else { 0 }
    $totalCandidates = if ($visibleQuality -and $visibleQuality.PSObject.Properties['TotalCandidates']) { [int]$visibleQuality.TotalCandidates } else { $checkResults.Count }
    
    Write-Host "Verification qualite terminee" -ForegroundColor Green
    Write-Host "Resultats de verification: $($checkResults.Count) slicers visibles analyses" -ForegroundColor White
    if ($totalCandidates -ne $checkResults.Count) {
        Write-Host "  - Total slicers (y compris caches): $totalCandidates" -ForegroundColor Gray
    }
    
    if ($removedHidden -gt 0) {
        Write-Host "  - Filtres caches ignores: $removedHidden" -ForegroundColor DarkGray
    }
    
    if ($checkResults.Count -gt 0) {
        $alertes = $checkResults | Where-Object { $_.Status -eq "ALERTE" }
        $ok = $checkResults | Where-Object { $_.Status -eq "OK" }
        Write-Host "  - OK: $($ok.Count)" -ForegroundColor Green
        Write-Host "  - ALERTES: $($alertes.Count)" -ForegroundColor Red
    }
    
    # SIMPLIFIED SYNC: No more grouping or complex analysis
    # Synchronization differences are now treated like any other differences
    
    # Generation du rapport HTML (avec comparaison et checks qualite)
    $differencesForHtml = Convert-ReportDifferencesForHtml -differences $differences
    $reportPath = BuildReportHTMLReport_Orange -differences $differencesForHtml -checkResults $checkResults -outputFolder $DOSSIER_SORTIE
    
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "COMPARAISON TERMINEE AVEC SUCCES" -ForegroundColor Green
    Write-Host "Rapport HTML genere: $reportPath" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    
    # Afficher le rapport
    Start-Process $reportPath
    Write-Host "Rapport ouvert dans le navigateur" -ForegroundColor Green
}
catch {
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "ERREUR LORS DE L'EXECUTION" -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "Message d'erreur: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red
}

if ($env:PBI_SKIP_PAUSE -eq '1') {
    Write-Host "Script termine." -ForegroundColor Gray
}
else {
    Write-Host "Script termine. Appuyez sur Entree pour fermer..."
    Read-Host
}
}

