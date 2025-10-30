# Power BI Report Quality Check Script (PBIR format)
# Analyzes slicers in the new project only for UX quality verification

# Function to load report data for quality checks
Function LoadReportDataForCheck {
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

# Function to load pages data for quality checks
Function LoadPagesDataForCheck {
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
                if (-not [string]::IsNullOrWhiteSpace($content)) {
                    $pageData = $content | ConvertFrom-Json
                    
                    # Extract readable information
                    $displayName = if ($pageData.displayName) { $pageData.displayName } else { "(No name)" }
                    Write-Host "    Analyse de la page: '$displayName' ($pageName)" -ForegroundColor Gray
                    
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
                                        
                                        # Extract visual type
                                        $visualType = if ($visualData.visual -and $visualData.visual.visualType) { 
                                            $visualData.visual.visualType 
                                        } else { "unknown" }
                                        
                                        $visualData | Add-Member -NotePropertyName "_extracted_type" -NotePropertyValue $visualType -Force
                                        
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
                }
            }
            catch {
                Write-Warning "    Erreur lors du chargement de la page $pageName : $($_.Exception.Message)"
            }
        }
    }
    
    return $pagesData
}

# Helper to extract literal string values from nested JSON structures
Function Get-LiteralStringValue {
    param($value)

    if ($null -eq $value) {
        return $null
    }

    if ($value -is [string]) {
        return $value
    }

    if ($value -is [bool] -or $value -is [int] -or $value -is [double]) {
        return $value.ToString()
    }

    try {
        $psValue = [System.Management.Automation.PSObject]$value

        foreach ($property in @('Literal', 'expr', 'Value')) {
            if ($psValue.PSObject.Properties[$property]) {
                $nested = $psValue.PSObject.Properties[$property].Value
                $result = Get-LiteralStringValue -value $nested
                if ($null -ne $result -and $result -ne '') {
                    return $result
                }
            }
        }
    }
    catch {
        return $null
    }

    return $null
}

# Function to extract the main field name used by the slicer
Function Get-SlicerFieldName {
    param([PSCustomObject] $visual)

    try {
        if ($visual.visual -and $visual.visual.query -and $visual.visual.query.queryState -and $visual.visual.query.queryState.Values -and $visual.visual.query.queryState.Values.projections) {
            foreach ($projection in $visual.visual.query.queryState.Values.projections) {
                if ($projection.field -and $projection.field.Column -and $projection.field.Column.Property) {
                    return $projection.field.Column.Property
                }
            }
        }
    }
    catch {
        Write-Warning "Erreur lors de l'extraction du champ principal: $($_.Exception.Message)"
    }

    return $null
}

# Function to extract the display name shown to end-users
Function Get-SlicerDisplayName {
    param(
        [PSCustomObject] $visual,
        [string] $fallbackField,
        [string] $fallbackGroup
    )

    $candidates = @()

    try {
        if ($visual.visual -and $visual.visual.objects -and $visual.visual.objects.header) {
            foreach ($headerItem in $visual.visual.objects.header) {
                if ($headerItem.properties -and $headerItem.properties.text) {
                    $label = Get-LiteralStringValue -value $headerItem.properties.text
                    if (-not [string]::IsNullOrWhiteSpace($label)) {
                        $candidates += $label
                    }
                }
            }
        }

        if ($candidates.Count -eq 0 -and $visual.visual -and $visual.visual.visualContainerObjects -and $visual.visual.visualContainerObjects.title) {
            foreach ($titleItem in $visual.visual.visualContainerObjects.title) {
                if ($titleItem.properties -and $titleItem.properties.text) {
                    $label = Get-LiteralStringValue -value $titleItem.properties.text
                    if (-not [string]::IsNullOrWhiteSpace($label)) {
                        $candidates += $label
                    }
                }
            }
        }

        if ($candidates.Count -eq 0 -and $visual.visual -and $visual.visual.objects -and $visual.visual.objects.general) {
            foreach ($generalItem in $visual.visual.objects.general) {
                if ($generalItem.properties -and $generalItem.properties.displayName) {
                    $label = Get-LiteralStringValue -value $generalItem.properties.displayName
                    if (-not [string]::IsNullOrWhiteSpace($label)) {
                        $candidates += $label
                    }
                }
            }
        }
    }
    catch {
        Write-Warning "Erreur lors de l'extraction du display name: $($_.Exception.Message)"
    }

    foreach ($candidate in $candidates) {
        $clean = $candidate.Trim("'").Trim()
        if (-not [string]::IsNullOrWhiteSpace($clean)) {
            return $clean
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($fallbackGroup)) {
        return $fallbackGroup
    }

    if (-not [string]::IsNullOrWhiteSpace($fallbackField)) {
        return $fallbackField
    }

    return "(Filtre sans nom)"
}

# Function to extract slicer group information
Function Get-SlicerGroupName {
    param(
        [PSCustomObject] $visual,
        [PSCustomObject] $pageData
    )
    
    try {
        # Check in visual.visualContainerObjects.general for group information
        if ($visual.visual -and $visual.visual.visualContainerObjects -and $visual.visual.visualContainerObjects.general) {
            foreach ($generalItem in $visual.visual.visualContainerObjects.general) {
                if ($generalItem.properties -and $generalItem.properties.displayName) {
                    if ($generalItem.properties.displayName.expr -and $generalItem.properties.displayName.expr.Literal) {
                        $groupName = $generalItem.properties.displayName.expr.Literal.Value -replace "'", ""
                        if ($groupName -and $groupName -ne "") {
                            return $groupName
                        }
                    }
                }
            }
        }
        
        # Traverse parent groups for displayName (e.g., groups ending with _filter or _period)
        if ($pageData -and $visual.PSObject.Properties['parentGroupName']) {
            $currentGroupId = [string]$visual.parentGroupName
            $visited = New-Object System.Collections.Generic.HashSet[string]

            while (-not [string]::IsNullOrWhiteSpace($currentGroupId)) {
                if (-not $visited.Add($currentGroupId)) {
                    break
                }

                if ($pageData.PSObject.Properties['visuals'] -and $pageData.visuals.ContainsKey($currentGroupId)) {
                    $groupVisual = $pageData.visuals[$currentGroupId]
                    if ($groupVisual -and $groupVisual.PSObject.Properties['visualGroup'] -and $groupVisual.visualGroup.PSObject.Properties['displayName']) {
                        $groupDisplayName = [string]$groupVisual.visualGroup.displayName
                        if (-not [string]::IsNullOrWhiteSpace($groupDisplayName)) {
                            $cleanGroupName = $groupDisplayName.Trim().Trim("'")
                            if (-not [string]::IsNullOrWhiteSpace($cleanGroupName)) {
                                return $cleanGroupName
                            }
                        }
                    }

                    if ($groupVisual.PSObject.Properties['parentGroupName'] -and $groupVisual.parentGroupName) {
                        $currentGroupId = [string]$groupVisual.parentGroupName
                        continue
                    }
                }

                break
            }
        }

        # Fallback to visual name
        if ($visual.name) {
            return ([string]$visual.name).Trim().Trim("'")
        }

        return $null
    }
    catch {
        Write-Warning "Erreur lors de l'extraction du nom de groupe: $($_.Exception.Message)"
        return $null
    }
}

# Function to check if slicer has selected values
Function Test-SlicerHasSelection {
    param([PSCustomObject] $visual)
    
    try {
        # Check in visual.query.queryState.Filters for selected values
        if ($visual.visual -and $visual.visual.query -and $visual.visual.query.queryState -and $visual.visual.query.queryState.Filters) {
            foreach ($filter in $visual.visual.query.queryState.Filters) {
                if ($filter.Filter) {
                    # Check for "In" conditions (selected values)
                    if ($filter.Filter.Where) {
                        foreach ($whereClause in $filter.Filter.Where) {
                            if ($whereClause.Condition -and $whereClause.Condition.In) {
                                if ($whereClause.Condition.In.Values -and $whereClause.Condition.In.Values.Count -gt 0) {
                                    return $true
                                }
                            }
                        }
                    }
                }
            }
        }
        
        # Check in visual.objects.general for filter properties
        if ($visual.visual -and $visual.visual.objects -and $visual.visual.objects.general) {
            foreach ($generalItem in $visual.visual.objects.general) {
                if ($generalItem.properties -and $generalItem.properties.filter) {
                    return $true
                }
            }
        }
        
        return $false
    }
    catch {
        Write-Warning "Erreur lors de la vérification des sélections: $($_.Exception.Message)"
        return $false
    }
}

# Function to identify search usage and retrieve captured text
Function Get-SlicerSearchInfo {
    param([PSCustomObject] $visual)
    
    $searchTexts = @()

    try {
        # Check in visual.objects.search for search value
        if ($visual.visual -and $visual.visual.objects -and $visual.visual.objects.search) {
            foreach ($searchItem in $visual.visual.objects.search) {
                if ($searchItem.properties -and $searchItem.properties.value) {
                    if ($searchItem.properties.value.expr -and $searchItem.properties.value.expr.Literal) {
                        $searchValue = $searchItem.properties.value.expr.Literal.Value
                        if (-not [string]::IsNullOrWhiteSpace($searchValue)) {
                            $searchTexts += $searchValue
                        }
                    }
                }
            }
        }
        
        # Check for other search-related properties and self filters
        if ($visual.visual -and $visual.visual.objects) {
            $searchProperties = $visual.visual.objects.PSObject.Properties | Where-Object { $_.Name -like "*search*" -or $_.Name -like "*Search*" }
            foreach ($searchProp in $searchProperties) {
                if ($searchProp.Value -and $searchProp.Value -ne $null) {
                    $rawValue = Get-LiteralStringValue -value $searchProp.Value
                    if (-not [string]::IsNullOrWhiteSpace($rawValue)) {
                        $searchTexts += $rawValue
                    }
                }
            }

            if ($visual.visual.objects.general) {
                foreach ($generalItem in $visual.visual.objects.general) {
                    if (-not $generalItem.properties) { continue }

                    foreach ($prop in $generalItem.properties.PSObject.Properties) {
                        $propName = $prop.Name
                        $propValue = $prop.Value

                        if ($propName -eq 'selfFilter' -and $propValue.filter -and $propValue.filter.Where) {
                            foreach ($whereClause in $propValue.filter.Where) {
                                $condition = $whereClause.Condition
                                if ($null -eq $condition) { continue }

                                foreach ($condProp in $condition.PSObject.Properties) {
                                    if ($condProp.Name -match 'Contains' -or $condProp.Name -match 'Like' -or $condProp.Name -match 'Search') {
                                        $rightValue = $condProp.Value.Right
                                        $literal = Get-LiteralStringValue -value $rightValue
                                        if (-not [string]::IsNullOrWhiteSpace($literal)) {
                                            $searchTexts += $literal
                                        }
                                    }
                                }
                            }
                        }
                        elseif ($propName -match 'search' -and $null -ne $propValue) {
                            $literal = Get-LiteralStringValue -value $propValue
                            if (-not [string]::IsNullOrWhiteSpace($literal)) {
                                $searchTexts += $literal
                            }
                        }
                    }
                }
            }
        }
        
        $cleanTexts = @()
        foreach ($value in $searchTexts) {
            $clean = ($value.ToString()).Trim("'").Trim()
            if (-not [string]::IsNullOrWhiteSpace($clean)) {
                $lower = $clean.ToLower()
                if ($lower -notin @('true','false','basic','dropdown')) {
                    $cleanTexts += $clean
                }
            }
        }

        $uniqueTexts = $cleanTexts | Select-Object -Unique

        return [PSCustomObject]@{
            HasSearch = ($uniqueTexts.Count -gt 0)
            SearchText = if ($uniqueTexts.Count -gt 0) { $uniqueTexts -join ', ' } else { '' }
        }
    }
    catch {
        Write-Warning "Erreur lors de la vérification de la recherche: $($_.Exception.Message)"
        return [PSCustomObject]@{
            HasSearch = $false
            SearchText = ''
        }
    }
}

# Function to check if slicer is in single select mode
Function Test-SlicerIsSingleSelect {
    param([PSCustomObject] $visual)
    
    try {
        # Check in visual.objects.general for singleSelect property
        if ($visual.visual -and $visual.visual.objects -and $visual.visual.objects.general) {
            foreach ($generalItem in $visual.visual.objects.general) {
                if ($generalItem.properties -and $generalItem.properties.singleSelect) {
                    if ($generalItem.properties.singleSelect.expr -and $generalItem.properties.singleSelect.expr.Literal) {
                        $singleSelectValue = $generalItem.properties.singleSelect.expr.Literal.Value
                        return ($singleSelectValue -eq "true" -or $singleSelectValue -eq $true)
                    }
                }
            }
        }
        
        return $false
    }
    catch {
        Write-Warning "Erreur lors de la vérification du mode single select: $($_.Exception.Message)"
        return $false
    }
}

# Check if slicer in _period group has only "Current Year" selected
Function Test-SlicerIsCurrentYearOnly {
    param([PSCustomObject] $visual)

    try {
        # Extract selected values from visual.objects.general filter
        $selectedValues = @()

        if ($visual.visual -and $visual.visual.objects -and $visual.visual.objects.general) {
            foreach ($generalItem in $visual.visual.objects.general) {
                if ($generalItem.properties -and $generalItem.properties.filter -and
                    $generalItem.properties.filter.filter -and $generalItem.properties.filter.filter.Where) {

                    foreach ($whereClause in $generalItem.properties.filter.filter.Where) {
                        if ($whereClause.Condition -and $whereClause.Condition.In -and
                            $whereClause.Condition.In.Values) {

                            foreach ($valueArray in $whereClause.Condition.In.Values) {
                                if ($valueArray) {
                                    foreach ($valueItem in $valueArray) {
                                        if ($valueItem.Literal -and $valueItem.Literal.Value) {
                                            $selectedValues += $valueItem.Literal.Value
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        # Normalize values (remove quotes and whitespace)
        $normalizedValues = @()
        foreach ($value in $selectedValues) {
            $normalized = $value.ToString().Trim().Trim("'").Trim('"')
            $normalizedValues += $normalized
        }

        # Check: exactly 1 value AND it's "Current Year"
        if ($normalizedValues.Count -eq 1 -and
            $normalizedValues[0] -eq 'Current Year') {
            return $true
        }

        return $false
    }
    catch {
        Write-Warning "Erreur lors de la verification Current Year: $($_.Exception.Message)"
        return $false
    }
}

# Extract all selected values from slicer for error reporting
Function Get-SlicerSelectedValues {
    param([PSCustomObject] $visual)

    $selectedValues = @()

    try {
        if ($visual.visual -and $visual.visual.objects -and $visual.visual.objects.general) {
            foreach ($generalItem in $visual.visual.objects.general) {
                if ($generalItem.properties -and $generalItem.properties.filter -and
                    $generalItem.properties.filter.filter -and $generalItem.properties.filter.filter.Where) {

                    foreach ($whereClause in $generalItem.properties.filter.filter.Where) {
                        if ($whereClause.Condition -and $whereClause.Condition.In -and
                            $whereClause.Condition.In.Values) {

                            foreach ($valueArray in $whereClause.Condition.In.Values) {
                                if ($valueArray) {
                                    foreach ($valueItem in $valueArray) {
                                        if ($valueItem.Literal -and $valueItem.Literal.Value) {
                                            $cleanValue = $valueItem.Literal.Value.ToString().Trim().Trim("'").Trim('"')
                                            $selectedValues += $cleanValue
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }
    catch {
        # Silent failure, return empty array
    }

    return $selectedValues
}

# Function to check the visibility of the filter pane (ruban)
Function Invoke-FilterPaneQualityCheck {
    param(
        [string] $reportPath # This is the ...\definition path
    )
    
    $results = @()
    
    Write-Host "  Analyse du volet de filtre (ruban)..." -ForegroundColor Yellow

    try {
        $reportJsonPath = Join-Path $reportPath "report.json"
        if (-not (Test-Path $reportJsonPath)) {
            # This is not a fatal error for the whole check process, just for this part.
            Write-Warning "Fichier report.json principal introuvable, impossible de vérifier le volet de filtre."
            return $results
        }

        $reportData = Get-Content $reportJsonPath -Raw -Encoding UTF8 | ConvertFrom-Json

        # Default to not visible, which is the desired state
        $isVisible = $false
        $paneVisibleValue = "false" # Default value if not found

        if ($reportData.objects -and $reportData.objects.outspacePane) {
            foreach ($pane in $reportData.objects.outspacePane) {
                if ($pane.properties -and $pane.properties.visible) {
                    # Get-LiteralStringValue is good for nested structures
                    $paneVisibleValue = Get-LiteralStringValue -value $pane.properties.visible
                    if ($paneVisibleValue -eq 'true') {
                        $isVisible = $true
                        break
                    }
                }
            }
        }

        $status = "OK"
        $message = "Le volet de filtre est masqué comme attendu."
        $messageKey = "filter_pane_hidden"

        if ($isVisible) {
            $status = "ALERTE"
            $message = "Le volet de filtre est visible, il devrait être masqué."
            $messageKey = "filter_pane_visible"
        }

        $result = [PSCustomObject]@{
            PageName        = "Rapport Global"
            DisplayName     = "Volet Filtre"
            VisualId        = "FilterPane"
            FieldName       = "N/A"
            GroupName       = "Configuration"
            HasSelected     = $null # Use null for non-applicable boolean
            HasSearch       = $null
            SearchText      = ""
            IsRadio         = $null
            Status          = $status
            Message         = $message
            MessageKey      = $messageKey
            MessageDetail   = "Visibilité: $($isVisible)"
        }
        $results += $result

        $statusColor = if ($status -eq "OK") { "Green" } else { "Red" }
        Write-Host "    Status: $status - $message" -ForegroundColor $statusColor
    }
    catch {
        Write-Warning "Erreur lors de la vérification du volet de filtre: $($_.Exception.Message)"
        $errorResult = [PSCustomObject]@{
            PageName        = "Rapport Global"
            DisplayName     = "Volet Filtre"
            VisualId        = "FilterPane"
            FieldName       = "N/A"
            GroupName       = "Configuration"
            HasSelected     = $null
            HasSearch       = $null
            SearchText      = ""
            IsRadio         = $null
            Status          = "ERREUR"
            Message         = "Erreur lors de l'analyse du volet de filtre: $($_.Exception.Message)"
            MessageKey      = "filter_pane_error"
            MessageDetail   = ""
        }
        $results += $errorResult
    }

    return $results
}

# Main function to perform slicer quality checks
Function Invoke-SlicerQualityCheck {
    param([string] $projectPath)
    
    Write-Host "=== Analyse qualité des slicers ===" -ForegroundColor Green
    Write-Host "Projet analysé: $projectPath" -ForegroundColor Cyan
    
    $results = @()
    
    try {
        # Find the .pbip file to get the report structure
        $pbipFiles = Get-ChildItem -Path $projectPath -Filter "*.pbip"
        
        if ($pbipFiles.Count -eq 0) {
            throw "Aucun fichier .pbip trouvé dans '$projectPath'"
        }
        
        if ($pbipFiles.Count -gt 1) {
            throw "Plusieurs fichiers .pbip trouvés dans '$projectPath'"
        }
        
        $pbipFile = $pbipFiles[0]
        $reportPath = Join-Path $projectPath "$($pbipFile.BaseName).Report\definition"
        
        Write-Host "Chemin du rapport: $reportPath" -ForegroundColor White
        
        if (-not (Test-Path $reportPath)) {
            throw "Le dossier Report n'existe pas dans '$projectPath'"
        }

        # Perform filter pane check
        $paneResults = Invoke-FilterPaneQualityCheck -reportPath $reportPath
        $results += $paneResults
        
        # Load pages data
        $pagesData = LoadPagesDataForCheck -reportPath $reportPath
        
        Write-Host "Pages chargées: $($pagesData.Count)" -ForegroundColor Green
        
        # Analyze each page for slicers
        foreach ($pageEntry in $pagesData.GetEnumerator()) {
            $pageName = $pageEntry.Key
            $pageData = $pageEntry.Value
            
            $pageDisplayName = if ($pageData.displayName) { $pageData.displayName } else { "(Pas de nom)" }
            
            Write-Host "  Analyse de la page: '$pageDisplayName' ($pageName)" -ForegroundColor Yellow
            
            # Check each visual on the page
            foreach ($visualEntry in $pageData.visuals.GetEnumerator()) {
                $visualId = $visualEntry.Key
                $visualData = $visualEntry.Value
                
                # Only process slicers
                if ($visualData._extracted_type -eq "slicer") {
                    $groupName = Get-SlicerGroupName -visual $visualData -pageData $pageData
                    $fieldName = Get-SlicerFieldName -visual $visualData
                    $displayName = Get-SlicerDisplayName -visual $visualData -fallbackField $fieldName -fallbackGroup $groupName
                    $searchInfo = Get-SlicerSearchInfo -visual $visualData

                    Write-Host "    Slicer trouvé: $displayName ($visualId)" -ForegroundColor Gray
                    
                    # Extract slicer properties
                    $hasSelected = Test-SlicerHasSelection -visual $visualData
                    $hasSearch = $searchInfo.HasSearch
                    $isSingleSelect = Test-SlicerIsSingleSelect -visual $visualData
                    
                    # Determine status and message
                    $status = "OK"
                    $messageKey = 'ok'
                    $message = "Le filtre est correctement configuré."
                    $searchText = if ($searchInfo.SearchText) { [string]$searchInfo.SearchText } else { $null }
                    $searchDetail = if ($searchText) { " ('{0}')" -f $searchText } else { "" }
                    $normalizedGroup = if ($groupName) { $groupName.ToString().Trim().ToLowerInvariant() } else { "" }
                    $isMenuFilter = $normalizedGroup -like '*_filter'
                    $isMenuPeriod = $normalizedGroup -like '*_period'

                    # Apply verification rules
                    if ($hasSearch) {
                        if ($isSingleSelect) {
                            $status = "ALERTE"
                            $messageKey = 'radio_search'
                            $message = "Le filtre en mode radio contient du texte dans la loupe$searchDetail (non autorisé)."
                        }
                        elseif ($hasSelected) {
                            $status = "ALERTE"
                            $messageKey = 'selection_search'
                            $message = "Le filtre contient une sélection et du texte dans la loupe$searchDetail."
                        }
                        else {
                            $status = "ALERTE"
                            $messageKey = 'search'
                            $message = "Le filtre contient du texte dans la loupe$searchDetail."
                        }
                    }
                    elseif ($isMenuFilter) {
                        if ($hasSelected) {
                            $status = "ALERTE"
                            $messageKey = 'menu_selected'
                            $message = "Ce filtre du groupe $groupName (terminant par _filter) ne doit pas contenir de sélection."
                        }
                        else {
                            $status = "OK"
                            $messageKey = 'menu_ok'
                            $message = "Ce filtre du groupe $groupName (terminant par _filter) est vide comme attendu."
                        }
                    }
                    elseif ($isMenuPeriod) {
                        # Groups ending with _period: ONLY "Current Year" allowed
                        $isCurrentYearOnly = Test-SlicerIsCurrentYearOnly -visual $visualData

                        if ($isCurrentYearOnly) {
                            $status = "OK"
                            $messageKey = 'menu_period_ok'
                            $message = "Ce filtre du groupe $groupName (terminant par _period) a uniquement 'Current Year' comme attendu."
                        }
                        else {
                            # Extract current values for detailed error message
                            $currentValues = Get-SlicerSelectedValues -visual $visualData

                            if ($currentValues.Count -eq 0) {
                                $status = "ALERTE"
                                $messageKey = 'menu_period_empty'
                                $message = "Ce filtre du groupe $groupName (terminant par _period) doit contenir uniquement 'Current Year' (actuellement vide)."
                            }
                            elseif ($currentValues.Count -gt 1) {
                                $valuesList = $currentValues -join ', '
                                $status = "ALERTE"
                                $messageKey = 'menu_period_multiple'
                                $message = "Ce filtre du groupe $groupName (terminant par _period) doit contenir uniquement 'Current Year' (actuellement: $valuesList)."
                            }
                            else {
                                $status = "ALERTE"
                                $messageKey = 'menu_period_invalid'
                                $message = "Ce filtre du groupe $groupName (terminant par _period) doit contenir uniquement 'Current Year' (actuellement: $($currentValues[0]))."
                            }
                        }
                    }
                    elseif ($isSingleSelect) {
                        if ($hasSelected) {
                            $status = "OK"
                            $messageKey = 'radio_ok'
                            $message = "Le filtre en mode radio a une sélection (autorisé) et pas de recherche."
                        }
                    }
                    else {
                        if ($hasSelected) {
                            $status = "OK"
                            $messageKey = 'selection_allowed'
                            $message = "Le filtre contient une sélection (autorisé)."
                        }
                    }
                    
                    # Create result object
                    $result = [PSCustomObject]@{
                        PageName        = $pageDisplayName
                        DisplayName     = $displayName
                        VisualId        = $visualId
                        FieldName       = if ($fieldName) { $fieldName } else { "" }
                        GroupName       = if ($groupName) { $groupName } else { "" }
                        HasSelected     = $hasSelected
                        HasSearch       = $hasSearch
                        SearchText      = $searchInfo.SearchText
                        IsRadio         = $isSingleSelect
                        Status          = $status
                        Message         = $message
                        MessageKey      = $messageKey
                        MessageDetail   = $searchDetail
                    }
                    
                    $results += $result
                    
                    # Display result
                    $statusColor = if ($status -eq "OK") { "Green" } else { "Red" }
                    Write-Host "      Status: $status - $message" -ForegroundColor $statusColor
                }
            }
        }
        
        Write-Host "=== Résultats de l'analyse ===" -ForegroundColor Green
        $totalSlicers = $results | Where-Object { $_.VisualId -ne "FilterPane" }
        Write-Host "Nombre total de slicers analysés: $($totalSlicers.Count)" -ForegroundColor White
        
        if ($results.Count -gt 0) {
            $alertes = $results | Where-Object { $_.Status -eq "ALERTE" }
            $ok = $results | Where-Object { $_.Status -eq "OK" }
            
            Write-Host "  - OK: $($ok.Count)" -ForegroundColor Green
            Write-Host "  - ALERTES: $($alertes.Count)" -ForegroundColor Red
            
            if ($alertes.Count -gt 0) {
                Write-Host "`nDétail des alertes:" -ForegroundColor Yellow
                foreach ($alerte in $alertes) {
                    $identifier = if ($alerte.VisualId -eq "FilterPane") { $alerte.DisplayName } else { "Visuel '$($alerte.VisualId)'" }
                    Write-Host "  - Page '$($alerte.PageName)' / $($identifier): $($alerte.Message)" -ForegroundColor Red
                }
            }
        } else {
            Write-Host "  Aucun slicer trouvé dans le projet." -ForegroundColor Yellow
        }
        
    }
    catch {
        Write-Error "Erreur lors de l'analyse qualité: $($_.Exception.Message)"
        
        # Return error result
        $errorResult = [PSCustomObject]@{
            PageName        = "ERREUR"
            VisualId        = "N/A"
            GroupName       = "N/A"
            HasSelected     = $false
            HasSearch       = $false
            IsRadio         = $false
            Status          = "ERREUR"
            Message         = "Erreur lors de l'analyse: $($_.Exception.Message)"
        }
        $results += $errorResult
    }
    
    return $results
}

# Function to export check results to HTML (bonus feature for independent testing)
Function Export-CheckResultsToHTML {
    param(
        [PSCustomObject[]] $checkResults,
        [string] $outputPath
    )
    
    $htmlContent = @"
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Rapport de Vérification Qualité - Slicers Power BI</title>
<style>
body { font-family: 'Segoe UI', Arial, sans-serif; margin: 20px; background-color: #f7f9fc; }
h1 { color: #2c3e50; text-align: center; border-bottom: 3px solid #3498db; padding-bottom: 10px; }
table { border-collapse: collapse; width: 100%; margin-top: 20px; background-color: white; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }
th { background-color: #34495e; color: white; padding: 12px; text-align: left; font-weight: 600; }
td { border: 1px solid #ddd; padding: 10px; vertical-align: top; }
tr:nth-child(even) { background-color: #f8f9fa; }
tr:hover { background-color: #e3f2fd; }
.ok { background-color: #d4edda !important; color: #155724; font-weight: bold; }
.alerte { background-color: #f8d7da !important; color: #721c24; font-weight: bold; }
.erreur { background-color: #f5c6cb !important; color: #721c24; font-weight: bold; }
</style>
</head>
<body>
<h1>Rapport de Vérification Qualité - Slicers Power BI</h1>
<p><strong>Généré le:</strong> $(Get-Date -Format "dd/MM/yyyy HH:mm:ss")</p>
<p><strong>Nombre de slicers analysés:</strong> $($checkResults.Count)</p>
<table>
<thead>
<tr>
<th>Page</th>
<th>ID Visuel</th>
<th>Nom de Groupe</th>
<th>A une Sélection</th>
<th>A une Recherche</th>
<th>Mode Radio</th>
<th>Status</th>
<th>Message</th>
</tr>
</thead>
<tbody>
"@

    foreach ($result in $checkResults) {
        $statusClass = switch ($result.Status) {
            "OK" { "ok" }
            "ALERTE" { "alerte" }
            "ERREUR" { "erreur" }
            default { "" }
        }
        
        $htmlContent += @"
<tr>
<td>$([System.Web.HttpUtility]::HtmlEncode($result.PageName))</td>
<td>$([System.Web.HttpUtility]::HtmlEncode($result.VisualId))</td>
<td>$([System.Web.HttpUtility]::HtmlEncode($result.GroupName))</td>
<td>$($result.HasSelected)</td>
<td>$($result.HasSearch)</td>
<td>$($result.IsRadio)</td>
<td class="$statusClass">$([System.Web.HttpUtility]::HtmlEncode($result.Status))</td>
<td>$([System.Web.HttpUtility]::HtmlEncode($result.Message))</td>
</tr>
"@
    }

    $htmlContent += @"
</tbody>
</table>
</body>
</html>
"@

    try {
        $htmlContent | Out-File -FilePath $outputPath -Encoding UTF8
        Write-Host "Rapport HTML généré: $outputPath" -ForegroundColor Green
        return $outputPath
    }
    catch {
        Write-Error "Erreur lors de la génération du rapport HTML: $($_.Exception.Message)"
        return $null
    }
}

# Functions are now available for external use via dot sourcing
