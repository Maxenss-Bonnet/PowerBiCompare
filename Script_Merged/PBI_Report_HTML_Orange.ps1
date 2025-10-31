Function Encode-HtmlContent {
    param([string] $Text)
    if ($null -eq $Text) { return "" }
    return [System.Web.HttpUtility]::HtmlEncode($Text)
}

Function Encode-HtmlAttribute {
    param([string] $Text)
    return Encode-HtmlContent $Text
}

Function Convert-ElementTypeToEnglish {
    param([string] $Type)
    switch ($Type) {
        "Page" { return "Page" }
        "Visuel" { return "Visual" }
        "Synchronisation" { return "Synchronization" }
        "Signet" { return "Bookmark" }
        "Configuration" { return "Configuration" }
        "Theme" { return "Theme" }
        "Bouton" { return "Button" }
        default { return $Type }
    }
}

Function Convert-ValueToEnglish {
    param([string] $Value)
    switch ($Value) {
        "Ancien" { return "Previous" }
        "Nouveau" { return "New" }
        "Aucun filtre" { return "No filter" }
        "Aucun champ" { return "No field" }
        "Absent" { return "Absent" }
        "Present" { return "Present" }
        "Pas de nom" { return "No name" }
        # displayOption values
        "FitToPage" { return "Fit to Page" }
        "Ajuster à la page" { return "Fit to Page" }
        "FitToWidth" { return "Fit to Width" }
        "Ajuster à la largeur" { return "Fit to Width" }
        "ActualSize" { return "Actual Size" }
        "Taille réelle" { return "Actual Size" }
        # verticalAlignment values
        "Top" { return "Top" }
        "Haut" { return "Top" }
        "Middle" { return "Middle" }
        "Centre" { return "Middle" }
        "Bottom" { return "Bottom" }
        "Bas" { return "Bottom" }
        # Interaction types
        "Filtrer" { return "Filter" }
        "Mettre en surbrillance" { return "Highlight" }
        "Aucun" { return "None" }
        "Non definie" { return "Not defined" }
        "Non définie" { return "Not defined" }
        "Supprimee" { return "Removed" }
        "Supprimée" { return "Removed" }
        default { return $Value }
    }
}

Function Convert-AdditionalInfoToEnglish {
    param([string] $Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $Text }
    $result = $Text
    $map = @{
        "Ancien:" = "Previous:";
        "Nouveau:" = "New:";
        "Taille" = "Size";
        "Visibilite" = "Visibility";
        "Propriete filtre" = "Filter property";
        "Tri des filtres" = "Filter sorting";
        "Lien activé" = "Link enabled";
        "Type" = "Type";
        "Champs" = "Fields";
        "Methode creation" = "Creation method";
        "Cible navigation" = "Navigation target";
        "Position" = "Position";
        "Taille:" = "Size:";
        "Visibilite:" = "Visibility:";
        "Ancien" = "Previous";
        "Nouveau" = "New";
        "Aucun filtre" = "No filter";
        # displayOption translations
        "FitToPage" = "Fit to Page";
        "Ajuster à la page" = "Fit to Page";
        "FitToWidth" = "Fit to Width";
        "Ajuster à la largeur" = "Fit to Width";
        "ActualSize" = "Actual Size";
        "Taille réelle" = "Actual Size";
        # verticalAlignment translations
        "Haut" = "Top";
        "Centre" = "Middle";
        "Bas" = "Bottom";
        # Interaction type translations
        "Filtrer" = "Filter";
        "Mettre en surbrillance" = "Highlight";
        "Aucun" = "None";
        "Non definie" = "Not defined";
        "Non définie" = "Not defined";
        "Supprimee" = "Removed";
        "Supprimée" = "Removed";
    }
    foreach ($entry in $map.GetEnumerator()) {
        $result = $result -replace [regex]::Escape($entry.Key), $entry.Value
    }
    return $result
}

Function Sanitize-AdditionalInfoForEndUser {
    <#
    .SYNOPSIS
    Nettoie le texte AdditionalInfo pour l'affichage utilisateur final.
    
    .DESCRIPTION
    Supprime les IDs techniques et nettoie le formatage pour une présentation propre dans le HTML.
    
    .PARAMETER Text
    Le texte à nettoyer
    #>
    param([string] $Text)
    
    if ([string]::IsNullOrWhiteSpace($Text)) {
        return ""
    }
    
    # Suppression des IDs techniques entre parenthèses (ex: "(a1b2c3d4e5f6)")
    $result = $Text -replace '\([a-f0-9]{16,}\)', ''
    
    # Suppression des doubles espaces
    $result = $result -replace '\s{2,}', ' '
    
    # Nettoyage des espaces avant ponctuation
    $result = $result -replace '\s+([,;:\.])', '$1'
    
    # Trim
    $result = $result.Trim()
    
    return $result
}

Function Build-LocalizedCell {
    param(
        [string] $ContentFr,
        [string] $ContentEn,
        [string] $Attributes = ""
    )

    if ([string]::IsNullOrEmpty($ContentEn)) {
        $ContentEn = $ContentFr
    }

    $attrFragment = ""
    if ($Attributes) {
        $attrFragment = " " + $Attributes
    }

    $encodedFr = Encode-HtmlAttribute $ContentFr
    $encodedEn = Encode-HtmlAttribute $ContentEn
    $display = Encode-HtmlContent $ContentFr

    return "<td$attrFragment data-i18n-fr='$encodedFr' data-i18n-en='$encodedEn'>$display</td>"
}

Function Build-LocalizedCellRaw {
    param(
        [string] $ContentFr,
        [string] $ContentEn,
        [string] $Attributes = ""
    )

    if ([string]::IsNullOrWhiteSpace($ContentFr)) {
        $ContentFr = "<span class='value-empty'>-</span>"
    }

    if ([string]::IsNullOrWhiteSpace($ContentEn)) {
        $ContentEn = "<span class='value-empty'>-</span>"
    }

    $attrFragment = ""
    if ($Attributes) {
        $attrFragment = " " + $Attributes
    }

    $encodedFr = Encode-HtmlAttribute $ContentFr
    $encodedEn = Encode-HtmlAttribute $ContentEn

    return "<td$attrFragment data-i18n-fr='$encodedFr' data-i18n-en='$encodedEn'>$ContentFr</td>"
}

Function Get-DifferenceCssClass {
    param([string] $DifferenceType)

    switch ($DifferenceType) {
        "Ajoute" { return "ok" }
        "Added" { return "ok" }
        "Supprime" { return "alerte" }
        "Supprimee" { return "alerte" }
        "Removed" { return "alerte" }
        "Modifie" { return "different" }
        "Modified" { return "different" }
        default { return "" }
    }
}

Function Get-DiffValuePair {
    param(
        [string] $Value,
        [string] $FallbackFr = "-",
        [string] $FallbackEn = "-"
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return @{ fr = $FallbackFr; en = $FallbackEn }
    }

    $cleanFr = Remove-TechnicalMarkers -Text $Value
    if ([string]::IsNullOrWhiteSpace($cleanFr)) {
        $cleanFr = $FallbackFr
    }

    $translated = Convert-ValueToEnglish $Value
    $cleanEn = Remove-TechnicalMarkers -Text $translated
    if ([string]::IsNullOrWhiteSpace($cleanEn)) {
        $cleanEn = if ([string]::IsNullOrWhiteSpace($translated)) { $FallbackEn } else { $translated }
    }

    if ([string]::IsNullOrWhiteSpace($cleanEn)) {
        $cleanEn = $cleanFr
    }

    return @{ fr = $cleanFr; en = $cleanEn }
}

Function Get-DefaultComparisonColumns {
    param(
        [string] $NameHeaderKey,
        [string] $NameHeaderText,
        [string] $TypeHeaderKey,
        [string] $TypeHeaderText,
        [string] $ValueHeaderKey,
        [string] $ValueHeaderText
    )

    return @(
        @{ Key = 'name'; OldHeaderKey = $NameHeaderKey; OldHeaderText = $NameHeaderText; NewHeaderKey = $NameHeaderKey; NewHeaderText = $NameHeaderText; OldAttributes = "class='cell-name'"; NewAttributes = "class='cell-name'" },
        @{ Key = 'type'; OldHeaderKey = $TypeHeaderKey; OldHeaderText = $TypeHeaderText; NewHeaderKey = $TypeHeaderKey; NewHeaderText = $TypeHeaderText; OldAttributes = "class='cell-type'"; NewAttributes = "class='cell-type'" },
        @{ Key = 'value'; OldHeaderKey = $ValueHeaderKey; OldHeaderText = $ValueHeaderText; NewHeaderKey = $ValueHeaderKey; NewHeaderText = $ValueHeaderText; OldAttributes = "class='cell-value'"; NewAttributes = "class='cell-value'" }
    )
}

Function New-RichCellContent {
    param(
        [string] $PrimaryFr,
        [string] $PrimaryEn,
        [string[]] $NotesFr = @(),
        [string[]] $NotesEn = @(),
        [string[]] $TagsFr = @(),
        [string[]] $TagsEn = @()
    )

    function New-RichCellContent_Internal {
        param(
            [string] $Primary,
            [string[]] $Notes,
            [string[]] $Tags
        )

        $segments = @()

        if (-not [string]::IsNullOrWhiteSpace($Primary) -and $Primary -ne "-") {
            $segments += "<span class='value-main'>$( [System.Web.HttpUtility]::HtmlEncode($Primary) )</span>"
        }

        foreach ($tag in $Tags) {
            if (-not [string]::IsNullOrWhiteSpace($tag)) {
                $segments += "<span class='value-tag'>$( [System.Web.HttpUtility]::HtmlEncode($tag) )</span>"
            }
        }

        foreach ($note in $Notes) {
            if (-not [string]::IsNullOrWhiteSpace($note)) {
                $segments += "<span class='value-meta'>$( [System.Web.HttpUtility]::HtmlEncode($note) )</span>"
            }
        }

        if ($segments.Count -eq 0) {
            return "<span class='value-empty'>-</span>"
        }

        return ($segments -join "")
    }

    $frContent = New-RichCellContent_Internal -Primary $PrimaryFr -Notes $NotesFr -Tags $TagsFr
    $enContent = New-RichCellContent_Internal -Primary $PrimaryEn -Notes $NotesEn -Tags $TagsEn

    return @{ fr = $frContent; en = $enContent }
}

Function Build-TwoBlockTable {
    param(
        [string] $TableId,
        [string] $TitleKey,
        [string] $TitleText,
        [string] $EmptyKey,
        [string] $EmptyText,
        [array] $Columns,
        [array] $Rows
    )

    if (-not $Columns) {
        return ""
    }

    $columnCount = $Columns.Count
    $totalColspan = ($columnCount * 2) + 1

    $tableHeader = "<tr class='table-header-row table-header-blocks'>"
    $tableHeader += "<th colspan='$columnCount' data-i18n-key='semantic.tables.headers.old_version'>Ancienne version</th>"
    $tableHeader += "<th data-i18n-key='semantic.tables.headers.status'>Statut</th>"
    $tableHeader += "<th colspan='$columnCount' data-i18n-key='semantic.tables.headers.new_version'>Nouvelle version</th>"
    $tableHeader += "</tr>"

    $tableHeader += "<tr class='table-header-row table-subheader-row'>"
    foreach ($col in $Columns) {
        $oldHeaderKey = $col.OldHeaderKey
        $oldHeaderText = $col.OldHeaderText
        $tableHeader += "<th data-i18n-key='$oldHeaderKey'>$oldHeaderText</th>"
    }
    $tableHeader += "<th data-i18n-key='semantic.tables.headers.status'>Statut</th>"
    foreach ($col in $Columns) {
        $newHeaderKey = $col.NewHeaderKey
        $newHeaderText = $col.NewHeaderText
        $tableHeader += "<th data-i18n-key='$newHeaderKey'>$newHeaderText</th>"
    }
    $tableHeader += "</tr>"

    $html = "<table id=""$TableId"" class=""hidden responsive-table two-block-table"">$tableHeader"

    if (-not $Rows -or $Rows.Count -eq 0) {
        $html += "<tr class='ok'><td colspan='$totalColspan' data-i18n-key='$EmptyKey'>$EmptyText</td></tr>"
        $html += "</table>`n"
        return $html
    }

    foreach ($row in $Rows) {
        $rowClass = "two-block-row"
        $status = if ($row.Status) { $row.Status } else { "-" }

        $html += "<tr class='$rowClass'>"
        
        # OLD VERSION CELLS (Ancienne version)
        foreach ($col in $Columns) {
            $cell = if ($row.Old.ContainsKey($col.Key)) { $row.Old[$col.Key] } else { $null }
            if (-not $cell) {
                $cell = @{ fr = $null; en = $null }
            }
            
            # Apply background color based on Figma rules
            $cellStyle = ""
            if ($status -eq "Supprime" -or $status -eq "Removed") {
                # Red background ONLY in old version cells for removed items
                $cellStyle = " style='background-color: var(--tbl-row-removed);'"
            }
            elseif ($status -eq "Modifie" -or $status -eq "Modified") {
                # Orange background on OLD cells that changed (matching NEW version logic)
                $newCell = if ($row.New.ContainsKey($col.Key)) { $row.New[$col.Key] } else { $null }
                $oldValue = if ($cell) { $cell.fr } else { "" }
                $newValue = if ($newCell) { $newCell.fr } else { "" }
                
                # Compare values to see if THIS specific cell changed
                if ($oldValue -ne $newValue) {
                    $cellStyle = " style='background-color: var(--tbl-cell-changed);'"
                }
            }
            
            # For "Added" items, leave OLD version cells completely empty
            if ($status -eq "Ajoute" -or $status -eq "Added") {
                $cell = @{ fr = ""; en = "" }
            }
            
            $oldAttributes = $col.OldAttributes
            if ($cellStyle) {
                $oldAttributes = if ($oldAttributes) { $oldAttributes + $cellStyle } else { $cellStyle.Trim() }
            }
            
            $html += Build-LocalizedCellRaw $cell.fr $cell.en $oldAttributes
        }

        # STATUS CELL (NEVER has background color per Figma)
        $encodedStatus = Encode-HtmlAttribute $status
        $html += "<td class='diff-type-cell two-block-status' data-diff-badge='true' data-i18n-diff-type='$encodedStatus'>$status</td>"

        # NEW VERSION CELLS (Nouvelle version)
        foreach ($col in $Columns) {
            $cell = if ($row.New.ContainsKey($col.Key)) { $row.New[$col.Key] } else { $null }
            if (-not $cell) {
                $cell = @{ fr = $null; en = $null }
            }
            
            # Apply background color based on Figma rules
            $cellStyle = ""
            if ($status -eq "Ajoute" -or $status -eq "Added") {
                # Green background ONLY in new version cells for added items
                $cellStyle = " style='background-color: var(--tbl-row-added);'"
            }
            elseif ($status -eq "Modifie" -or $status -eq "Modified") {
                # Orange background ONLY on cells that actually changed
                $oldCell = if ($row.Old.ContainsKey($col.Key)) { $row.Old[$col.Key] } else { $null }
                $oldValue = if ($oldCell) { $oldCell.fr } else { "" }
                $newValue = if ($cell) { $cell.fr } else { "" }
                
                # Compare values to see if THIS specific cell changed
                if ($oldValue -ne $newValue) {
                    $cellStyle = " style='background-color: var(--tbl-cell-changed);'"
                }
            }
            
            # For "Removed" items, leave NEW version cells completely empty
            if ($status -eq "Supprime" -or $status -eq "Removed") {
                $cell = @{ fr = ""; en = "" }
            }
            
            $newAttributes = $col.NewAttributes
            if ($cellStyle) {
                $newAttributes = if ($newAttributes) { $newAttributes + $cellStyle } else { $cellStyle.Trim() }
            }
            
            $html += Build-LocalizedCellRaw $cell.fr $cell.en $newAttributes
        }

        $html += "</tr>"
    }

    $html += "</table>`n"
    return $html
}

# Helper to strip technical identifiers from display text (GUID, technical labels, etc.)
Function Remove-TechnicalMarkers {
    param([string] $Text)

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return ""
    }

    $clean = $Text.Trim()

    $patterns = @(
        '\s*\((?:id|identifiant|identifiant technique|technical id|visual id|guid)[^)]*\)\s*$',
        '\s*[-–—]?\s*(?:id|identifiant|identifiant technique|technical id|visual id|guid)\s*[:=]?\s*[0-9a-fA-F\-]{3,}\s*$',
        '\s*\(([0-9a-fA-F]{6,}|[0-9a-fA-F\-]{6,})\)\s*$',
        '\s*\(ReportSection[0-9A-Za-z]+\)\s*$',
        '\s*\(([0-9A-Za-z\-]{10,})\)\s*$'
    )

    foreach ($pattern in $patterns) {
        $clean = [regex]::Replace($clean, $pattern, '', 'IgnoreCase')
    }

    # Also remove technical parentheses anywhere, not only at the end
    $clean = [regex]::Replace($clean, '\(ReportSection[0-9A-Za-z]+\)', '', 'IgnoreCase')
    $clean = [regex]::Replace($clean, '\(([0-9A-Za-z\-]{10,})\)', '', 'IgnoreCase')

    $clean = [regex]::Replace($clean, '\s{2,}', ' ')
    $clean = [regex]::Replace($clean, '^[\s\-\u2013\u2014:]+|[\s\-\u2013\u2014:]+$', '')

    if ($clean -match '^[0-9a-fA-F\-]{6,}$') {
        return ""
    }

    return $clean.Trim()
}

Function Get-ElementFallbackLabel {
    param([string] $ElementType)

    switch ($ElementType) {
        "Page"   { return @{ fr = "Page sans titre";    en = "Untitled page" } }
        "Visuel" { return @{ fr = "Visuel sans titre";  en = "Untitled visual" } }
        "Bouton" { return @{ fr = "Bouton sans titre";  en = "Untitled button" } }
        "Signet" { return @{ fr = "Signet sans titre";  en = "Untitled bookmark" } }
        default   { return @{ fr = "Nom non disponible"; en = "Name unavailable" } }
    }
}

Function Get-DisplayNamePair {
    param(
        [string] $DisplayName,
        [string] $TechnicalName,
        [string] $ElementType
    )

    $candidate = Remove-TechnicalMarkers $DisplayName
    if ([string]::IsNullOrWhiteSpace($candidate)) {
        $candidate = Remove-TechnicalMarkers $TechnicalName
    }

    if ([string]::IsNullOrWhiteSpace($candidate)) {
        return Get-ElementFallbackLabel $ElementType
    }

    return @{ fr = $candidate; en = $candidate }
}

Function Get-SafeTextPair {
    param(
        [string] $Text,
        [string] $DefaultFr = "-",
        [string] $DefaultEn = "-"
    )

    $clean = Remove-TechnicalMarkers $Text
    if ([string]::IsNullOrWhiteSpace($clean)) {
        return @{ fr = $DefaultFr; en = $DefaultEn }
    }

    return @{ fr = $clean; en = $clean }
}

#==========================================================================================================================================

# Generate hierarchical sync group table (replaces GenerateSynchronizationTable)
Function Generate-CauseConsequenceTable {
    param([hashtable] $causesIndex)

    if ($null -eq $causesIndex -or $causesIndex.Causes.Count -eq 0) {
        return @"
<table id="table_synchronization" class="hidden responsive-table">
<tr class="table-title-row"><th colspan="6" data-i18n-key="tables.table_synchronization.title">Synchronisation des Slicers (CAUSE → CONSEQUENCES)</th></tr>
<tr class="ok"><td colspan="6" data-i18n-key="tables.table_synchronization.empty">Aucune modification de synchronisation detectee</td></tr>
</table>
"@
    }

    $html = @"
<table id="table_synchronization" class="hidden responsive-table">
<tr class="table-title-row"><th colspan="6" data-i18n-key="tables.table_synchronization.title">Synchronisation des Slicers (CAUSE → CONSEQUENCES)</th></tr>
<tr class="table-header-row">
    <th data-i18n-key="tables.table_synchronization.headers.element">Slicer</th>
    <th data-i18n-key="tables.table_synchronization.headers.page">Page</th>
    <th data-i18n-key="tables.table_synchronization.headers.syncGroup">Groupe Sync</th>
    <th data-i18n-key="tables.table_synchronization.headers.action">Action</th>
    <th data-i18n-key="tables.table_synchronization.headers.difference">Différence</th>
    <th data-i18n-key="tables.table_synchronization.headers.impact">Impact / Consequences</th>
</tr>
"@

    $causes = $causesIndex.Causes

    if ($null -ne $causes -and $causes.Count -gt 0) {
        foreach ($causeKey in $causes.Keys) {
            $cause = $causes[$causeKey]
            $causeSlicer = $cause.CauseSlicer
            $consequenceCount = $cause.Consequences.Count

            # Determine CSS class based on cause type
            $causeRowClass = switch ($cause.CauseType) {
                "Added" { "cause-added" }
                "Modified" { "cause-modified" }
                "Removed" { "cause-removed" }
                default { "cause-header" }
            }

            # Format cause type display
            $causeTypeDisplay = switch ($cause.CauseType) {
                "Added" { "Ajouté" }
                "Modified" { "Modifié" }
                "Removed" { "Supprimé" }
                default { $cause.CauseType }
            }

            # Format difference description
            $differenceDisplay = ""

            if ($cause.CauseType -eq "Added") {
                # Check if this is a visibility change (isHidden)
                if ($cause.SyncGroupName -eq "Visibilité") {
                    # Visibility: showing a hidden slicer
                    $differenceDisplay = "Visuel affiché sur la page '<b>$($causeSlicer.PageDisplayName)</b>'"
                } else {
                    # Synchronisation ajoutée - distinguish between full sync group creation and page additions
                    $syncedPages = @($cause.Consequences | Where-Object { $_.ActionType -eq "Synchronized" })

                    if ($syncedPages.Count -gt 0) {
                        # Added pages to existing sync group
                        $pageNames = ($syncedPages | ForEach-Object { "'<b>$($_.PageDisplayName)</b>'" }) -join ", "
                        $pageWord = if ($syncedPages.Count -eq 1) { "page" } else { "pages" }
                        $differenceDisplay = "Ajout de synchronisation avec $($syncedPages.Count) $pageWord : $pageNames"
                    } elseif ($cause.NewConfig) {
                        # New sync group created
                        $syncOptions = @()
                        if ($cause.NewConfig.FieldChanges -eq "True") { $syncOptions += "synchronisation des champs" }
                        if ($cause.NewConfig.FilterChanges -eq "True") { $syncOptions += "synchronisation des filtres" }
                        $optionsText = if ($syncOptions.Count -gt 0) { $syncOptions -join " et " } else { "aucune option" }
                        $differenceDisplay = "Synchronisation ajoutée pour le groupe '<b>$($cause.NewConfig.SyncGroupName)</b>' avec $optionsText"
                    } elseif ($cause.AdditionalInfo) {
                        $differenceDisplay = $cause.AdditionalInfo
                    } else {
                        $differenceDisplay = "Synchronisation ajoutée"
                    }
                }

            } elseif ($cause.CauseType -eq "Removed") {
                # Check if this is a visibility change (isHidden)
                if ($cause.SyncGroupName -eq "Visibilité") {
                    # Visibility: hiding a visible slicer
                    $differenceDisplay = "Visuel masqué sur la page '<b>$($causeSlicer.PageDisplayName)</b>'"
                } else {
                    # Synchronisation supprimée - distinguish between full sync group removal and page removals
                    $desyncedPages = @($cause.Consequences | Where-Object { $_.ActionType -eq "Desynchronized" })

                    if ($desyncedPages.Count -gt 0) {
                        # Removed pages from sync group
                        $pageNames = ($desyncedPages | ForEach-Object { "'<b>$($_.PageDisplayName)</b>'" }) -join ", "
                        $pageWord = if ($desyncedPages.Count -eq 1) { "page" } else { "pages" }
                        $differenceDisplay = "Suppression de synchronisation avec $($desyncedPages.Count) $pageWord : $pageNames"
                    } elseif ($cause.OldConfig) {
                        # Full sync group removed
                        $syncOptions = @()
                        if ($cause.OldConfig.FieldChanges -eq "True") { $syncOptions += "synchronisation des champs" }
                        if ($cause.OldConfig.FilterChanges -eq "True") { $syncOptions += "synchronisation des filtres" }
                        $optionsText = if ($syncOptions.Count -gt 0) { $syncOptions -join " et " } else { "aucune option" }
                        $differenceDisplay = "Synchronisation supprimée du groupe '<b>$($cause.OldConfig.SyncGroupName)</b>' (avait $optionsText)"
                    } elseif ($cause.AdditionalInfo) {
                        $differenceDisplay = $cause.AdditionalInfo
                    } else {
                        $differenceDisplay = "Synchronisation supprimée"
                    }
                }

            } elseif ($cause.CauseType -eq "Modified") {
                # Synchronisation modifiée
                $changes = @()

                if ($cause.OldConfig -and $cause.NewConfig) {
                    # Vérifier si le groupe a changé
                    if ($cause.OldConfig.SyncGroupName -ne $cause.NewConfig.SyncGroupName) {
                        $changes += "groupe changé de '<b>$($cause.OldConfig.SyncGroupName)</b>' vers '<b>$($cause.NewConfig.SyncGroupName)</b>'"
                    }

                    # Vérifier si FieldChanges a changé
                    if ($cause.OldConfig.FieldChanges -ne $cause.NewConfig.FieldChanges) {
                        if ($cause.NewConfig.FieldChanges -eq "True") {
                            $changes += "synchronisation des champs <b>activée</b>"
                        } else {
                            $changes += "synchronisation des champs <b>désactivée</b>"
                        }
                    }

                    # Vérifier si FilterChanges a changé
                    if ($cause.OldConfig.FilterChanges -ne $cause.NewConfig.FilterChanges) {
                        if ($cause.NewConfig.FilterChanges -eq "True") {
                            $changes += "synchronisation des filtres <b>activée</b>"
                        } else {
                            $changes += "synchronisation des filtres <b>désactivée</b>"
                        }
                    }
                }

                if ($changes.Count -gt 0) {
                    $differenceDisplay = ($changes -join ", ")
                } elseif ($cause.AdditionalInfo) {
                    # Fallback sur AdditionalInfo si disponible
                    $differenceDisplay = $cause.AdditionalInfo
                } else {
                    # Fallback sur le nombre de conséquences
                    $conseqCount = if ($cause.Consequences) { $cause.Consequences.Count } else { 0 }
                    if ($conseqCount -gt 0) {
                        $differenceDisplay = "Synchronisation modifiée : $conseqCount page(s) affectée(s)"
                    } else {
                        $differenceDisplay = "Synchronisation modifiée"
                    }
                }
            } else {
                if ($cause.AdditionalInfo) {
                    $differenceDisplay = $cause.AdditionalInfo
                } else {
                    $differenceDisplay = "Modification de synchronisation détectée"
                }
            }

            # Impact badge
            $impactBadge = if ($consequenceCount -gt 0) {
                "<span class='impact-badge-consequence'>$consequenceCount page(s) impactée(s)</span>"
            } else {
                "<span class='impact-badge-neutral'>Aucune conséquence</span>"
            }

            # Format action display with color
            $actionDisplay = switch ($cause.CauseType) {
                "Added" { "<span style='color:#27ae60;font-weight:bold;'>Ajout</span>" }
                "Modified" { "<span style='color:#f39c12;font-weight:bold;'>Modification</span>" }
                "Removed" { "<span style='color:#e74c3c;font-weight:bold;'>Suppression</span>" }
                default { "<span style='color:#7f8c8d;'>$($cause.CauseType)</span>" }
            }

            # Encode HTML
            $slicerFieldDisplay = Encode-HtmlContent $causeSlicer.FieldName
            $slicerPageDisplay = Encode-HtmlContent $causeSlicer.PageDisplayName
            $syncGroupDisplay = Encode-HtmlContent $causeSlicer.SyncGroupName

            $html += @"
<tr class='$causeRowClass' data-cause-key='$causeKey' onclick='toggleCause(this)'>
    <td style='font-weight:bold;cursor:pointer;'><span class='expand-icon'>▶</span> $causeTypeDisplay $slicerFieldDisplay</td>
    <td style='font-weight:bold;'>$slicerPageDisplay</td>
    <td style='font-weight:bold;color:#2980b9;'>$syncGroupDisplay</td>
    <td>$actionDisplay</td>
    <td style='max-width:300px;word-wrap:break-word;font-size:11px;'>$differenceDisplay</td>
    <td>$impactBadge</td>
</tr>
"@

            # Niveau 2: CONSEQUENCES
            if ($consequenceCount -gt 0) {
                foreach ($consequence in $cause.Consequences) {
                    $conseqClass = switch ($consequence.ActionType) {
                        "Synchronized" { "consequence-synchronized" }
                        "Desynchronized" { "consequence-desynchronized" }
                        "Hidden" { "consequence-hidden" }
                        "Shown" { "consequence-shown" }
                        default { "consequence-row" }
                    }

                    $conseqIcon = switch ($consequence.ActionType) {
                        "Synchronized" { "→" }
                        "Desynchronized" { "×" }
                        "Hidden" { "×" }
                        "Shown" { "+" }
                        default { "•" }
                    }

                    $conseqPageDisplay = Encode-HtmlContent $consequence.PageDisplayName
                    $conseqFieldDisplay = Encode-HtmlContent $consequence.FieldName
                    $conseqDescDisplay = Encode-HtmlContent $consequence.Description

                    $html += @"
<tr class='consequence-row hidden $conseqClass' data-parent-cause='$causeKey'>
    <td style='padding-left:40px;font-size:11px;'>└─► $conseqIcon $conseqFieldDisplay</td>
    <td style='font-style:italic;color:#7f8c8d;'>$conseqPageDisplay</td>
    <td style='font-style:italic;color:#95a5a6;'>$syncGroupDisplay</td>
    <td colspan='2' style='font-size:11px;color:#7f8c8d;'>$conseqDescDisplay</td>
    <td style='font-size:10px;color:#95a5a6;'>$($consequence.ActionType)</td>
</tr>
"@
                }
            } else {
                # No consequences
                $html += @"
<tr class='consequence-row hidden consequence-none' data-parent-cause='$causeKey'>
    <td colspan='6' style='padding-left:40px;font-size:11px;color:#95a5a6;font-style:italic;'>Aucune conséquence sur d'autres pages</td>
</tr>
"@
            }
        }
    }

    $html += "</table>"
    return $html
}

#==========================================================================================================================================

# HTML report generation function
Function BuildReportHTMLReport_Orange {
    param (
        [ReportDifference[]] $differences,
        [PSCustomObject[]] $checkResults = @(),
        [string] $outputFolder,
        [PSCustomObject] $semanticComparisonResult = $null,
        [string] $configPath = ""
    )

    Write-Host "  [BuildReportHTMLReport_Orange] Semantic data received: $($null -ne $semanticComparisonResult)" -ForegroundColor Cyan
    if ($semanticComparisonResult) {
        Write-Host "  [BuildReportHTMLReport_Orange] Semantic results type: $($semanticComparisonResult.GetType().Name)" -ForegroundColor Cyan
        Write-Host "  [BuildReportHTMLReport_Orange] Semantic results count: $($semanticComparisonResult.Count)" -ForegroundColor Cyan
    } else {
        Write-Host "  [BuildReportHTMLReport_Orange] WARNING: semanticComparisonResult is NULL or empty!" -ForegroundColor Yellow
    }

    $reportFinal = @'
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Rapport de Comparaison des Rapports Power BI - Orange</title>
<style>
/* ========== ORANGE DESIGN SYSTEM ========== */
* {
    margin: 0;
    padding: 0;
    box-sizing: border-box;
}

:root {
    --tbl-text: #000000;
    --tbl-text-header: #333333;
    --tbl-border: #cccccc;
    --tbl-header-bg: #f4f4f4;
    --tbl-row-added: #ecfdef;
    --tbl-row-removed: #fde5e6;
    --tbl-cell-changed: #fffae7;
    --status-added: #3de35a;
    --status-removed: #e70002;
    --status-modified: #ffcd0b;
    --scroll-track: #f4f4f4;
    --scroll-thumb: #555555;
    --font-family: "Helvetica Neue", Helvetica, Arial, sans-serif;
    --brand-primary: #ff7900;
    --brand-primary-dark: #f15e00;
    --brand-primary-rgb: 255, 121, 0;
    --brand-surface-50: #fff7f0;
    --brand-surface-100: #ffe9d9;
    --neutral-600: #666666;
    --neutral-700: #2c3e50;
    --surface-default: #ffffff;
}

body {
    font-family: var(--font-family);
    color: var(--tbl-text);
    min-height: 100vh;
    display: flex;
    flex-direction: column;
    background-color: var(--surface-default);
}

/* ========== HEADER ORANGE ========== */
.page-header {
    background-color: #000000;
    color: white;
    padding: 20px 30px 15px 30px;
    display: flex;
    justify-content: space-between;
    height: 60px;
	gap: 10px;
}

.header-left {
    display: flex;
    align-items: center;
    gap: 30px;
}

.section-hidden {
    opacity: 0.8;
}

.header-right {
    display: flex;
    align-items: center;
    gap: 10px;
}

.logo {
    width: 30px;
    height: 30px;
    background-color: var(--brand-primary);
    display: flex;
    align-items: center;
    justify-content: center;
    font-weight: bold;
    color: #000;
}

.title {
    width: 295;
    height: 32;
    font-size: 24px;
    font-weight: 400;
    font-style: 75 Bold;
    font-size: 24px;
    line-height: 32px;
    letter-spacing: 0%;
}

.language-switch {
    background-color: transparent;
    border: 0px;
    padding: 4px 8px;
    border-radius: 5px;
    cursor: pointer;
	transition: filter 0.3s ease;
}


.language-switch:hover {
    filter: brightness(70%);
}

.settings-button {
    background-color: transparent;
    border: 2px solid var(--border-subtle);
    padding: 6px 12px;
    border-radius: 8px;
    cursor: pointer;
    transition: all 0.3s ease;
    display: flex;
    align-items: center;
    gap: 6px;
    font-size: 14px;
    color: var(--text-default);
    height: 38px;
}

.settings-button:hover {
    background-color: var(--surface-hovered);
    border-color: var(--brand-primary);
}

.settings-button svg {
    width: 18px;
    height: 18px;
}

/* ========== MAIN CONTENT ========== */
.main-content {
    flex: 1;
    padding: 20px 30px;
}

.all-buttons {
    min-height: 0;
    margin-bottom: 0;
}

/* ========== BOUTONS PRINCIPAUX ORANGE ========== */
.main-buttons {
    display: flex;
    gap: 10px;
    margin-bottom: 10px;
}

.main-btn {
    background-color: transparent;
    border: 1px solid transparent;
    color: #000000;
    font-weight: 700;
    font-size: 18px;
    line-height: 24px;
    padding: 15px 20px;
    cursor: pointer;
    position: relative;
    transition: all 0.3s ease;
}

.main-btn:hover {
    color: var(--scroll-thumb);
}

.main-btn.active {
    color: #000000;
}

.main-btn.active::after {
    content: '';
    position: absolute;
    bottom: -3px;
    left: 0;
    right: 0;
    height: 5px;
    background-color: var(--brand-primary-dark);
    z-index: 2;
}

/* ========== LIGNE DE SÉPARATION ========== */
.separator {
    height: 2px;
    background-color: #ddd;
    /* Keep the underline overlay but leave a small gap for the chip row */
    margin: -8px 0 12px 0;
    position: relative;
    z-index: 1;
	margin-bottom: 20px;
}
    
/* ========== SOUS-BOUTONS ORANGE ========== */
.sub-buttons-group {
    display: none;
    gap: 12px;
    margin-bottom: 20px;
    flex-wrap: wrap;
}

.sub-buttons-group.active {
    display: flex;
}

.sub-btn {
    background-color: transparent;
    color: var(--neutral-600);
    border: 2px solid var(--tbl-border);
    border-radius: 20px;
    padding: 5px 18px 7px 18px;
    min-width: 54px;
    min-height: 32px;
    font-weight: 400;
    font-size: 16px;
    line-height: 24px;
    cursor: pointer;
    transition: all 0.3s ease;
}

.sub-btn:hover {
    background-color: #f8f9fa;
    border-color: var(--scroll-thumb);
    color: var(--scroll-thumb);
}

.sub-btn.active {
    border-color: var(--brand-primary-dark);
    color: #000000;
}

.sub-btn.active::before {
    content: "✓";
    color: var(--brand-primary-dark);
    font-weight: bold;
    margin-right: 5px;
}

/* Distinguish hierarchical sub-button (Pages & Visuels) */
.sub-btn.has-children {
    position: relative;
    padding-right: 34px; /* room for caret */
}
.sub-btn.has-children::after {
    content: '▸';
    position: absolute;
    right: 12px;
    top: 50%;
    transform: translateY(-50%);
    color: var(--brand-primary-dark);
    font-weight: 900;
    transition: transform 0.2s ease;
}
.sub-btn.has-children.expanded::after {
    content: '▾';
}

/* ========== BOUTONS TROISIÈME NIVEAU ========== */
.third-buttons {
    display: none;
    gap: 12px;
    margin-top: 12px;
    flex-wrap: wrap;
}

.third-buttons.show {
    display: flex;
}

.third-btn {
    background-color: var(--tbl-header-bg);
    color: var(--tbl-text);
    border: none;
    border-radius: 20px;
    padding: 5px 10px 7px 18px;
    min-width: 84px;
    min-height: 32px;
    font-weight: 700;
    font-size: 14px;
    line-height: 20px;
    cursor: pointer;
    transition: all 0.3s ease;
}

.third-btn:hover {
    background-color: var(--brand-surface-50);
    color: var(--scroll-thumb);
}

.third-btn.active,
.third-level-buttons .third-btn.active,
.third-buttons .third-btn.active,
.third-btn.active:hover,
.third-btn.active:focus {
    color: var(--brand-primary-dark) !important;
}

/* ========== TABLES ORANGE ========== */

.table-container {
    overflow-x: auto;
    margin-bottom: 20px;
    max-width: 100%;
}

.responsive-table {
    width: 100%;
    border-collapse: separate;
    border-spacing: 0;
    background-color: #ffffff;
    border: 0.25px solid var(--tbl-border);
    border-top-width: 2px;
    border-top-color: var(--tbl-text-header);
    margin: 16px 0;
    font-family: var(--font-family);
    font-size: 14px;
    line-height: 20px;
    color: var(--tbl-text);
    box-shadow: none;
    table-layout: auto;
}

.responsive-table tr {
    transition: background-color 0.2s ease;
    height: auto;
}

.responsive-table th,
.responsive-table td {
    padding: 10px 8px;
    border-bottom: 0.25px solid var(--tbl-border);
    vertical-align: top;
    word-wrap: break-word;
    overflow-wrap: break-word;
    white-space: pre-wrap;
    max-width: 400px;
    height: auto;
    min-height: 40px;
    overflow: visible;
    text-overflow: clip;
    line-height: 1.5;
}

.responsive-table th + th,
.responsive-table td + td {
    border-left: 0.25px solid var(--tbl-border);
}

.responsive-table tr:last-child td {
    border-bottom: none;
}

.responsive-table .table-title-row th {
    background-color: var(--tbl-header-bg);
    color: var(--tbl-text-header);
    font-weight: 700;
    font-size: 16px;
    padding: 14px 16px;
    text-align: left;
    border-bottom: 0.25px solid var(--tbl-border);
}

.responsive-table .table-header-row th {
    background-color: var(--tbl-header-bg);
    color: var(--tbl-text-header);
    font-weight: 600;
    text-transform: none;
    letter-spacing: 0;
}

.responsive-table tr:hover {
    background-color: var(--brand-surface-50);
}

.two-block-table .table-header-blocks th {
    background-color: var(--tbl-header-bg);
    font-size: 13px;
    font-weight: 600;
    text-transform: none;
}

.two-block-table .table-subheader-row th {
    background-color: var(--surface-default);
    font-size: 12px;
    font-weight: 600;
    color: var(--neutral-600);
}

.two-block-table td {
    vertical-align: top;
}

.two-block-table .two-block-status {
    width: 130px;
    text-align: center;
}

.cell-name {
    min-width: 180px;
    max-width: 350px;
}

.cell-type {
    min-width: 160px;
    max-width: 300px;
}

.cell-value {
    min-width: 220px;
    max-width: 600px;
}

.value-main {
    display: block;
    font-weight: 600;
    color: var(--neutral-700);
    white-space: pre-wrap;
    word-break: break-word;
}

.value-meta {
    display: block;
    margin-top: 4px;
    font-size: 12px;
    color: var(--neutral-600);
}

.value-tag {
    display: inline-block;
    margin-top: 4px;
    padding: 2px 8px;
    border-radius: 12px;
    font-size: 11px;
    font-weight: 600;
    background-color: var(--tbl-header-bg);
    color: var(--neutral-700);
}

.value-empty {
    display: inline-block;
    color: var(--neutral-500);
    font-style: italic;
}

.table-main-header {
    background-color: var(--tbl-header-bg);
    color: var(--tbl-text-header);
    font-weight: 700;
}

td.diff-type-cell,
.responsive-table td[data-diff-badge="true"] {
    text-align: center;
    min-width: 88px;
    font-weight: 600;
}

.diff-badge {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    padding: 1px 10px 2px;
    border-radius: 60px;
    font-size: 12px;
    line-height: 18px;
    font-weight: 700;
    min-width: 72px;
    background-color: var(--tbl-header-bg);
    color: #000000;
}

.diff-badge--added {
    background-color: var(--status-added);
    color: #000000;
}

.diff-badge--removed {
    background-color: var(--status-removed);
    color: #ffffff;
}

.diff-badge--modified {
    background-color: var(--status-modified);
    color: #000000;
}

.diff-badge--neutral {
    background-color: #d9d9d9;
    color: #000000;
}

@media (max-width: 1080px) {
    .responsive-table {
        font-size: 13px;
    }

    .responsive-table th,
    .responsive-table td {
        padding: 8px 6px;
        white-space: normal;
    }
}

/* ========== STATUS COLORS ========== */
/* Row-level colors: ONLY for Quality Check table (exception) */
.responsive-table tr.ok {
    background-color: var(--tbl-row-added);
}

.responsive-table tr.alerte {
    background-color: var(--tbl-row-removed);
}

/* Cell-level colors for Rapport and MDD tables (applied via Build-TwoBlockTable and ReportCreateCompareTable):
   - Ajouté: green background ONLY in Nouvelle version cells
   - Supprimé: red background ONLY in Ancienne version cells  
   - Modifié: orange background ONLY on cells that actually changed
   - Statut column: NEVER has background color
*/

/* ========== CAPSULES ORANGE STATUS ========== */
.capsule {
    background-color: #ffffff;
    display: inline-block;
    text-align: center;
    padding: 8px 16px;
    border-radius: 20px;
    font-weight: bold;
    font-size: 12px;
    letter-spacing: 0.5px;
}

.capsule.statut-ok {
    background-color: var(--status-added) !important;
    color: #000000;
}

.capsule.statut-alerte {
    background-color: var(--status-removed) !important;
    color: #ffffff;
}

.capsule.statut-different {
    background-color: var(--status-modified) !important;
    color: #000000;
}

.capsule.statut-identical {
    background-color: var(--tbl-header-bg) !important;
    color: #000000;
}

.hidden {
    display: none;
}

.section-hidden {
    display: none;
}
/* ========== LEGACY SUPPORT (minimal) ========== */
            table_buttons: {

                title: "Button changes (actionButton)",

                headers: {

                    name: "Button name",

                    page: "Page",

                    property: "Property",

                    change: "Change type",

                    oldValue: "Old value",

                    newValue: "New value",

                    details: "Details",

                    description: "Description"

                },

                empty: "No button changes detected"

            },

h2 {
    color: var(--neutral-700);
    border-bottom: 2px solid var(--brand-primary-dark);
    padding-bottom: 10px;
    margin-bottom: 20px;
}

/* ========== SUMMARY STATS CARDS ========== */
.summary-wrapper {
    margin: 24px 0;
}

.summary-title {
    font-size: 18px;
    font-weight: 600;
    color: var(--neutral-700);
    margin-bottom: 12px;
}

.summary-stats {
    display: flex;
    gap: 20px;
    margin: 20px 0;
    flex-wrap: wrap;
}

.stat-item {
    flex: 1;
    min-width: 150px;
    background: var(--surface-default);
    border-left: 5px solid var(--brand-primary);
    padding: 20px;
    border-radius: 5px;
    box-shadow: 0 2px 5px rgba(var(--brand-primary-rgb), 0.1);
    transition: transform 0.2s ease, box-shadow 0.2s ease;
}

.stat-item:hover {
    transform: translateY(-2px);
    box-shadow: 0 4px 10px rgba(var(--brand-primary-rgb), 0.2);
}

.stat-number {
    font-size: 2.5em;
    color: var(--brand-primary);
    font-weight: bold;
    line-height: 1;
}

.stat-label {
    color: var(--neutral-600);
    margin-top: 8px;
    font-size: 0.9em;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}

.summary-stats--comparison .stat-item--added,
.summary-stats--quality .stat-item--ok {
    border-left-color: var(--status-added);
}

.summary-stats--comparison .stat-item--removed,
.summary-stats--quality .stat-item--error {
    border-left-color: var(--status-removed);
}

.summary-stats--comparison .stat-item--modified,
.summary-stats--quality .stat-item--alert {
    border-left-color: var(--status-modified);
}

.stat-item--added .stat-number,
.stat-item--ok .stat-number {
    color: var(--status-added);
}

.stat-item--removed .stat-number,
.stat-item--error .stat-number {
    color: var(--status-removed);
}

.stat-item--modified .stat-number,
.stat-item--alert .stat-number {
    color: var(--status-modified);
}

/* ========== SEARCH CONTAINER ========== */
.search-container {
    margin: 20px 0;
    display: flex;
    gap: 10px;
    align-items: center;
}

.search-box {
    flex: 1;
    padding: 12px 15px;
    border: 2px solid var(--brand-primary);
    border-radius: 5px;
    font-size: 14px;
    outline: none;
    transition: border-color 0.3s ease, box-shadow 0.3s ease;
}

.search-box:focus {
    border-color: var(--brand-primary-dark);
    box-shadow: 0 0 0 3px rgba(var(--brand-primary-rgb), 0.1);
}

.search-container button {
    background: var(--brand-primary);
    color: white;
    padding: 12px 20px;
    border: none;
    border-radius: 5px;
    cursor: pointer;
    font-weight: 600;
    transition: background-color 0.3s ease;
}

.search-container button:hover {
    background: var(--brand-primary-dark);
}

.search-results {
    margin-top: 10px;
    font-size: 0.9em;
    color: var(--neutral-600);
}

/* ========== QUALITY INFO BUTTON + PANEL ========== */
.quality-info-bar, .visuals-info-bar {
    display: flex;
    justify-content: flex-end;
    margin: 0; /* avoid creating vertical gap above the rules panel */
    /* Use normal flow so it doesn’t overlap language switches */
    position: static;
    align-self: flex-end;
}
.info-btn {
    background: var(--brand-surface-50);
    color: #000;
    border: 2px solid var(--brand-primary);
    border-radius: 16px;
    padding: 6px 12px;
    font-weight: 700;
    font-size: 13px;
    cursor: pointer;
	margin-bottom: 10px;
}
.info-btn:hover {
    background: var(--brand-surface-100);
    border-color: var(--brand-primary-dark);
}
.quality-info-panel {
    background: #FFF8E6;
    border-left: 4px solid var(--brand-primary-dark);
    padding: 16px 20px;
    border-radius: 4px;
    margin: 12px 0 20px 0;
    box-shadow: 0 1px 3px rgba(0,0,0,0.1);
}
.quality-info-panel h3 {
    margin: 0 0 12px 0;
    color: var(--neutral-700);
    font-size: 1.1em;
}
.quality-info-panel ul {
    margin: 0;
    padding-left: 24px;
    list-style-type: disc;
}
.quality-info-panel li {
    margin-bottom: 8px;
    line-height: 1.5;
    color: var(--neutral-600);
}
.quality-info-panel li:last-child {
    margin-bottom: 0;
}


/* ========== COMPARISON BUTTONS ========== */
.comparison-buttons {
    display: flex;
    gap: 10px;
    margin: 0;
    flex-wrap: wrap;
}

/* Caret indicator for third-level buttons that have children (e.g., Visuals) */
.third-btn.has-children {
    position: relative;
    padding-right: 22px; /* space for caret */
}
.third-btn.has-children::after {
    content: '▸';
    position: absolute;
    right: 8px;
    top: 50%;
    transform: translateY(-50%);
    font-size: 12px;
    color: var(--brand-primary);
}
.third-btn.has-children.expanded::after {
    content: '▾';
}

.third-level-buttons {
    display: flex;
    gap: 12px;
    margin: 12px 0 0 0;
    flex-wrap: wrap;
}

.third-level-buttons.hidden {
    display: none;
}

/* ========== FOURTH-LEVEL (VISUALS DETAILS) BUTTONS ========== */
.fourth-level-buttons {
    display: flex;
    gap: 8px;
    margin: 0 0 8px 44px; /* deeper indent to signal hierarchy under Visuals */
    flex-wrap: wrap;
    padding-left: 14px;
    border-left: 2px dashed var(--brand-primary);
}

.fourth-level-buttons.hidden {
    display: none;
}

.fourth-btn {
    background-color: #FAFAFA;
    color: var(--tbl-text-header);
    border: 1px solid #E6E6E6;
    border-radius: 16px;
    padding: 4px 12px 6px 12px;
    min-height: 28px;
    font-weight: 600;
    font-size: 13px;
    line-height: 18px;
    cursor: pointer;
    transition: all 0.2s ease;
}

.fourth-btn:hover {
    background-color: var(--tbl-header-bg);
    color: var(--scroll-thumb);
}

.fourth-btn.active,
.fourth-btn.active:hover,
.fourth-btn.active:focus {
    color: var(--brand-primary-dark) !important;
    border-color: var(--brand-primary-dark);
    background-color: var(--brand-surface-50);
}

/* ========== TEXT DETAILS BUTTON (Loupe) ========== */
.btn-text-details {
    width: 28px;
    height: 28px;
    border: none;
    border-radius: 50%;
    cursor: pointer;
    font-size: 15px;
    font-weight: 700;
    margin: 0 4px;
    transition: all 0.2s ease;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    vertical-align: middle;
    background-color: var(--tbl-header-bg);
    color: #000;
}
.btn-text-details:hover {
    background-color: var(--brand-primary-dark);
    color: #000;
}

/* ========== MODAL (Loupe details) ========== */
.modal {
    position: fixed;
    inset: 0;
    z-index: 2000;
    display: none; /* controlled via .hidden and .show */
    background-color: rgba(0,0,0,0.5);
}
.modal.show {
    display: flex;
    align-items: center;
    justify-content: center;
}
.modal-content {
    background-color: #fff;
    margin: 4.5% auto;
    padding: 24px 24px 16px 24px;
    border-radius: 10px;
    max-height: 90%;
    max-width: 85%;
    width: 85%;
    position: relative;
    box-shadow: 0 4px 20px rgba(0,0,0,0.3);
    animation: modalOpen 0.25s ease-out;
}
.modal-header {
    background: #000;
    color: #fff;
    padding: 8px 10px;
    display: flex;
    justify-content: space-between;
    align-items: center;
    border-radius: 8px 8px 0 0;
	margin-bottom: 10px;
}
.close {
    background: none;
    border: none;
    color: #fff;
    font-size: 22px;
    cursor: pointer;
    padding: 5px;
    border-radius: 50%;
    width: 32px;
    height: 32px;
    display: inline-flex;
    align-items: center;
    justify-content: center;
}
.close:hover { background-color: rgba(255,255,255,0.2); }
.modal-body {
    padding: 14px;
    max-height: 60vh;
    overflow-y: auto;
}
.modal-table { width: 100%; border-collapse: collapse; margin-top: 8px; margin-bottom: 8px; }
.modal-table th {
    background-color: var(--tbl-header-bg);
    color: var(--tbl-text-header);
    font-weight: 700;
    font-size: 16px;
    padding: 14px 16px;
    text-align: left;
    border-bottom: 0.25px solid var(--tbl-border);
}
.modal-table td {
	padding: 8px 10px;
	border-bottom: 1px solid #ddd;
}

@keyframes modalOpen {
    from { opacity: 0; transform: translateY(-16px); }
    to { opacity: 1; transform: translateY(0); }
}

/* ========== PRINT STYLES ========== */
@media print {
    .all-buttons, .main-buttons, .sub-buttons-group, .third-buttons, .third-level-buttons, .fourth-level-buttons {
        display: none !important;
    }
    .btn-text-details, #modal { display: none !important; }
    
    .hidden {
        display: table !important;
    }
    
    body {
        font-size: 12px;
    }
    
    .page-header, .footer {
        background-color: #000000 !important;
        -webkit-print-color-adjust: exact;
        print-color-adjust: exact;
    }
}

/* ========== FOOTER ORANGE ========== */
.footer {
    background-color: #000;
    color: white;
    padding: 15px 30px;
    display: flex;
    justify-content: space-between;
    align-items: center;
    flex-wrap: wrap;
    gap: 12px;
    margin-top: auto;
}

.footer-message {
    font-size: 13px;
    color: var(--tbl-border);
    display: flex;
    flex-wrap: wrap;
    gap: 10px;
    align-items: center;
}

.footer-brand {
    font-size: 14px;
    font-weight: 600;
    color: #fff;
}

.footer-generated {
    display: inline-flex;
    flex-wrap: wrap;
    gap: 6px;
    align-items: baseline;
    color: var(--tbl-header-bg);
    font-size: 13px;
}

.footer-generated time {
    font-variant-numeric: tabular-nums;
    color: #fff;
}

.footer-separator {
    color: var(--tbl-border);
}

/* ========== SYNC AND CAUSE-CONSEQUENCE STYLES ORANGE ========== */
.sync-primary, .cause-added, .cause-modified, .cause-removed {
    background-color: var(--tbl-cell-changed);
    border-left: 4px solid var(--brand-primary-dark);
    cursor: pointer;
    font-weight: bold;
}

.sync-consequence, .consequence-row {
    background-color: #f8f9fa;
    border-left: 2px solid #CCCCCC;
}

.consequence-row.hidden, .slicer-row.hidden {
    display: none;
}

.expand-icon {
    display: inline-block;
    transition: transform 0.3s ease;
    margin-right: 8px;
}

.expanded .expand-icon {
    transform: rotate(90deg);
}

.impact-badge, .impact-badge-consequence {
    background-color: var(--brand-primary);
    color: var(--tbl-text);
    padding: 4px 12px;
    border-radius: 12px;
    font-size: 11px;
    font-weight: bold;
    margin-left: 12px;
}

.impact-badge-neutral {
    background-color: var(--tbl-border);
    color: var(--tbl-text);
    padding: 4px 12px;
    border-radius: 12px;
    font-size: 11px;
    font-weight: bold;
    margin-left: 12px;
}

.truncateElement { 
    color: var(--brand-primary-dark);
    font-weight: bold;
}

</style>
<script>
!function(t,e){"object"==typeof exports&&"undefined"!=typeof module?module.exports=e():"function"==typeof define&&define.amd?define(e):(t="undefined"!=typeof globalThis?globalThis:t||self).i18next=e()}(this,(function(){"use strict";const t={type:"logger",log(t){this.output("log",t)},warn(t){this.output("warn",t)},error(t){this.output("error",t)},output(t,e){console&&console[t]&&console[t].apply(console,e)}};class e{constructor(t){let e=arguments.length>1&&void 0!==arguments[1]?arguments[1]:{};this.init(t,e)}init(e){let s=arguments.length>1&&void 0!==arguments[1]?arguments[1]:{};this.prefix=s.prefix||"i18next:",this.logger=e||t,this.options=s,this.debug=s.debug}log(){for(var t=arguments.length,e=new Array(t),s=0;s<t;s++)e[s]=arguments[s];return this.forward(e,"log","",!0)}warn(){for(var t=arguments.length,e=new Array(t),s=0;s<t;s++)e[s]=arguments[s];return this.forward(e,"warn","",!0)}error(){for(var t=arguments.length,e=new Array(t),s=0;s<t;s++)e[s]=arguments[s];return this.forward(e,"error","")}deprecate(){for(var t=arguments.length,e=new Array(t),s=0;s<t;s++)e[s]=arguments[s];return this.forward(e,"warn","WARNING DEPRECATED: ",!0)}forward(t,e,s,i){return i&&!this.debug?null:("string"==typeof t[0]&&(t[0]=`${s}${this.prefix} ${t[0]}`),this.logger[e](t))}create(t){return new e(this.logger,{prefix:`${this.prefix}:${t}:`,...this.options})}clone(t){return(t=t||this.options).prefix=t.prefix||this.prefix,new e(this.logger,t)}}var s=new e;class i{constructor(){this.observers={}}on(t,e){return t.split(" ").forEach((t=>{this.observers[t]=this.observers[t]||[],this.observers[t].push(e)})),this}off(t,e){this.observers[t]&&(e?this.observers[t]=this.observers[t].filter((t=>t!==e)):delete this.observers[t])}emit(t){for(var e=arguments.length,s=new Array(e>1?e-1:0),i=1;i<e;i++)s[i-1]=arguments[i];if(this.observers[t]){[].concat(this.observers[t]).forEach((t=>{t(...s)}))}if(this.observers["*"]){[].concat(this.observers["*"]).forEach((e=>{e.apply(e,[t,...s])}))}}}function n(){let t,e;const s=new Promise(((s,i)=>{t=s,e=i}));return s.resolve=t,s.reject=e,s}function o(t){return null==t?"":""+t}function r(t,e,s){function i(t){return t&&t.indexOf("###")>-1?t.replace(/###/g,"."):t}function n(){return!t||"string"==typeof t}const o="string"!=typeof e?[].concat(e):e.split(".");for(;o.length>1;){if(n())return{};const e=i(o.shift());!t[e]&&s&&(t[e]=new s),t=Object.prototype.hasOwnProperty.call(t,e)?t[e]:{}}return n()?{}:{obj:t,k:i(o.shift())}}function a(t,e,s){const{obj:i,k:n}=r(t,e,Object);i[n]=s}function l(t,e){const{obj:s,k:i}=r(t,e);if(s)return s[i]}function u(t,e,s){for(const i in e)"__proto__"!==i&&"constructor"!==i&&(i in t?"string"==typeof t[i]||t[i]instanceof String||"string"==typeof e[i]||e[i]instanceof String?s&&(t[i]=e[i]):u(t[i],e[i],s):t[i]=e[i]);return t}function h(t){return t.replace(/[\-\[\]\/\{\}\(\)\*\+\?\.\\\^\$\|]/g,"\\$&")}var c={"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;","/":"&#x2F;"};function p(t){return"string"==typeof t?t.replace(/[&<>"'\/]/g,(t=>c[t])):t}const g=[" ",",","?","!",";"];function d(t,e){let s=arguments.length>2&&void 0!==arguments[2]?arguments[2]:".";if(!t)return;if(t[e])return t[e];const i=e.split(s);let n=t;for(let t=0;t<i.length;++t){if(!n)return;if("string"==typeof n[i[t]]&&t+1<i.length)return;if(void 0===n[i[t]]){let o=2,r=i.slice(t,t+o).join(s),a=n[r];for(;void 0===a&&i.length>t+o;)o++,r=i.slice(t,t+o).join(s),a=n[r];if(void 0===a)return;if(null===a)return null;if(e.endsWith(r)){if("string"==typeof a)return a;if(r&&"string"==typeof a[r])return a[r]}const l=i.slice(t+o).join(s);return l?d(a,l,s):void 0}n=n[i[t]]}return n}function f(t){return t&&t.indexOf("_")>0?t.replace("_","-"):t}class m extends i{constructor(t){let e=arguments.length>1&&void 0!==arguments[1]?arguments[1]:{ns:["translation"],defaultNS:"translation"};super(),this.data=t||{},this.options=e,void 0===this.options.keySeparator&&(this.options.keySeparator="."),void 0===this.options.ignoreJSONStructure&&(this.options.ignoreJSONStructure=!0)}addNamespaces(t){this.options.ns.indexOf(t)<0&&this.options.ns.push(t)}removeNamespaces(t){const e=this.options.ns.indexOf(t);e>-1&&this.options.ns.splice(e,1)}getResource(t,e,s){let i=arguments.length>3&&void 0!==arguments[3]?arguments[3]:{};const n=void 0!==i.keySeparator?i.keySeparator:this.options.keySeparator,o=void 0!==i.ignoreJSONStructure?i.ignoreJSONStructure:this.options.ignoreJSONStructure;let r=[t,e];s&&"string"!=typeof s&&(r=r.concat(s)),s&&"string"==typeof s&&(r=r.concat(n?s.split(n):s)),t.indexOf(".")>-1&&(r=t.split("."));const a=l(this.data,r);return a||!o||"string"!=typeof s?a:d(this.data&&this.data[t]&&this.data[t][e],s,n)}addResource(t,e,s,i){let n=arguments.length>4&&void 0!==arguments[4]?arguments[4]:{silent:!1};const o=void 0!==n.keySeparator?n.keySeparator:this.options.keySeparator;let r=[t,e];s&&(r=r.concat(o?s.split(o):s)),t.indexOf(".")>-1&&(r=t.split("."),i=e,e=r[1]),this.addNamespaces(e),a(this.data,r,i),n.silent||this.emit("added",t,e,s,i)}addResources(t,e,s){let i=arguments.length>3&&void 0!==arguments[3]?arguments[3]:{silent:!1};for(const i in s)"string"!=typeof s[i]&&"[object Array]"!==Object.prototype.toString.apply(s[i])||this.addResource(t,e,i,s[i],{silent:!0});i.silent||this.emit("added",t,e,s)}addResourceBundle(t,e,s,i,n){let o=arguments.length>5&&void 0!==arguments[5]?arguments[5]:{silent:!1},r=[t,e];t.indexOf(".")>-1&&(r=t.split("."),i=s,s=e,e=r[1]),this.addNamespaces(e);let h=l(this.data,r)||{};i?u(h,s,n):h={...h,...s},a(this.data,r,h),o.silent||this.emit("added",t,e,s)}removeResourceBundle(t,e){this.hasResourceBundle(t,e)&&delete this.data[t][e],this.removeNamespaces(e),this.emit("removed",t,e)}hasResourceBundle(t,e){return void 0!==this.getResource(t,e)}getResourceBundle(t,e){return e||(e=this.options.defaultNS),"v1"===this.options.compatibilityAPI?{...this.getResource(t,e)}:this.getResource(t,e)}getDataByLanguage(t){return this.data[t]}hasLanguageSomeTranslations(t){const e=this.getDataByLanguage(t);return!!(e&&Object.keys(e)||[]).find((t=>e[t]&&Object.keys(e[t]).length>0))}toJSON(){return this.data}}var y={processors:{},addPostProcessor(t){this.processors[t.name]=t},handle(t,e,s,i,n){return t.forEach((t=>{this.processors[t]&&(e=this.processors[t].process(e,s,i,n))})),e}};const v={};class b extends i{constructor(t){let e=arguments.length>1&&void 0!==arguments[1]?arguments[1]:{};var i,n;super(),i=t,n=this,["resourceStore","languageUtils","pluralResolver","interpolator","backendConnector","i18nFormat","utils"].forEach((t=>{i[t]&&(n[t]=i[t])})),this.options=e,void 0===this.options.keySeparator&&(this.options.keySeparator="."),this.logger=s.create("translator")}changeLanguage(t){t&&(this.language=t)}exists(t){let e=arguments.length>1&&void 0!==arguments[1]?arguments[1]:{interpolation:{}};if(null==t)return!1;const s=this.resolve(t,e);return s&&void 0!==s.res}extractFromKey(t,e){let s=void 0!==e.nsSeparator?e.nsSeparator:this.options.nsSeparator;void 0===s&&(s=":");const i=void 0!==e.keySeparator?e.keySeparator:this.options.keySeparator;let n=e.ns||this.options.defaultNS||[];const o=s&&t.indexOf(s)>-1,r=!(this.options.userDefinedKeySeparator||e.keySeparator||this.options.userDefinedNsSeparator||e.nsSeparator||function(t,e,s){e=e||"",s=s||"";const i=g.filter((t=>e.indexOf(t)<0&&s.indexOf(t)<0));if(0===i.length)return!0;const n=new RegExp(`(${i.map((t=>"?"===t?"\\?":t)).join("|")})`);let o=!n.test(t);if(!o){const e=t.indexOf(s);e>0&&!n.test(t.substring(0,e))&&(o=!0)}return o}(t,s,i));if(o&&!r){const e=t.match(this.interpolator.nestingRegexp);if(e&&e.length>0)return{key:t,namespaces:n};const o=t.split(s);(s!==i||s===i&&this.options.ns.indexOf(o[0])>-1)&&(n=o.shift()),t=o.join(i)}return"string"==typeof n&&(n=[n]),{key:t,namespaces:n}}translate(t,e,s){if("object"!=typeof e&&this.options.overloadTranslationOptionHandler&&(e=this.options.overloadTranslationOptionHandler(arguments)),"object"==typeof e&&(e={...e}),e||(e={}),null==t)return"";Array.isArray(t)||(t=[String(t)]);const i=void 0!==e.returnDetails?e.returnDetails:this.options.returnDetails,n=void 0!==e.keySeparator?e.keySeparator:this.options.keySeparator,{key:o,namespaces:r}=this.extractFromKey(t[t.length-1],e),a=r[r.length-1],l=e.lng||this.language,u=e.appendNamespaceToCIMode||this.options.appendNamespaceToCIMode;if(l&&"cimode"===l.toLowerCase()){if(u){const t=e.nsSeparator||this.options.nsSeparator;return i?{res:`${a}${t}${o}`,usedKey:o,exactUsedKey:o,usedLng:l,usedNS:a,usedParams:this.getUsedParamsDetails(e)}:`${a}${t}${o}`}return i?{res:o,usedKey:o,exactUsedKey:o,usedLng:l,usedNS:a,usedParams:this.getUsedParamsDetails(e)}:o}const h=this.resolve(t,e);let c=h&&h.res;const p=h&&h.usedKey||o,g=h&&h.exactUsedKey||o,d=Object.prototype.toString.apply(c),f=void 0!==e.joinArrays?e.joinArrays:this.options.joinArrays,m=!this.i18nFormat||this.i18nFormat.handleAsObject;if(m&&c&&("string"!=typeof c&&"boolean"!=typeof c&&"number"!=typeof c)&&["[object Number]","[object Function]","[object RegExp]"].indexOf(d)<0&&("string"!=typeof f||"[object Array]"!==d)){if(!e.returnObjects&&!this.options.returnObjects){this.options.returnedObjectHandler||this.logger.warn("accessing an object - but returnObjects options is not enabled!");const t=this.options.returnedObjectHandler?this.options.returnedObjectHandler(p,c,{...e,ns:r}):`key '${o} (${this.language})' returned an object instead of string.`;return i?(h.res=t,h.usedParams=this.getUsedParamsDetails(e),h):t}if(n){const t="[object Array]"===d,s=t?[]:{},i=t?g:p;for(const t in c)if(Object.prototype.hasOwnProperty.call(c,t)){const o=`${i}${n}${t}`;s[t]=this.translate(o,{...e,joinArrays:!1,ns:r}),s[t]===o&&(s[t]=c[t])}c=s}}else if(m&&"string"==typeof f&&"[object Array]"===d)c=c.join(f),c&&(c=this.extendTranslation(c,t,e,s));else{let i=!1,r=!1;const u=void 0!==e.count&&"string"!=typeof e.count,p=b.hasDefaultValue(e),g=u?this.pluralResolver.getSuffix(l,e.count,e):"",d=e.ordinal&&u?this.pluralResolver.getSuffix(l,e.count,{ordinal:!1}):"",f=e[`defaultValue${g}`]||e[`defaultValue${d}`]||e.defaultValue;!this.isValidLookup(c)&&p&&(i=!0,c=f),this.isValidLookup(c)||(r=!0,c=o);const m=(e.missingKeyNoValueFallbackToKey||this.options.missingKeyNoValueFallbackToKey)&&r?void 0:c,y=p&&f!==c&&this.options.updateMissing;if(r||i||y){if(this.logger.log(y?"updateKey":"missingKey",l,a,o,y?f:c),n){const t=this.resolve(o,{...e,keySeparator:!1});t&&t.res&&this.logger.warn("Seems the loaded translations were in flat JSON format instead of nested. Either set keySeparator: false on init or make sure your translations are published in nested format.")}let t=[];const s=this.languageUtils.getFallbackCodes(this.options.fallbackLng,e.lng||this.language);if("fallback"===this.options.saveMissingTo&&s&&s[0])for(let e=0;e<s.length;e++)t.push(s[e]);else"all"===this.options.saveMissingTo?t=this.languageUtils.toResolveHierarchy(e.lng||this.language):t.push(e.lng||this.language);const i=(t,s,i)=>{const n=p&&i!==c?i:m;this.options.missingKeyHandler?this.options.missingKeyHandler(t,a,s,n,y,e):this.backendConnector&&this.backendConnector.saveMissing&&this.backendConnector.saveMissing(t,a,s,n,y,e),this.emit("missingKey",t,a,s,c)};this.options.saveMissing&&(this.options.saveMissingPlurals&&u?t.forEach((t=>{this.pluralResolver.getSuffixes(t,e).forEach((s=>{i([t],o+s,e[`defaultValue${s}`]||f)}))})):i(t,o,f))}c=this.extendTranslation(c,t,e,h,s),r&&c===o&&this.options.appendNamespaceToMissingKey&&(c=`${a}:${o}`),(r||i)&&this.options.parseMissingKeyHandler&&(c="v1"!==this.options.compatibilityAPI?this.options.parseMissingKeyHandler(this.options.appendNamespaceToMissingKey?`${a}:${o}`:o,i?c:void 0):this.options.parseMissingKeyHandler(c))}return i?(h.res=c,h.usedParams=this.getUsedParamsDetails(e),h):c}extendTranslation(t,e,s,i,n){var o=this;if(this.i18nFormat&&this.i18nFormat.parse)t=this.i18nFormat.parse(t,{...this.options.interpolation.defaultVariables,...s},s.lng||this.language||i.usedLng,i.usedNS,i.usedKey,{resolved:i});else if(!s.skipInterpolation){s.interpolation&&this.interpolator.init({...s,interpolation:{...this.options.interpolation,...s.interpolation}});const r="string"==typeof t&&(s&&s.interpolation&&void 0!==s.interpolation.skipOnVariables?s.interpolation.skipOnVariables:this.options.interpolation.skipOnVariables);let a;if(r){const e=t.match(this.interpolator.nestingRegexp);a=e&&e.length}let l=s.replace&&"string"!=typeof s.replace?s.replace:s;if(this.options.interpolation.defaultVariables&&(l={...this.options.interpolation.defaultVariables,...l}),t=this.interpolator.interpolate(t,l,s.lng||this.language,s),r){const e=t.match(this.interpolator.nestingRegexp);a<(e&&e.length)&&(s.nest=!1)}!s.lng&&"v1"!==this.options.compatibilityAPI&&i&&i.res&&(s.lng=i.usedLng),!1!==s.nest&&(t=this.interpolator.nest(t,(function(){for(var t=arguments.length,i=new Array(t),r=0;r<t;r++)i[r]=arguments[r];return n&&n[0]===i[0]&&!s.context?(o.logger.warn(`It seems you are nesting recursively key: ${i[0]} in key: ${e[0]}`),null):o.translate(...i,e)}),s)),s.interpolation&&this.interpolator.reset()}const r=s.postProcess||this.options.postProcess,a="string"==typeof r?[r]:r;return null!=t&&a&&a.length&&!1!==s.applyPostProcessor&&(t=y.handle(a,t,e,this.options&&this.options.postProcessPassResolved?{i18nResolved:{...i,usedParams:this.getUsedParamsDetails(s)},...s}:s,this)),t}resolve(t){let e,s,i,n,o,r=arguments.length>1&&void 0!==arguments[1]?arguments[1]:{};return"string"==typeof t&&(t=[t]),t.forEach((t=>{if(this.isValidLookup(e))return;const a=this.extractFromKey(t,r),l=a.key;s=l;let u=a.namespaces;this.options.fallbackNS&&(u=u.concat(this.options.fallbackNS));const h=void 0!==r.count&&"string"!=typeof r.count,c=h&&!r.ordinal&&0===r.count&&this.pluralResolver.shouldUseIntlApi(),p=void 0!==r.context&&("string"==typeof r.context||"number"==typeof r.context)&&""!==r.context,g=r.lngs?r.lngs:this.languageUtils.toResolveHierarchy(r.lng||this.language,r.fallbackLng);u.forEach((t=>{this.isValidLookup(e)||(o=t,!v[`${g[0]}-${t}`]&&this.utils&&this.utils.hasLoadedNamespace&&!this.utils.hasLoadedNamespace(o)&&(v[`${g[0]}-${t}`]=!0,this.logger.warn(`key "${s}" for languages "${g.join(", ")}" won't get resolved as namespace "${o}" was not yet loaded`,"This means something IS WRONG in your setup. You access the t function before i18next.init / i18next.loadNamespace / i18next.changeLanguage was done. Wait for the callback or Promise to resolve before accessing it!!!")),g.forEach((s=>{if(this.isValidLookup(e))return;n=s;const o=[l];if(this.i18nFormat&&this.i18nFormat.addLookupKeys)this.i18nFormat.addLookupKeys(o,l,s,t,r);else{let t;h&&(t=this.pluralResolver.getSuffix(s,r.count,r));const e=`${this.options.pluralSeparator}zero`,i=`${this.options.pluralSeparator}ordinal${this.options.pluralSeparator}`;if(h&&(o.push(l+t),r.ordinal&&0===t.indexOf(i)&&o.push(l+t.replace(i,this.options.pluralSeparator)),c&&o.push(l+e)),p){const s=`${l}${this.options.contextSeparator}${r.context}`;o.push(s),h&&(o.push(s+t),r.ordinal&&0===t.indexOf(i)&&o.push(s+t.replace(i,this.options.pluralSeparator)),c&&o.push(s+e))}}let a;for(;a=o.pop();)this.isValidLookup(e)||(i=a,e=this.getResource(s,t,a,r))})))}))})),{res:e,usedKey:s,exactUsedKey:i,usedLng:n,usedNS:o}}isValidLookup(t){return!(void 0===t||!this.options.returnNull&&null===t||!this.options.returnEmptyString&&""===t)}getResource(t,e,s){let i=arguments.length>3&&void 0!==arguments[3]?arguments[3]:{};return this.i18nFormat&&this.i18nFormat.getResource?this.i18nFormat.getResource(t,e,s,i):this.resourceStore.getResource(t,e,s,i)}getUsedParamsDetails(){let t=arguments.length>0&&void 0!==arguments[0]?arguments[0]:{};const e=["defaultValue","ordinal","context","replace","lng","lngs","fallbackLng","ns","keySeparator","nsSeparator","returnObjects","returnDetails","joinArrays","postProcess","interpolation"],s=t.replace&&"string"!=typeof t.replace;let i=s?t.replace:t;if(s&&void 0!==t.count&&(i.count=t.count),this.options.interpolation.defaultVariables&&(i={...this.options.interpolation.defaultVariables,...i}),!s){i={...i};for(const t of e)delete i[t]}return i}static hasDefaultValue(t){const e="defaultValue";for(const s in t)if(Object.prototype.hasOwnProperty.call(t,s)&&e===s.substring(0,12)&&void 0!==t[s])return!0;return!1}}function x(t){return t.charAt(0).toUpperCase()+t.slice(1)}class S{constructor(t){this.options=t,this.supportedLngs=this.options.supportedLngs||!1,this.logger=s.create("languageUtils")}getScriptPartFromCode(t){if(!(t=f(t))||t.indexOf("-")<0)return null;const e=t.split("-");return 2===e.length?null:(e.pop(),"x"===e[e.length-1].toLowerCase()?null:this.formatLanguageCode(e.join("-")))}getLanguagePartFromCode(t){if(!(t=f(t))||t.indexOf("-")<0)return t;const e=t.split("-");return this.formatLanguageCode(e[0])}formatLanguageCode(t){if("string"==typeof t&&t.indexOf("-")>-1){const e=["hans","hant","latn","cyrl","cans","mong","arab"];let s=t.split("-");return this.options.lowerCaseLng?s=s.map((t=>t.toLowerCase())):2===s.length?(s[0]=s[0].toLowerCase(),s[1]=s[1].toUpperCase(),e.indexOf(s[1].toLowerCase())>-1&&(s[1]=x(s[1].toLowerCase()))):3===s.length&&(s[0]=s[0].toLowerCase(),2===s[1].length&&(s[1]=s[1].toUpperCase()),"sgn"!==s[0]&&2===s[2].length&&(s[2]=s[2].toUpperCase()),e.indexOf(s[1].toLowerCase())>-1&&(s[1]=x(s[1].toLowerCase())),e.indexOf(s[2].toLowerCase())>-1&&(s[2]=x(s[2].toLowerCase()))),s.join("-")}return this.options.cleanCode||this.options.lowerCaseLng?t.toLowerCase():t}isSupportedCode(t){return("languageOnly"===this.options.load||this.options.nonExplicitSupportedLngs)&&(t=this.getLanguagePartFromCode(t)),!this.supportedLngs||!this.supportedLngs.length||this.supportedLngs.indexOf(t)>-1}getBestMatchFromCodes(t){if(!t)return null;let e;return t.forEach((t=>{if(e)return;const s=this.formatLanguageCode(t);this.options.supportedLngs&&!this.isSupportedCode(s)||(e=s)})),!e&&this.options.supportedLngs&&t.forEach((t=>{if(e)return;const s=this.getLanguagePartFromCode(t);if(this.isSupportedCode(s))return e=s;e=this.options.supportedLngs.find((t=>t===s?t:t.indexOf("-")<0&&s.indexOf("-")<0?void 0:0===t.indexOf(s)?t:void 0))})),e||(e=this.getFallbackCodes(this.options.fallbackLng)[0]),e}getFallbackCodes(t,e){if(!t)return[];if("function"==typeof t&&(t=t(e)),"string"==typeof t&&(t=[t]),"[object Array]"===Object.prototype.toString.apply(t))return t;if(!e)return t.default||[];let s=t[e];return s||(s=t[this.getScriptPartFromCode(e)]),s||(s=t[this.formatLanguageCode(e)]),s||(s=t[this.getLanguagePartFromCode(e)]),s||(s=t.default),s||[]}toResolveHierarchy(t,e){const s=this.getFallbackCodes(e||this.options.fallbackLng||[],t),i=[],n=t=>{t&&(this.isSupportedCode(t)?i.push(t):this.logger.warn(`rejecting language code not found in supportedLngs: ${t}`))};return"string"==typeof t&&(t.indexOf("-")>-1||t.indexOf("_")>-1)?("languageOnly"!==this.options.load&&n(this.formatLanguageCode(t)),"languageOnly"!==this.options.load&&"currentOnly"!==this.options.load&&n(this.getScriptPartFromCode(t)),"currentOnly"!==this.options.load&&n(this.getLanguagePartFromCode(t))):"string"==typeof t&&n(this.formatLanguageCode(t)),s.forEach((t=>{i.indexOf(t)<0&&n(this.formatLanguageCode(t))})),i}}let k=[{lngs:["ach","ak","am","arn","br","fil","gun","ln","mfe","mg","mi","oc","pt","pt-BR","tg","tl","ti","tr","uz","wa"],nr:[1,2],fc:1},{lngs:["af","an","ast","az","bg","bn","ca","da","de","dev","el","en","eo","es","et","eu","fi","fo","fur","fy","gl","gu","ha","hi","hu","hy","ia","it","kk","kn","ku","lb","mai","ml","mn","mr","nah","nap","nb","ne","nl","nn","no","nso","pa","pap","pms","ps","pt-PT","rm","sco","se","si","so","son","sq","sv","sw","ta","te","tk","ur","yo"],nr:[1,2],fc:2},{lngs:["ay","bo","cgg","fa","ht","id","ja","jbo","ka","km","ko","ky","lo","ms","sah","su","th","tt","ug","vi","wo","zh"],nr:[1],fc:3},{lngs:["be","bs","cnr","dz","hr","ru","sr","uk"],nr:[1,2,5],fc:4},{lngs:["ar"],nr:[0,1,2,3,11,100],fc:5},{lngs:["cs","sk"],nr:[1,2,5],fc:6},{lngs:["csb","pl"],nr:[1,2,5],fc:7},{lngs:["cy"],nr:[1,2,3,8],fc:8},{lngs:["fr"],nr:[1,2],fc:9},{lngs:["ga"],nr:[1,2,3,7,11],fc:10},{lngs:["gd"],nr:[1,2,3,20],fc:11},{lngs:["is"],nr:[1,2],fc:12},{lngs:["jv"],nr:[0,1],fc:13},{lngs:["kw"],nr:[1,2,3,4],fc:14},{lngs:["lt"],nr:[1,2,10],fc:15},{lngs:["lv"],nr:[1,2,0],fc:16},{lngs:["mk"],nr:[1,2],fc:17},{lngs:["mnk"],nr:[0,1,2],fc:18},{lngs:["mt"],nr:[1,2,11,20],fc:19},{lngs:["or"],nr:[2,1],fc:2},{lngs:["ro"],nr:[1,2,20],fc:20},{lngs:["sl"],nr:[5,1,2,3],fc:21},{lngs:["he","iw"],nr:[1,2,20,21],fc:22}],L={1:function(t){return Number(t>1)},2:function(t){return Number(1!=t)},3:function(t){return 0},4:function(t){return Number(t%10==1&&t%100!=11?0:t%10>=2&&t%10<=4&&(t%100<10||t%100>=20)?1:2)},5:function(t){return Number(0==t?0:1==t?1:2==t?2:t%100>=3&&t%100<=10?3:t%100>=11?4:5)},6:function(t){return Number(1==t?0:t>=2&&t<=4?1:2)},7:function(t){return Number(1==t?0:t%10>=2&&t%10<=4&&(t%100<10||t%100>=20)?1:2)},8:function(t){return Number(1==t?0:2==t?1:8!=t&&11!=t?2:3)},9:function(t){return Number(t>=2)},10:function(t){return Number(1==t?0:2==t?1:t<7?2:t<11?3:4)},11:function(t){return Number(1==t||11==t?0:2==t||12==t?1:t>2&&t<20?2:3)},12:function(t){return Number(t%10!=1||t%100==11)},13:function(t){return Number(0!==t)},14:function(t){return Number(1==t?0:2==t?1:3==t?2:3)},15:function(t){return Number(t%10==1&&t%100!=11?0:t%10>=2&&(t%100<10||t%100>=20)?1:2)},16:function(t){return Number(t%10==1&&t%100!=11?0:0!==t?1:2)},17:function(t){return Number(1==t||t%10==1&&t%100!=11?0:1)},18:function(t){return Number(0==t?0:1==t?1:2)},19:function(t){return Number(1==t?0:0==t||t%100>1&&t%100<11?1:t%100>10&&t%100<20?2:3)},20:function(t){return Number(1==t?0:0==t||t%100>0&&t%100<20?1:2)},21:function(t){return Number(t%100==1?1:t%100==2?2:t%100==3||t%100==4?3:0)},22:function(t){return Number(1==t?0:2==t?1:(t<0||t>10)&&t%10==0?2:3)}};const O=["v1","v2","v3"],w=["v4"],N={zero:0,one:1,two:2,few:3,many:4,other:5};class R{constructor(t){let e=arguments.length>1&&void 0!==arguments[1]?arguments[1]:{};this.languageUtils=t,this.options=e,this.logger=s.create("pluralResolver"),this.options.compatibilityJSON&&!w.includes(this.options.compatibilityJSON)||"undefined"!=typeof Intl&&Intl.PluralRules||(this.options.compatibilityJSON="v3",this.logger.error("Your environment seems not to be Intl API compatible, use an Intl.PluralRules polyfill. Will fallback to the compatibilityJSON v3 format handling.")),this.rules=function(){const t={};return k.forEach((e=>{e.lngs.forEach((s=>{t[s]={numbers:e.nr,plurals:L[e.fc]}}))})),t}()}addRule(t,e){this.rules[t]=e}getRule(t){let e=arguments.length>1&&void 0!==arguments[1]?arguments[1]:{};if(this.shouldUseIntlApi())try{return new Intl.PluralRules(f(t),{type:e.ordinal?"ordinal":"cardinal"})}catch(t){return}return this.rules[t]||this.rules[this.languageUtils.getLanguagePartFromCode(t)]}needsPlural(t){let e=arguments.length>1&&void 0!==arguments[1]?arguments[1]:{};const s=this.getRule(t,e);return this.shouldUseIntlApi()?s&&s.resolvedOptions().pluralCategories.length>1:s&&s.numbers.length>1}getPluralFormsOfKey(t,e){let s=arguments.length>2&&void 0!==arguments[2]?arguments[2]:{};return this.getSuffixes(t,s).map((t=>`${e}${t}`))}getSuffixes(t){let e=arguments.length>1&&void 0!==arguments[1]?arguments[1]:{};const s=this.getRule(t,e);return s?this.shouldUseIntlApi()?s.resolvedOptions().pluralCategories.sort(((t,e)=>N[t]-N[e])).map((t=>`${this.options.prepend}${e.ordinal?`ordinal${this.options.prepend}`:""}${t}`)):s.numbers.map((s=>this.getSuffix(t,s,e))):[]}getSuffix(t,e){let s=arguments.length>2&&void 0!==arguments[2]?arguments[2]:{};const i=this.getRule(t,s);return i?this.shouldUseIntlApi()?`${this.options.prepend}${s.ordinal?`ordinal${this.options.prepend}`:""}${i.select(e)}`:this.getSuffixRetroCompatible(i,e):(this.logger.warn(`no plural rule found for: ${t}`),"")}getSuffixRetroCompatible(t,e){const s=t.noAbs?t.plurals(e):t.plurals(Math.abs(e));let i=t.numbers[s];this.options.simplifyPluralSuffix&&2===t.numbers.length&&1===t.numbers[0]&&(2===i?i="plural":1===i&&(i=""));const n=()=>this.options.prepend&&i.toString()?this.options.prepend+i.toString():i.toString();return"v1"===this.options.compatibilityJSON?1===i?"":"number"==typeof i?`_plural_${i.toString()}`:n():"v2"===this.options.compatibilityJSON||this.options.simplifyPluralSuffix&&2===t.numbers.length&&1===t.numbers[0]?n():this.options.prepend&&s.toString()?this.options.prepend+s.toString():s.toString()}shouldUseIntlApi(){return!O.includes(this.options.compatibilityJSON)}}function $(t,e,s){let i=arguments.length>3&&void 0!==arguments[3]?arguments[3]:".",n=!(arguments.length>4&&void 0!==arguments[4])||arguments[4],o=function(t,e,s){const i=l(t,s);return void 0!==i?i:l(e,s)}(t,e,s);return!o&&n&&"string"==typeof s&&(o=d(t,s,i),void 0===o&&(o=d(e,s,i))),o}class P{constructor(){let t=arguments.length>0&&void 0!==arguments[0]?arguments[0]:{};this.logger=s.create("interpolator"),this.options=t,this.format=t.interpolation&&t.interpolation.format||(t=>t),this.init(t)}init(){let t=arguments.length>0&&void 0!==arguments[0]?arguments[0]:{};t.interpolation||(t.interpolation={escapeValue:!0});const e=t.interpolation;this.escape=void 0!==e.escape?e.escape:p,this.escapeValue=void 0===e.escapeValue||e.escapeValue,this.useRawValueToEscape=void 0!==e.useRawValueToEscape&&e.useRawValueToEscape,this.prefix=e.prefix?h(e.prefix):e.prefixEscaped||"{{",this.suffix=e.suffix?h(e.suffix):e.suffixEscaped||"}}",this.formatSeparator=e.formatSeparator?e.formatSeparator:e.formatSeparator||",",this.unescapePrefix=e.unescapeSuffix?"":e.unescapePrefix||"-",this.unescapeSuffix=this.unescapePrefix?"":e.unescapeSuffix||"",this.nestingPrefix=e.nestingPrefix?h(e.nestingPrefix):e.nestingPrefixEscaped||h("$t("),this.nestingSuffix=e.nestingSuffix?h(e.nestingSuffix):e.nestingSuffixEscaped||h(")"),this.nestingOptionsSeparator=e.nestingOptionsSeparator?e.nestingOptionsSeparator:e.nestingOptionsSeparator||",",this.maxReplaces=e.maxReplaces?e.maxReplaces:1e3,this.alwaysFormat=void 0!==e.alwaysFormat&&e.alwaysFormat,this.resetRegExp()}reset(){this.options&&this.init(this.options)}resetRegExp(){const t=`${this.prefix}(.+?)${this.suffix}`;this.regexp=new RegExp(t,"g");const e=`${this.prefix}${this.unescapePrefix}(.+?)${this.unescapeSuffix}${this.suffix}`;this.regexpUnescape=new RegExp(e,"g");const s=`${this.nestingPrefix}(.+?)${this.nestingSuffix}`;this.nestingRegexp=new RegExp(s,"g")}interpolate(t,e,s,i){let n,r,a;const l=this.options&&this.options.interpolation&&this.options.interpolation.defaultVariables||{};function u(t){return t.replace(/\$/g,"$$$$")}const h=t=>{if(t.indexOf(this.formatSeparator)<0){const n=$(e,l,t,this.options.keySeparator,this.options.ignoreJSONStructure);return this.alwaysFormat?this.format(n,void 0,s,{...i,...e,interpolationkey:t}):n}const n=t.split(this.formatSeparator),o=n.shift().trim(),r=n.join(this.formatSeparator).trim();return this.format($(e,l,o,this.options.keySeparator,this.options.ignoreJSONStructure),r,s,{...i,...e,interpolationkey:o})};this.resetRegExp();const c=i&&i.missingInterpolationHandler||this.options.missingInterpolationHandler,p=i&&i.interpolation&&void 0!==i.interpolation.skipOnVariables?i.interpolation.skipOnVariables:this.options.interpolation.skipOnVariables;return[{regex:this.regexpUnescape,safeValue:t=>u(t)},{regex:this.regexp,safeValue:t=>this.escapeValue?u(this.escape(t)):u(t)}].forEach((e=>{for(a=0;n=e.regex.exec(t);){const s=n[1].trim();if(r=h(s),void 0===r)if("function"==typeof c){const e=c(t,n,i);r="string"==typeof e?e:""}else if(i&&Object.prototype.hasOwnProperty.call(i,s))r="";else{if(p){r=n[0];continue}this.logger.warn(`missed to pass in variable ${s} for interpolating ${t}`),r=""}else"string"==typeof r||this.useRawValueToEscape||(r=o(r));const l=e.safeValue(r);if(t=t.replace(n[0],l),p?(e.regex.lastIndex+=r.length,e.regex.lastIndex-=n[0].length):e.regex.lastIndex=0,a++,a>=this.maxReplaces)break}})),t}nest(t,e){let s,i,n,r=arguments.length>2&&void 0!==arguments[2]?arguments[2]:{};function a(t,e){const s=this.nestingOptionsSeparator;if(t.indexOf(s)<0)return t;const i=t.split(new RegExp(`${s}[ ]*{`));let o=`{${i[1]}`;t=i[0],o=this.interpolate(o,n);const r=o.match(/'/g),a=o.match(/"/g);(r&&r.length%2==0&&!a||a.length%2!=0)&&(o=o.replace(/'/g,'"'));try{n=JSON.parse(o),e&&(n={...e,...n})}catch(e){return this.logger.warn(`failed parsing options string in nesting for key ${t}`,e),`${t}${s}${o}`}return delete n.defaultValue,t}for(;s=this.nestingRegexp.exec(t);){let l=[];n={...r},n=n.replace&&"string"!=typeof n.replace?n.replace:n,n.applyPostProcessor=!1,delete n.defaultValue;let u=!1;if(-1!==s[0].indexOf(this.formatSeparator)&&!/{.*}/.test(s[1])){const t=s[1].split(this.formatSeparator).map((t=>t.trim()));s[1]=t.shift(),l=t,u=!0}if(i=e(a.call(this,s[1].trim(),n),n),i&&s[0]===t&&"string"!=typeof i)return i;"string"!=typeof i&&(i=o(i)),i||(this.logger.warn(`missed to resolve ${s[1]} for nesting ${t}`),i=""),u&&(i=l.reduce(((t,e)=>this.format(t,e,r.lng,{...r,interpolationkey:s[1].trim()})),i.trim())),t=t.replace(s[0],i),this.regexp.lastIndex=0}return t}}function C(t){const e={};return function(s,i,n){const o=i+JSON.stringify(n);let r=e[o];return r||(r=t(f(i),n),e[o]=r),r(s)}}class j{constructor(){let t=arguments.length>0&&void 0!==arguments[0]?arguments[0]:{};this.logger=s.create("formatter"),this.options=t,this.formats={number:C(((t,e)=>{const s=new Intl.NumberFormat(t,{...e});return t=>s.format(t)})),currency:C(((t,e)=>{const s=new Intl.NumberFormat(t,{...e,style:"currency"});return t=>s.format(t)})),datetime:C(((t,e)=>{const s=new Intl.DateTimeFormat(t,{...e});return t=>s.format(t)})),relativetime:C(((t,e)=>{const s=new Intl.RelativeTimeFormat(t,{...e});return t=>s.format(t,e.range||"day")})),list:C(((t,e)=>{const s=new Intl.ListFormat(t,{...e});return t=>s.format(t)}))},this.init(t)}init(t){const e=(arguments.length>1&&void 0!==arguments[1]?arguments[1]:{interpolation:{}}).interpolation;this.formatSeparator=e.formatSeparator?e.formatSeparator:e.formatSeparator||","}add(t,e){this.formats[t.toLowerCase().trim()]=e}addCached(t,e){this.formats[t.toLowerCase().trim()]=C(e)}format(t,e,s){let i=arguments.length>3&&void 0!==arguments[3]?arguments[3]:{};return e.split(this.formatSeparator).reduce(((t,e)=>{const{formatName:n,formatOptions:o}=function(t){let e=t.toLowerCase().trim();const s={};if(t.indexOf("(")>-1){const i=t.split("(");e=i[0].toLowerCase().trim();const n=i[1].substring(0,i[1].length-1);"currency"===e&&n.indexOf(":")<0?s.currency||(s.currency=n.trim()):"relativetime"===e&&n.indexOf(":")<0?s.range||(s.range=n.trim()):n.split(";").forEach((t=>{if(!t)return;const[e,...i]=t.split(":"),n=i.join(":").trim().replace(/^'+|'+$/g,"");s[e.trim()]||(s[e.trim()]=n),"false"===n&&(s[e.trim()]=!1),"true"===n&&(s[e.trim()]=!0),isNaN(n)||(s[e.trim()]=parseInt(n,10))}))}return{formatName:e,formatOptions:s}}(e);if(this.formats[n]){let e=t;try{const r=i&&i.formatParams&&i.formatParams[i.interpolationkey]||{},a=r.locale||r.lng||i.locale||i.lng||s;e=this.formats[n](t,a,{...o,...i,...r})}catch(t){this.logger.warn(t)}return e}return this.logger.warn(`there was no format function for ${n}`),t}),t)}}class E extends i{constructor(t,e,i){let n=arguments.length>3&&void 0!==arguments[3]?arguments[3]:{};super(),this.backend=t,this.store=e,this.services=i,this.languageUtils=i.languageUtils,this.options=n,this.logger=s.create("backendConnector"),this.waitingReads=[],this.maxParallelReads=n.maxParallelReads||10,this.readingCalls=0,this.maxRetries=n.maxRetries>=0?n.maxRetries:5,this.retryTimeout=n.retryTimeout>=1?n.retryTimeout:350,this.state={},this.queue=[],this.backend&&this.backend.init&&this.backend.init(i,n.backend,n)}queueLoad(t,e,s,i){const n={},o={},r={},a={};return t.forEach((t=>{let i=!0;e.forEach((e=>{const r=`${t}|${e}`;!s.reload&&this.store.hasResourceBundle(t,e)?this.state[r]=2:this.state[r]<0||(1===this.state[r]?void 0===o[r]&&(o[r]=!0):(this.state[r]=1,i=!1,void 0===o[r]&&(o[r]=!0),void 0===n[r]&&(n[r]=!0),void 0===a[e]&&(a[e]=!0)))})),i||(r[t]=!0)})),(Object.keys(n).length||Object.keys(o).length)&&this.queue.push({pending:o,pendingCount:Object.keys(o).length,loaded:{},errors:[],callback:i}),{toLoad:Object.keys(n),pending:Object.keys(o),toLoadLanguages:Object.keys(r),toLoadNamespaces:Object.keys(a)}}loaded(t,e,s){const i=t.split("|"),n=i[0],o=i[1];e&&this.emit("failedLoading",n,o,e),s&&this.store.addResourceBundle(n,o,s),this.state[t]=e?-1:2;const a={};this.queue.forEach((s=>{!function(t,e,s,i){const{obj:n,k:o}=r(t,e,Object);n[o]=n[o]||[],i&&(n[o]=n[o].concat(s)),i||n[o].push(s)}(s.loaded,[n],o),function(t,e){void 0!==t.pending[e]&&(delete t.pending[e],t.pendingCount--)}(s,t),e&&s.errors.push(e),0!==s.pendingCount||s.done||(Object.keys(s.loaded).forEach((t=>{a[t]||(a[t]={});const e=s.loaded[t];e.length&&e.forEach((e=>{void 0===a[t][e]&&(a[t][e]=!0)}))})),s.done=!0,s.errors.length?s.callback(s.errors):s.callback())})),this.emit("loaded",a),this.queue=this.queue.filter((t=>!t.done))}read(t,e,s){let i=arguments.length>3&&void 0!==arguments[3]?arguments[3]:0,n=arguments.length>4&&void 0!==arguments[4]?arguments[4]:this.retryTimeout,o=arguments.length>5?arguments[5]:void 0;if(!t.length)return o(null,{});if(this.readingCalls>=this.maxParallelReads)return void this.waitingReads.push({lng:t,ns:e,fcName:s,tried:i,wait:n,callback:o});this.readingCalls++;const r=(r,a)=>{if(this.readingCalls--,this.waitingReads.length>0){const t=this.waitingReads.shift();this.read(t.lng,t.ns,t.fcName,t.tried,t.wait,t.callback)}r&&a&&i<this.maxRetries?setTimeout((()=>{this.read.call(this,t,e,s,i+1,2*n,o)}),n):o(r,a)},a=this.backend[s].bind(this.backend);if(2!==a.length)return a(t,e,r);try{const s=a(t,e);s&&"function"==typeof s.then?s.then((t=>r(null,t))).catch(r):r(null,s)}catch(t){r(t)}}prepareLoading(t,e){let s=arguments.length>2&&void 0!==arguments[2]?arguments[2]:{},i=arguments.length>3?arguments[3]:void 0;if(!this.backend)return this.logger.warn("No backend was added via i18next.use. Will not load resources."),i&&i();"string"==typeof t&&(t=this.languageUtils.toResolveHierarchy(t)),"string"==typeof e&&(e=[e]);const n=this.queueLoad(t,e,s,i);if(!n.toLoad.length)return n.pending.length||i(),null;n.toLoad.forEach((t=>{this.loadOne(t)}))}load(t,e,s){this.prepareLoading(t,e,{},s)}reload(t,e,s){this.prepareLoading(t,e,{reload:!0},s)}loadOne(t){let e=arguments.length>1&&void 0!==arguments[1]?arguments[1]:"";const s=t.split("|"),i=s[0],n=s[1];this.read(i,n,"read",void 0,void 0,((s,o)=>{s&&this.logger.warn(`${e}loading namespace ${n} for language ${i} failed`,s),!s&&o&&this.logger.log(`${e}loaded namespace ${n} for language ${i}`,o),this.loaded(t,s,o)}))}saveMissing(t,e,s,i,n){let o=arguments.length>5&&void 0!==arguments[5]?arguments[5]:{},r=arguments.length>6&&void 0!==arguments[6]?arguments[6]:()=>{};if(this.services.utils&&this.services.utils.hasLoadedNamespace&&!this.services.utils.hasLoadedNamespace(e))this.logger.warn(`did not save key "${s}" as the namespace "${e}" was not yet loaded`,"This means something IS WRONG in your setup. You access the t function before i18next.init / i18next.loadNamespace / i18next.changeLanguage was done. Wait for the callback or Promise to resolve before accessing it!!!");else if(null!=s&&""!==s){if(this.backend&&this.backend.create){const a={...o,isUpdate:n},l=this.backend.create.bind(this.backend);if(l.length<6)try{let n;n=5===l.length?l(t,e,s,i,a):l(t,e,s,i),n&&"function"==typeof n.then?n.then((t=>r(null,t))).catch(r):r(null,n)}catch(t){r(t)}else l(t,e,s,i,r,a)}t&&t[0]&&this.store.addResource(t[0],e,s,i)}}}function F(){return{debug:!1,initImmediate:!0,ns:["translation"],defaultNS:["translation"],fallbackLng:["dev"],fallbackNS:!1,supportedLngs:!1,nonExplicitSupportedLngs:!1,load:"all",preload:!1,simplifyPluralSuffix:!0,keySeparator:".",nsSeparator:":",pluralSeparator:"_",contextSeparator:"_",partialBundledLanguages:!1,saveMissing:!1,updateMissing:!1,saveMissingTo:"fallback",saveMissingPlurals:!0,missingKeyHandler:!1,missingInterpolationHandler:!1,postProcess:!1,postProcessPassResolved:!1,returnNull:!1,returnEmptyString:!0,returnObjects:!1,joinArrays:!1,returnedObjectHandler:!1,parseMissingKeyHandler:!1,appendNamespaceToMissingKey:!1,appendNamespaceToCIMode:!1,overloadTranslationOptionHandler:function(t){let e={};if("object"==typeof t[1]&&(e=t[1]),"string"==typeof t[1]&&(e.defaultValue=t[1]),"string"==typeof t[2]&&(e.tDescription=t[2]),"object"==typeof t[2]||"object"==typeof t[3]){const s=t[3]||t[2];Object.keys(s).forEach((t=>{e[t]=s[t]}))}return e},interpolation:{escapeValue:!0,format:t=>t,prefix:"{{",suffix:"}}",formatSeparator:",",unescapePrefix:"-",nestingPrefix:"$t(",nestingSuffix:")",nestingOptionsSeparator:",",maxReplaces:1e3,skipOnVariables:!0}}}function I(t){return"string"==typeof t.ns&&(t.ns=[t.ns]),"string"==typeof t.fallbackLng&&(t.fallbackLng=[t.fallbackLng]),"string"==typeof t.fallbackNS&&(t.fallbackNS=[t.fallbackNS]),t.supportedLngs&&t.supportedLngs.indexOf("cimode")<0&&(t.supportedLngs=t.supportedLngs.concat(["cimode"])),t}function D(){}class V extends i{constructor(){let t=arguments.length>0&&void 0!==arguments[0]?arguments[0]:{},e=arguments.length>1?arguments[1]:void 0;var i;if(super(),this.options=I(t),this.services={},this.logger=s,this.modules={external:[]},i=this,Object.getOwnPropertyNames(Object.getPrototypeOf(i)).forEach((t=>{"function"==typeof i[t]&&(i[t]=i[t].bind(i))})),e&&!this.isInitialized&&!t.isClone){if(!this.options.initImmediate)return this.init(t,e),this;setTimeout((()=>{this.init(t,e)}),0)}}init(){var t=this;let e=arguments.length>0&&void 0!==arguments[0]?arguments[0]:{},i=arguments.length>1?arguments[1]:void 0;"function"==typeof e&&(i=e,e={}),!e.defaultNS&&!1!==e.defaultNS&&e.ns&&("string"==typeof e.ns?e.defaultNS=e.ns:e.ns.indexOf("translation")<0&&(e.defaultNS=e.ns[0]));const o=F();function r(t){return t?"function"==typeof t?new t:t:null}if(this.options={...o,...this.options,...I(e)},"v1"!==this.options.compatibilityAPI&&(this.options.interpolation={...o.interpolation,...this.options.interpolation}),void 0!==e.keySeparator&&(this.options.userDefinedKeySeparator=e.keySeparator),void 0!==e.nsSeparator&&(this.options.userDefinedNsSeparator=e.nsSeparator),!this.options.isClone){let e;this.modules.logger?s.init(r(this.modules.logger),this.options):s.init(null,this.options),this.modules.formatter?e=this.modules.formatter:"undefined"!=typeof Intl&&(e=j);const i=new S(this.options);this.store=new m(this.options.resources,this.options);const n=this.services;n.logger=s,n.resourceStore=this.store,n.languageUtils=i,n.pluralResolver=new R(i,{prepend:this.options.pluralSeparator,compatibilityJSON:this.options.compatibilityJSON,simplifyPluralSuffix:this.options.simplifyPluralSuffix}),!e||this.options.interpolation.format&&this.options.interpolation.format!==o.interpolation.format||(n.formatter=r(e),n.formatter.init(n,this.options),this.options.interpolation.format=n.formatter.format.bind(n.formatter)),n.interpolator=new P(this.options),n.utils={hasLoadedNamespace:this.hasLoadedNamespace.bind(this)},n.backendConnector=new E(r(this.modules.backend),n.resourceStore,n,this.options),n.backendConnector.on("*",(function(e){for(var s=arguments.length,i=new Array(s>1?s-1:0),n=1;n<s;n++)i[n-1]=arguments[n];t.emit(e,...i)})),this.modules.languageDetector&&(n.languageDetector=r(this.modules.languageDetector),n.languageDetector.init&&n.languageDetector.init(n,this.options.detection,this.options)),this.modules.i18nFormat&&(n.i18nFormat=r(this.modules.i18nFormat),n.i18nFormat.init&&n.i18nFormat.init(this)),this.translator=new b(this.services,this.options),this.translator.on("*",(function(e){for(var s=arguments.length,i=new Array(s>1?s-1:0),n=1;n<s;n++)i[n-1]=arguments[n];t.emit(e,...i)})),this.modules.external.forEach((t=>{t.init&&t.init(this)}))}if(this.format=this.options.interpolation.format,i||(i=D),this.options.fallbackLng&&!this.services.languageDetector&&!this.options.lng){const t=this.services.languageUtils.getFallbackCodes(this.options.fallbackLng);t.length>0&&"dev"!==t[0]&&(this.options.lng=t[0])}this.services.languageDetector||this.options.lng||this.logger.warn("init: no languageDetector is used and no lng is defined");["getResource","hasResourceBundle","getResourceBundle","getDataByLanguage"].forEach((e=>{this[e]=function(){return t.store[e](...arguments)}}));["addResource","addResources","addResourceBundle","removeResourceBundle"].forEach((e=>{this[e]=function(){return t.store[e](...arguments),t}}));const a=n(),l=()=>{const t=(t,e)=>{this.isInitialized&&!this.initializedStoreOnce&&this.logger.warn("init: i18next is already initialized. You should call init just once!"),this.isInitialized=!0,this.options.isClone||this.logger.log("initialized",this.options),this.emit("initialized",this.options),a.resolve(e),i(t,e)};if(this.languages&&"v1"!==this.options.compatibilityAPI&&!this.isInitialized)return t(null,this.t.bind(this));this.changeLanguage(this.options.lng,t)};return this.options.resources||!this.options.initImmediate?l():setTimeout(l,0),a}loadResources(t){let e=arguments.length>1&&void 0!==arguments[1]?arguments[1]:D;const s="string"==typeof t?t:this.language;if("function"==typeof t&&(e=t),!this.options.resources||this.options.partialBundledLanguages){if(s&&"cimode"===s.toLowerCase()&&(!this.options.preload||0===this.options.preload.length))return e();const t=[],i=e=>{if(!e)return;if("cimode"===e)return;this.services.languageUtils.toResolveHierarchy(e).forEach((e=>{"cimode"!==e&&t.indexOf(e)<0&&t.push(e)}))};if(s)i(s);else{this.services.languageUtils.getFallbackCodes(this.options.fallbackLng).forEach((t=>i(t)))}this.options.preload&&this.options.preload.forEach((t=>i(t))),this.services.backendConnector.load(t,this.options.ns,(t=>{t||this.resolvedLanguage||!this.language||this.setResolvedLanguage(this.language),e(t)}))}else e(null)}reloadResources(t,e,s){const i=n();return t||(t=this.languages),e||(e=this.options.ns),s||(s=D),this.services.backendConnector.reload(t,e,(t=>{i.resolve(),s(t)})),i}use(t){if(!t)throw new Error("You are passing an undefined module! Please check the object you are passing to i18next.use()");if(!t.type)throw new Error("You are passing a wrong module! Please check the object you are passing to i18next.use()");return"backend"===t.type&&(this.modules.backend=t),("logger"===t.type||t.log&&t.warn&&t.error)&&(this.modules.logger=t),"languageDetector"===t.type&&(this.modules.languageDetector=t),"i18nFormat"===t.type&&(this.modules.i18nFormat=t),"postProcessor"===t.type&&y.addPostProcessor(t),"formatter"===t.type&&(this.modules.formatter=t),"3rdParty"===t.type&&this.modules.external.push(t),this}setResolvedLanguage(t){if(t&&this.languages&&!(["cimode","dev"].indexOf(t)>-1))for(let t=0;t<this.languages.length;t++){const e=this.languages[t];if(!(["cimode","dev"].indexOf(e)>-1)&&this.store.hasLanguageSomeTranslations(e)){this.resolvedLanguage=e;break}}}changeLanguage(t,e){var s=this;this.isLanguageChangingTo=t;const i=n();this.emit("languageChanging",t);const o=t=>{this.language=t,this.languages=this.services.languageUtils.toResolveHierarchy(t),this.resolvedLanguage=void 0,this.setResolvedLanguage(t)},r=(t,n)=>{n?(o(n),this.translator.changeLanguage(n),this.isLanguageChangingTo=void 0,this.emit("languageChanged",n),this.logger.log("languageChanged",n)):this.isLanguageChangingTo=void 0,i.resolve((function(){return s.t(...arguments)})),e&&e(t,(function(){return s.t(...arguments)}))},a=e=>{t||e||!this.services.languageDetector||(e=[]);const s="string"==typeof e?e:this.services.languageUtils.getBestMatchFromCodes(e);s&&(this.language||o(s),this.translator.language||this.translator.changeLanguage(s),this.services.languageDetector&&this.services.languageDetector.cacheUserLanguage&&this.services.languageDetector.cacheUserLanguage(s)),this.loadResources(s,(t=>{r(t,s)}))};return t||!this.services.languageDetector||this.services.languageDetector.async?!t&&this.services.languageDetector&&this.services.languageDetector.async?0===this.services.languageDetector.detect.length?this.services.languageDetector.detect().then(a):this.services.languageDetector.detect(a):a(t):a(this.services.languageDetector.detect()),i}getFixedT(t,e,s){var i=this;const n=function(t,e){let o;if("object"!=typeof e){for(var r=arguments.length,a=new Array(r>2?r-2:0),l=2;l<r;l++)a[l-2]=arguments[l];o=i.options.overloadTranslationOptionHandler([t,e].concat(a))}else o={...e};o.lng=o.lng||n.lng,o.lngs=o.lngs||n.lngs,o.ns=o.ns||n.ns,o.keyPrefix=o.keyPrefix||s||n.keyPrefix;const u=i.options.keySeparator||".";let h;return h=o.keyPrefix&&Array.isArray(t)?t.map((t=>`${o.keyPrefix}${u}${t}`)):o.keyPrefix?`${o.keyPrefix}${u}${t}`:t,i.t(h,o)};return"string"==typeof t?n.lng=t:n.lngs=t,n.ns=e,n.keyPrefix=s,n}t(){return this.translator&&this.translator.translate(...arguments)}exists(){return this.translator&&this.translator.exists(...arguments)}setDefaultNamespace(t){this.options.defaultNS=t}hasLoadedNamespace(t){let e=arguments.length>1&&void 0!==arguments[1]?arguments[1]:{};if(!this.isInitialized)return this.logger.warn("hasLoadedNamespace: i18next was not initialized",this.languages),!1;if(!this.languages||!this.languages.length)return this.logger.warn("hasLoadedNamespace: i18n.languages were undefined or empty",this.languages),!1;const s=e.lng||this.resolvedLanguage||this.languages[0],i=!!this.options&&this.options.fallbackLng,n=this.languages[this.languages.length-1];if("cimode"===s.toLowerCase())return!0;const o=(t,e)=>{const s=this.services.backendConnector.state[`${t}|${e}`];return-1===s||2===s};if(e.precheck){const t=e.precheck(this,o);if(void 0!==t)return t}return!!this.hasResourceBundle(s,t)||(!(this.services.backendConnector.backend&&(!this.options.resources||this.options.partialBundledLanguages))||!(!o(s,t)||i&&!o(n,t)))}loadNamespaces(t,e){const s=n();return this.options.ns?("string"==typeof t&&(t=[t]),t.forEach((t=>{this.options.ns.indexOf(t)<0&&this.options.ns.push(t)})),this.loadResources((t=>{s.resolve(),e&&e(t)})),s):(e&&e(),Promise.resolve())}loadLanguages(t,e){const s=n();"string"==typeof t&&(t=[t]);const i=this.options.preload||[],o=t.filter((t=>i.indexOf(t)<0));return o.length?(this.options.preload=i.concat(o),this.loadResources((t=>{s.resolve(),e&&e(t)})),s):(e&&e(),Promise.resolve())}dir(t){if(t||(t=this.resolvedLanguage||(this.languages&&this.languages.length>0?this.languages[0]:this.language)),!t)return"rtl";const e=this.services&&this.services.languageUtils||new S(F());return["ar","shu","sqr","ssh","xaa","yhd","yud","aao","abh","abv","acm","acq","acw","acx","acy","adf","ads","aeb","aec","afb","ajp","apc","apd","arb","arq","ars","ary","arz","auz","avl","ayh","ayl","ayn","ayp","bbz","pga","he","iw","ps","pbt","pbu","pst","prp","prd","ug","ur","ydd","yds","yih","ji","yi","hbo","men","xmn","fa","jpr","peo","pes","prs","dv","sam","ckb"].indexOf(e.getLanguagePartFromCode(t))>-1||t.toLowerCase().indexOf("-arab")>1?"rtl":"ltr"}static createInstance(){return new V(arguments.length>0&&void 0!==arguments[0]?arguments[0]:{},arguments.length>1?arguments[1]:void 0)}cloneInstance(){let t=arguments.length>0&&void 0!==arguments[0]?arguments[0]:{},e=arguments.length>1&&void 0!==arguments[1]?arguments[1]:D;const s=t.forkResourceStore;s&&delete t.forkResourceStore;const i={...this.options,...t,isClone:!0},n=new V(i);void 0===t.debug&&void 0===t.prefix||(n.logger=n.logger.clone(t));return["store","services","language"].forEach((t=>{n[t]=this[t]})),n.services={...this.services},n.services.utils={hasLoadedNamespace:n.hasLoadedNamespace.bind(n)},s&&(n.store=new m(this.store.data,i),n.services.resourceStore=n.store),n.translator=new b(n.services,i),n.translator.on("*",(function(t){for(var e=arguments.length,s=new Array(e>1?e-1:0),i=1;i<e;i++)s[i-1]=arguments[i];n.emit(t,...s)})),n.init(i,e),n.translator.options=i,n.translator.backendConnector.services.utils={hasLoadedNamespace:n.hasLoadedNamespace.bind(n)},n}toJSON(){return{options:this.options,store:this.store,language:this.language,languages:this.languages,resolvedLanguage:this.resolvedLanguage}}}const T=V.createInstance();return T.createInstance=V.createInstance,T}));
</script>
<script>
// ========== FONCTIONS INTERACTIVES AVANCEES ==========

// Variables globales pour les filtres et la langue
var currentTable = 'table_pages';
var tableParentMap = {};
var currentFocusedTableId = null;
var focusedContainer = null;
var tablesContainer = null;
var currentLang = 'fr';
var i18nInitialized = false;
var currentSearchText = '';
var currentSearchCount = 0;
var currentQualitySearchText = '';
var currentQualitySearchCount = 0;
var currentDiffTypeFilter = 'all';

const translations = {

    fr: {

        title: "Rapport de Comparaison des Rapports Power BI (PBIR)",

        summary: {
            title: "Synthese des differences",

            total: "Total des différences",

            added: "Éléments ajoutés",

            removed: "Éléments supprimés",

            modified: "Éléments modifiés"

        },

        qualitySummary: {
            title: "Resume des checks qualite",

            total: "Total des slicers",

            ok: "Slicers OK",

            alerte: "Alertes",

            erreur: "Erreurs"

        },

        interactiveControls: "Contrôles Interactifs",

        qualityControls: "Contrôles de Vérification Qualité",

        viewSwitch: {

            comparison: "Voir Comparaison",

            quality: "Voir Checks Qualité"

        },

        comparison: {

            title: "Comparaison des Rapports"

        },

        buttons: {

            print: "Imprimer le rapport"

        },

        diffTypes: {

            Ajoute: "Ajouté",

            Supprime: "Supprimé",

            Modifie: "Modifié",

            Identique: "Identique"

        },

        buttonsFilter: {

            all: "Tous",

            added: "+ Ajoutés",

            removed: "- Supprimés",

            modified: "~ Modifiés"

        },

    hierarchy: {

            pages: "Pages & Visuels"

        },

        qualityInfo: {
            button: "Informations",
            title: "Règles du Check Qualité",
            rule_search: "Aucun texte ne doit être présent dans la loupe (zone de recherche).",
            rule_selection: "Aucun filtre ne doit être coché, sauf pour les filtres en mode radio.",
            rule_menu: "Les filtres des groupes filter doivent être vides.",
            rule_period: "Les filtres des groupes period doivent contenir uniquement 'Current Year'.",
            rule_pane: "Le volet de filtre doit être masqué."
        },

        visualsInfo: {
            button: "Informations",
            title: "Aide - Visuels supprimés et recréés",
            message: "Si un même visuel apparaît à la fois comme supprimé et ajouté, cela indique généralement qu'il a été supprimé puis recréé entre les deux versions du rapport. Dans ce cas, il est nécessaire de vérifier que le visuel recréé possède bien les mêmes configurations, interactions et propriétés que l'ancien, ou que cette modification est intentionnelle."
        },

        comparisonButtons: {

            bookmarks: "Signets",

            config: "Configuration",

            themes: "Thèmes",

            pagesVisuels: "Pages & Visuels"

        },

        comparisonThirdButtons: {

            pages: "Pages",

            visuals: "Visuels",

            sync: "Synchronisation",

            interactions: "Interactions",

            fields: "Champs du visuel",

            buttons: "Boutons"

        },

        search: {

            placeholder: "Rechercher dans les différences...",

            clear: "Effacer",

            results: "{{count}} résultat(s) trouvé(s) pour \"{{query}}\""

        },

        qualitySearch: {

            placeholder: "Rechercher dans les vérifications...",

            clear: "Effacer",

            results: "{{count}} vérification(s) trouvée(s) pour \"{{query}}\""

        },

        chart: {

            title: "Graphique des Modifications",

            added: "Ajoutés",

            removed: "Supprimés",

            modified: "Modifiés"

        },

        mainButtons: {

            all: "Toutes les différences",

            pages: "Pages",

            visuals: "Visuels",

            synchronization: "Synchronisation et Visuel",

            visualFields: "Champs du visuel",

            visualButtons: "Boutons",

            bookmarks: "Signets",

            config: "Configuration",

            visualInteractions: "Interactions visuelles",

            themes: "Thèmes",

            /* Main top navigation */
            comparison: "Rapport",
            quality: "Check Qualité",
            powerQuery: "Data",
            desktop: "MDD-OBI Desktop"

        },

        tables: {

            common: {

                headers: {

                    name: "Nom",

                    type: "Type",

                    value: "Valeur / Détails",

                    notes: "Notes"

                }

            },

            table_pages: {

                title: "Modifications des pages du rapport",

                headers: {

                    name: "Nom de la page",

                    property: "Propriété",

                    change: "Type de modification",

                    oldValue: "Ancienne valeur",

                    newValue: "Nouvelle valeur",

                    details: "Détails",

                    description: "Description"

                },

                empty: "Aucune modification détectée dans les pages"

            },

            table_visuals: {

                title: "Modifications des visuels",

                headers: {

                    name: "Type et champs du visuel",

                    page: "Page",

                    property: "Propriété",

                    change: "Type de modification",

                    oldValue: "Ancienne valeur",

                    newValue: "Nouvelle valeur",

                    details: "Détails",

                    description: "Description"

                },

                empty: "Aucune modification détectée dans les visuels"

            },

            table_synchronization: {

                title: "Modifications de synchronisation des slicers",

                headers: {

                    element: "Slicer",

                    page: "Page",

                    syncGroup: "Groupe Sync",

                    action: "Action",

                    difference: "Différence",

                    impact: "Impact / Conséquences",

                    slicer: "Slicer",

                    property: "Propriété",

                    change: "Type de modification",

                    oldValue: "Ancienne valeur",

                    newValue: "Nouvelle valeur",

                    details: "Détails",

                    description: "Description"

                },

                empty: "Aucune modification de synchronisation détectée"

            },

            table_visual_fields: {

                title: "Modifications des champs de visuels",

                headers: {

                    name: "Visuel",

                    page: "Page",

                    change: "Type de modification",

                    oldValue: "Ancienne valeur",

                    newValue: "Nouvelle valeur",

                    details: "Détails",

                    description: "Description"

                },

                empty: "Aucun changement détecté sur les champs de visuels"

            },

            table_buttons: {

                title: "Modifications des boutons (actionButton)",

                headers: {

                    name: "Nom du bouton",

                    page: "Page",

                    property: "Propriété",

                    change: "Type de modification",

                    oldValue: "Ancienne valeur",

                    newValue: "Nouvelle valeur",

                    details: "Détails",

                    description: "Description"

                },

                empty: "Aucun changement détecté sur les boutons"

            },

            table_bookmarks: {

                title: "Modifications des signets",

                headers: {

                    name: "Nom du signet",

                    property: "Propriété",

                    change: "Type de modification",

                    oldValue: "Ancienne valeur",

                    newValue: "Nouvelle valeur",

                    path: "Chemin",

                    description: "Description"

                },

                empty: "Aucune modification détectée dans les signets"

            },

            table_config: {

                title: "Modifications de Configuration",

                headers: {

                    page: "Page",

                    category: "Catégorie",

                    property: "Paramètre",

                    before: "Avant",

                    after: "Après",

                    impact: "Impact"

                },

                empty: "Aucune modification détectée dans la configuration"

            },

            table_visual_interactions: {

                title: "Interactions Visuelles",

                headers: {

                    page: "Page",

                    source: "Visuel Source",

                    target: "Visuel Cible",

                    before: "Interaction Avant",

                    after: "Interaction Après",

                    changeType: "Type de Changement",

                    impact: "Impact"

                },

                empty: "Aucune modification détectée dans les interactions visuelles"

            },

            table_themes: {

                title: "Modifications des thèmes et styles",

                headers: {

                    element: "Élément de thème",

                    property: "Propriété",

                    change: "Type de modification",

                    oldValue: "Ancienne valeur",

                    newValue: "Nouvelle valeur",

                    path: "Chemin",

                    details: "Détails",

                    description: "Description"

                },

                empty: "Aucune modification détectée dans les thèmes"

            },

            table_quality: {

                title: "Vérification Qualité des Slicers",

                headers: {

                    page: "Page",

                    displayName: "Nom du filtre",

                    fieldName: "Champ",

                    hasSelected: "Sélection",

                    hasSearch: "Recherche",

                    searchText: "Texte recherché",

                    isRadio: "Mode radio",

                    status: "Statut",

                    message: "Message"

                },

                empty: "Aucun slicer trouvé pour vérification"

            }

        },

        qualityMessages: {

            ok: "Le filtre est correctement configuré.",

            selection: "Le filtre contient une sélection.",

            selection_allowed: "Le filtre contient une sélection (autorisé).",

            search: "Le filtre contient du texte dans la loupe{{detail}}.",

            selection_search: "Le filtre contient une sélection et du texte dans la loupe{{detail}}.",

            radio_ok: "Le filtre en mode radio a une sélection (autorisé) et pas de recherche.",

            radio_search: "Le filtre en mode radio contient du texte dans la loupe{{detail}} (non autorisé).",

            menu_ok: "Ce filtre du groupe menu_filter est vide comme attendu.",

            menu_selected: "Ce filtre du groupe menu_filter ne doit pas contenir de sélection.",

            menu_period_ok: "Ce filtre du groupe menu_period a uniquement 'Current Year' comme attendu.",

            menu_period_empty: "Ce filtre du groupe menu_period doit contenir uniquement 'Current Year' (actuellement vide).",

            menu_period_multiple: "Ce filtre du groupe menu_period doit contenir uniquement 'Current Year' (actuellement: {{values}}).",

            menu_period_invalid: "Ce filtre du groupe menu_period doit contenir uniquement 'Current Year' (actuellement: {{value}}).",

            filter_pane_hidden: "Le volet de filtre est masqué comme attendu.",

            filter_pane_visible: "Le volet de filtre est visible, il devrait être masqué.",

            filter_pane_error: "Erreur lors de l'analyse du volet de filtre."

        },

        statusLabels: {

            OK: "OK",

            ALERTE: "ALERTE",

            ERREUR: "ERREUR"

        },

        footer: {

            generatedPrefix: "Rapport généré le",

            generatedSuffix: "par PBI_Report_Compare.ps1",

            analysis: "Analyse des modifications entre deux versions de rapport Power BI (format PBIR)",

            newProject: "Nouveau projet :",

            oldProject: "Ancien projet :"

        },

        subButtons: {
        
            tableQuery: "Tables Query",

            parametersOther: "Paramètres",

            relationships: "Relations du modèle",

            measures: "Mesures DAX",

            rls: "RLS",
            
            perspectives: "Perspectives",

            calcGroups: "Groupes de calcul"
        
        },

        thirdButtons: {
        
            tables: "Tables",

            columns: "Colonnes",
        
            steps: "Étapes Power Query",

            parameters: "Paramètres"
        
        },

        Model: {

			Perspectives: {
			
				PerspectiveTables: {
				
					PerspectiveColumns: {
					
						CalculationGroup: {
						
							description: "Desc. Groupe de calcul",
							
							Table: {
							
								name: "Nom Groupe de calcul"
								
							}
							
						},
						
						description: "Desc. Objet de calcul",
						
						expression: "Expression",
						
						name: "Nom Objet de calcul",
						
						PerspectiveTable: {
						
							name: "Nom",
							
							Perspective: {
							
								name: "Perspective"
								
							}
							
						},
						
						state: "État"
						
					}
					
				}
				
			},
			
			relationships: {
			
				crossFilteringBehavior: "Cross Filtering Behavior",
				
				fromCardinality: {
				
					ToString: {
					
						toCardinality: {
						
							ToString: "Cardinalité"
							
						}
						
					}
					
				},
				
				fromTable: {
				
					name: {
					
						fromColumn: {
						
							name: "Table.colonne origine"
							
						}
						
					}
					
				},
				
				toCardinality: "Cardinalité destination",
				
				toColumn: {
				
					name : "Colonne destination"
					
				},
				
				toTable: {
				
					name: {
					
						toColumn: {
						
							name: "Table.colonne destination"
							
						}
						
					}
					
				}
				
			},
			
			roles: {
			
				TablePermissions: {
				
					filterExpression: "Expression du filtre",
					
					name: "Nom",
					
					role: {
					
						name: "Nom du rôle"
						
					},
					
					table: {
					
						name: "Nom de la table"
						
					},
					
				}
				
			},
			
			Tables: {
			
				changedProperties: {
				
					property: "Propriété changée"
					
				},
				
				Columns: {
				
					changedProperties: {
					
						property: "Propriété changée"
						
					},

                    table: {
                    
                        name: "Nom de la table"

                    },
					
					dataType: "Type de données",
					
					expression: "Expression",
					
					formatString: "Format",
					
					isAvailableInMdx: "Est disponible en MDX",
					
					isNameInferred: "Nom déduit de la source",
					
					lineageTag: "lineageTag",
					
					name: "Nom",
					
					sortByColumn: "Colonne de tri",
					
					sourceColumn: "Colonne source",
					
                    summarizeBy: {
                        title: "Synthese des differences",
                        ToString: "Résumé par"
                    },

                    type: "Type",
				},

                columns: {

                    count: "Nombre de colonnes"

                },
				
				ExcludeFromModelRefresh: "Exclue du rafraichissement",
				
				isHidden: "Est cachée",
				
				IsPrivate: "Est privée",
				
				IsRemoved: "Est supprimée",
				
				Measures: {
					
					displayFolder: "Dossier d'affichage",
					
					expression: "Expression",
					
					formatString: "Format",
					
					changedProperties: {
					
						property: "Propriété changée"
						
					},
					
					isHidden: "Est cachée",
					
					name: "Nom"
					
				},

                measures: {
                
                    count: "Nombre de mesures"

                },
				
				name: "Nom",
				
				partitions: {
				
					expression: "Expression",
					
					kind: "Genre",
					
					mode: "Mode",
					
					name: "Nom de la partition",
					
					queryGroup: {
					
						description: "Description",
						
						folder: "Dossier"
						
					},
					
					source: {
					
						expression: "Expression"
						
					},
					
					sourceType: "Type de source",
					
					table: {
					
						name: "Nom de la table"
						
					}
					
				}
				
			}
			
		},

		Specific: {

			parameter: {
			
				ExpressionMeta: "Métadata",
				
				ExpressionValue: "Valeur",
				
				kind: "Genre",
				
				name: "Nom",
				
				queryGroup: {
				
					folder: "Dossier"
					
				}
				
			}
			
		},

		semantic: {

			tables:{

				headers: {

					old_version: "Ancienne version",

					new_version: "Nouvelle version",

					status: "Statut"

				},
				
                empty: {
                
                    semantic_table_tables: "Aucune modification détectée dans les tables",

                    semantic_table_columns: "Aucune modification détectée dans les colonnes",

                    semantic_table_steps: "Aucune modification détectée dans les étapes Power Query",

                    semantic_table_paramValue: "Aucune modification détectée dans les paramètres",

                    semantic_table_relationships: "Aucune modification détectée dans les relations entre les tables",

                    semantic_table_measures: "Aucune modification détectée dans les mesures",

                    semantic_table_roles: "Aucune modification détectée dans les rôles",

                    semantic_table_perspectives: "Aucune modification détectée dans les perspectives",

                    semantic_table_calculGroups: "Aucune modification détectée dans les groupes de calcul"
                
                }

			},
            
            settings: {
                button: "Paramètres",
                title: "⚙️ Paramètres",
                intro: "Configurez le dossier de sortie par défaut.",
                outputPath: {
                    label: "📁 Dossier de sortie :",
                    placeholder: "Ex: C:\\Users\\VotreNom\\Documents\\Rapports_PowerBI",
                    hint: "💡 Dans l'Explorateur Windows, cliquez dans la barre d'adresse du dossier souhaité et copiez le chemin (Ctrl+C), puis cliquez \"Coller\" ci-dessus.",
                    hint2: "Laissez vide pour être sollicité à chaque exécution"
                },
                pasteButton: "� Coller",
                saveButton: "💾 Sauvegarder",
                resetButton: "🔄 Réinitialiser"
            }

		}

    },

    en: {

        title: "Power BI Reports Comparison Report (PBIR)",

        summary: {
            title: "Difference summary",

            total: "Total differences",

            added: "Items added",

            removed: "Items removed",

            modified: "Items modified"

        },

        qualitySummary: {
            title: "Quality check summary",

            total: "Total slicers",

            ok: "Slicers OK",

            alerte: "Alerts",

            erreur: "Errors"

        },

        interactiveControls: "Interactive Controls",

        qualityControls: "Quality Check Controls",

        viewSwitch: {

            comparison: "Show Comparison",

            quality: "Show Quality Checks"

        },

        comparison: {

            title: "Report comparison"

        },

        buttons: {

            print: "Print report"

        },

        diffTypes: {

            Ajoute: "Added",

            Supprime: "Removed",

            Modifie: "Modified",

            Identique: "Identical"

        },

        buttonsFilter: {

            all: "All",

            added: "+ Added",

            removed: "- Removed",

            modified: "~ Modified"

        },

    hierarchy: {

            pages: "Pages & visuals"

        },

        qualityInfo: {
            button: "Information",
            title: "Quality Check Rules",
            rule_search: "No text should be present in a slicer's search box.",
            rule_selection: "No selection should be active, except for radio-mode slicers.",
            rule_menu: "Slicers in the menu_filter group must be empty.",
            rule_period: "Slicers in the menu_period group must have only 'Current Year' selected.",
            rule_pane: "The filter pane must be hidden."
        },

        visualsInfo: {
            button: "Information",
            title: "Help - Removed and Recreated Visuals",
            message: "If the same visual appears as both removed and added, this typically indicates that it was deleted and then recreated between the two report versions. In this case, it is necessary to verify that the recreated visual has the same configurations, interactions, and properties as the old one, or that this change is intentional."
        },

        comparisonButtons: {

            bookmarks: "Bookmarks",

            config: "Configuration",

            themes: "Themes",

            pagesVisuels: "Pages & visuals"

        },

        comparisonThirdButtons: {

            pages: "Pages",

            visuals: "Visuals",

            sync: "Synchronization",

            interactions: "Interactions",

            fields: "Visual fields",

            buttons: "Buttons"

        },

        search: {

            placeholder: "Search differences...",

            clear: "Clear",

            results: "{{count}} result(s) found for \"{{query}}\""

        },

        qualitySearch: {

            placeholder: "Search quality checks...",

            clear: "Clear",

            results: "{{count}} check(s) found for \"{{query}}\""

        },

        chart: {

            title: "Changes overview",

            added: "Added",

            removed: "Removed",

            modified: "Modified"

        },

        mainButtons: {

            all: "All differences",

            pages: "Pages",

            visuals: "Visuals",

            synchronization: "Synchronization and Visual",

            visualFields: "Visual fields",

            visualButtons: "Buttons",

            bookmarks: "Bookmarks",

            config: "Configuration",

            visualInteractions: "Visual Interactions",

            themes: "Themes",

            /* Main top navigation */
            comparison: "Report",
            quality: "Quality checks",
            powerQuery: "Data",
            desktop: "MDD-OBI Desktop"

        },

        tables: {

            common: {

                headers: {

                    name: "Name",

                    type: "Type",

                    value: "Value / Details",

                    notes: "Notes"

                }

            },

            table_pages: {

                title: "Report page changes",

                headers: {

                    name: "Page name",

                    property: "Property",

                    change: "Change type",

                    oldValue: "Old value",

                    newValue: "New value",

                    details: "Details",

                    description: "Description"

                },

                empty: "No page changes detected"

            },

            table_visuals: {

                title: "Visual changes",

                headers: {

                    name: "Visual type and fields",

                    page: "Page",

                    property: "Property",

                    change: "Change type",

                    oldValue: "Old value",

                    newValue: "New value",

                    details: "Details",

                    description: "Description"

                },

                empty: "No visual changes detected"

            },

            table_synchronization: {

                title: "Slicer synchronization changes",

                headers: {

                    element: "Slicer",

                    page: "Page",

                    syncGroup: "Sync Group",

                    action: "Action",

                    difference: "Difference",

                    impact: "Impact / Consequences",

                    slicer: "Slicer",

                    property: "Property",

                    change: "Change type",

                    oldValue: "Old value",

                    newValue: "New value",

                    details: "Details",

                    description: "Description"

                },

                empty: "No synchronization changes detected"

            },

            table_visual_fields: {

                title: "Visual field changes",

                headers: {

                    name: "Visual",

                    page: "Page",

                    change: "Change type",

                    oldValue: "Old value",

                    newValue: "New value",

                    details: "Details",

                    description: "Description"

                },

                empty: "No visual field changes detected"

            },

            table_buttons: {

                title: "Button changes (actionButton)",

                headers: {

                    name: "Button name",

                    page: "Page",

                    property: "Property",

                    change: "Change type",

                    oldValue: "Old value",

                    newValue: "New value",

                    details: "Details",

                    description: "Description"

                },

                empty: "No button changes detected"

            },

            table_bookmarks: {

                title: "Bookmark changes",

                headers: {

                    name: "Bookmark name",

                    property: "Property",

                    change: "Change type",

                    oldValue: "Old value",

                    newValue: "New value",

                    path: "Path",

                    description: "Description"

                },

                empty: "No bookmark changes detected"

            },

            table_config: {

                title: "Configuration Changes",

                headers: {

                    page: "Page",

                    category: "Category",

                    property: "Parameter",

                    before: "Before",

                    after: "After",

                    impact: "Impact"

                },

                empty: "No configuration changes detected"

            },

            table_visual_interactions: {

                title: "Visual Interactions",

                headers: {

                    page: "Page",

                    source: "Source Visual",

                    target: "Target Visual",

                    before: "Interaction Before",

                    after: "Interaction After",

                    changeType: "Change Type",

                    impact: "Impact"

                },

                empty: "No visual interaction changes detected"

            },

            table_themes: {

                title: "Theme and style changes",

                headers: {

                    element: "Theme element",

                    property: "Property",

                    change: "Change type",

                    oldValue: "Old value",

                    newValue: "New value",

                    path: "Path",

                    details: "Details",

                    description: "Description"

                },

                empty: "No theme changes detected"

            },

            table_quality: {

                title: "Quality Check - Slicers",

                headers: {

                    page: "Page",

                    displayName: "Filter name",

                    fieldName: "Field",

                    hasSelected: "Has selection",

                    hasSearch: "Has search",

                    searchText: "Search text",

                    isRadio: "Single select",

                    status: "Status",

                    message: "Message"

                },

                empty: "No slicers found for quality check"

            }

        },

        qualityMessages: {

            ok: "The slicer is properly configured.",

            selection: "The slicer contains a selection.",

            selection_allowed: "The slicer contains a selection (allowed).",

            search: "The slicer contains text in the search box{{detail}}.",

            selection_search: "The slicer contains a selection and text in the search box{{detail}}.",

            radio_ok: "The slicer in single-select mode has a selection (allowed) and no search text.",

            radio_search: "The slicer in single-select mode contains text in the search box{{detail}} (not allowed).",

            menu_ok: "This menu_filter slicer is empty as expected.",

            menu_selected: "This menu_filter slicer must not have an active selection.",

            menu_period_ok: "This menu_period slicer has only 'Current Year' selected as expected.",

            menu_period_empty: "This menu_period slicer must have only 'Current Year' selected (currently empty).",

            menu_period_multiple: "This menu_period slicer must have only 'Current Year' selected (currently: {{values}}).",

            menu_period_invalid: "This menu_period slicer must have only 'Current Year' selected (currently: {{value}}).",

            filter_pane_hidden: "The filter pane is hidden as expected.",

            filter_pane_visible: "The filter pane is visible, but it should be hidden.",

            filter_pane_error: "Error analyzing the filter pane."

        },

        statusLabels: {

            OK: "OK",

            ALERTE: "ALERT",

            ERREUR: "ERROR"

        },

        footer: {

            generatedPrefix: "Report generated on",

            generatedSuffix: "by PBI_Report_Compare.ps1",

            analysis: "Analysis of changes between two Power BI report versions (PBIR format)",

            newProject: "New project:",

            oldProject: "Old project:"

        },

        subButtons: {
        
            tableQuery: "Table Query",

            parametersOther: "Parameters",

            relationships: "Relations tables",

            measures: "DAX Measures",

            rls: "RLS",
            
            perspectives: "Perspectives",

            calcGroups: "Calculation Groups"
        
        },

        thirdButtons: {
        
            tables: "Tables",

            columns: "Columns",
        
            steps: "Power Query Steps",

            parameters: "Parameters"
        
        },

        Model: {

			Perspectives: {
			
				PerspectiveTables: {
				
					PerspectiveColumns: {
					
						CalculationGroup: {
						
							description: "Calculation Group Desc.",
							
							Table: {
							
								name: "Calculation Group Name"
								
							}
							
						},
						
						description: "Calculation Item Desc.",
						
						expression: "Expression",
						
						name: "Calculation Item Name",
						
						PerspectiveTable: {
						
							name: "Name",
							
							Perspective: {
							
								name: "Perspective"
								
							}
							
						},
						
						state: "State"
						
					}
					
				}
				
			},
			
			relationships: {
			
				crossFilteringBehavior: "Cross Filtering Behavior",
				
				fromCardinality: {
				
					ToString: {
					
						toCardinality: {
						
							ToString: "Relationships cardinality"
							
						}
						
					}
					
				},
				
				fromTable: {
				
					name: {
					
						fromColumn: {
						
							name: "From table.column"
							
						}
						
					},
					
				},
				
				name: "Name",
				
				toCardinality: "To Cardinality",
				
				toColumn: {
				
					name : "To Column"
					
				},
				
				toTable: {
				
					name: {
					
						toColumn: {
						
							name: "To table.column"
							
						}
						
					},
					
				}
				
			},
			
			roles: {
			
				TablePermissions: {
				
					filterExpression: "Filter Expression",
					
					name: "Name",
					
					role: {
					
						name: "Role Name"
						
					},
					
					table: {
					
						name: "Table Name"
						
					},
					
				}
				
			},
			
			Tables: {
			
				changedProperties: {
				
					property: "Changed Property"
					
				},
				
				Columns: {
				
					changedProperties: {
					
						property: "Changed Property"
						
					},

                    table: {
                    
                        name: "Table name"

                    },
					
					dataType: "Data Type",
					
					expression: "Expression",
					
					formatString: "Format String",
					
					isAvailableInMdx: "Is Available In MDX",
					
					isNameInferred: "Is Name Inferred",
					
					lineageTag: "lineageTag",
					
					name: "Name",
					
					sortByColumn: "Sort By Column",
					
					sourceColumn: "Source Column",
					
                    summarizeBy: {
                        title: "Difference summary",
                        ToString: "Summarize By"
                    },

                    type: "Type",
				},

                columns: {

                    count: "Columns Count"

                },
				
				ExcludeFromModelRefresh: "Exclude From Model Refresh",
				
				isHidden: "Is Hidde",
				
				IsPrivate: "Is Private",
				
				IsRemoved: "Is Removed",
				
				Measures: {
				
					displayFolder: "Display Folder",
					
					expression: "Expression",
					
					formatString: "Format String",
					
					changedProperties: {
					
						property: "Property Changed"
						
					},
					
					isHidden: "Is Hidden",
					
					name: "Name"
					
				},

                measures: {

                    count: "Measures Count"

                },
				
				name: "Name",
				
				partitions: {
				
					expression: "Expression",
					
					kind: "Kind",
					
					mode: "Mode",
					
					name: "Partition Name",
					
					queryGroup: {
					
						description: "Description",
						
						folder: "Folder"
						
					},
					
					source: {
					
						expression: "Expression"
						
					},
					
					sourceType: "Source Type",
					
					table: {
					
						name: "Table Name"
						
					}
					
				}
				
			}
			
		},

		Specific: {

			parameter: {
			
				ExpressionMeta: "Metadata",
				
				ExpressionValue: "Value",
				
				kind: "Kind",
				
				name: "Name",
				
				queryGroup: {
				
					folder: "Folder"
					
				}
				
			}
			
		},

		semantic: {

			tables:{

				headers: {

					old_version: "Old version",

					new_version: "New version",

					status: "Status"

				},

                empty: {
                
                    semantic_table_tables: "No table changes detected",

                    semantic_table_columns: "No column changes detected",

                    semantic_table_steps: "No Power Query step changes detected",

                    semantic_table_paramValue: "No parameter changes detected",

                    semantic_table_relationships: "No changes in relationships between tables detected",

                    semantic_table_measures: "No measure changes detected",

                    semantic_table_roles: "No role changes detected",

                    semantic_table_perspectives: "No perspective changes detected",

                    semantic_table_calculGroups: "No calculation group changes detected"
                
                }

            },
            
            settings: {
                button: "Settings",
                title: "⚙️ Settings",
                intro: "Configure the default output folder.",
                outputPath: {
                    label: "📁 Output folder:",
                    placeholder: "Ex: C:\\Users\\YourName\\Documents\\PowerBI_Reports",
                    hint: "💡 In Windows Explorer, click the address bar of the desired folder and copy the path (Ctrl+C), then click \"Paste\" above.",
                    hint2: "Leave empty to be prompted each time"
                },
                pasteButton: "� Paste",
                saveButton: "💾 Save",
                resetButton: "🔄 Reset"
            }

		}

    }

};

const diffTypeTranslations = {
    Ajoute: { fr: "Ajoute", en: "Added" },
    Supprime: { fr: "Supprime", en: "Removed" },
    Modifie: { fr: "Modifie", en: "Modified" }
};

const diffTypeTranslationIndex = {};
Object.keys(diffTypeTranslations).forEach(function(key) {
    diffTypeTranslationIndex[normalizeDiffCode(key)] = diffTypeTranslations[key];
});

function ensureI18nInitialized(lang, callback) {
    var targetLang = lang || currentLang || 'fr';
    if (!window.i18next) {
        console.error('i18next library not available');
        if (typeof callback === 'function') {
            callback();
        }
        return;
    }
    if (i18nInitialized) {
        if (i18next.language !== targetLang) {
            i18next.changeLanguage(targetLang, function() {
                if (typeof callback === 'function') {
                    callback();
                }
            });
        } else if (typeof callback === 'function') {
            callback();
        }
        return;
    }
    i18next.init({
        resources: {
            fr: { translation: translations.fr },
            en: { translation: translations.en }
        },
        lng: targetLang,
        fallbackLng: 'fr',
        debug: false,
        interpolation: { escapeValue: false }
    }, function(err) {
        if (err) {
            console.error('i18next init error', err);
        }
        i18nInitialized = true;
        if (typeof callback === 'function') {
            callback();
        }
    });
}
document.addEventListener('DOMContentLoaded', function() {
    prepareTables();
    initializeLanguage();

    setTimeout(function() {
		initTableTruncate('semantic_table_steps');
		initTableTruncate('semantic_table_paramValue');
		initTableTruncate('semantic_table_measures');
		initTableTruncate('semantic_table_roles');
        initTableTruncate('semantic_table_calculGroups');
    }, 100);
});

function resolveTranslation(lang, path, options) {
    if (window.i18next && i18next.isInitialized) {
        var fixedT = i18next.getFixedT(lang);
        var opts = {};
        if (options && typeof options === 'object') {
            Object.keys(options).forEach(function(key) {
                opts[key] = options[key];
            });
        }
        if (typeof opts.defaultValue === 'undefined') {
            opts.defaultValue = null;
        }
        var translated = fixedT(path, opts);
        if (translated !== null && translated !== undefined) {
            return translated;
        }
    }
    var segments = path.split('.');
    var value = translations[lang];
    for (var i = 0; i < segments.length; i++) {
        if (!value) {
            return null;
        }
        value = value[segments[i]];
    }
    if (typeof value === 'string' && options && typeof options === 'object') {
        Object.keys(options).forEach(function(key) {
            var replacement = options[key];
            value = value.split('{{' + key + '}}').join(replacement);
            value = value.split('{' + key + '}').join(replacement);
        });
    }
    return value;
}

function escapeHtml(value) {
    if (value === null || value === undefined) {
        return '';
    }
    return String(value)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

function normalizeDiffCode(code) {
    if (!code) {
        return '';
    }
    var normalized = code.toString().toLowerCase();
    normalized = normalized
        .replace(/[éèêë]/g, 'e')
        .replace(/[àâ]/g, 'a')
        .replace(/[îï]/g, 'i')
        .replace(/[ôö]/g, 'o')
        .replace(/[ûü]/g, 'u');
    return normalized;
}

function getDiffBadgeVariant(code) {
    var normalized = normalizeDiffCode(code);
    if (normalized === 'ajoute' || normalized === 'added') {
        return 'added';
    }
    if (normalized === 'supprime' || normalized === 'supprimee' || normalized === 'removed') {
        return 'removed';
    }
    if (normalized === 'modifie' || normalized === 'modified') {
        return 'modified';
    }
    return 'neutral';
}

function renderDiffBadge(code, label) {
    var variant = getDiffBadgeVariant(code);
    return '<span class="diff-badge diff-badge--' + variant + '">' + escapeHtml(label) + '</span>';
}

function qualityBtnHasFocus() {
    var qualityBtn = document.getElementById('btn-quality');
    if (!qualityBtn) {
        return false;
    }
    if (qualityBtn.classList && qualityBtn.classList.contains('active')) {
        return true;
    }
    return document.activeElement === qualityBtn;
}

// Toggle quality information panel
function toggleQualityInfo() {
    var panel = document.getElementById('quality-info-panel');
    if (!panel) return;
    panel.classList.toggle('hidden');
}

// Toggle visuals information panel
function toggleVisualsInfo() {
    var panel = document.getElementById('visuals-info-panel');
    if (!panel) return;
    panel.classList.toggle('hidden');
}

function applyLanguage(lang) {
    if (!translations[lang]) {
        lang = 'fr';
    }
    if (window.i18next && i18next.isInitialized) {
        lang = i18next.language;
    }
    var documentLang = lang === 'en' ? 'en' : 'fr';
    document.documentElement.lang = documentLang;
    var titleValue = resolveTranslation(lang, 'title');
    if (typeof titleValue === 'string') {
        document.title = titleValue;
    }
    var btnComparison = document.getElementById('btn-comparison');
    var btnQuality = document.getElementById('btn-quality');
    if (btnComparison && btnQuality) {
        btnComparison.textContent = resolveTranslation(lang, 'viewSwitch.comparison') || 'Voir Comparaison';
        btnQuality.textContent = resolveTranslation(lang, 'viewSwitch.quality') || 'Voir Checks Qualite';
    }
    var btnComparison = document.getElementById('btn-comparison');
    var btnQuality = document.getElementById('btn-quality');
    if (btnComparison && btnQuality) {
        var comparisonText = resolveTranslation(lang, 'viewSwitch.comparison') || 'Voir Comparaison';
        var qualityText = resolveTranslation(lang, 'viewSwitch.quality') || 'Voir Checks Qualite';
        if (qualityBtnHasFocus()) {
            btnComparison.textContent = qualityText;
            btnQuality.textContent = comparisonText;
        } else {
            btnComparison.textContent = comparisonText;
            btnQuality.textContent = qualityText;
        }
    }
    document.querySelectorAll('[data-i18n-key]').forEach(function(el) {
        var key = el.getAttribute('data-i18n-key');
        var text = resolveTranslation(lang, key);
        if (typeof text === 'string') {
            var prefix = el.getAttribute('data-i18n-prefix') || '';
            el.textContent = prefix + text;
        }
    });
    document.querySelectorAll('[data-i18n-placeholder]').forEach(function(el) {
        var key = el.getAttribute('data-i18n-placeholder');
        var text = resolveTranslation(lang, key);
        if (typeof text === 'string') {
            el.setAttribute('placeholder', text);
        }
    });
    document.querySelectorAll('[data-i18n-chart]').forEach(function(el) {
        var key = el.getAttribute('data-i18n-chart');
        var text = resolveTranslation(lang, key);
        if (typeof text === 'string') {
            el.textContent = text;
        }
    });
    document.querySelectorAll('td[data-i18n-fr]').forEach(function(el) {
        var attr = lang === 'en' ? 'data-i18n-en' : 'data-i18n-fr';
        var value = el.getAttribute(attr);
        if (value !== null) {
            el.innerHTML = value;
        }
    });
    document.querySelectorAll('[data-i18n-diff-type]').forEach(function(el) {
        var code = el.getAttribute('data-i18n-diff-type');
        if (!code) {
            return;
        }
        var translated = resolveTranslation(lang, 'diffTypes.' + code);
        var label = (typeof translated === 'string' && translated.trim().length > 0) ? translated : '';
        if (!label) {
            var normalizedKey = normalizeDiffCode(code);
            var map = diffTypeTranslations[code] || diffTypeTranslationIndex[normalizedKey];
            if (map) {
                label = map[lang] || map.fr || map.en || '';
            }
        }
        if (!label) {
            label = code;
        }
        if (el.getAttribute('data-diff-badge') === 'true') {
            el.innerHTML = renderDiffBadge(code, label);
        } else {
            el.textContent = label;
        }
    });
    currentLang = lang;
    renderSearchResults();
    renderQualitySearchResults();

    document.querySelectorAll('[data-i18n-message]').forEach(function(el) {
        var key = el.getAttribute('data-i18n-message');
        if (!key) {
            return;
        }
        var detail = el.getAttribute('data-i18n-message-detail') || '';
        var text = resolveTranslation(lang, 'qualityMessages.' + key, { detail: detail });
        if (typeof text === 'string') {
            el.textContent = text;
        }
    });
    document.querySelectorAll('[data-i18n-status]').forEach(function(el) {
        var key = el.getAttribute('data-i18n-status');
        if (!key) {
            return;
        }
        var text = resolveTranslation(lang, 'statusLabels.' + key);
        if (typeof text === 'string') {
            el.textContent = text;
        }
    });
}

function updateLanguageButtons(lang) {
    document.querySelectorAll('.language-switcher button').forEach(function(btn) {
        var isActive = btn.dataset.lang === lang;
        btn.classList.toggle('active', isActive);
        btn.setAttribute('aria-pressed', isActive ? 'true' : 'false');
    });
}

function switchLanguage(lang) {
    if (!translations[lang]) {
        return;
    }
    ensureI18nInitialized(lang, function() {
        var effectiveLang = window.i18next && i18next.isInitialized ? i18next.language : lang;
        currentLang = effectiveLang;
        localStorage.setItem('reportLang', effectiveLang);
        applyLanguage(effectiveLang);
        updateLanguageButtons(effectiveLang);
    });
}

function initializeLanguage() {
    var stored = localStorage.getItem('reportLang');
    if (stored && translations[stored]) {
        currentLang = stored;
    } else if (typeof navigator !== 'undefined' && navigator.language && navigator.language.toLowerCase().startsWith('en')) {
        currentLang = 'en';
    }
    applyLanguage(currentLang);
    updateLanguageButtons(currentLang);
    document.querySelectorAll('.language-switcher button').forEach(function(btn) {
        btn.addEventListener('click', function() {
            switchLanguage(this.dataset.lang);
        });
    });
}

function renderSearchResults() {
    var resultDiv = document.getElementById('searchResults');
    if (!resultDiv) {
        return;
    }
    if (!currentSearchText || currentSearchText.trim() === '') {
        resultDiv.textContent = '';
        return;
    }
    var text = resolveTranslation(currentLang, 'search.results', {
        count: currentSearchCount,
        query: currentSearchText
    });
    if (typeof text === 'string') {
        resultDiv.textContent = text;
    }
}

function renderQualitySearchResults() {
    var resultDiv = document.getElementById('searchResultsQuality');
    if (!resultDiv) {
        var container = document.querySelector('#quality-section .search-container');
        if (!container) {
            return;
        }
        resultDiv = document.createElement('div');
        resultDiv.id = 'searchResultsQuality';
        resultDiv.className = 'search-results';
        container.appendChild(resultDiv);
    }

    if (!currentQualitySearchText || currentQualitySearchText.trim() === '') {
        resultDiv.textContent = '';
        return;
    }

    var text = resolveTranslation(currentLang, 'qualitySearch.results', {
        count: currentQualitySearchCount,
        query: currentQualitySearchText
    });
    if (typeof text === 'string') {
        resultDiv.textContent = text;
    } else {
        resultDiv.textContent = currentQualitySearchCount + ' verification(s) trouvee(s)';
    }
}

function prepareTables() {
    tablesContainer = document.getElementById('all-tables-container');
    focusedContainer = document.getElementById('focused-table-container');
    if (!tablesContainer) {
        return;
    }

    var tables = tablesContainer.querySelectorAll('table');
    var defaultTableId = 'table_pages';
    tables.forEach(function(table) {
        if (!table.id) {
            return;
        }

        if (!tableParentMap[table.id]) {
            var wrapper = document.createElement('div');
            wrapper.className = 'table-wrapper';
            wrapper.dataset.tableId = table.id;
            table.parentNode.insertBefore(wrapper, table);
            wrapper.appendChild(table);
            tableParentMap[table.id] = wrapper;
        }

        if (table.id === defaultTableId) {
            table.classList.remove('hidden');
        } else {
            table.classList.add('hidden');
        }
    });

    if (focusedContainer) {
        focusedContainer.classList.add('hidden');
        focusedContainer.innerHTML = '';
    }

    if (!document.getElementById(defaultTableId)) {
        var firstTable = tablesContainer.querySelector('table');
        if (firstTable) {
            currentTable = firstTable.id;
            firstTable.classList.remove('hidden');
        }
    } else {
        currentTable = defaultTableId;
    }
    initializeLanguage();
}

function restoreFocusedTable() {
    if (!currentFocusedTableId) {
        if (tablesContainer) {
            tablesContainer.style.display = 'block';
        }
        if (focusedContainer) {
            focusedContainer.classList.add('hidden');
            focusedContainer.innerHTML = '';
        }
        return;
    }

    var table = document.getElementById(currentFocusedTableId);
    var wrapper = tableParentMap[currentFocusedTableId];
    if (table && wrapper) {
        table.classList.add('hidden');
        wrapper.appendChild(table);
    }

    currentFocusedTableId = null;

    if (focusedContainer) {
        focusedContainer.classList.add('hidden');
        focusedContainer.innerHTML = '';
    }
    if (tablesContainer) {
        tablesContainer.style.display = 'block';
    }
}

function showTable(tableId) {
    if (!tablesContainer) {
        tablesContainer = document.getElementById('all-tables-container');
    }
    if (!focusedContainer) {
        focusedContainer = document.getElementById('focused-table-container');
    }

    restoreFocusedTable();

    var targetTable = document.getElementById(tableId);
    if (!targetTable) {
        return;
    }

    currentFocusedTableId = tableId;
    targetTable.classList.remove('hidden');

    if (focusedContainer) {
        focusedContainer.classList.remove('hidden');
        focusedContainer.innerHTML = '';
        focusedContainer.appendChild(targetTable);
    }

    if (tablesContainer) {
        tablesContainer.style.display = 'none';
    }

    currentTable = tableId;
    updateActiveButton(tableId);
    applyFilters();
}

function updateActiveButton(activeTableId) {
    var buttons = document.querySelectorAll('.button-set button[data-target]');
    buttons.forEach(function(btn) {
        btn.classList.remove('active');
    });

    if (!activeTableId) {
        return;
    }

    var guard = 0;
    var visited = {};
    var selector = '.button-set button[data-target="' + activeTableId + '"]';
    var current = document.querySelector(selector);

    while (current && guard < 10) {
        current.classList.add('active');
        var parentTarget = current.getAttribute('data-parent');
        if (!parentTarget || visited[parentTarget]) {
            break;
        }
        visited[parentTarget] = true;
        current = document.querySelector('.button-set button[data-target="' + parentTarget + '"]');
        guard++;
    }
}

function getModificationColumnIndex(tableId) {
    switch (tableId) {
        case 'table_pages':
            return 2;
        case 'table_visuals':
        case 'table_buttons':
            return 3;
        case 'table_visual_fields':
        case 'table_bookmarks':
        case 'table_config':
        case 'table_themes':
            return 2;
        case 'table_synchronization':
            return 3;
        default:
            return 5;
    }
}

function filterByType(type) {
    currentDiffTypeFilter = type || 'all';

    var filterButtons = document.querySelectorAll('.filter-buttons button');
    filterButtons.forEach(function(btn) {
        btn.classList.remove('active');
    });

    var activeBtn = document.querySelector('.filter-buttons button[onclick*="' + type + '"]');
    if (activeBtn) {
        activeBtn.classList.add('active');
    }

    searchInTable();
}

function searchInTable() {
    var searchInput = document.getElementById('searchBox');
    if (!searchInput) {
        return;
    }

    var searchText = (searchInput.value || '').toLowerCase();
    var rows = document.querySelectorAll('#' + currentTable + ' tr');
    var visibleCount = 0;
    
    rows.forEach(function(row, index) {
        // Ignorer les lignes d'en-tête (les 2 premières lignes)
        if (index < 2) return;
        
        var shouldShow = false;
        
        // Si pas de texte de recherche, afficher toutes les lignes
        if (searchText.trim() === '') {
            shouldShow = true;
        } else {
            // Chercher dans toutes les cellules
            for (var i = 0; i < row.cells.length; i++) {
                var cellText = row.cells[i].textContent.toLowerCase();
                if (cellText.indexOf(searchText) !== -1) {
                    shouldShow = true;
                    break;
                }
            }
        }
        
        if (shouldShow && currentDiffTypeFilter !== 'all') {
            var modifColumnIndex = getModificationColumnIndex(currentTable);
            var modifCell = row.cells[modifColumnIndex];
            var modifTypeValue = '';
            if (modifCell) {
                modifTypeValue = modifCell.getAttribute('data-i18n-diff-type') || modifCell.textContent.trim();
            }
            if (modifTypeValue !== currentDiffTypeFilter) {
                shouldShow = false;
            }
        }

        if (shouldShow) {
            row.style.display = '';
            visibleCount++;
        } else {
            row.style.display = 'none';
        }
    });
    
    // Mettre a jour le compteur de resultats
    updateSearchResults(visibleCount, searchText);
}

function updateSearchResults(count, searchText) {
    var resultDiv = document.getElementById('searchResults');
    if (!resultDiv) {
        resultDiv = document.createElement('div');
        resultDiv.id = 'searchResults';
        resultDiv.className = 'search-results';
        document.querySelector('.search-container').appendChild(resultDiv);
    }

    currentSearchText = searchText || '';
    currentSearchCount = count || 0;

    if (!currentSearchText.trim()) {
        resultDiv.textContent = '';
        return;
    }

    renderSearchResults();
}

function clearSearch() {
    document.getElementById('searchBox').value = '';
    searchInTable();
}

function applyFilters() {
    if (document.getElementById('searchBox')) {
        searchInTable();
    }
}


function printReport() {
    window.print();
}

function toggleDetails(elementId) {
    var element = document.getElementById(elementId);
    if (element) {
        element.style.display = element.style.display === 'none' ? 'block' : 'none';
    }
}

// ========== ORANGE DESIGN SYSTEM - NAVIGATION FUNCTIONS ==========

// Main button navigation - 4 sections (Rapport, Check Qualité, MDD-Power Query, MDD-OBI Desktop)
function showMainSection(target) {
    // Hide all main sections
    document.querySelectorAll('.comparison-section, .quality-section, .power-query-section, .desktop-section').forEach(section => {
        section.classList.add('section-hidden');
    });
    
    // Show target section
    var targetSection = document.getElementById(target + '-section');
    if (targetSection) {
        targetSection.classList.remove('section-hidden');
    }
    
    // Update main button active states
    document.querySelectorAll('.main-btn').forEach(btn => {
        btn.classList.remove('active');
    });
    var activeBtn = document.querySelector('.main-btn[data-target="' + target + '"]');
    if (activeBtn) {
        activeBtn.classList.add('active');
    }
    
    // Reset all sub-buttons and third-level buttons (FIX BUG: hide when switching sections)
    document.querySelectorAll('.sub-btn, .third-btn, .fourth-btn').forEach(btn => {
        btn.classList.remove('active');
    });
    
    // Hide all third-level and visuals-details button groups (FIX BUG: prevent buttons from staying visible)
    document.querySelectorAll('.third-buttons').forEach(group => { group.style.display = 'none'; });
    document.querySelectorAll('.third-level-buttons').forEach(group => { group.classList.add('hidden'); });
    document.querySelectorAll('.fourth-level-buttons').forEach(group => { group.classList.add('hidden'); });

    if (target === 'comparison') {
        var pagesVisuelsButtons = document.getElementById('pages-visuels-buttons');
        if (pagesVisuelsButtons) {
            pagesVisuelsButtons.classList.remove('hidden');
        }
        var pagesVisuelsBtn = document.querySelector('.comparison-buttons .sub-btn[data-target="pages_visuels"]');
        if (pagesVisuelsBtn) {
            pagesVisuelsBtn.classList.add('active');
            pagesVisuelsBtn.classList.add('expanded');
        }
        var pagesThirdBtn = document.querySelector('#pages-visuels-buttons .third-btn[data-target="pages"]');
        if (pagesThirdBtn) {
            pagesThirdBtn.classList.add('active');
            pagesThirdBtn.style.setProperty('color', 'var(--brand-primary-dark)', 'important');
        }
        filterComparisonTable('pages');
    }
    
    // Show/hide appropriate sub-button groups
    document.querySelectorAll('.sub-buttons-group').forEach(group => {
        group.style.display = 'none';
    });
    
    if (target === 'power-query') {
        document.getElementById('power-query-subs').style.display = 'flex';
    } else if (target === 'desktop') {
        document.getElementById('desktop-subs').style.display = 'flex';
    }
}

// Sub-button navigation - Show third-level buttons and content
function showSubSection(target) {
    // Update sub-button active states
    document.querySelectorAll('.sub-btn').forEach(btn => {
        btn.classList.remove('active');
    });
    var activeBtn = document.querySelector('.sub-btn[data-target="' + target + '"]');
    if (activeBtn) {
        activeBtn.classList.add('active');
    }
    
    // Reset caret expansion on all hierarchical sub buttons
    document.querySelectorAll('.sub-btn.has-children').forEach(btn => {
        btn.classList.remove('expanded');
    });

    // Show/hide third-level button groups
    document.querySelectorAll('.third-buttons').forEach(group => { group.style.display = 'none'; });
    // Always hide visuals fourth-level when switching sub-section; it will be shown on demand
    document.querySelectorAll('.fourth-level-buttons').forEach(group => { group.classList.add('hidden'); });
    
    var thirdGroup = document.getElementById(target + '-buttons');
    if (thirdGroup) {
        thirdGroup.style.display = 'flex';
        // If the active sub button has children, mark it expanded so its caret shows ▾
        if (activeBtn && activeBtn.classList.contains('has-children')) {
            activeBtn.classList.add('expanded');
        }
    }
}

// Third-level button navigation - Show specific semantic tables
function showSemanticTable(tableId) {
    // Hide all semantic tables
    document.querySelectorAll('div.table-container').forEach(table => {
        table.classList.add('hidden');
    });
    
    // Show target table
    var targetTable = document.getElementById(tableId);
    if (targetTable) {
        targetTable.classList.remove('hidden');
    }
    
    // Update third-level button active states
    document.querySelectorAll('.third-btn').forEach(btn => {
        btn.classList.remove('active');
    });
    var activeBtn = document.querySelector('.third-btn[data-table="' + tableId + '"]');
    if (activeBtn) {
        activeBtn.classList.add('active');
        // Force orange active color inline to avoid any cascade conflicts
        var container = activeBtn.closest('.third-buttons');
        if (container) {
            container.querySelectorAll('.third-btn').forEach(function(b){ b.style.removeProperty('color'); });
        } else {
            document.querySelectorAll('.third-btn').forEach(function(b){ b.style.removeProperty('color'); });
        }
    activeBtn.style.setProperty('color', 'var(--brand-primary-dark)', 'important');
    } else {
        // Clear any previous inline colors if no active button found
        document.querySelectorAll('.third-btn').forEach(function(b){ b.style.removeProperty('color'); });
    }
}

// ========== MODAL (Loupe) FUNCTIONS ==========
function resetModalTables() {
    var containers = document.querySelectorAll('#modal .modal-table-container');
    containers.forEach(function(c){ c.classList.add('hidden'); });
}

function openModal(tableId) {
    var modal = document.getElementById('modal');
    if (!modal) return;
    var body = document.body;
    resetModalTables();
    if (tableId) {
        var target = document.getElementById(tableId);
        if (target) target.classList.remove('hidden');
    }
    modal.classList.add('show');
    modal.classList.remove('hidden');
    if (body) { body.style.overflow = 'hidden'; }
}

function closeModalFunction() {
    var modal = document.getElementById('modal');
    if (!modal) return;
    modal.classList.remove('show');
    modal.classList.add('hidden');
    document.body && (document.body.style.overflow = '');
}

// ========== SETTINGS MODAL FUNCTIONS ==========
function openSettingsModal() {
    var modal = document.getElementById('settingsModal');
    if (!modal) return;
    modal.classList.remove('hidden');
    modal.classList.add('show');
    document.body && (document.body.style.overflow = 'hidden');
}

function closeSettingsModal() {
    var modal = document.getElementById('settingsModal');
    if (!modal) return;
    modal.classList.remove('show');
    modal.classList.add('hidden');
    document.body && (document.body.style.overflow = '');
}

async function pasteFromClipboard() {
    var currentLang = localStorage.getItem('reportLanguage') || 'fr';
    
    if ('clipboard' in navigator && 'readText' in navigator.clipboard) {
        try {
            const text = await navigator.clipboard.readText();
            if (text && text.trim()) {
                document.getElementById('outputPathInput').value = text.trim();
                var successMsg = currentLang === 'fr'
                    ? '✓ Chemin collé avec succès !'
                    : '✓ Path pasted successfully!';
                alert(successMsg);
            } else {
                var emptyMsg = currentLang === 'fr'
                    ? '⚠️ Le presse-papiers est vide.\n\nCopiez d\'abord le chemin du dossier depuis l\'Explorateur Windows.'
                    : '⚠️ Clipboard is empty.\n\nFirst copy the folder path from Windows Explorer.';
                alert(emptyMsg);
            }
        } catch (err) {
            console.error('Clipboard error:', err);
            var errorMsg = currentLang === 'fr'
                ? '❌ Impossible d\'accéder au presse-papiers.\n\nVeuillez coller manuellement le chemin avec Ctrl+V dans le champ ci-dessus.'
                : '❌ Cannot access clipboard.\n\nPlease paste the path manually with Ctrl+V in the field above.';
            alert(errorMsg);
            document.getElementById('outputPathInput').focus();
        }
    } else {
        // Fallback for browsers without Clipboard API
        var fallbackMsg = currentLang === 'fr'
            ? 'Votre navigateur ne supporte pas le collage automatique.\n\nVeuillez coller manuellement le chemin avec Ctrl+V dans le champ ci-dessus.'
            : 'Your browser does not support automatic pasting.\n\nPlease paste the path manually with Ctrl+V in the field above.';
        alert(fallbackMsg);
        document.getElementById('outputPathInput').focus();
    }
}

async function saveConfig() {
    var outputPath = document.getElementById('outputPathInput').value.trim();
    
    if (!outputPath) {
        var currentLang = localStorage.getItem('reportLanguage') || 'fr';
        var msg = currentLang === 'fr' 
            ? 'Veuillez saisir un chemin de dossier.' 
            : 'Please enter a folder path.';
        alert(msg);
        return;
    }
    
    // Create config.json object
    var configData = {
        version: "1.0",
        defaultOutputPath: outputPath,
        lastModified: new Date().toISOString(),
        autoOpenReport: true
    };
    
    // Save to localStorage for UI persistence
    localStorage.setItem('pbi-report-config', JSON.stringify(configData));
    
    var currentLang = localStorage.getItem('reportLanguage') || 'fr';
    var configJsonPath = window.INJECTED_CONFIG_PATH || '';
    
    // Try File System Access API (Chrome/Edge) for direct file writing
    if (window.showSaveFilePicker) {
        try {
            // Check if we already have a stored file handle
            var storedHandle = await getStoredFileHandle(configJsonPath);
            
            if (!storedHandle) {
                // First time: use showSaveFilePicker with suggested name and location
                // The browser will open directly in the right folder if possible
                var saveOptions = {
                    suggestedName: 'config.json',
                    types: [{
                        description: 'JSON Configuration',
                        accept: { 'application/json': ['.json'] }
                    }]
                };
                
                // Try to suggest the start directory (works in some browsers)
                if (configJsonPath) {
                    var folderPath = configJsonPath.substring(0, configJsonPath.lastIndexOf('\\'));
                    try {
                        // Note: startIn with path is not widely supported, but we try
                        saveOptions.startIn = folderPath;
                    } catch (e) {
                        // Ignore if not supported
                    }
                }
                
                // Show save dialog - user just needs to click "Save"!
                storedHandle = await window.showSaveFilePicker(saveOptions);
                await storeFileHandle(configJsonPath, storedHandle);
            }
            
            // Write directly to the file
            var writable = await storedHandle.createWritable();
            await writable.write(JSON.stringify(configData, null, 4));
            await writable.close();
            
            var successMsg = currentLang === 'fr'
                ? '✓ Configuration sauvegardée !\n\n📝 Le fichier config.json a été mis à jour.\n\n✨ Les prochaines sauvegardes seront instantanées !'
                : '✓ Configuration saved!\n\n📝 The config.json file has been updated.\n\n✨ Future saves will be instant!';
            alert(successMsg);
            closeSettingsModal();
            return;
            
        } catch (error) {
            console.error('File System Access API error:', error);
            // User cancelled or API not available, fall through to download method
            if (error.name === 'AbortError') {
                // User cancelled the file picker
                return;
            }
        }
    }
    
    // Fallback: Download method (Safari, Firefox, or if user denied permission)
    var jsonContent = JSON.stringify(configData, null, 4);
    var blob = new Blob([jsonContent], { type: 'application/json' });
    var url = URL.createObjectURL(blob);
    var link = document.createElement('a');
    link.href = url;
    link.download = 'config.json';
    link.style.display = 'none';
    document.body.appendChild(link);
    link.click();
    
    setTimeout(function() {
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
    }, 100);
    
    var fallbackMsg = currentLang === 'fr'
        ? '✓ Configuration sauvegardée !\n\n📥 Le fichier config.json a été téléchargé.\n\n📁 Déplacez-le dans Script_Merged/ (remplacez l\'ancien)\n\n💡 Utilisez Chrome ou Edge pour l\'écriture automatique !'
        : '✓ Configuration saved!\n\n📥 The config.json file has been downloaded.\n\n📁 Move it to Script_Merged/ (replace the old one)\n\n💡 Use Chrome or Edge for automatic writing!';
    alert(fallbackMsg);
    closeSettingsModal();
}

// Helper functions for File System Access API handle persistence
async function getStoredFileHandle(configPath) {
    try {
        var db = await openIndexedDB();
        var key = 'configFileHandle_' + configPath;
        return await new Promise((resolve, reject) => {
            var transaction = db.transaction(['fileHandles'], 'readonly');
            var store = transaction.objectStore('fileHandles');
            var request = store.get(key);
            request.onsuccess = () => resolve(request.result);
            request.onerror = () => reject(request.error);
        });
    } catch (error) {
        console.error('Error getting stored handle:', error);
        return null;
    }
}

async function storeFileHandle(configPath, fileHandle) {
    try {
        var db = await openIndexedDB();
        var key = 'configFileHandle_' + configPath;
        return await new Promise((resolve, reject) => {
            var transaction = db.transaction(['fileHandles'], 'readwrite');
            var store = transaction.objectStore('fileHandles');
            var request = store.put(fileHandle, key);
            request.onsuccess = () => resolve();
            request.onerror = () => reject(request.error);
        });
    } catch (error) {
        console.error('Error storing handle:', error);
    }
}

async function openIndexedDB() {
    return new Promise((resolve, reject) => {
        var request = indexedDB.open('PBIReportConfig', 1);
        request.onupgradeneeded = (event) => {
            var db = event.target.result;
            if (!db.objectStoreNames.contains('fileHandles')) {
                db.createObjectStore('fileHandles');
            }
        };
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
    });
}

async function resetConfig() {
    document.getElementById('outputPathInput').value = '';
    localStorage.removeItem('pbi-report-config');
    
    // Also remove stored file handle for this config path
    try {
        var configJsonPath = window.INJECTED_CONFIG_PATH || '';
        if (configJsonPath) {
            var db = await openIndexedDB();
            var key = 'configFileHandle_' + configJsonPath;
            var transaction = db.transaction(['fileHandles'], 'readwrite');
            var store = transaction.objectStore('fileHandles');
            store.delete(key);
        }
    } catch (error) {
        console.error('Error removing file handle:', error);
    }
    
    var currentLang = localStorage.getItem('reportLanguage') || 'fr';
    var confirmMsg = currentLang === 'fr' 
        ? '🔄 Configuration réinitialisée !\n\nLe chemin par défaut et l\'accès au fichier ont été supprimés.\nVous serez sollicité à la prochaine sauvegarde.'
        : '🔄 Configuration reset!\n\nDefault path and file access removed.\nYou will be prompted on next save.';
    alert(confirmMsg);
}

// Close settings modal on backdrop click
document.addEventListener('click', function(event) {
    var modal = document.getElementById('settingsModal');
    if (modal && event.target === modal) {
        closeSettingsModal();
    }
});

// Initialize Orange navigation on page load
document.addEventListener('DOMContentLoaded', function() {
    // Attach click handlers to main buttons
    document.querySelectorAll('.main-btn').forEach(btn => {
        btn.addEventListener('click', function() {
            var target = this.getAttribute('data-target');
            showMainSection(target);
        });
    });
    
    // Attach click handlers to sub buttons
    document.querySelectorAll('.sub-btn').forEach(btn => {
        btn.addEventListener('click', function() {
            var target = this.getAttribute('data-target');
            showSubSection(target);
        });
    });
    
    // Attach click handlers to third-level buttons for semantic sections ONLY
    // Important: limit to buttons that declare a data-table attribute to avoid
    // interfering with Rapport third-level filters (which use data-target)
    document.querySelectorAll('.third-btn[data-table]').forEach(btn => {
        btn.addEventListener('click', function() {
            var tableId = this.getAttribute('data-table');
            if (!tableId) return;
            showSemanticTable(tableId);
        });
    });
    
    // Start with quality check section (Check Qualité button active by default)
    showMainSection('quality');
});

// Attach backdrop click and Escape key to close modal
document.addEventListener('DOMContentLoaded', function(){
    var modal = document.getElementById('modal');
    if (!modal) return;
    modal.addEventListener('click', function(e){
        if (e.target === modal) { closeModalFunction(); }
    });
    document.addEventListener('keydown', function(e){
        if (e.key === 'Escape') { closeModalFunction(); }
    });
});

// ========== COMPARISON TABLE FILTER FUNCTIONS ==========

// Toggle Visuels second-level buttons (Interactions, Champs, Boutons)
function toggleVisualsButtons() {
    var visualsDetailsButtons = document.getElementById('visuals-details-buttons');
    var button = document.querySelector('.sub-btn[onclick*="toggleVisualsButtons"]');
    
    if (visualsDetailsButtons.classList.contains('hidden')) {
        // Show buttons
        visualsDetailsButtons.classList.remove('hidden');
        // Mark the toggler as active/expanded
        if (button) { button.classList.add('active'); button.classList.add('expanded'); }
        // CRITICAL: Also show the visuals table when opening sub-buttons
        filterComparisonTable('visuals');
    } else {
        // Hide buttons
        visualsDetailsButtons.classList.add('hidden');
        // Clear active states from sub-buttons
        document.querySelectorAll('.third-level-buttons .third-btn').forEach(function(b){ b.classList.remove('active'); });
        // Remove active/expanded state from the toggler
        if (button) { button.classList.remove('active'); button.classList.remove('expanded'); }
    }
}

// Filter comparison tables by type
function filterComparisonTable(filter) {
    var visualsDetailsButtons = document.getElementById('visuals-details-buttons');
    var isSecondLevelFilter = ['interactions', 'fields', 'buttons'].includes(filter);
    
    // Show or hide the visuals details bar depending on selection
    if (isSecondLevelFilter) {
        if (visualsDetailsButtons) visualsDetailsButtons.classList.remove('hidden');
        var visualsBtn = document.querySelector('.comparison-buttons .sub-btn[data-target="visuals"]');
        if (visualsBtn) { 
            visualsBtn.classList.add('active'); 
            visualsBtn.classList.add('expanded'); 
        }
    } else if (filter !== 'visuals') {
        // Hide visuals details if we're not clicking Visuels or its sub-buttons
        if (visualsDetailsButtons && !visualsDetailsButtons.classList.contains('hidden')) {
            visualsDetailsButtons.classList.add('hidden');
            document.querySelectorAll('.third-level-buttons .third-btn').forEach(function(b){ b.classList.remove('active'); });
            var visualsBtn2 = document.querySelector('.comparison-buttons .sub-btn[data-target="visuals"]');
            if (visualsBtn2) { visualsBtn2.classList.remove('expanded'); }
        }
    }
    
    // Update active button state (first and second levels)
    document.querySelectorAll('.comparison-buttons .sub-btn').forEach(btn => {
        btn.classList.remove('active');
    });
    document.querySelectorAll('.third-level-buttons .third-btn').forEach(btn => { btn.classList.remove('active'); });
    
    var activeBtn = document.querySelector('.comparison-buttons .sub-btn[data-target="' + filter + '"]');
    if (activeBtn) {
        activeBtn.classList.add('active');
    }
    if (isSecondLevelFilter) {
        var parentBtn = document.querySelector('.comparison-buttons .sub-btn[data-target="visuals"]');
        if (parentBtn) {
            parentBtn.classList.add('active');
            parentBtn.classList.add('expanded');
        }
    }
    var activeSecond = document.querySelector('.third-level-buttons .third-btn[data-target="' + filter + '"]');
    if (activeSecond) {
        activeSecond.classList.add('active');
    }

    // Show or hide tables based on selection
    
    // Get all comparison tables
    var bookmarksTable = document.getElementById('table_bookmarks');
    var configTable = document.getElementById('table_config');
    var themesTable = document.getElementById('table_themes');
    var pagesTable = document.getElementById('table_pages');
    var visualsTable = document.getElementById('table_visuals');
    var syncTable = document.getElementById('table_synchronization');
    var interactionsTable = document.getElementById('table_visual_interactions');
    var fieldsTable = document.getElementById('table_visual_fields');
    var buttonsTable = document.getElementById('table_buttons');
    var visualsInfoSection = document.getElementById('visuals-info-section');
    
    // Hide all tables first
    [bookmarksTable, configTable, themesTable, pagesTable, visualsTable, syncTable, interactionsTable, fieldsTable, buttonsTable].forEach(table => {
        if (table) table.classList.add('hidden');
    });
    
    // Hide visuals info section by default
    if (visualsInfoSection) visualsInfoSection.classList.add('hidden');
    
    // Show selected table(s)
    switch(filter) {
        case 'bookmarks':
            if (bookmarksTable) bookmarksTable.classList.remove('hidden');
            break;
        case 'config':
            if (configTable) configTable.classList.remove('hidden');
            break;
        case 'themes':
            if (themesTable) themesTable.classList.remove('hidden');
            break;
        case 'pages':
            if (pagesTable) pagesTable.classList.remove('hidden');
            break;
        case 'visuals':
            if (visualsTable) visualsTable.classList.remove('hidden');
            if (visualsInfoSection) visualsInfoSection.classList.remove('hidden');
            break;
        case 'sync':
            if (syncTable) syncTable.classList.remove('hidden');
            break;
        case 'interactions':
            if (interactionsTable) interactionsTable.classList.remove('hidden');
            if (visualsInfoSection) visualsInfoSection.classList.remove('hidden');
            break;
        case 'fields':
            if (fieldsTable) fieldsTable.classList.remove('hidden');
            if (visualsInfoSection) visualsInfoSection.classList.remove('hidden');
            break;
        case 'buttons':
            if (buttonsTable) buttonsTable.classList.remove('hidden');
            if (visualsInfoSection) visualsInfoSection.classList.remove('hidden');
            break;
    }
}

// Search in comparison tables
function searchInComparisonTable() {
    var searchTerm = document.getElementById('searchBoxComparison').value.toLowerCase();
    var resultsDiv = document.getElementById('searchResultsComparison');
    var tables = document.querySelectorAll('#comparison-section table');
    var matchCount = 0;
    
    tables.forEach(table => {
        var rows = table.querySelectorAll('tbody tr');
        rows.forEach(row => {
            var text = row.textContent.toLowerCase();
            if (text.includes(searchTerm)) {
                row.style.display = '';
                matchCount++;
            } else {
                row.style.display = 'none';
            }
        });
    });
    
    if (searchTerm) {
        resultsDiv.textContent = matchCount + ' résultat(s) trouvé(s)';
    } else {
        resultsDiv.textContent = '';
    }
}

// Clear comparison search
function clearComparisonSearch() {
    document.getElementById('searchBoxComparison').value = '';
    document.getElementById('searchResultsComparison').textContent = '';
    var tables = document.querySelectorAll('#comparison-section table');
    tables.forEach(table => {
        var rows = table.querySelectorAll('tbody tr');
        rows.forEach(row => {
            row.style.display = '';
        });
    });
}

// ========== QUALITY TABLE SEARCH FUNCTIONS ==========

function searchInQualityTable() {
    var input = document.getElementById('searchBoxQuality');
    if (!input) {
        return;
    }

    var rawValue = input.value || '';
    var filter = rawValue.toLowerCase();
    var table = document.getElementById('table_quality_checks');
    if (!table) {
        return;
    }

    var tr = table.getElementsByTagName('tr');
    var foundCount = 0;

    for (var i = 2; i < tr.length; i++) { // Skip header rows
        var row = tr[i];
        var visible = false;
        var td = row.getElementsByTagName('td');
        
        for (var j = 0; j < td.length; j++) {
            if (td[j]) {
                var txtValue = td[j].textContent || td[j].innerText || '';
                if (txtValue.toLowerCase().indexOf(filter) > -1) {
                    visible = true;
                    break;
                }
            }
        }
        
        if (visible) {
            row.style.display = '';
            foundCount++;
        } else {
            row.style.display = 'none';
        }
    }

    currentQualitySearchText = rawValue;
    currentQualitySearchCount = foundCount;

    if (!rawValue.trim()) {
        currentQualitySearchText = '';
        currentQualitySearchCount = 0;
    }

    renderQualitySearchResults();
}

function clearQualitySearch() {
    document.getElementById('searchBoxQuality').value = '';
    searchInQualityTable();
}

// ========== HIERARCHICAL SYNC GROUP FUNCTIONS ==========

// Toggle function for sync group hierarchical display
function toggleSyncGroup(element) {
    var groupName = element.getAttribute('data-sync-group');
    if (!groupName) return;

    var childRows = document.querySelectorAll('[data-parent-group="' + groupName + '"]');
    var expandIcon = element.querySelector('.expand-icon');

    if (element.classList.contains('expanded')) {
        // Collapse
        childRows.forEach(function(row) {
            row.classList.add('hidden');
        });
        element.classList.remove('expanded');
    } else {
        // Expand
        childRows.forEach(function(row) {
            row.classList.remove('hidden');
        });
        element.classList.add('expanded');
    }
}

// ========== CAUSE → CONSEQUENCE FUNCTIONS ==========

// Toggle function for cause consequence display
function toggleCause(element) {
    var causeKey = element.getAttribute('data-cause-key');
    if (!causeKey) return;

    var childRows = document.querySelectorAll('[data-parent-cause="' + causeKey + '"]');
    var expandIcon = element.querySelector('.expand-icon');

    if (element.classList.contains('expanded')) {
        // Collapse
        childRows.forEach(function(row) {
            row.classList.add('hidden');
        });
        element.classList.remove('expanded');
    } else {
        // Expand
        childRows.forEach(function(row) {
            row.classList.remove('hidden');
        });
        element.classList.add('expanded');
    }
}

// Initialize: expand all causes by default
document.addEventListener('DOMContentLoaded', function() {
    var syncGroupHeaders = document.querySelectorAll('.sync-group-header, .sync-group-header-modified');
    syncGroupHeaders.forEach(function(header) {
        // Auto-expand groups with modifications
        if (header.classList.contains('sync-group-header-modified')) {
            toggleSyncGroup(header);
        }
    });

    // Auto-expand all causes to show consequences
    var causeHeaders = document.querySelectorAll('.cause-added, .cause-modified, .cause-removed');
    causeHeaders.forEach(function(header) {
        toggleCause(header);
    });
});

function initTableTruncate(tableID) {
	const maxLength = 100; // Limite de caractères
	
	const table = document.getElementById(tableID);
    if (!table) {
        console.error(`Tableau avec l'ID "${tableID}" non trouvé`);
        return;
    }

	// Cibler toutes les cellules du tableau
	table.querySelectorAll('td').forEach(cellule => {
		
		// Récupérer tous les nœuds texte (pas les boutons)
		const noeudsTexte = [];
		
		cellule.childNodes.forEach(noeud => {
			if (noeud.nodeType === Node.TEXT_NODE) {
				noeudsTexte.push(noeud);
			}
		});
		
		// Traiter chaque nœud texte
		noeudsTexte.forEach(noeudTexte => {
			const texte = noeudTexte.textContent.trim();
			
			if (texte.length > maxLength) {
				// Tronquer seulement le texte
				noeudTexte.textContent = texte.substring(0, maxLength);
				
				const truncateElement = document.createElement('span');
				truncateElement.textContent = '…';
				truncateElement.classList.add('truncateElement');
				noeudTexte.parentNode.insertBefore(truncateElement, noeudTexte.nextSibling);
			}
		});
	});
}

</script>
</head>
<body>

<!-- Header Orange -->
<header class="page-header">
    <div class="header-left">
        <div class="logo">
            <svg width="30" height="30" viewBox="0 0 30 30" fill="none" xmlns="http://www.w3.org/2000/svg">
                <path fill-rule="evenodd" clip-rule="evenodd" d="M0 30H30V0H0V30Z" style="fill: var(--brand-primary);"/>
                <path fill-rule="evenodd" clip-rule="evenodd" d="M4.28564 25.7144H25.7143V21.4287H4.28564V25.7144Z" style="fill: var(--surface-default);"/>
            </svg>
        </div>
        <h1 class="title" data-i18n-key="title">Rapports Power BI (PBIR)</h1>
    </div>
    <div class="header-right">
        <button class="language-switch" onclick="switchLanguage('fr')" data-lang="fr">
            <svg width="40" height="30" viewBox="0 0 40 30" fill="none" xmlns="http://www.w3.org/2000/svg">
                <path d="M0 0H40V30H0V0Z" fill="white"/>
                <path d="M0 0H13.1343V30H0V0Z" fill="#00267F"/>
                <path d="M26.8657 0H40.0001V30H26.8657V0Z" fill="#F31830"/>
            </svg>
        </button>
        <button class="language-switch" onclick="switchLanguage('en')" data-lang="en">
            <svg width="40" height="30" viewBox="0 0 40 30" fill="none" xmlns="http://www.w3.org/2000/svg">
                    <mask id="mask0_30_6160" style="mask-type:luminance" maskUnits="userSpaceOnUse" x="0" y="0" width="40" height="30">
                        <path fill-rule="evenodd" clip-rule="evenodd" d="M0 0H40V30H0V0Z" fill="white"/>
                    </mask>
                    <g mask="url(#mask0_30_6160)">
                        <path d="M0 0H40V30H0V0Z" fill="#012169"/>
                        <path fill-rule="evenodd" clip-rule="evenodd" d="M0 3.35156V0H4.47266L20 11.6468L15.5295 15L0 3.35156ZM20 18.3532L4.47266 30H0V26.6484L15.5295 15L20 18.3532ZM20 18.3532L24.4705 15L40 26.6484V30H35.5273L20 18.3532ZM24.4705 15L20 11.6468L35.5273 0H40V3.35156L24.4705 15Z" fill="white"/>
                        <path fill-rule="evenodd" clip-rule="evenodd" d="M16.668 10.002V0H23.332V10.002H16.668ZM16.668 19.998H0V10.002H16.668V19.998ZM16.668 19.998H23.332V30H16.668V19.998ZM23.332 19.998V10.002H40V19.998H23.332Z" fill="white"/>
                        <path fill-rule="evenodd" clip-rule="evenodd" d="M18 12V0H22V12H40V18H22V30H18V18H0V12H18ZM0 30L13.332 19.998H16.3164L2.98047 30H0ZM13.332 10.002L0 0V2.23828L10.3516 10.002H13.332ZM23.6875 10.002L37.0195 0H40L26.668 10.002H23.6875ZM26.668 19.998L40 30V27.7617L29.6484 19.998H26.668Z" fill="#C8102E"/>
                    </g>
                </svg>
        </button>
    </div>
</header>

<!-- Contenu principal -->
<main class="main-content">
    <div class="all-buttons">
        <!-- Boutons principaux Orange -->
        <div class="main-buttons">
            <button class="main-btn active" data-target="quality" onclick="showMainSection('quality')" data-i18n-key="mainButtons.quality">Check Qualité</button>
            <button class="main-btn" data-target="comparison" onclick="showMainSection('comparison')" data-i18n-key="mainButtons.comparison">Rapport</button>
            <button class="main-btn" data-target="power-query" onclick="showMainSection('power-query')" data-i18n-key="mainButtons.powerQuery">MDD-Power Query</button>
            <button class="main-btn" data-target="desktop" onclick="showMainSection('desktop')" data-i18n-key="mainButtons.desktop">MDD-OBI Desktop</button>
        </div>

        <!-- Ligne de séparation -->
        <div class="separator"></div>
    </div>

'@

    # Start of comparison section (Rapport)
    $reportFinal += '<div id="comparison-section" class="comparison-section section-hidden">'
    
    # Calculate comparison statistics
    if (-not $differences) {
        $differences = @()
    }

    $reportFinal += @"
<div class="comparison-buttons">
    <button class="sub-btn active" data-target="pages" onclick="filterComparisonTable('pages')" data-i18n-key="comparisonThirdButtons.pages">Pages</button>
    <button class="sub-btn has-children" data-target="visuals" onclick="toggleVisualsButtons()" data-i18n-key="comparisonThirdButtons.visuals">Visuels</button>
    <button class="sub-btn" data-target="sync" onclick="filterComparisonTable('sync')" data-i18n-key="comparisonThirdButtons.sync">Synchronisation</button>
    <button class="sub-btn" data-target="bookmarks" onclick="filterComparisonTable('bookmarks')" data-i18n-key="comparisonButtons.bookmarks">Signets</button>
    <button class="sub-btn" data-target="config" onclick="filterComparisonTable('config')" data-i18n-key="comparisonButtons.config">Configuration</button>
    <button class="sub-btn" data-target="themes" onclick="filterComparisonTable('themes')" data-i18n-key="comparisonButtons.themes">Thèmes</button>
</div>

<div class="third-level-buttons hidden" id="visuals-details-buttons">
    <button class="third-btn" data-target="interactions" onclick="filterComparisonTable('interactions')" data-i18n-key="comparisonThirdButtons.interactions">Interactions</button>
    <button class="third-btn" data-target="fields" onclick="filterComparisonTable('fields')" data-i18n-key="comparisonThirdButtons.fields">Champs du visuel</button>
    <button class="third-btn" data-target="buttons" onclick="filterComparisonTable('buttons')" data-i18n-key="comparisonThirdButtons.buttons">Boutons</button>
</div>

<div class="interactive-controls">
</div>
"@

    # Generation des tables de comparaison (toutes visibles dans la section Rapport)
    $reportFinal += GeneratePagesTable -differences $differences
    
    # Visuals section with info button (only visible when visuals table is active)
    $reportFinal += @"
<div id="visuals-info-section" class="hidden">
    <div class="visuals-info-bar">
        <button class="info-btn" onclick="toggleVisualsInfo()" data-i18n-key="visualsInfo.button">Informations</button>
    </div>
    <div id="visuals-info-panel" class="quality-info-panel hidden">
        <h3 data-i18n-key="visualsInfo.title">Aide - Visuels supprimés et recréés</h3>
        <p data-i18n-key="visualsInfo.message">Si un même visuel apparaît à la fois comme supprimé et ajouté, cela indique généralement qu'il a été supprimé puis recréé entre les deux versions du rapport. Dans ce cas, il est nécessaire de vérifier que le visuel recréé possède bien les mêmes configurations, interactions et propriétés que l'ancien, ou que cette modification est intentionnelle.</p>
    </div>
</div>
"@
    
    $reportFinal += GenerateVisualsTable -differences $differences
    $reportFinal += GenerateSynchronizationTable -differences $differences
    $reportFinal += GenerateVisualFieldsTable -differences $differences
    $reportFinal += GenerateButtonsTable -differences $differences
    $reportFinal += GenerateBookmarksTable -differences $differences
    $reportFinal += GenerateConfigTable -differences $differences
    $reportFinal += GenerateVisualInteractionsTable -differences $differences
    $reportFinal += GenerateThemesTable -differences $differences
    $reportFinal += "</div>"
    
    # End of comparison section
    $reportFinal += "</div>"
    
    # Start of quality section
    $reportFinal += '<div id="quality-section" class="quality-section section-hidden">'
    
    $reportFinal += @"
<div class="interactive-controls">
    <div class="quality-info-bar">
        <button class="info-btn" onclick="toggleQualityInfo()" data-i18n-key="qualityInfo.button">Informations</button>
    </div>
</div>
<div id="quality-info-panel" class="quality-info-panel hidden">
    <h3 data-i18n-key="qualityInfo.title">Règles du Check Qualité</h3>
    <ul>
        <li data-i18n-key="qualityInfo.rule_search">Aucun texte ne doit être présent dans la loupe (zone de recherche).</li>
        <li data-i18n-key="qualityInfo.rule_selection">Aucun filtre ne doit être coché, sauf pour les filtres en mode radio.</li>
        <li data-i18n-key="qualityInfo.rule_menu">Les filtres des groupes filter doivent être vides.</li>
        <li data-i18n-key="qualityInfo.rule_period">Les filtres des groupes period doivent contenir uniquement 'Current Year'.</li>
        <li data-i18n-key="qualityInfo.rule_pane">Le volet de filtre doit être masqué.</li>
    </ul>
</div>
"@
    
    # Generation du tableau des checks qualite
    if (-not $checkResults) {
        $checkResults = @()
    }
    $reportFinal += GenerateQualityChecksTable -checkResults $checkResults
    
    # End of quality section
    $reportFinal += "</div>"

    # Prepare modal tables container for loupe details across sections
    $modalTables = ""

    # Start of MDD-Power Query section (Orange design)
    $reportFinal += '<div id="power-query-section" class="power-query-section section-hidden">'
    
    if ($semanticComparisonResult) {
        # Power Query section header
        $reportFinal += @"
<!-- Sous-boutons pour MDD-Power Query -->
<div class="sub-buttons-group" id="power-query-subs">
    <button class="sub-btn has-children" data-target="table_query" onclick="showSubSection('table_query')" data-i18n-key="subButtons.tableQuery">Table query</button>
    <button class="sub-btn has-children" data-target="parameters_other" onclick="showSubSection('parameters_other')" data-i18n-key="subButtons.parametersOther">Parameters and other</button>
</div>

<!-- Boutons troisième niveau pour Table query -->
<div class="third-buttons" id="table_query-buttons">
    <button class="third-btn" data-table="semantic_table_tables" onclick="showSemanticTable('semantic_table_tables')" data-i18n-key="thirdButtons.tables">Names and Properties</button>
    <button class="third-btn" data-table="semantic_table_columns" onclick="showSemanticTable('semantic_table_columns')" data-i18n-key="thirdButtons.columns">Columns</button>
    <button class="third-btn" data-table="semantic_table_steps" onclick="showSemanticTable('semantic_table_steps')" data-i18n-key="thirdButtons.steps">Steps</button>
</div>

<!-- Boutons troisième niveau pour Parameters and other -->
<div class="third-buttons" id="parameters_other-buttons">
    <button class="third-btn" data-table="semantic_table_paramValue" onclick="showSemanticTable('semantic_table_paramValue')" data-i18n-key="thirdButtons.parameters">Parameters</button>
</div>

<div class="interactive-controls">
</div>
"@

        # Generate semantic tables by sourcing the functions from PBI_MDD_extract.ps1
        # Tables
        $objectsToShow = GetObjectToShow -comparisonResult $semanticComparisonResult -objectToGet "Table"
    $tables_table = ReportCreateCompareTable -listOfObjects $objectsToShow -objectsType 'Model.Tables' -tableName 'semantic_table_tables' -listOfProperties @("name", "isHidden", "IsPrivate", "ExcludeFromModelRefresh", "IsRemoved", "columns.count", "partitions.name", "changedProperties.property", "measures.count")
    $reportFinal += $tables_table[0]
    $modalTables += $tables_table[1]

        # Columns
        $objectsToShow = GetObjectToShow -comparisonResult $semanticComparisonResult -objectToGet @("DataColumn", "CalculatedColumn")
    $columns_table = ReportCreateCompareTable -listOfObjects $objectsToShow -objectsType 'Model.Tables.Columns' -tableName "semantic_table_columns" -listOfProperties @("table.name", "name", "type", "expression", "dataType", "sourceColumn", "formatString", "isAvailableInMdx", "summarizeBy.ToString", "sortByColumn", "changedProperties.property")
    $reportFinal += $columns_table[0]
    $modalTables += $columns_table[1]

        # Steps
        $objectsToShow = GetObjectToShow -comparisonResult $semanticComparisonResult -objectToGet "Partition"
    $steps_table = ReportCreateCompareTable -listOfObjects $objectsToShow -objectsType 'Model.Tables.partitions' -tableName "semantic_table_steps" -listOfProperties @("table.name", "queryGroup.folder", "source.expression")
    $reportFinal += $steps_table[0]
    $modalTables += $steps_table[1]

        # Parameter values
        $objectsToShow = GetObjectToShow -comparisonResult $semanticComparisonResult -objectToGet "NamedExpression" -filter "Parameters"
    $param_table = ReportCreateCompareTable -listOfObjects $objectsToShow -objectsType 'Specific.parameter' -tableName "semantic_table_paramValue" -listOfProperties @("name", "ExpressionValue", "ExpressionMeta")
    $reportFinal += $param_table[0]
    $modalTables += $param_table[1]
    }
    
    # End of MDD-Power Query section
    $reportFinal += "</div>"
    
    # Start of MDD-OBI Desktop section (Orange design)
    $reportFinal += '<div id="desktop-section" class="desktop-section section-hidden">'
    
    if ($semanticComparisonResult) {
        # Desktop section header
        $reportFinal += @"
<!-- Sous-boutons pour MDD-OBI Desktop -->
<div class="sub-buttons-group" id="desktop-subs">
    <button class="third-btn" data-table="semantic_table_relationships" onclick="showSemanticTable('semantic_table_relationships')" data-i18n-key="subButtons.relationships">Relations tables</button>
    <button class="third-btn" data-table="semantic_table_measures" onclick="showSemanticTable('semantic_table_measures')" data-i18n-key="subButtons.measures">DAX Measures</button>
    <button class="third-btn" data-table="semantic_table_roles" onclick="showSemanticTable('semantic_table_roles')" data-i18n-key="subButtons.rls">RLS</button>
    <button class="third-btn" data-table="semantic_table_perspectives" onclick="showSemanticTable('semantic_table_perspectives')" data-i18n-key="subButtons.perspectives">Perspectives</button>
    <button class="third-btn" data-table="semantic_table_calculGroups" onclick="showSemanticTable('semantic_table_calculGroups')" data-i18n-key="subButtons.calcGroups">Calculation Groups</button>
</div>

<div class="interactive-controls">
</div>
"@

        # Relationships
        $objectsToShow = GetObjectToShow -comparisonResult $semanticComparisonResult -objectToGet "SingleColumnRelationship"
    $rel_table = ReportCreateCompareTable -listOfObjects $objectsToShow -objectsType 'Model.relationships' -tableName "semantic_table_relationships" -listOfProperties @(@{Properties = @("fromTable.name", "fromColumn.name"); Sep = "."}, @{Properties = @("fromCardinality.ToString", "toCardinality.ToString"); Sep = " to "}, @{Properties = @("toTable.name", "toColumn.name"); Sep = "."}, "crossFilteringBehavior")
    $reportFinal += $rel_table[0]
    $modalTables += $rel_table[1]

        # Measures DAX
        $objectsToShow = GetObjectToShow -comparisonResult $semanticComparisonResult -objectToGet "Measure"
    $measures_table = ReportCreateCompareTable -listOfObjects $objectsToShow -objectsType 'Model.Tables.Measures' -tableName "semantic_table_measures" -listOfProperties @("name", "expression", "formatString", "isHidden", "displayFolder", "changedProperties.property")
    $reportFinal += $measures_table[0]
    $modalTables += $measures_table[1]

        # Roles
        $objectsToShow = GetObjectToShow -comparisonResult $semanticComparisonResult -objectToGet "TablePermission"
    $roles_table = ReportCreateCompareTable -listOfObjects $objectsToShow -objectsType 'Model.roles.TablePermissions' -tableName "semantic_table_roles" -listOfProperties @("role.name", "table.name", "filterExpression")
    $reportFinal += $roles_table[0]
    $modalTables += $roles_table[1]

        # Perspectives
        $objectsToShow = GetObjectToShow -comparisonResult $semanticComparisonResult -objectToGet "PerspectiveColumn"
    $persp_table = ReportCreateCompareTable -listOfObjects $objectsToShow -objectsType 'Model.Perspectives.PerspectiveTables.PerspectiveColumns' -tableName "semantic_table_perspectives" -listOfProperties @("PerspectiveTable.Perspective.name", "PerspectiveTable.name", "name")
    $reportFinal += $persp_table[0]
    $modalTables += $persp_table[1]

        # Calculation groups
        $objectsToShow = GetObjectToShow -comparisonResult $semanticComparisonResult -objectToGet "CalculationItem"
    $calc_table = ReportCreateCompareTable -listOfObjects $objectsToShow -objectsType 'Model.Perspectives.PerspectiveTables.PerspectiveColumns' -tableName "semantic_table_calculGroups" -listOfProperties @("CalculationGroup.Table.name", "CalculationGroup.description", "name", "description", "state", "expression")
    $reportFinal += $calc_table[0]
    $modalTables += $calc_table[1]
    }
    
    # End of MDD-OBI Desktop section
    $reportFinal += "</div>"
    
        # Modal container for loupe details (inject all additional semantic tables here)
        $reportFinal += @"
<div id="modal" class="modal hidden">
    <div class="modal-content">
        <div class="modal-header">
            <span>Details</span>
            <button class="close" onclick="closeModalFunction()">&times;</button>
        </div>
        <div class="modal-body" id="modal-body">
            $modalTables
        </div>
    </div>
    <!-- backdrop is the modal itself -->
</div>

<!-- Modal Settings -->
<div id="settingsModal" class="modal hidden">
    <div class="modal-content" style="max-width: 500px;">
        <div class="modal-header">
            <span data-i18n-key="settings.title">⚙️ Paramètres</span>
            <button class="close" onclick="closeSettingsModal()">&times;</button>
        </div>
        <div class="modal-body" style="padding: 25px;">
            <p data-i18n-key="settings.intro" style="margin-bottom: 20px; color: var(--text-subtle); font-size: 14px;">
                Configurez le dossier de sortie par défaut.
            </p>
            
            <div style="margin-bottom: 25px;">
                <label for="outputPathInput" data-i18n-key="settings.outputPath.label" style="display: block; margin-bottom: 8px; font-weight: 600;">
                    📁 Dossier de sortie :
                </label>
                <div style="display: flex; gap: 10px; margin-bottom: 8px;">
                    <input 
                        type="text" 
                        id="outputPathInput" 
                        placeholder="Ex: C:\Users\VotreNom\Documents\Rapports_PowerBI" 
                        style="flex: 1; padding: 10px; border: 2px solid var(--border-default); border-radius: 6px; font-family: 'Consolas', monospace; font-size: 13px;"
                        data-i18n-placeholder="settings.outputPath.placeholder"
                    />
                    <button 
                        onclick="pasteFromClipboard()" 
                        style="background-color: transparent; color: var(--text-default); border: 2px solid var(--border-default); padding: 10px 20px; border-radius: 6px; cursor: pointer; font-size: 14px; font-weight: 600; transition: all 0.3s ease; white-space: nowrap;"
                        onmouseover="this.style.backgroundColor='var(--surface-hovered)'"
                        onmouseout="this.style.backgroundColor='transparent'"
                        data-i18n-key="settings.pasteButton"
                        title="Coller le chemin depuis le presse-papiers"
                    >📋 Coller</button>
                </div>
                <small data-i18n-key="settings.outputPath.hint" style="color: var(--text-subtle); font-size: 12px; display: block; margin-bottom: 8px;">
                    � Dans l'Explorateur Windows, cliquez dans la barre d'adresse du dossier souhaité et copiez le chemin (Ctrl+C), puis cliquez "Coller" ci-dessus.
                </small>
                <small data-i18n-key="settings.outputPath.hint2" style="color: var(--text-subtle); font-size: 11px;">
                    Laissez vide pour être sollicité à chaque exécution
                </small>
            </div>
            
            <div style="text-align: center; display: flex; gap: 10px; justify-content: center;">
                <button 
                    onclick="saveConfig()" 
                    style="background-color: var(--brand-primary); color: white; border: none; padding: 12px 30px; border-radius: 8px; cursor: pointer; font-size: 15px; font-weight: 600; transition: all 0.3s ease;"
                    onmouseover="this.style.backgroundColor='var(--brand-primary-hover)'"
                    onmouseout="this.style.backgroundColor='var(--brand-primary)'"
                    data-i18n-key="settings.saveButton"
                >Save</button>
                <button 
                    onclick="resetConfig()" 
                    style="background-color: transparent; color: var(--text-default); border: 2px solid var(--border-default); padding: 12px 24px; border-radius: 8px; cursor: pointer; font-size: 15px; font-weight: 600; transition: all 0.3s ease;"
                    onmouseover="this.style.backgroundColor='var(--surface-hovered)'"
                    onmouseout="this.style.backgroundColor='transparent'"
                    data-i18n-key="settings.resetButton"
                >Reset</button>
            </div>
        </div>
    </div>
</div>
"@

        # Close main content container
    $reportFinal += "</main>"

    # Orange footer
    $generatedDate = Get-Date -Format 'dd/MM/yyyy HH:mm:ss'
    $generatedDateIso = Get-Date -Format "yyyy-MM-ddTHH:mm:ssK"
    
    # Inject current output path and config path into the HTML for settings modal
    $escapedOutputFolder = $outputFolder.Replace('\', '\\')
    $escapedConfigPath = $configPath.Replace('\', '\\')
    
    $reportFinal += @"
<footer class="footer">
    <div class="footer-brand">© Orange Business 2025</div>
    <div class="footer-message">
        <span data-i18n-key="footer.analysis">Analyse des modifications entre deux versions de rapport Power BI (format PBIR)</span>
        <span class="footer-separator">•</span>
        <span class="footer-generated">
            <span data-i18n-key="footer.generatedPrefix">Rapport genere le</span>
            <time datetime="$generatedDateIso" data-footer-generated-date="true">$generatedDate</time>
            <span data-i18n-key="footer.generatedSuffix">par PBI_Report_Compare.ps1</span>
        </span>
    </div>
</footer>

<script>
// Inject config.json path from PowerShell
window.INJECTED_CONFIG_PATH = '$escapedConfigPath';

// Load and inject configuration for settings modal
(function() {
    document.addEventListener('DOMContentLoaded', function() {
        var inputField = document.getElementById('outputPathInput');
        if (!inputField) return;
        
        // Get the path used during this report generation (injected by PowerShell)
        var currentOutputPath = '$escapedOutputFolder';
        
        // Priority 1: Check if user has previously saved a custom path in localStorage
        var savedConfig = localStorage.getItem('pbi-report-config');
        if (savedConfig) {
            try {
                var config = JSON.parse(savedConfig);
                if (config.defaultOutputPath && config.defaultOutputPath.trim()) {
                    // User has a saved preference - use it
                    inputField.value = config.defaultOutputPath;
                    return;
                }
            } catch (e) {
                console.error('Error parsing saved config:', e);
            }
        }
        
        // Priority 2: No saved preference - use the path from current main.ps1 execution
        // This is the folder selected during main.ps1 steps 1-3
        if (currentOutputPath && currentOutputPath.trim()) {
            inputField.value = currentOutputPath;
            console.log('Pre-filled with current execution path:', currentOutputPath);
        }
    });
})();
</script>

</body>
</html>
"@

    # Ecriture du fichier HTML (version Orange)
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $outputPath = Join-Path $outputFolder "Rapport_PowerBI_Reports_Comparison_ORANGE_$timestamp.html"
    $reportFinal | Out-File -FilePath $outputPath -Encoding UTF8
    
    Write-Host "Rapport HTML Orange genere: $outputPath" -ForegroundColor Magenta
    
    # Créer un raccourci .url pour lancer configure.ps1 facilement depuis le rapport HTML
    try {
        $shortcutPath = Join-Path $outputFolder "Configurer-Chemin-Sortie.url"
        
        # Trouver le dossier Script_Merged en partant du configPath fourni
        if ($configPath -and (Test-Path $configPath)) {
            $scriptRoot = Split-Path -Parent -Path $configPath
            $configurePsPath = Join-Path $scriptRoot "configure.ps1"
            
            # Vérifier que configure.ps1 existe
            if (Test-Path $configurePsPath) {
                # Format Windows Internet Shortcut qui lance PowerShell pour exécuter le script
                $shortcutContent = @"
[InternetShortcut]
URL=file:///$($configurePsPath -replace '\\', '/')
IconIndex=0
"@
                Set-Content -Path $shortcutPath -Value $shortcutContent -Encoding ASCII
                Write-Host "  Raccourci de configuration cree: $shortcutPath" -ForegroundColor Cyan
            } else {
                Write-Host "  Note: configure.ps1 introuvable, raccourci non cree" -ForegroundColor Yellow
            }
        }
    } catch {
        Write-Host "  Erreur lors de la creation du raccourci: $($_.Exception.Message)" -ForegroundColor Yellow
    }
    
    return $outputPath
}

# Fonction utilitaire pour generer des descriptions claires pour l'utilisateur
Function Get-UserFriendlyDescription {
    param($diff)

    $elementName = if ($diff.ElementDisplayName) { $diff.ElementDisplayName } else { $diff.ElementName }
    $elementName = Remove-TechnicalMarkers -Text $elementName
    $elementTypeFr = $diff.ElementType
    $modificationType = $diff.DifferenceType
    $property = $diff.PropertyName

    $fr = ""
    $en = ""

    switch ("$elementTypeFr|$modificationType") {
        "Page|Ajoute" {
            $fr = "La page '$elementName' a été ajoutée au rapport"
            $en = "Page '$elementName' was added to the report"
        }
        "Page|Supprime" {
            $fr = "La page '$elementName' a été supprimée du rapport"
            $en = "Page '$elementName' was removed from the report"
        }
        "Page|Modifie" {
            if ($property -eq "filters") {
                $fr = "La page '$elementName' : les filtres ont été modifiés"
                $en = "Page '$elementName': filters were modified"
            }
            elseif ($property -eq "displayName") {
                $fr = "La page '$elementName' a été renommée"
                $en = "Page '$elementName' was renamed"
            }
            elseif ($property -eq "visibility") {
                $fr = "La page '$elementName' : la visibilité a été modifiée"
                $en = "Page '$elementName': visibility changed"
            }
            else {
                $fr = "La page '$elementName' : la propriété '$property' a été modifiée"
                $en = "Page '$elementName': property '$property' was modified"
            }
        }
        "Visuel|Ajoute" {
            if ($property -eq "fields") {
                # Champ ajouté au visuel
                $fieldName = Remove-TechnicalMarkers -Text $diff.NewValue
                $fr = "Champ '$fieldName' ajouté au visuel '$elementName'"
                $en = "Field '$fieldName' added to visual '$elementName'"
            }
            elseif ($property -eq "syncGroup") {
                # Synchronisation ajoutée au slicer - utiliser AdditionalInfo qui contient le contexte complet
                if ($diff.AdditionalInfo) {
                    $fr = Sanitize-AdditionalInfoForEndUser -Text $diff.AdditionalInfo
                    # Traduire pour l'anglais
                    $en = (Convert-AdditionalInfoToEnglish $diff.AdditionalInfo) -replace "Groupe de synchronisation", "Sync group" `
                                              -replace "page", "page" `
                                              -replace "ajoutee au groupe", "added to group" `
                                              -replace "Slicer", "Slicer"
                    $en = Sanitize-AdditionalInfoForEndUser -Text $en
                } else {
                    $pageName = if ($diff.ParentDisplayName) { Remove-TechnicalMarkers -Text $diff.ParentDisplayName } else { "page inconnue" }
                    $fr = "Page '$pageName' ajoutée au groupe de synchronisation"
                    $en = "Page '$pageName' added to synchronization group"
                }
            }
            else {
                $fr = "Le visuel '$elementName' a été ajouté"
                $en = "Visual '$elementName' was added"
            }
        }
        "Visuel|Supprime" {
            if ($property -eq "fields") {
                # Champ supprimé du visuel
                $fieldName = Remove-TechnicalMarkers -Text $diff.OldValue
                $fr = "Champ '$fieldName' supprimé du visuel '$elementName'"
                $en = "Field '$fieldName' removed from visual '$elementName'"
            }
            elseif ($property -eq "syncGroup") {
                # Synchronisation supprimée du slicer - utiliser AdditionalInfo qui contient le contexte complet
                if ($diff.AdditionalInfo) {
                    $fr = Sanitize-AdditionalInfoForEndUser -Text $diff.AdditionalInfo
                    # Traduire pour l'anglais
                    $en = (Convert-AdditionalInfoToEnglish $diff.AdditionalInfo) -replace "Groupe de synchronisation", "Sync group" `
                                              -replace "page", "page" `
                                              -replace "retiree du groupe", "removed from group" `
                                              -replace "Slicer", "Slicer"
                    $en = Sanitize-AdditionalInfoForEndUser -Text $en
                } else {
                    $pageName = if ($diff.ParentDisplayName) { Remove-TechnicalMarkers -Text $diff.ParentDisplayName } else { "page inconnue" }
                    $fr = "Page '$pageName' retirée du groupe de synchronisation"
                    $en = "Page '$pageName' removed from synchronization group"
                }
            }
            else {
                $fr = "Le visuel '$elementName' a été supprimé"
                $en = "Visual '$elementName' was removed"
            }
        }
        "Visuel|Modifie" {
            if ($property -eq "fields_summary") {
                # Résumé des changements de champs
                $summary = Remove-TechnicalMarkers -Text $diff.OldValue
                $fr = "Le visuel '$elementName' : $summary"
                $en = "Visual '$elementName': $summary"
            }
            elseif ($property -eq "syncGroup.groupName") {
                # Changement de groupe de synchronisation
                $oldGroup = Remove-TechnicalMarkers -Text $diff.OldValue
                $newGroup = Remove-TechnicalMarkers -Text $diff.NewValue
                $pageName = if ($diff.ParentDisplayName) { Remove-TechnicalMarkers -Text $diff.ParentDisplayName } else { "page inconnue" }
                $fr = "Groupe de synchronisation du slicer '$elementName' sur page '$pageName' changé : '$oldGroup' → '$newGroup'"
                $en = "Slicer '$elementName' sync group on page '$pageName' changed: '$oldGroup' → '$newGroup'"
            }
            elseif ($property -eq "syncGroup.fieldChanges") {
                # Changement de synchronisation des champs
                $status = if ($diff.NewValue -eq "True") { "activée" } else { "désactivée" }
                $statusEn = if ($diff.NewValue -eq "True") { "enabled" } else { "disabled" }
                $pageName = if ($diff.ParentDisplayName) { Remove-TechnicalMarkers -Text $diff.ParentDisplayName } else { "page inconnue" }
                $fr = "Synchronisation des champs $status pour le slicer '$elementName' sur page '$pageName'"
                $en = "Field synchronization $statusEn for slicer '$elementName' on page '$pageName'"
            }
            elseif ($property -eq "syncGroup.filterChanges") {
                # Changement de synchronisation des filtres
                $status = if ($diff.NewValue -eq "True") { "activée" } else { "désactivée" }
                $statusEn = if ($diff.NewValue -eq "True") { "enabled" } else { "disabled" }
                $pageName = if ($diff.ParentDisplayName) { Remove-TechnicalMarkers -Text $diff.ParentDisplayName } else { "page inconnue" }
                $fr = "Synchronisation des filtres $status pour le slicer '$elementName' sur page '$pageName'"
                $en = "Filter synchronization $statusEn for slicer '$elementName' on page '$pageName'"
            }
            elseif ($property -eq "position") {
                $fr = "Le visuel '$elementName' a été déplacé ou redimensionné"
                $en = "Visual '$elementName' was moved or resized"
            }
            elseif ($property -eq "query") {
                $fr = "Le visuel '$elementName' : les données sources ont été modifiées"
                $en = "Visual '$elementName': source data changed"
            }
            elseif ($property -eq "formatting") {
                $fr = "Le visuel '$elementName' : le formatage a été modifié"
                $en = "Visual '$elementName': formatting changed"
            }
            else {
                $fr = "Le visuel '$elementName' : la propriété '$property' a été modifiée"
                $en = "Visual '$elementName': property '$property' was modified"
            }
        }
        "Signet|Ajoute" {
            $fr = "Le signet '$elementName' a été créé"
            $en = "Bookmark '$elementName' was created"
        }
        "Signet|Supprime" {
            $fr = "Le signet '$elementName' a été supprimé"
            $en = "Bookmark '$elementName' was removed"
        }
        "Signet|Modifie" {
            $fr = "Le signet '$elementName' : la propriété '$property' a été modifiée"
            $en = "Bookmark '$elementName': property '$property' was modified"
        }
        "Configuration|Modifie" {
            if ($property -eq "displayOption") {
                # Traduction des valeurs displayOption
                $oldValueFr = switch ($diff.OldValue) {
                    "FitToPage" { "Ajuster à la page" }
                    "FitToWidth" { "Ajuster à la largeur" }
                    "ActualSize" { "Taille réelle" }
                    default { $diff.OldValue }
                }
                $newValueFr = switch ($diff.NewValue) {
                    "FitToPage" { "Ajuster à la page" }
                    "FitToWidth" { "Ajuster à la largeur" }
                    "ActualSize" { "Taille réelle" }
                    default { $diff.NewValue }
                }
                $oldValueEn = switch ($diff.OldValue) {
                    "FitToPage" { "Fit to page" }
                    "FitToWidth" { "Fit to width" }
                    "ActualSize" { "Actual size" }
                    default { $diff.OldValue }
                }
                $newValueEn = switch ($diff.NewValue) {
                    "FitToPage" { "Fit to page" }
                    "FitToWidth" { "Fit to width" }
                    "ActualSize" { "Actual size" }
                    default { $diff.NewValue }
                }
                $pageName = if ($diff.ElementDisplayName) { Remove-TechnicalMarkers -Text $diff.ElementDisplayName } else { "page inconnue" }
                $fr = "L'affichage de la page $pageName a changé de '$oldValueFr' à '$newValueFr'"
                $en = "Page $pageName display changed from '$oldValueEn' to '$newValueEn'"
            }
            elseif ($property -eq "canvasType") {
                $pageName = if ($diff.ElementDisplayName) { Remove-TechnicalMarkers -Text $diff.ElementDisplayName } else { "page inconnue" }
                $fr = "Le format du canevas de la page $pageName a changé de '$($diff.OldValue)' à '$($diff.NewValue)'"
                $en = "Canvas format of page $pageName changed from '$($diff.OldValue)' to '$($diff.NewValue)'"
            }
            elseif ($property -eq "verticalAlignment") {
                # Traduction des valeurs d'alignement
                $oldValueFr = switch ($diff.OldValue) {
                    "Top" { "Haut" }
                    "Middle" { "Centre" }
                    "Bottom" { "Bas" }
                    default { $diff.OldValue }
                }
                $newValueFr = switch ($diff.NewValue) {
                    "Top" { "Haut" }
                    "Middle" { "Centre" }
                    "Bottom" { "Bas" }
                    default { $diff.NewValue }
                }
                $oldValueEn = $diff.OldValue
                $newValueEn = $diff.NewValue
                $pageName = if ($diff.ElementDisplayName) { Remove-TechnicalMarkers -Text $diff.ElementDisplayName } else { "page inconnue" }
                $fr = "L'alignement vertical de la page $pageName a changé de '$oldValueFr' à '$newValueFr'"
                $en = "Vertical alignment of page $pageName changed from '$oldValueEn' to '$newValueEn'"
            }
            else {
                $fr = "Configuration du rapport : la propriété '$property' a été modifiée"
                $en = "Report configuration: property '$property' was modified"
            }
        }
        "Theme|Modifie" {
            $fr = "Thème du rapport : la propriété '$property' a été modifiée"
            $en = "Report theme: property '$property' was modified"
        }
        "Synchronisation|Ajoute" {
            $fr = "Synchronisation ajoutée pour '$elementName'"
            $en = "Synchronization added for '$elementName'"
        }
        "Synchronisation|Supprime" {
            $fr = "Synchronisation supprimée pour '$elementName'"
            $en = "Synchronization removed for '$elementName'"
        }
        "Synchronisation|Modifie" {
            # Utiliser AdditionalInfo qui contient le détail complet
            if ($diff.AdditionalInfo) {
                $fr = Sanitize-AdditionalInfoForEndUser -Text $diff.AdditionalInfo
                # Traduire pour l'anglais
                $en = (Convert-AdditionalInfoToEnglish $diff.AdditionalInfo) -replace "sur page", "on page" `
                                          -replace "ajoutée", "added" `
                                          -replace "retirée", "removed" `
                                          -replace "sync activée", "sync enabled" `
                                          -replace "sync désactivée", "sync disabled" `
                                          -replace "visible activée", "visible enabled" `
                                          -replace "visible désactivée", "visible disabled" `
                                          -replace "Modifications de synchronisation", "Synchronization changes" `
                                          -replace "Synchronisation modifiée", "Synchronization modified"
                $en = Sanitize-AdditionalInfoForEndUser -Text $en
            } else {
                $fr = "Synchronisation modifiée pour '$elementName'"
                $en = "Synchronization modified for '$elementName'"
            }
        }
        "Bouton|Ajoute" {
            $fr = "Le bouton '$elementName' a été ajouté"
            $en = "Button '$elementName' was added"
        }
        "Bouton|Supprime" {
            $fr = "Le bouton '$elementName' a été supprimé"
            $en = "Button '$elementName' was removed"
        }
        "Bouton|Modifie" {
            $fr = "Le bouton '$elementName' : la propriété '$property' a été modifiée"
            $en = "Button '$elementName': property '$property' was modified"
        }
        default {
            $actionFr = switch ($modificationType) {
                "Ajoute" { "a été ajouté(e)" }
                "Supprime" { "a été supprimé(e)" }
                "Modifie" { "a été modifié(e)" }
                default { "a subi des modifications" }
            }
            $actionEn = switch ($modificationType) {
                "Ajoute" { "was added" }
                "Supprime" { "was removed" }
                "Modifie" { "was modified" }
                default { "has changed" }
            }
            $fr = "$elementTypeFr '$elementName' $actionFr"
            $en = "$(Convert-ElementTypeToEnglish $elementTypeFr) '$elementName' $actionEn"
        }
    }

    return [pscustomobject]@{
        fr = $fr
        en = $en
    }
}

# Fonctions de generation des tables simplifiees
Function GeneratePagesTable {
    param([ReportDifference[]] $differences)

    # Include Page differences AND displayOption from Configuration (as it's a page property)
    $pageDifferences = $differences | Where-Object { $_.ElementType -eq 'Page' -or ($_.ElementType -eq 'Configuration' -and $_.PropertyName -eq 'displayOption') }
    $columns = Get-DefaultComparisonColumns `
        -NameHeaderKey 'tables.common.headers.name' -NameHeaderText 'Nom' `
        -TypeHeaderKey 'tables.table_pages.headers.property' -TypeHeaderText 'Propriete' `
        -ValueHeaderKey 'tables.common.headers.value' -ValueHeaderText 'Valeur / Details'

    $rows = @()

    foreach ($diff in $pageDifferences) {
        # Special handling for displayOption (coming from Configuration ElementType)
        if ($diff.PropertyName -eq 'displayOption') {
            # Extract page name from ElementDisplayName for displayOption
            $pageName = Get-PageNameFromElement -elementDisplayName $diff.ElementDisplayName
            if ([string]::IsNullOrEmpty($pageName) -or $pageName -eq $diff.ElementName) {
                $displayPair = Get-SafeTextPair -Text $diff.ElementName
            }
            else {
                $displayPair = Get-SafeTextPair -Text $pageName
            }
            $propertyPair = @{fr="Mode d'affichage"; en="Display Mode"}
        }
        else {
            $displayPair = Get-DisplayNamePair -DisplayName $diff.ElementDisplayName -TechnicalName $diff.ElementName -ElementType $diff.ElementType
            $propertyPair = Get-SafeTextPair -Text $diff.PropertyName
        }

        $oldValuePair = Get-DiffValuePair -Value $diff.OldValue
        $newValuePair = Get-DiffValuePair -Value $diff.NewValue

        # Translate displayOption values
        if ($diff.PropertyName -eq 'displayOption') {
            $mappingFr = @{ FitToPage = "Ajuster à la page"; FitToWidth = "Ajuster à la largeur"; ActualSize = "Taille réelle" }
            $mappingEn = @{ FitToPage = "Fit to Page"; FitToWidth = "Fit to Width"; ActualSize = "Actual Size" }
            if ($mappingFr.ContainsKey($diff.OldValue)) { $oldValuePair.fr = $mappingFr[$diff.OldValue] }
            if ($mappingEn.ContainsKey($diff.OldValue)) { $oldValuePair.en = $mappingEn[$diff.OldValue] }
            if ($mappingFr.ContainsKey($diff.NewValue)) { $newValuePair.fr = $mappingFr[$diff.NewValue] }
            if ($mappingEn.ContainsKey($diff.NewValue)) { $newValuePair.en = $mappingEn[$diff.NewValue] }
        }

        $additionalFr = $null
        $additionalEn = $null
        if ($diff.AdditionalInfo) {
            $additionalFr = Sanitize-AdditionalInfoForEndUser -Text $diff.AdditionalInfo
            $additionalEnRaw = Convert-AdditionalInfoToEnglish $diff.AdditionalInfo
            $additionalEn = if ([string]::IsNullOrWhiteSpace($additionalEnRaw)) { $additionalFr } else { Sanitize-AdditionalInfoForEndUser -Text $additionalEnRaw }
            if ($additionalFr -eq '-' -or [string]::IsNullOrWhiteSpace($additionalFr)) { $additionalFr = $null }
            if ($additionalEn -eq '-' -or [string]::IsNullOrWhiteSpace($additionalEn)) { $additionalEn = $null }
        }

        $descriptionPair = Get-UserFriendlyDescription -diff $diff

        $oldNotesFr = @()
        $oldNotesEn = @()
        $newNotesFr = @()
        $newNotesEn = @()

        if ($additionalFr) {
            switch ($diff.DifferenceType) {
                'Ajoute' {
                    $newNotesFr += $additionalFr
                    $newNotesEn += $additionalEn
                }
                'Supprime' {
                    $oldNotesFr += $additionalFr
                    $oldNotesEn += $additionalEn
                }
                default {
                    $oldNotesFr += $additionalFr
                    $oldNotesEn += $additionalEn
                    $newNotesFr += $additionalFr
                    $newNotesEn += $additionalEn
                }
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($descriptionPair.fr)) {
            $newNotesFr += $descriptionPair.fr
            $newNotesEn += $descriptionPair.en
        }

        $rows += [ordered]@{
            CssClass = Get-DifferenceCssClass $diff.DifferenceType
            Status = $diff.DifferenceType
            Old = @{
                name  = New-RichCellContent -PrimaryFr $displayPair.fr -PrimaryEn $displayPair.en
                type  = New-RichCellContent -PrimaryFr $propertyPair.fr -PrimaryEn $propertyPair.en
                value = New-RichCellContent -PrimaryFr $oldValuePair.fr -PrimaryEn $oldValuePair.en -NotesFr $oldNotesFr -NotesEn $oldNotesEn
            }
            New = @{
                name  = New-RichCellContent -PrimaryFr $displayPair.fr -PrimaryEn $displayPair.en
                type  = New-RichCellContent -PrimaryFr $propertyPair.fr -PrimaryEn $propertyPair.en
                value = New-RichCellContent -PrimaryFr $newValuePair.fr -PrimaryEn $newValuePair.en -NotesFr $newNotesFr -NotesEn $newNotesEn
            }
        }
    }

    return Build-TwoBlockTable `
        -TableId 'table_pages' `
        -TitleKey 'tables.table_pages.title' `
        -TitleText 'Modifications des pages du rapport' `
        -EmptyKey 'tables.table_pages.empty' `
        -EmptyText 'Aucune modification detectee dans les pages' `
        -Columns $columns `
        -Rows $rows
}

Function GenerateVisualsTable {
    param([ReportDifference[]] $differences)

    $visualDifferences = $differences | Where-Object { $_.ElementType -eq 'Visuel' -and $_.PropertyName -ne 'fields' }
    $columns = Get-DefaultComparisonColumns `
        -NameHeaderKey 'tables.common.headers.name' -NameHeaderText 'Nom' `
        -TypeHeaderKey 'tables.common.headers.type' -TypeHeaderText 'Type' `
        -ValueHeaderKey 'tables.common.headers.value' -ValueHeaderText 'Valeur / Details'

    $rows = @()

    foreach ($diff in $visualDifferences) {
        $displayPair = Get-DisplayNamePair -DisplayName $diff.ElementDisplayName -TechnicalName $diff.ElementName -ElementType $diff.ElementType
        $propertyPair = Get-SafeTextPair -Text $diff.PropertyName

        $pageRaw = if ($diff.ParentDisplayName) { $diff.ParentDisplayName } elseif ($diff.ParentElementName) { $diff.ParentElementName } else { '' }
        $pagePair = Get-SafeTextPair -Text $pageRaw -DefaultFr 'Page inconnue' -DefaultEn 'Unknown page'

        $oldValuePair = Get-DiffValuePair -Value $diff.OldValue
        $newValuePair = Get-DiffValuePair -Value $diff.NewValue

        $additionalFr = $null
        $additionalEn = $null
        if ($diff.AdditionalInfo) {
            $additionalFr = Sanitize-AdditionalInfoForEndUser -Text $diff.AdditionalInfo
            $additionalEnRaw = Convert-AdditionalInfoToEnglish $diff.AdditionalInfo
            $additionalEn = if ([string]::IsNullOrWhiteSpace($additionalEnRaw)) { $additionalFr } else { Sanitize-AdditionalInfoForEndUser -Text $additionalEnRaw }
            if ($additionalFr -eq '-' -or [string]::IsNullOrWhiteSpace($additionalFr)) { $additionalFr = $null }
            if ($additionalEn -eq '-' -or [string]::IsNullOrWhiteSpace($additionalEn)) { $additionalEn = $null }
        }

        $descriptionPair = Get-UserFriendlyDescription -diff $diff

        $nameTagsFr = @()
        $nameTagsEn = @()
        if (-not [string]::IsNullOrWhiteSpace($pagePair.fr) -and $pagePair.fr -ne '-') {
            $nameTagsFr += "Page : $($pagePair.fr)"
            $nameTagsEn += "Page: $($pagePair.en)"
        }

        $oldNotesFr = @()
        $oldNotesEn = @()
        $newNotesFr = @()
        $newNotesEn = @()

        if ($additionalFr) {
            switch ($diff.DifferenceType) {
                'Ajoute' {
                    $newNotesFr += $additionalFr
                    $newNotesEn += $additionalEn
                }
                'Supprime' {
                    $oldNotesFr += $additionalFr
                    $oldNotesEn += $additionalEn
                }
                default {
                    $oldNotesFr += $additionalFr
                    $oldNotesEn += $additionalEn
                    $newNotesFr += $additionalFr
                    $newNotesEn += $additionalEn
                }
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($descriptionPair.fr)) {
            $newNotesFr += $descriptionPair.fr
            $newNotesEn += $descriptionPair.en
        }

        $rows += [ordered]@{
            CssClass = Get-DifferenceCssClass $diff.DifferenceType
            Status = $diff.DifferenceType
            Old = @{
                name  = New-RichCellContent -PrimaryFr $displayPair.fr -PrimaryEn $displayPair.en -TagsFr $nameTagsFr -TagsEn $nameTagsEn
                type  = New-RichCellContent -PrimaryFr $propertyPair.fr -PrimaryEn $propertyPair.en
                value = New-RichCellContent -PrimaryFr $oldValuePair.fr -PrimaryEn $oldValuePair.en -NotesFr $oldNotesFr -NotesEn $oldNotesEn
            }
            New = @{
                name  = New-RichCellContent -PrimaryFr $displayPair.fr -PrimaryEn $displayPair.en -TagsFr $nameTagsFr -TagsEn $nameTagsEn
                type  = New-RichCellContent -PrimaryFr $propertyPair.fr -PrimaryEn $propertyPair.en
                value = New-RichCellContent -PrimaryFr $newValuePair.fr -PrimaryEn $newValuePair.en -NotesFr $newNotesFr -NotesEn $newNotesEn
            }
        }
    }

    return Build-TwoBlockTable `
        -TableId 'table_visuals' `
        -TitleKey 'tables.table_visuals.title' `
        -TitleText 'Modifications des visuels' `
        -EmptyKey 'tables.table_visuals.empty' `
        -EmptyText 'Aucune modification detectee dans les visuels' `
        -Columns $columns `
        -Rows $rows
}

Function GenerateSynchronizationTable {
    param([ReportDifference[]] $differences)

    $syncDifferences = $differences | Where-Object { $_.ElementType -eq 'Synchronisation' }
    $columns = Get-DefaultComparisonColumns `
        -NameHeaderKey 'tables.common.headers.name' -NameHeaderText 'Nom' `
        -TypeHeaderKey 'tables.common.headers.type' -TypeHeaderText 'Type' `
        -ValueHeaderKey 'tables.common.headers.value' -ValueHeaderText 'Valeur / Details'

    $rows = @()

    foreach ($diff in $syncDifferences) {
        $displayPair = Get-DisplayNamePair -DisplayName $diff.ElementDisplayName -TechnicalName $diff.ElementName -ElementType $diff.ElementType
        $propertyPair = Get-SafeTextPair -Text $diff.PropertyName

        $pageRaw = if ($diff.ParentDisplayName) { $diff.ParentDisplayName } elseif ($diff.ParentElementName) { $diff.ParentElementName } else { '' }
        $pagePair = Get-SafeTextPair -Text $pageRaw -DefaultFr 'Page inconnue' -DefaultEn 'Unknown page'

        $oldValuePair = Get-DiffValuePair -Value $diff.OldValue
        $newValuePair = Get-DiffValuePair -Value $diff.NewValue

        $additionalFr = $null
        $additionalEn = $null
        if ($diff.AdditionalInfo) {
            $additionalFr = Sanitize-AdditionalInfoForEndUser -Text $diff.AdditionalInfo
            $additionalEnRaw = Convert-AdditionalInfoToEnglish $diff.AdditionalInfo
            $additionalEn = if ([string]::IsNullOrWhiteSpace($additionalEnRaw)) { $additionalFr } else { Sanitize-AdditionalInfoForEndUser -Text $additionalEnRaw }
            if ($additionalFr -eq '-' -or [string]::IsNullOrWhiteSpace($additionalFr)) { $additionalFr = $null }
            if ($additionalEn -eq '-' -or [string]::IsNullOrWhiteSpace($additionalEn)) { $additionalEn = $null }
        }

        $descriptionPair = Get-UserFriendlyDescription -diff $diff

        $nameTagsFr = @()
        $nameTagsEn = @()
        if (-not [string]::IsNullOrWhiteSpace($pagePair.fr) -and $pagePair.fr -ne '-') {
            $nameTagsFr += "Page : $($pagePair.fr)"
            $nameTagsEn += "Page: $($pagePair.en)"
        }

        $oldNotesFr = @()
        $oldNotesEn = @()
        $newNotesFr = @()
        $newNotesEn = @()

        if ($additionalFr) {
            switch ($diff.DifferenceType) {
                'Ajoute' {
                    $newNotesFr += $additionalFr
                    $newNotesEn += $additionalEn
                }
                'Supprime' {
                    $oldNotesFr += $additionalFr
                    $oldNotesEn += $additionalEn
                }
                default {
                    $oldNotesFr += $additionalFr
                    $oldNotesEn += $additionalEn
                    $newNotesFr += $additionalFr
                    $newNotesEn += $additionalEn
                }
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($descriptionPair.fr)) {
            $newNotesFr += $descriptionPair.fr
            $newNotesEn += $descriptionPair.en
        }

        $rows += [ordered]@{
            CssClass = Get-DifferenceCssClass $diff.DifferenceType
            Status = $diff.DifferenceType
            Old = @{
                name  = New-RichCellContent -PrimaryFr $displayPair.fr -PrimaryEn $displayPair.en -TagsFr $nameTagsFr -TagsEn $nameTagsEn
                type  = New-RichCellContent -PrimaryFr $propertyPair.fr -PrimaryEn $propertyPair.en
                value = New-RichCellContent -PrimaryFr $oldValuePair.fr -PrimaryEn $oldValuePair.en -NotesFr $oldNotesFr -NotesEn $oldNotesEn
            }
            New = @{
                name  = New-RichCellContent -PrimaryFr $displayPair.fr -PrimaryEn $displayPair.en -TagsFr $nameTagsFr -TagsEn $nameTagsEn
                type  = New-RichCellContent -PrimaryFr $propertyPair.fr -PrimaryEn $propertyPair.en
                value = New-RichCellContent -PrimaryFr $newValuePair.fr -PrimaryEn $newValuePair.en -NotesFr $newNotesFr -NotesEn $newNotesEn
            }
        }
    }

    return Build-TwoBlockTable `
        -TableId 'table_synchronization' `
        -TitleKey 'tables.table_synchronization.title' `
        -TitleText 'Modifications de synchronisation des slicers' `
        -EmptyKey 'tables.table_synchronization.empty' `
        -EmptyText 'Aucune modification de synchronisation detectee' `
        -Columns $columns `
        -Rows $rows
}

Function GenerateVisualFieldsTable {
    param([ReportDifference[]] $differences)

    $fieldDifferences = $differences | Where-Object { $_.ElementType -eq 'Visuel' -and $_.PropertyName -eq 'fields' }
    $columns = Get-DefaultComparisonColumns `
        -NameHeaderKey 'tables.common.headers.name' -NameHeaderText 'Nom' `
        -TypeHeaderKey 'tables.common.headers.type' -TypeHeaderText 'Type' `
        -ValueHeaderKey 'tables.common.headers.value' -ValueHeaderText 'Valeur / Details'

    $rows = @()

    foreach ($diff in $fieldDifferences) {
        $displayPair = Get-DisplayNamePair -DisplayName $diff.ElementDisplayName -TechnicalName $diff.ElementName -ElementType $diff.ElementType
        $pageRaw = if ($diff.ParentDisplayName) { $diff.ParentDisplayName } elseif ($diff.ParentElementName) { $diff.ParentElementName } else { '' }
        $pagePair = Get-SafeTextPair -Text $pageRaw -DefaultFr 'Page inconnue' -DefaultEn 'Unknown page'

        $oldValuePair = Get-DiffValuePair -Value $diff.OldValue
        $newValuePair = Get-DiffValuePair -Value $diff.NewValue

        $additionalFr = $null
        $additionalEn = $null
        if ($diff.AdditionalInfo) {
            $additionalFr = Sanitize-AdditionalInfoForEndUser -Text $diff.AdditionalInfo
            $additionalEnRaw = Convert-AdditionalInfoToEnglish $diff.AdditionalInfo
            $additionalEn = if ([string]::IsNullOrWhiteSpace($additionalEnRaw)) { $additionalFr } else { Sanitize-AdditionalInfoForEndUser -Text $additionalEnRaw }
            if ($additionalFr -eq '-' -or [string]::IsNullOrWhiteSpace($additionalFr)) { $additionalFr = $null }
            if ($additionalEn -eq '-' -or [string]::IsNullOrWhiteSpace($additionalEn)) { $additionalEn = $null }
        }

        $descriptionPair = Get-UserFriendlyDescription -diff $diff

        $nameTagsFr = @()
        $nameTagsEn = @()
        if (-not [string]::IsNullOrWhiteSpace($pagePair.fr) -and $pagePair.fr -ne '-') {
            $nameTagsFr += "Page : $($pagePair.fr)"
            $nameTagsEn += "Page: $($pagePair.en)"
        }

        $oldNotesFr = @()
        $oldNotesEn = @()
        $newNotesFr = @()
        $newNotesEn = @()

        if ($additionalFr) {
            switch ($diff.DifferenceType) {
                'Ajoute' {
                    $newNotesFr += $additionalFr
                    $newNotesEn += $additionalEn
                }
                'Supprime' {
                    $oldNotesFr += $additionalFr
                    $oldNotesEn += $additionalEn
                }
                default {
                    $oldNotesFr += $additionalFr
                    $oldNotesEn += $additionalEn
                    $newNotesFr += $additionalFr
                    $newNotesEn += $additionalEn
                }
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($descriptionPair.fr)) {
            $newNotesFr += $descriptionPair.fr
            $newNotesEn += $descriptionPair.en
        }

        $rows += [ordered]@{
            CssClass = Get-DifferenceCssClass $diff.DifferenceType
            Status = $diff.DifferenceType
            Old = @{
                name  = New-RichCellContent -PrimaryFr $displayPair.fr -PrimaryEn $displayPair.en -TagsFr $nameTagsFr -TagsEn $nameTagsEn
                type  = New-RichCellContent -PrimaryFr 'Champs' -PrimaryEn 'Fields'
                value = New-RichCellContent -PrimaryFr $oldValuePair.fr -PrimaryEn $oldValuePair.en -NotesFr $oldNotesFr -NotesEn $oldNotesEn
            }
            New = @{
                name  = New-RichCellContent -PrimaryFr $displayPair.fr -PrimaryEn $displayPair.en -TagsFr $nameTagsFr -TagsEn $nameTagsEn
                type  = New-RichCellContent -PrimaryFr 'Champs' -PrimaryEn 'Fields'
                value = New-RichCellContent -PrimaryFr $newValuePair.fr -PrimaryEn $newValuePair.en -NotesFr $newNotesFr -NotesEn $newNotesEn
            }
        }
    }

    return Build-TwoBlockTable `
        -TableId 'table_visual_fields' `
        -TitleKey 'tables.table_visual_fields.title' `
        -TitleText 'Modifications des champs de visuels' `
        -EmptyKey 'tables.table_visual_fields.empty' `
        -EmptyText 'Aucun changement detecte sur les champs de visuels' `
        -Columns $columns `
        -Rows $rows
}

# OLD COMPLEX SYNC TABLE REMOVED - Now using Build-TwoBlockTable like other tables

Function GenerateButtonsTable {
    param([ReportDifference[]] $differences)

    $buttonDifferences = $differences | Where-Object ElementType -eq 'Bouton'
    $columns = Get-DefaultComparisonColumns `
        -NameHeaderKey 'tables.common.headers.name' -NameHeaderText 'Nom' `
        -TypeHeaderKey 'tables.table_buttons.headers.property' -TypeHeaderText 'Propriete' `
        -ValueHeaderKey 'tables.common.headers.value' -ValueHeaderText 'Valeur / Details'

    $rows = @()

    foreach ($diff in $buttonDifferences) {
        $displayPair = Get-DisplayNamePair -DisplayName $diff.ElementDisplayName -TechnicalName $diff.ElementName -ElementType $diff.ElementType
        $propertyPair = Get-SafeTextPair -Text $diff.PropertyName

        $pageRaw = if ($diff.ParentDisplayName) { $diff.ParentDisplayName } elseif ($diff.ParentElementName) { $diff.ParentElementName } else { '' }
        $pagePair = Get-SafeTextPair -Text $pageRaw -DefaultFr 'Page inconnue' -DefaultEn 'Unknown page'

        $oldValuePair = Get-DiffValuePair -Value $diff.OldValue
        $newValuePair = Get-DiffValuePair -Value $diff.NewValue

        $additionalFr = $null
        $additionalEn = $null
        if ($diff.AdditionalInfo) {
            $additionalFr = Sanitize-AdditionalInfoForEndUser -Text $diff.AdditionalInfo
            $additionalEnRaw = Convert-AdditionalInfoToEnglish $diff.AdditionalInfo
            $additionalEn = if ([string]::IsNullOrWhiteSpace($additionalEnRaw)) { $additionalFr } else { Sanitize-AdditionalInfoForEndUser -Text $additionalEnRaw }
            if ($additionalFr -eq '-' -or [string]::IsNullOrWhiteSpace($additionalFr)) { $additionalFr = $null }
            if ($additionalEn -eq '-' -or [string]::IsNullOrWhiteSpace($additionalEn)) { $additionalEn = $null }
        }

        $descriptionPair = Get-UserFriendlyDescription -diff $diff

        $nameTagsFr = @()
        $nameTagsEn = @()
        if (-not [string]::IsNullOrWhiteSpace($pagePair.fr) -and $pagePair.fr -ne '-') {
            $nameTagsFr += "Page : $($pagePair.fr)"
            $nameTagsEn += "Page: $($pagePair.en)"
        }

        $oldNotesFr = @()
        $oldNotesEn = @()
        $newNotesFr = @()
        $newNotesEn = @()

        if ($additionalFr) {
            switch ($diff.DifferenceType) {
                'Ajoute' {
                    $newNotesFr += $additionalFr
                    $newNotesEn += $additionalEn
                }
                'Supprime' {
                    $oldNotesFr += $additionalFr
                    $oldNotesEn += $additionalEn
                }
                default {
                    $oldNotesFr += $additionalFr
                    $oldNotesEn += $additionalEn
                    $newNotesFr += $additionalFr
                    $newNotesEn += $additionalEn
                }
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($descriptionPair.fr)) {
            $newNotesFr += $descriptionPair.fr
            $newNotesEn += $descriptionPair.en
        }

        $rows += [ordered]@{
            CssClass = Get-DifferenceCssClass $diff.DifferenceType
            Status = $diff.DifferenceType
            Old = @{
                name  = New-RichCellContent -PrimaryFr $displayPair.fr -PrimaryEn $displayPair.en -TagsFr $nameTagsFr -TagsEn $nameTagsEn
                type  = New-RichCellContent -PrimaryFr $propertyPair.fr -PrimaryEn $propertyPair.en
                value = New-RichCellContent -PrimaryFr $oldValuePair.fr -PrimaryEn $oldValuePair.en -NotesFr $oldNotesFr -NotesEn $oldNotesEn
            }
            New = @{
                name  = New-RichCellContent -PrimaryFr $displayPair.fr -PrimaryEn $displayPair.en -TagsFr $nameTagsFr -TagsEn $nameTagsEn
                type  = New-RichCellContent -PrimaryFr $propertyPair.fr -PrimaryEn $propertyPair.en
                value = New-RichCellContent -PrimaryFr $newValuePair.fr -PrimaryEn $newValuePair.en -NotesFr $newNotesFr -NotesEn $newNotesEn
            }
        }
    }

    return Build-TwoBlockTable `
        -TableId 'table_buttons' `
        -TitleKey 'tables.table_buttons.title' `
        -TitleText 'Modifications des boutons (actionButton)' `
        -EmptyKey 'tables.table_buttons.empty' `
        -EmptyText 'Aucun changement detecte sur les boutons' `
        -Columns $columns `
        -Rows $rows
}

Function GenerateBookmarksTable {
    param([ReportDifference[]] $differences)
    
    $bookmarkDifferences = $differences | Where-Object ElementType -eq 'Signet'
    $columns = Get-DefaultComparisonColumns `
        -NameHeaderKey 'tables.common.headers.name' -NameHeaderText 'Nom' `
        -TypeHeaderKey 'tables.table_bookmarks.headers.property' -TypeHeaderText 'Propriete' `
        -ValueHeaderKey 'tables.common.headers.value' -ValueHeaderText 'Valeur / Details'

    $rows = @()

    foreach ($diff in $bookmarkDifferences) {
        $displayPair = Get-DisplayNamePair -DisplayName $diff.ElementDisplayName -TechnicalName $diff.ElementName -ElementType $diff.ElementType
        $propertyPair = Get-SafeTextPair -Text $diff.PropertyName

        $oldValuePair = Get-DiffValuePair -Value $diff.OldValue
        $newValuePair = Get-DiffValuePair -Value $diff.NewValue

        $pathPair = $null
        if ($diff.ElementPath) {
            $cleanPath = Remove-TechnicalMarkers -Text $diff.ElementPath
            $pathPair = @{ fr = $cleanPath; en = $cleanPath }
        }

        $descriptionPair = Get-UserFriendlyDescription -diff $diff

        $commonNotesFr = @()
        $commonNotesEn = @()
        if ($pathPair) {
            $commonNotesFr += "Chemin : $($pathPair.fr)"
            $commonNotesEn += "Path: $($pathPair.en)"
        }

        $oldNotesFr = @()
        $oldNotesEn = @()
        $newNotesFr = @()
        $newNotesEn = @()

        if ($commonNotesFr.Count -gt 0) {
            switch ($diff.DifferenceType) {
                'Ajoute' {
                    $newNotesFr += $commonNotesFr
                    $newNotesEn += $commonNotesEn
                }
                'Supprime' {
                    $oldNotesFr += $commonNotesFr
                    $oldNotesEn += $commonNotesEn
                }
                default {
                    $oldNotesFr += $commonNotesFr
                    $oldNotesEn += $commonNotesEn
                    $newNotesFr += $commonNotesFr
                    $newNotesEn += $commonNotesEn
                }
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($descriptionPair.fr)) {
            $newNotesFr += $descriptionPair.fr
            $newNotesEn += $descriptionPair.en
        }

        $rows += [ordered]@{
            CssClass = Get-DifferenceCssClass $diff.DifferenceType
            Status = $diff.DifferenceType
            Old = @{
                name  = New-RichCellContent -PrimaryFr $displayPair.fr -PrimaryEn $displayPair.en
                type  = New-RichCellContent -PrimaryFr $propertyPair.fr -PrimaryEn $propertyPair.en
                value = New-RichCellContent -PrimaryFr $oldValuePair.fr -PrimaryEn $oldValuePair.en -NotesFr $oldNotesFr -NotesEn $oldNotesEn
            }
            New = @{
                name  = New-RichCellContent -PrimaryFr $displayPair.fr -PrimaryEn $displayPair.en
                type  = New-RichCellContent -PrimaryFr $propertyPair.fr -PrimaryEn $propertyPair.en
                value = New-RichCellContent -PrimaryFr $newValuePair.fr -PrimaryEn $newValuePair.en -NotesFr $newNotesFr -NotesEn $newNotesEn
            }
        }
    }
    
    return Build-TwoBlockTable `
        -TableId 'table_bookmarks' `
        -TitleKey 'tables.table_bookmarks.title' `
        -TitleText 'Modifications des signets' `
        -EmptyKey 'tables.table_bookmarks.empty' `
        -EmptyText 'Aucune modification detectee dans les signets' `
        -Columns $columns `
        -Rows $rows
}

# Helper function to categorize configuration changes
Function Get-ConfigurationCategory {
    param([string]$propertyName)

    switch -Wildcard ($propertyName) {
        "displayOption" { return @{fr="Affichage"; en="Page View"} }
        "canvasType" { return @{fr="Format du canevas"; en="Canvas Format"} }
        "verticalAlignment" { return @{fr="Alignement"; en="Alignment"} }
        "metadata.*" { return @{fr="Métadonnées"; en="Metadata"} }
        "datasetReference.*" { return @{fr="Références dataset"; en="Dataset References"} }
        default { return @{fr="Configuration"; en="Configuration"} }
    }
}

# Helper function to extract page name from ElementDisplayName
Function Get-PageNameFromElement {
    param([string]$elementDisplayName)

    # Extract page name from format: 'PageName' (pageId)
    if ($elementDisplayName -match "^'([^']+)'") {
        return $Matches[1]
    }
    elseif ($elementDisplayName -match "^([^\(]+)\s*\(") {
        return $Matches[1].Trim()
    }
    else {
        return $elementDisplayName
    }
}

# Helper function to get simplified property name
Function Get-SimplifiedPropertyName {
    param([string]$propertyName)

    $translations = @{
        "displayOption" = @{fr="Mode d'affichage"; en="Display Mode"}
        "canvasType" = @{fr="Type de canevas"; en="Canvas Type"}
        "verticalAlignment" = @{fr="Alignement vertical"; en="Vertical Alignment"}
        "metadata.displayName" = @{fr="Nom du projet"; en="Project Name"}
        "datasetReference.byPath.path" = @{fr="Chemin du dataset"; en="Dataset Path"}
    }

    if ($translations.ContainsKey($propertyName)) {
        return $translations[$propertyName]
    }
    else {
        return @{fr=$propertyName; en=$propertyName}
    }
}

Function GenerateConfigTable {
    param([ReportDifference[]] $differences)

    # Exclude displayOption as it's now handled in Pages table
    $configDifferences = $differences | Where-Object { $_.ElementType -eq 'Configuration' -and $_.PropertyName -ne 'displayOption' }
    if (-not $configDifferences) {
        $configDifferences = @()
    }

    $columns = Get-DefaultComparisonColumns `
        -NameHeaderKey 'tables.table_config.headers.page' -NameHeaderText 'Page' `
        -TypeHeaderKey 'tables.table_config.headers.property' -TypeHeaderText 'Parametre' `
        -ValueHeaderKey 'tables.common.headers.value' -ValueHeaderText 'Valeur / Details'

    $rows = @()

    foreach ($diff in $configDifferences) {
        $pageName = Get-PageNameFromElement -elementDisplayName $diff.ElementDisplayName
        if ([string]::IsNullOrEmpty($pageName) -or $pageName -eq $diff.ElementName) {
            if ($diff.ElementName -match '\.platform|definition\.pbir') {
                $pagePair = @{ fr = 'Configuration système'; en = 'System configuration' }
            }
            else {
                $pagePair = Get-SafeTextPair -Text $diff.ElementName -DefaultFr 'Configuration du rapport' -DefaultEn 'Report configuration'
            }
        }
        else {
            $pagePair = Get-SafeTextPair -Text $pageName -DefaultFr 'Configuration du rapport' -DefaultEn 'Report configuration'
        }

        if (-not $pagePair) {
            $pagePair = @{ fr = 'Configuration du rapport'; en = 'Report configuration' }
        }

        $categoryPair = Get-ConfigurationCategory -propertyName $diff.PropertyName
        $propertyPair = Get-SimplifiedPropertyName -propertyName $diff.PropertyName
        if (-not $propertyPair) {
            $propertyPair = Get-SafeTextPair -Text $diff.PropertyName
        }

        $oldValuePair = Get-DiffValuePair -Value $diff.OldValue
        $newValuePair = Get-DiffValuePair -Value $diff.NewValue

        switch ($diff.PropertyName) {
            'displayOption' {
                $mappingFr = @{ FitToPage = "Ajuster à la page"; FitToWidth = "Ajuster à la largeur"; ActualSize = "Taille réelle" }
                $mappingEn = @{ FitToPage = "Fit to Page"; FitToWidth = "Fit to Width"; ActualSize = "Actual Size" }
                if ($mappingFr.ContainsKey($diff.OldValue)) { $oldValuePair.fr = $mappingFr[$diff.OldValue] }
                if ($mappingEn.ContainsKey($diff.OldValue)) { $oldValuePair.en = $mappingEn[$diff.OldValue] }
                if ($mappingFr.ContainsKey($diff.NewValue)) { $newValuePair.fr = $mappingFr[$diff.NewValue] }
                if ($mappingEn.ContainsKey($diff.NewValue)) { $newValuePair.en = $mappingEn[$diff.NewValue] }
            }
            'verticalAlignment' {
                $mappingFr = @{ Top = 'Haut'; Middle = 'Centre'; Bottom = 'Bas' }
                $mappingEn = @{ Top = 'Top'; Middle = 'Middle'; Bottom = 'Bottom' }
                if ($mappingFr.ContainsKey($diff.OldValue)) { $oldValuePair.fr = $mappingFr[$diff.OldValue] }
                if ($mappingEn.ContainsKey($diff.OldValue)) { $oldValuePair.en = $mappingEn[$diff.OldValue] }
                if ($mappingFr.ContainsKey($diff.NewValue)) { $newValuePair.fr = $mappingFr[$diff.NewValue] }
                if ($mappingEn.ContainsKey($diff.NewValue)) { $newValuePair.en = $mappingEn[$diff.NewValue] }
            }
        }

        if ($diff.PropertyName -eq 'canvasType' -and -not [string]::IsNullOrWhiteSpace($diff.NewValue)) {
            $oldValuePair.en = Convert-ValueToEnglish $diff.OldValue
            if ([string]::IsNullOrWhiteSpace($oldValuePair.en)) { $oldValuePair.en = $oldValuePair.fr }
            $newValuePair.en = Convert-ValueToEnglish $diff.NewValue
            if ([string]::IsNullOrWhiteSpace($newValuePair.en)) { $newValuePair.en = $newValuePair.fr }
        }

        $additionalFr = $null
        $additionalEn = $null
        if ($diff.AdditionalInfo) {
            $additionalFr = Sanitize-AdditionalInfoForEndUser -Text $diff.AdditionalInfo
            $additionalEnRaw = Convert-AdditionalInfoToEnglish $diff.AdditionalInfo
            $additionalEn = if ([string]::IsNullOrWhiteSpace($additionalEnRaw)) { $additionalFr } else { Sanitize-AdditionalInfoForEndUser -Text $additionalEnRaw }
            if ($additionalFr -eq '-' -or [string]::IsNullOrWhiteSpace($additionalFr)) { $additionalFr = $null }
            if ($additionalEn -eq '-' -or [string]::IsNullOrWhiteSpace($additionalEn)) { $additionalEn = $null }
        }

        $descriptionPair = Get-UserFriendlyDescription -diff $diff

        $nameTagsFr = @()
        $nameTagsEn = @()
        if ($categoryPair -and -not [string]::IsNullOrWhiteSpace($categoryPair.fr)) {
            $nameTagsFr += "Catégorie : $($categoryPair.fr)"
            $nameTagsEn += "Category: $($categoryPair.en)"
        }

        $oldNotesFr = @()
        $oldNotesEn = @()
        $newNotesFr = @()
        $newNotesEn = @()

        if ($additionalFr) {
            switch ($diff.DifferenceType) {
                'Ajoute' {
                    $newNotesFr += $additionalFr
                    $newNotesEn += $additionalEn
                }
                'Supprime' {
                    $oldNotesFr += $additionalFr
                    $oldNotesEn += $additionalEn
                }
                default {
                    $oldNotesFr += $additionalFr
                    $oldNotesEn += $additionalEn
                    $newNotesFr += $additionalFr
                    $newNotesEn += $additionalEn
                }
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($descriptionPair.fr)) {
            $newNotesFr += $descriptionPair.fr
            $newNotesEn += $descriptionPair.en
        }

        $rows += [ordered]@{
            CssClass = Get-DifferenceCssClass $diff.DifferenceType
            Status = $diff.DifferenceType
            Old = @{
                name  = New-RichCellContent -PrimaryFr $pagePair.fr -PrimaryEn $pagePair.en -TagsFr $nameTagsFr -TagsEn $nameTagsEn
                type  = New-RichCellContent -PrimaryFr $propertyPair.fr -PrimaryEn $propertyPair.en
                value = New-RichCellContent -PrimaryFr $oldValuePair.fr -PrimaryEn $oldValuePair.en -NotesFr $oldNotesFr -NotesEn $oldNotesEn
            }
            New = @{
                name  = New-RichCellContent -PrimaryFr $pagePair.fr -PrimaryEn $pagePair.en -TagsFr $nameTagsFr -TagsEn $nameTagsEn
                type  = New-RichCellContent -PrimaryFr $propertyPair.fr -PrimaryEn $propertyPair.en
                value = New-RichCellContent -PrimaryFr $newValuePair.fr -PrimaryEn $newValuePair.en -NotesFr $newNotesFr -NotesEn $newNotesEn
            }
        }
    }

    return Build-TwoBlockTable `
        -TableId 'table_config' `
        -TitleKey 'tables.table_config.title' `
        -TitleText 'Modifications de Configuration' `
        -EmptyKey 'tables.table_config.empty' `
        -EmptyText 'Aucune modification detectee dans la configuration' `
        -Columns $columns `
        -Rows $rows
}

Function GenerateVisualInteractionsTable {
    param([ReportDifference[]] $differences)

    $interactionDifferences = $differences | Where-Object ElementType -eq 'VisualInteraction'

    $columns = @(
        @{ Key = 'page'; OldHeaderKey = 'tables.table_visual_interactions.headers.page'; OldHeaderText = 'Page'; NewHeaderKey = 'tables.table_visual_interactions.headers.page'; NewHeaderText = 'Page'; OldAttributes = "class='cell-name'"; NewAttributes = "class='cell-name'" },
        @{ Key = 'source'; OldHeaderKey = 'tables.table_visual_interactions.headers.source'; OldHeaderText = 'Visuel Source'; NewHeaderKey = 'tables.table_visual_interactions.headers.source'; NewHeaderText = 'Visuel Source'; OldAttributes = "class='cell-type'"; NewAttributes = "class='cell-type'" },
        @{ Key = 'target'; OldHeaderKey = 'tables.table_visual_interactions.headers.target'; OldHeaderText = 'Visuel Cible'; NewHeaderKey = 'tables.table_visual_interactions.headers.target'; NewHeaderText = 'Visuel Cible'; OldAttributes = "class='cell-type'"; NewAttributes = "class='cell-type'" },
        @{ Key = 'interaction'; OldHeaderKey = 'tables.table_visual_interactions.headers.before'; OldHeaderText = 'Type Interaction'; NewHeaderKey = 'tables.table_visual_interactions.headers.after'; NewHeaderText = 'Type Interaction'; OldAttributes = "class='cell-value'"; NewAttributes = "class='cell-value'" }
    )

    $rows = @()
    foreach ($diff in $interactionDifferences) {
        # Extract page name
        $pageName = if ($diff.ParentDisplayName) { $diff.ParentDisplayName } else { $diff.ParentElementName }

        # Extract source and target display names from ElementDisplayName (format: "source → target")
        $parts = $diff.ElementDisplayName -split '→'
        $sourceDisplay = if ($parts.Count -eq 2) { $parts[0].Trim() } else { "N/A" }
        $targetDisplay = if ($parts.Count -eq 2) { $parts[1].Trim() } else { "N/A" }

        $sourcePair = Get-DisplayNamePair -DisplayName $sourceDisplay -TechnicalName $sourceDisplay -ElementType 'Visuel'
        $targetPair = Get-DisplayNamePair -DisplayName $targetDisplay -TechnicalName $targetDisplay -ElementType 'Visuel'
        $pagePair = Get-SafeTextPair -Text $pageName -DefaultFr "Page inconnue" -DefaultEn "Unknown page"

        # Map interaction types without emojis
        $oldInteractionPair = Get-DiffValuePair -Value $diff.OldValue -FallbackFr "-" -FallbackEn "-"
        $newInteractionPair = Get-DiffValuePair -Value $diff.NewValue -FallbackFr "-" -FallbackEn "-"

        $rows += @{
            Old = @{
                page = $pagePair
                source = $sourcePair
                target = $targetPair
                interaction = $oldInteractionPair
            }
            New = @{
                page = $pagePair
                source = $sourcePair
                target = $targetPair
                interaction = $newInteractionPair
            }
            Status = $diff.DifferenceType
            CssClass = Get-DifferenceCssClass $diff.DifferenceType
        }
    }

    return Build-TwoBlockTable `
        -TableId 'table_visual_interactions' `
        -TitleKey 'tables.table_visual_interactions.title' `
        -TitleText 'Interactions Visuelles' `
        -EmptyKey 'tables.table_visual_interactions.empty' `
        -EmptyText 'Aucune modification détectée dans les interactions visuelles' `
        -Columns $columns `
        -Rows $rows
}

Function GenerateThemesTable {
    param([ReportDifference[]] $differences)
    
    $themeDifferences = $differences | Where-Object ElementType -eq 'Theme'

    $columns = @(
        @{ Key = 'element'; OldHeaderKey = 'tables.table_themes.headers.element'; OldHeaderText = 'Élément de thème'; NewHeaderKey = 'tables.table_themes.headers.element'; NewHeaderText = 'Élément de thème'; OldAttributes = "class='cell-name'"; NewAttributes = "class='cell-name'" },
        @{ Key = 'property'; OldHeaderKey = 'tables.table_themes.headers.property'; OldHeaderText = 'Propriété'; NewHeaderKey = 'tables.table_themes.headers.property'; NewHeaderText = 'Propriété'; OldAttributes = "class='cell-type'"; NewAttributes = "class='cell-type'" },
        @{ Key = 'value'; OldHeaderKey = 'tables.table_themes.headers.oldValue'; OldHeaderText = 'Valeur'; NewHeaderKey = 'tables.table_themes.headers.newValue'; NewHeaderText = 'Valeur'; OldAttributes = "class='cell-value'"; NewAttributes = "class='cell-value'" }
    )

    $rows = @()
    foreach ($diff in $themeDifferences) {
        $elementPair = Get-DisplayNamePair -DisplayName $diff.ElementDisplayName -TechnicalName $diff.ElementName -ElementType $diff.ElementType
        $propertyPair = Get-SafeTextPair -Text $diff.PropertyName -DefaultFr $diff.PropertyName -DefaultEn $diff.PropertyName
        
        $oldValueFr = if ([string]::IsNullOrWhiteSpace($diff.OldValue)) { "-" } else { Remove-TechnicalMarkers -Text $diff.OldValue }
        $newValueFr = if ([string]::IsNullOrWhiteSpace($diff.NewValue)) { "-" } else { Remove-TechnicalMarkers -Text $diff.NewValue }
        $oldValueEn = Remove-TechnicalMarkers -Text (Convert-ValueToEnglish $diff.OldValue)
        if ([string]::IsNullOrWhiteSpace($oldValueEn)) { $oldValueEn = $oldValueFr }
        $newValueEn = Remove-TechnicalMarkers -Text (Convert-ValueToEnglish $diff.NewValue)
        if ([string]::IsNullOrWhiteSpace($newValueEn)) { $newValueEn = $newValueFr }

        $rows += @{
            Old = @{
                element = $elementPair
                property = $propertyPair
                value = @{ fr = $oldValueFr; en = $oldValueEn }
            }
            New = @{
                element = $elementPair
                property = $propertyPair
                value = @{ fr = $newValueFr; en = $newValueEn }
            }
            Status = $diff.DifferenceType
            CssClass = Get-DifferenceCssClass $diff.DifferenceType
        }
    }

    return Build-TwoBlockTable `
        -TableId 'table_themes' `
        -TitleKey 'tables.table_themes.title' `
        -TitleText 'Modifications des thèmes et styles' `
        -EmptyKey 'tables.table_themes.empty' `
        -EmptyText 'Aucune modification détectée dans les thèmes' `
        -Columns $columns `
        -Rows $rows
}

# Function to generate quality checks table for slicers
Function GenerateQualityChecksTable {
    param([PSCustomObject[]] $checkResults)
    
    $html = @"
<table id="table_quality_checks" class="responsive-table">
<tr class="table-title-row"><th colspan="9" data-i18n-key="tables.table_quality.title">Verification Qualite des Slicers</th></tr>
<tr class="table-header-row">
    <th data-i18n-key="tables.table_quality.headers.page">Page</th>
    <th data-i18n-key="tables.table_quality.headers.displayName">Nom du filtre</th>
    <th data-i18n-key="tables.table_quality.headers.fieldName">Champ</th>
    <th data-i18n-key="tables.table_quality.headers.hasSelected">Selection</th>
    <th data-i18n-key="tables.table_quality.headers.hasSearch">Recherche</th>
    <th data-i18n-key="tables.table_quality.headers.searchText">Texte recherche</th>
    <th data-i18n-key="tables.table_quality.headers.isRadio">Mode radio</th>
    <th data-i18n-key="tables.table_quality.headers.status">Status</th>
    <th data-i18n-key="tables.table_quality.headers.message">Message</th>
</tr>
"@

    if ($checkResults.Count -eq 0) {
        $html += '<tr class="ok"><td colspan="9" data-i18n-key="tables.table_quality.empty">Aucun slicer trouve pour verification</td></tr>'
    }
    else {
        foreach ($result in $checkResults) {
            $cssClass = switch ($result.Status) {
                "OK" { "ok" }
                "ALERTE" { "alerte" }
                "ERREUR" { "alerte" }
                default { "" }
            }
            
            $hasSelectedIcon = if ($result.HasSelected) { "✓" } else { "✗" }
            $hasSearchIcon = if ($result.HasSearch) { "✓" } else { "✗" }
            $isRadioIcon = if ($result.IsRadio) { "🔘" } else { "☐" }
            $searchTextDisplay = if ($result.SearchText) {
                (( $result.SearchText -split ',' ) | ForEach-Object { Remove-TechnicalMarkers -Text ($_.Trim("'").Trim()) }) -join ', '
            } else {
                "-"
            }
            
            $messageKeyAttr = ''
            if ($result.PSObject.Properties['MessageKey']) {
                $keyValue = [string]$result.MessageKey
                if (-not [string]::IsNullOrWhiteSpace($keyValue)) {
                    $messageDetailValue = if ($result.PSObject.Properties['MessageDetail']) { [string]$result.MessageDetail } else { '' }
                    $messageKeyAttr = " data-i18n-message=`"$(Encode-HtmlAttribute $keyValue)`" data-i18n-message-detail=`"$(Encode-HtmlAttribute $messageDetailValue)`""
                }
            }

            $statusAttr = ''
            $statusValue = if ($result.PSObject.Properties['Status']) { [string]$result.Status } else { '' }
            if (-not [string]::IsNullOrWhiteSpace($statusValue)) {
                $statusAttr = " data-i18n-status=`"$(Encode-HtmlAttribute $statusValue)`""
            }

            $pageText = Remove-TechnicalMarkers -Text $result.PageName
            if ([string]::IsNullOrWhiteSpace($pageText)) { $pageText = "-" }
            $filterText = Remove-TechnicalMarkers -Text $result.DisplayName
            if ([string]::IsNullOrWhiteSpace($filterText)) { $filterText = "-" }
            $fieldText = Remove-TechnicalMarkers -Text $result.FieldName
            if ([string]::IsNullOrWhiteSpace($fieldText)) { $fieldText = "-" }
            $searchText = if ([string]::IsNullOrWhiteSpace($searchTextDisplay)) { "-" } else { $searchTextDisplay }
            $statusLabel = if ([string]::IsNullOrWhiteSpace($statusValue)) { "-" } else { $statusValue }
            $messageText = if ($result.Message) { $result.Message } else { "-" }

            $html += @"
<tr class="$cssClass">
    $(Build-LocalizedCell $pageText $pageText 'style="max-width:160px;word-wrap:break-word;font-weight:bold;"')
    $(Build-LocalizedCell $filterText $filterText 'style="max-width:180px;word-wrap:break-word;font-weight:bold;color:var(--neutral-700);"')
    $(Build-LocalizedCell $fieldText $fieldText 'style="max-width:140px;word-wrap:break-word;"')
    <td style="text-align:center;font-size:16px;">$hasSelectedIcon</td>
    <td style="text-align:center;font-size:16px;">$hasSearchIcon</td>
    $(Build-LocalizedCell $searchText $searchText 'style="max-width:140px;word-wrap:break-word;font-style:italic;"')
    <td style="text-align:center;font-size:16px;">$isRadioIcon</td>
    <td$statusAttr style="text-align:center;font-weight:bold;">$(Encode-HtmlContent $statusLabel)</td>
    <td$messageKeyAttr style="max-width:320px;word-wrap:break-word;">$(Encode-HtmlContent $messageText)</td>
</tr>
"@
        }
    }
    
    $html += "</table>`n"
    return $html
}
