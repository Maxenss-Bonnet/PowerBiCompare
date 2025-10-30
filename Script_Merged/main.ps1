# Hardcoded paths
$NewVersionFolderPath = "C:\Users\BQTR7546\.vscodeProject\nrt-powerbi\AdventureWorks2019 XLS après"
$OldVersionFolderPath = "C:\Users\BQTR7546\.vscodeProject\nrt-powerbi\AdventureWorks2019 XLS"
$BaseTargetFolder = "C:\Users\BQTR7546\.vscodeProject\nrt-powerbi"
$ReportOutputFolder = Join-Path $BaseTargetFolder "RapportHTML"

# Create the output directory if it doesn't exist
if (-not (Test-Path $ReportOutputFolder)) {
    New-Item -ItemType Directory -Path $ReportOutputFolder | Out-Null
}

# Get script path
$scriptRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
. (Join-Path $scriptRoot 'PBI_Classes.ps1')

# 1. Load DLLs
Write-Host "--- Chargement des DLLs ---" -ForegroundColor Green
. (Join-Path $scriptRoot 'PBI_load_dll.ps1')
LoadNeededDLL -path $scriptRoot

# 2. Load Class Definitions and Functions
Write-Host "--- Chargement des définitions et fonctions ---" -ForegroundColor Green
. (Join-Path $scriptRoot 'PBI_Report_HTML_Orange.ps1')
. (Join-Path $scriptRoot 'PBI_Report_Check.ps1')
. (Join-Path $scriptRoot 'PBI_Report_Compare.ps1')
. (Join-Path $scriptRoot 'PBI_MDD_extract.ps1')

# 3. Semantic Analysis
Write-Host "`n=== Analyse Sémantique ===" -ForegroundColor Green

try {
    # Suppress verbose output from semantic analysis functions
    $projetArray = LoadProjectVersionsPath -newVersionProjectRepertory $NewVersionFolderPath -oldVersionProjectRepertory $OldVersionFolderPath 2>$null
    $parsedArray = [PSCustomObject]@{
        ElementNewVersion = $projetArray[0].dataBase.Model
        ElementOldVersion = $projetArray[1].dataBase.Model
        PathNewVersion = 'DataBase'
        PathOldVersion = 'DataBase'
    }
    $semanticComparisonResult = CheckDifferenceInSubElement -element $parsedArray -elementChecked 'Model' 2>$null
    
    if ($semanticComparisonResult -and $semanticComparisonResult.Count -gt 0) {
        Write-Host "   Analyse semantique terminee: $($semanticComparisonResult.Count) differences detectees" -ForegroundColor Green
        Write-Host "  Type de donnees: $($semanticComparisonResult.GetType().Name)" -ForegroundColor Cyan
    } else {
        Write-Host "  ! Aucune difference semantique detectee" -ForegroundColor Yellow
    }
} catch {
    Write-Host "   ERREUR lors de l'analyse semantique: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "  Details: $($_.ScriptStackTrace)" -ForegroundColor Red
    $semanticComparisonResult = $null
}


# 4. Report Analysis
Write-Host "`n=== Analyse du Rapport ===" -ForegroundColor Green

try {
    # Load the Report projects
    $reportProjects = LoadReportProjectVersions -newVersionProjectDirectory $NewVersionFolderPath -oldVersionProjectDirectory $OldVersionFolderPath
    
    if ($reportProjects.Count -ne 2) {
        throw "Erreur lors du chargement des projets Report. Nombre de projets charges: $($reportProjects.Count)"
    }
    
    Write-Host "Projets Report charges avec succes" -ForegroundColor Green
} catch {
    Write-Host "ERREUR lors du chargement des projets Report: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Details: $($_.ScriptStackTrace)" -ForegroundColor Red
    throw
}

$newReportProject = $reportProjects[0]
$oldReportProject = $reportProjects[1]

# Compare Report projects
$reportDifferences = CompareReportProjects -newProject $newReportProject -oldProject $oldReportProject

Write-Host "  $($reportDifferences.Count) differences de rapport trouvees" -ForegroundColor White

# Quality check on visible slicers (new project only)
$visibleQuality = Invoke-VisibleSlicerQualityCheck -project $newReportProject -projectRoot $NewVersionFolderPath
$checkResults = @()

if ($visibleQuality -and $visibleQuality.PSObject.Properties['Results']) {
    $checkResults = $visibleQuality.Results
}

if ($checkResults.Count -gt 0) {
    $checkResults = Get-VisibleSlicerQualityResult -checkResults $checkResults -newProject $newReportProject
}

Write-Host "  $($checkResults.Count) slicers visibles analyses" -ForegroundColor White

# SIMPLIFIED SYNC: No more grouping or complex analysis
# Synchronization differences are now treated like any other differences

Write-Host "`n=== Generation des rapports HTML ===" -ForegroundColor Green
Write-Host "  Debug: semanticComparisonResult type = $($semanticComparisonResult.GetType().Name)" -ForegroundColor Cyan
Write-Host "  Debug: semanticComparisonResult count = $($semanticComparisonResult.Count)" -ForegroundColor Cyan

# Convert differences for HTML
$differencesForHtml = Convert-ReportDifferencesForHtml -differences $reportDifferences

# Generate ORANGE HTML report (same data, Orange branding)
Write-Host "`n--- Génération du rapport ORANGE ---" -ForegroundColor Magenta
$reportPathOrange = BuildReportHTMLReport_Orange -differences $differencesForHtml -checkResults $checkResults -outputFolder $ReportOutputFolder -semanticComparisonResult $semanticComparisonResult
Write-Host "   Rapport HTML Orange genere: $reportPathOrange" -ForegroundColor Magenta

Write-Host "`n=== Analyse terminée ===" -ForegroundColor Green
Write-Host "  → Rapport: $reportPathOrange" -ForegroundColor White
