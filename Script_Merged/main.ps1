# =============================================================================
# Interactive Folder Selection for Power BI Project Comparison
# =============================================================================

# Script parameters for non-interactive execution (e.g., compiled .exe)
param(
    [switch]$NonInteractive,
    [string]$NewVersionPath,
    [string]$OldVersionPath,
    [string]$OutputPath
)

# Load Windows Forms for file dialog (compatible with .exe compilation)
# Skip loading if running in non-interactive mode
# Determine script root directory
if ($PSScriptRoot) {
    $scriptRoot = $PSScriptRoot
} elseif ($MyInvocation.MyCommand.Path) {
    $scriptRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Path
} else {
    $scriptRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
}

# Only load Windows Forms if interactive mode (dialogs needed)
if (-not $NonInteractive) {
    try {
        # Try loading from GAC first (normal PowerShell execution)
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        Add-Type -AssemblyName System.Drawing -ErrorAction Stop
    } catch {
        # Fallback 1: Try loading from .NET Framework installation path
        Write-Host "⚠️  Loading Windows Forms from .NET Framework path..." -ForegroundColor Yellow
        
        try {
            $dotnetPath = "$env:SystemRoot\Microsoft.NET\Framework64\v4.0.30319"
            
            # Validate path before using it
            if (-not $dotnetPath -or -not (Test-Path $dotnetPath)) {
                throw ".NET Framework path not found or invalid"
            }
            
            $formsPath = Join-Path $dotnetPath "System.Windows.Forms.dll"
            $drawingPath = Join-Path $dotnetPath "System.Drawing.dll"
            
            if ((Test-Path $formsPath) -and (Test-Path $drawingPath)) {
                Add-Type -Path $formsPath -ErrorAction Stop
                Add-Type -Path $drawingPath -ErrorAction Stop
                Write-Host "✓ Windows Forms loaded from .NET Framework" -ForegroundColor Green
            } else {
                throw "DLLs not found in .NET Framework path"
            }
        } catch {
        # Fallback 2: Try loading from local lib folder (like other DLLs in PBI_load_dll.ps1)
        Write-Host "⚠️  Loading Windows Forms from local lib folder..." -ForegroundColor Yellow
        
        $libPath = Join-Path $scriptRoot "lib"
        $formsLibPath = Join-Path $libPath "System.Windows.Forms.dll"
        $drawingLibPath = Join-Path $libPath "System.Drawing.dll"
        
        if ((Test-Path $formsLibPath) -and (Test-Path $drawingLibPath)) {
            # Same validation as in PBI_load_dll.ps1
            foreach ($dllPath in @($formsLibPath, $drawingLibPath)) {
                # Check Authenticode signature
                $signature = Get-AuthenticodeSignature -FilePath $dllPath
                if ($signature.Status -ne 'Valid') {
                    Write-Warning "DLL signature not valid for $dllPath (status: $($signature.Status))"
                }
                
                # Unblock if necessary
                $zoneInfo = Get-Item -Path $dllPath -Stream Zone.Identifier -ErrorAction SilentlyContinue
                if ($null -ne $zoneInfo) {
                    Write-Host "DLL blocked by Windows, unblocking: $(Split-Path $dllPath -Leaf)" -ForegroundColor Yellow
                    Unblock-File -Path $dllPath
                }
                
                # Load the Assembly
                Add-Type -Path $dllPath
                Write-Host "✓ $(Split-Path $dllPath -Leaf) loaded from lib folder" -ForegroundColor Green
            }
        } else {
            throw "ERREUR: Impossible de charger System.Windows.Forms. Vérifiez que .NET Framework 4.x est installé ou que les DLLs sont présentes dans le dossier lib."
        }
    }
    }
} else {
    Write-Host "ℹ️  Mode non-interactif activé. Dialogs désactivés." -ForegroundColor Cyan
}

# =============================================================================
# LOADING WINDOW FUNCTIONS
# =============================================================================

# Variables globales pour la fenêtre de chargement
$script:progressForm = $null
$script:progressLabel = $null
$script:progressBar = $null
$script:progressPercent = $null

function Show-LoadingWindow {
    <#
    .SYNOPSIS
    Affiche une fenêtre de chargement professionnelle avec barre de progression
    #>
    
    # Créer le formulaire principal
    $form = New-Object System.Windows.Forms.Form
    $form.Width = 520
    $form.Height = 250  # Augmenté pour plus d'espace au-dessus du sous-titre
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'None'
    $form.TopMost = $true
    $form.BackColor = [System.Drawing.Color]::White
    
    # Panel avec bordure orange (3px)
    $borderPanel = New-Object System.Windows.Forms.Panel
    $borderPanel.Dock = 'Fill'
    $borderPanel.Padding = New-Object System.Windows.Forms.Padding(3)
    $borderPanel.BackColor = [System.Drawing.Color]::FromArgb(255, 127, 0) # Orange Business
    
    # Panel intérieur blanc
    $innerPanel = New-Object System.Windows.Forms.Panel
    $innerPanel.Dock = 'Fill'
    $innerPanel.BackColor = [System.Drawing.Color]::White
    
    # Logo Orange carré (imite le SVG du header)
    $logoPanel = New-Object System.Windows.Forms.Panel
    $logoPanel.Location = New-Object System.Drawing.Point(30, 18)
    $logoPanel.Size = New-Object System.Drawing.Size(30, 30)
    $logoPanel.BackColor = [System.Drawing.Color]::FromArgb(255, 127, 0) # Orange
    
    # Bande blanche en bas du logo (comme dans le SVG)
    $logoStripe = New-Object System.Windows.Forms.Panel
    $logoStripe.Location = New-Object System.Drawing.Point(2, 22)
    $logoStripe.Size = New-Object System.Drawing.Size(26, 6)
    $logoStripe.BackColor = [System.Drawing.Color]::White
    $logoPanel.Controls.Add($logoStripe)
    
    # Titre (décalé pour laisser place au logo)
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "Analyse en cours..."
    $titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
    $titleLabel.ForeColor = [System.Drawing.Color]::FromArgb(64, 64, 64)
    $titleLabel.Location = New-Object System.Drawing.Point(70, 18)
    $titleLabel.AutoSize = $true
    
    # Sous-titre Orange
    $subtitleLabel = New-Object System.Windows.Forms.Label
    $subtitleLabel.Text = "Comparateur Power BI - Orange Business Services"
    $subtitleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $subtitleLabel.ForeColor = [System.Drawing.Color]::FromArgb(255, 127, 0)
    $subtitleLabel.Location = New-Object System.Drawing.Point(70, 52)
    $subtitleLabel.AutoSize = $true
    
    # Label de statut (texte qui change)
    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Text = "Initialisation..."
    $statusLabel.Location = New-Object System.Drawing.Point(30, 100)
    $statusLabel.Size = New-Object System.Drawing.Size(460, 30)
    $statusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $statusLabel.ForeColor = [System.Drawing.Color]::FromArgb(64, 64, 64)
    
    # Barre de progression
    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location = New-Object System.Drawing.Point(30, 145)
    $progressBar.Size = New-Object System.Drawing.Size(460, 25)
    $progressBar.Style = 'Continuous'
    $progressBar.Minimum = 0
    $progressBar.Maximum = 100
    $progressBar.Value = 0
    
    # Label pourcentage
    $percentLabel = New-Object System.Windows.Forms.Label
    $percentLabel.Text = "0%"
    $percentLabel.Location = New-Object System.Drawing.Point(30, 180)
    $percentLabel.AutoSize = $true
    $percentLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $percentLabel.ForeColor = [System.Drawing.Color]::FromArgb(255, 127, 0)
    
    # Assembler les contrôles
    $innerPanel.Controls.AddRange(@($logoPanel, $titleLabel, $subtitleLabel, $statusLabel, $progressBar, $percentLabel))
    $borderPanel.Controls.Add($innerPanel)
    $form.Controls.Add($borderPanel)
    
    # Stocker dans variables globales
    $script:progressForm = $form
    $script:progressLabel = $statusLabel
    $script:progressBar = $progressBar
    $script:progressPercent = $percentLabel
    
    # Afficher la fenêtre sans bloquer
    $form.Show()
    $form.Refresh()
    [System.Windows.Forms.Application]::DoEvents()
}

function Update-LoadingWindow {
    <#
    .SYNOPSIS
    Met à jour le texte et la progression de la fenêtre de chargement avec animation fluide
    
    .PARAMETER Status
    Texte de statut à afficher
    
    .PARAMETER Percent
    Pourcentage d'avancement (0-100)
    #>
    param(
        [string]$Status,
        [int]$Percent
    )
    
    # Guard: vérifier que le form existe et n'est pas fermé
    if (-not $script:progressForm -or $script:progressForm.IsDisposed) {
        Write-Host "WARN: Loading window already closed, skipping update to $Percent%" -ForegroundColor Yellow
        return
    }
    
    # Mettre à jour le texte de statut
    $script:progressLabel.Text = $Status
    
    # Mise à jour instantanée et non-bloquante de la barre de progression
    $target = [Math]::Min([Math]::Max($Percent, 0), 100)
    $script:progressBar.Value = $target
    $script:progressPercent.Text = "$target%"
    
    # Rafraîchir l'interface
    [System.Windows.Forms.Application]::DoEvents()
}

function Update-LoadingWindowSmooth {
    <#
    .SYNOPSIS
    Met à jour la fenêtre avec progression fluide entre deux pourcentages
    
    .PARAMETER Status
    Texte de statut à afficher
    
    .PARAMETER FromPercent
    Pourcentage de départ (si -1, utilise la valeur actuelle)
    
    .PARAMETER ToPercent
    Pourcentage cible
    
    .PARAMETER DelayMs
    Délai en millisecondes entre chaque incrément (par défaut 60ms)
    #>
    param(
        [string]$Status,
        [int]$FromPercent = -1,
        [int]$ToPercent,
        [int]$DelayMs = 60
    )
    
    # Guard: vérifier que le form existe
    if (-not $script:progressForm -or $script:progressForm.IsDisposed) {
        return
    }
    
    # Mettre à jour le texte de statut
    if ($Status) {
        $script:progressLabel.Text = $Status
    }
    
    # Déterminer le point de départ
    $start = if ($FromPercent -ge 0) { $FromPercent } else { $script:progressBar.Value }
    $target = [Math]::Min([Math]::Max($ToPercent, 0), 100)
    
    # Progression fluide
    if ($start -lt $target) {
        for ($i = $start + 1; $i -le $target; $i++) {
            if (-not $script:progressForm -or $script:progressForm.IsDisposed) { break }
            
            $script:progressBar.Value = $i
            $script:progressPercent.Text = "$i%"
            
            # Fragmenter le délai pour permettre au spinner de tourner
            # Au lieu d'un seul gros Sleep, faire plusieurs petits avec DoEvents() entre
            $chunks = [Math]::Max(1, [int]($DelayMs / 15))
            for($k = 0; $k -lt $chunks; $k++) {
                [System.Windows.Forms.Application]::DoEvents()
                Start-Sleep -Milliseconds 15
            }
        }
    } else {
        # Si on doit descendre ou rester identique, mise à jour directe
        $script:progressBar.Value = $target
        $script:progressPercent.Text = "$target%"
        [System.Windows.Forms.Application]::DoEvents()
    }
}

function Close-LoadingWindow {
    <#
    .SYNOPSIS
    Ferme et libère la fenêtre de chargement
    #>
    
    if ($script:progressForm -and -not $script:progressForm.IsDisposed) {
        $script:progressForm.Close()
        $script:progressForm.Dispose()
        $script:progressForm = $null
        $script:progressLabel = $null
        $script:progressBar = $null
        $script:progressPercent = $null
    }
}

# =============================================================================
# CONFIGURATION MANAGEMENT
# =============================================================================

# Function to load configuration from config.json
function Load-Config {
    param(
        [string]$ConfigPath = (Join-Path $PSScriptRoot "config.json")
    )
    
    $defaultConfig = @{
        version = "1.0"
        defaultOutputPath = ""
        lastModified = (Get-Date -Format "o")
        autoOpenReport = $true
    }
    
    if (Test-Path $ConfigPath) {
        try {
            $config = Get-Content $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
            return $config
        } catch {
            Write-Host "⚠️  Fichier de configuration corrompu. Utilisation des paramètres par défaut." -ForegroundColor Yellow
            return $defaultConfig
        }
    }
    
    # Create default config file
    $defaultConfig | ConvertTo-Json | Set-Content $ConfigPath -Encoding UTF8
    return $defaultConfig
}

# Function to save configuration to config.json
function Save-Config {
    param(
        [Parameter(Mandatory=$true)]
        $Config,
        [string]$ConfigPath = (Join-Path $PSScriptRoot "config.json")
    )
    
    try {
        $Config.lastModified = (Get-Date -Format "o")
        $Config | ConvertTo-Json | Set-Content $ConfigPath -Encoding UTF8
        return $true
    } catch {
        Write-Host "❌ Erreur lors de la sauvegarde de la configuration: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Function to validate Power BI project structure
function Test-PowerBIProjectStructure {
    param(
        [string]$Path,
        [string]$ProjectType
    )
    
    if (-not (Test-Path $Path)) {
        Write-Host "  ✗ Le chemin n'existe pas: $Path" -ForegroundColor Red
        return $false
    }
    
    # Check for .pbip file
    $pbipFiles = Get-ChildItem -Path $Path -Filter "*.pbip" -File
    if ($pbipFiles.Count -eq 0) {
        Write-Host "  ✗ Aucun fichier .pbip trouve dans: $Path" -ForegroundColor Red
        return $false
    }
    
    # Check for Report folder
    $reportFolders = Get-ChildItem -Path $Path -Filter "*.Report" -Directory
    if ($reportFolders.Count -eq 0) {
        Write-Host "  ✗ Aucun dossier .Report trouve dans: $Path" -ForegroundColor Red
        return $false
    }
    
    # Check for SemanticModel folder
    $semanticFolders = Get-ChildItem -Path $Path -Filter "*.SemanticModel" -Directory
    if ($semanticFolders.Count -eq 0) {
        Write-Host "  ✗ Aucun dossier .SemanticModel trouve dans: $Path" -ForegroundColor Red
        return $false
    }
    
    Write-Host "  ✓ Structure Power BI valide pour $ProjectType" -ForegroundColor Green
    Write-Host "    - Fichier: $($pbipFiles[0].Name)" -ForegroundColor Gray
    Write-Host "    - Dossier Report: $($reportFolders[0].Name)" -ForegroundColor Gray
    Write-Host "    - Dossier SemanticModel: $($semanticFolders[0].Name)" -ForegroundColor Gray
    return $true
}

# Function to show Windows folder selection dialog
# Compatible avec .exe compilé - utilise les assemblies déjà chargées
function Select-Folder {
    param(
        [string]$Description,
        [string]$InitialDirectory = [Environment]::GetFolderPath('MyDocuments'),
        [string]$PresetPath  # If provided, skip dialog and use this path
    )
    
    # If preset path is provided (non-interactive mode), validate and return it
    if ($PresetPath) {
        if (Test-Path $PresetPath) {
            return $PresetPath
        } else {
            Write-Host "❌ Le chemin fourni est invalide: $PresetPath" -ForegroundColor Red
            return $null
        }
    }
    
    # If in non-interactive mode without preset path, return null
    if ($script:NonInteractiveMode) {
        Write-Host "❌ Mode non-interactif: un chemin doit être fourni via paramètres." -ForegroundColor Red
        return $null
    }
    
    try {
        # Utiliser directement les assemblies déjà chargées (pas de sous-processus)
        $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $dialog.Description = $Description
        $dialog.ShowNewFolderButton = $false
        
        if (Test-Path $InitialDirectory) {
            $dialog.SelectedPath = $InitialDirectory
        }
        
        # Try to set apartment state only if not already set
        try {
            if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne [System.Threading.ApartmentState]::STA) {
                [System.Threading.Thread]::CurrentThread.SetApartmentState([System.Threading.ApartmentState]::STA)
            }
        } catch {
            # Ignore if already set or cannot be changed
        }
        
        $result = $dialog.ShowDialog()
        
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            return $dialog.SelectedPath
        }
        
        return $null
        
    } catch {
        Write-Host "⚠️  Erreur lors de l'affichage du dialog: $($_.Exception.Message)" -ForegroundColor Yellow
        
        # In non-interactive environment, Read-Host might fail
        try {
            Write-Host "Veuillez entrer le chemin manuellement." -ForegroundColor Yellow
            $manualPath = Read-Host "Entrez le chemin complet du dossier"
            
            if ($manualPath -and (Test-Path $manualPath)) {
                return $manualPath
            } else {
                Write-Host "Chemin invalide ou inexistant." -ForegroundColor Red
                return $null
            }
        } catch {
            Write-Host "❌ Interaction utilisateur impossible. Utilisez -NonInteractive avec des paramètres." -ForegroundColor Red
            return $null
        }
    } finally {
        if ($dialog) {
            $dialog.Dispose()
        }
    }
}

Write-Host "`n╔════════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║  Comparateur de Rapports Power BI - Orange Business Services  ║" -ForegroundColor Cyan
Write-Host "╚════════════════════════════════════════════════════════════════╝`n" -ForegroundColor Cyan

# Set global flag for non-interactive mode
$script:NonInteractiveMode = $NonInteractive

# Initialize configuration (creates config.json if it doesn't exist)
# $scriptRoot already defined at top of script
$config = Load-Config -ConfigPath (Join-Path $scriptRoot "config.json")

# Step 1: Select NEW version project
Write-Host "[1/3] Sélection du projet NOUVEAU (version récente)" -ForegroundColor Yellow

if ($NonInteractive) {
    # Non-interactive mode: use provided parameters
    if (-not $NewVersionPath) {
        Write-Host "❌ Erreur: -NewVersionPath est requis en mode non-interactif." -ForegroundColor Red
        Write-Host "   Exemple: script.exe -NonInteractive -NewVersionPath 'C:\path' -OldVersionPath 'C:\path2' -OutputPath 'C:\output'" -ForegroundColor Yellow
        exit 1
    }
    $NewVersionFolderPath = $NewVersionPath
    Write-Host "   Utilisation du chemin fourni: $NewVersionFolderPath" -ForegroundColor Cyan
} else {
    # Interactive mode: show dialog
    $NewVersionFolderPath = Select-Folder -Description "Sélectionnez le dossier du projet Power BI NOUVEAU (version récente)"
}

if (-not $NewVersionFolderPath) {
    Write-Host "`n✗ Opération annulée par l'utilisateur." -ForegroundColor Red
    Write-Host "  Aucun dossier sélectionné pour la nouvelle version.`n" -ForegroundColor Red
    exit 1
}

Write-Host "`nChemin sélectionné: $NewVersionFolderPath" -ForegroundColor White
if (-not (Test-PowerBIProjectStructure -Path $NewVersionFolderPath -ProjectType "NOUVEAU")) {
    Write-Host "`n✗ Structure Power BI invalide pour la nouvelle version." -ForegroundColor Red
    Write-Host "  Le dossier doit contenir: fichier .pbip, dossier .Report, dossier .SemanticModel`n" -ForegroundColor Red
    exit 1
}

# Step 2: Select OLD version project
Write-Host "`n[2/3] Sélection du projet ANCIEN (version de référence)" -ForegroundColor Yellow

if ($NonInteractive) {
    # Non-interactive mode: use provided parameters
    if (-not $OldVersionPath) {
        Write-Host "❌ Erreur: -OldVersionPath est requis en mode non-interactif." -ForegroundColor Red
        exit 1
    }
    $OldVersionFolderPath = $OldVersionPath
    Write-Host "   Utilisation du chemin fourni: $OldVersionFolderPath" -ForegroundColor Cyan
} else {
    # Interactive mode: show dialog
    $OldVersionFolderPath = Select-Folder -Description "Sélectionnez le dossier du projet Power BI ANCIEN (version de référence)" -InitialDirectory (Split-Path $NewVersionFolderPath -Parent)
}

if (-not $OldVersionFolderPath) {
    Write-Host "`n✗ Opération annulée par l'utilisateur." -ForegroundColor Red
    Write-Host "  Aucun dossier sélectionné pour l'ancienne version.`n" -ForegroundColor Red
    exit 1
}

Write-Host "`nChemin sélectionné: $OldVersionFolderPath" -ForegroundColor White
if (-not (Test-PowerBIProjectStructure -Path $OldVersionFolderPath -ProjectType "ANCIEN")) {
    Write-Host "`n✗ Structure Power BI invalide pour l'ancienne version." -ForegroundColor Red
    Write-Host "  Le dossier doit contenir: fichier .pbip, dossier .Report, dossier .SemanticModel`n" -ForegroundColor Red
    exit 1
}

# Step 3: Select output folder (or use default from config)
Write-Host "`n[3/3] Sélection du dossier de sortie pour le rapport HTML" -ForegroundColor Yellow

if ($NonInteractive) {
    # Non-interactive mode: use provided parameter or config
    if ($OutputPath) {
        $BaseTargetFolder = $OutputPath
        Write-Host "   Utilisation du chemin fourni: $BaseTargetFolder" -ForegroundColor Cyan
    } elseif ($config.defaultOutputPath -and (Test-Path $config.defaultOutputPath)) {
        $BaseTargetFolder = $config.defaultOutputPath
        Write-Host "   Utilisation du chemin configuré: $BaseTargetFolder" -ForegroundColor Cyan
    } else {
        Write-Host "❌ Erreur: -OutputPath est requis en mode non-interactif si aucun chemin par défaut n'est configuré." -ForegroundColor Red
        exit 1
    }
} else {
    # Interactive mode: check config or ask user
    if ($config.defaultOutputPath -and (Test-Path $config.defaultOutputPath)) {
        $BaseTargetFolder = $config.defaultOutputPath
        Write-Host "✓ Utilisation du dossier configuré : $BaseTargetFolder" -ForegroundColor Cyan
        Write-Host "  (Pour modifier : exécutez configure.ps1)" -ForegroundColor Gray
    } else {
        # Ask user to select folder
        if ($config.defaultOutputPath) {
            Write-Host "  ⚠️  Le dossier configuré n'existe plus : $($config.defaultOutputPath)" -ForegroundColor Yellow
        }
        
        $BaseTargetFolder = Select-Folder -Description "Sélectionnez le dossier de sortie pour le rapport HTML" -InitialDirectory (Split-Path $NewVersionFolderPath -Parent)
        
        if (-not $BaseTargetFolder) {
            Write-Host "`n✗ Opération annulée par l'utilisateur." -ForegroundColor Red
            Write-Host "  Aucun dossier de sortie sélectionné.`n" -ForegroundColor Red
            exit 1
        }
        
        Write-Host "`nDossier de sortie: $BaseTargetFolder" -ForegroundColor White
        
        # Save this path as the new default in config.json
        $config.defaultOutputPath = $BaseTargetFolder
        if (Save-Config -Config $config -ConfigPath (Join-Path $scriptRoot "config.json")) {
            Write-Host "  ✓ Configuration sauvegardée pour les prochaines exécutions" -ForegroundColor Green
        }
    }
}

# Also update the config object with current path for HTML injection
$config.defaultOutputPath = $BaseTargetFolder

$ReportOutputFolder = Join-Path $BaseTargetFolder "RapportHTML"

# Summary of selections
Write-Host "`n╔════════════════════════════════════════════════════════════════╗" -ForegroundColor Green
Write-Host "║                     Résumé des sélections                      ║" -ForegroundColor Green
Write-Host "╚════════════════════════════════════════════════════════════════╝" -ForegroundColor Green
Write-Host "  → NOUVEAU projet : $NewVersionFolderPath" -ForegroundColor Cyan
Write-Host "  → ANCIEN projet  : $OldVersionFolderPath" -ForegroundColor Cyan
Write-Host "  → Sortie rapport : $ReportOutputFolder" -ForegroundColor Cyan
Write-Host ""

# Create the output directory if it doesn't exist
if (-not (Test-Path $ReportOutputFolder)) {
    New-Item -ItemType Directory -Path $ReportOutputFolder | Out-Null
}

# Load helper modules ($scriptRoot already defined at top of script)
. (Join-Path $scriptRoot 'PBI_Classes.ps1')

# ═══════════════════════════════════════════════════════════════════════════
# DÉBUT DE L'ANALYSE - Fenêtre de chargement
# ═══════════════════════════════════════════════════════════════════════════

# Afficher la fenêtre de chargement
Show-LoadingWindow
Start-Sleep -Milliseconds 300  # Pause pour que la fenêtre s'affiche correctement

try {
    # Forcer toutes les erreurs à être terminantes pour le try/catch
    $ErrorActionPreference = 'Stop'
    
    # Chargement silencieux des DLLs et modules (rapide, pas d'affichage)
    . (Join-Path $scriptRoot 'PBI_load_dll.ps1')
    LoadNeededDLL -path $scriptRoot
    . (Join-Path $scriptRoot 'PBI_Report_HTML_Orange.ps1')
    . (Join-Path $scriptRoot 'PBI_Report_Check.ps1')
    . (Join-Path $scriptRoot 'PBI_Report_Compare.ps1')
    . (Join-Path $scriptRoot 'PBI_MDD_extract.ps1')

    # 1. Semantic Analysis - Phase 1 : Chargement des projets (0% → 20%)
    Update-LoadingWindowSmooth -Status "Analyse sémantique : initialisation..." -ToPercent 3 -DelayMs 70
    Write-Host "`n=== Analyse Sémantique ===" -ForegroundColor Green

    try {
        Update-LoadingWindowSmooth -Status "Analyse sémantique : chargement des modèles..." -ToPercent 8 -DelayMs 60
        # Suppress verbose output from semantic analysis functions
        $projetArray = LoadProjectVersionsPath -newVersionProjectRepertory $NewVersionFolderPath -oldVersionProjectRepertory $OldVersionFolderPath 2>$null
        
        Update-LoadingWindowSmooth -Status "Analyse sémantique : préparation des modèles..." -ToPercent 15 -DelayMs 50
        
        if (-not $projetArray -or $projetArray.Count -lt 2) {
            Write-Host "  ! Impossible de charger les projets semantiques" -ForegroundColor Yellow
            $semanticComparisonResult = $null
        } else {
            Update-LoadingWindowSmooth -Status "Analyse sémantique : parsing des structures..." -ToPercent 18 -DelayMs 60
            
            $parsedArray = [PSCustomObject]@{
                ElementNewVersion = $projetArray[0].dataBase.Model
                ElementOldVersion = $projetArray[1].dataBase.Model
                PathNewVersion = 'DataBase'
                PathOldVersion = 'DataBase'
            }
            
            Update-LoadingWindowSmooth -Status "Analyse sémantique : comparaison en cours..." -ToPercent 22 -DelayMs 60
            $semanticComparisonResult = CheckDifferenceInSubElement -element $parsedArray -elementChecked 'Model' 2>$null
            
            Update-LoadingWindowSmooth -Status "Analyse sémantique : traitement des résultats..." -ToPercent 37 -DelayMs 50
            
            if ($semanticComparisonResult -and $semanticComparisonResult.Count -gt 0) {
                Write-Host "   Analyse semantique terminee: $($semanticComparisonResult.Count) differences detectees" -ForegroundColor Green
                Write-Host "  Type de donnees: $($semanticComparisonResult.GetType().Name)" -ForegroundColor Cyan
            } else {
                Write-Host "  ! Aucune difference semantique detectee" -ForegroundColor Yellow
            }
        }
    } catch {
        Write-Host "   ERREUR lors de l'analyse semantique: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "  Details: $($_.ScriptStackTrace)" -ForegroundColor Red
        $semanticComparisonResult = $null
    }
    
    Update-LoadingWindowSmooth -Status "Analyse sémantique terminée" -ToPercent 40 -DelayMs 40

    # 2. Report Analysis
    Update-LoadingWindowSmooth -Status "Chargement des définitions de rapport..." -ToPercent 42 -DelayMs 60
    Write-Host "`n=== Analyse du Rapport ===" -ForegroundColor Green

    try {
        # Load the Report projects
        Update-LoadingWindowSmooth -Status "Analyse des rapports : chargement..." -ToPercent 45 -DelayMs 50
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
    Update-LoadingWindowSmooth -Status "Analyse des rapports : comparaison..." -ToPercent 50 -DelayMs 60
    $reportDifferences = CompareReportProjects -newProject $newReportProject -oldProject $oldReportProject
    
    Update-LoadingWindowSmooth -Status "Analyse des rapports : traitement..." -ToPercent 58 -DelayMs 50

    Write-Host "  $($reportDifferences.Count) differences de rapport trouvees" -ForegroundColor White
    
    Update-LoadingWindowSmooth -Status "Analyse des rapports terminée" -ToPercent 60 -DelayMs 40

    # 3. Quality check on visible slicers (new project only)
    Update-LoadingWindowSmooth -Status "Vérification qualité des slicers..." -ToPercent 65 -DelayMs 50
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
    
    Update-LoadingWindowSmooth -Status "Vérification terminée" -ToPercent 75 -DelayMs 40

    # 4. Génération du rapport HTML
    Update-LoadingWindowSmooth -Status "Préparation des données pour le rapport..." -ToPercent 76 -DelayMs 70
    Write-Host "`n=== Generation des rapports HTML ===" -ForegroundColor Green
    
    # Debug avec protection contre null
    if ($semanticComparisonResult) {
        Write-Host "  Debug: semanticComparisonResult type = $($semanticComparisonResult.GetType().Name)" -ForegroundColor Cyan
        Write-Host "  Debug: semanticComparisonResult count = $($semanticComparisonResult.Count)" -ForegroundColor Cyan
    } else {
        Write-Host "  Debug: semanticComparisonResult = null" -ForegroundColor Cyan
    }

    # Convert differences for HTML
    Update-LoadingWindowSmooth -Status "Conversion des données..." -ToPercent 80 -DelayMs 60
    $differencesForHtml = Convert-ReportDifferencesForHtml -differences $reportDifferences

    # Generate ORANGE HTML report (same data, Orange branding)
    Update-LoadingWindowSmooth -Status "Génération du rapport HTML..." -ToPercent 85 -DelayMs 50
    Write-Host "`n--- Génération du rapport ORANGE ---" -ForegroundColor Magenta
    $configJsonPath = Join-Path $scriptRoot "config.json"
    
    Update-LoadingWindowSmooth -Status "Création du fichier HTML..." -ToPercent 90 -DelayMs 60
    $reportPathOrange = BuildReportHTMLReport_Orange -differences $differencesForHtml -checkResults $checkResults -outputFolder $ReportOutputFolder -semanticComparisonResult $semanticComparisonResult -configPath $configJsonPath
    
    Update-LoadingWindowSmooth -Status "Rapport HTML généré avec succès" -ToPercent 95 -DelayMs 40
    Write-Host "   Rapport HTML Orange genere: $reportPathOrange" -ForegroundColor Magenta

    # Finalisation
    Update-LoadingWindowSmooth -Status "Ouverture du rapport..." -ToPercent 100 -DelayMs 50
    Start-Sleep -Milliseconds 500  # Pause pour que l'utilisateur voie "100%"

    Write-Host "`n=== Analyse terminée ===" -ForegroundColor Green
    Write-Host "  → Rapport: $reportPathOrange" -ForegroundColor White
    
    # Ouvrir automatiquement le rapport dans le navigateur par défaut
    Write-Host "`n📂 Ouverture du rapport dans le navigateur..." -ForegroundColor Cyan
    Start-Process $reportPathOrange
    
} catch {
    # En cas d'erreur, afficher de multiples façons pour garantir visibilité
    Write-Host "`n" -ForegroundColor White
    Write-Host "╔══════════════════════════════════════════════════╗" -ForegroundColor Red
    Write-Host "║           ❌ ERREUR CRITIQUE DÉTECTÉE            ║" -ForegroundColor Red
    Write-Host "╚══════════════════════════════════════════════════╝" -ForegroundColor Red
    Write-Host ""
    Write-Host "Message: $($_.Exception.Message)" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Type d'erreur: $($_.Exception.GetType().FullName)" -ForegroundColor DarkYellow
    Write-Host ""
    Write-Host "Stack trace:" -ForegroundColor DarkYellow
    Write-Host $_.ScriptStackTrace -ForegroundColor Gray
    Write-Host ""
    
    # Afficher aussi dans une MessageBox Windows
    try {
        [System.Windows.Forms.MessageBox]::Show(
            "Une erreur s'est produite :`n`n$($_.Exception.Message)`n`nConsultez le terminal pour plus de détails.",
            "Erreur - Comparateur Power BI",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    } catch {
        # Si même la MessageBox échoue, au moins on a les Write-Host
    }
    
    Close-LoadingWindow
    throw
} finally {
    # Toujours fermer la fenêtre de chargement à la fin
    Close-LoadingWindow
}
