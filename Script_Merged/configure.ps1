# =============================================================================
# Configuration du Comparateur de Rapports Power BI
# =============================================================================
# Ce script permet de configurer les paramètres persistants de l'application.
# =============================================================================

# Load Windows Forms for folder selection
Add-Type -AssemblyName System.Windows.Forms

# Get script path
$scriptRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
$configPath = Join-Path $scriptRoot "config.json"

# Function to show Windows folder selection dialog (copied from main.ps1)
function Select-Folder {
    param(
        [string]$Description,
        [string]$InitialDirectory = [Environment]::GetFolderPath('MyDocuments')
    )
    
    $escapedDesc = $Description -replace "'", "''"
    $escapedInitDir = $InitialDirectory -replace "'", "''"
    
    $staScript = @"
Add-Type -AssemblyName System.Windows.Forms
`$dialog = New-Object System.Windows.Forms.FolderBrowserDialog
`$dialog.Description = '$escapedDesc'
`$dialog.ShowNewFolderButton = `$false
if (Test-Path '$escapedInitDir') {
    `$dialog.SelectedPath = '$escapedInitDir'
}
`$result = `$dialog.ShowDialog()
if (`$result -eq [System.Windows.Forms.DialogResult]::OK) {
    Write-Output `$dialog.SelectedPath
}
"@
    
    $selectedPath = pwsh -STA -NoProfile -Command $staScript
    
    if ($selectedPath) {
        return $selectedPath.Trim()
    }
    
    return $null
}

# Load or create default config
function Load-Config {
    $defaultConfig = @{
        version = "1.0"
        defaultOutputPath = ""
        lastModified = (Get-Date -Format "o")
        autoOpenReport = $true
    }
    
    if (Test-Path $configPath) {
        try {
            return Get-Content $configPath -Raw -Encoding UTF8 | ConvertFrom-Json
        } catch {
            Write-Host "⚠️  Fichier de configuration corrompu. Réinitialisation..." -ForegroundColor Yellow
            return $defaultConfig
        }
    }
    
    return $defaultConfig
}

# Save configuration
function Save-Config {
    param($Config)
    
    try {
        $Config.lastModified = (Get-Date -Format "o")
        $Config | ConvertTo-Json | Set-Content $configPath -Encoding UTF8
        Write-Host "✓ Configuration sauvegardée avec succès !" -ForegroundColor Green
        return $true
    } catch {
        Write-Host "❌ Erreur lors de la sauvegarde : $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Display header
Clear-Host
Write-Host "`n╔════════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║        Configuration - Comparateur de Rapports Power BI       ║" -ForegroundColor Cyan
Write-Host "╚════════════════════════════════════════════════════════════════╝`n" -ForegroundColor Cyan

# Load current configuration
$config = Load-Config

# Display current settings
Write-Host "📋 Paramètres actuels :" -ForegroundColor Yellow
Write-Host ""
if ($config.defaultOutputPath) {
    $status = if (Test-Path $config.defaultOutputPath) { "✓ Valide" } else { "⚠️  Introuvable" }
    Write-Host "  Dossier de sortie par défaut : $($config.defaultOutputPath)" -ForegroundColor White
    Write-Host "  Statut : $status" -ForegroundColor $(if (Test-Path $config.defaultOutputPath) { "Green" } else { "Yellow" })
} else {
    Write-Host "  Dossier de sortie par défaut : Non configuré (demandé à chaque exécution)" -ForegroundColor Gray
}
Write-Host ""

# Menu
Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "Actions disponibles :" -ForegroundColor Yellow
Write-Host "  [1] Configurer le dossier de sortie par défaut" -ForegroundColor White
Write-Host "  [2] Réinitialiser (demander à chaque fois)" -ForegroundColor White
Write-Host "  [0] Quitter" -ForegroundColor White
Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""

$choice = Read-Host "Votre choix"

switch ($choice) {
    "1" {
        Write-Host "`n📁 Sélection du dossier de sortie par défaut..." -ForegroundColor Yellow
        
        $initialDir = if ($config.defaultOutputPath -and (Test-Path $config.defaultOutputPath)) {
            $config.defaultOutputPath
        } else {
            [Environment]::GetFolderPath('MyDocuments')
        }
        
        $selectedPath = Select-Folder -Description "Sélectionnez le dossier de sortie par défaut pour les rapports HTML" -InitialDirectory $initialDir
        
        if ($selectedPath) {
            $config.defaultOutputPath = $selectedPath
            
            if (Save-Config -Config $config) {
                Write-Host "`n✅ Dossier par défaut configuré : $selectedPath" -ForegroundColor Green
                Write-Host "   Les prochaines exécutions utiliseront automatiquement ce dossier." -ForegroundColor Gray
            }
        } else {
            Write-Host "`n⚠️  Aucun dossier sélectionné. Configuration annulée." -ForegroundColor Yellow
        }
    }
    
    "2" {
        Write-Host "`n🔄 Réinitialisation du dossier par défaut..." -ForegroundColor Yellow
        $config.defaultOutputPath = ""
        
        if (Save-Config -Config $config) {
            Write-Host "`n✅ Configuration réinitialisée." -ForegroundColor Green
            Write-Host "   Le dossier de sortie sera demandé à chaque exécution." -ForegroundColor Gray
        }
    }
    
    "0" {
        Write-Host "`n👋 À bientôt !" -ForegroundColor Cyan
    }
    
    default {
        Write-Host "`n❌ Choix invalide." -ForegroundColor Red
    }
}

Write-Host ""
