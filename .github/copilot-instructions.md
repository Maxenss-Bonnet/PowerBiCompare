# Copilot — Instructions pour ce dépôt

Ce dépôt compare deux versions d’un rapport Power BI (PBIR Enhanced) et génère un rapport HTML interactif. Le code a été fusionné dans `Script_Merged` (merge de Script_Semantic et Script_Report).

## Exécution principale (Windows, pwsh)

- Script d’orchestration: `Script_Merged/main.ps1`.
- Éditer au début du fichier: `$NewVersionFolderPath`, `$OldVersionFolderPath`, `$BaseTargetFolder`, `$ReportOutputFolder`.
- Résultat: rapport HTML brandé ORANGE dans `RapportHTML/`.
- Après chaque exécution, vérifier le rapport avec MCP Chrome DevTools: ouvrir l’HTML, contrôler la console (aucune erreur JS), tester recherche/filtres/i18n et interactions. Utiliser systématiquement le MCP Chrome DevTools (chrome-devtools-mcp) pour valider l’UI et le DOM du rapport.

## Architecture modulaire (Script_Merged)

- Diff moteur: `PBI_Report_Compare.ps1` (fonction publique: `CompareReportProjects`).
- Qualité slicers: `PBI_Report_Check.ps1` (fonctions: `Invoke-VisibleSlicerQualityCheck`, `Invoke-SlicerQualityCheck`).
- Sémantique (TMDL): `PBI_MDD_extract.ps1` (fonctions: `LoadProjectVersionsPath`, `CheckDifferenceInSubElement`).
- HTML: `PBI_Report_HTML_Orange.ps1` — fonction publique: `BuildReportHTMLReport_Orange`.
- Modèles partagés: `PBI_Classes.ps1` (ex: `ReportProject`, `ReportDifference`).
- Chargement DLL: `PBI_load_dll.ps1` → `LoadNeededDLL -path $scriptRoot` (assemblies vendorisés dans `Script_Merged/lib`, ne pas télécharger en externe).

## Références design Figma

- Style master du tableau de comparaison: `FigmaExports/One-CI-CIR/table-style-spec.md` (extrait du frame Figma Accueil / Version A).
- Lors de toute évolution de rendu HTML/CSS, relire ce fichier et aligner les tokens/couleurs/typo indiqués.
- Toujours valider la conformité avec l’outil MCP `mcp_figma-mcp-ser_create_design_system_rules` en plus des captures `get_design_context`/`get_screenshot`.

Interfaces publiques à préserver (compat): `CompareReportProjects`, `Invoke-VisibleSlicerQualityCheck`/`Invoke-SlicerQualityCheck`, `BuildReportHTMLReport_Orange`, `LoadProjectVersionsPath`, `CheckDifferenceInSubElement`, `LoadReportProjectVersions`.

## Structure PBIR (Enhanced)

```
{Project}.Report/
├── .pbi/localSettings.json
├── definition/
│   ├── version.json
│   ├── pages/{pageId}/page.json + visuals/{visualId}/visual.json
│   ├── bookmarks/{bookmarkId}.bookmark.json
│   └── report.json
```
Identifiants stables: `pageId`, `visualId`, `bookmarkId`.

## Règle critique — Display name des visuels/slicers

- Toujours résoudre le titre via les helpers fournis; ne jamais afficher d’ID technique.
- JSON path source: `visual.objects.header[].properties.text.expr.Literal.Value` (valeur quotée: `'Titre'`).
- Fonctions: `Resolve-SlicerDisplayName` (dans `PBI_Report_Compare.ps1`) ou `Get-SlicerDisplayName` (dans `PBI_Report_Check.ps1`).
- Fallback autorisé uniquement si header manquant; ne pas exposer d’IDs techniques en HTML.

Exemple JSON complet:
```json
{
  "visual": {
    "objects": {
      "header": [
        {
          "properties": {
            "text": {
              "expr": {
                "Literal": { "Value": "'NomDuVisuel'" }
              }
            }
          }
        }
      ]
    }
  }
}
```

## Qualité — Slicers visibles seulement

- Cibler uniquement les slicers visibles et respecter la visibilité du volet Filtres.
- Conserver la sémantique: présence de sélection, mode single-select, texte de recherche, règles période/année courante, visibilité page/visuel.
- Utiliser `Invoke-VisibleSlicerQualityCheck`; filtrer les résultats pour l’HTML sans exposer d’IDs techniques.

## Comparaison — Rapport et Sémantique

- Rapport: inclure pages, visuels, bookmarks, config, thèmes, fichiers système; conserver l’analyse des interactions et des sync-groups (`Group-SynchronizationChanges`, `Analyze-SyncGroupCauses`).
- Sémantique (TMDL): charger via `LoadProjectVersionsPath`, comparer `Model` avec `CheckDifferenceInSubElement`, puis passer tel quel aux builders HTML.

## Entrées/Sorties, erreurs et performance

- Lecture JSON: `-Raw -Encoding UTF8`. Gérer nœuds manquants/vides défensivement; préférer des chemins absolus.
- Envelopper en `try/catch` avec messages clairs; éviter les re-parsings larges; préserver l’ordre d’extraction; garder des temps d’exécution courts.

## Rapport HTML (i18n FR/EN)

- Génération via `BuildReportHTMLReport_Orange` (UI interactive: recherche, filtres, i18n).
- Ne pas casser les clés i18n ni l’UI de recherche/filtrage.

## Commandes de développement (adaptées Script_Merged)

- Exécution principale:
```powershell
pwsh -File "C:\Users\BQTR7546\.vscodeProject\PowerBiCompare\Script_Merged\main.ps1"
```
- Lint (recommandé):
```powershell
Install-Module PSScriptAnalyzer -Scope CurrentUser -Force
Invoke-ScriptAnalyzer -Path "C:\Users\BQTR7546\.vscodeProject\PowerBiCompare\Script_Merged" -Recurse
```
- Vérification post-run (obligatoire): ouvrir le HTML généré et valider avec MCP Chrome DevTools (chrome-devtools-mcp) — console sans erreurs, filtres/recherche fonctionnels, i18n OK, interactions et totaux cohérents.

## Tests de non-régression (smoke)

Après toute modif:
1) Exécuter `Script_Merged/main.ps1`.
2) Vérifier: résumé des différences, rapport HTML ORANGE généré sans erreur.
3) Totaux attendus (références): 12 pages, ~150 visuels (test2), 224 signets; temps d’exécution cible < 20s.

## Règles de dev spécifiques

- Travailler sur: `Script_Merged/PBI_Report_Compare.ps1`, `PBI_Report_HTML_Orange.ps1`, `PBI_Report_Check.ps1`, `PBI_MDD_extract.ps1`, `PBI_load_dll.ps1`, `PBI_Classes.ps1`.
- Ne pas modifier: artefacts PBIR des dossiers d’échantillons; ne pas committer de secrets.
- Ne changez pas les noms de fonctions publiques sans MAJ de tous les call-sites et bindings HTML.
- Petites PRs: garder les diffs localisés, créer de petits helpers plutôt que dupliquer la logique.

## Standards PowerShell

- Encodage: UTF-8 obligatoire; erreurs: `try/catch`; commentaires courts et factuels.
- Chemins absolus privilégiés; éviter les traversées inutiles.
- Sécurité: pas de secrets en clair, pas de logs sensibles.

## Boucle anti-erreurs (obligatoire)

1) Lancer le script principal:
```powershell
pwsh -File "C:\Users\BQTR7546\.vscodeProject\PowerBiCompare\Script_Merged\main.ps1"
```
2) Capturer toutes les erreurs (stderr, stacktraces, exit code, warnings).
3) Corriger immédiatement et rejouer jusqu’à: aucune erreur/warning bloquant, HTML OK, totaux cohérents.
4) Ouvrir le HTML et valider via MCP Chrome DevTools (console, DOM, interactions, filtres, i18n).
5) Lancer le lint PSScriptAnalyzer avant de conclure.

## Configuration des chemins (main.ps1)

Variables à définir en tête de `main.ps1`:
```powershell
$NewVersionFolderPath = "...chemin du projet NOUVEAU..."
$OldVersionFolderPath = "...chemin du projet ANCIEN..."
$BaseTargetFolder     = "...racine de sortie..."
$ReportOutputFolder   = Join-Path $BaseTargetFolder "RapportHTML"
```

## Schémas et références

- Schémas Microsoft: présents via `$schema` dans les JSON PBIR.
- Format PBIR: Enhanced Report 3.0.0.
- Fichiers clés: `report.json`, `version.json`, `page.json`, `visual.json`, `bookmark.json`.
- Docs internes utiles: `PowerBI_Sync_Understanding.md` (syncGroup, fieldChanges, filterChanges, isHidden), autres docs PBIR si disponibles.

## Sécurité et performance

- Lecture seule des dossiers PBIR; sortie unique: fichiers HTML.
- Logs sobres; performance: lecture ciblée, `-Raw -Encoding UTF8`.
- Préserver l’ordre de chargement; paralléliser seulement si sûr.

## Notes

- Toutes les réponses dans les scripts/rapports doivent être en français sauf demande contraire explicite.
- Préférer éditer l’existant plutôt que multiplier les fichiers.

