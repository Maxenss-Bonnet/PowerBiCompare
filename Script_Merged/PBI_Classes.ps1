class ReportProject {
    [string] $name
    [string] $nameAndTimestamp
    [string] $path
    [string] $reportPath
    [PSCustomObject] $reportData
    [hashtable] $pagesData
    [hashtable] $bookmarksData
    [PSCustomObject] $bookmarkMetadata
    
    ReportProject() {
        $this.pagesData = @{}
        $this.bookmarksData = @{}
        $this.bookmarkMetadata = $null
    }
}

class ReportDifference {
    [string] $ElementType
    [string] $ElementName
    [string] $ElementDisplayName
    [string] $ElementPath
    [string] $ParentElementName
    [string] $ParentDisplayName
    [string] $PropertyName
    [string] $DifferenceType
    [string] $OldValue
    [string] $NewValue
    [string] $HierarchyLevel
    [string] $AdditionalInfo
    # Nouvelles propriétés pour la hiérarchisation des synchronisations
    [string] $SyncGroupName
    [bool] $IsPrimarySyncChange
    [string] $SyncGroupId
    [System.Collections.ArrayList] $RelatedSyncChanges
    [bool] $IsSynthetic

    ReportDifference() {
        $this.RelatedSyncChanges = [System.Collections.ArrayList]::new()
        $this.IsPrimarySyncChange = $false
        $this.IsSynthetic = $false
    }
}

if (-not ([System.Web.HttpUtility] -as [type])) {
    Add-Type -AssemblyName System.Web
}

$script:SyntheticSyncOrigins = @{}