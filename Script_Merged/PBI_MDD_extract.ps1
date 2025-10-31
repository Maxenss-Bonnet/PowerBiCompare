#This script analysis two different data model from two differents PBI project
#These project qhould be stored as PBIP files, and the folder where the two projects are stored should be provided


#==========================================================================================================================================


#Declaration
class Project {
    [string] $name;
    [string] $nameAndTimestamp;
    [string] $path;

    $dataBase;
}


#Specific classes to display things differently
class Parameter{
    [string]$Name
    [Microsoft.AnalysisServices.Tabular.ExpressionKind] $Kind
    [Microsoft.AnalysisServices.Tabular.QueryGroup] $QueryGroup
    [string]$ExpressionValue
    [string]$ExpressionMeta

    Parameter([string]$name, [Microsoft.AnalysisServices.Tabular.ExpressionKind] $kind, [Microsoft.AnalysisServices.Tabular.QueryGroup] $queryGroup, [string]$expression) {
        $this.Name = $name
        $this.Kind = $kind
        $this.QueryGroup = $queryGroup
        $this.ExpressionValue = ($expression -split 'meta' | Select-Object -First 1).Trim()
        $this.ExpressionMeta = ((($expression -split '\[')[1] -split '\]' | Select-Object -First 1).Trim()) -replace ',', "`n"
    }
}


#==========================================================================================================================================


Function LoadProjectVersionsPath {

    param (
        $newVersionProjectRepertory = $null,
        $oldVersionProjectRepertory = $null
    )

    $projectName = $null
    $errorCount = 0

    #Ask the repertory where the project is stored:
    if(-not $newVersionProjectRepertory) { $newVersionProjectRepertory = Read-Host "Please enter the repertory where the new .pbip is stored" }
    if(-not $oldVersionProjectRepertory) { $oldVersionProjectRepertory = Read-Host "Please enter the repertory where the old .pbip is stored" }
    Write-Host ("Doing the comparison between {0} and {1}" -f $newVersionProjectRepertory, $oldVersionProjectRepertory)

    #TODO verifier le formatage du nom de dossier (genre " et / à virer si entré en trop)

    $projectArray = @(
        $newVersionProjectRepertory,
        $oldVersionProjectRepertory
    )
    $returnArray = @()

    ForEach ( $version in $projectArray) {
        Push-Location -Path $version
        $checkedFile = Get-ChildItem -Filter "*.pbip"
        Pop-Location;

        if ($checkedFile -eq $null) {
            Write-Host ('The folder ' + $version + ' do not contains any .pbip file')
            $errorCount += 1
            continue
        }
        elseif ($checkedFile.count -gt 1){
            Write-Host ('The folder ' + $version + ' contains more than on .pbip file')
            $errorCount += 1
            continue
        }

        #todo faire la gestion de si il manque les deux répertoire (report et semantic model)

        else {
            $projet = [Project]::new()

            Write-Host ('Project named ' + $checkedFile.BaseName + ' found!')

            $projet.name = $checkedFile.BaseName
            $projet.nameAndTimestamp = $checkedFile.BaseName + '__' + $checkedFile.LastWriteTimeUtc
            $projet.path = $version
            write-host ('Version loader cheking for ' + $version + '\' + $checkedFile.BaseName + '.SemanticModel\definition')
            $projet.dataBase = [Microsoft.AnalysisServices.Tabular.TmdlSerializer]::DeserializeDatabaseFromFolder($version + '\' + $checkedFile.BaseName + '.SemanticModel\definition');

            $returnArray += $projet
        }
    }

    if($errorCount -eq 0){
        return $returnArray
    }
    else {
        return $null
    }

}


#==========================================================================================================================================


function GetElementToCheckList {

    param(
        [string[]]$modelPropertiesToCheckFullList
    )

    #In the list, if we have to put method (like count), dont put '()' into its name
    if (-not $modelPropertiesToCheckFullList){
        $modelPropertiesToCheckFullList = @(
            "model:name"
            ##"model:culture",
            ##"model.dataAccessOptions.fastCombine",
            ##"model.dataAccessOptions.legacyRedirects",
            ##"model.dataAccessOptions.returnErrorValuesAsNull",
            ##"model.defaultPowerBIDataSourceVersion",
            ##"model.discourageImplicitMeasures",
            ##"model.sourceQueryCulture",
            "model:tables",
            "model.tables:lineageTag",
            "model.tables:name",
            "model.tables:isHidden",
            "model.tables:IsPrivate"
            "model.tables:ExcludeFromModelRefresh"
            "model.tables:IsRemoved"
            "model.tables:columns",
            "model.tables:columns.count",
            "model.tables.columns:lineageTag",
            "model.tables.columns:type",
            "model.tables.columns:name",
            "model.tables.columns:expression",
            "model.tables.columns:dataType",
            #"model.tables.columns:isNameInferred",
            "model.tables.columns:sourceColumn",
            "model.tables.columns:formatString",
            "model.tables.columns:isAvailableInMdx",
            "model.tables.columns:summarizeBy.ToString",
            ##"model.tables.columns:annotation SummarizationSetBy",
            ##"model.tables.columns:annotation UnderlyingDateTimeDataType",
            "model.tables.columns:sortByColumn",
            "model.tables.columns:changedProperties.property",
            ##"model.tables.columns:annotations",
            ##"model.tables.columns.annotations:name",
            ##"model.tables.columns.annotations:value",
            "model.tables:partitions",
            "model.tables.partitions:name",
            #"model.tables.partitions:sourceType", #Pour le moment pas utile, mais peut être plus tard
            #"model.tables.partitions:mode", #Que de l'import pour le momen
            "model.tables.partitions:queryGroup",
            "model.tables.partitions.queryGroup:folder",
            #"model.tables.partitions.queryGroup:description",
            "model.tables.partitions:source.expression",
            "model.tables.partitions:kind",
            "model.tables.partitions:expression",
            "model.tables.partitions:table.name",
            "model.tables:changedProperties",
            "model.tables.changedProperties:property",
            ##"model.tables:annotations",
            ##"model.tables.annotations:name",
            ##"model.tables.annotations:value",
            "model.tables:measures",
            "model.tables:measures.count",
            "model.tables.measures:lineageTag",
            "model.tables.measures:name",
            "model.tables.measures:expression",
            "model.tables.measures:formatString",
            "model.tables.measures:isHidden",
            "model.tables.measures:displayFolder",
            "model.tables.measures:hangedProperties.property",
            ##"model.tables.measures:annotations",
            ##"model.tables.measures.annotations:name",
            ##"model.tables.measures.annotations:value",
            "model.tables:calculationgroup",
            "model.tables.calculationgroup:description",
            "model.tables.calculationgroup:calculationitems",
            "model.tables.calculationgroup.calculationitems:name",
            "model.tables.calculationgroup.calculationitems:description",
            "model.tables.calculationgroup.calculationitems:state",
            "model.tables.calculationgroup.calculationitems:expression",
            "model:relationships",
            "model.relationships:name",
            "model.relationships:fromCardinality",
            "model.relationships:toCardinality",
            "model.relationships:fromTable.name",
            "model.relationships:fromColumn.name",
            "model.relationships:toTable.name",
            "model.relationships:toColumn.name",
            "model.relationships:crossFilteringBehavior",
            "model:perspectives",
            "model.perspectives:name",
            "model.perspectives:perspectivetables",
            "model.perspectives.perspectivetables:name",
            "model.perspectives.perspectivetables:perspectivecolumns",
            "model.perspectives.perspectivetables.perspectivecolumns:name",
            "model.perspectives.perspectivetables.perspectivecolumns:perspectivetable.perspective.name",
            "model.perspectives.perspectivetables.perspectivecolumns:perspectivetable.name",
            "model.perspectives.perspectivetables.perspectivecolumns:CalculationGroup.Table.name",
            "model.perspectives.perspectivetables.perspectivecolumns:CalculationGroup.description",
            "model.perspectives.perspectivetables.perspectivecolumns:description",
            "model.perspectives.perspectivetables.perspectivecolumns:state",
            "model.perspectives.perspectivetables.perspectivecolumns:expression",
            "model:roles",
            "model.roles:name",
            ##"model.roles:modelPermission",
            "model.roles:tablePermissions",
            "model.roles.tablePermissions:name",
            "model.roles.tablePermissions:role.name",
            "model.roles.tablePermissions:table.name",
            "model.roles.tablePermissions:filterExpression",
            ##"model.roles:annotations",
            ##"model.roles.annotations:name",
            ##"model.roles.annotations:value",
            "model:expressions"
            "model.expressions:lineageTag",
            "model.expressions:name",
            "model.expressions:kind",
            "model.expressions:expression",
            "model.expressions:queryGroup.description"#,
            ##"model.expressions:annotations",
            ##"model.expressions.annotations:name",
            ##"model.expressions.annotations:value",
            #"model:queryGroups",
            #"model.queryGroups:folder",
            #"model.queryGroups:description"
            ##"model.queryGroups:annotations",
            ##"model.queryGroups.annotations:name",
            ##"model.queryGroups.annotations:value",
            ##"model:annotations",
            ##"model.annotations:name",
            ##"model.annotations:value"
        )
    }


    $modelPropertiesToCheck = @()
    Foreach ($line in $modelPropertiesToCheckFullList) {
        $segments = $line -split ":"
        $modelPropertiesToCheck += [PSCustomObject]@{
            hierarchyLevel = $segments[0..($segments.length - 2)] -join '.'
            element = $segments[-1]
        }
    }

    return $modelPropertiesToCheck

}


function GetElementToTextualyCheck {

    param(
        [string] $objectName,
        [string[]] $elementsList
    )

    if (-not $modelPropertiesTextualCheckFullList){
        $elementsList = @(
            "DataColumn:expression"
            "Partition:source.expression"
            "NamedExpression:expression"
            "Measure:expression"
            "TablePermission:filterExpression"
            "CalculationItem:expression"

            #Specific cases
            "Parameter:ExpressionValue"
            "Parameter:ExpressionMeta"
        )
    }


    $outputList = @()
    Foreach ($line in $elementsList) {
        $segments = $line -split ":"

        $segmentObjectName = $segments[0..($segments.length - 2)] -join '.'

        write-host "Checking $($segmentObjectName) with $($objectName))"
        if($segmentObjectName -ne $objectName){
            continue
        }
        else {
            $outputList += [PSCustomObject]@{
                objectName = $segmentObjectName
                element = $segments[-1]
            }
        }
    }

    return $outputList

}


#Generic function used to get element for a defined Object
function GetElementFromList {

    param (
        [string] $objectName,
        [String[]] $elementsList
    )

    if(-not $elementsList){
        write-host "!No list is defined for the call of GetElementFromList!"
        return $null
    }
    if(-not $objectName){
        write-host "!No object name defined!"
        return $null
    }

    $outputList = @()
    Foreach ($line in $elementsList) {
        $segments = $line -split ":"

        $segmentObjectName = $segments[0..($segments.length - 2)] -join '.'

        if($segmentObjectName -ne $objectName){
            continue
        }
        else {
            $outputList += [PSCustomObject]@{
                objectName = $segmentObjectName
                element = $segments[-1]
            }
        }
    }

    return $outputList
}


#Function in order to change the name of automatized html to better fitting names
#DEPRECIATED DEPRECIATED DEPRECIATED DEPRECIATED DEPRECIATED DEPRECIATED DEPRECIATED
function GetLocalization {
    param (
        [string] $inputString
    )

    $result = ""

    $localizationDB = @{
        'Model.Tables.Columns.lineageTag' = 'lineageTag'
        'Model.Tables.Columns.type' = 'Type'
        'Model.Tables.Columns.name' = 'Name'
        'Model.Tables.Columns.expression' = 'Expression'
        'Model.Tables.Columns.dataType' = 'Data Type'
        'Model.Tables.Columns.isNameInferred' = 'Is Name Inferred'
        'Model.Tables.Columns.sourceColumn' = 'Source Column'
        'Model.Tables.Columns.formatString' = 'Format String'
        'Model.Tables.Columns.isAvailableInMdx' = 'Is Available In MDX'
        'Model.Tables.Columns.summarizeBy.ToString' = 'Summarize By'
        'Model.Tables.Columns.sortByColumn' = 'Sort By Column'
        'Model.Tables.Columns.changedProperties.property' = 'Changed Property'
        'Model.Tables.name' = 'Name'
    'Model.Tables.isHidden' = 'Is Hidden'
        'Model.Tables.IsPrivate' = 'Is Private'
        'Model.Tables.ExcludeFromModelRefresh' = 'Exclude From Model Refresh'
        'Model.Tables.IsRemoved' = 'Is Removed'
        'Model.Tables.columns.count' = 'Column Count'
        'Model.Tables.partitions.name' = 'Partition Name'
        'Model.Tables.changedProperties.property' = 'Changed Property'
        'Model.Tables.measures.count' = 'Measures Count'
        'Model.Tables.partitions.sourceType' = 'Source Type'
        'Model.Tables.partitions.mode' = 'Mode'
        'Model.Tables.partitions.queryGroup.folder' = 'Folder'
        'Model.Tables.partitions.queryGroup.description' = 'Description'
        'Model.Tables.partitions.source.expression' = 'Expression'
        'Model.Tables.partitions.kind' = 'Kind'
        'Model.Tables.partitions.queryGroup' = 'Query Group'
        'Model.Tables.partitions.expression' = 'Expression'
        "Model.Tables.partitions.table.name" = 'Table Name'
        'Model.relationships.name' = 'Name'
        'Model.relationships.fromCardinality' = 'From Cardinality'
        'Model.relationships.fromTable.name' = 'From Table'
        'Model.relationships.fromColumn.name' = 'From Column'
        'Model.relationships.toCardinality' = 'To Cardinality'
        'Model.relationships.toTable.name' = 'To Table'
        'Model.relationships.toColumn.name ' = 'To Column'
        'Model.relationships.crossFilteringBehavior' = 'Cross Filtering Behavior'
        'Model.Tables.Measures.name' = 'Name'
        'Model.Tables.Measures.expression' = 'Expression'
        'Model.Tables.Measures.formatString' = 'Format String'
        'Model.Tables.Measures.isHidden' = 'Is Hidden'
        'Model.Tables.Measures.displayFolder' = 'Display Folder'
    'Model.Tables.Measures.hangedProperties.property' = 'Property Changed'
        'Model.roles.TablePermissions.role.name' = 'Role Name'
        'Model.roles.TablePermissions.table.name' = 'Table Name'
        'Model.roles.TablePermissions.name' = 'Name'
        'Model.roles.TablePermissions.filterExpression' = 'Filter Expression'
        'Model.Perspectives.PerspectiveTables.PerspectiveColumns.PerspectiveTable.Perspective.name' = 'Perspective'
        'Model.Perspectives.PerspectiveTables.PerspectiveColumns.PerspectiveTable.name' = 'Name'
        'Model.Perspectives.PerspectiveTables.PerspectiveColumns.name' = 'Calculation Item Name'
        'Model.Perspectives.PerspectiveTables.PerspectiveColumns.CalculationGroup.Table.name' = 'Calculation Group Name'
        'Model.Perspectives.PerspectiveTables.PerspectiveColumns.CalculationGroup.description' = 'Calculation Group Desc.'
        'Model.Perspectives.PerspectiveTables.PerspectiveColumns.description' = 'Calculation Item Desc.'
        'Model.Perspectives.PerspectiveTables.PerspectiveColumns.state' = 'State'
        'Model.Perspectives.PerspectiveTables.PerspectiveColumns.expression' = 'Expression'

        'Model.relationships.fromCardinality.ToString.toCardinality.ToString' = 'Relationships cardinality'
        'Model.relationships.fromTable.name.fromColumn.name' = 'From table.column'
        'Model.relationships.toTable.name.toColumn.name' = 'To table.column'


        #Specific
        'Specific.parameter.name' = 'Name'
        'Specific.parameter.kind' = 'Kind'
        'Specific.parameter.queryGroup.folder' = 'Folder'
        'Specific.parameter.ExpressionValue' = 'Value'
        'Specific.parameter.ExpressionMeta' = 'Expression meta'
    }

    $result = $localizationDB.$inputString

    if (-not $result){
        $inputString
    }

    return $result

}








Function GetDiffOutputObject {
    param (
        $objectNew,
        $objectOld,
        $hierarchyNewVersion,
        $hierarchyOldVersion,
        $path,
        $type,
        $result,
        [bool] $difference
    )

    write-host "...writing difference output..."

    $result = [PSCustomObject]@{
        ObjectNewVersion = $objectNew
        ObjectOldVersion = $objectOld
        HierarchyNewVersion = $hierarchyNewVersion
        HierarchyOldVersion = $hierarchyOldVersion
        Path = $path
        Type = $type
        Result = $result
        Difference = $difference
    }

    return $result

}


#==========================================================================================================================================


#Function to compute the distance of Levenshtein
function GetLevenshteinDistance {
    param (
        [string] $lineNew,
        [string] $lineOld
    )

    $lengthNew = $lineNew.Length
    $lengthOld = $lineOld.Length

    $matrix = @()
    for ($i = 0; $i -le $lengthNew; $i++) {
        $row = @()
        for ($j = 0; $j -le $lengthOld; $j++) {
            $row += 0
        }
        $matrix += ,$row
    }

    for ($i = 0; $i -le $lengthNew; $i++) {
        $matrix[$i][0] = $i
    }
    for ($j = 0; $j -le $lengthOld; $j++) {
        $matrix[0][$j] = $j
    }

    for ($i = 1; $i -le $lengthNew; $i++) {
        for ($j = 1; $j -le $lengthOld; $j++) {
            if ($lineNew[$i - 1] -eq $lineOld[$j - 1]) {
                $cost = 0
            } else {
                $cost = 1
            }

            $matrix[$i][$j] = [Math]::Min([Math]::Min($matrix[$i - 1][$j] + 1, $matrix[$i][$j - 1] + 1), $matrix[$i - 1][$j - 1] + $cost)
        }
    }

    return $matrix[$lengthNew][$lengthOld]
}


#Function created in order to be able to find, in a text, where a difference is present
function CompareStringElement {
    
    param (
        $textNew,
        $textOld,
        $distanceTolerance = 0.5
    )

    $diffResult = @()

    write-host "==Textual comparison=="

    $linesNew = $textNew -split "`n"
    $linesOld = $textOld -split "`n"

    #Creation of list in order to modify the treated text differently than the original one (addition of space for gaps)
    $techLinesNew = [System.Collections.Generic.List[string]]::new()
    $techLinesNew.AddRange($linesNew)
    $techLinesOld = [System.Collections.Generic.List[string]]::new()
    $techLinesOld.AddRange($linesOld)

    $maxLinesCount = [Math]::Max($linesNew.Count, $linesOld.Count)

    #Check all of the lines of both files
    for ($i = 0; $i -lt $maxLinesCount; $i++) {

        $currentLineNew = if($techLinesNew[$i]) { $techLinesNew[$i] } else {$null}
        $currentLineOld = if($techLinesOld[$i]) { $techLinesOld[$i] } else {$null}
        write-host "Lines checked : new ($($currentLineNew)) vs old ($($currentLineOld))"
        
        #First, check if the two lines exists
        if ($currentLineNew -and $currentLineOld) {
            #Check if the two lines are equal
            if($currentLineNew -eq  $currentLineOld) {
            write-host "No difference found on this line"
            $diffResult += New-Object PSObject -Property @{
                Index = $i
                TexteNew = $techLinesNew[$i]
                TexteOld = $techLinesOld[$i]
                Type     = "No change"
            }
            continue
            }
            else
            {
                $distance = GetLevenshteinDistance -lineNew $currentLineNew -lineOld $currentLineOld
                $distancePerCent = $distance / [Math]::Max($currentLineNew.Length, $currentLineOld.Length)

                write-host "The difference between the two lines is of $($distancePerCent), the step is $($distanceTolerance)"

                #If the distance is lesser than the defined step, then the two lines are considered as the same line but modified
                #If the size of the original text is 1, we consider that the line is modified whatever the distance
                if($distancePerCent -lt $distanceTolerance -or $maxLinesCount -eq 1) {
                    
                    if($maxLinesCount -ne 1){
                        write-host 'The difference is not enought to consider the two lines as trully different'
                    }else{
                        write-host 'Only one line in the original text, the line is considered modified'
                    }

                    $diffResult += New-Object PSObject -Property @{
                        Index = $i
                        TexteNew = $techLinesNew[$i]
                        TexteOld = $techLinesOld[$i]
                        Type     = "Modified"
                    }
                    continue

                }
                #Else we can't say where the lines are from
                #The line can be a pure addition, or it was displaced in the text
                else{
                    
                    write-host 'The difference is large enought to see the lines as different'

                    #Check if we can find the current lines in the other version remaining lines
                    $bNewInOld = $techLinesOld[$i..($techLinesOld.Count)].Contains($currentLineNew)
                    $bOldInNew = $techLinesNew[$i..($techLinesNew.Count)].Contains($currentLineOld)

                    if( $bNewInOld -and $bOldInNew ){
                        
                        #If both lines are found, that's mean that it is difficult to know which one of the two lines are the new or deleted
                        #So we will check where is the doppelganger line, the one with the farest is considered as the different line

                        $indexNewInOld = $i+1
                        $indexOldInNew = $i+1

                        write-host "Index new in old ($($indexNewInOld))"
                        while ($currentLineNew -ne $techLinesOld[$indexNewInOld] -or (-not $techLinesOld[$indexNewInOld])){
                            $indexNewInOld++
                            write-host "Index new in old ($($indexNewInOld))"
                        }
                        write-host "Index old in new ($($indexOldInNew))"
                        while ($currentLineOld -ne $techLinesNew[$indexOldInNew] -or (-not $techLinesNew[$indexOldInNew])){
                            $indexOldInNew++
                            write-host "Index old in new ($($indexOldInNew)) cheking ($($currentLineOld)) with ($($techLinesNew[$indexOldInNew]))"
                        }

                        #TODO If the index is outside the text (shouldn't appear), treat it

                        write-host "Cheking if the Index new in old ($($indexNewInOld)) is lesser than Index old in new ($($indexOldInNew))"
                        #if the new line in the old version is nearer of the new line in the new version than the old line in the new version, the new line is considered as different
                        if($indexNewInOld -gt $indexOldInNew -and $indexNewInOld -and $indexOldInNew){
                            #New line is different
                            write-host "The new line is different"
                            $techLinesOld.Insert($i,"")
                            $diffResult += New-Object PSObject -Property @{
                                Index = $i
                                TexteNew = $techLinesNew[$i]
                                TexteOld = $techLinesOld[$i]
                                Type     = "Added"
                            }
                        }
                        #Else is the other way around
                        else{
                            #old line is different
                            write-host "The old line is different"
                            $techLinesNew.Insert($i,"")
                            $diffResult += New-Object PSObject -Property @{
                                Index = $i
                                TexteNew = $techLinesNew[$i]
                                TexteOld = $techLinesOld[$i]
                                Type     = "Deleted"
                            }
                        }



                    }
                    elseif(-not $bNewInOld -and $bOldInNew) {
                        #If the old line exists in the new version but the new line don't exist in the old version, the new line is considered as a new line
                        write-host "The new line is different"
                        $techLinesOld.Insert($i,"")
                        $diffResult += New-Object PSObject -Property @{
                            Index = $i
                            TexteNew = $techLinesNew[$i]
                            TexteOld = $techLinesOld[$i]
                            Type     = "Added"
                        }

                    }
                    elseif($bNewInOld -and (-not $bOldInNew)) {
                        #If the new line exist in the old version but the old line don't exist in the new version, the old line is considered as a deleted line
                        write-host "The old line is different"
                        $techLinesNew.Insert($i,"")
                        $diffResult += New-Object PSObject -Property @{
                            Index = $i
                            TexteNew = $techLinesNew[$i]
                            TexteOld = $techLinesOld[$i]
                            Type     = "Deleted"
                        }

                    }
                    else {
                        #If both lines are not found in the other version, the old was deleted and the new added
                        write-host "Both lines are different"
                        $techLinesOld.Insert($i,"")
                        $diffResult += New-Object PSObject -Property @{
                            Index = $i
                            TexteNew = $techLinesNew[$i]
                            TexteOld = $techLinesOld[$i]
                            Type     = "Added"
                        }
                        $i++
                        $techLinesNew.Insert($i,"")
                        $diffResult += New-Object PSObject -Property @{
                            Index = $i
                            TexteNew = $techLinesNew[$i]
                            TexteOld = $techLinesOld[$i]
                            Type     = "Deleted"
                        }

                    }
                }

            }
        }
        #Else, thats mean that we are off limit for one of the two texts (with the addition of blanck lines for missing ones, this shouldn't happen exept for the last one)
        elseif (-not $currentLineNew -and $currentLineOld) {
            write-host "The old line is different because the new is null"
            $diffResult += New-Object PSObject -Property @{
                Index = $i
                TexteNew = $techLinesNew[$i]
                TexteOld = $techLinesOld[$i]
                Type     = "Deleted"
            }

        }
        #Else the line didn't exist for the old version (so the line is added by the new one)
        elseif (-not $currentLineOld -and $currentLineNew) {
            write-host "The new line is different because the old is null"
            $diffResult += New-Object PSObject -Property @{
                Index = $i
                TexteNew = $techLinesNew[$i]
                TexteOld = $techLinesOld[$i]
                Type     = "Added"
            }
        }
        #Else both lines are null (so we don't care) or something wrong happenend
        else{
            write-host "No interesting thing was found at index $i so we do nothing" -ForegroundColor DarkGray
        }
    }

    $textResult = @(
        $techLinesNew -join "`n"
        $techLinesOld -join "`n"
    )

    #We return first the differences messages, and then both text with space where missing lines exists
    return $diffResult, $textResult

}



#==========================================================================================================================================


#Get the value of the propertie searched for
function GetPropertyValue {
    param (
        $element, #One element
        $propertyPath #One property
    )
    $properties = $propertyPath -split '\.'
    $currentObject = $element
    foreach ($property in $properties) {
        if ($null -eq $currentObject) {
            return $null
        }
        if ( (-not $currentObject.PSObject.Methods[$property]) -and $currentObject.PSObject.Properties[$property]) {
            $currentObject = $currentObject.$property
        } 
        #If it is a function, we have to call it differently than a property
        elseif ($currentObject.PSObject.Methods[$property]) {
            $currentObject = $currentObject.$property()
        }
        else {
            return $null
        }
    }
    return $currentObject
}


#Function created in order to get the value of elements we want to check
function GetCheckableElementsList {
    param (
        $elementList,
        [string[]] $propertyToCheckList
    )

    # The first property of the list is considered as the key for the check between version
    $keyProperty = $propertyToCheckList[0]

    $returnPropertiesValues = @{}
    $returnObjects = @{}

    foreach ($element in $elementList) {

        $keyValue = GetPropertyValue -element $element -propertyPath $keyProperty
        if ($null -eq $keyValue) { 
            write-host "No value was found as key for this element!"
            continue 
        }

        $currentProperties = [PSCustomObject]@{}
        foreach ($property in $propertyToCheckList) {

            $value = GetPropertyValue -element $element -propertyPath $property
            
            $currentProperties | Add-Member -NotePropertyName $property -NotePropertyValue $value
        }

        
        #if (-not $returnPropertiesValues.ContainsKey($keyValue)) {
        #    $returnPropertiesValues[$keyValue] = [PSCustomObject]@{}
        #}
        #else {
        #    write-host "!!!! Trying to write properties for the $($keyValue) element, but it already exists !!!!"
        #}

        if (-not $returnPropertiesValues.ContainsKey($keyValue)) {
            $returnPropertiesValues[$keyValue] = $currentProperties
        }
        else {
            write-host "!!!! Trying to write properties for the $($keyValue) element, but it already exists !!!!"
        }
        
        #foreach ($property in $propertyToCheckList) {
        #    $returnPropertiesValues[$keyValue] | Add-Member -NotePropertyName $property -NotePropertyValue ($currentProperties | Select-Object -ExpandProperty $prop) -Force
        #}

        if (-not $returnObjects.ContainsKey($keyValue)) {
            #$returnObjects[$keyValue] = @()
            $returnObjects[$keyValue] = $element
        }
        else {
            write-host "!!!! Trying to write element for the $($keyValue) element, but it already exists !!!!"
        }
        #$returnObjects[$keyValue] += $element
    }

    return $returnPropertiesValues, $returnObjects

}


#Get the list of the different element of the table, and put in light the differences that can exists between the two versions
Function CheckDifferenceInSubElement {
    
    param (
        [PSCustomObject] $element,
        [string] $elementChecked,
        [string] $hierarchyLevel,
        [PSCustomObject[]] $hierarchies,
        [bool] $getAll = $false
    )

    
    function CheckIfSubElementCanBeChecked {
        param (
            $hierarchyLevel
        )
        #If no element to check list is given, then we get it from the global list from the hierarchy
        $subElementList = @()
        write-host "trying to know if there is element to check for $($hierarchyLevel + '.' + $subElmt)"
        GetElementToCheckList | Where-Object hierarchyLevel -eq $hierarchyLevel | Select-Object element | ForEach-Object { $subElementList += $_.element }

        #If no element are to be checked, stop here and return nothing
        write-host "Checking let us with $($subElementList) as sub element to check"
        return $subElementList #Will be null if no element is found
    }

    function LaunchRecursiveFunction {
        param (
            $valNew,
            $valOld,
            $pathNewVersion,
            $pathOldVersion,
            $elementChecked,
            $hierarchyLevel,
            $hierarchies
        )

        if (CheckIfSubElementCanBeChecked -hierarchyLevel ($hierarchyLevel + '.' + $subElmt)) {

            write-host "=Sub elements found to be checked (for $($hierarchyLevel)), let's do that!="

            $elementNext = @{}
            $elementNext = [PSCustomObject]@{
                ElementNewVersion = $valNew
                ElementOldVersion = $valOld
                PathNewVersion = $pathNewVersion
                PathOldVersion = $pathOldVersion
            }

            $result = CheckDifferenceInSubElement -element $elementNext -elementChecked $subElmt -hierarchyLevel $hierarchyLevel -hierarchies @($currentNewHierarchy, $currentOldHierarchy)

        }
        else{
            write-host "°°°No element to check here, so we continue°°°"
        }

        if($result){return $result}

    }


    write-host "===Check from CheckDifferenceInSubElement called!===" 

    $diffResult = @()

    write-host "We are checking at the level $($hierarchyLevel) for $($elementChecked)"

    #Get a value to know at which level in the model we found the difference
    if(-not $hierarchyLevel){
        $hierarchyLevel = $elementChecked
    }
    else{
        $hierarchyLevel = $hierarchyLevel + '.' + $elementChecked
    }

    write-host "We will look for $($hierarchyLevel)"

    #If no element to check list is given, then we get it from the global list from the hierarchy
    $subElementList = CheckIfSubElementCanBeChecked -hierarchyLevel $hierarchyLevel

    #If no element are to be checked, stop here and return nothing
    if($subElementList.Count -eq 0){
        write-host "°°°No element to check, so we end the check here°°°"
        return
    }

    write-host "Cheking for the following elements : $($subElementList)"

    $modelElementNew = $element.ElementNewVersion
    write-host "modelElementNew: $($modelElementNew)"
    $modelElementOld = $element.ElementOldVersion
    write-host "modelElementOld: $($modelElementOld)"
    $pathElementNew = $element.PathNewVersion
    write-host "pathElementNew: $($pathElementNew)"
    $pathElementOld = $element.PathOldVersion
    write-host "pathElementOld: $($pathElementOld)"

    #--------------

    write-host "Getting the GetCheckableElementsList for new version"
    $newVersionDicts = GetCheckableElementsList -elementList $modelElementNew -propertyToCheckList $subElementList

    $newVersionDictonnary = $newVersionDicts[0]
    $newVersionObjects = $newVersionDicts[1]

    write-host "Getting the GetCheckableElementsList for old version"
    $oldVersionDicts = GetCheckableElementsList -elementList $modelElementOld -propertyToCheckList $subElementList

    $oldVersionDictonnary = $oldVersionDicts[0]
    $oldVersionObjects = $oldVersionDicts[1]

    write-host "Creating the key dictionnary"
    $keysProperties = $newVersionDictonnary.Keys + $oldVersionDictonnary.Keys | Select-Object -Unique
    write-host "keysProperty: $($keysProperty)"


    #---
    #Here we initialise the hierarchies
    if(-not $hierarchies){
        $newHierarchy = [PSCustomObject]@{}
        $oldHierarchy = [PSCustomObject]@{}
    }
    else{
        $newHierarchy = $hierarchies[0]
        $oldHierarchy = $hierarchies[1]
    }
    #---


    #Since we get all the 'Key' for the element to compare, we check
    ForEach ($key in $keysProperties){

        write-host "Checking for the '$($key)' key (for $($elementChecked))"


        #Update hierarchies (adress of the previous objects serached in the recusrive search)
        if ($newVersionObjects[$key]){
            $currentNewHierarchy = $newHierarchy.PSObject.Copy() 
            $currentNewHierarchy | Add-Member -NotePropertyName $($hierarchyLevel) -NotePropertyValue $newVersionObjects[$key]
            write-host "New version : Updating the hierarchy of object from the following $($newHierarchy)"
            write-host "to $($currentNewHierarchy)"
        }
        else {
            $currentNewHierarchy = $newHierarchy
        }
        if ($oldVersionObjects[$key]){
            $currentOldHierarchy = $oldHierarchy.PSObject.Copy() 
            $currentOldHierarchy | Add-Member -NotePropertyName $($hierarchyLevel) -NotePropertyValue $oldVersionObjects[$key]
            write-host "Old version : Updating the hierarchy of object from the following $($oldHierarchy)"
            write-host "to $($currentOldHierarchy)"
        }
        else{
            $currentOldHierarchy = $oldHierarchy
        }



        #Firstly check if the key is valid for both versions
        if($newVersionDictonnary[$key] -and $oldVersionDictonnary[$key]){
            write-host "The key ($($key)) exists in both version"

            Foreach ($subElmt in $subElementList){

                write-host "Checking '$($subElmt)'"
                write-host "$($newVersionDictonnary[$key].GetType())"
                write-host "$newVersionDictonnary[$key].$subElmt"

                $valNew = $newVersionDictonnary[$key].$subElmt
                #write-host "New value ($($valNew))"
                $valOld = $oldVersionDictonnary[$key].$subElmt
                #write-host "Old value ($($valOld))"

                write-host "We have the new value of '$($valNew)' and old value of '$($valOld)'"

                #Check if the values are null
                if($null -eq $valNew -and $null -eq $valOld){
                    write-host "Both values are null, so no difference"

                    if($getAll){
                        $diffResult += GetDiffOutputObject `
                            -objectNew $newVersionObjects[$key] `
                            -objectOld $oldVersionObjects[$key] `
                            -hierarchyNewVersion $currentNewHierarchy `
                            -hierarchyOldVersion $currentOldHierarchy `
                            -path "$($pathElementNew).$($key)" `
                            -type "Both values null" `
                            -result "No difference : new value ($($valNew)) compared to old value ($($valOld))" `
                            -difference $false
                    }

                    continue #equal so we don't report it as a difference
                }
                #check if one of them is null
                elseif($null -eq $valNew -or $null -eq $valOld){
                    Write-Host "/!\ one of the value is null detection /!\"
                    
                    #If the new value is null, we will still try to check for the old elements as deleted
                    if($null -eq $valNew){
                        write-host "new is null, so we consider that old is deleted"

                        $diffResult += GetDiffOutputObject `
                            -objectNew $null `
                            -objectOld $oldVersionObjects[$key] `
                            -hierarchyNewVersion $null `
                            -hierarchyOldVersion $currentOldHierarchy `
                            -path "$($pathElementOld).$($key)" `
                            -type "Element removed in new version" `
                            -result "$($elementChecked) $($key) removed (the new value is null)."`
                            -difference $true

                        #We check then the next set of properties
                        write-host "Checking the next set of properties, with null as new"

                        $diffResult += LaunchRecursiveFunction `
                            -valNew $null `
                            -valOld $valOld `
                            -pathNewVersion $null `
                            -pathOldVersion "$($pathElementOld).$($key)" `
                            -elementChecked $subElmt `
                            -hierarchyLevel $hierarchyLevel `
                            -hierarchies @($null, $currentOldHierarchy)`

                        continue
                    }
                    #else = old is null and new is considered as an addition
                    else{
                        write-host "old is null, so we consider that new is an addition"

                        $diffResult += GetDiffOutputObject `
                            -objectNew $newVersionObjects[$key] `
                            -objectOld $null `
                            -hierarchyNewVersion $currentNewHierarchy `
                            -hierarchyOldVersion $null `
                            -path "$($pathElementNew).$($key)" `
                            -type "Element added in new version" `
                            -result "$($elementChecked) $($key) added (the old value is null)."`
                            -difference $true

                        #We check then the next set of properties
                        write-host "Checking the next set of properties, with null as old"

                        $diffResult += LaunchRecursiveFunction `
                            -valNew $valNew `
                            -valOld $null `
                            -pathNewVersion "$($pathElementNew).$($key)" `
                            -pathOldVersion $null `
                            -elementChecked $subElmt `
                            -hierarchyLevel $hierarchyLevel `
                            -hierarchies @($currentNewHierarchy, $null)`

                        continue
                    }
                }
                #Else both are not null
                else{
                    #Check depending of the type of the two values
                    if ($valNew.GetType() -eq $valOld.GetType()){

                        write-host "The type of the two values is equal"

                        if($valNew.GetType().isPrimitive -or $valNew -is [string] -or $valNew -is [DateTime]){

                            write-host "This type can be checked as it is ($($valNew.GetType()))"
                        
                            if ($valNew -eq $valOld){
                                Write-Host "no difference detected"

                                if($getAll){
                                    $diffResult += GetDiffOutputObject `
                                        -objectNew $newVersionObjects[$key] `
                                        -objectOld $oldVersionObjects[$key] `
                                        -hierarchyNewVersion $currentNewHierarchy `
                                        -hierarchyOldVersion $currentOldHierarchy `
                                        -path "$($pathElementNew).$($key)" `
                                        -type "Difference in value" `
                                        -result "No difference : new value ($($valNew)) compared to old value ($($valOld))" `
                                        -difference $false
                                }

                                continue #No difference here, move along
                            }
                            #Else = both values exists but they are different between versions
                            else {
                                Write-Host "/!\ value difference detection /!\"

                                $diffResult += GetDiffOutputObject `
                                    -objectNew $newVersionObjects[$key] `
                                    -objectOld $oldVersionObjects[$key] `
                                    -hierarchyNewVersion $currentNewHierarchy `
                                    -hierarchyOldVersion $currentOldHierarchy `
                                    -path "$($pathElementNew).$($key)" `
                                    -type "Difference in value" `
                                    -result "Difference : new value ($($valNew)) compared to old value ($($valOld))"`
                                    -difference $true
                            }

                        }
                        else
                        {
                            #If we don't have a primitive type, we do the same comparison with the element to test this time

                            write-host "This type can't be checked as it is ($($valNew.GetType())), so we redo the check with its sub elements"

                            $diffResult += LaunchRecursiveFunction `
                                -valNew $valNew `
                                -valOld $valOld `
                                -pathNewVersion "$($pathElementNew).$($key)" `
                                -pathOldVersion "$($pathElementOld).$($key)" `
                                -elementChecked $subElmt `
                                -hierarchyLevel $hierarchyLevel `
                                -hierarchies @($currentNewHierarchy, $currentOldHierarchy)`

                        }
                    }
                    #Else is for different type between the two values
                    #Probably not complete here, we should probably check if there are other elements that can be check under one or the other version
                    else {

                        Write-Host "/!\ different types detection /!\"

                        $diffResult += GetDiffOutputObject `
                        -objectNew $newVersionObjects[$key] `
                        -objectOld $oldVersionObjects[$key] `
                        -hierarchyNewVersion $currentNewHierarchy `
                        -hierarchyOldVersion $currentOldHierarchy `
                        -path "$($pathElementNew).$($key)" `
                        -type "Difference in type" `
                        -result "Difference : new value ($($valNew)) type ($($valNew.GetType())) compared to old value ($($valOld)) type ($($valOld.GetType()))"`
                        -difference $true

                        #We check then the next set of properties
                        write-host "Checking the next set of properties"

                        $diffResult += LaunchRecursiveFunction `
                            -valNew $valNew `
                            -valOld $valOld `
                            -pathNewVersion "$($pathElementNew).$($key)" `
                            -pathOldVersion "$($pathElementOld).$($key)" `
                            -elementChecked $subElmt `
                            -hierarchyLevel $hierarchyLevel `
                            -hierarchies @($currentNewHierarchy, $currentOldHierarchy)`
                    }
                }
            }
        }
        #Else if the key not present in the old version (probably a case more often found than the other way around)
        elseif (-not $oldVersionDictonnary[$key]){
            Write-Host "/!\ addition detection /!\"

            $diffResult += GetDiffOutputObject `
                -objectNew $newVersionObjects[$key] `
                -objectOld $null `
                -hierarchyNewVersion $currentNewHierarchy `
                -hierarchyOldVersion $null `
                -path "$($pathElementNew).$($key)" `
                -type "Element added in new version" `
                -result "$($elementChecked) $($key) added."`
                -difference $true

                #Then we will check in the next set of property for the new version
                Foreach ($subElmt in $subElementList){

                    write-host "Checking '$($subElmt)'"
                    write-host "$($newVersionDictonnary[$key].GetType())"
                    write-host "$newVersionDictonnary[$key].$subElmt"

                    $valNew = $newVersionDictonnary[$key].$subElmt
                    #write-host "New value ($($valNew))"
                    $valOld = $null
                    #write-host "Old value ($($valOld))"

                    write-host "We have the new value of '$($valNew)' and old value of '$($valOld)'"

                    #We check then the next set of properties
                    write-host "Checking the next set of properties, with null as old"

                    $diffResult += LaunchRecursiveFunction `
                        -valNew $valNew `
                        -valOld $null `
                        -pathNewVersion "$($pathElementNew).$($key)" `
                        -pathOldVersion $null `
                        -elementChecked $subElmt `
                        -hierarchyLevel $hierarchyLevel `
                        -hierarchies @($currentNewHierarchy, $null)`
                }

        }
        #Else (meaning that the key is not present in the new version)
        else{
            Write-Host "/!\ deletion detection /!\"

            $diffResult += GetDiffOutputObject `
                -objectNew $null `
                -objectOld $oldVersionObjects[$key] `
                -hierarchyNewVersion $null `
                -hierarchyOldVersion $currentOldHierarchy `
                -path "$($pathElementOld).$($key)" `
                -type "Element removed in new version" `
                -result "$($elementChecked) $($key) removed."`
                -difference $true

            #Then we will check in the next set of property for the old version
            Foreach ($subElmt in $subElementList){

                write-host "Checking '$($subElmt)'"
                write-host "$($oldVersionDictonnary[$key].GetType())"
                write-host "$oldVersionDictonnary[$key].$subElmt"

                $valNew = $null
                #write-host "New value ($($valNew))"
                $valOld = $oldVersionDictonnary[$key].$subElmt
                #write-host "Old value ($($valOld))"

                write-host "We have the new value of '$($valNew)' and old value of '$($valOld)'"

                #We check then the next set of properties
                write-host "Checking the next set of properties, with null as new"

                $diffResult += LaunchRecursiveFunction `
                    -valNew $null `
                    -valOld $valOld `
                    -pathNewVersion $null `
                    -pathOldVersion "$($pathElementOld).$($key)" `
                    -elementChecked $subElmt `
                    -hierarchyLevel $hierarchyLevel `
                    -hierarchies @($null, $currentOldHierarchy)`
            }
        }
    }

    return $diffResult;
}


#==========================================================================================================================================


function GetObjectToShow {

    param (
        $comparisonResult,
        [string[]] $objectToGet,
        [string] $filter
    )

    $listOfObject = @()
    $positiveResults = $comparisonResult | Where-Object Difference -eq $true

    ForEach ($result in $positiveResults) {
    
        $newHierarchies = $result.HierarchyNewVersion
        $oldHierarchies = $result.HierarchyOldVersion

        if($result.ObjectNewVersion) { $newType = $result.ObjectNewVersion.GetType().Name } else { $newType = $null }
        if($result.ObjectOldVersion) { $oldType = $result.ObjectOldVersion.GetType().Name } else { $oldType = $null }

        $currentResultOk = $false

        #Check for tables
        if ($objectToGet){
            #if ($newHierarchies[-1]=$($objectToGet) -or $oldHierarchies[-1].$($objectToGet)){
            if ($newType -in $objectToGet -or $oldType -in $objectToGet){
                $currentResult = [PSCustomObject] @{
                    type = if($newType){$newType}else{$oldType} #normally, the type should be the same for the new and old object (one of them can be null hovewer)
                    #new = $newHierarchies.$($objectToGet)
                    #old = $oldHierarchies.$($objectToGet)
                    new = $result.ObjectNewVersion
                    old = $result.ObjectOldVersion
                }
            }
            else { continue }
        }
        else {

            if ($newType -eq $oldType) {
                $typeToKeep = $newType
            }
            elseif ($newType -and (-not $oldType)) {
                $typeToKeep = $newType
            }
            elseif ($oldType -and (-not $newType)) {
                $typeToKeep = $oldType
            }
            else {
                $typeToKeep = "$($newType)|$($oldType)" 
            }

            $currentResult = [PSCustomObject] @{
                type = $typeToKeep
                #new = $newHierarchies.$($objectToGet)
                #old = $oldHierarchies.$($objectToGet)
                new = $result.ObjectNewVersion
                old = $result.ObjectOldVersion
            }
        }

        #Check if the element we want to get is already added in the result
        $checkExists = $listOfObject | Where-Object {
            $_.type -eq $currentResult.type -and
            $_.new -eq $currentResult.new -and
            $_.old -eq $currentResult.old
        }

        #If the object is not already added in the result, and if the result correspond to the filter we can add in param
        if(-not $checkExists)
        {
            if(-not $filter){
                $currentResultOk = $true
            }
            else{
                #Check the possible filter on the results
                switch($filter){
                    'Parameters' {if((($result.ObjectNewVersion | ForEach-Object { $_.QueryGroup } | Select-Object Folder -Unique).Folder -contains "DMT_GLOBAL PARAMETERS" -or ($result.ObjectOldVersion | ForEach-Object { $_.QueryGroup } | Select-Object Folder -Unique).Folder -contains "DMT_GLOBAL PARAMETERS")`
                        -or (($result.ObjectNewVersion | ForEach-Object { $_.QueryGroup } | Select-Object Folder -Unique).Folder -contains "Report Parameters" -or ($result.ObjectOldVersion | ForEach-Object { $_.QueryGroup } | Select-Object Folder -Unique).Folder -contains "Report Parameters")) {$currentResultOk = $true}}
                    'Not Parameters' {if( -not ((($result.ObjectNewVersion | ForEach-Object { $_.QueryGroup } | Select-Object Folder -Unique).Folder -contains "DMT_GLOBAL PARAMETERS" -or ($result.ObjectOldVersion | ForEach-Object { $_.QueryGroup } | Select-Object Folder -Unique).Folder -contains "DMT_GLOBAL PARAMETERS")`
                        -or (($result.ObjectNewVersion | ForEach-Object { $_.QueryGroup } | Select-Object Folder -Unique).Folder -contains "Report Parameters" -or ($result.ObjectOldVersion | ForEach-Object { $_.QueryGroup } | Select-Object Folder -Unique).Folder -contains "Report Parameters"))) {$currentResultOk = $true}}
                    default { 
                        Write-Host "Filter asked but none condition found in the script."
                    }
                }
            }
            if($currentResultOk -eq $true){
                #Check for the tables where a difference is seen
                $listOfObject += $currentResult
            }
        }

    }

    return $listOfObject
}


#Function used to create a table of comparison between two versions of a string element
function ReportCreateStringCompareTable {

    param (
        [String] $newString,
        [String] $oldString,
        [String] $tableName
    )

    $comparisonResult = CompareStringElement -textNew $newString -textOld $oldString
    $comparisonResultTable = $comparisonResult[0]

    #Initialization of the table
    $compareTable = "<table class=`"modal-table`">`n"

    #Table headers
    $compareTable += "<thead>"
    #$compareTable += "<th data-i18n-key=`"semantic.tables.headers.line_number`">semantic.tables.headers.line_number</th>"
    $compareTable += "<th data-i18n-key=`"semantic.tables.headers.old_version`">semantic.tables.headers.old_version</th>"
    $compareTable += "<th data-i18n-key=`"semantic.tables.headers.status`">semantic.tables.headers.status</th>"
    $compareTable += "<th data-i18n-key=`"semantic.tables.headers.new_version`">semantic.tables.headers.new_version</th>"
    $compareTable += "</thead>`n"


    #Creation of the HTML table
    forEach($line in $comparisonResultTable){

        #Get some style for the lines if they are added, deleted or modified
            if($line.Type -eq "Added"){
                $additionalStyle = " style='background-color: var(--tbl-row-added);'"
            }
            elseif($line.Type -eq "Deleted"){
                $additionalStyle = " style='background-color: var(--tbl-row-removed);'"
            }
            elseif($line.Type -eq "Modified"){
                $additionalStyle = " style='background-color: var(--tbl-cell-changed);'"
            }
            else{
                $additionalStyle = $null
            }

        #For each line, lets get their line in the table :
        $compareLine = "<tr>"
        #$compareLine += "<td$($additionalStyle)>$($line.index)</td>"
        $compareLine += "<td$($additionalStyle)>$($line.TexteOld)</td>"

        #Adding the type of the line
        $compareLine += "<td$($additionalStyle)>"
        if($line.Type -eq "Added"){
            $compareLine += "<div class=`"capsule statut-ok`" data-i18n-key=`"diffTypes.Ajoute`">"
        }
        elseif($line.Type -eq "Deleted"){
            $compareLine += "<div class=`"capsule statut-alerte`" data-i18n-key=`"diffTypes.Supprime`">"
        }
        elseif($line.Type -eq "Modified"){
            $compareLine += "<div class=`"capsule statut-different`" data-i18n-key=`"diffTypes.Modifie`">"
        }
        else{
            $compareLine += "<div class=`"capsule statut-identical`" data-i18n-key=`"diffTypes.Identique`">"
        }
        $compareLine += "$($line.Type)</div></td>"

        $compareLine += "<td$($additionalStyle)>$($line.TexteNew)</td>"
        $compareLine += "</tr>`n"

        #Adding the line to the table
        $compareTable += $compareLine
    }

    #End the table
    $compareTable += "</table>`n"

    return $compareTable
}


#Main function in order to create tables for the HTML report
function ReportCreateCompareTable {

    param (
        $listOfObjects, #This list should have pointers to the object we want in the new and old
        [string] $objectsType,
        [string] $tableName,
        [object] $listOfProperties
    )

    #SPECIFIC CASES
    #=============================

    #Parameters
    #-------------
    if ($objectsType -eq 'Specific.parameter') {

        #Transform the 
        Foreach($object in $listOfObjects){

            
            
            #For each NamedExpression object, we replace it with a custom class
            if($object.type -ne 'NamedExpression'){
                continue
            }
            else{
                
                $object.Type = 'Parameter'

                if($object.new) {
                    $newObject = [Parameter]::new($object.new.Name, $object.new.Kind, $object.new.QueryGroup, $object.new.Expression)
                    $object.new = $newObject
                }

                if($object.old){
                    $oldObject = [Parameter]::new($object.old.Name, $object.new.Kind, $object.old.QueryGroup, $object.old.Expression)
                    $object.old = $oldObject
                }
            }
        }
    }



    #=============================


    #We get all the properties we checked in the compare script 
    #En fait non : on risque d'avoir un soucis au niveau des types qui ne sont pas de base
    #Au pire est ce que ça pose problème ?
    if (-not $listOfProperties) {
        $listOfProperties = @()
        GetElementToCheckList | Where-Object hierarchyLevel -eq $objectsType | Select-Object element | ForEach-Object { $listOfProperties += $_.element }
    }

    #Table initialization
    $reportTable = "<div id=`"$($tableName)`" class=`"table-container hidden`">`n
    <table class=`"responsive-table two-block-table`"><tbody>`n"

    #Initialization of additional tables in case of textual check
    $additionalTable = ""

    #Initialization of the table header
    $reportHeader = "<tr class=`"table-header-row table-header-blocks`">"

    #Create the header of the table
    $reportHeader += "<th colspan=`"$($listOfProperties.Count)`" data-i18n-key=`"semantic.tables.headers.old_version`">semantic.tables.headers.old_version</th>"
    $reportHeader += "<th data-i18n-key=`"semantic.tables.headers.status`">semantic.tables.headers.status</th>"
    $reportHeader += "<th colspan=`"$($listOfProperties.Count)`" data-i18n-key=`"semantic.tables.headers.new_version`">semantic.tables.headers.old_version</th>"
    $reportHeader += "</tr>`n"
    # Column headers row using Orange style
    $reportHeader += "<tr class=`"table-header-row table-subheader-row`">"

    for ($j=0;$j -lt 2;$j++){

        #We do that two time, one for the old version, one for the new
        foreach($propertyList in $listOfProperties){

            if ($propertyList -is [array]) { #If multiple values, only keep one header with a specific localization key
                $propertyJoined = $propertyList -join "/"
            }
            elseif($propertyList -is [hashtable]){
                if($propertyList.Properties -and $propertyList.Sep){
                    $propertyJoined = $propertyList.Properties -join "."
                }
            }
            else{ #If only one string, don't do the join
                $propertyJoined = $propertyList
            }

            #$propertyLocalized = GetLocalization -input "$($objectsType).$($propertyJoined)"
            #write-host "The localization for $($objectsType).$($propertyJoined) is $propertyLocalized"

            $reportHeader += "<th data-i18n-key=`"$($objectsType).$($propertyJoined)`">$($objectsType).$($propertyJoined)</th>"
        }

        if($j -eq 0) { $reportHeader += "<th data-i18n-key=`"semantic.tables.headers.status`">semantic.tables.headers.status</th>" }
    }

    $reportHeader += "</tr>`n"

    #Adding the table header
    $reportTable += $reportHeader

    #Get the number of elements in the list of object
    #I don't remember why I did this, so perhaps removing it can cause some bugs
    #if(-not $listOfObjects -or $listOfObjects.GetType().Name -eq 'PSCustomObject'){
    #    $numObjects = 1
    #}
    #else{
    $numObjects = $listOfObjects.Length
    #}

    write-host "Go for $($numObjects) object(s)"

    #Then, for each line of the $listOfObjects we create a line of table in html (text)
    if ($numObjects -gt 0) {
        for ($i = 0; $i -lt $numObjects; $i++) {
        
            $currentNewObject = $listOfObjects[$i].new
            $currentOldObject = $listOfObjects[$i].old
            $currentObjectType = $listOfObjects[$i].type

            #Check if we have elements to textually check
            write-host "Getting elements to check if textual for $($currentObjectType)"
            $elementTexutalCheck = GetElementToTextualyCheck -objectName $currentObjectType
            write-host "We have the following list $($elementTexutalCheck)"

            #Check wich type of change we got here
            if($currentNewObject -and $currentOldObject){
                $diffType = "Different"
            }
            elseif($null -eq $currentNewObject){
                $diffType = "Deleted"
            }
            elseif($null -eq $currentOldObject){
                $diffType = "Added"
            }
            else{
                $diffType = 'Identical'
            }


            #Initialisation of the new report line (NO row-level class - colors applied at cell level per Figma)
            $reportLine = "<tr class=`"two-block-row`">"
            $reportLineNewPart = ""
            $reportLineOldPart = ""


            foreach($propertyList in $listOfProperties){

                write-host "$($propertyList)" -ForegroundColor blue

                $newVerisonPropertyList = @()
                $oldVerisonPropertyList = @()

                if($propertyList -is [hashtable]){
                    write-host $propertyList -BackgroundColor DarkYellow

                    foreach ($property in $propertyList.Properties) {
                        write-host $property -BackgroundColor DarkYellow
                        write-host $property.GetType() -BackgroundColor DarkYellow
                    }

                    if($propertyList.ContainsKey('Properties') -and $propertyList.ContainsKey('Sep') -and $propertyList.Sep.Count -eq 1){
                        
                        $separator = $propertyList.Sep

                        foreach ($property in $propertyList.Properties) {
                            $newVerisonPropertyList += [string](GetPropertyValue -element $currentNewObject -propertyPath $property)
                            $oldVerisonPropertyList += [string](GetPropertyValue -element $currentOldObject -propertyPath $property)
                        }
                    }
                    else{
                        write-host "An hashtable was defined but not with Properties or Sep" -ForegroundColor Red
                    }
                    
                }
                else {
                    
                    $separator = " "

                    foreach ($property in $propertyList) {
                        #write-host "$($currentNewObject)" -ForegroundColor Green
                        #write-host "$($currentOldObject)" -ForegroundColor DarkYellow
                        #write-host "$($property)"  -ForegroundColor Yellow
                        #write-host $(GetPropertyValue -element $currentNewObject -propertyPath $property) -ForegroundColor DarkGray
                        #write-host $(GetPropertyValue -element $currentOldObject -propertyPath $property) -ForegroundColor Gray
                        $newVerisonPropertyList += [string](GetPropertyValue -element $currentNewObject -propertyPath $property)
                        $oldVerisonPropertyList += [string](GetPropertyValue -element $currentOldObject -propertyPath $property)

                        #write-host $newVerisonPropertyList -ForegroundColor DarkGray
                        #write-host $oldVerisonPropertyList -ForegroundColor Gray
                    }

                }

                #write-host $newVerisonPropertyList.Count -BackgroundColor DarkRed

                #From a list, we want a concatenated string
                $countNotEmpty = $newVerisonPropertyList | Where-Object { $_ -ne $null -and [string]$_ -ne ''} | Measure-Object | Select-Object -ExpandProperty Count
                #write-host $countNotEmpty -BackgroundColor Darkblue
                if($countNotEmpty -eq 0){
                    $newVerisonProperty = $null
                }
                elseif($countNotEmpty -eq 1){
                    $newVerisonProperty = $newVerisonPropertyList[0]
                }
                else{
                    $newVerisonProperty = $newVerisonPropertyList -join $separator
                }

                $countNotEmpty = $oldVerisonPropertyList | Where-Object { $_ -ne $null -and [string]$_ -ne ''} | Measure-Object | Select-Object -ExpandProperty Count
                #write-host $countNotEmpty -BackgroundColor Darkblue
                if($countNotEmpty -eq 0){
                    $oldVerisonProperty = $null
                }
                elseif($countNotEmpty -eq 1){
                    $oldVerisonProperty = $oldVerisonPropertyList[0]
                }
                else{
                    $oldVerisonProperty = $oldVerisonPropertyList -join $separator
                }
                
                $propertyJoined = $propertyList -join "/"
                write-host $propertyJoined -ForegroundColor DarkCyan

                #If the property need the creation of a string comparison table, we create it here (if we have a element to check textually and have texts to compare)
                if($elementTexutalCheck.element -and $elementTexutalCheck.element.Contains($propertyJoined) -and ($newVerisonProperty -and $oldVerisonProperty)){
                    $additionalTableName = "table_add_$($tableName+$i+$propertyJoined)"
                    $additionalTable += "<div id=`"$($additionalTableName)`" class=`"modal-table-container hidden`">"
                    $additionalTable += ReportCreateStringCompareTable -newString $newVerisonProperty -oldString $oldVerisonProperty -tableName $additionalTableName
                    $additionalTable += "</div>"
                    $additionalButton = " <button class=`"btn-text-details`" onclick=`"openModal('$($additionalTableName)')`">🔍</button> "
                }
                else{
                    $additionalButton = $null
                }


                #Coloring the different elements depending if their are different or not
                #if($comparisonResult.ObjectNewVersion.Contains($newVerisonProperty)) {
                #    $repportLineNewPart += "<td class=`"different`">$($additionalButton)$($newVerisonProperty)</td>"
                #}
                #else {
                #    $repportLineNewPart += "<td>$($additionalButton)$($newVerisonProperty)</td>"
                #}

                #Coloring the different elements depending if their are different or not
                #if($comparisonResult.ObjectOldVersion.Contains($oldVerisonProperty)) {
                #    $repportLineOldPart += "<td class=`"different`">$($additionalButton)$($oldVerisonProperty)</td>"
                #}
                #else {
                #    $repportLineOldPart += "<td>$($additionalButton)$($oldVerisonProperty)</td>"
                #}
                
                # Apply Figma cell-level colors
                $oldCellStyle = ""
                $newCellStyle = ""
                
                if($diffType -eq "Added"){
                    # Added: green background ONLY in Nouvelle version (new) cells
                    $newCellStyle = " style=`"background-color:var(--tbl-row-added);`""
                }
                elseif($diffType -eq "Deleted"){
                    # Deleted: red background ONLY in Ancienne version (old) cells
                    $oldCellStyle = " style=`"background-color:var(--tbl-row-removed);`""
                }
                elseif($diffType -eq "Different"){
                    # Modified: orange background ONLY on cells that actually changed (old != new)
                    if($oldVerisonProperty -ne $newVerisonProperty){
                        $oldCellStyle = " style=`"background-color:var(--tbl-cell-changed);`""
                        $newCellStyle = " style=`"background-color:var(--tbl-cell-changed);`""
                    }
                }
                
                $reportLineNewPart += "<td$newCellStyle>$($additionalButton)$($newVerisonProperty)</td>"
                $reportLineOldPart += "<td$oldCellStyle>$($additionalButton)$($oldVerisonProperty)</td>"
            
            }

            #Old Version part
            $reportLine += $reportLineOldPart

            #Here the difference indicator
            if($diffType -eq "Different"){
                $diffIndice = "<td  class=`"diff-type-cell two-block-status`"><span class=`"diff-badge diff-badge--modified`" data-i18n-key=`"diffTypes.Modifie`">diffTypes.Modifie</span></td>"
            }
            elseif($diffType -eq "Deleted"){
                $diffIndice = "<td class=`"diff-type-cell two-block-status`"><span class=`"diff-badge diff-badge--removed`" data-i18n-key=`"diffTypes.Supprime`">diffTypes.Supprime</span></td>"
            }
            elseif($diffType -eq "Added"){
                $diffIndice = "<td class=`"diff-type-cell two-block-status`"><span class=`"diff-badge diff-badge--added`" data-i18n-key=`"diffTypes.Ajoute`">diffTypes.Ajoute</span></td>"
            }
            elseif ($diffType -eq "Identical"){
                $diffIndice = "<td class=`"diff-type-cell two-block-status`"><span class=`"diff-badge diff-badge--neutral`" data-i18n-key=`"diffTypes.Identique`">diffTypes.Identique</span></td>"
            }
            $reportLine += "$($diffIndice)"

            #New version part
            $reportLine += $reportLineNewPart

            #End of the line
            $reportLine += "</tr>`n"

            #Adding the line into the tabl
            $reportTable += $reportLine

        }
    }
    else{
        #Else, nothing to show, so creation of a line to inform that nothing appears there
        $reportTable += "<tr class='ok'><td colspan=`"$(($listOfProperties.Count * 2) + 1)`" data-i18n-key='semantic.tables.empty.$tableName'>semantic.tables.empty.$tableName</td></tr>"
    }

    #End of the table definition
    $reportTable += "</tbody></table>`n</div>`n"

    #if($additionalTable){
    #    $reportTable += $additionalTable
    #}

    return $reportTable, $additionalTable
}




#------------------------------


#Version of the function in order to be able to generate three tables that are put together to get the report
#OLDER from the current version of the single table, need update to be used
function ReportCreateCompareTables {

    param (
        $listOfObjects, #This list should have pointers to the object we want in the new and old
        [string] $objectsType,
        [string] $tableName,
        [string[]]$listOfProperties
    )

    #We get all the properties we checked in the compare script 
    #En fait non : on risque d'avoir un soucis au niveau des types qui ne sont pas de base
    #Au pire est ce que ça pose problème ?
    if (-not $listOfProperties) {
        $listOfProperties = @()
        GetElementToCheckList | Where-Object hierarchyLevel -eq $objectsType | Select-Object element | ForEach-Object { $listOfProperties += $_.element }
    }

    #Initialization of the table
    $repportTableLeft = "<table class=`"$($tableName) hidden`" style=`"min-width: 600px; border-collapse: collapse;`" border=`"1`">`n"
    $repportTableCenter = "<table class=`"$($tableName) hidden`" style=`"min-width: 600px; border-collapse: collapse;`" border=`"1`">`n"
    $repportTableRight = "<table class=`"$($tableName) hidden`" style=`"min-width: 600px; border-collapse: collapse;`" border=`"1`">`n"

    #Initialization of the table header
    $repportHeaderLeft = "<thead><tr>"
    $repportHeaderCenter = "<thead><tr>"
    $repportHeaderRight = "<thead><tr>"

    #Create the header of the table
    $repportHeaderLeft += "<th colspan=`"$($listOfProperties.Count)`" data-i18n-key=`"semantic.tables.headers.old_version`">semantic.tables.headers.old_version</th>"
    $repportHeaderCenter += "<th></th>"
    $repportHeaderRight += "<th colspan=`"$($listOfProperties.Count)`" data-i18n-key=`"semantic.tables.headers.new_version`">semantic.tables.headers.new_version</th>"

    #End the header initialisation
    $repportHeaderLeft += "</tr>`n"
    $repportHeaderCenter += "</tr>`n"
    $repportHeaderRight += "</tr>`n"

    #Creation of the column headers
    foreach($property in $listOfProperties){
        #$propertyLocalized = GetLocalization -input (+$property)

        $repportHeaderLeft += "<th data-i18n-key=`"$($property)`">$($property)</th>"
        $repportHeaderRight += "<th data-i18n-key=`"$($property)`">$($property)</th>"
    }
    $repportHeaderCenter += "<th data-i18n-key=`"semantic.tables.headers.status`">semantic.tables.headers.status</th>"

    $repportHeaderLeft += "</tr>`n"
    $repportHeaderCenter += "</tr>`n"
    $repportHeaderRight += "</tr>`n"

    #Adding the table header
    $repportTableLeft += $repportHeaderLeft
    $repportTableCenter += $repportHeaderCenter
    $repportTableRight += $repportHeaderRight

    #Get the number of elements in the list of object
    if(-not $listOfObjects -or $listOfObjects.GetType().Name -eq 'PSCustomObject'){
        $numObjects = 1
    }
    else{
        $numObjects = $listOfObjects.Length
    }

    write-host "Go for $($numObjects) object(s)"

    #Then, for each line of the $listOfObjects we create a line of table in html (text)
    if ($listOfObjects) {
        for ($i = 0; $i -lt $numObjects; $i++) {
        
            $currentNewObject = $listOfObjects[$i].new
            $currentOldObject = $listOfObjects[$i].old

            #Initialisation of the new report line
            $repportLineLeft = "<tr>"
            $repportLineCenter = "<tr>"
            $repportLineRight = "<tr>"

            $repportLineNewPart = ""
            $repportLineOldPart = ""


            foreach($property in $listOfProperties){

                write-host "$($currentNewObject)"
                write-host "$($property)"
                $newVerisonProperty = GetPropertyValue -element $currentNewObject -propertyPath $property
                $oldVerisonProperty = GetPropertyValue -element $currentOldObject -propertyPath $property


                #Coloring the different elements depending if their are different or not
                $repportLineNewPart += "<td>$($newVerisonProperty)</td>"
                $repportLineOldPart += "<td>$($oldVerisonProperty)</td>"
            
            

                #TODO We check if the checked property exist in the difference result (to know what element we have to highligth in the table)

            }

            #Old Version part
            $repportLineLeft += $repportLineOldPart

            #Here the difference indicator
            if($currentNewObject -and $currentOldObject){
                $diffIndice = "<td style=`"color: yellow;`" data-i18n-key=`"diffTypes.Modifie`">Different</td>"
            }
            elseif($null -eq $currentNewObject){
                $diffIndice = "<td style=`"color: red;`" data-i18n-key=`"diffTypes.Supprime`">Deleted</td>"
            }
            elseif($null -eq $currentOldObject){
                $diffIndice = "<td style=`"color: green;`" data-i18n-key=`"diffTypes.Ajoute`">Added</td>"
            }
            $repportLineCenter += "$($diffIndice)"

            #New version part
            $repportLineRight += $repportLineNewPart

            #End of the line
            $repportLineLeft += "</tr>`n"
            $repportLineCenter += "</tr>`n"
            $repportLineRight += "</tr>`n"

            #Adding the lines into the tables
            $repportTableLeft += $repportLineLeft
            $repportTableCenter += $repportLineCenter
            $repportTableRight += $repportLineRight

        }
    }

    #End of the table definition
    $repportTableLeft += "</table>`n"
    $repportTableCenter += "</table>`n"
    $repportTableRight += "</table>`n"

    return $repportTableLeft, $repportTableCenter, $repportTableRight
}


#==========================================================================================================================================


#Creation of the HTML code for the report
Function BuildHTMLReport {
    
    param (
        $comparisonResult,

        [string] $outputFolder
    )


    #Todo Check the output folder if it contain an html file name


    #Creation of the report

    $reportFinal = ""

    #Style definition
    $reportFinal += "<!DOCTYPE html PUBLIC `"-//W3C//DTD XHTML 1.0 Strict//EN`"  `"http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd`"> `
    <html xmlns=`"http://www.w3.org/1999/xhtml`"> `
    <head>
    <style>
    
* {
    margin: 0;
    padding: 0;
    box-sizing: border-box;
}

body {
    font-family: Helvetica, Arial, sans-serif;
    min-height: 100vh;
    display: flex;
    flex-direction: column;
}

/* Header */
.page-header {
    background-color: #000000;
    color: white;
    width: 1440;
    height: 60;
    padding-top: 20px;
    padding-right: 30px;
    padding-bottom: 15px;
    padding-left: 30px;
    gap: 10px;
    display: flex;
    justify-content: space-between;
    /*align-items: center;*/
    
}

.header-left {
    display: flex;
    align-items: center;
    gap: 15px;
    width: 515;
    height: 48;
    opacity: 1;
    gap: 30px;
}

.header-right {
    width: 855;
    height: 24;
    gap: 30px;
}

.logo {
    width: 30px;
    height: 30px;
    background-color: #FF7900;
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
}

/* Contenu principal */
.main-content {
    flex: 1;
    padding: 20px 30px;
}

.all-buttons {
    width: 1380;
    min-height: 150px;
    top: 100px;
    left: 30px;
    gap: 20px;
}

/* Boutons principaux */
.main-buttons {
    display: flex;
    gap: 10px;
    margin-bottom: 20px;
    border-bottom: #000;
    
}

.main-btn {
    width: 70;
    height: 64;
    justify-content: space-between;
    padding-top: 15px;
    border: 1px solid #ddd;
    position: relative;
    cursor: pointer;

    font-weight: 700;
    font-style: Bold;
    font-size: 18px;
    line-height: 24px;
    letter-spacing: 0px;
    text-align: center;
    background-color: transparent;
    border-color: transparent;
    color: #000000;
}

.main-btn:hover {
    width: 70;
    height: 64;
    justify-content: space-between;
    padding-top: 15px;

    font-weight: 700;
    font-style: Bold;
    font-size: 18px;
    line-height: 24px;
    letter-spacing: 0px;
    text-align: center;
    color: #555555;
}

.main-btn:active {
    
    font-weight: 700;
    font-style: Bold;
    font-size: 18px;
    line-height: 24px;
    letter-spacing: 0px;
    text-align: center;
    color: #F15E00;
}

/* État actif avec soulignement qui touche la séparation */
.main-btn.active {
    color: #000000;
}

.main-btn.active::after {
    content: '';
    position: absolute;
    bottom: -10px;
    left: 0;
    right: 0;
    height: 5px;
    background-color: #F15E00;
    z-index: 2;
}


/* Ligne de séparation */
.separator {
    height: 2px;
    background-color: #ddd;
    margin: -11px 0 20px 0;
    position: relative;
    z-index: 1;
}

/* Sous-boutons - style sans fond */
.sub-buttons-group {
    display: none;
    gap: 8px;
    margin-bottom: 20px;
}

.sub-buttons-group.active {
    display: flex;
}

.sub-btn {
    background-color: transparent;
    color: #666;
    border: 1px solid #CCCCCC;
    cursor: pointer;

    width: 184;
    height: 36;
    min-width: 54px;
    min-height: 32px;
    gap: 4px;
    padding-top: 5px;
    padding-right: 18px;
    padding-bottom: 7px;
    padding-left: 18px;
    border-radius: 20px;
    border-width: 2px;

    font-weight: 400;
    font-style: Bold;
    font-size: 16px;
    line-height: 24px;
}

.sub-btn:hover {
    background-color: #f8f9fa;
    border-color: #555555;
    color: #555555;
}

.sub-btn:active {
    background-color: #f8f9fa;
    border-color: #444444;
    color: #444444;
    outline: 2px solid #444444;
}

.sub-btn.active {
    border-color: #F15E00;
    color: #000000;
}

.image.check-mark{
    background-image: url('data:image/svg+xml;utf8,<svg width=`"18`" height=`"18`" viewBox=`"0 0 18 18`" fill=`"none`" xmlns=`"http://www.w3.org/2000/svg`"><path fill-rule=`"evenodd`" clip-rule=`"evenodd`" d=`"M14.9532 5.67744L8.12565 13.5525L8.12547 13.5523C7.92998 13.7775 7.62962 13.9219 7.29262 13.9219C6.95558 13.9219 6.65526 13.7775 6.45974 13.5523L6.45958 13.5524L3.04579 9.61494L3.04595 9.61482C2.89969 9.44632 2.81201 9.23268 2.81201 9C2.81201 8.72817 2.93141 8.4821 3.12446 8.30395L3.5512 7.9102C3.74425 7.73205 4.01096 7.62188 4.30555 7.62188C4.57067 7.62188 4.81305 7.71133 4.99964 7.85906L4.9998 7.85886L7.07924 9.9844L13.3906 4.34437L13.3908 4.34455C13.5815 4.17942 13.838 4.07812 14.1202 4.07812C14.7094 4.07812 15.187 4.51885 15.187 5.06252C15.187 5.29516 15.0993 5.50882 14.953 5.67732L14.9532 5.67744Z`" fill=`"#FF7900`"/></svg>');
}

.sub-btn.active::before {
    content: `"✓`";
    color: #F15E00;
    font-weight: bold;
    margin-right: 5px;
}

/* Troisième ligne de boutons (conditionnelle) */
.third-buttons {
    display: none;
    gap: 6px;
    margin-bottom: 20px;
}

.third-buttons.show {
    display: flex;
}

.third-btn {
    background-color: #F4F4F4;
    color: #000000;
    border: transparent;
    border-radius: 0;
    cursor: pointer;
    transition: all 0.3s ease;

    width: 153;
    height: 32;
    min-width: 84px;
    min-height: 32px;
    border-radius: 20px;
    gap: 4px;
    padding-top: 5px;
    padding-right: 10px;
    padding-bottom: 7px;
    padding-left: 18px;

    font-weight: 700;
    font-style: Bold;
    font-size: Font/Size/Body M (body 2);
    line-height: Font/Line-height/Body M (body 2);
    letter-spacing: 0px;

}

.third-btn:hover {
    background-color: #f0f0f0;
    color: #555555;
}

.third-btn.active {
    color: #F15E00;
}

.btn-text-details {
    width: 30px;
    height: 30px;
    border: none;
    border-radius: 50%;
    cursor: pointer;
    font-size: 16px;
    font-weight: bold;
    margin: 0 5px;
    transition: all 0.3s ease;
    align-items: center;
    justify-content: center;
    vertical-align: middle;
    background-color: #DDDDDD;
}

.btn-text-details:hover {
    background-color: #F15E00;
}

/* Tableau */
.table-container {
    overflow-x: scroll;
    margin-bottom: 30px;
}

table {
    width: 100%;
    border-collapse: collapse;
    background-color: white;
    box-shadow: 0 2px 4px rgba(0,0,0,0.1);
}

th, td {
    padding: 8px 10px ;
    text-align: left;
    border-bottom: 1px solid #ddd;
}

.table-main-header {
    background-color: #F4F4F4;
    font-weight: bold;
    color: #333;
}

/*
tr:hover {
    background-color: #f5f5f5;
}
*/

/* Styles pour le modal */
.modal {
    position: fixed;
    z-index: 1000;
    width: 100%;
    height: 100%;
    background-color: rgba(0,0,0,0.5);
    animation: fadeIn 0.3s ease;
}

.modal.show {
    display: flex;
    align-items: center;
    justify-content: center;
}

/* Contenu du modal */
.modal-content {
    background-color: white;
    margin: 4.5% auto;
    padding: 30px;
    border-radius: 10px;
    max-height: 85%;
    max-width: 85%;
    position: relative;
    box-shadow: 0 4px 20px rgba(0, 0, 0, 0.3);
    animation: modalOpen 0.3s ease-out;
}

.modal-header {
    background: #000000;
    color: white;
    padding: 20px;
    display: flex;
    justify-content: space-between;
    align-items: center;

    margin: 0;
    font-size: 0.8em;
    font-weight: bold;
    transition: opacity 0.3s ease;
}

/* Animation d'ouverture */
@keyframes modalOpen {
    from {
        opacity: 0;
        transform: translateY(-50px);
    }
    to {
        opacity: 1;
        transform: translateY(0);
    }
}


.close {
    background: none;
    border: none;
    color: white;
    font-size: 24px;
    cursor: pointer;
    padding: 5px;
    border-radius: 50%;
    width: 35px;
    height: 35px;
    display: flex;
    align-items: center;
    justify-content: center;
    transition: background-color 0.2s ease;
}

.close:hover {
    background-color: rgba(255,255,255,0.2);
}

.modal-body {
    padding: 20px;
    max-height: 600px;
    overflow-y: auto;
}

.modal-table {
    width: 100%;
    border-collapse: collapse;
    margin-top: 10px;
}


/* Information for the data results*/
.ok {
    background-color: #ECFDEF !important;
}
.alerte {
    background-color: #FDE5E6 !important;
}
.different {
    background-color: #FFFAE7 !important;
}

.capsule {
    background-color: #ffffff;
    display: inline-block;
    text-align: center;
    padding: 8px 16px;
    border-radius: 20px;
    font-weight: bold;
    font-size: 12px;
    letter-spacing: 0.5px;
    white-space: nowrap;
}

.capsule.statut-ok {
    background-color: #3DE35A !important;
    color: #000000;
}
.statut-alerte {
    background-color: #E70002 !important;
    color: #000000;
}
.statut-different {
    background-color: #FFCD0B !important;
    color: #000000;
}
.statut-identical {
    background-color: #f0f0f0 !important;
    color: #000000;
}


.hidden {
    display: none;
}



/* Pied de page */
.footer {
    background-color: #000;
    color: white;
    padding: 15px 30px;
    display: flex;
    justify-content: flex-end;
    align-items: center;
}

.footer-message {
    font-size: 14px;
    color: #ccc;
}


    "

    $reportFinal += "</style></head>"


    #Ajout des boutons

    $reportFinal += @"
<body style="overflow: auto">
    <!-- Entête -->
    <header class="page-header">
        <div class="header-left">
            <div class="logo">
                <svg width="30" height="30" viewBox="0 0 30 30" fill="none" xmlns="http://www.w3.org/2000/svg">
                    <path fill-rule="evenodd" clip-rule="evenodd" d="M0 30H30V0H0V30Z" fill="#FF7900"/>
                    <path fill-rule="evenodd" clip-rule="evenodd" d="M4.28564 25.7144H25.7143V21.4287H4.28564V25.7144Z" fill="white"/>
                </svg>
            </div>
            <h1 class="title">Rapports Power BI (PBIR)</h1>
        </div>
        <div class="header-right">
            <button class="language-switch" onclick="switchLanguage()">
                <svg width="40" height="30" viewBox="0 0 40 30" fill="none" xmlns="http://www.w3.org/2000/svg">
                    <path d="M0 0H40V30H0V0Z" fill="white"/>
                    <path d="M0 0H13.1343V30H0V0Z" fill="#00267F"/>
                    <path d="M26.8657 0H40.0001V30H26.8657V0Z" fill="#F31830"/>
                </svg>
            </button>
            <button class="language-switch" onclick="switchLanguage()">
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
            <!-- Boutons principaux -->
            <div class="main-buttons">
                <button class="main-btn" data-target="category_quality" onclick="showTable('table_quality_checks')">Check Qualité</button>
                <button class="main-btn" data-target="category_report">Rapport</button>
                <button class="main-btn" data-target="category_power_query">MDD-Power Query</button>
                <button class="main-btn" data-target="category_desktop">MDD-OBI Desktop</button>
            </div>

            <!-- Ligne de séparation -->
            <div class="separator"></div>

            <!-- Sous-boutons pour Report -->
            <div class="sub-buttons-group" id="category_report-subs">
                <button class="sub-btn" data-target="nouveau">Nouveau</button>
                <button class="sub-btn" data-target="modifier">Modifier</button>
                <button class="sub-btn" data-target="supprimer">Supprimer</button>
                <button class="sub-btn" data-target="categories">Catégories</button>
                <button class="sub-btn" data-target="stock">Stock</button>
            </div>

            <!-- Sous-boutons pour MDD-OBI Table query -->
            <div class="sub-buttons-group" id="category_power_query-subs">
                <button class="sub-btn" data-target="sub_category_table_query">Table query</button>
                <button class="sub-btn" data-target="sub_category_parameters_and_other">Parameters and other</button>
            </div>

            <!-- Sous-boutons pour MDD-OBI Desktop -->
            <div class="sub-buttons-group third-buttons" id="category_desktop-subs">
                <button class="third-btn" onclick="showTable('table_relationships')">Relations tables</button>
                <button class="third-btn" onclick="showTable('table_measures')">DAX Measures</button>
                <button class="third-btn" onclick="showTable('table_roles')">RLS</button>
                <button class="third-btn" onclick="showTable('table_perspectives')">Perspectives</button>
                <button class="third-btn" onclick="showTable('table_calculGroups')">Calculation Groups</button>
            </div>

            <!-- Troisième ligne de boutons pour Power Query -->
            <div class="third-buttons" id="sub_category_table_query-buttons">
                <button class="third-btn"  onclick="showTable('table_tables')">Names and Properties</button>
                <button class="third-btn" onclick="showTable('table_columns')">Columns</button>
                <button class="third-btn" onclick="showTable('table_steps')">Steps</button>
            </div>

            <div class="third-buttons" id="sub_category_parameters_and_other-buttons">
                <button class="third-btn"  onclick="showTable('table_paramValue')">Parameters</button>
            </div>

        </div>


"@

    #Ajout des tables

    #----Tables----
    $objectsToShow = GetObjectToShow -comparisonResult $comparisonResult -objectToGet "Table"
    $tables_table_1 = ReportCreateCompareTable -listOfObjects $objectsToShow -objectsType 'Model.Tables' -tableName 'table_tables' -listOfProperties @("name", "isHidden", "IsPrivate", "ExcludeFromModelRefresh", "IsRemoved", "columns.count", "partitions.name", "changedProperties.property", "measures.count")

    if($tables_table_1[0]) { $reportFinal += $tables_table_1[0] }


    #----Columns----
    $objectsToShow = GetObjectToShow -comparisonResult $comparisonResult -objectToGet @("DataColumn", "CalculatedColumn")
    $columns_table_2 = ReportCreateCompareTable -listOfObjects $objectsToShow -objectsType 'Model.Tables.Columns' -tableName "table_columns" -listOfProperties @("Table.name", "name", "type", "expression", "dataType", "sourceColumn", "formatString", "isAvailableInMdx", "summarizeBy.ToString", "sortByColumn", "changedProperties.property")

    if($columns_table_2[0]) { $reportFinal += $columns_table_2[0] }

    #----Steps----
    $objectsToShow = GetObjectToShow -comparisonResult $comparisonResult -objectToGet "Partition"
    $columns_table_3 = ReportCreateCompareTable -listOfObjects $objectsToShow -objectsType 'Model.Tables.partitions' -tableName "table_steps" -listOfProperties @("table.name", "queryGroup.folder", "source.expression")

    if($columns_table_3[0]) { $reportFinal += $columns_table_3[0] }

    #----Parameter values----
    $objectsToShow = GetObjectToShow -comparisonResult $comparisonResult -objectToGet "NamedExpression" -filter "Parameters"
    $columns_table_4 = ReportCreateCompareTable -listOfObjects $objectsToShow -objectsType 'Specific.parameter' -tableName "table_paramValue" -listOfProperties @("name", "ExpressionValue", "ExpressionMeta")

    if($columns_table_4[0]) { $reportFinal += $columns_table_4[0] }

    #----Other partitions----
    #$objectsToShow = GetObjectToShow -comparisonResult $comparisonResult -objectToGet "NamedExpression" -filter "Not Parameters"
    #$columns_table = ReportCreateCompareTable -listOfObjects $objectsToShow -objectsType 'Model.Tables.partitions' -tableName "table_otherValue" -listOfProperties @("name", "kind", "queryGroup.folder", "expression")

    #$reportFinal += $columns_table

    #----Relationship----
    $objectsToShow = GetObjectToShow -comparisonResult $comparisonResult -objectToGet "SingleColumnRelationship"
    $columns_table_5 = ReportCreateCompareTable -listOfObjects $objectsToShow -objectsType 'Model.relationships' -tableName "table_relationships" -listOfProperties @(@{Properties = @("fromCardinality.ToString", "toCardinality.ToString"); Sep = " to "}, @{Properties = @("fromTable.name", "fromColumn.name"); Sep = "."}, @{Properties = @("toTable.name", "toColumn.name"); Sep = "."}, "crossFilteringBehavior")

    if($columns_table_5[0]) { $reportFinal += $columns_table_5[0] }

    #----Measures DAX----
    $objectsToShow = GetObjectToShow -comparisonResult $comparisonResult -objectToGet "Measure"
    $columns_table_6 = ReportCreateCompareTable -listOfObjects $objectsToShow -objectsType 'Model.Tables.Measures' -tableName "table_measures" -listOfProperties @("name", "expression", "formatString", "isHidden", "displayFolder", "changedProperties.property")

    if($columns_table_6[0]) { $reportFinal += $columns_table_6[0] }

     #----Roles----
    $objectsToShow = GetObjectToShow -comparisonResult $comparisonResult -objectToGet "TablePermission"
    $columns_table_7 = ReportCreateCompareTable -listOfObjects $objectsToShow -objectsType 'Model.roles.TablePermissions' -tableName "table_roles" -listOfProperties @( "role.name", "table.name", "filterExpression")

    if($columns_table_7[0]) { $reportFinal += $columns_table_7[0] }

    #----Perspectives (using columns to get one line per column)----
    $objectsToShow = GetObjectToShow -comparisonResult $comparisonResult -objectToGet "PerspectiveColumn"
    $columns_table_8 = ReportCreateCompareTable -listOfObjects $objectsToShow -objectsType 'Model.Perspectives.PerspectiveTables.PerspectiveColumns' -tableName "table_perspectives" -listOfProperties @("PerspectiveTable.Perspective.name", "PerspectiveTable.name", "name")

    if($columns_table_8[0]) { $reportFinal += $columns_table_8[0] }

    #----Calculation Groups----
    $objectsToShow = GetObjectToShow -comparisonResult $comparisonResult -objectToGet "CalculationItem"
    $columns_table_9 = ReportCreateCompareTable -listOfObjects $objectsToShow -objectsType 'Model.Perspectives.PerspectiveTables.PerspectiveColumns' -tableName "table_calculGroups" -listOfProperties @("CalculationGroup.Table.name", "CalculationGroup.description", "name", "description", "state", "expression")

    if($columns_table_9[0]) { $reportFinal += $columns_table_9[0] }


    $reportFinal += "</main>"


    #Ajout de la partie modale

    $reportFinal += @"

    <div id="modal" class="modal hidden">
        <div class="modal-content">
            <div class="modal-header">
                <h2>Détails</h2>
                <button class="close" id="close-modal" onclick="closeModalFunction()">&times;</button>
            </div>
            <div class="modal-body">
"@

    #Adding the tables in the modal
    if ($tables_table_1[1]) { $reportFinal += $tables_table_1[1] }

    if ($columns_table_2[1]) { $reportFinal += $columns_table_2[1] }

    if ($columns_table_3[1]) { $reportFinal += $columns_table_3[1] }

    if ($columns_table_4[1]) { $reportFinal += $columns_table_4[1] }

    if ($columns_table_5[1]) { $reportFinal += $columns_table_5[1] }

    if ($columns_table_6[1]) { $reportFinal += $columns_table_6[1] }

    if ($columns_table_7[1]) { $reportFinal += $columns_table_7[1] }
    
    if ($columns_table_8[1]) { $reportFinal += $columns_table_8[1] }

    if ($columns_table_9[1]) { $reportFinal += $columns_table_9[1] }

    $reportFinal += @"
            </div>
        </div>
    </div>
"@





    #Ajout du pied de page

    $reportFinal += @"

    <footer class="footer">
        <div class="footer-message">
            © Orange Business 2025
        </div>
    </footer>
"@


    #Partie fonction (pour les boutons)
    $reportFinal += @"
    <script>

        // Variables globales
        let currentLang = 'fr';

        // Gestion des boutons principaux
        document.querySelectorAll('.main-btn').forEach(btn => {
            btn.addEventListener('click', function() {
                // Désactiver tous les boutons principaux
                document.querySelectorAll('.main-btn').forEach(b => b.classList.remove('active'));
                this.classList.add('active');
                
                // Masquer tous les groupes de sous-boutons
                document.querySelectorAll('.sub-buttons-group').forEach(group => {
                    group.classList.remove('active');
                });
                
                // Masquer toutes les troisièmes lignes
                document.querySelectorAll('.third-buttons').forEach(tb => tb.classList.remove('show'));

                // Afficher le groupe de sous-boutons correspondant
                const target = this.getAttribute('data-target');
                const targetGroup = document.getElementById(target + '-subs');
                if (targetGroup) {
                    targetGroup.classList.add('active');
                }
                
                // Désactiver tous les sous-boutons
                document.querySelectorAll('.sub-btn').forEach(sb => sb.classList.remove('active'));
                document.querySelectorAll('.third-btn').forEach(sb => sb.classList.remove('active'));
                
                //console.log('Bouton principal cliqué:', this.textContent);
            });
        });

        // Gestion des sous-boutons avec affichage de la troisième ligne
        document.querySelectorAll('.sub-btn').forEach(btn => {
            btn.addEventListener('click', function() {
                // Désactiver tous les sous-boutons du groupe actuel
                const activeGroup = document.querySelector('.sub-buttons-group.active');
                if (activeGroup) {
                    activeGroup.querySelectorAll('.sub-btn').forEach(b => b.classList.remove('active'));
                }
                
                // Activer le bouton cliqué
                this.classList.add('active');
                
                // Masquer toutes les troisièmes lignes
                document.querySelectorAll('.third-buttons').forEach(tb => tb.classList.remove('show'));
                
                // Afficher la troisième ligne correspondante
                const target = this.getAttribute('data-target');
                const targetButtons = document.getElementById(target + '-buttons');
                if (targetButtons) {
                    targetButtons.classList.add('show');
                }
                
                //console.log('Sous-bouton cliqué:', this.textContent);
            });
        });

        // Gestion des boutons de troisième niveau
        document.querySelectorAll('.third-btn').forEach(btn => {
            btn.addEventListener('click', function() {
                // Désactiver tous les boutons de troisième niveau dans le même groupe
                document.querySelectorAll('.third-btn').forEach(b => b.classList.remove('active'));
                
                // Activer le bouton cliqué
                this.classList.add('active');
                
                //console.log('Bouton de troisième niveau cliqué:', this.textContent);
            });
        });

        // Définir le premier bouton principal comme actif par défaut au chargement
        document.addEventListener('DOMContentLoaded', function() {
            const firstMainBtn = document.querySelector('.main-btn');
            if (firstMainBtn) {
                firstMainBtn.click(); // Simule un clic pour activer le premier bouton
            }
        });


        /* function for showing table or hide them */
	    function resetTables() {
          document.querySelectorAll('div.table-container').forEach(set => set.classList.add('hidden'));
        }

	    function showTable(id) {
          resetTables();
          document.getElementById(id).classList.remove('hidden');
        }


        function simpleOpenModalFunction() {
            modal.classList.remove('hidden');
            document.body.style.overflow = 'hidden';
        }

        /* function for showing table or hide them */
	    function resetModalTables() {
          document.querySelectorAll('div.modal-table-container').forEach(set => set.classList.add('hidden'));
        }

        // Fonction pour ouvrir le modal
        function openModal(id) {
            //Faire en sorte que le tableau soit visible et que les autres soient désactivés
            resetModalTables();
            
            // Afficher le modal
            simpleOpenModalFunction();

            document.getElementById(id).classList.remove('hidden');

            //On désactive la possibilité de défilement sur la page proincipale
            document.body.style.overflow = 'none';
        }

        // Fonction pour fermer le modal
        function closeModalFunction() {
            resetModalTables();

            modal.classList.add('hidden');
            document.body.style.overflow = 'auto';
        }

        // Fermer le modal en cliquant sur l'arrière-plan
        modal.addEventListener('click', function(userEvent) {
            if (userEvent.target === modal) {
                closeModalFunction();
            }
        });

        // Fermer le modal avec la touche Échap
        document.addEventListener('keydown', function(userEvent) {
            if (userEvent.key === 'Escape' && modal.classList.contains('hidden')) {
                closeModalFunction();
            }
        });


    </script>
"@





    $reportFinal += "</body></html>"

    #Save the report

    #add a html file name
    $outputFolder += "Report.html"

    # Saving the HTML content
    $reportFinal | Out-File -FilePath $outputFolder -Encoding UTF8

    # Ouvre automatiquement le rapport
    Start-Process $outputFolder


    return $reportFinal
}


#==========================================================================================================================================



# Main (harness désactivé par défaut)
# Ce bloc n'est exécuté que si le fichier est lancé directement ET si la variable d'env PBI_MDD_RUN_STANDALONE=1.
# Aucun chemin en dur n'est conservé ici pour éviter les conflits avec l'orchestrateur (main.ps1).

$__thisScriptPath = $MyInvocation.MyCommand.Path
$__callerPath = $PSCommandPath
$__runStandalone = $false
if ($__thisScriptPath -and $__callerPath) {
    $__runStandalone = ($__thisScriptPath -eq $__callerPath)
} else {
    $__runStandalone = ($MyInvocation.InvocationName -ne '.')
}

if ($__runStandalone -and $env:PBI_MDD_RUN_STANDALONE -eq '1') {
    try {
        $newDir = Read-Host "Chemin du projet NOUVEAU (.pbip parent)"
        $oldDir = Read-Host "Chemin du projet ANCIEN (.pbip parent)"
        $projetArray = LoadProjectVersionsPath -newVersionProjectRepertory $newDir -oldVersionProjectRepertory $oldDir

        $parsedArray = [PSCustomObject]@{
            ElementNewVersion = $projetArray[0].dataBase.Model
            ElementOldVersion = $projetArray[1].dataBase.Model
            PathNewVersion = 'DataBase'
            PathOldVersion = 'DataBase'
        }

        $comparisonResult = CheckDifferenceInSubElement -element $parsedArray -elementChecked 'Model'

        $chemin = Read-Host "Dossier de sortie HTML (laisser vide pour ignorer la génération)"
        if ($chemin) {
            BuildHTMLReport -outputFolder $chemin -comparisonResult $comparisonResult
            Start-Process (Join-Path $chemin 'Report.html') -ErrorAction SilentlyContinue
        }
    } catch {
        Write-Host "ERREUR (mode autonome PBI_MDD_extract.ps1): $($_.Exception.Message)" -ForegroundColor Red
    }
}