#==========================================================================================================================================

#Disable telemetry
$env:POWERSHELL_TELEMETRY_OPTOUT = "1"

#==========================================================================================================================================


Function LoadNeededDLL {

    param (
        [string] $path
    )

    Write-Information "The path is : $path"

    $dllFolderPath = $path + "\lib\"

    $dllPaths = @(
        ($dllFolderPath + "System.Buffers.dll"),
        ($dllFolderPath + "System.Diagnostics.DiagnosticSource.dll"),
        ($dllFolderPath + "System.Memory.dll"),
        ($dllFolderPath + "System.Numerics.Vectors.dll"),
        ($dllFolderPath + "System.Runtime.CompilerServices.Unsafe.dll"),
        #($dllFolderPath + "System.Security.Cryptography.Cng.dll"),
        #($dllFolderPath + "Newtonsoft.Json.dll"),
        ($dllFolderPath + "Microsoft.Identity.Client.dll"),
        ($dllFolderPath + "Microsoft.Identity.Client.NativeInterop.dll"),
        ($dllFolderPath + "Microsoft.IdentityModel.Abstractions.dll"),
        ($dllFolderPath + "Microsoft.Identity.Client.Broker.dll"),
        ($dllFolderPath + "Microsoft.AnalysisServices.Core.dll"),
        ($dllFolderPath + "Microsoft.AnalysisServices.Tabular.dll"),
        ($dllFolderPath + "Microsoft.AnalysisServices.Tabular.Json.dll")
    )

    foreach ($dllPath in $dllPaths){
    
        if (-not (Test-Path $dllPath)) {
            throw "in $path DLL not found for $dllPath"
        }

        # Check Authenticode signature
        $signature = Get-AuthenticodeSignature -FilePath $dllPath

        if ($signature.Status -ne 'Valid') {
            throw "La DLL n'a pas une signature valide (status: $($signature.Status))."
        }

        #TODO : Peut-être tester la signature en fonction de son contenu ?


        $zoneInfo = Get-Item -Path $dllPath -Stream Zone.Identifier -ErrorAction SilentlyContinue

        if ($null -ne $zoneInfo) {
            Write-Host "DLL blocked by Windows, unblocking"
            try {
                Unblock-File -Path $dllPath
                Write-Host "Unblocking done"
            }
            catch {
                throw "Impossible to unblock $dllPath"
            }
        }
        else {
            Write-Host "DLL already unblocked"
        }

        #Charge the Assembly
        try {
            Add-Type -Path $dllPath
            Write-Host "$dllPath loaded successfully"
        }
        catch {
            throw "Error when loading $dllPath"
        }
    }
}

#==========================================================================================================================================

#exit 0