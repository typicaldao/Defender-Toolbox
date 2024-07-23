# Global variables
$ModuleFolder = "$env:USERPROFILE\Documents\PowerShell\Modules\Defender-Toolbox"

function Get-LatestVersion{
    $versionListFileUri = "http://127.0.0.1/version_list" # for local test
    # $versionListFileUri = GitHut Remote Url
    $version = Invoke-RestMethod -Uri $versionListFileUri # should be a dymamic variable.
    return $version
}

function Get-LocalVersion{
    if (Test-Path $ModuleFolder){
        Import-Module Defender-Toolbox # Try to import the module
        $local_version = (Get-Module -Name Defender-Toolbox).Version
        Remove-Module -Name Defender-Toolbox
        return $local_version # Returns System.Version. Use ToString() to convert type.
    }
    else {
        return "0.0"
    }
}



function Download-DefenderToolbox([string]$version) {
    # Module links
    $ModuleFile = "Defender-Toolbox.psm1"
    $ModuleManifestFile = "Defender-Toolbox.psd1"
    $ModuleFileUri = "https://raw.githubusercontent.com/typicaldao/Defender-Toolbox/main/Modules/$version/$ModuleFile"
    $ModuleManifestUri = "https://raw.githubusercontent.com/typicaldao/Defender-Toolbox/main/Modules/$version/$ModuleManifestFile"
    
    Write-Host "Download the latest version from GitHub to your temp folder."

    try {
        Invoke-RestMethod -Uri $ModuleFileUri -OutFile $env:TEMP\$ModuleFile
        Invoke-RestMethod -Uri $ModuleManifestUri -OutFile $env:TEMP\$ModuleManifestFile
    }
    catch {
        Write-Host -ForegroundColor Red "Downloading failed:"
        Write-Host $_
        return
    }

    # Confirm module installation path.
    # OneDrive folder needs to be confirmed.
    if ($ModuleFolder -in $env:PSModulePath.Split(";")){
        Write-Host "Install Defender-Toolbox version $version at path: $ModuleFolder."
    }  
    # To be continued.  
}

function Copy-ModuleFiles{
    # Do this when the download is successful.
    Write-Host "Update module files to the module folder."
}

 # Main
 $local_version = Get-LocalVersion
 $version = Get-LatestVersion

 if ($version -gt $local_version){
    Download-DefenderToolbox
    Copy-ModuleFiles
 }

 # To be continued
