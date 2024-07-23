# Global variables
$ModuleFolder = [System.Environment]::GetFolderPath('MyDocuments') + "\PowerShell\Modules"
$ModuleName = "Defender-Toolbox"
$ModuleFile = "Defender-Toolbox.psm1"
$ModuleManifestFile = "Defender-Toolbox.psd1"

function Get-LatestVersion{
    # $versionListFileUri = "http://127.0.0.1/version_list" # for local test (python3 -m http.server 8080)
    $versionListFileUri = "https://raw.githubusercontent.com/typicaldao/Defender-Toolbox/main/version_list"
    $version = Invoke-RestMethod -Uri $versionListFileUri # should be a dymamic variable list in the future.
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
        return "0.0" # Returns a 0.0 version if module has never been installed.
    }
}

function Download-DefenderToolbox([string]$version) {
    # Module links
    $ModuleFileUri = "https://raw.githubusercontent.com/typicaldao/Defender-Toolbox/main/Modules/$version/$ModuleFile"
    $ModuleManifestUri = "https://raw.githubusercontent.com/typicaldao/Defender-Toolbox/main/Modules/$version/$ModuleManifestFile"
    
    Write-Host "Try to download the latest version from GitHub to your temp folder."

    try {
        Invoke-RestMethod -Uri $ModuleFileUri -OutFile $env:TEMP\$ModuleFile
        Invoke-RestMethod -Uri $ModuleManifestUri -OutFile $env:TEMP\$ModuleManifestFile
    }
    catch {
        Write-Host -ForegroundColor Red "Downloading failed:"
        Write-Host $_
        return
    }

    Write-Host -ForegroundColor Green "Download $ModuleFile and $ModuleManifestfile at $env:TEMP successfully."
}

function Copy-ModuleFiles([string]$version){
    # Do this when the download is successful.
    # OneDrive folder needs to be confirmed in the future. To be continued.

    # Create module folder with versions if it does not exists.
    if(-Not (Test-Path $ModuleFolder\$ModuleName\$version)){
        New-Item -ItemType Directory -Path $ModuleFolder\$ModuleName -Name $version
    }
    
    # Confirm module installation path.
    if ($ModuleFolder -in $env:PSModulePath.Split(";")){
        Write-Host "Trying to install Defender-Toolbox version $version at path: $ModuleFolder."
        try {
            Copy-Item -Path $env:TEMP\$ModuleFile -Destination $ModuleFolder\$ModuleName\$version
            Copy-Item -Path $env:TEMP\$ModuleManifestFile -Destination $ModuleFolder\\$ModuleName\$version
        }
        catch {
            Write-Host -ForegroundColor Red "Failed to copy files."
            return
        }
    }
    else{
        Write-Host -ForegroundColor Red "Module folder $ModuleFolder is not inside PowerShell module path."
        Write-Host -ForegroundColor Yellow "Please manually copy the files $env:TEMP\$ModuleFile and $env:TEMP\$ModuleManifestFile to one of the folders below."
        Write-Host $env:PSModulePath.Split(";")
        return
    }  
    
    Write-Host -ForegroundColor Green "Module files of Defender-Toolbox version $version are copied to the module folder."
}

 # Main
 $local_version = Get-LocalVersion
 $version = Get-LatestVersion

 if ($version -gt $local_version){
    Download-DefenderToolbox($version)
    Copy-ModuleFiles($version)
 }
 elseif ($version -eq $local_version) {
    Write-Host "You have already installed the latest version: $version"
    return
 }

 # function Update-PSUserProfile
 # Will add 'Import-Module Defender-Toolbox' in Profile so the module is automatically loaded.
