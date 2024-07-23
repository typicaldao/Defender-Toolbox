# Global variables
$ModuleFolder = [System.Environment]::GetFolderPath('MyDocuments') + "\PowerShell\Modules"
$ModuleName = "Defender-Toolbox"
$ModuleFile = "$ModuleName.psm1"
$ModuleManifestFile = "$ModuleName.psd1"

function Get-LatestVersion{
    # $versionListFileUri = "http://127.0.0.1/version_list" # for local test (python3 -m http.server 8080)
    $versionListFileUri = "https://raw.githubusercontent.com/typicaldao/$ModuleName/main/version_list"
    $version = Invoke-RestMethod -Uri $versionListFileUri # should be a dymamic variable list in the future.
    return [string]$version
}

function Get-LocalVersion{
    if (Test-Path $ModuleFolder\$ModuleName){
        try {
            Import-Module $ModuleName -ErrorAction Stop # Try to import the module
            $local_version = (Get-Module -Name $ModuleName).Version
            Remove-Module -Name $ModuleName
            return $local_version # Returns System.Version. Use ToString() to convert type.
        }
        catch {
            Write-Host "Failed to import module $ModuleName."
        }
    }
    else {
        return "0.0" # Returns a 0.0 version if module has never been installed.
    }
}

function Download-DefenderToolbox([string]$version) {
    # Module links
    $ModuleFileUri = "https://raw.githubusercontent.com/typicaldao/$ModuleName/main/Modules/$version/$ModuleFile"
    $ModuleManifestUri = "https://raw.githubusercontent.com/typicaldao/$ModuleName/main/Modules/$version/$ModuleManifestFile"
    
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
        Write-Host "Trying to install $ModuleName version $version at path: $ModuleFolder."
        try {
            Copy-Item -Path $env:TEMP\$ModuleFile -Destination $ModuleFolder\$ModuleName\$version
            Copy-Item -Path $env:TEMP\$ModuleManifestFile -Destination $ModuleFolder\\$ModuleName\$version
        }
        catch {
            Write-Host -ForegroundColor Red "Failed to copy files."
            Remove-Item $ModuleFolder\$ModuleName\$version # clean up the folder.
            return
        }
    }
    else{
        Write-Host -ForegroundColor Red "Module folder $ModuleFolder is not inside PowerShell module path."
        Remove-Item $ModuleFolder\$ModuleName\$version # clean up the folder.
        Write-Host -ForegroundColor Yellow "Please manually copy the files $env:TEMP\$ModuleFile and $env:TEMP\$ModuleManifestFile to one of the folders below."
        Write-Host $env:PSModulePath.Split(";")
        return
    }  
    
    Write-Host -ForegroundColor Green "Module files of $ModuleName version $version are copied to the module folder."
}

function Update-PsUserProfile {
    $command = "Import-Module -Name $ModuleName"
    $imported = $false
    if (Test-Path $PROFILE){
        $ProfileContent = Get-Content $PROFILE
        $ProfileContent | ForEach-Object {
            if ($_.Trim() -eq $command) { 
                $imported = $true
                Write-Host "$ModuleName has been imported automatically."
                return 
            }
        }
        if (!$imported){
            Write-Host "Trying to add a new line to import module in your existing profile."
            $command | Out-File $PROFILE -Append -Confirm
        }
    }
    else{
        Write-Host "You do not have a PowerShell profile. Would you create one and import the module automatially?"
        New-Item -ItemType File -Value $command -Confirm $PROFILE 
    }
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

 Update-PsUserProfile  # Will add 'Import-Module -Name Defender-Toolbox' in Profile so the module is automatically loaded.
