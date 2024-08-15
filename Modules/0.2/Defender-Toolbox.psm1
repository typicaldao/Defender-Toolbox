function Convert-MpOperationalEventLogTxt {
    param(
        [string]$Path="$PWD\MpOperationalEvents.txt",
        [string]$OutFile="$PWD\MpOperationalEvents.csv"
    )
    
    if ((Test-Path $Path) -eq $false ) {
        Write-Host "MpOperationalEvents.txt is not found! Exiting."
        return
    }
    
    #Read Data
    $LogData = (Get-Content -Path $Path) + "*****"
    
    # Some variables
    $Counter = -1
    $Details = ""
    $Result = [array][PSCustomObject]@() 
    $Description = ""
    
    $i = 0 # Counter for Write-Progress
    $TotalLines = $Logdata.Count
    
    $LogData | ForEach-Object {
        if ($_ -match '^[*]+') {
            if ($Counter -ne -1) {
               $Result += [PSCustomObject]@{
                Time = $Timestamp
                ErrorLevel = $Level
                EventId = $Id
                EventDescription = $Description
                EventDetails = $Details
               }
            }
            # Initialization of event block
            $Details = ""
            $Counter = 0 
        } 
        # Parse the first line of event info
        if ($Counter -eq 1) {  
            $Timestamp = $_.Substring(0,22) 
            $Level,$Id = $_.Split()[5],$_.Split()[8] 
        } 
        # One-line description
        elseif ($Counter -eq 2) {
            $Description = $_
        }
        # Multi-lines details
        elseif ($Counter -gt 2) {
            $Details += $_ + "`n"
        }
        $Counter++ 
        $i++
        Write-Progress -Activity "Parsing" -Status "$i of $TotalLines lines parsed" -PercentComplete ($i/($TotalLines)*100)
    }
    
    $Result | Export-Csv -Encoding utf8BOM -Path $OutFile 
}
Export-ModuleMember -Function Convert-MpOperationalEventLogTxt

function Convert-MacDlpPolicyBin {

    param (
        [string]$Path="$PWD\dlp_policy.bin",
        [string]$OutFile="$PWD\dlp_policy.json"
    )

    if ((Test-Path $Path) -eq $false ) {
        Write-Host "Policy file not found. Exiting."
        exit
    }
    $byteArray = Get-Content -AsByteStream -Path $Path

    $memoryStream = New-Object System.IO.MemoryStream(,$byteArray)
    $deflateStream = New-Object System.IO.Compression.DeflateStream($memoryStream,  [System.IO.Compression.CompressionMode]::Decompress)
    $streamReader =  New-Object System.IO.StreamReader($deflateStream, [System.Text.Encoding]::utf8)
    $policyStr = $streamReader.ReadToEnd()
    $policy = $policyStr | ConvertFrom-Json


    $policyBodyCmd = ($policy.body | ConvertFrom-Json).cmd 
    if ($policyBodyCmd) {Set-Content -Path $OutFile $policyBodyCmd} 
}
Export-ModuleMember -Function Convert-MacDlpPolicyBin

function Convert-UnixTime([string]$unixTime){
    # Unix time converting
    $base = [datetime]'1970-01-01 00:00:00'
    $timezone = (Get-TimeZone).BaseUtcOffset
    $baseTime = $base + $timezone

    switch ($unixTime.length) {
        10 { $converter = $baseTime.Addseconds($unixTime)}
        13 { $converter = $baseTime.AddMilliseconds($unixTime)}
        16 { $converter = $baseTime.AddMicroseconds($unixTime)}
        Default { Write-Host "Not a valid Unix Time!";break}
    }
    return $converter
}
Export-ModuleMember -Function Convert-UnixTime

function Convert-Wdavhistory {

    param(
        $Path="$PWD\wdavhistory",
        $OutFile="$PWD\wdavhistory.log"
    )

    $Logdata = Get-Content -Raw $Path | ConvertFrom-Json

    $Logdata.scans | ForEach-Object {
        $_.endTime = Convert-UnixTime $_.endTime
        $_.startTime = Convert-UnixTime $_.startTime
        $duration = $_.endTime - $_.startTime
        $_ | Add-Member -NotePropertyName scanDuration -NotePropertyValue $duration.ToString()
        $_.threats = $_.threats | Out-String
    }

    $Logdata.scans | Format-Table -Property startTime, endTime, scanDuration, filesScanned, scheduled, state, type, threats| Out-File $OutFile
}
Export-ModuleMember -Function Convert-Wdavhistory

function Convert-MdavRtpDiag() {
    param(
        [Parameter(Mandatory=$true)]$rtplog,
        $OutFile = "$rtplog.txt"
    )
    if (Test-Path $rtpLog){
        $val = Get-Content $rtpLog | ConvertFrom-Json
        $val.counters | ForEach-Object {$_.totalFilesScanned = [int32]$_.totalFilesScanned} # Convert totalFileScanned type from string to int.
        $val.counters | Where-Object {$PSItem.totalFilesScanned -gt 0} | Select-Object id, name, totalFilesScanned, path | Sort-Object -Property totalFilesScanned -Descending | Format-Table | Out-File -Force $OutFile
    }
    else {
        Write-Host "Cannot find the file $rtplog. Exit."
        return
    }
    
    Write-Host "File saved as $OutFile successfully."
}
Export-ModuleMember -Function Convert-MdavRtpDiag

function Convert-MpRegistrytxtToJson {
    param(
        [string]$Path="$PWD\MpRegistry.txt",
        [string]$OutFile="$PWD\MpRegistry.json"
    )

    $lines = Get-Content $Path
    $json_result = [ordered]@{}

    $this_line = 0
    $depth = 1
    $regs = ""
    $key = ""
    $value = ""
    
    :DefenderAV foreach ($line in $lines) {
        # Match the root category
        if ($line.StartsWith("Current configuration options for location")){
            $policy = ($line -split '"')[1]
            $json_result.$policy = [ordered]@{}
        }

        elseif ($line.StartsWith("Windows Setup keys from")) {
            $this_line++
            break DefenderAV # exit when Defender AV regs are finished
        }
        
        # Match a category like '[*]'
        elseif ($line -match '^\[(.+)\]$') {
            $regs = $matches[1] -split "\\"
            $depth = $regs.Length
            
            # Length of 'keys' is the depth of json
            switch ($depth){
                1 { $json_result.$policy.($regs[0]) = [ordered]@{}; break }
                2 { $json_result.$policy.($regs[0]).($regs[1]) = [ordered]@{}; break }
                {$line -eq '[NIS\Consumers\IPS]'} { $json_result.$policy.($regs[0]).($regs[1]) = [ordered]@{} }
                3 { $json_result.$policy.($regs[0]).($regs[1]).($regs[2]) = [ordered]@{}; break }
                4 { $json_result.$policy.($regs[0]).($regs[1]).($regs[2]).($regs[3]) = [ordered]@{}; break }
            }
        } 

        # Match a line like *key* : *value*
        elseif ($line -match '^\s{4}(\S+)\s+\[.+\]\s+:\s(.+)$') {
            $key = $matches[1]
            $value = $matches[2]
            if ($value -in ("<NO VALUE>","(null)")) { $value = "" }
            switch ($depth){
                1 { $json_result.$policy.($regs[0]).$key = $value; break }
                2 { $json_result.$policy.($regs[0]).($regs[1]).$key = $value; break}
                3 { $json_result.$policy.($regs[0]).($regs[1]).($regs[2]).$key = $value; break }
                4 { $json_result.$policy.($regs[0]).($regs[1]).($regs[2]).($regs[3]).$key = $value; break }
            }
        }

        # Try to merge Device Control policy into single line. It will break when the DC policy format is not expected.
        elseif ($line -match '\s{2}<.+>$' ) {
            $json_result.$policy.($regs[0]).($regs[1]).$key += $line
        }

        $this_line++
        }

        # Deal with other registry keys
        # Initializing
        $json_result["Others"]=[ordered]@{}
        $json_result["WindowsUpdate"]=[ordered]@{}
        $policy = "Others"
        $reg_path = ""


        for ($i = $this_line; $i -lt $lines.Count; $i++) {

            switch -Wildcard ($lines[$i]){
                '`[HKEY_*WindowsUpdate*`]' { 
                    $policy = "WindowsUpdate"
                    $reg_path = $lines[$i].Substring(1,$lines[$i].Length-2)
                    $json_result.$policy.$reg_path = [ordered]@{}
                    break
                }
                '`[HKEY*`]' { 
                    $policy = "Others"
                    $reg_path = $lines[$i].Substring(1,$lines[$i].Length-2)
                    $json_result.$policy.$reg_path = [ordered]@{}
                    break
                }
                '*`[REG_*`]*' {
                    if ($lines[$i] -match '^\s{4}(\S+)\s+\[.+\]\s+:\s(.+)$'){
                        $key = $matches[1]
                        $value = $matches[2]
                        $json_result.$policy.$reg_path.$key = $value
                    }

                }
            }
        }

    $json = $json_result | ConvertTo-Json -Depth 4 -WarningAction Ignore | ForEach-Object {
        [Regex]::Replace($_, "\\u(?<Value>[a-fA-F0-9]{4})", { param($m) ([char]([int]::Parse($m.Groups['Value'].Value, [System.Globalization.NumberStyles]::HexNumber))).ToString() } )}       
        # Converto-Json in PowerShell 5.1 does not have escape handling options, so we do it manually.
        # Reference: https://stackoverflow.com/questions/47779157/convertto-json-and-convertfrom-json-with-special-characters/47779605#47779605
    $json | Out-File -FilePath $OutFile -Force
}
Export-ModuleMember -Function Convert-MpRegistrytxtToJson

function Update-DefenderToolbox {
    # Global variables
    $PSVersion = [string]$PSVersionTable.PSVersion
    $PSModuleUserProfile = "\WindowsPowerShell\Modules"
    if ($PSVersion -gt "7.1"){$PSModuleUserProfile = "\PowerShell\Modules"}
    $ModuleFolder = [System.Environment]::GetFolderPath('MyDocuments') + $PSModuleUserProfile
    $ModuleName = "Defender-Toolbox"
    $ModuleFile = "$ModuleName.psm1"
    $ModuleManifestFile = "$ModuleName.psd1"
    $result = $false

    function Get-LatestVersion{
        # $versionListFileUri = "http://127.0.0.1/version_list" # for local test (python3 -m http.server 8080)
        $versionListFileUri = "https://raw.githubusercontent.com/typicaldao/$ModuleName/main/version_list"
        try {
            $version = Invoke-RestMethod -Uri $versionListFileUri
            $latest_version = $version.Split()[-1]
            return [string]$latest_version
        }
        catch {
            Write-Host "Failed to get remote version."
            Write-Host -ForegroundColor Red "Error: $_"
            return "0"
        }
    }

    function Get-LocalVersion{
        if (Test-Path $ModuleFolder\$ModuleName){
            try {
                Import-Module $ModuleName -ErrorAction Stop # Try to import the module
                $local_version = (Get-Module -Name $ModuleName).Version
                # Remove-Module -Name $ModuleName # The module should not be removed when used as a local module.
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
            return $false
        }

        Write-Host -ForegroundColor Green "Download $ModuleFile and $ModuleManifestfile at $env:TEMP successfully."
        return $true
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
                return $false
            }
        }
        else{
            Write-Host -ForegroundColor Red "Module folder $ModuleFolder is not inside PowerShell module path."
            Remove-Item $ModuleFolder\$ModuleName\$version # clean up the folder.
            Write-Host -ForegroundColor Yellow "Please manually copy the files $env:TEMP\$ModuleFile and $env:TEMP\$ModuleManifestFile to one of the folders below."
            Write-Host $env:PSModulePath.Split(";")
            return $false
        }  
        
        Write-Host -ForegroundColor Green "Module files of $ModuleName version $version are copied to the module folder."
        return $true
    }

    function Update-PsUserProfile {
        $command = "`nImport-Module -Name $ModuleName"
        $imported = $false
        if (Test-Path $PROFILE){
            $ProfileContent = Get-Content $PROFILE
            $ProfileContent | ForEach-Object {
                if ($_.Trim() -eq $command.Trim()) { 
                    $imported = $true
                    Write-Host "$ModuleName has been imported automatically."
                    return 
                }
            }
            if (!$imported){
                Write-Host -ForegroundColor Cyan "Trying to add a new line to import module in your existing profile. Please confirm."
                $command | Out-File $PROFILE -Append -Confirm
            }
        }
        else{
            Write-Host -ForegroundColor Cyan "You do not have a PowerShell profile. Would you create one and import the module automatially?"
            New-Item -ItemType File -Value $command -Confirm $PROFILE 
        }
    }

    # Main
    $local_version = Get-LocalVersion
    $version = Get-LatestVersion

    if ($version -gt $local_version){
        $result = Download-DefenderToolbox($version)
        $result = Copy-ModuleFiles($version)
    }
    elseif ($version -eq $local_version) {
        Write-Host "You have already installed the latest version: $version"
        $result = $true
    }

    if ($result) { Update-PsUserProfile }   # Will add 'Import-Module -Name Defender-Toolbox' in Profile so the module is automatically loaded.
}
Export-ModuleMember -Function Update-DefenderToolbox