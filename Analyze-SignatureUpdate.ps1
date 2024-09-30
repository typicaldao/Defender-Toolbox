# param(
#     [string]$Path="$PWD\MpSupportFiles.cab"
# )
# Analyze MpSupportFiles.cab to get a result of signature update issue

$Path="$PWD\MpSupportFiles.cab"

# const
$MPLOG_FOLDER = "$PWD\MpSupportFiles"


# Extract Defender logs
if (!(Test-Path $Path)){
    return
}
if (!(Test-Path $MPLOG_FOLDER)){ 
    New-Item -Name MpSupportFiles -ItemType Directory 
}
# Expand MpSupportFiles.cab
# Source: C:\Windows\system32\expand.exe
expand -i -f:* $Path Destination $MPLOG_FOLDER
Set-Location $MPLOG_FOLDER

# Organize the files
$dynamic_signatures = Get-ChildItem | Where-Object Name -Match "[a-f0-9]{40}"

# Convert MpRegistry.txt to JSON.
Convert-MpRegistrytxtToJson

# Merge and convert MpOperationalEvents.txt.bak
$events_log = "MpOperationalEvents.txt"
if (Test-Path MpOperationalEvents.txt.bak){
    Get-Content MpOperationalEvents.txt.bak, MpOperationalEvents.txt | Set-Content MpOperationalEvents-merged.txt
    $events_log = "MpOperationalEvents-merged.txt"
}

Convert-MpOperationalEventLogTxt -Path $events_log

# Get Signature information and fallback order
$defender_configs = Get-Content MpRegistry.json | ConvertFrom-Json
$signature_configs = $defender_configs."effective policy"."Signature Updates"
$shared_signatureroot = $signature_configs."SharedSignatureRoot"
$fallback_order = $signature_configs."FallbackOrder"


# Windows Update log
Get-WindowsUpdateLog -ETLPath $MPLOG_FOLDER -LogPath $MPLOG_FOLDER # default output: WindowsUpdate.log
# Get Windows Update configurations

# Validate them in logs.