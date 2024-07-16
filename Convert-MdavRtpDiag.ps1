function Convert-MdavRtpDiag() {
    param(
        [Parameter(Mandatory=$true)]$rtplog,
        $OutFile = "$rtplog.txt"
    )
    if (Test-Path $rtpLog){
        $val = Get-Content $rtpLog | ConvertFrom-Json
        $val.counters | ForEach-Object {$_.totalFilesScanned = [int32]$_.totalFilesScanned} # Convert totalFileScanned type from string to integer.
        $val.counters | Where-Object {$PSItem.totalFilesScanned -gt 0} | Select-Object id, name, totalFilesScanned, path | Sort-Object -Property totalFilesScanned -Descending | Format-Table | Out-File -Force $OutFile
    }
    else {
        Write-Host "Cannot find the file $rtplog. Exit."
        return
    }
    
    Write-Host "File saved as $OutFile successfully."
}