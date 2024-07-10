# Plan: Add event ID as filter in parameter.

function Convert-MpOperationalEventLogTxt {
    param(
        [string]$Path="$PWD\MpOperationalEvents.txt",
        [string]$OutFile="$PWD\MpOperationalEvents.csv"
    )
    
    if ((Test-Path $Path) -eq $false ) {
        Write-Host "$Path is not found! Exit"
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
            $Level, $Id = $_.Split()[5], $_.Split()[8] 
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