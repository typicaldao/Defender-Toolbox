<#
.SYNOPSIS
    Read Windows Defender registry export file MpRegistry.txt

.DESCRIPTION
    This script reads a Windows Defender registry export file and returns an array of objects containing the registry keys and values.

.NOTES
File Name      : Read-MpRegistry.ps1
version        : 0.1

.EXAMPLE
    C:\'Program Files'\'Windows Defender'\MpCmdRun.exe -GetFiles
    copy 'C:\ProgramData\Microsoft\Windows Defender\Support\MpSupportFiles.cab'
    md $pwd\MpSupportFiles
    expand -R -I $pwd\MpSupportFiles.cab -F:* $pwd\MpSupportFiles

    To generate the MpRegistry.txt file

.EXAMPLE
    Read-MpRegistry -regFilePath "C:\temp\MpRegistry.reg"
    Read the MpRegistry.reg file and return the registry keys and values

.EXAMPLE
    Read-MpRegistry -regFilePath "C:\temp\MpRegistry.reg" -truncateBinaryValues -truncateLength 100
    Read the MpRegistry.reg file and return the registry keys and values with truncated binary values

.PARAMETER regFilePath
    The path to the registry export file MpRegistry.reg

.PARAMETER truncateBinaryValues
    Truncate binary values that exceed the specified length

.PARAMETER truncateLength
    The maximum length of binary values to display

#>
[cmdletbinding()]
param(
  [Parameter(Mandatory = $true)]
  [string]$regFilePath, # Path to the registry export file MpRegistry.reg
  [bool]$truncateBinaryValues = $true,
  [int]$truncateLength = 100
)

$global:mpRegistry = [System.Collections.ArrayList]::New()
$scriptName = "$psscriptroot\$($MyInvocation.MyCommand.Name)"

function Main() {
  try {
    if (!(Test-Path $regFilePath)) {
      Get-Help $scriptName -Examples
      Write-Error "The specified registry export file does not exist: $regFilePath"
      return
    }

    $regContent = [IO.File]::ReadAllLines($regFilePath)
    $optionPattern = "Current configuration options for location `"(?<configurationOption>.+)`""
    # Detect and process registry values
    $propertyPattern = '\s+(?<propertyName>.+?)\s+?\[(?<propertyType>\w+?)\]\s+:\s(?<propertyValue>.+)'

    # Read each line of the registry export file
    for ($i = 0; $i -lt $regContent.Length; $i++) {
      $line = $regContent[$i]
      Write-Verbose "processing line: $line"

      if ($line -match $optionPattern) {
        $configurationOption = $matches['configurationOption']
        continue
      }

      # Detect and process registry keys
      if ($line -match '^\[(.+)\]$') {
        $currentKey = $matches[1]
        $keyName = $currentKey

        if ($currentKey.Contains("\")) {
          $keyName = Split-Path -Path $currentKey -Leaf
        }
        Write-Host "current key: $currentKey" -ForegroundColor Green
        continue
      }

      $regexMatch = [regex]::Match($line, $propertyPattern)
      if ($regexMatch.Success) {
        $propertyName = $regexMatch.Groups['propertyName'].Value
        $propertyType = $regexMatch.Groups['propertyType'].Value
        $propertyValue = $regexMatch.Groups['propertyValue'].Value

        # Check for Multi-String values that span multiple lines
        if ($propertyType -eq 'REG_MULTI_SZ') {
          Write-Verbose "processing multi-string value: $line"
          $propertyValue = @($propertyValue)
          $lineIndex = $regexMatch.Groups['propertyValue'].Index

          for ($j = $i + 1; $j -lt $regContent.Length; $j++) {
            $nextLine = $regContent[$j]
            $regexMatch = [regex]::Match($nextLine, '\s+(?<propertyValue>.+)')

            if ($regexMatch.Success -and ($regexMatch.Groups['propertyValue'].Index -eq $lineIndex)) {
              $propertyValue += $regexMatch.Groups['propertyValue'].Value
            }
            else {
              $i = $j - 1
              break
            }
          }
        }
        elseif ($propertyType -eq "REG_BINARY" -and $propertyValue -and $truncateBinaryValues -and $propertyValue.Length -gt $truncateLength) {
          $propertyValue = $propertyValue.ToString().Substring(0, $truncateLength) + "..."
          $line = $line.Substring(0, $line.IndexOf($propertyValue) + $propertyValue.Length)
        }

        Write-Verbose "processing value: $line"
        $entry = [ordered]@{
          ConfigurationOption = $configurationOption
          Key                 = $currentKey
          KeyName             = $keyName
          PropertyName        = $propertyName
          PropertyType        = $propertyType
          PropertyValue       = $propertyValue
        }

        $global:mpRegistry += $entry
      }
    }

    return $global:mpRegistry
  }
  catch {
    Write-Host "exception::$($psitem.Exception.Message)`r`n$($psitem.scriptStackTrace)" -ForegroundColor Red
    return $null
  }
}

Main