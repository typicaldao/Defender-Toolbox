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