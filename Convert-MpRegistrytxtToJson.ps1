function Convert-MpRegistrytxtToJson {
    param(
        [string]$Path="$PWD\MpRegistry.txt",
        [string]$OutFile="$PWD\MpRegistry.json"
    )

    $lines = Get-Content $Path
    $json_result = [ordered]@{}

    $Section = ""
    $next_Section= ""
    $this_line = 0
    $depth = 1
    $regs = ""
    $other_regs = $false

    # NIS is a special one. Initialize it here.
    
    :DefenderAV foreach ($line in $lines) {
        # Match the root category
        if ($line.StartsWith("Current configuration options for location")){
            $policy = ($line -split '"')[1]
            $json_result.$policy = [ordered]@{}
        }

        elseif ($line.StartsWith("Windows Setup keys from")) {
            $other_regs = $true
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
            switch ($depth){
                1 { $json_result.$policy.($regs[0]).$key = $value; break }
                2 { $json_result.$policy.($regs[0]).($regs[1]).$key = $value; break}
                3 { $json_result.$policy.($regs[0]).($regs[1]).($regs[2]).$key = $value; break }
                4 { $json_result.$policy.($regs[0]).($regs[1]).($regs[2]).($regs[3]).$key = $value; break }
            }
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

    $json = $json_result | ConvertTo-Json -Depth 4 -WarningAction Ignore
    $json | Out-File -FilePath $OutFile -Force
}