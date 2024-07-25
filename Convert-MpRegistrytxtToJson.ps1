function Convert-MpRegistrytxtToJson {
    param(
        [string]$Path="$PWD\MpRegistry.txt",
        [string]$OutFile="$PWD\MpRegistry.json"
    )

    $lines = Get-Content $Path
    $jsonObject = [ordered]@{}

    # Split text data into lines
    # $lines = $MpRegistryText -split "\r?\n"

    $currentSection = ""
    foreach ($line in $lines) {
        if ($line.StartsWith("Current configuration options for location")){
            $policy = ($line -split '"')[1]
            $jsonObject.$policy = [ordered]@{}
        }
        elseif ($line -match '^\[(.+)\]$') {
            $currentSection = $matches[1]
            $jsonObject.$policy.$currentSection = [ordered]@{}
        } 
        elseif ($line -match '^\s{4}(\S+)\s+\[.+\]\s+:\s(.+)$') {
            $key = $matches[1]
            $value = $matches[2]
            $jsonObject.$policy.$currentSection.$key = $value
        }
}
    $json = $jsonObject | ConvertTo-Json
    $json | Out-File -FilePath $OutFile -Force
}