function Convert-MacDlpPolicyBin {
    # inspired by MDE client analyzer: MDEClientAnalyzer.ps1
    param (
        [string]$Path="$PWD\dlp_policy.bin",
        [string]$OutFile="$PWD\configuration.txt"
    )
    if ((Test-Path $Path) -eq $false ) {
        Write-Host "Policy file not found. Exiting."
        return
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