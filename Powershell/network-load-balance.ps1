[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"

$targets = @(
    'bbc.co.uk'
    'cnn.com'
    'fxohub.tpicap.com'
    'netflix.com'
    'microsoft.com'
)

$providers = @{
    '192.168.1.1' = "Trooli"
    '195.166.130.249' = "Plusnet"
    '195.166.130.250' = "Plusnet"
}

$longestTarget = ( $targets | ForEach-Object { $_.Length } | Measure-Object -Maximum ).Maximum

foreach ($i in @(1..2))
{
    foreach ($target in $targets)
    {
        $result = Test-Connection `
            -Traceroute $target `
            -IPv4 `
            -ResolveDestination:$false `
            -MaxHops 2 `
            -ErrorAction SilentlyContinue
        
        $secondHopHost = $result[-1].Hostname
        $provider = "Unknown ($($secondHopHost))"
        if ($providers.ContainsKey($secondHopHost))
        {
            $provider = $providers[$secondHopHost]
        }

        Write-Host "$($target.PadLeft($longestTarget)) : $($provider)"
    }
}