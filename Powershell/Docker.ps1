function Get-DockerClientBundle()
{
    <#
    .SYNOPSIS
        Downloads the Docker EE client bundle used to connect to the remote docker swarm and configured
        the current process to communicate to that swarm
    #>
    [CmdletBinding()]
    param(
        # Directory to store the client bundle.  Defaults to a random directory in the temp folder
        [Parameter(Mandatory = $false)]
        [string]
        $WorkingDirectory = $(Join-Path -Path ([IO.Path]::GetTempPath()) -ChildPath ([IO.Path]::GetRandomFileName()))
    )

    $clientBundleUrl = Get-DockerApiUrl "api/clientbundle"
    $headers = Get-DockerApiAuthHeader

    $clientBundleZipPath = Join-Path -Path $WorkingDirectory -ChildPath "docker-client-bundle.zip"
    $clientBundlePath = Join-Path -Path $WorkingDirectory -ChildPath "docker-client-bundle"

    Invoke-DockerApi {
        Write-Host "Downloading Docker client bundle to $clientBundleZipPath"
        Invoke-WebRequest -Uri $clientBundleUrl -Method Get -Headers $headers -OutFile $clientBundleZipPath
    }

    Write-Host "Expanding client bundle to $clientBundlePath"
    Expand-Archive -Path $clientBundleZipPath -DestinationPath $clientBundlePath -Force

    $clientBundleInitialisationPath = Join-Path -Path $clientBundlePath -ChildPath "env.ps1"
    Invoke-Expression $clientBundleInitialisationPath
}

function Get-DockerApiAuthHeader()
{
    <#
    .SYNOPSIS
        Builds the Docker authorization header used to communicate with the Docker UCP REST API

    .OUTPUTS
        [hashtable] Authorization header
    #>
    [CmdletBinding()]
    param()

    $authToken = Get-DockerAuthToken
    $authHeaderValue = 'Bearer {0}' -f $authToken
    @{ Authorization = $authHeaderValue }
}

function Get-DockerAuthToken()
{
    <#
    .SYNOPSIS
        Retrieves the auth token for communicating with the docker API

    .OUTPUTS
        [string] Authentication token which can be used to communicate with the API
    #>
    [CmdletBinding()]
    param(
        # Credentials used for the Docker API
        [Parameter(Mandatory = $true)]
        [pscredential]
        $Credential
    )

    if ( $script:dockerAuthToken -ne $null )
    {
        return $script:dockerAuthToken
    }

    $dockerUser = $Credential.UserName
    $dockerPassword = $Credential.GetNetworkCredential().Password

    $authUri = Get-DockerApiUrl "auth/login"
    $credentials = '{{ "username" : "{0}", "password" : "{1}" }}' -f $dockerUser, $dockerPassword

    . Invoke-DockerApi {
        Write-Host "Getting docker auth token for user $dockerUser"
        $authResponse = Invoke-WebRequest -Uri $authUri -Method Post -Body $credentials -ContentType "application/json"
        $authToken = ConvertFrom-Json $authResponse.Content
    }

    $script:dockerAuthToken = $authToken.auth_token
    $script:dockerAuthToken
}

function Invoke-DockerApi()
{
    <#
    .SYNOPSIS
        Calls a Docker API configuring TLS as required
    #>
    [CmdletBinding()]
    param(
        # Scriptblock representing the call to the API
        [Parameter(Position = 1, Mandatory = $true)]
        [scriptblock]
        $ApiCall
    )

    # Docker API needs TLS 1.2 and has self-signed certificates
    Enable-TlsAllProtocols
    $revertTlsCertCheck = Disable-TlsCertificateChecks

    try
    {
        . Invoke-WebRequestWithExceptionLogging $ApiCall
    }
    finally
    {
        . $revertTlsCertCheck
    }
}

function Get-DockerApiUrl()
{
    <#
    .SYNOPSIS
        Builds a URL to a docker API endpoint

    .OUTPUTS
        [uri] Uri to the Docker API endpoint
    #>
    [CmdletBinding()]
    param
    (
        # Root Docker URL
        [Parameter(Position = 1, Mandatory = $true)]
        [string]
        $DockerUrl,
        # Relative URL to the desired endpoint
        [Parameter(Position = 2, Mandatory = $true)]
        [string]
        $ApiUrl
    )

    $authUriBuilder = New-Object System.UriBuilder -ArgumentList $DockerUrl
    $authUriBuilder.Path += $ApiUrl

    $authUriBuilder.Uri
}