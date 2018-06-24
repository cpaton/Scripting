function Exec() {
    <#
    .SYNOPSIS
        Runs a command line tool and throws an exception depending on the exit code
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        # Description of the command, this is used in error messages if there is a failure
        [Parameter( Position = 1, Mandatory = $false )]
        [string] $Description,
        # String representing the command to run
        [Parameter( Position = 2, Mandatory = $true )]
        [string] $Command,
        # Additional exit codes other than 0 that should not trigger an exception to be thrown
        [Parameter( Position = 3, Mandatory = $false )]
        [array] $IgnoreExitCodes = $(@()),
        # Values which should be masked and not included in command output
        [Parameter(Mandatory = $false)]
        [array]
        $MaskedValues = $(@())
    )

    $safeCommand  = $Command
    foreach ($maskValue in $MaskedValues)
    {
        $safeCommand = $safeCommand.Replace($maskValue, "*****")
    }

    Write-Verbose $safeCommand
    if ( $PSCmdlet.ShouldProcess($Description, "Execute")) {
        Invoke-Expression $Command

        $exitCode = $LASTEXITCODE
        if ( ( $exitCode -ne 0 ) -and ( $IgnoreExitCodes -notcontains $exitCode ) ) {
            Write-Host $safeCommand
            throw ( "Command failed ({0}). {1}" -f $exitCode, $Description )
        }
    }
}

function Invoke-WebRequestWithExceptionLogging()
{
    <#
    .SYNOPSIS
    Sends a web request within an exception handler.  Writes more information about the exception to the log and then rethrows
    #>
    [CmdletBinding()]
    param(
        # Function that will send the web request
        [Parameter(Position = 1, Mandatory = $true)]
        [scriptblock]
        $RequestSender
    )

    try
    {
        . $RequestSender
    }
    catch [WebException]
    {
        $message = $_.Exception.Message
        Write-Host ( '{0} - {1}' -f $_.Exception.Response.ResponseUri, $message )
    
        if ( ($_.Exception.Response.ContentType -eq "application/json") -and ($_.Exception.Response.ContentLength -gt 0) )
        {
            $responseStream = $_.Exception.Response.GetResponseStream()
            $responseStream.Position = 0;
            $reader = New-Object System.IO.StreamReader($responseStream)
            $response = $reader.ReadToEnd()
            $reader.Close()        
    
            $jsonResponse = ConvertFrom-Json $response
            Write-Host $jsonResponse
        }
    
        throw
    }
}