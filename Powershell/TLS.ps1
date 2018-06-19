function Disable-TlsCertificateChecks()
{
    <#
    .SYNOPSIS
        Disable TLS certificate checking by trusting all certificates

    .DESCRIPTION
        Useful within a trusted network when communicating with services using self-signed certificates

    .OUTPUTS
        [Scriptblock] Function which when executed returns TLS certificate checks back to its previous setting
    #>
    [CmdletBinding()]
    param ()

    if ( -not ( ([System.Management.Automation.PSTypeName]'TrustAllCertsPolicy').Type ) )
    {
        Add-Type @"
using System.Net;
using System.Security.Cryptography.X509Certificates;
public class TrustAllCertsPolicy : ICertificatePolicy {
    public bool CheckValidationResult(
        ServicePoint srvPoint, X509Certificate certificate,
        WebRequest request, int certificateProblem) {
        return true;
    }
}
"@
    }

    $existingSetting = [System.Net.ServicePointManager]::CertificatePolicy

    [System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy

    { [System.Net.ServicePointManager]::CertificatePolicy = $existingSetting }
}


function Enable-TlsAllProtocols()
{
    <#
    .SYNOPSIS
        By default .Net only supports SSL v3 and TLS 1.0.  This function also enables TLS 1.1 and TLS 1.2
    #>
    [CmdletBinding()]
    param()

    $AllProtocols = [System.Net.SecurityProtocolType]'Ssl3,Tls,Tls11,Tls12'
    [System.Net.ServicePointManager]::SecurityProtocol = $AllProtocols
}