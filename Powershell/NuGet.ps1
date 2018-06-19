function Install-NugetPackage() {
    <#
    .SYNOPSIS
        Installs a NugGet package

    .OUTPUTS
        [string] Full path to the installed nuget package
    #>
    [CmdletBinding( SupportsShouldProcess = $true )]
    param (
        # ID of the package to install
        [Parameter(Position = 1, Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $PackageId,
        # Version of the package to install
        [Parameter(Position = 2, Mandatory = $true)]
        [version]
        $PackageVersion,
        # Nuget source to use to locate the package
        [Parameter(Position = 3, Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $NugetSource,
        # Directory where the package should be installed
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $PackagesDirectory,
        # Optional path to nuget.exe. If not specified nuget must be available on the path
        [Parameter(Mandatory = $false)]
        [string]
        $NugetExePath
    )

    if ( -not $NugetExePath)
    {
        $nugetCommand = Get-Command nuget.exe -ErrorAction SilentlyContinue
        if ( $nugetCommand )
        {
            $NugetExePath = $nugetCommand.Definition
        }
        else
        {
            throw "nuget.exe not found"
        }
    }

    $installedPackagePath = Join-Path -Path $PackagesDirectory -ChildPath ("{0}.{1}" -f $PackageId, $PackageVersion )

    if ( -not ( Test-Path $installedPackagePath ) ) {
        if ( $PSCmdlet.ShouldProcess( ( "{0} v{1}" -f $PackageId, $PackageVersion ), "Install" ) )
        {
            $nugetCommand = '& "{0}" install {1} -Version {2} -Source "{3}" -OutputDirectory "{4}" -PackageSaveMode nuspec -NonInteractive' -f $NugetExePath, $PackageId, $PackageVersion.ToString(), $NugetSource, $PackagesDirectory
            Exec ( "Installing nuget package {0} v{1}" -f $PackageId, $PackageVersion ) $nugetCommand
        }
    }

    $installedPackagePath
}