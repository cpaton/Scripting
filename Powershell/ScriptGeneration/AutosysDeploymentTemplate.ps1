[CmdletBinding( SupportsShouldProcess = $true )]
param (
)

$ErrorActionPreference = "Stop"

function Deploy() {
	[CmdletBinding( SupportsShouldProcess = $true )]
	param(
	)
	
	# Deploy content here
}

function New-AutosysJob() {
	[CmdletBinding( SupportsShouldProcess = $true )]
	param (
		[Parameter( Position = 1, Mandatory = $true )]
		$JilFileName
	)
	
	$jilName = [IO.Path]::GetFileNameWithoutExtension( $JilFileName )
	if ( $PSCmdlet.ShouldProcess( $jilName, "Create job" ) ) {
		Log ( "Creating job from {0}" -f $JilFileName )
	}
}

function Update-AutosysJob() {
	[CmdletBinding( SupportsShouldProcess = $true )]
	param (
		[Parameter( Position = 1, Mandatory = $true )]
		$JilFileName
	)
	
	$jilName = [IO.Path]::GetFileNameWithoutExtension( $JilFileName )
	if ( $PSCmdlet.ShouldProcess( $jilName, "Update job" ) ) {
		Log ( "Updating job from {0}" -f $JilFileName )
	}
}

function Set-AutoSysJobStatus() {
	[CmdletBinding()]
	param (
		[Parameter( Position = 1, Mandatory = $true )]
		$JobName,
		[Parameter( Position = 2, Mandatory = $true )]
		$Event
	)
	
	if ( $PSCmdlet.ShouldProcess( $JobName, ( "Send event {0}" -f $Event ) ) ) {
		Log ( "Sending event {0} to job {1}" -f $Event, $JobName )
	}
}

function Log() {
	[CmdletBinding()]
	param(
		[Parameter( Position = 1, Mandatory = $true )]
		$message
	)
	
	Write-Host ( "[{0:yyyy-MM-dd HH:mm:ss}] {1}" -f [DateTime]::Now, $message )
	
}

function GetScriptPath() {
	[CmdletBinding()]
	param(
	)
	
	$PSCmdlet.MyInvocation.ScriptName
}
$deployScript = GetScriptPath 	
$deployFolder = Split-Path -Path $deployScript -Parent
$deployScriptName = [IO.Path]::GetFileNameWithoutExtension( $deployScript )
$logFilePath = Join-Path $deployFolder ( "{0}-{1:yyyyMMddHHmmss}.log" -f $deployScriptName, [DateTime]::Now )

Start-Transcript -Path $logFilePath -WhatIf:$false

Log (  "*" * 100 )
Log ( "Deploy Script : {0}" -f $deployScript )
Log ( "Run           : {0:dd MMMM yyyy HH:mm:ss}" -f [DateTime]::Now )
Log ( "Run On        : {0}" -f $env:COMPUTERNAME )
Log ( "Run By        : {0}\{1}" -f $env:USERDOMAIN, $env:USERNAME )
Log (  "*" * 100 )

Push-Location $deployFolder
try {
	Deploy
}
finally {
	Pop-Location
}