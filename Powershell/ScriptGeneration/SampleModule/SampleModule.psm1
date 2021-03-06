function New-AutoSysJob() {
	[CmdletBinding( DefaultParameterSetName = "AutoSysPackage" )]
	param (
		[Parameter( Position = 1, Mandatory = $true, ParameterSetName = "AutoSysPackage" )]
		$JilPath,
		[Parameter( Position = 1, Mandatory = $true, ParameterSetName = "Module" )]
		$JobName,
		[Parameter( Position = 2, Mandatory = $true, ParameterSetName = "Module" )]
		$TargetEnvironment
	)
	
	"Runing New-AutoSysJob with parameter set {0}" -f $PSCmdlet.ParameterSetName
}
Export-ModuleMember -Function New-AutoSysJob

function Update-AutoSysJob() {
	[CmdletBinding( DefaultParameterSetName = "AutoSysPackage" )]
	param (
		[Parameter( Position = 1, Mandatory = $true, ParameterSetName = "AutoSysPackage" )]
		$JilPath,
		[Parameter( Position = 1, Mandatory = $true, ParameterSetName = "Module" )]
		$JobName,
		[Parameter( Position = 2, Mandatory = $true, ParameterSetName = "Module" )]
		$TargetEnvironment
	)
	
	"Runing Update-AutoSysJob with parameter set {0}" -f $PSCmdlet.ParameterSetName
}
Export-ModuleMember -Function Update-AutoSysJob

function Set-AutoSysJobStatus() {
	[CmdletBinding( DefaultParameterSetName = "AutoSysPackage" )]
	param (
		[Parameter( Position = 1, Mandatory = $true, ParameterSetName = "AutoSysPackage" )]
		$JobName,
		[Parameter( Position = 2, Mandatory = $true, ParameterSetName = "AutoSysPackage" )]
		$Status,
		[Parameter( Position = 3, Mandatory = $false, ParameterSetName = "Module" )]
		$TargetEnvironment
	)
	
	"Runing Set-AutoSysJobStatus with parameter set {0}" -f $PSCmdlet.ParameterSetName
}
Export-ModuleMember -Function Set-AutoSysJobStatus
