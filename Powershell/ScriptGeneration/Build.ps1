Task Default -Depends RunPackage

Task PowershellModule {
	$modulePath = Join-Path -Path $psake.build_script_dir -ChildPath "SampleModule\SampleModule.psm1"
	Import-Module -Name $modulePath -Global -Force 
}

Task RunPackage -Depends GeneratePackage {
	$autoSysDeploymentScript = $script:AutoSysDeploymentByEnvironment["PROD"].UpgradeScriptPath
	Powershell -NoProfile -Command "& `"$autoSysDeploymentScript`""
}

Task GeneratePackage -Depends PowershellModule {
	$script:AutoSysUpgradeFolder = Join-Path -Path $psake.build_script_dir -ChildPath "AutoSysUpgrade"
	$script:ScriptOutputPath = Join-Path -Path $AutoSysUpgradeFolder -ChildPath "ScriptOutput.ps1"
	$autosysUpgradeTemplatePath = Join-Path $psake.build_script_dir -ChildPath "AutosysDeploymentTemplate.ps1"
	$script:AutoSysUpgradeScriptPath = Join-Path $AutoSysUpgradeFolder -ChildPath "AutoSysDeploy.ps1"
	
	$script:AutoSysRootFolder = Join-Path $psake.build_script_dir -ChildPath "AutoSys"
	$upgradeScriptPath = Join-Path $AutoSysRootFolder -ChildPath "Upgrade\1.2.0.0\AutoSysUpgrade.ps1"
	
	
	if ( Test-Path $AutoSysUpgradeFolder ) {
		Remove-Item $AutoSysUpgradeFolder -Recurse
	}
	New-Item -Path $AutoSysUpgradeFolder -ItemType Directory | Out-Null
	
	$environmentNames = @( "PROD", "PRE", "UAT01" )
	$script:AutoSysDeploymentByEnvironment = @{}
	foreach ( $environment in $environmentNames ) {
		$environmentSpecificUpgradeFolder = Join-Path -Path $AutoSysUpgradeFolder -ChildPath $environment
		$environmentSpecificUpgradeScriptPath = Join-Path -Path $environmentSpecificUpgradeFolder -ChildPath "AutoSysUpgrade.ps1"
		New-Item -Path $environmentSpecificUpgradeFolder -ItemType Directory | Out-Null
	
		$AutoSysDeploymentByEnvironment[$environment] = @{
			UpgradeFolder = $environmentSpecificUpgradeFolder
			UpgradeScriptPath = $environmentSpecificUpgradeScriptPath
			DeployCommands = @()
		}
	}
	
	#
	# Override the deployment functions with versions that can be used to
	# generate the actual autosys deployment scripts
	#
	function New-AutosysJob() {
		[CmdletBinding( )]
		param (
			[Parameter( Position = 1, Mandatory = $true )]
			$JilFileName
		)
		
		$JilPath = Join-Path -Path $script:AutoSysRootFolder -ChildPath ( "Jil\{0}" -f $JilFileName )
		
		if ( -not ( Test-Path $JilPath ) ) {
			throw ( "JIL {0} doesn't exist at {1}" -f $JilFileName, $JilPath )
		}
		
		foreach ( $environment in $script:AutoSysDeploymentByEnvironment.Keys ) {
			$environmentSpecificJilPath = Join-Path $script:AutoSysDeploymentByEnvironment[$environment].UpgradeFolder -ChildPath $JilFileName
			Copy-Item -Path $JilPath -Destination $environmentSpecificJilPath
			
			$script:AutoSysDeploymentByEnvironment[$environment].DeployCommands += ( "{0} {1}" -f $PSCmdlet.MyInvocation.MyCommand, $JilFileName )				
		}
	}
	
	function Update-AutosysJob() {
		[CmdletBinding( SupportsShouldProcess = $true )]
		param (
			[Parameter( Position = 1, Mandatory = $true )]
			$JilFileName
		)
		
		$JilPath = Join-Path -Path $script:AutoSysRootFolder -ChildPath ( "Jil\{0}" -f $JilFileName )
		
		if ( -not ( Test-Path $JilPath ) ) {
			throw ( "JIL {0} doesn't exist at {1}" -f $JilFileName, $JilPath )
		}
		
		foreach ( $environment in $script:AutoSysDeploymentByEnvironment.Keys ) {
			$environmentSpecificJilPath = Join-Path $script:AutoSysDeploymentByEnvironment[$environment].UpgradeFolder -ChildPath $JilFileName
			Copy-Item -Path $JilPath -Destination $environmentSpecificJilPath
			
			$script:AutoSysDeploymentByEnvironment[$environment].DeployCommands += ( "{0} {1}" -f $PSCmdlet.MyInvocation.MyCommand, $JilFileName )				
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
		
		foreach ( $environment in $script:AutoSysDeploymentByEnvironment.Keys ) {
			$environmentSpecificJobName = $JobName.Replace( "eee", $environment )
			$script:AutoSysDeploymentByEnvironment[$environment].DeployCommands += ( "{0} {1} {2}" -f $PSCmdlet.MyInvocation.MyCommand, $environmentSpecificJobName, $Event )				
		}
	}

	#
	# Execute the upgrade script in the context of the re-defined functions - this will enable the real
	# deployment scripts by environment to be generated
	#
	Invoke-Expression ". `"$upgradeScriptPath`""
	
	#
	# Write out the deployment scripts for each environment
	#
	$upgradeTemplate = Get-Content -Path $autosysUpgradeTemplatePath | Out-String
	foreach ( $environment in $AutoSysDeploymentByEnvironment.Keys ) {
		$deployContent = [string]::Join( "`n	", $AutoSysDeploymentByEnvironment[$environment].DeployCommands )
		$deployScriptContent = $upgradeTemplate.Replace( "# Deploy content here", $deployContent )
		$deployScriptPath = Join-Path -Path $AutoSysDeploymentByEnvironment[$environment].UpgradeFolder -ChildPath ""
		Set-Content -Path $AutoSysDeploymentByEnvironment[$environment].UpgradeScriptPath -Value $deployScriptContent	
	}
}