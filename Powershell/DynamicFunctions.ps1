function GenerateFunctionAndExportIntoCallingModule() {
	[CmdletBinding()]
	param (
		[Parameter( Position = 1, Mandatory = $true )]
		$FunctionName,
		[Parameter( Position = 2, Mandatory = $true )]
		$FunctionContent,
		[Parameter( Position = 3, Mandatory = $true )]
		$CallingModuleScope
		
	)
	
	& $CallingModuleScope New-Item -Path "Function:$FunctionName" -Value $FunctionContent -ItemType Function -OutVariable $discarded 
	& $CallingModuleScope Export-ModuleMember -Function $FunctionName	
}

function New-EnvironmentManagementFunctionsForComponent() {
	[CmdletBinding()]
	param (
		$Component
	)	
	
	CreateLocalEnvironmentUpdateFunctions $Component $PSCmdlet.SessionState.Module
	CreateLocalRestoreSnapshotFunctions $Component $PSCmdlet.SessionState.Module
}

function CreateLocalRestoreSnapshotFunctions() {
	[CmdletBinding()]
	param (
		$Component,
		$callingModuleScope
	)
	
	$FunctionName = "Update-Local" + $Component.ComponentName.Replace( "-", "" )  + "DatabaseFromIntegrationTestSnapshot"
	$FunctionContent = { 
			param(
				[int]
				$Snapshot = $(0)
			)
			Update-LocalDatabaseFromIntegrationTestSnapshot $Component @psBoundParameters	
		}
		
	GenerateFunctionAndExportIntoCallingModule $FunctionName $FunctionContent $callingModuleScope		
}

function CreateLocalEnvironmentUpdateFunctions() {
	[CmdletBinding()]
	param (
		$Component,
		$callingModuleScope
	)
	
	$FunctionName = "Update-Local" + $Component.ComponentName.Replace( "-", "" )  + "Environment"
	
	if ( $Component.HasSsas ) {
		$FunctionContent = { 
			param (
				[SWITCH] $Database,
				[SWITCH] $Ssis,
				[SWITCH] $PackageConfiguration,
				[SWITCH] $EtlJobDefinitions,
				[SWITCH] $SqlJob,
				[SWITCH] $Ssas,
				[SWITCH] $All
			)	
			
			Update-LocalEnvironment $Component @psBoundParameters	
		}
	}
	else {
		$FunctionContent = { 
			param (
				[SWITCH] $Database,
				[SWITCH] $Ssis,
				[SWITCH] $PackageConfiguration,
				[SWITCH] $EtlJobDefinitions,
				[SWITCH] $SqlJob,
				[SWITCH] $All
			)	
			
			Update-LocalEnvironment $Component @psBoundParameters	
		}	
	}
	
	& $callingModuleScope New-Item -Path "Function:$FunctionName" -Value $FunctionContent -ItemType Function
	& $callingModuleScope Export-ModuleMember -Function $FunctionName
}