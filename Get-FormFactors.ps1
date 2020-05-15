function Get-FormFactors {
	
	param(
		[Parameter(Mandatory=$true)]
		[string]$Collection
	)
	
	function log {
		param(
			[string]$msg
		)
		
		Write-Host $msg
	}
	
	log $Collection
}