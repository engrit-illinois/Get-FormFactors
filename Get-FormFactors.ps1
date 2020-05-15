function Get-FormFactors {
	
	param(
		[Parameter(Mandatory=$true)]
		[string]$OUDN,
		
		[int]$Verbosity=0,
		
		[int]$CIMTimeoutSec=15,
		
		[string]$LogPath="c:\engrit\logs\Get-FormFactors_$(Get-Date -Format `"yyyy-MM-dd_HH-mm-ss-ffff`").log",
		
		[switch]$DisableCIMFallbacks
	)
	
	$CSVPATH = $LogPath -replace "\.log",".csv"
	
	function log {
		param (
			[string]$msg,
			[int]$l=0, # level (of indentation)
			[int]$v=0, # verbosity level
			[switch]$nots, # omit timestamp
			[switch]$nnl # No newline after output
		)
		
		if(!(Test-Path -PathType leaf -Path $LogPath)) {
			$shutup = New-Item -ItemType File -Force -Path $LogPath
		}
		
		for($i = 0; $i -lt $l; $i += 1) {
			$msg = "    $msg"
		}
		if(!$nots) {
			$ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss:ffff"
			$msg = "[$ts] $msg"
		}
		
		if($v -le $Verbosity) {
			if($nnl) {
				Write-Host $msg -NoNewline
			}
			else {
				Write-Host $msg
			}
			
			if(!$NoLog) {
				if($nnl) {
					$msg | Out-File $LogPath -Append -NoNewline
				}
				else {
					$msg | Out-File $LogPath -Append
				}
			}
		}
	}

	
	function Get-CompNames {
		log "Getting list of computer names in OU: `"$OUDN`"..."
		$compNames = (Get-ADComputer -SearchBase $OUDN -Filter "*" | Select Name).Name
		if($compNames) {
			log "Found $(@($compNames).count) computers in given OU." -l 1
		}
		else {
			log "No computers found in given OU!" -l 1
		}
		$compNames
	}

	
	# Make array of objects representing computers
	function Get-CompObjects($comps) {
		log "Making array of objects to represent each computer..."
	
		# Make sure $comp is treated as an array, even if it has only one member
		# Not sure if this is necessary, but better safe than sorry
		$comps = @($comps)
		
		# Make new array to hold objects representing computers
		$compObjects = @()
		
		foreach($thisComp in $comps) {
			$thisCompHash = @{
				"Name" = $thisComp
				"CS_Manufacturer" = $null
				"CS_Model" = $null
				"CS_ChassisSKUNumber" = $null
				"CS_SystemFamily" = $null
				"CS_SystemSKUNumber" = $null
				"CS_SystemType" = $null
				"CS_TotalPhysicalMemory" = $null
				"SE_Manufacturer" = $null
				"SE_Model" = $null
				"SE_ChassisTypes" = $null
				"SE_SerialNumber" = $null
				"SE_SMBIOSAssetTag" = $null
				
			}
			$thisCompObject = New-Object PSObject -Property $thisCompHash
			$compObjects += @($thisCompObject)
		}
		
		log "Done making computer object array." -v 2
		$compObjects
	}
	
	function Get-CompData($comps) {
		log " " -nots
		log "Getting data for all computers..."
		$num = 1
		foreach($comp in $comps) {
			$thisCompName = $comp.name
			$count = @($comps).count
			$completion = ([math]::Round($num / $count, 2)) * 100
			log "Getting data for computer $num/$count ($completion%): `"$thisCompName`"..." -l 1
			$num += 1
			
			$comp = Get-Model $comp
						
			log " " -nots -v 1
			log "Done getting data for computer: `"$thisCompName`"." -l 1
			
			log " " -nots
			log " " -nots -v 1
			log " " -nots -v 1
		}
		log "Done getting data for all computers."
		
		$comps
	}
	
	function Get-Model($comp) {
		$compName = $comp.name
		log " " -nots -v 1
		log "Getting model data for computer: `"$compName`"..." -l 2
		
		if(Test-Connection $compName -Quiet -Count 1) {
			log "Computer `"$compName`" responded." -l 3 -v 2
			
			$cimClass = "Win32_ComputerSystem"
			$cimErrorAction = "Stop"
			
			# Try CIM first as it's the easiest
			log "Trying CIM..." -l 3 -v 1
			try {
				$info = Get-CIMInstance -ComputerName $compName -Class $cimClass -ErrorAction $cimErrorAction -OperationTimeoutSec $CIMTimeoutSec
				$make = $info.Manufacturer
				$model = $info.Model
			}
			catch {
				log "CIM didn't work." -l 3 -v 1
				Log-Error $_ -v 2
			}
			
			# If CIM doesn't work
			if((!$info) -and (!$DisableCIMFallbacks)) {
				log "Trying Invoke-Command to use WMI locally..." -l 3 -v 1
				try {
					# Try Invoke-Command
					# In some cases (Win7 + PSv2), CIM and remote WMI were not working, but this did for some reason
					
					$info = Invoke-Command -ComputerName $compName -ErrorAction $cimErrorAction -ScriptBlock { Get-WMIObject -Class $cimClass -ErrorAction $cimErrorAction }
					$make = $info.Manufacturer
					$model = $info.Model
				}
				catch {
					log "Invoke-Command didn't work." -l 3 -v 1
					Log-Error $_ -v 2
				}
			}
				
			# If Invoke-Command doesn't work
			if((!$info) -and (!$DisableCIMFallbacks)) {
				log "Trying WMI..." -l 3 -v 1
				try {
					# Fall back to WMI
					$info = Get-WMIObject -ComputerName $compName -Class $cimClass -ErrorAction $cimErrorAction
					$make = $info.Manufacturer
					$model = $info.Model
				}
				catch {
					log "WMI didn't work. I give up." -l 3 -v 1
					Log-Error $_ -v 2
				}
			}
			
			$comp.Make = $make
			$comp.Model = $model
			
			if($info) {
				log "Model is `"$make $model`"." -l 3
			}
			else {
				log "Model not retrieved from computer: `"$compName`"!" -l 3
			}
		}
		else {
			log "Computer `"$compName`" did not respond!" -l 3
		}
		log "Done getting Model for computer: `"$compName`"..." -l 2 -v 2
		
		$comp
	}
	
	function Get-ChassisType($comp) {
		$compName = $comp.name
		log " " -nots -v 1
		log "Getting chassis type data for computer: `"$compName`"..." -l 2
		
		if(Test-Connection $compName -Quiet -Count 1) {
			log "Computer `"$compName`" responded." -l 3 -v 2
			
			$cimClass = "Win32_SystemEnclosure"
			$cimErrorAction = "Stop"
			
			# Try CIM first as it's the easiest
			log "Trying CIM..." -l 3 -v 1
			try {
				$info = Get-CIMInstance -ComputerName $compName -Class $cimClass -ErrorAction $cimErrorAction -OperationTimeoutSec $CIMTimeoutSec
				$make = $info.Manufacturer
				$model = $info.Model
			}
			catch {
				log "CIM didn't work." -l 3 -v 1
				Log-Error $_ -v 2
			}
			
			# If CIM doesn't work
			if((!$info) -and (!$DisableCIMFallbacks)) {
				log "Trying Invoke-Command to use WMI locally..." -l 3 -v 1
				try {
					# Try Invoke-Command
					# In some cases (Win7 + PSv2), CIM and remote WMI were not working, but this did for some reason
					
					$info = Invoke-Command -ComputerName $compName -ErrorAction $cimErrorAction -ScriptBlock { Get-WMIObject -Class $cimClass -ErrorAction $cimErrorAction }
					$make = $info.Manufacturer
					$model = $info.Model
				}
				catch {
					log "Invoke-Command didn't work." -l 3 -v 1
					Log-Error $_ -v 2
				}
			}
				
			# If Invoke-Command doesn't work
			if((!$info) -and (!$DisableCIMFallbacks)) {
				log "Trying WMI..." -l 3 -v 1
				try {
					# Fall back to WMI
					$info = Get-WMIObject -ComputerName $compName -Class $cimClass -ErrorAction $cimErrorAction
					$make = $info.Manufacturer
					$model = $info.Model
				}
				catch {
					log "WMI didn't work. I give up." -l 3 -v 1
					Log-Error $_ -v 2
				}
			}
			
			$comp.Make = $make
			$comp.Model = $model
			
			if($info) {
				log "Model is `"$make $model`"." -l 3
			}
			else {
				log "Model not retrieved from computer: `"$compName`"!" -l 3
			}
		}
		else {
			log "Computer `"$compName`" did not respond!" -l 3
		}
		log "Done getting Model for computer: `"$compName`"..." -l 2 -v 2
		
		$comp
	}

	function Do-Stuff {
		log " " -nots
		
		$compNames = Get-CompNames
		if($compNames) {
			$comps = Get-CompObjects $compNames
			$comps = Get-CompData $comps
		}
	}
	
	Do-Stuff
	
	log "EOF"
	log " " -nots
	
	log "test"

}