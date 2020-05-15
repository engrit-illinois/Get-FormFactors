function Get-FormFactors {
	
	param(
		[Parameter(Mandatory=$true)]
		[string]$OUDN,
		
		[int]$Verbosity=0,
		
		[int]$CIMTimeoutSec=15,
		
		[string]$LogPath="c:\engrit\logs\Get-FormFactors_$(Get-Date -Format `"yyyy-MM-dd_HH-mm-ss-ffff`").log",
		
		[switch]$DisableCIMFallbacks
	)
	
	$CSV_PATH = $LogPath -replace "\.log",".csv"
	
	# https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/win32-systemenclosure#members
	$CHASSIS_TYPES = @(
		[PSCustomObject]@{"Id" = 1; "Name" = "Other"},
		[PSCustomObject]@{"Id" = 2; "Name" = "Unknown"},
		[PSCustomObject]@{"Id" = 3; "Name" = "Desktop"},
		[PSCustomObject]@{"Id" = 4; "Name" = "Low Profile Desktop"},
		[PSCustomObject]@{"Id" = 5; "Name" = "Pizza Box"},
		[PSCustomObject]@{"Id" = 6; "Name" = "Mini Tower"},
		[PSCustomObject]@{"Id" = 7; "Name" = "Tower"},
		[PSCustomObject]@{"Id" = 8; "Name" = "Portable"},
		[PSCustomObject]@{"Id" = 9; "Name" = "Laptop"},
		[PSCustomObject]@{"Id" = 10; "Name" = "Notebook"},
		[PSCustomObject]@{"Id" = 11; "Name" = "Hand Held"},
		[PSCustomObject]@{"Id" = 12; "Name" = "Docking Station"},
		[PSCustomObject]@{"Id" = 13; "Name" = "All in One"},
		[PSCustomObject]@{"Id" = 14; "Name" = "Sub Notebook"},
		[PSCustomObject]@{"Id" = 15; "Name" = "Space-Saving"},
		[PSCustomObject]@{"Id" = 16; "Name" = "Lunch Box"},
		[PSCustomObject]@{"Id" = 17; "Name" = "Main System Chassis"},
		[PSCustomObject]@{"Id" = 18; "Name" = "Expansion Chassis"},
		[PSCustomObject]@{"Id" = 19; "Name" = "SubChassis"},
		[PSCustomObject]@{"Id" = 20; "Name" = "Bus Expansion Chassis"},
		[PSCustomObject]@{"Id" = 21; "Name" = "Peripheral Chassis"},
		[PSCustomObject]@{"Id" = 22; "Name" = "Storage Chassis"},
		[PSCustomObject]@{"Id" = 23; "Name" = "Rack Mount Chassis"},
		[PSCustomObject]@{"Id" = 24; "Name" = "Sealed-Case PC"}
	)

	
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
	
	function Log-Error {
		param(
			[string]$e,
			[int]$v=0
		)
		
		if($v -le $Verbosity) {
			log "$($e.Exception.Message)" -l 3
			log "$($e.InvocationInfo.PositionMessage.Split("`n")[0])" -l 4
		}
	}
	
	function Get-ChassisTypeFriendlyName($type) {
		($CHASSIS_TYPES | Where { $_.Id -eq $type } | Select Name).Name
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
			$thisComp = [PSCustomObject]@{
				"Name" = $thisComp
				"Error" = $false
				"ErrorMsg" = $null
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
				"SE_ChassisTypesFriendly" = $null
				"SE_SerialNumber" = $null
				"SE_SMBIOSAssetTag" = $null
			}
			$compObjects += @($thisComp)
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
			
			if(Test-Connection -ComputerName $thisCompName -Count 1 -Quiet) {
				$comp = Get-ComputerSystem $comp
				if(!$comp.Error) {
					$comp = Get-SystemEnclosure $comp
				}
			}
			else {
				log "Computer did not respond to ping!" -l 2
				$comp.Error = $true
				$comp.ErrorMsg = "No pong"
			}
			
			log " " -nots -v 1
			log "Done getting data for computer: `"$thisCompName`"." -l 1
			
			log " " -nots
			log " " -nots -v 1
			log " " -nots -v 1
		}
		log "Done getting data for all computers."
		
		$comps
	}
	
	function Get-ComputerSystem($comp) {
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
				}
				catch {
					log "WMI didn't work. I give up." -l 3 -v 1
					Log-Error $_ -v 2
				}
			}
			
			if($info) {
				$comp.CS_Manufacturer = $info.Manufacturer
				$comp.CS_Model = $info.Model
				$comp.CS_ChassisSKUNumber = $info.ChassisSKUNumber
				$comp.CS_SystemFamily = $info.SystemFamily
				$comp.CS_SystemSKUNumber = $info.SystemSKUNumber
				$comp.CS_SystemType = $info.SystemType
				$comp.CS_TotalPhysicalMemory = $info.TotalPhysicalMemory
				
				log "Model is `"$($comp.CS_Manufacturer)`" `"$($comp.CS_Model)`"." -l 3
			}
			else {
				log "Data not retrieved from computer: `"$compName`"!" -l 3
				$comp.Error = $true
				$comp.ErrorMsg = "Could not retrieve Win32_ComputerSystem data"
			}
		}
		else {
			log "Computer `"$compName`" did not respond!" -l 3
		}
		log "Done getting model for computer: `"$compName`"..." -l 2 -v 2
		
		$comp
	}
	
	function Get-SystemEnclosure($comp) {
		$compName = $comp.name
		log " " -nots -v 1
		log "Getting Win32_SystemEnclosure data for computer: `"$compName`"..." -l 2
		
		if(Test-Connection $compName -Quiet -Count 1) {
			log "Computer `"$compName`" responded." -l 3 -v 2
			
			$cimClass = "Win32_SystemEnclosure"
			$cimErrorAction = "Stop"
			
			# Try CIM first as it's the easiest
			log "Trying CIM..." -l 3 -v 1
			try {
				$info = Get-CIMInstance -ComputerName $compName -Class $cimClass -ErrorAction $cimErrorAction -OperationTimeoutSec $CIMTimeoutSec
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
				}
				catch {
					log "WMI didn't work. I give up." -l 3 -v 1
					Log-Error $_ -v 2
				}
			}
			
			if($info) {
				
				$comp.SE_Manufacturer = $info.Manufacturer
				$comp.SE_Model = $info.Model
				$comp.SE_SerialNumber = $info.SerialNumber
				$comp.SE_SMBIOSAssetTag = $info.SMBIOSAssetTag
				
				if(@($comp.SE_ChassisTypesFriendly).count -eq 1) {
					$comp.SE_ChassisTypes = [int]$info.ChassisTypes[0]
					$comp.SE_ChassisTypesFriendly = Get-ChassisTypeFriendlyName $comp.SE_ChassisTypes[0]
				}
				else {
					log "Computer has more than one ChassisTypes!"
					$chassisTypesString = ""
					foreach($type in $comp.SE_ChassisTypes) {
						$thisChassisType = [int]$type
						$chassisTypesString += "$thisChassisType "
					}
					$chassisTypesString.TrimEnd()
					$comp.SE_ChassisTypes = $chassisTypesString
				}
				
				log "Model is `"$($comp.SE_Manufacturer)`" `"$($comp.SE_Model)`". ChassisTypes is `"$($comp.SE_ChassisTypes)`" (`"$($comp.SE_ChassisTypesFriendlyName)`")." -l 3
			}
			else {
				log "Data not retrieved from computer: `"$compName`"!" -l 3
			}
		}
		else {
			log "Computer `"$compName`" did not respond!" -l 3
		}
		log "Done getting Win32_SystemEnclosure data for computer: `"$compName`"..." -l 2 -v 2
		
		$comp
	}
	
	function Export-Comps($comps) {
		log "Exporting data to `"$CSV_PATH`"..."
		$comps = $comps | Sort Name
		#$comps = $comps | Select 
		$comps | Export-Csv -Encoding Ascii -NoTypeInformation -Path $CSV_PATH
		log "Done exporting assignments." -v 2
	}
	
	function Do-Stuff {
		log " " -nots
		
		$compNames = Get-CompNames
		if($compNames) {
			$comps = Get-CompObjects $compNames
			$comps = Get-CompData $comps
			Export-Comps $comps
		}
	}
	
	Do-Stuff
	
	log "EOF"
	log " " -nots

}