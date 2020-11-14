# Summary
Takes an AD OU. Queries all computers in that OU for their make, model, and other chassis-related info, and outputs the results in a CSV.

# Instructions
1. Download `Get-FormFactors.psm1`
2. Import it as a module: `Import-Module "c:\path\to\script\Get-FormFactors.psm1"`
3. Run it: `Get-FormFactors -OUDN "OU=YourOU,OU=Desktops,OU=Engineering,OU=Urbana,DC=ad,DC=uillinois,DC=edu"`

# Parameters

### -OUDN
Required string.  
The distinguished name of the OU containing the computers you want to query.  

### -LogPath
Optional string.  
The full path to the logfile which will be generated.  
Default is `c:\engrit\logs\Get-FormFactors_yyyy-MM-dd_HH-mm-ss-ffff.log`.  
The output CSV file will have the same filename, but with a `.csv` extension.  

### -CIMTimeoutSec
Optional integer.  
The number of seconds to wait before timing out CIM calls when querying computers.  
Default is `15`.  

### -DisableCIMFallbacks
Optional switch.  
By default, if a CIM query fails, the script will fall back on alternate query methods utilizing WMI and/or `Invoke-Command`. These methods do not have an intelligent timeout mechanism, so they can cause the script to hang indefinitely in some cases.  
When this switch is specified, and the initial CIM query fails (after the timeout specified by `CIMTimeoutSec`), the script will skip the fallback methods and move onto the next computer.  

# Notes
- By mseng3
