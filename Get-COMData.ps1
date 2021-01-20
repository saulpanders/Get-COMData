function Get-COMData {
<#
    .SYNOPSIS
        Enumerates COM objects on host, currently drops 2 artifacts (mandatory)in the path where script is run
    .DESCRIPTION
		TBD - research thanks to hamlton https://www.fireeye.com/blog/threat-research/2019/06/hunting-com-objects.html
		
		TODO: 
				improve args/logic to remove temporary files (if desired)
				add automatic keyword search w/ word bank for interesting objects (execute, exec, spawn, launch, and run)
				
#>
[CmdletBinding()]
[ValidateNotNullorEmpty()]
param(
	[Parameter(Mandatory=$false)]
	[switch]$ScriptUpdates,
	
	[Parameter(Mandatory=$true)]
    [string]$CLSIDFileName,
	
	[Parameter(Mandatory=$true)]
    [string]$OutputFileName
)
	New-PSDrive -PSProvider registry -Root HKEY_CLASSES_ROOT -Name HKCR | out-null
	if ($ScriptUpdates) {
		Write-Host "[+] Enumerating CLSIDs ..."
	}
	$clsids = (Get-ChildItem -Path HKCR:\CLSID -Name | Select -Skip 1)
	if ($CLSIDFileName) {
		Out-File -FilePath "$CLSIDFileName" -InputObject $clsids
		Write-Host "[*] Wrote CLSIDs to $CLSIDFileName"
	}
	
	if ($ScriptUpdates) {
		Write-Host "[+] Instantiating COM objects ..."
	}
	
	$Position = 1
	
	# This situation can be improved
	ForEach($CLSID in Get-Content $CLSIDFileName) {
		if ($ScriptUpdates) {
			Write-Output "$($Position) - $($CLSID)"
		}
		Write-Output "$($Position) - $($CLSID)" | Out-File $OutputFileName -Append
		Write-Output "------------------------" | Out-File $OutputFileName -Append
		Write-Output $($CLSID) | Out-File $OutputFileName -Append
		
		#add try-catch here
		try {
			$handle = [activator]::CreateInstance([type]::GetTypeFromCLSID($CLSID))	
			$handle | Get-Member | Out-File $OutputFileName -Append
		}
		catch {
			if ($ScriptUpdates) {
				Write-Output "Error accessing COM Object: $($CLSID)"
			}
			Write-Output "Error accessing COM Object: $($CLSID) `n" | Out-File $OutputFileName -Append
			continue
		}
		finally {
			$Position +=1
		}
	
	}
}