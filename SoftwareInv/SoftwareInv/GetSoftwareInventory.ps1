Param(
	[alias("CN")]
    [String[]]$ComputerName = $ENV:ComputerName
)

# ---------------------------------------------------------------------------
#	Settings

$InvFile = "$($pwd.path)\INSinv.csv"
$machines = $ComputerName


# ---------------------------------------------------------------------------
# 	MAIN

# prepare data file
if (!(Test-Path $InvFile)) { '"Computer";"Software"' > $InvFile }

foreach ($machine in $ComputerName) {
	
	Write-Host "Quering $($machine) ..."
	
	# get OS
	$query = "Select Caption from Win32_OperatingSystem"
	if ($machine -eq $Env:ComputerName) {
		$wmi = gwmi -Query $query -EA 0
	} else {
		$wmi = gwmi -Query $query -Computer $machine -EA 0 # -Credential $cred 
	}
	
	if ($wmi) { 
	
		$os = $wmi.Caption.trim()
		("'$($machine)';'$($os)'").replace("'",'"') >> $InvFile
	
		# get software
		$query = "Select Caption from Win32_Product"
		if ($machine -eq $Env:ComputerName) {
			$wmi = gwmi -Query $query -EA 0
		} else {
			$wmi = gwmi -Query $query -Computer $machine -EA 0 # -Credential $cred 
		}	
	
		if ($wmi) {
			$wmi | ?{ $_.Caption } | %{ $_.Caption.trim() } | sort | Get-Unique | %{
				("'$($machine)';'$($_)'").replace("'",'"') >> $InvFile
			}
		}	
	}
}

Write-Host "All done."



break

#*** end of script ***#

# ---------------------------------------------------------------------------
#	Additional used functions
	
# get computer from dns query
$systems = @()
$ipbase = "192.168.13."
2..254 | %{
	$dns = nslookup "$ipbase$($_)"
	if ($dns.length -gt 5) {
		$systems += $dns[3].split(':')[1].trim()
	}
}

