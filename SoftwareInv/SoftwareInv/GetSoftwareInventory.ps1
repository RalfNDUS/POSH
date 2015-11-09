#
# GetSoftwareInventory.ps1
#
# 	Author: INSIGMA IT Engineering GmbH, Ralf Nelle
# 	Last Change: 09.11.2015
#

param([string[]]$ComputerName,[switch]$CleanOnly, [switch]$ShowCleanedProducts)

# ---------------------------------------------------------------------------
#     Settings

$ComputersFile = $pwd.path + "\hosts.txt"
$XmlFile = $pwd.path + "\SoftwareInv.xml"
$XmlCleanFile = $pwd.path + "\SoftwareInv_cleaned.xml"
$PatternFile = $pwd.path + "\CleanPatterns.txt"
$CleaningReportFile = $pwd.path + "\CleaningReport.txt"

# prepare xml store
[Xml]$Xml = @'
<?xml version="1.0" encoding="iso-8859-1"?>
<Inventory>
  <Computers />
</Inventory>
'@


# ---------------------------------------------------------------------------
#     Functions

# ---------------------------------------------------------------------------
Function New-XmlElement($XmlParent, $Name, $Child, $InnerText, $Attribute ) {

  $XmlElement = $Xml.CreateElement($Name)
  [void]$XmlParent.AppendChild($XmlElement)

  # add inner text, attribute or child element
  if ($Innertext) { $XmlElement.Set_Innertext($Innertext) }			
  if ($Attribute) { 
	foreach ($key in $Attribute.Keys) {
	  $XmlElement.SetAttribute($key, $Attribute[$key])
	}
  }
  if ($Child) {
	$Child.psobject.properties | Foreach { 
	  New-XmlElement $XmlElement $_.Name -Innertext $_.Value | Out-Null
    }   		
  }	
	
  return $XmlElement
}

# ---------------------------------------------------------------------------
Function Get-Operatingsystem {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
        [Parameter(ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, Position=0)]
        [string[]]$ComputerName = $env:COMPUTERNAME
    )

    process {
        foreach ($Computer in $ComputerName) {
		
			Write-Host "Query operatingsystem on '$Computer' ... " -nonewline -fore gray		
		
			try { 
				$Result = gwmi -q "SELECT Caption FROM Win32_Operatingsystem" -ComputerName $Computer -ErrorAction Stop
				New-Object -TypeName PSCustomObject -Property @{
				  'ComputerName' = $Computer; 'Name' = $result.Caption
				}
			}
			catch { Write-Host "$($Computer): $($_.Exception.Message)" -fore red -back black } 		

			Write-Host "done." -fore gray			
		}
	}
}

# ---------------------------------------------------------------------------
Function Get-RemoteSoftware ($ComputerName = $ENV:COMPUTERNAME) {

  # Registry settings
  $HKLM = [UInt32] "0x80000002"
  $UNINSTALL_KEYS = `
	"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", `
	"SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"

  # Iterate the computer name(s).
  foreach ($Computer in $ComputerName) {

    Write-Host "Query software on '$Computer' ... " -nonewline -fore gray

	try {
		# Connect to the StdRegProv class on the computer.
		$regProv = [WMIClass] "\\$Computer\root\default:StdRegProv"

		# Enumerate the Uninstall subkey.
		&{ 
			foreach ($UNINSTALL_KEY in $UNINSTALL_KEYS) {

			  $subkeys = $regProv.EnumKey($HKLM, $UNINSTALL_KEY).sNames
			  foreach ($subkey in $subkeys) {
			  
				# Get the application's display name.
				$name = $regProv.GetStringValue($HKLM,(join-path $UNINSTALL_KEY $subkey), "DisplayName").sValue
				if ($name -ne $NULL) {
				
					$version = $regProv.GetStringValue($HKLM,(join-path $UNINSTALL_KEY $subkey), "DisplayVersion").sValue;
					if ($version) { $name = ("{0} [{1}]" -f $name.trim(),$version.trim()) }

					New-Object -TypeName PSCustomObject -Property @{
					  'ComputerName' = $Computer;
					  'AppName' = $name.trim();
					  'Vendor' = $regProv.GetStringValue($HKLM,(join-path $UNINSTALL_KEY $subkey), "Publisher").sValue;
					}
				}
			  }
			}
		} | select -Unique ComputerName,AppName,Vendor
	}
	catch { Write-Host "$($Computer): $($_.Exception.Message)" -fore red -back black }
	
	Write-Host "done." -fore gray
  }
}

# ---------------------------------------------------------------------------
Function Clean-InventoryData {

	if (Test-Path $PatternFile) {

		# build cleaning rules
		$pattern = "(" + ([System.IO.File]::ReadAllLines($PatternFile) -join '|') + ")"	

		# clean inventory file
		[xml]$Xml = gc $XmlFile
		foreach ($Software in $Xml.SelectNodes("//Software")) {

			# check by pattern
			if (!($Software.InnerText) -or ($Software.InnerText -match $pattern)) {
				[void]$Software.ParentNode.RemoveChild($Software)
			}
		}
		$Xml.Save($XmlCleanFile)
		Write-Host "Inventory result optimization done." -fore gray

	} else {
		Write-Host "Missing pattern file. Inventory result optimization skipped." -fore gray 
	}
}

# ---------------------------------------------------------------------------
Function Show-CleanedProducts {
	Compare-Object (gc .\SoftwareInv.xml) (gc .\SoftwareInv_cleaned.xml) | %{ 
		$_.InputObject.replace("<Software>","").replace("</Software>","").trim() 
	} | sort | Get-Unique | sc $CleaningReportFile
	notepad $CleaningReportFile
}

# ---------------------------------------------------------------------------
Function Get-InventoryData {

	# get hosts
	if ($ComputerName) { 
		[string[]]$Computers = $ComputerName 
	} else { 
		if (!(Test-Path $ComputersFile)) {
			Write-Host "'$ComputersFile' not found. Using localhost." -fore yellow
			[string[]]$Computers = $env:COMPUTERNAME
		} else {
			[string[]]$Computers = gc $ComputersFile | ?{ $_ -notlike "#*" }
		}
	}

	# get exiting xml file
	if (Test-Path $XmlFile) { [xml]$Xml = gc $XmlFile }

	# get root note
	$XmlComputers = $Xml.SelectSingleNode("//Computers")
	$AvailableComputers = @()

	# get operating system
	Write-Host "Getting host OS ..." -fore gray
	foreach ($Result in (Get-Operatingsystem $Computers | Sort ComputerName)) {

		# remove existing computer node
		$XmlComputer = $Xml.SelectSingleNode("//Computer[@Name='$($Result.ComputerName)']")
		if ($XmlComputer) { [Void]$XmlComputer.ParentNode.RemoveChild($XmlComputer) }

		# add computer node
		$AvailableComputers += $Result.ComputerName
		$XmlComputer = New-XmlElement $XmlComputers -Name "Computer" -Attribute @{ 'Name' = $Result.ComputerName }
		New-XmlElement $XmlComputer -Name 'Software' -InnerText $Result.Name | Out-Null

		# add timestamp
		$XmlComputer.SetAttribute("TimeStamp",(Get-Date).ToString("dd.MM.yyyy HH:mm:ss"))

	}
	$Xml.Save($XmlFile)


	# get details from active hosts
	$Results = Get-RemoteSoftware $AvailableComputers
	$InstalledSoftware = $Results | Group ComputerName -AsHashTable -AsString

	foreach ($Computer in $InstalledSoftware.Keys) {

		Write-Host "Examing '$Computer' software ..." -fore gray
		$XmlComputer = $Xml.SelectSingleNode("//Computer[@Name='$($Computer)']")
		foreach ($Software in ($InstalledSoftware[$Computer] | %{ $_.AppName } | sort | get-unique )) {
			New-XmlElement $XmlComputer -Name 'Software' -InnerText $Software | Out-Null
		}
	}
	$Xml.Save($XmlFile)
}


# ===========================================================================
#   MAIN

if ($ShowCleanedProducts) { Show-CleanedProducts; break; }

if (!($CleanOnly)) { Get-InventoryData }
	
Clean-InventoryData

Write-Host "All done."

