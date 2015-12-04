#
# GetSoftwareInventory.ps1
#
# 	Author: INSIGMA IT Engineering GmbH, Ralf Nelle
# 	Last Change: 09.11.2015
#

param(
	[string]$XmlFile,
	[string[]]$ComputerName,
	[string]$HostsFile,
	[switch]$OnlyAddNew = $true,
	[switch]$CleanOnly, 
	[switch]$ShowCleaningReport,
	[switch]$ShowHosts
)

# ---------------------------------------------------------------------------
#     Fixed Variables

$PatternFile = "$($pwd.path)\CleaningPatterns.txt"
$CleaningReportFile = "$($pwd.path)\CleaningReport.txt" 

# prepare xml store
[Xml]$Xml = @'
<?xml version="1.0" encoding="iso-8859-1"?>
<Inventory>
	<Computers />
</Inventory>
'@



# ---------------------------------------------------------------------------
#     Functions

# removes remarks
Filter Get-NameFromLine { 
  if (($_.trim() -eq "") -or $_.StartsWith('#')) { 
  } else {
	if ($_.contains('#')) { $_.split('#')[0].trim() } else { $_ } 
  }
}

Filter Clean-CustomString {
  if ($_ -ne $null) {
	while ($_ -match "\s\s") { $_ = $_.replace('  ',' ') }
    $_.replace('__',' ').replace('_ ',' ').trim('_').trim()
  } else { $_ }
}

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
				$OS = gwmi -q "SELECT Caption FROM Win32_Operatingsystem" -ComputerName $Computer -ErrorAction Stop
				$CS = gwmi -q "SELECT Domain,Manufacturer,Model,UserName FROM Win32_Computersystem" -ComputerName $Computer -ErrorAction Stop
				$Enc = gwmi -q "SELECT SerialNumber FROM Win32_SystemEnclosure" -ComputerName $Computer -ErrorAction Stop
				$serial = $Enc.SerialNumber
				if ($Enc.SerialNumber.trim() -eq "") {
					$Board = gwmi -q "SELECT SerialNumber FROM Win32_BaseBoard" -ComputerName $Computer -ErrorAction Stop
					$serial = $Board.SerialNumber
				}
				
				$props = [ordered]@{
				  'ComputerName' = $Computer
				  'Domain' = $CS.Domain
				  'Name' = $OS.Caption
				  'Manufacturer' = $CS.Manufacturer
				  'Model' = $CS.Model
				  'SerialNumber' = $Serial
				  'LastUser' = $CS.UserName
				}
				New-Object PSObject -prop $props
				
				Write-Host "done." -fore gray			

			}
			catch { 
				Write-Host
				Write-Host "$($Computer): $($_.Exception.Message)" -fore red -back black 
				#$props = [ordered]@{
				#  'ComputerName' = $Computer
				#  'Error' = $_.Exception.Message
				#}
				#New-Object PSObject -prop $props
			} 		
		}
	}
}

# ---------------------------------------------------------------------------
Function Get-SQLServerVersion {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
        [Parameter(ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, Position=0)]
        [string[]]$ComputerName = $env:COMPUTERNAME
    )

    process {
		$NameSpace = "root\Microsoft\SqlServer\ComputerManagement10"
	
        foreach ($Computer in $ComputerName) {
		
			Write-Host "Query SQL Server version on '$Computer' ... " -nonewline -fore gray		
		
			try { 
				
				$WmiSql = gwmi -Namespace $NameSpace -Class SqlServiceAdvancedProperty -ComputerName $Computer -ErrorAction Stop | 
					?{ "VERSION","SKUNAME" -contains $_.PropertyName } | `
					Select ServiceName,PropertyName,PropertyStrValue | group ServiceName -AsHashTable -AsString
				
				foreach ($item in $WmiSql.Keys) { 
					New-Object -TypeName PSCustomObject -Property @{ 
					  'ComputerName' = $Computer;
					  'SQLVersion' = ($WmiSql[$item] | %{  $_.PropertyStrValue }) -join " "
					}				
				}
				
				Write-Host "done." -fore gray	
			}
			catch { 
				$ErrorLogMsg = "$($Computer): $($_.Exception.Message)"
				Write-Host
				Write-Host $ErrorLogMsg -fore red -back black 
				$ErrorLogMsg >> ".\Errors.log"			
			}		
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
	catch { 
		$ErrorLogMsg = "$($Computer): $($_.Exception.Message)"
		Write-Host $ErrorLogMsg -fore red -back black 
		$ErrorLogMsg >> ".\Errors.log"
	}
	
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

		# clean model and serial 
		foreach ($Computer in $Xml.SelectNodes("//Computer")) {
	
			while ($Computer.Model -match "\s\s") { $Computer.Model = $Computer.Model.replace('  ',' ') }
			$Computer.Model = $Computer.Model.replace('__',' ').replace('_ ',' ').trim()
			$Computer.SerialNumber = $Computer.SerialNumber.trim()
		}
		$Xml.Save($XmlCleanFile)

		Write-Host "Inventory result optimization done." -fore gray

	} else {
		Write-Host "Missing pattern file. Inventory result optimization skipped." -fore gray 
	}
}

# ---------------------------------------------------------------------------
Function Show-CleanedProducts {
	Compare-Object (gc $XmlFile) (gc $XmlCleanFile) | %{ 
		$_.InputObject.replace("<Software>","").replace("</Software>","").trim() 
	} | sort | Get-Unique | sc $CleaningReportFile
	notepad $CleaningReportFile
}

# ---------------------------------------------------------------------------
Function Show-Hosts {

	Write-Host "List of gathered systems`n" -ForegroundColor Green

	$Props = 
		@{Name="Name"; Expression={$_.Name}},
		@{Name="OS";Expression={ $_.Software[0] }},
		@{Name="Description"; Expression={$_.Description}},
		@{Name="IP"; Expression={$_.IP}},
		@{Name="Model"; Expression={$_.Model}},
		@{Name="LastUser"; Expression={$_.LastUser}}

	$nodes = $Xml.SelectNodes("//Computer") | sort Name | Select-Object $Props
	$nodes | Export-Csv Computer.csv -NoTypeInformation -Delimiter ";" -NoClobber -Encoding UTF8
		
}



# ---------------------------------------------------------------------------
Function Get-InventoryData {

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
		$XmlComputer = New-XmlElement $XmlComputers -Name "Computer" -Attribute @{ 'Name' = $Result.ComputerName }
		
		if ($Result.Error) {
			$XmlComputer.SetAttribute('AccessError',$Result.Error)
		} else {
			$XmlComputer.SetAttribute('Domain',$Result.Domain)
			$XmlComputer.SetAttribute('Model',$Result.Manufacturer + ' ' + $Result.Model)	
			$XmlComputer.SetAttribute('SerialNumber',$Result.SerialNumber)
			$XmlComputer.SetAttribute('LastUser',$Result.LastUser)			
				
			New-XmlElement $XmlComputer -Name 'Software' -InnerText $Result.Name | Out-Null

			$AvailableComputers += $Result.ComputerName
		}
			
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

# get exiting xml file
if (($XmlFile) -and (Test-Path $XmlFile)) {
	$XmlFile = (Resolve-Path $XmlFile).path
	[xml]$Xml = gc $XmlFile
} else { $XmlFile = ("{0}\SwInv_{1}.xml" -f $pwd.path,$ENV:USERDOMAIN.replace('.','_')) }
$XmlCleanFile = $XmlFile.replace(".xml","_Cleaned.xml")
Write-Host "Using '$XmlFile' ..." -fore cyan

# special ops
if ($ShowCleaningReport) { Show-CleanedProducts; break; }
if ($ShowHosts) { Show-Hosts; break; }
if ($CleanOnly) { Clean-InventoryData; break; }

# get hosts
[string[]]$Computers = $env:COMPUTERNAME
if ($ComputerName) { 
	[string[]]$Computers = $ComputerName 
} elseif ($HostsFile) {
	if (Test-Path $HostsFile) {
		$HostsFile = (Resolve-Path $HostsFile).path
		[string[]]$Computers = gc $HostsFile | Get-NameFromLine
	} else {
		Write-Host "'$HostsFile' not found. Using localhost." -fore yellow
	}
}


# filter existing computer
if ($OnlyAddNew) {
	$ExistingHosts = $Xml.SelectNodes("//Computer") | sort Name | %{ $_.Name }
	if ($ExistingHosts) {
		$Computers = diff $Computers $ExistingHosts | ?{ $_.sideindicator -eq "<=" } | %{ $_.InputObject }
	}
}


Get-InventoryData
Clean-InventoryData

Write-Host "All done."

break



# ===========================================================================
#   Additional stuff (not script relevant)

$hosts = gc .\hosts_INSIGMAWEB_DMZ.txt | ?{ -not($_.StartsWith('#')) } 
$hosts = $hosts | ?{ $_.IndexOf('#') -gt 0 }

$systems = &{ 
  $hosts | %{ 
	$item = $_.Split('#') 
	New-Object -TypeName PSCustomObject -Property @{ 
		'ComputerName' = $item[0].trim();
		'Description' = $item[1].trim();
	}
  }
}

$systems = $systems | ?{ $_.Description }


foreach ($XmlComputer in $Xml.SelectNodes("//Computer")) {
	$Description = ""
	$system = $systems | ?{ $_.ComputerName -eq $XmlComputer.Name }
	if ($system) { $Description = $system.Description }
	$XmlComputer.SetAttribute("Description",$Description)
}
	

# generate computer short list
[xml]$Xml = gc $XmlFile

foreach ($Software in $Xml.SelectNodes("//Software")) {
	[void]$Software.ParentNode.RemoveChild($Software)
}
$Xml.Save("C:\Temp\SWInv\Computers.xml")

foreach ($node in $Xml.SelectNodes("//Computer")) {
	$node.Model = $node.Model | Clean-CustomString
	$node.SerialNumber = $node.SerialNumber | Clean-CustomString
}

