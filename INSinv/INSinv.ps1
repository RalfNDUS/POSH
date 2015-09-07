#
# INSinv.ps1
#
#	Author: Ralf Nelle
#	Last Change: 21.08.2015
#

Param(
    [String[]]$ComputerName = 'localhost',
    [String]$NetworkName = $ENV:USERDOMAIN,	
    [alias("Cred")]
    $Credential,
	[Switch]$All = $false
)

# ---------------------------------------------------------------------------
#   Settings

$xmlFile = "$($pwd.path)\INSinvDB.xml"

$xmlTemplate = @'
<?xml version="1.0" encoding="iso-8859-1"?>
<Networks>
<Network Name="{0}" />
</Networks>
'@


# ---------------------------------------------------------------------------
#	Property definitions
# 	see also http://msdn.microsoft.com/en-us/library/aa394084(VS.85).aspx

$CompCol = New-Object PsObject -Property @{
	'Query' = "Select * from Win32_ComputerSystem"
	'Props' = 
		@{Name="HostName"; Expression={$_.Name}},
		@{Name="HostDomain"; Expression={$_.Domain}},
		@{Name="Manufacturer"; Expression={$_.Manufacturer.trim().trim('_')}},
		@{Name="Model"; Expression={$_.Model.trim().trim('_')}},
		@{Name="Memory"; Expression={$_.TotalPhysicalMemory/1MB -as [int64]}}
}

$OsCol = New-Object PsObject -Property @{
	'Query' = "Select * from Win32_OperatingSystem"
	'Props' = 
		@{Name="OsName"; Expression={$_.Caption}},
		@{Name="OsServicePack"; Expression={$_.CSDVersion}},
		@{Name="OsInstalldate"; Expression={($_.ConvertToDateTime($_.InstallDate))}},
		@{Name="OsLastBootUpTime"; Expression={($_.ConvertToDateTime($_.LastBootUpTime))}},
		@{Name="Osx64"; Expression={$_.OSArchitecture -like "64-Bit"}}	
}

$PropCol = @()


$PropCol += New-Object PsObject -Property @{
	'Query' = "Select * from Win32_Processor"
	'Props' = 
		@{Name="CpuName"; Expression={(TrimSpace $_.Name)}},
		@{Name="CpuCaption"; Expression={$_.Caption}},
		@{Name="CpuCores"; Expression={$_.NumberOfCores}},
		@{Name="CpuLogicalProcessors"; Expression={$_.NumberOfLogicalProcessors}},
		@{Name="CpuMaxClockSpeed"; Expression={$_.MaxClockSpeed}}
	'Multiple' = $true		
}

$PropCol += New-Object PsObject -Property @{
	'Query' = "Select * from Win32_PhysicalMemory"
	'Props' = 
		@{Name="RamTag"; Expression={$_.Tag}},
		@{Name="RamCapacity"; Expression={[int64]$_.Capacity/1MB}},
		@{Name="RamSpeed"; Expression={$_.Speed}}
	'Multiple' = $true		

}

$PropCol += New-Object PsObject -Property @{
	'Query' = "Select * from Win32_LogicalDisk Where DriveType='3'"
	'Props' = 
		@{Name="LDiskName"; Expression={$_.Name}},
		@{Name="LDiskVolumeName"; Expression={$_.VolumeName}},
		@{Name="LDiskSize"; Expression={$_.Size/1MB -as [Int64]}},
		@{Name="LDiskFree"; Expression={$_.FreeSpace/1MB -as [Int64]}},
		@{Name="LDiskUsage"; Expression={(1-($_.FreeSpace/$_.Size))*100 -as [int]}}		
	'Multiple' = $true
}

$PropCol += New-Object PsObject -Property @{
	'Query' = 
		"Select * from Win32_NetworkAdapterConfiguration Where IPEnabled='True'"
	'Props' = 
		@{Name="NicName"; Expression={$_.Description}},
		@{Name="NicMac"; Expression={$_.MACAddress}},
		@{Name="NicServiceName"; Expression={$_.ServiceName}},
		@{Name="NicIpAddress"; Expression={($_.IPAddress -like '*.*.*.*') -join ' '}},
		@{Name="NicIpSubnet"; Expression={($_.IPSubnet -like '*.*.*.*') -join ' '}},
		@{Name="NicIpGateway"; Expression={$_.DefaultIPGateway}},
		@{Name="NicDhcpEnabled"; Expression={$_.DHCPEnabled}},
		@{Name="NicDnsServer"; Expression={($_.DNSServerSearchOrder -like '*.*.*.*') -join ' '}},
		@{Name="NicDnsDomain"; Expression={$_.DNSDomain}},
		@{Name="NicDnsDomainSuffixSearchOrder"; Expression={$_.DNSDomainSuffixSearchOrder}}
	'Multiple' = $true
}

if ($All) {
	$PropCol += New-Object PsObject -Property @{
		'Query' = 
			"Select * from Win32_Share"
		'Props' = 
			@{Name="ShareName"; Expression={$_.Name}},
			@{Name="SharePath"; Expression={$_.Path}},
			@{Name="ShareDescription"; Expression={$_.Description}}
		'Multiple' = $true
	}


	$PropCol += New-Object PsObject -Property @{
		'Query' = 
			"Select Name,ID,ParentID from Win32_ServerFeature"
		'Props' = 
			@{Name="ServerFeatureName"; Expression={$_.Name}},
			@{Name="ServerFeatureID"; Expression={$_.ID}},
			@{Name="ServerFeatureParentID"; Expression={$_.ParentID}}
		'Multiple' = $true
	}

	# $PropCol += New-Object PsObject -Property @{
		# 'Query' = 
			# "Select * from Win32_Product"
		# 'Props' = 
			# @{Name="ProductName"; Expression={$_.Name}},
			# @{Name="ProductVersion"; Expression={$_.Version}},
			# @{Name="ProductVendor"; Expression={$_.Vendor}}
		# 'Multiple' = $true
	# }
}


# ---------------------------------------------------------------------------
#     Functions

#-- TrimSpace
	function TrimSpace($str) { 
		while ($str -match " {2,}") { $str = $str.replace('  ',' ') }; $str
	}

#-- ConvertHereString
	function ConvertHereString($str) { 
		$str -split (',') | %{ $_.trim() } | sort
	}

#-- Start-WmiPing
	function Start-WmiPing($ComputerName) {
		
		write-progress -Activity "Avaibility check for $($ComputerName)" -status "Waiting"
		
		try { 
			$result = gwmi -query "Select * from Win32_PingStatus Where Address='$($ComputerName)'"
			if ($result.statuscode -eq 0) { $success = $true }	else { $success = $false }
		}
		catch { $success = $false; $result = $_ }

		@{ 'Success' = $success; 'Result' = $result }
	}

#-- Start-WmiQuery
	Function Start-WmiQuery($ComputerName, $Query, $Credential, $SelectObj, $StatusMsg) {
		
		If ($StatusMsg) {
			Write-Progress -Activity "Starting WMI query on $($ComputerName)" -status $StatusMsg
			$PSBoundParameters.remove('StatusMsg') | Out-Null
		}
		$PSBoundParameters.remove('SelectObj') | Out-Null
		
		try {
			$result = gwmi @PSBoundParameters -ea 0
			if ($result) { $success = $true; $result = $result | select-object $SelectObj } 
				else { $success = $false; $result = "No response on WMI query."	}
		}
		catch { $success = $false; $result = $_ }
		
		@{ 'Success' = $success; 'Result' = $result }
	}

#    Add-XmlInnerText   -----------------------------------------------------
Function Add-XmlInnerText($ParentNode,$PropObj) {
    Foreach($Property in $PropObj) { 
        $Property.psobject.properties | %{ 
            $PropNode = $Xml.CreateElement($_.Name)
            [void]$ParentNode.AppendChild($PropNode)
            $PropNode.Set_Innertext($_.Value)
        }            
    }    
}

#   Add-XmlPropNode   ------------------------------------------------------
Function Add-XmlPropNode($Name, $WMIdata, $multiple) {
    If ($multiple) {
		$PropNode = $xml.CreateElement($Name + "s"); $i=0
        Foreach ($WMIrecord in $WMIdata) {
            $PropSubNode = $xml.CreateElement($Name)
            $PropSubNode.SetAttribute("Id", $i++)    
            [void]$PropNode.AppendChild($PropSubNode)
            Add-XmlInnerText $PropSubNode $WMIrecord
        }
	} else {
        $PropNode = $xml.CreateElement($Name)
        Add-XmlInnerText $PropNode $WMIdata
    }
	[void]$ComputerNode.AppendChild($PropNode)
}

#	New-XmlElement   ------------------------------------------------------- 	
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

	

# ===========================================================================    
#   Main Script

#-- Prepare xml file
If (!(Test-Path $xmlFile)) { $xmlTemplate -f $NetworkName | Out-File $XmlFile }
[xml]$Xml = Get-Content $xmlFile
$XmlRoot = $Xml.Networks

$NetworkNode = $Xml.SelectSingleNode("//Network[@Name='$($NetworkName)']")
If (-not($NetworkNode)) {
	$NetworkNode = $xml.CreateElement("Network")
	[void]$xmlRoot.AppendChild($NetworkNode)
	$NetworkNode.SetAttribute("Name", $NetworkName)
}

$ComputersNode = $xml.SelectSingleNode("//Network[@Name='$($NetworkName)']/Computers")
If (-not($ComputersNode)) {
	$ComputersNode = $xml.CreateElement("Computers")
	[void]$NetworkNode.AppendChild($ComputersNode)
}

$ServersNode = $xml.SelectSingleNode("//Network[@Name='$($NetworkName)']/Servers")
If (-not($ServersNode)) {
	$ServersNode = $xml.CreateElement("Servers")
	[void]$NetworkNode.AppendChild($ServersNode)
}

#-- Gather WMI data from all systems	
$HostNames = $ComputerName
Foreach ($HostName in $HostNames) {

	$HostArgs = @{ 'ComputerName' = $HostName }
	If ($HostName -notlike 'localhost') { 
		If ($Credential) { $HostArgs += @{'Credential' = $Credential} }
	}

	# check availibility
	if (-not((Start-WmiPing @HostArgs).success)) { 
		write-host "$($HostArgs.ComputerName): Zielhost nicht erreichbar." -fore gray
	}

	# try wmi access
	else {

		$HostArgs.Query = $CompCol.Query
		$HostArgs.SelectObj = $CompCol.Props
		$HostArgs.StatusMsg = "Processing Win32_ComputerSystem"
		$Wmi = Start-WmiQuery @HostArgs

		If ($Wmi.Success -eq $true) { 
		
			$ComputerSystem = $Wmi.Result
		
			$HostArgs.Query = $OsCol.Query
			$HostArgs.SelectObj = $OsCol.Props
			$HostArgs.StatusMsg = "Processing Win32_OperatingSystem"
			$Wmi = Start-WmiQuery @HostArgs
			$OperatingSystem = $Wmi.Result
		
			if ($OperatingSystem.OsName -like "*Server*") { 
				$ComputerNode = $NetworkNode.SelectSingleNode("//Server[@Name='$($ComputerSystem.HostName)']")
				if (-not($ComputerNode)) { 
					$ComputerNode = $Xml.CreateElement("Server")
					[void]$ServersNode.AppendChild($ComputerNode)
				}
			} else { 
				$ComputerNode = $NetworkNode.SelectSingleNode("//Computer[@Name='$($ComputerSystem.HostName)']")
				if (-not($ComputerNode)) { 
					$ComputerNode = $Xml.CreateElement("Computer")
					[void]$ComputersNode.AppendChild($ComputerNode)
				}
			}
			
			# remove old properties and add computer and os properties
			$ComputerNode.RemoveAll()
			$ComputerNode.SetAttribute("Name", $ComputerSystem.HostName)        
			$isVM = "VMware Virtual Platform","Virtual Machine" -contains $cs.Model
			$ComputerNode.SetAttribute("VirtualMaschine", $isVM)        
			$ComputerNode.SetAttribute("TimeStamp",(Get-Date).ToString("dd.MM.yyyy HH:mm"))
			
			Add-XmlPropNode "ComputerSystem" $ComputerSystem $false
			Add-XmlPropNode "OperatingSystem" $OperatingSystem $false
			
			foreach ($PropSet in $PropCol) { 
						
				$HostArgs.Query = $PropSet.Query
				$HostArgs.SelectObj = $PropSet.Props
				if ($PropSet.Query -match "Win32_\w+") { 
					$HostArgs.StatusMsg = $class = $matches[0] 
				} 
				$Wmi = Start-WmiQuery @HostArgs

				If ($Wmi.Success) { 
					Add-XmlPropNode $class.replace('Win32_','') $Wmi.Result $PropSet.Multiple
				}
			}
			
			# new code
			$HostArgs
			$XmlProducts = New-XmlElement $ComputerNode -Name "Products" -Attribute @{ 'Name' = $ComputerName }
			# $XmlWebs = New-XmlElement $XmlSite -Name "SPWebs"
			
			$xml.save($xmlFile); Write-host "$($HostName): Inventory done." 
		}
	}
}

# End of script





