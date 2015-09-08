<#
    GetADInventory.ps1
	
    .INFO
    Author: INSIGMA IT Engineering GmbH, Ralf Nelle
    Last Change: 25.02.2015
#>

# abort, if module not found
# Import-Module ActiveDirectory -EA 0
# if (!(Get-Module ActiveDirectory)) {
	# Write-Host "Module 'ActiveDirectory' not found. Script aborted.`n" -fore red -back black
	# break
# }

# ---------------------------------------------------------------------------
#     Settings

$url = "http://intranet"

$site = Get-SPSite $url
$FileName = "{0}\SPGInv_{1}_{2:yyyy-MM-dd}.xml" -f $pwd.path, $site.HostName, [DateTime]::now
	

# ---------------------------------------------------------------------------
#     Functions

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
#   MAIN
	
# prepare xml store
[Xml]$Xml = @'
<?xml version="1.0" encoding="iso-8859-1"?>
<SPFarm>
  <SPGroups />
</SPFarm>
'@

	
# ---------------------------------------------------------------------------	
# 	Get groups

$web  = $site.RootWeb

# query site groups
try {
	Write-Progress -Activity "Getting inventory from domain ''" -status "Gathering user information ..."
	$SiteGroups = $web.SiteGroups | sort name | Select-Object name,owner,users
} catch {
    Write-Warning "$($_.Exception.Message)"
	break
}
  
# add groups to xml store
if ($SiteGroups.Count -gt 0) {  

	Write-Progress -Activity "Getting inventory from domain ''" -status "Storing user information ..."

	# select users node
	$XmlGroups = $Xml.SelectSingleNode("//SPGroups")
	#$XmlGroups.SetAttribute('Count',$SiteGroups.count)

	foreach ($Group in $SiteGroups) {

		$Users = $Group | Select-Object -ExpandProperty Users | Select UserLogin,DisplayName
		
		# add xml element
		$XmlGroup = New-XmlElement $XmlGroups -Name "SPGroup" -Attribute @{Name=$Group.Name;Owner=$Group.Owner} # -Child $Users
		
		if ($Users.Count) {
			$XmlUsers = New-XmlElement $XmlGroup -Name "SPUsers" -Attribute @{Count=$Users.Count} 
			$Users | %{
			   $XmlUser = New-XmlElement $XmlUsers -Name 'SPUser' -Attribute @{UserLogin=$_.UserLogin;DisplayName=$_.DisplayName}
			}
		}
		


	}
	$Xml.Save($FileName)
}

break

$UsedGroups = $site.AllWebs| % {$_.RoleAssignments | Select-Object Member } # ,RoleDefinitionBindings,Parent)
$UG = $UsedGroups | %{ $_.member.Tostring() }
$UG = $UG | Sort-Object | Get-Unique
$UG = $UG |? {$_ -notmatch '\\'}


		#$XmlUser.SetAttribute('Name',$Name)
		#$XmlUser.SetAttribute('OU',$OU)
		#$XmlUser.SetAttribute('Created',$Created)
		#$XmlUser.SetAttribute('LastLogonDate',$LastLogonDate)
		$XmlMemberOfs = Add-XmlNode $XmlUser -Name "MemberOfs"
		
		# add MemberOf property
		if ($MemberOfs) {
			$XmlMemberOfs.SetAttribute('Count',$MemberOfs.Count)
			$MemberOfs | sort | foreach { 
				Add-XmlNode $XmlMemberOfs -Name 'MemberOf' -InnerText $_ | Out-Null
			}
		} else { $XmlMemberOfs.SetAttribute('Count','0') }
		
# ---------------------------------------------------------------------------	
# 	Get groups

# select properties
$ADQueryProperties = 
	@{Name="SamAccountName"; Expression={$_.SamAccountName}},
	@{Name="MemberOf"; Expression={(($_.MemberOf | %{ $_ -replace "(CN=)(.*?),.*",'$2' }) -join ';')}},
	@{Name="OU"; Expression={($_.DistinguishedName | RegEx '(?i)OU=\w{1,}\b') -join ','}},
	@{Name="DistinguishedName"; Expression={$_.DistinguishedName}},
	@{Name="Created"; Expression={$_.Created}}
$ADQueryPropertiesNames = $ADQueryProperties.GetEnumerator() | %{ $_.Name }
$ADQueryFilter = '*'

# get groups
try {
	Write-Progress -Activity "Getting inventory from domain '$NetBIOSName'" -status "Gathering group information ..."
	$ADGroups = Get-ADGroup -SearchBase $SearchBase -Filter $ADQueryFilter -Properties $ADQueryPropertiesNames | `
    Select-Object $ADQueryProperties
} catch {
    Write-Warning "$($_.Exception.Message)"
	break
}
  
# add groups to xml store  
if ($ADGroups.count -gt 0) {  

	Write-Progress -Activity "Getting inventory from domain '$NetBIOSName'" -status "Storing group information ..."

	# select groups node
	$XmlGroups = $Xml.SelectSingleNode("//ADGroups")
	$XmlGroups.SetAttribute('Count',$ADGroups.count)

	foreach ($Group in ($ADGroups | Sort SamAccountName)) {

		# temporary store
		$Name = $Group.SamAccountName
		$OU = $Group.OU
		$MemberOfs = $Group.MemberOf -split ';'
		$Created = $Group.Created

		# removed used properties 
		"SamAccountName","DistinguishedName","OU","MemberOf","Created" | foreach {
			$Group.PSObject.Properties.Remove($_)
		}

		# create xml group node
		$XmlGroup = Add-XmlNode $XmlGroups -Name "ADGroup" -Child $Group
		$XmlGroup.SetAttribute('Name',$Name)
		$XmlGroup.SetAttribute('OU',$OU)
		$XmlGroup.SetAttribute('Created',$Created)
		
		# add MemberOf property
		if ($MemberOfs) {
			$XmlMemberOfs = Add-XmlNode $XmlGroup -Name "MemberOfs"
			$XmlMemberOfs.SetAttribute('Count',$MemberOfs.Count)
			$MemberOfs | sort | foreach { 
				Add-XmlNode -XmlParent $XmlMemberOfs -Name 'MemberOf' -Innertext $_ | Out-Null
			}
		}
	}
	$Xml.Save($FileName)
}
  
break
# ===========================================================================
# 	End of script
