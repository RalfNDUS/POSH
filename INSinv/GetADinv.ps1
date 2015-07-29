#
# GetADInv.ps1
#
# Author: INSIGMA IT Engineering GmbH, Ralf Nelle
# Last Change: 29.07.2015
#

Param(
    [String]$SearchBase
)


# ---------------------------------------------------------------------------
#     Settings

$NetBIOSName = (Get-ADDomain).NetBIOSName
$FileName = "{0}\ADInv_{1}_{2:yyyy-MM-dd}.xml" -f $pwd.path, $NetBIOSName, [DateTime]::now

If (!$SearchBase) {   
	$SearchBase = (Get-ADDomain).DistinguishedName  
}
	

# ---------------------------------------------------------------------------
#     Functions

Function Split-DistinguishedName ($dn) {
  $cn = ($dn -split "OU=")[0] 
  $dn = $dn.replace($cn,'')
  $ou = ($dn -split "DC=")[0] 
  $dn = $dn.replace($ou,'')

  New-Object PSObject -Property @{
	'Name' = $cn.replace("CN=",'').trim(',')
	'OU' = $ou.trim(',')
	'DC' = $dn
  }
}

Filter RegEx($Pattern) { ([RegEx]::Matches($_,$Pattern)) }

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

# abort, if module not found
Import-Module ActiveDirectory -EA 0
if (!(Get-Module ActiveDirectory)) {
	Write-Host "Module 'ActiveDirectory' not found. Script aborted.`n" -fore red -back black
	break
}
	
# prepare xml store
[Xml]$Xml = @'
<?xml version="1.0" encoding="iso-8859-1"?>
<ADDomain Name="{0}">
  <ADUsers />
  <ADGroups />
</ADDomain>
'@ -f $NetBIOSName

	
# ---------------------------------------------------------------------------	
# 	Get active users

# select properties
$ADQueryProperties = 
	@{Name="SamAccountName"; Expression={$_.SamAccountName}},
	@{Name="SurName"; Expression={$_.SurName}},
	@{Name="GivenName"; Expression={$_.GivenName}},
	@{Name="DisplayName"; Expression={$_.DisplayName}},
	@{Name="UserPrincipalName"; Expression={$_.UserPrincipalName}},
	@{Name="EmailAddress"; Expression={$_.EmailAddress}},
	@{Name="MemberOf"; Expression={(($_.MemberOf | %{ $_ -replace "(CN=)(.*?),.*",'$2' }) -join ';')}},
	@{Name="DistinguishedName"; Expression={$_.DistinguishedName}},
	@{Name="OU"; Expression={($_.DistinguishedName | RegEx '(?i)OU=\w{1,}\b') -join ','}},
	@{Name="LastLogonDate"; Expression={$_.LastLogonDate}},
	@{Name="Created"; Expression={$_.Created}}
$ADQueryPropertiesNames = $ADQueryProperties.GetEnumerator() | %{ $_.Name }
$ADQueryFilter = 'Enabled -eq $true'

# run ad query
try {
	Write-Progress -Activity "Getting inventory from domain '$NetBIOSName'" -status "Gathering user information ..."
	$ADUsers = Get-ADUser -SearchBase $SearchBase -Filter $ADQueryFilter -Properties $ADQueryPropertiesNames | `
    Select-Object $ADQueryProperties
} catch {
    Write-Warning "$($_.Exception.Message)"
	break
}
  
# add users to xml store
if ($ADUsers.count -gt 0) {  

	Write-Progress -Activity "Getting inventory from domain '$NetBIOSName'" -status "Storing user information ..."

	# select users node
	$XmlUsers = $Xml.SelectSingleNode("//ADUsers")
	$XmlUsers.SetAttribute('Count',$ADUsers.count)

	foreach ($User in ($ADUsers | Sort SamAccountName)) {

		# user properties
		$UserProp = @{}
		$UserProp.Name = $User.SamAccountName
		$UserProp.OU = $User.OU
		$UserProp.Created = $User.Created
		$UserProp.LastLogonDate = $User.LastLogonDate
		$MemberOfs = $User.MemberOf -split ';'

		# removed used properties 
		"SamAccountName","DistinguishedName","OU","MemberOf","Created","LastLogonDate" | foreach {
			$User.PSObject.Properties.Remove($_)
		}

		# add xml user element
		$XmlUser = New-XmlElement $XmlUsers -Name "ADUser" -Attribute $UserProp -Child $User
		$XmlMemberOfs = New-XmlElement $XmlUser -Name "MemberOfs"
		
		# add MemberOf property
		if ($MemberOfs) {
			$XmlMemberOfs.SetAttribute('Count',$MemberOfs.Count)
			$MemberOfs | sort | foreach {
				New-XmlElement $XmlMemberOfs -Name 'MemberOf' -InnerText $_ | Out-Null
			}
		} else { $XmlMemberOfs.SetAttribute('Count','0') }
	}
	$Xml.Save($FileName)
}


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
		$GroupProp = @{}
		$GroupProp.Name = $Group.SamAccountName
		$GroupProp.OU = $Group.OU
		$GroupProp.Created = $Group.Created
		$MemberOfs = $Group.MemberOf -split ';'
		
		# removed used properties 
		"SamAccountName","DistinguishedName","OU","MemberOf","Created" | foreach {
			$Group.PSObject.Properties.Remove($_)
		}

		# create xml group node
		$XmlGroup = New-XmlElement $XmlGroups -Name "ADGroup" -Attribute $GroupProp -Child $Group
		
		# add MemberOf property
		if ($MemberOfs) {
			$XmlMemberOfs = New-XmlElement $XmlGroup -Name "MemberOfs" -Attribute @{ 'Count' = $MemberOfs.Count }
			$MemberOfs | sort | foreach { 
				New-XmlElement -XmlParent $XmlMemberOfs -Name 'MemberOf' -Innertext $_ | Out-Null
			}
		}
	}
	$Xml.Save($FileName)
}
  
break
# ===========================================================================
# 	End of script
