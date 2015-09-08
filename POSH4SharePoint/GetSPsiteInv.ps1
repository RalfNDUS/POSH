<#
	.Synopsis
		Exports SharePoint site permissions

    .Parameter Site
		Gibt die URL der Websitesammlung an, von der die Unterwebsites aufgelistet werden.

    .Notes
        Author: INSIGMA IT Engineering GmbH, Ralf Nelle
		Last Change: 10.06.2015
#>

Param(
	[Parameter(Mandatory=$true)]
    [String]$Site,
	[String[]]$Exclude
)


# ---------------------------------------------------------------------------
#     Settings

$FileName = "{0}\SPsiteInv_{1:yyyy-MM-dd}.xml" -f $pwd.path,[DateTime]::now


# ---------------------------------------------------------------------------
#     Functions

Function isNotExcluded($str,$patterns){
    foreach($pattern in $patterns) { if($str -like $pattern) { return $false; } }
    return $true;
}

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
Add-PSSnapIn Microsoft.SharePoint.PowerShell -EA 0
if (!(Get-PSSnapIn Microsoft.SharePoint.PowerShell)) {
	Write-Host "PSSnapIn 'Microsoft.SharePoint.PowerShell' not found. Script aborted.`n" -fore red -back black
	break
}
	
# prepare xml store
[Xml]$Xml = @'
<?xml version="1.0" encoding="iso-8859-1"?>
<SPSites />
'@ 

	
# ---------------------------------------------------------------------------	
# 	Get site

# get site
try {
	$spSite = Get-SPSite $Site
} catch {
    Write-Warning "$($_.Exception.Message)"
	break
}


# add site to xml store
$XmlSites = $Xml.SelectSingleNode("//SPSites")
$XmlSite = New-XmlElement $XmlSites -Name "SPSite" -Attribute @{ 'Url' = $spSite.Url }
$XmlWebs = New-XmlElement $XmlSite -Name "SPWebs"

$AllWebs = $spSite.AllWebs
if ($Exclude) { $AllWebs = $AllWebs | ?{ isNotExcluded $_.Url $Exclude }}

foreach ($web in $AllWebs) {
	$XmlWeb = New-XmlElement $XmlWebs -Name "SPWeb" -Attribute @{ 'Url' = $web.Url }
	$XmlWeb.SetAttribute('HasUniquePerm',"$($web.HasUniquePerm)")
	
	if ($web.HasUniquePerm) {
	
		$XmlRoleAssignments = New-XmlElement $XmlWeb -Name "RoleAssignments"
		
		$RoleAssignments = $web.RoleAssignments | Select-Object Member,RoleDefinitionBindings
		foreach ($RoleAssignment in $RoleAssignments) {
		
			$XmlRoleAssignment = New-XmlElement $XmlRoleAssignments -Name "RoleAssignment" -Attribute @{ 
				'Member' = $RoleAssignment.Member.LoginName
				'Permissions' = ($RoleAssignment.RoleDefinitionBindings | %{ $_.Name }) -join ';'
			}
			
			if ($RoleAssignment.Member.GetType().FullName -eq "Microsoft.SharePoint.SPGroup") {
				foreach ($User in $RoleAssignment.Member.Users) {
					New-XmlElement $XmlRoleAssignment -Name "SPUser" -InnerText $User.UserLogin | Out-Null
				}
			}
		}
	}
}

$Xml.Save($FileName)

break
# ===========================================================================
# 	End of script
