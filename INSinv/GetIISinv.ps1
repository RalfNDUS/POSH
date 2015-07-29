#
# GetIISinv.ps1
#
#	Author: INSIGMA IT Engineering GmbH, Ralf Nelle
#	Last Change: 29.07.2015
#

# ---------------------------------------------------------------------------
#     Settings

$FileName = "{0}\{1}_Websites.xml" -f $pwd.ProviderPath,$ENV:COMPUTERNAME	

[Xml]$Xml = @'
<?xml version="1.0" encoding="iso-8859-1"?>
<WebAdministration>
    <Websites />
</WebAdministration>
'@
	

# ---------------------------------------------------------------------------
#     Functions

function Get-IPAddress{
  param(
    [Parameter(ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
    [String[]]$Name
   )  
  process
  { 
	  $Name | ForEach-Object { 
		  try { 
			  [System.Net.DNS]::GetHostByName($_) 
		  } 
		  catch { } 
	  }
  }
}


# ===========================================================================
#   MAIN

# abort, if module not found
Import-Module WebAdministration  -EA 0
if (!(Get-Module WebAdministration)) {
	Write-Host "Module 'WebAdministration' not found. Script aborted.`n" -fore red -back black
	break
}
	

# get websites, apppools and ssl bindings
$Sites = Get-ChildItem IIS:\Sites | sort Name
$AppPools = Get-ChildItem IIS:\AppPools | sort Name

# get certificate store
$Certs = Get-ChildItem Cert:\LocalMachine\My | select Thumbprint,FriendlyName,NotAfter,Issuer

# get ssl bindings
$SslBindings = dir IIS:\SslBindings

# set site information object

$WebsiteCol = @()
foreach ($Site in $Sites) {

	$SiteObj = $Site | Select-Object `
		@{Name="Name"; Expression={$_.name}},
		@{Name="Bindings";Expression={($_.bindings.Collection | %{ $_.protocol + ' ' + $_.bindingInformation }) -join ','}},
		@{Name="PhysicalPath"; Expression={$_.physicalPath}},
		@{Name="LogfileDirectory"; Expression={$_.logFile.directory}},
		@{Name="ApplicationPool"; Expression={$_.applicationPool}},
		@{Name="managedRuntimeVersion"; Expression={''}},
		@{Name="ID"; Expression={$_.id}},
		@{Name="State"; Expression={$_.state}},
		@{Name="CertificateName"; Expression={''}},
		@{Name="NotAfter"; Expression={''}},
		@{Name="Issuer"; Expression={''}},
		@{Name="Hostname"; Expression={$ENV:COMPUTERNAME}},
		@{Name="Timestamp"; Expression={[DateTime]::now}}	

	# get application pool details
	$AppPool = $AppPools | ?{ $_.Name -eq $Site.Name }
	$SiteObj.managedRuntimeVersion = $AppPool.managedRuntimeVersion
	
	# get SSL certificate
	$SslBinding = $Site.bindings.Collection | ?{ $_.protocol -eq "https" }
	$Thumbprint = ($SslBindings | ?{ $_.Sites -eq $Site.Name }).Thumbprint
	if ($Thumbprint) {

		$SiteCert = $certs | where { $_.Thumbprint -eq $Thumbprint }
		if ($SiteCert) {
			$SiteObj.CertificateName = $SiteCert.FriendlyName
			$SiteObj.NotAfter = $SiteCert.NotAfter
			$SiteObj.Issuer = $SiteCert.Issuer
		}
	}
	
	$WebsiteCol += $SiteObj
}


# export to xml
$XmlRoot = $Xml.SelectSingleNode("//Websites")

foreach ($Website in $WebSiteCol) {

	$xmlSite = $xml.CreateElement("Website")
	[void]$XmlRoot.AppendChild($xmlSite)

	$WebSite.psobject.properties | Foreach { 
		$xmlNode = $xml.CreateElement($_.Name)
		$xmlNode.Set_Innertext($_.Value)
		[void]$xmlSite.AppendChild($xmlNode)
	}
	
	# add public IP
	$xmlNode = $xml.CreateElement("PublicIP")
	[void]$xmlSite.AppendChild($xmlNode)
	
	# resolve DNS name
	try { 
		$dns = ([System.Net.DNS]::GetHostByName($Website.Name)).AddressList
		if ($dns) { 
			$PublicIP = ($dns | %{ $_.IPAddressToString }) -join ','	
			$xmlNode.Set_Innertext($PublicIP) } 
		}
    catch { }
}

$xml.Save($FileName)
	
