# load the needed PowerShell snapins
Add-PSSnapIn SqlServer* -EA 0
Add-PsSnapin Microsoft.SharePoint.PowerShell -EA 0


# Hash Tables 
$SQLServerVersion = @{
	"10.0"  = "SQL Server 2008"; 
	"10.50" = "SQL Server 2008 R2"; 
	"11.0"  = "SQL Server 2012"; 
	"12.0"  = "SQL Server 2014" 
}

# Get SharePoint products Guid 
# Reference: http://msdn.microsoft.com/en-us/library/ff721969.aspx
$Products = @{
	"BEED1F75-C398-4447-AEF1-E66E1F0DF91E" = "SharePoint Foundation 2010";
	"1328E89E-7EC8-4F7E-809E-7E945796E511" = "Search Server Express 2010";
	"3FDFBCC8-B3E4-4482-91FA-122C6432805C" = "SharePoint Server 2010 Standard";
	"D5595F62-449B-4061-B0B2-0CBAD410BB51" = "SharePoint Server 2010 Enterprise";
	"9FF54EBC-8C12-47D7-854F-3865D4BE8118" = "SharePoint Foundation 2013";
	"C5D855EE-F32B-4A1C-97A8-F0A28CE02F9C" = "SharePoint Server 2013 Standard";
	"B7D84C2B-0754-49E4-B7BE-7EE321DCE0A9" = "SharePoint Server 2013 Enterprise"
}

# Get the SharePoint version 
# Reference: http://www.toddklindt.com/blog/Lists/Posts/Post.aspx?ID=224
$Versions = @{
  '14.0.5114' = 'June 2010 CU';
  '14.0.5123' = 'August 2010 CU';
  '14.0.5128' = 'October 2010 CU';
  '14.0.5130' = 'December 2010 CU';
  '14.0.5136' = 'February 2011 CU';
  '14.0.5138' = 'April 2011 CU';
  '14.0.6029' = 'Service Pack 1';
  '14.0.6105' = 'June 2011 CU';
  '14.0.6106' = 'June 2011 CU Mark 2';
  '14.0.6109' = 'August 2011 CU';
  '14.0.6112' = 'October 2011 CU';
  '14.0.6114' = 'December 2011 CU';
  '14.0.6117' = 'February 2012 CU';
  '14.0.6120' = 'April 2012 CU';
  '14.0.6123' = 'June 2012 CU';
  '14.0.6126' = 'August 2012 CU';
  '14.0.6129' = 'October 2012 CU';
  '14.0.6131' = 'December 2012 CU';
  '14.0.6134' = 'February 2013 CU';
  '14.0.6137' = 'April 2013 CU';
  '14.0.7011' = 'SP2 Public Beta';
  '14.0.7015' = 'Service Pack 2';
  '14.0.7102' = 'June 2013 CU Mark 1';
  '14.0.7106' = 'August 2013 CU';
  '14.0.7110' = 'October 2013 CU';
  '14.0.7113' = 'December 2013 CU';
  '14.0.7116' = 'February 2014 CU';
  '14.0.7121' = 'April 2014 CU';
  '14.0.7123' = 'MS14-022';
  '14.0.7125' = 'June 2014 CU';
  '14.0.7128' = 'July 2014 CU';
  '14.0.7130' = 'August 2014 CU';
  '14.0.7132' = 'September 2014 CU';
  '14.0.7134' = 'October 2014 CU';
  '14.0.7137' = 'November 2014 CU';
  '14.0.7140' = 'December 2014 CU';
  '14.0.7143' = 'February 2015 CU';
  '14.0.7145' = 'March 2015 CU'; 
  '14.0.7147' = 'April 2015 CU'; 
  '14.0.7149' = 'Mai 2015 CU'; 
}

### get SQL Server information
Write-Progress -Activity "Gathering SQL Server information ..." -status "Processing"
$sql = dir SQLSERVER:\sql\$($ENV:COMPUTERNAME)

Write-Host
Write-Host "Installed SQL Server:" -fore yellow
$sql | Select-Object `
	@{Name="InstanceName"; Expression={$_.Name}},
	@{Name="Name"; Expression={$SQLServerVersion[$_.VersionMajor,$_.VersionMinor -join '.']}},
	@{Name="Edition"; Expression={$_.Edition}},
	@{Name="Language"; Expression={$_.Language}},
	@{Name="ProductLevel"; Expression={$_.ProductLevel}},
	@{Name="VersionString"; Expression={$_.VersionString}}

### get SharePoint information
Write-Progress -Activity "Gathering SharePoint information ..." -status "Processing"

Write-Host "Installed SharePoint Products:" -fore yellow
(Get-SPFarm).Products | foreach { $Products[$_.Guid] }

Write-Host "`r`nSharePoint Patchlevel:" -fore yellow
$BuildVersion = (Get-SPFarm).BuildVersion.ToString().SubString(0,9)
$PatchLevel = $Versions[$BuildVersion]
if (!$PatchLevel) { $PatchLevel = "unknown (Build Version: $BuildVersion)" }
$PatchLevel
Write-Host

### get SharePoint languages
$ca = Get-SPWebApplication -IncludeCentralAdministration | where { $_.IsAdministrationWebApplication }
$caRegionalSettings = (Get-SPWeb $ca.Url).RegionalSettings

Write-Host "SharePoint Server Language:" -fore yellow
$caRegionalSettings.ServerLanguage.DisplayName

Write-Host "`r`nInstalled Languages:" -fore yellow
$caRegionalSettings.InstalledLanguages | %{ $_.DisplayName }

  
