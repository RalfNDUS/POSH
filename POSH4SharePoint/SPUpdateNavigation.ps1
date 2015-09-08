<#
    SPUpdateNavigation
	SharePoint 2010 Navigation Bar Update
	Author: Ralf Nelle
	Last Change: 04.09.2014
#>

Param(
	[parameter(Mandatory=$true)][String]$Url
)

Add-PSSnapIn Microsoft.Sharepoint.PowerShell -EA 0

function ProcessSubWebs($currentWeb)
{       
   foreach($sub in $currentWeb.Webs)
   {
      if($sub.Webs.Count -ge 0)
      {
         Write-Host -ForegroundColor gray $sub.Url
         UpdateNavigation($sub)
         ProcessSubWebs($sub)
         $sub.Update()
         $sub.Dispose()
      }            
   }        
}

function UpdateNavigation($web)
{
    $pubWeb = [Microsoft.SharePoint.Publishing.PublishingWeb]::GetPublishingWeb($web)
    Write-Host -ForegroundColor yellow $pubWeb.Navigation.GlobalIncludeSubSites
    $pubWeb.Navigation.InheritGlobal = $true
    $pubWeb.Navigation.GlobalIncludeSubSites = $false
    $pubWeb.Navigation.GlobalIncludePages = $false
    $web.AllowUnsafeUpdates = 1;
	$web.Update();
}

ProcessSubWebs(Get-SPWeb -identity $Url)
Write-Host -ForegroundColor green "FINISHED"
