<#
  .SYNOPSIS
  List NTFS permissions
  .INFO
  INSIGMA IT Engineering GmbH, Ralf Nelle
  Last Change: 03.03.2015
#>

Function Get-NTFSPermission {

  Param(
    [Parameter(Mandatory=$true)][String]$Path,
	[Switch]$Recurse = $false
  )

  if (!(Test-Path $Path)) {
    Write-Host "Path '$Path' not found. Script aborted.`n" -fore red -back black
    break
  }

  # get all folders
  Write-Progress -Activity "Gathering folders in '$($Path)' ..." -status "Processing"
  $folders = Get-ChildItem -Path $Path -Recurse | `
    ?{ $_.PSIsContainer } | Select-Object -ExpandProperty FullName
	
  # include base path 
  $folders = ($folders += $Path) | sort  

  # get permissions  
  Write-Progress -Activity "Getting folder permissions ..." -status "Processing"

  Foreach ($folder in $folders) {
	$acl = Get-ACL $folder
	if (($folder -like $Path) -or ($acl.AreAccessRulesProtected -eq $True)) {
	
	  $strAccess = $acl.AccessToString -split '\n' | ?{ $_ -notlike "CREATOR OWNER*" } | sort | get-unique 
	  New-Object psobject -Property @{
		'Path' = $folder
		'Access' = $strAccess -join "`r`n"
	  }
	}
  }
}
