<#
    Profile.ps1
	
    .INFO
    PowerShell Profile (SharePoint 2013)
    Author: INSIGMA IT Engineering GmbH, Ralf Nelle
    Last Change: 09.09.2015
#>


# ---------------------------------------------------------------------------
#     Settings

$Transcript = "{0}\PowerShell_transcript.{1:yyyyMMdd}.txt" -f `
  [environment]::GetFolderPath("MyDocuments"),[DateTime]::now

### Mail settings 

$CN = $ENV:COMPUTERNAME
$UN = $ENV:USERNAME
$DN = ($ENV:USERDNSDOMAIN).ToLower()

$from = "$UN@$DN"
$to = "RNelle@insigma.de"
$subject = "Message from $CN.$DN"  

if ($CN -eq "SHAREPOINT01") {
  $SMTPServer = "192.168.9.15" 
} else {  
  $SMTPServer = "INSIGMA14v.insigma.int"
}
  

# ---------------------------------------------------------------------------
#     Funktionen

Filter grep($keyword) { if( ($_ | Out-String) -like "*$keyword*") { $_ } }

Function pro { notepad $profile.AllUsersAllHosts }

Function script { notepad $transcript }

# ---------------------------------------------------------------------------
Function send
([Parameter(Mandatory=$true)][string[]]$file) 
{
  Begin {}
  Process {
    try {
      Send-MailMessage -From $from -To $to -Subject $subject `
        -SmtpServer $SmtpServer -Attachments $file -ErrorAction Stop
      Write-Host "Message send to '$to'." -fore gray  
    } catch {
      Write-Warning "Cannot send mail: $($_.Exception.Message)"
      return
    }
  }
  End {}
}

# ---------------------------------------------------------------------------
Function TopSyslogEvents
($ComputerName=$ENV:COMPUTERNAME) 
{
  Begin {}
  Process {
    try {
		Get-EventLog -Log System -After ((Get-Date).AddHours(-24)) -EntryType Error,Warning `
		  -Newest 10 -ComputerName $ComputerName | `
		    fl TimeGenerated,MachineName,EventID,EntryType,Source,Message
    } catch {
      Write-Warning "Cannot access system log: $($_.Exception.Message)"
      return
    }
  }
  End {}
}

# ---------------------------------------------------------------------------
Function TopApplogEvents
($ComputerName=$ENV:COMPUTERNAME) 
{
  Begin {}
  Process {
    try {
		Get-EventLog -Log Application -After ((Get-Date).AddHours(-24)) -EntryType Error,Warning `
		  -Newest 10 -ComputerName $ComputerName | `
		    fl TimeGenerated,MachineName,EventID,EntryType,Source,Message
    } catch {
      Write-Warning "Cannot access application log: $($_.Exception.Message)"
      return
    }
  }
  End {}
}

# ---------------------------------------------------------------------------
Function SetMasterPageFromParent ([Parameter(Mandatory=$true)][string]$Url) {
  $web = Get-SPWeb $Url
  if ($web -eq $null) { break; }
		
  $web.MasterUrl = $web.ParentWeb.MasterUrl
  $web.CustomMasterUrl = $web.ParentWeb.CustomMasterUrl
  $web.update()
  Write-Host "$($web.url) MasterUrl/CustomMasterUrl changed." -fore green
} 

# ---------------------------------------------------------------------------
Function loadSPSnapIn {
  if ($host.Version -ne "2.0") { 
    Stop-Transcript
    Clear-Host
    powershell.exe -version 2.0 
  }
}
	
# ---------------------------------------------------------------------------
Function AddUserprofile ($UserName) {
  $objUser = [ADSI]("WinNT://" + $UserName.replace("\","/"))
  $objGroup = [ADSI]("WinNT://$($ENV:COMPUTERNAME)/Administrators")
  $objGroup.PSBase.Invoke("Add",$objUser.PSBase.Path)

  runas /user:$($UserName) /profile cmd
  Read-Host "Type 'exit' in cmd to continue script" | Out-Null

  $objGroup.PSBase.Invoke("Remove",$objUser.PSBase.Path)
}
	

# -----------------------------------------------------------------------	
#	MAIN

# load SharePoint SnapIn
if ($host.Version -eq "2.0") {
  Write-Progress -Activity "PS-Modules and -SnapIns" -status "Loading PSSnapIn: Microsoft.Sharepoint.PowerShell"
  Add-PSSnapIn Microsoft.SharePoint.PowerShell -EA 0
} else {
  Write-Host "Type 'loadSPSnapIn' for SharePoint-PSSnapIn ...`n" -fore gray
}

# Goto Script dir
cd C:\PSScripts

# Start-Transcript
Write-Host (&{ Start-Transcript -Append }) -fore gray; Write-Host
