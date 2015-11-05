#
# GetSoftwareInventory.ps1
#
# 	Author: INSIGMA IT Engineering GmbH, Ralf Nelle
# 	Last Change: 05.11.2015
#


# ---------------------------------------------------------------------------
#     Settings

$ComputersFile = $pwd.path + "\hosts.txt"
$XmlFile = $pwd.path + "\SoftwareInv.xml"
$XmlCleanFile = $pwd.path + "\SoftwareInv_cleaned.xml"
$NoiseProductsFile = $pwd.path + "\NoiseProducts.txt"

# prepare xml store
[Xml]$Xml = @'
<?xml version="1.0" encoding="iso-8859-1"?>
<Inventory>
  <Computers />
</Inventory>
'@


# ---------------------------------------------------------------------------
#     Functions

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
			try { 
				$Result = gwmi -q "SELECT Caption FROM Win32_Operatingsystem" -ComputerName $Computer -ErrorAction Stop
				New-Object -TypeName PSCustomObject -Property @{
				  'ComputerName' = $Computer; 'Name' = $result.Caption
				}
			}
			catch { Write-Host "$($Computer): $($_.Exception.Message)" -fore red -back black } 					
		}
	}
}

# ---------------------------------------------------------------------------
Function Get-RemoteSoftware ($ComputerName = $ENV:COMPUTERNAME) {

  # Registry settings
  $HKLM = [UInt32] "0x80000002"
  $UNINSTALL_KEY = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"

  # Create a hash table containing the requested application properties.
  $propertyList = @{}

  # Iterate the computer name(s).
  foreach ($machine in $ComputerName) {
    $err = $NULL

    # If WMI throws a RuntimeException exception,
    # save the error and continue to the next statement.
    #CALLOUT B
    trap [System.Management.Automation.RuntimeException] {
      set-variable err $ERROR[0] -scope 1
      continue
    }
    #END CALLOUT B

    # Connect to the StdRegProv class on the computer.
    #CALLOUT C
    $regProv = [WMIClass] "\\$machine\root\default:StdRegProv"

    # In case of an exception, write the error
    # record and continue to the next computer.
    if ($err) {
      write-error -errorrecord $err
      continue
    }
    #END CALLOUT C

    # Enumerate the Uninstall subkey.
    $subkeys = $regProv.EnumKey($HKLM, $UNINSTALL_KEY).sNames
    foreach ($subkey in $subkeys) {
      # Get the application's display name.
      $name = $regProv.GetStringValue($HKLM,
        (join-path $UNINSTALL_KEY $subkey), "DisplayName").sValue
      # Only continue of the application's display name isn't empty.
      if ($name -ne $NULL) {
        # Create an object representing the installed application.
        $output = new-object PSObject
        $output | add-member NoteProperty ComputerName -value $machine
        #$output | add-member NoteProperty AppID -value $subkey
        $output | add-member NoteProperty AppName -value $name
        $output | add-member NoteProperty Publisher -value `
          $regProv.GetStringValue($HKLM,
          (join-path $UNINSTALL_KEY $subkey), "Publisher").sValue
        $output | add-member NoteProperty Version -value `
          $regProv.GetStringValue($HKLM,
          (join-path $UNINSTALL_KEY $subkey), "DisplayVersion").sValue
        # If the property list is empty, output the object;
        # otherwise, try to match all named properties.
        if ($propertyList.Keys.Count -eq 0) {
          $output
        } else {
          #CALLOUT D
          $matches = 0
          foreach ($key in $propertyList.Keys) {
            if ($output.$key -like $propertyList.$key) {
              $matches += 1
            }
          }
          # If all properties matched, output the object.
          if ($matches -eq $propertyList.Keys.Count) {
            $output
            # If -matchall is missing, break out of the foreach loop.
            if (-not $MatchAll) {
              break
            }
          }
          #END CALLOUT D
        }
      }
    }
  }
}


# ===========================================================================
#   MAIN

# get list from file
if (!(Test-Path $ComputersFile)) {
	Write-Host "'$ComputersFile' not found. Script aborted." -fore red -back black; break;
}
$Computers = gc $ComputersFile | ?{ $_ -notlike "#*" }

# get exiting xml file
if (Test-Path $XmlFile) { [xml]$Xml = gc $XmlFile }

# get root note
$XmlComputers = $Xml.SelectSingleNode("//Computers")


# get operating system
Write-Host "Getting host OS ..." -fore gray
foreach ($Result in (Get-Operatingsystem $Computers | Sort ComputerName)) {

	# add computer node
	$XmlComputer = New-XmlElement $XmlComputers -Name "Computer" -Attribute @{ 'Name' = $Result.ComputerName }
	New-XmlElement $XmlComputer -Name 'Software' -InnerText $Result.Name | Out-Null

}
$Xml.Save($XmlFile)


# get details from active hosts
$Results = Get-RemoteSoftware ($xml.SelectNodes("//Computer") | %{ $_.Name })
$InstalledSoftware = $Results | Group ComputerName -AsHashTable -AsString

foreach ($Computer in $InstalledSoftware.Keys) {

	Write-Host "Examing '$Computer' software ..." -fore gray
	$XmlComputer = $Xml.SelectSingleNode("//Computer[@Name='$($Computer)']")
	foreach ($Software in ($InstalledSoftware[$Computer] | %{ $_.AppName } | sort | Get-Unique)) {
		New-XmlElement $XmlComputer -Name 'Software' -InnerText $Software | Out-Null
	}
}
$Xml.Save($XmlFile)
	
# clean product list 	
$NoiseProducts = gc $NoiseProductsFile -Encoding String | %{ $_.trim() }
foreach ($Software in $Xml.SelectNodes("//Software")) {
	[string]$SwCaption = $Software."#text"
	if ($NoiseProducts -contains $SwCaption.trim()) { 
		[Void]$Software.ParentNode.RemoveChild($Software)
	}
}
$Xml.Save($XmlCleanFile)


