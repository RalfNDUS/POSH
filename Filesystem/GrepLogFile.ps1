<#
    GrepLogFile
	Written by:  Ralf Nelle
	Last Change: 28.08.2015
#>

# Functions
Filter grep($keyword) { if( ($_ | Out-String) -like "*$keyword*") { $_ } }

Function Grep-LogFile {

	Param(
		[Parameter(Mandatory=$true)]
		[string]$FilePath = $pwd.Path,
		[Parameter(Mandatory=$true)]
		[string]$Keyword,	
		[int]$Days = -1,
		$ResultFile = "$($pwd.path)\Grep-LogFile.txt"
	)

	# MAIN
	Write-Progress -Activity "Searching for '$($Keyword)' in '$($FilePath)' ..." -status " "

	$LogFiles = get-childitem $FilePath -Include *.log -Recurse | `
	  ?{ $_.CreationTime -ge (Get-Date).AddDays($days) } | %{ $_.FullName }

	$Result = &{ $LogFiles | %{ gc $_ | grep $keyword } }

	$Result | sort | tee-object -filepath $ResultFile 
	notepad $ResultFile

}