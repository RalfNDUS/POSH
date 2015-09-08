# ---------------------------------------------
#
#  ShrinkAllSqlDB
#
#  Get all databases from SQL instance and shrink them
#
#  Author:  Ralf Nelle
#  Last change: 09.12.2010
#
# ---------------------------------------------

Add-PSSnapIn SqlServer* -EA 0
if ((Get-PSSnapin "SqlServer*").Count -ne 2) {
	Write-Error "PowerShell SQL-SnapIns not found. Script aborted!"
	break
}

# ===========================================================================
# 	SQL Maintenance job definition
$SqlCmd = @"
/*
    Only use this script for SQL Server development servers!
    Script must be executed as sysadmin

    This script will execute the following actions on all databases
        - trucate log file
        - shrink log file
*/

use [master]
go

-- Declare container variabels for each column we select in the cursor
declare @databaseName nvarchar(128)

-- Define the cursor name
declare databaseCursor cursor
-- Define the dataset to loop
for
select [name] from sys.databases

-- Start loop
open databaseCursor

-- Get information from the first row
fetch next from databaseCursor into @databaseName

-- Loop until there are no more rows
while @@fetch_status = 0
begin
    print 'Shrinking logfile for database [' + @databaseName + ']'
    exec('
    use [' + @databaseName + '];' +'

    declare @logfileName nvarchar(128);
    set @logfileName = (
        select top 1 [name] from sys.database_files where [type] = 1
    );
    dbcc shrinkfile(@logfileName,1);
    ')
	
    -- Get information from next row
    fetch next from databaseCursor into @databaseName
end

-- End loop and clean up
close databaseCursor
deallocate databaseCursor
go

"@


# ===========================================================================
# 	Get SQL instances on local machine

$sql = dir SQLSERVER:\sql\$($ENV:COMPUTERNAME) | where { $_.Version -like "10.5*" }

# 	show information
@"
Server Instance: $($sql.Name)
Edition: $($sql.Edition)
Version: $($sql.VersionString)
ProductLevel: $($sql.ProductLevel)
"@ | Write-Host -fore gray



# ===========================================================================
# 	Execute job 

$Result = Invoke-Sqlcmd -Query $SqlCmd -ServerInstance $sql
$Result




