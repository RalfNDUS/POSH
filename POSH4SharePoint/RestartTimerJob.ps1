# ===================================================================================
# Func: Restarts the timer service in the farm
# ===================================================================================
function RestartTimer
{
    Write-Host  " - Restarting OWSTIMER instances on Farm"
    $farm = Get-SPFarm
    $farm.TimerService.Instances | foreach {$_.Stop();$_.Start();}

	Get-Process OWSTimer| Stop-Process –Force;
	start-sleep 5
}

