Function CheckTimeServers {
BEGIN {}
PROCESS {
$Server = "$_"
if ($_ -ne "") {
Write-host "Checking time source of $Server, pasting output in $($LogFile)"
$TimeServer = w32tm /query /computer:$Server /source
Write-output "[$(Get-Date -format g)]: Server $($Server) gets its time from $($TimeServer)" >> $LogFile
}
}
END {}
}
cls
$ScriptPath = Split-Path $MyInvocation.MyCommand.Path
$LogFile = $Scriptpath+"\Output.txt"
Get-Content $Scriptpath"\ServerList.txt" | CheckTimeServers