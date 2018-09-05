Import-Module ActiveDirectory

$dcs = (Get-ADDomain).ReplicaDirectoryServers
$svcs = "adws","dns","kdc","netlogon","dhcp"
$filter = "(Name = 'ADS' OR Name = 'DNS' OR Name = 'KDC' OR Name = 'NetLogon') AND State<>'Running'"

Get-Service -name $svcs -ComputerName $dcs | Sort Machinename | Format-Table -group @{Name="Computername";Expression={$_.Machinename.toUpper()}} -Property Name,Displayname,Status
Get-WmiObject -Class Win32_service -filter $filter -ComputerName $dcs | Select PSComputername,Name,Displayname,State,StartMode | format-table -autosize