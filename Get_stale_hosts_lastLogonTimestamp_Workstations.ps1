# Gets time stamps for all computers in the domain that have NOT logged in since after specified date
 
$time = Read-host "Enter a date in format mm/dd/yyyy"
$time = get-date ($time)
$date = get-date ($time) -UFormat %d.%m.%y
 
# Get all AD computers with lastLogonTimestamp less than our time
Get-ADComputer -Filter {LastLogonTimeStamp -lt $time} -SearchBase "OU=Workstations,OU=Emirates Steel,DC=eisf,DC=co,DC=ae" -Properties lastLogonTimeStamp, CanonicalName |
 
# Output hostname and lastLogonTimestamp into CSV
select-object Name, @{Name="Last Logon TimeStamp"; Expression={[DateTime]::FromFileTime($_.lastLogonTimestamp)}}, CanonicalName | export-csv Stale-Hosts-Workstations.csv -notypeinformation


