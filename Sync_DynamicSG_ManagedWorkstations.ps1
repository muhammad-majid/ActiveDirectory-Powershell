Import-Module ActiveDirectory
$groupname = 'DSG_ManagedWorkstations'
$OUname = 'OU=Managed Workstations,OU=iEmirates Steel,DC=ieisf,DC=co,DC=ae'
$workstations = Get-ADComputer -Filter * -SearchBase $OUname

$members = Get-ADGroupMember -Identity $groupname
foreach($member in $members)
{
  if($member.distinguishedname -notlike "*$OUname*")
  {
    Remove-ADGroupMember -Identity $groupname -Member $member.samaccountname
  }
}

foreach($workstation in $workstations)
{
  Add-ADGroupMember -Identity $groupname -Member $workstation.samaccountname -ErrorAction SilentlyContinue
}
