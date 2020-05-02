Import-Module ActiveDirectory
$groupname = 'DSG_ManagedUsers'
$OUname = 'OU=Managed Users,OU=iEmirates Steel,DC=ieisf,DC=co,DC=ae'
$users = Get-ADUser -Filter * -SearchBase $OUname

$members = Get-ADGroupMember -Identity $groupname
foreach($member in $members)
{
  if($member.distinguishedname -notlike "*$OUname*")
  {
    Remove-ADGroupMember -Identity $groupname -Member $member.samaccountname
  }
}

foreach($user in $users)
{
  Add-ADGroupMember -Identity $groupname -Member $user.samaccountname -ErrorAction SilentlyContinue
}
