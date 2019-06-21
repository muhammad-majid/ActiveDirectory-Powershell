$OUBase = "OU=Users, OU=Crisis Support Services, DC=css, DC=local"
$users = Get-ADUser -filter * -SearchBase $OUBase
ForEach ($User in $users)
{
	$aduserobject = [ADSI]"LDAP://$($User.DistinguishedName)"
	$aduserobject.putex(1, "msNPAllowDialin", $null)
	$aduserobject.SetInfo()
}
