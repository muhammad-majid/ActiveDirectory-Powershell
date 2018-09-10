 Import-Module ActiveDirectory  

$users = Import-Csv -Path c:\update.csv

foreach ($user in $users)
{
	Get-ADUser -Filter "employeeID -eq '$($user.employeeID)'" -Properties * -SearchBase "ou=Test,ou=OurUsers,ou=Logins,dc=domain,dc=com" | 
	Set-ADUser -employeeNumber $($user.employeeNumber) -department $($user.department) -title $($user.title) -office $($user.office) -streetAddress $($user.streetAddress) -City $($user.City) -state $($user.state) -postalCode $($user.postalCode) -OfficePhone $($user.OfficePhone) -mobile $($user.mobile) -Fax $($user.Fax) -replace @{"extensionAttribute1"=$user.extensionAttribute1; "extensionAttribute2"=$user.extensionAttribute2; "extensionAttribute3"=$user.extensionAttribute3}
}