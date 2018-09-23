$addn     = (Get-ADDomain).DistinguishedName
$dnsroot  = (Get-ADDomain).DNSRoot
$GivenName="trial"
$LastName="1"
$Name=$GivenName+" "+$LastName
$sam=$GivenName+"."+$LastName
$userPrincipalName=$sam+"@"+$dnsroot

$propertiesImported = @("department", "title", "description", "distinguishedName", "physicalDeliveryOfficeName", "streetAddress", "l", "postalCode", "st", "co", "company", "manager")

$propertiesToExport = @{
			"givenName"=$GivenName
			"surname"=$LastName
			#"name"=$Name
			"displayName"=$Name
			"UserPrincipalName"=$userPrincipalName
			}
			
$exists = Get-ADUser -Identity "user1" -Properties $propertiesImported
New-ADUser $sam
Set-ADUser -Identity $sam @propertiesToExport
