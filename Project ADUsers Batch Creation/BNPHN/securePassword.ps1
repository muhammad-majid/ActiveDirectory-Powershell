Function New-SecurePassword
{
	$Pass = "!?@#$%^&*23456789ABCDEFGHJKMNPQRSTUVWXYZ_abcdefghijkmnpqrstuvwxyz".tochararray()
	($Pass | Get-Random -Count 8) -Join ''
}


$password = (New-SecurePassword)
#Write-Host "$($password)`r`n"
Write-Host "$(New-SecurePassword)`r`n"
