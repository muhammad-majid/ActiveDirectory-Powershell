function Get-RandomCharacters($length, $characters)
{
    $random = 1..$length | ForEach-Object { Get-Random -Maximum $characters.length }
    $private:ofs=""
    return [String]$characters[$random]
}
 
function Scramble-String([string]$inputString)
{
    $characterArray = $inputString.ToCharArray()   
    $scrambledStringArray = $characterArray | Get-Random -Count $characterArray.Length     
    $outputString = -join $scrambledStringArray
    return $outputString 
}

function New-RandomPassword ($len)
{ 
	$pass = Get-RandomCharacters -length 1 -characters 'ABCDEFGHKMNPRSTUVWXYZ'
	$pass += Get-RandomCharacters -length 1 -characters '23456789'
	$pass += Get-RandomCharacters -length 1 -characters '!$%&/()=?}][{@#*+'
	$pass += Get-RandomCharacters -length ($len-3) -characters 'abcdefghikmnprstuvwxyz'
	$pass = Scramble-String ($pass)
	return $pass
}

$password = New-RandomPassword(20)
Write-Host $password
