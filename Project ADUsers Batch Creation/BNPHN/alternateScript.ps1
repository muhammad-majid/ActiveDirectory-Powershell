Create new AD accounts#

write-host "Creating Student Accounts"
foreach ($student in $students){
$fname = (RemoveSpecials($student.Firstname))
$lname = (RemoveSpecials($student.Lastname)) 
$Fullname = $fname + " " + $lname
$gradyr = $student."Graduation Year"
$id = $student."Student ID" 
$username = $fname + "." + $lname + $gradyr.substring(2)
if ($username.length -gt 20) { $username = $username.substring(0,20) } #shorten username to 20 characters for sAMAccountName
$emailadd = $fname+"."+$lname+ $gradyr.substring(2) + "@" + $stuemail
$principalname = $fname+"."+$lname+ $gradyr.substring(2) + "@" + $stuemail
$homedir = $stuhomedir1 + "\" + $gradyr+ "\"+ $username
$building = $student."Current Building" #Edit this to match your CSV If your header is not exactly Current Building

Switch ($Student."Current Building"){
    "8" {$stubuildingou ='ou=elementary'}
    "9" {$stubuildingou ='ou=Highschool'}
    "11" {$stubuildingou ='ou=MiddleSchool'}
    }

Write-host $Fullanme $username $password $building
	#create the new user with long samaccountname
	Write-Host `nCreating new user: $username in $gradyr
New-Aduser `
-sAMAccountName $username `
-givenName $fname `
-Surname $lname `
-UserPrincipalName $principalname `
-DisplayName $fullname `
-name $fullname -homeDrive "h:" `
-homeDirectory $homedir `
-scriptPath "logon.bat" `
-EmailAddress $emailadd `
-EmployeeID	 $id `
-ChangePasswordAtLogon $true `
-AccountPassword (ConvertTo-SecureString "$password" -AsPlainText -force) `
-Enabled $true `
-Path "ou=$gradyr,$stubuildingou,$stuou"`
-Department 'student'

If (Test-Path $homedir -PathType Container)
    {Write-host "$homedir already exists"}
    Else
    {New-Item -path $homedir -ItemType directory -Force}


$IdentityReference=$Domainacl+$username

$AccessRule=NEW-OBJECT System.Security.AccessControl.FileSystemAccessRule($IdentityReference,"FullControl",”ContainerInherit, ObjectInherit”,"None","Allow")

# Get current Access Rule from Home Folder for User
$HomeFolderACL = Get-acl -Path $homedir

$HomeFolderACL.AddAccessRule($AccessRule)

SET-ACL –path $homedir -AclObject $HomeFolderACL


If (($bldnotification) -eq $true){
#Email Building Staff New account information
  $body ="
    <p> A new account has been created for
    <p> $fullname,$id,$emailadd,$password <br>
    <p> The student will need to log onto a windows computer here at the school to set their password. Their username is also their email address. <br>
    <p> The password should be atleast 8 characters long with a capital and a number, once the student has set their password remind them to NOT SHARE it with other students or staff <br>     
    </P>"
if ($building -eq $hsbuild1){  #Edit this to match your building numbers for high school
Send-MailMessage -SmtpServer $smtpserver -From $mailfrom -To $hsbldcontact -Subject 'New student Account' -Body $body -BodyAsHtml 
} elseif ($building -eq $msbuild1){ #Edit this to match your building number for middle school
Send-MailMessage -SmtpServer $smtpserver -From $mailfrom -To $msbldcontact -Subject 'New student Account' -Body $body -BodyAsHtml 
} elseif ($building -eq $elembuild1){ #Edit this to match your building number for elementary
Send-MailMessage -SmtpServer $smtpserver -From $mailfrom -To $elembldcontact -Subject 'New student Account' -Body $body -BodyAsHtml 
}
start-sleep -Milliseconds 15
continue
}
}