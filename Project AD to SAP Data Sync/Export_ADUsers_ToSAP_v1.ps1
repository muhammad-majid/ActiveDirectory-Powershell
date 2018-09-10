#Takes 4 parameters.
#1 Folder Path
#2 AD Delta File name.
#3 From email address.
#4 To email address.
#Generates ADDeltaFileName to be fed to SAP for updates.  

[CmdletBinding()]
Param
(
   
    [Parameter(Mandatory=$True,Position=1)]
    [string]$FolderPath,

	[Parameter(Mandatory=$True,Position=2)]
    [string]$ADDeltaFileName,
	
    [Parameter(Mandatory=$True,Position=3)]
    [string]$MailFrom,

    [Parameter(Mandatory=$True,Position=4)]
    [string]$MailTo
)

#Create a timestamp
$myTimeStamp = Get-Date -Format yyyyMMddHHmm

$smtpsettings = @{
  To =  $MailTo
  From = $MailFrom
  Subject = "AD to SAP Employee Sync Report " + $myTimeStamp
  SmtpServer = "mail.emiratessteel.com"
  }
  
 $smtpsettings2 = @{
  To =  "muhammad.majid@emiratessteel.com"
  From = $MailFrom
  Subject = "AD to SAP Employee Sync Report " + $myTimeStamp
  SmtpServer = "mail.emiratessteel.com"
  }

$logfile = $FolderPath + "\" + "AD to SAP Data Sync Report $myTimeStamp.log"
$MailBodyContent = "AD to SAP Employee Sync Report... <br/><br/><br/>"  

$FinalCsv = "$FolderPath\$ADDeltaFileName"
(Get-Date -Format yyyyMMddHHmm) + " : " + "Deleting older $FinalCsv.."  | Out-File $logfile -Append
Remove-Item $FinalCsv -Force

(Get-Date -Format yyyyMMddHHmm) + " : " + "Searching for .Zip file at $FolderPath.."  | Out-File $logfile -Append

$zipFile = Get-ChildItem -Path "$FolderPath\*.*" -File -Include *.zip | Sort-Object LastAccessTime -Descending | Select-Object -First 1		#Get the latest .zip file
If ($zipFile.count -eq 0)
{
	(Get-Date -Format yyyyMMddHHmm) + " : " + "---------------------------------------------------------------------------------------------------------------------------------"  | Out-File $logfile -Append
	(Get-Date -Format yyyyMMddHHmm) + " : " + "No zip file found..."  | Out-File $logfile -Append
	(Get-Date -Format yyyyMMddHHmm) + " : " + "---------------------------------------------------------------------------------------------------------------------------------"  | Out-File $logfile -Append
	(Get-Date -Format yyyyMMddHHmm) + " : " + "Exiting script.."  | Out-File $logfile -Append
	Send-MailMessage @smtpsettings -Body "No zip file found...<br/> NO .CSV File Generated...<br/> Exiting Script..." -BodyAsHtml -Encoding ([System.Text.Encoding]::UTF8)
	Send-MailMessage @smtpsettings2 -Body "No zip file found...<br/> NO .CSV File Generated...<br/> Exiting Script..." -BodyAsHtml -Encoding ([System.Text.Encoding]::UTF8)
	exit
}

(Get-Date -Format yyyyMMddHHmm) + " : " + "Zip file found as: " + $zipFile | Out-File $logfile -Append

$unZippedDir = "$FolderPath\unZipped-$myTimeStamp"
If(-NOT (Test-Path -Path $unZippedDir)){ New-Item -ItemType directory -Path $unZippedDir }

(Get-Date -Format yyyyMMddHHmm) + " : " + "Created unZipped Folder as: " + $unZippedDir  | Out-File $logfile -Append
(Get-Date -Format yyyyMMddHHmm) + " : " + "Expanding $zipFile.."  | Out-File $logfile -Append

Expand-Archive $zipFile -DestinationPath $unZippedDir

(Get-Date -Format yyyyMMddHHmm) + " : " + "Expanded Successfully at $unZippedDir.."  | Out-File $logfile -Append
(Get-Date -Format yyyyMMddHHmm) + " : " + "Renaming $ZipFile"  | Out-File $logfile -Append

#Remove-Item $ZipFile
Rename-Item $ZipFile "$ZipFile.bkp"

(Get-Date -Format yyyyMMddHHmm) + " : " + "$ZipFile renamed to .bkp .."  | Out-File $logfile -Append
(Get-Date -Format yyyyMMddHHmm) + " : " + "Searching for .csv file at $unZippedDir.."  | Out-File $logfile -Append

$myCsvFile = Get-ChildItem -Path "$unZippedDir\*.*" -File -Include *.csv

(Get-Date -Format yyyyMMddHHmm) + " : " + "CSV file found as: " + $myCsvFile | Out-File $logfile -Append
(Get-Date -Format yyyyMMddHHmm) + " : " + "Getting contents of .csv file.." | Out-File $logfile -Append
(Get-Date -Format yyyyMMddHHmm) + " : " + "Checking for updates..."  | Out-File $logfile -Append

#check if there are any updates since last run.

#$noDataAvailable = "No data available to generate the"
$noDataAvailable = "Number of Records : 0"

if (Select-String -Path $myCsvFile -Pattern $noDataAvailable -Quiet)
{
	(Get-Date -Format yyyyMMddHHmm) + " : " + "---------------------------------------------------------------------------------------------------------------------------------"  | Out-File $logfile -Append
	(Get-Date -Format yyyyMMddHHmm) + " : " + "$myCsvFile has nothing to update.."  | Out-File $logfile -Append
	(Get-Date -Format yyyyMMddHHmm) + " : " + "---------------------------------------------------------------------------------------------------------------------------------"  | Out-File $logfile -Append
	(Get-Date -Format yyyyMMddHHmm) + " : " + "Deleting folder $unZippedDir recursively.."  | Out-File $logfile -Append
	Remove-Item -Recurse -Force $unZippedDir
	(Get-Date -Format yyyyMMddHHmm) + " : " + "Deletion complete, exiting script.."  | Out-File $logfile -Append
	Send-MailMessage @smtpsettings -Body "Nothing to Update...<br/> NO .CSV File Generated...<br/> Exiting Script..." -BodyAsHtml -Encoding ([System.Text.Encoding]::UTF8)
	Send-MailMessage @smtpsettings2 -Body "Nothing to Update...<br/> NO .CSV File Generated...<br/> Exiting Script..." -BodyAsHtml -Encoding ([System.Text.Encoding]::UTF8)
	exit
}


(Get-Date -Format yyyyMMddHHmm) + " : " + "Updates found.."  | Out-File $logfile -Append
$a = Get-Content $myCsvFile
(Get-Date -Format yyyyMMddHHmm) + " : " + "Skipping top 8 lines, and saving as temp.csv" | Out-File $logfile -Append

$a[8..($a.length-1)] > "$unZippedDir\temp.csv"

(Get-Date -Format yyyyMMddHHmm) + " : " + "$unZippedDir\temp.csv saved successfully.." | Out-File $logfile -Append
(Get-Date -Format yyyyMMddHHmm) + " : " + "Modifying $unZippedDir\temp.csv to sort and remove duplicate.."  | Out-File $logfile -Append

$a = import-csv "$unZippedDir\temp.csv"
$a | Sort-Object | Group-Object "User Name" | Foreach-Object {$_.Group | Select-Object -Last 1} | export-csv "$unZippedDir\temp2.csv" -NoTypeInformation

(Get-Date -Format yyyyMMddHHmm) + " : " + "Saved as $unZippedDir\temp2.csv.."  | Out-File $logfile -Append
(Get-Date -Format yyyyMMddHHmm) + " : " + "Looping through $unZippedDir\temp2.csv to fetch user information from Active Directory.."  | Out-File $logfile -Append

#Import-Csv $unZippedDir\temp2.csv | ForEach-Object {Get-ADUser $_."User Name" -Properties employeeID, emailAddress, telephoneNumber, mobile | select employeeID, emailAddress, telephoneNumber, mobile} | Export-Csv $FinalCsv -NoTypeInformation

$delimiter = '~'
#Import-Csv $unZippedDir\temp2.csv | ForEach-Object {Get-ADUser $_."User Name" -Properties employeeID, emailAddress, telephoneNumber, mobile | select employeeID, emailAddress, telephoneNumber, mobile} | ConvertTo-Csv -Delimiter $delimiter -NoTypeInformation | foreach { $_ -replace '^"','' -replace "`"$delimiter`"",$delimiter -replace '"$','' } | Out-File $FinalCsv -fo -en ascii
Import-Csv $unZippedDir\temp2.csv | ForEach-Object {Get-ADUser $_."User Name" -Properties employeeID, emailAddress, telephoneNumber, mobile | select employeeID, emailAddress, telephoneNumber, mobile} | ConvertTo-Csv -Delimiter $delimiter -NoTypeInformation | % {$_ -replace '"',''} | Out-File $FinalCsv -fo -en ascii

#Get-Process | ConvertTo-Csv -Delimiter $delimiter -NoTypeInformation | foreach { $_ -replace '^"','' -replace "`"$delimiter`"",$delimiter -replace '"$','' }



$Records =import-csv $FinalCsv -Delimiter '~'
(Get-Date -Format yyyyMMddHHmm) + " : " + "---------------------------------------------------------------------------------------------------------------------------------"  | Out-File $logfile -Append
foreach ($Record in $Records)
{
	(Get-Date -Format yyyyMMddHHmm) + " : " + $Record | Out-File $logfile -Append
	$MailBodyContent += "$Record<br/>"
	#$EmployeeID = ($Record."User Name")
	#$Message = $Record.Message
	#get adObject info
}
(Get-Date -Format yyyyMMddHHmm) + " : " + "---------------------------------------------------------------------------------------------------------------------------------"  | Out-File $logfile -Append
(Get-Date -Format yyyyMMddHHmm) + " : " + "User informating successfully updated at $FinalCsv"  | Out-File $logfile -Append
(Get-Date -Format yyyyMMddHHmm) + " : " + "Deleting folder $unZippedDir recursively.."  | Out-File $logfile -Append
Remove-Item -Recurse -Force $unZippedDir
(Get-Date -Format yyyyMMddHHmm) + " : " + "Deletion complete, exiting script.."  | Out-File $logfile -Append
Send-MailMessage @smtpsettings -Body  $MailBodyContent -BodyAsHtml -Encoding ([System.Text.Encoding]::UTF8) -Attachments $logfile
Send-MailMessage @smtpsettings2 -Body  $MailBodyContent -BodyAsHtml -Encoding ([System.Text.Encoding]::UTF8) -Attachments $logfile