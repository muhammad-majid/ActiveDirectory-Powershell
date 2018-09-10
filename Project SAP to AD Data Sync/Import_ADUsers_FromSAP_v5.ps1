# Embedded CSV File location with NO Try and Catch.
# Get script Start Time (used to measure run time)
$startDTM = (Get-Date)

$path = "\\eisf017\SAP_Automation_Data\"

#Create log date and user disabled date

$logdate = Get-Date -Format ddMMyyyy

#Define CSV and log file location variables
#they have to be on the same location as the script

$csvfile = $path + "Employee_List.csv"
$logfile = $path + "$logdate.log"

"Importing CSV file...." | Out-File $logfile -Append
$records = Import-Csv -path $csvfile 
"CSV file Imported Successfully...." | Out-File $logfile -Append

foreach ($record in $records)
{
    $EmployeeID = $record.EmployeeID
    "EmployeeID read as: " + $EmployeeID | Out-File $logfile -Append

    $Title = $record.JobTitle
    "Title read as: " + $Title | Out-File $logfile -Append

    $Department = $record.Department
    "Department read as: " + $Department | Out-File $logfile -Append

    $ManagerID = $record.ManagerID
    "ManagerID read as: " + $ManagerID | Out-File $logfile -Append
	
	"Trying to verify " + $EmployeeID + " existance in AD" | Out-File $logfile -Append
	
    #Try { $SAMinAD = Get-ADUser -LDAPFilter "(description=$EmployeeID)"}
	#Try { $SAMinAD = Get-ADUser -filter "description -eq $EmployeeID"}
	#Catch { }
    
    #Execute set-aduser below only if $sam is in AD and also is in the excel file, else ignore#
		#If($SAMinAD -ne $null -and $SAMinAD -ne '')
		#{

			#added the 'if clause' to ensure that blank fields in the CSV are ignored.
			#the object names must be the LDAP names. get values using ADSI Edit
	
			#Manager did not accept the -Replace switch. It works with the -manager switch
			IF ($ManagerID -eq '' -or $ManagerID -eq '0') {$ManagerID = $null}
			"Trying to get Manager from AD " | Out-File $logfile -Append
            $ManagerAlias = Get-ADUser -LDAPFilter "(description=$ManagerID)" -Properties SamAccountName | Select-Object SamAccountName
            $ManagerAlias | Out-File $logfile -Append
			"Attempting to modify record of " + $EmployeeID + " in AD" | Out-File $logfile -Append
            Get-ADUser -LDAPFilter "(description=$EmployeeID)" | Set-ADUser -Replace @{Title=$Title;Department=$Department;} -Manager $ManagerAlias
			$EmployeeID + " modified as " + $Title + " under " + $Department + " reporting to " + $ManagerID | Out-File $logfile -Append
			"Fetching updated record from AD..." | Out-File $logfile -Append
			Get-ADUser -LDAPFilter "(description=$EmployeeID)" | Out-File $logfile -Append
            			
			#Set a flag to indicate that the user has been updated on AD.
			#When I export, I will omit all users with thie flag enabled 
}

		#Else

		#{ #Log error for users that are not in Active Directory or with no Logon name in excel file
			#$EmployeeID + " Not modified because it does not exist in AD" | Out-File $logfile -Append
		#}
#}

#Finish
#The lins below calculates how long it takes to run this script
#Get End Time
$endDTM = (Get-Date)

# Echo Time elapsed
"Elapsed Time: $(($endDTM-$startDTM).totalseconds) seconds"
"Elapsed Time: $(($endDTM-$startDTM).totalminutes) minutes"

#SCRIPT ENDS
