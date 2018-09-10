###########################################################
# AUTHOR  : Victor Ashiedu
# WEBSITE : iTechguides.com
# BLOG    : iTechguides.com/blog-2/
# CREATED : 08-08-2014 
# UPDATED : 19-09-2014
# VERSION : 1.3
# COMMENT : Sometimes when users are created in Active Directory, some attributes are left blank. 
#           This PowerShel Script updates blank user attributes like email address, physical address
#           Manager and more using a CSV file as imput. 
#           If you find this script useful, please take time to rate it via the link below: 
#           http://gallery.technet.microsoft.com/PowerShell-script-to-376e9462
###########################################################
#SCRIPT BEGINS
#The line below measures the lenght of time it takes to
#execute this script

# Get script Start Time (used to measure run time)
$startDTM = (Get-Date)

#Define location of my script variable
#the -parent switch returns one directory lower from directory defined. 
#below will return up to ImportADUsers folder 
#and since my files are located here it will find it.
#It failes withpout appending "*.*" at the end
#This file is required to update fields for existing users
#Modify this script to create new users in UnifiedGov domain


$path = "C:\ShellWorkingFolder\"

#Create log date and user disabled date

$logdate = Get-Date -Format ddmmyyyy
$userdisableddate = Get-Date

#Define CSV and log file location variables
#they have to be on the same location as the script

$csvfile = $path + "Employee List.csv"
$logfile = $path + "$logdate.LogFile.txt"
$scriptrunrime = $path + "$logdate.ScriptTime.txt"


#define searchbase variable

#$SearchBase = "OU=Deleted IDs APR-JUN,OU=Non-Phone Directory Users,DC=kingston,DC=gov,DC=uk"

#Import CSV file and update users in the OU with details from the file
#Create the function script to update the users

Function Update-ADUsers
{
	Import-Csv -path $csvfile | 
	ForEach-Object
	{ 
		$sam = $_.Description
		$Title = $_.Title
		$Department = $_.Department
		$Manager = $_.Manager
 
		#Included the If clause below to ignore execution if the $Manager variable
		#from the csv is blank. Avoids throwing errors and saves execution time
		#Used different possible displaynames to search for a managername

		##First check whether $sam exisits in AD

		Try { $SAMinAD = Get-ADUser -LDAPFilter "(sAMAccountName=$sam)"} 
		Catch { }

		#Execute set-aduser below only if $sam is in AD and also is in the excel file, else ignore#
		If($SAMinAD -ne $null -and $sam -ne '' -and $sam -ne '0')
		{

			#added the 'if clause' to ensure that blank fields in the CSV are ignored.
			#the object names must be the LDAP names. get values using ADSI Edit

			Set-ADUser -Identity $sam -Replace @{Title=$Title}
			Set-ADUser -Identity $sam -Replace @{Department=$Department}
			#Manager did not accept the -Replace switch. It works with the -manager switch
			IF ($Manager -eq '') {$Manager = $null}
			IF ($Manager -eq '0') {$Manager = $null}
			Set-ADUser -Identity $sam -Manager 'Get-ADUser -LDAPFilter "(sAMAccountName=$Manager)"'

			#Set a flag to indicate that the user has been updated on AD.
			#When I export, I will omit all users with thie flag enabled 
		}

		Else

		{ #Log error for users that are not in Active Directory or with no Logon name in excel file
			$DisplayName + " Not modified because it does not exist in AD or LogOn name field is empty on excel file" | Out-File $logfile -Append
		}

	}
}
   
# Run the function script 
Update-ADUsers
#Finish
#The lins below calculates how long
#it takes to run this script
# Get End Time
$endDTM = (Get-Date)

# Echo Time elapsed
"Elapsed Time: $(($endDTM-$startDTM).totalseconds) seconds"
"Elapsed Time: $(($endDTM-$startDTM).totalminutes) minutes"

#send the information to a text file

"$(($endDTM-$startDTM).totalseconds) seconds" > $scriptrunrime

#Append the minutes value to the text file

Add-Content -path $scriptrunrime "$(($endDTM-$startDTM).totalminutes) minutes"
#SCRIPT ENDS