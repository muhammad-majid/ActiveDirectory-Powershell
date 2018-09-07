###########################################################
# AUTHOR  : Muhammad Majid
# DATE    : 26-04-2012
# EDIT    : 07-09-2018
# COMMENT : This script creates new Active Directory users,
#           including different kind of properties, based
#           on an input_create_ad_users.csv.
# VERSION : 1.4
###########################################################
##
# CHANGELOG
# Version 1.2: 15-04-2014 - Changed the code for better
# - Added better Error Handling and Reporting.
# - Changed input file with more logical headers.
# - Added functionality for account Enabled,
#   PasswordNeverExpires, ProfilePath, ScriptPath,
#   HomeDirectory and HomeDrive
# - Added the option to move every user to a different OU.
# Version 1.3: 08-07-2014
# - Added functionality for ProxyAddresses
# Version 1.4: 07-09-2018
# - Added functionality to copy user from existing user


# ERROR REPORTING ALL
Set-StrictMode -Version latest

#----------------------------------------------------------
# LOAD ASSEMBLIES AND MODULES
#----------------------------------------------------------
Try
{
  Import-Module ActiveDirectory -ErrorAction Stop
}
Catch
{
  Write-Host "[ERROR]`t ActiveDirectory Module couldn't be loaded. Script will stop!"
  Exit 1
}

#----------------------------------------------------------
#STATIC VARIABLES
#----------------------------------------------------------
$path     = Split-Path -parent $MyInvocation.MyCommand.Definition
$newpath  = $path + "\import_create_ad_users.csv"
$log      = $path + "\create_ad_users.log"
$date     = Get-Date
$addn     = (Get-ADDomain).DistinguishedName
$dnsroot  = (Get-ADDomain).DNSRoot
$i        = 1

#----------------------------------------------------------
#START FUNCTIONS
#----------------------------------------------------------
Function Start-Commands
{
  Create-Users
}

Function Create-Users
{
  "Processing started (on " + $date + "): " | Out-File $log -append
  "--------------------------------------------" | Out-File $log -append
  Import-CSV $newpath | ForEach-Object {
  
	Write-Host "Running iteration $i`r`n"
	"Running iteration $i.." | Out-File $log -append
    If (($_.Implement.ToLower()) -ne "yes")
	{
		Write-Host "[SKIP]`t User ($($_.GivenName) $($_.LastName)) will be skipped for processing!`r`n"
		"[SKIP]`t User ($($_.GivenName) $($_.LastName)) will be skipped for processing!" | Out-File $log -append
		$i++
		continue
	}

	Write-Host "Implement is set to yes`r`n"
	"Implement is set to yes" | Out-File $log -append
	
	If (($_.GivenName -eq "") -Or ($_.LastName -eq ""))
	{
		Write-Host "[ERROR]`t Please provide valid GivenName and LastName. Processing skipped for line $($i)`r`n"
        "[ERROR]`t Please provide valid GivenName, LastName. Processing skipped for line $($i)`r`n" | Out-File $log -append
		$i++
		continue
	}

	Write-Host "Given name and LastName are not empty..`r`n"
	"Given name and LastName are not empty.." | Out-File $log -append
	
	If (($_.CopyFrom -ne ""))
	{
		Write-Host "CopyFrom is not empty`r`n"
		"CopyFrom is not empty.." | Out-File $log -append
		$userInstance = Get-ADUser -Identity $_.CopyFrom
		$replace = $_.Lastname.Replace(".","")
		$sam = $_.GivenName.ToLower() + "." + $_.LastName.ToLower()
        
		Try   { $exists = Get-ADUser -LDAPFilter "(sAMAccountName=$sam)" }
        Catch
		{
			Write-Host "[SKIP]`t CopyFrom User $($sam) ($($_.GivenName) $($_.LastName)) already exists or returned an error!`r`n"
			"[SKIP]`t CopyFrom User $($sam) ($($_.GivenName) $($_.LastName)) already exists or returned an error!" | Out-File $log -append
		}
	
		If(!$exists)
        {
			Write-Host "User $($sam) ($($_.GivenName) $($_.LastName)) doesn't already exist`r`n"
			"User $($sam) ($($_.GivenName) $($_.LastName)) doesn't already exist.." | Out-File $log -append
          # Set all variables according to the table names in the Excel
          # sheet / import CSV. The names can differ in every project, but
          # if the names change, make sure to change it below as well.
          $setpass = ConvertTo-SecureString -AsPlainText $_.Password -force

          Try
          {
            Write-Host "[INFO]`t Creating user : $($sam)"
            "[INFO]`t Creating user : $($sam)" | Out-File $log -append
            
			$Copy = Get-ADUser $_.CopyFrom
			
			Write-Host "Running create user command`r`n"
            "Running create user command.." | Out-File $log -append
			
			New-ADUser -SAMAccountName $sam -Instance $Copy -DisplayName $_.GivenName + " " + $_.LastName
			
			Write-Host "Completed create user command`r`n"
            "Completed create user command.." | Out-File $log -append
			
			If (($_.Enabled.ToLower()) -eq "true") { $enabled = $True } Else { $enabled = $False }
			If (($_.PasswordNeverExpires.ToLower()) -eq "true") { $expires = $True } Else { $expires = $False }
			
			#Set-ADUser -Identity $sam.SamAccountName -GivenName $_.GivenName -Surname $_.LastName -Enabled $enabled -PasswordNeverExpires $expires
			Set-ADUser -Identity "$($sam)" -GivenName $_.GivenName -Surname $_.LastName -Enabled $enabled -PasswordNeverExpires $expires
			
            Write-Host "[INFO]`t Created new user : $($sam)"
            "[INFO]`t Created new user : $($sam)" | Out-File $log -append

            # Rename the object to a good looking name (otherwise you see
            # the 'ugly' shortened sAMAccountNames as a name in AD. This
            # can't be set right away (as sAMAccountName) due to the 20
            # character restriction
            $newdn = (Get-ADUser $sam).DistinguishedName
            Rename-ADObject -Identity $newdn -NewName ($_.GivenName + " " + $_.LastName)
            Write-Host "[INFO]`t Renamed $($sam) to $($_.GivenName) $($_.LastName)`r`n"
            "[INFO]`t Renamed $($sam) to $($_.GivenName) $($_.LastName)`r`n" | Out-File $log -append
          }
          Catch
          {
            Write-Host "[ERROR]`t Oops, something went wrong when attempting to create user $($sam) : $($_.Exception.Message)`r`n"
			"[ERROR]`t Oops, something went wrong when attempting to create user $($sam) : $($_.Exception.Message)" | Out-File $log -append
          }
		  
		Write-Host "Moving to next iteration`r`n"
		"Moving to next iteration`r`n" | Out-File $log -append
		$i++
		continue
        }
	}
    
	Else	#CopyFrom is empty, create a fresh new user from rest of the provided details in excel.
     {
		Write-Host "CopyFrom is empty, Createing new user from details in .csv file`r`n"
		"CopyFrom is empty, Createing new user from details in .csv file.." | Out-File $log -append
		
        # Set the target OU
        #$location = $_.TargetOU + ",$($addn)"
		$location = $_.TargetOU

        # Set the Enabled and PasswordNeverExpires properties
        If (($_.Enabled.ToLower()) -eq "true") { $enabled = $True } Else { $enabled = $False }
        If (($_.PasswordNeverExpires.ToLower()) -eq "true") { $expires = $True } Else { $expires = $False }

        # A check for the country, because those were full names and need
        # to be land codes in order for AD to accept them. I used Netherlands
        # as example
        If($_.Country -eq "Netherlands") {$_.Country = "NL"} Else { $_.Country = "EN" }
		
        # Replace dots / points (.) in names, because AD will error when a
        # name ends with a dot (and it looks cleaner as well)
        $replace = $_.Lastname.Replace(".","")
		
		Write-Host "Applying naming convention to get username to be created..`r`n"
		"Applying naming convention to get username to be created.." | Out-File $log -append
				
        # Create sAMAccountName according to this 'naming convention':
        # <FirstLetterInitials><FirstFourLettersLastName> for example
        #$sam = $_.Initials.substring(0,1).ToLower() + $lastname.ToLower()
        $sam = $_.GivenName.ToLower() + "." + $_.LastName.ToLower()
        Try   { $exists = Get-ADUser -LDAPFilter "(sAMAccountName=$sam)" }
        Catch {	}
		
		If($exists)
		{
			Write-Host "[SKIP]`t User $($sam) ($($_.GivenName) $($_.LastName)) already exists or returned an error!`r`n"
			"[SKIP]`t User $($sam) ($($_.GivenName) $($_.LastName)) already exists or returned an error!" | Out-File $log -append
		}
		
        If(!$exists)
        {
			Write-Host "[SKIP]`t User $($sam) ($($_.GivenName) $($_.LastName)) not found among existing users`r`n"
			 "[SKIP]`t User $($sam) ($($_.GivenName) $($_.LastName)) not found among existing users.." | Out-File $log -append
          # Set all variables according to the table names in the Excel
          # sheet / import CSV. The names can differ in every project, but
          # if the names change, make sure to change it below as well.
          $setpass = ConvertTo-SecureString -AsPlainText $_.Password -force

          Try
          {
            Write-Host "[INFO]`t Creating user : $($sam)"
            "[INFO]`t Creating user : $($sam)" | Out-File $log -append
            New-ADUser $sam -GivenName $_.GivenName -Initials $_.Initials `
            -Surname $_.LastName -DisplayName ($_.GivenName + " " + $_.LastName) `
            -Office $_.OfficeName -Description $_.Description -EmailAddress $_.Mail `
            -StreetAddress $_.StreetAddress -City $_.City -State $_.State `
            -PostalCode $_.PostalCode -Country $_.Country -UserPrincipalName ($sam + "@" + $dnsroot) `
            -Company $_.Company -Department $_.Department -EmployeeID $_.EmployeeID `
            -Title $_.Title -OfficePhone $_.Phone -AccountPassword $setpass -Manager $_.Manager `
            -profilePath $_.ProfilePath -scriptPath $_.ScriptPath -homeDirectory $_.HomeDirectory `
            -homeDrive $_.homeDrive -Enabled $enabled -PasswordNeverExpires $expires
            Write-Host "[INFO]`t Created new user : $($sam)"
            "[INFO]`t Created new user : $($sam)" | Out-File $log -append

            $dn = (Get-ADUser $sam).DistinguishedName
            # Set an ExtensionAttribute
            If ($_.ExtensionAttribute1 -ne "" -And $_.ExtensionAttribute1 -ne $Null)
            {
              $ext = [ADSI]"LDAP://$dn"
              $ext.Put("extensionAttribute1", $_.ExtensionAttribute1)
              Try   { $ext.SetInfo() }
              Catch { Write-Host "[ERROR]`t Couldn't set the Extension Attribute : $($_.Exception.Message)" }
            }

            # Set ProxyAdresses
            If ($_.proxyAddresses -ne "")
            {
              Try { $dn | Set-ADUser -Add @{proxyAddresses = ($_.ProxyAddresses -split ";")} -ErrorAction Stop }
              Catch { Write-Host "[ERROR]`t Couldn't set the ProxyAddresses Attributes : $($_.Exception.Message)" }
            }

            # Move the user to the OU ($location) you set above. If you don't
            # want to move the user(s) and just create them in the global Users
            # OU, comment the string below
            If ([adsi]::Exists("LDAP://$($location)"))
            {
              Move-ADObject -Identity $dn -TargetPath $location
              Write-Host "[INFO]`t User $sam moved to target OU : $($location)"
              "[INFO]`t User $sam moved to target OU : $($location)" | Out-File $log -append
            }
            Else
            {
              Write-Host "[ERROR]`t Targeted OU couldn't be found. Newly created user wasn't moved!"
              "[ERROR]`t Targeted OU couldn't be found. Newly created user wasn't moved!" | Out-File $log -append
            }

            # Rename the object to a good looking name (otherwise you see
            # the 'ugly' shortened sAMAccountNames as a name in AD. This
            # can't be set right away (as sAMAccountName) due to the 20
            # character restriction
            $newdn = (Get-ADUser $sam).DistinguishedName
            Rename-ADObject -Identity $newdn -NewName ($_.GivenName + " " + $_.LastName)
            Write-Host "[INFO]`t Renamed $($sam) to $($_.GivenName) $($_.LastName)`r`n"
            "[INFO]`t Renamed $($sam) to $($_.GivenName) $($_.LastName)`r`n" | Out-File $log -append
          }
          Catch
          {
            Write-Host "[ERROR]`t Oops, something went wrong: $($_.Exception.Message)`r`n"
          }
        }
		Write-Host "Moving to next iteration`r`n"
		"Moving to next iteration`r`n" | Out-File $log -append
		$i++
    }
  }
  "--------------------------------------------" + "`r`n" | Out-File $log -append
}

Write-Host "STARTED SCRIPT`r`n"
Start-Commands
Write-Host "STOPPED SCRIPT"
