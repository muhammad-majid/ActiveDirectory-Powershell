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
# Version 2: 08-09-2018
# - Removed copy user, replicating group memberships instead.
# - Simplified if/else and process flow
# - Improved log output


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
$log      = $path + "\Batch_AD_Creation.log"
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

Function insertTimeStamp { return (Get-Date).ToString('yyyy-MM-dd HH:mm:ss') + ' : ' }

Function Create-Users
{
  (insertTimeStamp) + "Processing started.." | Out-File $log -append
  "--------------------------------------------" | Out-File $log -append
  Import-CSV $newpath | ForEach-Object {
  
	
	$GivenName = $_.GivenName.Trim()
	$LastName = $_.LastName.Trim()
	$CopyUserFrom = $_.CopyUserFrom
	$Password = ConvertTo-SecureString -AsPlainText $_.Password -force
	$Email = $_.Email.Trim()
	$Department = $_.Department.Trim()
	$Title = $_.Title.Trim()
	$Phone = $_.Phone.TrimStart('0')
	$Description = $_.Description.Trim()
	$PasswordNeverExpires = $_.PasswordNeverExpires.ToLower()
	$AccountIsEnabled = $_.AccountIsEnabled.ToLower()
	$Manager = $_.Manager.Trim()
	$TargetOU = $_.TargetOU.Trim()
	$OfficeName = $_.OfficeName.Trim()
	$StreetAddress = $_.StreetAddress.Trim()
	$City = $_.City
	$PostalCode = $_.PostalCode
	$State = $_.State.Trim()
	$Country = $_.Country.Trim()
	$Company = $_.Company.Trim()
	$ProfilePath = $_.ProfilePath.Trim()
	$ScriptPath = $_.ScriptPath.Trim()
	$HomeDirectory = $_.HomeDirectory.Trim()
	$HomeDrive = $_.HomeDrive.Trim()
	$ProxyAddresses = $_.ProxyAddresses.Trim()
	
	$sam = $GivenName.ToLower() + "." + $LastName.ToLower()
	$name = $GivenName + " " + $LastName
    
	Write-Host "Running iteration $i for $GivenName $LastName`r`n"
	Write-Host "Subject username is $($sam)`r`n"
	(insertTimeStamp) + "Running iteration $i for $GivenName $LastName .." | Out-File $log -append
	(insertTimeStamp) + "Subject username is $($sam).." | Out-File $log -append
	
	
	If (($GivenName -eq "") -Or ($LastName -eq ""))
	{
		Write-Host "[ERROR]`t Please provide valid GivenName and LastName. Processing skipped for line $($i)`r`n"
        (insertTimeStamp) + "[ERROR]`t Please provide valid GivenName, LastName. Processing skipped for line $($i)`r`n" | Out-File $log -append
		$i++
		return
	}
	
	
	Try { $exists = Get-ADUser -LDAPFilter "(sAMAccountName=$sam)" }
	Catch { }
	If($exists)
	{
		Write-Host "$($sam) already exists in AD..`r`n"
		Write-Host "skipping to next iteration..`r`n"
		(insertTimeStamp) + "Template user found in AD as $($CopyUserFrom).." | Out-File $log -append
		(insertTimeStamp) + "skipping to next iteration.."
		$i++
		return
	}
			
	
	Write-Host "Creating User`r`n"
	(insertTimeStamp) + "Creating User.." | Out-File $log -append
	New-ADUser $sam
	Write-Host "$($sam) created successfully`r`n"
	(insertTimeStamp) + "$($sam) created successfully.." | Out-File $log -append
	
		
	If (($CopyUserFrom -ne ""))
	{
		Write-Host "Looking for Template $($CopyUserFrom) in AD..`r`n"		
		(insertTimeStamp) + "Looking for Template $($CopyUserFrom) in AD.." | Out-File $log -append
		Try { $exists = Get-ADUser -LDAPFilter "(sAMAccountName=$CopyUserFrom)" }
		Catch { }
		
		If($exists)
		{
			Write-Host "Template user found in AD as $($CopyUserFrom)..`r`n"
			(insertTimeStamp) + "Template user found in AD as $($CopyUserFrom).." | Out-File $log -append
			#call function to replicate user
			#call function to replicate user
			(insertTimeStamp) + "Copying group memberships from $($CopyUserFrom).." | Out-File $log -append
			Get-ADUser -Identity $CopyUserFrom -Properties memberof | Select-Object -ExpandProperty memberof | Add-ADGroupMember -Members $sam
			#call function to replicate user
			#call function to replicate user
			
			Write-Host "$($sam) replicated successfully from template $($CopyUserFrom)..`r`n"
			#Write-Host "proceeding with user modification (if any) from rest of the attributes in .csv`r`n"
			(insertTimeStamp) + "$($sam) replicated successfully from template $($CopyUserFrom).." | Out-File $log -append
			#(insertTimeStamp) + "proceeding with user modification (if any) from rest of the attributes in .csv" | Out-File $log -append
		}
		
		else
		{
			Write-Host "Template $($CopyUserFrom) NOT found in AD..`r`n"		
			(insertTimeStamp) + "Template $($CopyUserFrom) NOT found in AD.." | Out-File $log -append
		}

	}
	
	else
	{
		Write-Host "No Template to copy from in .csv`r`n"		
		(insertTimeStamp) + "No Template to copy from in .csv`r`n.." | Out-File $log -append
	}
	
	
	Write-Host "proceeding with user modification (if any) from rest of the attributes in .csv`r`n"
	(insertTimeStamp) + "proceeding with user modification (if any) from rest of the attributes in .csv" | Out-File $log -append	
	
	$params = @{
		"givenName"=$GivenName
		"sn"=$LastName
		"name"=$name
	}
	
	
	If($Password) { $params.Add("AccountPassword",$Password) }
	If($Email) { $params.Add("mail",$Email) }
	If($Department) { $Params.Add("department",$Department) }
	If($Title) { $Params.Add("title",$Title) }
	If($Phone) { $Params.Add("telephoneNumber",$Phone) }
	If($Description) { $Params.Add("description",$Description) }
	#If($PasswordNeverExpires -eq "true") { $Params.Add("??",$True) }
	If($AccountIsEnabled -eq "true") { { $Params.Add("enabled",$True) }
	If($TargetOU) { $params.Add("Path",$TargetOU)}
	If($OfficeName) { $params.Add("physicalDeliveryOfficeName",$OfficeName) }
	If($StreetAddress) { $params.Add("streetAddress",$StreetAddress) }
	If($City) { $params.Add("l",$City) }
	If($PostalCode) { $params.Add("postalCode",$PostalCode) }
	If($State) { $params.Add("st",$State) }
	
	#If($Country -eq "Australia") {$Country = "NL"} Else { $Country = "EN" }
	If($Country) { $params.Add("co",$Country) }
	
	If($Company) { $params.Add("company",$Company) }
	If($ProfilePath) { $params.Add("profilePath",$ProfilePath) }
	If($ScriptPath) { $params.Add("scriptPath",$ScriptPath) }
	#$HomeDirectory = $_.HomeDirectory.Trim()
	#$HomeDrive = $_.HomeDrive.Trim()
	#$ProxyAddresses = $_.ProxyAddresses.Trim()
	
	Try { $ManagerExists = Get-ADUser -LDAPFilter "(sAMAccountName=$Manager)" }
	Catch { }
		
	If($ManagerExists){ $params.Add("manager",$Manager) }
	
	Try
	{
		Set-ADUser -identity $sam @params
		Write-Host "$($sam) set up successfully`r`n"
		(insertTimeStamp)+"$($sam) set up successfully.." | Out-File $log -append
		
	}
	Catch
	{
		Write-Host "[ERROR]`t Oops, something went wrong: $($Exception.Message)`r`n"
		(insertTimeStamp)+"Oops, something went wrong: $($Exception.Message)" | Out-File $log -append
	}
	
	$newdn = (Get-ADUser $sam).DistinguishedName
    Rename-ADObject -Identity $newdn -NewName ($GivenName + " " + $LastName)
	Write-Host "Renamed to include a space`r`n"
	(insertTimeStamp)+"Renamed to include a space.." | Out-File $log -append
	
	<#
	"Instance"=$template_obj
	
	"DisplayName"=$name
	"GivenName"=$givenname
	"SurName"=$surname
	"AccountPassword"=$password_ss
	"Enabled"=$enabled
	"ChangePasswordAtLogon"=$changepw
	}
	#>
	
		# Set the Enabled and PasswordNeverExpires properties
        
        

        # A check for the country, because those were full names and need
        # to be land codes in order for AD to accept them. I used Netherlands
        # as example
        #If($Country -eq "Netherlands") {$Country = "NL"} Else { $Country = "EN" }
		
        # Replace dots / points (.) in names, because AD will error when a
        # name ends with a dot (and it looks cleaner as well)
        #$replace = $Lastname.Replace(".","")
		
		#Write-Host "Applying naming convention to get username to be created..`r`n"
		#(insertTimeStamp) + "Applying naming convention to get username to be created.." | Out-File $log -append
				
        # Create sAMAccountName according to this 'naming convention':
        # <FirstLetterInitials><FirstFourLettersLastName> for example
        #$sam = $Initials.substring(0,1).ToLower() + $lastname.ToLower()
        <#
        Try   { $exists = Get-ADUser -LDAPFilter "(sAMAccountName=$sam)" }
        Catch {	}
		
		If($exists)
		{
			Write-Host "[SKIP]`t User $($sam) ($($GivenName) $($LastName)) already exists or returned an error!`r`n"
			Write-Host "No changes applied - skipping to next iteration`r`n"
			(insertTimeStamp) + "[SKIP]`t User $($sam) ($($GivenName) $($LastName)) already exists or returned an error!" | Out-File $log -append
			(insertTimeStamp) + "No changes applied - skipping to next iteration.." | Out-File $log -append
			$i++
			return
		}
		
          # Set all variables according to the table names in the Excel
          # sheet / import CSV. The names can differ in every project, but
          # if the names change, make sure to change it below as well.
          #$setpass = ConvertTo-SecureString -AsPlainText $Password -force

          Try
          {
            Write-Host "[INFO]`t Creating user : $($sam)"
            (insertTimeStamp) + "[INFO]`t Creating user : $($sam)" | Out-File $log -append
            New-ADUser $sam -GivenName $GivenName -Initials $Initials `
            -Surname $LastName -DisplayName ($GivenName + " " + $LastName) `
            -Office $OfficeName -Description $Description -EmailAddress $Mail `
            -StreetAddress $StreetAddress -City $City -State $State `
            -PostalCode $PostalCode -Country $Country -UserPrincipalName ($sam + "@" + $dnsroot) `
            -Company $Company -Department $Department -EmployeeID $EmployeeID `
            -Title $Title -OfficePhone $Phone -AccountPassword $setpass -Manager $Manager `
            -profilePath $ProfilePath -scriptPath $ScriptPath -homeDirectory $HomeDirectory `
            -homeDrive $homeDrive -Enabled $enabled -PasswordNeverExpires $expires
            Write-Host "[INFO]`t Created new user : $($sam)"
            (insertTimeStamp) + "[INFO]`t Created new user : $($sam)" | Out-File $log -append

            $dn = (Get-ADUser $sam).DistinguishedName
            # Set an ExtensionAttribute
            If ($ExtensionAttribute1 -ne "" -And $ExtensionAttribute1 -ne $Null)
            {
              $ext = [ADSI]"LDAP://$dn"
              $ext.Put("extensionAttribute1", $ExtensionAttribute1)
              Try   { $ext.SetInfo() }
              Catch { Write-Host "[ERROR]`t Couldn't set the Extension Attribute : $($Exception.Message)" }
            }

            # Set ProxyAdresses
            If ($proxyAddresses -ne "")
            {
              Try { $dn | Set-ADUser -Add @{proxyAddresses = ($ProxyAddresses -split ";")} -ErrorAction Stop }
              Catch { Write-Host "[ERROR]`t Couldn't set the ProxyAddresses Attributes : $($Exception.Message)" }
            }

            # Move the user to the OU ($location) you set above. If you don't
            # want to move the user(s) and just create them in the global Users
            # OU, comment the string below
            If ([adsi]::Exists("LDAP://$($location)"))
            {
              Move-ADObject -Identity $dn -TargetPath $location
              Write-Host "[INFO]`t User $sam moved to target OU : $($location)"
              (insertTimeStamp) + "[INFO]`t User $sam moved to target OU : $($location)" | Out-File $log -append
            }
            Else
            {
              Write-Host "[ERROR]`t Targeted OU couldn't be found. Newly created user wasn't moved!"
              (insertTimeStamp) + "[ERROR]`t Targeted OU couldn't be found. Newly created user wasn't moved!" | Out-File $log -append
            }

            # Rename the object to a good looking name (otherwise you see
            # the 'ugly' shortened sAMAccountNames as a name in AD. This
            # can't be set right away (as sAMAccountName) due to the 20
            # character restriction
            $newdn = (Get-ADUser $sam).DistinguishedName
            Rename-ADObject -Identity $newdn -NewName ($GivenName + " " + $LastName)
            Write-Host "[INFO]`t Renamed $($sam) to $($GivenName) $($LastName)`r`n"
            (insertTimeStamp) + "[INFO]`t Renamed $($sam) to $($GivenName) $($LastName)`r`n" | Out-File $log -append
			
		  }
          Catch
          {
            Write-Host "[ERROR]`t Oops, something went wrong: $($Exception.Message)`r`n"
          }
		  
		 #>
		Write-Host "Moving to next iteration`r`n"
		(insertTimeStamp) + "Moving to next iteration`r`n" | Out-File $log -append
		$i++
    }
  "--------------------------------------------" + "`r`n" | Out-File $log -append
 }
}
Write-Host "STARTED SCRIPT`r`n"
Start-Commands
Write-Host "STOPPED SCRIPT"
