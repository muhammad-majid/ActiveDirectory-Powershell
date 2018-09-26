
# ERROR REPORTING ALL
Set-StrictMode -Version latest

#----------------------------------------------------------
# LOAD ASSEMBLIES AND MODULES
#----------------------------------------------------------
Try { Import-Module ActiveDirectory -ErrorAction Stop }
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

Function insertTimeStamp { return (Get-Date).ToString('yyyy-MM-dd HH:mm:ss') + ' : ' }

Write-Host "STARTED SCRIPT`r`n"
(insertTimeStamp) + "Processing started.." | Out-File $log -append
"--------------------------------------------" | Out-File $log -append

Import-CSV $newpath | ForEach-Object {

	$GivenName = $_.GivenName.Trim()
	$LastName = $_.LastName.Trim()
	$CopyUserFrom = $_.CopyUserFrom.Trim()
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
	$Name = $GivenName+" "+$LastName
	$userPrincipalName = $sam+"@"+$dnsroot

	#poperties that will be imported from template user ($CopyUserFrom) when necessary.
	#$propertiesImported = @("department", "title", "description", "distinguishedName", "physicalDeliveryOfficeName", "streetAddress", "l", "postalCode", "st", "co", "company", "manager")
	$propertiesImported = @("l", "company", "c", "department", "description", "wWWHomePage", "manager", "physicalDeliveryOfficeName", "o", "postOfficeBox", "postalCode", "st", "streetAddress", "title")
	#as well as group memberhips at the time of replication.
	
	Write-Host "Running iteration $i for $GivenName $LastName`r`n"
	Write-Host "Subject username is $($sam)`r`n"
	(insertTimeStamp) + "Running iteration $i for $GivenName $LastName .." | Out-File $log -append
	(insertTimeStamp) + "Subject username is $($sam).." | Out-File $log -append

	#if first name, last name or password are missing, skip move to next iteration.
	If ($GivenName -eq '' -Or $GivenName -eq $null -Or $LastName -eq '' -Or $LastName -eq $null -Or $_.Password -eq '' -Or $_.Password -eq $null)
	{
		Write-Host "[ERROR]`t Please provide valid GivenName, LastName and Password. Processing skipped for line $($i)`r`n"
		(insertTimeStamp) + "[ERROR]`t Please provide valid GivenName, LastName and Password. Processing skipped for line $($i)`r`n" | Out-File $log -append
		$i++
		return
	}


	Try { $exists = Get-ADUser -LDAPFilter "(sAMAccountName=$sam)" }		
	Catch { }
	If($exists -ne $null -and $exists -ne '')		#if account already exists, skip and move to next iteration.
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
	New-ADuser -sAMAccountName $sam -Name $Name -GivenName $GivenName -Surname $LastName -DisplayName $Name -AccountPassword $Password
	Write-Host "$($sam) created successfully`r`n"
	(insertTimeStamp) + "$($sam) created successfully.." | Out-File $log -append
	$propertiesToExport = @{}
	<#$propertiesToExport = @{
			"givenName"=$GivenName
			"surname"=$LastName
			#"name"=$Name
			"displayName"=$Name
			"UserPrincipalName"=$userPrincipalName
			#"AccountPassword"=$Password
			}
	#>
	
	If ($CopyUserFrom -ne $null -and $CopyUserFrom -ne '')				#if csv has template for the new user
	{
		Write-Host "Looking for Template $($CopyUserFrom) in AD..`r`n"		
		(insertTimeStamp) + "Looking for Template $($CopyUserFrom) in AD.." | Out-File $log -append
		Try { $exists = Get-ADUser -LDAPFilter "(sAMAccountName=$CopyUserFrom)" }
		Catch { }
		
		If($exists)					#if the template exists, copy from template first, then continue from csv entries.
		{
		
			Write-Host "Template user found in AD as $($CopyUserFrom)..`r`nAttempting to copy...`r`n"
			(insertTimeStamp) + "Template user found in AD as $($CopyUserFrom).." | Out-File $log -append
			(insertTimeStamp) + "Attempting to copy.." | Out-File $log -append
			
			$exists = Get-ADUser -Identity $CopyUserFrom -Properties $propertiesImported
			
			If($exists.l -ne '' -and $exists.l -ne $null) { $propertiesToExport.Add("l",$exists.l) }
			If($exists.company -ne '' -and $exists.company -ne $null) { $propertiesToExport.Add("company",$exists.company) }
			If($exists.c -ne '' -and $exists.c -ne $null) { $propertiesToExport.Add("Country",$exists.c) }							#problematic
			If($exists.department -ne '' -and $exists.department -ne $null) { $propertiesToExport.Add("department",$exists.department) }
			If($exists.description -ne '' -and $exists.description -ne $null) { $propertiesToExport.Add("description",$exists.description) }
			If($exists.wWWHomePage -ne '' -and $exists.wWWHomePage -ne $null) { $propertiesToExport.Add("HomePage",$exists.wWWHomePage) }
			If($exists.physicalDeliveryOfficeName -ne '' -and $exists.physicalDeliveryOfficeName -ne $null) { $propertiesToExport.Add("office",$exists.physicalDeliveryOfficeName) }
			If($exists.o -ne '' -and $exists.o -ne $null) { $propertiesToExport.Add("o",$exists.o) }
			If($exists.postOfficeBox -ne '' -and $exists.postOfficeBox -ne $null) { $propertiesToExport.Add("postOfficeBox",$exists.postOfficeBox) }
			If($exists.postalCode -ne '' -and $exists.postalCode -ne $null) { $propertiesToExport.Add("postalCode",$exists.postalCode) }
			If($exists.st -ne '' -and $exists.st -ne $null) { $propertiesToExport.Add("state",$exists.st) }
			If($exists.streetAddress -ne '' -and $exists.streetAddress -ne $null) { $propertiesToExport.Add("streetAddress",$exists.streetAddress) }
			If($exists.title -ne '' -and $exists.title -ne $null) { $propertiesToExport.Add("title",$exists.title) }
			
			$TemplateOU = (Get-AdUser $CopyUserFrom).distinguishedName.Split(',',2)[1]	#get the OU template is in
			
			Try
			{
				#Set-aduser -Identity $sam -Department $CopyUserFrom.department -title $CopyUserFrom.title -description $CopyUserFrom.description -path $CopyUserFrom.path -physicalDeliveryOfficeName $CopyUserFrom.physicalDeliveryOfficeName -streetAddress $CopyUserFrom.streetAddress -l $CopyUserFrom.l -postalCode $CopyUserFrom.postalCode -st $CopyUserFrom.st -co $CopyUserFrom.co -company $CopyUserFrom.company -manager $CopyUserFrom.manager -ErrorAction Stop write host "Success"
				Set-ADUser -identity $sam @propertiesToExport
				Write-Host "All properties copied successfully from $($CopyUserFrom) successfully`r`n"
				(insertTimeStamp)+"All properties copied successfully from $($CopyUserFrom) successfully.." | Out-File $log -append
				Get-AdUser -Identity $sam | Move-ADObject -TargetPath $TemplateOU
				Write-Host "[INFO]`t User $sam moved to target OU : $($TemplateOU)"
				(insertTimeStamp) + "[INFO]`t User $sam moved to target OU : $($TemplateOU)" | Out-File $log -append
			}
			
			Catch
			{
				Write-Host "[ERROR]`t Oops, something went wrong when copying from template: $($_.Exception.Message)`r`n"
				(insertTimeStamp)+"Oops, something went wrong when copying from template: $($_.Exception.Message)" | Out-File $log -append
			}
			
			(insertTimeStamp) + "Copying group memberships from $($CopyUserFrom).." | Out-File $log -append		# copy group memberships
			Get-ADUser -Identity $CopyUserFrom -Properties memberof | Select-Object -ExpandProperty memberof | Add-ADGroupMember -Members $sam
			Write-Host "$($sam) replicated successfully from template $($CopyUserFrom)..`r`n"
			(insertTimeStamp) + "$($sam) replicated successfully from template $($CopyUserFrom).." | Out-File $log -append
		}
		
		else					#if template not found, notify, then continue from csv entries.
		{
			Write-Host "Template $($CopyUserFrom) NOT found in AD..`r`n"		
			(insertTimeStamp) + "Template $($CopyUserFrom) NOT found in AD.." | Out-File $log -append
		}

	}

	else						#if csv has no template entry, notify and continue from csv entries.
	{
		Write-Host "No Template to copy from in .csv`r`n"		
		(insertTimeStamp) + "No Template to copy from in .csv`r`n.." | Out-File $log -append
	}

					#continuing from csv entries...
	Write-Host "proceeding with user modification (if any) from rest of the attributes in .csv`r`n"
	(insertTimeStamp) + "proceeding with user modification (if any) from rest of the attributes in .csv" | Out-File $log -append	

	#If($Email -ne '' -and $Email -ne $null) { $propertiesToExport.Add("mail",$Email) }
	If($Department -ne '' -and $Department -ne $null) { $propertiesToExport.Add("department",$Department) }
	If($Title -ne '' -and $Title -ne $null) { $propertiesToExport.Add("title",$Title) }
	#If($Phone -ne '' -and $Phone -ne $null) { $propertiesToExport.Add("telephoneNumber",$Phone) }
	If($Description -ne '' -and $Description -ne $null) { $propertiesToExport.Add("description",$Description) }
	#If($PasswordNeverExpires -eq "true") { $propertiesToExport.Add("??",$True) }
	
	#If($OfficeName -ne '' -and $OfficeName -ne $null) { $propertiesToExport.Add("physicalDeliveryOfficeName",$OfficeName) }
	#If($StreetAddress -ne '' -and $StreetAddress -ne $null) { $propertiesToExport.Add("streetAddress",$StreetAddress) }
	#If($City -ne '' -and $City -ne $null) { $propertiesToExport.Add("l",$City) }
	#If($PostalCode -ne '' -and $PostalCode -ne $null) { $propertiesToExport.Add("postalCode",$PostalCode) }
	#If($State -ne '' -and $State -ne $null) { $propertiesToExport.Add("st",$State) }

	#If($Country -eq "Australia") {$Country = "NL"} Else { $Country = "EN" }
	#If($Country -ne '' -and $Country -ne $null) { $propertiesToExport.Add("co",$Country) }

	#If($Company -ne '' -and $Company -ne $null) { $propertiesToExport.Add("company",$Company) }
	#If($ProfilePath -ne '' -and $ProfilePath -ne $null) { $propertiesToExport.Add("profilePath",$ProfilePath) }
	#If($ScriptPath -ne '' -and $ScriptPath -ne $null) { $propertiesToExport.Add("scriptPath",$ScriptPath) }
	#$HomeDirectory = $_.HomeDirectory.Trim()
	#$HomeDrive = $_.HomeDrive.Trim()
	#$ProxyAddresses = $_.ProxyAddresses.Trim()

	Try { $ManagerExists = Get-ADUser -LDAPFilter "(sAMAccountName=$Manager)" }
	Catch { }
		
	If($ManagerExists -ne $null -and $ManagerExists -ne ''){ $propertiesToExport.Add("manager",$Manager) }

	Try
	{
		Set-ADUser -identity $sam @propertiesToExport
		Write-Host "$($sam) set up successfully`r`n"
		(insertTimeStamp)+"$($sam) set up successfully.." | Out-File $log -append
		
		If($AccountIsEnabled -eq "true")
		{
			Enable-ADAccount -Identity $sam
			Write-Host "$($sam) enabled successfully`r`n"
			(insertTimeStamp)+"$($sam) enabled successfully.." | Out-File $log -append
		}
		
        Write-Host "Attempting OU move to $($TargetOU)`r`n"
		(insertTimeStamp)+"Attempting OU move to $($TargetOU).." | Out-File $log -append
		
		
		If($TargetOU -ne '' -and $TargetOU -ne $null -and [adsi]::Exists("LDAP://$($TargetOU)"))
		{
			Get-AdUser -Identity $sam | Move-ADObject -TargetPath $TargetOU
			Write-Host "[INFO]`t User $sam moved to target OU : $($TargetOU)"
			(insertTimeStamp) + "[INFO]`t User $sam moved to target OU : $($TargetOU)" | Out-File $log -append
		}
		Else
		{
		  Write-Host "[ERROR]`t Targeted OU $($TargetOU) couldn't be found. Newly created user wasn't moved!"
		  "[ERROR]`t Targeted OU $($TargetOU) couldn't be found. Newly created user wasn't moved!" | Out-File $log -append
		}
	}
	
	Catch
	{
		#Write-Host "[ERROR]`t Oops, something went wrong: $($_.Exception.Message)`r`n"
		Write-Host "[ERROR]`t Oops, something went wrong: $($_)`r`n"
		(insertTimeStamp)+"Oops, something went wrong: $($_.Exception.Message)" | Out-File $log -append
	}

	#$newdn = (Get-ADUser $sam).DistinguishedName
	#Rename-ADObject -Identity $newdn -NewName ($GivenName + " " + $LastName)
	#Write-Host "$($sam) renamed to include a space`r`n"
	#(insertTimeStamp)+"$($sam) renamed to include a space.." | Out-File $log -append

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
			  Catch { Write-Host "[ERROR]`t Couldn't set the Extension Attribute : $($_.Exception.Message)" }
			}

			# Set ProxyAdresses
			If ($proxyAddresses -ne "")
			{
			  Try { $dn | Set-ADUser -Add @{proxyAddresses = ($ProxyAddresses -split ";")} -ErrorAction Stop }
			  Catch { Write-Host "[ERROR]`t Couldn't set the ProxyAddresses Attributes : $($_.Exception.Message)" }
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
			Write-Host "[ERROR]`t Oops, something went wrong: $($_.Exception.Message)`r`n"
		  }
		  
		 #>
		Write-Host "Moving to next iteration`r`n"
		(insertTimeStamp) + "Moving to next iteration`r`n" | Out-File $log -append
		$i++
	}
	"--------------------------------------------" + "`r`n" | Out-File $log -append

Write-Host "STOPPED SCRIPT"