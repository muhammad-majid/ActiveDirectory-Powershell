<#
Created by Muhammad Majid, HuonIT
Order of execution:
01. Check FistName, Lastname and Password fields are not blank, proceed an create a new user (firstname.lastname)
02. Once created, check copyUserFrom field, find that user, and copy all properties of that user to the newly created user.
03. Irrespective of 'copyUserFrom' field being empty or not, proceed and modify or the newly created user with rest of the entries in the csv.
[This means that entries in csv will overwrite the copyUserFrom properties set]
04. Log output is recorded in the same folder script was run from.
#>


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
$dnsroot  = (Get-ADDomain).DNSRoot
$i        = 0

#----------------------------------------------------------
#START FUNCTIONS
#----------------------------------------------------------

Function insertTimeStamp { return (Get-Date).ToString('yyyy-MM-dd HH:mm:ss') + ' : ' }

Write-Host "SCRIPT STARTED `r`n"
(insertTimeStamp) + "Create AD Users in batches from CSV - v3.5" | Out-File $log -append
(insertTimeStamp) + "Created by Muhammad Majid, HuonIT.." | Out-File $log -append
(insertTimeStamp) + "Processing started.." | Out-File $log -append

Import-CSV $newpath | ForEach-Object{
	
	" " | Out-File $log -append
	"--------------------------------------------" | Out-File $log -append
	
	$i++;
	$GivenName = $_.GivenName.Trim()
	$LastName = $_.LastName.Trim()
	$CopyUserFrom = $_.CopyUserFrom.Trim()
	$Password = ConvertTo-SecureString -AsPlainText $_.Password -force
	$Email = $_.Email.Trim()
	$Department = $_.Department.Trim()
	$Title = $_.Title.Trim()
	$Phone = $_.Phone.TrimStart('0')
	$Mobile = $_.Mobile.TrimStart('0')
	$Description = $_.Description.Trim()
	$PasswordNeverExpires = $_.PasswordNeverExpires.ToLower()
	$AccountIsEnabled = $_.AccountIsEnabled.ToLower()
	$Manager = $_.Manager.Trim()
	$TargetOU = $_.TargetOU.Trim()
	$OfficeName = $_.OfficeName.Trim()
	$StreetAddress = $_.StreetAddress.Trim()
	$POBox = $_.POBox.Trim()
	$City = $_.City
	$State = $_.State.Trim()
	$PostalCode = $_.PostalCode.Trim()
	$Country = ($_.Country.Trim()).toLower()
	$Company = $_.Company.Trim()
	$WebPage = $_.WebPage.Trim()
	$ProfilePath = $_.ProfilePath.Trim()
	$ScriptPath = $_.ScriptPath.Trim()
	$HomeDirectory = $_.HomeDirectory.Trim()
	$HomeDrive = $_.HomeDrive.Trim()
	$ProxyAddresses = $_.ProxyAddresses.Trim()

	$sam = $GivenName.ToLower() + "." + $LastName.ToLower()
	$Name = $GivenName+" "+$LastName
	$userPrincipalName = $sam+"@"+$dnsroot

	#poperties that will be imported from template user ($CopyUserFrom) when necessary.
	$propertiesImported = @("description", "physicalDeliveryOfficeName", "wWWHomePage", "streetAddress", "postOfficeBox", "l", "st", "postalCode", "c", "title", "department", "company", "manager", "enabled")
	#as well as group memberhips at the time of replication.
	
	Write-Host "Running iteration $i for $GivenName $LastName`r`n"
	Write-Host "Subject username is $($sam)`r`n"
	"" | Out-File $log -append
	(insertTimeStamp) + "Running iteration $i [$GivenName $LastName] .." | Out-File $log -append
	(insertTimeStamp) + "Subject username would be $($sam).." | Out-File $log -append
	(insertTimeStamp) + "-----------PHASE 1-----------" | Out-File $log -append	

	#if first name, last name or password are missing, skip move to next iteration.
	If ($GivenName -eq '' -Or $GivenName -eq $null -Or $LastName -eq '' -Or $LastName -eq $null -Or $_.Password -eq '' -Or $_.Password -eq $null)
	{
		Write-Host "[WARNING]`t Please provide valid GivenName, LastName and Password.`r`nProcessing skipped for Record $($i) : $($Name)`r`n"
		(insertTimeStamp) + "[WARNING]`t Please provide valid GivenName, LastName and Password" | Out-File $log -append
		(insertTimeStamp) + "Processing skipped for Record $($i) : $($Name)" | Out-File $log -append
		return
	}


	Try { $exists = Get-ADUser -LDAPFilter "(sAMAccountName=$sam)" }		
	Catch { }
	If($exists -ne $null -and $exists -ne '')		#if account already exists, skip and move to next iteration.
	{
		Write-Host "[WARNING]`t Record $($i) : $($sam) already exists in AD..`r`n Skipping to next iteration..`r`n"
		(insertTimeStamp) + "[WARNING]`t Record $($i): $($sam) already exists in AD.." | Out-File $log -append
		(insertTimeStamp) + "Processing skipped for Record $($i) : $($Name)" | Out-File $log -append
		return
	}

	Write-Host "Creating User`r`n"
	(insertTimeStamp) + "Creating User.." | Out-File $log -append
	New-ADuser -sAMAccountName $sam -UserPrincipalName $userPrincipalName -Name $Name -GivenName $GivenName -Surname $LastName -DisplayName $Name -AccountPassword $Password
	Write-Host "$($sam) created successfully`r`n"
	(insertTimeStamp) + "$($sam) created successfully.." | Out-File $log -append
	$propertiesToExport = @{}
		
	#If ($CopyUserFrom -ne $null -and $CopyUserFrom -ne '')		#if csv has template for the new user
	If ($CopyUserFrom -eq $null -or $CopyUserFrom -eq '')		#if no template user mentioned in csv
	{
		Write-Host "No Template to copy from in .csv`r`n"		
		(insertTimeStamp) + "No Template to copy from in .csv.." | Out-File $log -append
	}

	else #if csv has template for the new user
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

			If($exists.description -ne '' -and $exists.description -ne $null)
			{
				Try{ Set-ADUser -identity $sam -description $exists.description
				(insertTimeStamp) + "Description set successfully to: $($exists.description)" | Out-File $log -append
				}
				catch{(insertTimeStamp)+"Oops, something went wrong when copying description: $($_.Exception.Message)" | Out-File $log -append}
			}

			If($exists.physicalDeliveryOfficeName -ne '' -and $exists.physicalDeliveryOfficeName -ne $null)
			{
				Try{ Set-ADUser -identity $sam -office $exists.physicalDeliveryOfficeName
				(insertTimeStamp) + "Office set successfully to: $($exists.physicalDeliveryOfficeName)" | Out-File $log -append
				}
				catch{(insertTimeStamp)+"Oops, something went wrong when copying office: $($_.Exception.Message)" | Out-File $log -append}
			}

			If($exists.wWWHomePage -ne '' -and $exists.wWWHomePage -ne $null)
			{
				Try{ Set-ADUser -identity $sam -homepage $exists.wWWHomePage
				(insertTimeStamp) + "HomePage set successfully to: $($exists.wWWHomePage)" | Out-File $log -append
				}
				catch{(insertTimeStamp)+"Oops, something went wrong when copying homepage: $($_.Exception.Message)" | Out-File $log -append}
			}

			If($exists.streetAddress -ne '' -and $exists.streetAddress -ne $null)
			{
				Try{ Set-ADUser -identity $sam -streetAddress $exists.streetAddress
				(insertTimeStamp) + "Street Address set successfully to: $($exists.streetAddress)" | Out-File $log -append
				}
				catch{(insertTimeStamp)+"Oops, something went wrong when copying street: $($_.Exception.Message)" | Out-File $log -append}
			}

			If($exists.postOfficeBox -ne '' -and $exists.postOfficeBox -ne $null)
			{
				Try{ Set-ADUser -identity $sam -pobox $exists.postOfficeBox
				(insertTimeStamp) + "P.O Box set successfully to: $($exists.postOfficeBox)" | Out-File $log -append
				}
				catch{(insertTimeStamp)+"Oops, something went wrong when copying POBox: $($_.Exception.Message)" | Out-File $log -append}
			}

			If($exists.l -ne '' -and $exists.l -ne $null)
			{
				Try{ Set-ADUser -identity $sam -city $exists.l
				(insertTimeStamp) + "City set successfully to: $($exists.l)" | Out-File $log -append
				}
				catch{(insertTimeStamp)+"Oops, something went wrong when copying city: $($_.Exception.Message)" | Out-File $log -append}
			}

			If($exists.st -ne '' -and $exists.st -ne $null)
			{
				Try{ Set-ADUser -identity $sam -state $exists.st
				(insertTimeStamp) + "State set successfully to: $($exists.st)" | Out-File $log -append
				}
				catch{(insertTimeStamp)+"Oops, something went wrong when copying state / province: $($_.Exception.Message)" | Out-File $log -append}
			}

			If($exists.postalCode -ne '' -and $exists.postalCode -ne $null)
			{
				Try{ Set-ADUser -identity $sam -postalCode $exists.postalCode
				(insertTimeStamp) + "Postal Code set successfully to: $($exists.postalCode)" | Out-File $log -append
				}
				catch{(insertTimeStamp)+"Oops, something went wrong when copying zip / postal code: $($_.Exception.Message)" | Out-File $log -append}
			}

			If($exists.c -ne '' -and $exists.c -ne $null)
			{
				Try{ Set-ADUser -identity $sam -country $exists.country
				(insertTimeStamp) + "Country set successfully to: $($exists.c)" | Out-File $log -append
				}
				catch{(insertTimeStamp)+"Oops, something went wrong when copying country: $($_.Exception.Message)" | Out-File $log -append}
			}

			If($exists.enabled -eq $true)
			{
				Try{ Set-ADUser -identity $sam -enabled $true
				(insertTimeStamp) + "Account Enabled successfully" | Out-File $log -append
				}
				catch{(insertTimeStamp)+"Oops, something went wrong when enabling account: $($_.Exception.Message)" | Out-File $log -append}
			}

			If($exists.title -ne '' -and $exists.title -ne $null)
			{
				Try{ Set-ADUser -identity $sam -title $exists.title
				(insertTimeStamp) + "Job Title set successfully to: $($exists.title)" | Out-File $log -append
				}
				catch{(insertTimeStamp)+"Oops, something went wrong when copying job title: $($_.Exception.Message)" | Out-File $log -append}
			}

			If($exists.department -ne '' -and $exists.department -ne $null)
			{
				Try{ Set-ADUser -identity $sam -department $exists.department
				(insertTimeStamp) + "Departemnt set successfully to: $($exists.department)" | Out-File $log -append
				}
				catch{(insertTimeStamp)+"Oops, something went wrong when copying department: $($_.Exception.Message)" | Out-File $log -append}
			}
			
			If($exists.company -ne '' -and $exists.company -ne $null)
			{
				Try{ Set-ADUser -identity $sam -company $exists.company
				(insertTimeStamp) + "Company set successfully to: $($exists.company)" | Out-File $log -append
				}
				catch{(insertTimeStamp)+"Oops, something went wrong when copying company: $($_.Exception.Message)" | Out-File $log -append}
			}
			
			If($exists.manager -ne '' -and $exists.manager -ne $null)
			{
				Try{ Set-ADUser -identity $sam -manager $exists.manager
				(insertTimeStamp) + "Manager set successfully" | Out-File $log -append
				}
				catch{(insertTimeStamp)+"Oops, something went wrong when copying manager: $($_.Exception.Message)" | Out-File $log -append}
			}
			
			$TemplateOU = (Get-AdUser $CopyUserFrom).distinguishedName.Split(',',2)[1]	#get the OU template is in

			Try{ Get-AdUser -Identity $sam | Move-ADObject -TargetPath $TemplateOU
				Write-Host "User $sam moved to target OU : $($TemplateOU)"
				(insertTimeStamp) + "User $sam moved to target OU : $($TemplateOU)" | Out-File $log -append
			}
			catch{(insertTimeStamp)+"Oops, something went wrong when changing OU: $($_.Exception.Message)" | Out-File $log -append}

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

	#continuing from csv entries...
	Write-Host "proceeding with user modification (if any) from rest of the entries in .csv file`r`n"
	(insertTimeStamp) + "-----------PHASE 2-----------" | Out-File $log -append	
	(insertTimeStamp) + "proceeding with user modification (if any) from rest of the attributes in .csv file.." | Out-File $log -append	

	$propertiesToExport2 = @{}

	If($Email -ne '' -and $Email -ne $null) { $propertiesToExport2.Add("emailaddress",$Email) }
	If($Department -ne '' -and $Department -ne $null) { $propertiesToExport2.Add("department",$Department) }
	If($Title -ne '' -and $Title -ne $null) { $propertiesToExport2.Add("title",$Title) }
	If($Phone -ne '' -and $Phone -ne $null) { $propertiesToExport2.Add("OfficePhone",$Phone) }
	If($Mobile -ne '' -and $Mobile -ne $null) { $propertiesToExport2.Add("mobilephone",$Mobile) }
	If($Description -ne '' -and $Description -ne $null) { $propertiesToExport2.Add("description",$Description) }
	If($PasswordNeverExpires -eq $True) { $propertiesToExport2.Add("passwordneverexpires",$True) } else { $propertiesToExport2.Add("passwordneverexpires",$False) }
	If($AccountIsEnabled -eq $True) { $propertiesToExport2.Add("enabled",$True) } else { $propertiesToExport2.Add("enabled",$False) }
	#Manager  done below
	#TargetOU  done below
	If($OfficeName -ne '' -and $OfficeName -ne $null) { $propertiesToExport2.Add("Office",$OfficeName) }
	If($StreetAddress -ne '' -and $StreetAddress -ne $null) { $propertiesToExport2.Add("streetAddress",$StreetAddress) }
	If($POBox -ne '' -and $POBox -ne $null) { $propertiesToExport2.Add("pobox",$POBox) }
	If($City -ne '' -and $City -ne $null) { $propertiesToExport2.Add("city",$City) }
	If($State -ne '' -and $State -ne $null) { $propertiesToExport2.Add("state",$State) }
	If($PostalCode -ne '' -and $PostalCode -ne $null) { $propertiesToExport2.Add("postalCode",$PostalCode) }

	If($Country -eq "australia" -or $Country -eq '' -or $Country -eq $null)	{ $Country = "AU" } else { $Country = "EN" }
	$propertiesToExport2.Add("country",$Country)

	If($Company -ne '' -and $Company -ne $null) { $propertiesToExport2.Add("company",$Company) }
	If($WebPage -ne '' -and $WebPage -ne $null) { $propertiesToExport2.Add("HomePage",$WebPage) }
	If($ProfilePath -ne '' -and $ProfilePath -ne $null) { $propertiesToExport2.Add("profilePath",$ProfilePath) }
	If($ScriptPath -ne '' -and $ScriptPath -ne $null) { $propertiesToExport2.Add("scriptPath",$ScriptPath) }
	#$HomeDirectory = $_.HomeDirectory.Trim()
	#$HomeDrive = $_.HomeDrive.Trim()
	#$ProxyAddresses = $_.ProxyAddresses.Trim()

	#---------------------------------------------------------------------------------------------	
	Try { $ManagerExists = Get-ADUser -LDAPFilter "(sAMAccountName=$Manager)" }
	Catch { }
		
	If($ManagerExists -ne $null -and $ManagerExists -ne ''){ $propertiesToExport2.Add("manager",$Manager) }

	Write-Host "Below properties will be copied over based on csv entries..`r`n"
	(insertTimeStamp) + "Below properties will be copied over based on csv entries.." | Out-File $log -append

	$PropertiesToExport2.GetEnumerator() | ForEach-Object {
		Write-Host "        Key: $($_.Key), Value: $($_.Value)"
		(insertTimeStamp) + "        Key: $($_.Key), Value: $($_.Value)" | Out-File $log -append
	}

	Try
	{
		Set-ADUser -identity $sam @propertiesToExport2
		Write-Host "$($sam) set up successfully`r`n"
		(insertTimeStamp)+"$($sam) set up successfully.." | Out-File $log -append
		
        Write-Host "Attempting OU move to $($TargetOU)`r`n"
		(insertTimeStamp)+"Attempting OU move to $($TargetOU).." | Out-File $log -append
		
		
		If($TargetOU -ne '' -and $TargetOU -ne $null -and [adsi]::Exists("LDAP://$($TargetOU)"))
		{
			Get-AdUser -Identity $sam | Move-ADObject -TargetPath $TargetOU
			Write-Host "User $sam moved successfully to target OU : $($TargetOU)"
			(insertTimeStamp) + "User $sam moved successfully to target OU : $($TargetOU)" | Out-File $log -append
		}
		Else
		{
		  Write-Host "[Warning]`t Targete OU left blank, or couldn't be found. Newly created user wasn't moved!"
		  "[Warning]`t Targete OU left blank, or couldn't be found. Newly created user wasn't moved!" | Out-File $log -append
		}
	}
	
	Catch
	{
		Write-Host "[ERROR]`t Oops, something went wrong: $($_.Exception.Message)`r`n"
		(insertTimeStamp)+"Oops, something went wrong: $($_.Exception.Message)" | Out-File $log -append
	}
	
	Write-Host "Moving to next iteration`r`n"
	(insertTimeStamp) + "Moving to next iteration" | Out-File $log -append
}

Write-Host "SCRIPT STOPPED"
