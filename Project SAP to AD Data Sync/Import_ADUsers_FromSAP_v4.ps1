# Takes NO Parameters and only applies changes.

    
    $EmployeeID = $_.EmployeeID
    $Title = $_.JobTitle
    $Department = $_.Department
    $ManagerID = $_.ManagerID

    Try { $SAMinAD = Get-ADUser -LDAPFilter "(description=$EmployeeID)"} 
	Catch { }
    
    #Execute set-aduser below only if $sam is in AD and also is in the excel file, else ignore#
		If($SAMinAD -ne $null -and $SAMinAD -ne '')
		{

			#added the 'if clause' to ensure that blank fields in the CSV are ignored.
			#the object names must be the LDAP names. get values using ADSI Edit
	
			#Manager did not accept the -Replace switch. It works with the -manager switch
			IF ($ManagerID -eq '' -or $ManagerID -eq '0') {$ManagerID = $null}

            $ManagerAlias = Get-ADUser -LDAPFilter "(description=$ManagerID)" -Properties SamAccountName | Select-Object SamAccountName

            Get-ADUser -LDAPFilter "(description=$EmployeeID)" | Set-ADUser -Replace @{Title=$Title;Department=$Department;} -Manager $ManagerAlias
            #$EmployeeID + " modified as " + $Title + " under " + $Department + " reporting to " + $ManagerID | Out-File $logfile -Append

            #Set a flag to indicate that the user has been updated on AD.
			#When I export, I will omit all users with thie flag enabled 
		}

		Else

		{ #Log error for users that are not in Active Directory or with no Logon name in excel file
			#$EmployeeID + " Not modified because it does not exist in AD" | Out-File $logfile -Append
		}
