﻿#--------------------------------------------------------------------------------------------------------
#Changes in Version 8.5
#Saving 'Section' from SAP records as 'Office' in AD record.
#--------------------------------------------------------------------------------------------------------
#Changes in Version 8.4
#Fixed manager not getting cleared if managerID is 000000
#Fixed output logs and email to show braces [] for all modified entries.
#New feature, modify the SAP delta file name with timestamp to avoid repeated processing the next time this script is run.
#--------------------------------------------------------------------------------------------------------

#Takes 5 parameters.
#1 Folder Path
#2 SAP Delta File name.
#3 Departments-Sections file name.
#4 From email address.
#5 To email address.
#Generates MergedUpdates.csv file to be fed to ADManager for updates.  

param($FolderPath, $SAPDeltaFileName, $DeptSectionsFileName, $MailFrom, $MailTo)

function Join-Object
{
    <#
    .SYNOPSIS
        Join data from two sets of objects based on a common value
    .DESCRIPTION
        Join data from two sets of objects based on a common value
        For more details, see the accompanying blog post:
            http://ramblingcookiemonster.github.io/Join-Object/
        For even more details,  see the original code and discussions that this borrows from:
            Dave Wyatt's Join-Object - http://powershell.org/wp/forums/topic/merging-very-large-collections
            Lucio Silveira's Join-Object - http://blogs.msdn.com/b/powershell/archive/2012/07/13/join-object.aspx
    .PARAMETER Left
        'Left' collection of objects to join.  You can use the pipeline for Left.
        The objects in this collection should be consistent.
        We look at the properties on the first object for a baseline.
    
    .PARAMETER Right
        'Right' collection of objects to join.
        The objects in this collection should be consistent.
        We look at the properties on the first object for a baseline.
    .PARAMETER LeftJoinProperty
        Property on Left collection objects that we match up with RightJoinProperty on the Right collection
    .PARAMETER RightJoinProperty
        Property on Right collection objects that we match up with LeftJoinProperty on the Left collection
    .PARAMETER LeftProperties
        One or more properties to keep from Left.  Default is to keep all Left properties (*).
        Each property can:
            - Be a plain property name like "Name"
            - Contain wildcards like "*"
            - Be a hashtable like @{Name="Product Name";Expression={$_.Name}}.
                 Name is the output property name
                 Expression is the property value ($_ as the current object)
                
                 Alternatively, use the Suffix or Prefix parameter to avoid collisions
                 Each property using this hashtable syntax will be excluded from suffixes and prefixes
    .PARAMETER RightProperties
        One or more properties to keep from Right.  Default is to keep all Right properties (*).
        Each property can:
            - Be a plain property name like "Name"
            - Contain wildcards like "*"
            - Be a hashtable like @{Name="Product Name";Expression={$_.Name}}.
                 Name is the output property name
                 Expression is the property value ($_ as the current object)
                
                 Alternatively, use the Suffix or Prefix parameter to avoid collisions
                 Each property using this hashtable syntax will be excluded from suffixes and prefixes
    .PARAMETER Prefix
        If specified, prepend Right object property names with this prefix to avoid collisions
        Example:
            Property Name                   = 'Name'
            Suffix                          = 'j_'
            Resulting Joined Property Name  = 'j_Name'
    .PARAMETER Suffix
        If specified, append Right object property names with this suffix to avoid collisions
        Example:
            Property Name                   = 'Name'
            Suffix                          = '_j'
            Resulting Joined Property Name  = 'Name_j'
    .PARAMETER Type
        Type of join.  Default is AllInLeft.
        AllInLeft will have all elements from Left at least once in the output, and might appear more than once
          if the where clause is true for more than one element in right, Left elements with matches in Right are
          preceded by elements with no matches.
          SQL equivalent: outer left join (or simply left join)
        AllInRight is similar to AllInLeft.
        
        OnlyIfInBoth will cause all elements from Left to be placed in the output, only if there is at least one
          match in Right.
          SQL equivalent: inner join (or simply join)
         
        AllInBoth will have all entries in right and left in the output. Specifically, it will have all entries
          in right with at least one match in left, followed by all entries in Right with no matches in left, 
          followed by all entries in Left with no matches in Right.
          SQL equivalent: full join
    .EXAMPLE
        #
        #Define some input data.
        $l = 1..5 | Foreach-Object {
            [pscustomobject]@{
                Name = "jsmith$_"
                Birthday = (Get-Date).adddays(-1)
            }
        }
        $r = 4..7 | Foreach-Object{
            [pscustomobject]@{
                Department = "Department $_"
                Name = "Department $_"
                Manager = "jsmith$_"
            }
        }
        #We have a name and Birthday for each manager, how do we find their department, using an inner join?
        Join-Object -Left $l -Right $r -LeftJoinProperty Name -RightJoinProperty Manager -Type OnlyIfInBoth -RightProperties Department
            # Name    Birthday             Department  
            # ----    --------             ----------  
            # jsmith4 4/14/2015 3:27:22 PM Department 4
            # jsmith5 4/14/2015 3:27:22 PM Department 5
    .EXAMPLE  
        #
        #Define some input data.
        $l = 1..5 | Foreach-Object {
            [pscustomobject]@{
                Name = "jsmith$_"
                Birthday = (Get-Date).adddays(-1)
            }
        }
        $r = 4..7 | Foreach-Object{
            [pscustomobject]@{
                Department = "Department $_"
                Name = "Department $_"
                Manager = "jsmith$_"
            }
        }
        #We have a name and Birthday for each manager, how do we find all related department data, even if there are conflicting properties?
        $l | Join-Object -Right $r -LeftJoinProperty Name -RightJoinProperty Manager -Type AllInLeft -Prefix j_
            # Name    Birthday             j_Department j_Name       j_Manager
            # ----    --------             ------------ ------       ---------
            # jsmith1 4/14/2015 3:27:22 PM                                    
            # jsmith2 4/14/2015 3:27:22 PM                                    
            # jsmith3 4/14/2015 3:27:22 PM                                    
            # jsmith4 4/14/2015 3:27:22 PM Department 4 Department 4 jsmith4  
            # jsmith5 4/14/2015 3:27:22 PM Department 5 Department 5 jsmith5  
    .EXAMPLE
        #
        #Hey!  You know how to script right?  Can you merge these two CSVs, where Path1's IP is equal to Path2's IP_ADDRESS?
        
        #Get CSV data
        $s1 = Import-CSV $Path1
        $s2 = Import-CSV $Path2
        #Merge the data, using a full outer join to avoid omitting anything, and export it
        Join-Object -Left $s1 -Right $s2 -LeftJoinProperty IP_ADDRESS -RightJoinProperty IP -Prefix 'j_' -Type AllInBoth |
            Export-CSV $MergePath -NoTypeInformation
    .EXAMPLE
        #
        # "Hey Warren, we need to match up SSNs to Active Directory users, and check if they are enabled or not.
        #  I'll e-mail you an unencrypted CSV with all the SSNs from gmail, what could go wrong?"
        
        # Import some SSNs. 
        $SSNs = Import-CSV -Path D:\SSNs.csv
        #Get AD users, and match up by a common value, samaccountname in this case:
        Get-ADUser -Filter "samaccountname -like 'wframe*'" |
            Join-Object -LeftJoinProperty samaccountname -Right $SSNs `
                        -RightJoinProperty samaccountname -RightProperties ssn `
                        -LeftProperties samaccountname, enabled, objectclass
    .NOTES
        This borrows from:
            Dave Wyatt's Join-Object - http://powershell.org/wp/forums/topic/merging-very-large-collections/
            Lucio Silveira's Join-Object - http://blogs.msdn.com/b/powershell/archive/2012/07/13/join-object.aspx
        Changes:
            Always display full set of properties
            Display properties in order (left first, right second)
            If specified, add suffix or prefix to right object property names to avoid collisions
            Use a hashtable rather than ordereddictionary (avoid case sensitivity)
    .LINK
        http://ramblingcookiemonster.github.io/Join-Object/
    .FUNCTIONALITY
        PowerShell Language
    #>
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipeLine = $true)]
        [object[]] $Left,

        # List to join with $Left
        [Parameter(Mandatory=$true)]
        [object[]] $Right,

        [Parameter(Mandatory = $true)]
        [string] $LeftJoinProperty,

        [Parameter(Mandatory = $true)]
        [string] $RightJoinProperty,

        [object[]]$LeftProperties = '*',

        # Properties from $Right we want in the output.
        # Like LeftProperties, each can be a plain name, wildcard or hashtable. See the LeftProperties comments.
        [object[]]$RightProperties = '*',

        [validateset( 'AllInLeft', 'OnlyIfInBoth', 'AllInBoth', 'AllInRight')]
        [Parameter(Mandatory=$false)]
        [string]$Type = 'AllInLeft',

        [string]$Prefix,
        [string]$Suffix
    )
    Begin
    {
        function AddItemProperties($item, $properties, $hash)
        {
            if ($null -eq $item)
            {
                return
            }

            foreach($property in $properties)
            {
                $propertyHash = $property -as [hashtable]
                if($null -ne $propertyHash)
                {
                    $hashName = $propertyHash["name"] -as [string]         
                    $expression = $propertyHash["expression"] -as [scriptblock]

                    $expressionValue = $expression.Invoke($item)[0]
            
                    $hash[$hashName] = $expressionValue
                }
                else
                {
                    foreach($itemProperty in $item.psobject.Properties)
                    {
                        if ($itemProperty.Name -like $property)
                        {
                            $hash[$itemProperty.Name] = $itemProperty.Value
                        }
                    }
                }
            }
        }

        function TranslateProperties
        {
            [cmdletbinding()]
            param(
                [object[]]$Properties,
                [psobject]$RealObject,
                [string]$Side)

            foreach($Prop in $Properties)
            {
                $propertyHash = $Prop -as [hashtable]
                if($null -ne $propertyHash)
                {
                    $hashName = $propertyHash["name"] -as [string]         
                    $expression = $propertyHash["expression"] -as [scriptblock]

                    $ScriptString = $expression.tostring()
                    if($ScriptString -notmatch 'param\(')
                    {
                        Write-Verbose "Property '$HashName'`: Adding param(`$_) to scriptblock '$ScriptString'"
                        $Expression = [ScriptBlock]::Create("param(`$_)`n $ScriptString")
                    }
                
                    $Output = @{Name =$HashName; Expression = $Expression }
                    Write-Verbose "Found $Side property hash with name $($Output.Name), expression:`n$($Output.Expression | out-string)"
                    $Output
                }
                else
                {
                    foreach($ThisProp in $RealObject.psobject.Properties)
                    {
                        if ($ThisProp.Name -like $Prop)
                        {
                            Write-Verbose "Found $Side property '$($ThisProp.Name)'"
                            $ThisProp.Name
                        }
                    }
                }
            }
        }

        function WriteJoinObjectOutput($leftItem, $rightItem, $leftProperties, $rightProperties)
        {
            $properties = @{}

            AddItemProperties $leftItem $leftProperties $properties
            AddItemProperties $rightItem $rightProperties $properties

            New-Object psobject -Property $properties
        }

        #Translate variations on calculated properties.  Doing this once shouldn't affect perf too much.
        foreach($Prop in @($LeftProperties + $RightProperties))
        {
            if($Prop -as [hashtable])
            {
                foreach($variation in ('n','label','l'))
                {
                    if(-not $Prop.ContainsKey('Name') )
                    {
                        if($Prop.ContainsKey($variation) )
                        {
                            $Prop.Add('Name',$Prop[$Variation])
                        }
                    }
                }
                if(-not $Prop.ContainsKey('Name') -or $Prop['Name'] -like $null )
                {
                    Throw "Property is missing a name`n. This should be in calculated property format, with a Name and an Expression:`n@{Name='Something';Expression={`$_.Something}}`nAffected property:`n$($Prop | out-string)"
                }


                if(-not $Prop.ContainsKey('Expression') )
                {
                    if($Prop.ContainsKey('E') )
                    {
                        $Prop.Add('Expression',$Prop['E'])
                    }
                }
            
                if(-not $Prop.ContainsKey('Expression') -or $Prop['Expression'] -like $null )
                {
                    Throw "Property is missing an expression`n. This should be in calculated property format, with a Name and an Expression:`n@{Name='Something';Expression={`$_.Something}}`nAffected property:`n$($Prop | out-string)"
                }
            }        
        }

        $leftHash = @{}
        $rightHash = @{}

        # Hashtable keys can't be null; we'll use any old object reference as a placeholder if needed.
        $nullKey = New-Object psobject
        
        $bound = $PSBoundParameters.keys -contains "InputObject"
        if(-not $bound)
        {
            [System.Collections.ArrayList]$LeftData = @()
        }
    }
    Process
    {
        #We pull all the data for comparison later, no streaming
        if($bound)
        {
            $LeftData = $Left
        }
        Else
        {
            foreach($Object in $Left)
            {
                [void]$LeftData.add($Object)
            }
        }
    }
    End
    {
        foreach ($item in $Right)
        {
            $key = $item.$RightJoinProperty

            if ($null -eq $key)
            {
                $key = $nullKey
            }

            $bucket = $rightHash[$key]

            if ($null -eq $bucket)
            {
                $bucket = New-Object System.Collections.ArrayList
                $rightHash.Add($key, $bucket)
            }

            $null = $bucket.Add($item)
        }

        foreach ($item in $LeftData)
        {
            $key = $item.$LeftJoinProperty

            if ($null -eq $key)
            {
                $key = $nullKey
            }

            $bucket = $leftHash[$key]

            if ($null -eq $bucket)
            {
                $bucket = New-Object System.Collections.ArrayList
                $leftHash.Add($key, $bucket)
            }

            $null = $bucket.Add($item)
        }

        $LeftProperties = TranslateProperties -Properties $LeftProperties -Side 'Left' -RealObject $LeftData[0]
        $RightProperties = TranslateProperties -Properties $RightProperties -Side 'Right' -RealObject $Right[0]

        #I prefer ordered output. Left properties first.
        [string[]]$AllProps = $LeftProperties

        #Handle prefixes, suffixes, and building AllProps with Name only
        $RightProperties = foreach($RightProp in $RightProperties)
        {
            if(-not ($RightProp -as [Hashtable]))
            {
                Write-Verbose "Transforming property $RightProp to $Prefix$RightProp$Suffix"
                @{
                    Name="$Prefix$RightProp$Suffix"
                    Expression=[scriptblock]::create("param(`$_) `$_.'$RightProp'")
                }
                $AllProps += "$Prefix$RightProp$Suffix"
            }
            else
            {
                Write-Verbose "Skipping transformation of calculated property with name $($RightProp.Name), expression:`n$($RightProp.Expression | out-string)"
                $AllProps += [string]$RightProp["Name"]
                $RightProp
            }
        }

        $AllProps = $AllProps | Select -Unique

        Write-Verbose "Combined set of properties: $($AllProps -join ', ')"

        foreach ( $entry in $leftHash.GetEnumerator() )
        {
            $key = $entry.Key
            $leftBucket = $entry.Value

            $rightBucket = $rightHash[$key]

            if ($null -eq $rightBucket)
            {
                if ($Type -eq 'AllInLeft' -or $Type -eq 'AllInBoth')
                {
                    foreach ($leftItem in $leftBucket)
                    {
                        WriteJoinObjectOutput $leftItem $null $LeftProperties $RightProperties | Select $AllProps
                    }
                }
            }
            else
            {
                foreach ($leftItem in $leftBucket)
                {
                    foreach ($rightItem in $rightBucket)
                    {
                        WriteJoinObjectOutput $leftItem $rightItem $LeftProperties $RightProperties | Select $AllProps
                    }
                }
            }
        }

        if ($Type -eq 'AllInRight' -or $Type -eq 'AllInBoth')
        {
            foreach ($entry in $rightHash.GetEnumerator())
            {
                $key = $entry.Key
                $rightBucket = $entry.Value

                $leftBucket = $leftHash[$key]

                if ($null -eq $leftBucket)
                {
                    foreach ($rightItem in $rightBucket)
                    {
                        WriteJoinObjectOutput $null $rightItem $LeftProperties $RightProperties | Select $AllProps
                    }
                }
            }
        }
    }
}

# Get script Start Time (used to measure run time)
$startDTM = (Get-Date)

#Create log date and user disabled date
$logdate = Get-Date -Format yyyyMMddHHmm

#Define CSV and log file location variables
#they have to be on the same location as the script

$smtpsettings = @{
  To =  $MailTo
  From = $MailFrom
    Subject = "AD SAP Employee Sync Report " + $logdate
  SmtpServer = "mail.emiratessteel.com"
  }

$LeftFile = Import-Csv -path "$FolderPath\$SAPDeltaFileName" -Delimiter '~'
$RightFile = Import-Csv -path "$FolderPath\$DeptSectionsFileName"

Join-Object -Left $LeftFile -Right $RightFile -LeftJoinProperty Section -RightJoinProperty Section -Type AllInLeft -RightProperties Department | Export-Csv -NoTypeInformation "$FolderPath\MergedUpdates.csv" -Delimiter '~'

$MergedUpdatesFile = $FolderPath + "\MergedUpdates.csv"
Rename-Item "$FolderPath\$SAPDeltaFileName" -NewName "$FolderPath\$SAPDeltaFileName - Processed $logdate.csv"

$logfile = $FolderPath + "\" + "AD SAP Data Sync Report $logdate.log"
$MailBodyContent = "SAP AD Employee Sync Report... <br/><br/><br/>"

$UpdatesRecords = Import-Csv -path $MergedUpdatesFile -Delimiter '~'

foreach ($Record in $UpdatesRecords)
{
    $EmployeeID = ($Record.EmployeeID).TrimStart('0')
    $Title = $Record.JobTitle
    $Department = $Record.Department
    $ManagerID = ($Record.ManagerID).TrimStart('0')
	$Section = $Record.Section

    Try { $SAMinAD = Get-ADUser -LDAPFilter "(description=$EmployeeID)"} 
	Catch { }
    
    #Execute set-aduser below only if $sam is in AD and also is in the excel file, else ignore#
		If($SAMinAD -ne $null -and $SAMinAD -ne '')
		{
			#added the 'if clause' to ensure that blank fields in the CSV are ignored.
			#the object names must be the LDAP names. get values using ADSI Edit
	
			#Modify the Manager using -manager switch as it does not accept the -Replace switch.
			If ($ManagerID -eq '' -or $ManagerID -eq '0') {$ManagerID = $null}
            $ManagerAlias = $null
            $ManagerAlias = Get-ADUser -LDAPFilter "(description=$ManagerID)" -Properties SamAccountName | Select-Object SamAccountName
            Get-ADUser -LDAPFilter "(description=$EmployeeID)" | Set-ADUser -Manager $ManagerAlias

            #Modify the Title
            If ($Title -eq '' -or $Title -eq '0') {Get-ADUser -LDAPFilter "(description=$EmployeeID)" | Set-ADUser -Clear Title}
            Else {Get-ADUser -LDAPFilter "(description=$EmployeeID)" | Set-ADUser -Title $Title}
            
            #Modify the Department
			If ($Department -eq '' -or $Department -eq '0') {Get-ADUser -LDAPFilter "(description=$EmployeeID)" | Set-ADUser -Clear Department}
			Else {Get-ADUser -LDAPFilter "(description=$EmployeeID)" | Set-ADUser -Department $Department}

            #Modify the Office
            If ($Section -eq '' -or $Section -eq '0') {$Section = $null}
			Get-ADUser -LDAPFilter "(description=$EmployeeID)" | Set-ADUser -Office $Section

			$TimeStamp = Get-Date -Format yyyyMMddHHmm
			$TimeStamp + " : " + $EmployeeID + " modified to Title: [" + $Title + "] under Department: [" + $Department + "], Section: [" + $Section + "] reporting to Manager: [" + $ManagerID + "]"  | Out-File $logfile -Append
            $MailBodyContent += $EmployeeID + " modified to Title: [" + $Title + "] under Department: [" + $Department + "], Section: [" + $Section + "] reporting to Manager: [" + $ManagerID + "]<br/>"

            #Set a flag to indicate that the user has been updated on AD.
			#When I export, I will omit all users with thie flag enabled 
		}

		Else

		{ #Log error for users that are not in Active Directory or with no Logon name in excel file
			$TimeStamp = Get-Date -Format yyyyMMddHHmm
            $TimeStamp + " : " + $EmployeeID + " not modified because it does not exist in AD" | Out-File $logfile -Append
			$MailBodyContent += $EmployeeID + " not modified because it does not exist in AD <br/>"
		}
}

#$LogFileContent = Get-Content $logfile | Out-String

Send-MailMessage @smtpsettings -Body  $MailBodyContent -BodyAsHtml -Encoding ([System.Text.Encoding]::UTF8) -Attachments $logfile

#Finish
#The lins below calculates how long it takes to run this script
#Get End Time
$endDTM = (Get-Date)

# Echo Time elapsed
"Elapsed Time: $(($endDTM-$startDTM).totalseconds) seconds"
"Elapsed Time: $(($endDTM-$startDTM).totalminutes) minutes"

#SCRIPT ENDS
