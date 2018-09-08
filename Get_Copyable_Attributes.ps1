$directoryEntryType = 'System.DirectoryServices.DirectoryEntry'
$directorySearcherType = 'System.DirectoryServices.DirectorySearcher'

$rootdse = New-Object -TypeName $directoryEntryType -ArgumentList "LDAP://RootDSE"

$schema = New-Object -TypeName $directoryEntryType -ArgumentList @(
    "LDAP://cn=schema,$($rootdse.configurationNamingContext)"
)

$filter = "(&(objectCategory=attributeSchema)(searchflags:1.2.840.113556.1.4.803:=16))"

$props = @("Name","searchFlags")

$searcher = New-Object -TypeName $directorySearcherType -ArgumentList @(
    $schema
    $filter
    ,$props
    'SubTree'
)
 
$searcher.FindAll() | 
Select -ExpandProperty Properties |
ForEach-Object {
    New-Object -TypeName psobject -Property @{Name=$_.name[0]}
} | Sort-Object -Property Name