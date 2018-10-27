Import-Module ActiveDirectory
$Groups = Get-ADGroup -Filter {GroupCategory -eq "Distribution"} -Properties Members
ForEach ($g in $Groups) {
    Write-Host $g.name
    Write-Host $g.members `n
}