$myArray = @{big=""; medium="m"; small="s"}
$i=1
foreach ($item in $myArray.Values)
{
 Write-Host "Iteration $($i)"
 $item
 if(($item.length) -eq 0) { $myArray.remove($item) }
 $i++
}
$myArray
