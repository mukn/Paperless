<## -- Move-DuplicatesInSharePoint --

This script is used to move files from the old service technician library to
the new one. This is a one-off script and should not be run again unless the 
new library is deleted.

Updated 26 May 2020.

#>

## Functions

## Declared variables
$NewLibraryUri = "https://nacgroup.sharepoint.com/Sites/Service9/Shared%20Documents/Technicians/"
$NewLibraryPath = "~\Noyes Air Conditioning\Service - Technicians"
$OldLibraryUri1 = "https://nacgroup.sharepoint.com/Projects%200C/"
$OldLibraryPath1 = "~\Noyes Air Conditioning\Noyes Air Conditioning - Projects (0-C)"
$OldLibraryUri2 = "https://nacgroup.sharepoint.com/Projects%20CN/"
$OldLibraryPath2 = "~\Noyes Air Conditioning\Noyes Air Conditioning - Projects (D-N)"
$OldLibraryUri3 = "https://nacgroup.sharepoint.com/Projects%20OZ/"
$OldLibraryPath3 = "~\Noyes Air Conditioning\Noyes Air Conditioning - Projects (O-Z)"

Set-Location -Path $NewLibraryPath
$dirsService = Get-ChildItem | Select-Object -ExpandProperty Name
Set-Location -Path $OldLibraryPath3
$dirsProjects = Get-ChildItem | Select-Object -ExpandProperty Name

$matchCounter = 0
$matchList = @()
$mismatchList = @()
ForEach ($sd in $dirsService) {
    $name = $sd.Split("-")[0].Trim()
    if ($name | ?{$dirsProjects -match $_}) {
        Write-Host "Match found for $sd."
        New-Item -ItemType File -Name "Updated-$(Get-Date -Format dd-MM-yyyy).txt" -Path "$NewLibraryPath\$sd\"
        $matchCounter++
        $matchList += $sd
        $dirContents = Get-ChildItem "$OldLibraryPath3\$name" -Recurse -Exclude "Quoted Projects"
        Copy-Item -Path $OldLibraryPath3\$name\* -Destination "$NewLibraryPath\$sd\" -Recurse -Force
        }
    else {
        $mismatch += $sd
        $mismatchList += $sd
        }
}