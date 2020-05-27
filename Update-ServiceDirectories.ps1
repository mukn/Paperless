<## -- Update-ServiceDirectories --

This script reviews each of the directories in the Service - Technicians 
document library with the intent of moving directories for sites tagged
as inactive or unused as well as creating new site directories, as 
necessary.

#>

# Declare the standard variables. These assume that the script is run from a computer with the libraries 
# synchronized.
$PathStatusA = "~\Noyes Air Conditioning\Service - Technicians"
$PathStatusI = "~\Noyes Air Conditioning\Service - Technicians\Inactive sites"
$PathStatusU = "~\Noyes Air Conditioning\Service - Technicians\Unused sites"
$SiteCodes = @()
$directories = @()
$directories = Get-ChildItem -Path $PathStatusA
$directories += Get-ChildItem -Path $PathStatusI
$directories += Get-ChildItem -Path $PathStatusU

# Begin direcotry processing.
ForEach ($d in $directories) {
    # Isolate site codes.
    $SiteCode = $d.Name.Split("-").Trim()
    $SiteCode = $SiteCode[1]
    $SiteCodes += $SiteCode
    # Determine site status with SQL query.
    $query = @"
        SELECT *
        FROM [dbo].[Z_EKTO_REFERENCE_SITE_MASTER]
        WHERE Site_Code = '$SiteCode'
"@
    $result = Invoke-Sqlcmd -Query $query -ServerInstance "spectrum.nacgroup.com" -Database "Forefront"
    $SiteStatus = $result.Status

    # Conditional checks begin.
    if ($SiteStatus -match "A") {
        # For active sites: search inactive and unused dirs for folders. If 
        # present, move the existing folder, else create a new folder tree.
        if (Test-Path "$PathStatusA\*$SiteCode") { Write-Host "The directory for $SiteCode already exists." }
        elseif (Test-Path "$PathStatusI\*$SiteCode") { Move-Item -Path "$PathStatusI\*$SiteCode" -Destination "$PathStatusA\" }
        elseif (Test-Path "$PathStatusU\*$SiteCode") { Move-Item -Path "$PathStatusU\*$SiteCode" -Destination "$PathStatusA\" }
        else { 
            # Make-ServiceDirectories.ps1
            $SiteDir = $result.Site_Name.Trim() + " - " + $result.Site_Code.Trim()
            Write-Host "Creating directory tree for $SiteDir."
            New-Item -Type Directory -Name $SiteDir -Path "$PathStatusA"
			New-Item -Type Directory -Name "Quoted projects" -Path "$PathStatusA\$SiteDir"
			New-Item -Type Directory -Name "Scheduled maintenance reports" -Path "$PathStatusA\$SiteDir"
			New-Item -Type Directory -Name "Service contract" -Path "$PathStatusA\$SiteDir"
			New-Item -Type Directory -Name "Work order reports" -Path "$PathStatusA\$SiteDir"
            if (Test-Path "$PathStatusA\$SiteDir\Quoted projects") { 
                New-Item -Type Directory -Name "Daily job site reports" -Path "$PathStatusA\$SiteDir\Quoted projects"
                }
            else {
                Start-Sleep -Seconds 3
                New-Item -Type Directory -Name "Daily job site reports" -Path "$PathStatusA\$SiteDir\Quoted projects"
                }
            }
        }
    elseif ($SiteStatus -match "I") {
        # For inactive sites.
        if (Test-Path "$PathStatusA\*$SiteCode") { Move-Item -Path "$PathStatusA\*$SiteCode" -Destination "$PathStatusI" }
        elseif (Test-Path "$PathStatusI\*$SiteCode") { Write-Host "The directory for $SiteCode already exists." }
        elseif (Test-Path "$PathStatusU\*$SiteCode") { Move-Item -Path "$PathStatusU\*$SiteCode" -Destination "$PathStatusI" }
        else { # Make-ServiceDirectories.ps1
            $SiteDir = $result.Site_Name.Trim() + " - " + $result.Site_Code.Trim()
            Write-Host "Creating directory tree for $SiteDir."
            New-Item -Type Directory -Name $SiteDir -Path "$PathStatusI"
			New-Item -Type Directory -Name "Quoted projects" -Path "$PathStatusI\$SiteDir"
			New-Item -Type Directory -Name "Scheduled maintenance reports" -Path "$PathStatusI\$SiteDir"
			New-Item -Type Directory -Name "Service contract" -Path "$PathStatusI\$SiteDir"
			New-Item -Type Directory -Name "Work order reports" -Path "$PathStatusI\$SiteDir"
            if (Test-Path "$PathStatusI\$SiteDir\Quoted projects") { 
                New-Item -Type Directory -Name "Daily job site reports" -Path "$PathStatusI\$SiteDir\Quoted projects"
                }
            else {
                Start-Sleep -Seconds 3
                New-Item -Type Directory -Name "Daily job site reports" -Path "$PathStatusI\$SiteDir\Quoted projects"
                }
            }
        }
    elseif ($SiteStatus -match "U") {
        # For unused sites.
        if (Test-Path "$PathStatusA\*$SiteCode") { Move-Item -Path "$PathStatusA\*$SiteCode" -Destination "$PathStatusU" }
        elseif (Test-Path "$PathStatusI\*$SiteCode") { Move-Item -Path "$PathStatusI\*$SiteCode" -Destination "$PathStatusU" }
        elseif (Test-Path "$PathStatusU\*$SiteCode") { Write-Host "The directory for $SiteCode already exists." }
        else { # Make-ServiceDirectories.ps1
            $SiteDir = $result.Site_Name.Trim() + " - " + $result.Site_Code.Trim()
            Write-Host "Creating directory tree for $SiteDir."
            New-Item -Type Directory -Name $SiteDir -Path $PathStatusU
			New-Item -Type Directory -Name "Quoted projects" -Path "$PathStatusU\$SiteDir"
			New-Item -Type Directory -Name "Scheduled maintenance reports" -Path "$PathStatusU\$SiteDir"
			New-Item -Type Directory -Name "Service contract" -Path "$PathStatusU\$SiteDir"
			New-Item -Type Directory -Name "Work order reports" -Path "$PathStatusU\$SiteDir"
            if (Test-Path "$PathStatusU\$SiteDir\Quoted projects") { 
                New-Item -Type Directory -Name "Daily job site reports" -Path "$PathStatusU\$SiteDir\Quoted projects"
                }
            else {
                Start-Sleep -Seconds 3
                New-Item -Type Directory -Name "Daily job site reports" -Path "$PathStatusU\$SiteDir\Quoted projects"
                }
            }
        }
    else {
        # Error handling.

        }

    }