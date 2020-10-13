<## -- Sort-SharePointFiles --

This script moves files from the unsorted directory in the new service library
to the final location in the service library. It identifies the correct work 
number, looks that up against SQL, pulls job information from SQL about that 
job, and sorts the file with that information.

This script also tests if the files already exist. If not, the file tree is 
created before the file is put into place.

#>

# Declare the standard variables. These assume that the script is run from a computer with the libraries 
# synchronized.
$SourcePath = "~\Noyes Air Conditioning\Service - GoCanvas work reports (unsorted)\"
$DestinationPath = "~\Noyes Air Conditioning\Service - Technicians"


# Pull the list of current work orders and job numbers from the Spectrum database.
# $query = @"
#     SELECT * 
#     FROM [dbo].[Z_EKTO_REFERENCE_MERGED_JOBS_AND_WORK_ORDERS]
# "@
# $results = Invoke-Sqlcmd -Query $query -ServerInstance "spectrum.nacgroup.com" -Database "Forefront"



# Pull all file names for GoCanvas reports into a variable. This again assumes the script is run 
# from a computer synchronized with document library holding unsorted GoCanvas reports.
$files = Get-ChildItem -Path $SourcePath -File
$WorkNumbers = @()
$WorkSites = @()
$WorkReports = @()

# Processing each file.
ForEach ($f in $files) {
    # Read the file name and split it into various reference variables.
    $WorkNumber = $f.Name.Split("-")[0]
    $WorkNumbers += $WorkNumber
    $WorkSite = $f.Name.Split("-")[1]
    $WorkSites += $WorkSite
    $workReport = $f.Name.Split("-")[2]
    $WorkReports += $WorkReport
    # Using the work order or job number find the corresponding information from Spectrum.
    $query = @"
        SELECT * 
        FROM [dbo].[Z_EKTO_REFERENCE_MERGED_JOBS_AND_WORK_ORDERS]
        WHERE Job_Number = '$WorkNumber'
"@
    $result = Invoke-Sqlcmd -Query $query -ServerInstance "spectrum.nacgroup.com" -Database "Forefront"
    # User the resulting data to get the identifier for the site directory.
    $SiteCode = $result.Site_Code.Trim()
    # Search for the site directory.
    if (Test-Path "$DestinationPath\*$SiteCode") {
        # If the search is successful save the absolute path of the site directory.
        $Path = Get-ChildItem -Path "$DestinationPath\*$SiteCode"
        if ($WorkNumber.Length -eq 10) {
            # Reports with 10-digit job numbers go directly to the associated job folder.
            if (Test-Path "$Path\Quoted projects\$WorkNumber*") {
		$ProjectDirName = $(Get-Item -Path "$Path\Quoted projects\$WorkNumber*").Name
                Copy-Item "$SourcePath\$f" -Destination "$Path\Quoted projects\$ProjectDirName"
		Start-Sleep -Seconds 3
		if (Test-Path "$Path\Quoted projects\$ProjectDirName\$f") {
			Write-Host "I moved $f to $Path\Quoted projects\$WorkNumber."
			Move-Item "$SourcePath\$f" -Destination "$SourcePath\Sorted"
			}
                }
            elseif {
                # If no job folder exists create one and move the report.
                $WorkDescription = $result.Description.Trim()
                New-Item -Type Directory -Path "$Path\Quoted projects\" -Name "$WorkNumber - $WorkDescription" | Out-Null
                Write-Host "I created a new directory ($WorkNumber - $WorkDescription)"
		Start-Sleep -Seconds 3
                Copy-Item "$SourcePath\$f" -Destination "$Path\Quoted projects\$WorkNumber*"
		if (Test-Path "$Path\Quoted projects\$ProjectDirName\$f") {
			Write-Host "I moved $f to $Path\Quoted projects\$WorkNumber."
			Move-Item "$SourcePath\$f" -Destination "$SourcePath\Sorted"
			}
                }
	    else {
          	 # Error handling if job number isn't 5- or 10-digits long.
                 Write-Host "I can't write $f anywhere. I'll move it to the `"Needs review`" directory."
	         Move-Item "$SourcePath\$f" -Destination "$SourcePath\Needs review"
                 }
            }
        if ($WorkNumber.Length -lt 10) {
            # Reports with 5-digit job numbers can land in several different directories based on the 
            # report type.
            if ($WorkReport -like "Work Order*") {$ReportDir = "Work order reports"}
            elseif ($WorkReport -like "Scheduled Maintenance*") {$ReportDir = "Scheduled maintenance reports"}
            elseif ($WorkReport -like "Scheduled PM*") {$ReportDir = "Scheduled maintenance reports"}
            elseif ($WorkReport -like "Chiller Service*") {$ReportDir = "Scheduled maintenance reports"}
            elseif ($WorkReport -like "Chiller Overhaul*") {$ReportDir = "Scheduled maintenance reports"}
            elseif ($WorkReport -like "Plant Engineering*") {$ReportDir = "Scheduled maintenance reports"}
            else {$ReportDir = "Scheduled maintenance reports"}
            Copy-Item "$SourcePath\$f" -Destination "$Path\$ReportDir"
            if (Test-Path "$Path\$ReportDir\$f") { Move-Item "$SourcePath\$f" -Destination "$SourcePath\Sorted" }
            Write-Host "I moved $f to $Path\$ReportDir."
            }
        else {
            # Error handling if job number isn't 5- or 10-digits long.
            Write-Host "I can't write $f anywhere. I'll move it to the `"Needs review`" directory."
	    Move-Item "$SourcePath\$f" -Destination "$SourcePath\Needs review"
            }
        }
    }
