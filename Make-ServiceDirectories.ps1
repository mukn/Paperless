<## -- Make-ServiceDirectories --

This script is used to create service directories for each site in Spectrum. It
sorts based on whether the site is active, inactive, or unused but does not check
existing directories beforehand.

Reviewed 26 May 2020.

#>

<## -- Reminders
## This section should only be run when run manually rather than in an automated fashion.
$UserResponse = Read-Host -Prompt "Is this running from the correct file path (y/n)?"
if ($UserResponse -eq "n") {
	$targetPath = Read-Host -Prompt "Input the desired file path."
	Test-Path -Path $targetPath
	Set-Location -Path $targetPath
	}

#>

## Functions
# Remove-SpecialChars - https://stackoverflow.com/q/30778098
function Remove-SpecialChars {
	param(
		[string]$InputString,
		[string]$Replacement = "",
		[string]$SpecialChars = "!#%&*<>?/\|{}"
	)
	$rePattern = ($specialChars.ToCharArray() | ForEach-Object { [regex]::Escape($_) }) -join "|"
	$inputString -replace $rePattern,$replacement
	}

# New-SharePointFolders
Function New-SharePointFolders {
	Param(
		[array]$InputDataTable,
		[string]$Path
	)
	$resultsArray = @()
	ForEach ($row in $InputDataTable) {
		$resultsArray += $row
	}
    $rootDirectory = Get-Location
	Set-Location -Path $Path
	$i = 0
	$currentDirs = Get-ChildItem -Path .\ | Select-Object -ExpandProperty Name
	ForEach ($item in $resultsArray.Site_Code) {
		# Remove blank space from Site_Name & Site_Code
		$name = $resultsArray[$i].Site_Name.Trim()
		$name = Remove-SpecialChars -InputString $name
		$code = $resultsArray[$i].Site_Code.Trim()
		$dir = "$name - $code"
		If ($currentDirs -match $code) {
			Write-Host "Match found in directory $dir."
		} Else 
		{
			# Create directories
			New-Item -Type Directory -Name $dir | Out-Null
			New-Item -Type Directory -Name "Quoted projects" -Path $dir | Out-Null
			New-Item -Type Directory -Name "Scheduled maintenance reports" -Path $dir | Out-Null
			New-Item -Type Directory -Name "Service contract" -Path $dir | Out-Null
			New-Item -Type Directory -Name "Work order reports" -Path $dir | Out-Null
			New-Item -Type Directory -Name "Daily job site reports" -Path "$dir\Quoted projects" | Out-Null
			Write-Host "Directory tree created under $dir."
		}
		$i++
	}
	Set-Location -Path $rootDirectory
}

## 
# Pull active sites from SQL and save to variable.
$query = @"
	SELECT * FROM [dbo].[Z_PAPERLESS_REFERENCE_SITE_MASTER]
	WHERE Status = 'A'
	ORDER BY Site_Code
"@
$results = Invoke-Sqlcmd -Query $query -ServerInstance spectrum.nacgroup.com -Database Forefront
New-SharePointFolders -InputDataTable $results -Path .\

# Pull inactive sites from SQL and save to variable.
$query = @"
	SELECT * FROM [dbo].[Z_PAPERLESS_REFERENCE_SITE_MASTER]
	WHERE Status = 'I'
	ORDER BY Site_Code
"@
$results = Invoke-Sqlcmd -Query $query -ServerInstance spectrum.nacgroup.com -Database Forefront
New-SharePointFolders -InputDataTable $results -Path ".\Inactive sites"

# Pull unused sites from SQL and save to variable.
$query = @"
	SELECT * FROM [dbo].[Z_PAPERLESS_REFERENCE_SITE_MASTER]
	WHERE Status = 'N'
	ORDER BY Site_Code
"@
$results = Invoke-Sqlcmd -Query $query -ServerInstance spectrum.nacgroup.com -Database Forefront
# New-SharePointFolders -InputDataTable $results -Path ".\Unused sites"
