param(
	[string]$apiKey,
	[string]$octopusUrl,
	[string]$octopusProjectId,
	[string]$nugetPackageVersions	# needs to be a JSON array of {StepName, Version} objects. If not provided, we'll try to get the latest
)
$octopusHeader =  @{ "X-Octopus-ApiKey" = $apiKey }
if (-not $octopusUrl.EndsWith("/")) {$octopusUrl += "/" }

# Make an Octopus request
function Get-OctopusWebRequest ($relativeUrl) {
	return Invoke-WebRequest -Uri "$octopusUrl$relativeUrl" -Headers $octopusHeader | ConvertFrom-Json
}


# Get the selected packages we need
function Get-SelectedPackages($deploymentTemplate, $packageVersions) {
	if ($packageVersions -EQ $null) {
		$packageVersions = @()
	}
	$deploymentTemplate.Packages | ForEach-Object {
		$thisStep = $_
		Write-Verbose "Evaluating package version for step '$($thisStep.StepName)' which uses package '$($thisStep.NugetPackageId)'"
		$matching = $packageVersions | Where-Object -Property StepName -EQ -Value $thisStep.StepName
		if ($matching.Count -eq 0) {
			# the version hasn't been specified
			if ($thisStep.NugetFeedId -ne "feeds-builtin") {
				Write-Warning "  No version specified for step '$($thisStep.StepName)' and as it doesn't use the built-in feed, we can't reliably get the latest version :("
				
				$matchingFeeds = Get-OctopusWebRequest "api/feeds/$($thisStep.NugetFeedId)/packages?packageId=$($thisStep.NugetPackageId)&partialMatch=False&take=1"
				#$matchingFeeds = Invoke-WebRequest -Uri "$octoBaseUrl/api/feeds/$($thisStep.NugetFeedId)/packages?packageId=$($thisStep.NugetPackageId)&partialMatch=False&take=1" -Headers $octopusHeader | ConvertFrom-Json
				if ($matchingFeeds.Count -eq 1) {
					Write-Verbose "  Step '$($thisStep.StepName)' will use discovered version $($matchingFeeds[0].Version) on the assumption it's the latest"
					$packageVersions += @{
						StepName = $thisStep.StepName
						Version = $matchingFeeds[0].Version
					}
				} else {
					Write-Warning "Couldn't find the latest version for package required in step '$($thisStep.StepName)'"
					Write-Verbose "  Tried calling api/feeds/$($thisStep.NugetFeedId)/packages?packageId=$($thisStep.NugetPackageId)&partialMatch=False&take=1"
					exit 1
				}
			} else {
				Write-Verbose "  No version specified for step '$($thisStep.StepName)', but as it uses the built-in feed, we should be able to safely get the latest :)"
				
				# retrieve the latest version from Octopus
				$matchingFeeds = Get-OctopusWebRequest "api/packages?latest=true&filter=$($thisStep.NugetPackageId)"
				#$matchingFeeds = Invoke-WebRequest -Uri "$octoBaseUrl/api/packages?latest=true&filter=$($thisStep.NugetPackageId)" -Headers $octopusHeader | ConvertFrom-Json
				$latestPackage = $matchingFeeds.Items | Where-Object -Property NugetPackageId -EQ -Value $thisStep.NuGetPackageId
				if ($latestPackage) {
					Write-Verbose "  Step '$($thisStep.StepName)' will use discovered latest version $($latestPackage.Version)"
					$packageVersions += @{
						StepName = $thisStep.StepName
						Version = $latestPackage.Version
					}
				} else {
					Write-Warning "Couldn't find the latest version for package required in step '$($thisStep.StepName)'"
					Write-Verbose "  Tried calling api/packages?latest=true&filter=$($thisStep.NugetPackageId)"
					exit 1
				}
			}
		} else {
			Write-Verbose "  Step '$($thisStep.StepName)' will use specified version $($matching.Version)"
		}
	}
	
	return $packageVersions
}

### Execution Starts Here ###

# Get the Octopus Project
$octoProject = Get-OctopusWebRequest "api/projects/$octopusProjectId"
$resolvedProjectId = $octoProject.Id
$templateUrl = $octoProject.Links.DeploymentProcess + "/template"
	
# Get deployment template to get next version number
$deploymentTemplate = Get-OctopusWebRequest $templateUrl 

# Get package versions
if (-not [string]::IsNullOrEmpty($nugetPackageVersions))
{
	$packageVersions = $nugetPackageVersions | ConvertFrom-Json
}
$packageVersions = Get-SelectedPackages $deploymentTemplate $packageVersions

#todo: Release Notes

# Create Release
$releaseBody = 
@{
	ProjectId = $resolvedProjectId
	SelectedPackages = [array]$packageVersions
	Version = $deploymentTemplate.NextVersionIncrement
	
} | ConvertTo-Json

Write-Output "Calling $octopusUrl$releasesUrl with: "
Write-Output ($releaseBody | ConvertTo-Json)

$release = Invoke-WebRequest -Uri "$octopusUrl/api/releases" -Method Post -Headers $octopusHeader -Body $releaseBody | ConvertFrom-Json

Write-Output "Successfully created Release: $($release.Id)"