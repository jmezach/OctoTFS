Import-Module "$PSScriptRoot\..\BuildTaskTestHelper.psm1"

Describe "Create Octopus Release" {
	InModuleScope BuildTaskTestHelper {
		$sut = "$PSScriptRoot\..\..\source\VSTSExtensions\OctopusBuildTasks\CreateOctopusRelease\task.json"
		$octoExe = Get-Item -Path $PSScriptRoot\..\..\source\VSTSExtensions\OctopusBuildTasks\CreateOctopusRelease\Octo.exe
		
		# Ensure no further modules are being loaded
		Mock Import-Module
		
		# Dummy function definitions so we can mock these
		Function Get-ServiceEndpoint { param($Name, $Context) }
		Function Invoke-Tool { param($Path, $Arguments) }
		
		# Set-up TFS variables
		$env:BUILD_STAGINGDIRECTORY = "TestDrive:\"
		$env:SYSTEM_TEAMFOUNDATIONCOLLECTIONURI = "http://dummtfsurl/dummytpc/"
		$env:SYSTEM_TEAMPROJECTID = "dummytpid"
		$env:BUILD_BUILDID = "611"
		
		# Setup a Mock for retrieving the Octopus Deploy service endpoint 
		Mock Get-ServiceEndpoint -Verifiable -ParameterFilter { $Name -eq "Octopus Deploy Server" } {
			return [PSCustomObject]@{
				"Url"="http://dummyoctopusurl";
				"Authorization"=[PSCustomObject]@{
					"Parameters"=[PSCustomObject]@{
						"Password"="API-1234567890"
					}
				}
			}
		}
		
		# Setup a Mock for retrieving the VSO/TFS service endpoint
		Mock Get-ServiceEndpoint -Verifiable -ParameterFilter { $Name -eq "SystemVssConnection" } {
				return [PSCustomObject]@{
					"Url"="http://dummtfsurl";
					"Authorization"=[PSCustomObject]@{
						"Parameters"=[PSCustomObject]@{
							"AccessToken"="DummyAccessToken"
						}
					}
				}
			}
		
		Function MockTfVersionControlChanges
		{
			$env:BUILD_REPOSITORY_PROVIDER = "TfsVersionControl"
			
			Mock Invoke-WebRequest -Verifiable -ParameterFilter { $Uri.AbsolutePath.EndsWith("/changes") } {
				$changes = [PSCustomObject]@{
					"value"=@(
						[PSCustomObject]@{
							"id"=1;
							"message"="My changeset";
							"location"="http://dummtfsurl/changesets/changeset/1";
							"author"=[PSCustomObject]@{
								"displayName"="John Doe"
							}
						})
				}
				
				return [PSCustomObject]@{ "Content"=(ConvertTo-Json $changes -Depth 3) }
			}
		}
		
		Function MockGitChanges
		{
			$env:BUILD_REPOSITORY_PROVIDER = "Git"
			
			Mock Invoke-WebRequest -Verifiable -ParameterFilter { $Uri.AbsolutePath.EndsWith("/changes") } {
				$changes = [PSCustomObject]@{
					"value"=@(
						[PSCustomObject]@{
							"id"=1;
							"message"="My commit";
							"location"="http://dummtfsurl/commits/commit/1";
							"author"=[PSCustomObject]@{
								"displayName"="John Doe"
							}
						})
				}
				
				return [PSCustomObject]@{ "Content"=(ConvertTo-Json $changes -Depth 3) }
			}
		}
		
		Function MockRelatedWorkItems
		{	
			Mock Invoke-WebRequest -Verifiable -ParameterFilter { $Uri.AbsolutePath.EndsWith("/workitems") } {
				$workitems = [PSCustomObject]@{
					"count"=1;
					"value"=@([PSCustomObject]@{ "id"=1 })
				}
				
				return [PSCustomObject]@{ "Content"=(ConvertTo-Json $workitems -Depth 3) }
			}
			
			Mock Invoke-WebRequest -Verifiable -ParameterFilter { $Uri.AbsolutePath.EndsWith("/wit/workItems") } {
				$workitemDetails = [PSCustomObject]@{
					"count"=1;
					"value"=@(
						[PSCustomObject]@{
							"id"=1;
							"fields"=[PSCustomObject]@{
								"System.Title"="Dummy work item";
								"System.State"="Done"
							}
						}
					)
				}
				
				return [PSCustomObject]@{ "Content"=(ConvertTo-Json $workitemDetails -Depth 3) }
			}
		}
		
		Context "Create simple release" {
			# Arrange
			Mock Invoke-Tool -Verifiable -ParameterFilter {
				$Path -eq $octoExe.FullName -and
				$Arguments -like "create-release*" -and
				$Arguments -like "*--project=`"Some Project`"*" -and
				$Arguments -like "*--server=http://dummyoctopusurl*" -and
				$Arguments -like "*--apikey=API-1234567890*"
			}
			
			# Act
			Invoke-BuildTask -TaskDefinitionFile $sut -- -ConnectedServiceName "Octopus Deploy Server" -ProjectName "Some Project"
			
			# Assert
			It "invokes Octo.exe with the correct project, server and API key" {
				Assert-VerifiableMocks
			}
		}

		Context "Including changeset comments" {
			# Arrange
			MockTfVersionControlChanges
			Mock Invoke-Tool -Verifiable -ParameterFilter {
				$Path -eq $octoExe.FullName -and
				$Arguments -like "*--project=`"Some Project`"*" -and
				$Arguments -like "*--releaseNotesFile=`"TestDrive:\*"
			}
			
			
			# Act
			Invoke-BuildTask -TaskDefinitionFile $sut -- -ConnectedServiceName "Octopus Deploy Server" -ProjectName "Some Project" -ChangesetCommentReleaseNotes "true"
			
			# Assert
			It "invokes Octo.exe with a release notes files" {
				Assert-VerifiableMocks
			}
			It "writes the changeset comments into the release notes file" {
				$releaseNotesFile = Get-ChildItem -Path TestDrive:\ -Filter *.md
				$releaseNotesFile.FullName | Should Contain "Changeset Comments:"
				$releaseNotesFile.FullName | Should Contain "My changeset"
				$releaseNotesFile.FullName | Should Contain "John Doe"
				$releaseNotesFile.FullName | Should Contain "$($env:SYSTEM_TEAMFOUNDATIONCOLLECTIONURI)$($env:SYSTEM_TEAMPROJECTID)/_versionControl/changeset/1"
			}
		}
		
		Context "Including commit messages" {
			# Arrange
			MockGitChanges
			Mock Invoke-Tool -Verifiable -ParameterFilter {
				$Path -eq $octoExe.FullName -and
				$Arguments -like "*--project=`"Some Project`"*" -and
				$Arguments -like "*--releaseNotesFile=`"TestDrive:\*"
			}
			
			# Act
			Invoke-BuildTask -TaskDefinitionFile $sut -- -ConnectedServiceName "Octopus Deploy Server" -ProjectName "Some Project" -ChangesetCommentReleaseNotes "true"
			
			# Assert
			It "invokes Octo.exe with a release notes files" {
				Assert-VerifiableMocks
			}
			It "writes the changeset comments into the release notes file" {
				$releaseNotesFile = Get-ChildItem -Path TestDrive:\ -Filter *.md
				$releaseNotesFile.FullName | Should Contain "Commit Messages:"
				$releaseNotesFile.FullName | Should Contain "My commit"
				$releaseNotesFile.FullName | Should Contain "John Doe"
				$releaseNotesFile.FullName | Should Contain "$($env:SYSTEM_TEAMFOUNDATIONCOLLECTIONURI)$($env:SYSTEM_TEAMPROJECTID)/_git/commits/commit/1"
			}
		}
		
		Context "Including work items" {
			# Arrange
			MockGitChanges
			MockRelatedWorkItems
			Mock Invoke-Tool -Verifiable -ParameterFilter {
				$Path -eq $octoExe.FullName -and
				$Arguments -like "*--project=`"Some Project`"*" -and
				$Arguments -like "*--releaseNotesFile=`"TestDrive:\*"
			}
			
			# Act
			Invoke-BuildTask -TaskDefinitionFile $sut -- -ConnectedServiceName "Octopus Deploy Server" -ProjectName "Some Project" -WorkItemReleaseNotes "true"
		
			# Assert
			It "invokes Octo.exe with a release notes files" {
				Assert-VerifiableMocks
			}
			It "writes the changeset comments into the release notes file" {
				$releaseNotesFile = Get-ChildItem -Path TestDrive:\ -Filter *.md
				$releaseNotesFile.FullName | Should Contain "Work Items:"
				$releaseNotesFile.FullName | Should Contain "[1]"
				$releaseNotesFile.FullName | Should Contain "Done"
				$releaseNotesFile.FullName | Should Contain "$($env:SYSTEM_TEAMFOUNDATIONCOLLECTIONURI)$($env:SYSTEM_TEAMPROJECTID)/_workitems/edit/1"
			}
		}
	}
}