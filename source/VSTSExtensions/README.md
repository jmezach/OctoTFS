Team Build Preview Custom Steps
===============================

Custom Build Steps for [Team Build vNext](http://aka.ms/tfbuild)

These need to be uploaded to your Visual Studio or TFS instance before they can be used. To upload this custom task, use the `tfs-cli` tool available from [https://www.npmjs.com/package/tfx-cli](https://github.com/Microsoft/tfs-cli).

*Note: You can still use [OctoPack](http://docs.octopusdeploy.com/display/OD/Using+OctoPack) as part of your MSBuild task to package and push Nuget packages.*

[Create Octopus Release](CreateOctopusRelease)
----------------------
Creates a new Release in Octopus Deploy.

### Instructions for use

Detailed installation instructions can be found at [http://docs.octopusdeploy.com/display/OD/Use+the+Team+Foundation+Build+Custom+Task](http://docs.octopusdeploy.com/display/OD/Use+the+Team+Foundation+Build+Custom+Task)

There are a number of configuration options available.

 ![Configure Custom Build Step](../../img/tfsbuild-configurebuildstep.png)
 
 Options include:
 * **Octopus Deploy Server**:  Dropdown for selecting your Octopus Server (step 1)
 * **Project Name**:  The name of the project to create a release for
 * **Include Changeset comments**:  Whether to include changeset comments in the release notes
 * **Include Work Items**:  Whether to include work item titles in the release notes
 * **Deploy Release To**:  Optional environment to automatically deploy to (uses the [`--deployTo` argument in octo.exe](http://docs.octopusdeploy.com/display/OD/Creating+releases))
 * **Additional Octo.exe Arguments**:  Any additional [Octo.exe arguments](http://docs.octopusdeploy.com/display/OD/Creating+releases) to include
 
### Release Notes:

The *Release Notes* options, if selected, will result in nicely formatted release notes with deep links to Team Foundation Server or Visual Studio Online.

![Release Notes in Octopus Deploy Release](../../img/tfsbuild-releasenotes.png)

