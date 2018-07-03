## Phishing Intelligence Engine Outlook Add-In (Button)
<img src="https://user-images.githubusercontent.com/16614015/38744998-92b3109e-3f00-11e8-89bb-b6b8ee3d27a4.png" width="125px" alt="PIE">

```
Phishing Intelligence Engine Outlook Add-In (Button)
LogRhythm Strategic Integrations Team
zack . rowland@logrhythm . com
v2.0 -- April, 2018
```
Copyright 2018 LogRhythm Inc. - See licensing details below

## [About]

Phishing Intelligence Engine companion Microsoft Outlook add-in that allows for easy reporting of suspected phishing attacks with one click.

```diff
- This add-in/framework is not officially supported by LogRhythm - use at your own risk!
```

## [Additional Information]

Currently, the add-in is only compatible with Microsoft Windows Outlook applications. The add-in has been tested and is known to be compatible with:
* Microsoft Outlook 2013 &#40;Windows&#41;
* Microsoft Outlook 2016 &#40;Windows&#41;

## [Build Procedure]

The project must be built using one of the following Visual Studio editions:
* Visual Studio 2017 Professional
* Visual Studio 2017 Enterprise

Visual Studio Community edition unfortunately does not include the necessary Microsoft Office Interop assemblies that the add-in requires.

Prior to building the project, it is strongly recommended that you have a digital code signing certificate (that is trusted by your organization) installed on the build system. Properly signing the add-in at build time greatly reduces the complexity of deploying the add-in to target systems/end-users by allowing for scriptable/silent installation (with no user intervention necessary). To build the add-in (and sign the add-in at build time):
1. With the project open in Visual Studio, right click the "PIEButton" project in "Solution Explorer" and select "Properties"
2. Select the "Signing" tab on the left
3. Check the "Sign the ClickOnce manifests" checkbox
4. If the code signing certificate has already been added to the Windows certifcate store, click the "Select from Store" button and select the appropriate certificate
5. If the code signing certificate is available as a file, select "Select from File" and navigate to the certificate location in the pop-up window
6. In "Solution Explorer", double click the "app.config" file to open it in the VS editor
7. Change the "ReportTargetEmail" value entry to the target e-mail address of your choice. This address will receive submitted phishing reports.
8. Once the above steps have been completed, click "File" > "Save All" to make sure the changes are saved.
9. At the top menu bar, select "Build" > "Build Solution"

Once the solution has finished building, the add-in will appear in the project "bin/debug" or "bin/release" folder (depending on whether you've selected the debug or release build). If the add-in has been built and properly signed at build time, it can be installed by simply double clicking the "PIEButton.vsto" installer file (after copying the add-in folder to a target machine). Please see the "Install and Usage" section below for detailed installation/deployment options.

## [Install and Usage]

Once the add-in has been built, it can be deployed to target workstations/hosts "by hand" using Windows Explorer, or programmatically using the Windows CLI (the CLI install/uninstall procedures should be compatible with most mass-deployment methodologies).
### Add-In Deployment via Windows Explorer/GUI
#### Install
1. Copy add-in folder to target workstation/host
2. Open add-in folder in Windows Explorer, double-click "PIEButton.vsto"
3. Installer window opens, select "Install" button
#### Uninstall
1. In Windows Explorer, click the "Start" menu/button
2. Select "Control Panel" > "Programs" > "Programs and Features"
3. Scroll through the program list to find "PIEButton"
4. Click the "PIEButton" entry to select it
5. Click the "Uninstall" button


### Add-In Deployment via Windows CLI
All CLI deployment operations utilize an existing Microsoft Office helper program, "VSTOInstaller.exe". This program should be present by default (on any host that has the Microsoft Office software suite installed) at the following location:
```
C:\Program Files (x86)\Common Files\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe
```
#### Install
`VSTOInstaller.exe /install \\C:\Full\path\to\PIEButton.vsto`
#### Uninstall
`VSTOInstaller.exe /uninstall \\C:\Full\path\to\PIEButton.vsto`
#### Silent Install
`VSTOInstaller.exe /install \\C:\Full\path\to\PIEButton.vsto /silent`
#### Silent Uninstall
`VSTOInstaller.exe /uninstall \\C:\Full\path\to\PIEButton.vsto /silent`

None of the above CLI install/uninstall operations should require administrator privileges if the add-in was properly signed during the build procedure, however the silent install will silently fail if the add-in's digital code signing certificate/publisher has not been added to the target host's "Trusted Publishers" certificate store. It's also worth mentioning that the add-in folder/files does not *have* to be copied locally to a target workstation to use the above commands; the VSTO installer supports add-ins located on a network share/shared storage. The VSTO installer *is*, however, a little bit particular in how the add-in path is specified (the `\\` prefix in path examples above is required even in the case of a local path).

## [License]

Copyright 2018 LogRhythm Inc.   

C# code is Licensed under the MIT License. See LICENSE file in the project root for full license information.

LogRhythm sample code (C# code) is licensed pursuant to the LogRhythm End User License Agreement located at [https://logrhythm.com/about/logrhythm-terms-and-conditions/](https://logrhythm.com/about/logrhythm-terms-and-conditions/) (“License Agreement”) and by downloading and using this content you agree to the terms and conditions of the License Agreement unless you have a separate signed end user license agreement with LogRhythm in which case that signed agreement shall govern your licensed use of this content. For purposes of the applicable end user license agreement, this content constitutes LogRhythm Software.
