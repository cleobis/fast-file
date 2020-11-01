# Fast File for Outlook

Fast filing plugin for Outlook.

Dev branch status: [![Build Status](https://img.shields.io/endpoint.svg?url=https%3A%2F%2Factions-badge.atrox.dev%2Fcleobis%2Ffast-file%2Fbadge%3Fref%3Ddev&style=flat)](https://actions-badge.atrox.dev/cleobis/fast-file/goto?ref=dev)

## Install
Download the latest release. Un-zip the archive and run setup.exe. It seems to be necessary to save the vsto file in the location you install from for later uninstallation.

If the installer is blocked due to a certificate error, you can:
* Right click on setup.exe > Properties > Digtial Signatures
* Click the signature, then click Details
* Click View Certificate
* Click Install Certificate...
* Install the certificate in the Trusted Root Certification Authorities store

## Upgrade
To upgrade, first uninstall, then install the new version.

## Uninstall
Uninstall by using the Uninstall Applications control panel.

If you prefer to uninstall manually or without user permissions, check the registry (regedit.exe) for HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall to get the uninstallation command. Then run the command in the terminal. ([ref](https://social.technet.microsoft.com/Forums/ie/en-US/8d920ece-614a-4109-afae-a408abbcbdf0/uninstalling-office-vsto-addins?forum=officeitproprevious)) For my installation, the command is:

    "C:\Program Files\Common Files\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe" /Uninstall file:///C:/Users/<username>/Downloads/vsto/QuickFile.vsto
