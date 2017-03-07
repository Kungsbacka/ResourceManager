# Resource Manager

Resource Manager is a tool for creating and managing user resources. It can currently:
* Create mailboxes - both on-prem and in Office 365
* Enable Office 365 sync and set licenses
* Enable Skype on-prem
* Create home folders
* Create SAML identities

Just like Lifecycle Manager, Resource Manager is built specifically for Kungsbacka kommun and is not meant for general use.

I had to make a few workarounds to get the scripts to run in a scheduled task under a Group Managed Service Account (gMSA).
Under these specific conditions the environment doesn't get properly initialized. Default modules are not loaded and 
environment variables are not set properly. You can find more information in
[this forum post](https://powershell.org/forums/topic/command-exist-and-does-not-exist-at-the-same-time/#post-58156) (powershell.org).

Note: One "gMSA workaround" must be done manually before deploying the script. Connect-AzureAD writes its logs to
%LOCALAPPDATA%\AzureAD\Powershell. If the environment it not properly initialied LOCALAPPDATA is going to point
to the public AppData folder (C:\Users\Public\AppData\Local). Make sure this location exists and is writable by
the gMSA. Alternatively you can create the subfolders and only give write permissions to the Powershell folder.
