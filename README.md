# Resource Manager

Resource Manager is a tool for creating user resources. It can currently:
* Create mailboxes - both on-prem and in Office 365
* Enable Office 365 sync and set licenses
* Enable Skype on-prem
* Create home folders
* Create SAML identities

Just like Lifecycle Manager, Resource Manager is built specifically for Kungsbacka kommun and is not meant for general use.

I had to make a few workarounds to get the scripts to run as a scheduled task under a Group Managed Service Account (gMSA).
Under these specific conditions the environment doesn't get properly initialized. Default modules are not loaded and 
environment variables are not set properly. You can find more information in
[this forum post](https://powershell.org/forums/topic/command-exist-and-does-not-exist-at-the-same-time/#post-58156) (powershell.org).
