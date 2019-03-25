# Resource Manager

Resource Manager is a tool for creating and managing user resources. It can currently:
* Create mailboxes - both on-prem and in Office 365
* Enable Office 365 sync and set license
* Create SAML identities

Just like Lifecycle Manager, Resource Manager is built specifically for Kungsbacka kommun and is not meant for general use.

Tasks are stored as JSON in attribute carLicense on each user object processed by Resource manager. This attribute was chosen
because it's not indexed and replicated to neither Global Catalog nor Azure AD. [Kungsbacka.AccountTasks](https://github.com/Kungsbacka/Kungsbacka.AccountTasks)
can be used to generate task JSON or parse task JSON into objects.
