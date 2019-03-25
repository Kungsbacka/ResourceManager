$Script:TaskDefinitions = @{
    EnableMailbox = @{
        Command = 'Enable-RmOnpremMailbox'
        Initializer = 'Connect-KBAExchangeOnprem'
        Parameters = @('Type')
        OptionalParameters = @()
    }
    EnableRemoteMailbox = @{
        Command = 'Enable-RmOnpremRemoteMailbox'
        Initializer = 'Connect-KBAExchangeOnprem'
        Parameters = @()
        OptionalParameters = @()
    }
    DisableMailbox = @{
        Command = 'Disable-RmOnpremMailbox'
        Initializer = 'Connect-KBAExchangeOnprem'
        Parameters = @()
        OptionalParameters = @()
    }
    ConnectMailbox = @{
        Command = 'Connect-RmOnpremMailbox'
        Initializer = 'Connect-KBAExchangeOnprem'
        Parameters = @()
        OptionalParameters = @()
    }
    ConfigureMailbox = @{
        Command = 'Set-RmOnpremMailbox'
        Initializer = 'Connect-KBAExchangeOnprem'
        Parameters = @('Type')
        OptionalParameters = @()
    }
    ConfigureRemoteMailbox = @{
        Command = 'Set-RmOnpremRemoteMailbox'
        Initializer = 'Connect-KBAExchangeOnprem'
        Parameters = @()
        OptionalParameters = @()
    }
    ConfigureOwa = @{
        Command = 'Set-RmOnpremOwa'
        Initializer = 'Connect-KBAExchangeOnprem'
        Parameters = @('Type')
        OptionalParameters = @()
    }
    ConfigureCalendar = @{
        Command = 'Set-RmOnpremCalendar'
        Initializer = 'Connect-KBAExchangeOnprem'
        Parameters = @('Type')
        OptionalParameters = @()
    }
    ConfigureMessage = @{
        Command = 'Set-RmOnpremMailboxMessageConfiguration'
        Initializer = 'Connect-KBAExchangeOnprem'
        Parameters = @()
        OptionalParameters = @()
    }
    CleanupMailbox = @{
        Command = 'Cleanup-RmOnpremMailbox'
        Initializer = 'Connect-KBAExchangeOnprem'
        Parameters = @()
        OptionalParameters = @()
    }
    ConfigureMailboxAutoReplyTask = @{
        Command = 'Set-RmOnpremMailboxAutoReplyState'
        Initializer = 'Connect-KBAExchangeOnprem'
        Parameters = @('Enabled')
        OptionalParameters = @('Message')
    }
    SendWelcomeMail = @{
        Command = 'Send-RmWelcomeMail'
        Initializer = $null
        Parameters = @('Type')
        OptionalParameters = @()
    }
    ConfigureOnlineMailbox = @{
        Command = 'Set-RmOnlineMailbox'
        Initializer = 'Connect-KBAExchangeOnline'
        Parameters = @('Type')
        OptionalParameters = @()
    }
    ConfigureOnlineOwa = @{
        Command = 'Set-RmOnlineOwa'
        Initializer = 'Connect-KBAExchangeOnline'
        Parameters = @()
        OptionalParameters = @()
    }
    SamlId = @{
        Command = 'Set-RmSamlId'
        Initializer = $null
        Parameters = @()
        OptionalParameters = @()
    }
    MsolEnableSync = @{
        Command = 'Enable-RmMsolSync'
        Initializer = $null
        Parameters = @()
        OptionalParameters = @()
    }
    MsolLicense = @{
        Command = 'Set-RmMsolUserLicense'
        Initializer = 'Connect-KBAAzureAD'
        Parameters = @('License')
        OptionalParameters = @()
    }
    MsolRemoveLicense = @{
        Command = 'Remove-RmMsolUserLicense'
        Initializer = 'Connect-KBAAzureAD'
        Parameters = @()
        OptionalParameters = @()
    }
    MsolRestoreLicense = @{
        Command = 'Restore-RmMsolUserLicense'
        Initializer = 'Connect-KBAAzureAD'
        Parameters = @()
        OptionalParameters = @()
    }
    MsolLicenseGroup = @{
        Command = 'Set-RmLicenseGroupMembership'
        Initializer = 'Update-LicenseGroupCache'
        Parameters = @('LicenseGroups')
        OptionalParameters = @('SkipSyncCheck','SkipDynamicGroupCheck')
    }
    MsolRemoveLicenseGroup = @{
        Command = 'Remove-RmLicenseGroupMembership'
        Initializer = 'Update-LicenseGroupCache'
        Parameters = @('LicenseGroups')
        OptionalParameters = @('SkipBaseLicenseCheck')
    }
    MsolRemoveAllLicenseGroup = @{
        Command = 'Remove-RmAllLicenseGroupMembership'
        Initializer = 'Update-LicenseGroupCache'
        Parameters = @()
        OptionalParameters = @('SkipStashLicense')
    }
    MsolRestoreLicenseGroup = @{
        Command = 'Restore-RmLicenseGroupMembership'
        Initializer = 'Update-LicenseGroupCache'
        Parameters = @()
        OptionalParameters = @('SkipSyncCheck')
    }
    Wait = @{
        Command = $null
        Initializer = $null
        Parameters = @('Minutes')
        OptionalParameters = @()
    }
    Debug = @{
        Command = 'Set-Content'
        Initializer = $null
        Parameters = @('Path')
        OptionalParameters = @()
    }
}
