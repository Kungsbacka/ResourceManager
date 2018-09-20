$Script:TaskDefinitions = @{
    EnableMailbox = @{
        Command = 'Enable-KBAOnpremMailbox'
        Initializer = 'Connect-KBAExchangeOnprem'
        Parameters = @('Type')
        OptionalParameters = @()
    }
    EnableRemoteMailbox = @{
        Command = 'Enable-KBAOnpremRemoteMailbox'
        Initializer = 'Connect-KBAExchangeOnprem'
        Parameters = @()
        OptionalParameters = @()
    }
    DisableMailbox = @{
        Command = 'Disable-KBAOnpremMailbox'
        Initializer = 'Connect-KBAExchangeOnprem'
        Parameters = @()
        OptionalParameters = @()
    }
    ConnectMailbox = @{
        Command = 'Connect-KBAOnpremMailbox'
        Initializer = 'Connect-KBAExchangeOnprem'
        Parameters = @()
        OptionalParameters = @()
    }
    ConfigureMailbox = @{
        Command = 'Set-KBAOnpremMailbox'
        Initializer = 'Connect-KBAExchangeOnprem'
        Parameters = @('Type')
        OptionalParameters = @()
    }
    ConfigureRemoteMailbox = @{
        Command = 'Set-KBAOnpremRemoteMailbox'
        Initializer = 'Connect-KBAExchangeOnprem'
        Parameters = @()
        OptionalParameters = @()
    }
    ConfigureOwa = @{
        Command = 'Set-KBAOnpremOwa'
        Initializer = 'Connect-KBAExchangeOnprem'
        Parameters = @('Type')
        OptionalParameters = @()
    }
    ConfigureCalendar = @{
        Command = 'Set-KBAOnpremCalendar'
        Initializer = 'Connect-KBAExchangeOnprem'
        Parameters = @('Type')
        OptionalParameters = @()
    }
    ConfigureMessage = @{
        Command = 'Set-KBAOnpremMailboxMessageConfiguration'
        Initializer = 'Connect-KBAExchangeOnprem'
        Parameters = @()
        OptionalParameters = @()
    }
    CleanupMailbox = @{
        Command = 'Cleanup-KBAOnpremMailbox'
        Initializer = 'Connect-KBAExchangeOnprem'
        Parameters = @()
        OptionalParameters = @()
    }
    ConfigureMailboxAutoReplyTask = @{
        Command = 'Set-KBAOnpremMailboxAutoReplyState'
        Initializer = 'Connect-KBAExchangeOnprem'
        Parameters = @('Enabled')
        OptionalParameters = @('Message')
    }
    SendWelcomeMail = @{
        Command = 'Send-KBAOnpremWelcomeMail'
        Initializer = $null
        Parameters = @('Type')
        OptionalParameters = @()
    }
    ConfigureOnlineMailbox = @{
        Command = 'Set-KBAOnlineMailbox'
        Initializer = 'Connect-KBAExchangeOnline'
        Parameters = @('Type')
        OptionalParameters = @()
    }
    ConfigureOnlineOwa = @{
        Command = 'Set-KBAOnlineOwa'
        Initializer = 'Connect-KBAExchangeOnline'
        Parameters = @()
        OptionalParameters = @()
    }
    HomeFolder = @{
        Command = 'New-HomeFolder'
        Initializer = $null
        Parameters = @('Path')
        OptionalParameters = @()
    }
    SamlId = @{
        Command = 'Set-SamlId'
        Initializer = $null
        Parameters = @()
        OptionalParameters = @()
    }
    MsolEnableSync = @{
        Command = 'Enable-KBAMsolSync'
        Initializer = $null
        Parameters = @()
        OptionalParameters = @()
    }
    MsolLicense = @{
        Command = 'Set-KBAMsolUserLicense'
        Initializer = 'Connect-KBAAzureAD'
        Parameters = @('License')
        OptionalParameters = @()
    }
    MsolRemoveLicense = @{
        Command = 'Remove-KBAMsolUserLicense'
        Initializer = 'Connect-KBAAzureAD'
        Parameters = @()
        OptionalParameters = @()
    }
    MsolRestoreLicense = @{
        Command = 'Restore-KBAMsolUserLicense'
        Initializer = 'Connect-KBAAzureAD'
        Parameters = @()
        OptionalParameters = @()
    }
    MsolLicenseGroup = @{
        Command = 'Set-LicenseGroupMembership'
        Initializer = $null
        Parameters = @('LicenseGroups')
        OptionalParameters = @('SkipSyncCheck')
        
    }
    EnableCSUser = @{
        Command = 'Enable-KBAOnpremCSUser'
        Initializer = 'Import-KBASkypeOnpremModule'
        Parameters = @()
        OptionalParameters = @()
    }
    GrantCSConferencingPolicy = @{
        Command = 'Grant-KBAOnpremCSConferencingPolicy'
        Initializer = 'Import-KBASkypeOnpremModule'
        Parameters = @('ConferencingPolicy')
        OptionalParameters = @()
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
