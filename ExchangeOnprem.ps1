# Make all error terminating errors
$Global:ErrorActionPreference = 'Stop'
. "$PSScriptRoot\Config.ps1"

Enum TestMailboxResult
{
    None
    Onprem
    Remote
    Online
    Both
}

function Connect-KBAExchangeOnprem
{
    $server = $Script:Config.ExchangeOnprem.Servers | Get-Random
    Get-PSSession -Name 'KBAExchOnprem' -ErrorAction SilentlyContinue | Remove-PSSession
    $credential = New-Object -TypeName 'System.Management.Automation.PSCredential' -ArgumentList @(
        $Script:Config.ExchangeOnprem.User
        $Script:Config.ExchangeOnprem.Password | ConvertTo-SecureString
    )
    $params = @{
        Name = 'KBAExchOnprem'
        ConfigurationName = 'Microsoft.Exchange'
        ConnectionUri = "http://$server/PowerShell/"
        Authentication = 'Kerberos'
        Credential = $credential
    }
    try
    {
        $session = New-PSSession @params
        $params = @{
            Session = $session
            Prefix = 'Onprem'
            DisableNameChecking = $true
            CommandName = @(
                'Enable-Mailbox'
                'Enable-RemoteMailbox'
                'Get-Mailbox'
                'Get-RemoteMailbox'
                'Get-MailboxFolderStatistics'
                'Get-MailboxDatabase'
                'Set-Mailbox'
                'Set-RemoteMailbox'
                'Set-MailboxAutoReplyConfiguration'
                'Set-MailboxCalendarConfiguration'
                'Set-CasMailbox'
                'Set-MailboxRegionalConfiguration'
                'Set-MailboxFolderPermission'
                'Set-MailboxMessageConfiguration'
            )
        }
        Import-PSSession @params | Out-Null
    }
    catch
    {
        throw [pscustomobject]@{
            Target     = "http://$server/PowerShell/"
            Activity   = $MyInvocation.MyCommand.Name
            Reason     = 'New/Import-PSSession failed'
            Message    = $_.Exception.Message
            RetryCount = 50
            Delay      = 10
        }
    }
}

function Test-KBAOnpremMailbox
{
    [CmdLetBinding()]
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]$UserPrincipalName
    )
    process
    {
        $result = [TestMailboxResult]::None
        try
        {
            if (Get-OnpremMailbox -Identity $UserPrincipalName)
            {
                $result = [TestMailboxResult]::Onprem
            }
        }
        catch
        {
            if ($_.CategoryInfo.Reason -ne 'ManagementObjectNotFoundException')
            {
                throw
            }
        }
        try
        {
            if (Get-OnpremRemoteMailbox -Identity $UserPrincipalName)
            {
                if ($result -eq [TestMailboxResult]::None)
                {
                    $result = [TestMailboxResult]::Remote
                }
                else
                {
                    $result = [TestMailboxResult]::Both
                }
            }
        }
        catch
        {
            if ($_.CategoryInfo.Reason -ne 'ManagementObjectNotFoundException')
            {
                throw
            }
        }
        $result
    }
}

function Set-KBAOnpremOwa
{
    [CmdLetBinding()]
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]$UserPrincipalName,
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Employee', 'Student', 'Shared')]
        [string]$Type
    )
    process
    {
        try
        {
            $result = Test-KBAOnpremMailbox $UserPrincipalName
        }
        catch
        {
            throw [pscustomobject]@{
                Target     = $UserPrincipalName
                Activity   = $MyInvocation.MyCommand.Name
                Reason     = 'Test-KBAOnpremMailbox failed'
                Message    = $_.Exception.Message
                RetryCount = 3
                Delay      = 5
            }
        }
        if ($result -ne [TestMailboxResult]::Onprem)
        {
            throw [pscustomobject]@{
                Target     = $UserPrincipalName
                Activity   = $MyInvocation.MyCommand.Name
                Reason     = 'On-prem mailbox not found'
                Message    = 'An on-prem mailbox not found for target.'
                RetryCount = 0
                Delay      = 0
            }
        }
        $params = @{
            Identity = $UserPrincipalName
        }
        if ($Type -eq 'Student')
        {
            $params.OwaMailboxPolicy = $Script:Config.ExchangeOnprem.Owa.Student.OwaMailboxPolicy
        }
        else # Employee or Shared
        {
            $params.OwaMailboxPolicy = $Script:Config.ExchangeOnprem.Owa.OwaMailboxPolicy
        }
        try
        {
            Set-OnpremCasMailbox @params
        }
        catch
        {
            throw [pscustomobject]@{
                Target     = $UserPrincipalName
                Activity   = $MyInvocation.MyCommand.Name
                Reason     = 'Set-CasMailbox failed'
                Message    = $_.Exception.Message
                RetryCount = 3
                Delay      = 10
            }
        }
        $params = @{
            Identity = $UserPrincipalName
            DateFormat = $Script:Config.ExchangeOnprem.Owa.DateFormat
            TimeFormat = $Script:Config.ExchangeOnprem.Owa.TimeFormat
            TimeZone = $Script:Config.ExchangeOnprem.Owa.TimeZone
            Language = $Script:Config.ExchangeOnprem.Owa.Language
            LocalizeDefaultFolderName = $Script:Config.ExchangeOnprem.Owa.LocalizeDefaultFolderName
        }
        try
        {
            Set-OnpremMailboxRegionalConfiguration @params
        }
        catch
        {
            throw [pscustomobject]@{
                Target     = $UserPrincipalName
                Activity   = $MyInvocation.MyCommand.Name
                Reason     = 'Set-MailboxRegionalConfiguration failed'
                Message    = $_.Exception.Message
                RetryCount = 3
                Delay      = 10
            }
        }
    }
}

function Set-KBAOnpremCalendar
{
    [CmdLetBinding()]
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]$UserPrincipalName,
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Employee', 'Student', 'Shared')]
        [string]$Type
    )
    process
    {
        try
        {
            $result = Test-KBAOnpremMailbox $UserPrincipalName
        }
        catch
        {
            throw [pscustomobject]@{
                Target     = $UserPrincipalName
                Activity   = $MyInvocation.MyCommand.Name
                Reason     = 'Test-KBAOnpremMailbox failed'
                Message    = $_.Exception.Message
                RetryCount = 3
                Delay      = 5
            }
        }
        if ($result -ne [TestMailboxResult]::Onprem)
        {
            throw [pscustomobject]@{
                Target     = $UserPrincipalName
                Activity   = $MyInvocation.MyCommand.Name
                Reason     = 'On-prem mailbox not found'
                Message    = 'An on-prem mailbox not found for target.'
                RetryCount = 0
                Delay      = 0
            }
        }
        $calendarIdentities = @()
        $params = @{
            Identity = $UserPrincipalName
            FolderScope = 'Calendar'
        }
        try
        {
            $calendarFolders = Get-OnpremMailboxFolderStatistics @params
        }
        catch
        {
            throw [pscustomobject]@{
                Target     = $UserPrincipalName
                Activity   = $MyInvocation.MyCommand.Name
                Reason     = 'Get-MailboxFolderStatistics failed'
                Message    = $_.Exception.Message
                RetryCount = 3
                Delay      = 15
            }
        }
        if ($calendarFolders.Name -contains 'Kalender')
        {
            $calendarIdentities += '{0}:\Kalender' -f $UserPrincipalName
        }
        if ($calendarFolders.Name -contains 'Calendar')
        {
            $calendarIdentities += '{0}:\Calendar' -f $UserPrincipalName
        }
        if ($calendarIdentities.Count -eq 0)
        {
            throw [pscustomobject]@{
                Target     = $UserPrincipalName
                Activity   = $MyInvocation.MyCommand.Name
                Reason     = 'No calendar folder found'
                Message    = ''
                RetryCount = 3
                Delay      = 15
            }
        }
        $params = @{
            Identity = $UserPrincipalName
            WorkingHoursTimeZone = $Script:Config.ExchangeOnprem.Calendar.WorkingHoursTimeZone
            ShowWeekNumbers = $Script:Config.ExchangeOnprem.Calendar.ShowWeekNumbers
            WeekStartDay = $Script:Config.ExchangeOnprem.Calendar.WeekStartDay
            FirstWeekOfYear = $Script:Config.ExchangeOnprem.Calendar.FirstWeekOfYear
        }
        if ($Type -eq 'Student')
        {
            $params.WorkingHoursStartTime = $Script:Config.ExchangeOnprem.Calendar.Student.WorkingHoursStartTime
            $params.WorkingHoursEndTime = $Script:Config.ExchangeOnprem.Calendar.Student.WorkingHoursEndTime
        }
        else # Employee or Shared
        {
            $params.WorkingHoursStartTime = $Script:Config.ExchangeOnprem.Calendar.WorkingHoursStartTime
            $params.WorkingHoursEndTime = $Script:Config.ExchangeOnprem.Calendar.WorkingHoursEndTime
        }
        try
        {
            Set-OnpremMailboxCalendarConfiguration @params
        }
        catch
        {
            throw [pscustomobject]@{
                Target     = $UserPrincipalName
                Activity   = $MyInvocation.MyCommand.Name
                Reason     = 'Set-MailboxCalendarConfiguration failed'
                Message    = $_.Exception.Message
                RetryCount = 3
                Delay      = 15
            }
        }
        if ($Type -in 'Employee', 'Shared')
        {
            foreach($calendarIdentity in $calendarIdentities)
            {
                $params = @{
                    Identity = $calendarIdentity
                    User = 'Default'
                    AccessRights = $Script:Config.ExchangeOnprem.Calendar.DefaultCalendarPermission
                }
                try
                {
                    Set-OnpremMailboxFolderPermission @params
                }
                catch
                {
                    throw [pscustomobject]@{
                        Target     = $UserPrincipalName
                        Activity   = $MyInvocation.MyCommand.Name
                        Reason     = 'Set-MailboxFolderPermission failed'
                        Message    = $_.Exception.Message
                        RetryCount = 3
                        Delay      = 15
                    }
                }
            }
        }
    }
}

# The mailbox configuration is moved to Enable-KBAOnpremMailbox.
# This function is kept around to be able to run mailbox configuration
# as a separate task after the mailbox is created.
function Set-KBAOnpremMailbox
{
    [CmdLetBinding()]
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]$UserPrincipalName,
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Employee', 'Student', 'Shared')]
        [string]$Type
    )
    process
    {
        try
        {
            $result = Test-KBAOnpremMailbox $UserPrincipalName
        }
        catch
        {
            throw [pscustomobject]@{
                Target     = $UserPrincipalName
                Activity   = $MyInvocation.MyCommand.Name
                Reason     = 'Test-KBAOnpremMailbox failed'
                Message    = $_.Exception.Message
                RetryCount = 3
                Delay      = 5
            }
        }
        if ($result -ne [TestMailboxResult]::Onprem)
        {
            throw [pscustomobject]@{
                Target     = $UserPrincipalName
                Activity   = $MyInvocation.MyCommand.Name
                Reason     = 'On-prem mailbox not found'
                Message    = 'An on-prem mailbox not found for target.'
                RetryCount = 0
                Delay      = 0
            }
        }
        if ($Type -eq 'Student')
        {
            $params = @{
                Identity = $UserPrincipalName
                EmailAddressPolicyEnabled = $Script:Config.ExchangeOnprem.Mailbox.EmailAddressPolicyEnabled
                RetentionPolicy = $Script:Config.ExchangeOnprem.Mailbox.Student.RetentionPolicy
                AddressBookPolicy = $Script:Config.ExchangeOnprem.Mailbox.Student.AddressBookPolicy
                IssueWarningQuota = $Script:Config.ExchangeOnprem.Mailbox.Student.IssueWarningQuota
                ProhibitSendQuota = $Script:Config.ExchangeOnprem.Mailbox.Student.ProhibitSendQuota
                ProhibitSendReceiveQuota = $Script:Config.ExchangeOnprem.Mailbox.Student.ProhibitSendReceiveQuota
                UseDatabaseQuotaDefaults = $Script:Config.ExchangeOnprem.Mailbox.UseDatabaseQuotaDefaults
            }
        }
        else # Employee or Shared
        {
            $params = @{
                Identity = $UserPrincipalName
                EmailAddressPolicyEnabled = $Script:Config.ExchangeOnprem.Mailbox.EmailAddressPolicyEnabled
                RetentionPolicy = $Script:Config.ExchangeOnprem.Mailbox.RetentionPolicy
                AddressBookPolicy = $Script:Config.ExchangeOnprem.Mailbox.AddressBookPolicy
                IssueWarningQuota = $Script:Config.ExchangeOnprem.Mailbox.IssueWarningQuota
                ProhibitSendQuota = $Script:Config.ExchangeOnprem.Mailbox.ProhibitSendQuota
                ProhibitSendReceiveQuota = $Script:Config.ExchangeOnprem.Mailbox.ProhibitSendReceiveQuota
                UseDatabaseQuotaDefaults = $Script:Config.ExchangeOnprem.Mailbox.UseDatabaseQuotaDefaults
            }
        }
        try
        {
            Set-OnpremMailbox @params
        }
        catch
        {
            throw [pscustomobject]@{
                Target     = $UserPrincipalName
                Activity   = $MyInvocation.MyCommand.Name
                Reason     = 'Set-Mailbox failed'
                Message    = $_.Exception.Message
                RetryCount = 10
                Delay      = 3
            }
        }
    }
}

function Enable-KBAOnpremMailbox
{
    [CmdLetBinding()]
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]$UserPrincipalName,
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]$SamAccountName,
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Employee', 'Student', 'Shared')]
        [string]$Type
    )
    process
    {
        try
        {
            $result = Test-KBAOnpremMailbox $UserPrincipalName
        }
        catch
        {
            throw [pscustomobject]@{
                Target     = $UserPrincipalName
                Activity   = $MyInvocation.MyCommand.Name
                Reason     = 'Test-KBAOnpremMailbox failed'
                Message    = $_.Exception.Message
                RetryCount = 3
                Delay      = 5
            }
        }
        if ($result -eq [TestMailboxResult]::Onprem)
        {
            throw [pscustomobject]@{
                Target     = $UserPrincipalName
                Activity   = $MyInvocation.MyCommand.Name
                Reason     = 'Mailbox already exists'
                Message    = 'An on-prem mailbox already exists for target.'
                RetryCount = 0
                Delay      = 0
            }
        }
        if ($result -eq [TestMailboxResult]::Remote)
        {
            throw [pscustomobject]@{
                Target     = $UserPrincipalName
                Activity   = $MyInvocation.MyCommand.Name
                Reason     = 'Mailbox already exists'
                Message    = 'A remote mailbox already exsits for target.'
                RetryCount = 0
                Delay      = 0
            }
        }
        $params = @{
            Identity = $UserPrincipalName
            PrimarySmtpAddress = $UserPrincipalName
            Alias = $SamAccountName
        }
        if ($Type -eq 'Student')
        {
            try
            {
                $database =
                    Get-OnpremMailboxDatabase |
                    Where-Object -Property 'Name' -Like -Value 'S_DB*' |
                    Get-Random
            }
            catch
            {
                throw [pscustomobject]@{
                    Target     = $UserPrincipalName
                    Activity   = $MyInvocation.MyCommand.Name
                    Reason     = 'Get-MailboxDatabase failed'
                    Message    = $_.Exception.Message
                    RetryCount = 3
                    Delay      = 10
                }
            }
            $params.Database = $database.Name
        }
        try
        {
            Enable-OnpremMailbox @params | Out-Null
        }
        catch
        {
            throw [pscustomobject]@{
                Target     = $UserPrincipalName
                Activity   = $MyInvocation.MyCommand.Name
                Reason     = 'Enable-Mailbox failed'
                Message    = $_.Exception.Message
                RetryCount = 3
                Delay      = 10
            }
        }
        Start-Sleep -Seconds 5
        $retryCount = 0
        while (-not (Get-OnpremMailbox -Identity $UserPrincipalName -ErrorAction SilentlyContinue) -and $retryCount++ -lt 10)
        {
            Start-Sleep -Seconds 5
        }
        if ($Type -eq 'Student')
        {
            $params = @{
                Identity = $UserPrincipalName
                EmailAddressPolicyEnabled = $Script:Config.ExchangeOnprem.Mailbox.EmailAddressPolicyEnabled
                RetentionPolicy = $Script:Config.ExchangeOnprem.Mailbox.Student.RetentionPolicy
                AddressBookPolicy = $Script:Config.ExchangeOnprem.Mailbox.Student.AddressBookPolicy
                IssueWarningQuota = $Script:Config.ExchangeOnprem.Mailbox.Student.IssueWarningQuota
                ProhibitSendQuota = $Script:Config.ExchangeOnprem.Mailbox.Student.ProhibitSendQuota
                ProhibitSendReceiveQuota = $Script:Config.ExchangeOnprem.Mailbox.Student.ProhibitSendReceiveQuota
                UseDatabaseQuotaDefaults = $Script:Config.ExchangeOnprem.Mailbox.UseDatabaseQuotaDefaults
            }
        }
        else # Employee or Shared
        {
            $params = @{
                Identity = $UserPrincipalName
                EmailAddressPolicyEnabled = $Script:Config.ExchangeOnprem.Mailbox.EmailAddressPolicyEnabled
                RetentionPolicy = $Script:Config.ExchangeOnprem.Mailbox.RetentionPolicy
                AddressBookPolicy = $Script:Config.ExchangeOnprem.Mailbox.AddressBookPolicy
                IssueWarningQuota = $Script:Config.ExchangeOnprem.Mailbox.IssueWarningQuota
                ProhibitSendQuota = $Script:Config.ExchangeOnprem.Mailbox.ProhibitSendQuota
                ProhibitSendReceiveQuota = $Script:Config.ExchangeOnprem.Mailbox.ProhibitSendReceiveQuota
                UseDatabaseQuotaDefaults = $Script:Config.ExchangeOnprem.Mailbox.UseDatabaseQuotaDefaults
            }
        }
        $params.Add('EmailAddresses', @{add=$SamAccountName + '@' + $Script:Config.ExchangeOnprem.ExchangeOnlineMailDomain})
        try
        {
            Set-OnpremMailbox @params
        }
        catch
        {
            throw [pscustomobject]@{
                Target     = $UserPrincipalName
                Activity   = $MyInvocation.MyCommand.Name
                Reason     = 'Set-Mailbox failed'
                Message    = $_.Exception.Message
                RetryCount = 0
                Delay      = 0
            }
        }
    }
}

function Set-KBAOnpremRemoteMailbox
{
    [CmdLetBinding()]
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]$UserPrincipalName
    )
    process
    {
        try
        {
            $result = Test-KBAOnpremMailbox $UserPrincipalName
        }
        catch
        {
            throw [pscustomobject]@{
                Target     = $UserPrincipalName
                Activity   = $MyInvocation.MyCommand.Name
                Reason     = 'Test-KBAOnpremMailbox failed'
                Message    = $_.Exception.Message
                RetryCount = 3
                Delay      = 5
            }
        }
        if ($result -ne [TestMailboxResult]::Remote)
        {
            throw [pscustomobject]@{
                Target     = $UserPrincipalName
                Activity   = $MyInvocation.MyCommand.Name
                Reason     = 'Not a remote mailbox'
                Message    = 'This is not a remote mailbox.'
                RetryCount = 0
                Delay      = 0
            }
        }
        try
        {
            $mailbox = Get-OnpremRemoteMailbox -Identity $UserPrincipalName
        }
        catch
        {
            throw [pscustomobject]@{
                Target     = $UserPrincipalName
                Activity   = $MyInvocation.MyCommand.Name
                Reason     = 'Get-RemoteMailbox failed'
                Message    = $_.Exception.Message
                RetryCount = 3
                Delay      = 10
            }
        }
        if ($mailbox.RemoteRoutingAddress -cnotmatch "^SMTP:[^@]+@$($Script:Config.ExchangeOnprem.ExchangeOnlineMailDomain)$")
        {
            throw [pscustomobject]@{
                Target     = $UserPrincipalName
                Activity   = $MyInvocation.MyCommand.Name
                Reason     = 'RemoteRoutingAddress is invalid'
                Message    = 'The Remote Routing Address is not valid.'
                RetryCount = 0
                Delay      = 0
            }
        }
        if ($mailbox.RemoteRoutingAddress -notin $mailbox.EmailAddresses)
        {
            $rra = 'smtp:' + $mailbox.RemoteRoutingAddress.Substring(5)
            $params = @{
                Identity = $UserPrincipalName
                EmailAddresses = @{'Add' = $rra}
            }
            try
            {
                Set-OnpremRemoteMailbox @params
            }
            catch
            {
                throw [pscustomobject]@{
                    Target     = $UserPrincipalName
                    Activity   = $MyInvocation.MyCommand.Name
                    Reason     = 'Set-RemoteMailbox failed'
                    Message    = $_.Exception.Message
                    RetryCount = 3
                    Delay      = 10
                }
            }
        }
    }
}

function Enable-KBAOnpremRemoteMailbox
{
    [CmdLetBinding()]
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]$UserPrincipalName,
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]$SamAccountName
    )
    process
    {
        try
        {
            $result = Test-KBAOnpremMailbox $UserPrincipalName
        }
        catch
        {
            throw [pscustomobject]@{
                Target     = $UserPrincipalName
                Activity   = $MyInvocation.MyCommand.Name
                Reason     = 'Test-KBAOnpremMailbox failed'
                Message    = $_.Exception.Message
                RetryCount = 3
                Delay      = 5
            }
        }
        if ($result -eq [TestMailboxResult]::Onprem)
        {
            throw [pscustomobject]@{
                Target     = $UserPrincipalName
                Activity   = $MyInvocation.MyCommand.Name
                Reason     = 'Mailbox already exists'
                Message    = 'An on-prem mailbox already exists for target.'
                RetryCount = 0
                Delay      = 0
            }
        }
        if ($result -eq [TestMailboxResult]::Remote)
        {
            throw [pscustomobject]@{
                Target     = $UserPrincipalName
                Activity   = $MyInvocation.MyCommand.Name
                Reason     = 'Mailbox already exists'
                Message    = 'A remote mailbox already exsits for target.'
                RetryCount = 0
                Delay      = 0
            }
        }
        $params = @{
            Identity = $UserPrincipalName
            PrimarySmtpAddress = $UserPrincipalName
            Alias = $SamAccountName
            RemoteRoutingAddress = $SamAccountName + '@' +
                $Script:Config.ExchangeOnprem.ExchangeOnlineMailDomain
        }
        try
        {
            Enable-OnpremRemoteMailbox @params | Out-Null
        }
        catch
        {
            throw [pscustomobject]@{
                Target     = $UserPrincipalName
                Activity   = $MyInvocation.MyCommand.Name
                Reason     = 'Enable-RemoteMailbox failed'
                Message    = $_.Exception.Message
                RetryCount = 3
                Delay      = 10
            }
        }
		# Not sure why this is needed, but sometimes Enable-RemoteMailbox fails
		# if this is not present. Initially 10.
        Start-Sleep -Seconds 5
    }
}

function Set-KBAOnpremMailboxAutoReplyState
{
        [CmdLetBinding()]
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]$UserPrincipalName,
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [bool]$Enabled,
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string]$Message
        
    )
    process
    {
        try
        {
            $result = Test-KBAOnpremMailbox $UserPrincipalName
        }
        catch
        {
            throw [pscustomobject]@{
                Target     = $UserPrincipalName
                Activity   = $MyInvocation.MyCommand.Name
                Reason     = 'Test-KBAOnpremMailbox failed'
                Message    = $_.Exception.Message
                RetryCount = 3
                Delay      = 5
            }
        }
        if ($result -ne [TestMailboxResult]::Onprem)
        {
            throw [pscustomobject]@{
                Target     = $UserPrincipalName
                Activity   = $MyInvocation.MyCommand.Name
                Reason     = 'On-prem mailbox not found'
                Message    = 'An on-prem mailbox not found for target.'
                RetryCount = 0
                Delay      = 0
            }
        }
        if ($Enabled)
        {
            $state = 'Enabled'
        }
        else
        {
            $state = 'Disabled'
        }
        if ($Enabled -and $Message)
        {
            $internalAndExternalMessage = $Message
        }
        elseif ($Enabled)
        {
            $internalAndExternalMessage = $Script:Config.ExchangeOnprem.AutoReply.DefaultMessage
        }
        else
        {
            $internalAndExternalMessage = $null    
        }
        $params = @{
            Identity = $UserPrincipalName
            AutoReplyState = $state
            InternalMessage = $internalAndExternalMessage
            ExternalMessage = $internalAndExternalMessage
        }
        try
        {
            Set-OnpremMailboxAutoReplyConfiguration @params | Out-Null
        }
        catch
        {
            throw [pscustomobject]@{
                Target     = $UserPrincipalName
                Activity   = $MyInvocation.MyCommand.Name
                Reason     = 'Set-MailboxAutoReplyConfiguration failed'
                Message    = $_.Exception.Message
                RetryCount = 3
                Delay      = 10
            }
        }
    }
}

function Send-KBAOnpremWelcomeMail
{
    [CmdLetBinding()]
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]$UserPrincipalName,
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Employee', 'Student', 'Shared')]
        [string]$Type
    )
    process
    {
        if ($Type -ne 'Employee')
        {
            throw [pscustomobject]@{
              Target     = $UserPrincipalName
              Activity   = $MyInvocation.MyCommand.Name
              Reason     = 'Wrong type'
              Message    = 'Welcome mail is only sent to employees.'
              RetryCount = 0
              Delay      = 0
            }
        }
        # Use SmtpClient instead of Send-MailMessage since the latter
        # always tries to authenticate with default credentials. A gMSA is
        # not allowed to authenticate to our Exchange SMTP receive connector.
        $smtpClient = New-Object -TypeName 'System.Net.Mail.SmtpClient'
        $smtpClient.UseDefaultCredentials = $false
        $smtpClient.Host = $Script:Config.ExchangeOnprem.WelcomeMail.Server
        $msg = New-Object -TypeName 'System.Net.Mail.MailMessage'
        $msg.BodyEncoding = [System.Text.Encoding]::UTF8
        $msg.SubjectEncoding = [System.Text.Encoding]::UTF8
        $msg.IsBodyHtml = $true
        $msg.From = $Script:Config.ExchangeOnprem.WelcomeMail.From
        $msg.Subject = $Script:Config.ExchangeOnprem.WelcomeMail.Subject
        $msg.Body = $Script:Config.ExchangeOnprem.WelcomeMail.Body
        $msg.To.Add($UserPrincipalName)
        try
        {
            $smtpClient.Send($msg)
        }
        catch
        {
            throw [pscustomobject]@{
                Target     = $UserPrincipalName
                Activity   = $MyInvocation.MyCommand.Name
                Reason     = 'SmtpClient Send failed'
                Message    = $_.Exception.Message
                RetryCount = 3
                Delay      = 10
            }
        }
    }
}

function Set-KBAOnpremMailboxMessageConfiguration
{
    [CmdLetBinding()]
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]$UserPrincipalName
    )
    process
    {
        $params = @{
            Identity = $UserPrincipalName
            IsReplyAllTheDefaultResponse = $false
        }
        try
        {
            Set-OnpremMailboxMessageConfiguration @params
        }
        catch
        {       
            throw [pscustomobject]@{
                Target     = $UserPrincipalName
                Activity   = $MyInvocation.MyCommand.Name
                Reason     = 'Set-MailboxMessageConfiguration failed'
                Message    = $_.Exception.Message
                RetryCount = 3
                Delay      = 10
            }
        }
    }
}
