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
    param
    (
        [pscredential]
        [System.Management.Automation.Credential()]
        $Credential
    )
    $server = $Script:Config.ExchangeOnprem.Servers | Get-Random
    Get-PSSession -Name 'KBAExchOnprem' -ErrorAction SilentlyContinue | Remove-PSSession
    if ($Credential -eq $null)
    {
        $Credential = New-Object -TypeName 'System.Management.Automation.PSCredential' -ArgumentList @(
            $Script:Config.ExchangeOnprem.User
            $Script:Config.ExchangeOnprem.Password | ConvertTo-SecureString
        )
    }
    $params = @{
        Name = 'KBAExchOnprem'
        ConfigurationName = 'Microsoft.Exchange'
        ConnectionUri = "http://$server/PowerShell/"
        Authentication = 'Kerberos'
        Credential = $Credential
    }
    $session = New-PSSession @params
    $params = @{
        Session = $session
        Prefix = 'Onprem'
        DisableNameChecking = $true
        CommandName = @(
            'Connect-Mailbox'
            'Disable-Mailbox'
            'Enable-Mailbox'
            'Enable-RemoteMailbox'
            'Get-Mailbox'
            'Get-RemoteMailbox'
            'Get-MailboxStatistics'
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

function Test-KBAOnpremMailbox
{
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]
        $UserPrincipalName
    )
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

function Set-KBAOnpremOwa
{
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]
        $UserPrincipalName,
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Employee', 'Student', 'Shared')]
        [string]
        $Type
    )
    $result = Test-KBAOnpremMailbox $UserPrincipalName
    if ($result -ne [TestMailboxResult]::Onprem)
    {
        throw 'Target account has no on-prem mailbox'
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
    Set-OnpremCasMailbox @params
    $params = @{
        Identity = $UserPrincipalName
        DateFormat = $Script:Config.ExchangeOnprem.Owa.DateFormat
        TimeFormat = $Script:Config.ExchangeOnprem.Owa.TimeFormat
        TimeZone = $Script:Config.ExchangeOnprem.Owa.TimeZone
        Language = $Script:Config.ExchangeOnprem.Owa.Language
        LocalizeDefaultFolderName = $Script:Config.ExchangeOnprem.Owa.LocalizeDefaultFolderName
    }
    Set-OnpremMailboxRegionalConfiguration @params
}

function Set-KBAOnpremCalendar
{
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]
        $UserPrincipalName,
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Employee', 'Student', 'Shared')]
        [string]
        $Type
    )
    $result = Test-KBAOnpremMailbox $UserPrincipalName
    if ($result -ne [TestMailboxResult]::Onprem)
    {
        throw 'Target account has no on-prem mailbox'
    }
    $calendarIdentities = @()
    $params = @{
        Identity = $UserPrincipalName
        FolderScope = 'Calendar'
    }
    $calendarFolders = Get-OnpremMailboxFolderStatistics @params
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
        throw 'No calendar folder found'
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
    Set-OnpremMailboxCalendarConfiguration @params
    if ($Type -in 'Employee', 'Shared')
    {
        foreach($calendarIdentity in $calendarIdentities)
        {
            $params = @{
                Identity = $calendarIdentity
                User = 'Default'
                AccessRights = $Script:Config.ExchangeOnprem.Calendar.DefaultCalendarPermission
            }
            Set-OnpremMailboxFolderPermission @params
        }
    }
}

# The mailbox configuration is moved to Enable-KBAOnpremMailbox.
# This function is kept around to be able to run mailbox configuration
# as a separate task after the mailbox is created.
function Set-KBAOnpremMailbox
{
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]
        $UserPrincipalName,
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Employee', 'Student', 'Shared')]
        [string]
        $Type
    )
    $result = Test-KBAOnpremMailbox $UserPrincipalName
    if ($result -ne [TestMailboxResult]::Onprem)
    {
        throw 'Target account has no on-prem mailbox'
    }
    if ($Type -eq 'Student')
    {
        $params = @{
            Identity = $UserPrincipalName
            PrimarySmtpAddress = $UserPrincipalName
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
            PrimarySmtpAddress = $UserPrincipalName
            EmailAddressPolicyEnabled = $Script:Config.ExchangeOnprem.Mailbox.EmailAddressPolicyEnabled
            RetentionPolicy = $Script:Config.ExchangeOnprem.Mailbox.RetentionPolicy
            AddressBookPolicy = $Script:Config.ExchangeOnprem.Mailbox.AddressBookPolicy
            IssueWarningQuota = $Script:Config.ExchangeOnprem.Mailbox.IssueWarningQuota
            ProhibitSendQuota = $Script:Config.ExchangeOnprem.Mailbox.ProhibitSendQuota
            ProhibitSendReceiveQuota = $Script:Config.ExchangeOnprem.Mailbox.ProhibitSendReceiveQuota
            UseDatabaseQuotaDefaults = $Script:Config.ExchangeOnprem.Mailbox.UseDatabaseQuotaDefaults
        }
    }
    Set-OnpremMailbox @params
}

function Enable-KBAOnpremMailbox
{
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]
        $UserPrincipalName,
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]
        $SamAccountName,
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Employee', 'Student', 'Shared')]
        [string]
        $Type
    )
    $result = Test-KBAOnpremMailbox $UserPrincipalName
    if ($result -eq [TestMailboxResult]::Onprem)
    {
        throw 'Target alread has an on-prem mailbox'
    }
    if ($result -eq [TestMailboxResult]::Remote)
    {
        throw 'Target alread has a remote mailbox'
    }
    $params = @{
        Identity = $UserPrincipalName
        PrimarySmtpAddress = $UserPrincipalName
        Alias = $SamAccountName
    }
    if ($Type -eq 'Student')
    {
        $database =
            Get-OnpremMailboxDatabase |
            where 'Name' -Like -Value 'S_DB*' |
            Get-Random
        $params.Database = $database.Name
    }
    Enable-OnpremMailbox @params | Out-Null
    # Sleep until mailbox is available
    $retryCount = 0
    do 
    {
        Start-Sleep -Seconds 5
    }
    while (-not (Get-OnpremMailbox -Identity $UserPrincipalName -ErrorAction SilentlyContinue) -and ++$retryCount -lt 10)
    if ($retryCount -eq 10)
    {
        throw 'Mailbox not created in time'
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
    Set-OnpremMailbox @params
}

function Set-KBAOnpremRemoteMailbox
{
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]
        $UserPrincipalName
    )
    $result = Test-KBAOnpremMailbox $UserPrincipalName
    if ($result -ne [TestMailboxResult]::Remote)
    {
        throw 'Target mailbox is not a remote mailbox'
    }
    $mailbox = Get-OnpremRemoteMailbox -Identity $UserPrincipalName
    if ($mailbox.RemoteRoutingAddress -cnotmatch "^SMTP:[^@]+@$($Script:Config.ExchangeOnprem.ExchangeOnlineMailDomain)$")
    {
        throw "Invalid RemoteRoutingAddress: $($mailbox.RemoteRoutingAddress)"
    }
    if ($mailbox.RemoteRoutingAddress -notin $mailbox.EmailAddresses)
    {
        $rra = 'smtp:' + $mailbox.RemoteRoutingAddress.Substring(5)
        $params = @{
            Identity = $UserPrincipalName
            EmailAddresses = @{'Add' = $rra}
        }
        Set-OnpremRemoteMailbox @params
    }
}

function Enable-KBAOnpremRemoteMailbox
{
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]
        $UserPrincipalName,
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]
        $SamAccountName
    )
    $result = Test-KBAOnpremMailbox $UserPrincipalName
    if ($result -eq [TestMailboxResult]::Onprem)
    {
        throw 'Target alread has an on-prem mailbox'
    }
    if ($result -eq [TestMailboxResult]::Remote)
    {
        throw 'Target alread has an remote mailbox'
    }
    $params = @{
        Identity = $UserPrincipalName
        PrimarySmtpAddress = $UserPrincipalName
        Alias = $SamAccountName
        RemoteRoutingAddress = $SamAccountName + '@' +
            $Script:Config.ExchangeOnprem.ExchangeOnlineMailDomain
    }
    Enable-OnpremRemoteMailbox @params | Out-Null
    # Not sure why this is needed, but sometimes Enable-RemoteMailbox
    # fails if this is not present. Initially 10.
    Start-Sleep -Seconds 5
}

function Disable-KBAOnpremMailbox
{
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]
        $UserPrincipalName
    )
    $result = Test-KBAOnpremMailbox $UserPrincipalName
    if ($result -ne [TestMailboxResult]::Onprem)
    {
        throw 'Target account has no on-prem mailbox'
    }
    Disable-OnpremMailbox -Identity $UserPrincipalName -Confirm:$false | Out-Null
}

function Connect-KBAOnpremMailbox
{
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]
        $UserPrincipalName,
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]
        $SamAccountName
    )
    $result = Test-KBAOnpremMailbox $UserPrincipalName
    if ($result -eq [TestMailboxResult]::Onprem)
    {
        throw 'Target alread has an on-prem mailbox'
    }
    if ($result -eq [TestMailboxResult]::Remote)
    {
        throw 'Target alread has an remote mailbox'
    }
    $user = Get-ADUser -Identity $SamAccountName -Properties LegacyExchangeDn
    if (-not $user.LegacyExchangeDn)
    {
        throw 'No disconnected mailbox exists for target'
    }
    $disconnectedMailboxes = Get-OnpremMailboxDatabase | 
        Get-OnpremMailboxStatistics -Filter "LegacyDn -eq '$($user.LegacyExchangeDn)'" -NoADLookup | 
        Sort-Object -Property DisconnectDate -Descending

    if ($disconnectedMailboxes.Count -eq 0 -or $disconnectedMailboxes[0].DisconnectDate -eq $null)
    {
        throw 'No disconnected mailbox exists for target'
    }
    $param = @{
        Identity = $user.LegacyExchangeDn
        Database = $disconnectedMailboxes[0].Database.ToString()
        Force = $true
    }
    Connect-OnpremMailbox @param
}

function Cleanup-KBAOnpremMailbox
{
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]
        $UserPrincipalName
    )
    $result = Test-KBAOnpremMailbox $UserPrincipalName
    if ($result -ne [TestMailboxResult]::Onprem)
    {
        throw 'Target account has no on-prem mailbox'
    }
    $mailbox = Get-OnpremMailbox -Identity $UserPrincipalName
    $addresses = $mailbox.EmailAddresses
    $currentO365Address = $addresses | where {$_ -like "*@$($Script:Config.ExchangeOnprem.ExchangeOnlineMailDomain)"}
    $correctO365Address = "smtp:$($mailbox.SamAccountName)@$($Script:Config.ExchangeOnprem.ExchangeOnlineMailDomain)"
    if ($currentO365Address)
    {
        if ($currentO365Address -ne $correctO365Address)
        {
            $addresses = $addresses | where {$_ -ne $currentO365Address}
            $addresses += $correctO365Address
        }
    }
    else
    {
        $addresses += $correctO365Address
    }
    $alias = ''
    foreach ($c in [char[]]($mailbox.Alias.Normalize('FormD')))
    {
        if ([System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($c) -ne [System.Globalization.UnicodeCategory]::NonSpacingMark)
        {
            $alias += $c
        }
    }
    $badSecondaryAddress = $addresses | where {$_ -clike "smtp:$alias@*"}
    if ($badSecondaryAddress)
    {
        $addresses = $addresses | where {$_ -ne $badSecondaryAddress}
    }
    Set-OnpremMailbox -Identity $UserPrincipalName -EmailAddresses $addresses
    if ($mailbox.Alias -ne $mailbox.SamAccountName)
    {
        # Alias did not change when i tried to set EmailAddress and Alias at the same time
        Set-OnpremMailbox -Identity $UserPrincipalName -Alias $mailbox.SamAccountName
    }
}

function Set-KBAOnpremMailboxAutoReplyState
{
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]
        $UserPrincipalName,
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [bool]
        $Enabled,
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string]
        $Message
        
    )
    $result = Test-KBAOnpremMailbox $UserPrincipalName
    if ($result -ne [TestMailboxResult]::Onprem)
    {
        throw 'Target account has no on-prem mailbox'
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
    Set-OnpremMailboxAutoReplyConfiguration @params | Out-Null
}

function Send-KBAOnpremWelcomeMail
{
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]
        $UserPrincipalName,
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Employee', 'Student', 'Shared')]
        [string]
        $Type
    )
    if ($Type -ne 'Employee')
    {
        throw 'Wrong mailbox type. Welcome mail is only sent to employees.'
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
    $smtpClient.Send($msg)
}

function Set-KBAOnpremMailboxMessageConfiguration
{
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
        Set-OnpremMailboxMessageConfiguration @params
    }
}
