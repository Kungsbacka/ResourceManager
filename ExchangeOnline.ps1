# Make all error terminating errors
$Global:ErrorActionPreference = 'Stop'

Enum TestMailboxResult
{
    None
    Onprem
    Remote
    OnpremAndRemote
    OnpremDisabled
    Online
}

function Connect-KBAExchangeOnline
{
    $params = @{
        CertificateFilePath = $Script:Config.ExchangeOnline.AppCertificatePath
        CertificatePassword = ($Script:Config.ExchangeOnline.AppCertificatePassword | ConvertTo-SecureString)
        AppId = $Script:Config.ExchangeOnline.AppId
        Organization = $Script:Config.ExchangeOnline.Organization
        CommandName = @(
            'Get-EXOMailbox'
            'Set-Mailbox'
            'Set-MailboxRegionalConfiguration'
        )
        ShowBanner = $false
        ShowProgress = $false
        Prefix = 'Online'
    }
    Connect-ExchangeOnline @params
}

function Test-KBAOnlineMailbox
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
        if (Get-EXOMailbox -Identity $UserPrincipalName)
        {
            $result = [TestMailboxResult]::Online
        }
    }
    catch
    {
        if ($_.Exception.Message -notlike '*ManagementObjectNotFoundException*')
        {
            throw
        }
    }
    $result
}

function Set-RmOnlineOwa
{
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]
        $UserPrincipalName
    )
    $result = Test-KBAOnlineMailbox $UserPrincipalName
    if ($result -ne [TestMailboxResult]::Online)
    {
        throw 'Target account has no Exchange Online mailbox'
    }
    $params = @{
        Identity = $UserPrincipalName
        DateFormat = $Script:Config.ExchangeOnline.Owa.DateFormat
        TimeFormat = $Script:Config.ExchangeOnline.Owa.TimeFormat
        TimeZone = $Script:Config.ExchangeOnline.Owa.TimeZone
        Language = $Script:Config.ExchangeOnline.Owa.Language
    }
    Set-OnlineMailboxRegionalConfiguration @params
}

function Set-RmOnlineMailbox
{
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]
        $UserPrincipalName,
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Employee', 'Faculty', 'Student', 'Shared')]
        [string]
        $Type
    )
    $result = Test-KBAOnlineMailbox $UserPrincipalName
    if ($result -ne [TestMailboxResult]::Online)
    {
        throw 'Target account has no Exchange Online mailbox'
    }
    $params = @{
        Identity = $UserPrincipalName
        Languages = $Script:Config.ExchangeOnline.Mailbox.Languages
        AuditEnabled = $false
    }
    if ($Type -eq 'Employee' -or $Type -eq 'Shared')
    {
        $params.RetentionPolicy = $Script:Config.ExchangeOnline.Mailbox.RetentionPolicy
        $params.AddressBookPolicy = $Script:Config.ExchangeOnline.Mailbox.AddressBookPolicy
    }
    elseif ($Type -eq 'Faculty')
    {
        $params.RetentionPolicy = $Script:Config.ExchangeOnline.Mailbox.RetentionPolicy
    }
    else # Student
    {
        $params.RetentionPolicy = $Script:Config.ExchangeOnline.Mailbox.Student.RetentionPolicy
        $params.AddressBookPolicy = $Script:Config.ExchangeOnline.Mailbox.Student.AddressBookPolicy
    }
    Set-OnlineMailbox @params

    # Toggle 'AuditEnabled' off and on again. This is apparently neccesary for auditing to start working
    $params = @{
        Identity = $UserPrincipalName
        AuditEnabled = $true
    }
    Set-OnlineMailbox @params
}

function Set-RmOnlineMailboxType
{
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]
        $UserPrincipalName,
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Regular','Room','Equipment','Shared')]
        [string]
        $MailboxType
    )
    $result = Test-KBAOnlineMailbox $UserPrincipalName
    if ($result -ne [TestMailboxResult]::Online)
    {
        throw 'Target account has no Exchange Online mailbox'
    }
    Set-OnlineMailbox -Identity $UserPrincipalName -Type $MailboxType
}
