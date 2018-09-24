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
    Get-PSSession -Name 'KBAExchOnline' -ErrorAction SilentlyContinue | Remove-PSSession
    $credential = New-Object -TypeName 'System.Management.Automation.PSCredential' -ArgumentList @(
        $Script:Config.Office365.User
        $Script:Config.Office365.Password | ConvertTo-SecureString
    )
    $params = @{
        Name = 'KBAExchOnline'
        ConfigurationName = 'Microsoft.Exchange'
        ConnectionUri = 'https://outlook.office365.com/powershell-liveid/'
        Authentication = 'Basic'
        AllowRedirection = $true
        Credential = $credential
    }
    $session = New-PSSession @params
    $params = @{
        Session = $session
        Prefix = 'Online'
        DisableNameChecking = $true
        CommandName = @(
            'Get-Mailbox'
            'Set-Mailbox'
            'Set-MailboxRegionalConfiguration'
        )
    }
    Import-PSSession @params | Out-Null
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
        if (Get-OnlineMailbox -Identity $UserPrincipalName)
        {
            $result = [TestMailboxResult]::Online
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

function Set-KBAOnlineOwa
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
        throw 'Target account has no Office 365 mailbox'
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

function Set-KBAOnlineMailbox
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
        throw 'Target account has no Office 365 mailbox'
    }
    $params = @{
        Identity = $UserPrincipalName
        Languages = $Script:Config.ExchangeOnline.Mailbox.Languages
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
        $params.AddressBookPolicy = $Script:Config.ExchangeOnline.Mailbox.Student.AddressBookPolicy
    }
    Set-OnlineMailbox @params
}