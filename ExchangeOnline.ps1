# Make all error terminating errors
$Global:ErrorActionPreference = 'Stop'

Enum TestMailboxResult
{
    None
    Onprem
    Remote
    Online
    Both
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
        DateFormat = 'yyyy-MM-dd'
        TimeFormat = 'HH:mm'
        TimeZone = 'W. Europe Standard Time'
        Language = 'sv-SE'
    }
    Set-OnlineMailboxRegionalConfiguration @params
}
