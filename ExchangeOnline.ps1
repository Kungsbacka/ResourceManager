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
    try
    {
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
    catch
    {
        throw [pscustomobject]@{
            Target     = 'https://outlook.office365.com/powershell-liveid/'
            Activity   = $MyInvocation.MyCommand.Name
            Reason     = 'New/Import-PSSession failed'
            Message    = $_.Exception.Message
            RetryCount = 50
            Delay      = 10
        }
    }
}

function Test-KBAOnlineMailbox
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
}

function Set-KBAOnlineOwa
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
            $result = Test-KBAOnlineMailbox $UserPrincipalName
        }
        catch
        {
            throw [pscustomobject]@{
                Target     = $UserPrincipalName
                Activity   = $MyInvocation.MyCommand.Name
                Reason     = 'Test-KBAOnlineMailbox failed'
                Message    = $_.Exception.Message
                RetryCount = 3
                Delay      = 5
            }
        }
        if ($result -ne [TestMailboxResult]::Online)
        {
            throw [pscustomobject]@{
                Target     = $UserPrincipalName
                Activity   = $MyInvocation.MyCommand.Name
                Reason     = 'Online mailbox not found'
                Message    = 'An Office 365 mailbox not found for target.'
                RetryCount = 0
                Delay      = 0
            }
        }
        $params = @{
            Identity = $UserPrincipalName
            DateFormat = 'yyyy-MM-dd'
            TimeFormat = 'HH:mm'
            TimeZone = 'W. Europe Standard Time'
            Language = 'sv-SE'
        }
        try
        {
            Set-OnlineMailboxRegionalConfiguration @params
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
