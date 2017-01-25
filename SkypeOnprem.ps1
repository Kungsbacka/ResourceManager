# Make all error terminating errors
$Global:ErrorActionPreference = 'Stop'
. "$PSScriptRoot\Config.ps1"

function Import-KBALyncOnpremModule
{
    try
    {
        Import-Module -Name 'SkypeForBusiness' -Prefix 'Onprem'
    }
    catch
    {
        throw [pscustomobject]@{
            Target     = 'Lync module'
            Activity   = $MyInvocation.MyCommand.Name
            Reason     = 'Import-Module failed'
            Message    = $_.Exception.Message
            RetryCount = 3
            Delay      = 5
        }

    }

}

function Test-KBAOnpremCSUser
{
    [CmdLetBinding()]
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]$UserPrincipalName
    )
    process
    {
        $isEnabled = $false
        try
        {
            if ((Get-OnpremCsUser -Identity $UserPrincipalName))
            {
                $isEnabled = $true
            }
        }
        catch
        {
            if ($_.Exception.Data['LyncId'] -notlike '*ObjectNotFoundIdentity')
            {
                throw
            }
        }
        $isEnabled
    }
}

function Grant-KBAOnpremCSConferencingPolicy
{
    [CmdLetBinding()]
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]$UserPrincipalName,
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]$ConferencingPolicy
    )
    process
    {
        try
        {
            $enabled = Test-KBAOnpremCSUser $UserPrincipalName
        }
        catch
        {
            throw [pscustomobject]@{
                Target     = $UserPrincipalName
                Activity   = $MyInvocation.MyCommand.Name
                Reason     = 'Test-KBAOnpremCSUser failed'
                Message    = $_.Exception.Message
                RetryCount = 3
                Delay      = 5
            }
        }
        if (-not $enabled)
        {
            throw [pscustomobject]@{
                Target     = $UserPrincipalName
                Activity   = $MyInvocation.MyCommand.Name
                Reason     = 'Not Lync enabled'
                Message    = 'Target is not Lync enabled.'
                RetryCount = 3
                Delay      = 5
            }
        }
        $params = @{
            Identity = $UserPrincipalName
            PolicyName = $ConferencingPolicy
        }
        try
        {
            Grant-OnpremCsConferencingPolicy @params
        }
        catch
        {
            throw [pscustomobject]@{
                Target     = $UserPrincipalName
                Activity   = $MyInvocation.MyCommand.Name
                Reason     = 'Grant-CsConferencingPolicy failed'
                Message    = $_.Exception.Message
                RetryCount = 3
                Delay      = 5
            }
        }
    }
}

function Enable-KBAOnpremCSUser
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
            $enabled = Test-KBAOnpremCSUser $UserPrincipalName
        }
        catch
        {
            throw [pscustomobject]@{
                Target     = $UserPrincipalName
                Activity   = $MyInvocation.MyCommand.Name
                Reason     = 'Test-KBAOnpremCSUser failed'
                Message    = $_.Exception.Message
                RetryCount = 3
                Delay      = 5
            }
        }
        if ($enabled)
        {
            throw [pscustomobject]@{
                Target     = $UserPrincipalName
                Activity   = $MyInvocation.MyCommand.Name
                Reason     = 'Already Lync enabled'
                Message    = 'Target is already Lync enabled.'
                RetryCount = 0
                Delay      = 0
            }
        }
        $params = @{
            Identity = $UserPrincipalName
            RegistrarPool = $Script:Config.LyncOnprem.RegistrarPool
            SipAddress = 'sip:' + $UserPrincipalName
        }
        try
        {
            Enable-OnpremCSUser @params
        }
        catch
        {
            throw [pscustomobject]@{
                Target     = $UserPrincipalName
                Activity   = $MyInvocation.MyCommand.Name
                Reason     = 'Enable-CSUser failed'
                Message    = $_.Exception.Message
                RetryCount = 3
                Delay      = 5
            }
        }
    }
}
