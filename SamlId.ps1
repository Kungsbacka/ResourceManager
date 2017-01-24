# Make all error terminating errors
$Global:ErrorActionPreference = 'Stop'
. "$PSScriptRoot\Config.ps1"

function Set-SamlId
{
    [CmdLetBinding()]
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]$SamAccountName,
        [switch]$Force
    )
    process
    {
        try
        {
            $params = @{
                Identity = $SamAccountName
                Properties = 'ExtensionAttribute14'
            }
            $user = Get-ADUser @params
        }
        catch
        {
            throw [pscustomobject]@{
                Target     = $SamAccountName
                Activity   = $MyInvocation.MyCommand.Name
                Reason     = 'Get-ADUser failed with identity'
                Message    = $_.Exception.Message
                RetryCount = 3
                Delay      = 5
            }
        }
        $hasValue = $user.ExtensionAttribute14 -ne $null
        if ($hasValue)
        {
            if ($Force)
            {
                Write-Warning 'User already has a value in ExtensionAttribute14. This value will be overwritten!'
            }
            else
            {
                throw [pscustomobject]@{
                    Target     = $SamAccountName
                    Activity   = $MyInvocation.MyCommand.Name
                    Reason     = 'ExtensionAttribute14 is not empty'
                    Message    = 'User already has a value in ExtensionAttribute14. Force was not specified.'
                    RetryCount = 0
                    Delay      = 0
                }
            }
        }
        $set = 'abcdefghijkmnpqrstuvxyz23456789'
        $exists = $false
        do
        {
            $id = ''
            for ($i = 0; $i -lt 16; $i++)
            {
                $id += $set.GetEnumerator() | Get-Random
            }
            $id += '@' + $Script:Config.SamlId.Domain
            try
            {
                if ((Get-ADUser -Filter {ExtensionAttribute14 -eq $id}))
                {
                    $exists = $true
                }
            }
            catch
            {
                throw [pscustomobject]@{
                    Target     = $SamAccountName
                    Activity   = $MyInvocation.MyCommand.Name
                    Reason     = 'Get-ADUser failed with filter'
                    Message    = $_.Exception.Message
                    RetryCount = 3
                    Delay      = 5
                }
            }
        }
        while ($exists)
        if ($hasValue -and $Force)
        {
            try
            {
                $params = @{
                    Identity = $SamAccountName
                    Replace = @{ExtensionAttribute14 = $id}
                }
                Set-ADUser @params
            }
            catch
            {
                throw [pscustomobject]@{
                    Target     = $SamAccountName
                    Activity   = $MyInvocation.MyCommand.Name
                    Reason     = 'Set-ADUser failed with replace'
                    Message    = $_.Exception.Message
                    RetryCount = 3
                    Delay      = 5
                }
            }
        }
        else
        {
            try
            {
                $params = @{
                    Identity = $SamAccountName
                    Add = @{ExtensionAttribute14 = $id}
                }
                Set-ADUser @params
            }
            catch
            {
                throw [pscustomobject]@{
                    Target     = $SamAccountName
                    Activity   = $MyInvocation.MyCommand.Name
                    Reason     = 'Set-ADUser failed with add'
                    Message    = $_.Exception.Message
                    RetryCount = 3
                    Delay      = 5
                }
            }
        }
    }
}
