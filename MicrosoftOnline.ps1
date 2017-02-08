# Make all error terminating errors
$Global:ErrorActionPreference = 'Stop'
. "$PSScriptRoot\Config.ps1"

function Connect-KBAAzureAD
{
    Import-Module -Name 'AzureAD' -DisableNameChecking
    $credential = New-Object -TypeName 'System.Management.Automation.PSCredential' -ArgumentList @(
        $Script:Config.Office365.User
        $Script:Config.Office365.Password | ConvertTo-SecureString
    )
    try
    {
        Connect-AzureAD -Credential $credential | Out-Null
    }
    catch
    {
        throw [pscustomobject]@{
            Time       = [DateTime]::Now
            Target     = 'AzureAD'
            Activity   = $MyInvocation.MyCommand.Name
            Reason     = 'Connect-AzureAD failed'
            Message    = $_.Exception.Message
            RetryCount = 20
            Delay      = 5
        }
    }
}

function Enable-KBAMsolSync
{
    [CmdLetBinding()]
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [Alias('SamAccountName','ObjectGUID','DistinguishedName')]
        [string]$Identity
    )
    process
    {
        $params = @{
            Identity = $Identity
            Properties = @('ExtensionAttribute11')
        }
        try
        {
            $user = Get-ADUser @params
        }
        catch
        {
            throw [pscustomobject]@{
                Target     = $Identity
                Activity   = $MyInvocation.MyCommand.Name
                Reason     = 'Get-ADUser failed'
                Message    = $_.Exception.Message
                RetryCount = 3
                Delay      = 5
            }
        }
        if ($user.ExtensionAttribute11 -eq $null)
        {
            $params = @{
                Identity = $Identity
                Add = @{
                    ExtensionAttribute11 = 'SYNC_ME'
                }
                Replace = @{
                    MsExchUsageLocation = 'SE' # Synced to Usage Location in Azure AD
                }
            }
            try
            {
                Set-ADUser @params
            }
            catch
            {
                throw [pscustomobject]@{
                    Target     = $Identity
                    Activity   = $MyInvocation.MyCommand.Name
                    Reason     = 'Set-ADUser failed'
                    Message    = $_.Exception.Message
                    RetryCount = 3
                    Delay      = 5
                }
            }
        }
    }
}

function Set-KBAMsolUserLicense
{
    [CmdLetBinding()]
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]$UserPrincipalName,
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [object]$License
    )
    process
    {
        try
        {
            $msolUser = Get-AzureADUser -ObjectId $UserPrincipalName
        }
        catch
        {
            throw [pscustomobject]@{
                Target     = $UserPrincipalName
                Activity   = $MyInvocation.MyCommand.Name
                Reason     = 'Get-AzureADUser failed'
                Message    = $_.Exception.Message
                RetryCount = 3
                Delay      = 20
            }
        }
        $removeLicenses = New-Object 'System.Collections.Generic.List[string]'
        foreach ($item in $msolUser.AssignedLicenses)
        {
            $removeLicenses.Add($item.SkuId)
        }
        $addLicenses = New-Object 'System.Collections.Generic.List[Microsoft.Open.AzureAD.Model.AssignedLicense]'
        foreach ($item in $License)
        {
            try
            {
                $addLicenses.Add((New-Object -TypeName 'Microsoft.Open.AzureAD.Model.AssignedLicense' -ArgumentList @($item.DisabledPlans, $item.SkuId)))
            }
            catch
            {
                throw [pscustomobject]@{
                    Target     = $UserPrincipalName
                    Activity   = $MyInvocation.MyCommand.Name
                    Reason     = 'New [Microsoft.Open.AzureAD.Model.AssignedLicense]'
                    Message    = $_.Exception.Message
                    RetryCount = 0
                    Delay      = 0
                }
            }
        }
        try
        {
            $assignedLicenses = New-Object -TypeName 'Microsoft.Open.AzureAD.Model.AssignedLicenses' -ArgumentList @($addLicenses, $removeLicenses)
        }
        catch
        {
                throw [pscustomobject]@{
                    Target     = $UserPrincipalName
                    Activity   = $MyInvocation.MyCommand.Name
                    Reason     = 'New [Microsoft.Open.AzureAD.Model.AssignedLicenses]'
                    Message    = $_.Exception.Message
                    RetryCount = 0
                    Delay      = 0
                }
        }
        try
        {
            Set-AzureADUserLicense -ObjectId $UserPrincipalName -AssignedLicenses $assignedLicenses
        }
        catch
        {
            throw [pscustomobject]@{
                Target     = $UserPrincipalName
                Activity   = $MyInvocation.MyCommand.Name
                Reason     = 'Set-AzureADUserLicense failed'
                Message    = $_.Exception.Message
                RetryCount = 3
                Delay      = 5
            }
        }
    }
}

function Remove-KBAMsolUserLicense
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
            $msolUser = Get-AzureADUser -ObjectId $UserPrincipalName
        }
        catch
        {
            throw [pscustomobject]@{
                Target     = $UserPrincipalName
                Activity   = $MyInvocation.MyCommand.Name
                Reason     = 'Get-AzureADUser failed'
                Message    = $_.Exception.Message
                RetryCount = 3
                Delay      = 20
            }
        }
        $licenseJson = ConvertTo-Json -InputObject @($msolUser.AssignedLicenses) -Compress
        $licenseJson = $licenseJson.ToString()
        if ($licenseJson -eq '[]')
        {
            # User has no license
            return
        }
        try
        {
            Set-ADUser -Identity $SamAccountName -Replace @{ExtensionAttribute1=$licenseJson}
        }
        catch
        {
            throw [pscustomobject]@{
                Target     = $UserPrincipalName
                Activity   = $MyInvocation.MyCommand.Name
                Reason     = 'Set-ADUser failed'
                Message    = $_.Exception.Message
                RetryCount = 3
                Delay      = 20
            }
        }
        $removeLicenses = New-Object 'System.Collections.Generic.List[string]'
        foreach ($item in $msolUser.AssignedLicenses)
        {
            $removeLicenses.Add($item.SkuId)
        }
        try
        {
            $assignedLicenses = New-Object -TypeName 'Microsoft.Open.AzureAD.Model.AssignedLicenses' -ArgumentList @($null, $removeLicenses)
        }
        catch
        {
                throw [pscustomobject]@{
                    Target     = $UserPrincipalName
                    Activity   = $MyInvocation.MyCommand.Name
                    Reason     = 'New [Microsoft.Open.AzureAD.Model.AssignedLicenses]'
                    Message    = $_.Exception.Message
                    RetryCount = 0
                    Delay      = 0
                }
        }
        try
        {
            Set-AzureADUserLicense -ObjectId $UserPrincipalName -AssignedLicenses $assignedLicenses
        }
        catch
        {
            throw [pscustomobject]@{
                Target     = $UserPrincipalName
                Activity   = $MyInvocation.MyCommand.Name
                Reason     = 'Set-AzureADUserLicense failed'
                Message    = $_.Exception.Message
                RetryCount = 3
                Delay      = 5
            }
        }
    }
}

function Restore-KBAMsolUserLicense
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
            $adUser = Get-ADUser -Identity $SamAccountName -Properties @('ExtensionAttribute1')
        }
        catch
        {
            throw [pscustomobject]@{
                Target     = $SamAccountName
                Activity   = $MyInvocation.MyCommand.Name
                Reason     = 'Get-ADUSer failed'
                Message    = $_.Exception.Message
                RetryCount = 3
                Delay      = 20
            }
        }
        if ($adUser.ExtensionAttribute1 -eq $null)
        {
            throw [pscustomobject]@{
                Target     = $SamAccountName
                Activity   = $MyInvocation.MyCommand.Name
                Reason     = 'ExtensionAttribute1 is empty'
                Message    = 'A stashed license does not exists in ExtensionAttribute1'
                RetryCount = 0
                Delay      = 0
            }
        }
        
        try
        {
            $license = $adUser.ExtensionAttribute1 | ConvertFrom-Json
        }
        catch
        {
            throw [pscustomobject]@{
                Target     = $UserPrincipalName
                Activity   = $MyInvocation.MyCommand.Name
                Reason     = 'ConvertFrom-Json failed'
                Message    = $_.Exception.Message
                RetryCount = 0
                Delay      = 0
            }
        }
        try
        {
            Set-KBAMsolUserLicense -UserPrincipalName $UserPrincipalName -License $license
        }
        catch
        {
            throw [pscustomobject]@{
                Target     = $UserPrincipalName
                Activity   = $MyInvocation.MyCommand.Name
                Reason     = 'Set-KBAMsolUserLicense failed'
                Message    = $_.Exception.Message
                RetryCount = 3
                Delay      = 20
            }
        }
        try
        {
            Set-ADUser -Identity $SamAccountName -Clear 'ExtensionAttribute1'
        }
        catch
        {
            throw [pscustomobject]@{
                Target     = $UserPrincipalName
                Activity   = $MyInvocation.MyCommand.Name
                Reason     = 'Set-ADUser failed'
                Message    = $_.Exception.Message
                RetryCount = 3
                Delay      = 20
            }
        }
    }
}
