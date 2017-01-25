# Make all error terminating errors
$Global:ErrorActionPreference = 'Stop'
. "$PSScriptRoot\Config.ps1"

function Connect-KBAMsolService
{
    Import-Module -Name 'MSOnline' -DisableNameChecking
    $credential = New-Object -TypeName 'System.Management.Automation.PSCredential' -ArgumentList @(
        $Script:Config.Office365.User
        $Script:Config.Office365.Password | ConvertTo-SecureString
    )
    try
    {
        Connect-MsolService -Credential $credential
    }
    catch
    {
        throw [pscustomobject]@{
            Time       = [DateTime]::Now
            Target     = 'MSOnline'
            Activity   = $MyInvocation.MyCommand.Name
            Reason     = 'Connect-MsolService failed'
            Message    = $_.Exception.Message
            RetryCount = 100
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
            Properties = 'ExtensionAttribute11'
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
                Delay      = 0
            }
        }
        if ($user.ExtensionAttribute11 -eq $null)
        {
            $params = @{
                Identity = $Identity
                Add = @{
                    ExtensionAttribute11 = 'SYNC_ME'
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
                    Delay      = 0
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
        [string]$License
    )
    process
    {
        try
        {
            $msolUser = Get-MsolUser -UserPrincipalName $UserPrincipalName
        }
        catch
        {
            throw [pscustomobject]@{
                Target     = $UserPrincipalName
                Activity   = $MyInvocation.MyCommand.Name
                Reason     = 'Get-MsolUser failed'
                Message    = $_.Exception.Message
                RetryCount = 10
                Delay      = 20
            }
        }
        $currentLicense = ''
        foreach ($lic in $msolUser.Licenses)
        {
            $currentLicense += $lic.AccountSkuId -replace "$($Script:Config.MicrosoftOnline.AccountName):", ''
            $disabledPlans = '('
            foreach($status in $lic.ServiceStatus)
            {
                if ($status.ProvisioningStatus -eq 'Disabled')
                {
                    $disabledPlans = $disabledPlans + $status.ServicePlan.ServiceName + ','
                }
            }
            $disabledPlans = $disabledPlans.TrimEnd(',') + ')'
            if ($disabledPlans.Length -gt 2)
            {
                $currentLicense += $disabledPlans
            }
            $currentLicense += '+'
        }
        if ($License -eq $currentLicense.TrimEnd('+'))
        {
            Write-Verbose -Message "User account ($UserPrincipalName) already has this license"
            return
        }
        # Set-MsolUserLicense fails if usage location is not set
        if ($msolUser.UsageLocation -eq $null)
        {
            try
            {
                Set-MsolUser -UserPrincipalName $UserPrincipalName -UsageLocation 'SE'
            }
            catch
            {
                throw [pscustomobject]@{
                    Target     = $UserPrincipalName
                    Activity   = $MyInvocation.MyCommand.Name
                    Reason     = 'Set-MsolUser failed'
                    Message    = $_.Exception.Message
                    RetryCOunt = 3
                    Delay      = 5
                }
            }
        }   
        $licensesToRemove = $msolUser.Licenses.AccountSkuId
        $licensesToAdd = @()
        $licenseOptions = @()
        $skus = $License -split '\+'
        foreach ($sku in $skus)
        {
            if ($sku -match '\(([^)]+)\)')
            {
                $disabledPlans = $Matches[1] -split ','
                $sku = $sku -replace '\([^)]+\)', ''
                $licenseOptions += New-MsolLicenseOptions -AccountSkuId "$($Script:Config.MicrosoftOnline.AccountName):$sku" -DisabledPlans $disabledPlans
            }
            $licensesToAdd += "$($Script:Config.MicrosoftOnline.AccountName):$sku"
        }
        try
        {
            # This is the only way I was able to replace all licenses.
            # Only one call to Set-MsolUserLicense with all parameters did not work reliably.
            Set-MsolUserLicense -UserPrincipalName $UserPrincipalName -RemoveLicenses $licensesToRemove
            Set-MsolUserLicense -UserPrincipalName $UserPrincipalName -AddLicenses $licensesToAdd
            Set-MsolUserLicense -UserPrincipalName $UserPrincipalName -LicenseOptions $licenseOptions
        }
        catch
        {
            throw [pscustomobject]@{
                Target     = $UserPrincipalName
                Activity   = $MyInvocation.MyCommand.Name
                Reason     = 'Set-MsolUserLicense failed'
                Message    = $_.Exception.Message
                RetryCount = 3
                Delay      = 5
            }
        }
    }
}
