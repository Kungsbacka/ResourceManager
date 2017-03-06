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
    Connect-AzureAD -Credential $credential -LogLevel None | Out-Null
}

function Enable-KBAMsolSync
{
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [Alias('SamAccountName','ObjectGUID','DistinguishedName')]
        [string]
        $Identity
    )
    $params = @{
        Identity = $Identity
        Properties = @('ExtensionAttribute11')
    }
    $user = Get-ADUser @params
    if ($user.ExtensionAttribute11 -eq $null)
    {
        $params = @{
            Identity = $Identity
            Add = @{
                ExtensionAttribute11 = 'SYNC_ME'
            }
            Replace = @{
                # MsExchUsageLocation is synced to UsageLocation in Azure AD
                MsExchUsageLocation = $Script:Config.MicrosoftOnline.UsageLocation
            }
        }
        Set-ADUser @params
    }
}

function Set-KBAMsolUserLicense
{
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]
        $UserPrincipalName,
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [object]
        $License
    )
    $aadUser = Get-AzureADUser -ObjectId $UserPrincipalName
    if ($aadUser.UsageLocation -ne 'SE')
    {
        # If MsExchUsageLocation hasn't synced yet, we set UsageLocation explicitly
        Set-AzureADUser -ObjectId $UserPrincipalName -UsageLocation $Script:Config.MicrosoftOnline.UsageLocation
    }
    $removeLicenses = New-Object 'System.Collections.Generic.List[string]'
    foreach ($item in $aadUser.AssignedLicenses)
    {
        $removeLicenses.Add($item.SkuId)
    }
    $addLicenses = New-Object 'System.Collections.Generic.List[Microsoft.Open.AzureAD.Model.AssignedLicense]'
    foreach ($item in $License)
    {
        $addLicenses.Add((New-Object -TypeName 'Microsoft.Open.AzureAD.Model.AssignedLicense' -ArgumentList @($item.DisabledPlans, $item.SkuId)))
    }
    $assignedLicenses = New-Object -TypeName 'Microsoft.Open.AzureAD.Model.AssignedLicenses' -ArgumentList @($addLicenses, $removeLicenses)
    Set-AzureADUserLicense -ObjectId $UserPrincipalName -AssignedLicenses $assignedLicenses
}

function Remove-KBAMsolUserLicense
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
    $msolUser = Get-AzureADUser -ObjectId $UserPrincipalName
    $licenseJson = ConvertTo-Json -InputObject @($msolUser.AssignedLicenses) -Compress
    $licenseJson = $licenseJson.ToString()
    if ($licenseJson -eq '[]')
    {
        return # User has no license
    }
    Set-ADUser -Identity $SamAccountName -Replace @{ExtensionAttribute1=$licenseJson}
    $removeLicenses = New-Object 'System.Collections.Generic.List[string]'
    foreach ($item in $msolUser.AssignedLicenses)
    {
        $removeLicenses.Add($item.SkuId)
    }
    $assignedLicenses = New-Object -TypeName 'Microsoft.Open.AzureAD.Model.AssignedLicenses' -ArgumentList @($null, $removeLicenses)
    Set-AzureADUserLicense -ObjectId $UserPrincipalName -AssignedLicenses $assignedLicenses
}

function Restore-KBAMsolUserLicense
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
    $adUser = Get-ADUser -Identity $SamAccountName -Properties @('ExtensionAttribute1')
    if ($adUser.ExtensionAttribute1 -eq $null)
    {
        throw 'No stashed license exists in extensionAttribute1'
    }
    $license = $adUser.ExtensionAttribute1 | ConvertFrom-Json
    Set-KBAMsolUserLicense -UserPrincipalName $UserPrincipalName -License $license
    Set-ADUser -Identity $SamAccountName -Clear 'ExtensionAttribute1'
}
