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
        Connect-AzureAD -Credential $credential -LogLevel None | Out-Null
    }
    catch
    {
        $_.Exception.Data.Add('RetryCount', 20)
        $_.Exception.Data.Add('Delay', 5)
        throw
    }
}

function Enable-KBAMsolSync
{
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [Alias('SamAccountName','ObjectGUID','DistinguishedName')]
        [string]$Identity
    )
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
        $_.Exception.Data.Add('Parameters', $PSBoundParameters)
        $_.Exception.Data.Add('RetryCount', 1)
        $_.Exception.Data.Add('Delay', 3)
        throw
    }
    if ($user.ExtensionAttribute11 -eq $null)
    {
        $params = @{
            Identity = $Identity
            Add = @{
                ExtensionAttribute11 = 'SYNC_ME'
            }
            Replace = @{
                MsExchUsageLocation = $Script:Config.MicrosoftOnline.UsageLocation # Synced to Usage Location in Azure AD
            }
        }
        try
        {
            Set-ADUser @params
        }
        catch
        {
            $_.Exception.Data.Add('Parameters', $PSBoundParameters)
            $_.Exception.Data.Add('RetryCount', 1)
            $_.Exception.Data.Add('Delay', 3)
            throw
        }
    }
}

function Set-KBAMsolUserLicense
{
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]$UserPrincipalName,
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [object]$License
    )
    try
    {
        $aadUser = Get-AzureADUser -ObjectId $UserPrincipalName
    }
    catch
    {
        $_.Exception.Data.Add('Parameters', $PSBoundParameters)
        $_.Exception.Data.Add('RetryCount', 3)
        $_.Exception.Data.Add('Delay', 10)
        throw
    }
    if (aadUser.UsageLocation -ne 'SE')
    {
        try
        {
            # If MsExchUsageLocation hasn't synced yet we set UsageLocation explicitly
            Set-AzureADUser -ObjectId $UserPrincipalName -UsageLocation $Script:Config.MicrosoftOnline.UsageLocation
        }
        catch
        {
            $_.Exception.Data.Add('Parameters', $PSBoundParameters)
            $_.Exception.Data.Add('RetryCount', 3)
            $_.Exception.Data.Add('Delay', 10)
            throw            
        }
    }
    $removeLicenses = New-Object 'System.Collections.Generic.List[string]'
    foreach ($item in $aadUser.AssignedLicenses)
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
            $_.Exception.Data.Add('Parameters', $PSBoundParameters)
            throw
        }
    }
    try
    {
        $assignedLicenses = New-Object -TypeName 'Microsoft.Open.AzureAD.Model.AssignedLicenses' -ArgumentList @($addLicenses, $removeLicenses)
    }
    catch
    {
        $_.Exception.Data.Add('Parameters', $PSBoundParameters)
        throw
    }
    try
    {
        Set-AzureADUserLicense -ObjectId $UserPrincipalName -AssignedLicenses $assignedLicenses
    }
    catch
    {
        $_.Exception.Data.Add('Parameters', $PSBoundParameters)
        $_.Exception.Data.Add('RetryCount', 3)
        $_.Exception.Data.Add('Delay', 5)
        throw
    }
}

function Remove-KBAMsolUserLicense
{
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]$UserPrincipalName,
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]$SamAccountName
    )
    try
    {
        $msolUser = Get-AzureADUser -ObjectId $UserPrincipalName
    }
    catch
    {
        $_.Exception.Data.Add('Parameters', $PSBoundParameters)
        $_.Exception.Data.Add('RetryCount', 3)
        $_.Exception.Data.Add('Delay', 10)
        throw
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
        $_.Exception.Data.Add('Parameters', $PSBoundParameters)
        $_.Exception.Data.Add('RetryCount', 1)
        $_.Exception.Data.Add('Delay', 3)
        throw
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
        $_.Exception.Data.Add('Parameters', $PSBoundParameters)
        throw
    }
    try
    {
        Set-AzureADUserLicense -ObjectId $UserPrincipalName -AssignedLicenses $assignedLicenses
    }
    catch
    {
        $_.Exception.Data.Add('Parameters', $PSBoundParameters)
        $_.Exception.Data.Add('RetryCount', 3)
        $_.Exception.Data.Add('Delay', 10)
        throw
    }
}

function Restore-KBAMsolUserLicense
{
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]$UserPrincipalName,
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]$SamAccountName
    )
    try
    {
        $adUser = Get-ADUser -Identity $SamAccountName -Properties @('ExtensionAttribute1')
    }
    catch
    {
        $_.Exception.Data.Add('Parameters', $PSBoundParameters)
        $_.Exception.Data.Add('RetryCount', 3)
        $_.Exception.Data.Add('Delay', 10)
        throw
    }
    if ($adUser.ExtensionAttribute1 -eq $null)
    {
        $e = New-Object -TypeName 'System.Exception' -ArgumentList @('No stashed license exists for user')
        $e.Data.Add('Parameters', $PSBoundParameters)
        throw $e
    }
    try
    {
        $license = $adUser.ExtensionAttribute1 | ConvertFrom-Json
    }
    catch
    {
        $_.Exception.Data.Add('Parameters', $PSBoundParameters)
        throw
    }
    try
    {
        Set-KBAMsolUserLicense -UserPrincipalName $UserPrincipalName -License $license
    }
    catch
    {
        $_.Exception.Data.Add('Parameters', $PSBoundParameters)
        $_.Exception.Data.Add('RetryCount', 3)
        $_.Exception.Data.Add('Delay', 10)
        throw
    }
    try
    {
        Set-ADUser -Identity $SamAccountName -Clear 'ExtensionAttribute1'
    }
    catch
    {
        $_.Exception.Data.Add('Parameters', $PSBoundParameters)
        $_.Exception.Data.Add('RetryCount', 1)
        $_.Exception.Data.Add('Delay', 3)
        throw
    }
}

function Get-KBAMsolPredefinedLicensePackage
{
    param
    (
        # TODO: turn into DymanicParam and use reflection to populate a ValidateSet:
        # $t = 'Kungsbacka.AccountTasks.MsolPredefinedLicensePackage' -as [type]
        # $t.GetMembers() | ? {$_.FieldType -eq ('Kungsbacka.AccountTasks.MsolLicense[]' -as [type])} | % Name
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [string[]]$Package
    )
    begin
    {
        Add-Type -Path 'Kungsbacka.AccountTasks.dll'
    }
    process
    {
        foreach ($item in $Package)
        {
            [Kungsbacka.AccountTasks.MsolPredefinedLicensePackage]::GetPackageFromName($Package)
        }
    }
}
