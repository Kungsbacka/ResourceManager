# Make all error terminating errors
$Global:ErrorActionPreference = 'Stop'

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
    if ($null -eq $user.ExtensionAttribute11)
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
        if ($item.SkuId -notin $License.SkuId)
        {
            $removeLicenses.Add($item.SkuId)
        }
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
    if ($null -eq $adUser.ExtensionAttribute1)
    {
        throw 'No stashed license exists in extensionAttribute1'
    }
    $license = $adUser.ExtensionAttribute1 | ConvertFrom-Json
    Set-KBAMsolUserLicense -UserPrincipalName $UserPrincipalName -License $license
    Set-ADUser -Identity $SamAccountName -Clear 'ExtensionAttribute1'
}

function Set-LicenseGroupMembership
{
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]
        $SamAccountName,
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string[]]
        $LicenseGroups,
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [bool]
        $SkipSyncCheck = $false,
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [bool]
        $SkipDynamicGroupCheck = $false
    )
    function ParseGroup($group)
    {
        $json = $group.Location.Substring(8)
        $obj = ConvertFrom-Json -InputObject $json
        Add-Member -InputObject $obj -NotePropertyName 'Guid' -NotePropertyValue $group.ObjectGUID
        Add-Member -InputObject $obj -NotePropertyName 'Dn' -NotePropertyValue $group.DistinguishedName
        $obj
    }
    $adUser = Get-ADUser -Identity $SamAccountName -Properties @('MemberOf','ExtensionAttribute11')
    if (-not $SkipSyncCheck -and $null -eq $adUser.ExtensionAttribute11)
    {
        throw 'User is not synced to AzureAD'
    }
    $allLicenseGroups = Get-ADGroup -Filter "Location -like 'license:*'" -Properties @('Location')
    $guidHash = @{}
    $dnHash = @{}
    foreach ($group in $allLicenseGroups)
    {
        $parsedGroup = ParseGroup $group
        $guidHash.Add($group.ObjectGUID.ToString(), $parsedGroup)
        $dnHash.Add($group.DistinguishedName, $parsedGroup)
    }
    $currentMemberships = @{}
    foreach ($dn in $adUser.MemberOf)
    {
        $group = $dnHash[$dn]
        if ($null -ne $group)
        {
            $currentMemberships.Add($group.Guid, $group)
        }
    }
    $addTo = @()
    $removeFrom = @()
    $categories = @{}
    foreach ($groupGuid in $LicenseGroups)
    {
        $group = $guidHash[$groupGuid] 
        if ($null -eq $group)
        {
            throw 'Unknown license group'
        }
        if ($categories[$group.Category])
        {
            throw 'Cannot add a user to more than one license group in each category'
        }
        else
        {
            $categories[$group.Category] = $true
        }
        if ($group.Dynamic -and -not $SkipDynamicGroupCheck)
        {
            throw 'Cannot add a user to a dynamic license group'
        }
        if ($currentMemberships.ContainsKey($group.Guid))
        {
            continue
        }
        foreach ($memberGroup in $currentMemberships.Values)
        {
            if ($memberGroup.Category -eq $group.Category)
            {
                if ($memberGroup.Dynamic)
                {
                    throw 'Cannot remove a user from a dynamic license group'
                }
                $removeFrom += $memberGroup
            }
        }
        $addTo += $group
    }
    if (Compare-Object -ReferenceObject $addTo -DifferenceObject $removeFrom -Property 'Guid')
    {
        if ($removeFrom.Count -gt 0)
        {
            foreach ($group in $removeFrom)
            {
                Remove-ADGroupMember -Identity $group.Dn -Members $adUser.DistinguishedName -Confirm:$false
            }
        }
        if ($addTo.Count -gt 0)
        {
            foreach ($group in $addTo)
            {
                Add-ADGroupMember -Identity $group.Dn -Members $adUser.DistinguishedName
            }
        }
    }
}
