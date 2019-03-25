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

function Update-LicenseGroupCache
{
    function ParseGroup($group)
    {
        $json = $group.Location.Substring(8)
        $obj = ConvertFrom-Json -InputObject $json
        Add-Member -InputObject $obj -NotePropertyName 'Guid' -NotePropertyValue $group.ObjectGUID
        Add-Member -InputObject $obj -NotePropertyName 'Dn' -NotePropertyValue $group.DistinguishedName
        $obj
    }

    if ($Script:LicenseGroupCache)
    {
        return
    }
    $Script:LicenseGroupCache = @{
        GroupsByGuid = @{}
        GroupsByDn = @{}
    }
    $licenseGroups = Get-ADGroup -Filter "Location -like 'license:*'" -Properties @('Location')
    foreach ($group in $licenseGroups)
    {
        $metadata = ConvertFrom-Json -InputObject $group.Location.Substring(8)
        Add-Member -InputObject $metadata -NotePropertyName 'Guid' -NotePropertyValue $group.ObjectGUID
        Add-Member -InputObject $metadata -NotePropertyName 'Dn' -NotePropertyValue $group.DistinguishedName
        $Script:LicenseGroupCache.GroupsByGuid.Add($group.ObjectGUID.ToString(), $metadata)
        $Script:LicenseGroupCache.GroupsByDn.Add($group.DistinguishedName, $metadata)
    }
}

function Enable-RmMsolSync
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
        Properties = @('extensionAttribute11')
    }
    $user = Get-ADUser @params
    if ($null -eq $user.extensionAttribute11)
    {
        $params = @{
            Identity = $Identity
            Add = @{
                extensionAttribute11 = 'SYNC_ME'
            }
            Replace = @{
                # MsExchUsageLocation is synced to UsageLocation in Azure AD
                MsExchUsageLocation = $Script:Config.MicrosoftOnline.UsageLocation
            }
        }
        Set-ADUser @params
    }
}

function Set-RmMsolUserLicense
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

function Remove-RmMsolUserLicense
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
    Set-ADUser -Identity $SamAccountName -Replace @{'msDS-cloudExtensionAttribute1'=$licenseJson}
    $removeLicenses = New-Object 'System.Collections.Generic.List[string]'
    foreach ($item in $msolUser.AssignedLicenses)
    {
        $removeLicenses.Add($item.SkuId)
    }
    $assignedLicenses = New-Object -TypeName 'Microsoft.Open.AzureAD.Model.AssignedLicenses' -ArgumentList @($null, $removeLicenses)
    Set-AzureADUserLicense -ObjectId $UserPrincipalName -AssignedLicenses $assignedLicenses
}

function Restore-RmMsolUserLicense
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
    $adUser = Get-ADUser -Identity $SamAccountName -Properties @('msDS-cloudExtensionAttribute1')
    if ($null -eq $adUser.'msDS-cloudExtensionAttribute1')
    {
        throw 'No stashed license exists in msDS-cloudExtensionAttribute1'
    }
    $license = $adUser.'msDS-cloudExtensionAttribute1' | ConvertFrom-Json
    Set-RmMsolUserLicense -UserPrincipalName $UserPrincipalName -License $license
    Set-ADUser -Identity $SamAccountName -Clear 'msDS-cloudExtensionAttribute1'
}

function Set-RmLicenseGroupMembership
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
        [switch]
        $SkipSyncCheck,
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [switch]
        $SkipDynamicGroupCheck
    )
    $adUser = Get-ADUser -Identity $SamAccountName -Properties @('MemberOf','extensionAttribute11')
    if (-not $SkipSyncCheck -and $null -eq $adUser.extensionAttribute11)
    {
        throw 'User is not synced to AzureAD'
    }
    $currentMemberships = @{}
    foreach ($dn in $adUser.MemberOf)
    {
        $group = $Script:LicenseGroupCache.GroupsByDn[$dn]
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
        $group = $Script:LicenseGroupCache.GroupsByGuid[$groupGuid]
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

function Remove-RmLicenseGroupMembership
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
        [switch]
        $SkipBaseLicenseCheck
    )
    $adUser = Get-ADUser -Identity $SamAccountName -Properties @('MemberOf')
    $currentMemberships = @{}
    foreach ($dn in $adUser.MemberOf)
    {
        $group = $Script:LicenseGroupCache.GroupsByDn[$dn]
        if ($null -ne $group)
        {
            $currentMemberships.Add($group.Guid, $group)
        }
    }
    if ($currentMemberships.Count -eq 0)
    {
        return
    }
    $removeFrom = @()
    foreach ($groupGuid in $LicenseGroups)
    {
        $currentMemberships.Remove($groupGuid)
        $group = $Script:LicenseGroupCache.GroupsByGuid[$groupGuid]
        if ($null -eq $group)
        {
            throw 'Unknown license group'
        }
        $removeFrom += $group
    }
    if ($removeFrom.Count -eq 0)
    {
        return
    }
    if (-not $SkipBaseLicenseCheck)
    {
        $baseLicensePresent = $false
        foreach ($group in $currentMemberships)
        {
            if ($group.category -eq 'A')
            {
                $baseLicensePresent = $true
                break
            }
        }
        if (-not $baseLicensePresent)
        {
            throw 'Removing licenses would leave user without a base license. Use SkipBaseLicenseCheck to force removal.'
        }
    }
    foreach ($group in $removeFrom)
    {
        Remove-ADGroupMember -Identity $group.Dn -Members $adUser.DistinguishedName -Confirm:$false
    }
}

function Remove-RmAllLicenseGroupMembership
{
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]
        $SamAccountName,
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [switch]
        $SkipStashLicense
    )
    $adUser = Get-ADUser -Identity $SamAccountName -Properties @('MemberOf')
    $currentMemberships = @{}
    foreach ($dn in $adUser.MemberOf)
    {
        $group = $Script:LicenseGroupCache.GroupsByDn[$dn]
        if ($null -ne $group)
        {
            $currentMemberships.Add($group.Guid, $group)
        }
    }
    if ($currentMemberships.Count -eq 0)
    {
        return
    }
    if (-not $SkipStashLicense)
    {
        $serializedLicenses = $currentMemberships.Values.Guid -join ','
        Set-ADUser -Identity $SamAccountName -Add @{'msDS-cloudExtensionAttribute1'=$serializedLicenses}
    }
    foreach ($group in $currentMemberships.Values)
    {
        Remove-ADGroupMember -Identity $group.Dn -Members $adUser.DistinguishedName -Confirm:$false
    }
}

function Restore-RmLicenseGroupMembership
{
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]
        $SamAccountName,
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [switch]
        $SkipSyncCheck
    )
    $adUser = Get-ADUser -Identity $SamAccountName -Properties @('msDS-cloudExtensionAttribute1','extensionAttribute11')
    if (-not $SkipSyncCheck -and $null -eq $adUser.extensionAttribute11)
    {
        throw 'User is not synced to AzureAD'
    }
    if ($null -eq $adUser.'msDS-cloudExtensionAttribute1')
    {
        throw 'User has no stashed licenses in msDS-cloudExtensionAttribute1'
    }
    $deserializedLicenses = $adUser.'msDS-cloudExtensionAttribute1' -split ','
    Set-RmLicenseGroupMembership -SamAccountName $SamAccountName -LicenseGroups $deserializedLicenses -SkipDynamicGroupCheck -SkipSyncCheck:$SkipSyncCheck
    Set-ADUser -Identity $SamAccountName -Clear 'msDS-cloudExtensionAttribute1'
}
