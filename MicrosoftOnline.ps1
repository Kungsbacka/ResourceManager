# Make all error terminating errors
$Global:ErrorActionPreference = 'Stop'

function Connect-KBAAzureAD {
    Import-Module -Name 'AzureAD' -DisableNameChecking
    $credential = New-Object -TypeName 'System.Management.Automation.PSCredential' -ArgumentList @(
        $Script:Config.AzuerAD.User
        $Script:Config.AzureAD.Password | ConvertTo-SecureString
    )
    Connect-AzureAD -Credential $credential -LogLevel None | Out-Null
}

function Update-LicenseGroupCache {
    if ($Script:LicenseGroupCache) {
        return
    }
    $Script:LicenseGroupCache = @{
        GroupsByGuid = @{}
        GroupsByDn   = @{}
    }
    $licenseGroups = Get-ADGroup -Filter "Location -like 'license:*'" -Properties @('Location')
    foreach ($group in $licenseGroups) {
        $metadata = ConvertFrom-Json -InputObject $group.Location.Substring(8)
        Add-Member -InputObject $metadata -NotePropertyName 'Guid' -NotePropertyValue $group.ObjectGUID
        Add-Member -InputObject $metadata -NotePropertyName 'Dn' -NotePropertyValue $group.DistinguishedName
        $Script:LicenseGroupCache.GroupsByGuid.Add($group.ObjectGUID.ToString(), $metadata)
        $Script:LicenseGroupCache.GroupsByDn.Add($group.DistinguishedName, $metadata)
    }
}

function Enable-RmMsolSync {
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('SamAccountName', 'ObjectGUID', 'DistinguishedName')]
        [string]
        $Identity
    )
    $params = @{
        Identity   = $Identity
        Properties = @('extensionAttribute11')
    }
    $user = Get-ADUser @params
    if ($null -eq $user.extensionAttribute11) {
        $params = @{
            Identity = $Identity
            Add      = @{
                extensionAttribute11 = 'SYNC_ME'
            }
            Replace  = @{
                # MsExchUsageLocation is synced to UsageLocation in Azure AD
                MsExchUsageLocation = $Script:Config.MicrosoftOnline.UsageLocation
            }
        }
        Set-ADUser @params
    }
}

function Set-RmMsolUserLicense {
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]
        $UserPrincipalName,
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [object]
        $License
    )
    $aadUser = Get-AzureADUser -ObjectId $UserPrincipalName
    if ($aadUser.UsageLocation -ne 'SE') {
        # If MsExchUsageLocation hasn't synced yet, we set UsageLocation explicitly
        Set-AzureADUser -ObjectId $UserPrincipalName -UsageLocation $Script:Config.MicrosoftOnline.UsageLocation
    }
    $removeLicenses = New-Object 'System.Collections.Generic.List[string]'
    foreach ($item in $aadUser.AssignedLicenses) {
        if ($item.SkuId -notin $License.SkuId) {
            $removeLicenses.Add($item.SkuId)
        }
    }
    $addLicenses = New-Object 'System.Collections.Generic.List[Microsoft.Open.AzureAD.Model.AssignedLicense]'
    foreach ($item in $License) {
        $addLicenses.Add((New-Object -TypeName 'Microsoft.Open.AzureAD.Model.AssignedLicense' -ArgumentList @($item.DisabledPlans, $item.SkuId)))
    }
    $assignedLicenses = New-Object -TypeName 'Microsoft.Open.AzureAD.Model.AssignedLicenses' -ArgumentList @($addLicenses, $removeLicenses)
    Set-AzureADUserLicense -ObjectId $UserPrincipalName -AssignedLicenses $assignedLicenses
}

function Remove-RmMsolUserLicense {
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]
        $UserPrincipalName,
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]
        $SamAccountName
    )
    $msolUser = Get-AzureADUser -ObjectId $UserPrincipalName
    $licenseJson = ConvertTo-Json -InputObject @($msolUser.AssignedLicenses) -Compress
    $licenseJson = $licenseJson.ToString()
    if ($licenseJson -eq '[]') {
        return # User has no license
    }
    Set-ADUser -Identity $SamAccountName -Replace @{'msDS-cloudExtensionAttribute1' = $licenseJson }
    $removeLicenses = New-Object 'System.Collections.Generic.List[string]'
    foreach ($item in $msolUser.AssignedLicenses) {
        $removeLicenses.Add($item.SkuId)
    }
    $assignedLicenses = New-Object -TypeName 'Microsoft.Open.AzureAD.Model.AssignedLicenses' -ArgumentList @($null, $removeLicenses)
    Set-AzureADUserLicense -ObjectId $UserPrincipalName -AssignedLicenses $assignedLicenses
}

function Restore-RmMsolUserLicense {
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]
        $UserPrincipalName,
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]
        $SamAccountName
    )
    $adUser = Get-ADUser -Identity $SamAccountName -Properties @('msDS-cloudExtensionAttribute1')
    if ($null -eq $adUser.'msDS-cloudExtensionAttribute1') {
        throw 'No stashed license exists in msDS-cloudExtensionAttribute1'
        return
    }
    $license = $adUser.'msDS-cloudExtensionAttribute1' | ConvertFrom-Json
    Set-RmMsolUserLicense -UserPrincipalName $UserPrincipalName -License $license
    Set-ADUser -Identity $SamAccountName -Clear 'msDS-cloudExtensionAttribute1'
}

function Set-RmLicenseGroupMembership {
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]
        $SamAccountName,
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string[]]
        $LicenseGroups,
        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [switch]
        $SkipSyncCheck,
        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [switch]
        $SkipDynamicGroupCheck
    )
    $adUser = Get-ADUser -Identity $SamAccountName -Properties @('MemberOf', 'extensionAttribute11')
    if (-not $SkipSyncCheck -and $null -eq $adUser.extensionAttribute11) {
        throw 'User is not synced to AzureAD'
        return
    }
    $requestedLicenseGroups = @{}
    $categories = @{}
    foreach ($guid in $LicenseGroups) {
        $group = $Script:LicenseGroupCache.GroupsByGuid[$guid]
        if ($null -eq $group) {
            throw 'Unknown license group'
            return
        }
        if ($categories[$group.Category]) {
            throw 'Cannot add a user to more than one license group in each category'
            return
        }
        else {
            $categories[$group.Category] = $true
        }
        if ($group.Dynamic -and -not $SkipDynamicGroupCheck) {
            throw 'Cannot add a user to a dynamic license group'
            return
        }
        $requestedLicenseGroups.Add($group.Guid, $group)
    }
    $currentLicenseGroups = @{}
    foreach ($dn in $adUser.MemberOf) {
        $group = $Script:LicenseGroupCache.GroupsByDn[$dn]
        if ($null -ne $group) {
            $currentLicenseGroups.Add($group.Guid, $group)
        }
    }
    $removeFrom = @()
    foreach ($group in $currentLicenseGroups.Values) {
        if (-not $requestedLicenseGroups.ContainsKey($group.Guid)) {
            if ($group.Dynamic) {
                throw 'Cannot remove a user from a dynamic license group'
                return
            }
            $removeFrom += $group
        }
    }
    $addTo = @()
    foreach ($group in $requestedLicenseGroups.Values) {
        if (-not $currentLicenseGroups.ContainsKey($group.Guid)) {
            $addTo += $group
        }
    }
    foreach ($group in $removeFrom) {
        Remove-ADGroupMember -Identity $group.Dn -Members $adUser.DistinguishedName -Confirm:$false
    }
    foreach ($group in $addTo) {
        Add-ADGroupMember -Identity $group.Dn -Members $adUser.DistinguishedName
    }
}

function Remove-RmLicenseGroupMembership {
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]
        $SamAccountName,
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string[]]
        $LicenseGroups,
        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [switch]
        $SkipBaseLicenseCheck
    )
    $adUser = Get-ADUser -Identity $SamAccountName -Properties @('MemberOf')
    $currentMemberships = @{}
    foreach ($dn in $adUser.MemberOf) {
        $group = $Script:LicenseGroupCache.GroupsByDn[$dn]
        if ($null -ne $group) {
            $currentMemberships.Add($group.Guid, $group)
        }
    }
    if ($currentMemberships.Count -eq 0) {
        return
    }
    $removeFrom = @()
    foreach ($groupGuid in $LicenseGroups) {
        $currentMemberships.Remove($groupGuid)
        $group = $Script:LicenseGroupCache.GroupsByGuid[$groupGuid]
        if ($null -eq $group) {
            throw 'Unknown license group'
            return
        }
        $removeFrom += $group
    }
    if ($removeFrom.Count -eq 0) {
        return
    }
    if (-not $SkipBaseLicenseCheck) {
        $baseLicensePresent = $false
        foreach ($group in $currentMemberships) {
            if ($group.category -eq 'A') {
                $baseLicensePresent = $true
                break
            }
        }
        if (-not $baseLicensePresent) {
            throw 'Removing licenses would leave user without a base license. Use SkipBaseLicenseCheck to force removal.'
            return
        }
    }
    foreach ($group in $removeFrom) {
        Remove-ADGroupMember -Identity $group.Dn -Members $adUser.DistinguishedName -Confirm:$false
    }
}

function Remove-RmAllLicenseGroupMembership {
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]
        $SamAccountName,
        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [switch]
        $SkipStashLicense
    )
    $adUser = Get-ADUser -Identity $SamAccountName -Properties @('MemberOf')
    $currentMemberships = @{}
    foreach ($dn in $adUser.MemberOf) {
        $group = $Script:LicenseGroupCache.GroupsByDn[$dn]
        if ($null -ne $group) {
            $currentMemberships.Add($group.Guid, $group)
        }
    }
    if ($currentMemberships.Count -eq 0) {
        return
    }
    if (-not $SkipStashLicense) {
        $serializedLicenses = $currentMemberships.Values.Guid -join ','
        Set-ADUser -Identity $SamAccountName -Add @{'msDS-cloudExtensionAttribute1' = $serializedLicenses }
    }
    foreach ($group in $currentMemberships.Values) {
        Remove-ADGroupMember -Identity $group.Dn -Members $adUser.DistinguishedName -Confirm:$false
    }
}

function Restore-RmLicenseGroupMembership {
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]
        $SamAccountName,
        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [switch]
        $SkipSyncCheck
    )
    $adUser = Get-ADUser -Identity $SamAccountName -Properties @('msDS-cloudExtensionAttribute1', 'extensionAttribute11', 'memberOf')
    if (-not $SkipSyncCheck -and $null -eq $adUser.extensionAttribute11) {
        throw 'User is not synced to AzureAD'
        return
    }

    if ($null -eq $adUser.'msDS-cloudExtensionAttribute1') {
        $isMemberOfLicenseGroup = $false
        foreach ($group in $adUser.memberOf) {
            $licenseGroup = $Script:LicenseGroupCache.GroupsByDn[$group]
            if ($null -ne $licenseGroup -and $licenseGroup.Category -eq 'A' -and $licenseGroup.MailEnabled) {
                $isMemberOfLicenseGroup = $true
                break
            }
        }
        
        if (-not $isMemberOfLicenseGroup) {
            throw 'User has no stashed licenses in msDS-cloudExtensionAttribute1 and is not a member of a mail-enabled license group.'
            return
        }
        
    }
    else {
        $deserializedLicenses = $adUser.'msDS-cloudExtensionAttribute1' -split ','
        Set-RmLicenseGroupMembership -SamAccountName $SamAccountName -LicenseGroups $deserializedLicenses -SkipDynamicGroupCheck -SkipSyncCheck:$SkipSyncCheck
        Set-ADUser -Identity $SamAccountName -Clear 'msDS-cloudExtensionAttribute1'
    }
}
