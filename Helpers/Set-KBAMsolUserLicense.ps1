[CmdLetBinding(SupportsShouldProcess=$true,ConfirmImpact='Medium')]
param
(
    [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
    [ValidateNotNullOrEmpty()]
    [string]
    $UserPrincipalName,
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]
    $AccountName,
    [Parameter(Mandatory=$true,ParameterSetName='CustomLicense')]
    [ValidateNotNullOrEmpty()]
    [string]
    $CustomLicense,
    [Parameter(Mandatory=$true,ParameterSetName='PredefinedLicense')]
    [ValidateSet('Faculty','Student','EMS')]
    [string]
    $PredefinedLicense
)
begin
{
    if ($CustomLicense)
    {
        $licenseString = $CustomLicense
    }
    else
    {
        switch ($PredefinedLicense)
        {
            'Faculty' 
                {$licenseString = 'OFFICESUBSCRIPTION_FACULTY+STANDARDWOFFPACK_FACULTY(MCOSTANDARD,EXCHANGE_S_STANDARD)'}
            'Student'
                {$licenseString = 'OFFICESUBSCRIPTION_STUDENT+STANDARDWOFFPACK_STUDENT'}
            'EMS'
                {$licenseString = 'EMS'}
        }
    }
    $licensesToAdd = @()
    $licenseOptions = @()
    $skus = $licenseString -split '\+'
    foreach ($sku in $skus)
    {
        if ($sku -match '\((?<disabled>[^)]+)\)')
        {
            $disabledPlans = $Matches['disabled'] -split ','
            $sku = $sku -replace '\([^)]+\)', ''
            $licenseOptions += New-MsolLicenseOptions -AccountSkuId "$($AccountName):$sku" -DisabledPlans $disabledPlans
        }
        $licensesToAdd += "$($AccountName):$sku"
    }
}
process
{
    foreach ($User in $UserPrincipalName)
    {
        $msolUser = Get-MsolUser -UserPrincipalName $User
        # Sorting is necessary to be able to compare license strings
        $licenseArray = $msolUser.Licenses | Sort-Object -Property AccountSkuId
        $currentLicense = ''
        foreach ($license in $msolUser.Licenses)
        {
            $currentLicense += $license.AccountSkuId -replace "$($AccountName):", ''
            $disabledPlans = '('
            foreach($status in $license.ServiceStatus)
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
        if ($LicenseString -eq $currentLicense.TrimEnd('+'))
        {
            Write-Verbose -Message "$User already has this license. No change required."
            return
        }
        # Set-MsolUserLicense fails if usage location is not set
        if ($msolUser.UsageLocation -eq $null)
        {
            Write-Verbose -Message "Setting usage location to 'SE' for $User"
            if ($PSCmdlet.ShouldProcess($User, 'Set usage location to "SE"'))
            {
                Set-MsolUser -UserPrincipalName $User -UsageLocation 'SE'
            }
        }   
        $licensesToRemove = $msolUser.Licenses.AccountSkuId
        Write-Verbose -Message "Setting $($licensesToAdd.Count) and removing $($licensesToRemove.Count) license(s) for $User"
        if ($PSCmdlet.ShouldProcess($User, 'Replace licenses'))
        {
            # This is the only way I was able to replace all licenses.
            # Calling Set-MsolUserLicense with all parameters at once did not work reliably.
            Set-MsolUserLicense -UserPrincipalName $UserPrincipalName -RemoveLicenses $licensesToRemove
            Set-MsolUserLicense -UserPrincipalName $UserPrincipalName -AddLicenses $licensesToAdd
            Set-MsolUserLicense -UserPrincipalName $UserPrincipalName -LicenseOptions $licenseOptions
        }
    }
}
