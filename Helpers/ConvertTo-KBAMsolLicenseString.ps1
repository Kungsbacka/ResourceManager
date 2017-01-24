[CmdLetBinding()]
param
(
    [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
    [object[]]$Licenses
)
begin
{
    $licenseArray = @()
}
process
{
    foreach ($license in $Licenses)
    {
        $licenseArray += $license
    }
}
end
{
    # Sorting is necessary to be able to compare license strings
    $licenseArray = $licenseArray | Sort-Object -Property AccountSkuId
    $currentLicense = ''
    foreach ($license in $LicenseArray)
    {
        $currentLicense += $license.AccountSkuId -replace "^[^:]+:", ''
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
    $currentLicense.TrimEnd('+')
}
