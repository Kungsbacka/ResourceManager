[CmdLetBinding()]
param
(
    [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
    [object[]]$LicenseString,
    [Parameter(Mandatory=$true)]
    [string]$AccountName
)
process
{
    foreach ($str in $LicenseString)
    {
        $licenses = @()
        $licenseOptions = @()
        $skuIds = $str -split '\+'
        foreach ($skuId in $skuIds)
        {
            if ($skuId -match '\((?<plans>[^)]+)\)')
            {
                $disabledPlans = $Matches['plans'] -split ','
                $skuId = $skuId -replace '\([^)]+\)', ''
                $licenseOptions += New-MsolLicenseOptions -AccountSkuId "$($AccountName):$skuId" -DisabledPlans $disabledPlans
            }
            $licenses += "$($AccountName):$skuId"
        }
        [pscustomobject]@{
            Licenses = $licenses
            LicenseOptions = $licenseOptions
        }
    }
}
