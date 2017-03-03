param
(
    # TODO: turn into DymanicParam and use reflection to populate a ValidateSet:
    # $t = 'Kungsbacka.AccountTasks.MsolPredefinedLicensePackage' -as [type]
    # $t.GetMembers() | ? {$_.FieldType -eq ('Kungsbacka.AccountTasks.MsolLicense[]' -as [type])} | % Name
    [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
    [string[]]
    $Package
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
