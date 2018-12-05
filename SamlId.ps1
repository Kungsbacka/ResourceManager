# Make all error terminating errors
$Global:ErrorActionPreference = 'Stop'

function Set-SamlId
{
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]
        $SamAccountName,
        [switch]
        $Force
    )
    $user = Get-ADUser -Identity $SamAccountName -Properties 'msDS-cloudExtensionAttribute14'
    $hasValue = $user.'msDS-cloudExtensionAttribute14' -ne $null
    if ($hasValue -and -not $Force)
    {
        throw "msDS-cloudExtensionAttribute14 is not empty. Use Force to overwrite."
    }
    $set = 'abcdefghijkmnpqrstuvxyz23456789'
    do
    {
        $id = ''
        for ($i = 0; $i -lt 16; $i++)
        {
            $id += $set.GetEnumerator() | Get-Random
        }
        $id += '@' + $Script:Config.SamlId.Domain
    }
    while ((Get-ADUser -Filter "msDS-cloudExtensionAttribute14 -eq '$id'"))
    if ($hasValue)
    {
        if ($Force)
        {
            Set-ADUser -Identity $SamAccountName -Replace @{'msDS-cloudExtensionAttribute14' = $id}
        }
    }
    else
    {
        Set-ADUser -Identity $SamAccountName -Add @{'msDS-cloudExtensionAttribute14' = $id}
    }
}
