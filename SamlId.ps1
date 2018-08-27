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
    $user = Get-ADUser -Identity $SamAccountName -Properties 'ExtensionAttribute14'
    $hasValue = $user.ExtensionAttribute14 -ne $null
    if ($hasValue -and -not $Force)
    {
        throw "ExtensionAttribute14 is not empty. Use Force to overwrite."
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
    while ((Get-ADUser -Filter {ExtensionAttribute14 -eq $id}))
    if ($hasValue)
    {
        if ($Force)
        {
            Set-ADUser -Identity $SamAccountName -Replace @{ExtensionAttribute14 = $id}
        }
    }
    else
    {
        Set-ADUser -Identity $SamAccountName -Add @{ExtensionAttribute14 = $id}
    }
}
