# Make all error terminating errors
$Global:ErrorActionPreference = 'Stop'

function Get-SamlIdFromPool
{
    try
    {
        $conn = New-Object -TypeName 'System.Data.SqlClient.SqlConnection' -ArgumentList @($Script:Config.MetaDirectory.ConnectionString)
        $conn.Open()
        $cmd = New-Object -TypeName 'System.Data.SqlClient.SqlCommand'
        $cmd.Connection = $conn
        $cmd.CommandType = 'StoredProcedure'
        $cmd.CommandText = 'dbo.spGetSamlIdFromPool'
        $rdr = $cmd.ExecuteReader()
        if ($rdr.Read())
        {
            $rdr.GetString(0)
        }
    }
    finally
    {
        if ($cmd)
        {
            $cmd.Dispose()
            $cmd = $null
        }
        if ($conn)
        {
            $conn.Dispose()
            $conn = $null
        }
    }
}

function Set-RmSamlId
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
    $hasValue = $null -ne $user.'msDS-cloudExtensionAttribute14'
    if ($hasValue -and -not $Force)
    {
        throw "msDS-cloudExtensionAttribute14 is not empty. Use Force to overwrite."
    }
    $samlId = Get-SamlIdFromPool
    if ($hasValue)
    {
        if ($Force)
        {
            Set-ADUser -Identity $SamAccountName -Replace @{'msDS-cloudExtensionAttribute14' = $samlId}
        }
    }
    else
    {
        Set-ADUser -Identity $SamAccountName -Add @{'msDS-cloudExtensionAttribute14' = $samlId}
    }
}
