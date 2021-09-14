# Make all error terminating errors
$Global:ErrorActionPreference = 'Stop'

function New-RmSamlId
{
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [string]
        $Domain
    )
    $rng = [System.Security.Cryptography.RNGCryptoServiceProvider]::Create()
    [byte[]]$bytes = [byte[]]::new(4)
    $sb = [System.Text.StringBuilder]::new()

    # Borrowed from the CryptoRandom class found here:
    # https://docs.microsoft.com/en-us/archive/msdn-magazine/2007/september/net-matters-tales-from-the-cryptorandom
    function Rnd([int32]$minValue, [int32]$maxValue)
    {
        if ($minValue -gt $maxValue)
        {
            throw 'Min cannot be greater than max'
        }
        if ($minValue -eq $maxValue)
        {
            $minValue
            return
        }
        [int64]$diff = $maxValue - $minValue
        while ($true)
        {
            $rng.GetBytes($bytes)
            [uint32]$rand = [System.BitConverter]::ToUInt32($bytes, 0)
            [int64]$max = 1 + [long][int32]::MaxValue
            [int64]$remainder = $max % $diff
            if ($rand -lt $max - $remainder)
            {
                [int32]($minValue + ($rand % $diff))
                return
            }
        }
    }

    $set = 'abcdefghijkmnpqrstuvxyz23456789'
    $len = $set.Length
    for ($i = 0; $i -lt 16; $i++) {
        $n = Rnd -minValue 0 -maxValue $len
        $null = $sb.Append($set[$n])
    }
    if ($Domain -notlike '@*')
    {
        $null = $sb.Append('@')
    }
    $null = $sb.Append($Domain)
    $sb.ToString()
}

function Test-RmSamlId
{
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [string]
        $SamlId
    )
    try
    {
        $conn = New-Object -TypeName 'System.Data.SqlClient.SqlConnection' -ArgumentList @($Script:Config.MetaDirectory.ConnectionString)
        $conn.Open()
        $cmd = New-Object -TypeName 'System.Data.SqlClient.SqlCommand'
        $cmd.Connection = $conn
        $cmd.CommandType = 'StoredProcedure'
        $cmd.CommandText = 'dbo.spIsSamlIdReserved'
        $null = $cmd.Parameters.AddWithValue('@samlId', $SamlId)
        $param = $cmd.Parameters.Add('@rc', 'int')
        $param.Direction = 'ReturnValue'
        $null = $cmd.ExecuteNonQuery()
        if ($cmd.Parameters["@rc"].Value -eq 1) # 1 = SAML-ID i reserved and cannot be reused
        {
            $false
            return
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

    if ((Get-ADUser -Filter "msDS-cloudExtensionAttribute14 -eq '$SamlId'"))
    {
        $false
        return
    }

    $true
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
    do
    {
        $samlId = New-RmSamlId -Domain $Script:Config.SamlId.Domain
    }
    while (-not (Test-RmSamlId -SamlId $samlId))
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
