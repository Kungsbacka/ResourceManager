# Make all error terminating errors
$Global:ErrorActionPreference = 'Stop'

function Import-KBASkypeOnpremModule
{
    Import-Module -Name 'SkypeForBusiness' -Prefix 'Onprem'
}

function Test-KBAOnpremCSUser
{
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]
        $UserPrincipalName
    )
    $isEnabled = $false
    try
    {
        if ((Get-OnpremCsUser -Identity $UserPrincipalName))
        {
            $isEnabled = $true
        }
    }
    catch
    {
        if ($_.Exception.Message -notlike 'Management object not found*')
        {
            throw
        }
    }
    $isEnabled
}

function Grant-KBAOnpremCSConferencingPolicy
{
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]
        $UserPrincipalName,
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]
        $ConferencingPolicy
    )
    if (-not (Test-KBAOnpremCSUser $UserPrincipalName))
    {
        throw 'Target is not Skype enabled'
    }
    Grant-OnpremCsConferencingPolicy -Identity $UserPrincipalName -PolicyName $ConferencingPolicy
}

function Enable-KBAOnpremCSUser
{
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]
        $UserPrincipalName
    )
    process
    {
        if (Test-KBAOnpremCSUser $UserPrincipalName)
        {
            throw 'Target is already Skype enabled.'
        }
        $params = @{
            Identity = $UserPrincipalName
            RegistrarPool = $Script:Config.SkypeOnprem.RegistrarPool
            SipAddress = 'sip:' + $UserPrincipalName
        }
        Enable-OnpremCSUser @params
    }
}
