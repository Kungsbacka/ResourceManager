# Make all error terminating errors
$Global:ErrorActionPreference = 'Stop'

function Add-OnpremGroupMember
{
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]$SamAccountName,
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]$Group
    )
    $params = @{
        Identity = $Group
        Members = @($SamAccountName)
    }
    Add-ADGroupMember @params
}

function Remove-OnpremGroupMember
{
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]$SamAccountName,
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]$Group
    )
    $params = @{
        Identity = $Group
        Members = @($SamAccountName)
        Confirm = $false
    }
    Remove-ADGroupMember @params
}
