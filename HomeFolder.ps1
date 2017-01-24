# Make all error terminating errors
$Global:ErrorActionPreference = 'Stop'
. "$PSScriptRoot\Config.ps1"

function New-HomeFolder
{
    [CmdLetBinding()]
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]$SamAccountName,
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]$Path
    )
    process
    {
        if ((Test-Path (Join-Path -Path $Path -ChildPath $SamAccountName)))
        {
            throw [pscustomobject]@{
                Target     = $SAMAccountName
                Activity   = $MyInvocation.MyCommand.Name
                Reason     = 'Folder already exists'
                Message    = "A folder with the name '$SamAccountName' already exists."
                RetryCount = 0
                Delay      = 0
            }
        }
        try
        {
            $homeFolder = New-Item -Path $Path -Name $SamAccountName -ItemType Directory
        }
        catch
        {
            throw [pscustomobject]@{
                Target     = (Join-Path $Path $SamAccountName)
                Activity   = $MyInvocation.MyCommand.Name
                Reason     = 'New-Item failed'
                Message    = $_.Exception.Message
                RetryCount = 3
                Delay      = 5
            }
        }
        try
        {
            @('Documents', 'Desktop', 'Favorites') | ForEach-Object {
                New-Item -Path $homeFolder.FullName -Name $_ -ItemType Directory | Out-Null
            }
        }
        catch
        {
            throw [pscustomobject]@{
                Target     = $homeFolder.FullName
                Activity   = $MyInvocation.MyCommand.Name
                Reason     = 'New-Item failed'
                Message    = $_.Exception.Message
                RetryCount = 0
                Delay      = 0
            }
        }
        icacls.exe "$($homeFolder.FullName)" /grant "$($Script:Config.HomeFolder.Domain)\$($SamAccountName):(OI)(CI)F" /Q | Out-Null
        if ($LASTEXITCODE -ne 0)
        {
            throw [pscustomobject]@{
                Target     = $homeFolder.FullName
                Activity   = $MyInvocation.MyCommand.Name
                Reason     = 'icacls.exe failed'
                Message    = "Grant full permissions failed with exit code $LASTEXITCODE"
                RetryCount = 0
                Delay      = 0
            }
        }
        icacls.exe "$($homeFolder.FullName)" /setowner "$($Script:Config.HomeFolder.Domain)\$($User.SamAccountName)" /T /Q | Out-Null
        if ($LASTEXITCODE -ne 0)
        {
            throw [pscustomobject]@{
                Target     = $homeFolder.FullName
                Activity   = $MyInvocation.MyCommand.Name
                Reason     = 'icacls.exe failed'
                Message    = "Set owner failed with exit code $LASTEXITCODE"
                RetryCount = 0
                Delay      = 0
            }
        }
        if ($homeFolder.FullName -like "$($Script:Config.HomeFolder.StudentShare)\*")
        {
            $param = @{
                Identity = $SamAccountName
                HomeDirectory = (Join-Path $homeFolder.FullName 'Documents')
                HomeDrive = 'H:'
            }
        }
        else
        {
            $param = @{
                Identity = $SamAccountName
                HomeDirectory = $null
                HomeDrive = $null
            }
        }
        try
        {
            Set-ADUser @param
        }
        catch
        {
            throw [pscustomobject]@{
                Target     = $SamAccountName
                Activity   = $MyInvocation.MyCommand.Name
                Reason     = 'Set-ADUser failed'
                Message    = $_.Exception.Message
                RetryCount = 0
                Delay      = 0
            }
        }
    }
}
