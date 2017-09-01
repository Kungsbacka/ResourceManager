# Make all error terminating errors
$Global:ErrorActionPreference = 'Stop'
. "$PSScriptRoot\Config.ps1"

function New-HomeFolder
{
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]
        $SamAccountName,
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]
        $Path
    )
    $homeFolderPath = Join-Path -Path $Path -ChildPath $SamAccountName
    if ((Test-Path $homeFolderPath))
    {
        throw "Folder already exists: $homeFolderPath"
    }
    New-Item -Path $homeFolderPath -ItemType Directory | Out-Null
    @('Documents', 'Desktop', 'Favorites') | ForEach-Object {
        New-Item -Path $homeFolderPath -Name $_ -ItemType Directory | Out-Null
    }
    icacls.exe "$($homeFolderPath)" /grant "$($Script:Config.HomeFolder.Domain)\$($SamAccountName):(OI)(CI)F" /Q | Out-Null
    if ($LASTEXITCODE -ne 0)
    {
        throw "icacls failed ($LASTEXITCODE) to grant permissions on folder: $homeFolderPath"
    }
    icacls.exe "$($homeFolderPath)" /setowner "$($Script:Config.HomeFolder.Domain)\$($User.SamAccountName)" /T /Q | Out-Null
    if ($LASTEXITCODE -ne 0)
    {
        throw "icacls failed ($LASTEXITCODE) to set owner on folder: $homeFolderPath"
    }
}
