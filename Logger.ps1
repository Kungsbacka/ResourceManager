# Make all error terminating errors
$Global:ErrorActionPreference = 'Stop'
. "$PSScriptRoot\Config.ps1"

Enum TaskResult
{
    Success
    Failure
    Wait
}

function Update-TaskLogEntry
{
    param
    (
        [Parameter(Mandatory=$true)]
        [int]
        $TaskId,
        [Parameter(Mandatory=$true)]
        [TaskResult]
        $Result,
        [switch]
        $EndTask
    )
    process
    {
        if ($EndTask)
        {
            $end = 1
        }
        else
        {
            $end = 0
        }
        Invoke-StoredProcedure -Procedure 'dbo.spUpdateTaskLogEntry' -Parameters @{
            TaskId = $TaskId
            Status = $Result.ToString()
            End = $end
        }
    }
}

function New-TaskLogEntry
{
    param
    (
        [Parameter(Mandatory=$true)]
        [string]
        $Task,
        [string]
        $Target,
        [TaskResult]
        $Result
    )
    process
    {
        $id = Invoke-StoredProcedure -Procedure 'dbo.spInsertNewTaskLogEntry' -Scalar -Parameters @{
            Task = $Task
            Target = $Target
        }
        if ($Result)
        {
            Invoke-StoredProcedure -Procedure 'dbo.spUpdateTaskLogEntry' -Parameters @{
                TaskId = $id
                Status = $Result.ToString()
                End = 1
            }
        }
        else
        {
            $id
        }
    }
}

function Write-ErrorLog
{
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,ParameterSetName='ErrorRecord')]
        [object]
        $ErrorRecord,
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true,ParameterSetName='CustomError')]
        [string]
        $Message,
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string]
        $Target,
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string]
        $TaskJson
    )
    if ($Message)
    {
        $params = @{
            Target = $Target
            Message = $Message
            TaskJson = $TaskJson
        }
    }
    else
    {
        $params = @{
            Target = $Target
            Message = $ErrorRecord.Exception.ToString()
            ScriptStackTrace = $ErrorRecord.ScriptStackTrace
            TaskJson = $TaskJson
        }
    }
    $text = Get-ErrorText @params
    $currentLog = Get-ChildItem -Path $Script:Config.Logger.LogPath -Filter '*.log' |
        Sort-Object -Property LastWriteTime -Descending | Select-Object -First 1
    if ($currentLog.LastWriteTime -lt ((Get-Date).AddMonths(-1)))
    {
        $fileName = 'rmgr_' + (Get-Date).ToString('yyyyMMdd_HHmmss') + '.log'
        $newLogFile = Join-Path -Path $Script:Config.Logger.LogPath -ChildPath $fileName
        $currentLog = New-Item -Path $newLogFile -ItemType File
    }
    $text | Out-File -FilePath $currentLog.FullName -Encoding UTF8 -Append
}

function Get-ErrorText
{
    param
    (
        [string]$Target,
        [string]$Message,
        [string]$ScriptStackTrace,
        [string]$TaskJson
    )
    if ($TaskJson)
    {
        # Make sure JSON is properly formatted
        $obj = $TaskJson | ConvertFrom-Json
        $TaskJson = $obj | ConvertTo-NewtonsoftJson -Formatting Indented
    }
@"
---------------------------------------------------
---- {0}
---------------------------------------------------
##Target: {1}

##Exception:
{2}

##Script stacktrace:
{3}

##Task JSON:
{4}

"@ -f (Get-Date), $Target, $Message, $ScriptStackTrace, $TaskJson
}

function Invoke-StoredProcedure
{
    param
    (
        [Parameter(Mandatory = $true)]
        [string]
        $Procedure,
        [Parameter(Mandatory = $true)]
        [object]
        $Parameters,
        [Parameter(Mandatory = $false)]
        [switch]
        $Scalar
    )
    process
    {
        $conn = New-Object -TypeName 'System.Data.SqlClient.SqlConnection'
        $conn.ConnectionString = $Script:Config.Logger.ConnectionString
        $conn.Open()
        $cmd = New-Object -TypeName 'System.Data.SqlClient.SqlCommand'
        $cmd.Connection = $conn
        $cmd.CommandText = $Procedure
        $cmd.CommandType = [System.Data.CommandType]::StoredProcedure
        foreach ($key in $Parameters.Keys)
        {
            [void]$cmd.Parameters.AddWithValue($key, $Parameters[$key])
        }
        if ($Scalar)
        {
            $cmd.ExecuteScalar()
        }
        else
        {
            [void]$cmd.ExecuteNonQuery()
        }
    }
}

function ConvertTo-NewtonsoftJson
{
    param
    (
        [Parameter(Mandatory=$true, ValueFromPipeline=$true, Position=0)]
        [object]
        $InputObject,
        [Parameter(ValueFromPipelineByPropertyName=$true, Position=1)]
        [Newtonsoft.Json.Formatting]
        $Formatting = [Newtonsoft.Json.Formatting]::Indented
    )
    [Newtonsoft.Json.JsonConvert]::SerializeObject($InputObject, $Formatting) | Write-Output
}
