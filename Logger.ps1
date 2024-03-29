﻿# Make all error terminating errors
$Global:ErrorActionPreference = 'Stop'

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
        $EndTask,
        [string]
        $Details
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
        Invoke-StoredProcedure -Procedure 'dbo.spRmUpdateLogEntry' -Parameters @{
            TaskId = $TaskId
            Status = $Result.ToString()
            End = $end
            Details = $Details
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
        $Result,
        [string]
        $Details
    )
    process
    {
        $id = Invoke-StoredProcedure -Procedure 'dbo.spRmNewLogEntry' -Scalar -Parameters @{
            Task = $Task
            Target = $Target
        }
        if ($Result)
        {
            Invoke-StoredProcedure -Procedure 'dbo.spRmUpdateLogEntry' -Parameters @{
                TaskId = $id
                Status = $Result.ToString()
                End = 1
                Details = $Details
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
        [int]
        $TaskId,
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
            TaskId = $TaskId
            TaskJson = $TaskJson
        }
    }
    else
    {
        $params = @{
            Target = $Target
            Message = $ErrorRecord.Exception.ToString()
            TaskId = $TaskId
            ScriptStackTrace = $ErrorRecord.ScriptStackTrace
            TaskJson = $TaskJson
        }
    }
    $text = Get-ErrorText @params
    $currentLog = Get-ChildItem -Path $Script:Config.Logger.LogPath -Filter 'rmgr_*.log' |
        Sort-Object -Property LastWriteTime -Descending | Select-Object -First 1
    if ($currentLog.CreationTime -lt ((Get-Date).AddDays(-7)))
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
        [string]$TaskId,
        [string]$Target,
        [string]$Message,
        [string]$ScriptStackTrace,
        [string]$TaskJson
    )
    if ($TaskJson)
    {
        # Make sure JSON is properly formatted
        # May be invalid JSON
        try
        {
        $obj = $TaskJson | ConvertFrom-Json
        }
        catch
        {
            $obj = [pscustomobject]@{InvalidJson = $TaskJson}
        }
        $TaskJson = $obj | ConvertTo-Json -Depth 4
    }
    if ($TaskId)
    {
        $idString = 'ID #' + $TaskId
    }
    else
    {
        $idString = ''
    }
@"
---------------------------------------------------
---- {0}  {1}
---------------------------------------------------
##Target: {2}

##Exception:
{3}

##Script stacktrace:
{4}

##Task JSON:
{5}

"@ -f (Get-Date), $idString, $Target, $Message, $ScriptStackTrace, $TaskJson
}

function Close-UnclosedFailedTasks
{
    Invoke-StoredProcedure -Procedure 'dbo.spRmLogCleanup'
}

function Invoke-StoredProcedure
{
    param
    (
        [Parameter(Mandatory = $true)]
        [string]
        $Procedure,
        [Parameter(Mandatory = $false)]
        [object]
        $Parameters,
        [Parameter(Mandatory = $false)]
        [switch]
        $Scalar
    )
    process
    {
        try
        {
            $conn = New-Object -TypeName 'System.Data.SqlClient.SqlConnection'
            $conn.ConnectionString = $Script:Config.Logger.ConnectionString
            $conn.Open()
            $cmd = New-Object -TypeName 'System.Data.SqlClient.SqlCommand'
            $cmd.Connection = $conn
            $cmd.CommandText = $Procedure
            $cmd.CommandType = [System.Data.CommandType]::StoredProcedure
            if ($Parameters)
            {
                foreach ($key in $Parameters.Keys)
                {
                    [void]$cmd.Parameters.AddWithValue($key, $Parameters[$key])
                }
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
        finally {
            if ($conn)
            {
                $conn.Dispose()
            }
            if ($cmd)
            {
                $cmd.Dispose()
            }
        }
    }
}
