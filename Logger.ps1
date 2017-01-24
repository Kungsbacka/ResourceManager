# Make all error terminating errors
$Global:ErrorActionPreference = 'Stop'
. "$PSScriptRoot\Config.ps1"

function Update-TaskLogEntry
{
    param
    (
        [Parameter(Mandatory = $true)]
        [int]
        $TaskId,
        [Parameter(Mandatory = $true)]
        [AccountTasks.TaskResult]
        $Result,
        [Parameter(Mandatory = $false)]
        [switch]
        $EndTask
    )
    process
    {
        Invoke-StoredProcedure -Procedure 'dbo.spUpdateTaskLogEntry' -Parameters @{
            TaskId = $TaskId
            Status = $Result.ToString()
            EndTask = [int]($EndTask.ToBool())
        }
    }
}

function New-TaskLogEntry
{
    param
    (
        [Parameter(Mandatory = $true)]
        [string]
        $Task,
        [Parameter(Mandatory = $true)]
        [string]
        $Target
    )
    process
    {
        Invoke-StoredProcedure -Procedure 'dbo.spInsertNewTaskLogEntry' -Scalar -Parameters @{
            Task = $Task
            Target = $Target
        }
    }
}

function Write-ErrorLogEntry
{
    param
    (
        [Parameter(Mandatory = $true)]
        [int]
        $TaskId,
        [Parameter(Mandatory = $true)]
        [string]
        $TaskName,
        [Parameter(Mandatory = $true)]
        [System.Object]
        $Object
    )
    process
    {
        Invoke-StoredProcedure -Procedure 'dbo.spInsertNewTaskErrorLogEntry' -Parameters @{
            TaskId = $TaskId
            Task = $TaskName
            Target = $Object.Target
            Activity = $Object.Activity
            Reason = $Object.Reason
            Message = $Object.Message
        }
  
    }
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
