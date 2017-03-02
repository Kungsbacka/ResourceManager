# Make all error terminating errors
$Global:ErrorActionPreference = 'Stop'
. "$PSScriptRoot\Config.ps1"

Enum TaskResult
{
    Success
    Failure
    Retry
    Wait
}

function Update-TaskLogEntry
{
    param
    (
        [Parameter(Mandatory = $true)]
        [int]
        $TaskId,
        [Parameter(Mandatory = $true)]
        [TaskResult]
        $Result,
        [Parameter(Mandatory = $false)]
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
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true,ParameterSetName='CustomError')]
        [string]
        $Target,
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string]
        $TaskJson
    )
    if ($Message)
    {
@"
---------------------------------------------------
---- {0}
---------------------------------------------------
##Target: {1}

##Exception:
{2}
"@ -f (Get-Date),

    }


    if ($ErrorRecord.Exception.Data.Contains('Parameters'))
    {
        $params = $ErrorRecord.Exception.Data.Parameters
        $paramsObject = $params | ConvertFrom-Json
        if ($paramsObject.UserPrincipalName)
        {
            $target = $paramsObject.UserPrincipalName
        }
        elseif ($paramsObject.Identity)
        {
            $target = $paramsObject.Identity
        }
        else
        {
            $target = 'Unknown'
        }
    }
    else
    {
        $params = 'No parameters'
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

##Parameters:
{4}

"@ -f  (Get-Date),
        $target,
        $_.Exception.ToString(),
        $_.ScriptStackTrace,
        $params | Out-File -FilePath $Script:Config.Logger.LogPath -Encoding UTF8 -Append
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

function New-Exception
{
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Message,
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [object]$Parameters
    )

    $exception = New-Object -TypeName 'System.Exception' -ArgumentList @($Message)
    if ($Parameters)
    {
        $serializedParams = $Parameters | ConvertTo-NewtonsoftJson -Formatting Indented
        $exception.Data.Add('Parameters', $serializedParams)
    }
    Write-Output -InputObject $exception
}
