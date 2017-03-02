# Make all error terminating errors
$ErrorActionPreference = 'Stop'

# Import-Module statements below are a workaround to make the script run with a gMSA
Import-Module -Name 'Microsoft.PowerShell.Utility'
Import-Module -Name 'Microsoft.PowerShell.Management'
Import-Module -Name 'Microsoft.PowerShell.Security'
Import-Module -Name 'ActiveDirectory'

. "$PSScriptRoot\Config.ps1"
. "$PSScriptRoot\TaskDefinition.ps1"
. "$PSScriptRoot\ExchangeOnline.ps1"
. "$PSScriptRoot\ExchangeOnprem.ps1"
. "$PSScriptRoot\MicrosoftOnline.ps1"
. "$PSScriptRoot\SkypeOnprem.ps1"
. "$PSScriptRoot\HomeFolder.ps1"
. "$PSScriptRoot\SamlId.ps1"
. "$PSScriptRoot\Logger.ps1"

Enum TaskResult
{
    Success
    Failure
    Retry
    Wait
}

$ExecutedInitializers = @()

function Start-Task
{
    param
    (
        [Parameter(Mandatory = $true)]
        [object]
        $Task,
        [Parameter(Mandatory = $true)]
        [object]
        $Target,
        [Parameter(Mandatory = $false)]
        [object]
        $SequenceTask = $null
    )
    process
    {
        $result = [TaskResult]::Failure
        if ($SequenceTask -eq $null)
        {
            $taskId = $Task.Id
        }
        else
        {
            $taskId = $SequenceTask.Id
        }
        # Do we know how to do this task?
        if (-not $Script:TaskDefinitions.ContainsKey($Task.TaskName))
        {
            Write-ErrorLogEntry -TaskName $Task.TaskName -TaskId $taskId -Object ([pscustomobject]@{
                Target    = $Target.UserPrincipalName
                Activity  = 'Start-Task'
                Reason    = 'Unknown task'
                Message   = 'This task is not defined in the dispatcher.'
            })
            [TaskResult]::Failure
            return
        }
        # Do we need to wait?
        if ($Task.WaitUntil -ne $null)
        {
            $waitUntil = Get-Date -Date $Task.WaitUntil
            if ((Get-Date) -lt $waitUntil)
            {
                if ($Task.TaskName -eq 'Wait')
                {
                    $result = [TaskResult]::Wait
                }
                else
                {
                    $result = [TaskResult]::Retry
                }
                $result
                return
            }
        }
        # Add any parameters as properties on the Target object if they don't exist.
        foreach ($paramName in $Script:TaskDefinitions[$Task.TaskName].Parameters)
        {
            if ($Target.$paramName -ne $null)
            {
                continue
            }
            if ($Task.$paramName -ne $null)
            {
                $paramValue = $Task.$paramName
            }
            elseif ($SequenceTask -ne $null -and $SequenceTask.$paramName -ne $null)
            {
                $paramValue = $SequenceTask.$paramName
            }
            else
            {
                Write-ErrorLogEntry -TaskName $Task.TaskName -TaskId $taskId -Object ([pscustomobject]@{
                    Target    = $Target.UserPrincipalName
                    Activity  = 'Start-Task'
                    Reason    = 'Parameter missing'
                    Message   = "Mandatory parameter '$paramName' is missing"
                })
                [TaskResult]::Failure
                return
            }
            $params = @{
                InputObject = $Target
                NotePropertyName = $paramName
                NotePropertyValue = $paramValue
            }
            Add-Member @params
        }
        # Add any optional parameters as properties on the Target object if they don't exist.
        foreach ($paramName in $Script:TaskDefinitions[$Task.TaskName].OptionalParameters)
        {
            if ($Target.$paramName -ne $null)
            {
                continue
            }
            if ($Task.$paramName -ne $null)
            {
                $paramValue = $Task.$paramName
            }
            elseif ($SequenceTask -ne $null -and $SequenceTask.$paramName -ne $null)
            {
                $paramValue = $SequenceTask.$paramName
            }
            else 
            {
                continue
            }
            $params = @{
                InputObject = $Target
                NotePropertyName = $paramName
                NotePropertyValue = $paramValue
            }
            Add-Member @params
        }
        if ($Task.TaskName -eq 'Wait')
        {
            if ($Task.WaitUntil -eq $null)
            {
                $params = @{
                    InputObject = $Task
                    NotePropertyName = 'WaitUntil'
                    NotePropertyValue = (Get-Date).AddMinutes($Task.Minutes).ToString('s')
                }
                Add-Member @params
                $result = [TaskResult]::Wait
            }
            else
            {
                $result = [TaskResult]::Success
            }
            $result
            return
        }
        try
        {
            # Execute initializer if needed
            $initializer = $Script:TaskDefinitions[$Task.TaskName].Initializer
            if ($initializer -ne $null -and $initializer -notin $Script:ExecutedInitializers)
            {
                &$initializer
                $Script:ExecutedInitializers += $initializer
            }
            # Execute task
            $Target | &$Script:TaskDefinitions[$Task.TaskName].Command
            $result = [TaskResult]::Success
        }
        catch
        {
            Write-ErrorLogEntry -TaskName $Task.TaskName -TaskId $taskId -Object $_.TargetObject
            if ($_.TargetObject.RetryCount -gt 0)
            {
                $params = @{
                    InputObject = $Task
                    NotePropertyName = 'WaitUntil'
                    NotePropertyValue = (Get-Date).AddMinutes($_.TargetObject.Delay).ToString('s')
                    Force = $true
                }
                Add-Member @params
                if ($Task.RetryCount -eq $null)
                {
                    $params = @{
                        InputObject = $Task
                        NotePropertyName = 'RetryCount'
                        NotePropertyValue = $_.TargetObject.RetryCount
                    }
                    Add-Member @params
                    $result = [TaskResult]::Retry
                }
                else
                {
                    if (--$Task.RetryCount -gt 0)
                    {
                        $result = [TaskResult]::Retry
                    }
                    else
                    {
                        Write-ErrorLogEntry -TaskName $Task.TaskName -TaskId $taskId -Object ([pscustomobject]@{
                            Target    = $Target.UserPrincipalName
                            Activity  = 'Start-Task'
                            Reason    = 'Task has failed completely'
                            Message   = 'Retry limit reached.'
                        })
                    }
                }
            }
            else
            {
                Write-ErrorLogEntry -TaskName $Task.TaskName -TaskId $taskId -Object ([pscustomobject]@{
                    Target    = $Target.UserPrincipalName
                    Activity  = 'Start-Task'
                    Reason    = 'Task has failed completely'
                    Message   = 'This task will not be retried.'
                })   
            }
        }
        $result
        return
    }
}

# ExtensionAttribute9 contains the task objects serialized as json
$params = @{
    Filter = {ExtensionAttribute9 -like '*' -and Enabled -eq $true}
    Properties = @(
        'Department'
        'DisplayName'
        'ExtensionAttribute9'
        'SamAccountName'
        'TelephoneNumber'
        'Title'
        'UserPrincipalName'
    )
}
try
{
    $users = Get-ADUser @params |
        ForEach-Object -Process {
            [pscustomobject]@{
                Identity = $_.ObjectGUID
                UserPrincipalName = $_.UserPrincipalName
                SamAccountName = $_.SamAccountName
                ExtensionAttribute9 = $_.ExtensionAttribute9
                DisplayName = $_.DisplayName
                Department = $_.Department
                Title = $_.Title
                TelephoneNumber = $_.TelephoneNumber
            }
        }
}
catch
{
    Write-ErrorLogEntry -TaskName 'GetUsers' -TaskId -1 -Object ([pscustomobject]@{
        Target    = '(&(extensionAttribute9=*)(!userAccountControl:1.2.840.113556.1.4.803:=2))'
        Activity  = 'Get-ADUser'
        Reason    = 'Failed to get users from Active Directory'
        Message   = $_.Exception.Message
    })
    exit
}
foreach ($user in $users)
{
    try
    {
        $deserializedTasks = ConvertFrom-Json -InputObject $user.ExtensionAttribute9
    }
    catch
    {
        Write-ErrorLogEntry -TaskName 'ProcessTasks' -TaskId -1 -Object ([pscustomobject]@{
            Target    = $user.UserPrincipalName
            Activity  = 'ConvertFrom-Json'
            Reason    = 'Failed to deserialize the contents of ExtensionAttribute9'
            Message   = $_.Exception.Message
        })
        continue
    }
    $remainingTasks = @()
    for ($i = 0; $i -lt $deserializedTasks.Count; $i++)
    {
        $currentTask = $deserializedTasks[$i]
        $isSequenceTask = ($currentTask.Tasks -ne $null -and $currentTask.Tasks.Count -gt 0)
        if ($currentTask.Id -eq $null)
        {
            $taskId = New-TaskLogEntry -Task $currentTask.TaskName -Target $user.UserPrincipalName
            $params = @{
                InputObject = $currentTask
                NotePropertyName = 'Id'
                NotePropertyValue = $taskId
            }
            Add-Member @params
        }
        # Sequence Task
        if ($isSequenceTask)
        {
            while ($currentTask.Tasks.Count -gt 0)
            {
                $result = Start-Task -Task $currentTask.Tasks[0] -Target $user -SequenceTask $currentTask
                if ($result -eq [TaskResult]::Success)
                {
                    if ($currentTask.Tasks.Count -gt 1)
                    {
                        Update-TaskLogEntry -TaskId $currentTask.Id -Result $result
                        $currentTask.Tasks = $currentTask.Tasks[1..($currentTask.Tasks.Count - 1)]
                    }
                    else
                    {
                        Update-TaskLogEntry -TaskId $currentTask.Id -Result $result -EndTask
                        $currentTask.Tasks = @()
                    }
                }
                elseif ($result -eq [TaskResult]::Retry -or $result -eq [TaskResult]::Wait)
                {
                    $remainingTasks += $currentTask
                    Update-TaskLogEntry -TaskId $currentTask.Id -Result $result
                    break
                }
                else # [TaskResult]::Failure
                {
                    Update-TaskLogEntry -TaskId $currentTask.Id -Result $result -EndTask
                    break
                }
            }
        }
        # Single task
        else
        {
            $result = Start-Task -Task $currentTask -Target $user
            if ($result -eq [TaskResult]::Retry)
            {
                $remainingTasks += $currentTask
                Update-TaskLogEntry -TaskId $currentTask.Id -Result $result
            }
            elseif ($result -eq [TaskResult]::Wait)
            {
                $remainingTasks += $deserializedTasks[$i..($deserializedTasks.Count - 1)]
                Update-TaskLogEntry -TaskId $currentTask.Id -Result $result
                break
            }
            else
            {
                Update-TaskLogEntry -TaskId $currentTask.Id -Result $result -EndTask
            }
        }
    }
    # If we have tasks left, save them back to ExtensionAttribute9. If this attribute was cleared
    # while the tasks were executed, we assume someone don't want to execute the remaining tasks.
    try
    {
        if ($remainingTasks.Count -gt 0)
        {
            $json = ConvertTo-Json -InputObject $remainingTasks -Depth 3 -Compress
            $params = @{
                Identity = $user.Identity
                Properties = @('ExtensionAttribute9')
                ErrorAction = 'SilentlyContinue'
            }
            if ($null -ne (Get-ADUser @params).ExtensionAttribute9)
            {
                Set-ADUser -Identity $user.Identity -Replace @{ExtensionAttribute9=$json}
            }
        }
        else
        {
            Set-ADUser -Identity $user.Identity -Clear 'ExtensionAttribute9'
        }
    }
    catch
    {
        Write-ErrorLogEntry -TaskName 'ProcessTasks' -TaskId -1 -Object ([pscustomobject]@{
            Target    = $user.UserPrincipalName
            Activity  = 'Save remaining tasks'
            Reason    = ''
            Message   = $_.Exception.Message
        })
    }
}
