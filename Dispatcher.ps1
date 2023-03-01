# Make all error terminating errors
$ErrorActionPreference = 'Stop'

Import-Module -Name 'ActiveDirectory'

. "$PSScriptRoot\Config.ps1"
. "$PSScriptRoot\TaskDefinition.ps1"
. "$PSScriptRoot\Exchange.ps1"
. "$PSScriptRoot\ExchangeOnline.ps1"
. "$PSScriptRoot\ExchangeOnprem.ps1"
. "$PSScriptRoot\MicrosoftOnline.ps1"
. "$PSScriptRoot\OnpremAd.ps1"
. "$PSScriptRoot\SamlId.ps1"
. "$PSScriptRoot\Logger.ps1"

Enum TaskResult
{
    Success
    Failure
    Wait
}

$ExecutedInitializers = @()

function Start-Task
{
    param
    (
        [Parameter(Mandatory=$true)]
        [object]
        $Task,
        [Parameter(Mandatory=$true)]
        [object]
        $Target,
        [object]
        $SequenceTask = $null
    )
    process
    {
        # Do we know how to do this task?
        if (-not $Script:TaskDefinitions.ContainsKey($Task.TaskName))
        {
            throw "Unknown task: $($Task.Name)"
        }
        # Do we need to wait?
        if ($null -ne $Task.WaitUntil)
        {
            $waitUntil = Get-Date -Date $Task.WaitUntil
            if ((Get-Date) -lt $waitUntil)
            {
                [TaskResult]::Wait
                return
            }
        }
        # Add any parameters as properties on the Target object if they don't exist.
        foreach ($paramName in $Script:TaskDefinitions[$Task.TaskName].Parameters)
        {
            if ($null -ne $Target.$paramName)
            {
                continue
            }
            if ($null -ne $Task.$paramName)
            {
                $paramValue = $Task.$paramName
            }
            elseif ($null -ne $SequenceTask -and $null -ne $SequenceTask.$paramName)
            {
                $paramValue = $SequenceTask.$paramName
            }
            else
            {
                throw "Mandatory parameter missing: $paramName"
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
            if ($null -ne $Target.$paramName)
            {
                continue
            }
            if ($null -ne $Task.$paramName)
            {
                $paramValue = $Task.$paramName
            }
            elseif ($null -ne $SequenceTask -and $null -ne $SequenceTask.$paramName)
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
            if ($null -eq $Task.WaitUntil)
            {
                $params = @{
                    InputObject = $Task
                    NotePropertyName = 'WaitUntil'
                    NotePropertyValue = (Get-Date).AddMinutes($Task.Minutes).ToString('s')
                }
                Add-Member @params
                [TaskResult]::Wait
            }
            else
            {
                [TaskResult]::Success
            }
            return
        }
        # Execute initializer if needed
        $initializer = $Script:TaskDefinitions[$Task.TaskName].Initializer
        if ($null -ne $initializer -and $initializer -notin $Script:ExecutedInitializers)
        {
            &$initializer
            $Script:ExecutedInitializers += $initializer
        }
        # Execute task
        $Target | &$Script:TaskDefinitions[$Task.TaskName].Command
        [TaskResult]::Success
        return
    }
}

# CarLicense contains the task objects serialized as json
$params = @{
    Filter = "CarLicense -like '*' -and (Enabled -eq 'True' -or (Enabled -eq 'False' -and AccountExpirationDate -le '$([DateTime]::Now.ToString('yyyy-MM-dd HH:mm:ss'))'))"
    Properties = @(
        'Department'
        'DisplayName'
        'CarLicense'
        'SamAccountName'
        'TelephoneNumber'
        'Title'
        'UserPrincipalName'
    )
    ResultSetSize = 200 # Only do 200 before exiting to avoid problems with Exchange Online
}
try
{
    $users = Get-ADUser @params |
        ForEach-Object -Process {
            [pscustomobject]@{
                Identity = $_.ObjectGuid
                UserPrincipalName = $_.UserPrincipalName
                SamAccountName = $_.SamAccountName
                CarLicense = $_.CarLicense
                DisplayName = $_.DisplayName
                Department = $_.Department
                Title = $_.Title
                TelephoneNumber = $_.TelephoneNumber
            }
        }
}
catch
{
    New-TaskLogEntry -Task 'QueryActiveDirectory' -Result ([TaskResult]::Failure)
    Write-ErrorLog -ErrorRecord $_
    exit
}
foreach ($user in $users)
{
    try
    {
        $user.CarLicense = $user.CarLicense[0] # Multivalued attribute
        $deserializedTasks = ConvertFrom-Json -InputObject $user.CarLicense
    }
    catch
    {
        New-TaskLogEntry -Task 'DeserializeTaskJson' -Target $user.UserPrincipalName -Result ([TaskResult]::Failure)
        Write-ErrorLog -ErrorRecord $_ -Target $user.UserPrincipalName -TaskJson $user.CarLicense
        continue
    }
    $remainingTasks = @()
    for ($i = 0; $i -lt $deserializedTasks.Count; $i++)
    {
        $currentTask = $deserializedTasks[$i]
        if (-not $currentTask.TaskName)
        {
            try
            {
                throw 'TaskName is empty or missing'
            }
            catch
            {
                New-TaskLogEntry -Task 'ValidateDeserializedTaskJson' -Target $user.UserPrincipalName -Result ([TaskResult]::Failure)
                Write-ErrorLog -ErrorRecord $_ -Target $user.UserPrincipalName -TaskJson $user.CarLicense
                exit
            }
        }
        $isSequenceTask = ($null -ne $currentTask.Tasks -and $currentTask.Tasks.Count -gt 0)
        if ($null -eq $currentTask.Id)
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
                try
                {
                    $result = Start-Task -Task $currentTask.Tasks[0] -Target $user -SequenceTask $currentTask
                }
                catch
                {
                    Write-ErrorLog -ErrorRecord $_ -TaskId $currentTask.Id -Target $user.UserPrincipalName -TaskJson $user.CarLicense
                    Update-TaskLogEntry -TaskId $currentTask.Id -Result ([TaskResult]::Failure) -EndTask -Details $_.Exception.ToString()
                    break
                }
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
                else # [TaskResult]::Wait
                {
                    $remainingTasks += $currentTask
                    Update-TaskLogEntry -TaskId $currentTask.Id -Result $result
                    break
                }
            }
        }
        # Single task
        else
        {
            try
            {
                $result = Start-Task -Task $currentTask -Target $user
            }
            catch
            {
                Write-ErrorLog -ErrorRecord $_ -TaskId $currentTask.Id -Target $user.UserPrincipalName -TaskJson $user.CarLicense
                Update-TaskLogEntry -TaskId $currentTask.Id -Result ([TaskResult]::Failure) -EndTask -Details $_.Exception.ToString()
                continue
            }
            if ($result -eq [TaskResult]::Wait)
            {
                $remainingTasks += $deserializedTasks[$i..($deserializedTasks.Count - 1)]
                Update-TaskLogEntry -TaskId $currentTask.Id -Result $result
            }
            else # [TaskResult]::Success
            {
                Update-TaskLogEntry -TaskId $currentTask.Id -Result $result -EndTask
            }
        }
    }
    # If we have tasks left, save them back to CarLicense. If this attribute was cleared
    # while the tasks were executed, we assume someone didn't want to execute the remaining tasks.
    try
    {
        $user2 = Get-ADUser -Filter "ObjectGuid -eq '$($user.Identity)' -and CarLicense -like '*'" -Properties 'CarLicense'
        if ($null -eq $user2)
        {
            continue
        }
        if ($remainingTasks.Count -gt 0)
        {
            if ($null -ne $user2.CarLicense)
            {
                $json = ConvertTo-Json -InputObject $remainingTasks -Depth 4 -Compress
                Set-ADUser -Identity $user.Identity -Replace @{'CarLicense'=$json}
            }
        }
        else
        {
            Set-ADUser -Identity $user.Identity -Clear 'CarLicense'
        }
    }
    catch
    {
        New-TaskLogEntry -Task 'SaveRemaningTasks' -Target $user.UserPrincipalName -Result ([TaskResult]::Failure)
        Write-ErrorLog -ErrorRecord $_ -Target $user.UserPrincipalName
    }
}

# Clean up stale log entries
Close-UnclosedFailedTasks
