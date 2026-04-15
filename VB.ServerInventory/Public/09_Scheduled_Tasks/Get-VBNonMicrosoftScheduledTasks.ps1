# ============================================================
# FUNCTION : Get-VBNonMicrosoftScheduledTasks
# VERSION  : 1.0.2
# CHANGED  : 10-04-2026 -- Initial VB-compliant release
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Retrieve non-Microsoft scheduled tasks from target computer
# ENCODING : UTF-8 with BOM
# ============================================================

<#
.SYNOPSIS
    Retrieve non-Microsoft scheduled tasks from target computer(s).

.DESCRIPTION
    Queries scheduled tasks and filters out those created by Microsoft or related
    to OneDrive. Returns custom/third-party task details including name, state,
    author, path, run times, actions, and description.

.PARAMETER ComputerName
    Target computer(s). Defaults to local machine. Accepts pipeline input.
    Aliases: Name, Server, Host

.PARAMETER Credential
    Alternate credentials for remote execution.

.EXAMPLE
    Get-VBNonMicrosoftScheduledTasks
    Retrieves non-Microsoft tasks from local computer.

.EXAMPLE
    Get-VBNonMicrosoftScheduledTasks -ComputerName SERVER01
    Retrieves non-Microsoft tasks from SERVER01.

.EXAMPLE
    'SERVER01', 'SERVER02' | Get-VBNonMicrosoftScheduledTasks
    Retrieves non-Microsoft tasks from multiple computers via pipeline.

.OUTPUTS
    [PSCustomObject]: ComputerName, TaskName, State, Author, TaskPath, LastRunTime,
                     NextRunTime, Actions, Description, Status, CollectionTime

.NOTES
    Version  : 1.0.2
    Author   : Vibhu Bhatnagar
    Modified : 10-04-2026
    Category : Scheduled Tasks
#>

function Get-VBNonMicrosoftScheduledTasks {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('Name', 'Server', 'Host')]
        [string[]]$ComputerName = $env:COMPUTERNAME,
        [PSCredential]$Credential
    )
    process {
        foreach ($computer in $ComputerName) {
            try {
                # Step 1 -- Define remote script block
                $scriptBlock = {
                    # Step 1a -- Get all scheduled tasks
                    $tasks = Get-ScheduledTask | Where-Object {
                        ($_.TaskName -notmatch 'Microsoft') -and
                        ($_.TaskPath -notmatch 'Microsoft') -and
                        ($_.TaskName -notmatch 'OneDrive')
                    }

                    # Step 1b -- Process each task
                    foreach ($task in $tasks) {
                        $definition = $task.Definition

                        # Filter out Microsoft-authored tasks
                        if ($definition.Author -notmatch 'Microsoft') {
                            $taskInfo = Get-ScheduledTaskInfo -TaskName $task.TaskName -TaskPath $task.TaskPath -ErrorAction SilentlyContinue

                            [PSCustomObject]@{
                                TaskName = $task.TaskName
                                State = $task.State
                                Author = $definition.Author
                                TaskPath = $task.TaskPath
                                LastRunTime = if ($taskInfo.LastRunTime) { $taskInfo.LastRunTime } else { 'Never' }
                                NextRunTime = if ($taskInfo.NextRunTime) { $taskInfo.NextRunTime } else { 'N/A' }
                                Actions = ($definition.Actions | ForEach-Object { $_.Execute }) -join ', '
                                Description = $definition.Description
                            }
                        }
                    }
                }

                # Step 2 -- Execute locally or remotely
                if ($computer -eq $env:COMPUTERNAME) {
                    $results = & $scriptBlock
                } else {
                    $splat = @{
                        ComputerName = $computer
                        ScriptBlock = $scriptBlock
                    }
                    if ($Credential) {
                        $splat['Credential'] = $Credential
                    }
                    $results = Invoke-Command @splat
                }

                # Step 3 -- Output results with metadata
                if ($results) {
                    foreach ($item in $results) {
                        [PSCustomObject]@{
                            ComputerName = $computer
                            TaskName = $item.TaskName
                            State = $item.State
                            Author = $item.Author
                            TaskPath = $item.TaskPath
                            LastRunTime = $item.LastRunTime
                            NextRunTime = $item.NextRunTime
                            Actions = $item.Actions
                            Description = $item.Description
                            Status = 'Success'
                            CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                        }
                    }
                } else {
                    [PSCustomObject]@{
                        ComputerName = $computer
                        TaskName = 'None found'
                        State = 'N/A'
                        Author = 'N/A'
                        TaskPath = 'N/A'
                        LastRunTime = 'N/A'
                        NextRunTime = 'N/A'
                        Actions = 'N/A'
                        Description = 'No non-Microsoft scheduled tasks found'
                        Status = 'Success'
                        CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                    }
                }
            }
            catch {
                [PSCustomObject]@{
                    ComputerName = $computer
                    TaskName = 'N/A'
                    Status = 'Failed'
                    Error = $_.Exception.Message
                    CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                }
            }
        }
    }
}
