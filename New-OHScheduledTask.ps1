Function New-OHScheduledTask {
    <#
    .SYNOPSIS
    Creates a new scheduled task on the local computer.
    
    .DESCRIPTION
    The New-OHScheduledTask function creates a new scheduled task on the local computer. You can create a task that runs a PowerShell script or an executable program with arguments.
    
    .PARAMETER TaskName
    Specifies the name of the scheduled task.
    
    .PARAMETER TaskDescription
    Specifies a description of the scheduled task. This parameter is mandatory."
    
    .PARAMETER Trigger
    Specifies the trigger for the scheduled task. The accepted values are "AtLogon" and "AtStartup".
    
    .PARAMETER AllowedUser
    Specifies the user account that is allowed to run the scheduled task. The accepted values are "BUILTIN\Users" and "NT AUTHORITY\SYSTEM".
    
    .PARAMETER ScriptPath
    Specifies the path to the PowerShell script that the scheduled task runs. This parameter is mandatory if the parameter set name is "Script".
    
    .PARAMETER Program
    Specifies the path to the program that the scheduled task runs. This parameter is mandatory if the parameter set name is "Program".
    
    .PARAMETER Arguments
    Specifies the arguments that the program uses. This parameter is optional.
    
    .PARAMETER StartIn
    Specifies the starting directory for the program. This parameter is optional.
    
    .PARAMETER Action
    Specifies the action for the scheduled task. The accepted values are "add" and "replace".
    
    .PARAMETER DelayTask
    Specifies the delay time for the scheduled task. The accepted values are "30s", "1m", "30m", and "1h". This parameter is optional.
    
    .PARAMETER TaskFolder
    Specifies the folder for the scheduled task. The default value is "OHTesting".
    
    .PARAMETER RunWithHighestPrivilege
    Runs the scheduled task with the highest privileges. This parameter is optional.
    
    .PARAMETER StartTaskImmediately
    Starts the scheduled task immediately. This parameter is optional.
    
    .PARAMETER Delete
    Deletes the scheduled task. This parameter is optional.
    
    .EXAMPLE
    New-OHScheduledTask -TaskName "MyTask" -TaskDescription "Runs a PowerShell script with lots of arguments." -Trigger AtStartup -AllowedUser 'NT AUTHORITY\SYSTEM' -Program "C:\Windows\System32\WindowsPowerShell\v1.0\PowerShell.exe" -Arguments '-noprofile -executionpolicy bypass -command "& { . c:\ohtemp\test2.ps1; Show-Text -textToDisplay "Important Text!" }' -Action add -RunWithHighestPrivilege
    
    This example creates a new scheduled task named "MyTask" that runs a PowerShell script with lots of arguments at system startup. The task is set to add a new task as long as the task name does not already exist, and it will run with the highest privileges. The PowerShell script in this example displays the text "Important Text!" using the Show-Text function contained in the script.
    
    .EXAMPLE
    New-OHScheduledTask -TaskName "MyTask" -TaskDescription "Runs a simple PowerShell script." -Trigger AtStartup -AllowedUser 'NT AUTHORITY\SYSTEM' -ScriptPath "c:\ohtemp\test2.ps1" -Action replace -RunWithHighestPrivilege
    
    This example creates a new scheduled task named "MyTask" that runs a simple PowerShell script at system startup. The task is set to replace any existing task with the same name, and it will run with the highest privileges.
    
    .EXAMPLE
    New-OHScheduledTask -TaskName "MyTask" -Delete

    This example deletes the scheduled task named, "MyTask" if it exists.
    
    .NOTES
    Author: owen.heaume
    Date: 14-May-2023
    Version: 1.6
    #>
    
    [CmdletBinding(SupportsShouldProcess = $true, DefaultParameterSetName = "Script")]
    param(
        [Parameter()]
        [string]$TaskName,
    
        [Parameter(Mandatory = $true, ParameterSetName = "Program")]
        [Parameter(Mandatory = $true, ParameterSetName = "Script")]
        [string]$TaskDescription,
    
        [Parameter(Mandatory = $true, ParameterSetName = "Program")]
        [Parameter(Mandatory = $true, ParameterSetName = "Script")]
        [ValidateSet('AtLogon', 'AtStartup')]
        [string]$Trigger = "AtStartup",
    
        [Parameter(Mandatory = $true, ParameterSetName = "Program")]
        [Parameter(Mandatory = $true, ParameterSetName = "Script")]
        [ValidateSet('BUILTIN\Users', 'NT AUTHORITY\SYSTEM')]
        [string]$AllowedUser,
    
        [Parameter(Mandatory = $true, ParameterSetName = "Script")]
        [ValidateScript({ $_ -match '\.ps1$' })]
        [string]$ScriptPath,
    
        [Parameter(Mandatory = $true, ParameterSetName = "Program")]
        [ValidateScript({ $_ -match '\.(exe|bat|cmd)$' })]      
        [string]$Program,
    
        [Parameter(Mandatory = $false, ParameterSetName = "Program")]        
        [string]$Arguments,
    
        [Parameter(Mandatory = $false, ParameterSetName = "Program")]        
        [string]$StartIn,
    
        [Parameter(Mandatory = $true, ParameterSetName = "Program")]
        [Parameter(Mandatory = $true, ParameterSetName = "Script")]
        [ValidateSet('add', 'replace')]
        [string]$Action,
    
        [Parameter(Mandatory = $false, ParameterSetName = "Program")]
        [Parameter(Mandatory = $false, ParameterSetName = "Script")]
        [ValidateSet('30s', '1m', '30m', '1h')]
        [string]$DelayTask,
    
        [Parameter(Mandatory = $false, ParameterSetName = "Program")]
        [Parameter(Mandatory = $false, ParameterSetName = "Script")]
        [string]$TaskFolder = "OHTesting",
    
        [Parameter(Mandatory = $false, ParameterSetName = "Program")]
        [Parameter(Mandatory = $false, ParameterSetName = "Script")]
        [switch]$RunWithHighestPrivilege,
    
        [Parameter(Mandatory = $false, ParameterSetName = "Program")]
        [Parameter(Mandatory = $false, ParameterSetName = "Script")]
        [switch]$StartTaskImmediately,
    
        [Parameter(ParameterSetName = "Delete")]
        [switch]$Delete
    )    
        
    if ($PSCmdlet.ParameterSetName -eq "Script") {
        if (!(Test-Path $ScriptPath)) {
            Write-host "Script path not found: $ScriptPath" -ForegroundColor Yellow
            return
        }
    
        # If a -TaskName was not used, then get the leaf name of the script path to use instead
        if ($TaskName -eq "") {
            $TaskName = $(Split-Path $ScriptPath -Leaf -Resolve).Replace('.ps1', "")
        }
    }
    elseif ($PSCmdlet.ParameterSetName -eq "Program") {
        if (!(Test-Path $Program)) {
            Write-host "Program path not found: $Program" -ForegroundColor Yellow
            return
        }
    
        # If a -TaskName was not used, then get the leaf name of the program path to use instead
        if ($TaskName -eq "") {
            $TaskName = $(Split-Path $Program -Leaf -Resolve)
        }
    }
    
    # If the action is "delete" or "replace," then delete the task if it already exists (by task name)
    if ($delete -or $Action -eq "Replace") {
        Try {
            if (Get-ScheduledTask -TaskName $TaskName -ErrorAction Stop) {
                Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
                Write-Host "Removed Scheduled Task $taskname" -ForegroundColor Green
            }
        }
        Catch {
            Write-Host "$TaskName not found. No task to delete" -ForegroundColor Cyan        
        }
    }
    
    # If the action is "add" or "replace," then check if the task already exists
    if ($Action -eq "Add" -or $Action -eq "Replace") {
        if (!(Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue)) {
            # Task action
            if ($PSCmdlet.ParameterSetName -eq "Script") {
                $taskAction = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument "-File `"$ScriptPath`""
            }
            elseif ($PSCmdlet.ParameterSetName -eq "Program") {
                $taskAction = New-ScheduledTaskAction -Execute $Program
                if ($Arguments -ne $null -and $Arguments -ne '') {
                    #$taskAction.Argument = $Arguments  <<<--- For some reason, this doesn't work
                    $taskAction = New-ScheduledTaskAction -Execute $Program -Argument $Arguments
                }
                if ($StartIn -ne $null -and $StartIn -ne '') {
                    $taskAction.WorkingDirectory = $StartIn  #<<<--- Yet this does!
                }
            }
    
            # Task Trigger
            switch ($trigger) {
                'AtStartup' { $taskTrigger = New-ScheduledTaskTrigger -AtStartup }
                'AtLogon' { $taskTrigger = New-ScheduledTaskTrigger -AtLogOn }
            }
    
            # Task principal
            $taskPrincipal = New-ScheduledTaskPrincipal -GroupId $AllowedUser
    
            # task settings
            $taskSettings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -WakeToRun
            if ($RunWithHighestPrivilege) {
                $taskSettings.ExecutionTimeLimit = 'PT0S'
                $taskSettings.RunOnlyIfNetworkAvailable = $false
                $taskSettings.Hidden = $false
                $taskSettings.Priority = 7
            }
    
            # Add a task delay if selected
            if ($DelayTask) {
                $taskDelay = New-TimeSpan -Seconds 1
                switch ($DelayTask) {
                    '30s' { $taskDelay = New-TimeSpan -Seconds 30 }
                    '1m' { $taskDelay = New-TimeSpan -Minutes 1 }
                    '30m' { $taskDelay = New-TimeSpan -Minutes 30 }
                    '1h' { $taskDelay = New-TimeSpan -Hours 1 }
                }
                $delayTime = "PT" + $taskDelay.ToString('hh') + "H" + $taskDelay.ToString('mm') + "M" + $taskDelay.ToString('ss') + "S"
                $taskTrigger.Delay = $delayTime
            }
    
            # Run with highest privileges if selected
            if ($RunWithHighestPrivilege) {
                $taskPrincipal.RunLevel = "Highest"
            }
            else {
                $taskPrincipal.RunLevel = "Limited"
            }
    
            # Register the new PowerShell scheduled task
            Register-ScheduledTask `
                -TaskName $TaskName `
                -Action $taskAction `
                -Trigger $taskTrigger `
                -Principal $taskPrincipal `
                -Description $TaskDescription `
                -TaskPath "\$taskFolder"
    
            # Verify the task has been registered
            try {
                if ($(Get-ScheduledTask -TaskName $TaskName -ErrorAction stop -ev x).TaskName -eq $TaskName) {
                    write-host "Task registered successfully!" -ForegroundColor Green
                    # If the switch has been used, then start the task straight away
                    if ($StartTaskImmediately) {
                        Write-host "Starting the task immediately." -ForegroundColor Cyan
                        Start-ScheduledTask -TaskName $TaskName -TaskPath "\$taskFolder"
                    }
                }
            }
            catch {
                Write-Warning "The task was not registered"
                $x
                exit
            } 
        }
        else {
            Write-Host "Scheduled Task $taskname already exists so taking no action" -ForegroundColor Yellow
        }
    }   
}    