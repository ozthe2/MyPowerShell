Function New-OHScheduledTask {

    <#
.SYNOPSIS
    Creates a new scheduled task in PowerShell.

.DESCRIPTION
    The New-OHScheduledTask function allows you to create a new scheduled task in PowerShell. 
    It provides various parameters to customize the task, including the task name, description, trigger, 
    allowed user, script or program path, task delay, task folder, run with highest privilege, and 
    start task immediately options.

.PARAMETER TaskName
    Specifies the name of the scheduled task.

.PARAMETER TaskDescription
    Sets the description for the scheduled task.

.PARAMETER Trigger
    Determines when the task should be triggered. It accepts the following values:
    - AtLogon: Trigger the task when a user logs on.
    - AtStartup: Trigger the task when the system starts up.

.PARAMETER AllowedUser
    Specifies the user who is allowed to run the scheduled task. It accepts the following values:
    - BUILTIN\Users: Any user in the built-in Users group.
    - NT AUTHORITY\SYSTEM: The system account.

.PARAMETER ScriptPath
    Specifies the path to the PowerShell script file to be executed by the scheduled task.

.PARAMETER Program
    Specifies the path to the program or script file to be executed by the scheduled task.

.PARAMETER Arguments
    Specifies any arguments to be passed to the program or script.

.PARAMETER StartIn
    Sets the working directory for the program or script to be executed.

.PARAMETER DelayTask
    Adds a delay before the task is executed. It accepts the following values:
    - 30s: 30 seconds delay.
    - 1m: 1 minute delay.
    - 30m: 30 minutes delay.
    - 1h: 1 hour delay.

.PARAMETER TaskFolder
    Sets the folder where the scheduled task will be created. Default is "OHTesting".

.PARAMETER RunWithHighestPrivilege
    Specifies whether the task should run with the highest privileges.

.PARAMETER StartTaskImmediately
    Determines whether to start the task immediately after registering it.

.EXAMPLE
    New-OHScheduledTask -TaskName "MyTask" -TaskDescription "My task description" -Trigger "AtLogon" -AllowedUser "BUILTIN\Users" -ScriptPath "C:\Scripts\MyScript.ps1"

    Creates a new scheduled task named "MyTask" that triggers at user logon, allows any user in the built-in Users group to run the task, and executes the PowerShell script located at "C:\Scripts\MyScript.ps1".

.EXAMPLE
    New-OHScheduledTask -TaskName "ProgramTask" -TaskDescription "Program task description" -Trigger "AtStartup" -AllowedUser "NT AUTHORITY\SYSTEM" -Program "C:\Programs\MyProgram.exe" -Arguments "-param1 value1"

    Creates a new scheduled task named "ProgramTask" that triggers at system startup, allows the system account to run the task, and executes the program located at "C:\Programs\MyProgram.exe" with the specified arguments.

.NOTES
    Author: owen.heaume
    Date: 26-May-2023
    Version: 
        1.0 - Initial Script
        1.1 - Add -NoLogo -NonInteractive switches to $taskAction

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
    
        [Parameter(Mandatory = $false, ParameterSetName = "Program")]
        [Parameter(Mandatory = $false, ParameterSetName = "Script")]
        [ValidateSet('30s', '1m', '30m', '1h')]
        [string]$DelayTask,
    
        [Parameter(Mandatory = $false, ParameterSetName = "Program")]
        [Parameter(Mandatory = $false, ParameterSetName = "Script")]
        [string]$TaskFolder = "Specsavers",
    
        [Parameter(Mandatory = $false, ParameterSetName = "Program")]
        [Parameter(Mandatory = $false, ParameterSetName = "Script")]
        [switch]$RunWithHighestPrivilege,
    
        [Parameter(Mandatory = $false, ParameterSetName = "Program")]
        [Parameter(Mandatory = $false, ParameterSetName = "Script")]
        [switch]$StartTaskImmediately
    )

    if ($PSCmdlet.ParameterSetName -eq "Script") {
        if (!(Test-Path $ScriptPath)) {
            Write-Warning "Script path not found: $ScriptPath"
            return
        }
    
        # If a -TaskName was not used, then get the leaf name of the script path to use instead
        if ($TaskName -eq "") {
            $TaskName = $(Split-Path $ScriptPath -Leaf -Resolve).Replace('.ps1', "")
        }
    } elseif ($PSCmdlet.ParameterSetName -eq "Program") {
        if (!(Test-Path $Program)) {
            Write-Warning "Program path not found: $Program"
            return
        }
    
        # If a -TaskName was not used, then get the leaf name of the program path to use instead
        if ($TaskName -eq "") {
            $TaskName = $(Split-Path $Program -Leaf -Resolve)
        }
    }


    if (!(Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue)) {
        # Task action
        if ($PSCmdlet.ParameterSetName -eq "Script") {
            $taskAction = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument "-NoLogo -NonInteractive -File `"$ScriptPath`""
        } elseif ($PSCmdlet.ParameterSetName -eq "Program") {
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
        } else {
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
    
        # Verify the task has been registered and if so start immediately if the switch was 
        try {
            if ($(Get-ScheduledTask -TaskName $TaskName -ErrorAction stop -ev x).TaskName -eq $TaskName) {
                write-verbose "Task registered successfully!"
                # If the switch has been used, then start the task straight away
                if ($StartTaskImmediately) {
                    Write-Verbose "Starting the task immediately."
                    Start-ScheduledTask -TaskName $TaskName -TaskPath "\$taskFolder"
                }
            }
        } catch {
            Write-Warning "The task was not registered"
            $x
            exit
        } 
    } else {
        write-verbose "Scheduled Task $taskname already exists so taking no action"
    }   
}