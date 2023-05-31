BeforeAll {
    . "..\New-OHScheduledTask.ps1"
}

Describe "New-OHScheduledTask" {    

    BeforeEach {
        # Create a temporary script file
        $scriptContent = @"
        Write-Host 'Hello, World!'
"@
        $scriptPath = [System.IO.Path]::GetTempFileName() + ".ps1"
        $scriptContent | Set-Content -Path $scriptPath -Force

        $scriptArguments = "-File `"$(Resolve-Path $scriptPath)`""

        # Create a temporary program file
        $programPath = [System.IO.Path]::GetTempFileName() + ".exe"
        $programContent = @"
        echo Hello, World!
"@
        $programContent | Set-Content -Path $programPath -Force

        # Set the program file path in a variable for later use
        $script:programPath = $programPath
        $program = $script:programPath
        $Arguments = "-Param1 Value1 -Param2 Value2"
        $taskDescription = "MyTask Description"
        $trigger = "AtLogon"
        $allowedUser = "BUILTIN\Users"
        $DelayTask = "1h"
        $taskName = "MyTask"
    }
    
    AfterEach {
        # Clean up the scheduled task
        if ($taskName) {
            Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue
        }

        # Remove the temporary program file
        if ($script:programPath) {
            Remove-Item -Path $script:programPath -Force
        }
    } 

    Context "When called with a -TaskName parameter" {
        It "Should create a scheduled task with the specified name" {
            New-OHScheduledTask -TaskName $taskName -TaskDescription $taskDescription -Trigger $trigger -AllowedUser $allowedUser -ScriptPath $scriptPath

            $task = Get-ScheduledTask -TaskName $taskName
            $task.TaskName | Should -Be $taskName

            # Remove the temporary script file
            Remove-Item -Path $scriptPath -Force
        }
    }

    Context "When called with a -ScriptPath parameter" -Tag 'ScriptTest' { # passing
        It "Should create a scheduled task that runs the specified script" {
            New-OHScheduledTask -TaskName $taskName -TaskDescription $taskDescription -Trigger $trigger -AllowedUser $allowedUser -ScriptPath $scriptPath

            $task = Get-ScheduledTask -TaskName $taskName
            $task.actions.Arguments | Should -Be "-NoLogo -NonInteractive $scriptArguments"

            # Remove the temporary script file
            Remove-Item -Path $scriptPath -Force
        }
    }

    Context "When called with a -Program parameter" {
        It "Should create a scheduled task that runs the specified program" {
            New-OHScheduledTask -TaskName $taskName -TaskDescription $taskDescription -Trigger $trigger -AllowedUser $allowedUser -Program $program
            
            $task = Get-ScheduledTask -TaskName $taskName
            
            $programPath = $task.Actions.execute
            $programPath | Should -Be $program
        }
    }

    Context "When called with a -TaskDescription parameter" {
        It "Should create a scheduled task with the specified task description" {        

            New-OHScheduledTask -TaskName $taskName -TaskDescription $taskDescription -Trigger $trigger -AllowedUser $allowedUser -ScriptPath $scriptPath

            $task = Get-ScheduledTask -TaskName $taskName

            # Verify the task description
            $task.Description | Should -Be $taskDescription
        }
    }

    Context "When called with -AllowedUser 'BUILTIN\Users'" {
        It "Should create a scheduled task with the allowed user" {

            New-OHScheduledTask -TaskName $taskName -TaskDescription $taskDescription -Trigger $trigger -AllowedUser $allowedUser -ScriptPath $scriptPath

            $task = Get-ScheduledTask -TaskName $taskName

            # Verify the allowed user
            $task.Principal.GroupId | Should -Be 'Users'
        }
    }

    Context "When called with -AllowedUser 'NT Authority\SYSTEM'" {
        It "Should create a scheduled task with the allowed user" {
            $alloweduser = 'NT AUTHORITY\SYSTEM'

            New-OHScheduledTask -TaskName $taskName -TaskDescription $taskDescription -Trigger $trigger -AllowedUser $allowedUser -ScriptPath $scriptPath

            $task = Get-ScheduledTask -TaskName $taskName

            # Verify the allowed user
            $task.Principal.userid | Should -Be 'SYSTEM'
        }
    }

    Context "When called with a -arguments parameter" {
        It "Should create a scheduled task with the specified arguments" {
            
            New-OHScheduledTask -TaskName $taskName -TaskDescription $taskDescription -Trigger $trigger -AllowedUser $allowedUser -program $program -Arguments $Arguments

            $task = Get-ScheduledTask -TaskName $taskName

            # Verify the arguments
            $task.Actions.Arguments | Should -Be $Arguments
        }
    }

    Context "When called with a -DelayTask parameter" {
        It "Should create a scheduled task with the specified delay" {
            New-OHScheduledTask -TaskName $taskName -TaskDescription $taskDescription -Trigger $trigger -AllowedUser $allowedUser -ScriptPath $scriptPath -DelayTask $DelayTask

            $task = Get-ScheduledTask -TaskName $taskName

            # Verify the task delay
            $task.Triggers.delay | Should -Be 'PT1H'
        }
    }

    Context "When called with the -RunWithHighestPrivilege switch" {
        It "Should create a scheduled task with the 'Run with highest privileges' option enabled" {
            #$RunWithHighestPrivilege = $true

            New-OHScheduledTask -TaskName $taskName -TaskDescription $taskDescription -Trigger $trigger -AllowedUser $allowedUser -ScriptPath $scriptPath -RunWithHighestPrivilege

            $task = Get-ScheduledTask -TaskName $taskName

            # Verify that the RunLevel is set to Highest
            $task.Principal.RunLevel | Should -Be 'Highest'
        }
    }   

    Context "When not called with the -RunWithHighestPrivilege switch" {
        It "Should create a scheduled task without 'Run with highest privileges' option enabled" {            

            New-OHScheduledTask -TaskName $taskName -TaskDescription $taskDescription -Trigger $trigger -AllowedUser $allowedUser -ScriptPath $scriptPath

            $task = Get-ScheduledTask -TaskName $taskName

            # Verify that the RunLevel is set to Highest
            $task.Principal.RunLevel | Should -Not -Be 'Highest'
        }
    }  
 
    Context "When specifying a valid StartIn path" -tag 'work' {
        It "Should create a scheduled task with the specified StartIn path" {
               
            $startIn = "C:\Scripts"
              
            New-OHScheduledTask -TaskName $taskName -TaskDescription $taskDescription -Trigger $trigger -AllowedUser $allowedUser -Program $program -StartIn $startIn -TaskFolder $taskFolder
                    
            $task = Get-ScheduledTask -TaskName $taskName
                    
            $task.actions.workingdirectory | Should -Be $startIn
        }
    }

    Context "When not specifying a StartIn path" {
        It "Should create a scheduled task with no StartIn path" { 
               
            New-OHScheduledTask -TaskName $taskName -TaskDescription $taskDescription -Trigger $trigger -AllowedUser $allowedUser -Program $program -TaskFolder $taskFolder
    
            $task = Get-ScheduledTask -TaskName $taskName
                                    
            $task.actions.workingdirectory | Should -BeNullOrEmpty
        }
    }

    Context "When using the 'Script' parameter set" {
        It "Should set the TaskName based on the script path" {          
            
            $filename = Split-Path $scriptpath -Leaf
            $filenameWithoutExtension = $filename.Replace(".ps1", "")
            $expectedTaskName = $filenameWithoutExtension
           
            New-OHScheduledTask -TaskDescription $taskDescription -Trigger $trigger -AllowedUser $allowedUser -ScriptPath $scriptPath

           
            $task = Get-ScheduledTask -TaskName $expectedtaskName
           
            $task.taskName | Should -Be $expectedTaskName
        }
    }

    Context "When using the -trigger 'AtStartup' parameter" -tag "test" {
        It "Should set the trigger based on the supplied value" {    
                       
            $trigger = 'AtStartup'
           
            New-OHScheduledTask -TaskName $taskname -TaskDescription $taskDescription -Trigger $trigger -AllowedUser $allowedUser -ScriptPath $scriptPath

            [xml]$taskXML = Export-ScheduledTask -TaskName "Specsavers\$taskname"

            if ($taskxml.task.triggers.LogonTrigger -eq "") {                
                $triggerResult = "AtLogon"
            } elseif ($taskxml.task.triggers.BootTrigger -eq "") {                
                $triggerResult = "AtStartup"
            }
           
            $triggerResult | Should -Be 'AtStartup'
        }
    }

    Context "When using the -trigger 'AtLogon' parameter" -tag "test" {
        It "Should set the trigger based on the supplied value" {    
                       
            $trigger = 'AtLogon'
           
            New-OHScheduledTask -TaskName $taskname -TaskDescription $taskDescription -Trigger $trigger -AllowedUser $allowedUser -ScriptPath $scriptPath

            [xml]$taskXML = Export-ScheduledTask -TaskName "Specsavers\$taskname"

            if ($taskxml.task.triggers.LogonTrigger -eq "") {                
                $triggerResult = "AtLogon"
            } elseif ($taskxml.task.triggers.BootTrigger -eq "") {                
                $triggerResult = "AtStartup"
            }
           
            $triggerResult | Should -Be 'AtLogon'
        }
    }

    Context "When ommitting the -TaskPath parameter" {
        It "Should set the scheduled task folder to 'Specsavers'" {
           
            New-OHScheduledTask -TaskName $taskname -TaskDescription $taskDescription -Trigger $trigger -AllowedUser $allowedUser -ScriptPath $scriptPath

            $taskdetails = Get-ScheduledTaskInfo -TaskName "Specsavers\$taskname" -ErrorAction SilentlyContinue
           
            if ($taskdetails) {
                $result = $true
            } else {
                $result = $false
            }

            $result | Should -BeTrue
        }
    }
}