BeforeAll {
    . "..\Remove-OHScheduledTask.ps1"
}

Describe "My test" {
    $taskName = 'OHPesterTest'
    It "should return false if the task does not exist" {
        $taskName = 'OHPesterTest'

        $taskExists = Remove-OHScheduledTask -TaskName $taskName
        $taskExists | Should -BeFalse -Because "the task does not exist to delete."
    }

    It "it should delete the task" {
        $taskName = 'OHPesterTest'

        # Set up the scheduled task for the tests
        $action = New-ScheduledTaskAction -Execute "Taskmgr.exe"
        $trigger = New-ScheduledTaskTrigger -AtLogon
        $settings = New-ScheduledTaskSettingsSet
        $task = New-ScheduledTask -Action $action  -Trigger $trigger -Settings $settings
        Register-ScheduledTask -TaskName $taskName -InputObject $task -ErrorAction SilentlyContinue

        Remove-OHScheduledTask -TaskName $taskName
        
        # Check if the task exists after deletion
        $taskExistsAfterDeletion = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
        $taskExistsAfterDeletion | Should -BeNullOrEmpty -Because "The task should not exist after deletion"
    } 

    It "it should delete the task via pipeline input" {
        $taskName = 'OHPesterTest'
       
        # Set up the scheduled task for the tests
        $action = New-ScheduledTaskAction -Execute "Taskmgr.exe"
        $trigger = New-ScheduledTaskTrigger -AtLogon
        $settings = New-ScheduledTaskSettingsSet
        $task = New-ScheduledTask -Action $action  -Trigger $trigger -Settings $settings
        Register-ScheduledTask -TaskName $taskName -InputObject $task -ErrorAction SilentlyContinue        

        $taskName | Remove-OHScheduledTask       

        # Check if the task exists after deletion
        $taskExistsAfterDeletion = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
        $taskExistsAfterDeletion | Should -BeNullOrEmpty -Because "The task should not exist after deletion" 
    }
}