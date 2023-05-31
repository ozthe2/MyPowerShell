Function Remove-OHScheduledTask {
    <#
    .SYNOPSIS
    Removes scheduled tasks from the local system.

    .DESCRIPTION
    The Remove-OHScheduledTask function removes one or more scheduled tasks from the local system. It uses the Get-ScheduledTask cmdlet to check if the task exists, and if so, it unregisters the task using the Unregister-ScheduledTask cmdlet.

    .PARAMETER TaskName
    Specifies the name of the scheduled task to remove. This parameter accepts pipeline input.

    .INPUTS
    System.String

    .OUTPUTS
    None

    .EXAMPLE
    PS C:\> Remove-OHScheduledTask -TaskName "Task1"
    Removes the scheduled task with the name "Task1" from the local system.

    .EXAMPLE
    PS C:\> "Task1", "Task2" | Remove-OHScheduledTask
    Removes multiple scheduled tasks named "Task1" and "Task2" from the local system.

    .NOTES
    Author: owen.heaume
    Date: 23-May-2023
    Version: 1.0
    #>

    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(ValueFromPipeline = $true)]
        [string]$TaskName
    )

    Process {
        foreach ($name in $TaskName) {
            try {
                if (Get-ScheduledTask -TaskName $name -ErrorAction Stop) {
                    if ($PSCmdlet.ShouldProcess($name, 'Remove')) {
                        Unregister-ScheduledTask -TaskName $name -Confirm:$false
                        Write-Verbose "Removed Scheduled Task $name"
                    }
                }
            } catch {
                Write-Warning "$name not found. No task to delete"
                return $false
            }
        }
    }
}