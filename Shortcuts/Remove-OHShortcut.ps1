Function Remove-OHShortcut {
    <#
    .SYNOPSIS
    Removes a shortcut from the specified location(s).

    .DESCRIPTION
    The Remove-OHShortcut function is used to delete a shortcut from the desktop, start menu, or both locations.

    .PARAMETER ShortcutName
    Specifies the name of the shortcut to delete.

    .PARAMETER DeleteLocation
    Specifies the location of the shortcut to delete. Valid values are "Desktop", "StartMenu", or "Both".

    .PARAMETER UseLoggedInUsersDesktop
    Specifies whether to use the logged-in user's desktop location for deleting the shortcut. If this switch is specified, the shortcut will be deleted from the logged-in user's desktop. Otherwise, it will be deleted from the common desktop directory.

    .PARAMETER UseLoggedInUsersStartMenu
    Specifies whether to use the logged-in user's start menu location for deleting the shortcut. If this switch is specified, the shortcut will be deleted from the logged-in user's start menu. Otherwise, it will be deleted from the common programs directory.

    .EXAMPLE
    Remove-OHShortcut -ShortcutName "MyApp" -DeleteLocation "Desktop"
    Removes the shortcut named "MyApp" from the desktop.

    .EXAMPLE
    Remove-OHShortcut -ShortcutName "MyApp" -DeleteLocation "StartMenu"
    Removes the shortcut named "MyApp" from the start menu.

    .EXAMPLE
    Remove-OHShortcut -ShortcutName "MyApp" -DeleteLocation "Both"
    Removes the shortcut named "MyApp" from both the desktop and start menu.

    .EXAMPLE
    Remove-OHShortcut -ShortcutName "MyApp" -DeleteLocation "Desktop" -UseLoggedInUsersDesktop
    Removes the shortcut named "MyApp" from the logged-in user's desktop.

    .EXAMPLE
    Remove-OHShortcut -ShortcutName "MyApp" -DeleteLocation "StartMenu" -UseLoggedInUsersStartMenu
    Removes the shortcut named "MyApp" from the logged-in user's start menu.

    .NOTES
    Author: owen.heaume
    Date: 20 May 2023
    Version: 1.0
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true, HelpMessage = "Specifies the name of the shortcut.")]
        [string]$ShortcutName,        
    
        [Parameter(Mandatory = $true, ParameterSetName = "Delete", HelpMessage = "Specifies the location of the shortcut to delete.")]
        [ValidateSet("Desktop", "StartMenu", "Both")]
        [string]$DeleteLocation,

        [Switch]$UseLoggedInUsersDesktop,

        [Switch]$UseLoggedInUsersStartMenu
    )
   
    # Construct paths to the desktop and start menu folders
    if ($UseLoggedInUsersDesktop) {
        $desktopPath = [Environment]::GetFolderPath("Desktop") # Logged in user's desktop
    } else {
        $desktopPath = [Environment]::GetFolderPath("CommonDesktopDirectory") # Public desktop
    }

    if ($UseLoggedInUsersStartMenu) {
        $startMenuPath = "$env:APPDATA\Microsoft\Windows\Start Menu"
    } else {
        $startMenuPath = [Environment]::GetFolderPath("CommonPrograms")
    }

    # Create a WScript Shell object to work with
    $wshell = New-Object -ComObject WScript.Shell
    
    # Construct the full path to the shortcut file
    $shortcutPath = Join-Path $desktopPath "$ShortcutName.lnk"    
    
    # Check if the shortcut file exists
    $shortcutExists = Test-Path $shortcutPath
    
    switch ($DeleteLocation) {
        "Desktop" {
            if (Test-Path $shortcutPath) {
                Remove-Item $shortcutPath
                Write-Verbose "Desktop shortcut '$ShortcutName' has been deleted."
            } else {
                Write-Warning "Desktop shortcut '$ShortcutName' does not exist to delete."
            }
        }
    
        "StartMenu" {
            $shortcutPath = Join-Path $startMenuPath "$ShortcutName.lnk"
            if (Test-Path $shortcutPath) {
                Remove-Item $shortcutPath
                Write-Verbose "Start menu shortcut '$ShortcutName' has been deleted."
            } else {
                Write-Warning "Start menu shortcut '$ShortcutName' does not exist to delete."
            }
        }
    
        "Both" {
            if (Test-Path $shortcutPath) {
                Remove-Item $shortcutPath
                Write-Verbose "Desktop shortcut '$ShortcutName' has been deleted."
            } else {
                Write-Warning "Desktop shortcut '$ShortcutName' does not exist to delete."
            }
    
            $shortcutPath = Join-Path $startMenuPath "$ShortcutName.lnk"
            if (Test-Path $shortcutPath) {
                Remove-Item $shortcutPath
                Write-Verbose "Start menu shortcut '$ShortcutName' has been deleted."
            } else {
                Write-Warning "Start menu shortcut '$ShortcutName' does not exist to delete."
            }
        }
    }
}
