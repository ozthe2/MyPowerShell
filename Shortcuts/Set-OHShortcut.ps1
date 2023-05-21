Function Set-OHShortcut {
    <#
    .SYNOPSIS
    Updates an existing shortcut with new settings.

    .DESCRIPTION
    The Set-OHShortcut function is used to update an existing shortcut with new settings. This function can update shortcuts on the desktop, start menu, or both.

    .PARAMETER ShortcutName
    Specifies the name of the shortcut that needs to be updated.

    .PARAMETER TargetPath
    Specifies the path of the item that the shortcut opens.

    .PARAMETER UpdateLocation
    Specifies the location where the shortcut should be updated. Valid values are 'Desktop', 'StartMenu', or 'Both'. The default value is 'Desktop'.

    .PARAMETER WorkingDirectory
    Specifies the working directory for the target item.

    .PARAMETER IconLocation
    Specifies the location of the icon for the shortcut.

    .PARAMETER WindowsStyle
    Specifies the window style for the shortcut. Valid values are 3, 7, or 4. The default value is 7.

    .PARAMETER Arguments
    Specifies the arguments to use when opening the target item.

    .PARAMETER UseLoggedInUsersDesktop
    Switch parameter. If specified, the shortcut will be updated on the logged-in user's desktop. Otherwise, it will be updated on the public desktop.

    .PARAMETER UseLoggedInUsersStartMenu
    Switch parameter. If specified, the shortcut will be updated in the logged-in user's start menu. Otherwise, it will be updated in all users' start menu.

    .EXAMPLE
    Set-OHShortcut -ShortcutName "MyShortcut" -TargetPath "C:\MyApp.exe" -UpdateLocation 'Both' -WorkingDirectory "C:\MyApp" -IconLocation "C:\MyApp\icon.ico" -WindowsStyle 3 -Arguments "-option1 -option2"

    Updates the shortcut named "MyShortcut" with the specified settings. The shortcut will be updated on both the desktop and start menu. The target path is set to "C:\MyApp.exe", the working directory is set to "C:\MyApp", the icon location is set to "C:\MyApp\icon.ico", the window style is set to 3, and the arguments are set to "-option1 -option2".

    .NOTES
    Author: owen.heaume
    Date: 23-May-2023
    Version: 1.0
    #>    
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true, HelpMessage = "Specifies the name of the shortcut.")]
        [string]$ShortcutName,

        [Parameter(Mandatory = $true, HelpMessage = "Specifies the path of the item that the shortcut opens.")]
        [string]$TargetPath,

        [Parameter(Mandatory = $false, HelpMessage = "Specifies the location that the shortcut is created on.")]
        [ValidateSet('Desktop', 'StartMenu', 'Both')]
        [string]$UpdateLocation = 'Desktop',

        [Parameter(Mandatory = $false, HelpMessage = "Specifies the working directory for the target.")]
        [string]$WorkingDirectory = "",

        [Parameter(Mandatory = $false, HelpMessage = "Specifies the location of the icon for the shortcut.")]
        [string]$IconLocation = 'imageres.dll,4',

        [Parameter(Mandatory = $false, HelpMessage = "Specifies the window style for the shortcut.")]
        [ValidateSet(3, 7, 4)]
        [int]$WindowsStyle = 7,

        [Parameter(Mandatory = $false, HelpMessage = "Specifies the arguments to use when opening the target.")]
        [string]$Arguments,

        [Switch]$UseLoggedInUsersDesktop,

        [Switch]$UseLoggedInUsersStartMenu
    )

    # Construct paths to the desktop and start menu folders
    if ($UseLoggedInUsersDesktop) {
        $desktopPath = [Environment]::GetFolderPath("Desktop") # Logged-in user's desktop
    } else {
        $desktopPath = [Environment]::GetFolderPath("CommonDesktopDirectory") # Public desktop
    }

    if ($UseLoggedInUsersStartMenu) {
        $StartMenuPath = "$env:APPDATA\Microsoft\Windows\Start Menu" # Logged-in user's start menu
    } else {
        $StartMenuPath = [Environment]::GetFolderPath("CommonPrograms") # All users' start menu
    }

    $wshell = New-Object -ComObject WScript.Shell

    $shortcutPath = Join-Path $desktopPath "$ShortcutName.lnk"
    $startmenuPath = Join-Path $StartMenuPath "$ShortcutName.lnk"

    # Update shortcut on desktop
    if (($UpdateLocation -eq 'Desktop' -or $UpdateLocation -eq 'Both') -and (Test-Path $shortcutPath)) {
        try {
            $shortcut = $wshell.CreateShortcut($shortcutPath)
            if ($TargetPath) { $shortcut.TargetPath = $TargetPath }
            if ($PSBoundParameters.ContainsKey('WorkingDirectory')) {
                $workingDirectory = $PSBoundParameters['WorkingDirectory']
                if ($workingDirectory -ne $null) {
                    $shortcut.WorkingDirectory = $workingDirectory  # Assign new working directory
                }
            }

            if ($IconLocation -eq "") { 
                $shortcut.IconLocation = 'imageres.dll,4' # set default icon
            } elseif ($IconLocation) {
                $shortcut.iconLocation = $IconLocation
            }
            if ($WindowsStyle) { $shortcut.WindowStyle = $WindowsStyle }
            if ($PSBoundParameters.ContainsKey('Arguments')) {
                $arguments = $PSBoundParameters['Arguments']
                if ($arguments -ne $null) {
                    $shortcut.Arguments = $arguments  # Assign new argument
                }
            }        
            $shortcut.Save()
            Write-Verbose "Shortcut '$ShortcutName' on the desktop updated."
        } catch {
            Write-Error "An error occurred: $($_.Exception.Message)"
        }
    } elseif (($UpdateLocation -eq 'Desktop' -or $UpdateLocation -eq 'Both') -and (!(Test-Path $shortcutPath))) {
        Write-Warning "Unable to modify desktop shortcut: $($shortcut) as it does not exist."
    }
    
   # Update shortcut on start menu
    if (($UpdateLocation -eq 'StartMenu' -or $UpdateLocation -eq 'Both') -and (Test-Path $startmenuPath)) {
        try {
            $shortcut = $wshell.CreateShortcut($startmenuPath)
            if ($TargetPath) { $shortcut.TargetPath = $TargetPath }
            if ($PSBoundParameters.ContainsKey('WorkingDirectory')) {
                $workingDirectory = $PSBoundParameters['WorkingDirectory']
                if ($workingDirectory -ne $null) {
                    $shortcut.WorkingDirectory = $workingDirectory  # Assign new working directory
                }
            }

            if ($IconLocation -eq "") { 
                $shortcut.IconLocation = 'imageres.dll,4' # set default icon
            } elseif ($IconLocation) {
                $shortcut.iconLocation = $IconLocation
            }
            if ($WindowsStyle) { $shortcut.WindowStyle = $WindowsStyle }
            if ($PSBoundParameters.ContainsKey('Arguments')) {
                $arguments = $PSBoundParameters['Arguments']
                if ($arguments -ne $null) {
                    $shortcut.Arguments = $arguments  # Assign new argument
                }
            }
            $shortcut.Save()
            Write-Verbose "Shortcut '$ShortcutName' in the start menu updated."
        } catch {
            Write-Error "An error occurred: $($_.Exception.Message)"
        }
    } elseif (($UpdateLocation -eq 'StartMenu' -or $UpdateLocation -eq 'Both') -and (!(Test-Path $startmenuPath))) {
        Write-Warning "Unable to modify start menu shortcut: $shortcut as it does not exist."
    }
}
