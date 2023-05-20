Function New-OHshortcut {
    <#
    .SYNOPSIS
    Creates a shortcut on the desktop or start menu.

    .DESCRIPTION
    The New-OHShortcut function allows you to create a shortcut on the desktop or start menu of a Windows operating system. You can specify the name of the shortcut, the path of the item that the shortcut opens, the location where the shortcut is created, and additional properties such as the working directory, icon location, window style, and arguments.

    .PARAMETER ShortcutName
    Specifies the name of the shortcut to be created.

    .PARAMETER TargetPath
    Specifies the path of the item that the shortcut opens.

    .PARAMETER CreateLocation
    Specifies the location where the shortcut is created. Valid values are 'Desktop', 'StartMenu', or 'Both'. The default value is 'Desktop'.

    .PARAMETER WorkingDirectory
    Specifies the working directory for the target of the shortcut.

    .PARAMETER IconLocation
    Specifies the location of the icon for the shortcut.

    .PARAMETER WindowsStyle
    Specifies the window style for the shortcut. Valid values are 3 (Maximized), 7 (Minimized), or 4 (Normal). The default value is 7 (Minimized).

    .PARAMETER Arguments
    Specifies the arguments to use when opening the target of the shortcut.

    .PARAMETER UseLoggedInUsersDesktop
    Creates the shortcut on the logged in user's desktop. By default, the shortcut is created on the public desktop.

    .PARAMETER UseLoggedInUsersStartMenu
    Creates the shortcut in the logged in user's start menu. By default, the shortcut is created in the all users start menu.

    .EXAMPLE
    New-OHShortcut -ShortcutName "MyApp" -TargetPath "C:\Path\To\MyApp.exe" -CreateLocation Desktop -IconLocation "C:\Path\To\Icon.ico" -WindowsStyle 4
    Creates a shortcut named "MyApp" on the desktop that opens the "MyApp.exe" file located at "C:\Path\To\MyApp.exe". The shortcut uses a custom icon located at "C:\Path\To\Icon.ico" and has a normal window style.

    .EXAMPLE
    New-OHShortcut -ShortcutName "MyApp" -TargetPath "C:\Path\To\MyApp.exe" -CreateLocation StartMenu -UseLoggedInUsersStartMenu -WindowsStyle 3
    Creates a shortcut named "MyApp" in the logged in user's start menu that opens the "MyApp.exe" file located at "C:\Path\To\MyApp.exe". The shortcut has a maximized window style.

    .NOTES
    Created by: owen.heaume
    Date: 20-May-2023
    Version: 1.0

    #>
    [CmdletBinding(DefaultParameterSetName = "Add")]
    Param(
        [Parameter(Mandatory = $true, HelpMessage = "Specifies the name of the shortcut.")]
        [string]$ShortcutName,
    
        [Parameter(Mandatory = $true, HelpMessage = "Specifies the path of the item that the shortcut opens.")]
        [string]$TargetPath,

        [Parameter(Mandatory = $false, HelpMessage = "Specifies the location that the shortcut is created on.")]
        [ValidateSet('Desktop', 'StartMenu', 'Both')]
        [string]$CreateLocation = 'Desktop',

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

        [switch]$UseLoggedInUsersStartMenu
    )  

    # Set default shortcut icon if one was not given
    if (!$IconLocation) { $IconLocation = 'imageres.dll,4' }

    # Construct paths to the desktop and start menu folders
    If ($UseLoggedInUsersDesktop) {
        $desktopPath = [Environment]::GetFolderPath("Desktop") # Logged in users desktop
    } else {
        $desktopPath = [Environment]::GetFolderPath("CommonDesktopDirectory") # Public desktop
    }

    if ($UseLoggedInUsersStartMenu) {
        $StartMenuPath = "$env:APPDATA\Microsoft\Windows\Start Menu" # Logged in users start menu
    } else {
        $StartMenuPath = [Environment]::GetFolderPath("CommonPrograms") # All users start menu
    }

    $wshell = New-Object -ComObject WScript.Shell

    $shortcutPath = Join-Path $desktopPath "$ShortcutName.lnk" 
    $startmenuPath = Join-Path $StartMenuPath "$ShortcutName.lnk"   

    $shortcutExists = Test-Path $shortcutPath
    $startmenuExists = Test-Path $StartMenuPath

    # Create shortcut on desktop
    if (($CreateLocation -eq 'Desktop' -or $CreateLocation -eq 'Both') -and (!$shortcutExists)) {
        try {
            $shortcut = $wshell.CreateShortcut($shortcutPath)
            $shortcut.TargetPath = $TargetPath
            $shortcut.WorkingDirectory = $WorkingDirectory
            $shortcut.IconLocation = $IconLocation
            $shortcut.WindowStyle = $WindowsStyle
            $shortcut.Arguments = $Arguments
            $shortcut.Save()
            Write-Verbose "Shortcut '$ShortcutName' created on Desktop."
        } catch {
            Write-Error "An error occurred: $($_.Exception.Message)"
        }
    } elseif ($CreateLocation -eq 'Desktop' -or $CreateLocation -eq 'Both') {
        Write-Warning "The shortcut $ShortcutName already exists on the desktop."
    }

    # Create shortcut on start menu
    if (($CreateLocation -eq 'StartMenu' -or $CreateLocation -eq 'Both') -and (!(Test-Path $startmenuPath))) {
        try {
            $shortcut = $wshell.CreateShortcut($startmenuPath)
            $shortcut.TargetPath = $TargetPath
            $shortcut.WorkingDirectory = $WorkingDirectory
            $shortcut.IconLocation = $IconLocation
            $shortcut.WindowStyle = $WindowsStyle
            $shortcut.Arguments = $Arguments
            $shortcut.Save()
            Write-Verbose "Shortcut '$ShortcutName' created in the start menu."
        } catch {
            Write-Error "An error occurred: $($_.Exception.Message)"
        }
    } elseif ($CreateLocation -eq 'StartMenu' -or $CreateLocation -eq 'Both') {
        Write-Warning "The shortcut $ShortcutName already exists in the start menu."
    }
}
