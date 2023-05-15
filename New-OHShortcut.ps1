function New-OHShortcut {
    <#
    .SYNOPSIS
    Creates or deletes a shortcut to a specified item on the desktop and/or start menu.
    
    .DESCRIPTION
    The function creates or deletes a shortcut to a specified item on the desktop and/or start menu. 
    It uses Windows Script Host to create the shortcut file.
    
    .PARAMETER ShortcutName
    Specifies the name of the shortcut. This parameter is mandatory.
    
    .PARAMETER TargetPath
    Specifies the path of the item that the shortcut opens. This parameter is mandatory.
    
    .PARAMETER WorkingDirectory
    Specifies the working directory for the target. This is an optional parameter. 
    
    .PARAMETER IconLocation
    Specifies the location of the icon for the shortcut. If not provided then a default folder icon is used. This is an optional parameter.
    
    .PARAMETER WindowsStyle
    Specifies the window style for the shortcut. 3 = Maximised, 4 = Normal, 7 = Minimised. Defaults to '7'. This is an optional parameter. 
    
    .PARAMETER AddLocation
    Specifies the location of the shortcut to add: Desktop, StartMenu, or Both. This parameter is mandatory
    
    .PARAMETER ReplaceLocation
    Specifies the location of the shortcut to replace: Desktop, StartMenu, or Both. This parameter is mandatory
    
    .PARAMETER DeleteLocation
    Specifies the location of the shortcut to delete: Desktop, StartMenu, or Both. This parameter is mandatory
    
    .PARAMETER Arguments
    Specifies the arguments to use when opening the target. This is an optional parameter.
    
    .EXAMPLE
    New-OHShortcut -ShortcutName "Notepad" -TargetPath "C:\Windows\System32\notepad.exe" -AddLocation Desktop -IconLocation "C:\Windows\System32\notepad.exe"
    Creates a shortcut to Notepad on the desktop.
    
    .EXAMPLE
    New-OHShortcut -ShortcutName "Calculator" -TargetPath "C:\Windows\System32\calc.exe" -AddLocation StartMenu -IconLocation "C:\Windows\System32\calc.exe"
    Creates a shortcut to Calculator in the Start menu.
    
    .EXAMPLE
    New-OHShortcut -ShortcutName "MyApp" -TargetPath "C:\Program Files\MyApp\MyApp.exe" -AddLocation Both -IconLocation "C:\Program Files\MyApp\MyApp.exe"
    Creates a shortcut to MyApp on the desktop and in the Start menu.
    
    .EXAMPLE
    New-OHShortcut -ShortcutName "MyApp" -ReplaceLocation Both -TargetPath "C:\Program Files\MyApp\MyApp.exe" -IconLocation "C:\Program Files\MyApp\MyApp.exe"
    Replaces the shortcut to MyApp on the desktop and in the Start menu.
    
    .EXAMPLE
    New-OHShortcut -ShortcutName "MyApp" -DeleteLocation Both
    Deletes the shortcut to MyApp on the desktop and in the Start menu.
    
    .NOTES
    This function uses Windows Script Host to create the shortcut file. The script may fail if WScript is not installed or if the user account does not have sufficient privileges to create files in the specified locations.
    
    Created By: owen.heaume
    Date: 12-May-2023
    Version: 1.5 (15-May-2023)
            - Changed all Write-Host to Write-Verbose
            - Added default icon if path to one was not supplied  
    #>
    
    [CmdletBinding(DefaultParameterSetName = "Add")]
    Param(
        [Parameter(Mandatory = $true, HelpMessage = "Specifies the name of the shortcut.")]
        [string]$ShortcutName,
    
        [Parameter(Mandatory = $true,ParameterSetName = "Add", HelpMessage = "Specifies the path of the item that the shortcut opens.")]        
        [Parameter(Mandatory = $true,ParameterSetName = "Replace")]
        [string]$TargetPath,
    
        [Parameter(Mandatory = $false,ParameterSetName = "Add", HelpMessage = "Specifies the working directory for the target.")]        
        [Parameter(Mandatory = $false,ParameterSetName = "Replace")]
        [string]$WorkingDirectory = "",
    
        [Parameter(Mandatory = $false, ParameterSetName = "Add",HelpMessage = "Specifies the location of the icon for the shortcut.")]        
        [Parameter(Mandatory = $false,ParameterSetName = "Replace")]
        [string]$IconLocation,
    
        [Parameter(Mandatory = $false, ParameterSetName = "Add",HelpMessage = "Specifies the window style for the shortcut.")]        
        [Parameter(Mandatory = $false,ParameterSetName = "Replace")]
        [ValidateSet(3, 7, 4)]
        [int]$WindowsStyle = 7,
    
        [Parameter(Mandatory = $true, ParameterSetName = "Add", HelpMessage = "Specifies the location of the shortcut to add.")]
        [ValidateSet("Desktop", "StartMenu", "Both")]
        [string]$AddLocation,
    
        [Parameter(Mandatory = $true, ParameterSetName = "Replace", HelpMessage = "Specifies the location of the shortcut to replace.")]
        [ValidateSet("Desktop", "StartMenu", "Both")]
        [string]$ReplaceLocation,
    
        [Parameter(Mandatory = $true, ParameterSetName = "Delete", HelpMessage = "Specifies the location of the shortcut to delete.")]
        [ValidateSet("Desktop", "StartMenu", "Both")]
        [string]$DeleteLocation,
    
        [Parameter(Mandatory = $false, ParameterSetName = "Add",HelpMessage = "Specifies the arguments to use when opening the target.")]        
        [Parameter(Mandatory = $false,ParameterSetName = "Replace")]
        [string]$Arguments
    )     
    
    # Set default shortcut icon if one was not given
    if (!$IconLocation) { $IconLocation = 'imageres.dll,4' }

    # Construct paths to the desktop and start menu folders
    $desktopPath = [Environment]::GetFolderPath("Desktop")
    $StartMenuPath = [Environment]::GetFolderPath("CommonPrograms")
    
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
                Write-Verbose "DELETE: Desktop shortcut '$ShortcutName' has been deleted."
            }
            else {
                Write-Verbose "DELETE: Desktop shortcut '$ShortcutName' does not exist."
            }
        }
    
        "StartMenu" {
            $shortcutPath = Join-Path $startMenuPath "$ShortcutName.lnk"
            if (Test-Path $shortcutPath) {
                Remove-Item $shortcutPath
                Write-Verbose "DELETE: Start menu shortcut '$ShortcutName' has been deleted."
            }
            else {
                Write-Verbose "DELETE: Start menu shortcut '$ShortcutName' does not exist."
            }
        }
    
        "Both" {
            if (Test-Path $shortcutPath) {
                Remove-Item $shortcutPath
                Write-Verbose "DELETE: Desktop shortcut '$ShortcutName' has been deleted."
            }
            else {
                Write-Verbose "DELETE: Desktop shortcut '$ShortcutName' does not exist."
            }
    
            $shortcutPath = Join-Path $startMenuPath "$ShortcutName.lnk"
            if (Test-Path $shortcutPath) {
                Remove-Item $shortcutPath
                Write-Verbose "DELETE: Start menu shortcut '$ShortcutName' has been deleted."
            }
            else {
                Write-Verbose "DELETE: Start menu shortcut '$ShortcutName' does not exist."
            }
        }        
    }
    
    switch ($AddLocation) {
        "Desktop" {
            if ($shortcutExists) {
                Write-Verbose "ADD: Shortcut '$ShortcutName' already exists on Desktop."
            }
            else {
                # Create desktop shortcut
                $shortcut = $wshell.CreateShortcut($shortcutPath)
                $shortcut.TargetPath = $TargetPath
                $shortcut.WorkingDirectory = $WorkingDirectory
                $shortcut.IconLocation = $IconLocation
                $shortcut.WindowStyle = $WindowsStyle
                $shortcut.Arguments = $Arguments
                $shortcut.Save()
                Write-Verbose "ADD: Desktop shortcut '$ShortcutName' created."
            }
        }
        "StartMenu" {
            $shortcutPath = Join-Path $StartMenuPath "$ShortcutName.lnk"
            $shortcutExists = Test-Path $shortcutPath
            if ($shortcutExists) {
                Write-Verbose "ADD: Shortcut '$ShortcutName'already exists in Start Menu."
            }
            else {
                # Create start menu shortcut
                $shortcut = $wshell.CreateShortcut($shortcutPath)
                $shortcut.TargetPath = $TargetPath
                $shortcut.WorkingDirectory = $WorkingDirectory
                $shortcut.IconLocation = $IconLocation
                $shortcut.WindowStyle = $WindowsStyle
                $shortcut.Arguments = $Arguments
                $shortcut.Save()
                Write-Verbose "ADD: Start Menu shortcut '$ShortcutName' created."
            }
        }
        "Both" {
            # Create desktop shortcut
            $shortcutPath = Join-Path $desktopPath "$ShortcutName.lnk"
            $shortcutExists = Test-Path $shortcutPath
            if ($shortcutExists) {
                Write-Verbose "ADD: Shortcut '$ShortcutName' already exists on desktop."
            }
            else {
                $shortcut = $wshell.CreateShortcut($shortcutPath)
                $shortcut.TargetPath = $TargetPath
                $shortcut.WorkingDirectory = $WorkingDirectory
                $shortcut.IconLocation = $IconLocation
                $shortcut.WindowStyle = $WindowsStyle
                $shortcut.Arguments = $Arguments
                $shortcut.Save()
                Write-Verbose "ADD: Desktop shortcut '$ShortcutName' created."
            }
            
            # Create start menu shortcut
            $shortcutPath = Join-Path $StartMenuPath "$ShortcutName.lnk"
            $shortcutExists = Test-Path $shortcutPath
            if ($shortcutExists) {
                Write-Verbose "ADD: Shortcut '$ShortcutName' already exists in Start Menu."
            }
            else {
                $shortcut = $wshell.CreateShortcut($shortcutPath)
                $shortcut.TargetPath = $TargetPath
                $shortcut.WorkingDirectory = $WorkingDirectory
                $shortcut.IconLocation = $IconLocation
                $shortcut.WindowStyle = $WindowsStyle
                $shortcut.Arguments = $Arguments
                $shortcut.Save()
                Write-Verbose "ADD: Start Menu shortcut created."
            }
            
        }
    }
    
    switch ($ReplaceLocation) {
        "Desktop" {
            if ($shortcutExists) {
                Remove-Item $shortcutPath
                $shortcut = $wshell.CreateShortcut($shortcutPath)
                $shortcut.TargetPath = $TargetPath
                $shortcut.WorkingDirectory = $WorkingDirectory
                $shortcut.IconLocation = $IconLocation
                $shortcut.WindowStyle = $WindowsStyle
                $shortcut.Arguments = $Arguments
                $shortcut.Save()
                Write-Verbose "REPLACE: Shortcut '$ShortcutName' on desktop replaced successfully."
            }
            else {
                $shortcut = $wshell.CreateShortcut($shortcutPath)
                $shortcut.TargetPath = $TargetPath
                $shortcut.WorkingDirectory = $WorkingDirectory
                $shortcut.IconLocation = $IconLocation
                $shortcut.WindowStyle = $WindowsStyle
                $shortcut.Arguments = $Arguments
                $shortcut.Save()
                Write-Verbose "REPLACE: Shortcut '$ShortcutName' created successfully on desktop."
            }
        }
        "StartMenu" {
            $shortcutPath = Join-Path $startMenuPath "$ShortcutName.lnk"
            $shortcutExists = Test-Path $shortcutPath
            if ($shortcutExists) {
                Remove-Item $shortcutPath
                $shortcut = $wshell.CreateShortcut($shortcutPath)
                $shortcut.TargetPath = $TargetPath
                $shortcut.WorkingDirectory = $WorkingDirectory
                $shortcut.IconLocation = $IconLocation
                $shortcut.WindowStyle = $WindowsStyle
                $shortcut.Arguments = $Arguments
                $shortcut.Save()
                Write-Verbose "REPLACE: Shortcut '$ShortcutName' in start menu replaced successfully."
            }
            else {
                $shortcut = $wshell.CreateShortcut($shortcutPath)
                $shortcut.TargetPath = $TargetPath
                $shortcut.WorkingDirectory = $WorkingDirectory
                $shortcut.IconLocation = $IconLocation
                $shortcut.WindowStyle = $WindowsStyle
                $shortcut.Arguments = $Arguments
                $shortcut.Save()
                Write-Verbose "REPLACE: Shortcut '$ShortcutName' created successfully in start menu."
            }
            break
        }
        "Both" {
            $desktopShortcutPath = Join-Path $desktopPath "$ShortcutName.lnk"
            $startMenuShortcutPath = Join-Path $startMenuPath "$ShortcutName.lnk"
            $desktopShortcutExists = Test-Path $desktopShortcutPath
            $startMenuShortcutExists = Test-Path $startMenuShortcutPath
            if ($desktopShortcutExists -and $startMenuShortcutExists) {
                Remove-Item $desktopShortcutPath
                Remove-Item $startMenuShortcutPath
                $desktopShortcut = $wshell.CreateShortcut($desktopShortcutPath)
                $desktopShortcut.TargetPath = $TargetPath
                $desktopShortcut.WorkingDirectory = $WorkingDirectory
                $desktopShortcut.IconLocation = $IconLocation
                $desktopShortcut.WindowStyle = $WindowsStyle
                $desktopShortcut.Arguments = $Arguments
                $desktopShortcut.Save()
                $startMenuShortcut = $wshell.CreateShortcut($startMenuShortcutPath)
                $startMenuShortcut.TargetPath = $TargetPath
                $startMenuShortcut.WorkingDirectory = $WorkingDirectory
                $startMenuShortcut.IconLocation = $IconLocation
                $startMenuShortcut.WindowStyle = $WindowsStyle
                $startMenuShortcut.Arguments = $Arguments
                $startMenuShortcut.Save()
                Write-Verbose "REPLACE: Shortcut '$ShortcutName' replaced successfully on both desktop and start menu."
            }
            else {
                $desktopShortcut = $wshell.CreateShortcut($desktopShortcutPath)
                $desktopShortcut.TargetPath = $TargetPath
                $desktopShortcut.WorkingDirectory = $WorkingDirectory
                $desktopShortcut.IconLocation = $IconLocation
                $desktopShortcut.WindowStyle = $WindowsStyle
                $desktopShortcut.Arguments = $Arguments
                $desktopShortcut.Save()
                $startMenuShortcut = $wshell.CreateShortcut($startMenuShortcutPath)
                $startMenuShortcut.TargetPath = $TargetPath
                $startMenuShortcut.WorkingDirectory = $WorkingDirectory
                $startMenuShortcut.IconLocation = $IconLocation
                $startMenuShortcut.WindowStyle = $WindowsStyle
                $startMenuShortcut.Arguments = $Arguments
                $startMenuShortcut.Save()
                Write-Verbose "REPLACE: Shortcut '$ShortcutName' created successfully on both desktop and start menu."
            }
            break
        }
    }   
}
