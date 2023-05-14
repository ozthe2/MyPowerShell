function New-OHShortcut {
    <#
    .SYNOPSIS
    Creates or deletes a shortcut to a specified item on the desktop and/or start menu.
    
    .DESCRIPTION
    The function creates or deletes a shortcut to a specified item on the desktop and/or start menu. 
    It uses Windows Script Host to create the shortcut file.
    
    .PARAMETER ShortcutName
    Specifies the name of the shortcut.
    
    .PARAMETER TargetPath
    Specifies the path of the item that the shortcut opens.
    
    .PARAMETER WorkingDirectory
    Specifies the working directory for the target. 
    
    .PARAMETER IconLocation
    Specifies the location of the icon for the shortcut.
    
    .PARAMETER WindowsStyle
    Specifies the window style for the shortcut. 3 = Maximised, 4 = Normal, 7 = Minimised. 
    
    .PARAMETER AddLocation
    Specifies the location of the shortcut to add: Desktop, StartMenu, or Both.
    
    .PARAMETER ReplaceLocation
    Specifies the location of the shortcut to replace: Desktop, StartMenu, or Both.
    
    .PARAMETER DeleteLocation
    Specifies the location of the shortcut to delete: Desktop, StartMenu, or Both.
    
    .PARAMETER Arguments
    Specifies the arguments to use when opening the target.
    
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
    Version: 1.2
    #>
    
    [CmdletBinding(DefaultParameterSetName = "Add")]
    Param(
        [Parameter(Mandatory = $true, HelpMessage = "Specifies the name of the shortcut.")]
        [string]$ShortcutName,
    
        [Parameter(Mandatory = $true, HelpMessage = "Specifies the path of the item that the shortcut opens.")]
        [string]$TargetPath,
    
        [Parameter(Mandatory = $false, HelpMessage = "Specifies the working directory for the target.")]
        [string]$WorkingDirectory = "",
    
        [Parameter(Mandatory = $true, HelpMessage = "Specifies the location of the icon for the shortcut.")]
        [string]$IconLocation,
    
        [Parameter(Mandatory = $false, HelpMessage = "Specifies the window style for the shortcut.")]
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
    
        [Parameter(Mandatory = $false, HelpMessage = "Specifies the arguments to use when opening the target.")]
        [string]$Arguments
    )     
    
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
                Write-Host "Desktop shortcut '$ShortcutName' has been deleted." -ForegroundColor Yellow
            }
            else {
                Write-Host "Desktop shortcut '$ShortcutName' does not exist." -ForegroundColor Cyan
            }
        }
    
        "StartMenu" {
            $shortcutPath = Join-Path $startMenuPath "$ShortcutName.lnk"
            if (Test-Path $shortcutPath) {
                Remove-Item $shortcutPath
                Write-Host "Start menu shortcut '$ShortcutName' has been deleted." -ForegroundColor Yellow
            }
            else {
                Write-Host "Start menu shortcut '$ShortcutName' does not exist." -ForegroundColor Cyan
            }
        }
    
        "Both" {
            if (Test-Path $shortcutPath) {
                Remove-Item $shortcutPath
                Write-Host "Desktop shortcut '$ShortcutName' has been deleted." -ForegroundColor Yellow
            }
            else {
                Write-Host "Desktop shortcut '$ShortcutName' does not exist." -ForegroundColor Cyan
            }
    
            $shortcutPath = Join-Path $startMenuPath "$ShortcutName.lnk"
            if (Test-Path $shortcutPath) {
                Remove-Item $shortcutPath
                Write-Host "Start menu shortcut '$ShortcutName' has been deleted." -ForegroundColor Yellow
            }
            else {
                Write-Host "Start menu shortcut '$ShortcutName' does not exist." -ForegroundColor Cyan
            }
        }        
    }
    
    switch ($AddLocation) {
        "Desktop" {
            if ($shortcutExists) {
                Write-Host "Shortcut already exists on Desktop." -ForegroundColor Cyan
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
                Write-Host "Desktop shortcut created." -ForegroundColor Green
            }
        }
        "StartMenu" {
            $shortcutPath = Join-Path $StartMenuPath "$ShortcutName.lnk"
            $shortcutExists = Test-Path $shortcutPath
            if ($shortcutExists) {
                Write-Host "Shortcut already exists in Start Menu." -ForegroundColor Cyan
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
                Write-Host "Start Menu shortcut created." -ForegroundColor Green
            }
        }
        "Both" {
            $desktopShortcutPath = Join-Path $desktopPath "$ShortcutName.lnk"
            $startMenuShortcutPath = Join-Path $startMenuPath "$ShortcutName.lnk"
            $desktopShortcutExists = Test-Path $desktopShortcutPath
            $startMenuShortcutExists = Test-Path $startMenuShortcutPath

            if ($desktopShortcutExists -and $startMenuShortcutExists) {
                Write-Host "Shortcut already exists on Desktop and in Start Menu." -ForegroundColor Cyan
                break
            }
            
                # Create desktop shortcut
                $shortcutPath = Join-Path $desktopPath "$ShortcutName.lnk"
                $shortcutExists = Test-Path $shortcutPath
                if ($shortcutExists) {
                    Write-Host "Shortcut already exists on desktop." -ForegroundColor Cyan
                }
                else {
                    $shortcut = $wshell.CreateShortcut($shortcutPath)
                    $shortcut.TargetPath = $TargetPath
                    $shortcut.WorkingDirectory = $WorkingDirectory
                    $shortcut.IconLocation = $IconLocation
                    $shortcut.WindowStyle = $WindowsStyle
                    $shortcut.Arguments = $Arguments
                    $shortcut.Save()
                    Write-Host "Desktop shortcut created." -ForegroundColor Green
                }
            
                # Create start menu shortcut
                $shortcutPath = Join-Path $StartMenuPath "$ShortcutName.lnk"
                $shortcutExists = Test-Path $shortcutPath
                if ($shortcutExists) {
                    Write-Host "Shortcut already exists in Start Menu." -ForegroundColor Cyan
                }
                else {
                    $shortcut = $wshell.CreateShortcut($shortcutPath)
                    $shortcut.TargetPath = $TargetPath
                    $shortcut.WorkingDirectory = $WorkingDirectory
                    $shortcut.IconLocation = $IconLocation
                    $shortcut.WindowStyle = $WindowsStyle
                    $shortcut.Arguments = $Arguments
                    $shortcut.Save()
                    Write-Host "Start Menu shortcut created." -ForegroundColor Green
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
                Write-Host "Shortcut '$ShortcutName' on desktop replaced successfully." -ForegroundColor Green
            }
            else {
                $shortcut = $wshell.CreateShortcut($shortcutPath)
                $shortcut.TargetPath = $TargetPath
                $shortcut.WorkingDirectory = $WorkingDirectory
                $shortcut.IconLocation = $IconLocation
                $shortcut.WindowStyle = $WindowsStyle
                $shortcut.Arguments = $Arguments
                $shortcut.Save()
                Write-Host "Shortcut '$ShortcutName' created successfully on desktop." -ForegroundColor Green
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
                Write-Host "Shortcut '$ShortcutName' in start menu replaced successfully." -ForegroundColor Green
            }
            else {
                $shortcut = $wshell.CreateShortcut($shortcutPath)
                $shortcut.TargetPath = $TargetPath
                $shortcut.WorkingDirectory = $WorkingDirectory
                $shortcut.IconLocation = $IconLocation
                $shortcut.WindowStyle = $WindowsStyle
                $shortcut.Arguments = $Arguments
                $shortcut.Save()
                Write-Host "Shortcut '$ShortcutName' created successfully in start menu." -ForegroundColor Green
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
                Write-Host "Shortcut '$ShortcutName' replaced successfully on both desktop and start menu." -ForegroundColor Green
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
                Write-Host "Shortcut '$ShortcutName' created successfully on both desktop and start menu." -ForegroundColor Green
            }
            break
        }
    }   
}
