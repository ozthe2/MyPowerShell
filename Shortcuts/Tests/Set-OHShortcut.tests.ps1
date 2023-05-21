BeforeAll {
    . ".\Set-OHShortcut.ps1"    
}

Describe "Set-OHShortcut" {
    Context "When changing shortcut arguments" {
        BeforeEach {
            # Create a test shortcut for each scenario
            $desktopShortcutName = "Shortcut1"           
            $desktopShortcutPath = [System.IO.Path]::Combine([Environment]::GetFolderPath("Desktop"), "$desktopShortcutName.lnk") # Logged-in user's desktop
            $shell = New-Object -ComObject WScript.Shell
            $shortcut = $shell.CreateShortcut($desktopShortcutPath)
            $shortcut.Save()
            
            $startMenuShortcutName = "Shortcut2"           
            $startMenuShortcutPath = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\$startMenuShortcutName.lnk" # Logged-in user's start menu
            $shell = New-Object -ComObject WScript.Shell
            $shortcut = $shell.CreateShortcut($startMenuShortcutPath)
            $shortcut.Save()
        }

        AfterEach {
            # Remove the test shortcuts
            $desktopShortcutPath = [System.IO.Path]::Combine([Environment]::GetFolderPath("Desktop"), "$desktopShortcutName.lnk")
            $startMenuShortcutPath = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\$startMenuShortcutName.lnk"
            
            Remove-Item $desktopShortcutPath -Force -ErrorAction SilentlyContinue
            Remove-Item $startMenuShortcutPath -Force -ErrorAction SilentlyContinue
        }

        It "should update shortcut arguments on the desktop" {
            # Arrange
            $desktopShortcutName = "Shortcut1"
            $targetPath = "C:\Path\to\Target.exe"
            $arguments = "-newArg"

            # Act
            Set-OHShortcut -ShortcutName $desktopShortcutName -TargetPath $targetPath -Arguments $arguments -UseLoggedInUsersDesktop

            # Assert
            $shortcut = (New-Object -ComObject WScript.Shell).CreateShortcut($desktopShortcutPath)
            $shortcut.Arguments | Should -Be $arguments
        }

        It "should update shortcut arguments in the start menu" {
            # Arrange
            $startMenuShortcutName = "Shortcut2"
            $targetPath = "C:\Path\to\Target.exe"
            $arguments = "-newArg"

            # Act
            Set-OHShortcut -ShortcutName $startMenuShortcutName -TargetPath $targetPath -Arguments $arguments -UpdateLocation StartMenu -UseLoggedInUsersStartMenu

            # Assert
            $shortcut = (New-Object -ComObject WScript.Shell).CreateShortcut($startMenuShortcutPath)
            $shortcut.Arguments | Should -Be $arguments
        }

        It "should update working directory arguments on the desktop" {
            # Arrange
            $desktopShortcutName = "Shortcut1"
            $targetPath = "C:\Path\to\Target.exe"
            $WorkingDir = "c:\ohtemp"

            # Act
            Set-OHShortcut -ShortcutName $desktopShortcutName -TargetPath $targetPath -WorkingDirectory $WorkingDir -UpdateLocation Desktop -UseLoggedInUsersDesktop

            # Assert
            $shortcut = (New-Object -ComObject WScript.Shell).CreateShortcut($desktopShortcutPath)
            $shortcut.workingDirectory | Should -Be $WorkingDir
        }

        It "should update working directory arguments in the start menu" {
            # Arrange
            $desktopShortcutName = "Shortcut1"
            $targetPath = "C:\Path\to\Target.exe"
            $WorkingDir = "c:\ohtemp"

            # Act
            Set-OHShortcut -ShortcutName $startmenuShortcutName -TargetPath $targetPath -WorkingDirectory $WorkingDir -UpdateLocation StartMenu -UseLoggedInUsersStartMenu

            # Assert
            $shortcut = (New-Object -ComObject WScript.Shell).CreateShortcut($startmenuShortcutPath)
            $shortcut.workingDirectory | Should -Be $WorkingDir
        }
        
    }
}
