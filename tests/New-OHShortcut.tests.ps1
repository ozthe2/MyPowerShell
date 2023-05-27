Describe "New-OHShortcut" {
    BeforeAll {
        . "$PSScriptRoot\..\Shortcuts\New-OHshortcut.ps1"       
        $shortcutName = "TestShortcut"
        $targetPath = "C:\windows\system32\notepad.exe"

        $shortcutPathLI = [System.IO.Path]::Combine([Environment]::GetFolderPath("Desktop"), "$shortcutName.lnk")
        $shortcutPathAU = [System.IO.Path]::Combine([Environment]::GetFolderPath("CommonDesktopDirectory"), "$shortcutName.lnk")
        $startMenuPathLI = "$env:APPDATA\Microsoft\Windows\Start Menu\$shortcutname.lnk"
        $startMenuPathAU = [System.IO.Path]::Combine([Environment]::GetFolderPath("CommonPrograms"), "$shortcutName.lnk")

        # Test cases for parameters with null or empty values
        $nullEmptyTestCases = @(
            @{ ArgumentValue = $null
               IconValue = $null
            },
            @{ ArgumentValue = ""
               IconValue = ""
            }
        )
    }

    AfterEach {
        Remove-Item $shortcutPathLI, $shortcutPathAU, $startMenuPathLI, $startMenuPathAU -Force -ErrorAction SilentlyContinue -Filter "*.lnk"
    }

    Context "Creating shortcut on the desktop" {
        It "Should create a shortcut on the logged-in user's desktop" {
            New-OHShortcut -ShortcutName $shortcutName -TargetPath $targetPath -CreateLocation Desktop -UseLoggedInUsersDesktop
            $shortcutExists = Test-Path $shortcutPathLI
            $shortcutExists | Should -Be $true

            # Should not create shortcut in logged in users start menu
            $startMenuExists = Test-Path $startMenuPathLI
            $startMenuExists | Should -Be $false -Because "it should not have been created"

            # Should not create shortcut in all users start menu
            $startMenuExists = Test-Path $startMenuPathAU
            $startMenuExists | Should -Be $false -Because "it should not have been created"
        }        

        It "Should create a shortcut on the all users desktop" {
            New-OHShortcut -ShortcutName $shortcutName -TargetPath $targetPath -CreateLocation Desktop
            $shortcutExists = Test-Path $shortcutPathAU
            $shortcutExists | Should -Be $true

            # Should not create shortcut in logged in users start menu
            $startMenuExists = Test-Path $startMenuPathLI
            $startMenuExists | Should -Be $false -Because "it should not have been created"

            # Should not create shortcut in all users start menu
            $startMenuExists = Test-Path $startMenuPathAU
            $startMenuExists | Should -Be $false -Because "it should not have been created"
        }       
    }

    Context "Creating shortcut on the start menu" {
        It "Should create a shortcut on the logged-in user's start menu" {
            New-OHShortcut -ShortcutName $shortcutName -TargetPath $targetPath -CreateLocation startmenu -UseLoggedInUsersStartMenu
            $shortcutExists = Test-Path $startMenuPathLI
            $shortcutExists | Should -Be $true

            # Should not create a shortcut on the logged-in user's desktop
            $startMenuExists = Test-Path $shortcutPathLI
            $startMenuExists | Should -Be $false -Because "it should not have been created"

            # Should not create a shortcut on the all users desktop
            $startMenuExists = Test-Path $shortcutPathAU
            $startMenuExists | Should -Be $false -Because "it should not have been created"
        }        

        It "Should create a shortcut on the all users start menu" {
            New-OHShortcut -ShortcutName $shortcutName -TargetPath $targetPath -CreateLocation startmenu
            $shortcutExists = Test-Path $startMenuPathAU
            $shortcutExists
            $shortcutExists | Should -Be $true

            # Should not create a shortcut on the logged-in user's desktop
            $startMenuExists = Test-Path $shortcutPathLI
            $startMenuExists | Should -Be $false -Because "it should not have been created"

            # Should not create a shortcut on the all users desktop
            $startMenuExists = Test-Path $shortcutPathAU
            $startMenuExists | Should -Be $false -Because "it should not have been created"
        }        
    }

    Context "Creating shortcut on both desktop and start menu" {
        It "Should create a shortcut on the logged-in users desktop and start menu" {
            New-OHShortcut -ShortcutName $shortcutName -TargetPath $targetPath -CreateLocation Both -UseLoggedInUsersDesktop -UseLoggedInUsersStartMenu
            $shortcutExists = Test-Path $shortcutPathLI
            $shortcutExists | Should -Be $true
            $startMenuExists = Test-Path $startMenuPathLI
            $startMenuExists | Should -Be $true
        }

        It "Should create a shortcut on the all users desktop and start menu" {
            New-OHShortcut -ShortcutName $shortcutName -TargetPath $targetPath -CreateLocation Both
            $shortcutExists = Test-Path $shortcutPathAU
            $shortcutExists | Should -Be $true
            $startMenuExists = Test-Path $startMenuPathAU
            $startMenuExists | Should -Be $true
        }
    }

    Context "Parameters with null or empty values" {
        It "Should still create the shortcut if -argument or -icon is Null or empty" -TestCases $nullEmptyTestCases {
            
            param ($arguments,$icon)
        
            New-OHShortcut -ShortcutName $shortcutName -TargetPath $targetPath -CreateLocation startmenu -arguments $arguments -icon $icon
            $shortcutExists = Test-Path $startMenuPathAU
            $shortcutExists
            $shortcutExists | Should -Be $true -because "the shortcut should still be created"

        }
    }
}
