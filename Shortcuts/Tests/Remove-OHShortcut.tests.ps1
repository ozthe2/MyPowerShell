Describe "Remove-OHShortcut" {
    BeforeAll {
        # Set the values of the variables once before all test cases
        . "$PSScriptRoot\..\Shortcuts\Remove-OHShortcut.ps1"
        . "$PSScriptRoot\..\Shortcuts\New-OHshortcut.ps1"
        $shortcutName = "TestShortcut"
        $targetPath = "C:\windows\system32\notepad.exe"

        # Define the paths of the desktop shortcuts to be deleted
        $shortcutPathLI = [System.IO.Path]::Combine([Environment]::GetFolderPath("Desktop"), "$shortcutName.lnk") # Logged in user (LI)
        $shortcutPathAU = [System.IO.Path]::Combine([Environment]::GetFolderPath("CommonDesktopDirectory"), "$shortcutName.lnk") # All users (AU)

        # Define the paths of the start menu shortcuts to be deleted
        $startMenuPathLI = "$env:APPDATA\Microsoft\Windows\Start Menu\$shortcutname.lnk" # Logged in user (LI)
        $startMenuPathAU = [System.IO.Path]::Combine([Environment]::GetFolderPath("CommonPrograms"), "$shortcutName.lnk") # All users (AU)
    }

    AfterEach {
        # Clean up: Delete the shortcuts if they exist
        Remove-Item $shortcutPathLI, $shortcutPathAU, $startMenuPathLI, $startMenuPathAU -Force -ErrorAction SilentlyContinue -Filter "*.lnk"
    }

    Context "Deleting shortcut" {
        It "Should delete the shortcut from the logged-in users desktop" {
            # Create the shortcut first
            New-OHShortcut -ShortcutName $shortcutName -TargetPath $targetPath -CreateLocation Desktop -UseLoggedInUsersDesktop

            # Ensure the shortcut exists before deleting
            $shortcutExists = Test-Path $shortcutPathLI
            $shortcutExists | Should -Be $true -Because "the shortcut should exist before deleting"

            # Delete the shortcut
            Remove-OHShortcut -ShortcutName $shortcutName -DeleteLocation Desktop -UseLoggedInUsersDesktop

            # Verify that the shortcut has been deleted
            $shortcutExists = Test-Path $shortcutPathLI
            $shortcutExists | Should -Be $false -Because "the shortcut should have been deleted"
        }

        It "Should delete the shortcut from the all users desktop" {
            # Create the shortcut first
            New-OHShortcut -ShortcutName $shortcutName -TargetPath $targetPath -CreateLocation Desktop

            # Ensure the shortcut exists before deleting
            $shortcutExists = Test-Path $shortcutPathAU
            $shortcutExists | Should -Be $true -Because "the shortcut should exist before deleting"

            # Delete the shortcut
            Remove-OHShortcut -ShortcutName $shortcutName -DeleteLocation Desktop

            # Verify that the shortcut has been deleted
            $shortcutExists = Test-Path $shortcutPathAU
            $shortcutExists | Should -Be $false -Because "the shortcut should have been deleted"
        }

        It "Should delete the shortcut from the logged-in user's start menu" {
            # Create the shortcut first
            New-OHShortcut -ShortcutName $shortcutName -TargetPath $targetPath -CreateLocation StartMenu -UseLoggedInUsersStartMenu

            # Ensure the shortcut exists before deleting
            $shortcutExists = Test-Path $startMenuPathLI
            $shortcutExists | Should -Be $true -Because "the shortcut should exist"

            # Delete the shortcut
            Remove-OHShortcut -ShortcutName $shortcutName -DeleteLocation StartMenu -UseLoggedInUsersStartMenu

            # Verify that the shortcut has been deleted
            $shortcutExists = Test-Path $startMenuPathLI
            $shortcutExists | Should -Be $false -Because "the shortcut should have been deleted"
        }

        It "Should delete the shortcut from the all users start menu" {
            # Create the shortcut first
            New-OHShortcut -ShortcutName $shortcutName -TargetPath $targetPath -CreateLocation StartMenu

            # Ensure the shortcut exists before deleting
            $shortcutExists = Test-Path $startMenuPathAU
            $shortcutExists | Should -Be $true -Because "the shortcut should exist before deleting"

            # Delete the shortcut
            Remove-OHShortcut -ShortcutName $shortcutName -DeleteLocation StartMenu

            # Verify that the shortcut has been deleted
            $shortcutExists = Test-Path $startMenuPathAU
            $shortcutExists | Should -Be $false -Because "the shortcut should have been deleted"
        }

        It "Should delete the shortcut from both the logged-in user's desktop and start menu" {
            # Create the shortcut first
            New-OHShortcut -ShortcutName $shortcutName -TargetPath $targetPath -CreateLocation Both -UseLoggedInUsersDesktop -UseLoggedInUsersStartMenu

            # Ensure the shortcuts exist before deleting
            $shortcutExists = Test-Path $shortcutPathLI
            $shortcutExists | Should -Be $true -Because "the shortcut on the desktop should exist before deleting"

            $startMenuExists = Test-Path $startMenuPathLI
            $startMenuExists | Should -Be $true -Because "the shortcut on the start menu should exist before deleting"

            # Delete the shortcuts
            Remove-OHShortcut -ShortcutName $shortcutName -DeleteLocation Both -UseLoggedInUsersDesktop -UseLoggedInUsersStartMenu

            # Verify that the shortcuts have been deleted
            $shortcutExists = Test-Path $shortcutPathLI
            $shortcutExists | Should -Be $false -Because "the shortcut on the desktop should have been deleted"

            $startMenuExists = Test-Path $startMenuPathLI
            $startMenuExists | Should -Be $false -Because "the shortcut on the start menu should have been deleted"
        }

        It "Should delete the shortcut from both the all users desktop and start menu" {
            # Create the shortcut first
            New-OHShortcut -ShortcutName $shortcutName -TargetPath $targetPath -CreateLocation Both

            # Ensure the shortcuts exist before deleting
            $shortcutExists = Test-Path $shortcutPathAU
            $shortcutExists | Should -Be $true -Because "the shortcut on the all users desktop should exist before deleting"

            $startMenuExists = Test-Path $startMenuPathAU
            $startMenuExists | Should -Be $true -Because "the shortcut on the all users start menu should exist before deleting"

            # Delete the shortcuts
            Remove-OHShortcut -ShortcutName $shortcutName -DeleteLocation Both

            # Verify that the shortcuts have been deleted
            $shortcutExists = Test-Path $shortcutPathAU
            $shortcutExists | Should -Be $false -Because "the shortcut on the all users desktop should have been deleted"

            $startMenuExists = Test-Path $startMenuPathAU
            $startMenuExists | Should -Be $false -Because "the shortcut on the all users start menu should have been deleted"
        }
    }
}
