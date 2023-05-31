BeforeAll {
    . ".\Remove-OHRegistryKey.ps1"
}

Describe "Remove-OHRegistryKey Tests" {
    # Mock the ShouldProcess function to simulate user confirmation
    Mock ShouldProcess { $true }

    Context "When the registry value exists" {
        BeforeEach {
            # Create a test registry key and value for the context
            $testKeyPath = "HKCU:\Software\Test"
            $testValueName = "TestValue"
            $testValueData = "TestData"
            New-Item -Path $testKeyPath -Force | Out-Null
            New-ItemProperty -Path $testKeyPath -Name $testValueName -Value $testValueData -PropertyType String -Force | Out-Null
        }

        AfterEach {
            # Remove the test registry key and value after each test
            Remove-Item -Path $testKeyPath -Recurse -Force | Out-Null
        }

        It "Should remove the registry value and return true" {
            $result = Remove-OHRegistryKey -KeyPath $testKeyPath -ValueName $testValueName
            $result | Should -Be $true
        }

        It "Should remove the registry value from the specified key" {
            Remove-OHRegistryKey -KeyPath $testKeyPath -ValueName $testValueName
            $valueExists = (Get-ItemProperty -Path $testKeyPath -Name $testValueName -ErrorAction SilentlyContinue)
            $valueExists | Should -Be $null
        }
    }

    Context "When the registry value does not exist" {
        BeforeEach {
            # Create a test registry key without the value for the context
            $testKeyPath = "HKCU:\Software\Test"
            New-Item -Path $testKeyPath -Force | Out-Null
        }

        AfterEach {
            # Remove the test registry key after each test
            Remove-Item -Path $testKeyPath -Recurse -Force | Out-Null
        }

        It "Should return false" {
            $result = Remove-OHRegistryKey -KeyPath $testKeyPath -ValueName "NonExistentValue"
            $result | Should -Be $false
        }
    }

    Context "When the registry key does not exist" {
        It "Should return false" {
            $result = Remove-OHRegistryKey -KeyPath "HKCU:\Software\NonExistentKey" -ValueName "NonExistentValue"
            $result | Should -Be $false
        }
    }
}