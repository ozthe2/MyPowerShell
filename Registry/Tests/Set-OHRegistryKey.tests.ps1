BeforeAll {
    . "$PSScriptRoot\..\..\Registry\Set-OHRegistryKey.ps1"
}

Describe "Set-OHRegistryKey" {
    BeforeEach {
        $keyPath = "HKCU:\Software\Test"
        $valueName = "TestValue"
        $valueData = "Test Data"
    }

    AfterEach {
        Remove-Item -Path $keyPath -ErrorAction SilentlyContinue
    }

    Context "When the registry key and value exist" {
        It "Should set the registry value and return $true" {            
            $valueType = "String"

            # Create the registry key and value for testing
            New-Item -Path $keyPath -Force | Out-Null
            New-ItemProperty -Path $keyPath -Name $valueName -Value "Initial Data" -PropertyType String | Out-Null

            $result = Set-OHRegistryKey -KeyPath $keyPath -ValueName $valueName -ValueData $valueData -ValueType $valueType
            $modifiedValue = Get-ItemProperty -Path $keyPath -Name $valueName

            $result | Should -Be $true
            $modifiedValue.$valueName | Should -Be $valueData
        }
    }

    Context "When the registry key does not exist" {
        It "Should return $false" {
            $keyPath = "HKCU:\Software\NonExistentKey"
            $valueType = "String"

            $result = Set-OHRegistryKey -KeyPath $keyPath -ValueName $valueName -ValueData $valueData -ValueType $valueType

            $result | Should -Be $false
        }
    }

    Context "When the registry value does not exist" {
        It "Should return $false" {           
            $valueName = "NonExistentValue"            
            $valueType = "String"

            # Create the registry key for testing
            New-Item -Path $keyPath -Force | Out-Null

            $result = Set-OHRegistryKey -KeyPath $keyPath -ValueName $valueName -ValueData $valueData -ValueType $valueType

            $result | Should -Be $false
        }
    }

    Context "When an error occurs during registry modification" {
        It "Should throw an error" {            
            $valueData = "Invalid Data"
            $valueType = "DWord"

            # Create the registry key for testing
            New-Item -Path $keyPath -Force | Out-Null
            New-ItemProperty -Path $KeyPath -Name $ValueName -Value $ValueData -PropertyType String -Force

            $errorActionPreference = "Stop"

            { Set-OHRegistryKey -KeyPath $keyPath -ValueName $valueName -ValueData $valueData -ValueType $valueType } | Should -throw
        }
    }
}