BeforeAll {
    . "$PSScriptRoot\..\Registry\New-OHRegistryKey.ps1"
}

Describe "New-OHRegistryKey" {
    BeforeEach {
        $keyPath = "HKCU:\Software\MyCompany"
        $valueName = "MyValue"
        $valueData = "Hello, world!"
        $valueType = "string"        
    }

    AfterEach {
        Remove-Item -Path $keyPath
    }

    it "should create a new registry key with a STRING value" {
        # Arrange
        $valueType = "string"

        # Act
        $result = New-OHRegistryKey -KeyPath $keyPath -ValueName $valueName -ValueData $valueData -ValueType $valueType

        # Assert
        $result | Should -Be $true

        (Get-ItemProperty -Path $keyPath -Name $valueName).myvalue | Should -Be $valueData

        $property = (Get-ItemProperty -Path $keyPath -Name $valueName).myValue
        $property.GetType().Name | Should -Be 'String'
    }    

    it "should create a new registry key with a DWORD value" {
        # Arrange
        $valueData = '1'   
        $valueType = "dword"

        # Act
        $result = New-OHRegistryKey -KeyPath $keyPath -ValueName $valueName -ValueData $valueData -ValueType $valueType

        # Assert
        $result | Should -Be $true

        (Get-ItemProperty -Path $keyPath -Name $valueName).myvalue | Should -Be $valueData

        $property = (Get-ItemProperty -Path $keyPath -Name $valueName).myValue
        $property.GetType().Name | Should -Be 'Int32'
    }
  
    it "should return false if the registry path and value already exists" {
        # Arrange
        New-Item -Path "HKCU:\Software\MyCompany"
        New-ItemProperty -Path "HKCU:\Software\MyCompany" -Name "MyValue" -Value "Hello, world!"

        # Act
        $result = New-OHRegistryKey -KeyPath "HKCU:\Software\MyCompany" -ValueName "MyValue" -ValueData "Hello, world!" -ValueType $valuetype

        # Assert
        $result | Should -Be $false
    }
}