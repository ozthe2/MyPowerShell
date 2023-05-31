function Set-SpecRegistryKey {
    <# 
    .SYNOPSIS
    Sets the value of a registry key in the Windows Registry.

    .DESCRIPTION
    The Set-SpecRegistryKey function allows you to modify the value of a registry key in the Windows Registry. It checks if the specified registry key and value exist before modifying them. The function supports two value types: String and DWord.

    .PARAMETER KeyPath
    Specifies the registry key path where the value should be modified. This parameter is mandatory.

    .PARAMETER ValueName
    Specifies the name of the value to be modified. This parameter is mandatory.

    .PARAMETER ValueData
    Specifies the new data to be set for the registry value. This parameter is mandatory.

    .PARAMETER ValueType
    Specifies the value type to be set for the registry value. Valid options are 'String' and 'DWord'. This parameter is mandatory.

    .OUTPUTS
    System.Boolean
    The function returns $true if the registry value is successfully modified, and $false otherwise.

    .EXAMPLE
    Set-SpecRegistryKey -KeyPath "HKCU:\Software\MyApp" -ValueName "Version" -ValueData "1.0" -ValueType "String"
    This example sets the registry value named "Version" under the "HKCU:\Software\MyApp" key to the string value "1.0".
    The function returns $true if the modification is successful.

    .EXAMPLE
    Set-SpecRegistryKey -KeyPath "HKLM:\Software\MyApp" -ValueName "Enabled" -ValueData "1" -ValueType "DWord"
    This example sets the registry value named "Enabled" under the "HKLM:\Software\MyApp" key to the DWORD value 1.
    The function returns $true if the modification is successful.

    .NOTES
    Author: owen.heaume
    Date: 31-May-2023
    Version:
        1.0 - Initial script creation
    #>

    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [String]$KeyPath,
        
        [Parameter(Mandatory = $true)]
        [String]$ValueName,
        
        [Parameter(Mandatory = $true)]
        [String]$ValueData,
        
        [Parameter(Mandatory = $true)]
        [ValidateSet('String', 'DWord')]
        [String]$ValueType
    )
    
    try {
        # Check if the registry key exists
        $keyExists = Test-Path -Path $KeyPath
    
        if (-not $keyExists) {
            Write-Verbose "$KeyPath does not exist."
            return $false
        }
        
        # Check if the registry value exists
        $valueExists = Get-ItemProperty -Path $KeyPath -Name $ValueName -ErrorAction SilentlyContinue
        
        if (-not $valueExists) {
            Write-Verbose "$ValueName not found in $KeyPath."
            return $false
        }
        
        # Modify the existing registry value with the new data and type
        if ($ValueType -eq 'String') {
            write-verbose "Setting registry to: $keypath $ValueName=$ValueData of type String"
            Set-ItemProperty -Path $KeyPath -Name $ValueName -Value $ValueData -Type String -Force | Out-Null
        }
        elseif ($ValueType -eq 'DWord') {
            write-verbose "Setting registry to: $keypath $ValueName=$ValueData of type DWord"
            Set-ItemProperty -Path $KeyPath -Name $ValueName -Value ([int]$ValueData) -Type DWord -Force | Out-Null
        }
        
        return $true
    }
    catch {
        Write-Error "Failed to modify registry value: $_"
    }
}