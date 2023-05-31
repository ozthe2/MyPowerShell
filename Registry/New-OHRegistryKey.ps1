function New-OHRegistryKey {
    <#
    .SYNOPSIS
    Creates a new registry key, value, and data if they do not already exist.

    .DESCRIPTION
    The New-OHRegistryKey function creates a new registry key, value, and data if they do not already exist. It checks for the existence of the registry key and value and creates them if necessary.

    .PARAMETER KeyPath
    Specifies the registry key path where the value will be created.

    .PARAMETER ValueName
    Specifies the name of the registry value.

    .PARAMETER ValueData
    Specifies the data for the registry value.

    .PARAMETER ValueType
    Specifies the data type of the registry value. Valid values are 'String' or 'DWord'.

    .OUTPUTS
    System.Boolean
    The function returns $true if the registry key or value is created successfully. It returns $false if the registry key and value already exist, and no changes are made to the registry.

    .EXAMPLE
    New-OHRegistryKey -KeyPath "HKCU:\Software\MyApp" -ValueName "Setting1" -ValueData "Value1" -ValueType "String"
    Creates a new registry key "HKCU:\Software\MyApp" if it does not exist, and then creates a new string registry value "Setting1" with the data "Value1". Returns $true if the registry key or value is created.

    .EXAMPLE
    New-OHRegistryKey -KeyPath "HKLM:\Software\MyApp" -ValueName "Setting2" -ValueData "123" -ValueType "DWord"
    Creates a new registry key "HKLM:\Software\MyApp" if it does not exist, and then creates a new DWORD registry value "Setting2" with the data 123. Returns $true if the registry key or value is created.    

    .NOTES
        Author: owen.heaume
        Date: 31-May-2023
        Version:
            1.0 - Initial script creation
    #>

    [CmdletBinding(SupportsShouldProcess = $true)]
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
            Write-Verbose "$KeyPath does not exist so creating"
            # Create the new registry key
            New-Item -Path $KeyPath -Force | Out-Null
            
        }
        
        # Check if the registry value exists
        $valueExists = Get-ItemProperty -Path $KeyPath -Name $ValueName -ErrorAction SilentlyContinue
        
        if (-not $valueExists) {
            write-verbose "$ValueName not found so creating"
            # Create the new registry value with data and type
            if ($ValueType -eq 'String') {
                New-ItemProperty -Path $KeyPath -Name $ValueName -Value $ValueData -PropertyType String -Force | Out-Null
            } elseif ($ValueType -eq 'DWord') {
                New-ItemProperty -Path $KeyPath -Name $ValueName -Value ([int]$ValueData) -PropertyType DWord -Force | Out-Null
            }
            
            return $true
        }
        write-verbose "$keypath and $ValueName ($ValueType) = $ValueData already exist - no changes to the registry are required"
        return $false
    } catch {
        Write-Error "Failed to create registry key or value: $_"
    }
}