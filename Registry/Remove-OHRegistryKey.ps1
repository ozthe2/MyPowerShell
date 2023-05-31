function Remove-OHRegistryKey {
    <#
    .SYNOPSIS
    Removes a registry value from a specified key.

    .DESCRIPTION
    The Remove-OHRegistryKey function removes a registry value from the specified key. It returns true if the deletion is successful, or false if the value does not exist.

    .PARAMETER KeyPath
    Specifies the path of the registry key.

    .PARAMETER ValueName
    Specifies the name of the registry value to remove.

    .OUTPUTS
    System.Boolean
    Returns true if the registry value is successfully removed. Returns false if the value does not exist or if any errors occur.

    .EXAMPLE
    PS> Remove-OHRegistryKey -KeyPath "HKCU:\Software\Test" -ValueName "TestValue"
    True

    .EXAMPLE
    PS> $result = Remove-OHRegistryKey -KeyPath "HKCU:\Software\Test" -ValueName "NonExistentValue"
    $result
    False

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
        [String]$ValueName
    )

    try {
        # Check if the registry key exists
        $keyExists = Test-Path -Path $KeyPath

        if (-not $keyExists) {
            Write-Verbose "$KeyPath does not exist"
            return $false
        }

        # Check if the registry value exists
        $value = Get-ItemProperty -Path $KeyPath -Name $ValueName -ErrorAction SilentlyContinue

        if (-not $value) {
            Write-Verbose "$ValueName not found"
            return $false
        }

        # Remove the registry value
        if ($PSCmdlet.ShouldProcess("$KeyPath\$ValueName", "Remove")) {
            Remove-ItemProperty -Path $KeyPath -Name $ValueName -Force | Out-Null
            return $true
        }
    } catch {
        Write-Error "Failed to remove registry value: $_"
    }
}