function Modify-OHRegistry {
<#
.SYNOPSIS
Modifies the Windows registry by creating, replacing or deleting a specified registry key value.

.DESCRIPTION
This function modifies a registry key value by creating a new one, replacing an existing one or deleting a specific one.
The function takes five parameters: the registry key path, the name of the registry key, the data type of the registry key value,
the value to be associated with the key, and the action to be performed on the registry (create, replace or delete).

.PARAMETER path
The path to the registry key (e.g. HKCU:\Software\Microsoft). This parameter is mandatory.

.PARAMETER name
The name of the registry item (e.g. '(Default)' or 'UninstallString'). This parameter is mandatory.

.PARAMETER type
The data type of the registry key value (e.g. 'DWORD' or 'STRING'). This parameter is mandatory.
Accepted values are: 'String','ExpandString','Dword','Binary' and 'MultiString'.

.PARAMETER value
The value associated with the name (e.g. 0 or 'C:\program files\myProgram\uninstall.exe'). This parameter is mandatory.

.PARAMETER Action
The action to perform on the registry: 'Create', 'Replace' or 'Delete'. This parameter is mandatory.

.NOTES
Author: OH
Date: 29-April-2023
Version: 1.0
 - Initial script creation

.EXAMPLE
Modify-OHRegistry -path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" -name "MyApp" -type "String" -value "C:\Program Files\MyApp\MyApp.exe" -Action "create"
Creates a new registry value named "MyApp" with the value "C:\Program Files\MyApp\MyApp.exe" under the registry key "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run".

.EXAMPLE
Modify-OHRegistry -path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" -name "MyApp" -type "String" -value "C:\Program Files\MyApp\MyApp.exe" -Action "replace"
Replaces the value of the existing registry value named "MyApp" with the value "C:\Program Files\MyApp\MyApp.exe" under the registry key "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run".

.EXAMPLE
Modify-OHRegistry -path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" -name "MyApp" -type "String" -value "C:\Program Files\MyApp\MyApp.exe" -Action "delete"
Deletes the registry value named "MyApp" under the registry key "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run".

#>

    [CmdletBinding(SupportsShouldProcess=$true)]
    Param (
        [Parameter(Mandatory=$true,
        HelpMessage="The path to the registry key eg HKCU:\Software\Microsoft")]
        [string]$path,

        [Parameter(Mandatory=$true,
        HelpMessage="The name of the registry item eg '(Default)' or 'UninstallString'")]
        [string]$name,

        [Parameter(Mandatory=$true,
        HelpMessage="The data type eg 'DWORD' or 'STRING'")]
        [ValidateSet("String","ExpandString","Dword","Binary","MultiString")]
        [string]$type,

        [Parameter(Mandatory=$true,
        HelpMessage="The value associated with the name eg 0 or 'C:\program files\myProgram\uninstall.exe'")]
        [string]$value,

        [Parameter(Mandatory=$true,
        HelpMessage="The action to perform on the registry: 'Create', 'Replace' or 'Delete'")]
        #[ValidateSet("Create","Replace","Delete")]
        [string]$Action
    )

    switch ($Action) {
        'create' {
            if (!(Test-Path $path)) {
                # path does not exist, so create it and the data\value
                write-warning "The path $path does not exist, so creating path, name, type and value..."
                New-Item -Path $path
                New-ItemProperty -Path $path -Name $name -PropertyType $type -Value $value
                Write-host "Created $path with name: $name and value: $value"
            } elseif ((Get-ItemProperty -Path $path -Name $name -ErrorAction SilentlyContinue) -eq $null) {
                # the path exists but there is no data \ value
                Write-warning "the path exists but there is no data \ value. Creating..."
                Set-ItemProperty -Path $path -Name $name -Value $value -type $type
                write-host "Created new name: $name with value: $value at location: $path"    
            } else {
                # the path, name and value already exists so do nothing
                Write-Host "Action set to $Action, and path and name already exist. Nothing to do, so exiting." -ForegroundColor Green
            }
        }
        'replace' {
            if (!(Test-Path $path)) {
                # path does not exist, so create it and the data\value\type
                write-warning "The path $path does not exist, so creating path, name and value..."              
                New-Item -Path $path                
                New-ItemProperty -Path $path -Name $name -PropertyType $type -Value $value
                Write-host "Created $path with name: $name and value: $value"
            } elseif ((Get-ItemProperty -Path $path -Name $name -ErrorAction SilentlyContinue) -eq $null) {
                # the path exists but there is no data \ value
                Write-warning "the path exists but there is no data \ value. Creating..."
                Set-ItemProperty -Path $path -Name $name -Value $value -type $type
                write-host "Created new name: $name with value: $value at location: $path"    
            } else {
                # the path, name and value exists but regardless of the type and value, overwrite with new values
                write-warning "Action set to: $Action, so replacing type and value at $path\$name"
                if ($PSCmdlet.ShouldProcess("type:$type and value:$value","replace")) {
                    Set-ItemProperty -Path $path -Name $name -Value $value -type $type -Force
                }
                Write-Host "Replaced type with $type and value with $value at $path::$name"
            } 
        }
        'delete' {
            if ((Get-ItemProperty -Path $path -Name $name -ErrorAction SilentlyContinue) -eq $null) {
                #  only delete the name and data, not the path (registry key). Registry 'name' does not exist, so nothing to delete
                Write-Host "Action set to: $action, but the name: $name at location: $path does not exist, nothing to delete." -ForegroundColor Green                
            } else {
                # delete the name
                write-warning "Action set to: $action, deleting name entry..."
                if ($PSCmdlet.ShouldProcess("$path\$name","delete")) {
                    Remove-ItemProperty -Path $path -Name $name -Force
                }
                write-host "Deleted name: $name at location: $path"
            }
        }
        default {write-warning "$action is not a valid action type.  The action must be one of: 'Create', 'Replace' or 'Delete'."}
    }
} 
