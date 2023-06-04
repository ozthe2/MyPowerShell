## New-OHRegistryKey

- Create a new registry key, value, and data if they do not already exist.
- Checks for the existence of the registry key and value and creates them if necessary.
- Supports creating registry keys and values with different data types.

### Syntax

```powershell
New-OHRegistryKey -KeyPath <String> -ValueName <String> -ValueData <String> -ValueType <String>
```

## Set-OHRegistryKey

- Sets the value of a registry key in the Windows Registry.
- Checks if the specified registry key and value exist before modifying them.
- Supports modifying registry values with different data types.

### Syntax

```powershell
Set-OHRegistryKey -KeyPath <String> -ValueName <String> -ValueData <String> -ValueType <String>
```

## Remove-OHRegistryKey

- Removes a registry value from a specified key.
- Returns `true` if the deletion is successful, or `false` if the value does not exist.

### Syntax

```powershell
Remove-OHRegistryKey -KeyPath <String> -ValueName <String>
```
