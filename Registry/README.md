## New-OHRegistryKey

- Create a new registry key, value, and data if they do not already exist.
- Checks for the existence of the registry key and value and creates them if necessary.
- Supports creating registry keys and values with different data types.

### Syntax

```powershell
New-OHRegistryKey -KeyPath <String> -ValueName <String> -ValueData <String> -ValueType <String>
