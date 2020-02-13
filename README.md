# Octopus Variable Set Export

Export Variable Sets from Octopus to an Excel file.

## Dependencies

* [ImportExcel](https://www.powershellgallery.com/packages/ImportExcel#install-item)
* [Octopus.Client](https://octopus.com/docs/octopus-rest-api/octopus.client#Octopus.Client-Gettingstarted)

## Usage

```
.\OctopusVariableSetExport.ps1 -OctopusUrl 'https://localhost' -UserApiKey 'your-api-key-here' -VariableSetNames 'variable-set-1', 'variable-set-2'
```
