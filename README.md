# Octopus Variable Set Export

Export Variable Sets from Octopus to an Excel file. Each variable set will be on its own worksheet.

## Dependencies

* [ImportExcel](https://www.powershellgallery.com/packages/ImportExcel#install-item)
* [Octopus.Client](https://octopus.com/docs/octopus-rest-api/octopus.client#Octopus.Client-Gettingstarted)

### Install Dependencies

```Install-Module -Name ImportExcel```

```Install-Package Octopus.Client -source https://www.nuget.org/api/v2 -SkipDependencies```

## Usage

```
.\OctopusVariableSetExport.ps1 -OctopusUrl 'https://localhost' -UserApiKey 'your-api-key-here' -VariableSetNames 'variable-set-1', 'variable-set-2'
```
