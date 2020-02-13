Param(
    [Parameter(Mandatory = $true, HelpMessage = 'Octopus URL')][string] $OctopusUrl,    
    [Parameter(Mandatory = $true, HelpMessage = 'User API Key')][string] $UserApiKey,
    [Parameter(Mandatory = $true, HelpMessage = 'Variable Set Names')][string[]] $VariableSetNames
)

$excelFile = "$env:TEMP\Export.xlsx"
Write-Verbose -Verbose -Message "Save location: $excelFile"
Remove-Item $excelFile -ErrorAction Ignore

# Add Octopus.Client .NET library
$path = Join-Path (Get-Item ((Get-Package Octopus.Client).source)).Directory.FullName 'lib/net45/Octopus.Client.dll'
Add-Type -Path $path

function GetVariableSet {
    Param (
        [Parameter(Mandatory = $true)][string] $OctopusUrl,
        [Parameter(Mandatory = $true)][string] $UserApiKey,
        [Parameter(Mandatory = $true)][string] $VariableSetName
    )
    Process {

        $export = @()

        $endpoint = New-Object Octopus.Client.OctopusServerEndpoint($OctopusUrl, $UserApiKey)
        $repository = New-Object Octopus.Client.OctopusRepository($endpoint)

        $library = $repository.LibraryVariableSets.FindByName($VariableSetName)
        $variableset = $repository.VariableSets.Get($library.VariableSetId);
        
        $variableset.Variables | ForEach-Object {

            $environments = @()

            $_.Scope.Values | ForEach-Object {
                $environment = $repository.Environments.Get($_)
                $environments += $environment.Name
            }

            $export += [PSCustomObject]@{
                Name  = $_.Name;
                Value = $_.Value;
                Scope = $environments -join ','
            }

        }

        return $export
        
    }
}

$VariableSetNames | ForEach-Object {
    GetVariableSet -OctopusUrl $OctopusUrl -UserApiKey $UserApiKey -VariableSetName $_ |
    Export-Excel $excelFile -WorksheetName $_ -AutoSize
}