Set-StrictMode -Version 3

[System.Reflection.Assembly]::LoadFile("$PSScriptRoot\WindowsBase.dll")
[System.Reflection.Assembly]::LoadFile("$PSScriptRoot\DocumentFormat.OpenXml.dll")

Add-Type -Path .\ExcelReader.cs -ReferencedAssemblies "$(pwd)\DocumentFormat.OpenXml.dll", "$(pwd)\WindowsBase.dll"

function Is-Xslx {
    param ([System.IO.FileInfo] $File)
    $File.Extension.ToLowerInvariant() -eq '.xlsx'
}

function Create-Reader {
    param ([string] $FilePath)
    New-Object ExcelReader($FilePath)
}

function Read-Xlsx {
    [CmdletBinding()]
    param (
        [Parameter(
            Mandatory,
            ValueFromPipeline,
            ValueFromPipelineByPropertyName)]
        [Alias('FullName')]
        [System.IO.FileInfo[]] $FilePath,
        $Sheet = 0 # name or idx
    )

    process {
        $FilePath | ? { Is-Xslx $_ } | % {
            try {
                $reader = Create-Reader $_.FullName
                $colNames = @()
                $reader.ReadSheet($Sheet) | % {
                    $row = $_
                    if ($_.Get('A').Row -eq 1) {
                        $row | % { $colNames += @{ Column = $_.Column; Name = $_.Value } }
                        $null
                    } else {
                        $rowConverted = @{}
                        $colNames | % {
                            $rowConverted.($_.Name) = $row.Get($_.Column)
                        }
                        [PSCustomObject] $rowConverted
                    }
                } | ? { $_ -ne $null }
            } finally {
                $reader.Dispose()
            }
        }
    }
}

function Get-XlsxInfo {
    [CmdletBinding()]
    param (
        [Parameter(
            Mandatory,
            ValueFromPipeline,
            ValueFromPipelineByPropertyName)]
        [Alias('FullName')]
        [System.IO.FileInfo[]] $FilePath
    )

    process {
        $FilePath | ? { Is-Xslx $_ } | % {
            try {
                $reader = Create-Reader $_.FullName
                [PSCustomObject] @{
                    FilePath = $_.FullName
                    Sheets = $reader.Sheets | % {
                        [PSCustomObject] @{
                            Name = $_.Name
                            Columns = $reader.ReadSheet($_.Name) | select -First 1 | % {
                                $_ | % { $_.Value }
                            }
                        }
                    }
                }
            } finally {
                $reader.Dispose()
            }
        }
    }
}

Export-ModuleMember -Function Read-Xlsx
Export-ModuleMember -Function Get-XlsxInfo