Set-StrictMode -Version 3

[System.Reflection.Assembly]::LoadFile("$PSScriptRoot\WindowsBase.dll")
[System.Reflection.Assembly]::LoadFile("$PSScriptRoot\DocumentFormat.OpenXml.dll")

Add-Type -Path $PSScriptRoot\ExcelReader.cs -ReferencedAssemblies "$PSScriptRoot\DocumentFormat.OpenXml.dll", "$PSScriptRoot\WindowsBase.dll"

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

    begin {
        $FilePath = $FilePath | % {
            if (-not [System.IO.Path]::IsPathRooted($_)) {
                "$(pwd)\$_"
            } else { $_ }
        }
    }

    process {
        $FilePath | ? { Is-Xslx $_ } | % {
            if (-not (Test-Path $_.FullName)) {
                Write-Error ("File not found: {0}" -f $_.FullName)
                return
            }

            $reader = $false
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
                if ($reader) { $reader.Dispose() }
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