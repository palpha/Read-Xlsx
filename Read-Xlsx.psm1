Set-StrictMode -Version 3

[System.IO.Directory]::SetCurrentDirectory($PSScriptRoot)
[System.Reflection.Assembly]::LoadFile("$PSScriptRoot\DocumentFormat.OpenXml.dll")
[System.Reflection.Assembly]::LoadFile("$PSScriptRoot\Panagora.Office.dll")

function Is-Xslx {
    param ([System.IO.FileInfo] $File)
    $File.Extension.ToLowerInvariant() -eq '.xlsx'
}

function Create-Reader {
    param ([string] $FilePath)
    New-Object Panagora.Office.ExcelReader($FilePath)
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
        [string] $SheetName = 'Sheet1'
    )

    process {
        $FilePath | ? { Is-Xslx $_ } | % {
            Write-Host $_
            $reader = Create-Reader $_.FullName
            Write-Host $reader
            $colNames = @()
            $reader.ReadSheet($SheetName) | % {
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
        }
    }
}

Export-ModuleMember -Function Read-Xlsx