Read-Xlsx
=========

Description
-----------
A simple XLSX file reader for PowerShell. It might not work with complex Excel documents, but it should be sufficient for normal ETL purposes.

Prerequisites
-------------
PowerShell v3 is needed, simply because I haven't bothered to test
it with v2.

Typical installation
--------------------
Clone in a directory in your module path (see $env:PSModulePath).

Usage
-----
    Import-Module Read-Xlsx
    help Read-Xlsx -Full
    help Get-XlsxInfo -Full

Read-Xlsx will produce a PSCustomObject per row, with property names based on the first row of the sheet.

Get-XlsxInfo will give you sheet and column names.

Both functions accept pipeline input (by value and by property name).