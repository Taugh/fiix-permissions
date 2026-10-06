<#
.SYNOPSIS
  Rebuilds the Fiix permission workbooks from the files in this folder.

.DESCRIPTION
  1. Runs the Python builder into a temporary folder (Python cannot always write straight into OneDrive).
  2. Opens the result in Excel, recalculates every formula, and stops if any cell shows an error.
  3. Saves the finished workbook next to this script.

  Needs Python 3 with the openpyxl package, and desktop Excel.

.EXAMPLE
  .\rebuild.ps1                  # rebuild both workbooks
  .\rebuild.ps1 -What groups     # only Fiix_v5_Group_Build_Sheet.xlsx
  .\rebuild.ps1 -What reference  # only Fiix_Permissions_Reference.xlsx

  Close the workbooks in Excel first, or the save will fail.
#>
param([ValidateSet('all', 'groups', 'reference')][string]$What = 'all')

$ErrorActionPreference = 'Stop'
$Here = $PSScriptRoot
$Mgmt = Join-Path $Here 'source\IREN_Fiix_Enterprise_Permission_Model_Updated_2026-10-02.xlsx'   # change when management sends a new version
$Tmp = Join-Path $env:TEMP 'fiix-permissions-build'
New-Item -ItemType Directory -Force -Path $Tmp | Out-Null
$env:PYTHONDONTWRITEBYTECODE = '1'
$env:PYTHONIOENCODING = 'utf-8'

function Build-Workbook([string]$Script, [string]$OutName, [string[]]$ExtraArgs) {
    $tmpOut = Join-Path $Tmp $OutName
    $dest = Join-Path $Here $OutName
    if (Test-Path $tmpOut) { Remove-Item $tmpOut -Force }

    Write-Host "Building $OutName ..."
    & python (Join-Path $Here $Script) $tmpOut @ExtraArgs
    if ($LASTEXITCODE -ne 0 -or -not (Test-Path $tmpOut)) { throw "$Script failed." }

    $xl = $null
    try {
        $xl = New-Object -ComObject Excel.Application
        $xl.Visible = $false
        $xl.DisplayAlerts = $false
        $wb = $xl.Workbooks.Open($tmpOut)
        $xl.CalculateFull()
        $errors = 0
        foreach ($ws in $wb.Worksheets) {
            try { $errors += $ws.UsedRange.SpecialCells(-4123, 16).Count } catch { }   # formulas that evaluate to an error
        }
        if ($errors -gt 0) { $wb.Close($false); throw "$OutName has $errors formula errors. Not saved." }
        $wb.Worksheets.Item(1).Activate()
        $wb.Save()
        $wb.SaveCopyAs($dest)
        $wb.Close($false)
        Write-Host "  saved $dest (no formula errors)"
    }
    finally {
        if ($xl) { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
    }
}

if ($What -in 'all', 'reference') {
    Build-Workbook -Script 'build_workbook.py' -OutName 'Fiix_Permissions_Reference.xlsx' -ExtraArgs @()
}
if ($What -in 'all', 'groups') {
    if (-not (Test-Path $Mgmt)) { throw "Management workbook not found: $Mgmt" }
    Build-Workbook -Script 'build_group_design.py' -OutName 'Fiix_v5_Group_Build_Sheet.xlsx' -ExtraArgs @($Mgmt)
}
Write-Host 'Done.'
