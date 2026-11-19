# Build-Workbook.ps1
# Turns Performance_Tracker.xlsx + vba\*.bas into Performance_Tracker.xlsm with the
# macros imported and the buttons installed. Run from the folder that holds this
# script (double-click Build-Workbook.cmd, or:
#   powershell -ExecutionPolicy Bypass -File .\Build-Workbook.ps1)
param(
    [string]$Source = "Performance_Tracker.xlsx",
    [string]$Target = "Performance_Tracker.xlsm"
)
$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$src  = Join-Path $here $Source
$dst  = Join-Path $here $Target
$vba  = Join-Path $here "vba"

if (-not (Test-Path $src)) { throw "Not found: $src" }
$modules = Get-ChildItem -Path $vba -Filter *.bas
if ($modules.Count -lt 8) { throw "Expected eight .bas files in $vba" }

# Excel only lets a script import code if 'Trust access to the VBA project object
# model' is on. Offer to switch it on for this build and restore it afterwards.
$probe = New-Object -ComObject Excel.Application
$ver = $probe.Version
$probe.Quit()
[void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($probe)
$key = "HKCU:\Software\Microsoft\Office\$ver\Excel\Security"
if (-not (Test-Path $key)) { New-Item -Path $key -Force | Out-Null }
$old = (Get-ItemProperty -Path $key -Name AccessVBOM -ErrorAction SilentlyContinue).AccessVBOM
$changed = $false
if ($old -ne 1) {
    Write-Host "Excel setting 'Trust access to the VBA project object model' is OFF."
    $ans = Read-Host "Turn it on just for this build and restore it afterwards? (Y/N)"
    if ($ans -notmatch '^[Yy]') {
        Write-Host "Stopped. Import the .bas files by hand instead (see the Start sheet, Option B)."
        exit 1
    }
    Set-ItemProperty -Path $key -Name AccessVBOM -Value 1 -Type DWord
    $changed = $true
}

$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Open($src)
    foreach ($m in $modules) {
        Write-Host "Importing $($m.Name)"
        [void]$wb.VBProject.VBComponents.Import($m.FullName)
    }
    $excel.Run("'" + $wb.Name + "'!InstallButtons")
    if (Test-Path $dst) { Remove-Item $dst -Force }
    $wb.SaveAs($dst, 52)   # 52 = xlOpenXMLWorkbookMacroEnabled
    $wb.Close($false)
    Write-Host ""
    Write-Host "Built $dst"
    Write-Host "Open it, enable macros, then click '1. Setup / create database' on the Start sheet."
}
finally {
    if ($excel) {
        $excel.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    if ($changed) {
        if ($null -eq $old) { Remove-ItemProperty -Path $key -Name AccessVBOM -ErrorAction SilentlyContinue }
        else { Set-ItemProperty -Path $key -Name AccessVBOM -Value $old -Type DWord }
        Write-Host "Restored the VBA trust setting."
    }
}
