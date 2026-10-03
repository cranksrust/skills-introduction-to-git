# Builds StepwiseRegression.xlam from StepwiseRegression.bas using desktop Excel on Windows.
# One-time prerequisite: Excel > File > Options > Trust Center > Trust Center Settings >
# Macro Settings > tick "Trust access to the VBA project object model".
# Usage: powershell -ExecutionPolicy Bypass -File .\Build-Addin.ps1

$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$bas  = Join-Path $here "StepwiseRegression.bas"
$out  = Join-Path $here "StepwiseRegression.xlam"

$excel = New-Object -ComObject Excel.Application
$excel.DisplayAlerts = $false
try {
    $wb = $excel.Workbooks.Add()
    $null = $wb.VBProject.VBComponents.Import($bas)
    if (Test-Path $out) { Remove-Item $out }
    $wb.SaveAs($out, 55)   # 55 = xlOpenXMLAddIn
    $wb.Close($false)
    Write-Host "Built $out"
}
finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
