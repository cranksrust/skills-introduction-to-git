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
    # Workbook_Open fires reliably for add-ins loaded at startup; Auto_Open does not always
    $code = "Private Sub Workbook_Open()`r`n    Auto_Open`r`nEnd Sub`r`n`r`nPrivate Sub Workbook_BeforeClose(Cancel As Boolean)`r`n    Auto_Close`r`nEnd Sub`r`n"
    $wb.VBProject.VBComponents.Item($wb.CodeName).CodeModule.AddFromString($code)
    if (Test-Path $out) { Remove-Item $out }
    $wb.SaveAs($out, 55)   # 55 = xlOpenXMLAddIn
    $wb.Close($false)
    Write-Host "Built $out"
}
finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
