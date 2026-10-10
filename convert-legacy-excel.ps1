param([string]$InputPath,[string]$OutputPath)
$ErrorActionPreference='Stop'
$excel=$null
$book=$null
try {
  $excel=New-Object -ComObject Excel.Application
  $excel.Visible=$false
  $excel.DisplayAlerts=$false
  $excel.AutomationSecurity=3
  $excel.AskToUpdateLinks=$false
  $book=$excel.Workbooks.Open($InputPath,0,$true)
  $book.SaveAs($OutputPath,51)
} finally {
  if($book){$book.Close($false);[Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)|Out-Null}
  if($excel){$excel.Quit();[Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)|Out-Null}
}
