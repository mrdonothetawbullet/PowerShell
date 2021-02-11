$xl = New-Object -COM "Excel.Application"
$xl.Visible = $true

$wb = $xl.Workbooks.Open("C:\path\to\your.xlsx")
$ws = $wb.Sheets.Item(1)

$rows = $ws.UsedRange.Rows.Count

foreach ( $col in "A", "B", "C" ) {
  $xl.WorksheetFunction.CountIf($ws.Range($col + "1:" + $col + $rows), "<>") - 1
}

$wb.Close()
$xl.Quit()
