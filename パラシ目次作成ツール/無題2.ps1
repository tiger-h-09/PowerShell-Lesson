$excel = New-Object -ComObject Excel.Application

# Excelを見えるようにする
$excel.visible = $true

$book = $excel.workbooks.Open("C:\Users\bvs20005\Documents\03_基盤講習\PowerShell\パラシもどき_ちゃんとそろってる.xlsx")  



$Book.Worksheets.Add([System.Reflection.Missing]::Value, $Book.Sheets(2))

$book.WorkSheets.item(3).name = "目次"