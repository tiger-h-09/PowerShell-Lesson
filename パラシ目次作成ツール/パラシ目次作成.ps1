# パラシから送られてきたオブジェクトの中でパラシだけ受け取る
$parashies = $input |? name -Match "パラシもどき"
 
# Excelを開く
$excel = New-Object -ComObject Excel.Application

# Excelを見えるようにする
$excel.visible = $true

# 目次作成
foreach($parashi in $parashies){

  # パラシを開く
  $book = $excel.workbooks.Open($parashi.fullname)  

  # 3シート目に新規のシートを追加する
  #$Book.Worksheets.Add([System.Reflection.Missing]::Value, $Book.Sheets(2)) | Out-Null
  $book.Worksheets.Add($book.Sheets(3)) | Out-Null

  # 追加したシート名を「目次」に変更する
  $book.WorkSheets.item(3).name = "目次"

  # 大見出しを太字にする
    $book.sheets(3).range("B1:B1000").font.bold = $true

  # 目次の行数カウンタ
  $wroteCount = 2

  # 目次作成（繰り返し）
  for($sheetcount = 4;$sheetcount -le $book.worksheets.count;$sheetcount++){

    # 大見出しをコピペ
    $book.sheets(3).cells.item($wroteCount,2) = $book.sheets($sheetcount).cells.item(2,2)

    $wroteCount++

    # 小見出しをコピペ（繰り返し）
    for($rowCount = 1;$rowCount -le 100;$rowCount++){
      if($book.sheets($sheetcount).cells.item($rowCount,3).text -match '^[0-9]{1,2}-[0-9]{1,2}'){
        $book.sheets(3).cells.item($wroteCount,3) = $book.sheets($sheetcount).cells.item($rowCount,3)
        $wroteCount++
      }
    }
  }

  # 上書き保存
  $book.Save()
}

# Excelを終了
$Excel.Quit()

# 変数の解放
$excel = $null
$book　= $null
$parashies = $null
$parashi = $null
