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
  $Book.Worksheets.Add([System.Reflection.Missing]::Value, $Book.Sheets(2)) | Out-Null

  # 追加したシート名を「目次」に変更する
  $book.WorkSheets.item(3).name = "目次"

  # 目次の行数カウンタ
  $wroteCount = 2

  # 目次作成（繰り返し）
  for($i = 4;$i -le $book.worksheets.count;$i++){

    # 大見出しをコピペ
    $book.sheets(3).cells.item($wroteCount,2) = $book.sheets($i).cells.item(2,2)

    $wroteCount++

    # 大見出しを太字にする
    $book.sheets(3).range("B1:B1000").font.bold = $true

    # 小見出しをコピペ（繰り返し）
    for($j = 1;$j -le 100;$j++){
      if($book.sheets($i).cells.item($j,3).text -match '^[0-9]{1,2}-[0-9]{1,2}'){
        $book.sheets(3).cells.item($wroteCount,3) = $book.sheets($i).cells.item($j,3)
        $wroteCount++
      }
      if($book.sheets($i).cells.item($j,4).text -match '^[0-9]{1,2}-[0-9]{1,2}'){
        $book.sheets(3).cells.item($wroteCount,4) = $book.sheets($i).cells.item($j,3)
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
