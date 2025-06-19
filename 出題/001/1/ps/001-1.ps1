$FILE_PATH = "C:\【VBA】001_スケジュールへの着色.xlsm"

# Excel COMオブジェクト作成
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true

try {
    $workbook = $excel.Workbooks.Open($FILE_PATH)
} catch {
    Write-Host "ファイルを開けませんでした" -ForegroundColor Red
    # exit 1
}

# セル塗りつぶしの色を取得
$targetColor = $workbook.Sheets.Item("設定").Range("D4").Interior.Color

$workbook.Sheets.Item("スケジュール").activate()

$standardRange = $workbook.Sheets.Item("スケジュール").Range("B3")

$offsetCol = 1
$cellHeader = $standardRange.Offset(0, $offsetCol)

# ヘッダーが空白ではないor結合セルである場合は継続する
while (-not [System.String]::IsNullOrEmpty($cellHeader.Value2()) -or 
       $cellHeader.MergeCells) {

    # 最初のデータ(ヘッダの次の行)セルを取得する
    $offsetRow=1
    $cell = $cellHeader.Offset($offsetRow, 0)

    # 日付列が空白であればループを中止する
    while (-not [System.String]::IsNullOrEmpty($standardRange.Offset($offsetRow, 0).Value2())) {
        # 現在参照しているセルが数値であればセルに色を塗る
        Write-Host ($offsetRow + 3)
        if([int]::TryParse($cell.Value2(),[ref]$null)){
            $cell.Interior.Color = $targetColor
            $standardRange.Offset($offsetRow, 0).Interior.Color = $targetColor

            $cellHeader.Address()
            if ($cellHeader.MergeCells) {
                $cellHeader.MergeArea.Cells.Item(1, 1).Interior.Color = $targetColor
            } else {
                $cellHeader.Interior.Color = $targetColor
            }

            # 次の列に移動するためにbreakする
            Write-Host $cell.Address()
            break;
        }
        $offsetRow++
        $cell = $cellHeader.Offset($offsetRow, 0)
    }

    # 次のセル参照を代入
    # offset指定すると結合セルの場合は１つのセルとみなされるため
    # offset指定すると結合セルの最初の列のみとなってしまう
    # $cellHeader = $cellHeader.Offset(0,1)
    $offsetCol++
    $cellHeader = $standardRange.Offset(0, $offsetCol)
}

# 終了処理
$workbook.Save()
$workbook.Close()
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
