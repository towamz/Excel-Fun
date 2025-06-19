#　データ定義
####################
# $isUsedCols = @{}   

# key=列番号
# value=ヘッダーのセル番地/hidden/used
# (データのある列は必ずkeyが存在する)

# 数値セルが見つかったら$isUsedCols[(列番号)]のvalueで
# rangeオブジェクト取得をtryし
# 取得できた=塗りつぶし対象
# 取得できない=塗りつぶし対象外(hidden/used)

# <例>
# ヘッダーが3行目
# 列がB列の時
# $isUsedCols[2]=$B$3

# C,D,E列が結合
# $isUsedCols[3]=$C$3
# $isUsedCols[4]=$C$3
# $isUsedCols[5]=$C$3
#
# F列が非表示の場合
# $isUsedCols[6]="hidden"
# 
# B列で数値が見つかった時
# $isUsedCols[2]=$B$3 → $isUsedCols[2]="used"

####################
# $isUsedHeaders = @{}
# key=ヘッダーの列番号
# value=塗りつぶしを実施した行番号
# (塗りつぶしを実施していない列番号はkey自体が存在しない)
# 
# 数値セルが見つかったら$isUsedHeaders[(列番号)]のvalueが
# 数値セルの行番号と  一致=塗りつぶし対象
# 数値セルの行番号と不一致=塗りつぶし対象外
#
# <例>
# c列が10行目で初めて数値があったとき
# $isUsedHeaders[3]=10


# セル塗りつぶしの色を取得
function getColorsArray(){
    [long[]]$targetColors = @()

    $cell = $workbook.Sheets.Item("設定").Range("F4")

    # 塗りつぶしの色が白になるまで取得する
    while ($cell.Interior.Color -ne 16777215) {
        $targetColors += $cell.Interior.Color
        $cell = $cell.Offset(1,0)
    }
    return($targetColors)
}


# function getColsStatus($worksheet,$standardRange) {
#     $isUsedCols = @{}
#     Write-Host $standardRange.Address()
#     $rowNumber = $standardRange.Row
#     $colNumber = $standardRange.Offset(0,1).Column

#     $cellHeader = $worksheet.Cells($rowNumber, $colNumber)

#     while (-not [System.String]::IsNullOrEmpty($cellHeader.Value2()) -or 
#         $cellHeader.MergeCells) {

#         Write-Host $cellHeader.Address()

#         if($cellHeader.EntireColumn.Hidden){
#             $isUsedCols[$cellHeader.EntireColumn.Address()] = 'hidden'
#         }else{
#             $isUsedCols[$cellHeader.EntireColumn.Address()] = 'notUsed'
#         }

#         $colNumber++
#         $cellHeader = $worksheet.Cells($rowNumber, $colNumber)
#     }

#     foreach($key in $isUsedCols.Keys){
#         Write-Host $key ":" $isUsedCols[$key] 
#     }
#     return $isUsedCols
# }


$FILE_PATH = "C:\sampleMacro\ps1\課題\001\2\【VBA】001_スケジュールへの着色.xlsm"

# Excel COMオブジェクト作成
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true

try {
    $workbook = $excel.Workbooks.Open($FILE_PATH)
} catch {
    Write-Host "ファイルを開けませんでした" -ForegroundColor Red
    # exit 1
}

$worksheet = $workbook.Sheets.Item("スケジュール")
$standardRange = $worksheet.Range("B3")
$worksheet.activate()

# 連想配列を関数で作成するとvalueが空白になる
# 原因がわからないので関数化しない
# $isUsedCols = @{}
# Write-Host $worksheet.name()
# Write-Host $standardRange.Address()
# $isUsedCols = getColsStatus($worksheet, $standardRange)

# 各列の状態を保持する連想配列を生成
$isUsedCols = @{}   
$isUsedHeaders = @{}
$numberRow = $standardRange.Row
$numberCol = $standardRange.column + 1
$cellHeader = $worksheet.Cells($numberRow,$numberCol)
while (-not [System.String]::IsNullOrEmpty($cellHeader.Value2()) -or 
       $cellHeader.MergeCells) {
    $offsetCol
    
    Write-Host $cellHeader.Address()
    if($cellHeader.EntireColumn.Hidden){
        $isUsedCols[$cellHeader.column] = 'hidden'
    }else{  
        if ($cellHeader.MergeCells){
            $isUsedCols[$cellHeader.column] = $cellHeader.MergeArea.Cells.Item(1, 1).Address()
        }else{
            $isUsedCols[$cellHeader.column] = $cellHeader.Address()
        }
    }

    $numberCol++
    $cellHeader = $worksheet.Cells($numberRow,$numberCol)
}

$lastNumberCols = $numberCol - 1
Write-Host $lastNumberCols

foreach($key in $isUsedCols.Keys){
    Write-Host $key ":" $isUsedCols[$key] 
}


$targetColors = getColorsArray
$colorIndex = 0
$numberRow = $standardRange.Row + 1
$numberCol = $standardRange.Column
$cellDate = $worksheet.Cells($numberRow,$numberCol)


while (-not [System.String]::IsNullOrEmpty($cellDate.Value2())) {
    $baseDate = Get-Date -Year 1899 -Month 12 -Day 30  # Excelの起点
    $date = $baseDate.AddDays($cellDate.Value2())
    Write-Host $date.ToString("yyyy/MM/dd")

    for ($i = $numberCol+1; $i -le $lastNumberCols; $i++) {
        $cell = $worksheet.Cells($numberRow,$i)
        if ([long]::TryParse($cell.Value2(), [ref]$null)){
            Write-Host $cell.Address() "-->" $cell.column ":" $isUsedCols[$cell.column]
            try {
                $cellHeader = $worksheet.Range($isUsedCols[$cell.column])

                # 色の塗りつぶしを実行
                # (条件1)ヘッダーセルの塗りつぶしが実行されていない
                if($cellHeader.Interior.Color -eq $standardRange.Interior.Color){
                    # Write-Host $colorIndex % $targetColors.Length ":" $targetColors[$colorIndex % $targetColors.Length]
                    # 16777215 = 塗りつぶしなし
                    if($cellDate.Interior.Color -eq 16777215){
                        [long]$color = $targetColors[$colorIndex % $targetColors.Length]
                        $colorIndex++
                    }else{
                        [long]$color = $cellDate.Interior.Color
                    }
                    $cell.Interior.Color = [long]$color
                    $cellHeader.Interior.Color = [long]$color
                    $cellDate.Interior.Color = [long]$color
                    $isUsedHeaders[$cellHeader.Column] = $cellDate.Row
                # (条件2)結合セルandセルヘッダーを塗りつぶした日付と同じ場合
                }elseif($cellHeader.MergeCells -and
                       $isUsedHeaders[$cellHeader.Column] -eq $cellDate.Row){
                    # [long]でキャストしないとエラーになる
                    $cell.Interior.Color = [long]$cellHeader.Interior.Color
                }
                $isUsedCols[$cell.column] = "used"
            # hidden/usedの場合は対象外なので何もしない
            } catch {
                Write-Host "エラー発生: $($_.FullyQualifiedErrorId) : $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    }

    $numberRow++
    $cellDate = $worksheet.Cells($numberRow,$numberCol)
}
