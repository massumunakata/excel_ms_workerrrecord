Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.Windows.Forms

# 1) ユーザーに数字を入力させる
$number = [Microsoft.VisualBasic.Interaction]::InputBox(
    "数字を入力してください",   # メッセージ
    "入力確認",                 # タイトル
    "0"                         # デフォルト値
)

[int]$num = 0
if (-not [int]::TryParse($number, [ref]$num)) {
    [System.Windows.Forms.MessageBox]::Show("数字が入力されませんでした")
    return
}

# 2)親フォルダのパス
# 実行中のスクリプトのフルパス
$scriptPath = $MyInvocation.MyCommand.Path

# スクリプトのあるフォルダ
$scriptDir = Split-Path $scriptPath

# その一つ上のフォルダ
$baseDir = Split-Path $scriptDir

Write-Host "親フォルダは: $baseDir"


# "10_"で始まるフォルダを検索（最初の1件を取得）
$targetFolder = Get-ChildItem -Path $baseDir -Directory |
                Where-Object { $_.Name -like "10_*" } |
                Select-Object -First 1

# 見つかった場合に $folderPath を設定
if ($targetFolder) {
    $folderPath = $targetFolder.FullName
    Write-Host "対象フォルダ: $folderPath"
} else {
    Write-Host "10_で始まるフォルダが見つかりませんでした。"
}


# 3) Excel COM 起動
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# 4) Rangeを1次元配列に変換する関数
function Flatten-Range2D {
    param($range2D)
    if (-not $range2D) { return @() }
    if (-not ($range2D -is [System.Array])) { return @($range2D) }
    $rows = $range2D.GetLength(0)
    $cols = $range2D.GetLength(1)
    $out  = New-Object System.Collections.Generic.List[object]
    for ($i = 1; $i -le $rows; $i++) {
        for ($j = 1; $j -le $cols; $j++) {
            $val = $range2D.GetValue($i, $j)
            $out.Add($val)
        }
    }
    return $out.ToArray()
}

# 5) 全配列を格納するリスト
$finalArrays = @()

try {
    $files = Get-ChildItem -Path $folderPath -File -Filter *.xls*

    foreach ($file in $files) {
        $wb = $excel.Workbooks.Open($file.FullName)
        $ws = $wb.Sheets.Item(1)

        # A4, R4, AA4 を取得
        $a4  = $ws.Range("A4").Value()
        $r4  = $ws.Range("R4").Value()
        $aa4 = $ws.Range("AA4").Value()

        # AI列と AJ列を配列化
        $aiArr = Flatten-Range2D ($ws.Range("AI1:AI400").Value())
        $ajArr = Flatten-Range2D ($ws.Range("AJ1:AJ400").Value())

        # 入力された数字がどこにあるか探す
        for ($i = 0; $i -lt $aiArr.Count; $i++) {
            if ($aiArr[$i] -eq $num) {
                $start = $i + 1
                $end   = [math]::Min($start + 6, $ajArr.Count - 1)
                $subArr = if ($end -ge $start) { $ajArr[$start..$end] } else { @() }

                # 7要素に不足があれば空文字で補充
                while ($subArr.Count -lt 7) { $subArr += "" }

                # 11要素の配列を作成
                # [A4, 入力数字, R4, AA4, AJ1..AJ7]
                $combinedArr = @($a4, $num, $r4, $aa4) + $subArr[0..6]

                # 最終配列リストに追加
                $finalArrays += ,$combinedArr
            }
        }

        $wb.Close($false)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb)
    }
}
catch {
    Write-Output "エラー: $($_.Exception.Message)"
}
finally {
    $excel.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    [gc]::Collect(); [gc]::WaitForPendingFinalizers()
}

# 6) 表示（最大10件）
$finalArrays | ForEach-Object {
    Write-Output ("配列(11要素): " + ($_ -join ', '))
}



# 7) CSV保存（同じ階層に出力）
if ($finalArrays.Count -gt 0) {
    # 最初の配列を使ってファイル名を決定
    $firstElement  = if ($finalArrays[0][0]) { $finalArrays[0][0].ToString() } else { "配列" }
    $secondElement = if ($finalArrays[0][1]) { $finalArrays[0][1].ToString() } else { "要素" }
    $fileName = "${firstElement}_${secondElement}_salary.csv"
    $csvPath  = Join-Path $baseDir $fileName   # 同じ階層に保存

    # 全配列をCSVに出力（複数行）
    $rows = @()
    foreach ($arr in $finalArrays) {
        $rows += [PSCustomObject]@{
            年度       = $arr[0]
            月度       = $arr[1]
            社員番号   = $arr[2]
            名前       = $arr[3]
            稼働日数   = $arr[4]
            勤務日数   = $arr[5]
            稼働時間数 = $arr[6]
            勤務時間数 = $arr[7]
            超過時間   = $arr[8]
            有給休暇   = $arr[9]
            時間有給   = $arr[10]
        }
    }

    $rows | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    Write-Output "CSVを書き出しました → $csvPath"
}
else {
    Write-Output "保存対象の配列がありません。"
}
