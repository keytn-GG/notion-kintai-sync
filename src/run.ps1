<#
  notion-kintai-sync：Notion のタイムカードから、会社の勤怠Excelを作る・更新する

  動かし方（-Mode）
    Auto      タスクスケジューラから呼ぶ。画面には何も出さず、ログと通知だけ
    Now       今すぐ更新する（今月分。月初の数日は前の月も）
    Finalize  確定する（最後にもう一度更新してから、以降は自動で更新しない）
    Rebuild   月を指定して作り直す（-Year -Month。確定も取り消す）
    Menu      run.cmd から呼ぶ。メニューを出して選んでもらう

  仕組み（2段階方式）
    A：カレンダーを作る（マウスを使う・月に1回）。席を外しているとき（操作がない）に実行。
       20日を過ぎても機会がなければ確認画面を出す
    B：時刻を書き込む（画面に出さない・何度でも）。Notion と前回書いた値を比べ、変わったセルだけ書く
#>
param(
    [ValidateSet('Auto', 'Now', 'Finalize', 'Rebuild', 'Menu')][string]$Mode = 'Auto',
    [int]$Year = 0,
    [int]$Month = 0,
    [switch]$Open,
    [switch]$ResetToken
)

$ErrorActionPreference = 'Stop'
# PowerShell 7 のターミナルから run.cmd を動かすと、PowerShell 7 用のモジュールの場所が引き継がれ、
# Windows PowerShell 5.1 が標準のコマンド（ConvertTo-SecureString など）を読み込めなくなる。7 用の場所を外す
if ($PSVersionTable.PSVersion.Major -le 5 -and $env:PSModulePath) {
    $env:PSModulePath = (($env:PSModulePath -split ';') | Where-Object { $_ -and $_ -notmatch '\\PowerShell\\' }) -join ';'
    Import-Module (Join-Path $PSHOME 'Modules\Microsoft.PowerShell.Security') -ErrorAction SilentlyContinue
}
try { [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch {}
$SrcDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$RepoRoot = Split-Path -Parent $SrcDir
foreach ($lib in 'Common', 'Config', 'Notion', 'DialogAutomation', 'Excel') { . (Join-Path $SrcDir "lib\$lib.ps1") }

$LogDir = Join-Path $RepoRoot 'logs'
Initialize-Log $LogDir
$script:Quiet = ($Mode -eq 'Auto')

# ------------------------------------------------------------
# メニュー（run.cmd から）
# ------------------------------------------------------------
function Read-YearMonth([string]$Prompt, [DateTime]$Default) {
    $ans = Read-Host ("  {0} [{1:yyyy-MM}]" -f $Prompt, $Default)
    if (-not $ans) { return $Default }
    if ($ans -notmatch '^\s*(\d{4})-(\d{1,2})\s*$') { throw '「2026-10」の形で入力してください' }
    return (Get-Date -Year ([int]$Matches[1]) -Month ([int]$Matches[2]) -Day 1)
}
function Show-Menu {
    $now = Get-JstNow
    Write-Rule
    Write-Host '  notion-kintai-sync'
    Write-Rule
    Write-Host ("  1. 今すぐ更新する（{0}月分）" -f $now.Month)
    Write-Host '  2. 確定する（これ以上自動で更新しない）'
    Write-Host '  3. 月を指定して作り直す'
    Write-Host '  0. 終了'
    $n = Read-Host '  番号を入力してください [1]'
    if (-not $n) { $n = '1' }
    switch ($n) {
        '1' { return @{ Mode = 'Now' } }
        '2' { $d = Read-YearMonth '確定する月' $now; return @{ Mode = 'Finalize'; Year = $d.Year; Month = $d.Month } }
        '3' { $d = Read-YearMonth '作り直す月' $now; return @{ Mode = 'Rebuild'; Year = $d.Year; Month = $d.Month } }
        default { return $null }
    }
}

# ------------------------------------------------------------
# A を実行してよいか（Auto のとき）
# ------------------------------------------------------------
function Test-CanRunA($Cfg, $State, [DateTime]$Today) {
    if ($State.skipDate -eq $Today.ToString('yyyy-MM-dd')) { Write-Info '今日は「やめる」が選ばれているので、カレンダー作成は明日にします'; return 'no' }
    $idle = Get-IdleSeconds
    if ($idle -ge [int]$Cfg.idleMinutes * 60) { Write-Info ("{0:N0} 秒操作がないので、席を外していると判断しました" -f $idle); return 'yes' }
    if ($Today.Day -lt [int]$Cfg.forceConfirmDay) { Write-Info 'PCが操作中なので、カレンダー作成はあとにします'; return 'no' }
    for ($i = 0; $i -lt 3; $i++) {
        Show-Toast 'notion-kintai-sync' '勤怠Excelのカレンダーを作ります。画面右下の確認を見てください'
        $ans = Show-ConfirmDialog ("今月の勤怠Excelのカレンダーを作ります（約10秒）。`n実行中はマウスとキーボードに触らないでください。") ([int]$Cfg.confirmSeconds)
        Write-Info "確認画面の答え：$ans"
        if ($ans -eq 'run') { return 'yes' }
        if ($ans -eq 'skip') { return 'skip' }
        Start-Sleep -Seconds 600
    }
    return 'no'
}

# ------------------------------------------------------------
# 1か月分の処理
# ------------------------------------------------------------
function Sync-Month([int]$Y, [int]$M, [string]$RunMode, [bool]$IsPrevious) {
    $cfg = $script:App.Config; $prof = $script:App.Profile
    $today = (Get-JstNow).Date
    Write-Rule
    Write-Ok ("{0}年{1}月分" -f $Y, $M)
    $state = Read-MonthState $Y $M
    $outFile = Get-OutputPath $cfg $Y $M

    if ($RunMode -eq 'Rebuild') {
        if (Test-Path -LiteralPath $outFile) {
            if (Test-FileLocked $outFile) { throw "ファイルが Excel で開かれています。閉じてから実行してください：$outFile" }
            $bak = $outFile + ('.bak_{0:yyyyMMddHHmmss}' -f (Get-Date))
            Move-Item -LiteralPath $outFile -Destination $bak
            Write-Info "前のファイルは名前を変えて残しました：$bak"
        }
        Remove-MonthState $Y $M
        $state = New-MonthState
    }
    if ($state.finalized -and $RunMode -in 'Auto', 'Now') {
        Write-Info '確定済みなので更新しません（作り直すときは、メニューの「3. 月を指定して作り直す」）'
        return
    }

    # Notion から読む
    $ids = Get-NotionPropertyIds $cfg.dataSourceId
    $rows = Get-MonthRows $cfg.dataSourceId $ids $Y $M
    $regular = Get-RegularHours $cfg.dataSourceId $ids $rows
    $days = ConvertTo-DayRecords $rows
    Write-Ok ("Notion から {0} 行を読みました（定時 {1}）" -f $rows.Count, $regular.Text)

    # A：カレンダーがなければ作る
    if (-not (Test-Path -LiteralPath $outFile)) {
        if ($IsPrevious) { Write-Info '前の月のファイルがないので、何もしません'; return }
        if ($RunMode -eq 'Auto') {
            $can = Test-CanRunA $cfg $state $today
            if ($can -eq 'skip') { $state.skipDate = $today.ToString('yyyy-MM-dd'); Save-MonthState $Y $M $state; return }
            if ($can -ne 'yes') { return }
        } else {
            Write-Warn 'カレンダーを作ります。マウスが動くので、終わるまで触らないでください（やめるときは Ctrl+C）'
            for ($i = 5; $i -ge 1; $i--) { Write-Host "     $i..."; Start-Sleep -Seconds 1 }
        }
        $notice = Start-BusyNotice $LogDir
        try { [void](New-MonthWorkbook $Y $M $outFile $regular $cfg $prof $LogDir) }
        finally { Stop-BusyNotice $notice }
        $state = New-MonthState
        Save-MonthState $Y $M $state
    }

    # B：変わったセルだけ書く
    $desired = Get-DesiredCells $Y $M $days $regular $cfg $prof
    $changes = Get-ChangedCells $desired $state.cells
    # 手直しで書かなかったセルは、Notion の値がそのときから変わっていなければ、もう一度は知らせない
    foreach ($cell in @($state.manual.Keys)) {
        if ($state.cells -and (Test-SameValue $desired[$cell] $state.cells[$cell])) { $state.manual.Remove($cell); continue }   # Notion が元に戻った
        if (Test-SameValue $desired[$cell] $state.manual[$cell]) { $changes = @($changes | Where-Object { $_ -ne $cell }) }
    }
    if ($state.manual.Count -gt 0) { Write-Info ("Excel で手直しされたままのセル：{0}（Notion の値は書いていません）" -f (@($state.manual.Keys) -join ', ')) }
    $latest = ($days.Values | Where-Object { $null -ne $_.Out } | Measure-Object -Property Day -Maximum).Maximum
    if ($changes.Count -eq 0) {
        Write-Ok '変化がないので、Excel は開きませんでした'
        Save-MonthState $Y $M $state   # 手直しの記録が変わっていることがあるので保存しておく
    } elseif (Test-FileLocked $outFile) {
        Write-Warn 'ファイルが Excel で開かれているので、今回は書きませんでした（閉じれば次の回に書きます）'
        if ($RunMode -ne 'Auto') { throw "閉じてから、もう一度実行してください：$outFile" }
        return
    } else {
        Write-Info ("書き換えが必要なセル：{0} 個" -f $changes.Count)
        $res = Update-MonthWorkbook $outFile $desired $changes $state.cells $prof $Y $M
        $state.cells = $res.NewCells
        foreach ($cell in $changes) { if ($res.Skipped -contains $cell) { $state.manual[$cell] = $desired[$cell] } else { $state.manual.Remove($cell) } }
        $state.updated = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        $state.latestDay = $latest
        Save-MonthState $Y $M $state
        $msg = if ($latest) { "{0}/{1} の退勤まで反映しました" -f $M, $latest } else { '反映しました' }
        if ($res.Skipped.Count -gt 0) { $msg += ("。Excel で手直しされたセル {0} 個は書き換えていません" -f $res.Skipped.Count) }
        Write-Ok ("{0} 個のセルを書きました。{1}" -f $res.Written, $msg)
        Show-Toast ("{0}月分の勤怠を更新しました" -f $M) $msg
    }

    # 月の最終日の夜：出勤はあるのに退勤がない
    $lastDay = [DateTime]::DaysInMonth($Y, $M)
    if ($RunMode -eq 'Auto' -and -not $IsPrevious -and $today.Day -eq $lastDay -and (Get-JstNow).Hour -ge [int]$cfg.lastDayAlertHour -and -not $state.lastDayAlerted) {
        if ($days.ContainsKey($lastDay) -and $null -eq $days[$lastDay].Out) {
            Show-Toast '退勤が未打刻です' ("今日（{0}/{1}）の退勤がまだ打刻されていません。打刻すると勤怠Excelに反映されます" -f $M, $lastDay)
            $state.lastDayAlerted = $true
            Save-MonthState $Y $M $state
        }
    }

    if ($RunMode -eq 'Finalize') {
        $state.finalized = $true
        Save-MonthState $Y $M $state
        Write-Ok ("{0}年{1}月分を確定しました。これ以降は自動で更新しません" -f $Y, $M)
    }
    if ($RunMode -ne 'Auto' -and -not $IsPrevious -and (Test-Path -LiteralPath $outFile)) {
        Write-Info "ファイル：$outFile"
        if ($Open) { Invoke-Item -LiteralPath $outFile }   # -Open を付けたときだけ Excel で開く
    }
}

# ------------------------------------------------------------
# 本体
# ------------------------------------------------------------
try {
    if ($Mode -eq 'Menu') {
        $sel = Show-Menu
        if (-not $sel) { exit 0 }
        $Mode = $sel.Mode
        if ($sel.Year) { $Year = $sel.Year; $Month = $sel.Month }
    }
    Write-Log 'RUN' "開始（$Mode）"
    $script:App = Read-AppConfig $RepoRoot
    $cfg = $script:App.Config

    $now = Get-JstNow
    if ($Mode -eq 'Auto') {
        $h = $cfg.activeHours -split '-'
        $from = ConvertTo-Minutes $h[0]; $to = ConvertTo-Minutes $h[1]
        $nowM = $now.Hour * 60 + $now.Minute
        if ($nowM -lt $from -or $nowM -ge $to) { Write-Log 'RUN' '動く時間帯の外なので終了'; exit 0 }
    }

    $token = Get-NotionToken -Interactive:($Mode -ne 'Auto') -Reset:$ResetToken
    Initialize-Notion $token
    $token = $null

    if ($Mode -in 'Finalize', 'Rebuild') {
        if (-not $Year) { $Year = $now.Year; $Month = $now.Month }
        Sync-Month $Year $Month $Mode $false
    } else {
        $prev = $now.AddMonths(-1)
        if ($now.Day -le [int]$cfg.previousMonthDays) { Sync-Month $prev.Year $prev.Month $Mode $true }
        Sync-Month $now.Year $now.Month $Mode $false
    }
    Write-Rule
    Write-Ok ("完了（Notion {0:N1} KB・{1} ミリ秒）" -f ($script:NotionBytes / 1KB), $script:NotionMs)
    Write-Log 'RUN' '終了'
    exit 0
}
catch {
    Write-Ng $_.Exception.Message
    Write-Log 'NG' ("{0} 行目：{1}" -f $_.InvocationInfo.ScriptLineNumber, $_.InvocationInfo.Line.Trim())
    Show-Toast 'notion-kintai-sync：失敗しました' ("{0}（ログ：{1}）" -f $_.Exception.Message, $script:LogFile)
    exit 1
}
