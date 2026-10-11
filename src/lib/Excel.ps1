# 勤怠Excelの操作
#   A：カレンダーを作る（複製 → 年月 → 6行目 → 日付設定）。マウスを使う・月に1回
#   B：時刻を書き込む（画面に出さずに開き、変わったセルだけ書く）。何度でも

Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class NksProc {
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);
}
'@

# ------------------------------------------------------------
# 値の比較と書き込み
# ------------------------------------------------------------
function Test-Blank($v) { return ($null -eq $v -or ($v -is [string] -and $v -eq '')) }
function Test-SameValue($a, $b) {
    if ((Test-Blank $a) -and (Test-Blank $b)) { return $true }
    if ((Test-Blank $a) -or (Test-Blank $b)) { return $false }
    $na = 0.0; $nb = 0.0
    $inv = [Globalization.CultureInfo]::InvariantCulture
    $isNa = [double]::TryParse([string]$a, [Globalization.NumberStyles]::Float, $inv, [ref]$na)
    $isNb = [double]::TryParse([string]$b, [Globalization.NumberStyles]::Float, $inv, [ref]$nb)
    if ($isNa -and $isNb -and -not ($a -is [string] -and $b -is [string])) { return [Math]::Abs($na - $nb) -lt 1e-6 }
    return ([string]$a -eq [string]$b)
}
function Format-CellValue($v) {
    if (Test-Blank $v) { return '（空欄）' }
    $n = 0.0
    if (-not ($v -is [string]) -and [double]::TryParse([string]$v, [ref]$n) -and $n -gt 0 -and $n -lt 2) { return (Format-Minutes ([int][Math]::Round($n * 1440))) }
    return [string]$v
}
function Set-CellValue($Range, $Value) {
    if (Test-Blank $Value) { [void]$Range.ClearContents(); return }
    if ($Value -is [string]) {
        # 文字は、普通の代入だと「キャストが無効」になることがあるので、Excelに直接渡す（試作3で判明）
        [void][System.__ComObject].InvokeMember('Value2', [Reflection.BindingFlags]::SetProperty, $null, $Range, [object[]]@((-join $Value.ToCharArray())))
        return
    }
    $Range.Value2 = [double]$Value
}

# ------------------------------------------------------------
# 書き込むべき値を作る（セル番地 → 値）
# ------------------------------------------------------------
function Get-TokenValues($Config, $Prof, $Regular) {
    return @{
        regularStart   = $Regular.Start
        regularEnd     = $Regular.End
        lateNightStart = ConvertTo-Minutes $Prof.lateNightStart
        employeeNo     = [string]$Config.employeeNo
        name           = [string]$Config.name
        project        = [string]$Config.project
    }
}
function Resolve-InitialValue($Item, $Tokens) {
    $key = ([string]$Item.value).Trim('{', '}')
    if (-not $Tokens.ContainsKey($key)) { throw "profile の initialCells に分からない値があります：$($Item.value)" }
    $v = $Tokens[$key]
    switch ($Item.type) {
        'time'   { return [double]$v / 1440.0 }
        'number' { return [double]$v }
        default  { return [string]$v }
    }
}
function Get-DesiredCells([int]$Year, [int]$Month, $Days, $Regular, $Config, $Prof, [string]$Phase = '') {
    $tokens = Get-TokenValues $Config $Prof $Regular
    $desired = [ordered]@{}
    foreach ($item in $Prof.initialCells) {
        if ($Phase -and $item.phase -ne $Phase) { continue }
        $desired[$item.cell] = Resolve-InitialValue $item $tokens
    }
    if ($Phase) { return $desired }
    $last = [DateTime]::DaysInMonth($Year, $Month)
    for ($d = 1; $d -le $last; $d++) {
        $r = [int]$Prof.firstDayRow + $d - 1
        $inCell = "$($Prof.inColumn)$r"; $outCell = "$($Prof.outColumn)$r"
        if (-not $Days.ContainsKey($d) -or $Days[$d].Kind -eq $script:PaidLeave) {
            $desired[$inCell] = $null; $desired[$outCell] = $null; continue
        }
        $o = $Days[$d]
        $wIn  = if ($null -ne $o.In)  { $o.In }  else { $Regular.Start }
        $wOut = if ($null -ne $o.Out) { $o.Out } else { $Regular.End }
        $desired[$inCell] = $wIn / 1440.0; $desired[$outCell] = $wOut / 1440.0
    }
    return $desired
}
function Get-ChangedCells($Desired, $StateCells) {
    $changes = @()
    foreach ($cell in $Desired.Keys) {
        if ($null -eq $StateCells -or -not $StateCells.ContainsKey($cell) -or -not (Test-SameValue $StateCells[$cell] $Desired[$cell])) { $changes += $cell }
    }
    return ,$changes
}

function Test-Calendar($ws, $Prof, [int]$Year, [int]$Month) {
    $first = Get-Date -Year $Year -Month $Month -Day 1
    $want = ('日','月','火','水','木','金','土')[[int]$first.DayOfWeek]
    $got = $ws.Range("$($Prof.dowColumn)$($Prof.firstDayRow)").Text
    return @{ Ok = ($got -eq $want); Want = $want; Got = $got }
}

function Close-Excel($xl, $wb, $ws) {
    if ($wb) { try { $wb.Close($false) } catch {} }
    if ($xl) { try { $xl.Quit() } catch {} }
    foreach ($o in $ws, $wb, $xl) { if ($o) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch {} } }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

# ------------------------------------------------------------
# A：カレンダーを作る
# ------------------------------------------------------------
function New-MonthWorkbook([int]$Year, [int]$Month, [string]$OutFile, $Regular, $Config, $Prof, [string]$LogDir) {
    if (-not (Test-Path -LiteralPath $Config.template)) { throw "勤怠Excel（テンプレート）が見つかりません：$($Config.template)" }
    $dir = Split-Path -Parent $OutFile
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
    Copy-Item -LiteralPath $Config.template -Destination $OutFile -Force
    Unblock-File -LiteralPath $OutFile
    Write-Ok "複製しました：$OutFile"

    $xl = $null; $wb = $null; $ws = $null
    try {
        $xl = New-Object -ComObject Excel.Application
        $xl.Visible = $true
        $xl.DisplayAlerts = $false
        $wb = $xl.Workbooks.Open($OutFile, 0)    # 0 ＝ ブックの外へのリンクを更新しない
        $ws = $wb.Worksheets.Item($Prof.sheet)
        $ws.Activate()
        $ws.Range($Prof.yearCell).Value2  = [string]$Year
        $ws.Range($Prof.monthCell).Value2 = [string]$Month
        foreach ($kv in (Get-DesiredCells $Year $Month $null $Regular $Config $Prof 'before').GetEnumerator()) {
            Set-CellValue $ws.Range($kv.Key) $kv.Value
        }
        Write-Ok ("{0}年{1}月と、日付設定の前に入れる項目を書きました" -f $Year, $Month)

        $done = Invoke-DateSetup $xl $ws $Prof.dateSetup $LogDir
        if (-not $done) { throw '日付設定を自動で押せませんでした' }
        $cal = Test-Calendar $ws $Prof $Year $Month
        if (-not $cal.Ok) { throw ("日付設定が反映されていません（1日の曜日が「{0}」。正しくは「{1}」）" -f $cal.Got, $cal.Want) }
        $wb.Save()
        Write-Ok ("{0}年{1}月のカレンダーを作りました" -f $Year, $Month)
        return $true
    }
    catch {
        # 途中で失敗したファイルは残さない（次の回にやり直す）
        Close-Excel $xl $wb $ws; $xl = $null; $wb = $null; $ws = $null
        Remove-Item -LiteralPath $OutFile -Force -ErrorAction SilentlyContinue
        throw
    }
    finally { Close-Excel $xl $wb $ws }
}

# ------------------------------------------------------------
# B：時刻を書き込む（変わったセルだけ。手直しされたセルは書かない）
# ------------------------------------------------------------
function Update-MonthWorkbook([string]$OutFile, $Desired, [string[]]$Changes, $StateCells, $Prof, [int]$Year, [int]$Month, [int]$TimeoutSec = 90) {
    $xl = $null; $wb = $null; $ws = $null; $watch = $null
    $result = @{ Written = 0; Skipped = @(); NewCells = @{} }
    foreach ($k in $Desired.Keys) { $result.NewCells[$k] = $Desired[$k] }
    try {
        $xl = New-Object -ComObject Excel.Application
        $xl.Visible = $false
        $xl.DisplayAlerts = $false
        $xl.ScreenUpdating = $false
        $xl.AutomationSecurity = 1     # マクロを動かしたまま開く。日ごとの計算はマクロが行う（試作4で判明）
        $xl.EnableEvents = $true

        # 見張り役：時間内に終わらなければ、この Excel だけを強制終了する
        $excelPid = [uint32]0
        [void][NksProc]::GetWindowThreadProcessId([IntPtr]$xl.Hwnd, [ref]$excelPid)
        $watch = Start-Job -ArgumentList $excelPid, $TimeoutSec -ScriptBlock {
            param($procId, $sec)
            Start-Sleep -Seconds $sec
            if (Get-Process -Id $procId -ErrorAction SilentlyContinue) { Stop-Process -Id $procId -Force; 'killed' }
        }

        $wb = $xl.Workbooks.Open($OutFile, 0)
        $ws = $wb.Worksheets.Item($Prof.sheet)
        $cal = Test-Calendar $ws $Prof $Year $Month
        if (-not $cal.Ok) { throw ("このファイルは{0}年{1}月のカレンダーになっていません（1日の曜日が「{2}」）。作り直してください" -f $Year, $Month, $cal.Got) }

        foreach ($cell in $Changes) {
            $range = $ws.Range($cell)
            if ($null -ne $StateCells -and $StateCells.ContainsKey($cell)) {
                $now = $range.Value2
                if (-not (Test-SameValue $now $StateCells[$cell])) {
                    Write-Warn ("{0} は Excel で手直しされているので書きません（今：{1}、前回書いた値：{2}）" -f $cell, (Format-CellValue $now), (Format-CellValue $StateCells[$cell]))
                    $result.Skipped += $cell
                    $result.NewCells[$cell] = $StateCells[$cell]   # 記録は前回の値のまま。次の回も手直しとして扱う
                    continue
                }
            }
            Set-CellValue $range $Desired[$cell]
            $result.Written++
        }
        $xl.Calculate()
        $wb.Save()
        return $result
    }
    catch {
        if ($watch -and ((Receive-Job -Job $watch -ErrorAction SilentlyContinue) -contains 'killed')) {
            throw "Excel が $TimeoutSec 秒以内に終わらなかったので、強制終了しました（隠れた画面で止まっていた可能性があります）"
        }
        throw
    }
    finally {
        if ($watch) { try { Stop-Job $watch; Remove-Job $watch -Force } catch {} }
        Close-Excel $xl $wb $ws
    }
}
