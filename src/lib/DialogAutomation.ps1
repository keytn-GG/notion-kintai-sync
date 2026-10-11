# 日付設定の自動化：［日付設定］ボタンをマウスで押し、出てくる2つの画面を別のプロセスで押す
# （VBAにはパスワードがかかっているので、マクロを直接は呼べない。パスワードの解除はしない）

$script:DateSetupClicker = {
    param($TimeoutSec, $FormTitle, $StartButton, $ConfirmTitle, $OkButton, $LogPath, $StartRelX, $StartRelY, $FormClass, $ConfirmClass, $StartKeys)

    function Write-ClickerLog([string]$m) {
        Add-Content -LiteralPath $LogPath -Value ("{0:HH:mm:ss} {1}" -f (Get-Date), $m) -Encoding UTF8
    }
    Write-ClickerLog 'INFO|画面を押す係を起動しました'

    Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes, System.Windows.Forms
    Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class NksMouse {
    [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern IntPtr FindWindow(string cls, string title);
    [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr h, uint msg, IntPtr w, IntPtr l);
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] public static extern void keybd_event(byte vk, byte scan, uint flags, UIntPtr extra);
    public static bool Focus(IntPtr h) {
        // 別のプロセスから画面を手前に出すときの定番の手順（Altキーを一度押して離す）
        keybd_event(0x12, 0, 0, UIntPtr.Zero);
        keybd_event(0x12, 0, 2, UIntPtr.Zero);
        SetForegroundWindow(h);
        System.Threading.Thread.Sleep(300);
        return GetForegroundWindow() == h;
    }
    [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
    [DllImport("user32.dll")] public static extern void mouse_event(uint f, uint x, uint y, uint d, UIntPtr e);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
    public static void Click(int x, int y) {
        SetCursorPos(x, y);
        mouse_event(0x0002, 0, 0, 0, UIntPtr.Zero);
        mouse_event(0x0004, 0, 0, 0, UIntPtr.Zero);
    }
}
'@
    [NksMouse]::SetProcessDPIAware() | Out-Null
    $AE    = [System.Windows.Automation.AutomationElement]
    $Scope = [System.Windows.Automation.TreeScope]

    function Find-Window([string]$cls, [string]$title, [int]$sec = $TimeoutSec) {
        # Win32 の FindWindow で探す（ほかの画面に「持たれている」画面も見つかる）
        $deadline = (Get-Date).AddSeconds($sec)
        while ((Get-Date) -lt $deadline) {
            $h = [NksMouse]::FindWindow($cls, $title)
            if ($h -ne [IntPtr]::Zero) { return $AE::FromHandle($h) }
            Start-Sleep -Milliseconds 300
        }
        return $null
    }

    function Show-Tree($win) {
        Write-ClickerLog "INFO|--- 画面「$($win.Current.Name)」の中身（クラス: $($win.Current.ClassName)、位置: $($win.Current.BoundingRectangle)） ---"
        $walker = [System.Windows.Automation.TreeWalker]::RawViewWalker
        $stack = New-Object System.Collections.Stack
        $stack.Push(@($win, 0))
        $count = 0
        while ($stack.Count -gt 0 -and $count -lt 80) {
            $item = $stack.Pop(); $el = $item[0]; $depth = $item[1]
            $c = $el.Current
            Write-ClickerLog ("INFO|  {0}{1} / 名前「{2}」 / クラス {3}" -f ('  ' * $depth), $c.ControlType.ProgrammaticName, $c.Name, $c.ClassName)
            $count++
            $child = $walker.GetFirstChild($el)
            $kids = @()
            while ($child) { $kids += $child; $child = $walker.GetNextSibling($child) }
            [array]::Reverse($kids)
            foreach ($k in $kids) { $stack.Push(@($k, $depth + 1)) }
        }
    }

    function Invoke-DialogButton($win, [string]$name, $relX, $relY) {
        $cond = New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, $name)
        $btn = $win.FindFirst($Scope::Descendants, $cond)
        if ($btn) {
            try {
                $p = $btn.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)
                $p.Invoke()
                return "名前で見つけて押した（Invoke）"
            } catch {
                $r = $btn.Current.BoundingRectangle
                if ($r.Width -gt 0) {
                    [NksMouse]::SetForegroundWindow([IntPtr]$win.Current.NativeWindowHandle) | Out-Null
                    Start-Sleep -Milliseconds 200
                    [NksMouse]::Click([int]($r.X + $r.Width / 2), [int]($r.Y + $r.Height / 2))
                    return "名前で見つけて押した（マウス）"
                }
            }
        }
        if ($null -ne $relX) {
            # 名前で見つからないときは、画面の大きさに対する割合の位置をクリックする
            $wr = $win.Current.BoundingRectangle
            $x = [int]($wr.X + $wr.Width * $relX)
            $y = [int]($wr.Y + $wr.Height * $relY)
            [NksMouse]::SetForegroundWindow([IntPtr]$win.Current.NativeWindowHandle) | Out-Null
            Start-Sleep -Milliseconds 300
            [NksMouse]::Click($x, $y)
            return "名前では見つからず、位置（x=$x, y=$y）で押した"
        }
        return "見つからない"
    }

    # 1つ目の画面
    $form = Find-Window $FormClass $FormTitle
    if (-not $form) { Write-ClickerLog "NG|「$FormTitle」の画面が $TimeoutSec 秒以内に見つかりませんでした"; return }
    Write-ClickerLog "OK|「$FormTitle」の画面を見つけました"
    Show-Tree $form
    Start-Sleep -Milliseconds 500
    # ［設定開始］はキーボードで押す。だめなら位置でクリックする
    $hForm = [IntPtr]$form.Current.NativeWindowHandle
    if ([NksMouse]::Focus($hForm)) {
        [System.Windows.Forms.SendKeys]::SendWait($StartKeys)
        Write-ClickerLog "OK|［$StartButton］：キーボード（$StartKeys）で押した"
    } else {
        Write-ClickerLog "NG|「$FormTitle」の画面を手前に出せなかったので、キーボードは使いませんでした"
    }
    if (-not (Find-Window $ConfirmClass $ConfirmTitle 3)) {
        if ([NksMouse]::FindWindow($FormClass, $FormTitle) -ne [IntPtr]::Zero) {
            Write-ClickerLog "INFO|3秒たっても「$ConfirmTitle」の画面が出ないので、位置でクリックします"
            $r1 = Invoke-DialogButton $form $StartButton $StartRelX $StartRelY
            Write-ClickerLog "OK|［$StartButton］：$r1"
        }
    }

    # 2つ目の画面
    $dlg = Find-Window $ConfirmClass $ConfirmTitle
    if (-not $dlg) { Write-ClickerLog "NG|「$ConfirmTitle」の画面が $TimeoutSec 秒以内に見つかりませんでした（［$StartButton］が押せていない可能性）。手で押してください"; return }
    Write-ClickerLog "OK|「$ConfirmTitle」の画面を見つけました"
    Show-Tree $dlg
    Start-Sleep -Milliseconds 300
    $r2 = Invoke-DialogButton $dlg $OkButton $null $null
    if ($r2 -eq '見つからない') {
        # Windows標準のメッセージ画面なので、「OKが押された」という合図を直接送る
        [NksMouse]::SendMessage([IntPtr]$dlg.Current.NativeWindowHandle, 0x0111, [IntPtr]1, [IntPtr]::Zero) | Out-Null
        $r2 = 'ボタンは見つからず、OKの合図を直接送った'
    }
    Write-ClickerLog "OK|［$OkButton］：$r2"
    Start-Sleep -Seconds 1
    if ([NksMouse]::FindWindow($ConfirmClass, $ConfirmTitle) -ne [IntPtr]::Zero) { Write-ClickerLog "NG|「$ConfirmTitle」の画面がまだ残っています。手で［$OkButton］を押してください" }
    else { Write-ClickerLog "OK|「$ConfirmTitle」の画面が閉じました" }
    Write-ClickerLog 'INFO|画面を押す係を終了します'
}

Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class NksMain {
    [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
    [DllImport("user32.dll")] public static extern uint GetDpiForWindow(IntPtr h);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
    [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
    [DllImport("user32.dll")] public static extern void mouse_event(uint f, uint x, uint y, uint d, UIntPtr e);
    public static void Click(int x, int y) {
        SetCursorPos(x, y);
        mouse_event(0x0002, 0, 0, 0, UIntPtr.Zero);
        mouse_event(0x0004, 0, 0, 0, UIntPtr.Zero);
    }
}
'@

Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class NksWin {
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern IntPtr FindWindow(string cls, string title);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
    public static bool Shown(string cls, string title) { IntPtr h = FindWindow(cls, title); return h != IntPtr.Zero && IsWindowVisible(h); }
}
'@

function Invoke-DateSetup($xl, $ws, $ds, [string]$LogDir) {
    # 戻り値：画面の処理が終わったら $true（カレンダーができたかは呼んだ側で確かめる）
    $job = $null
    try {
        # --------------------------------------------------------
        # ［日付設定］ボタンを探す（ActiveXのボタン）
        # --------------------------------------------------------
        Write-Info '--- シート上のActiveXのボタン ---'
        $btnShape = $null
        foreach ($ole in $ws.OLEObjects()) {
            $caption = ''
            try { $caption = $ole.Object.Caption } catch {}
            Write-Info ("名前「{0}」 / 表示「{1}」" -f $ole.Name, $caption)
            if (-not $btnShape -and ($ole.Name -eq $ds.buttonName -or $caption -like "*$($ds.buttonText)*")) { $btnShape = $ws.Shapes.Item($ole.Name) }
        }
        if (-not $btnShape) { Write-Ng "［$($ds.buttonText)］ボタンが見つかりませんでした"; return $false }
        Write-Ok ("［{0}］ボタンを見つけました（名前 {1}）" -f $ds.buttonText, $btnShape.Name)

        # --------------------------------------------------------
        # 画面を押す係を先に起動してから、ボタンを押す
        # --------------------------------------------------------
        $logDir = $LogDir
        if (-not (Test-Path -LiteralPath $logDir)) { New-Item -ItemType Directory -Path $logDir | Out-Null }
        $clickLog = Join-Path $logDir ('clicker_{0:yyyyMMdd_HHmmss}.log' -f (Get-Date))
        $job = Start-Job -ScriptBlock $script:DateSetupClicker -ArgumentList ([int]$ds.timeoutSec), $ds.formTitle, $ds.startButton, $ds.confirmTitle, $ds.okButton, $clickLog, ([double]$ds.startRelX), ([double]$ds.startRelY), $ds.formClass, $ds.confirmClass, $ds.startKeys
        Write-Info '画面を押す係を起動しています（3秒待ちます）...'
        Start-Sleep -Seconds 3

        # ボタンの画面上の位置を求めてクリックする
        [NksMain]::SetProcessDPIAware() | Out-Null
        $win = $xl.ActiveWindow
        $win.ScrollRow = 1
        $win.ScrollColumn = 1
        $hwnd = [IntPtr]$xl.Hwnd
        $dpi = [NksMain]::GetDpiForWindow($hwnd); if (-not $dpi) { $dpi = 96 }
        $scale = $dpi / 72.0 * $win.Zoom / 100.0
        $cx = [int]($win.PointsToScreenPixelsX(0) + ($btnShape.Left + $btnShape.Width  / 2) * $scale)
        $cy = [int]($win.PointsToScreenPixelsY(0) + ($btnShape.Top  + $btnShape.Height / 2) * $scale)
        Write-Info ("ボタンの位置: x={0}, y={1}（DPI {2}、倍率 {3}%）" -f $cx, $cy, $dpi, $win.Zoom)

        [NksMain]::SetForegroundWindow($hwnd) | Out-Null
        Start-Sleep -Milliseconds 500
        [NksMain]::SetCursorPos($cx, $cy) | Out-Null
        Write-Info 'マウスをボタンの上に移動しました。1秒後にクリックします'
        Start-Sleep -Seconds 1
        $started = Get-Date
        [NksMain]::Click($cx, $cy)
        Write-Info 'クリックしました。画面が出ても、しばらく触らずに待ってください'
        Write-Info '（ボタンが押されなかったときは、手で［日付設定］を押してください）'

        # 押す係が終わるまで待つ（その間は Excel に話しかけない）。報告は届いたそばから表示する
        Write-Info '--- 画面を押す係の報告 ---'
        $shown = 0
        $limit = (Get-Date).AddSeconds([int]$ds.timeoutSec * 2 + 20)
        while ($true) {
            if (Test-Path -LiteralPath $clickLog) {
                $lines = @(Get-Content -LiteralPath $clickLog -Encoding UTF8)
                for ($i = $shown; $i -lt $lines.Count; $i++) {
                    $l = $lines[$i]; $body = $l.Substring(9)
                    if ($body -like 'OK|*')     { Write-Ok ($l.Substring(0, 9) + $body.Substring(3)) }
                    elseif ($body -like 'NG|*') { Write-Ng ($l.Substring(0, 9) + $body.Substring(3)) }
                    else                        { Write-Info ($l.Substring(0, 9) + $body.Substring(5)) }
                }
                $shown = $lines.Count
            }
            if ($job.State -ne 'Running') { break }
            if ((Get-Date) -gt $limit) { Write-Warn '待ち時間を過ぎました'; break }
            Start-Sleep -Milliseconds 500
        }
        $jobErr = Receive-Job -Job $job -ErrorAction SilentlyContinue 2>&1
        foreach ($e in $jobErr) { Write-Warn "押す係のエラー: $e" }
        Remove-Job -Job $job -Force
        $job = $null
        # 画面が残っている間は Excel を触らない（手で押してもらうのを待つ）
        $warned = $false
        while ([NksWin]::Shown($ds.formClass, $ds.formTitle) -or [NksWin]::Shown($ds.confirmClass, $ds.confirmTitle)) {
            if (-not $warned) { Write-Warn '画面がまだ残っています。手で押してください。閉じるまでここで待ちます'; $warned = $true }
            Start-Sleep -Milliseconds 500
        }
        Start-Sleep -Seconds 2
        Write-Ok ("画面の処理が終わりました（{0:N1} 秒）" -f ((Get-Date) - $started).TotalSeconds)
        return $true
    }
    finally {
        if ($job) { try { Stop-Job $job; Remove-Job $job -Force } catch {} }
    }
}
