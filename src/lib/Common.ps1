# 共通の道具：表示とログ、通知、操作の有無、確認画面、時刻の変換

$script:LogFile = $null
$script:Quiet = $false   # タスクから動くときは画面に出さない（ログだけ）

function Initialize-Log([string]$Dir) {
    if (-not (Test-Path -LiteralPath $Dir)) { New-Item -ItemType Directory -Path $Dir | Out-Null }
    $script:LogFile = Join-Path $Dir ('run_{0:yyyyMM}.log' -f (Get-Date))
}
function Write-Log([string]$Level, [string]$Message) {
    if ($script:LogFile) {
        try { Add-Content -LiteralPath $script:LogFile -Value ('{0:yyyy-MM-dd HH:mm:ss} [{1}] {2}' -f (Get-Date), $Level, $Message) -Encoding UTF8 } catch {}
    }
}
# 状態は [OK] [!!] [NG] の文字で表す。色は赤と緑を避ける
function Write-Ok($m)   { Write-Log 'OK' $m; if (-not $script:Quiet) { Write-Host "  [OK] $m" -ForegroundColor Cyan } }
function Write-Warn($m) { Write-Log '!!' $m; if (-not $script:Quiet) { Write-Host "  [!!] $m" -ForegroundColor Yellow } }
function Write-Ng($m)   { Write-Log 'NG' $m; if (-not $script:Quiet) { Write-Host "  [NG] $m" -ForegroundColor Magenta } }
function Write-Info($m) { Write-Log '  ' $m; if (-not $script:Quiet) { Write-Host "     $m" } }
function Write-Rule     { if (-not $script:Quiet) { Write-Host ('-' * 60) } }

# ------------------------------------------------------------
# 時刻
# ------------------------------------------------------------
$script:Tokyo = [TimeZoneInfo]::FindSystemTimeZoneById('Tokyo Standard Time')
function Get-JstNow { return [TimeZoneInfo]::ConvertTime([DateTimeOffset]::Now, $script:Tokyo).DateTime }
function ConvertTo-Jst($v) {
    # APIの日時（文字列。PowerShell 7 では日時に変換済みのこともある）→ 日本時間
    if ($v -is [DateTime]) { $dto = [DateTimeOffset]$v }
    else { $dto = [DateTimeOffset]::Parse([string]$v, [Globalization.CultureInfo]::InvariantCulture) }
    return [TimeZoneInfo]::ConvertTime($dto, $script:Tokyo).DateTime
}
function ConvertTo-DateOnly($v) {
    if ($v -is [DateTime]) { return $v.Date }
    return [DateTime]::ParseExact(([string]$v).Substring(0, 10), 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
}
function ConvertTo-Minutes([string]$hhmm) {
    if ($hhmm -notmatch '^\s*(\d{1,2}):(\d{2})\s*$') { throw "時刻の形が正しくありません：「$hhmm」（例：10:00）" }
    return [int]$Matches[1] * 60 + [int]$Matches[2]
}
function Format-Minutes($m) {
    if ($null -eq $m) { return '-' }
    return ('{0}:{1:00}' -f [int][Math]::Floor($m / 60), ($m % 60))
}

# ------------------------------------------------------------
# 最後にPCが操作されてからの秒数
# ------------------------------------------------------------
Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class NksIdle {
    [StructLayout(LayoutKind.Sequential)]
    struct LASTINPUTINFO { public uint cbSize; public uint dwTime; }
    [DllImport("user32.dll")] static extern bool GetLastInputInfo(ref LASTINPUTINFO p);
    public static double IdleSeconds() {
        LASTINPUTINFO l = new LASTINPUTINFO();
        l.cbSize = (uint)Marshal.SizeOf(l);
        if (!GetLastInputInfo(ref l)) { return 0; }
        return ((uint)Environment.TickCount - l.dwTime) / 1000.0;
    }
}
'@
function Get-IdleSeconds { return [NksIdle]::IdleSeconds() }

# ------------------------------------------------------------
# 通知（Windowsのトースト）。出せなくても処理は止めない
# ------------------------------------------------------------
function Show-Toast([string]$Title, [string]$Body) {
    Write-Log 'TOAST' "$Title / $Body"
    try {
        [void][Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime]
        [void][Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime]
        $t = [Security.SecurityElement]::Escape($Title)
        $b = [Security.SecurityElement]::Escape($Body)
        $xml = New-Object Windows.Data.Xml.Dom.XmlDocument
        $xml.LoadXml("<toast><visual><binding template='ToastGeneric'><text>$t</text><text>$b</text></binding></visual></toast>")
        $toast = New-Object Windows.UI.Notifications.ToastNotification $xml
        $appId = '{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\WindowsPowerShell\v1.0\powershell.exe'
        [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($appId).Show($toast)
    } catch {
        Write-Log '!!' ("通知を出せませんでした：{0}" -f $_.Exception.Message)
    }
}

# ------------------------------------------------------------
# 手前に出るが、キー入力を奪わない画面
# ------------------------------------------------------------
Add-Type -AssemblyName System.Windows.Forms, System.Drawing
Add-Type -WarningAction SilentlyContinue -ReferencedAssemblies System.Windows.Forms, System.Drawing @'
using System.Windows.Forms;
public class NksNoActivateForm : Form {
    protected override bool ShowWithoutActivation { get { return true; } }
    protected override CreateParams CreateParams {
        get {
            CreateParams cp = base.CreateParams;
            cp.ExStyle |= 0x08000000 | 0x00000008;   // WS_EX_NOACTIVATE | WS_EX_TOPMOST
            return cp;
        }
    }
}
'@

function Show-ConfirmDialog([string]$Message, [int]$Seconds = 30) {
    # 戻り値：run（実行）／later（10分後）／skip（今日はやめる）
    $form = New-Object NksNoActivateForm
    $form.Text = 'notion-kintai-sync'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false; $form.MinimizeBox = $false
    $form.StartPosition = 'Manual'
    $form.ClientSize = New-Object Drawing.Size(420, 170)
    $wa = [Windows.Forms.Screen]::PrimaryScreen.WorkingArea
    $form.Location = New-Object Drawing.Point(($wa.Right - 440), ($wa.Bottom - 200))
    $font = New-Object Drawing.Font('Yu Gothic UI', 10)

    $label = New-Object Windows.Forms.Label
    $label.Text = $Message
    $label.Font = $font
    $label.Location = New-Object Drawing.Point(16, 12)
    $label.Size = New-Object Drawing.Size(390, 70)
    $form.Controls.Add($label)

    $count = New-Object Windows.Forms.Label
    $count.Font = $font
    $count.Location = New-Object Drawing.Point(16, 86)
    $count.Size = New-Object Drawing.Size(390, 22)
    $form.Controls.Add($count)

    $script:ConfirmResult = 'run'
    $buttons = @(@('今すぐ実行', 'run'), @('10分後', 'later'), @('今日はやめる', 'skip'))
    $x = 16
    foreach ($b in $buttons) {
        $btn = New-Object Windows.Forms.Button
        $btn.Text = $b[0]; $btn.Tag = $b[1]; $btn.Font = $font
        $btn.Location = New-Object Drawing.Point($x, 120)
        $btn.Size = New-Object Drawing.Size(124, 34)
        $btn.Add_Click({ $script:ConfirmResult = $this.Tag; $this.FindForm().Close() })
        $form.Controls.Add($btn)
        $x += 132
    }

    $script:ConfirmLeft = $Seconds
    $count.Text = "$Seconds 秒後に自動で始めます"
    $timer = New-Object Windows.Forms.Timer
    $timer.Interval = 1000
    $timer.Add_Tick({
        $script:ConfirmLeft--
        $count.Text = "$($script:ConfirmLeft) 秒後に自動で始めます"
        if ($script:ConfirmLeft -le 0) { $script:ConfirmResult = 'run'; $form.Close() }
    })
    $timer.Start()
    [Windows.Forms.Application]::Run($form)
    $timer.Stop(); $timer.Dispose(); $form.Dispose()
    return $script:ConfirmResult
}

# 実行中に「触らないでください」を出しておく（別のプロセスで表示し、合図のファイルが消えたら閉じる）
function Start-BusyNotice([string]$FlagDir) {
    $flag = Join-Path $FlagDir ('busy_{0}.flag' -f [Guid]::NewGuid().ToString('N'))
    Set-Content -LiteralPath $flag -Value 'busy' -Encoding ASCII
    $job = Start-Job -ArgumentList $flag -ScriptBlock {
        param($flag)
        Add-Type -AssemblyName System.Windows.Forms, System.Drawing
        Add-Type -WarningAction SilentlyContinue -ReferencedAssemblies System.Windows.Forms, System.Drawing @'
using System.Windows.Forms;
public class NksBusyForm : Form {
    protected override bool ShowWithoutActivation { get { return true; } }
    protected override CreateParams CreateParams { get { CreateParams cp = base.CreateParams; cp.ExStyle |= 0x08000000 | 0x00000008; return cp; } }
}
'@
        $f = New-Object NksBusyForm
        $f.FormBorderStyle = 'None'
        $f.ShowInTaskbar = $false
        $f.BackColor = [Drawing.Color]::FromArgb(40, 40, 40)
        $f.ClientSize = New-Object Drawing.Size(360, 56)
        $f.StartPosition = 'Manual'
        $wa = [Windows.Forms.Screen]::PrimaryScreen.WorkingArea
        $f.Location = New-Object Drawing.Point(($wa.Right - 380), ($wa.Top + 20))
        $l = New-Object Windows.Forms.Label
        $l.Text = "勤怠Excelを作成中です`nマウスとキーボードに触らないでください"
        $l.ForeColor = [Drawing.Color]::White
        $l.Font = New-Object Drawing.Font('Yu Gothic UI', 10)
        $l.Dock = 'Fill'; $l.TextAlign = 'MiddleCenter'
        $f.Controls.Add($l)
        $t = New-Object Windows.Forms.Timer
        $t.Interval = 500
        $t.Add_Tick({ if (-not (Test-Path -LiteralPath $flag)) { $f.Close() } })
        $t.Start()
        [Windows.Forms.Application]::Run($f)
    }
    return @{ Flag = $flag; Job = $job }
}
function Stop-BusyNotice($Notice) {
    if (-not $Notice) { return }
    Remove-Item -LiteralPath $Notice.Flag -Force -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 700
    try { Stop-Job $Notice.Job; Remove-Job $Notice.Job -Force } catch {}
}

function Test-FileLocked([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) { return $false }
    try { $fs = [IO.File]::Open($Path, 'Open', 'ReadWrite', 'None'); $fs.Close(); return $false }
    catch { return $true }
}
