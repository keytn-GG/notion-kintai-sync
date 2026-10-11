# 設定ファイル・書式の設定（profile）・APIキー・記録（state）の読み書き

function Get-AppDataDir {
    $dir = Join-Path $env:APPDATA 'notion-kintai-sync'
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
    return $dir
}

function Read-JsonFile([string]$Path) {
    return (Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json)
}

function Read-AppConfig([string]$RepoRoot) {
    $cfgPath = Join-Path $RepoRoot 'config\config.json'
    if (-not (Test-Path -LiteralPath $cfgPath)) {
        throw "設定ファイルがありません：$cfgPath（config.example.json をコピーして作ってください）"
    }
    $cfg = Read-JsonFile $cfgPath
    foreach ($k in 'dataSourceId', 'template', 'profile', 'employeeNo', 'name', 'project') {
        if (-not $cfg.$k) { throw "設定ファイルに「$k」がありません：$cfgPath" }
    }
    # 既定値
    $defaults = @{
        outputDir          = 'output'
        outputName         = 'kintai_{yyyy}{MM}'
        previousMonthDays  = 3
        idleMinutes        = 3
        forceConfirmDay    = 20
        confirmSeconds     = 30
        activeHours        = '07:00-24:00'
        lastDayAlertHour   = 22
    }
    foreach ($k in $defaults.Keys) {
        if ($null -eq $cfg.$k -or "$($cfg.$k)" -eq '') { $cfg | Add-Member -NotePropertyName $k -NotePropertyValue $defaults[$k] -Force }
    }
    # パスは、リポジトリの一番上からの相対パスでも書ける
    foreach ($k in 'template', 'outputDir') {
        if (-not [IO.Path]::IsPathRooted($cfg.$k)) { $cfg.$k = Join-Path $RepoRoot $cfg.$k }
    }
    $profPath = Join-Path $RepoRoot ('config\profiles\{0}.profile.json' -f $cfg.profile)
    if (-not (Test-Path -LiteralPath $profPath)) { throw "書式の設定（profile）がありません：$profPath" }
    $prof = Read-JsonFile $profPath
    return @{ Config = $cfg; Profile = $prof; ConfigPath = $cfgPath; ProfilePath = $profPath }
}

# ------------------------------------------------------------
# APIキー（Windowsの機能で暗号化。このPCのこのユーザーだけが読める）
# ------------------------------------------------------------
function Get-NotionToken([switch]$Interactive, [switch]$Reset) {
    $file = Join-Path (Get-AppDataDir) 'token.dat'
    if ($Reset -and (Test-Path -LiteralPath $file)) { Remove-Item -LiteralPath $file -Force }
    if (-not (Test-Path -LiteralPath $file)) {
        if (-not $Interactive) { throw 'NotionのAPIキーが保存されていません。run.cmd から一度実行して登録してください' }
        Write-Info 'NotionのAPIキーがまだ保存されていません'
        $sec = Read-Host '  APIキーを貼り付けて Enter（画面には表示されません）' -AsSecureString
        $plain = [Runtime.InteropServices.Marshal]::PtrToStringBSTR([Runtime.InteropServices.Marshal]::SecureStringToBSTR($sec))
        if (-not $plain.StartsWith('ntn_')) { throw 'APIキーは ntn_ で始まる文字列です' }
        $sec | ConvertFrom-SecureString | Set-Content -LiteralPath $file -Encoding ASCII
        Write-Ok 'APIキーを暗号化して保存しました'
        return $plain
    }
    $sec = (Get-Content -LiteralPath $file -Raw).Trim() | ConvertTo-SecureString
    return [Runtime.InteropServices.Marshal]::PtrToStringBSTR([Runtime.InteropServices.Marshal]::SecureStringToBSTR($sec))
}

# ------------------------------------------------------------
# 月ごとの記録：前回書いた値・確定したか・今日はやめるか など
# ------------------------------------------------------------
function Get-StatePath([int]$Year, [int]$Month) {
    $dir = Join-Path (Get-AppDataDir) 'state'
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
    return (Join-Path $dir ('{0:0000}{1:00}.json' -f $Year, $Month))
}
function New-MonthState { return @{ version = 1; finalized = $false; cells = $null; manual = @{}; updated = $null; skipDate = $null; lastDayAlerted = $false; latestDay = $null } }
function Read-MonthState([int]$Year, [int]$Month) {
    $path = Get-StatePath $Year $Month
    $st = New-MonthState
    if (-not (Test-Path -LiteralPath $path)) { return $st }
    $obj = Read-JsonFile $path
    foreach ($k in 'finalized', 'updated', 'skipDate', 'lastDayAlerted', 'latestDay') { if ($null -ne $obj.$k) { $st[$k] = $obj.$k } }
    if ($obj.cells) {
        $cells = @{}
        foreach ($p in $obj.cells.PSObject.Properties) { $cells[$p.Name] = $p.Value }
        $st.cells = $cells
    }
    if ($obj.manual) {
        # 手直しされていて書かなかったセルと、そのときに書こうとした値（同じ通知を何度も出さないため）
        foreach ($p in $obj.manual.PSObject.Properties) { $st.manual[$p.Name] = $p.Value }
    }
    return $st
}
function Save-MonthState([int]$Year, [int]$Month, $State) {
    $State.version = 1
    $State | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Get-StatePath $Year $Month) -Encoding UTF8
}
function Remove-MonthState([int]$Year, [int]$Month) {
    Remove-Item -LiteralPath (Get-StatePath $Year $Month) -Force -ErrorAction SilentlyContinue
}

function Get-OutputPath($Config, [int]$Year, [int]$Month) {
    $name = $Config.outputName.Replace('{yyyy}', ('{0:0000}' -f $Year)).Replace('{MM}', ('{0:00}' -f $Month)).Replace('{employeeNo}', [string]$Config.employeeNo)
    $ext = [IO.Path]::GetExtension($Config.template)
    return (Join-Path $Config.outputDir ($name + $ext))
}
