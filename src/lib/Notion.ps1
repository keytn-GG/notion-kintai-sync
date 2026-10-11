# Notion API から timecard.db の行と「定時」を読む（読み取りだけ）

$script:NotionHeaders = $null
$script:NotionBytes = 0
$script:NotionMs = 0
$script:NotionProps = @{ date = '日付'; in = '出勤打刻'; out = '退勤打刻'; kind = '区分'; regular = '定時' }
$script:PaidLeave = '有休'

function Initialize-Notion([string]$Token) {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $script:NotionHeaders = @{ Authorization = "Bearer $Token"; 'Notion-Version' = '2025-09-03' }
}

function Invoke-Notion([string]$Method, [string]$Url, $Body, [string]$Label) {
    $params = @{ Uri = $Url; Method = $Method; Headers = $script:NotionHeaders; UseBasicParsing = $true }
    if ($null -ne $Body) {
        $params.Body = [Text.Encoding]::UTF8.GetBytes(($Body | ConvertTo-Json -Depth 10 -Compress))
        $params.ContentType = 'application/json; charset=utf-8'
    }
    $sw = [Diagnostics.Stopwatch]::StartNew()
    try { $r = Invoke-WebRequest @params }
    catch {
        $status = $null; $detail = $_.ErrorDetails.Message
        try { $status = [int]$_.Exception.Response.StatusCode } catch {}
        $hint = switch ($status) {
            401 { 'APIキーが違うか、無効になっています。run.cmd から入れ直してください' }
            404 { 'データベースが見つかりません。timecard.db にコネクトが接続されているか確認してください' }
            429 { '短い時間に呼びすぎました。少し待ってから実行してください' }
            default { 'ネットワークに接続できているか確認してください' }
        }
        throw ("Notion API でエラー（{0}、{1}）：{2} {3}" -f $Label, $status, $hint, $detail)
    }
    $bytes = $r.RawContentStream.ToArray()
    $sw.Stop()
    $script:NotionBytes += $bytes.Length; $script:NotionMs += $sw.ElapsedMilliseconds
    return ([Text.Encoding]::UTF8.GetString($bytes) | ConvertFrom-Json)
}

# 項目のID（応答を小さくするために使う）。一度取ったら保存しておく
function Get-NotionPropertyIds([string]$DataSourceId) {
    $cache = Join-Path (Get-AppDataDir) ('props_{0}.json' -f $DataSourceId)
    $ids = $null
    if (Test-Path -LiteralPath $cache) {
        $obj = Read-JsonFile $cache
        $ids = @{}
        foreach ($p in $obj.PSObject.Properties) { $ids[$p.Name] = $p.Value }
        foreach ($k in $script:NotionProps.Keys) { if (-not $ids[$k]) { $ids = $null; break } }
    }
    if (-not $ids) {
        $schema = Invoke-Notion 'GET' "https://api.notion.com/v1/data_sources/$DataSourceId" $null 'データベースの項目一覧'
        $ids = @{}
        foreach ($k in $script:NotionProps.Keys) {
            $name = $script:NotionProps[$k]
            $prop = $schema.properties.$name
            if (-not $prop) { throw "Notion の timecard.db に項目「$name」がありません。テンプレートが古い可能性があります" }
            $ids[$k] = $prop.id
        }
        $ids | ConvertTo-Json | Set-Content -LiteralPath $cache -Encoding UTF8
    }
    return $ids
}

function Invoke-NotionQuery([string]$DataSourceId, $PropIds, $Body, [string]$Label) {
    $url = "https://api.notion.com/v1/data_sources/$DataSourceId/query"
    if ($PropIds) { $url += '?' + (($PropIds | ForEach-Object { "filter_properties=$_" }) -join '&') }
    $rows = @(); $page = 1
    while ($true) {
        $res = Invoke-Notion 'POST' $url $Body ("{0}（{1}）" -f $Label, $page)
        $rows += @($res.results)
        if (-not $res.has_more -or $Body.page_size -eq 1) { break }
        $Body.start_cursor = $res.next_cursor; $page++
    }
    return ,$rows
}

function Get-MonthRows([string]$DataSourceId, $Ids, [int]$Year, [int]$Month) {
    $first = Get-Date -Year $Year -Month $Month -Day 1
    $last  = $first.AddMonths(1).AddDays(-1)
    $body = @{
        filter = @{ and = @(
            @{ property = $script:NotionProps.date; date = @{ on_or_after  = $first.ToString('yyyy-MM-dd') } }
            @{ property = $script:NotionProps.date; date = @{ on_or_before = $last.ToString('yyyy-MM-dd') } }
        ) }
        page_size = 100
    }
    return (Invoke-NotionQuery $DataSourceId @($Ids.date, $Ids.in, $Ids.out, $Ids.kind, $Ids.regular) $body ("{0}年{1}月の行" -f $Year, $Month))
}

function Get-RegularText($Row) {
    $f = $Row.properties.($script:NotionProps.regular).formula
    if ($f -and $f.string) { return [string]$f.string }
    return $null
}

# 「定時」を読む。その月に行がなければ、いちばん新しい行から読む
function Get-RegularHours([string]$DataSourceId, $Ids, $MonthRows) {
    $text = $null
    foreach ($r in $MonthRows) { $text = Get-RegularText $r; if ($text) { break } }
    if (-not $text) {
        $body = @{ sorts = @(@{ property = $script:NotionProps.date; direction = 'descending' }); page_size = 1 }
        $latest = Invoke-NotionQuery $DataSourceId @($Ids.regular) $body 'いちばん新しい行'
        if ($latest.Count -gt 0) { $text = Get-RegularText $latest[0] }
    }
    if (-not $text) { throw 'Notion の「定時」が読めません。timecard.db に1行もない場合は、一度ボタンを押して行を作ってください' }
    if ($text -notmatch '^\s*(\d{1,2}):(\d{2})\s*-\s*(\d{1,2}):(\d{2})\s*$') {
        throw "Notion の「定時」の形が正しくありません：「$text」（例：10:00-19:00）"
    }
    $start = [int]$Matches[1] * 60 + [int]$Matches[2]
    $end   = [int]$Matches[3] * 60 + [int]$Matches[4]
    if ($end -le $start) { throw "Notion の「定時」の終わりが始まりより前です：「$text」" }
    return @{ Text = $text.Trim(); Start = $start; End = $end }
}

# 行を日ごとにまとめる（同じ日に2行あれば、出勤は早いほう・退勤は遅いほう）
function ConvertTo-DayRecords($Rows) {
    $P = $script:NotionProps
    $days = @{}
    foreach ($row in $Rows) {
        $pr = $row.properties
        $dv = $pr.($P.date).date
        if (-not $dv -or -not $dv.start) { Write-Warn ("日付が空の行を飛ばしました（{0}）" -f $row.id); continue }
        $date = ConvertTo-DateOnly $dv.start
        $kind = if ($pr.($P.kind).select) { $pr.($P.kind).select.name } else { '' }
        $inM = $null; $outM = $null
        $iv = $pr.($P.in).date;  if ($iv -and $iv.start) { $inM  = [int][Math]::Floor(((ConvertTo-Jst $iv.start) - $date).TotalMinutes) }
        $ov = $pr.($P.out).date; if ($ov -and $ov.start) { $outM = [int][Math]::Floor(((ConvertTo-Jst $ov.start) - $date).TotalMinutes) }
        $d = $date.Day
        if ($days.ContainsKey($d)) {
            Write-Warn ("{0}日の行が2つ以上あります。出勤は早いほう、退勤は遅いほうを使います" -f $d)
            $o = $days[$d]
            if ($null -ne $inM  -and ($null -eq $o.In  -or $inM  -lt $o.In))  { $o.In = $inM }
            if ($null -ne $outM -and ($null -eq $o.Out -or $outM -gt $o.Out)) { $o.Out = $outM }
            if ($kind) { $o.Kind = $kind }
        } else {
            $days[$d] = [pscustomobject]@{ Day = $d; Kind = $kind; In = $inM; Out = $outM }
        }
    }
    return $days
}
