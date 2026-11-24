#requires -version 5.1
# 보세구역장치화물 보수작업 승인(신청)서 작성 프로그램
# - 업체 DB(CSV, 공유폴더)에서 상호 선택 시 성명/주소 자동 입력
# - B/L번호로 UNI-PASS Open API(화물통관진행정보조회)를 호출해 화물정보 자동 입력 (반입신고는 가장 최근 건 기준)
#   ※ 브라우저를 띄우지 않음. [화면에서 가져오기]만 예비용으로 브라우저를 사용
# - 엑셀 서식 파일에 값을 채워 PDF로 저장 (Excel 필요)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
# 직인 이미지의 흰(또는 투명) 여백 잘라내기
Add-Type -ReferencedAssemblies System.Drawing -TypeDefinition @"
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
public static class SealTrim {
    // 직인 이미지: 흰 바탕을 투명하게 바꾸고(서식의 선이 가려지지 않게) 둘레 여백을 잘라 PNG로 저장
    public static string Trim(string src, string dst) {
        using (Image img0 = Image.FromFile(src))
        using (Bitmap bmp = new Bitmap(img0.Width, img0.Height, PixelFormat.Format32bppArgb)) {
            using (Graphics g = Graphics.FromImage(bmp)) { g.Clear(Color.Transparent); g.DrawImage(img0, 0, 0, img0.Width, img0.Height); }
            int w = bmp.Width, h = bmp.Height;
            Rectangle all = new Rectangle(0, 0, w, h);
            BitmapData d = bmp.LockBits(all, ImageLockMode.ReadWrite, PixelFormat.Format32bppArgb);
            int stride = d.Stride; byte[] px = new byte[stride * h];
            Marshal.Copy(d.Scan0, px, 0, px.Length);
            int minX = w, minY = h, maxX = -1, maxY = -1;
            const int WHITE = 235, FADE = 180;   // 235 이상 = 완전 투명, 180~235 = 점점 투명(가장자리 부드럽게)
            for (int y = 0; y < h; y++) {
                int row = y * stride;
                for (int x = 0; x < w; x++) {
                    int i = row + x * 4;
                    int mn = Math.Min(px[i], Math.Min(px[i + 1], px[i + 2]));
                    int a = px[i + 3];
                    if (mn >= WHITE) a = 0;
                    else if (mn > FADE) a = a * (WHITE - mn) / (WHITE - FADE);
                    px[i + 3] = (byte)a;
                    if (a > 40) { if (x < minX) minX = x; if (x > maxX) maxX = x; if (y < minY) minY = y; if (y > maxY) maxY = y; }
                }
            }
            Marshal.Copy(px, 0, d.Scan0, px.Length);
            bmp.UnlockBits(d);
            if (maxX < 0) return src;
            int pad = Math.Max(2, (int)(Math.Max(maxX - minX, maxY - minY) * 0.02));
            minX = Math.Max(0, minX - pad); minY = Math.Max(0, minY - pad);
            maxX = Math.Min(w - 1, maxX + pad); maxY = Math.Min(h - 1, maxY + pad);
            using (Bitmap c = bmp.Clone(new Rectangle(minX, minY, maxX - minX + 1, maxY - minY + 1), PixelFormat.Format32bppArgb)) c.Save(dst, ImageFormat.Png);
            return dst;
        }
    }
}
"@
[System.Windows.Forms.Application]::EnableVisualStyles()

$ScriptDir    = Split-Path -Parent $MyInvocation.MyCommand.Path
$AppDir       = Join-Path $env:APPDATA 'BosuForm'
if (-not (Test-Path $AppDir)) { New-Item -ItemType Directory -Path $AppDir | Out-Null }
$SettingsPath = Join-Path $AppDir 'settings.json'
$UnipassUrl   = 'https://unipass.customs.go.kr/csp/index.do?tgMenuId=MYC_MNU_00000450'
$CargoApiUrl  = 'https://unipass.customs.go.kr:38010/ext/rest/cargCsclPrgsInfoQry/retrieveCargCsclPrgsInfo'
$DefaultApiKey = ''   # 공개 저장소이므로 키는 코드에 넣지 않음 → 각 PC의 [설정]에서 입력(settings.json에 저장)
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
$CdpPort      = 9333
$Utf8Bom      = New-Object System.Text.UTF8Encoding($true)
$UiFont       = New-Object System.Drawing.Font('Malgun Gothic', 9)

# ---------------------------------------------------------------- 공통
function Msg([string]$text, [string]$icon = 'Information') {
    [void][System.Windows.Forms.MessageBox]::Show($text, '보수작업신청서', 'OK', $icon)
}
function Ask([string]$text) {
    return ([System.Windows.Forms.MessageBox]::Show($text, '보수작업신청서', 'YesNo', 'Question') -eq 'Yes')
}
function Pump { [System.Windows.Forms.Application]::DoEvents() }

# ---------------------------------------------------------------- 설정
function Load-Settings {
    $d = [ordered]@{ templatePath = ''; impDbPath = ''; appDbPath = ''; outDir = ''; openPdf = $true; saveXlsx = $false;
                     lastApplicant = ''; purposes = @(); methods = @(); apiKey = '' }
    $s = $null
    if (Test-Path $SettingsPath) {
        try {
            $s = Get-Content -Path $SettingsPath -Raw -Encoding UTF8 | ConvertFrom-Json
            foreach ($k in @($d.Keys)) { if ($s.PSObject.Properties[$k]) { $d[$k] = $s.$k } }
        } catch { }
    }
    # 이전 버전(업체 DB 1개) 설정은 수출입자 DB로 이어서 사용
    if ($s -and $s.PSObject.Properties['dbPath'] -and $s.dbPath -and -not $d.impDbPath) { $d.impDbPath = $s.dbPath }
    if (-not $d.templatePath) {
        $p = Join-Path $ScriptDir '보수작업신청서_서식.xlsx'
        if (Test-Path $p) { $d.templatePath = $p }
    }
    if (-not $d.outDir) { $d.outDir = [Environment]::GetFolderPath('Desktop') }
    if (-not $d.apiKey) { $d.apiKey = $DefaultApiKey }
    $d.purposes = @($d.purposes | Where-Object { $_ })
    $d.methods  = @($d.methods  | Where-Object { $_ })
    return $d
}
function Save-Settings {
    try { [IO.File]::WriteAllText($SettingsPath, ($Settings | ConvertTo-Json -Depth 4), $Utf8Bom) } catch { }
}
$Settings = Load-Settings

function Push-History([string]$key, [string]$value) {
    $value = $value.Trim()
    if (-not $value) { return }
    $list = @($value) + @($Settings[$key] | Where-Object { $_ -ne $value })
    if ($list.Count -gt 20) { $list = $list[0..19] }
    $Settings[$key] = $list
}

# ---------------------------------------------------------------- 업체 DB (CSV: 상호,성명,주소,직인파일)
# 수출(입)자 DB(imp)와 신청인 DB(app)를 따로 관리
$script:Dbs = @{ imp = @(); app = @() }
$DbLabel = @{ imp = '수출(입)자'; app = '신청인' }
function Get-DbPath([string]$kind) { return [string]$Settings["${kind}DbPath"] }
# UNI-PASS API 키: 업체 DB 폴더(공유폴더)의 키 파일을 우선 사용 → 키가 바뀌면 이 파일 하나만 고치면 모든 PC에 적용
$KeyFileName = 'UNI-PASS_API키.txt'
function Get-SharedKeyPath {
    $db = Get-DbPath 'app'
    if (-not $db) { $db = Get-DbPath 'imp' }
    if (-not $db) { return $null }
    return (Join-Path (Split-Path -Parent $db) $KeyFileName)
}
function Read-SharedKey {
    $p = Get-SharedKeyPath
    if ($p -and (Test-Path -LiteralPath $p)) { try { return ([IO.File]::ReadAllText($p)).Trim() } catch { } }
    return ''
}
function Get-ApiKey {
    $k = Read-SharedKey
    if ($k) { return $k }
    return ([string]$Settings.apiKey).Trim()
}
function Read-DbFile([string]$path) {
    if (-not $path -or -not (Test-Path $path)) { return }
    Import-Csv -Path $path -Encoding UTF8 | Where-Object { $_.'상호' -and $_.'상호'.Trim() }
}
function Load-Db {
    foreach ($k in @('imp', 'app')) {
        try { $script:Dbs[$k] = @(Read-DbFile (Get-DbPath $k)) }
        catch { $script:Dbs[$k] = @(); Msg ("$($DbLabel[$k]) DB 파일을 읽지 못했습니다.`n" + $_.Exception.Message) 'Warning' }
    }
    Refresh-PartyCombos
}
function Csv-Quote([string]$s) { if ($null -eq $s) { $s = '' }; return '"' + $s.Replace('"', '""') + '"' }
function Write-DbFile([string]$path, $rows) {
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('상호,성명,주소,직인파일')
    foreach ($r in ($rows | Sort-Object { $_.'상호' })) {
        [void]$sb.AppendLine(((Csv-Quote $r.'상호'.Trim()), (Csv-Quote $r.'성명'), (Csv-Quote $r.'주소'), (Csv-Quote $r.'직인파일')) -join ',')
    }
    [IO.File]::WriteAllText($path, $sb.ToString(), $Utf8Bom)
}
function Find-Company([string]$kind, [string]$name) {
    $n = $name.Trim()
    if (-not $n) { return $null }
    return ($script:Dbs[$kind] | Where-Object { $_.'상호'.Trim() -eq $n } | Select-Object -First 1)
}
# 직인파일: 신청인 DB 폴더 기준 상대경로(예: 직인\회사명.png) 또는 전체 경로
function Resolve-Seal([string]$rel) {
    $rel = $rel.Trim().Trim('"')
    if (-not $rel) { return $null }
    if ([IO.Path]::IsPathRooted($rel)) { if (Test-Path -LiteralPath $rel) { return $rel } else { return $null } }
    # 상대경로·파일 이름만 적은 경우: 신청인 DB 폴더 → 그 안의 "직인" 폴더 순서로 찾음
    $dir = Split-Path -Parent (Get-DbPath 'app')
    foreach ($p in @((Join-Path $dir $rel), (Join-Path (Join-Path $dir '직인') $rel))) { if (Test-Path -LiteralPath $p) { return $p } }
    return $null
}
# 한 업체의 직인파일 값만 바로 파일에 저장 (다른 줄은 파일에 있는 그대로 둠)
function Set-CompanySeal([string]$name, [string]$person, [string]$addr, [string]$rel) {
    $path = Get-DbPath 'app'
    $rows = [System.Collections.ArrayList]@(Read-DbFile $path)
    $hit = $rows | Where-Object { $_.'상호'.Trim() -eq $name.Trim() } | Select-Object -First 1
    if ($hit) {
        if ($hit.PSObject.Properties['직인파일']) { $hit.'직인파일' = $rel } else { $hit | Add-Member -NotePropertyName '직인파일' -NotePropertyValue $rel }
    } else { [void]$rows.Add([pscustomobject]@{ '상호' = $name.Trim(); '성명' = $person; '주소' = $addr; '직인파일' = $rel }) }
    Write-DbFile $path $rows
    Load-Db
}
function Get-SealFor([string]$name) {
    $r = Find-Company 'app' $name
    if (-not $r) { return $null }
    return (Resolve-Seal ([string]$r.'직인파일'))
}
# 이미지를 신청인 DB 폴더의 "직인" 폴더로 복사하고 상대경로를 돌려줌 (팀원 PC에서도 같이 쓰도록)
function Import-SealFile([string]$name, [string]$src) {
    $dir = Join-Path (Split-Path -Parent (Get-DbPath 'app')) '직인'
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
    $fn = (($name -replace '[\\/:*?"<>|]', '_').Trim()) + [IO.Path]::GetExtension($src).ToLower()
    $dst = Join-Path $dir $fn
    if ([IO.Path]::GetFullPath($src) -ne [IO.Path]::GetFullPath($dst)) { Copy-Item -LiteralPath $src -Destination $dst -Force }
    return ('직인\' + $fn)
}
# 공유폴더 동시 사용을 고려해 저장 직전에 파일을 다시 읽고 해당 업체만 추가/수정
function Upsert-Company([string]$kind, [string]$name, [string]$person, [string]$addr) {
    $path = Get-DbPath $kind
    $rows = [System.Collections.ArrayList]@(Read-DbFile $path)
    $hit = $rows | Where-Object { $_.'상호'.Trim() -eq $name.Trim() } | Select-Object -First 1
    if ($hit) { $hit.'성명' = $person; $hit.'주소' = $addr }
    else { [void]$rows.Add([pscustomobject]@{ '상호' = $name.Trim(); '성명' = $person; '주소' = $addr; '직인파일' = '' }) }
    Write-DbFile $path $rows
    Load-Db
}

# ---------------------------------------------------------------- 브라우저(UNI-PASS) 연결 - Chrome DevTools Protocol
$script:Ws = $null; $script:WsTarget = ''; $script:CdpId = 0

function New-WebClient {
    $wc = New-Object System.Net.WebClient
    $wc.Proxy = $null
    $wc.Encoding = [System.Text.Encoding]::UTF8
    return $wc
}
function Get-Targets {
    try {
        $json = (New-WebClient).DownloadString("http://127.0.0.1:$CdpPort/json/list")
        $list = ConvertFrom-Json $json
        return ,@($list | ForEach-Object { $_ })
    } catch { return $null }
}
function Find-Browser {
    $c = @(
        (Join-Path ${env:ProgramFiles(x86)} 'Microsoft\Edge\Application\msedge.exe'),
        (Join-Path $env:ProgramFiles 'Microsoft\Edge\Application\msedge.exe'),
        (Join-Path $env:LOCALAPPDATA 'Microsoft\Edge\Application\msedge.exe'),
        (Join-Path $env:ProgramFiles 'Google\Chrome\Application\chrome.exe'),
        (Join-Path ${env:ProgramFiles(x86)} 'Google\Chrome\Application\chrome.exe'),
        (Join-Path $env:LOCALAPPDATA 'Google\Chrome\Application\chrome.exe'))
    foreach ($p in $c) { if ($p -and (Test-Path $p)) { return $p } }
    throw 'Edge 또는 Chrome을 찾지 못했습니다.'
}
function Close-Ws {
    if ($script:Ws) { try { $script:Ws.Dispose() } catch { } }
    $script:Ws = $null; $script:WsTarget = ''
}
function Connect-Browser {
    $t = Get-Targets
    if ($null -eq $t) {
        $exe  = Find-Browser
        $prof = Join-Path $AppDir 'browser'
        Start-Process -FilePath $exe -ArgumentList @("--remote-debugging-port=$CdpPort", "--user-data-dir=`"$prof`"",
            '--no-first-run', '--no-default-browser-check', '--new-window', $UnipassUrl)
        for ($i = 0; $i -lt 60; $i++) {
            Start-Sleep -Milliseconds 250; Pump
            $t = Get-Targets
            if ($null -ne $t -and ($t | Where-Object { $_.type -eq 'page' })) { break }
        }
        if ($null -eq $t) { throw '조회용 브라우저를 열지 못했습니다.' }
    }
    $page = $t | Where-Object { $_.type -eq 'page' -and $_.url -like '*unipass.customs.go.kr*' } | Select-Object -First 1
    if (-not $page) { $page = $t | Where-Object { $_.type -eq 'page' } | Select-Object -First 1 }
    if (-not $page) {
        $page = (New-WebClient).UploadString("http://127.0.0.1:$CdpPort/json/new?$UnipassUrl", 'PUT', '') | ConvertFrom-Json
    }
    if ($script:Ws -and $script:Ws.State -eq 'Open' -and $script:WsTarget -eq $page.id) { return }
    Close-Ws
    $ws = New-Object System.Net.WebSockets.ClientWebSocket
    try { $ws.Options.Proxy = $null } catch { }
    $task = $ws.ConnectAsync([Uri]$page.webSocketDebuggerUrl, [Threading.CancellationToken]::None)
    $deadline = (Get-Date).AddSeconds(10)
    while (-not $task.Wait(100)) { Pump; if ((Get-Date) -gt $deadline) { throw '브라우저 연결 시간 초과' } }
    $script:Ws = $ws; $script:WsTarget = $page.id
}
function Receive-Ws([datetime]$deadline) {
    $buf = New-Object byte[] 65536
    $ms  = New-Object System.IO.MemoryStream
    do {
        $seg  = New-Object 'System.ArraySegment[byte]' -ArgumentList @(, $buf)
        $task = $script:Ws.ReceiveAsync($seg, [Threading.CancellationToken]::None)
        while (-not $task.Wait(100)) {
            Pump
            if ((Get-Date) -gt $deadline) { Close-Ws; throw '브라우저 응답 시간 초과' }
        }
        $r = $task.Result
        if ($r.MessageType -eq 'Close') { Close-Ws; throw '브라우저 연결이 끊어졌습니다.' }
        $ms.Write($buf, 0, $r.Count)
    } while (-not $r.EndOfMessage)
    return [System.Text.Encoding]::UTF8.GetString($ms.ToArray())
}
function Invoke-Cdp([string]$method, $params = @{}, [int]$timeoutSec = 30) {
    $id = ++$script:CdpId
    $msg = @{ id = $id; method = $method; params = $params } | ConvertTo-Json -Depth 6 -Compress
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($msg)
    $seg = New-Object 'System.ArraySegment[byte]' -ArgumentList @(, $bytes)
    [void]$script:Ws.SendAsync($seg, [System.Net.WebSockets.WebSocketMessageType]::Text, $true, [Threading.CancellationToken]::None).Wait(5000)
    $deadline = (Get-Date).AddSeconds($timeoutSec)
    while ($true) {
        $o = (Receive-Ws $deadline) | ConvertFrom-Json
        if ($o.id -eq $id) {
            if ($o.error) { throw ('브라우저 오류: ' + $o.error.message) }
            return $o.result
        }
    }
}
function Invoke-Js([string]$expr, [int]$timeoutSec = 40) {
    $r = Invoke-Cdp 'Runtime.evaluate' @{ expression = $expr; awaitPromise = $true; returnByValue = $true } $timeoutSec
    if ($r.exceptionDetails) { throw ('스크립트 오류: ' + $r.exceptionDetails.text) }
    return $r.result.value
}
function Test-PageReady {
    try { return [bool](Invoke-Js "!!document.getElementById('MYC0405101Q_hblNoTab1')" 10) } catch { return $false }
}
function Ensure-UnipassPage {
    if (Test-PageReady) { return }
    try { [void](Invoke-Cdp 'Page.navigate' @{ url = $UnipassUrl }) } catch { }
    for ($i = 0; $i -lt 40; $i++) {
        Start-Sleep -Milliseconds 500; Pump
        if (Test-PageReady) { return }
    }
    throw 'UNI-PASS 수입화물진행정보 화면을 열지 못했습니다.'
}

$LookupJs = @'
(async()=>{
const BL='__BL__', YY='__YY__', KIND='__KIND__';
const sleep=ms=>new Promise(r=>setTimeout(r,ms));
const $=i=>document.getElementById(i);
const clean=s=>(s||'').replace(/\s+/g,' ').trim();
function extract(){
  const g=k=>clean(($('MYC0405102Q_'+k)||{}).textContent);
  let box=$('MYC0405102Q_printResultListL');
  if(!box||!box.querySelector('td[name=brngBssNo]')) box=$('MYC0405102Q_resultListL');
  const hist=box?[...box.querySelectorAll('td[name=brngBssNo]')].map(td=>{
    const tr=td.parentElement, nx=tr.nextElementSibling;
    return {step:clean(tr.cells[1]&&tr.cells[1].textContent),
            sgn:clean((tr.querySelector('[name=snarSgn]')||{}).textContent),
            nm:clean(nx&&(nx.querySelector('[name=snarKoreNm]')||{}).textContent),
            no:clean(td.textContent),
            dt:clean(nx&&nx.cells[0]&&nx.cells[0].textContent)};
  }):[];
  const ban=hist.find(h=>/^반입신고/.test(h.step))||null;
  return {state:'found',cargMtNo:g('cargMtNo'),mbl:g('mblNo'),hbl:g('hblNo'),sanm:g('sanm'),
          etprDt:g('etprDt'),prnm:g('prnm'),pkg:g('cmdtGcnt'),wght:g('cmdtWght'),ban:ban,histCount:hist.length};
}
if(KIND==='GRAB'){
  const c=$('MYC0405102Q_cargMtNo');
  if(!c||!clean(c.textContent)) return JSON.stringify({state:'nodetail'});
  return JSON.stringify(extract());
}
if(!$('MYC0405101Q_hblNoTab1')) return JSON.stringify({state:'notready'});
const cnt=document.querySelector('strong[data-recordCountTab1]');
const old=$('MYC0405102Q_cargMtNo'); if(old) old.textContent='';
if(cnt) cnt.textContent='-';
$('MYC0405101Q_qryTpValue2Tab1').click();
$('MYC0405101Q_mblNoTab1').value=(KIND==='M')?BL:'';
$('MYC0405101Q_hblNoTab1').value=(KIND==='H')?BL:'';
const sel=$('MYC0405101Q_blYyTab1'); sel.value=YY; sel.dispatchEvent(new Event('change',{bubbles:true}));
$('MYC0405101Q_searchBtnTab1').click();
for(let i=0;i<60;i++){
  await sleep(400);
  const c=$('MYC0405102Q_cargMtNo');
  if(c&&clean(c.textContent)){
    for(let j=0;j<15;j++){ const b=$('MYC0405102Q_printResultListL')||$('MYC0405102Q_resultListL'); if(b&&b.querySelector('td[name=brngBssNo]')) break; await sleep(300); }
    return JSON.stringify(extract());
  }
  const n=cnt?clean(cnt.textContent):'';
  if(n&&n!=='-'){ const k=parseInt(n.replace(/,/g,''),10); if(k===0) return JSON.stringify({state:'notfound'}); if(k>1) return JSON.stringify({state:'multi',count:k}); }
}
return JSON.stringify({state:'timeout'});
})()
'@

# ---------------------------------------------------------------- UNI-PASS Open API (화물통관진행정보조회)
function Get-XmlText($node, [string]$name) {
    $c = $node.SelectSingleNode($name)
    if ($c) { return $c.InnerText.Trim() } else { return '' }
}
function Format-Num([string]$s) {
    $v = 0.0
    if ([double]::TryParse(($s -replace ',', ''), [ref]$v)) { return $v.ToString('#,0.###') }
    return $s
}
function Format-ApiDate([string]$s) {
    # 20261120 → 2026-11-20, 20261120153000 → 2026-11-20 15:30
    if ($s -match '^(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})') { return "$($Matches[1])-$($Matches[2])-$($Matches[3]) $($Matches[4]):$($Matches[5])" }
    if ($s -match '^(\d{4})(\d{2})(\d{2})$') { return "$($Matches[1])-$($Matches[2])-$($Matches[3])" }
    return $s
}
# 반환: @{ count = 건수; vos = 화물 요약 노드들; dtls = 처리이력 노드들 }
function Invoke-CargoApi([hashtable]$q) {
    $key = Get-ApiKey
    if (-not $key) { throw "UNI-PASS API 키가 없습니다.`n[설정] 창의 'UNI-PASS API 키' 칸에 인증키를 입력하고 저장하세요." }
    $qs = @("crkyCn=$([Uri]::EscapeDataString($key))")
    foreach ($k in $q.Keys) { $qs += "$k=$([Uri]::EscapeDataString([string]$q[$k]))" }
    $wc = New-Object System.Net.WebClient
    $wc.Encoding = [System.Text.Encoding]::UTF8
    if ($wc.Proxy) { $wc.Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials }
    $task = $wc.DownloadStringTaskAsync("$($CargoApiUrl)?" + ($qs -join '&'))
    $deadline = (Get-Date).AddSeconds(20)
    while (-not $task.IsCompleted) {
        Start-Sleep -Milliseconds 50; Pump
        if ((Get-Date) -gt $deadline) { $wc.CancelAsync(); throw 'UNI-PASS API 응답이 없습니다. 잠시 후 다시 시도하세요.' }
    }
    if ($task.IsFaulted) { throw ('UNI-PASS API 연결 실패: ' + $task.Exception.GetBaseException().Message) }
    $x = New-Object System.Xml.XmlDocument
    $x.LoadXml($task.Result)
    $root = $x.DocumentElement
    $ntce = Get-XmlText $root 'ntceInfo'
    $cnt = 0; [void][int]::TryParse((Get-XmlText $root 'tCnt'), [ref]$cnt)
    $vos = @($root.SelectNodes('cargCsclPrgsInfoQryVo'))
    # 여러 건이면 "[N00] 조회 결과가 다건입니다..." 안내와 함께 화물 목록만 옴 → 오류 아님
    if ($ntce -and $vos.Count -eq 0) {
        $hint = $(if ($ntce -match '인증키') { "`n`n인증키가 만료되었거나 바뀌었다면 [설정]에서 새 키를 입력하고`n'공유폴더에 저장'을 체크해 저장하세요. (모든 PC에 적용)" } else { '' })
        throw ('UNI-PASS API: ' + $ntce + $hint)
    }
    return @{ count = $cnt
              vos   = $vos
              dtls  = @($root.SelectNodes('cargCsclPrgsInfoDtlQryVo')) }
}
# API 응답 → 화면 입력용 결과 (기존 Apply-Result 형식과 동일)
function Convert-CargoResult($vo, $dtls) {
    $pkg = Get-XmlText $vo 'pckGcnt'
    $pkg = $(if ($pkg) { ((Format-Num $pkg) + ' ' + (Get-XmlText $vo 'pckUt')).Trim() } else { '' })
    $wg = Get-XmlText $vo 'ttwg'
    $wg = $(if ($wg) { ((Format-Num $wg) + ' ' + (Get-XmlText $vo 'wghtUt')).Trim() } else { '' })
    # 처리이력 중 반입신고: 처리일시가 가장 최근인 건
    $ban = $dtls | Where-Object { (Get-XmlText $_ 'cargTrcnRelaBsopTpcd') -like '반입신고*' } |
           Sort-Object { Get-XmlText $_ 'prcsDttm' } -Descending | Select-Object -First 1
    $banObj = $null
    if ($ban) {
        $no = Get-XmlText $ban 'dclrNo'
        if (-not $no) { $no = Get-XmlText $ban 'rlbrBssNo' }
        $banObj = [pscustomobject]@{ sgn = (Get-XmlText $ban 'shedSgn'); nm = (Get-XmlText $ban 'shedNm'); no = $no
                                     dt = (Format-ApiDate (Get-XmlText $ban 'prcsDttm')) }
    }
    return [pscustomobject]@{
        state = 'found'; cargMtNo = (Get-XmlText $vo 'cargMtNo'); mbl = (Get-XmlText $vo 'mblNo'); hbl = (Get-XmlText $vo 'hblNo')
        sanm = (Get-XmlText $vo 'shipNm'); etprDt = (Format-ApiDate (Get-XmlText $vo 'etprDt')); prnm = (Get-XmlText $vo 'prnm')
        pkg = $pkg; wght = $wg; ban = $banObj; histCount = @($dtls).Count }
}
# 조회 결과가 여러 건일 때 고르는 창 → 선택한 화물관리번호 반환 (취소 시 $null)
function Select-Cargo($vos) {
    $f = New-Object System.Windows.Forms.Form
    $f.Text = '화물관리번호 선택'; $f.Font = $UiFont; $f.StartPosition = 'CenterParent'
    $f.ClientSize = New-Object System.Drawing.Size(1110, 330); $f.MinimizeBox = $false; $f.MaximizeBox = $false
    $lb = New-Object System.Windows.Forms.Label
    $lb.Text = "이 B/L로 화물관리번호가 $(@($vos).Count)건 조회되었습니다. 줄을 클릭하면 선명·포장·품명·장치장소를 불러옵니다. 해당 화물을 더블클릭하거나 선택 후 [선택]을 누르세요. (입항일 최신순)"
    $lb.Location = New-Object System.Drawing.Point(10, 10); $lb.Size = New-Object System.Drawing.Size(1090, 18); $f.Controls.Add($lb)
    $lv = New-Object System.Windows.Forms.ListView
    $lv.View = 'Details'; $lv.FullRowSelect = $true; $lv.MultiSelect = $false; $lv.HideSelection = $false; $lv.GridLines = $true
    $lv.Location = New-Object System.Drawing.Point(10, 32); $lv.Size = New-Object System.Drawing.Size(1090, 250)
    # 0 화물관리번호 1 MBL 2 HBL 3 입항일 4 양륙항 5 선사 | 6~9는 줄을 클릭하면 상세 조회로 채움
    foreach ($c in @(@('화물관리번호', 150), @('Master B/L', 120), @('House B/L', 105), @('입항일', 78), @('양륙항', 60), @('선사', 110),
                     @('선(기)명', 115), @('포장 / 중량', 120), @('장치장소', 110), @('품명', 200))) {
        [void]$lv.Columns.Add($c[0], $c[1])
    }
    foreach ($v in ($vos | Sort-Object { Get-XmlText $_ 'etprDt' } -Descending)) {
        $it = New-Object System.Windows.Forms.ListViewItem((Get-XmlText $v 'cargMtNo'))
        foreach ($n in @('mblNo', 'hblNo')) { [void]$it.SubItems.Add((Get-XmlText $v $n)) }
        [void]$it.SubItems.Add((Format-ApiDate (Get-XmlText $v 'etprDt')))
        foreach ($n in @('dsprNm', 'shcoFlco', 'shipNm')) { [void]$it.SubItems.Add((Get-XmlText $v $n)) }
        foreach ($i in 1..3) { [void]$it.SubItems.Add('') }
        $it.Tag = [pscustomobject]@{ mt = (Get-XmlText $v 'cargMtNo'); loaded = $false }
        [void]$lv.Items.Add($it)
    }
    # 줄 선택 시 해당 화물 상세를 조회해 빈 칸 채우기 (한 번 불러온 줄은 다시 조회하지 않음)
    $script:SelLoading = $false
    $lv.Add_SelectedIndexChanged({
        if ($script:SelLoading -or $lv.SelectedItems.Count -eq 0) { return }
        $it = $lv.SelectedItems[0]
        if ($it.Tag.loaded) { return }
        $script:SelLoading = $true
        $f.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            $it.SubItems[9].Text = '불러오는 중...'; Pump
            $d = Invoke-CargoApi @{ cargMtNo = $it.Tag.mt }
            if ($d.vos.Count -gt 0) {
                $r = Convert-CargoResult $d.vos[0] $d.dtls
                $it.SubItems[6].Text = $r.sanm
                $it.SubItems[7].Text = (@($r.pkg, $r.wght) | Where-Object { $_ }) -join ' / '
                $it.SubItems[8].Text = $(if ($r.ban) { $r.ban.nm } else { '' })
                $it.SubItems[9].Text = $r.prnm
                $it.Tag.loaded = $true
            } else { $it.SubItems[9].Text = '' }
        } catch { $it.SubItems[9].Text = '(불러오기 실패)' }
        finally { $f.Cursor = [System.Windows.Forms.Cursors]::Default; $script:SelLoading = $false }
    })
    $f.Controls.Add($lv)
    $ok = New-Object System.Windows.Forms.Button; $ok.Text = '선택'; $ok.Location = New-Object System.Drawing.Point(924, 291); $ok.Size = New-Object System.Drawing.Size(85, 30); $f.Controls.Add($ok)
    $cancel = New-Object System.Windows.Forms.Button; $cancel.Text = '취소'; $cancel.Location = New-Object System.Drawing.Point(1015, 291); $cancel.Size = New-Object System.Drawing.Size(85, 30); $f.Controls.Add($cancel)
    $f.Add_Shown({ if ($lv.Items.Count) { $lv.Items[0].Selected = $true; $lv.Items[0].Focused = $true }; $lv.Focus() })
    $f.AcceptButton = $ok; $f.CancelButton = $cancel
    $ok.Add_Click({ $f.DialogResult = 'OK'; $f.Close() })
    $cancel.Add_Click({ $f.DialogResult = 'Cancel'; $f.Close() })
    $lv.Add_DoubleClick({ $f.DialogResult = 'OK'; $f.Close() })
    if ($f.ShowDialog($form) -eq 'OK' -and $lv.SelectedItems.Count) { return [string]$lv.SelectedItems[0].Tag.mt }
    return $null
}

function Run-Lookup([string]$bl, [string]$year, [string]$kind) {
    $js = $LookupJs.Replace('__BL__', $bl).Replace('__YY__', $year).Replace('__KIND__', $kind)
    $txt = Invoke-Js $js 45
    if (-not $txt) { return [pscustomobject]@{ state = 'timeout' } }
    return ($txt | ConvertFrom-Json)
}

# ---------------------------------------------------------------- 엑셀 → PDF
function Text-Width([string]$s) {
    $w = 0
    foreach ($ch in $s.ToCharArray()) { if ([int]$ch -gt 127) { $w += 2 } else { $w += 1 } }
    return $w
}
# 칸 너비 대비 글자가 많으면 줄바꿈 + 글자 축소
$FitLimit = @{ 'F6' = 34; 'S6' = 34; 'F7' = 34; 'S7' = 34; 'F8' = 36; 'S8' = 36;
               'A10' = 22; 'G10' = 26; 'R9' = 36; 'A12' = 22; 'G12' = 26; 'R11' = 36;
               'A14' = 22; 'G14' = 26; 'R13' = 16; 'W13' = 16;
               'A17' = 22; 'G17' = 18; 'L17' = 18; 'Q17' = 18; 'V17' = 18;
               'A22' = 26; 'H22' = 20; 'N22' = 20; 'T22' = 26; 'Q31' = 40 }

function Export-Form($vals, [string]$pdfPath, [string]$xlsxPath, [string]$sealPath) {
    $xl = $null; $wb = $null
    try {
        $xl = New-Object -ComObject Excel.Application
        $xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false
        $wb = $xl.Workbooks.Open($Settings.templatePath, 0, $true)
        $ws = $null
        foreach ($s in $wb.Worksheets) { if ($s.Name -eq '보수작업신청서') { $ws = $s } }
        if (-not $ws) { $ws = $wb.Worksheets.Item(1) }
        foreach ($k in $vals.Keys) {
            $v = [string]$vals[$k]
            if ($null -eq $v -or $v -eq '') { continue }
            $c = $ws.Range($k)
            $c.NumberFormat = '@'
            $c.Value2 = $v
            if ($FitLimit.ContainsKey($k)) {
                $w = Text-Width $v; $lim = $FitLimit[$k]
                if ($w -gt $lim) {
                    $area = $c.MergeArea
                    $area.WrapText = $true
                    $area.Font.Size = $(if ($w -gt $lim * 2) { 7 } else { 8 })
                }
            }
            if ($k -in @('Q17', 'V17', 'N22', 'T22')) { $c.MergeArea.HorizontalAlignment = -4108 }
            if ($k -eq 'H29') { $c.HorizontalAlignment = -4152 }
        }
        if ($sealPath) {
            # 명판직인: 흰 바탕을 투명하게 하고 둘레 여백을 잘라낸 뒤, '신청인' 오른쪽 칸(Q~Z열, 27~33행) 안에 최대한 크게 배치
            $img = $sealPath
            try { $img = [SealTrim]::Trim($sealPath, (Join-Path $env:TEMP ('seal_' + [guid]::NewGuid().ToString('N') + '.png'))) } catch { }
            $boxL = [double]$ws.Range('Q31').Left + 6
            $boxR = [double]$ws.Range('Z31').Left + [double]$ws.Range('Z31').Width - 6
            $boxT = [double]$ws.Range('Q27').Top + 2
            $boxB = [double]$ws.Range('Q34').Top - 3
            $pic = $ws.Shapes.AddPicture($img, 0, -1, $boxL, $boxT, -1, -1)
            $pic.LockAspectRatio = -1
            $ratio = [double]$pic.Width / [double]$pic.Height
            $h = $boxB - $boxT; $w = $h * $ratio
            if ($w -gt ($boxR - $boxL)) { $w = $boxR - $boxL; $h = $w / $ratio }
            $pic.Height = $h
            $pic.Left = $boxL + (($boxR - $boxL) - [double]$pic.Width) / 2
            $pic.Top  = $boxT + (($boxB - $boxT) - [double]$pic.Height) / 2
            $pic.Placement = 3
            if ($img -ne $sealPath) { $script:TrimTemp = $img }
        }
        $ws.ExportAsFixedFormat(0, $pdfPath)
        if ($xlsxPath) { $wb.SaveCopyAs($xlsxPath) }
    } finally {
        if ($script:TrimTemp) { Remove-Item -LiteralPath $script:TrimTemp -ErrorAction SilentlyContinue; $script:TrimTemp = $null }
        if ($wb) { try { $wb.Close($false) } catch { } }
        if ($xl) { try { $xl.Quit() } catch { } }
        foreach ($o in @($ws, $wb, $xl)) { if ($o) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch { } } }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}

# ---------------------------------------------------------------- 화면 구성
$form = New-Object System.Windows.Forms.Form
$form.Text = '보수작업 승인(신청)서 작성'
$VerFile = Join-Path $ScriptDir 'version.txt'
if (Test-Path $VerFile) { $form.Text += '  (버전 ' + ([IO.File]::ReadAllText($VerFile)).Trim() + ')' }
$form.Font = $UiFont
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedSingle'
$form.MaximizeBox = $false
$form.ClientSize = New-Object System.Drawing.Size(860, 618)

function New-Group([string]$text, [int]$x, [int]$y, [int]$w, [int]$h) {
    $g = New-Object System.Windows.Forms.GroupBox
    $g.Text = $text; $g.Location = New-Object System.Drawing.Point($x, $y); $g.Size = New-Object System.Drawing.Size($w, $h)
    $form.Controls.Add($g); return $g
}
function Add-Label($parent, [string]$text, [int]$x, [int]$y, [int]$w = 60) {
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $text; $l.Location = New-Object System.Drawing.Point($x, ($y + 3)); $l.Size = New-Object System.Drawing.Size($w, 18)
    $parent.Controls.Add($l); return $l
}
function Add-Text($parent, [int]$x, [int]$y, [int]$w) {
    $t = New-Object System.Windows.Forms.TextBox
    $t.Location = New-Object System.Drawing.Point($x, $y); $t.Size = New-Object System.Drawing.Size($w, 23)
    $parent.Controls.Add($t); return $t
}
function Add-Combo($parent, [int]$x, [int]$y, [int]$w, [switch]$List) {
    $c = New-Object System.Windows.Forms.ComboBox
    $c.Location = New-Object System.Drawing.Point($x, $y); $c.Size = New-Object System.Drawing.Size($w, 23)
    if ($List) { $c.DropDownStyle = 'DropDownList' } else { $c.DropDownStyle = 'DropDown' }
    $parent.Controls.Add($c); return $c
}
function Add-Date($parent, [int]$x, [int]$y, [int]$w = 120) {
    $d = New-Object System.Windows.Forms.DateTimePicker
    $d.Location = New-Object System.Drawing.Point($x, $y); $d.Size = New-Object System.Drawing.Size($w, 23)
    $d.Format = 'Custom'; $d.CustomFormat = 'yyyy-MM-dd'
    $parent.Controls.Add($d); return $d
}
function Add-Button($parent, [string]$text, [int]$x, [int]$y, [int]$w, [int]$h = 27) {
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $text; $b.Location = New-Object System.Drawing.Point($x, $y); $b.Size = New-Object System.Drawing.Size($w, $h)
    $parent.Controls.Add($b); return $b
}

# 수출입자 / 신청인
$gImp = New-Group '수출(입)자' 10 8 415 115
Add-Label $gImp '상   호' 10 22 | Out-Null; $cImpName = Add-Combo $gImp 70 22 330
Add-Label $gImp '성   명' 10 52 | Out-Null; $tImpPerson = Add-Text $gImp 70 52 330
Add-Label $gImp '주   소' 10 82 | Out-Null; $tImpAddr = Add-Text $gImp 70 82 330

$gApp = New-Group '신청인' 435 8 415 115
Add-Label $gApp '상   호' 10 22 | Out-Null; $cAppName = Add-Combo $gApp 70 22 330
Add-Label $gApp '성   명' 10 52 | Out-Null; $tAppPerson = Add-Text $gApp 70 52 330
Add-Label $gApp '주   소' 10 82 | Out-Null; $tAppAddr = Add-Text $gApp 70 82 330

foreach ($c in @($cImpName, $cAppName)) { $c.AutoCompleteMode = 'SuggestAppend'; $c.AutoCompleteSource = 'ListItems' }

# 화물정보
$gCargo = New-Group '화물정보 (UNI-PASS 조회)' 10 128 840 182
Add-Label $gCargo 'B/L번호' 10 22 | Out-Null; $tBL = Add-Text $gCargo 70 22 170
$tBL.CharacterCasing = 'Upper'
$cYear = Add-Combo $gCargo 246 22 70 -List
$thisYear = (Get-Date).Year
for ($y = $thisYear + 1; $y -ge $thisYear - 5; $y--) { [void]$cYear.Items.Add([string]$y) }
$cYear.SelectedItem = [string]$thisYear
$cKind = Add-Combo $gCargo 322 22 100 -List
[void]$cKind.Items.AddRange(@('자동(H→M)', 'House B/L', 'Master B/L')); $cKind.SelectedIndex = 0
$bLookup = Add-Button $gCargo 'UNI-PASS 조회' 430 20 120
$bGrab   = Add-Button $gCargo '화면에서 가져오기' 556 20 130
$lStatus = Add-Label $gCargo '' 692 22 140
$lStatus.ForeColor = [System.Drawing.Color]::DimGray

Add-Label $gCargo '입항일자' 10 56 | Out-Null;        $tArrive = Add-Text $gCargo 70 56 150
Add-Label $gCargo '선(기)명' 240 56 | Out-Null;       $tShip = Add-Text $gCargo 300 56 170
Add-Label $gCargo '내외국물품' 490 56 70 | Out-Null;  $cGoodsTp = Add-Combo $gCargo 565 56 120
[void]$cGoodsTp.Items.AddRange(@('외국물품', '내국물품')); $cGoodsTp.Text = '외국물품'
Add-Label $gCargo '장치장소' 10 88 | Out-Null;        $tStore = Add-Text $gCargo 70 88 400
Add-Label $gCargo '반입신고번호' 485 88 80 | Out-Null; $tInNo = Add-Text $gCargo 565 88 260
Add-Label $gCargo '포   장' 10 120 | Out-Null;        $tPkg = Add-Text $gCargo 70 120 150
Add-Label $gCargo '품   명' 240 120 | Out-Null;       $tGoods = Add-Text $gCargo 300 120 525
$lBanInfo = Add-Label $gCargo '' 70 150 750
$lBanInfo.ForeColor = [System.Drawing.Color]::DimGray

# 작업내용
$gWork = New-Group '작업내용 / 물품' 10 315 840 118
Add-Label $gWork '작업목적' 10 22 | Out-Null;  $cPurpose = Add-Combo $gWork 70 22 340
Add-Label $gWork '작업방법' 425 22 | Out-Null; $cMethod = Add-Combo $gWork 485 22 340
Add-Label $gWork '작업기간' 10 54 | Out-Null;  $dFrom = Add-Date $gWork 70 54
Add-Label $gWork '~' 196 54 15 | Out-Null;     $dTo = Add-Date $gWork 214 54
Add-Label $gWork '신청일자' 425 54 | Out-Null; $dApply = Add-Date $gWork 485 54
Add-Label $gWork '규   격' 10 86 | Out-Null;   $tSpec = Add-Text $gWork 70 86 200
Add-Label $gWork '수   량' 290 86 | Out-Null;  $tQty = Add-Text $gWork 350 86 160
Add-Label $gWork '가   격' 530 86 | Out-Null;  $tPrice = Add-Text $gWork 590 86 235

# 사용재료
$gMat = New-Group '사용 재료 명세' 10 438 840 58
Add-Label $gMat '품   명' 10 22 | Out-Null;  $tMName = Add-Text $gMat 70 22 220
Add-Label $gMat '규격' 305 22 35 | Out-Null; $tMSpec = Add-Text $gMat 342 22 150
Add-Label $gMat '수량' 505 22 35 | Out-Null; $tMQty = Add-Text $gMat 542 22 110
Add-Label $gMat '가격' 665 22 35 | Out-Null; $tMPrice = Add-Text $gMat 702 22 123

# 하단
$chkOpen = New-Object System.Windows.Forms.CheckBox
$chkOpen.Text = 'PDF 바로 열기'; $chkOpen.Location = New-Object System.Drawing.Point(15, 508); $chkOpen.AutoSize = $true
$chkOpen.Checked = [bool]$Settings.openPdf; $form.Controls.Add($chkOpen)
$chkXlsx = New-Object System.Windows.Forms.CheckBox
$chkXlsx.Text = '엑셀 파일도 저장'; $chkXlsx.Location = New-Object System.Drawing.Point(130, 508); $chkXlsx.AutoSize = $true
$chkXlsx.Checked = [bool]$Settings.saveXlsx; $form.Controls.Add($chkXlsx)

$bDbImp = Add-Button $form '수출입자 DB' 270 503 100 30
$bDbApp = Add-Button $form '신청인 DB' 375 503 100 30
$bSet   = Add-Button $form '설정' 480 503 70 30
$bNew   = Add-Button $form '새로 작성' 555 503 85 30
$bPrev  = Add-Button $form '미리보기' 645 503 95 30
$bPdf   = Add-Button $form 'PDF 생성' 745 503 105 30
$bPdf.Font = New-Object System.Drawing.Font('Malgun Gothic', 10, [System.Drawing.FontStyle]::Bold)

$lPaths = New-Object System.Windows.Forms.Label
$lPaths.Location = New-Object System.Drawing.Point(12, 542); $lPaths.Size = New-Object System.Drawing.Size(838, 72)
$lPaths.ForeColor = [System.Drawing.Color]::Gray
$form.Controls.Add($lPaths)

function Update-PathLabel {
    $lPaths.Text = "서식: $($Settings.templatePath)`n수출입자 DB: $($Settings.impDbPath)  ($($script:Dbs.imp.Count)개)`n신청인 DB: $($Settings.appDbPath)  ($($script:Dbs.app.Count)개)`n저장 폴더: $($Settings.outDir)"
}
function Refresh-PartyCombos {
    foreach ($pair in @(@($cImpName, 'imp'), @($cAppName, 'app'))) {
        $c = $pair[0]; $keep = $c.Text
        $c.BeginUpdate(); $c.Items.Clear()
        foreach ($r in ($script:Dbs[$pair[1]] | Sort-Object { $_.'상호' })) { [void]$c.Items.Add($r.'상호') }
        $c.EndUpdate(); $c.Text = $keep
    }
    Update-PathLabel
    Update-SealHint
}
function Refresh-HistoryCombos {
    foreach ($pair in @(@($cPurpose, 'purposes'), @($cMethod, 'methods'))) {
        $c = $pair[0]; $keep = $c.Text
        $c.Items.Clear(); foreach ($v in $Settings[$pair[1]]) { [void]$c.Items.Add($v) }
        $c.Text = $keep
    }
}
function Fill-Party([string]$kind, $combo, $tPerson, $tAddr) {
    $r = Find-Company $kind $combo.Text
    if ($r) { $tPerson.Text = $r.'성명'; $tAddr.Text = $r.'주소' }
    Update-SealHint
}
function Update-SealHint {
    if (Get-SealFor $cAppName.Text) { $gApp.Text = '신청인  (명판직인 등록됨 - PDF에 자동 삽입)'; $gApp.ForeColor = [System.Drawing.Color]::DarkGreen }
    else { $gApp.Text = '신청인'; $gApp.ForeColor = [System.Drawing.SystemColors]::ControlText }
}
# 작업기간 = 신청일자 ~ 신청일자+7일 (신청일자를 바꾸면 같이 바뀜)
function Sync-WorkPeriod {
    $dFrom.Value = $dApply.Value.Date
    $dTo.Value   = $dApply.Value.Date.AddDays(7)
}
$dApply.Add_ValueChanged({ Sync-WorkPeriod })
Sync-WorkPeriod

$cImpName.Add_SelectedIndexChanged({ Fill-Party 'imp' $cImpName $tImpPerson $tImpAddr })
$cImpName.Add_Leave({ Fill-Party 'imp' $cImpName $tImpPerson $tImpAddr })
$cAppName.Add_SelectedIndexChanged({ Fill-Party 'app' $cAppName $tAppPerson $tAppAddr })
$cAppName.Add_Leave({ Fill-Party 'app' $cAppName $tAppPerson $tAppAddr })

function Set-Status([string]$s) { $lStatus.Text = $s; Pump }
function Set-Busy([bool]$busy) {
    foreach ($b in @($bLookup, $bGrab, $bPdf, $bPrev)) { $b.Enabled = -not $busy }
    $form.Cursor = $(if ($busy) { [System.Windows.Forms.Cursors]::WaitCursor } else { [System.Windows.Forms.Cursors]::Default })
    Pump
}

function Apply-Result($r) {
    $tArrive.Text = $r.etprDt
    $tShip.Text   = $r.sanm
    $tPkg.Text    = (@($r.pkg, $r.wght) | Where-Object { $_ }) -join ' / '
    $tGoods.Text  = $r.prnm
    if ($r.ban) {
        $code = ([string]$r.ban.sgn).Split('/')[0].Trim()
        $nm = ([string]$r.ban.nm).Trim()
        $tStore.Text = $(if ($nm -and $code) { "$nm($code)" } elseif ($nm) { $nm } else { $code })
        $tInNo.Text  = $r.ban.no
        $lBanInfo.Text = "화물관리번호 $($r.cargMtNo) · 반입신고(최상단) $($r.ban.dt) · 처리이력 $($r.histCount)건"
        $lBanInfo.ForeColor = [System.Drawing.Color]::DimGray
    } else {
        $tStore.Text = ''; $tInNo.Text = ''
        $lBanInfo.Text = "화물관리번호 $($r.cargMtNo) · 처리이력에 반입신고가 없습니다. 장치장소/반입신고번호를 직접 입력하세요."
        $lBanInfo.ForeColor = [System.Drawing.Color]::Firebrick
    }
}

function Do-Lookup {
    $bl = ($tBL.Text.ToUpper() -replace '[^A-Z0-9\-]', '')
    if (-not $bl) { Msg 'B/L번호를 입력하세요.'; return }
    $tBL.Text = $bl
    Set-Busy $true
    try {
        $kinds = switch ($cKind.SelectedIndex) { 1 { @('H') } 2 { @('M') } default { @('H', 'M') } }
        $years = @($cYear.SelectedItem)
        if ([string]$cYear.SelectedItem -eq [string]$thisYear) { $years += [string]($thisYear - 1) }
        foreach ($y in $years) {
            foreach ($k in $kinds) {
                Set-Status "조회 중 ($k B/L, $y)"
                $q = @{ blYy = $y }
                if ($k -eq 'H') { $q.hblNo = $bl } else { $q.mblNo = $bl }
                $res = Invoke-CargoApi $q
                # ※ tCnt는 처리이력 줄 수라서 화물 건수는 cargCsclPrgsInfoQryVo 개수로 판단
                if ($res.vos.Count -eq 0) { continue }
                if ($res.vos.Count -gt 1) {
                    Set-Status "결과 $($res.vos.Count)건"
                    $mt = Select-Cargo $res.vos
                    if (-not $mt) { Set-Status '선택 취소'; return }
                    Set-Status '선택한 화물 조회 중...'
                    $res = Invoke-CargoApi @{ cargMtNo = $mt }
                    if ($res.vos.Count -eq 0) { throw '선택한 화물을 다시 조회하지 못했습니다.' }
                }
                Apply-Result (Convert-CargoResult $res.vos[0] $res.dtls)
                Set-Status "조회 완료 ($k B/L, $y)"
                return
            }
        }
        Set-Status '조회 결과 없음'
        Msg "UNI-PASS에서 $bl 을(를) 찾지 못했습니다.`n(조회 연도: $($years -join ', '))" 'Warning'
    } catch {
        Set-Status '조회 실패'
        Msg ("조회 중 오류가 발생했습니다.`n" + $_.Exception.Message) 'Error'
    } finally { Set-Busy $false }
}

function Do-Grab {
    Set-Busy $true
    try {
        Set-Status '화면 읽는 중...'
        Connect-Browser
        $r = Run-Lookup '' '' 'GRAB'
        if ($r.state -ne 'found') { Set-Status ''; Msg 'UNI-PASS 브라우저에 화물 상세 화면이 열려 있지 않습니다.' 'Warning'; return }
        Apply-Result $r
        if (-not $tBL.Text) { $tBL.Text = $(if ($r.hbl) { $r.hbl } else { $r.mbl }) }
        Set-Status '화면에서 가져옴'
    } catch {
        Set-Status '실패'
        Msg ("화면을 읽지 못했습니다.`n" + $_.Exception.Message) 'Error'
    } finally { Set-Busy $false }
}

function Check-Party([string]$kind, $combo, $tPerson, $tAddr) {
    $label = $DbLabel[$kind]
    $name = $combo.Text.Trim()
    if (-not $name -or -not (Get-DbPath $kind)) { return }
    $r = Find-Company $kind $name
    try {
        if (-not $r) {
            if (Ask "'$name' 이(가) $label DB에 없습니다.`n성명/주소와 함께 $label DB에 추가할까요?") { Upsert-Company $kind $name $tPerson.Text $tAddr.Text }
        } elseif ($r.'성명' -ne $tPerson.Text -or $r.'주소' -ne $tAddr.Text) {
            if (Ask "'$name' 의 성명/주소가 $label DB와 다릅니다.`n지금 입력한 내용으로 $label DB를 수정할까요?") { Upsert-Company $kind $name $tPerson.Text $tAddr.Text }
        }
    } catch { Msg ("$label DB 저장 실패: " + $_.Exception.Message) 'Warning' }
}

function Safe-FileName([string]$s) { return ($s -replace '[\\/:*?"<>|]', '_').Trim() }

# 화면 입력값 → 서식 칸 값
function Get-FormValues {
    $ap = $dApply.Value
    $sign = (($cAppName.Text.Trim() + ' ' + $tAppPerson.Text.Trim()).Trim())
    $seal = Get-SealFor $cAppName.Text
    $vals = [ordered]@{
        'F6' = $cImpName.Text.Trim(); 'F7' = $tImpPerson.Text.Trim(); 'F8' = $tImpAddr.Text.Trim()
        'S6' = $cAppName.Text.Trim(); 'S7' = $tAppPerson.Text.Trim(); 'S8' = $tAppAddr.Text.Trim()
        'A10' = $tArrive.Text.Trim(); 'G10' = $tBL.Text.Trim(); 'R9' = $cPurpose.Text.Trim()
        'A12' = $tShip.Text.Trim(); 'G12' = $cGoodsTp.Text.Trim(); 'R11' = $cMethod.Text.Trim()
        'A14' = $tStore.Text.Trim(); 'G14' = $tInNo.Text.Trim()
        'R13' = $dFrom.Value.ToString('yyyy-MM-dd'); 'W13' = $dTo.Value.ToString('yyyy-MM-dd')
        'A17' = $tPkg.Text.Trim(); 'G17' = $tGoods.Text.Trim(); 'L17' = $tSpec.Text.Trim()
        'Q17' = $tQty.Text.Trim(); 'V17' = $tPrice.Text.Trim()
        'A22' = $tMName.Text.Trim(); 'H22' = $tMSpec.Text.Trim(); 'N22' = $tMQty.Text.Trim(); 'T22' = $tMPrice.Text.Trim()
        'H29' = $ap.ToString('yyyy'); 'L29' = [string]$ap.Month; 'N29' = [string]$ap.Day
        'Q31' = $(if ($seal) { '' } elseif ($sign) { "$sign (인)" } else { '' })
    }
    return @{ vals = $vals; seal = $seal }
}

function Do-Pdf {
    if (-not $Settings.templatePath -or -not (Test-Path $Settings.templatePath)) { Msg '서식 파일(xlsx)을 [설정]에서 지정하세요.' 'Warning'; return }
    if (-not (Test-Path $Settings.outDir)) { Msg '저장 폴더를 [설정]에서 지정하세요.' 'Warning'; return }

    Check-Party 'imp' $cImpName $tImpPerson $tImpAddr
    Check-Party 'app' $cAppName $tAppPerson $tAppAddr

    $fv = Get-FormValues; $vals = $fv.vals; $seal = $fv.seal

    $base = $(if ($tBL.Text.Trim()) { (Safe-FileName $tBL.Text) + '_보수작업신청서' } else { '보수작업신청서_' + (Get-Date -Format 'yyyyMMdd_HHmm') })
    $pdf = Join-Path $Settings.outDir ($base + '.pdf')
    $n = 2
    while (Test-Path $pdf) { $pdf = Join-Path $Settings.outDir ("{0}_{1}.pdf" -f $base, $n); $n++ }
    $xlsx = $(if ($chkXlsx.Checked) { [IO.Path]::ChangeExtension($pdf, '.xlsx') } else { $null })

    Set-Busy $true
    try {
        Set-Status 'PDF 만드는 중...'
        Export-Form $vals $pdf $xlsx $seal
        Push-History 'purposes' $cPurpose.Text; Push-History 'methods' $cMethod.Text
        $Settings.lastApplicant = $cAppName.Text.Trim()
        $Settings.openPdf = $chkOpen.Checked; $Settings.saveXlsx = $chkXlsx.Checked
        Save-Settings; Refresh-HistoryCombos
        Set-Status 'PDF 생성 완료'
        if ($chkOpen.Checked) { Start-Process -FilePath $pdf }
        else { Msg "저장했습니다.`n$pdf" }
    } catch {
        Set-Status 'PDF 생성 실패'
        Msg ("PDF를 만들지 못했습니다.`n" + $_.Exception.Message) 'Error'
    } finally { Set-Busy $false }
}

# ---------------------------------------------------------------- 미리보기
# 엑셀로 임시 PDF를 만든 뒤, 윈도우 내장 PDF 기능(Windows.Data.Pdf)으로 그림으로 바꿔 창에 보여줌
function Await-WinRt($op, [Type]$type) {
    if (-not $script:AsTaskOp) {
        Add-Type -AssemblyName System.Runtime.WindowsRuntime
        $ms = [System.WindowsRuntimeSystemExtensions].GetMethods()
        $script:AsTaskOp  = ($ms | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]
        $script:AsTaskAct = ($ms | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncAction' })[0]
    }
    if ($type) { $t = $script:AsTaskOp.MakeGenericMethod($type).Invoke($null, @($op)) }
    else { $t = $script:AsTaskAct.Invoke($null, @($op)) }
    if (-not $t.Wait(30000)) { throw '시간 초과' }
    if ($type) { return $t.Result }
}
function Render-PdfPage([string]$pdf, [string]$png, [int]$width) {
    Add-Type -AssemblyName System.Runtime.WindowsRuntime
    [void][Windows.Storage.StorageFile,Windows.Storage,ContentType=WindowsRuntime]
    [void][Windows.Storage.StorageFolder,Windows.Storage,ContentType=WindowsRuntime]
    [void][Windows.Storage.CreationCollisionOption,Windows.Storage,ContentType=WindowsRuntime]
    [void][Windows.Storage.FileAccessMode,Windows.Storage,ContentType=WindowsRuntime]
    [void][Windows.Storage.Streams.IRandomAccessStream,Windows.Storage.Streams,ContentType=WindowsRuntime]
    [void][Windows.Data.Pdf.PdfDocument,Windows.Data.Pdf,ContentType=WindowsRuntime]
    [void][Windows.Data.Pdf.PdfPageRenderOptions,Windows.Data.Pdf,ContentType=WindowsRuntime]
    $file = Await-WinRt ([Windows.Storage.StorageFile]::GetFileFromPathAsync($pdf)) ([Windows.Storage.StorageFile])
    $doc  = Await-WinRt ([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($file)) ([Windows.Data.Pdf.PdfDocument])
    $folder = Await-WinRt ([Windows.Storage.StorageFolder]::GetFolderFromPathAsync((Split-Path -Parent $png))) ([Windows.Storage.StorageFolder])
    $out  = Await-WinRt ($folder.CreateFileAsync((Split-Path -Leaf $png), [Windows.Storage.CreationCollisionOption]::ReplaceExisting)) ([Windows.Storage.StorageFile])
    $stream = Await-WinRt ($out.OpenAsync([Windows.Storage.FileAccessMode]::ReadWrite)) ([Windows.Storage.Streams.IRandomAccessStream])
    $page = $doc.GetPage(0)
    try {
        $opt = [Windows.Data.Pdf.PdfPageRenderOptions]::new()
        $opt.DestinationWidth = [uint32]$width
        Await-WinRt ($page.RenderToStreamAsync($stream, $opt)) $null
    } finally {
        try { $page.Dispose() } catch { }
        try { $stream.Dispose() } catch { }
    }
}
function Show-Preview([string]$pdf) {
    $png = [IO.Path]::ChangeExtension($pdf, '.png')
    $img = $null
    try {
        Render-PdfPage $pdf $png 1240
        $bytes = [IO.File]::ReadAllBytes($png)
        $img = [System.Drawing.Image]::FromStream((New-Object System.IO.MemoryStream(, $bytes)))
    } catch {
        # 그림으로 바꾸지 못하면 PDF 보기 프로그램으로 엶
        Start-Process -FilePath $pdf
        return
    }
    $f = New-Object System.Windows.Forms.Form
    $f.Text = '미리보기 (저장되지 않은 상태입니다)'; $f.Font = $UiFont; $f.StartPosition = 'CenterParent'
    $f.ClientSize = New-Object System.Drawing.Size(820, ([Math]::Min(1000, [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Height - 80)))
    $panel = New-Object System.Windows.Forms.Panel
    $panel.Dock = 'Fill'; $panel.AutoScroll = $true; $panel.BackColor = [System.Drawing.Color]::FromArgb(230, 230, 230)
    $pic = New-Object System.Windows.Forms.PictureBox
    $pic.SizeMode = 'Zoom'; $pic.Image = $img; $pic.BackColor = [System.Drawing.Color]::White
    $panel.Controls.Add($pic)
    $bar = New-Object System.Windows.Forms.Panel
    $bar.Dock = 'Bottom'; $bar.Height = 46
    $zoom = New-Object System.Windows.Forms.ComboBox
    $zoom.DropDownStyle = 'DropDownList'; [void]$zoom.Items.AddRange(@('화면 너비에 맞춤', '75%', '100%', '125%', '150%'))
    $zoom.Location = New-Object System.Drawing.Point(10, 11); $zoom.Size = New-Object System.Drawing.Size(140, 23)
    $bSave = New-Object System.Windows.Forms.Button; $bSave.Text = 'PDF 생성'; $bSave.Size = New-Object System.Drawing.Size(110, 30)
    $bSave.Font = New-Object System.Drawing.Font('Malgun Gothic', 10, [System.Drawing.FontStyle]::Bold)
    $bOpen = New-Object System.Windows.Forms.Button; $bOpen.Text = 'PDF 보기로 열기'; $bOpen.Size = New-Object System.Drawing.Size(120, 30)
    $bClose = New-Object System.Windows.Forms.Button; $bClose.Text = '닫기(수정하기)'; $bClose.Size = New-Object System.Drawing.Size(110, 30)
    foreach ($b in @($bClose, $bOpen, $bSave)) { $b.Anchor = 'Top,Right' }
    $bar.Controls.AddRange(@($zoom, $bOpen, $bClose, $bSave))
    $f.Controls.Add($panel); $f.Controls.Add($bar)
    $layout = {
        $avail = $panel.ClientSize.Width - 30
        $w = switch ($zoom.SelectedIndex) { 1 { 595 * 1.33 * 0.75 } 2 { 595 * 1.33 } 3 { 595 * 1.33 * 1.25 } 4 { 595 * 1.33 * 1.5 } default { $avail } }
        $w = [int][Math]::Max(200, $w)
        $pic.Size = New-Object System.Drawing.Size($w, [int]($w * $img.Height / $img.Width))
        $pic.Location = New-Object System.Drawing.Point([int][Math]::Max(10, ($avail - $w) / 2 + 10), 10)
        $bSave.Location  = New-Object System.Drawing.Point(($bar.ClientSize.Width - 120), 8)
        $bClose.Location = New-Object System.Drawing.Point(($bar.ClientSize.Width - 240), 8)
        $bOpen.Location  = New-Object System.Drawing.Point(($bar.ClientSize.Width - 370), 8)
    }
    $zoom.Add_SelectedIndexChanged($layout)
    $f.Add_Resize($layout)
    $f.Add_Shown({ $zoom.SelectedIndex = 0; & $layout })
    $script:PreviewSave = $false
    $bSave.Add_Click({ $script:PreviewSave = $true; $f.Close() })
    $bClose.Add_Click({ $f.Close() })
    $bOpen.Add_Click({ Start-Process -FilePath $pdf })
    [void]$f.ShowDialog($form)
    $img.Dispose()
    if ($script:PreviewSave) { Do-Pdf }
}
function Do-Preview {
    if (-not $Settings.templatePath -or -not (Test-Path $Settings.templatePath)) { Msg '서식 파일(xlsx)을 [설정]에서 지정하세요.' 'Warning'; return }
    $dir = Join-Path $env:TEMP 'BosuForm'
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
    Get-ChildItem -Path $dir -Filter 'preview_*' -ErrorAction SilentlyContinue | Remove-Item -ErrorAction SilentlyContinue
    $pdf = Join-Path $dir ('preview_' + (Get-Date -Format 'yyyyMMdd_HHmmss') + '.pdf')
    $fv = Get-FormValues
    Set-Busy $true
    try {
        Set-Status '미리보기 만드는 중...'
        Export-Form $fv.vals $pdf $null $fv.seal
        Set-Status ''
    } catch {
        Set-Status '미리보기 실패'
        Msg ("미리보기를 만들지 못했습니다.`n" + $_.Exception.Message) 'Error'
        return
    } finally { Set-Busy $false }
    Show-Preview $pdf
}

function Do-New {
    foreach ($t in @($tBL, $tArrive, $tShip, $tStore, $tInNo, $tPkg, $tGoods, $tSpec, $tQty, $tPrice, $tMName, $tMSpec, $tMQty, $tMPrice, $tImpPerson, $tImpAddr)) { $t.Text = '' }
    $cImpName.Text = ''; $cPurpose.Text = ''; $cMethod.Text = ''; $cGoodsTp.Text = '외국물품'
    $dApply.Value = (Get-Date).Date; Sync-WorkPeriod
    $lBanInfo.Text = ''; $lStatus.Text = ''
    [void]$tBL.Focus()
}

# ---------------------------------------------------------------- 설정 창
function Show-Settings {
    $f = New-Object System.Windows.Forms.Form
    $f.Text = '설정'; $f.Font = $UiFont; $f.StartPosition = 'CenterParent'; $f.FormBorderStyle = 'FixedDialog'
    $f.MaximizeBox = $false; $f.MinimizeBox = $false; $f.ClientSize = New-Object System.Drawing.Size(650, 283)
    $rows = @(@('서식 파일(xlsx)', $Settings.templatePath), @('수출입자 DB(csv)', $Settings.impDbPath),
              @('신청인 DB(csv)', $Settings.appDbPath), @('PDF 저장 폴더', $Settings.outDir),
              @('UNI-PASS API 키', (Get-ApiKey)))
    $boxes = @()
    for ($i = 0; $i -lt 5; $i++) {
        $y = 15 + $i * 38
        $l = New-Object System.Windows.Forms.Label; $l.Text = $rows[$i][0]; $l.Location = New-Object System.Drawing.Point(10, ($y + 4)); $l.Size = New-Object System.Drawing.Size(115, 20); $f.Controls.Add($l)
        $t = New-Object System.Windows.Forms.TextBox; $t.Text = $rows[$i][1]; $t.Location = New-Object System.Drawing.Point(125, $y); $t.Size = New-Object System.Drawing.Size(355, 23); $f.Controls.Add($t)
        $boxes += $t
    }
    $tTpl = $boxes[0]; $tImp = $boxes[1]; $tApp = $boxes[2]; $tOut = $boxes[3]; $tKey = $boxes[4]
    function New-Btn($text, $x, $y, $w) {
        $b = New-Object System.Windows.Forms.Button; $b.Text = $text; $b.Location = New-Object System.Drawing.Point($x, $y); $b.Size = New-Object System.Drawing.Size($w, 26)
        $f.Controls.Add($b); return $b
    }
    $bTpl  = New-Btn '찾기' 487 14 70
    $bImp  = New-Btn '찾기' 487 52 70;  $bImpN = New-Btn '새로 만들기' 561 52 82
    $bApp  = New-Btn '찾기' 487 90 70;  $bAppN = New-Btn '새로 만들기' 561 90 82
    $bOut  = New-Btn '찾기' 487 128 70
    $chkShareKey = New-Object System.Windows.Forms.CheckBox
    $chkShareKey.Text = "API 키를 공유폴더(업체 DB 폴더의 $KeyFileName)에 저장해 모든 PC가 같이 사용"
    $chkShareKey.Location = New-Object System.Drawing.Point(125, 192); $chkShareKey.Size = New-Object System.Drawing.Size(515, 22)
    $chkShareKey.Checked = $true
    $f.Controls.Add($chkShareKey)
    $lKeySrc = New-Object System.Windows.Forms.Label
    $lKeySrc.Location = New-Object System.Drawing.Point(142, 214); $lKeySrc.Size = New-Object System.Drawing.Size(500, 18)
    $lKeySrc.ForeColor = [System.Drawing.Color]::Gray
    $lKeySrc.Text = $(if (Read-SharedKey) { "현재: 공유 키 파일 사용 중 ($(Get-SharedKeyPath))" } elseif ($Settings.apiKey) { '현재: 이 PC에만 저장된 키 사용 중' } else { '현재: 키 없음' })
    $f.Controls.Add($lKeySrc)
    $ok = New-Object System.Windows.Forms.Button; $ok.Text = '저장'; $ok.Location = New-Object System.Drawing.Point(462, 243); $ok.Size = New-Object System.Drawing.Size(85, 30); $f.Controls.Add($ok)
    $cancel = New-Object System.Windows.Forms.Button; $cancel.Text = '취소'; $cancel.Location = New-Object System.Drawing.Point(553, 243); $cancel.Size = New-Object System.Drawing.Size(85, 30); $f.Controls.Add($cancel)
    $f.AcceptButton = $ok; $f.CancelButton = $cancel

    $bTpl.Add_Click({
        $d = New-Object System.Windows.Forms.OpenFileDialog; $d.Filter = '엑셀 파일 (*.xlsx;*.xls)|*.xlsx;*.xls'
        if ($d.ShowDialog() -eq 'OK') { $tTpl.Text = $d.FileName }
    })
    $pickCsv = {
        param($box)
        $d = New-Object System.Windows.Forms.OpenFileDialog; $d.Filter = 'CSV 파일 (*.csv)|*.csv'
        if ($d.ShowDialog() -eq 'OK') { $box.Text = $d.FileName }
    }
    $newCsv = {
        param($box, $defName)
        $d = New-Object System.Windows.Forms.SaveFileDialog; $d.Filter = 'CSV 파일 (*.csv)|*.csv'; $d.FileName = $defName
        if ($d.ShowDialog() -eq 'OK') {
            if (-not (Test-Path $d.FileName)) { [IO.File]::WriteAllText($d.FileName, "상호,성명,주소,직인파일`r`n", $Utf8Bom) }
            $box.Text = $d.FileName
        }
    }
    $bImp.Add_Click({ & $pickCsv $tImp })
    $bApp.Add_Click({ & $pickCsv $tApp })
    $bImpN.Add_Click({ & $newCsv $tImp '수출입자DB.csv' })
    $bAppN.Add_Click({ & $newCsv $tApp '신청인DB.csv' })
    $bOut.Add_Click({
        $d = New-Object System.Windows.Forms.FolderBrowserDialog
        if ($tOut.Text -and (Test-Path $tOut.Text)) { $d.SelectedPath = $tOut.Text }
        if ($d.ShowDialog() -eq 'OK') { $tOut.Text = $d.SelectedPath }
    })
    $ok.Add_Click({ $f.DialogResult = 'OK'; $f.Close() })
    $cancel.Add_Click({ $f.DialogResult = 'Cancel'; $f.Close() })
    if ($f.ShowDialog($form) -eq 'OK') {
        $Settings.templatePath = $tTpl.Text.Trim(); $Settings.impDbPath = $tImp.Text.Trim()
        $Settings.appDbPath = $tApp.Text.Trim(); $Settings.outDir = $tOut.Text.Trim()
        $newKey = $tKey.Text.Trim()
        $Settings.apiKey = $(if ($newKey) { $newKey } else { $DefaultApiKey })
        Save-Settings; Load-Db
        # API 키 공유: DB 경로가 바뀐 뒤 기준으로 공유 키 파일에 저장
        $kp = Get-SharedKeyPath
        if ($chkShareKey.Checked -and $newKey) {
            if (-not $kp) { Msg "업체 DB(csv) 경로가 없어 API 키를 공유폴더에 저장하지 못했습니다.`n이 PC에만 저장했습니다." 'Warning' }
            elseif ($newKey -ne (Read-SharedKey)) {
                try { [IO.File]::WriteAllText($kp, $newKey, (New-Object System.Text.UTF8Encoding($false))) }
                catch { Msg ("API 키를 공유폴더에 저장하지 못했습니다. 이 PC에만 저장했습니다.`n" + $_.Exception.Message) 'Warning' }
            }
        } elseif ($newKey -and (Read-SharedKey) -and $newKey -ne (Read-SharedKey)) {
            Msg "공유폴더에 다른 API 키 파일이 있어 그 키가 우선 사용됩니다.`n($kp)`n`n모든 PC의 키를 바꾸려면 '공유폴더에 저장'을 체크하고 저장하세요." 'Warning'
        }
    }
}

# ---------------------------------------------------------------- 업체 DB 관리 창
function Show-DbManager([string]$kind) {
    $dbPath = Get-DbPath $kind
    $isApp = ($kind -eq 'app')
    if (-not $dbPath) { Msg "먼저 [설정]에서 $($DbLabel[$kind]) DB(csv) 파일을 지정하세요." 'Warning'; return }
    $f = New-Object System.Windows.Forms.Form
    $f.Text = "$($DbLabel[$kind]) DB 관리 - $dbPath"; $f.Font = $UiFont; $f.StartPosition = 'CenterParent'
    $f.ClientSize = New-Object System.Drawing.Size(820, 480); $f.MinimizeBox = $false

    $dt = New-Object System.Data.DataTable
    foreach ($c in @('상호', '성명', '주소', '직인파일')) { [void]$dt.Columns.Add($c) }
    # 창을 연 시점의 내용(다른 사람이 그 사이 파일을 고쳤을 때 덮어쓰지 않도록 비교용)
    $snap = @{}
    try {
        foreach ($r in (Read-DbFile $dbPath)) {
            $v = @(([string]$r.'상호').Trim(), [string]$r.'성명', [string]$r.'주소', [string]$r.'직인파일')
            [void]$dt.Rows.Add($v[0], $v[1], $v[2], $v[3]); $snap[$v[0]] = ($v -join "`t")
        }
        $dt.AcceptChanges()
    }
    catch { Msg ("DB 파일을 읽지 못했습니다.`n" + $_.Exception.Message) 'Error'; return }

    $search = New-Object System.Windows.Forms.TextBox
    $search.Location = New-Object System.Drawing.Point(10, 10); $search.Size = New-Object System.Drawing.Size(300, 23)
    $f.Controls.Add($search)
    $hint = New-Object System.Windows.Forms.Label; $hint.Text = '← 상호 검색 · 맨 아래 빈 줄에 입력하면 추가, 행 선택 후 Delete 키로 삭제'
    $hint.Location = New-Object System.Drawing.Point(318, 13); $hint.AutoSize = $true; $hint.ForeColor = [System.Drawing.Color]::Gray
    $f.Controls.Add($hint)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Location = New-Object System.Drawing.Point(10, 40); $grid.Size = New-Object System.Drawing.Size(800, 390)
    $grid.Anchor = 'Top,Bottom,Left,Right'
    $grid.DataSource = $dt
    $grid.AutoSizeColumnsMode = 'Fill'
    $grid.RowHeadersWidth = 30
    $f.Controls.Add($grid)
    $f.Add_Shown({
        $grid.Columns[0].FillWeight = 25; $grid.Columns[1].FillWeight = 12; $grid.Columns[2].FillWeight = 43; $grid.Columns[3].FillWeight = 20
        $grid.Columns[3].ReadOnly = $true
        $grid.Columns[3].HeaderText = '직인파일 (더블클릭·끌어놓기)'
        if (-not $isApp) { $grid.Columns[3].Visible = $false }
    })
    $search.Add_TextChanged({
        $s = $search.Text.Replace("'", "''").Replace('[', '[[]').Replace('%', '[%]').Replace('*', '[*]')
        $dt.DefaultView.RowFilter = $(if ($s) { "[상호] LIKE '%$s%'" } else { '' })
    })

    $save = New-Object System.Windows.Forms.Button; $save.Text = '저장'; $save.Size = New-Object System.Drawing.Size(90, 30)
    $save.Location = New-Object System.Drawing.Point(624, 440); $save.Anchor = 'Bottom,Right'; $f.Controls.Add($save)
    $close = New-Object System.Windows.Forms.Button; $close.Text = '닫기'; $close.Size = New-Object System.Drawing.Size(90, 30)
    $close.Location = New-Object System.Drawing.Point(720, 440); $close.Anchor = 'Bottom,Right'; $f.Controls.Add($close)

    $commit = { $grid.EndEdit(); try { $f.BindingContext[$dt].EndCurrentEdit() } catch { } }
    $save.Add_Click({
        & $commit
        $rows = @()
        $seen = @{}
        # 파일을 다시 읽어서, 이 창에서 손대지 않은 줄은 파일의 최신 내용(다른 사람이 고친 것)을 그대로 씀
        $file = @{}
        try { foreach ($x in (Read-DbFile $dbPath)) { $file[([string]$x.'상호').Trim()] = $x } } catch { }
        foreach ($r in $dt.Rows) {
            if ($r.RowState -eq 'Deleted') { continue }
            $n = ([string]$r['상호']).Trim()
            if (-not $n) { continue }
            if ($seen.ContainsKey($n)) { Msg "상호가 중복됩니다: $n" 'Warning'; return }
            $seen[$n] = 1
            $mine = @($n, ([string]$r['성명']).Trim(), ([string]$r['주소']).Trim(), ([string]$r['직인파일']).Trim())
            if ($snap.ContainsKey($n) -and $snap[$n] -eq ($mine -join "`t") -and $file.ContainsKey($n)) {
                $x = $file[$n]; $mine = @($n, [string]$x.'성명', [string]$x.'주소', [string]$x.'직인파일')
            }
            $rows += [pscustomobject]@{ '상호' = $mine[0]; '성명' = $mine[1]; '주소' = $mine[2]; '직인파일' = $mine[3] }
        }
        # 그 사이 다른 사람이 추가한 업체도 유지 (이 창에서 지운 업체는 제외)
        foreach ($k in $file.Keys) {
            if (-not $seen.ContainsKey($k) -and -not $snap.ContainsKey($k)) {
                $x = $file[$k]; $rows += [pscustomobject]@{ '상호' = $k; '성명' = [string]$x.'성명'; '주소' = [string]$x.'주소'; '직인파일' = [string]$x.'직인파일' }
            }
        }
        try {
            Write-DbFile $dbPath $rows; Load-Db
            # 저장된 내용으로 다시 채움
            $dt.Rows.Clear(); $snap.Clear()
            foreach ($x in $rows) { [void]$dt.Rows.Add($x.'상호', $x.'성명', $x.'주소', $x.'직인파일'); $snap[$x.'상호'] = (@($x.'상호', $x.'성명', $x.'주소', $x.'직인파일') -join "`t") }
            $dt.AcceptChanges()
            Msg "저장했습니다. ($($rows.Count)개 업체)"
        }
        catch { Msg ("저장하지 못했습니다. 다른 사람이 파일을 열고 있는지 확인하세요.`n" + $_.Exception.Message) 'Error' }
    })
    $sealSet = New-Object System.Windows.Forms.Button; $sealSet.Text = '선택한 업체에 명판직인 지정'; $sealSet.Size = New-Object System.Drawing.Size(200, 30)
    $sealSet.Location = New-Object System.Drawing.Point(10, 440); $sealSet.Anchor = 'Bottom,Left'; $f.Controls.Add($sealSet)
    $sealDel = New-Object System.Windows.Forms.Button; $sealDel.Text = '직인 지정 해제'; $sealDel.Size = New-Object System.Drawing.Size(110, 30)
    $sealDel.Location = New-Object System.Drawing.Point(216, 440); $sealDel.Anchor = 'Bottom,Left'; $f.Controls.Add($sealDel)
    $sealLbl = New-Object System.Windows.Forms.Label; $sealLbl.Text = '직인은 지정하는 즉시 저장됩니다 (이미지를 줄에 끌어다 놓아도 됨)'
    $sealLbl.Location = New-Object System.Drawing.Point(334, 447); $sealLbl.AutoSize = $true; $sealLbl.ForeColor = [System.Drawing.Color]::Gray
    $sealLbl.Anchor = 'Bottom,Left'; $f.Controls.Add($sealLbl)
    foreach ($ctl in @($sealSet, $sealDel, $sealLbl)) { $ctl.Visible = $isApp }
    # 직인 지정: 이미지를 DB 폴더의 "직인" 폴더로 복사하고 그 업체 줄만 바로 저장
    $applySeal = {
        param($row, [string]$src)
        if (-not $row -or $row.IsNewRow) { Msg '업체 행을 먼저 선택하세요.'; return }
        & $commit
        $name = ([string]$row.Cells[0].Value).Trim()
        if (-not $name) { Msg '상호를 먼저 입력하세요.'; return }
        if (-not $src) {
            $d = New-Object System.Windows.Forms.OpenFileDialog
            $d.Filter = '이미지 파일 (*.png;*.jpg;*.jpeg;*.bmp;*.gif)|*.png;*.jpg;*.jpeg;*.bmp;*.gif'
            $d.Title = "$name 명판직인 이미지 선택"
            if ($d.ShowDialog() -ne 'OK') { return }
            $src = $d.FileName
        }
        try {
            $rel = $(if ($src -eq '') { '' } else { Import-SealFile $name $src })
            Set-CompanySeal $name ([string]$row.Cells[1].Value) ([string]$row.Cells[2].Value) $rel
            $row.Cells[3].Value = $rel; & $commit
            $dv = $row.DataBoundItem; if ($dv) { $snap[$name] = (@($name, [string]$dv['성명'], [string]$dv['주소'], $rel) -join "`t") }
            $sealLbl.Text = "$name : 직인 저장됨 ($rel)"
        } catch { Msg ("직인을 저장하지 못했습니다.`n" + $_.Exception.Message) 'Error' }
    }
    $sealSet.Add_Click({ & $applySeal $grid.CurrentRow $null })
    $sealDel.Add_Click({
        $row = $grid.CurrentRow
        if (-not $row -or $row.IsNewRow) { return }
        $name = ([string]$row.Cells[0].Value).Trim()
        try { Set-CompanySeal $name ([string]$row.Cells[1].Value) ([string]$row.Cells[2].Value) ''; $row.Cells[3].Value = ''; & $commit
              $snap[$name] = (@($name, [string]$row.Cells[1].Value, [string]$row.Cells[2].Value, '') -join "`t"); $sealLbl.Text = "$name : 직인 지정 해제됨" }
        catch { Msg ("저장하지 못했습니다.`n" + $_.Exception.Message) 'Error' }
    })
    if ($isApp) {
        $grid.Add_CellDoubleClick({ if ($_.RowIndex -ge 0 -and $_.ColumnIndex -eq 3) { & $applySeal $grid.Rows[$_.RowIndex] $null } })
        $grid.AllowDrop = $true
        $grid.Add_DragEnter({ if ($_.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) { $_.Effect = 'Copy' } })
        $grid.Add_DragDrop({
            $files = @($_.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)) | Where-Object { $_ -match '\.(png|jpe?g|bmp|gif)$' }
            if (-not $files) { Msg '이미지 파일(png, jpg, bmp, gif)만 넣을 수 있습니다.'; return }
            $pt = $grid.PointToClient((New-Object System.Drawing.Point($_.X, $_.Y)))
            $hit = $grid.HitTest($pt.X, $pt.Y)
            if ($hit.RowIndex -lt 0) { Msg '이미지를 업체 줄 위에 놓아 주세요.'; return }
            & $applySeal $grid.Rows[$hit.RowIndex] $files[0]
        })
    }
    $close.Add_Click({ $f.Close() })
    [void]$f.ShowDialog($form)
}

# ---------------------------------------------------------------- 이벤트
$bLookup.Add_Click({ Do-Lookup })
$bGrab.Add_Click({ Do-Grab })
$bPdf.Add_Click({ Do-Pdf })
$bPrev.Add_Click({ Do-Preview })
$bNew.Add_Click({ Do-New })
$bSet.Add_Click({ Show-Settings })
$bDbImp.Add_Click({ Show-DbManager 'imp' })
$bDbApp.Add_Click({ Show-DbManager 'app' })
$tBL.Add_KeyDown({ if ($_.KeyCode -eq 'Enter') { $_.SuppressKeyPress = $true; Do-Lookup } })

$form.Add_Shown({
    Load-Db
    Refresh-HistoryCombos
    if ($Settings.lastApplicant) { $cAppName.Text = $Settings.lastApplicant; Fill-Party 'app' $cAppName $tAppPerson $tAppAddr }
    if (-not $Settings.templatePath -or -not (Test-Path $Settings.templatePath) -or -not $Settings.impDbPath -or -not $Settings.appDbPath) {
        Msg '서식 파일, 수출입자 DB, 신청인 DB(공유폴더의 csv), 저장 폴더를 지정해 주세요.'
        Show-Settings
    }
    [void]$tBL.Focus()
})
$form.Add_FormClosed({ Close-Ws; Save-Settings })

[void]$form.ShowDialog()
