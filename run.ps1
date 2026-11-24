# 실행기: (1) GitHub에 새 버전이 있으면 자동 업데이트 (2) 본 프로그램 실행, 오류가 나면 내용을 보여주고 기록합니다.
$log = Join-Path $env:APPDATA 'BosuForm\error.log'

# ---------------------------------------------------------------- 자동 업데이트 (GitHub)
# GitHub 저장소의 version.txt 가 내 version.txt 보다 새 것이면 최신 파일을 받아 덮어씀.
# 인터넷이 안 되거나 실패하면 조용히 넘어가고 지금 버전으로 실행.
# 관리자 개발 폴더(.git 있음)에서는 업데이트하지 않음.
$Repo = 'ghj3815-jpg/bosu-form'; $Branch = 'main'
function Read-Text([string]$p) {
    if (Test-Path -LiteralPath $p) { return ([IO.File]::ReadAllText($p)).Trim() }
    return ''
}
$here = $PSScriptRoot
$tmp = $null
try {
    if (-not (Test-Path -LiteralPath (Join-Path $here '.git'))) {
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        $req = [Net.HttpWebRequest]::Create("https://raw.githubusercontent.com/$Repo/$Branch/version.txt?t=$([DateTime]::UtcNow.Ticks)")
        $req.Timeout = 4000; $req.ReadWriteTimeout = 4000; $req.UserAgent = 'bosu-form-updater'
        $resp = $req.GetResponse()
        try { $remote = (New-Object IO.StreamReader($resp.GetResponseStream())).ReadToEnd().Trim() } finally { $resp.Close() }
        $local = Read-Text (Join-Path $here 'version.txt')
        if ($remote -match '^\d{8}-\d{4}$' -and [string]::CompareOrdinal($remote, $local) -gt 0) {
            $tmp = Join-Path $env:TEMP ('bosu_update_' + [guid]::NewGuid().ToString('N'))
            New-Item -ItemType Directory -Path $tmp | Out-Null
            $zip = Join-Path $tmp 'update.zip'
            $wc = New-Object Net.WebClient
            $wc.Headers['User-Agent'] = 'bosu-form-updater'
            $wc.DownloadFile("https://codeload.github.com/$Repo/zip/refs/heads/$Branch", $zip)
            Expand-Archive -LiteralPath $zip -DestinationPath (Join-Path $tmp 'x') -Force
            $root = Get-ChildItem -LiteralPath (Join-Path $tmp 'x') -Directory | Select-Object -First 1
            if (-not (Test-Path -LiteralPath (Join-Path $root.FullName 'bosu_form.ps1'))) { throw '받은 파일에 bosu_form.ps1 이 없습니다.' }
            # 관리자용 파일은 받지 않음, version.txt 는 마지막에 (중간에 실패하면 다음 실행 때 다시 시도)
            $skip = @('version.txt', '.gitignore', '.gitattributes', 'README.md', '업데이트배포.bat', 'deploy.ps1')
            foreach ($f in Get-ChildItem -LiteralPath $root.FullName -File) {
                if ($skip -notcontains $f.Name) { Copy-Item -LiteralPath $f.FullName -Destination (Join-Path $here $f.Name) -Force }
            }
            Copy-Item -LiteralPath (Join-Path $root.FullName 'version.txt') -Destination (Join-Path $here 'version.txt') -Force
            $note = Read-Text (Join-Path $here '업데이트내용.txt')
            Add-Type -AssemblyName System.Windows.Forms
            [void][System.Windows.Forms.MessageBox]::Show(("새 버전으로 업데이트했습니다.`n`n" + $note), '보수작업신청서 업데이트', 'OK', 'Information')
        }
    }
} catch {
    try {
        New-Item -ItemType Directory -Force -Path (Split-Path $log) | Out-Null
        ("[{0}] 업데이트 확인 실패(지금 버전으로 실행): {1}`r`n" -f (Get-Date), $_.Exception.Message) | Out-File -FilePath $log -Append -Encoding UTF8
    } catch { }
} finally {
    if ($tmp) { Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue }
}

# ---------------------------------------------------------------- 본 프로그램 실행
try {
    & (Join-Path $PSScriptRoot 'bosu_form.ps1')
} catch {
    $e = $_
    try {
        New-Item -ItemType Directory -Force -Path (Split-Path $log) | Out-Null
        ("[{0}] {1}`r`n{2}`r`n{3}`r`n" -f (Get-Date), $e.Exception.Message, $e.InvocationInfo.PositionMessage, $e.ScriptStackTrace) |
            Out-File -FilePath $log -Append -Encoding UTF8
    } catch { }
    Add-Type -AssemblyName System.Windows.Forms
    [void][System.Windows.Forms.MessageBox]::Show(
        ("프로그램 실행 중 오류가 발생했습니다.`n`n" + $e.Exception.Message + "`n`n" + $e.InvocationInfo.PositionMessage + "`n`n(기록: $log)"),
        '보수작업신청서', 'OK', 'Error')
    exit 2
}
exit 0
