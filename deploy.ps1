# 관리자용: 수정한 프로그램을 GitHub에 올려 팀원 PC가 자동 업데이트 받게 함
# 실행: 업데이트배포.bat 더블클릭
#  1) 업데이트 내용 입력 → 2) version.txt / 업데이트내용.txt 작성 → 3) 안전 점검 → 4) git commit + push
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic
$ErrorActionPreference = 'Stop'
$RepoUrl = 'https://github.com/ghj3815/bosu-form.git'
$here = $PSScriptRoot
Set-Location -LiteralPath $here
$Utf8Bom = New-Object System.Text.UTF8Encoding($true)
$Utf8 = New-Object System.Text.UTF8Encoding($false)

function Say([string]$t, [string]$icon = 'Information') { [void][System.Windows.Forms.MessageBox]::Show($t, '업데이트 배포', 'OK', $icon) }
function Fail([string]$t) { Say $t 'Error'; exit 1 }
function Invoke-Git { & git.exe -c core.quotepath=false @args; if ($LASTEXITCODE -ne 0) { throw "git $($args -join ' ') 실패 (코드 $LASTEXITCODE)" } }

try {
    if (-not (Get-Command git -ErrorAction SilentlyContinue)) { Fail 'git이 설치되어 있지 않습니다. https://git-scm.com 에서 설치하세요.' }

    # 처음 한 번: 이 폴더를 저장소로 연결
    if (-not (Test-Path -LiteralPath (Join-Path $here '.git'))) {
        Invoke-Git init -q
        Invoke-Git symbolic-ref HEAD refs/heads/main
        Invoke-Git remote add origin $RepoUrl
    }
    # 커밋 작성자 정보가 없으면 이 저장소에만 설정 (실제 이메일 대신 GitHub 비공개 주소)
    & git.exe config user.name *> $null; if ($LASTEXITCODE -ne 0) { Invoke-Git config user.name 'ghj3815' }
    & git.exe config user.email *> $null; if ($LASTEXITCODE -ne 0) { Invoke-Git config user.email 'ghj3815@users.noreply.github.com' }

    # 안전 점검 1: 코드에 API 키가 들어 있으면 중단
    $form = [IO.File]::ReadAllText((Join-Path $here 'bosu_form.ps1'))
    if ($form -match "(?m)^\`$DefaultApiKey\s*=\s*'[^']+'") { Fail "bosu_form.ps1 의 `$DefaultApiKey 에 API 키가 들어 있습니다.`n공개 저장소이므로 '' 로 비운 뒤 다시 배포하세요." }
    $myKey = ''
    $sp = Join-Path $env:APPDATA 'BosuForm\settings.json'
    if (Test-Path $sp) { try { $myKey = [string](Get-Content $sp -Raw -Encoding UTF8 | ConvertFrom-Json).apiKey } catch { } }

    $note = [Microsoft.VisualBasic.Interaction]::InputBox("이번 업데이트 내용을 적어주세요.`n(팀원 PC에서 업데이트될 때 안내창에 표시됩니다)`n`n취소하면 배포하지 않습니다.", '업데이트 배포', '')
    if (-not $note.Trim()) { exit 0 }

    $ver = Get-Date -Format 'yyyyMMdd-HHmm'
    [IO.File]::WriteAllText((Join-Path $here 'version.txt'), $ver, $Utf8)
    [IO.File]::WriteAllText((Join-Path $here '업데이트내용.txt'), "버전 $ver`r`n$($note.Trim())", $Utf8Bom)

    Invoke-Git add -A
    # 안전 점검 2: 올라가면 안 되는 파일 / 내 API 키가 들어간 파일이 있으면 중단
    $staged = @(& git.exe -c core.quotepath=false diff --cached --name-only)
    $bad = $staged | Where-Object { $_ -match '\.(csv|png|jpe?g|bmp|gif|zip|log)$' -or $_ -match '(^|/)직인/' -or $_ -match 'settings\.json$' }
    if ($bad) { Invoke-Git reset -q; Fail ("올리면 안 되는 파일이 포함되어 배포를 중단했습니다:`n" + ($bad -join "`n")) }
    if ($myKey) {
        foreach ($f in $staged) {
            $p = Join-Path $here $f
            if ((Test-Path -LiteralPath $p) -and ([IO.File]::ReadAllText($p)).Contains($myKey)) { Invoke-Git reset -q; Fail "API 키가 들어 있는 파일이 있어 배포를 중단했습니다: $f" }
        }
    }

    $msgFile = Join-Path $env:TEMP 'bosu_commit_msg.txt'
    [IO.File]::WriteAllText($msgFile, "$ver $($note.Trim())", $Utf8)
    Invoke-Git commit -q -F $msgFile
    Remove-Item -LiteralPath $msgFile -ErrorAction SilentlyContinue
    Write-Host 'GitHub에 올리는 중... (처음이면 GitHub 로그인 창이 뜹니다)'
    Invoke-Git push -u origin main
    Say "배포 완료: 버전 $ver`n`n팀원 PC는 다음에 프로그램을 켤 때 자동으로 업데이트됩니다."
} catch {
    Fail ("배포 중 오류가 발생했습니다.`n`n" + $_.Exception.Message + "`n`n이 창을 캡처해서 알려주세요.")
}
