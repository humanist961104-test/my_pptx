<#
.SYNOPSIS
  my_pptx 작업 환경 준비: 디스크 확인 → Node 확보(없으면 휴대용 zip 설치) → 스킬 복사(영문 경로) → npm 패키지 설치.
.EXAMPLE
  powershell -ExecutionPolicy Bypass -File setup.ps1
  powershell -ExecutionPolicy Bypass -File setup.ps1 -WorkDir C:\work\pptx -NodeDir C:\tools\node-v22.11.0-win-x64
.NOTES
  - WorkDir는 영문 경로여야 한다 (한글 경로에서 node가 비정상 종료하는 사례가 있음).
  - 휴대용 Node는 nodejs.org 공식 배포본을 받아 SHASUMS256.txt로 해시를 검증한 뒤 푼다.
  - 끝나면 이후 명령에 쓸 경로를 출력한다.
#>
param(
  [string]$WorkDir = (Join-Path $env:TEMP "my_pptx_work"),
  [string]$NodeDir = "",
  [string]$NodeVersion = "22.11.0",
  [int]$MinFreeMB = 500
)
$ErrorActionPreference = "Stop"
$SkillRoot = Split-Path -Parent $PSScriptRoot
function Step($m) { Write-Output "`n== $m" }

# 1. 경로·디스크
Step "작업 폴더 확인"
if ($WorkDir -match '[^\x00-\x7F]') { throw "WorkDir에 한글 등 비ASCII 문자가 있음: $WorkDir — 영문 경로를 지정하세요." }
New-Item -ItemType Directory -Force $WorkDir | Out-Null
$WorkDir = (Resolve-Path $WorkDir).Path
$drive = Get-PSDrive ($WorkDir.Substring(0, 1))
$freeMB = [int]($drive.Free / 1MB)
Write-Output "작업 폴더: $WorkDir  (드라이브 여유 $freeMB MB)"
if ($freeMB -lt $MinFreeMB) { throw "여유 공간 부족 ($freeMB MB < $MinFreeMB MB). 다른 드라이브의 -WorkDir 를 지정하세요." }
$projDrive = Get-PSDrive ((Get-Location).Path.Substring(0, 1))
Write-Output ("현재 폴더 드라이브 {0}: 여유 {1} MB (최종 .pptx만 저장 권장)" -f $projDrive.Name, [int]($projDrive.Free / 1MB))

# 2. Node
Step "Node.js 확인"
$node = $null
if ($NodeDir) { $node = Join-Path $NodeDir "node.exe"; if (-not (Test-Path $node)) { throw "node.exe 없음: $node" } }
if (-not $node) { $cmd = Get-Command node -ErrorAction SilentlyContinue; if ($cmd) { $node = $cmd.Source } }
$portable = Join-Path $WorkDir "node-v$NodeVersion-win-x64"
if (-not $node -and (Test-Path "$portable\node.exe")) { $node = "$portable\node.exe" }
if (-not $node) {
  $zipName = "node-v$NodeVersion-win-x64.zip"
  $base = "https://nodejs.org/dist/v$NodeVersion"
  $zip = Join-Path $WorkDir $zipName
  Write-Output "휴대용 Node 다운로드: $base/$zipName"
  [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
  Invoke-WebRequest "$base/$zipName" -OutFile $zip -UseBasicParsing
  $sums = (Invoke-WebRequest "$base/SHASUMS256.txt" -UseBasicParsing).Content
  $expected = ($sums -split "`n" | Where-Object { $_ -match [regex]::Escape($zipName) + '$' }) -replace '\s+.*$', ''
  $actual = (Get-FileHash $zip -Algorithm SHA256).Hash.ToLower()
  if (-not $expected -or $actual -ne $expected.Trim().ToLower()) { Remove-Item $zip -Force; throw "SHA256 불일치 — 다운로드 파일 삭제함" }
  Write-Output "SHA256 확인 완료"
  Expand-Archive $zip -DestinationPath $WorkDir -Force
  Remove-Item $zip -Force
  $node = "$portable\node.exe"
}
$nodeDirFinal = Split-Path -Parent $node
$env:PATH = "$nodeDirFinal;$env:PATH"
Write-Output ("node {0}  ({1})" -f (& $node -v), $node)

# 3. 스킬 복사 (영문 경로)
Step "스킬 파일 복사"
$skillCopy = Join-Path $WorkDir "skill"
New-Item -ItemType Directory -Force $skillCopy | Out-Null
foreach ($d in "scripts", "examples", "references") { Copy-Item -Recurse -Force (Join-Path $SkillRoot $d) $skillCopy }
Copy-Item -Force (Join-Path $SkillRoot "theme.json") $skillCopy
Write-Output "복사 위치: $skillCopy"

# 4. npm 패키지
Step "npm 패키지 확인"
$pkgs = "pptxgenjs", "react", "react-dom", "react-icons", "sharp"
$missing = $pkgs | Where-Object { -not (Test-Path (Join-Path $WorkDir "node_modules\$_")) }
if ($missing) {
  Write-Output "설치: $($missing -join ', ')"
  Push-Location $WorkDir
  try {
    if (-not (Test-Path package.json)) { '{ "name": "my-pptx-work", "private": true }' | Set-Content package.json -Encoding ascii }
    & (Join-Path $nodeDirFinal "npm.cmd") install --no-audit --no-fund @pkgs
    if ($LASTEXITCODE -ne 0) { throw "npm install 실패 (exit $LASTEXITCODE)" }
  } finally { Pop-Location }
} else { Write-Output "이미 설치됨" }

# 5. Python
Step "Python 확인"
$py = Get-Command python -ErrorAction SilentlyContinue
if ($py) { Write-Output ("{0} — 스크립트 실행 시 PYTHONUTF8=1 을 붙이세요" -f (& python --version)) } else { Write-Output "python 없음 — brand_lint.py를 쓰려면 Python 3 설치 필요" }

# 6. 요약
Step "준비 완료"
Write-Output "WORK  = $WorkDir"
Write-Output "NODE  = $node"
Write-Output "SKILL = $skillCopy"
Write-Output ""
Write-Output "# 빌드 (작업 폴더에서 실행)"
Write-Output "cd `"$WorkDir`"; & `"$node`" `"$skillCopy\scripts\build.js`" deck.json out.pptx"
Write-Output "# 검사 / 렌더링"
Write-Output "`$env:PYTHONUTF8='1'; python `"$skillCopy\scripts\brand_lint.py`" out.pptx"
Write-Output "powershell -ExecutionPolicy Bypass -File `"$skillCopy\scripts\render.ps1`" -Pptx out.pptx"
