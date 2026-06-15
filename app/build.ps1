# build.ps1 — 드릴다운 표 생성기 EXE 빌드
# 엔진(drilldown_table.py, v14)을 스킬 폴더에서 그대로 번들 — 재구현/복사본 없음(단일 소스).
#
# 사용:
#   .\build.ps1                # 기본: onedir + 콘솔숨김(release)
#   .\build.ps1 -Console       # 콘솔 표시(디버그/CLI 검증용)
#   .\build.ps1 -OneFile       # 단일 .exe (배포 간편, 시작 느림·AV오탐 ↑)
param(
  [switch]$Console,   # 콘솔 창 표시(오류 확인/CLI 검증)
  [switch]$OneFile    # onefile 단일 exe (기본은 onedir)
)

$ErrorActionPreference = "Stop"
$ENGINE = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path   # v14 엔진 단일 소스
$NAME   = "드릴다운표생성기"

$flags = @(
  "--noconfirm", "--clean", "--name", $NAME,
  "--paths", $ENGINE, "--hidden-import", "drilldown_table",
  "--collect-data", "pptx", "--collect-data", "docx", "--collect-data", "openpyxl",
  "--collect-data", "customtkinter",
  "--hidden-import", "win32com.client", "--hidden-import", "pywintypes", "--hidden-import", "pythoncom",
  "--add-data", "assets;assets",
  # 안전 제외(미사용 전이 의존 — 약 -28MB). PIL/lxml은 미리보기·OOXML에 필요하므로 유지.
  "--exclude-module", "numpy", "--exclude-module", "pandas", "--exclude-module", "matplotlib",
  "--exclude-module", "scipy", "--exclude-module", "test", "--exclude-module", "unittest",
  "--exclude-module", "pydoc", "--exclude-module", "setuptools", "--exclude-module", "pip",
  "--exclude-module", "distutils"
)
if ($OneFile) { $flags += "--onefile" } else { $flags += "--onedir" }
if ($Console) { $flags += "--console" } else { $flags += "--windowed" }
$flags += "app.py"

Write-Host "엔진 단일소스: $ENGINE" -ForegroundColor Cyan
Write-Host "빌드: py -m PyInstaller $($flags -join ' ')" -ForegroundColor Cyan
py -m PyInstaller @flags

# 배포물에 입력양식 동봉
$distRoot = if ($OneFile) { "dist" } else { "dist\$NAME" }
py -X utf8 -c "import excel_io,os; excel_io.make_template(os.path.join(r'$distRoot','입력양식.xlsx')); print('입력양식.xlsx 동봉 ->', r'$distRoot')"

Write-Host "`n완료. 배포물: $distRoot" -ForegroundColor Green
Write-Host "코드서명(권장): Azure Trusted Signing 또는 signtool 로 $NAME.exe 서명 후 배포." -ForegroundColor Yellow
