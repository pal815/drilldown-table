# sign.ps1 — 드릴다운 표 생성기 EXE 코드서명 + 해시 매니페스트.
# 인증서가 준비되면 한 줄로 서명. SmartScreen은 평판 기반이라 서명+배포 누적이 핵심이며,
# 서명 자체로 Defender 오탐·일부 정책을 완화한다. (네이티브/파이썬 무관 — 언어 아닌 '서명'이 본질.)
#
# 사용 예:
#   .\sign.ps1 -AzureTrustedSigning -Endpoint "https://eus.codesigning.azure.net" -Account "myacct" -Profile "myprofile"
#   .\sign.ps1 -Pfx cert.pfx -Password (Read-Host -AsSecureString)
#   .\sign.ps1 -Thumbprint A1B2C3...      # 인증서 저장소에 설치된 인증서
#   .\sign.ps1 -SelfSignTest             # 자체서명(SmartScreen엔 무효, 파이프라인 점검용)
#   .\sign.ps1 -HashOnly                 # 서명 없이 SHA256 매니페스트만(IT 화이트리스트 요청용)
param(
  [string]$Exe = "dist\DrilldownTable\DrilldownTable.exe",
  [switch]$AzureTrustedSigning, [string]$Endpoint, [string]$Account, [string]$Profile,
  [string]$Pfx, [System.Security.SecureString]$Password,
  [string]$Thumbprint,
  [switch]$SelfSignTest,
  [switch]$HashOnly,
  [string]$Timestamp = "http://timestamp.acs.microsoft.com"
)
$ErrorActionPreference = "Stop"
if (-not (Test-Path $Exe)) { throw "EXE 없음: $Exe  (먼저 .\build.ps1 실행)" }

function Find-SignTool {
  $c = Get-Command signtool.exe -ErrorAction SilentlyContinue
  if ($c) { return $c.Source }
  $hit = Get-ChildItem "C:\Program Files (x86)\Windows Kits\10\bin" -Recurse -Filter signtool.exe -ErrorAction SilentlyContinue |
         Where-Object FullName -like "*x64*" | Sort-Object FullName -Descending | Select-Object -First 1
  if ($hit) { return $hit.FullName }
  throw "signtool.exe 미발견 — Windows SDK 설치 필요(또는 -HashOnly 로 해시만 생성)."
}

if (-not $HashOnly) {
  $st = Find-SignTool
  if ($AzureTrustedSigning) {
    # 권장: Azure Trusted Signing ($9.99/월). 사전: 'dotnet tool install --global AzureSignTool' 또는
    # Trusted Signing dlib + 메타데이터. 아래는 dlib 방식 예시(경로/계정은 환경에 맞게).
    $meta = Join-Path $PSScriptRoot "trusted-signing-metadata.json"
    if (-not (Test-Path $meta)) {
      @{ Endpoint = $Endpoint; CodeSigningAccountName = $Account; CertificateProfileName = $Profile } |
        ConvertTo-Json | Set-Content $meta -Encoding utf8
      Write-Host "메타데이터 생성: $meta (Endpoint/Account/Profile 확인)" -ForegroundColor Yellow
    }
    $dlib = (Get-ChildItem "$env:USERPROFILE\.nuget\packages\microsoft.trusted.signing.client" -Recurse -Filter "Azure.CodeSigning.Dlib.dll" -ErrorAction SilentlyContinue | Select-Object -First 1).FullName
    if (-not $dlib) { throw "Trusted Signing dlib 미발견 — 'nuget install Microsoft.Trusted.Signing.Client' 후 재시도(또는 AzureSignTool 사용)." }
    & $st sign /v /fd SHA256 /tr $Timestamp /td SHA256 /dlib $dlib /dmdf $meta $Exe
  }
  elseif ($Pfx) {
    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password)
    $plain = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
    & $st sign /v /fd SHA256 /tr $Timestamp /td SHA256 /f $Pfx /p $plain $Exe
  }
  elseif ($Thumbprint) {
    & $st sign /v /fd SHA256 /tr $Timestamp /td SHA256 /sha1 $Thumbprint $Exe
  }
  elseif ($SelfSignTest) {
    Write-Host "자체서명(테스트) — SmartScreen엔 무효. 파이프라인 점검용." -ForegroundColor Yellow
    $cert = New-SelfSignedCertificate -Type CodeSigningCert -Subject "CN=DrilldownTable Test" -CertStoreLocation Cert:\CurrentUser\My
    Set-AuthenticodeSignature -FilePath $Exe -Certificate $cert -TimestampServer $Timestamp | Out-Null
  }
  else { throw "서명 방식 미지정 — -AzureTrustedSigning / -Pfx / -Thumbprint / -SelfSignTest / -HashOnly 중 하나." }

  Write-Host "서명 검증:" -ForegroundColor Cyan
  & $st verify /pa /v $Exe
}

# SHA256 매니페스트 (IT 보안팀 해시 화이트리스트 요청용)
$dir = Split-Path $Exe -Parent
$manifest = Join-Path $dir "SHA256SUMS.txt"
Get-ChildItem $dir -Recurse -File | Get-FileHash -Algorithm SHA256 |
  ForEach-Object { "{0}  {1}" -f $_.Hash, ($_.Path.Replace((Resolve-Path $dir).Path + "\", "")) } |
  Set-Content $manifest -Encoding utf8
Write-Host "SHA256 매니페스트: $manifest" -ForegroundColor Green
Write-Host "→ 이 파일을 IT 보안팀에 제출해 빌드 해시 화이트리스트(AppLocker/WDAC 예외)를 요청하세요." -ForegroundColor Green
