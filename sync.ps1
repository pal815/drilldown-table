# sync.ps1 — 이 리포(엔진 정본)를 설치된 Claude/Codex 스킬 사본으로 동기화.
#
# 사용:
#   .\sync.ps1            # Claude + Codex 둘 다
#   .\sync.ps1 -Claude    # Claude 스킬만
#   .\sync.ps1 -Codex     # Codex 스킬만
#
# 동기화 대상: 엔진(drilldown_table.py) · 검수 렌더 · examples.
#  - Claude(~/.claude/skills/drilldown-table): SKILL.md · .claude-plugin · commands 까지 전부.
#  - Codex (~/.codex/skills/drilldown-table) : 엔진/예시/렌더만. Codex의 SKILL.md·agents/ 는
#    프런트매터 컨벤션이 달라 보존(덮어쓰지 않음).
param([switch]$Claude, [switch]$Codex)
$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
if (-not $Claude -and -not $Codex) { $Claude = $true; $Codex = $true }

$engineFiles = @('drilldown_table.py', 'render_png.py', 'render_pptx_png.py', 'render_docx_png.py')

function Copy-Dir([string]$srcDir, [string]$dstDir) {
  New-Item -ItemType Directory $dstDir -Force | Out-Null
  Copy-Item (Join-Path $srcDir '*') $dstDir -Recurse -Force
}
function Sync-Common([string]$dst) {
  New-Item -ItemType Directory $dst -Force | Out-Null
  foreach ($f in $engineFiles) { Copy-Item (Join-Path $root $f) (Join-Path $dst $f) -Force }
  Copy-Dir (Join-Path $root 'examples') (Join-Path $dst 'examples')
  # 프리빌트 Rust 바이너리(xlsx/docx Python 불필요). 있으면 동봉.
  if (Test-Path (Join-Path $root 'bin')) { Copy-Dir (Join-Path $root 'bin') (Join-Path $dst 'bin') }
}

if ($Claude) {
  $d = Join-Path $HOME '.claude\skills\drilldown-table'
  Sync-Common $d
  Copy-Item (Join-Path $root 'SKILL.md') (Join-Path $d 'SKILL.md') -Force
  Copy-Dir (Join-Path $root '.claude-plugin') (Join-Path $d '.claude-plugin')
  Copy-Dir (Join-Path $root 'commands')       (Join-Path $d 'commands')
  Write-Host "  Claude 스킬 동기화 완료 -> $d" -ForegroundColor Green
}
if ($Codex) {
  $d = Join-Path $HOME '.codex\skills\drilldown-table'
  Sync-Common $d
  Write-Host "  Codex 스킬 동기화 완료(엔진/예시/렌더만; SKILL.md·agents 보존) -> $d" -ForegroundColor Green
}
Write-Host "동기화 완료. 정본 = $root" -ForegroundColor Cyan
