# 드릴다운 표 생성기 — 릴리스 / 배포 가이드

언어모델·파이썬 설치 없이 일반 사무직이 쓰는 두 채널:
- **EXE**(주력) — `dist\DrilldownTable\` (onedir, 약 52MB). 더블클릭 GUI.
- **웹**(폴백) — `drilldown_web.html` 1개. 설치 0, 브라우저에서 동일 엔진(Pyodide) 실행. 잠긴 PC용.

> 끝사용자는 **파이썬 불필요**. EXE는 인터프리터를 번들, 웹은 브라우저에서 Pyodide로 실행한다.

## 1. EXE 빌드 → 서명 → 배포
```powershell
.\build.ps1                 # release(창모드 onedir). -OneFile 단일exe, -Console 디버그
.\sign.ps1 -HashOnly        # 서명 전: SHA256SUMS.txt 생성(IT 화이트리스트 요청용)
# 인증서 준비 후 한 줄 서명:
.\sign.ps1 -AzureTrustedSigning -Endpoint <...> -Account <...> -Profile <...>
#   또는  .\sign.ps1 -Pfx cert.pfx -Password (Read-Host -AsSecureString)
```
**서명이 핵심**: 무서명 exe는 SmartScreen "평판 없음" 경고가 거의 100% 뜨고, AppLocker/WDAC "서명 exe만"·Smart App Control 환경에선 무권한 사용자가 아예 실행 못 한다. 권장 = **Azure Trusted Signing($9.99/월, 개인 등록 가능)**. 단 서명해도 SmartScreen '평판'은 다운로드 누적으로만 쌓이므로(2024.3 이후 EV 즉시특권 폐지), **IT에 `SHA256SUMS.txt` 제출 → 해시 화이트리스트** 등록을 병행한다. ※ 이 서명 파이프라인은 나중에 Rust 바이너리에도 그대로 재사용된다(언어 무관, 매몰비용 없음).

**배포 경로**: 메일첨부/임의 다운로드는 MOTW가 붙어 마찰↑ → IT 승인 공유위치(신뢰 경로)로 전달. 동봉물 = `DrilldownTable.exe` + `_internal\` + `입력양식.xlsx`.

## 2. 웹 폴백(잠긴 PC용)
- `py build_web.py` → `drilldown_web.html` 재생성(엔진 바뀌면 다시 실행 = 단일 소스 유지).
- **검증됨**: Pyodide v314로 `openpyxl`(micropip)·`lxml`/`Pillow`(번들)·`python-docx`/`python-pptx`(micropip)가 동작해 동일 v14 산출물(row/both × xlsx/docx/pptx) 생성 확인.
- 호스팅: 사내 인트라넷/공유드라이브에 `.html` 1개. **오프라인/외부차단 환경**은 `build_web.py`의 `PYODIDE_BASE`를 사내 호스팅한 Pyodide 자산 경로로 바꾼다(CDN `*.jsdelivr` 차단 대비). 데이터는 브라우저 밖으로 나가지 않음(DLP 안전).

## 3. 용량 줄이기(선택)
현재 52MB. 추가 절감:
- **PIL 제외**(이미지 미사용): `build.ps1`에 `--exclude-module PIL` 추가 → 약 39MB(엔진이 이미지 안 넣으므로 안전).
- 더 작게: Python embeddable-zip 동봉(~30MB) 또는 Nuitka `--standalone`(AV 오탐↓).
- **진짜 저자원(2~6MB 단일 네이티브)**: Rust 재구현(별도 트랙). xlsx=`rust_xlsxwriter`, docx=`docx-rs`, pptx=`zip`+`quick-xml` 직접 OOXML. v14 동결이라 이식=단일 소스 교체.

## 구성 파일
- `app.py`(GUI+CLI) · `excel_io.py`(입력양식↔모델) · `build.ps1` · `sign.ps1` · `build_web.py`/`drilldown_web.html` · `pyodide_verify.mjs`(웹 검증 재현).
- 엔진은 `리포 루트의 drilldown_table.py(단일 소스)`(v14 단일 소스)를 빌드시 번들/주입.
