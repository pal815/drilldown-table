# drilldown-table — 드릴다운(┌ / 코너헤더) 계층형 표 생성기

상위 항목을 좌상단 코너에 두고 세부(drill-through)를 옆 열·아래 행으로 **┌ 모양**으로 펼쳐 묶는,
기업 보고서·실적 자료식 **계층형 표(일명 "ㄱ자 표")** 를 **Excel(.xlsx)·Word(.docx)·PowerPoint(.pptx)** 로 생성합니다.

**데이터만 주면** 동일한 v14 양식으로 3종 문서를 만들고, **방향 3 × 테마 3**을 옵션으로 받습니다.

- **방향** — `row`(항목 세로) · `column`(항목 가로) · `both`(행·열 2D 교차표)
- **테마** — `color`(네이비/연파랑) · `grey`(그레이스케일) · `mono`(테두리만)
- **다단계** — 대분류>중분류>… 깊이 2~4단계 자동 인식(행·열), 합계 줄 생략 시 자동 합산, 가지별 깊이 달라도 됨
- 단일 IR(공용 중간표현) → 3 렌더러. 양식은 v14로 동결.

> 샘플 출력은 [`samples/`](samples/)에, GUI 미리보기 썸네일은 [`app/assets/previews/`](app/assets/previews/)에 있습니다.

---

## 네 가지 사용 방법

이 리포는 **하나의 엔진**([`drilldown_table.py`](drilldown_table.py))을 여러 방식으로 감쌉니다.

### 1) Claude Code / Codex 스킬 (플러그인)
리포 루트가 그대로 Claude Code 스킬·플러그인입니다([`SKILL.md`](SKILL.md), [`.claude-plugin/plugin.json`](.claude-plugin/plugin.json), [`commands/`](commands/)).
- 개인 스킬: 리포를 `~/.claude/skills/drilldown-table`로 두거나 플러그인으로 설치.
- "사업부별 실적 표 만들어 / 드릴다운 표 / ㄱ자 표" 같은 요청에 발동.
- 데이터 스키마는 [`examples/`](examples/)의 `sample_column.yaml`(row·column) / `sample_both.yaml`(both) 참고.

### 2) 데스크톱 앱 (EXE) — 언어모델·파이썬 불필요
일반 사무직용 GUI. 엑셀 첨부 또는 **클립보드 붙여넣기** → 방향·테마·포맷 선택 → 생성.
DRM/다중시트 엑셀은 **Excel COM**으로 자동 처리. 빌드·배포는 [`app/`](app/) 참조.
```powershell
cd app
.\build.ps1          # PyInstaller onedir EXE (창모드). -OneFile / -Console 옵션
```

### 3) 웹 (브라우저, 설치 0) — 잠긴 PC용
단일 HTML 한 파일을 브라우저에서 열면 **Pyodide로 동일 엔진**이 돕니다(데이터는 PC 밖으로 안 나감).
```powershell
cd app
py build_web.py      # → app/drilldown_web.html 생성(엔진 바뀌면 다시 실행)
```

### 4) Python에서 직접
```python
import drilldown_table as G
G.generate(model, "out.xlsx", "row", "color")          # 단일
# 또는 CLI:  py drilldown_table.py data.yaml out.pptx column grey
```

---

## 리포 구조

```
drilldown-table/
├─ drilldown_table.py        # ★ 엔진(단일 소스) — IR + 3 렌더러 + 레이아웃 빌더
├─ SKILL.md                  # Claude/Codex 스킬 정의
├─ .claude-plugin/           # 플러그인 매니페스트
├─ commands/                 # /drilldown-table 슬래시 커맨드
├─ examples/                 # 입력 데이터 스키마 예시(YAML)
├─ samples/                  # 쇼케이스 출력물(v14 × 3테마 × 3포맷)
├─ render_png.py / render_pptx_png.py / render_docx_png.py   # 검수용 PNG 렌더
└─ app/                      # 독립 배포 wrapper(EXE + 웹)
   ├─ app.py · excel_io.py   # GUI/CLI + 입력양식↔모델 어댑터
   ├─ build.ps1 · build_web.py · sign.ps1
   ├─ assets/                # GUI 아이콘·미리보기 썸네일
   ├─ RELEASE.md · 사용법.txt · 웹버전_안내.txt
   └─ make_*.py · pyodide_verify.mjs
```

**엔진 단일 소스**: `drilldown_table.py`가 정본이며, `app/`(빌드 시 번들)과 Claude/Codex 스킬 설치본이 이를 사용합니다.
별도 위치(`~/.claude/skills/…`, `~/.codex/skills/…`)에 설치된 사본은 [`sync.ps1`](sync.ps1)로 동기화합니다:
```powershell
.\sync.ps1            # Claude + Codex 둘 다 (엔진·예시·렌더, Claude는 SKILL/plugin 포함)
```

## 요구 사항
- Python 3.10+ (개발/스킬 사용 시). 끝사용자는 EXE/웹이면 **불필요**.
- 패키지: `openpyxl`, `python-docx`, `python-pptx`, `PyYAML`(스킬 입력), `customtkinter`·`Pillow`(앱 GUI), `pywin32`(Excel COM·검수 렌더).
- Excel COM 입력/PNG 검수 렌더는 Windows + Microsoft Excel 필요.

## 라이선스
[MIT](LICENSE)
