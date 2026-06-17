---
description: 드릴다운(corner-header / ┌) 계층형 표를 Excel/Word/PPT로 생성 (방향·테마 옵션).
argument-hint: <데이터 설명 또는 데이터 파일 경로> [orient: row|column|both] [theme: color|grey|mono] [format: xlsx|docx|pptx]
---

드릴다운(corner-header / ┌, 일명 ㄱ자) 표를 만든다: $ARGUMENTS

`drilldown-table` 스킬을 그대로 따른다. 스킬 루트 `$SKILL` = `${CLAUDE_PLUGIN_ROOT}`(플러그인) 또는 레포 디렉터리.
생성 파일은 현재 작업 디렉터리에 쓴다.

1. 요청에서 **orient**(row/column/both)·**theme**(color/grey/mono)·**format**(xlsx/docx/pptx)을 파싱한다.
   불명확하면 기본값(row·color·xlsx)으로 진행하되 한 줄로 알린다. 데이터가 모호하면 먼저 질문한다.
2. 데이터를 `$SKILL/examples/`의 스키마에 맞춰 `data.yaml`로 작성한다
   (row/column → `sample_column.yaml`, both → `sample_both.yaml`). 퍼센트는 `"21%"` 문자열, 계/총계는 both에서 자동 합산.
3. 생성 — **xlsx·docx는 프리빌트 바이너리(Python 불필요), pptx는 Python**:
   - xlsx/docx: `"$SKILL/bin/<platform>/drilldown-table[.exe]" data.yaml OUT.<ext> <orient> <theme>` (플랫폼 분기는 SKILL.md §3 참조). 바이너리 없으면 `py "$SKILL/drilldown_table.py" …`로 폴백.
   - pptx: `py "$SKILL/drilldown_table.py" data.yaml OUT.pptx <orient> <theme>`.
   - 여러 방향/테마/시트(Python): `import drilldown_table as G` 후 `G.generate(...)` 또는 `G.build_workbook([...], path, theme)`.
4. 검수: `py "$SKILL/render_png.py|render_pptx_png.py|render_docx_png.py" OUT` 로 PNG 렌더 → 직접 확인.
   필요한 패키지(openpyxl/python-docx/python-pptx/PyYAML)가 없으면 `py -m pip install` 안내.
5. 결과(파일 경로 + 사용한 옵션 + 렌더 이미지)를 보고한다.

양식은 v14 확정본이다. 사용자가 양식 변경을 요청하면 `drilldown_table.py`의 레이아웃 빌더만 수정하고 IR/렌더 구조는 보존한다.
