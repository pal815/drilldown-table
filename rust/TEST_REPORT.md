# drilldown-table Rust 포트 — 테스트 리포트

## 상태: 통과 (xlsx + docx)

- `cargo build --release` 성공 → `target/release/drilldown-table.exe` (~1.46 MB). 배포본 = `bin/windows-x64/`.
- **IR 15/15 MATCH** — `--dump-ir` 출력이 Python 엔진의 기준 IR(`fixtures/ir/`)과 셀 단위로 동일.
  레이아웃·다단계(depth 2~4)·테마(color/grey/mono)·서식(천단위·퍼센트) 패리티 입증.
- **xlsx 렌더 15/15** 무오류. 행/다단계/2D × 3테마를 Excel COM PNG로 Python 베이스라인과 시각 대조 → 동일.
- **docx 렌더 15/15** 무오류. 다단계를 Word COM PDF로 v14와 대조 → 동일(navy 헤더·세로 띠·안쪽 ┌·음영·맑은 고딕).
- **pptx**: 의도적 미지원. 바이너리는 `pptx is not supported in the Rust build; use the Python engine.`로 거부(rc=1) → 스킬이 Python 엔진으로 폴백.

## 검증 매트릭스 (IR / xlsx / docx)

| Fixture | IR | xlsx | docx |
|---|---|---|---|
| twolevel_{row,column} × {color,grey,mono} | MATCH | OK | OK |
| both_both × {color,grey,mono} | MATCH | OK | OK |
| multilevel_{row,column} × {color,grey,mono} | MATCH | OK | OK |

## 의도적 편차

- `w:noWrap`은 docx-rs가 미지원이라 생략(고정 레이아웃 + 명시 열폭으로 줄바꿈 억제). 시각 영향 미미.
- Excel 아웃라인(드릴 +/- 버튼)은 미구현(스타일 무관). 필요 시 후속.

## 재현
```
cargo build --release
# IR: drilldown-table <fixtures/inputs/NAME.yaml> out.xlsx <orient> <theme> --dump-ir | py fixtures/compare_ir.py - fixtures/ir/TAG.json
```
