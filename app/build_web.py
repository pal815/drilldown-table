# -*- coding: utf-8 -*-
"""
build_web.py — 웹 폴백(단일 HTML) 생성기.

기존 Python 엔진(drilldown_table.py)과 excel_io.py를 HTML에 '주입'해
drilldown_web.html 한 파일을 만든다. 브라우저에서 Pyodide로 그 엔진을 그대로 실행 →
JS 재포팅 없이 동일 v14 양식을 .xlsx/.docx/.pptx 로 생성·다운로드.

엔진이 바뀌면 이 스크립트만 다시 돌리면 HTML이 갱신됨(단일 소스 유지).
잠긴 인트라넷/오프라인은 Pyodide 자산을 사내 호스팅(아래 PYODIDE_BASE 교체)하면 됨.
"""
import json
import os

SKILL = os.path.join(os.path.dirname(os.path.abspath(__file__)), os.pardir)
APP = os.path.dirname(os.path.abspath(__file__))
PYODIDE_BASE = "https://cdn.jsdelivr.net/pyodide/v314.0.0/full/"  # 사내 오프라인이면 인트라넷 경로로 교체

engine = open(os.path.join(SKILL, "drilldown_table.py"), encoding="utf-8").read()
excelio = open(os.path.join(APP, "excel_io.py"), encoding="utf-8").read()

HTML = r"""<!doctype html>
<html lang="ko">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>드릴다운 표 생성기 (웹)</title>
<style>
  body{font-family:"맑은 고딕",system-ui,sans-serif;max-width:680px;margin:24px auto;padding:0 16px;color:#1a1a1a}
  h1{font-size:20px;margin:0 0 4px} .sub{color:#666;margin:0 0 16px}
  fieldset{border:1px solid #d6d6d6;border-radius:8px;margin:12px 0;padding:12px 14px}
  legend{font-weight:700;padding:0 6px}
  button{font:inherit;padding:8px 14px;border:1px solid #1f6feb;background:#1f6feb;color:#fff;border-radius:6px;cursor:pointer}
  button.sec{background:#fff;color:#1f6feb}
  button:disabled{opacity:.5;cursor:not-allowed}
  label.opt{display:inline-block;margin:2px 12px 2px 0}
  #status{padding:10px 12px;border-radius:6px;background:#eef4ff;color:#1f4ea8;margin:8px 0;font-size:14px}
  #log{white-space:pre-wrap;font-size:13px;color:#444;margin-top:8px}
  .file{margin:6px 0}
  #go{padding:10px 28px;font-size:16px;margin-top:6px}
</style>
</head>
<body>
<h1>드릴다운(ㄱ자/코너헤더) 표 생성기 — 웹</h1>
<p class="sub">설치 0. 브라우저에서 엔진(Python)을 그대로 실행해 Excel·Word·PPT 표를 만듭니다. 데이터는 PC 밖으로 나가지 않습니다.</p>

<div id="status">초기화 중… 최초 1회 엔진 로딩에 수십 초 걸립니다.</div>

<fieldset>
  <legend>1. 입력양식</legend>
  <button id="tpl" class="sec" disabled>① 입력양식(.xlsx) 받기</button>
  <div class="file"><label>② 채운 입력양식 불러오기: <input type="file" id="file" accept=".xlsx" disabled></label></div>
</fieldset>

<fieldset>
  <legend>2. 방향</legend>
  <label class="opt"><input type="radio" name="orient" value="row" checked> 행 기준</label>
  <label class="opt"><input type="radio" name="orient" value="column"> 열 기준</label>
  <label class="opt"><input type="radio" name="orient" value="both"> 행+열(2D)</label>
</fieldset>

<fieldset>
  <legend>3. 테마</legend>
  <label class="opt"><input type="radio" name="theme" value="color" checked> 컬러</label>
  <label class="opt"><input type="radio" name="theme" value="grey"> 그레이</label>
  <label class="opt"><input type="radio" name="theme" value="mono"> 테두리만</label>
</fieldset>

<fieldset>
  <legend>4. 출력 포맷 (복수)</legend>
  <label class="opt"><input type="checkbox" class="fmt" value="xlsx" checked> Excel</label>
  <label class="opt"><input type="checkbox" class="fmt" value="docx"> Word</label>
  <label class="opt"><input type="checkbox" class="fmt" value="pptx"> PowerPoint</label>
</fieldset>

<button id="go" disabled>표 생성</button>
<div id="log"></div>

<script src="__PYODIDE_JS__"></script>
<script>
const ENGINE_SRC = __ENGINE__;
const EXCELIO_SRC = __EXCELIO__;
const $ = s => document.querySelector(s);
const status = m => $("#status").textContent = m;
const log = m => $("#log").textContent += m + "\n";

function download(bytes, name){
  const blob = new Blob([bytes], {type:"application/octet-stream"});
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob); a.download = name; a.click();
  setTimeout(()=>URL.revokeObjectURL(a.href), 4000);
}

let py;
async function init(){
  py = await loadPyodide({indexURL:"__PYODIDE_BASE__"});
  status("패키지 로딩(lxml·openpyxl·Pillow)…");
  await py.loadPackage(["micropip","lxml","Pillow"]);
  status("openpyxl·python-docx·python-pptx 설치…");
  const mp = py.pyimport("micropip");
  await mp.install(["openpyxl","python-docx","python-pptx","XlsxWriter"]);
  py.FS.writeFile("/drilldown_table.py", ENGINE_SRC);
  py.FS.writeFile("/excel_io.py", EXCELIO_SRC);
  py.runPython('import sys; sys.path.insert(0,"/")');
  status("준비 완료 — ① 입력양식 받기 → 채워서 ② 불러오기 → [표 생성]");
  $("#tpl").disabled = false; $("#file").disabled = false; $("#go").disabled = false;
}

$("#tpl").onclick = () => {
  py.runPython('import excel_io; excel_io.make_template("/입력양식.xlsx")');
  download(py.FS.readFile("/입력양식.xlsx"), "입력양식.xlsx");
  log("입력양식.xlsx 내려받음 — 시트(행·열 / 행+열)에 데이터를 채워 ②로 불러오세요.");
};

$("#go").onclick = async () => {
  const f = $("#file").files[0];
  if(!f){ alert("채운 입력양식(.xlsx)을 먼저 불러오세요."); return; }
  const orient = document.querySelector('input[name=orient]:checked').value;
  const theme  = document.querySelector('input[name=theme]:checked').value;
  const fmts   = [...document.querySelectorAll('.fmt:checked')].map(c=>c.value);
  if(!fmts.length){ alert("출력 포맷을 하나 이상 선택하세요."); return; }
  try{
    status("생성 중…");
    py.FS.writeFile("/in.xlsx", new Uint8Array(await f.arrayBuffer()));
    py.globals.set("ORIENT", orient); py.globals.set("THEME", theme);
    py.globals.set("FMTS", py.toPy(fmts));
    const names = py.runPython(`
import excel_io, drilldown_table as G
m = excel_io.read_model("/in.xlsx", ORIENT)
out=[]
for ext in FMTS:
    p = "/드릴다운표_%s_%s.%s" % (ORIENT, THEME, ext)
    G.generate(m, p, ORIENT, THEME); out.append(p)
out
`).toJs();
    for(const p of names){
      const name = p.split("/").pop();
      download(py.FS.readFile(p), name);
      log("생성: " + name);
    }
    status("완료 — 다운로드 폴더를 확인하세요.");
  }catch(e){
    status("오류: " + e); log(String(e));
  }
};

init().catch(e => { status("초기화 실패: " + e); log(String(e)); });
</script>
</body>
</html>
"""

out = (HTML
       .replace("__PYODIDE_JS__", PYODIDE_BASE + "pyodide.js")
       .replace("__PYODIDE_BASE__", PYODIDE_BASE)
       .replace("__ENGINE__", json.dumps(engine))
       .replace("__EXCELIO__", json.dumps(excelio)))

dst = os.path.join(APP, "drilldown_web.html")
with open(dst, "w", encoding="utf-8") as f:
    f.write(out)
print("saved", dst, "(", round(len(out) / 1024), "KB )")
