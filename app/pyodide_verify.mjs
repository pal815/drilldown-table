// Pyodide headless 검증: 기존 엔진(drilldown_table.py)이 브라우저 파이썬에서 그대로
// .xlsx/.docx/.pptx 를 만드는지(특히 lxml 의존 python-docx/pptx) 확인.
// node는 바이너리 휠을 로컬에 안 두므로, lock에서 의존성 클로저를 풀어 CDN→로컬 dist로 받은 뒤 기본 로더로 실행.
import { loadPyodide } from "pyodide";
import { readFileSync, writeFileSync, existsSync } from "node:fs";
import { createRequire } from "node:module";

const require = createRequire(import.meta.url);
const DIST = require.resolve("pyodide/package.json").replace(/package\.json$/, "");
const CDN = "https://cdn.jsdelivr.net/pyodide/v314.0.0/full/";
const lock = JSON.parse(readFileSync("./_lock.json", "utf8")).packages;

// 의존성 클로저 (소문자 매칭)
const key = n => lock[n] ? n : Object.keys(lock).find(k => k.toLowerCase() === n.toLowerCase());
const closure = new Set(), q = ["micropip", "lxml", "Pillow"];
while (q.length) {
  const k = key(q.pop()); if (!k || closure.has(k)) continue;
  closure.add(k); for (const d of (lock[k].depends || [])) q.push(d);
}
console.log("[0] 클로저:", [...closure].join(", "));

// 누락 휠 다운로드
for (const k of closure) {
  const fn = lock[k].file_name;
  const dst = DIST + fn;
  if (existsSync(dst)) continue;
  const r = await fetch(CDN + fn);
  if (!r.ok) { console.log("  다운로드 실패", fn, r.status); continue; }
  writeFileSync(dst, Buffer.from(await r.arrayBuffer()));
  console.log("  받음:", fn);
}

console.log("[1] Pyodide 로드(로컬 dist)…");
const py = await loadPyodide();
console.log("[2] 바이너리 패키지 로드(micropip/lxml/Pillow)…");
await py.loadPackage(["micropip", "lxml", "Pillow"]);
console.log("[3] micropip 로 openpyxl/python-docx/python-pptx/XlsxWriter 설치(PyPI)…");
await py.pyimport("micropip").install(["openpyxl", "python-docx", "python-pptx", "XlsxWriter"]);

console.log("[4] 엔진 주입 + 6개 생성(row/both x xlsx/docx/pptx)…");
py.FS.writeFile("/drilldown_table.py",
  readFileSync("C:/Users/pal81/.claude/skills/drilldown-table/drilldown_table.py", "utf8"));
const out = py.runPython(`
import sys, os
sys.path.insert(0, "/")
import drilldown_table as G
for ext in ("xlsx","docx","pptx"):
    G.generate(G.SAMPLE,   "/row_"+ext+"."+ext,  "row",  "color")
    G.generate(G.SAMPLE_2D, "/both_"+ext+"."+ext, "both", "grey")
str({f: os.path.getsize("/"+f) for f in sorted(os.listdir("/")) if f.startswith(("row_","both_"))})
`);
console.log("[5] 생성 결과:", out);
console.log("PYODIDE_VERIFY_OK");
