// DSD엑셀변환.html 만들기: template.html에 라이브러리(src/lib의 JSZip 3.10.2, ExcelJS 4.4.0)와 dsd-core.js·dsd-checks.js를 넣어 한 파일로 묶는다.
// 실행: node src/build.cjs
// 라이브러리는 인터넷 없이 쓰도록 파일 안에 그대로 넣는다(JSZip MIT/GPLv3 이중, ExcelJS MIT. 머리말 유지, 원문은 src/lib/LICENSE-*).
const fs = require("fs");
const path = require("path");

const here = __dirname;
const read = (p) => fs.readFileSync(path.join(here, p), "utf8");
const strip = (s) => s.replace(/\n\/\/# sourceMappingURL=.*$/m, "");   // 원본 지도 파일 참조 제거
const safe = (s) => s.replace(/<\/script/gi, "<\\/script");             // 스크립트 안의 </script> 방지

const d = new Date();
const stamp = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;

const html = read("template.html")
  .replace("/*__JSZIP__*/", () => safe(strip(read("lib/jszip.min.js"))))
  .replace("/*__EXCELJS__*/", () => safe(strip(read("lib/exceljs.min.js"))))
  .replace("/*__CORE__*/", () => safe(read("dsd-core.js") + "\n" + read("dsd-checks.js")))
  .replace("__BUILD_DATE__", stamp)
  .replace("__SAMPLE_B64__", () => fs.readFileSync(path.join(here, "sample", "샘플전자_감사보고서_2025_예시.dsd")).toString("base64"));   // [예시 파일로 해 보기]용 가상 회사 감사보고서

const out = path.join(here, "..", "DSD엑셀변환.html");
fs.writeFileSync(out, html, "utf8");
console.log("만듦:", out, Math.round(fs.statSync(out).size / 1024), "KB");
