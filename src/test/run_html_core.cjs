// 배포 HTML 안의 변환 코드(dsd-core)를 꺼내 Node에서 18개 샘플을 변환한다.
const fs = require("fs"), path = require("path"), vm = require("vm");
const JSZip = require("jszip"), ExcelJS = require("exceljs"), { DOMParser } = require("@xmldom/xmldom");
const [html, srcDir, outDir] = process.argv.slice(2);
const m = /<script id="dsd-core">([\s\S]*?)<\/script>/.exec(fs.readFileSync(html, "utf8"));
vm.runInThisContext(m[1]);
(async () => {
  fs.mkdirSync(outDir, { recursive: true });
  const files = fs.readdirSync(srcDir).filter((f) => f.toLowerCase().endsWith(".dsd")).sort();
  for (let i = 0; i < files.length; i++) {
    const buf = fs.readFileSync(path.join(srcDir, files[i]));
    const t = Date.now();
    try {
      const res = await DSDCore.convert(buf, files[i], { JSZip, ExcelJS, DOMParser, now: new Date(2026, 9, 6, 12, 0) });
      fs.writeFileSync(path.join(outDir, `j${String(i).padStart(2, "0")}.xlsx`), res.data);
      const ck = res.checks ? Object.entries(res.checks).map(([k, v]) => `${k}${v}`).join(" ") : "";
      console.log(`${Date.now() - t}ms 시트${res.sheets.length} 주석${res.notes} [${ck}] ${files[i].slice(0, 26)}`);
    } catch (e) { console.log("실패", files[i], e.stack); }
  }
})();
