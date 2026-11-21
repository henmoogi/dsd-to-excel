// 배포 HTML 안의 변환 코드로 폴더의 .dsd를 '한 번에 넣은 묶음'처럼 변환한다(같은 회사 두 해 파일은 전기 대사).
// 사용: node src/test/run_prior.cjs DSD엑셀변환.html <dsd폴더> <출력폴더>
//  → j00.xlsx… (run_html_core.cjs와 같은 번호), prior_jNN.json(전기 대사 결과), 화면에 파일별 요약
const fs = require("fs"), path = require("path"), vm = require("vm");
const JSZip = require("jszip"), ExcelJS = require("exceljs"), { DOMParser } = require("@xmldom/xmldom");
const [html, srcDir, outDir] = process.argv.slice(2);
const m = /<script id="dsd-core">([\s\S]*?)<\/script>/.exec(fs.readFileSync(html, "utf8"));
vm.runInThisContext(m[1]);
(async () => {
  fs.mkdirSync(outDir, { recursive: true });
  const files = fs.readdirSync(srcDir).filter((f) => f.toLowerCase().endsWith(".dsd")).sort();
  const items = files.map((f) => ({ name: f, read: async () => fs.readFileSync(path.join(srcDir, f)) }));
  const deps = { JSZip, ExcelJS, DOMParser, now: new Date(2026, 9, 6, 12, 0) };
  const t0 = Date.now();
  await DSDCore.convertAll(items, deps, {
    onDone: async (i, res, err) => {
      const tag = `j${String(i).padStart(2, "0")}`;
      if (err) { console.log(tag, "실패", files[i], err.stack); return; }
      fs.writeFileSync(path.join(outDir, `${tag}.xlsx`), res.data);
      const ck = res.checks ? Object.entries(res.checks).map(([k, v]) => `${k}${v}`).join(" ") : "";
      console.log(`${tag} 시트${res.sheets.length} 주석${res.notes} [${ck}] ${files[i].slice(0, 26)}`);
      if (res.pairNote) console.log("     짝:", res.pairNote);
      if (res.prior) {
        console.log(`     전기 대사(작년 ${files.indexOf(res.priorName)}번): ${res.prior.sentence} ${res.prior.rowsLine}`);
        fs.writeFileSync(path.join(outDir, `prior_${tag}.json`), JSON.stringify(res.prior, null, 1));
      }
    },
  });
  console.log(`전체 ${Date.now() - t0}ms`);
})();
