/* DSD → 엑셀 변환 핵심 (브라우저·Node 공용)
 *
 * .dsd(DART 편집기 파일)는 ZIP이고, 안의 contents.xml에 본문이 있다.
 * 사용: DSDCore.convert(arrayBuffer, 파일이름, {JSZip, ExcelJS, DOMParser, now})
 *   → {data: Uint8Array(xlsx), sheets: [시트이름], notes: 주석제목수}
 *
 * 규칙은 처음 만든 Python 판(_보관_python/dsd_to_excel.py)과 같다. 고칠 때는 이 파일만 고친다.
 */
(function (root) {
  "use strict";
  const HOST = root;   // 전역(window). Converter 안의 root(XML)와 헷갈리지 않게 따로 둔다

  const FONT_NAME = "맑은 고딕";
  const F_NORMAL = { name: FONT_NAME, size: 10 };
  const F_BOLD = { name: FONT_NAME, size: 10, bold: true };
  const F_TITLE = { name: FONT_NAME, size: 12, bold: true };
  const F_LINK = { name: FONT_NAME, size: 10, color: { argb: "FF0563C1" }, underline: true };
  const F_NOTE = { name: FONT_NAME, size: 9, color: { argb: "FF666666" } };
  const THIN = { style: "thin", color: { argb: "FF808080" } };
  const BOX = { top: THIN, left: THIN, bottom: THIN, right: THIN };
  const HEAD_FILL = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF2F2F2" } };
  const STYLES = { F_NORMAL, F_BOLD, F_TITLE, F_LINK, F_NOTE, BOX, HEAD_FILL, FONT_NAME };

  // 출력하지 않는 태그: 편집기 안내문, 라이브러리 이름, 쪽 나눔, 원문 목차(목차 시트를 따로 만든다)
  const SKIP = new Set(["COMMENT", "WARNING", "FILENAME", "LIBRARYLIST", "PGBRK", "TOC", "DOCUMENT-HEADER", "COLGROUP", "COL"]);
  const CELL_TAGS = new Set(["TD", "TH", "TE", "TU"]);
  const BLOCKS = new Set(["SECTION-1", "SECTION-2", "SECTION-3", "TABLE-GROUP"]);
  const PAGE_PX = 600;      // DART 본문 폭(px). 문단은 이 폭만큼 셀을 합쳐 줄바꿈한다
  const PX_PER_CHAR = 7.0;  // 엑셀 열 너비 1 ≈ 7px
  const LINE_PT = 16.5;     // 맑은 고딕 10pt 한 줄 높이
  const H_ALIGN = { left: "left", center: "center", right: "right", fill: "fill", justify: "justify",
                    centercontinuous: "centerContinuous", distributed: "distributed" };

  const NUM_RE = /^(\()?(-)?([\d,]*\d)(\.\d+)?(\))?(%)?$/;
  const FS_NAME_RE = /^(연결)?(요약)?(반기|분기)?(재무상태표|대차대조표|포괄손익계산서|손익계산서|자본변동표|현금흐름표|이익잉여금처분계산서|결손금처리계산서)$/;
  const NOTE_RE = /^(\d{1,2})\s*\.\s*(?=\D)/;

  // ------------------------------------------------------------ DOM 도우미
  const elems = (e) => { const out = []; for (let n = e.firstChild; n; n = n.nextSibling) if (n.nodeType === 1) out.push(n); return out; };
  const attr = (e, k) => e.getAttribute(k) || "";
  const marks = (e) => attr(e, "USERMARK").split(/\s+/).filter(Boolean);
  const cps = (s) => Array.from(s);
  const round1 = (x) => Math.round(x * 10) / 10;

  // ------------------------------------------------------------ 글자 처리
  /** 요소를 줄 단위로 나눈다 → [[줄 글자, 줄 첫머리가 굵은지]].
   *  DART 편집기의 줄바꿈은 '&cr;'이고, XML 원문의 실제 줄바꿈은 태그 사이 정리용이라 버린다. 하위 P는 줄을 나눈다. */
  function rawText(e) {
    const segs = [];
    (function rec(x, bold) {
      const b = bold || marks(x).includes("B");
      for (let n = x.firstChild; n; n = n.nextSibling) {
        if (n.nodeType === 3 || n.nodeType === 4) {
          if (n.nodeValue) segs.push([n.nodeValue, b, true]);
        } else if (n.nodeType === 1 && !SKIP.has(n.nodeName)) {
          const isP = n.nodeName === "P";
          if (isP) segs.push(["\n", b, false]);
          rec(n, b);
          if (isP) segs.push(["\n", b, false]);
        }
      }
    })(e, false);
    const lines = [];
    let cur = "", curBold = null;
    for (const [t0, b, fromXml] of segs) {
      const text = fromXml ? t0.replace(/[\r\n]+/g, "").split("&cr;").join("\n") : t0;
      text.split("\n").forEach((part, i) => {
        if (i) { lines.push([cur, !!curBold]); cur = ""; curBold = null; }
        cur += part;
        if (curBold === null && part.trim()) curBold = b;
      });
    }
    lines.push([cur, !!curBold]);
    return lines.map(([t, bd]) => [t.replace(/[ \t　\xa0]{2,}/g, " ").trim(), bd]);
  }

  /** 빈 줄을 뺀 글자(표 셀·제목용). */
  const textOf = (e, sep = "\n") => rawText(e).map(([t]) => t).filter(Boolean).join(sep);

  /** 엑셀 열 너비 단위로 본 글자 폭(한글·전각은 2). */
  function widthUnits(text) {
    let n = 0;
    for (const ch of text) n += ch.codePointAt(0) > 0x2e7f ? 2 : 1;
    return n;
  }

  /** 표 셀 글자를 [값, 표시형식]으로. 숫자가 아니면 [글자, null]. */
  function toValue(text) {
    const t = text.trim().split(" ").join("");
    if (t === "-" || t === "－" || t === "—") return [0, '#,##0;(#,##0);"-"'];
    const m = NUM_RE.exec(t);
    if (!m) return [text, null];
    const [, lp, minus, ip, dec, rp, pct] = m;
    if (!!lp !== !!rp) return [text, null];
    const digits = ip.split(",").join("");
    const comma = ip.includes(",");
    if (comma && !/^\d{1,3}(,\d{3})*$/.test(ip)) return [text, null];
    if (!comma && !dec && (digits.length >= 8 || (digits.length > 1 && digits.startsWith("0")))) {
      return [text, null];   // 사업자번호·코드 등은 글자로 둔다
    }
    let val = Number(digits + (dec || ""));
    if (lp || minus) val = -val;
    const nd = dec ? dec.length - 1 : 0;
    const base = (comma ? "#,##0" : "0") + (nd ? "." + "0".repeat(nd) : "");
    if (pct) return [val / 100, base.split("#,##").join("") + "%"];
    if (lp) return [val, `${base};(${base})`];
    return [val, base];
  }

  /** 표시형식에서 최소 단위(반올림 허용 오차 계산용): '#,##0' → 1, '0.00' → 0.01, '0.0%' → 0.001 */
  function precOf(fmt) {
    if (!fmt) return 1;
    const m = /0\.(0+)/.exec(fmt);
    const nd = m ? m[1].length : 0;
    return Math.pow(10, -(nd + (fmt.includes("%") ? 2 : 0)));
  }

  /** '(단위 : 천원)' → '천원' */
  function parseUnit(text) {
    const m = /단위\s*[:：]?\s*([^)\s,，]+)/.exec(text);
    return m ? m[1] : null;
  }

  function sheetNameOf(title) {
    let name = title.replace(/\s+/g, "").replace(/^\(첨부\)/, "").replace(/[\[\]:*?\/\\]/g, "");
    if (name.includes("감사보고서") && name.includes("감사인")) name = "감사보고서";
    return cps(name || "시트").slice(0, 31).join("");
  }

  // ------------------------------------------------------------ 시트 쓰기
  class SheetWriter {
    constructor(ws) {
      this.ws = ws;
      this.row = 3;             // 1행: 목차로 돌아가기 링크, 2행: 빈 줄
      this.colPx = new Map();   // 열 번호 → 표에서 본 최대 폭(px)
      this.paragraphs = [];     // [행, 글자] — 나중에 셀 합치기·높이 계산
      this.blank = true;        // 직전 행이 빈 줄인지
      this.freeze = null;       // [xSplit, ySplit]
      this.maxRow = 0;
      this.used = false;
      this.unit = null;         // 마지막으로 본 '(단위: …)'
      this.tables = [];         // 검증용 표 구조: {sheet, top, unit, aclass, bordered, rows:[{r, head, cells:[{c,cs,rs,text,value,prec,head}]}]}
    }

    cellAt(r, c) {
      if (r > this.maxRow) this.maxRow = r;
      this.used = true;
      return this.ws.getCell(r, c);
    }

    put(r, c, value) {
      const cell = this.cellAt(r, c);
      cell.value = value;
      return cell;
    }

    merge(t, l, b, r) {
      try { this.ws.mergeCells(t, l, b, r); } catch (e) { /* 원문 표가 어긋나 겹치는 병합은 건너뛴다 */ }
    }

    gap() {
      if (!this.blank) { this.row += 1; this.blank = true; }
    }

    title(text, big = true) {
      if (!text) return null;
      this.gap();
      const c = this.put(this.row, 1, text);
      c.font = big ? F_TITLE : F_BOLD;
      this.row += 1;
      this.blank = false;
      return this.row - 1;
    }

    /** 문단의 한 줄을 한 행에 쓴다. 빈 줄은 연달아 나오면 하나로 줄인다. */
    line(text, bold = false) {
      if (!text) { this.gap(); return null; }
      const u = parseUnit(text);
      if (u) this.unit = u;
      const c = this.put(this.row, 1, text);
      c.font = bold ? F_BOLD : F_NORMAL;
      this.paragraphs.push([this.row, text]);
      this.row += 1;
      this.blank = false;
      return this.row - 1;
    }

    table(tbl) {
      const bordered = (attr(tbl, "BORDER") || "0") !== "0";
      const cols = tbl.getElementsByTagName("COL");
      for (let i = 0; i < cols.length; i++) {
        const w = parseInt(attr(cols[i], "WIDTH") || "0", 10) || 0;
        this.colPx.set(i + 1, Math.max(this.colPx.get(i + 1) || 0, w));
      }
      const rows = [];
      for (const part of elems(tbl)) {
        if (part.nodeName === "THEAD" || part.nodeName === "TBODY") {
          for (const tr of elems(part)) if (tr.nodeName === "TR") rows.push([tr, part.nodeName === "THEAD"]);
        } else if (part.nodeName === "TR") {
          rows.push([part, false]);
        }
      }
      if (!rows.length) return;
      if (!this.blank) this.row += 1;
      const top = this.row;
      const taken = new Set();
      let last = null;
      let bodyStart = null;
      const model = { sheet: this.ws.name, writer: this, top, unit: this.unit, ownUnit: null,
                      aclass: attr(tbl, "ACLASS"), bordered, rows: [] };
      rows.forEach(([tr, inHead], rOff) => {
        const r = top + rOff;
        if (!inHead && bodyStart === null) bodyStart = r;
        const rowModel = { r, head: inHead, cells: [] };
        model.rows.push(rowModel);
        let c = 1;
        for (const cell of elems(tr)) {
          if (!CELL_TAGS.has(cell.nodeName)) continue;
          while (taken.has(r + "," + c)) c += 1;
          const cs = parseInt(attr(cell, "COLSPAN") || "1", 10) || 1;
          const rs = parseInt(attr(cell, "ROWSPAN") || "1", 10) || 1;
          for (let dr = 0; dr < rs; dr++) {
            for (let dc = 0; dc < cs; dc++) taken.add((r + dr) + "," + (c + dc));
            if (last === null || r + dr > last) last = r + dr;
          }
          const text = textOf(cell);
          const head = inHead || cell.nodeName === "TH";
          const [value, fmt] = (c === 1 || head) ? [text, null] : toValue(text);
          rowModel.cells.push({ c, cs, rs, text, head, value: typeof value === "number" ? value : null, prec: precOf(fmt) });
          const u = text.length < 40 ? parseUnit(text) : null;
          if (u) model.ownUnit = u;
          const xc = this.put(r, c, value !== "" ? value : null);
          xc.font = head ? F_BOLD : F_NORMAL;
          if (fmt) xc.numFmt = fmt;
          let align = attr(cell, "ALIGN").toLowerCase();
          if (head && !align) align = "center";
          if (fmt && !align) align = "right";
          const al = { vertical: "middle",
                       wrapText: typeof value === "string" && (value.includes("\n") || widthUnits(value) > 30) };
          if (H_ALIGN[align]) al.horizontal = H_ALIGN[align];
          xc.alignment = al;
          const fill = /BC0X([0-9A-Fa-f]{6})/.exec(attr(cell, "USERMARK"));
          if (fill) xc.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + fill[1].toUpperCase() } };
          else if (head && bordered) xc.fill = HEAD_FILL;
          if (cs > 1 || rs > 1) this.merge(r, c, r + rs - 1, c + cs - 1);
          if (bordered) {
            for (let dr = 0; dr < rs; dr++) for (let dc = 0; dc < cs; dc++) this.cellAt(r + dr, c + dc).border = BOX;
          }
          c += cs;
        }
      });
      if (last === null) last = top;
      if (model.ownUnit) { model.unit = model.ownUnit; this.unit = model.ownUnit; }   // 재무제표는 머리 표에 단위가 있다
      this.tables.push(model);
      const isFs =attr(tbl, "ACLASS") === "FINANCE" || (bordered && FS_NAME_RE.test(this.ws.name.replace(/\(\d+\)$/, "")));
      if (isFs && this.freeze === null && bodyStart && bodyStart > top) this.freeze = [1, bodyStart - 1];
      this.row = last + 1;
      this.blank = false;
    }

    /** 열 너비·문단 줄바꿈·틀 고정을 정하고, 1행에 목차 링크를 단다.
     *  heads: 주석 번호 제목 행들. 제목 아래 내용을 제목별로 묶는다(그룹, 수준 1). */
    finish(heads) {
      const ws = this.ws;
      const back = this.put(1, 1, { text: "◀ 목차로 돌아가기", hyperlink: "'목차'!A1" });
      back.font = F_LINK;
      heads = [...heads].sort((a, b) => a - b);
      heads.forEach((h, i) => {
        const end = i + 1 < heads.length ? heads[i + 1] - 1 : this.maxRow;
        for (let r = h + 1; r <= end; r++) {
          const row = ws.getRow(r);
          row.outlineLevel = 1;
          if (!row.hasValues && !row.height) ws.getCell(r, 1).font = F_NORMAL;   // 빈 행도 저장되어야 접힌다
        }
      });
      if (heads.length) ws.properties.outlineProperties = { summaryBelow: false, summaryRight: true };

      const ncol = this.colPx.size ? Math.max(...this.colPx.keys()) : 1;
      const widths = {};
      for (let i = 1; i <= ncol; i++) {
        const px = this.colPx.get(i) || 100;
        widths[i] = Math.max(8, Math.min(60, round1(px / PX_PER_CHAR)));
        ws.getColumn(i).width = widths[i];
      }
      // 문단: 본문 폭(600px)만큼 셀을 합치고 줄바꿈, 높이는 글자 수로 어림한다
      let span = 1, acc = widths[1];
      while (acc * PX_PER_CHAR < PAGE_PX && span < ncol) { span += 1; acc += widths[span]; }
      if (acc * PX_PER_CHAR < PAGE_PX) {   // 열이 모자라면 마지막 열을 넓힌다
        const extra = round1(PAGE_PX / PX_PER_CHAR - acc);
        widths[span] += extra;
        acc += extra;
        ws.getColumn(span).width = widths[span];
      }
      for (const [r, text] of this.paragraphs) {
        if (span > 1) this.merge(r, 1, r, span);
        ws.getCell(r, 1).alignment = { wrapText: true, vertical: "top" };
        const lines = text.split("\n").reduce((s, ln) => s + Math.max(1, Math.ceil(widthUnits(ln) / Math.max(acc - 1, 1))), 0);
        if (lines > 1) ws.getRow(r).height = round1(lines * LINE_PT);
      }
      const [xs, ys] = this.freeze || [0, 1];   // 재무제표는 머리글 아래, 나머지는 1행(목차 링크) 고정
      ws.views = [{ state: "frozen", xSplit: xs, ySplit: ys, showGridLines: false }];
    }
  }

  // ------------------------------------------------------------ 문서 순회
  class Converter {
    constructor(root, ExcelJS) {
      this.root = root;
      this.wb = new ExcelJS.Workbook();
      this.tocWs = this.wb.addWorksheet("목차");
      this.checkWs = HOST.DSDChecks ? this.wb.addWorksheet("검증") : null;   // 합계·정합성 검증 결과(목차 다음)
      this.checkSummary = null;
      this.writers = [];   // SheetWriter
      this.notes = [];     // [시트이름, 행, 제목, 굵게 여부, 번호]
      this.cur = null;
      this.newSheet("표지");
    }

    newSheet(name) {
      const base = name;
      const used = new Set(this.writers.map((w) => w.ws.name).concat(["목차", "검증"]));
      let n = 2;
      while (used.has(name)) {
        const suffix = `(${n})`;
        name = cps(base).slice(0, 31 - suffix.length).join("") + suffix;
        n += 1;
      }
      const w = new SheetWriter(this.wb.addWorksheet(name));
      this.writers.push(w);
      this.cur = w;
      return w;
    }

    walk(e, skipNode) {
      for (const c of elems(e)) {
        if (c === skipNode) continue;
        const tag = c.nodeName;
        if (SKIP.has(tag)) continue;
        if (BLOCKS.has(tag)) {
          const t = elems(c).find((x) => x.nodeName === "TITLE") || null;
          if (t && attr(t, "ATOC") === "Y") {
            const title = textOf(t, " ");
            this.newSheet(sheetNameOf(title)).title(title);
          } else if (t) {
            this.cur.title(textOf(t, " "), false);
          }
          this.walk(c, t);
        } else if (tag === "TITLE" || tag === "COVER-TITLE") {
          this.cur.title(textOf(c, " "));
        } else if (tag === "P" || tag === "SPAN") {
          for (const [text, bold] of rawText(c)) {
            const r = this.cur.line(text, bold);
            const m = NOTE_RE.exec(text);
            if (r && m && cps(text).length <= 60) this.notes.push([this.cur.ws.name, r, text, bold, parseInt(m[1], 10)]);
          }
        } else if (tag === "TABLE") {
          const fs = Converter.fsTitle(c);
          if (fs && !this.cur.ws.name.includes("주석") && this.cur.ws.name !== sheetNameOf(fs)) this.newSheet(sheetNameOf(fs));
          this.cur.table(c);
        } else {
          this.walk(c);   // DOCUMENT, BODY, COVER, INSERTION, LIBRARY 등 감싸는 태그
        }
      }
    }

    /** 제목 그룹 없이 표 첫 칸에 재무제표 이름을 쓴 경우(예: 자산운용사 서식) 그 이름. */
    static fsTitle(tbl) {
      const all = tbl.getElementsByTagName("*");
      for (let i = 0; i < all.length; i++) {
        if (CELL_TAGS.has(all[i].nodeName)) {
          const name = textOf(all[i]).replace(/\s+/g, "");
          return FS_NAME_RE.test(name) ? name : null;
        }
      }
      return null;
    }

    /** 주석 번호 제목(1. 2. 3. …)을 순서대로 골라낸다.
     *  같은 번호 후보가 여럿이면(본문 속 나열 번호 등) 첫머리가 굵은 문단을 우선한다. 번호 하나가 빠지면 한 번은 건너뛴다. */
    noteIndex() {
      const cand = this.notes.filter((n) => n[0].includes("주석"));
      const picked = [];
      let pos = 0, expect = 1, skipped = false;
      for (;;) {
        let nxtBold = cand.length;
        for (let i = pos; i < cand.length; i++) if (cand[i][4] === expect + 1 && cand[i][3]) { nxtBold = i; break; }
        const idx = [];
        for (let i = pos; i < nxtBold; i++) if (cand[i][4] === expect) idx.push(i);
        if (!idx.length) {
          if (skipped || !picked.length) break;
          expect += 1;
          skipped = true;
          continue;
        }
        skipped = false;
        const choose = idx.find((i) => cand[i][3]) ?? idx[0];
        picked.push(cand[choose].slice(0, 3));
        pos = choose + 1;
        expect += 1;
      }
      return picked;
    }

    buildToc(srcName, notes, now) {
      const ws = this.tocWs;
      const header = elems(this.root).find((x) => x.nodeName === "DOCUMENT-HEADER");
      const child = (tag) => {
        const x = header && elems(header).find((y) => y.nodeName === tag);
        return x ? x.textContent : "";
      };
      const set = (addr, v, font) => { const c = ws.getCell(addr); c.value = v; c.font = font; return c; };
      set("A1", `${child("COMPANY-NAME")} ${child("DOCUMENT-NAME")}`.trim(), F_TITLE);
      set("A2", `원본: ${srcName}`, F_NOTE);
      const p = (x) => String(x).padStart(2, "0");
      const stamp = `${now.getFullYear()}-${p(now.getMonth() + 1)}-${p(now.getDate())} ${p(now.getHours())}:${p(now.getMinutes())}`;
      set("A3", `변환: ${stamp}  (숫자는 값으로 변환, 첫 열은 글자 유지)`, F_NOTE);
      if (this.checkSummary) set("A4", { text: `▶ 검증: ${this.checkSummary}`, hyperlink: "'검증'!A1" }, F_LINK);
      let r = 5;
      set(`A${r}`, "시트", F_BOLD);
      set(`B${r}`, "주석 시트는 주석별로 행이 묶여 있습니다. 왼쪽 위 1 / 2 단추로 제목만 보기·모두 펼치기", F_NOTE);
      for (const w of this.writers) {
        r += 1;
        set(`A${r}`, { text: w.ws.name, hyperlink: `'${w.ws.name}'!A1` }, F_LINK);
        for (const [sheet, row, text] of notes) {
          if (sheet === w.ws.name) {
            r += 1;
            set(`B${r}`, { text, hyperlink: `'${sheet}'!A${row}` }, F_LINK);
          }
        }
      }
      ws.getColumn(1).width = 28;
      ws.getColumn(2).width = 60;
      ws.views = [{ showGridLines: false }];
    }

    run(srcName, now) {
      const body = elems(this.root).find((x) => x.nodeName === "BODY");
      this.walk(body || this.root);
      const notes = this.noteIndex();
      for (const w of [...this.writers]) {
        if (!w.used) {
          this.wb.removeWorksheet(w.ws.id);
          this.writers.splice(this.writers.indexOf(w), 1);
        } else {
          w.finish(notes.filter((n) => n[0] === w.ws.name).map((n) => n[1]));
        }
      }
      if (this.checkWs) {
        const tables = [].concat(...this.writers.map((w) => w.tables));
        const paras = [].concat(...this.writers.map((w) => w.paragraphs.map(([r, text]) => ({ sheet: w.ws.name, r, text }))));
        const res = HOST.DSDChecks.run({ tables, notes, paras, ws: this.checkWs, styles: STYLES });
        this.checkSummary = res.summary;
        this.checkCounts = res.counts;
        this.checkFindings = res.findings;
      }
      this.buildToc(srcName, notes, now);
      return notes.length;
    }
  }

  // ------------------------------------------------------------ 파일 처리
  /** contents.xml 바이트를 글자로. XML 선언의 인코딩을 따른다(대부분 utf-8). */
  function decodeXml(bytes) {
    const head = new TextDecoder("latin1").decode(bytes.subarray(0, 200));
    const m = /encoding=["']([^"']+)["']/i.exec(head);
    let enc = m ? m[1].toLowerCase() : "utf-8";
    try { return new TextDecoder(enc).decode(bytes); } catch (e) { return new TextDecoder("utf-8").decode(bytes); }
  }

  /** ExcelJS는 시트 안 이동 링크에도 외부 링크 관계(r:id)를 같이 넣어, 엑셀이 다른 파일을 열려고 할 수 있다.
   *  location이 있는 링크에서 r:id와 그 관계를 지운다. */
  async function fixInternalLinks(buf, JSZip) {
    const zip = await JSZip.loadAsync(buf);
    const sheets = Object.keys(zip.files).filter((n) => /^xl\/worksheets\/sheet\d+\.xml$/.test(n));
    for (const sf of sheets) {
      let xml = await zip.file(sf).async("string");
      const drop = new Set();
      xml = xml.replace(/<hyperlink\b[^>]*\/>/g, (tag) => {
        if (!/\slocation="/.test(tag)) return tag;
        const m = /\sr:id="([^"]+)"/.exec(tag);
        if (m) drop.add(m[1]);
        return tag.replace(/\sr:id="[^"]*"/, "");
      });
      if (!drop.size) continue;
      zip.file(sf, xml);
      const relName = sf.replace("worksheets/", "worksheets/_rels/") + ".rels";
      const rel = zip.file(relName);
      if (rel) {
        const rels = (await rel.async("string")).replace(/<Relationship\b[^>]*\/>/g, (tag) => {
          const m = /\sId="([^"]+)"/.exec(tag);
          return m && drop.has(m[1]) ? "" : tag;
        });
        zip.file(relName, rels);
      }
    }
    return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  }

  async function convert(arrayBuffer, srcName, deps) {
    const { JSZip, ExcelJS, DOMParser } = deps;
    let zip;
    try {
      zip = await JSZip.loadAsync(arrayBuffer);
    } catch (e) {
      throw new Error("DSD(압축) 형식이 아닙니다.");
    }
    const entry = Object.values(zip.files).find((f) => f.name.toLowerCase() === "contents.xml");
    if (!entry) throw new Error("contents.xml이 없습니다. DART 편집기 파일이 맞는지 확인하세요.");
    const text = decodeXml(await entry.async("uint8array"));
    const doc = new DOMParser().parseFromString(text, "application/xml");
    if (doc.getElementsByTagName("parsererror").length) throw new Error("본문(XML)을 읽지 못했습니다.");
    const conv = new Converter(doc.documentElement, ExcelJS);
    const notes = conv.run(srcName, deps.now || new Date());
    const raw = await conv.wb.xlsx.writeBuffer();
    const data = await fixInternalLinks(raw, JSZip);
    return { data, sheets: conv.wb.worksheets.map((w) => w.name), notes, checks: conv.checkCounts || null, findings: conv.checkFindings || [] };
  }

  root.DSDCore = { convert, _internal: { rawText, toValue, sheetNameOf, parseUnit, FS_NAME_RE } };
})(typeof window !== "undefined" ? window : globalThis);
