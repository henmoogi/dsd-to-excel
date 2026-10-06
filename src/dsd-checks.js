/* 합계 검증·정합성 체크 (DSD 엑셀 변환기 확장)
 *
 * DSDCore가 변환하면서 모은 표 구조(tables)를 검사해 '검증' 시트에 결과를 쓰고, 불일치 셀을 연한 빨강으로 칠한다.
 * 검사 종류
 *   세로 합계   : 합계·총계·소계 행 = 위 항목 합 (기초→기말 증감표 포함, 'Ⅲ.합계(Ⅰ+Ⅱ)' 같은 식 표기 포함)
 *   상위 항목   : 'Ⅰ.유동자산' 같은 상위 행 = 바로 아래 하위 항목 합
 *   가로 합계   : 합계 열 = 왼쪽 열들의 합, 기초 열 + 증감 열 = 기말 열
 *   재무제표 간 : 자산총계=부채와자본총계, 당기순이익(손익계산서=현금흐름표=자본변동표=이익잉여금처분계산서),
 *                기말현금=현금및현금성자산, 자본변동표 기말=자본총계, 처분계산서 미처분이익잉여금=재무상태표
 *   본문↔주석   : 재무제표 과목에 적힌 주석 번호의 주석에 같은 금액이 있는지
 * 결과: 일치 / 단수차이(천원 단위 등 반올림 범위) / 불일치 / 확인필요(자동으로 단정하기 어려운 것)
 *       / 미확인(본문 금액을 주석 표·문장에서 찾지 못해 자동 대사를 못 함)
 */
(function (root) {
  "use strict";
  const OK = "일치", ROUND = "단수차이", BAD = "불일치", REVIEW = "확인필요", MISSING = "미확인";
  const RANK = { [BAD]: 0, [REVIEW]: 1, [ROUND]: 2, [MISSING]: 3, [OK]: 4 };
  const BAD_FILL = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFC7CE" } };
  const STATUS_COLOR = { [BAD]: "FFC00000", [REVIEW]: "FFB45F06", [ROUND]: "FF666666", [MISSING]: "FF57606A", [OK]: "FF1A7F37" };
  const EXACT_UNITS = new Set(["원", "주", "명", "시간", "건", "개"]);
  const UNIT_SCALE = { "원": 1, "천원": 1e3, "백만원": 1e6, "억원": 1e8, "십억원": 1e9 };

  // ------------------------------------------------------------ 행 이름 해석
  const MARK_RE = /^\s*(?:[IVX]{1,5}\s*[.．]|[Ⅰ-Ⅻ]\s*[.．]?|\(\d{1,2}\)|\d{1,2}\s*[.．)](?!\d)|[①-⑳]|[가-하]\s*[.．)])\s*/;
  const DETAIL = 9;   // 번호 없는 세부 항목의 깊이
  const ROMAN = { "Ⅰ": "I", "Ⅱ": "II", "Ⅲ": "III", "Ⅳ": "IV", "Ⅴ": "V", "Ⅵ": "VI", "Ⅶ": "VII", "Ⅷ": "VIII", "Ⅸ": "IX", "Ⅹ": "X", "Ⅺ": "XI", "Ⅻ": "XII" };
  const TOTAL_RE = /^(합계|총계|계|소계|합|총합계)$|(합계|총계|소계)$/;
  const SUB_RE = /소계$/;
  const OPEN_RE = /^(기초|당기초|전기초)(잔액|금액|장부금액|순장부금액)?$|\((당기초|전기초|기초)\)$/;
  const CLOSE_RE = /^(기말|당기말|전기말)(잔액|금액|장부금액|순장부금액)?$|\((당기말|전기말|기말)\)$/;
  const HEAD_OPEN_RE = /^기초(잔액|금액|장부금액|순장부금액)?$/;
  const HEAD_CLOSE_RE = /^기말(잔액|금액|장부금액|순장부금액)?$/;
  const HEADER_LABEL_RE = /^(구분|과목|계정과목|항목|내역|내용|과목명)$/;
  const DEDUCT_RE = /(차감|감소|처분|매각|상환|손상|상각|유출|환입|타계정|에누리|환출|반품|기말.*재고|재고.*기말|대손|결손|퇴직연금운용자산|국민연금전환금)/;
  const EXPENSE_RE = /(비용|손실|원가|판매비와관리비)/;
  const CONTRA_RE = /^(대손충당금|감가상각누계액|손상차손누계액|상각누계액|감액손실누계액|현재가치할인차금|정부보조금|국고보조금|사채할인발행차금|사채할증발행차금|전환권조정|신주인수권조정|평가충당금|재고자산평가충당금|퇴직연금운용자산|국민연금전환금)$/;
  const PROFIT_RE = /^(?!영업외)(매출총|영업|법인세비용차감전|법인세차감전|계속영업|중단영업|당기순|반기순|분기순|총포괄|반기총포괄|분기총포괄).*(이익|손실|손익)/;
  const EPS_RE = /주당/;
  const NONADD_RE = /(주식수|주수|비율|지분율|이자율|단가|%|만기|기간|일자|연이율|금리)/;
  const REF_RE = /\((?:주석|주)\s*([\d,\s~\-]+)\)/;

  function markerType(s) {
    s = s.trim();
    if (/^(?:[IVX]{1,5}\s*[.．]|[Ⅰ-Ⅻ])/.test(s)) return "R";
    if (/^\(\d{1,2}\)/.test(s)) return "P";
    if (/^\d{1,2}\s*[.．](?!\d)/.test(s)) return "D";
    if (/^\d{1,2}\)/.test(s)) return "Q";
    if (/^[①-⑳]/.test(s)) return "C";
    if (/^[가-하]\s*[.．)]/.test(s)) return "K";
    return null;
  }
  function markerKey(s) {
    let m = /^\s*([IVX]{1,5})\s*[.．]/.exec(s) || /^\s*([Ⅰ-Ⅻ])/.exec(s);
    if (m) return ROMAN[m[1]] || m[1];
    m = /^\s*\(?(\d{1,2})[.．)]/.exec(s);
    return m ? m[1] : null;
  }
  /** 이름 끝의 '(Ⅰ+Ⅱ)' 같은 식 → [{sign, key}] */
  function formulaOf(label) {
    label = label.replace(/\((?:주석|주)\s*[\d,\s~\-]+\)/g, "").trim();
    const m = /\(([^()]*[+\-－−][^()]*)\)\s*$/.exec(label);
    if (!m) return null;
    const terms = m[1].replace(/\s+/g, "").match(/[+\-－−]?[^+\-－−]+/g) || [];
    const out = terms.map((t) => ({ sign: /^[\-－−]/.test(t) ? -1 : 1, key: t.replace(/^[+\-－−]/, "") }));
    if (!out.length || !out.every((t) => /^([Ⅰ-Ⅻ]|[IVX]{1,5}|\d{1,2})$/.test(t.key))) return null;
    out.forEach((t) => { t.key = ROMAN[t.key] || t.key; });
    return { terms: out, text: m[1] };
  }
  function normLabel(label) {
    return label.replace(MARK_RE, "").replace(MARK_RE, "")
      .replace(/\((?:주석|주)\s*[\d,\s~\-]+\)/g, "")
      .replace(/\(([^()]*[+\-－−][^()]*)\)\s*$/, "")
      .replace(/\(\*\d*\)|\*\d+/g, "")
      .replace(/\s+/g, "");
  }
  function refsOf(text) {
    const m = REF_RE.exec(text || "");
    if (!m) return [];
    const out = [];
    for (const part of m[1].split(",")) {
      const r = /^\s*(\d+)\s*(?:[~\-]\s*(\d+))?\s*$/.exec(part);
      if (!r) continue;
      const a = +r[1], b = r[2] ? +r[2] : a;
      for (let n = a; n <= b && n - a < 20; n++) out.push(n);
    }
    return out;
  }
  const addr = (r, c) => {
    let s = "", n = c;
    while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
    return s + r;
  };

  // ------------------------------------------------------------ 표 해석
  function analyze(t) {
    const rows = t.rows;
    const numCols = new Set();
    for (const row of rows) for (const x of row.cells) if (x.value !== null) numCols.add(x.c);
    if (!numCols.size) return null;
    const firstNum = Math.min(...numCols);
    const firstText = (row) => { const x = row.cells.find((y) => y.c < firstNum && y.text); return x ? normLabel(x.text) : ""; };
    let h = 0;
    while (h < rows.length && (rows[h].head || !rows[h].cells.some((x) => x.value !== null) || HEADER_LABEL_RE.test(firstText(rows[h])))) h++;
    if (h >= rows.length) return null;
    const headRows = rows.slice(0, h), body = rows.slice(h);
    const cols = [...numCols].sort((a, b) => a - b);
    const covers = (x, c) => x.c <= c && c < x.c + x.cs;
    const bottomHead = (c) => {
      for (let i = headRows.length - 1; i >= 0; i--) { const x = headRows[i].cells.find((y) => covers(y, c) && y.text); if (x) return x.text; }
      return "";
    };

    // 기간 묶음(예: '제21(당)기말'이 두 칸을 덮음). 묶음 안에서 한 행에 값이 하나뿐이면 재무제표식 두 칸 배치로 보고 한 줄로 합친다.
    const spans = (x) => x.cs > 1 && cols.filter((c) => covers(x, c)).length > 1 && !/단위/.test(x.text || "");
    const whole = (x) => cols.every((c) => covers(x, c));
    let gIdx = headRows.findIndex((row) => row.cells.some((x) => spans(x) && !whole(x)));
    if (gIdx < 0) gIdx = headRows.findIndex((row) => row.cells.some((x) => spans(x) && x.text));
    const groupOf = new Map();
    if (gIdx >= 0) {
      for (const x of headRows[gIdx].cells) {
        const gc = cols.filter((c) => covers(x, c));
        if (spans(x) && gc.length) { const g = { id: "g" + x.c, label: x.text, cols: gc }; gc.forEach((c) => groupOf.set(c, g)); }
      }
    }
    cols.forEach((c) => { if (!groupOf.has(c)) groupOf.set(c, { id: "c" + c, label: null, cols: [c] }); });
    for (const g of new Set(groupOf.values())) {
      if (g.cols.length < 2) { g.merged = false; continue; }
      const split = headRows.slice(gIdx + 1).some((row) => row.cells.some((x) => x.text && g.cols.some((c) => covers(x, c)) && !g.cols.every((c) => covers(x, c))));
      const multi = body.some((row) => row.cells.filter((x) => x.value !== null && x.value !== 0 && g.cols.includes(x.c)).length > 1);
      g.merged = !split && !multi;
    }
    const series = [];
    const seen = new Set();
    for (const c of cols) {
      const g = groupOf.get(c);
      if (g.merged) {
        if (seen.has(g.id)) continue;
        seen.add(g.id);
        series.push({ cols: g.cols, merged: true, group: g, label: (g.label || "").replace(/\s+/g, " "), head: normLabel(g.label || "") });
      } else {
        const ht = bottomHead(c);
        const label = [g.label, ht].filter(Boolean).filter((v, i, a) => a.indexOf(v) === i).join(" ").replace(/\s+/g, " ");
        series.push({ cols: [c], merged: false, group: g, label, head: normLabel(ht) });
      }
    }
    const usable = series.filter((s) => !/^(주석|주)$/.test(s.head));
    const noteSeries = series.find((s) => /^(주석|주)$/.test(s.head));

    // 행 이름(이름 칸이 위에서 병합되어 내려온 경우도 반영)
    const allCells = [];
    rows.forEach((row) => row.cells.forEach((x) => allCells.push({ r: row.r, x })));
    const info = body.map((row) => {
      const parts = [];
      let own = null, grp = null;
      for (let c = 1; c < firstNum; c++) {
        const x = row.cells.find((y) => y.c === c);
        if (x) {
          if (x.rs > 1) grp = `${row.r}:${c}`;
          if (x.text) { parts.push(x.text); if (!own) own = x.text; }
          continue;
        }
        const up = allCells.find((y) => y.r < row.r && y.r + y.x.rs > row.r && covers(y.x, c));
        if (up && up.x.rs > 1) grp = `${up.r}:${up.x.c}`;
        if (up && up.x.text && !parts.includes(up.x.text)) parts.push(up.x.text);
      }
      const label = parts.join(" ").replace(/\s+/g, " ").trim();
      const lead = own || label;
      const norm = normLabel(label);
      const hasVal = row.cells.some((x) => x.value !== null && usable.some((s) => s.cols.includes(x.c)));
      const noteText = noteSeries ? (row.cells.find((x) => noteSeries.cols.includes(x.c)) || {}).text : "";
      return {
        r: row.r, label, norm, mtype: markerType(lead), key: markerKey(lead), hasVal, grp, contra: CONTRA_RE.test(norm),
        profit: PROFIT_RE.test(norm), eps: EPS_RE.test(norm),
        formula: formulaOf(label), total: TOTAL_RE.test(norm), sub: SUB_RE.test(norm),
        open: OPEN_RE.test(norm), close: CLOSE_RE.test(norm),
        refs: refsOf(label).concat(noteText ? refsOf(`(주석${noteText})`) : []),
      };
    });
    const rank = new Map(info.some((it) => it.mtype === "R") ? [["R", 1]] : []);
    let prevLevel = DETAIL;
    for (const it of info) {
      if (it.mtype && !rank.has(it.mtype)) rank.set(it.mtype, rank.size + 1);
      it.level = it.mtype ? rank.get(it.mtype) : it.contra ? prevLevel : DETAIL;
      if (it.hasVal && !it.contra) prevLevel = it.level;
    }
    let seenOpen = false;
    for (const it of info) {
      if (!it.hasVal) it.kind = it.label ? "S" : "B";
      else if (it.total || it.formula) it.kind = "T";
      else if (it.close && seenOpen) { it.kind = "T"; it.roll = true; }
      else it.kind = "N";
      if (it.open && it.hasVal) seenOpen = true;
    }
    return { t, body, info, series: usable, unit: t.unit };
  }

  // ------------------------------------------------------------ 판정
  function makeJudge(unit) {
    return function judge(stated, cands, prec) {
      cands = cands.filter((c) => c && c.n > 0);
      if (!cands.length) return null;
      const exact = EXACT_UNITS.has(unit) && prec >= 1;
      const eps = 1e-7 * Math.max(1, Math.abs(stated)) + prec * 1e-6;
      let best = null;
      for (const c of cands) {
        const diff = stated - c.sum;
        if (Math.abs(diff) <= eps) return { status: OK, calc: c.sum, diff: 0, basis: c.basis };
        if (!best || Math.abs(diff) < Math.abs(best.diff)) best = { calc: c.sum, diff, basis: c.basis, n: c.n };
      }
      if (!exact && Math.abs(best.diff) <= prec * Math.max(1, Math.ceil(best.n / 2)) + eps) return { status: ROUND, ...best };
      return { status: BAD, ...best };
    };
  }
  const sumOf = (xs) => xs.reduce((a, b) => a + b, 0);

  // ------------------------------------------------------------ 표 안 검사
  function tableChecks(A, out) {
    const { t, info, series } = A;
    const judge = makeJudge(A.unit);
    const sheet = t.sheet;
    for (const s of series) {
      const cell = (i) => {
        let first = null;
        for (const x of A.body[i].cells) if (x.value !== null && s.cols.includes(x.c)) { if (x.value !== 0) return x; if (!first) first = x; }
        return first;
      };
      const push = (i, kind, res, x) => {
        if (!res) return;
        out.push({ status: res.status, kind, sheet, r: info[i].r, c: x.c, item: info[i].label || "(이름 없음)", basisCol: s.label,
                   stated: x.value, calc: res.calc, diff: res.diff, basis: res.basis, ws: t.writer.ws });
      };
      const topLevel = (idx) => {
        const res = [], stack = [];
        for (const j of idx) {
          const lv = info[j].level;
          while (stack.length && stack[stack.length - 1] >= lv) stack.pop();
          const x = cell(j);
          if (!stack.length && x) res.push(j);
          if (x && lv < DETAIL) stack.push(lv);
        }
        return res;
      };
      const cand = (idx, basis, mode) => {
        const xs = idx.map((j) => {
          const x = cell(j);
          if (!x) return null;
          const v = x.value, nm = info[j].norm;
          if (mode === "deduct" && v > 0 && DEDUCT_RE.test(nm)) return -v;
          if (mode === "net" && v > 0 && EXPENSE_RE.test(nm)) return -v;
          return v;
        }).filter((v) => v !== null);
        if (!xs.length) return null;
        const sum = sumOf(xs);
        return { sum: mode === "flip" ? -sum : sum, n: xs.length, basis };
      };
      const variants = (idx, what) => [
        cand(idx, what),
        cand(idx, `${what}(차감 항목은 뺌)`, "deduct"),
        cand(idx, `${what}(수익 − 비용)`, "net"),
      ];
      const precOf = (i) => cell(i).prec;

      // 상위 항목 = 하위 항목 합
      for (let i = 0; i < info.length; i++) {
        const it = info[i];
        if (it.kind !== "N" || it.level >= DETAIL || it.contra || !cell(i) || it.profit || it.eps) continue;
        const kids = [];
        for (let j = i + 1; j < info.length; j++) {
          const jt = info[j];
          if (jt.kind === "T" || jt.kind === "S" || jt.eps) break;
          if (jt.kind === "B") continue;
          if (jt.level <= it.level) break;
          kids.push(j);
        }
        const direct = topLevel(kids);
        if (!direct.length) continue;
        push(i, "상위 항목 합계", judge(cell(i).value, [
          ...variants(direct, `아래 하위 항목 ${direct.length}개 합`),
          cand(direct, `아래 하위 항목 ${direct.length}개 합(상위 행은 부호를 반대로 표시)`, "flip"),
        ], precOf(i)), cell(i));
      }

      // 합계·총계·소계·기말 행
      for (let i = 0; i < info.length; i++) {
        const it = info[i];
        if (it.kind !== "T" || !cell(i)) continue;
        const cands = [];
        if (it.formula) {
          const parts = it.formula.terms.map((tm) => {
            const j = info.findIndex((x, k) => k < i && x.key === tm.key && cell(k));
            return j >= 0 ? tm.sign * cell(j).value : null;
          });
          if (parts.every((p) => p !== null)) cands.push({ sum: sumOf(parts), n: parts.length, basis: `표기된 식 (${it.formula.text})` });
        }
        if (!cands.length && it.roll) {
          let k = i - 1;
          while (k >= 0 && !(info[k].open && info[k].hasVal)) k--;
          const idx = [];
          for (let j = Math.max(k, 0); j < i; j++) if (info[j].kind === "N") idx.push(j);
          cands.push(cand(idx, "기초 + 증감 = 기말"), cand(idx, "기초 + 증가 − 감소 = 기말", "deduct"));
        }
        const formulaOnly = cands.length > 0;
        if (!formulaOnly && !it.roll && it.grp) {
          const same = [];
          for (let j = 0; j < i; j++) if (info[j].grp === it.grp && info[j].kind === "N") same.push(j);
          if (same.length) cands.push(...variants(same, `같은 묶음(${(it.label.split(" ")[0] || "").slice(0, 12)}) ${same.length}개 합`));
        }
        if (!formulaOnly && !it.roll) {
          const starts = new Set();
          let k = i - 1;
          for (; k >= 0; k--) { const kt = info[k]; if (kt.kind === "S" || (kt.kind === "T" && (it.sub || !kt.sub))) break; }
          starts.add(k + 1);
          k = i - 1;
          for (; k >= 0; k--) { const kt = info[k]; if (kt.kind === "T" && (it.sub || !kt.sub)) break; }
          starts.add(k + 1);
          starts.add(0);
          for (const st of starts) {
            const block = [];
            for (let j = st; j < i; j++) block.push(j);
            const nRows = block.filter((j) => info[j].kind === "N");
            const where = st === 0 ? "표 처음부터" : "위";
            const top = topLevel(nRows);
            cands.push(...variants(top, `${where} 항목 ${top.length}개 합`));
            cands.push(cand(nRows, `${where} 모든 행 합`));
            if (!it.sub) {
              const subs = block.filter((j) => info[j].kind === "T" && info[j].sub && cell(j));
              if (subs.length) {
                const lastSub = subs[subs.length - 1];
                const rest = topLevel(nRows.filter((j) => j > lastSub));
                cands.push(cand(subs.concat(rest), `소계 ${subs.length}개` + (rest.length ? ` + 나머지 ${rest.length}개` : "") + " 합"));
              }
            }
          }
          const combo = /^(.+?)(및|와|과)(.+)총계$/.exec(it.norm);
          if (combo) {
            const names = [combo[1] + "총계", combo[3] + "총계"];
            const idx = names.map((nm) => { for (let j = i - 1; j >= 0; j--) if (info[j].norm === nm && cell(j)) return j; return -1; });
            if (idx.every((j) => j >= 0)) cands.push(cand(idx, names.join(" + ")));
          }
        }
        push(i, it.roll ? "증감(기초→기말)" : "세로 합계", judge(cell(i).value, cands, precOf(i)), cell(i));
      }
    }

    // 가로: 합계 열 = 왼쪽 열 합, 기초 열 + 증감 열 = 기말 열
    const single = series.filter((s) => !s.merged);
    for (let k = 0; k < single.length; k++) {
      const s = single[k];
      const isTot = TOTAL_RE.test(s.head);
      const isClose = HEAD_CLOSE_RE.test(s.head);
      if ((!isTot && !isClose) || NONADD_RE.test(s.head)) continue;
      const adds = [];
      let openAt = -1;
      for (let j = k - 1; j >= 0; j--) {
        const p = single[j];
        if (p.group.id !== s.group.id && s.group.cols.length > 1) break;
        if (s.group.cols.length === 1 && p.group.cols.length > 1) break;
        if (isTot && TOTAL_RE.test(p.head)) break;
        if (NONADD_RE.test(p.head)) continue;
        adds.unshift(p);
        if (isClose && HEAD_OPEN_RE.test(p.head)) { openAt = 0; break; }
      }
      if (isClose && openAt < 0) continue;
      if (adds.length < (isClose ? 2 : 2)) continue;
      const judge = makeJudge(A.unit);
      for (let i = 0; i < info.length; i++) {
        if (info[i].kind === "S" || info[i].kind === "B") continue;
        const row = A.body[i];
        const tx = row.cells.find((x) => x.value !== null && x.c === s.cols[0]);
        if (!tx) continue;
        const vals = adds.map((p) => row.cells.find((x) => x.value !== null && x.c === p.cols[0])).map((x, j) => x ? { v: x.value, head: adds[j].head } : null).filter(Boolean);
        if (!vals.length) continue;
        const cands = [{ sum: sumOf(vals.map((v) => v.v)), n: vals.length, basis: isClose ? "기초 + 증감 = 기말 (가로)" : `왼쪽 ${adds.length}개 열 합` }];
        if (isClose) cands.push({ sum: sumOf(vals.map((v) => (v.v > 0 && DEDUCT_RE.test(v.head) ? -v.v : v.v))), n: vals.length, basis: "기초 + 증가 − 감소 = 기말 (가로)" });
        const res = judge(tx.value, cands, tx.prec);
        if (res) out.push({ status: res.status, kind: isClose ? "증감(기초→기말)" : "가로 합계", sheet, r: row.r, c: tx.c,
                            item: info[i].label || "(이름 없음)", basisCol: s.label, stated: tx.value, calc: res.calc, diff: res.diff, basis: res.basis, ws: t.writer.ws });
      }
    }
  }

  // ------------------------------------------------------------ 재무제표 간·본문↔주석
  function fsKind(name) {
    const base = name.replace(/\(\d+\)$/, "");
    if (!root.DSDCore._internal.FS_NAME_RE.test(base)) return null;
    if (/재무상태표|대차대조표/.test(base)) return "BS";
    if (/손익계산서/.test(base)) return "IS";
    if (/현금흐름표/.test(base)) return "CF";
    if (/자본변동표/.test(base)) return "SCE";
    if (/이익잉여금처분계산서|결손금처리계산서/.test(base)) return "RE";
    return null;
  }

  function crossChecks(As, notes, paras, out) {
    const fs = {};
    for (const A of As) {
      const k = fsKind(A.t.sheet);
      if (!k || !A.t.bordered) continue;
      if (!fs[k] || A.body.length > fs[k].body.length) fs[k] = A;
    }
    if (!fs.RE) fs.RE = As.find((A) => A.info.some((x) => /^미처분이익잉여금$/.test(x.norm)) && A.info.some((x) => /차기이월/.test(x.norm))) || null;
    const NAMES = { BS: "재무상태표", IS: "손익계산서", CF: "현금흐름표", SCE: "자본변동표", RE: "이익잉여금처분계산서" };
    const PERIOD = ["당기", "전기"];

    const at = (A, i, p) => {   // p번째 기간 값 칸
      const s = A && A.series[p];
      if (!s || i < 0) return null;
      return A.body[i].cells.find((x) => x.value !== null && s.cols.includes(x.c)) || null;
    };
    const find = (A, re) => (A ? A.info.findIndex((x) => re.test(x.norm) && x.hasVal) : -1);
    const findAll = (A, re) => (A ? A.info.map((x, i) => (re.test(x.norm) && x.hasVal ? i : -1)).filter((i) => i >= 0) : []);
    const lastSeriesCell = (A, i) => {
      const s = A.series[A.series.length - 1];
      return A.body[i].cells.find((x) => x.value !== null && s.cols.includes(x.c)) || null;
    };
    const cmp = (A1, i1, x1, A2, i2, x2, what, period, softNote) => {
      if (!x1 || !x2) return;
      const s1 = UNIT_SCALE[A1.unit] || 1, s2 = UNIT_SCALE[A2.unit] || 1;
      const other = (x2.value * s2) / s1;
      const judge = makeJudge(s1 === s2 ? A1.unit : null);
      const res = judge(x1.value, [{ sum: other, n: 1, basis: `${A2.t.sheet}!${addr(A2.info[i2].r, x2.c)} ${A2.info[i2].label}` + (s1 !== s2 ? ` (${A2.unit}→${A1.unit} 환산)` : "") }],
                        Math.max(Math.min(x1.prec, x2.prec), s2 / s1));
      if (!res) return;
      if (res.status === BAD && softNote) { res.status = REVIEW; res.basis += ` — ${softNote}`; }
      out.push({ status: res.status, kind: "재무제표 간", sheet: A1.t.sheet, r: A1.info[i1].r, c: x1.c, item: A1.info[i1].label,
                 basisCol: period, stated: x1.value, calc: res.calc, diff: res.diff, basis: `${what}: ${res.basis}`, ws: A1.t.writer.ws });
    };

    const { BS, IS, CF, SCE, RE } = fs;
    for (let p = 0; p < 2; p++) {
      const P = PERIOD[p];
      // 1) 자산총계 = 부채와자본총계
      const ia = find(BS, /^자산총계$/), il = find(BS, /^부채(와|및)자본(의)?총계$/);
      cmp(BS, ia, at(BS, ia, p), BS, il, at(BS, il, p), "자산총계 = 부채와자본총계", P);
      // 2) 당기순이익
      const NI = /^당기순(이익|손실|손익)/;
      const ii = find(IS, NI);
      const ic = find(CF, NI);
      cmp(IS, ii, at(IS, ii, p), CF, ic, at(CF, ic, p), "당기순이익(손익계산서 = 현금흐름표)", P);
      if (SCE) {
        const rows = findAll(SCE, NI);
        const j = p === 0 ? rows[rows.length - 1] : rows.length > 1 ? rows[rows.length - 2] : -1;
        if (j !== undefined && j >= 0) cmp(IS, ii, at(IS, ii, p), SCE, j, lastSeriesCell(SCE, j), "당기순이익(손익계산서 = 자본변동표)", P);
      }
      if (RE && RE !== IS) {
        const j = find(RE, NI);
        cmp(IS, ii, at(IS, ii, p), RE, j, at(RE, j, p), "당기순이익(손익계산서 = 이익잉여금처분계산서)", P);
        const ju = find(RE, /^미처분이익잉여금$/), ib = find(BS, /^미처분이익잉여금/);
        cmp(BS, ib, at(BS, ib, p), RE, ju, at(RE, ju, p), "미처분이익잉여금(재무상태표 = 처분계산서)", P);
      }
      // 3) 기말 현금
      const icash = find(BS, /^현금및현금성자산$|^현금및현금등가물$|^현금및예치금$/);
      const iend = find(CF, /기말.*현금|현금.*기말/);
      const soft = BS && icash >= 0 && /예치금/.test(BS.info[icash].norm) ? "현금흐름표의 현금 범위와 과목이 달라 차이가 날 수 있음" : null;
      cmp(BS, icash, at(BS, icash, p), CF, iend, at(CF, iend, p), "기말 현금(재무상태표 = 현금흐름표)", P, soft);
      // 4) 자본변동표 기말 = 자본총계
      if (SCE) {
        const closes = SCE.info.map((x, i) => (x.close && x.hasVal ? i : -1)).filter((i) => i >= 0);
        const j = p === 0 ? closes[closes.length - 1] : closes.length > 1 ? closes[closes.length - 2] : -1;
        const ie = find(BS, /^자본총계$/);
        if (j !== undefined && j >= 0) cmp(BS, ie, at(BS, ie, p), SCE, j, lastSeriesCell(SCE, j), "자본총계(재무상태표 = 자본변동표 기말)", P);
      }
    }

    // 본문 과목의 주석 번호 → 그 주석에 같은 금액이 있는지
    const sections = notes.map(([sheet, row, text]) => ({ sheet, row, n: parseInt(text, 10) })).sort((a, b) => a.row - b.row);
    if (!sections.length) return;
    const noteOf = (A) => {
      if (!/주석/.test(A.t.sheet)) return null;
      let n = null;
      for (const s of sections) if (s.sheet === A.t.sheet && s.row <= A.t.top) n = s.n;
      return n;
    };
    const pool = [];   // {n, v(원 환산), scale, sheet, r, c}
    for (const A of As) {
      const n = noteOf(A);
      if (n === null) continue;
      const scale = UNIT_SCALE[A.unit] || 1;
      for (const row of A.body) for (const x of row.cells) if (x.value !== null && x.value !== 0) pool.push({ n, v: Math.abs(x.value) * scale, scale, sheet: A.t.sheet, r: row.r, c: x.c });
    }
    for (const pr of paras) {
      if (!/주석/.test(pr.sheet)) continue;
      let n = null;
      for (const s of sections) if (s.sheet === pr.sheet && s.row <= pr.r) n = s.n;
      if (n === null) continue;
      const re = /(\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?\s*(천원|백만원|억원|원)?/g;
      let m;
      while ((m = re.exec(pr.text))) {
        if (!m[2] && !m[1].includes(",")) continue;   // 연도·번호 같은 숫자는 뺀다
        const scale = m[2] ? UNIT_SCALE[m[2]] : 1;
        const v = Number(m[1].replace(/,/g, "")) * scale;
        if (v) pool.push({ n, v, scale, sheet: pr.sheet, r: pr.r, c: 1 });
      }
    }
    for (const k of ["BS", "IS", "CF"]) {
      const A = fs[k];
      if (!A) continue;
      const scale = UNIT_SCALE[A.unit] || 1;
      A.info.forEach((it, i) => {
        if (!it.refs.length) return;
        for (let p = 0; p < Math.min(2, A.series.length); p++) {
          const x = at(A, i, p);
          if (!x || !x.value) continue;
          const v = Math.abs(x.value) * scale;
          const hit = (q) => Math.abs(q.v - v) <= Math.max(q.scale, scale) / 2 + 1e-6;
          const inRef = pool.find((q) => it.refs.includes(q.n) && hit(q));
          const elsewhere = inRef ? null : pool.find(hit);
          const status = inRef ? OK : MISSING;   // 다른 주석의 같은 금액은 우연일 수 있어 참고로만 적는다
          const basis = inRef ? `주석 ${inRef.n}에 같은 금액 (${inRef.sheet}!${addr(inRef.r, inRef.c)})`
            : elsewhere ? `참조한 주석 ${it.refs.join(",")}에서는 못 찾음. 주석 ${elsewhere.n}에 같은 금액이 있음(${elsewhere.sheet}!${addr(elsewhere.r, elsewhere.c)}) — 참조 번호 확인`
            : `참조한 주석 ${it.refs.join(",")}의 표·문장에서 같은 금액을 찾지 못함(세부 내역만 있거나 당기만 표시한 경우 등)`;
          out.push({ status, kind: "본문↔주석", sheet: A.t.sheet, r: it.r, c: x.c, item: it.label, basisCol: PERIOD[p],
                     stated: x.value, calc: inRef ? x.value : null, diff: null, basis, ws: A.t.writer.ws });
        }
      });
    }
  }

  // ------------------------------------------------------------ 결과 시트
  function writeSheet(ws, findings, S, order) {
    const set = (r, c, v, font) => { const x = ws.getCell(r, c); x.value = v; if (font) x.font = font; return x; };
    const counts = { [BAD]: 0, [REVIEW]: 0, [ROUND]: 0, [MISSING]: 0, [OK]: 0 };
    findings.forEach((f) => { counts[f.status]++; });
    const summary = `불일치 ${counts[BAD]} · 확인필요 ${counts[REVIEW]} · 단수차이 ${counts[ROUND]} · 미확인 ${counts[MISSING]} · 일치 ${counts[OK]}`;
    set(1, 1, { text: "◀ 목차로 돌아가기", hyperlink: "'목차'!A1" }, S.F_LINK);
    set(3, 1, "합계·정합성 검증", S.F_TITLE);
    set(4, 1, summary, S.F_BOLD);
    set(5, 1, "자동 검사입니다. 서식이 특이한 표는 잘못 판단할 수 있으니 '불일치'·'확인필요'는 원문과 대조하세요. "
      + "'미확인'은 본문 금액을 주석에서 찾지 못해 자동 대사를 못 한 항목입니다. "
      + "위치를 누르면 해당 셀로 갑니다. 불일치 셀은 각 시트에서 연한 빨강으로 칠해 두었습니다.", S.F_NOTE);
    const heads = ["결과", "검사", "위치", "항목", "기준", "기재값", "계산·대사값", "차이", "근거"];
    heads.forEach((h, i) => { const x = set(7, i + 1, h, S.F_BOLD); x.fill = S.HEAD_FILL; x.border = S.BOX; x.alignment = { horizontal: "center", vertical: "middle" }; });
    findings.sort((a, b) => RANK[a.status] - RANK[b.status] || (order.get(a.sheet) || 0) - (order.get(b.sheet) || 0) || a.r - b.r || a.c - b.c);
    const num = (v) => (Number.isInteger(v) ? "#,##0;-#,##0" : "#,##0.0###;-#,##0.0###");
    let r = 8;
    for (const f of findings) {
      const cells = [
        set(r, 1, f.status, { ...S.F_BOLD, color: { argb: STATUS_COLOR[f.status] } }),
        set(r, 2, f.kind, S.F_NORMAL),
        set(r, 3, { text: `${f.sheet}!${addr(f.r, f.c)}`, hyperlink: `'${f.sheet}'!${addr(f.r, f.c)}` }, S.F_LINK),
        set(r, 4, f.item, S.F_NORMAL),
        set(r, 5, f.basisCol || "", S.F_NORMAL),
        set(r, 6, f.stated, S.F_NORMAL),
        set(r, 7, f.calc === null || f.calc === undefined ? null : f.calc, S.F_NORMAL),
        set(r, 8, f.diff === null || f.diff === undefined || f.status === OK ? null : f.diff, S.F_NORMAL),
        set(r, 9, f.basis || "", S.F_NORMAL),
      ];
      [6, 7, 8].forEach((c) => { const v = cells[c - 1].value; if (typeof v === "number") cells[c - 1].numFmt = num(v); });
      cells.forEach((x) => { x.border = S.BOX; x.alignment = { vertical: "top", wrapText: x.col === 4 || x.col === 9 }; });
      if (f.status === BAD && f.ws) f.ws.getCell(f.r, f.c).fill = BAD_FILL;
      r++;
    }
    if (!findings.length) set(8, 1, "검사할 합계·대사 항목을 찾지 못했습니다.", S.F_NORMAL);
    [9, 14, 18, 32, 18, 17, 17, 13, 60].forEach((w, i) => { ws.getColumn(i + 1).width = w; });
    if (findings.length) ws.autoFilter = { from: { row: 7, column: 1 }, to: { row: r - 1, column: 9 } };
    ws.views = [{ state: "frozen", xSplit: 0, ySplit: 7, showGridLines: false }];
    const list = findings.map((f) => ({ status: f.status, kind: f.kind, sheet: f.sheet, addr: addr(f.r, f.c), item: f.item,
                                       basisCol: f.basisCol || "", stated: f.stated, calc: f.calc ?? null,
                                       diff: f.status === OK ? null : f.diff ?? null, basis: f.basis || "" }));
    return { summary, counts, findings: list };   // findings는 화면(HTML) 표시용
  }

  function run({ tables, notes, paras, ws, styles }) {
    const As = tables.map(analyze).filter(Boolean);
    const findings = [];
    for (const A of As) tableChecks(A, findings);
    crossChecks(As, notes, paras || [], findings);
    const order = new Map();
    tables.forEach((t) => { if (!order.has(t.sheet)) order.set(t.sheet, order.size); });
    return writeSheet(ws, findings, styles, order);
  }

  root.DSDChecks = { run, _internal: { analyze, normLabel, markerType, formulaOf, refsOf } };
})(typeof window !== "undefined" ? window : globalThis);
