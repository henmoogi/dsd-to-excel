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
      g.split = split;   // 묶음 아래 하위 머리글이 있는지(전기 대사에서 열을 맞출 때 씀)
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
    return { t, body, info, series: usable, unit: t.unit, headRows, firstNum };
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
      if (res.status === BAD && softNote) { res.status = REVIEW; res.basis += ` - ${softNote}`; }
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
            : elsewhere ? `참조한 주석 ${it.refs.join(",")}에서는 못 찾음. 주석 ${elsewhere.n}에 같은 금액이 있음(${elsewhere.sheet}!${addr(elsewhere.r, elsewhere.c)}) - 참조 번호 확인`
            : `참조한 주석 ${it.refs.join(",")}의 표·문장에서 같은 금액을 찾지 못함(세부 내역만 있거나 당기만 표시한 경우 등)`;
          out.push({ status, kind: "본문↔주석", sheet: A.t.sheet, r: it.r, c: x.c, item: it.label, basisCol: PERIOD[p],
                     stated: x.value, calc: inRef ? x.value : null, diff: null, basis, ws: A.t.writer.ws });
        }
      });
    }
  }

  // ------------------------------------------------------------ 전기 대사(같은 회사 작년 보고서와 비교)
  // 올해 보고서의 '전기' 칸과 작년 보고서의 '당기' 칸을 표마다 맞춰 본다.
  // 표의 기간은 ① 열 머리글(당기·전기·제 N(당)기·기말일) ② 행(자본변동표 '(전기초)'~'(전기말)', 첫 열의 '당기말' 묶음)
  // ③ 표 앞 글('<당 기>', '전기말 현재 …') 순서로 가린다. 표 짝은 같은 구역(재무제표 종류·주석 제목)에서 행 이름이 가장 많이 겹치는 것.
  const CUR = "당기", PRI = "전기";
  const PD = { DIFF: "차이", ROUND: "단수차이", NEW: "작년에 없던 행", GONE: "빠진 행", RENAMED: "이름 바뀐 행",
               NOTFOUND: "표 못 찾음", NOPERIOD: "기간 표시 없음", OK: "일치" };
  const P_RANK = { [PD.DIFF]: 0, [PD.ROUND]: 1, [PD.NEW]: 2, [PD.GONE]: 3, [PD.RENAMED]: 4, [PD.NOTFOUND]: 5, [PD.NOPERIOD]: 6, [PD.OK]: 7 };
  const P_COLOR = { [PD.DIFF]: "FFC00000", [PD.ROUND]: "FF666666", [PD.NEW]: "FFB45F06", [PD.GONE]: "FFB45F06", [PD.RENAMED]: "FF57606A",
                    [PD.NOTFOUND]: "FF57606A", [PD.NOPERIOD]: "FF57606A", [PD.OK]: "FF1A7F37" };
  const NOT_AFTER = "(?!순|손익|법인세|이월|요금|료|공사|설비|장치|전자|오류|분)";   // 당기순이익·전기이월 같은 낱말은 기간 표시가 아니다
  const CUR_TOK = new RegExp(`\\(당\\)|(?<![해상])당(?:반|분)?기${NOT_AFTER}`);
  const PRI_TOK = new RegExp(`\\(전\\)|(?<![충발변송배])전(?:반|분)?기${NOT_AFTER}`);
  const TOK_G = new RegExp(`제\\d+(?:\\((?:당|전)\\))?(?:반|분)?기(?:말|초)?|\\((?:당|전)\\)|(?:당|전)(?:반|분)?기(?:말|초)?${NOT_AFTER}`, "g");
  const DATE_G = /((?:19|20)\d{2})[.\-\/년](\d{1,2})[.\-\/월](\d{1,2})일?\.?/g;
  const YEAR_RE = /((?:19|20)\d{2})년(?:도|말)?(?!이후|이내|부터|까지|~|초과|미만)/;
  const PUNCT_G = /[()（）\[\]<>【】＜＞《》〈〉「」『』.,，·ㆍ:：~\-－–_\/\\'"■□▶▷◆◇○●*※]/g;
  const HEADING_RE = /^\s*(?:\d{1,2}\s*[.．]\s*\D|\(\d{1,2}\)|\d{1,2}\)|[가-하]\s*[.)]|[①-⑳])/;
  const NUMBERING_RE = /^(?:\(\d{1,2}\)|\d{1,2}[.)]|[가-하][.)]|[①-⑳])/;

  function periodCtx(info) {
    const ok = (s) => (s && /^\d{8}$/.test(s) ? s : null);
    const to = ok(info && info.to), from = ok(info && info.from);
    const back = (d) => (d ? String(+d.slice(0, 4) - 1) + d.slice(4) : null);
    return { curEnd: to, curStart: from, priEnd: back(to), priStart: back(from), year: to ? +to.slice(0, 4) : null,
             term: info && info.term ? +info.term : null };
  }
  const ymd = (y, m, d) => y + m.padStart(2, "0") + d.padStart(2, "0");
  /** 기간 경계 날짜면 [기간, '기초'|'기말'] */
  function dateKind(s, ctx) {
    if (s === ctx.curEnd) return [CUR, "기말"];
    if (s === ctx.curStart) return [CUR, "기초"];
    if (s === ctx.priEnd) return [PRI, "기말"];
    if (s === ctx.priStart) return [PRI, "기초"];
    return null;
  }
  /** 머리글의 기간: 당기 / 전기 / null(없거나 둘 다) */
  function periodOf(label, ctx) {
    const t = (label || "").replace(/\s+/g, "");
    if (!t) return null;
    const c = CUR_TOK.test(t), p = PRI_TOK.test(t);
    if (c || p) return c && p ? null : c ? CUR : PRI;
    const tm = /제(\d+)(?:반|분)?기/.exec(t);
    if (tm && ctx.term) { if (+tm[1] === ctx.term) return CUR; if (+tm[1] === ctx.term - 1) return PRI; }
    let per = null;
    for (const m of t.matchAll(DATE_G)) {
      const k = dateKind(ymd(m[1], m[2], m[3]), ctx);
      if (!k || (per && per !== k[0])) return null;
      per = k[0];
    }
    if (per) return per;
    const y = YEAR_RE.exec(t);
    if (y && ctx.year) { if (+y[1] === ctx.year) return CUR; if (+y[1] === ctx.year - 1) return PRI; }
    return null;
  }
  /** 열을 맞출 때 쓰는 이름: 머리글에서 기간 낱말·날짜를 뺀 나머지('당기말 장부금액' → '장부금액') */
  function colKey(path) {
    const k = (path || "").replace(/\s+/g, "").replace(TOK_G, "").replace(DATE_G, "").replace(/(?:19|20)\d{2}년(?:도|말)?/g, "")
      .replace(/현재/g, "").replace(PUNCT_G, "");
    return /^(?:금액|잔액)$/.test(k) ? "" : k;
  }
  /** 행을 맞출 때 쓰는 이름. 각주 표시('(*1,2)')·'(감소)' 같은 괄호 표기를 빼고, 손익 결과 행은 이익·손실 표기 차이를 없앤다
   *  ('당기순이익(손실)' = '당기순손실', '법인세차감전순이익' = '법인세비용차감전순이익', '법인세등' = '법인세비용'). */
  function baseKey(norm, profit) {
    let k = norm.replace(/\(\*[^)]*\)|\*\d+(?:[,~\-]\d+)*/g, "").replace(/\((?:증가|감소|이익|손실|손익|환입|전입|수익|비용)\)/g, "");
    if (profit) k = k.replace(/(이익|손실|손익)(?:\((?:이익|손실|손익)\))?/g, "손익").replace(/법인세비용차감전/g, "법인세차감전");
    return k.replace(/^법인세(?:비용|등)$/, "법인세비용").replace(PUNCT_G, "");
  }
  const OLDNAME_RE = /[(（]\s*(?:구|舊)\s*[,，.:]?\s*([^)）]+)[)）]/;
  /** 행 이름(맞출 때 쓰는 것). 첫 열 쪽 글자 칸 중 머리글에 기간이 있는 칸(예: 당기말 외화금액)은 뺀다. */
  function keyLabels(A, ctx) {
    const skip = new Set();
    for (let c = 1; c < A.firstNum; c++) {
      const parts = [];
      for (const row of A.headRows) for (const x of row.cells) if (x.text && x.c <= c && c < x.c + x.cs && !/단위/.test(x.text)) parts.push(x.text);
      if (periodOf(parts.join(" "), ctx)) skip.add(c);
    }
    if (!skip.size) return A.info.map((it) => it.label);
    const allCells = [];
    A.t.rows.forEach((row) => row.cells.forEach((x) => allCells.push({ r: row.r, x })));
    return A.body.map((row) => {
      const parts = [];
      for (let c = 1; c < A.firstNum; c++) {
        if (skip.has(c)) continue;
        const x = row.cells.find((y) => y.c === c);
        if (x) { if (x.text) parts.push(x.text); continue; }
        const up = allCells.find((y) => y.r < row.r && y.r + y.x.rs > row.r && y.x.c <= c && c < y.x.c + y.x.cs);
        if (up && up.x.text && !parts.includes(up.x.text)) parts.push(up.x.text);
      }
      return parts.join(" ").replace(/\s+/g, " ").trim();
    });
  }
  /** 같은 이름이 여러 번 나오면 순서 번호를 붙인다(대손충당금은 바로 위 과목 이름을 앞에 붙여 구분). 'X(구, Y)'는 Y로도 맞춘다(alias). */
  function plainKeys(A, ctx) {
    let last = "";
    const labels = keyLabels(A, ctx);
    return A.info.map((it, i) => {
      const nm = normLabel(labels[i].replace(/\(\s*\*[^)]*\)/g, ""));   // 각주 표시 '(*1,2)'는 먼저 뺀다
      const old = OLDNAME_RE.exec(labels[i]);
      let k = baseKey(nm.replace(/[(（](?:구|舊)[,，.:]?[^)）]+[)）]/g, ""), PROFIT_RE.test(nm));
      let alias = old ? baseKey(normLabel(old[1]), false) : null;
      if (it.contra && last) { k = last + "/" + k; alias = null; }
      else if (it.hasVal) last = k;
      return { key: k, alias };
    });
  }
  function keyed(list) {
    const seen = new Map();
    return list.map((k) => { const n = seen.get(k) || 0; seen.set(k, n + 1); return n ? `${k}#${n}` : k; });
  }

  /** '<당 기>', '(전기)', '제 20 기', '2024년'처럼 기간만 적은 글 → 기간. 기간 글이 아니면 undefined. */
  function markPeriod(s, ctx) {
    const t = s.replace(/[(（]?\s*단\s*위\s*[:：]?[^)）]*[)）]?/g, "").replace(/\s+/g, "").replace(NUMBERING_RE, "");
    const bare = t.replace(PUNCT_G, "");
    if (!bare) return undefined;
    let m = /^(당|전)(?:반|분)?기(?:말|초)?(?:현재|중)?$/.exec(bare);
    if (m) return m[1] === "당" ? CUR : PRI;
    m = /^제(\d+)(당|전)?(?:반|분)?기(?:말|초)?(?:현재)?$/.exec(bare);
    if (m) {
      if (m[2]) return m[2] === "당" ? CUR : PRI;
      if (ctx.term) return +m[1] === ctx.term ? CUR : +m[1] === ctx.term - 1 ? PRI : null;
      return undefined;
    }
    let per = null, dates = 0;
    const rest = t.replace(DATE_G, (x, y, mo, d) => { dates++; const k = dateKind(ymd(y, mo, d), ctx); per = k && (!per || per === k[0]) ? k[0] : "x"; return ""; })
      .replace(/현재|부터|까지|~/g, "").replace(PUNCT_G, "");
    if (dates && !rest) return per === "x" ? null : per;
    m = /^((?:19|20)\d{2})년(?:도|말)?(?:현재)?$/.exec(bare);
    if (m && ctx.year) return +m[1] === ctx.year ? CUR : +m[1] === ctx.year - 1 ? PRI : null;
    return undefined;
  }
  /** 문장 속 기간 낱말('당기말 현재 …', '… 당기 중 …') → 기간. 당기·전기를 함께 말하면 null, 해당 없으면 undefined.
   *  낱말로 떨어진 것만 센다(당기순이익·전기통신·해당기간·충전기는 아님). */
  const WORD_END = "(?=[\\s,.)\\]]|$|중|에|의|와|과|및|부터|까지|동안|현재)";
  const W_CUR = new RegExp(`(?<![가-힣])당\\s?(?:반|분)?기(?:\\s?[말초])?${WORD_END}`);
  const W_PRI = new RegExp(`(?<![가-힣])전\\s?(?:반|분)?기(?:\\s?[말초])?${WORD_END}`);
  function sentencePeriod(s) {
    const t = s.replace(/\s+/g, " ").trim();
    if (/^보고기간(?:종료일|말)현재/.test(t.replace(/\s+/g, "").replace(NUMBERING_RE, ""))) return CUR;
    const c = W_CUR.test(t), p = W_PRI.test(t);
    if (!c && !p) return undefined;
    return c && p ? null : c ? CUR : PRI;
  }
  /** 표 앞 글(가까운 것부터)에서 표 전체의 기간을 찾는다. 소제목·주석 제목을 만나면 멈춘다. */
  function ctxPeriod(texts, ctx) {
    for (let k = texts.length - 1; k >= 0; k--) {
      const s = texts[k] || "";
      const a = markPeriod(s, ctx);
      if (a !== undefined) return a;
      const b = sentencePeriod(s);
      if (b !== undefined) return b;
      if (HEADING_RE.test(s)) return null;
    }
    return null;
  }
  /** 행의 기간: '2024.01.01(전기초)', '(전기말)', '<당기>' 같은 표시 행(marker), 또는 '당기말 현금…'처럼 기간으로 시작하는 행 */
  function rowPeriod(label, ctx) {
    const raw = label || "";
    let dk = null, bad = false;
    const t = raw.replace(/\s+/g, "").replace(DATE_G, (x, y, mo, d) => { const k = dateKind(ymd(y, mo, d), ctx); if (k) dk = k; else bad = true; return ""; });
    const bare = t.replace(PUNCT_G, "");
    let m = /^(당|전)(?:반|분)?기(초|말)?(?:잔액|금액|현재)?$/.exec(bare);
    if (m) return { period: m[1] === "당" ? CUR : PRI, key: m[2] ? "기" + m[2] : dk ? dk[1] : "", marker: true };
    if (dk && !bad && (!bare || /^(?:기초|기말)(?:잔액|금액)?$/.test(bare))) return { period: dk[0], key: dk[1], marker: true };
    // '2024.01.01.(당기초) (감사받지 아니한 재무제표)'처럼 괄호 속 기간 표시 + 덧붙인 설명
    m = /[(（](당|전)(?:반|분)?기(초|말)?[)）]/.exec(t);
    if (m && !bad && !t.replace(/[(（][^)）]*[)）]/g, "").replace(PUNCT_G, "")) {
      return { period: m[1] === "당" ? CUR : PRI, key: m[2] ? "기" + m[2] : dk ? dk[1] : "", marker: true };
    }
    m = /^((?:19|20)\d{2})년(?:도|말)?$/.exec(bare);
    if (m && ctx.year && !dk && !bad) {
      const y = +m[1];
      if (y === ctx.year || y === ctx.year - 1) return { period: y === ctx.year ? CUR : PRI, key: "", marker: true };
    }
    m = /^[\[(<【]?\s*(당|전)\s*(?:반|분)?\s*기\s*(?:초|말)?\s*[\])>】]?\s+(\S.*)$/.exec(raw);
    if (m && !new RegExp(`^[\\[(<【]?(?:당|전)(?:반|분)?기(?:초|말)?[\\])>】]?(?:순|손익|법인세|이월|요금|료|공사|설비|장치|전자|오류|분)`).test(raw.replace(/\s+/g, ""))) {
      return { period: m[1] === "당" ? CUR : PRI, rest: m[2], marker: false };
    }
    return null;
  }
  /** 행마다 {기간, 맞출 이름, 표시 행 여부}. 표시 행 사이의 행은 앞 표시 행(기초·기간)의 기간을, '…말' 표시 행 뒤에서는 다음 '…말' 표시 행의 기간을 따른다. */
  function rowInfo(A, ctx, plain) {
    const rp = A.info.map((it) => rowPeriod(it.label, ctx));
    const out = A.info.map((it, i) => ({ period: null, key: plain[i].key, alias: plain[i].alias, marked: false }));
    let markers = 0, leads = 0;
    rp.forEach((r, i) => {
      if (!r) return;
      out[i].period = r.period;
      if (r.marker) { markers++; out[i].marked = true; out[i].key = r.key; out[i].alias = null; }
      else { leads++; const nm = normLabel(r.rest); out[i].key = baseKey(nm, PROFIT_RE.test(nm)); out[i].alias = null; }
    });
    const isClose = (k) => rp[k].key === "기말";
    for (let i = 0; i < rp.length; i++) {
      if (rp[i]) continue;
      let prev = -1, next = -1;
      for (let k = i - 1; k >= 0; k--) if (rp[k] && rp[k].marker) { prev = k; break; }
      for (let k = i + 1; k < rp.length; k++) if (rp[k] && rp[k].marker) { next = k; break; }
      if (prev >= 0 && !isClose(prev)) out[i].period = rp[prev].period;
      else if (next >= 0 && isClose(next)) out[i].period = rp[next].period;
    }
    return { rows: out, rowType: markers >= 1 || leads >= 2 };
  }

  /** 기간을 가릴 열 단위. 재무제표처럼 기간 머리글 하나가 두 칸(세부·합계)을 덮고 하위 머리글이 없으면 두 칸을 한 단위로 본다. */
  function unitsOf(A, ctx) {
    const out = [], seen = new Set();
    const mk = (cols, list) => {
      const parts = [];
      for (const row of A.headRows) for (const x of row.cells) {
        if (!x.text || /단위/.test(x.text)) continue;
        if (cols.every((c) => x.c <= c && c < x.c + x.cs) && !parts.includes(x.text)) parts.push(x.text);
      }
      const label = parts.join(" ").replace(/\s+/g, " ").trim();
      return { cols, list, label, period: periodOf(label, ctx), key: colKey(label) };
    };
    for (const s of A.series) {
      const g = s.group;
      if (g.cols.length > 1 && !g.split && periodOf(g.label, ctx)) {
        if (seen.has(g.id)) continue;
        seen.add(g.id);
        out.push(mk(g.cols.filter((c) => A.series.some((x) => x.cols.includes(c))), true));
      } else out.push(mk(s.cols, false));
    }
    return out;
  }
  const cellsOf = (A, i, u) => A.body[i].cells.filter((x) => x.value !== null && u.cols.includes(x.c)).sort((a, b) => a.c - b.c);
  const repCell = (xs) => xs.find((x) => x.value !== 0) || xs[0] || null;

  /** 보고서 하나의 표를 기간 조각(part)으로 나눈다. → {parts, noPeriod} */
  function partsOf(snap, As) {
    const ctx = periodCtx(snap.info);
    const sections = (snap.notes || []).map(([sheet, row, text]) => ({ sheet, row, text })).sort((a, b) => a.row - b.row);
    const parts = [], noPeriod = [];
    As.forEach((A) => {
      if (!A) return;
      const t = A.t;
      let note = null;
      if (/주석/.test(t.sheet)) for (const s of sections) if (s.sheet === t.sheet && s.row <= t.top) note = s.text;
      const fk = fsKind(t.sheet);
      const section = fk ? "FS:" + fk : note ? "N:" + note.replace(/^\s*\d{1,2}\s*[.．]\s*/, "").replace(/\s+/g, "").replace(PUNCT_G, "")
        : "S:" + t.sheet.replace(/\(\d+\)$/, "");
      const name = cps(fk ? t.sheet : note ? `주석 ${note}` : t.sheet.replace(/\(\d+\)$/, "")).slice(0, 30).join("");
      const units = unitsOf(A, ctx);
      const plain = plainKeys(A, ctx);
      const valRows = (us) => A.info.map((it, i) => i).filter((i) => A.info[i].hasVal && us.some((u) => cellsOf(A, i, u).length));
      const plainRow = (i) => ({ i, key: plain[i].key, alias: plain[i].alias, marked: false });
      const add = (period, us, rows, markerOnly) => {
        if (!us.length || !rows.length) return;
        const keys = keyed(rows.map((r) => r.key));
        rows.forEach((r, k) => { r.key = keys[k]; });
        parts.push({ A, t, section, name, period, units: us, ukeys: keyed(us.map((u) => u.key)), rows, keys, markerOnly: !!markerOnly });
      };
      const kinds = new Set(units.map((u) => u.period).filter(Boolean));
      if (kinds.size === 2) {   // 열에 당기·전기가 함께 있는 표
        for (const P of [CUR, PRI]) { const us = units.filter((u) => u.period === P); add(P, us, valRows(us).map(plainRow)); }
        return;
      }
      const ri = rowInfo(A, ctx, plain);
      if (kinds.size === 0 && ri.rowType) {   // 행으로 기간을 나눈 표(자본변동표 등)
        for (const P of [CUR, PRI]) {
          const rows = valRows(units).filter((i) => ri.rows[i].period === P)
            .map((i) => ({ i, key: ri.rows[i].key, alias: ri.rows[i].alias, marked: ri.rows[i].marked }));
          add(P, units, rows, rows.length && rows.every((r) => r.marked));
        }
        return;
      }
      let P = kinds.size === 1 ? [...kinds][0] : null;
      if (!P) {
        const own = [];
        const lastCol = Math.max(...t.rows.map((rw) => Math.max(0, ...rw.cells.map((x) => x.c + x.cs - 1))));
        for (const row of A.headRows) for (const x of row.cells) if (x.text && (x.c < A.firstNum || (x.c === 1 && x.c + x.cs - 1 >= lastCol))) own.push(x.text);
        P = ctxPeriod((t.pre || []).concat(own), ctx);
      }
      if (!P && /외부감사실시내용/.test(t.sheet)) P = CUR;   // 올해 감사 내용
      if (P) { add(P, units, valRows(units).map(plainRow)); return; }
      if ((fk || /주석/.test(t.sheet)) && A.info.some((it) => it.hasVal)) noPeriod.push({ A, t, section, name });
    });
    return { parts, noPeriod };
  }

  function priorCompare(self, curAs, prior) {
    const C = partsOf(self, curAs), Pp = partsOf(prior, prior.tables.map(analyze));
    const As = C.parts.filter((p) => p.period === PRI), Bs = Pp.parts.filter((p) => p.period === CUR);
    const overlap = (x, y) => { const s = new Set(y); return x.filter((k) => s.has(k)).length; };
    const unitScale = (A) => UNIT_SCALE[A.unit] || 1;
    const rowVals = (p, us, i) => [].concat(...us.map((u) => cellsOf(p.A, i, u)));
    const usedA = new Set(), usedB = new Set(), pairs = [], best = new Map();
    const assign = (cand) => {
      cand.sort((x, y) => y.score - x.score);
      for (const c of cand) if (!usedA.has(c.ai) && !usedB.has(c.bi)) { usedA.add(c.ai); usedB.add(c.bi); pairs.push(c); }
    };
    // 1) 행 이름으로 짝: 같은 구역에서 행 이름이 절반 이상 겹침(다른 구역이면 80% 이상·3행 이상). 재무제표끼리는 같은 종류만.
    const cand = [];
    As.forEach((a, ai) => Bs.forEach((b, bi) => {
      const fa = a.section.startsWith("FS:"), fb = b.section.startsWith("FS:");
      if (fa && fb && a.section !== b.section) return;
      const akeys = a.keys.map((k, n) => (b.keys.includes(k) ? k : a.rows[n].alias || k));
      const inter = overlap(akeys, b.keys);
      const sim = (2 * inter) / (a.keys.length + b.keys.length);
      const colHit = overlap(a.ukeys, b.ukeys);
      const same = a.section === b.section;
      const ok = (colHit > 0 || a.units.length === b.units.length) && sim >= 0.5 && (same || (sim >= 0.8 && inter >= 3));
      if (ok) cand.push({ ai, bi, score: sim + (same ? 0.5 : 0) + (0.2 * colHit) / Math.max(a.units.length, 1), how: "name" });
      else if (!best.has(ai) || best.get(ai).sim < sim) best.set(ai, { sim, b });
    }));
    assign(cand);
    // 2) 행 이름이 거의 다 바뀐 표(거래 상대방 이름 변경 등)는 같은 구역 안에서 금액으로 짝: 0이 아닌 금액 2개 이상, 80% 이상 같음
    const amounts = (p) => {
      const sc = unitScale(p.A);
      return [].concat(...p.rows.map((r) => rowVals(p, p.units, r.i).filter((x) => x.value !== 0 && x.prec >= 1).map((x) => Math.round(x.value * sc))));
    };
    const cand2 = [];
    As.forEach((a, ai) => {
      if (usedA.has(ai)) return;
      const va = amounts(a);
      if (va.length < 2) return;
      Bs.forEach((b, bi) => {
        if (usedB.has(bi) || a.section !== b.section) return;
        const left = new Map();
        for (const v of amounts(b)) left.set(v, (left.get(v) || 0) + 1);
        let common = 0;
        for (const v of va) if (left.get(v)) { common++; left.set(v, left.get(v) - 1); }
        const ratio = common / va.length;
        if (common >= 2 && ratio >= 0.8) cand2.push({ ai, bi, score: ratio, how: "amount" });
      });
    });
    assign(cand2);

    const items = [];
    let cells = 0, zeroNew = 0, zeroGone = 0;
    const loc = (p, i) => p.A.info[i].r;
    const labelOf = (p, r) => p.A.info[r.i].label || "(이름 없음)";
    const unitLabel = (p, x) => (x ? (p.units.find((u) => u.cols.includes(x.c)) || {}).label || "" : "");
    for (const { ai, bi, how } of pairs) {
      const a = As[ai], b = Bs[bi];
      const bBy = new Map(b.units.map((u, k) => [b.ukeys[k], u]));
      let umap = a.units.map((u, k) => [u, bBy.get(a.ukeys[k])]).filter(([, y]) => y);
      if (!umap.length && a.units.length === b.units.length) umap = a.units.map((u, k) => [u, b.units[k]]);
      const bRows = a.markerOnly ? b.rows.filter((r) => r.marked) : b.rows;
      const byKey = new Map(bRows.map((r) => [r.key, r]));
      const s1 = unitScale(a.A), s2 = unitScale(b.A);
      const used = new Set(), matched = [];
      let leftA = [];
      for (const ra of a.rows) {
        let rb = byKey.get(ra.key);
        if ((!rb || used.has(rb.key)) && ra.alias) rb = byKey.get(ra.alias);
        if (rb && !used.has(rb.key)) { used.add(rb.key); matched.push([ra, rb, false]); } else leftA.push(ra);
      }
      let leftB = bRows.filter((r) => !used.has(r.key));
      // 이름만 바뀐 행: 남은 행끼리 0이 아닌 금액이 모두 같고 서로 하나뿐이면 잇는다
      if (s1 === s2 && leftA.length && leftB.length && umap.length) {
        const sig = (p, r, side) => {
          const v = [].concat(...umap.map(([ua, ub]) => cellsOf(p.A, r.i, side ? ub : ua))).map((x) => x.value).filter((x) => x !== 0);
          return v.length ? v.map((x) => Math.round(x * 1e6) / 1e6).join("|") : null;
        };
        const group = (list, p, side) => { const m = new Map(); for (const r of list) { const s = sig(p, r, side); if (s) m.set(s, (m.get(s) || []).concat([r])); } return m; };
        const ga = group(leftA, a, 0), gb = group(leftB, b, 1);
        const doneA = new Set(), doneB = new Set();
        for (const [s, ras] of ga) {
          const rbs = gb.get(s);
          if (ras.length === 1 && rbs && rbs.length === 1) { matched.push([ras[0], rbs[0], true]); doneA.add(ras[0]); doneB.add(rbs[0]); }
        }
        leftA = leftA.filter((r) => !doneA.has(r));
        leftB = leftB.filter((r) => !doneB.has(r));
      }
      for (const [ra, rb, renamed] of matched) {
        const label = labelOf(a, ra);
        if (renamed) {
          const x = repCell(rowVals(a, a.units, ra.i).filter((y) => y.value !== 0));
          items.push({ status: PD.RENAMED, table: a.name, sheet: a.t.sheet, r: loc(a, ra.i), c: x ? x.c : a.units[0].cols[0], item: label, basisCol: unitLabel(a, x),
                       stated: x ? x.value : null, calc: null, diff: null, mine: true,
                       basis: `작년 이름: ${labelOf(b, rb)} (작년 ${b.t.sheet}!A${loc(b, rb.i)}). 금액은 같음` + (how === "amount" ? ". 표는 금액으로 짝지음" : "") });
        }
        for (const [ua, ub] of umap) {
          const xa = cellsOf(a.A, ra.i, ua), xb = cellsOf(b.A, rb.i, ub);
          if (!xa.length && !xb.length) continue;
          const prs = xa.length === xb.length ? xa.map((x, k) => [x, xb[k]]) : [[repCell(xa), repCell(xb)]];
          for (const [ca, cb] of prs) {
            const scaleOk = (!ca || ca.prec >= 1) && (!cb || cb.prec >= 1);
            const k1 = scaleOk ? s1 : 1, k2 = scaleOk ? s2 : 1;
            const va = ca ? ca.value : 0, vb = ((cb ? cb.value : 0) * k2) / k1;
            const prec = Math.min(ca ? ca.prec : 1, cb ? cb.prec : 1);
            const eps = 1e-7 * Math.max(1, Math.abs(va)) + prec * 1e-6;
            const d = va - vb;
            const status = Math.abs(d) <= eps ? PD.OK : k1 !== k2 && Math.abs(d) <= (0.5 * Math.max(k1, k2)) / k1 + eps ? PD.ROUND : PD.DIFF;
            cells++;
            const bc = cb ? cb.c : ub.cols[0];
            items.push({ status, table: a.name, sheet: a.t.sheet, r: loc(a, ra.i), c: ca ? ca.c : ua.cols[0], item: label, basisCol: ua.label,
                         stated: ca ? ca.value : null, calc: cb ? vb : null, diff: status === PD.OK ? null : d, mine: true, ws: a.t.writer && a.t.writer.ws,
                         basis: `작년 ${b.t.sheet}!${addr(loc(b, rb.i), bc)}` + (ub.label ? ` (${ub.label})` : "") + (k1 !== k2 ? ` (${b.A.unit}→${a.A.unit} 환산)` : "")
                           + (renamed ? ` (작년 이름: ${labelOf(b, rb)})` : "") });
          }
        }
      }
      for (const ra of leftA) {   // 작년에 없던 행(금액이 모두 0이면 개수만 센다)
        const xs = rowVals(a, a.units, ra.i);
        if (xs.every((y) => y.value === 0)) { zeroNew++; continue; }
        const x = repCell(xs);
        items.push({ status: PD.NEW, table: a.name, sheet: a.t.sheet, r: loc(a, ra.i), c: x.c, item: labelOf(a, ra), basisCol: unitLabel(a, x),
                     stated: x.value, calc: null, diff: null, mine: true, basis: `작년 같은 표(${b.t.sheet}!A${b.t.top})에 이 행이 없음` });
      }
      for (const rb of leftB) {   // 올해 전기 칸에서 빠진 행
        const xs = rowVals(b, b.units, rb.i);
        if (xs.every((y) => y.value === 0)) { zeroGone++; continue; }
        const x = repCell(xs);
        items.push({ status: PD.GONE, table: a.name, sheet: b.t.sheet, r: loc(b, rb.i), c: x.c, item: labelOf(b, rb), basisCol: unitLabel(b, x),
                     stated: null, calc: x.value, diff: null, mine: false, basis: `작년 당기 칸에는 있으나 올해 전기 칸(${a.t.sheet}!A${a.t.top} 표)에 없음` });
      }
    }
    const missed = new Set();
    As.forEach((a, ai) => {
      if (usedA.has(ai) || missed.has(a.t)) return;
      missed.add(a.t);
      const hint = best.get(ai);
      items.push({ status: PD.NOTFOUND, table: a.name, sheet: a.t.sheet, r: a.t.top, c: 1, item: `첫 행: ${labelOf(a, a.rows[0])}`, basisCol: "",
                   stated: null, calc: null, diff: null, mine: true,
                   basis: `작년 보고서에서 같은 표를 찾지 못함(전기 행 ${a.rows.length}개)` + (hint && hint.sim > 0 ? `. 가장 비슷한 표: 작년 ${hint.b.t.sheet}!A${hint.b.t.top}(행 이름 ${Math.round(hint.sim * 100)}% 일치)` : "") });
    });
    for (const np of C.noPeriod) {
      const first = np.A.info.find((it) => it.hasVal);
      items.push({ status: PD.NOPERIOD, table: np.name, sheet: np.t.sheet, r: np.t.top, c: 1, item: `첫 행: ${(first && first.label) || "(이름 없음)"}`, basisCol: "",
                   stated: null, calc: null, diff: null, mine: true, basis: "표 머리글·앞 글에서 당기·전기를 가리지 못해 비교하지 않음" });
    }
    const counts = {};
    Object.values(PD).forEach((s) => { counts[s] = 0; });
    items.forEach((x) => { counts[x.status]++; });
    counts.tables = new Set(pairs.map((p) => As[p.ai].t)).size;
    counts.cells = cells;
    counts.zeroRows = zeroNew + zeroGone;
    const sentence = `비교한 표 ${counts.tables}개, 숫자 ${cells}칸, 차이 ${counts[PD.DIFF]}건` + (counts[PD.ROUND] ? `(단수차이 ${counts[PD.ROUND]}건 별도)` : "")
      + `. 비교하지 못한 표: 같은 표를 못 찾음 ${counts[PD.NOTFOUND]}개, 기간 표시 없음 ${counts[PD.NOPERIOD]}개.`;
    const rowsLine = `작년에 없던 행 ${counts[PD.NEW]}개, 올해 전기 칸에서 빠진 행 ${counts[PD.GONE]}개, 이름만 바뀐 행 ${counts[PD.RENAMED]}개(금액 같음).`
      + (zeroNew + zeroGone ? ` 금액이 모두 0인 행 ${zeroNew + zeroGone}개(작년에 없던 행 ${zeroNew}, 빠진 행 ${zeroGone})는 목록에서 뺌.` : "");
    return { priorName: prior.name, priorCompany: (prior.info && prior.info.company) || "", counts, sentence, rowsLine, items };
  }
  const cps = (s) => Array.from(s);
  const PRIOR_NOTE = "재분류·재작성으로 일부러 바꾼 숫자도 차이로 나옵니다. 차이는 전기 재작성·재분류 여부와 함께 원문을 확인하세요. "
    + "표는 재무제표 종류·주석 제목과 행 이름으로 짝지었고(주석 번호가 바뀌어도 찾음), 행 이름이 거의 다 바뀐 표는 같은 주석 안에서 금액으로 짝지었습니다.";

  /** 검증 시트 아래쪽에 전기 대사 구역을 쓰고, 차이 칸을 칠한다. → 화면 표시용 결과 */
  function writePrior(ws, P, S, order, start) {
    const set = (r, c, v, font) => { const x = ws.getCell(r, c); x.value = v; if (font) x.font = font; return x; };
    set(6, 1, { text: `▶ 전기 대사: ${P.sentence} (아래 ${start}행)`, hyperlink: `'검증'!A${start}` }, S.F_LINK);
    set(start, 1, "전기 대사", S.F_TITLE);
    set(start + 1, 1, `작년 보고서: ${P.priorName}` + (P.priorCompany ? ` (${P.priorCompany})` : "") + ". 올해 보고서의 전기 칸과 작년 보고서의 당기 칸을 표마다 비교했습니다.", S.F_NOTE);
    set(start + 2, 1, P.sentence, S.F_BOLD);
    set(start + 3, 1, P.rowsLine, S.F_BOLD);
    set(start + 4, 1, PRIOR_NOTE, S.F_NOTE);
    const h = start + 6;
    ["결과", "표", "위치", "항목", "기준", "올해 전기", "작년 당기", "차이", "근거"].forEach((t, i) => {
      const x = set(h, i + 1, t, S.F_BOLD); x.fill = S.HEAD_FILL; x.border = S.BOX; x.alignment = { horizontal: "center", vertical: "middle" };
    });
    const sheetOrder = (f) => (order.get(f.sheet) || 0) + (f.mine ? 0 : 0.5);
    const items = P.items.slice().sort((a, b) => P_RANK[a.status] - P_RANK[b.status] || sheetOrder(a) - sheetOrder(b) || a.r - b.r || a.c - b.c);
    const num = (v) => (Number.isInteger(v) ? "#,##0;-#,##0" : "#,##0.0###;-#,##0.0###");
    let r = h + 1;
    for (const f of items) {
      const where = `${f.sheet}!${addr(f.r, f.c)}`;
      const cells = [
        set(r, 1, f.status, { ...S.F_BOLD, color: { argb: P_COLOR[f.status] } }),
        set(r, 2, f.table, S.F_NORMAL),
        set(r, 3, f.mine ? { text: where, hyperlink: `'${f.sheet}'!${addr(f.r, f.c)}` } : `(작년) ${where}`, f.mine ? S.F_LINK : S.F_NORMAL),
        set(r, 4, f.item, S.F_NORMAL),
        set(r, 5, f.basisCol || "", S.F_NORMAL),
        set(r, 6, f.stated ?? null, S.F_NORMAL),
        set(r, 7, f.calc ?? null, S.F_NORMAL),
        set(r, 8, f.diff ?? null, S.F_NORMAL),
        set(r, 9, f.basis || "", S.F_NORMAL),
      ];
      [6, 7, 8].forEach((c) => { const v = cells[c - 1].value; if (typeof v === "number") cells[c - 1].numFmt = num(v); });
      cells.forEach((x) => { x.border = S.BOX; x.alignment = { vertical: "top", wrapText: [1, 2, 4, 9].includes(x.col) }; });
      if (cps(f.status).length > 4) ws.getRow(r).height = 30;
      if (f.status === PD.DIFF && f.ws) f.ws.getCell(f.r, f.c).fill = BAD_FILL;
      r++;
    }
    if (!items.length) set(h + 1, 1, "짝지은 표가 없습니다.", S.F_NORMAL);
    return { priorName: P.priorName, sentence: P.sentence, rowsLine: P.rowsLine, note: PRIOR_NOTE, counts: P.counts,
             items: items.map((f) => ({ status: f.status, table: f.table, sheet: f.sheet, addr: addr(f.r, f.c), mine: f.mine, item: f.item,
                                       basisCol: f.basisCol || "", stated: f.stated ?? null, calc: f.calc ?? null, diff: f.diff ?? null, basis: f.basis || "" })) };
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
    return { summary, counts, findings: list, nextRow: (findings.length ? r : 9) + 2 };   // findings는 화면(HTML) 표시용, nextRow는 전기 대사 구역 시작 행
  }

  /** self·prior: 올해·작년 보고서의 표 정보(DSDCore snapshot). prior가 없으면 전기 대사를 하지 않는다(결과는 예전과 같음). */
  function run({ tables, notes, paras, ws, styles, self, prior }) {
    const all = tables.map(analyze);
    const As = all.filter(Boolean);
    const findings = [];
    for (const A of As) tableChecks(A, findings);
    crossChecks(As, notes, paras || [], findings);
    const order = new Map();
    tables.forEach((t) => { if (!order.has(t.sheet)) order.set(t.sheet, order.size); });
    const res = writeSheet(ws, findings, styles, order);
    if (prior && self) res.prior = writePrior(ws, priorCompare(self, all, prior), styles, order, res.nextRow);
    return res;
  }

  root.DSDChecks = { run, _internal: { analyze, normLabel, markerType, formulaOf, refsOf, periodOf, markPeriod, sentencePeriod, rowPeriod, partsOf, priorCompare } };
})(typeof window !== "undefined" ? window : globalThis);
