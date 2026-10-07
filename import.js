/* นำเข้าข้อมูล OT จากไฟล์ Excel / CSV / PDF / รูปภาพ (OCR ในเบราว์เซอร์)
 * โมเดลกลาง: rows = [ [ {t:'ข้อความ', x:ซ้าย, c:กึ่งกลาง}, ... ], ... ]
 * - Excel/CSV: x=c=ลำดับคอลัมน์   - PDF/รูป: x,c = พิกเซลตำแหน่งจริง
 */
(function (root) {
  'use strict';

  /* ---------- ตรวจตารางรายเดือน (ชื่อ x วันที่ 1..31) ---------- */
  function isInt(s) { return /^\d{1,2}$/.test(String(s).trim()); }

  function findHeaders(rows) {
    const out = [];
    rows.forEach(function (r, ri) {
      const nums = r.filter(function (c) { return isInt(c.t) && +c.t >= 1 && +c.t <= 31; });
      let best = null, run = [];
      function close() { if (run.length >= 5 && (!best || run.length > best.length)) best = run; }
      nums.forEach(function (c) {
        if (run.length && +c.t === +run[run.length - 1].t + 1) run.push(c); else { close(); run = [c]; }
      });
      close();
      if (best) out.push({ ri: ri, days: best });
    });
    return out;
  }
  function findHeader(rows) { return findHeaders(rows)[0] || null; }

  let FIXDEC = false;   // OCR มักทำจุดทศนิยมหาย (8.0 -> 80)
  function num(s) {
    const t = String(s).replace(/,/g, '').trim();
    if (!/^\d+(\.\d+)?$/.test(t)) return null;
    if (FIXDEC && /^\d{2,3}$/.test(t) && /0$/.test(t)) return parseFloat(t) / 10;
    return parseFloat(t);
  }
  function cleanName(s) { return String(s).replace(/\s*[-–]\s*$/, '').replace(/\s+/g, ' ').trim(); }

  /* เพศ: แปลงทุกรูปแบบเป็น 'ชาย' / 'หญิง' (อื่น ๆ = '' ให้ผู้ใช้ตรวจ) */
  function normSex(v) {
    const t = String(v || '').trim().toLowerCase().replace(/\./g, '');
    if (/^(m|male|man|ชาย|นาย|ช)$/.test(t)) return 'ชาย';
    if (/^(f|female|woman|หญิง|นาง|นางสาว|น\.?ส|ญ)$/.test(t)) return 'หญิง';
    return '';
  }
  const TH_MONTH = ['มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];
  /* ข้อมูลหัวไฟล์: เวลาเริ่ม OT, เดือน/ปี, รหัสโครงการ (PRJ-2026-J-074 -> J74 : ตัวอักษร + เลขท้าย 2 ตัว) */
  function parseMeta(rows) {
    const meta = { start: '', year: 0, month: 0, project: '', projectRaw: '' };
    rows.slice(0, 8).forEach(function (r) {
      const line = r.map(function (c) { return c.t; }).join(' ');
      let m;
      if (!meta.start && (m = /เวลาเริ่ม\s*OT\s*[:：]?\s*(\d{1,2})\s*[.:]\s*(\d{2})/i.exec(line))) meta.start = ('0' + m[1]).slice(-2) + ':' + m[2];
      if (!meta.month && (m = /เดือน\s*(\S+)\s*(\d{4})/.exec(line))) {
        const mi = TH_MONTH.indexOf(m[1]);
        if (mi >= 0) { meta.month = mi + 1; meta.year = +m[2] > 2400 ? +m[2] - 543 : +m[2]; }
      }
      if (!meta.project && (m = /PRJ[-_\s]*\d{4}[-_\s]*([A-Za-z])[-_\s]*(\d+)/i.exec(line))) {
        meta.projectRaw = m[0]; meta.project = m[1].toUpperCase() + m[2].slice(-2);
      }
    });
    return meta;
  }
  /* วันเริ่มต้นที่ควรเลือก: วันล่าสุดที่มีข้อมูล แต่ไม่เกินวันปัจจุบัน (todayIso = yyyy-mm-dd เวลาไทย) */
  function defaultDay(grid, todayIso) {
    const t = String(todayIso).split('-').map(Number), meta = grid.meta || {};
    let cutoff = 31;
    if (meta.year && meta.month) {
      const a = meta.year * 12 + meta.month, b = t[0] * 12 + t[1];
      cutoff = a < b ? 31 : a === b ? t[2] : 0;
    } else cutoff = t[2];
    const days = grid.days.filter(function (d) { return d <= cutoff && grid.counts[d] > 0; });
    return days.length ? Math.max.apply(null, days) : 0;
  }

  /* ชนิดของกลุ่มจากหัวข้อ: Supply -> SUPPLY | Worker/แรงงาน -> TEAM | Staff/พนักงาน -> STAFF | ไม่ระบุ: มีคอลัมน์ชื่อชุด = TEAM */
  function secType(title, hasTeamCol) {
    const t = String(title || '');
    if (/supply|ซัพพลาย/i.test(t)) return 'SUPPLY';
    if (/worker|แรงงาน/i.test(t)) return 'TEAM';
    if (/staff|พนักงาน/i.test(t)) return 'STAFF';
    return hasTeamCol ? 'TEAM' : 'STAFF';
  }

  /* อ่านตารางรายเดือน อาจมีหลายกลุ่มในไฟล์/รูปเดียว: Staff / Worker / Supply Contract
   * - ตารางที่มีคอลัมน์ "ชื่อชุด" (Worker, Supply Contract) -> รวมเป็นชุดช่าง (TEAM)
   * - ตารางที่ไม่มี (Staff) -> รายบุคคล (STAFF) */
  function parseGrid(rows, opts) {
    FIXDEC = !!(opts && opts.fixDecimal);
    const heads = findHeaders(rows);
    if (!heads.length) return FIXDEC ? parseHeadless(rows) : null;
    const warnings = [], sections = [];
    heads.forEach(function (h, k) {
      const endRi = k + 1 < heads.length ? heads[k + 1].ri : rows.length;
      const days = h.days.map(function (c) { return { d: +c.t, c: c.c }; });
      let spacing = Infinity;
      for (let i = 1; i < days.length; i++) spacing = Math.min(spacing, Math.abs(days[i].c - days[i - 1].c));
      const tol = spacing / 2 + 0.01;
      const firstDayX = Math.min.apply(null, h.days.map(function (c) { return c.x; })) - tol;

      // ชื่อกลุ่ม (Staff/Worker/Supply Contract) จากแถวเหนือหัวตาราง
      let title = '';
      for (let j = h.ri - 1; j >= Math.max(0, h.ri - 3) && !title; j--) {
        rows[j].forEach(function (c) { const m = /(supply\s*(?:contract|manpower)?|ซัพพลาย|worker|แรงงาน|staff|พนักงาน)/i.exec(c.t); if (m && !title) title = m[1]; });
      }

      const head = rows[h.ri].concat(rows[h.ri - 1] && !title ? [] : []).filter(function (c) { return c.c < firstDayX; });
      const nameH = head.filter(function (c) { return /ชื่อ-|ชื่อ|ชือ|นามสกุล|name/i.test(c.t) && !/ชุด/.test(c.t); })[0];
      const posH = head.filter(function (c) { return /ตำแหน่ง|ตําแหน่ง|แหน่ง|position/i.test(c.t); })[0];
      const codeH = head.filter(function (c) { return /รหัส|code|emp/i.test(c.t); })[0];
      const teamH = head.filter(function (c) { return /ชื่อชุด|ชุด|team|crew/i.test(c.t); })[0];
      const sexH = head.filter(function (c) { return /^(เพศ|gender|sex)$/i.test(c.t.trim()); })[0];
      const colH = head.filter(function (c) { return c.t.trim(); }).sort(function (a, b) { return a.x - b.x; });
      if (!nameH) warnings.push((title || 'ตาราง') + ': ไม่พบหัวคอลัมน์ "ชื่อ" จึงเดาคอลัมน์ชื่อจากข้อความทางซ้าย');

      const pick = function (left, hc) {
        const idx = colH.indexOf(hc);
        return left.filter(function (c) {
          let own = -1;
          for (let i = 0; i < colH.length; i++) {
            const gap = i ? colH[i].x - colH[i - 1].x : 20;
            if (colH[i].x <= c.x + 0.25 * gap) own = i;
          }
          return own === idx;
        }).map(function (c) { return c.t.trim(); }).join(' ');
      };

      const people = [];
      for (let ri = h.ri + 1; ri < endRi; ri++) {
        const r = rows[ri];
        const left = r.filter(function (c) { return c.c < firstDayX && c.t.trim(); });
        let name = '', pos = '', team = '', sex = '', code = '';
        if (nameH) {
          name = pick(left, nameH);
          if (posH) pos = pick(left, posH);
          if (teamH) team = pick(left, teamH);
          if (sexH) sex = pick(left, sexH);
          if (codeH) code = pick(left, codeH);
        } else {
          const th = left.filter(function (c) { return /[฀-๿]{3,}|[A-Za-z]{3,}/.test(c.t); }).sort(function (a, b) { return b.t.length - a.t.length; })[0];
          name = th ? th.t.trim() : '';
        }
        name = cleanName(name);
        team = cleanName(team).replace(/\.{2,}|…/g, '').trim();
        if (!name || !/[฀-๿a-zA-Z]{2,}/.test(name)) continue;
        if (/^(staff|worker|supply|พนักงาน|แรงงาน|ซัพพลาย)/i.test(name)) continue;
        const hours = {};
        days.forEach(function (dc) {
          const cell = r.filter(function (c) { return Math.abs(c.c - dc.c) <= tol; })[0];
          const v = cell ? num(cell.t) : null;
          if (v) hours[dc.d] = v;
        });
        people.push({ name: name, position: pos, code: String(code || '').trim(), team: team, sex: normSex(sex), sexRaw: sex, hours: hours });
      }
      if (people.length) sections.push({ title: title || (teamH ? 'Worker' : 'Staff'), type: secType(title, teamH), people: people, days: days.map(function (d) { return d.d; }) });
    });
    if (!sections.length) return null;
    const dayList = sections[0].days.slice();
    sections.forEach(function (sc) { sc.days.forEach(function (d) { if (dayList.indexOf(d) < 0) dayList.push(d); }); });
    dayList.sort(function (a, b) { return a - b; });
    const counts = {};
    dayList.forEach(function (d) { counts[d] = sections.reduce(function (n, sc) { return n + sc.people.filter(function (p) { return p.hours[d]; }).length; }, 0); });
    return { kind: 'grid', sections: sections, days: dayList, counts: counts, warnings: warnings, meta: parseMeta(rows) };
  }


  /* รูปที่ OCR ไม่เจอหัวตาราง (ตัวหนังสือขาวบนพื้นม่วง): หาคอลัมน์วันจากตำแหน่งตัวเลขในแถวข้อมูลแทน
   * ถือว่าคอลัมน์แรกที่เห็นคือวันที่ 1 (ต้องให้ผู้ใช้ตรวจ) */
  function parseHeadless(rows) {
    const isCode = function (t) { return /^[A-Za-z]{0,3}\d{4,7}$/.test(t); };
    const isHr = function (t) { return /^\d{1,2}\.\d$/.test(t) || /^\d{2,3}$/.test(t) && /0$/.test(t) && +t <= 240; };
    const wordy = function (t) { return /[฀-๿]{2,}|[A-Za-z]{3,}/.test(t); };
    const data = [];
    rows.forEach(function (r, ri) {
      const ci = r.findIndex(function (c) { return isCode(c.t); });
      if (ci < 0) return;
      const after = r.slice(ci + 1);
      const wi = after.findIndex(function (c) { return wordy(c.t); });
      if (wi < 0) return;
      const nums = after.slice(wi + 1).filter(function (c) { return isHr(c.t); });
      if (nums.length >= 2) data.push({ ri: ri, ci: ci });
    });
    if (data.length < 3) return null;
    // ขอบซ้ายของโซนวัน = ขวาสุดของข้อความ (ชื่อ/ตำแหน่ง/ป้าย) ในแถวข้อมูลทั้งหมด
    let textRight = 0;
    data.forEach(function (d) {
      const r = rows[d.ri];
      for (let i = d.ci + 1; i < r.length; i++) if (wordy(r[i].t) && !isHr(r[i].t)) textRight = Math.max(textRight, r[i].x1 || r[i].c);
    });
    const cells = [];
    data.forEach(function (d) { rows[d.ri].forEach(function (c) { if (c.c > textRight && isHr(c.t)) cells.push(c.c); }); });
    if (cells.length < 6) return null;
    cells.sort(function (a, b) { return a - b; });
    const cl = [];
    cells.forEach(function (x) { const l = cl[cl.length - 1]; if (l && x - l.s / l.n < 14) { l.s += x; l.n++; } else cl.push({ s: x, n: 1 }); });
    const cen = cl.map(function (c) { return c.s / c.n; });
    if (cen.length < 5) return null;
    const diffs = []; for (let i = 1; i < cen.length; i++) diffs.push(cen[i] - cen[i - 1]);
    const minD = Math.min.apply(null, diffs), near = diffs.filter(function (d) { return d < minD * 1.4; }).sort(function (a, b) { return a - b; });
    const pitch = near[Math.floor(near.length / 2)];
    const idxOf = function (x) { return Math.round((x - cen[0]) / pitch) + 1; };
    const tol = pitch * 0.45;
    const maxDay = Math.min(31, idxOf(cen[cen.length - 1]));
    const days = []; for (let d = 1; d <= maxDay; d++) days.push(d);
    // แบ่งกลุ่มจากหัวข้อ Staff/Worker/Supply Contract ที่อยู่เหนือแถวข้อมูล
    const secs = {}, order = [];
    data.forEach(function (d) {
      let title = '';
      for (let j = d.ri - 1; j >= 0 && !title; j--) {
        const m = rows[j].map(function (c) { return c.t; }).join(' ').match(/(supply\s*(?:contract|manpower)?|ซัพพลาย|worker|แรงงาน|staff|พนักงาน)/i);
        if (m) title = m[1]; else if (data.some(function (x) { return x.ri === j; })) { title = '__prev'; }
      }
      secs[d.ri] = title;
    });
    let cur = 'Staff';
    const people = {};
    data.forEach(function (d) {
      if (secs[d.ri] && secs[d.ri] !== '__prev') cur = secs[d.ri].replace(/\s+/g, ' ');
      const r = rows[d.ri], team = /worker|supply/i.test(cur);
      const left = r.slice(d.ci + 1).filter(function (c) { return c.c <= textRight && wordy(c.t); });
      const name = cleanName((left[0] || { t: '' }).t), pos = (left[1] || { t: '' }).t;
      if (!name) return;
      const hours = {};
      r.forEach(function (c) { if (c.c > textRight && isHr(c.t)) { const v = num(c.t); if (v) hours[idxOf(c.c)] = v; } });
      (people[cur] = people[cur] || (order.push(cur), [])).push({ name: name, position: pos, team: team ? cleanName((left[2] || { t: '' }).t).replace(/\.{2,}|…/g, '') : '', hours: hours });
    });
    const sections = order.map(function (t) { return { title: t, type: secType(t, false), people: people[t], days: days }; });
    const counts = {};
    days.forEach(function (d) { counts[d] = sections.reduce(function (n, sc) { return n + sc.people.filter(function (p) { return p.hours[d]; }).length; }, 0); });
    return { kind: 'grid', sections: sections, days: days, counts: counts,
             warnings: ['อ่านหัวคอลัมน์วันที่จากรูปไม่ได้ จึงนับคอลัมน์แรกที่เห็นเป็นวันที่ 1 — ตรวจสอบจำนวนคนของแต่ละวันให้ตรงกับรูปก่อนนำเข้า'] };
  }

  /* ผลลัพธ์ของวันที่เลือก -> รายการสำหรับฟอร์ม
   * STAFF: ต่อคน | TEAM: รวมตามชื่อชุด+จำนวนชั่วโมง (ไม่มีข้อมูลเพศในตาราง ให้ผู้ใช้แบ่งชาย/หญิงเอง) */
  function itemsForDay(grid, day) {
    const out = [];
    grid.sections.forEach(function (sc) {
      if (sc.type === 'STAFF') {
        sc.people.forEach(function (p) { if (p.hours[day]) out.push({ type: 'STAFF', name: p.name, position: p.position || '', dc: /^dc/i.test(String(p.code || '')), hours: p.hours[day], section: sc.title }); });
      } else {
        const m = {}, order = [];
        sc.people.forEach(function (p) {
          if (!p.hours[day]) return;
          const t = p.team || '(ไม่ระบุชุด)', ps = String(p.position || '').trim(), k = t + '|' + ps + '|' + p.hours[day];   // แยกแถวตามตำแหน่ง เพื่อคิดค่าแรงตามตำแหน่งได้
          if (!m[k]) { m[k] = { type: sc.type, name: t, position: ps, dc: true, hours: p.hours[day], count: 0, male: 0, female: 0, unknown: 0, section: sc.title }; order.push(k); }
          m[k].count++;
          if (p.sex === 'ชาย') m[k].male++; else if (p.sex === 'หญิง') m[k].female++; else m[k].unknown++;
        });
        order.forEach(function (k) { out.push(m[k]); });
      }
    });
    return out;
  }

  /* ---------- แปลง rows -> ข้อความ ส่งให้ตัวอ่านข้อความรายวันเดิม ---------- */
  function rowsToText(rows) {
    return rows.map(function (r) { return r.map(function (c) { return c.t; }).join(' ').replace(/\s+/g, ' ').trim(); })
      .filter(Boolean).join('\n');
  }

  /* ---------- ตัวโหลดไลบรารีตามต้องการ ---------- */
  const loaded = {};
  function loadScript(url) {
    if (loaded[url]) return loaded[url];
    loaded[url] = new Promise(function (ok, bad) {
      const s = document.createElement('script'); s.src = url; s.onload = ok;
      s.onerror = function () { bad(new Error('โหลดไลบรารีไม่สำเร็จ (ตรวจสอบอินเทอร์เน็ต)')); };
      document.head.appendChild(s);
    });
    return loaded[url];
  }
  const CDN = 'https://cdn.jsdelivr.net/npm/';

  /* ---------- Excel / CSV ---------- */
  async function readSheet(file) {
    await loadScript(CDN + 'xlsx@0.18.5/dist/xlsx.full.min.js');
    const wb = XLSX.read(await file.arrayBuffer(), { type: 'array' });
    const out = [];
    wb.SheetNames.forEach(function (n) {
      const aoa = XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: false, defval: '' });
      aoa.forEach(function (row) {
        const cells = [];
        row.forEach(function (v, j) { const t = String(v).trim(); if (t) cells.push({ t: t, x: j, c: j }); });
        out.push(cells);
      });
    });
    return out;
  }

  /* ---------- PDF (ที่มีตัวอักษรจริง ไม่ใช่รูปสแกน) ---------- */
  async function readPdf(file) {
    await loadScript(CDN + 'pdfjs-dist@3.11.174/build/pdf.min.js');
    pdfjsLib.GlobalWorkerOptions.workerSrc = CDN + 'pdfjs-dist@3.11.174/build/pdf.worker.min.js';
    const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(await file.arrayBuffer()) }).promise;
    const rows = [];
    for (let p = 1; p <= pdf.numPages; p++) {
      const tc = await (await pdf.getPage(p)).getTextContent();
      const items = tc.items.filter(function (i) { return i.str.trim(); })
        .map(function (i) { return { t: i.str.trim(), x: i.transform[4], c: i.transform[4] + (i.width || 0) / 2, y: i.transform[5] }; })
        .sort(function (a, b) { return b.y - a.y || a.x - b.x; });
      let cur = null;
      items.forEach(function (i) {
        if (!cur || Math.abs(cur.y - i.y) > 3) { cur = { y: i.y, cells: [] }; rows.push(cur.cells); }
        cur.cells.push({ t: i.t, x: i.x, c: i.c });
      });
    }
    if (!rows.length) throw new Error('PDF นี้ไม่มีข้อความ (อาจเป็นรูปสแกน) กรุณาแคปเป็นรูปภาพแล้วอัปโหลดแทน');
    return rows;
  }

  /* ---------- รูปภาพ -> OCR ---------- */
  async function readImage(file, onProgress) {
    await loadScript(CDN + 'tesseract.js@5/dist/tesseract.min.js');
    const res = await Tesseract.recognize(file, 'tha+eng', {
      logger: function (m) { if (onProgress && m.status === 'recognizing text') onProgress(m.progress); }
    });
    let rows = wordsToRows(res.data.words || [], {});
    if (!findHeader(rows)) { if (onProgress) onProgress(0.97); try { rows = await recoverHeader(file, rows); } catch (e) { /* ใช้ผลเดิม */ } }
    return rows;
  }
  function wordsToRows(words, o) {
    o = o || {};
    const sc = o.scale || 1, ox = o.left || 0, oy = o.top || 0;
    const w = words.filter(function (x) { return x.text && x.text.trim() && x.confidence > 20; })
      .map(function (x) { return { t: x.text.trim(), x0: ox + x.bbox.x0 / sc, x1: ox + x.bbox.x1 / sc, yc: oy + (x.bbox.y0 + x.bbox.y1) / 2 / sc, h: (x.bbox.y1 - x.bbox.y0) / sc }; });
    if (!w.length) return [];
    const hs = w.map(function (x) { return x.h; }).sort(function (a, b) { return a - b; });
    const mh = o.mh || hs[Math.floor(hs.length / 2)] || 12;
    w.sort(function (a, b) { return a.yc - b.yc; });
    const lines = [];
    w.forEach(function (x) {
      const l = lines[lines.length - 1];
      if (l && Math.abs(l.yc - x.yc) < mh * 0.6) { l.w.push(x); l.yc = (l.yc * (l.w.length - 1) + x.yc) / l.w.length; }
      else lines.push({ yc: x.yc, w: [x] });
    });
    return lines.map(function (l) {
      l.w.sort(function (a, b) { return a.x0 - b.x0; });
      const cells = [];
      l.w.forEach(function (x) {
        const last = cells[cells.length - 1];
        const isNum = /^[\d.]+$/.test(x.t);
        const gap = last ? x.x0 - last.x1 : 0;
        if (last && !isNum && !last.num && gap < mh * 0.9) {
          last.t += (gap < mh * 0.35 ? '' : ' ') + x.t; last.x1 = Math.max(last.x1, x.x1);   // ภาษาไทย OCR ตัดเป็นตัวอักษร: ช่องไฟเล็กให้ต่อกัน
        } else cells.push({ t: x.t, x0: x.x0, x1: x.x1, num: isNum });
      });
      const out = cells.map(function (c) { return { t: c.t, x: c.x0, c: (c.x0 + c.x1) / 2, x1: c.x1 }; });
      out.y = l.yc; out.mh = mh;
      return out;
    });
  }

  /* ใต้พื้นสีเข้ม ตัวหนังสือขาว OCR อ่านไม่ออก -> ตัดแถบหัวตารางมาทำภาพขาว-ดำก่อนอ่านซ้ำ */
  async function recoverHeader(file, rows) {
    const isDec = function (t) { return /^\d+(\.\d)?$/.test(t); };
    const data = [];
    rows.forEach(function (r, i) {
      if (r.filter(function (c) { return isDec(c.t); }).length >= 3 && r.some(function (c) { return /[\u0E00-\u0E7F]{3,}/.test(c.t); })) data.push(i);
    });
    if (!data.length) return rows;
    const i0 = data[0], r0 = rows[i0], pitch = data.length > 1 ? Math.abs(rows[data[1]].y - r0.y) : r0.mh * 3;
    const textRight = Math.max.apply(null, r0.filter(function (c) { return !isDec(c.t) && !/^[\d.]+$/.test(c.t); }).map(function (c) { return c.x1; }));
    const top = Math.max(0, Math.floor(r0.y - pitch * 1.5)), hgt = Math.ceil(pitch);
    const bmp = await createImageBitmap(file);
    function band(x0, x1) {
      const cv = document.createElement('canvas'); const sc = 2;
      cv.width = Math.max(1, Math.floor((x1 - x0) * sc)); cv.height = hgt * sc;
      const cx = cv.getContext('2d'); cx.drawImage(bmp, x0, top, x1 - x0, hgt, 0, 0, cv.width, cv.height);
      const d = cx.getImageData(0, 0, cv.width, cv.height), px = d.data;
      for (let i = 0; i < px.length; i += 4) { const g = (px[i] + px[i + 1] + px[i + 2]) / 3; const v = g > 215 ? 0 : 255; px[i] = px[i + 1] = px[i + 2] = v; }
      cx.putImageData(d, 0, 0); return cv;
    }
    const dayLeft = Math.floor(textRight + 4);
    const firstNum = Math.min.apply(null, r0.filter(function (c) { return c.x > textRight; }).map(function (c) { return c.x; }));
    const W = bmp.width;
    const hdr = [];
    // ซ้าย: ชื่อหัวคอลัมน์ (ไทย+อังกฤษ)
    const leftX0 = Math.max(0, Math.floor(Math.min.apply(null, r0.map(function (c) { return c.x; })) - 20));
    const wk = await Tesseract.createWorker('tha+eng');
    await wk.setParameters({ tessedit_pageseg_mode: '7' });
    const L = await wk.recognize(band(leftX0, Math.min(W, dayLeft)));
    wordsToRows(L.data.words, { scale: 2, left: leftX0, top: 0, mh: r0.mh }).forEach(function (r) { r.forEach(function (c) { hdr.push(c); }); });
    // ขวา: เลขวันที่ (ตัวเลขล้วน)
    await wk.setParameters({ tessedit_pageseg_mode: '7', tessedit_char_whitelist: '0123456789' });
    const R = await wk.recognize(band(dayLeft, W));
    R.data.words.forEach(function (x) {
      if (!/^\d{1,2}$/.test(x.text)) return;
      const a = dayLeft + x.bbox.x0 / 2, b = dayLeft + x.bbox.x1 / 2;
      hdr.push({ t: x.text, x: a, c: (a + b) / 2, x1: b });
    });
    await wk.terminate();
    hdr.sort(function (a, b) { return a.x - b.x; });
    const copy = rows.slice(); copy.splice(i0, 0, hdr);
    return copy;
  }

  /* ---------- จุดเข้าหลัก ---------- */
  async function readFile(file, onProgress) {
    const n = (file.name || '').toLowerCase();
    let rows;
    if (/\.(xlsx|xls|csv)$/.test(n)) rows = await readSheet(file);
    else if (/\.pdf$/.test(n) || file.type === 'application/pdf') rows = await readPdf(file);
    else if ((file.type || '').indexOf('image/') === 0) rows = await readImage(file, onProgress);
    else throw new Error('ไม่รองรับไฟล์ชนิดนี้ (รองรับ Excel, CSV, PDF, รูปภาพ)');
    const isImg = (file.type || '').indexOf('image/') === 0;
    const grid = parseGrid(rows, { fixDecimal: isImg });
    return { rows: rows, grid: grid, text: rowsToText(rows), fromImage: (file.type || '').indexOf('image/') === 0 };
  }

  root.OtImport = { parseGrid: parseGrid, parseMeta: parseMeta, defaultDay: defaultDay, normSex: normSex, itemsForDay: itemsForDay, _recoverHeader: function(f, r) { return recoverHeader(f, r); }, rowsToText: rowsToText, wordsToRows: wordsToRows, readFile: readFile };
  if (typeof module !== 'undefined') module.exports = root.OtImport;
})(typeof window !== 'undefined' ? window : globalThis);
