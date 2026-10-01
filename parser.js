/* แปลงข้อความ OT ที่ทีมส่งในไลน์ให้เป็นรายการ (ใช้ได้ทั้งในเบราว์เซอร์และ Node) */
(function (root) {
  var WIN = /^(\d{1,2})[.:](\d{2})\s*-\s*(\d{1,2})[.:](\d{2})\s*น?\.?$/;
  var TEAM = /^-?\s*ชุด\s*(.+?)\s*\(\s*(\d{1,2})[.:](\d{2})\s*-\s*(\d{1,2})[.:](\d{2})\s*น?\.?\s*\)/;
  var COUNT = /ชาย\s*(\d+)\s*[,，]?\s*หญิง\s*(\d+)/;
  var TASK = /รายละ.*?งาน\s*[:：]\s*(.*)$/;
  var STAFF = /^(\d+)\s*\.\s*(.+?)\s*\((.+)\)\s*$/;
  var HEAD = /(\d{1,2})\/(\d{1,2})\/(\d{4})/;
  var DAYS = ['อาทิตย์', 'จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์', 'เสาร์'];
  function p2(n) { return ('0' + n).slice(-2); }
  function clean(s) { return String(s).replace(/\s+/g, ' ').trim(); }

  function parseOtText(text) {
    var res = { date: '', project: '', items: [], warnings: [] };
    var lines = String(text).replace(/\r/g, '').split('\n').map(function (s) { return s.trim(); });
    var win = null, team = null;
    lines.forEach(function (ln, idx) {
      if (!ln) return;
      var m;
      if (!res.date && (m = HEAD.exec(ln))) {
        var y = Number(m[3]); if (y > 2400) y -= 543;
        res.date = y + '-' + p2(m[2]) + '-' + p2(m[1]);
        res.project = clean(ln.split(/\s+/)[0]);
        var wd = /วัน(อาทิตย์|จันทร์|อังคาร|พุธ|พฤหัสบดี|ศุกร์|เสาร์)/.exec(ln);
        if (wd) {
          var real = DAYS[new Date(Date.UTC(y, Number(m[2]) - 1, Number(m[1]))).getUTCDay()];
          if (real !== wd[1]) res.warnings.push('ข้อความระบุ "วัน' + wd[1] + '" แต่วันที่ ' + res.date + ' ตรงกับวัน' + real + ' กรุณาตรวจสอบวันที่');
        }
        return;
      }
      if (/^STAFF$/i.test(ln)) { team = null; return; }
      if ((m = WIN.exec(ln))) { win = { s: p2(m[1]) + ':' + m[2], e: p2(m[3]) + ':' + m[4] }; team = null; return; }
      if ((m = TEAM.exec(ln))) {
        team = { type: 'TEAM', group: clean(m[1]), name: '', start: p2(m[2]) + ':' + m[3], end: p2(m[4]) + ':' + m[5],
                 male: 0, female: 0, task: '', _line: idx + 1 };
        res.items.push(team); win = null; return;
      }
      if (team && (m = COUNT.exec(ln))) { team.male = Number(m[1]); team.female = Number(m[2]); team._c = true; return; }
      if (team && (m = TASK.exec(ln))) { team.task = clean(m[1]); return; }
      if (win && !team && (m = STAFF.exec(ln))) {
        res.items.push({ type: 'STAFF', group: '', name: clean(m[2]), start: win.s, end: win.e, male: 0, female: 0, task: clean(m[3]) });
        return;
      }
      res.warnings.push('บรรทัด ' + (idx + 1) + ' อ่านไม่ออก: ' + ln.slice(0, 40));
    });
    res.items.forEach(function (it) {
      if (it.type === 'TEAM' && !it._c) res.warnings.push('ชุด ' + it.group + ' ไม่มีจำนวนชาย/หญิง');
      delete it._c; delete it._line;
    });
    if (!res.date) res.warnings.push('ไม่พบวันที่ในบรรทัดแรก (รูปแบบ วว/ดด/ปปปป)');
    if (!res.items.length) res.warnings.push('ไม่พบรายการ OT');
    return res;
  }

  function summarize(items) {
    var t = { people: 0, hours: 0, staff: 0, teamPeople: 0, teams: 0, male: 0, female: 0, byWindow: {} };
    items.forEach(function (it) {
      var hc = it.type === 'TEAM' ? it.male + it.female : 1;
      var a = it.start.split(':'), b = it.end.split(':');
      var hr = (Number(b[0]) * 60 + Number(b[1]) - Number(a[0]) * 60 - Number(a[1])) / 60;
      var k = it.start + '-' + it.end;
      var w = t.byWindow[k] = t.byWindow[k] || { people: 0, hours: 0 };
      w.people += hc; w.hours += hc * hr; t.people += hc; t.hours += hc * hr;
      if (it.type === 'TEAM') { t.teamPeople += hc; t.teams++; t.male += it.male; t.female += it.female; } else t.staff++;
    });
    return t;
  }

  var api = { parseOtText: parseOtText, summarize: summarize };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.OtParser = api;
})(typeof window !== 'undefined' ? window : this);
