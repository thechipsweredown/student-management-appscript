// ============================================================
// PANDA HAPPY — Quản Lý Học Sinh | Google Apps Script
// Schema mới: students / enrollments / classes / monthly_bills / payments
// ============================================================

var TIMEZONE = 'Asia/Ho_Chi_Minh';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Panda Happy — Quản Lý Học Sinh')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getSheet(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function fmtDate(d) {
  if (d instanceof Date) return Utilities.formatDate(d, TIMEZONE, 'dd/MM/yyyy');
  // Date serial từ Sheets API batchGet (UNFORMATTED_VALUE)
  if (typeof d === 'number' && d > 1000) {
    var dt = new Date(Math.round((d - 25569) * 86400000));
    return Utilities.formatDate(dt, TIMEZONE, 'dd/MM/yyyy');
  }
  if (d && String(d).match(/^\d{4}-\d{2}-\d{2}/)) {
    var p = String(d).split('T')[0].split('-');
    return p[2] + '/' + p[1] + '/' + p[0];
  }
  return d ? String(d).substring(0, 10) : '';
}

// Đọc nhiều sheet trong 1 API call.
// Nếu Google Sheets API advanced service chưa bật → fallback sang đọc từng sheet.
function batchReadSheets(names) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (typeof Sheets !== 'undefined') {
    var resp = Sheets.Spreadsheets.Values.batchGet(ss.getId(), {
      ranges: names,
      majorDimension: 'ROWS',
      valueRenderOption: 'UNFORMATTED_VALUE',
      dateTimeRenderOption: 'SERIAL_NUMBER'
    });
    var map = {};
    (resp.valueRanges || []).forEach(function(vr, i) {
      map[names[i]] = vr.values || [];
    });
    return map;
  }
  // Fallback: đọc từng sheet
  var map = {};
  names.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    map[name] = sheet ? sheet.getDataRange().getValues() : [];
  });
  return map;
}

// Parse raw 2D array thành array of objects dùng row 0 làm header
function parseFromValues(rows) {
  if (!rows || rows.length < 2) return [];
  var headers = rows[0].map(function(h) { return String(h || '').trim(); });
  var result = [];
  for (var i = 1; i < rows.length; i++) {
    var row = rows[i];
    if (row[0] === '' && row[1] === '') continue;
    if (row[0] === undefined && row[1] === undefined) continue;
    var obj = { _row: i + 1 };
    headers.forEach(function(h, ci) { if (h) obj[h] = ci < row.length ? row[ci] : ''; });
    result.push(obj);
  }
  return result;
}

// Tính điểm danh 7 ngày gần nhất từ raw parsed data
function _computeAtt7days(rawSessions, rawAtt) {
  var isoDate = function(v) {
    if (v instanceof Date) return Utilities.formatDate(v, TIMEZONE, 'yyyy-MM-dd');
    if (typeof v === 'number' && v > 1000) return Utilities.formatDate(new Date(Math.round((v - 25569) * 86400000)), TIMEZONE, 'yyyy-MM-dd');
    return String(v).substring(0, 10);
  };
  var today = new Date(), days = [];
  for (var i = 6; i >= 0; i--) { var d = new Date(today.getTime()); d.setDate(d.getDate() - i); days.push(Utilities.formatDate(d, TIMEZONE, 'yyyy-MM-dd')); }
  var sesDate = {};
  (rawSessions || []).forEach(function(s) { var iso = isoDate(s.date); if (days.indexOf(iso) !== -1) sesDate[String(s.id || '')] = iso; });
  var cnt = {};
  days.forEach(function(dy) { cnt[dy] = { present: 0, absent: 0, late: 0 }; });
  (rawAtt || []).forEach(function(a) {
    var date = sesDate[String(a.session_id || '')]; if (!date) return;
    var st = String(a.status || '');
    if (st === 'Có mặt') cnt[date].present++; else if (st === 'Vắng') cnt[date].absent++; else if (st === 'Đi muộn') cnt[date].late++;
  });
  return days.map(function(dy) { return { date: dy.substring(5), present: cnt[dy].present, absent: cnt[dy].absent, late: cnt[dy].late }; });
}

// Parse sheet thành array of objects, dùng row 1 làm header
function parseSheet(name) {
  var sheet = getSheet(name);
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0].map(function(h) { return String(h || '').trim(); });
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0] && !data[i][1]) continue;
    var obj = { _row: i + 1 };
    headers.forEach(function(h, ci) { if (h) obj[h] = data[i][ci]; });
    result.push(obj);
  }
  return result;
}

function genId(sheetName, prefix) {
  var sheet = getSheet(sheetName);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return prefix + '0001';
  var col1 = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var maxN = 0;
  col1.forEach(function(r) {
    var m = String(r[0]).match(new RegExp('^' + prefix + '(\\d+)'));
    if (m) maxN = Math.max(maxN, parseInt(m[1]));
  });
  return prefix + String(maxN + 1).padStart(4, '0');
}

// ── Auth ─────────────────────────────────────────────────────
// TODO: real auth bị lỗi khi cấu hình nhiều admin — tạm hardcode cho demo
// Uncomment đoạn dưới và xoá dòng return hardcode khi fix xong
function getCurrentUser() {
  var email = Session.getActiveUser().getEmail();
  // var users = parseSheet('Phân quyền');
  // for (var i = 0; i < users.length; i++) {
  //   var u = users[i];
  //   if (String(u['email'] || '').toLowerCase() === email.toLowerCase()
  //       && String(u['status']) === 'active') {
  //     return {
  //       email     : email,
  //       role      : String(u['role'] || 'none'),
  //       name      : String(u['display_name'] || ''),
  //       teacher_id: String(u['teacher_id'] || '')
  //     };
  //   }
  // }
  // return { email: email, role: 'none', name: '', teacher_id: '' };

  return { email: 'dangductungcfc@gmail.com', role: 'admin', name: 'Quản Lý', teacher_id: '' };
}

// ── getAllData ────────────────────────────────────────────────
function getAllData(termId) {
  var T = { start: Date.now(), steps: [] };
  function mark(label) { T.steps.push({ label: label, ms: Date.now() - T.start }); }

  var SHEET_NAMES = ['students','enrollments','classes','teachers','monthly_bills','payments','DS Lớp','terms','sessions','attendance'];
  var raw = batchReadSheets(SHEET_NAMES);
  mark('batchRead:all_sheets');
  T.read = Date.now();

  var rawStudents = parseFromValues(raw['students']);      mark('parse:students('   + rawStudents.length + ')');
  var rawEnroll   = parseFromValues(raw['enrollments']);   mark('parse:enrollments(' + rawEnroll.length   + ')');
  var rawClasses  = parseFromValues(raw['classes']);       mark('parse:classes('    + rawClasses.length   + ')');
  var rawTeachers = parseFromValues(raw['teachers']);      mark('parse:teachers('   + rawTeachers.length  + ')');
  var rawBills    = parseFromValues(raw['monthly_bills']); mark('parse:bills('      + rawBills.length     + ')');
  var rawPays     = parseFromValues(raw['payments']);      mark('parse:payments('   + rawPays.length      + ')');
  var rawLop      = raw['DS Lớp'];
  var rawTermsArr = parseFromValues(raw['terms']);         mark('parse:terms('      + rawTermsArr.length  + ')');
  mark('parse:done');

  // Auto-detect kỳ hiện tại nếu không truyền termId
  var resolvedTermId = termId || '';
  if (!resolvedTermId && rawTermsArr.length) {
    var todaySerial = new Date().getTime() / 86400000 + 25569;
    var latestPastId = '', latestPastEnd = 0;
    rawTermsArr.forEach(function(t) {
      var s = typeof t.start_date === 'number' ? t.start_date : 0;
      var e = typeof t.end_date   === 'number' ? t.end_date   : 0;
      if (s && e && todaySerial >= s && todaySerial <= e) {
        resolvedTermId = String(t.id || '');
      } else if (e && e < todaySerial && e > latestPastEnd) {
        latestPastEnd = e;
        latestPastId  = String(t.id || '');
      }
    });
    if (!resolvedTermId && latestPastId) resolvedTermId = latestPastId;
    mark('detect:currentTerm(' + resolvedTermId + ')');
  }

  if (resolvedTermId) {
    rawEnroll = rawEnroll.filter(function(e) {
      return String(e.term_id || '') === String(resolvedTermId);
    });
    mark('filter:enrollments_by_term(' + rawEnroll.length + ')');
  }

  // ── Lookup maps ──────────────────────────────────────────
  var classMap = {};
  rawClasses.forEach(function(c) { if (c.id) classMap[String(c.id)] = c; });

  var teacherMap = {};
  rawTeachers.forEach(function(t) { if (t.id) teacherMap[String(t.id)] = t; });

  var enrollByStudent = {};
  rawEnroll.forEach(function(e) {
    var sid = String(e.student_id || '');
    if (!sid) return;
    if (!enrollByStudent[sid]) enrollByStudent[sid] = [];
    enrollByStudent[sid].push(e);
  });

  var billByEnroll = {};
  rawBills.forEach(function(b) {
    var eid = String(b.enrollment_id || '');
    if (!eid) return;
    if (!billByEnroll[eid]) billByEnroll[eid] = [];
    billByEnroll[eid].push(b);
  });

  var billMap = {};
  rawBills.forEach(function(b) { if (b.id) billMap[String(b.id)] = b; });

  var enrollMap = {};
  rawEnroll.forEach(function(e) { if (e.id) enrollMap[String(e.id)] = e; });

  var studentMap = {};
  rawStudents.forEach(function(s) { if (s.id) studentMap[String(s.id)] = s; });

  mark('compute:lookup_maps');

  var currentMonth = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM');

  // ── Parse students ────────────────────────────────────────
  var students = rawStudents.filter(function(s) {
    return s.id && String(s.id).match(/^HS\d+/i);
  }).map(function(s) {
    var enrolls = enrollByStudent[String(s.id)] || [];
    var classes = enrolls.map(function(e) {
      var cls     = classMap[String(e.class_id || '')] || {};
      var teacher = teacherMap[String(cls.teacher_id || '')] || {};
      var enrollTuition = parseFloat(e.tuition) || 0;
      var effectiveTuition = enrollTuition || (parseFloat(cls.tuition_per_month) || 0);
      var bills     = (billByEnroll[String(e.id || '')] || []).filter(function(b) {
        return !b.month || String(b.month) <= currentMonth;
      });
      var billDue   = bills.reduce(function(a, b) { return a + (parseFloat(b.amount_due)  || 0); }, 0);
      var paid      = bills.reduce(function(a, b) { return a + (parseFloat(b.amount_paid) || 0); }, 0);
      var debt      = billDue > 0 ? Math.max(0, billDue - paid) : 0;
      return {
        enrollId    : String(e.id || ''),
        classId     : String(e.class_id || ''),
        className   : String(cls.name || ''),
        teacherName : String(teacher.name || ''),
        tuition     : effectiveTuition,
        enrollStatus: String(e.status || ''),
        billDue     : billDue,
        billPaid    : paid,
        billDebt    : debt
      };
    });
    var primary   = classes.find(function(c) { return c.enrollStatus === 'Đang học'; }) || classes[0] || {};
    var totalDebt = classes.reduce(function(a, c) { return a + c.billDebt; }, 0);
    return {
      rowIndex       : s._row,
      id             : String(s.id),
      name           : String(s.name || ''),
      dob            : fmtDate(s.dob),
      gender         : String(s.gender || ''),
      photoUrl       : String(s.photo_url || ''),
      status         : String(s.status || 'Đang học'),
      address        : String(s.address || ''),
      classes        : classes,
      primaryClass   : String(primary.className || ''),
      primaryTeacher : String(primary.teacherName || ''),
      totalDebt      : totalDebt
    };
  });

  mark('compute:students(' + students.length + ')');

  // ── Parse payments ────────────────────────────────────────
  var payments = rawPays.filter(function(p) {
    return p.date instanceof Date || (p.date && String(p.date).match(/\d/));
  }).map(function(p) {
    var bill = billMap[String(p.bill_id || '')] || {};
    var stu  = studentMap[String(p.student_id || '')] || {};
    var cls  = classMap[String(p.class_id   || '')] || {};
    return {
      id         : String(p.id || ''),
      billId     : String(p.bill_id || ''),
      date       : fmtDate(p.date),
      studentId  : String(p.student_id || ''),
      studentName: String(stu.name || ''),
      classId    : String(p.class_id || ''),
      className  : String(cls.name || ''),
      month      : String(bill.month || ''),
      amount     : parseFloat(p.amount) || 0,
      method     : String(p.method || ''),
      staff      : String(p.staff  || ''),
      note       : String(p.note   || '')
    };
  });
  payments.sort(function(a, b) {
    var da = a.date.split('/').reverse().join('');
    var db = b.date.split('/').reverse().join('');
    return da > db ? -1 : 1;
  });
  mark('compute:payments(' + payments.length + ')');

  // ── Classes list ──────────────────────────────────────────
  // Build per-class aggregates in single O(n) pass over enrollments
  var paidByClass  = {};
  var debtByClass  = {};
  var countByClass = {};
  rawEnroll.forEach(function(e) {
    var cid = String(e.class_id || '');
    if (!cid) return;
    if (String(e.status) === 'Đang học') countByClass[cid] = (countByClass[cid] || 0) + 1;
    var bills = (billByEnroll[String(e.id || '')] || []).filter(function(b) {
      return !b.month || String(b.month) <= currentMonth;
    });
    bills.forEach(function(b) {
      paidByClass[cid] = (paidByClass[cid] || 0) + (parseFloat(b.amount_paid) || 0);
      debtByClass[cid] = (debtByClass[cid] || 0) + (parseFloat(b.debt)        || 0);
    });
  });

  var classes = rawClasses.map(function(c) {
    var teacher = teacherMap[String(c.teacher_id || '')] || {};
    var cid     = String(c.id || '');
    return {
      id        : cid,
      name      : String(c.name || ''),
      teacherId : String(c.teacher_id || ''),
      teacher   : String(teacher.name || ''),
      tuition   : parseFloat(c.tuition_per_month) || 0,
      subject   : String(c.subject || ''),
      status    : String(c.status || ''),
      count     : countByClass[cid] || 0,
      paid      : paidByClass[cid]  || 0,
      debt      : debtByClass[cid]  || 0
    };
  });

  mark('compute:classes(' + classes.length + ')');

  // ── DS Lớp classGroups ────────────────────────────────────
  var classGroups = [];
  if (rawLop.length >= 2) {
    for (var col = 0; col < Math.min(rawLop[0].length, 2); col++) {
      var header = String(rawLop[0][col] || '').trim();
      if (!header) continue;
      var grpCls = [];
      for (var row = 1; row < rawLop.length; row++) {
        var v = rawLop[row][col];
        if (!v || v instanceof Date) continue;
        var sv = String(v).trim();
        if (sv) grpCls.push(sv);
      }
      if (grpCls.length) classGroups.push({ name: header, classes: grpCls });
    }
  }

  // ── Stats ─────────────────────────────────────────────────
  var totalRevenue   = payments.reduce(function(a, p) { return a + p.amount; }, 0);
  var totalDebt      = students.reduce(function(a, s) { return a + s.totalDebt; }, 0);
  var activeStudents = students.filter(function(s) { return s.status === 'Đang học'; }).length;
  var activeClasses  = classes.filter(function(c) { return c.count > 0; }).length;
  var teacherCount   = rawTeachers.filter(function(t) { return String(t.status || '') === 'Đang dạy'; }).length;

  var monthlyMap = {};
  payments.forEach(function(p) {
    if (p.month) monthlyMap[p.month] = (monthlyMap[p.month] || 0) + p.amount;
  });
  var monthlyRevenue = Object.keys(monthlyMap).sort().map(function(m) {
    return { month: m, amount: monthlyMap[m] };
  });

  var debtMonthMap = {};
  rawEnroll.forEach(function(e) {
    var bills = (billByEnroll[String(e.id || '')] || []).filter(function(b) {
      return !b.month || String(b.month) <= currentMonth;
    });
    bills.forEach(function(b) {
      if (b.month) debtMonthMap[String(b.month)] = (debtMonthMap[String(b.month)] || 0) + (parseFloat(b.debt) || 0);
    });
  });
  var monthlyDebt = Object.keys(debtMonthMap).sort().map(function(m) {
    return { month: m, amount: debtMonthMap[m] };
  });

  var classRevMap = {};
  payments.forEach(function(p) {
    if (p.className) classRevMap[p.className] = (classRevMap[p.className] || 0) + p.amount;
  });

  mark('done');
  T.done = Date.now();

  var att7days = _computeAtt7days(parseFromValues(raw['sessions'] || []), parseFromValues(raw['attendance'] || []));

  var teacherList = rawTeachers.map(function(t) {
    return { id: String(t.id||''), name: String(t.name||''), status: String(t.status||'') };
  });

  return {
    students   : students,
    payments   : payments,
    classes    : classes,
    classGroups: classGroups,
    teachers   : teacherList,
    terms      : rawTermsArr.map(function(t) {
      return { id: String(t.id||''), school_year: String(t.school_year||''), term_name: String(t.term_name||''), start_date: fmtDate(t.start_date), end_date: fmtDate(t.end_date) };
    }),
    stats: {
      totalStudents : students.length,
      activeStudents: activeStudents,
      activeClasses : activeClasses,
      teacherCount  : teacherCount,
      totalRevenue  : totalRevenue,
      totalDebt     : totalDebt,
      totalPayments : payments.length,
      monthlyRevenue: monthlyRevenue,
      monthlyDebt   : monthlyDebt,
      classRevenue  : classRevMap
    },
    att7days     : att7days,
    currentTermId: resolvedTermId,
    _timing: { total: T.done - T.start, read: T.read - T.start, compute: T.done - T.read, steps: T.steps }
  };
}

// ── getQuickData: chỉ đọc 7 sheets nhẹ, trả về ngay để hiện UI ──
// Không có enrollments/monthly_bills/payments → debt=0, revenue=0
// getAllData chạy song song ở background để cập nhật đầy đủ sau
function getQuickData(termId) {
  var T = { start: Date.now(), steps: [] };
  function mark(label) { T.steps.push({ label: label, ms: Date.now() - T.start }); }

  var raw = batchReadSheets(['students', 'classes', 'teachers', 'parents', 'DS Lớp', 'terms', 'sessions', 'attendance']);
  mark('batchRead');

  var rawStudents = parseFromValues(raw['students']);
  var rawClasses  = parseFromValues(raw['classes']);
  var rawTeachers = parseFromValues(raw['teachers']);
  var rawParents  = parseFromValues(raw['parents']);
  var rawTermsArr = parseFromValues(raw['terms']);
  var rawLop      = raw['DS Lớp'] || [];
  mark('parse');

  // Resolve current term (cùng logic với getAllData)
  var resolvedTermId = termId || '';
  if (!resolvedTermId && rawTermsArr.length) {
    var todaySerial = new Date().getTime() / 86400000 + 25569;
    var latestPastId = '', latestPastEnd = 0;
    rawTermsArr.forEach(function(t) {
      var s = typeof t.start_date === 'number' ? t.start_date : 0;
      var e = typeof t.end_date   === 'number' ? t.end_date   : 0;
      if (s && e && todaySerial >= s && todaySerial <= e) resolvedTermId = String(t.id || '');
      else if (e && e < todaySerial && e > latestPastEnd) { latestPastEnd = e; latestPastId = String(t.id || ''); }
    });
    if (!resolvedTermId && latestPastId) resolvedTermId = latestPastId;
  }

  var teacherMap = {};
  rawTeachers.forEach(function(t) { if (t.id) teacherMap[String(t.id)] = t; });

  var students = rawStudents.filter(function(s) {
    return s.id && String(s.id).match(/^HS\d+/i);
  }).map(function(s) {
    return {
      rowIndex: s._row, id: String(s.id), name: String(s.name || ''), dob: fmtDate(s.dob),
      gender: String(s.gender || ''), photoUrl: String(s.photo_url || ''),
      status: String(s.status || 'Đang học'), address: String(s.address || ''),
      classes: [], primaryClass: '', primaryTeacher: '', totalDebt: 0
    };
  });

  var classes = rawClasses.map(function(c) {
    var teacher = teacherMap[String(c.teacher_id || '')] || {};
    return {
      id: String(c.id || ''), name: String(c.name || ''), teacherId: String(c.teacher_id || ''),
      teacher: String(teacher.name || ''), tuition: parseFloat(c.tuition_per_month) || 0,
      subject: String(c.subject || ''), status: String(c.status || ''),
      count: 0, paid: 0, debt: 0
    };
  });

  var classGroups = [];
  if (rawLop.length >= 2) {
    for (var col = 0; col < Math.min(rawLop[0].length, 2); col++) {
      var header = String(rawLop[0][col] || '').trim();
      if (!header) continue;
      var grpCls = [];
      for (var row = 1; row < rawLop.length; row++) {
        var v = rawLop[row][col];
        if (!v || v instanceof Date) continue;
        var sv = String(v).trim(); if (sv) grpCls.push(sv);
      }
      if (grpCls.length) classGroups.push({ name: header, classes: grpCls });
    }
  }

  var att7days = _computeAtt7days(parseFromValues(raw['sessions'] || []), parseFromValues(raw['attendance'] || []));
  mark('compute');

  var activeStudents = students.filter(function(s) { return s.status === 'Đang học'; }).length;
  var activeClasses  = classes.filter(function(c) { return c.status === 'Đang hoạt động'; }).length;
  var teacherCount   = rawTeachers.filter(function(t) { return String(t.status || '') === 'Đang dạy'; }).length;

  T.done = Date.now();
  return {
    students: students, payments: [], classes: classes, classGroups: classGroups,
    parents: rawParents.map(function(p) { return { id: String(p.id||''), name: String(p.name||''), phone: String(p.phone||''), email: String(p.email||''), address: String(p.address||'') }; }),
    teachers: rawTeachers.map(function(t) { return { id: String(t.id||''), name: String(t.name||''), status: String(t.status||'') }; }),
    terms: rawTermsArr.map(function(t) { return { id: String(t.id||''), school_year: String(t.school_year||''), term_name: String(t.term_name||''), start_date: fmtDate(t.start_date), end_date: fmtDate(t.end_date) }; }),
    att7days: att7days,
    stats: {
      totalStudents: students.length, activeStudents: activeStudents,
      activeClasses: activeClasses,  teacherCount: teacherCount,
      totalRevenue: 0, totalDebt: 0, totalPayments: 0,
      monthlyRevenue: [], monthlyDebt: [], classRevenue: {}
    },
    currentTermId: resolvedTermId,
    _phase: 'quick',
    _timing: { total: T.done - T.start, steps: T.steps }
  };
}

// ── Student CRUD ──────────────────────────────────────────────
function addStudent(d) {
  var sheet = getSheet('students');
  var id = genId('students', 'HS');
  sheet.appendRow([id, d.name, d.dob ? new Date(d.dob) : '', d.gender || '', '', d.status || 'Đang học', d.address || '', new Date()]);
  if (d.classId) {
    var eid = genId('enrollments', 'EN');
    getSheet('enrollments').appendRow([eid, id, d.classId, d.termId || '', new Date(), '', d.status || 'Đang học', '']);
  }
  if (d.parentId) {
    var spSheet = getSheet('student_parents');
    if (spSheet) spSheet.appendRow([id, d.parentId, 'Phụ huynh']);
  }
  return { success: true, id: id };
}

function updateStudent(d) {
  var sheet   = getSheet('students');
  var data    = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h) { return String(h).trim(); });
  var col = {};
  headers.forEach(function(h, i) { col[h] = i + 1; });
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][col['id'] - 1]) !== String(d.id)) continue;
    var row = data[i].slice();
    row[col['name']    - 1] = d.name;
    row[col['dob']     - 1] = d.dob ? new Date(d.dob) : '';
    row[col['gender']  - 1] = d.gender || '';
    row[col['status']  - 1] = d.status || '';
    row[col['address'] - 1] = d.address || '';
    sheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
    return { success: true };
  }
  return { success: false, error: 'Không tìm thấy học sinh' };
}

// ── Payment ───────────────────────────────────────────────────
function addPayment(d) {
  // d: {studentId, classId, month, amount, method, staff, note}
  var enrollSheet = getSheet('enrollments');
  var billSheet   = getSheet('monthly_bills');
  var paySheet    = getSheet('payments');
  var clsSheet    = getSheet('classes');

  // Batch-read 3 sheets in 1 call
  var rawMap     = batchReadSheets(['enrollments', 'classes', 'monthly_bills']);
  var enrollData = rawMap['enrollments'];
  var clsData    = rawMap['classes'];
  var billData   = rawMap['monthly_bills'];

  // 1. Find or create enrollment
  var eHdr = enrollData[0].map(function(h) { return String(h).trim(); });
  var eid = '';
  for (var i = 1; i < enrollData.length; i++) {
    var er = {};
    eHdr.forEach(function(h, ci) { er[h] = enrollData[i][ci]; });
    var termMatch = !d.termId || String(er.term_id || '') === String(d.termId);
    if (String(er.student_id) === String(d.studentId) && String(er.class_id) === String(d.classId) && termMatch) {
      eid = String(er.id); break;
    }
  }
  if (!eid) {
    eid = genId('enrollments', 'EN');
    enrollSheet.appendRow([eid, d.studentId, d.classId, d.termId || '', new Date(), '', 'Đang học', '']);
  }

  // 2. Find class tuition
  var cHdr    = clsData[0].map(function(h) { return String(h).trim(); });
  var tuition = 0;
  for (var k = 1; k < clsData.length; k++) {
    var cr = {};
    cHdr.forEach(function(h, ci) { cr[h] = clsData[k][ci]; });
    if (String(cr.id) === String(d.classId)) { tuition = parseFloat(cr.tuition_per_month) || 0; break; }
  }

  // 3. Find or create monthly_bill
  var bHdr     = billData[0].map(function(h) { return String(h).trim(); });
  var bid = '', billRow = -1, prevPaid = 0, amtDue = tuition;
  for (var j = 1; j < billData.length; j++) {
    var br = {};
    bHdr.forEach(function(h, ci) { br[h] = billData[j][ci]; });
    if (String(br.enrollment_id) === eid && String(br.month) === String(d.month)) {
      bid = String(br.id);
      billRow = j + 1;
      prevPaid = parseFloat(br.amount_paid) || 0;
      amtDue   = parseFloat(br.amount_due)  || tuition;
      break;
    }
  }

  var amt = parseFloat(d.amount) || 0;
  if (!bid) {
    bid = genId('monthly_bills', 'BILL');
    var debt   = Math.max(0, amtDue - amt);
    var status = debt === 0 ? 'Đã thanh toán' : (amt > 0 ? 'Thanh toán một phần' : 'Chưa thanh toán');
    billSheet.appendRow([bid, eid, d.month, amtDue, amt, debt, status, '']);
  } else {
    var newPaid = prevPaid + amt;
    var newDebt = Math.max(0, amtDue - newPaid);
    var newStatus = newDebt === 0 ? 'Đã thanh toán' : 'Thanh toán một phần';
    var pCol = bHdr.indexOf('amount_paid') + 1;
    var dCol = bHdr.indexOf('debt') + 1;
    var sCol = bHdr.indexOf('status') + 1;
    // Nếu 3 cột liền nhau (amount_paid, debt, status thường adjacent) → 1 API call
    var minC = Math.min(pCol, dCol, sCol);
    var maxC = Math.max(pCol, dCol, sCol);
    if (maxC - minC === 2) {
      var rowVals = [];
      for (var ci = minC; ci <= maxC; ci++) {
        rowVals.push(ci === pCol ? newPaid : ci === dCol ? newDebt : newStatus);
      }
      billSheet.getRange(billRow, minC, 1, 3).setValues([rowVals]);
    } else {
      billSheet.getRange(billRow, pCol).setValue(newPaid);
      billSheet.getRange(billRow, dCol).setValue(newDebt);
      billSheet.getRange(billRow, sCol).setValue(newStatus);
    }
  }

  // 4. Write payment (student_id, class_id lưu thẳng để JOIN không phụ thuộc bill chain)
  var pid = genId('payments', 'PAY');
  paySheet.appendRow([pid, bid, d.studentId, d.classId, new Date(), amt, d.method || '', d.staff || '', d.note || '']);

  return { success: true, paymentId: pid };
}

function updatePayment(d) {
  // d: {id, amount, method, staff, note}
  var paySheet  = getSheet('payments');
  var billSheet = getSheet('monthly_bills');
  var payData   = paySheet.getDataRange().getValues();
  var pH        = payData[0].map(function(h) { return String(h).trim(); });
  var pIdCol    = pH.indexOf('id');
  var pAmtCol   = pH.indexOf('amount');
  var pMetCol   = pH.indexOf('method');
  var pStaCol   = pH.indexOf('staff');
  var pNtCol    = pH.indexOf('note');
  var pBidCol   = pH.indexOf('bill_id');

  var oldAmt = 0, bidVal = '';
  for (var i = 1; i < payData.length; i++) {
    if (String(payData[i][pIdCol]) !== String(d.id)) continue;
    oldAmt = parseFloat(payData[i][pAmtCol]) || 0;
    bidVal = String(payData[i][pBidCol] || '');
    paySheet.getRange(i+1, pAmtCol+1).setValue(parseFloat(d.amount) || 0);
    paySheet.getRange(i+1, pMetCol+1).setValue(d.method || '');
    paySheet.getRange(i+1, pStaCol+1).setValue(d.staff  || '');
    paySheet.getRange(i+1, pNtCol+1).setValue(d.note   || '');
    break;
  }
  if (bidVal) _recalcBill(billSheet, bidVal);
  return { success: true };
}

function deletePayment(id) {
  var paySheet  = getSheet('payments');
  var billSheet = getSheet('monthly_bills');
  var payData   = paySheet.getDataRange().getValues();
  var pH        = payData[0].map(function(h) { return String(h).trim(); });
  var pIdCol    = pH.indexOf('id');
  var pBidCol   = pH.indexOf('bill_id');
  var bidVal    = '';
  for (var i = 1; i < payData.length; i++) {
    if (String(payData[i][pIdCol]) !== String(id)) continue;
    bidVal = String(payData[i][pBidCol] || '');
    paySheet.deleteRow(i + 1);
    break;
  }
  if (bidVal) _recalcBill(billSheet, bidVal);
  return { success: true };
}

function _recalcBill(billSheet, billId) {
  var billData = billSheet.getDataRange().getValues();
  var bH       = billData[0].map(function(h) { return String(h).trim(); });
  var bIdCol   = bH.indexOf('id');
  var bDueCol  = bH.indexOf('amount_due');
  var bPdCol   = bH.indexOf('amount_paid');
  var bDtCol   = bH.indexOf('debt');
  var bStCol   = bH.indexOf('status');
  for (var i = 1; i < billData.length; i++) {
    if (String(billData[i][bIdCol]) !== String(billId)) continue;
    var due  = parseFloat(billData[i][bDueCol]) || 0;
    // Sum all payments for this bill
    var paySheet = getSheet('payments');
    var pays     = parseSheet('payments').filter(function(p) { return String(p.bill_id||'') === String(billId); });
    var paid     = pays.reduce(function(a, p) { return a + (parseFloat(p.amount)||0); }, 0);
    var debt     = Math.max(0, due - paid);
    var status   = debt === 0 && paid > 0 ? 'Đã thanh toán' : paid > 0 ? 'Thanh toán một phần' : 'Chưa thanh toán';
    billSheet.getRange(i+1, bPdCol+1).setValue(paid);
    billSheet.getRange(i+1, bDtCol+1).setValue(debt);
    billSheet.getRange(i+1, bStCol+1).setValue(status);
    break;
  }
}

// ── Class management ─────────────────────────────────────────
function saveClass(d) {
  // d: {id?, name, teacherId, tuition, subject, status, groupName}
  var sheet = getSheet('classes');
  var data  = sheet.getDataRange().getValues();
  var hdr   = data[0].map(function(h) { return String(h).trim(); });

  // Ensure subject column exists
  if (hdr.indexOf('subject') === -1) {
    sheet.getRange(1, hdr.length + 1).setValue('subject');
    hdr.push('subject');
  }

  var idCol  = hdr.indexOf('id');
  var nmCol  = hdr.indexOf('name');
  var tcCol  = hdr.indexOf('teacher_id');
  var tuiCol = hdr.indexOf('tuition_per_month');
  var subCol = hdr.indexOf('subject');
  var stCol  = hdr.indexOf('status');

  if (d.id) {
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idCol]) !== String(d.id)) continue;
      var row = data[i].slice();
      row[nmCol]  = d.name || '';
      row[tcCol]  = d.teacherId || '';
      row[tuiCol] = parseFloat(d.tuition) || 0;
      row[subCol] = d.subject || '';
      row[stCol]  = d.status || 'Đang hoạt động';
      sheet.getRange(i+1, 1, 1, row.length).setValues([row]);
      return { success: true, id: d.id };
    }
  }
  // New class
  var id = genId('classes', 'LOP');
  var row = [];
  hdr.forEach(function(h) {
    if (h === 'id')                row.push(id);
    else if (h === 'name')         row.push(d.name || '');
    else if (h === 'teacher_id')   row.push(d.teacherId || '');
    else if (h === 'tuition_per_month') row.push(parseFloat(d.tuition) || 0);
    else if (h === 'subject')      row.push(d.subject || '');
    else if (h === 'status')       row.push(d.status || 'Đang hoạt động');
    else row.push('');
  });
  sheet.appendRow(row);
  if (d.groupName) _addClassNameToGroup(d.name, d.groupName);
  return { success: true, id: id };
}

function deleteClass(classId) {
  var sheet = getSheet('classes');
  var data  = sheet.getDataRange().getValues();
  var hdr   = data[0].map(function(h) { return String(h).trim(); });
  var idCol = hdr.indexOf('id');
  var nmCol = hdr.indexOf('name');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) !== String(classId)) continue;
    var cname = String(data[i][nmCol] || '');
    sheet.deleteRow(i + 1);
    _removeClassNameFromGroup(cname);
    return { success: true };
  }
  return { success: false };
}

function saveClassGroup(groupName, oldGroupName) {
  var sheet = getSheet('DS Lớp');
  if (!sheet) return { success: false };
  var data  = sheet.getDataRange().getValues();
  var hdr   = data[0];

  // Rename existing group
  if (oldGroupName) {
    for (var c = 0; c < hdr.length; c++) {
      if (String(hdr[c]).trim() === String(oldGroupName).trim()) {
        sheet.getRange(1, c + 1).setValue(groupName);
        return { success: true };
      }
    }
  }
  // Add new group column
  var newCol = hdr.length + 1;
  sheet.getRange(1, newCol).setValue(groupName);
  return { success: true };
}

function deleteClassGroup(groupName) {
  var sheet = getSheet('DS Lớp');
  if (!sheet) return { success: false };
  var data  = sheet.getDataRange().getValues();
  var hdr   = data[0];
  for (var c = 0; c < hdr.length; c++) {
    if (String(hdr[c]).trim() === String(groupName).trim()) {
      sheet.deleteColumn(c + 1);
      return { success: true };
    }
  }
  return { success: false };
}

function _addClassNameToGroup(className, groupName) {
  var sheet = getSheet('DS Lớp');
  if (!sheet) return;
  var data  = sheet.getDataRange().getValues();
  var hdr   = data[0];
  for (var c = 0; c < hdr.length; c++) {
    if (String(hdr[c]).trim() !== String(groupName).trim()) continue;
    // Find first empty row in this column
    for (var r = 1; r <= data.length; r++) {
      var val = r < data.length ? String(data[r][c] || '').trim() : '';
      if (!val) { sheet.getRange(r + 1, c + 1).setValue(className); return; }
    }
    sheet.getRange(data.length + 1, c + 1).setValue(className);
    return;
  }
  // Group not found — create it
  saveClassGroup(groupName, '');
  sheet.getRange(2, sheet.getLastColumn()).setValue(className);
}

function _removeClassNameFromGroup(className) {
  var sheet = getSheet('DS Lớp');
  if (!sheet) return;
  var data  = sheet.getDataRange().getValues();
  for (var r = 1; r < data.length; r++) {
    for (var c = 0; c < data[r].length; c++) {
      if (String(data[r][c] || '').trim() === String(className).trim()) {
        sheet.getRange(r + 1, c + 1).setValue('');
        return;
      }
    }
  }
}

function addStudentToClass(studentId, classId, termId) {
  var enrollSheet = getSheet('enrollments');
  var data = enrollSheet.getDataRange().getValues();
  var hdr  = data[0].map(function(h) { return String(h).trim(); });
  var sidC = hdr.indexOf('student_id');
  var cidC = hdr.indexOf('class_id');
  var tidC = hdr.indexOf('term_id');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][sidC]) === String(studentId) &&
        String(data[i][cidC]) === String(classId) &&
        (!termId || String(data[i][tidC]) === String(termId))) {
      return { success: false, error: 'Học sinh đã trong lớp' };
    }
  }
  var eid = genId('enrollments', 'EN');
  enrollSheet.appendRow([eid, studentId, classId, termId || '', new Date(), '', 'Đang học', '']);
  return { success: true, enrollId: eid };
}

function removeStudentFromClass(enrollId) {
  var sheet = getSheet('enrollments');
  var data  = sheet.getDataRange().getValues();
  var hdr   = data[0].map(function(h) { return String(h).trim(); });
  var idCol = hdr.indexOf('id');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) !== String(enrollId)) continue;
    sheet.deleteRow(i + 1);
    return { success: true };
  }
  return { success: false };
}

// ── Avatar ────────────────────────────────────────────────────
function getAvatarFolder() {
  var name = 'Panda Happy - Avatars';
  var f = DriveApp.getFoldersByName(name);
  return f.hasNext() ? f.next() : DriveApp.createFolder(name);
}

function uploadAvatar(studentId, base64Data, mimeType) {
  var sheet   = getSheet('students');
  var data    = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h) { return String(h).trim(); });
  var idCol   = headers.indexOf('id');
  var urlCol  = headers.indexOf('photo_url');
  var rowIdx  = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === String(studentId)) { rowIdx = i + 1; break; }
  }
  if (rowIdx === -1) return { success: false, error: 'Không tìm thấy học sinh' };

  var folder = getAvatarFolder();
  var ext    = mimeType.includes('png') ? '.png' : mimeType.includes('gif') ? '.gif' : '.jpg';
  var fname  = studentId + ext;
  var old    = folder.getFilesByName(fname);
  while (old.hasNext()) old.next().setTrashed(true);

  var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, fname);
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  var url = 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w120';
  sheet.getRange(rowIdx, urlCol + 1).setValue(url);
  return { success: true, url: url };
}

// ── Settings ──────────────────────────────────────────────────
function updateClassTuition(classId, tuition) {
  var sheet   = getSheet('classes');
  var data    = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h) { return String(h).trim(); });
  var idCol   = headers.indexOf('id');
  var tuiCol  = headers.indexOf('tuition_per_month');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === String(classId)) {
      sheet.getRange(i + 1, tuiCol + 1).setValue(parseFloat(tuition) || 0);
      return { success: true };
    }
  }
  return { success: false };
}

function updateEnrollmentTuition(enrollId, tuition) {
  var sheet = getSheet('enrollments');
  var data  = sheet.getDataRange().getValues();
  var hdr   = data[0].map(function(h) { return String(h).trim(); });
  var idCol  = hdr.indexOf('id');
  var tuiCol = hdr.indexOf('tuition');
  if (tuiCol === -1) {
    tuiCol = hdr.length;
    sheet.getRange(1, tuiCol + 1).setValue('tuition');
  }
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === String(enrollId)) {
      sheet.getRange(i + 1, tuiCol + 1).setValue(parseFloat(tuition) || 0);
      return { success: true };
    }
  }
  return { success: false };
}

function getUsers() {
  return parseSheet('Phân quyền').map(function(u) {
    return { email: String(u.email || ''), role: String(u.role || ''), name: String(u.display_name || ''), teacher_id: String(u.teacher_id || ''), status: String(u.status || 'active') };
  });
}

function saveUser(d) {
  var sheet   = getSheet('Phân quyền');
  var data    = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h) { return String(h).trim(); });
  var eCol    = headers.indexOf('email');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][eCol]).toLowerCase() === String(d.email).toLowerCase()) {
      sheet.getRange(i + 1, 1, 1, 6).setValues([[d.email, d.role, d.name, d.teacher_id || '', d.status || 'active', '']]);
      return { success: true };
    }
  }
  sheet.appendRow([d.email, d.role, d.name, d.teacher_id || '', d.status || 'active', '']);
  return { success: true };
}

function deleteUser(email) {
  var sheet   = getSheet('Phân quyền');
  var data    = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h) { return String(h).trim(); });
  var eCol    = headers.indexOf('email');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][eCol]).toLowerCase() === String(email).toLowerCase()) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false };
}

// ── Terms / Kỳ học ────────────────────────────────────────────
function getTerms() {
  var sheet = getSheet('terms');
  if (!sheet) return [];
  return parseSheet('terms').map(function(t) {
    return {
      id         : String(t.id         || ''),
      school_year: String(t.school_year || ''),
      term_name  : String(t.term_name   || ''),
      start_date : fmtDate(t.start_date),
      end_date   : fmtDate(t.end_date)
    };
  });
}

function addTerm(d) {
  var sheet = getSheet('terms');
  if (!sheet) return { success: false, error: 'Sheet "terms" chưa tồn tại trong Spreadsheet' };
  var id = genId('terms', 'TERM');
  sheet.appendRow([
    id, d.school_year, d.term_name,
    d.start_date ? new Date(d.start_date) : '',
    d.end_date   ? new Date(d.end_date)   : ''
  ]);
  return { success: true, id: id };
}

function deleteTerm(id) {
  var sheet = getSheet('terms');
  if (!sheet) return { success: false };
  var data  = sheet.getDataRange().getValues();
  var hdr   = data[0].map(function(h) { return String(h).trim(); });
  var idCol = hdr.indexOf('id');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false };
}

// ── Sessions & Attendance ─────────────────────────────────────
function getSessions(classId) {
  var rawMap = batchReadSheets(['sessions', 'attendance']);
  var raw = parseFromValues(rawMap['sessions']);
  if (classId) raw = raw.filter(function(s) { return String(s.class_id||'') === String(classId); });

  // Build attendance counts per session
  var att = parseFromValues(rawMap['attendance']);
  var counts = {};
  att.forEach(function(a) {
    var sid = String(a.session_id||'');
    if (!counts[sid]) counts[sid] = { present:0, absent:0, late:0 };
    var st = String(a.status||'');
    if (st === 'Có mặt') counts[sid].present++;
    else if (st === 'Vắng')    counts[sid].absent++;
    else if (st === 'Đi muộn') counts[sid].late++;
  });

  return raw.map(function(s) {
    var c = counts[String(s.id||'')] || { present:0, absent:0, late:0 };
    return { id: String(s.id||''), class_id: String(s.class_id||''), date: fmtDate(s.date), status: String(s.status||''), note: String(s.note||''), present: c.present, absent: c.absent, late: c.late };
  }).sort(function(a, b) { return a.date < b.date ? 1 : -1; });
}

function getAttendanceLast7Days() {
  var rawMap = batchReadSheets(['sessions', 'attendance']);
  return _computeAtt7days(parseFromValues(rawMap['sessions']), parseFromValues(rawMap['attendance']));
}

// Chỉ đọc, không tạo mới — dùng khi load form điểm danh
function getSessionOnly(classId, dateStr) {
  var sheet = getSheet('sessions');
  var data  = sheet.getDataRange().getValues();
  var hdr   = data[0].map(function(h) { return String(h).trim(); });
  var cidC  = hdr.indexOf('class_id');
  var dateC = hdr.indexOf('date');
  var idC   = hdr.indexOf('id');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][cidC]) !== String(classId)) continue;
    var d = data[i][dateC];
    var rowISO = d instanceof Date ? Utilities.formatDate(d, TIMEZONE, 'yyyy-MM-dd') : String(d).substring(0, 10);
    if (rowISO === dateStr) {
      var sid = String(data[i][idC]);
      return { sessionId: sid, attendance: getAttendanceForSession(sid) };
    }
  }
  return { sessionId: null, attendance: [] };
}

function getOrCreateSession(classId, dateStr) {
  var sheet = getSheet('sessions');
  var data  = sheet.getDataRange().getValues();
  var hdr   = data[0].map(function(h) { return String(h).trim(); });
  var cidC  = hdr.indexOf('class_id');
  var dateC = hdr.indexOf('date');
  var idC   = hdr.indexOf('id');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][cidC]) !== String(classId)) continue;
    var d = data[i][dateC];
    var rowISO = d instanceof Date ? Utilities.formatDate(d, TIMEZONE, 'yyyy-MM-dd') : String(d).substring(0, 10);
    if (rowISO === dateStr) {
      return String(data[i][idC]);
    }
  }
  var sid = genId('sessions', 'SES');
  sheet.appendRow([sid, classId, new Date(dateStr + 'T00:00:00'), '', '', 'Đã học', '']);
  return sid;
}

// Tạo session (nếu chưa có) rồi lưu điểm danh — gọi khi submit
function saveAttendanceForClass(classId, dateStr, records) {
  var sessionId = getOrCreateSession(classId, dateStr);
  return saveAttendance(sessionId, records);
}

function getAttendanceForSession(sessionId) {
  return parseSheet('attendance').filter(function(a) {
    return String(a.session_id||'') === String(sessionId);
  }).map(function(a) {
    return { student_id: String(a.student_id||''), status: String(a.status||''), note: String(a.note||'') };
  });
}

function saveAttendance(sessionId, records) {
  // records: [{student_id, status, note}]
  var sheet = getSheet('attendance');
  var data  = sheet.getDataRange().getValues();
  var hdr   = data[0].map(function(h) { return String(h).trim(); });
  var idC   = hdr.indexOf('id');
  var sesC  = hdr.indexOf('session_id');
  var stuC  = hdr.indexOf('student_id');
  var stC   = hdr.indexOf('status');
  var ntC   = hdr.indexOf('note');

  var maxN = 0;
  data.slice(1).forEach(function(r) { var m = String(r[idC]).match(/^ATT(\d+)/); if (m) maxN = Math.max(maxN, parseInt(m[1])); });

  var existingRow = {};
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][sesC]) === String(sessionId)) existingRow[String(data[i][stuC])] = i + 1;
  }

  var toAdd = [];
  records.forEach(function(rec) {
    var stuId = String(rec.student_id);
    if (existingRow[stuId]) {
      var rowIdx  = existingRow[stuId];
      var rowData = data[rowIdx - 1].slice();
      rowData[stC] = rec.status || '';
      rowData[ntC] = rec.note  || '';
      sheet.getRange(rowIdx, 1, 1, rowData.length).setValues([rowData]);
    } else {
      maxN++;
      toAdd.push(['ATT' + String(maxN).padStart(4,'0'), sessionId, stuId, rec.status||'', rec.note||'']);
    }
  });
  if (toAdd.length) sheet.getRange(sheet.getLastRow()+1, 1, toAdd.length, 5).setValues(toAdd);
  return { success: true };
}

function getAttendanceStats(classId, termId) {
  var rawMap     = batchReadSheets(['sessions', 'attendance', 'students', 'enrollments']);
  var sessions   = parseFromValues(rawMap['sessions']);
  var attendance = parseFromValues(rawMap['attendance']);
  var students   = parseFromValues(rawMap['students']);
  var enrollments= parseFromValues(rawMap['enrollments']);

  // Valid class IDs for filter
  var validClasses = {};
  if (classId) { validClasses[classId] = true; }
  else if (termId) {
    enrollments.filter(function(e) { return String(e.term_id||'') === String(termId); })
      .forEach(function(e) { validClasses[String(e.class_id||'')] = true; });
  }

  var validSessions = {};
  sessions.forEach(function(s) {
    var cid = String(s.class_id||'');
    if (!classId && !termId) validSessions[String(s.id||'')] = cid;
    else if (validClasses[cid]) validSessions[String(s.id||'')] = cid;
  });

  var stuMap = {};
  students.forEach(function(s) { if (s.id) stuMap[String(s.id)] = String(s.name||''); });

  var stats = {};
  attendance.forEach(function(a) {
    var sesId = String(a.session_id||'');
    if (!validSessions[sesId]) return;
    var stuId = String(a.student_id||'');
    if (!stats[stuId]) stats[stuId] = { total:0, present:0, absent:0, late:0 };
    stats[stuId].total++;
    var st = String(a.status||'');
    if (st === 'Có mặt') stats[stuId].present++;
    else if (st === 'Vắng')    stats[stuId].absent++;
    else if (st === 'Đi muộn') stats[stuId].late++;
  });

  return Object.keys(stats).map(function(stuId) {
    var s = stats[stuId];
    return { studentId: stuId, name: stuMap[stuId]||stuId, total: s.total, present: s.present, absent: s.absent, late: s.late };
  }).sort(function(a,b) { return b.absent - a.absent; });
}

// ── Teachers ──────────────────────────────────────────────────
// Sheet columns: id, name, phone, email, subject_ids, status
function getTeachers() {
  return parseSheet('teachers').map(function(t) {
    return { id: String(t.id||''), name: String(t.name||''), phone: String(t.phone||''), email: String(t.email||''), subject: String(t.subject_ids||''), status: String(t.status||'') };
  });
}

function saveTeacher(d) {
  var sheet = getSheet('teachers');
  if (d.id) {
    var data = sheet.getDataRange().getValues();
    var hdr  = data[0].map(function(h) { return String(h).trim(); });
    var idCol = hdr.indexOf('id');
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idCol]) === String(d.id)) {
        sheet.getRange(i+1, 1, 1, 6).setValues([[d.id, d.name, d.phone||'', d.email||'', d.subject||'', d.status||'Đang dạy']]);
        return { success: true };
      }
    }
  }
  var id = genId('teachers', 'GV');
  sheet.appendRow([id, d.name, d.phone||'', d.email||'', d.subject||'', d.status||'Đang dạy']);
  return { success: true, id: id };
}

function deleteTeacher(id) {
  var sheet = getSheet('teachers');
  var data  = sheet.getDataRange().getValues();
  var hdr   = data[0].map(function(h) { return String(h).trim(); });
  var idCol = hdr.indexOf('id');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === String(id)) { sheet.deleteRow(i+1); return { success: true }; }
  }
  return { success: false };
}

// ── Parents ───────────────────────────────────────────────────
function getParents() {
  return parseSheet('parents').map(function(p) {
    return { id: String(p.id||''), name: String(p.name||''), phone: String(p.phone||''), email: String(p.email||''), address: String(p.address||'') };
  });
}

function saveParent(d) {
  var sheet = getSheet('parents');
  if (d.id) {
    var data = sheet.getDataRange().getValues();
    var hdr  = data[0].map(function(h) { return String(h).trim(); });
    var idCol = hdr.indexOf('id');
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idCol]) === String(d.id)) {
        sheet.getRange(i+1, 1, 1, 5).setValues([[d.id, d.name, d.phone||'', d.email||'', d.address||'']]);
        return { success: true };
      }
    }
  }
  var id = genId('parents', 'PAR');
  sheet.appendRow([id, d.name, d.phone||'', d.email||'', d.address||'']);
  return { success: true, id: id };
}

function deleteParent(id) {
  var sheet = getSheet('parents');
  var data  = sheet.getDataRange().getValues();
  var hdr   = data[0].map(function(h) { return String(h).trim(); });
  var idCol = hdr.indexOf('id');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === String(id)) { sheet.deleteRow(i+1); return { success: true }; }
  }
  return { success: false };
}

function linkStudentParent(d) {
  var sheet = getSheet('student_parents');
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(d.studentId) && String(data[i][1]) === String(d.parentId)) return { success: true };
  }
  sheet.appendRow([d.studentId, d.parentId, d.relationship||'Phụ huynh']);
  return { success: true };
}

function unlinkStudentParent(studentId, parentId) {
  var sheet = getSheet('student_parents');
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(studentId) && String(data[i][1]) === String(parentId)) {
      sheet.deleteRow(i+1); return { success: true };
    }
  }
  return { success: false };
}

function getStudentParents(studentId) {
  var all     = parseSheet('student_parents');
  var parents = parseSheet('parents');
  var pMap    = {};
  parents.forEach(function(p) { if (p.id) pMap[String(p.id)] = p; });
  return all.filter(function(r) { return String(r.student_id||'') === String(studentId); })
    .map(function(r) {
      var p = pMap[String(r.parent_id||'')] || {};
      return { parentId: String(r.parent_id||''), name: String(p.name||''), phone: String(p.phone||''), relationship: String(r.relationship||'') };
    });
}
