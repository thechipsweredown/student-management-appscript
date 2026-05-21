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
  if (d && String(d).match(/^\d{4}-\d{2}-\d{2}/)) {
    var p = String(d).split('T')[0].split('-');
    return p[2] + '/' + p[1] + '/' + p[0];
  }
  return d ? String(d).substring(0, 10) : '';
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
function getCurrentUser() {
  var email = Session.getActiveUser().getEmail();
  var users = parseSheet('Phân quyền');
  for (var i = 0; i < users.length; i++) {
    var u = users[i];
    if (String(u['email'] || '').toLowerCase() === email.toLowerCase()
        && String(u['status']) === 'active') {
      return {
        email     : email,
        role      : String(u['role'] || 'none'),
        name      : String(u['display_name'] || ''),
        teacher_id: String(u['teacher_id'] || '')
      };
    }
  }
  return { email: email, role: 'none', name: '', teacher_id: '' };
}

// ── getAllData ────────────────────────────────────────────────
function getAllData(termId) {
  var T = { start: Date.now() };

  var rawStudents = parseSheet('students');
  var rawEnroll   = parseSheet('enrollments');
  var rawClasses  = parseSheet('classes');
  var rawTeachers = parseSheet('teachers');
  var rawBills    = parseSheet('monthly_bills');
  var rawPays     = parseSheet('payments');
  var rawLop      = getSheet('DS Lớp').getDataRange().getValues();
  T.read = Date.now();

  // Lọc enrollment theo kỳ nếu có termId
  if (termId) {
    rawEnroll = rawEnroll.filter(function(e) {
      return String(e.term_id || '') === String(termId);
    });
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

  // ── Parse students ────────────────────────────────────────
  var students = rawStudents.filter(function(s) {
    return s.id && String(s.id).match(/^HS\d+/i);
  }).map(function(s) {
    var enrolls = enrollByStudent[String(s.id)] || [];
    var classes = enrolls.map(function(e) {
      var cls     = classMap[String(e.class_id || '')] || {};
      var teacher = teacherMap[String(cls.teacher_id || '')] || {};
      var bills   = billByEnroll[String(e.id || '')] || [];
      var due  = bills.reduce(function(a, b) { return a + (parseFloat(b.amount_due)  || 0); }, 0);
      var paid = bills.reduce(function(a, b) { return a + (parseFloat(b.amount_paid) || 0); }, 0);
      var debt = bills.reduce(function(a, b) { return a + (parseFloat(b.debt)        || 0); }, 0);
      return {
        enrollId    : String(e.id || ''),
        classId     : String(e.class_id || ''),
        className   : String(cls.name || ''),
        teacherName : String(teacher.name || ''),
        tuition     : parseFloat(cls.tuition_per_month) || 0,
        enrollStatus: String(e.status || ''),
        billDue     : due,
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

  // ── Parse payments ────────────────────────────────────────
  var payments = rawPays.filter(function(p) {
    return p.date instanceof Date || (p.date && String(p.date).match(/\d/));
  }).map(function(p) {
    var bill   = billMap[String(p.bill_id || '')] || {};
    var enroll = enrollMap[String(bill.enrollment_id || '')] || {};
    var cls    = classMap[String(enroll.class_id || '')] || {};
    var stu    = studentMap[String(enroll.student_id || '')] || {};
    return {
      id         : String(p.id || ''),
      billId     : String(p.bill_id || ''),
      date       : fmtDate(p.date),
      studentId  : String(enroll.student_id || ''),
      studentName: String(stu.name  || ''),
      classId    : String(enroll.class_id || ''),
      className  : String(cls.name  || ''),
      month      : String(bill.month || ''),
      amount     : parseFloat(p.amount) || 0,
      method     : String(p.method  || ''),
      staff      : String(p.staff   || ''),
      note       : String(p.note    || '')
    };
  });
  payments.sort(function(a, b) {
    var da = a.date.split('/').reverse().join('');
    var db = b.date.split('/').reverse().join('');
    return da > db ? -1 : 1;
  });

  // ── Classes list ──────────────────────────────────────────
  var classes = rawClasses.map(function(c) {
    var teacher = teacherMap[String(c.teacher_id || '')] || {};
    var count   = Object.keys(enrollByStudent).filter(function(sid) {
      return (enrollByStudent[sid] || []).some(function(e) {
        return String(e.class_id) === String(c.id) && String(e.status) === 'Đang học';
      });
    }).length;
    return {
      id     : String(c.id || ''),
      name   : String(c.name || ''),
      teacher: String(teacher.name || ''),
      tuition: parseFloat(c.tuition_per_month) || 0,
      status : String(c.status || ''),
      count  : count
    };
  });

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
  var totalRevenue  = payments.reduce(function(a, p) { return a + p.amount; }, 0);
  var totalDebt     = students.reduce(function(a, s) { return a + s.totalDebt; }, 0);
  var activeStudents = students.filter(function(s) { return s.status === 'Đang học'; }).length;

  var monthlyMap = {};
  payments.forEach(function(p) {
    if (p.month) monthlyMap[p.month] = (monthlyMap[p.month] || 0) + p.amount;
  });
  var monthlyRevenue = Object.keys(monthlyMap).sort().map(function(m) {
    return { month: m, amount: monthlyMap[m] };
  });

  var classRevMap = {};
  payments.forEach(function(p) {
    if (p.className) classRevMap[p.className] = (classRevMap[p.className] || 0) + p.amount;
  });

  T.done = Date.now();
  return {
    students   : students,
    payments   : payments,
    classes    : classes,
    classGroups: classGroups,
    terms      : getTerms(),
    stats: {
      totalStudents : students.length,
      activeStudents: activeStudents,
      totalRevenue  : totalRevenue,
      totalDebt     : totalDebt,
      totalPayments : payments.length,
      monthlyRevenue: monthlyRevenue,
      classRevenue  : classRevMap
    },
    _timing: { total: T.done - T.start, read: T.read - T.start, compute: T.done - T.read }
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
    sheet.getRange(i + 1, col['name']).setValue(d.name);
    sheet.getRange(i + 1, col['dob']).setValue(d.dob ? new Date(d.dob) : '');
    sheet.getRange(i + 1, col['gender']).setValue(d.gender || '');
    sheet.getRange(i + 1, col['status']).setValue(d.status || '');
    sheet.getRange(i + 1, col['address']).setValue(d.address || '');
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

  // 1. Find or create enrollment
  var enrollData = enrollSheet.getDataRange().getValues();
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
  var clsData = clsSheet.getDataRange().getValues();
  var cHdr    = clsData[0].map(function(h) { return String(h).trim(); });
  var tuition = 0;
  for (var k = 1; k < clsData.length; k++) {
    var cr = {};
    cHdr.forEach(function(h, ci) { cr[h] = clsData[k][ci]; });
    if (String(cr.id) === String(d.classId)) { tuition = parseFloat(cr.tuition_per_month) || 0; break; }
  }

  // 3. Find or create monthly_bill
  var billData = billSheet.getDataRange().getValues();
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
    billSheet.getRange(billRow, pCol).setValue(newPaid);
    billSheet.getRange(billRow, dCol).setValue(newDebt);
    billSheet.getRange(billRow, sCol).setValue(newStatus);
  }

  // 4. Write payment
  var pid = genId('payments', 'PAY');
  paySheet.appendRow([pid, bid, new Date(), amt, d.method || '', d.staff || '', d.note || '']);

  return { success: true, paymentId: pid };
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
