// ============================================================
// نظام متابعة الرياضيات — Google Apps Script (موحد)
// مديرية مطروح الأزهرية
// يشمل جميع المراحل من الأولى إلى الخامسة
// ============================================================

var SHEET_NAMES = {
  INSTITUTES: 'institutes',
  CLASSES:    'classes',
  TEACHERS:   'teachers',
  VISITS:     'visits',
  SETTINGS:   'settings'
};

var DEFAULT_SETTINGS = {
  cadre_grades: [
    'معلم مساعد',
    'معلم',
    'معلم أول',
    'معلم أول (أ)',
    'معلم خبير',
    'كبير معلمين'
  ],
  grades_list: [
    'الأول الإعدادي',
    'الثاني الإعدادي',
    'الثالث الإعدادي',
    'الأول الثانوي',
    'الثاني الثانوي',
    'الثالث الثانوي'
  ]
};

var GRADE_ORDER = [
  'الأول الإعدادي',
  'الثاني الإعدادي',
  'الثالث الإعدادي',
  'الأول الثانوي',
  'الثاني الثانوي',
  'الثالث الثانوي'
];

// ============================================================
// نقطة الدخول — GET (اختبار الاتصال)
// ============================================================
function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets().map(function(s) { return s.getName(); });
  return buildResponse({ status: 'ok', message: 'نظام متابعة الرياضيات — الخادم يعمل', sheets: sheets });
}

// ============================================================
// نقطة الدخول — POST (جميع العمليات)
// ============================================================
function doPost(e) {
  var result;
  try {
    var body    = JSON.parse(e.postData.contents);
    var action  = body.action;
    var payload = body.payload || {};

    switch (action) {
      // ── تهيئة ──
      case 'initSheets':           result = initSheets();                                          break;

      // ── المعاهد ──
      case 'createInstitute':      result = createInstitute(payload);                             break;
      case 'getInstitutes':        result = getInstitutes();                                      break;
      case 'updateInstitute':      result = updateInstitute(payload);                             break;
      case 'deleteInstitute':      result = deleteInstitute(payload.id);                          break;

      // ── الفصول ──
      case 'createClasses':        result = createClasses(payload.institute_id, payload.grade, payload.count); break;
      case 'getClasses':           result = getClasses(payload.institute_id);                     break;
      case 'assignTeacher':        result = assignTeacher(payload.class_id, payload.teacher_id); break;
      case 'unassignTeacher':      result = unassignTeacher(payload.class_id);                   break;
      case 'deleteClass':          result = deleteClass(payload.id);                              break;
      case 'deleteClassesByGrade': result = deleteClassesByGrade(payload.institute_id, payload.grade); break;

      // ── المعلمون ──
      case 'createTeacher':        result = createTeacher(payload);                               break;
      case 'getTeachers':          result = getTeachers(payload.institute_id);                   break;
      case 'updateTeacher':        result = updateTeacher(payload);                               break;
      case 'deleteTeacher':        result = deleteTeacher(payload.id);                           break;
      case 'getTeacherClasses':    result = getTeacherClasses(payload.teacher_id);               break;

      // ── الإحصائيات ──
      case 'getInstituteStats':    result = getInstituteStats(payload.institute_id);             break;
      case 'getGlobalStats':       result = getGlobalStats();                                    break;
      case 'getGradeAnalysis':     result = getGradeAnalysis();                                  break;

      // ── الزيارات ──
      case 'createVisit':          result = createVisit(payload);                                break;
      case 'getVisits':            result = getVisits(payload.institute_id);                     break;
      case 'getVisitHistory':      result = getVisitHistory(payload.institute_id);               break;
      case 'getAllVisits':          result = getAllVisits();                                       break;

      // ── الإعدادات ──
      case 'getSettings':          result = getSettings();                                       break;
      case 'updateSettings':       result = updateSettings(payload);                             break;

      // ── النسخ الاحتياطي ──
      case 'importAll':            result = importAll(payload);                                  break;

      default:
        result = { status: 'error', message: 'إجراء غير معروف: ' + action };
    }
  } catch (err) {
    result = { status: 'error', message: err.toString() };
  }
  return buildResponse(result);
}

// ============================================================
// بناء الاستجابة
// ============================================================
function buildResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// تهيئة قاعدة البيانات
// ============================================================
function initSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  ensureSheet(ss, SHEET_NAMES.INSTITUTES, ['id','name','type','notes','created_at']);
  ensureSheet(ss, SHEET_NAMES.CLASSES,    ['id','institute_id','grade','class_number','teacher_id','status','created_at']);
  ensureSheet(ss, SHEET_NAMES.TEACHERS, [
    'id','name','institute_id','teacher_type','cadre_grade',
    'legal_quota','actual_hours','math_hours','deficit_hours',
    'extra_hours','supervision','notes','created_at'
  ]);
  ensureSheet(ss, SHEET_NAMES.VISITS, [
    'id','institute_id','visit_date','visitor_name',
    'sufficient_classes','needs_convoy_classes','notes','snapshot_json','created_at'
  ]);
  var settingsSheet = ensureSheet(ss, SHEET_NAMES.SETTINGS, ['key','value']);
  initSettings(settingsSheet);

  return { status: 'ok', message: 'تم تهيئة قاعدة البيانات بنجاح (5 شيتات)' };
}

function ensureSheet(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#1a5276')
      .setFontColor('#ffffff');
  }
  return sheet;
}

function initSettings(sheet) {
  var data = sheet.getDataRange().getValues();
  var keys = data.slice(1).map(function(r) { return r[0]; });

  if (keys.indexOf('cadre_grades') === -1) {
    sheet.appendRow(['cadre_grades', JSON.stringify(DEFAULT_SETTINGS.cadre_grades)]);
  }
  if (keys.indexOf('grades_list') === -1) {
    sheet.appendRow(['grades_list', JSON.stringify(DEFAULT_SETTINGS.grades_list)]);
  }
}

// ============================================================
// دالة مساعدة: تحميل كل البيانات دفعة واحدة
// ============================================================
function _loadAllData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var institutes = [], classes = [], teachers = [];

  var instSheet = ss.getSheetByName(SHEET_NAMES.INSTITUTES);
  if (instSheet) {
    var d = instSheet.getDataRange().getValues();
    for (var i = 1; i < d.length; i++) {
      if (!d[i][0]) continue;
      institutes.push({ id: d[i][0], name: d[i][1], type: d[i][2], notes: d[i][3] });
    }
  }

  var clsSheet = ss.getSheetByName(SHEET_NAMES.CLASSES);
  if (clsSheet) {
    var d2 = clsSheet.getDataRange().getValues();
    for (var i = 1; i < d2.length; i++) {
      if (!d2[i][0]) continue;
      classes.push({
        id: d2[i][0], institute_id: d2[i][1], grade: d2[i][2],
        class_number: d2[i][3], teacher_id: d2[i][4] || '', status: d2[i][5]
      });
    }
  }

  var tchSheet = ss.getSheetByName(SHEET_NAMES.TEACHERS);
  if (tchSheet) {
    var d3 = tchSheet.getDataRange().getValues();
    for (var i = 1; i < d3.length; i++) {
      if (!d3[i][0]) continue;
      var lq = Number(d3[i][5]) || 0;
      var ah = Number(d3[i][6]) || 0;
      teachers.push({
        id: d3[i][0], name: d3[i][1], institute_id: d3[i][2],
        teacher_type: d3[i][3], cadre_grade: d3[i][4],
        legal_quota: lq, actual_hours: ah,
        math_hours: Number(d3[i][7]) || 0,
        deficit_hours: Number(d3[i][8]) || 0,
        extra_hours: Math.max(0, ah - lq),
        supervision: Number(d3[i][10]) || 0,
        notes: d3[i][11] || ''
      });
    }
  }

  return { institutes: institutes, classes: classes, teachers: teachers };
}

// ============================================================
// المعاهد
// ============================================================
function createInstitute(payload) {
  if (!payload.id || !payload.name || !payload.type) {
    return { status: 'error', message: 'البيانات المطلوبة ناقصة (id, name, type)' };
  }
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.INSTITUTES);
  if (!sheet) return { status: 'error', message: 'الشيت غير موجود — نفّذ تهيئة قاعدة البيانات أولاً' };

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === payload.id) return { status: 'error', message: 'المعهد موجود مسبقًا' };
  }

  sheet.appendRow([payload.id, payload.name, payload.type, payload.notes || '', new Date().toISOString()]);
  return { status: 'ok', message: 'تم إضافة المعهد بنجاح', id: payload.id };
}

// يُرجع مصفوفة مباشرة (بدون غلاف) متوافقةً مع phase5
function getInstitutes() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.INSTITUTES);
  if (!sheet) return [];

  var data    = sheet.getDataRange().getValues();
  var headers = data[0];
  return data.slice(1)
    .filter(function(r) { return r[0] !== ''; })
    .map(function(row) {
      var obj = {};
      headers.forEach(function(h, i) { obj[h] = row[i]; });
      return obj;
    });
}

function updateInstitute(payload) {
  if (!payload.id) return { status: 'error', message: 'معرّف المعهد مطلوب' };
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.INSTITUTES);
  if (!sheet) return { status: 'error', message: 'الشيت غير موجود' };

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === payload.id) {
      if (payload.name  !== undefined) sheet.getRange(i + 1, 2).setValue(payload.name);
      if (payload.type  !== undefined) sheet.getRange(i + 1, 3).setValue(payload.type);
      if (payload.notes !== undefined) sheet.getRange(i + 1, 4).setValue(payload.notes);
      return { status: 'ok', message: 'تم تحديث المعهد بنجاح' };
    }
  }
  return { status: 'error', message: 'المعهد غير موجود' };
}

function deleteInstitute(id) {
  if (!id) return { status: 'error', message: 'معرّف المعهد مطلوب' };
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.INSTITUTES);
  if (!sheet) return { status: 'error', message: 'الشيت غير موجود' };

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.deleteRow(i + 1);
      return { status: 'ok', message: 'تم حذف المعهد بنجاح' };
    }
  }
  return { status: 'error', message: 'المعهد غير موجود' };
}

// ============================================================
// الفصول
// ============================================================
function createClasses(institute_id, grade, count) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.CLASSES);
  if (!sheet) return { status: 'error', message: 'شيت classes غير موجود' };

  var countNum = parseInt(count, 10);
  if (isNaN(countNum) || countNum < 1) return { status: 'error', message: 'عدد الفصول غير صحيح' };

  // أكبر رقم فصل موجود لهذا الصف في هذا المعهد
  var data = sheet.getDataRange().getValues();
  var existingNums = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === institute_id && data[i][2] === grade) {
      existingNums.push(parseInt(data[i][3], 10));
    }
  }
  var lastNum = existingNums.length > 0 ? Math.max.apply(null, existingNums) : 0;

  var created = [];
  for (var n = 1; n <= countNum; n++) {
    var newId    = Utilities.getUuid();
    var classNum = lastNum + n;
    sheet.appendRow([newId, institute_id, grade, classNum, '', 'يحتاج قوافل', new Date().toISOString()]);
    created.push({ id: newId, institute_id: institute_id, grade: grade, class_number: classNum, teacher_id: '', status: 'يحتاج قوافل' });
  }
  return created; // مصفوفة مباشرة
}

// يُرجع مصفوفة مباشرة
function getClasses(institute_id) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.CLASSES);
  if (!sheet) return [];

  var data   = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    if (institute_id && data[i][1] !== institute_id) continue;
    result.push({
      id: data[i][0], institute_id: data[i][1], grade: data[i][2],
      class_number: data[i][3], teacher_id: data[i][4] || '', status: data[i][5]
    });
  }
  return result;
}

function assignTeacher(class_id, teacher_id) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.CLASSES);
  if (!sheet) return { status: 'error', message: 'شيت classes غير موجود' };

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === class_id) {
      var newStatus = teacher_id ? 'مكتفي' : 'يحتاج قوافل';
      sheet.getRange(i + 1, 5).setValue(teacher_id || '');
      sheet.getRange(i + 1, 6).setValue(newStatus);
      return { status: 'ok', newStatus: newStatus };
    }
  }
  return { status: 'error', message: 'الفصل غير موجود' };
}

function unassignTeacher(class_id) {
  return assignTeacher(class_id, '');
}

function deleteClass(id) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.CLASSES);
  if (!sheet) return { status: 'error', message: 'شيت classes غير موجود' };

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.deleteRow(i + 1);
      return { status: 'ok', message: 'تم حذف الفصل' };
    }
  }
  return { status: 'error', message: 'الفصل غير موجود' };
}

function deleteClassesByGrade(institute_id, grade) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.CLASSES);
  if (!sheet) return { status: 'error', message: 'شيت classes غير موجود' };

  var data         = sheet.getDataRange().getValues();
  var rowsToDelete = [];
  for (var i = data.length - 1; i >= 1; i--) {
    if (data[i][1] === institute_id && data[i][2] === grade) rowsToDelete.push(i + 1);
  }
  for (var r = 0; r < rowsToDelete.length; r++) sheet.deleteRow(rowsToDelete[r]);
  return { status: 'ok', deleted: rowsToDelete.length };
}

// ============================================================
// المعلمون
// ============================================================
function createTeacher(payload) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.TEACHERS);
  if (!sheet) return { status: 'error', message: 'شيت teachers غير موجود' };

  var id         = payload.id || Utilities.getUuid();
  var lq         = Number(payload.legal_quota)  || 0;
  var ah         = Number(payload.actual_hours) || 0;
  var extraHours = Math.max(0, ah - lq);
  var defHours   = Math.max(0, lq - ah);

  sheet.appendRow([
    id,
    payload.name          || '',
    payload.institute_id  || '',
    payload.teacher_type  || '',
    payload.cadre_grade   || '',
    lq,
    ah,
    Number(payload.math_hours)   || 0,
    defHours,
    extraHours,
    Number(payload.supervision)  || 0,
    payload.notes         || '',
    new Date().toISOString()
  ]);
  return { status: 'ok', id: id };
}

// يُرجع مصفوفة مباشرة
function getTeachers(institute_id) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.TEACHERS);
  if (!sheet) return [];

  var data   = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    if (institute_id && data[i][2] !== institute_id) continue;
    var lq = Number(data[i][5]) || 0;
    var ah = Number(data[i][6]) || 0;
    result.push({
      id:           data[i][0],
      name:         data[i][1],
      institute_id: data[i][2],
      teacher_type: data[i][3],
      cadre_grade:  data[i][4],
      legal_quota:  lq,
      actual_hours: ah,
      math_hours:   Number(data[i][7])  || 0,
      deficit_hours: Number(data[i][8]) || 0,
      extra_hours:   Math.max(0, ah - lq),
      supervision:   Number(data[i][10]) || 0,
      notes:         data[i][11] || ''
    });
  }
  return result;
}

function updateTeacher(payload) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.TEACHERS);
  if (!sheet) return { status: 'error', message: 'شيت teachers غير موجود' };

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === payload.id) {
      var lq = Number(payload.legal_quota)  || 0;
      var ah = Number(payload.actual_hours) || 0;
      sheet.getRange(i + 1, 2).setValue(payload.name          || '');
      sheet.getRange(i + 1, 3).setValue(payload.institute_id  || '');
      sheet.getRange(i + 1, 4).setValue(payload.teacher_type  || '');
      sheet.getRange(i + 1, 5).setValue(payload.cadre_grade   || '');
      sheet.getRange(i + 1, 6).setValue(lq);
      sheet.getRange(i + 1, 7).setValue(ah);
      sheet.getRange(i + 1, 8).setValue(Number(payload.math_hours)  || 0);
      sheet.getRange(i + 1, 9).setValue(Math.max(0, lq - ah));   // deficit_hours
      sheet.getRange(i + 1, 10).setValue(Math.max(0, ah - lq));  // extra_hours
      sheet.getRange(i + 1, 11).setValue(Number(payload.supervision) || 0);
      sheet.getRange(i + 1, 12).setValue(payload.notes || '');
      return { status: 'ok', message: 'تم تحديث المعلم بنجاح' };
    }
  }
  return { status: 'error', message: 'المعلم غير موجود' };
}

function deleteTeacher(id) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.TEACHERS);
  if (!sheet) return { status: 'error', message: 'شيت teachers غير موجود' };

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.deleteRow(i + 1);
      // إزالة هذا المعلم من أي فصول مرتبطة به
      var clsSheet = ss.getSheetByName(SHEET_NAMES.CLASSES);
      if (clsSheet) {
        var clsData = clsSheet.getDataRange().getValues();
        for (var j = 1; j < clsData.length; j++) {
          if (clsData[j][4] === id) {
            clsSheet.getRange(j + 1, 5).setValue('');
            clsSheet.getRange(j + 1, 6).setValue('يحتاج قوافل');
          }
        }
      }
      return { status: 'ok', message: 'تم حذف المعلم بنجاح' };
    }
  }
  return { status: 'error', message: 'المعلم غير موجود' };
}

function getTeacherClasses(teacher_id) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.CLASSES);
  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][4] === teacher_id) {
      result.push({
        id: data[i][0], institute_id: data[i][1], grade: data[i][2],
        class_number: data[i][3], teacher_id: data[i][4], status: data[i][5]
      });
    }
  }
  return result;
}

// ============================================================
// الإحصائيات
// ============================================================

// إحصائيات معهد واحد
function getInstituteStats(institute_id) {
  var all      = _loadAllData();
  var classes  = all.classes.filter(function(c)  { return c.institute_id === institute_id; });
  var teachers = all.teachers.filter(function(t) { return t.institute_id === institute_id; });

  var totalClasses  = classes.length;
  var sufficient    = classes.filter(function(c) { return c.teacher_id; }).length;
  var needsConvoy   = totalClasses - sufficient;
  var pctSufficient = totalClasses > 0 ? Math.round(sufficient / totalClasses * 100) : 0;

  var mathTeachers    = teachers.filter(function(t) { return t.teacher_type === 'معلم رياضيات'; });
  var deficitTeachers = teachers.filter(function(t) { return t.teacher_type === 'معلم يسد عجز'; });

  var totalExtra = 0, totalDeficit = 0, totalSupervision = 0;
  teachers.forEach(function(t) {
    totalExtra      += Math.max(0, t.actual_hours - t.legal_quota);
    totalDeficit    += Math.max(0, t.legal_quota - t.actual_hours);
    totalSupervision += t.supervision;
  });

  // الفصول المحتاجة للقوافل مرتبة
  var convoyClasses = classes.filter(function(c) { return !c.teacher_id; });
  convoyClasses.sort(function(a, b) {
    return GRADE_ORDER.indexOf(a.grade) - GRADE_ORDER.indexOf(b.grade) || a.class_number - b.class_number;
  });

  // تحليل بالصف
  var gradeStats = {};
  GRADE_ORDER.forEach(function(g) { gradeStats[g] = { total: 0, sufficient: 0, needs_convoy: 0 }; });
  classes.forEach(function(c) {
    if (!gradeStats[c.grade]) gradeStats[c.grade] = { total: 0, sufficient: 0, needs_convoy: 0 };
    gradeStats[c.grade].total++;
    if (c.teacher_id) gradeStats[c.grade].sufficient++;
    else gradeStats[c.grade].needs_convoy++;
  });

  return {
    status: 'ok',
    institute_id: institute_id,
    total_classes: totalClasses,
    sufficient_classes: sufficient,
    needs_convoy_classes: needsConvoy,
    pct_sufficient: pctSufficient,
    math_teachers: mathTeachers.length,
    deficit_teachers: deficitTeachers.length,
    total_teachers: teachers.length,
    total_extra_hours: totalExtra,
    total_deficit_hours: totalDeficit,
    total_supervision: totalSupervision,
    convoy_classes: convoyClasses,
    grade_stats: gradeStats,
    teachers: teachers,
    classes: classes
  };
}

// إحصائيات جميع المعاهد — حقول الأسماء متوافقة مع phase5.html
function getGlobalStats() {
  var all        = _loadAllData();
  var institutes = all.institutes;
  var classes    = all.classes;
  var teachers   = all.teachers;

  var totalInstitutes   = institutes.length;
  var totalClasses      = classes.length;
  var sufficientClasses = classes.filter(function(c) { return c.teacher_id; }).length;
  var needsConvoy       = totalClasses - sufficientClasses;

  var totalExtra = 0, totalDeficit = 0, totalSupervision = 0;
  teachers.forEach(function(t) {
    totalExtra      += Math.max(0, t.actual_hours - t.legal_quota);
    totalDeficit    += Math.max(0, t.legal_quota  - t.actual_hours);
    totalSupervision += t.supervision;
  });

  // ملخص لكل معهد
  var instituteRows = institutes.map(function(inst) {
    var iClasses  = classes.filter(function(c)  { return c.institute_id === inst.id; });
    var iTeachers = teachers.filter(function(t) { return t.institute_id === inst.id; });
    var iSuff     = iClasses.filter(function(c) { return c.teacher_id; }).length;
    var iNeeds    = iClasses.length - iSuff;
    var iPct      = iClasses.length > 0 ? Math.round(iSuff / iClasses.length * 100) : 0;
    return {
      id:           inst.id,
      name:         inst.name,
      type:         inst.type,
      teachers:     iTeachers.length,
      total:        iClasses.length,
      sufficient:   iSuff,
      needs_convoy: iNeeds,
      pct:          iPct
    };
  });

  return {
    total_institutes:    totalInstitutes,
    total_classes:       totalClasses,
    sufficient_classes:  sufficientClasses,
    needs_convoy_classes: needsConvoy,
    total_teachers:      teachers.length,
    total_extra_hours:   totalExtra,
    total_deficit_hours: totalDeficit,
    total_supervision:   totalSupervision,
    institutes:          instituteRows
  };
}

// تحليل الفصول حسب الصف — حقول الأسماء متوافقة مع phase5.html
function getGradeAnalysis() {
  var all     = _loadAllData();
  var classes = all.classes;

  return GRADE_ORDER.map(function(grade) {
    var gc    = classes.filter(function(c) { return c.grade === grade; });
    var suff  = gc.filter(function(c) { return c.teacher_id; }).length;
    var needs = gc.length - suff;
    var pct   = gc.length > 0 ? Math.round(suff / gc.length * 100) : 0;
    return {
      grade:         grade,
      stage:         grade.indexOf('الإعدادي') !== -1 ? 'إعدادي' : 'ثانوي',
      total:         gc.length,
      sufficient:    suff,
      needs_convoy:  needs,
      pct_sufficient: pct
    };
  });
}

// ============================================================
// الزيارات
// ============================================================
function createVisit(payload) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.VISITS);
  if (!sheet) return { status: 'error', message: 'شيت visits غير موجود' };

  var id = payload.id || Utilities.getUuid();
  var snapshotStr = '';
  if (payload.snapshot_json) {
    snapshotStr = typeof payload.snapshot_json === 'string'
      ? payload.snapshot_json
      : JSON.stringify(payload.snapshot_json);
  }

  sheet.appendRow([
    id,
    payload.institute_id         || '',
    payload.visit_date           || '',
    payload.visitor_name         || '',
    payload.sufficient_classes   || 0,
    payload.needs_convoy_classes || 0,
    payload.notes                || '',
    snapshotStr,
    new Date().toISOString()
  ]);
  return { status: 'ok', id: id };
}

// يُرجع مصفوفة مباشرة
function getVisits(institute_id) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.VISITS);
  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    if (institute_id && data[i][1] !== institute_id) continue;
    var snapshot = null;
    try { snapshot = data[i][7] ? JSON.parse(data[i][7]) : null; } catch(e) {}
    result.push({
      id:                   data[i][0],
      institute_id:         data[i][1],
      visit_date:           data[i][2],
      visitor_name:         data[i][3],
      sufficient_classes:   Number(data[i][4]) || 0,
      needs_convoy_classes: Number(data[i][5]) || 0,
      notes:                data[i][6] || '',
      snapshot_json:        snapshot ? JSON.stringify(snapshot) : ''
    });
  }
  return result;
}

function getVisitHistory(institute_id) {
  var visits = getVisits(institute_id);
  visits.sort(function(a, b) { return new Date(a.visit_date) - new Date(b.visit_date); });

  return visits.map(function(v, idx) {
    var prev = idx > 0 ? visits[idx - 1] : null;
    return {
      id:                   v.id,
      visit_date:           v.visit_date,
      visitor_name:         v.visitor_name,
      sufficient_classes:   v.sufficient_classes,
      needs_convoy_classes: v.needs_convoy_classes,
      notes:                v.notes,
      diff_sufficient:      prev ? (v.sufficient_classes    - prev.sufficient_classes)    : null,
      diff_needs_convoy:    prev ? (v.needs_convoy_classes  - prev.needs_convoy_classes)  : null
    };
  });
}

function getAllVisits() {
  return getVisits(null);
}

// ============================================================
// الإعدادات
// ============================================================
function getSettings() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  if (!sheet) return {};

  var data     = sheet.getDataRange().getValues();
  var settings = {};
  data.slice(1).forEach(function(row) {
    if (row[0]) {
      try { settings[row[0]] = JSON.parse(row[1]); }
      catch(e) { settings[row[0]] = row[1]; }
    }
  });
  return settings;
}

function updateSettings(payload) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  if (!sheet) return { status: 'error', message: 'شيت settings غير موجود' };

  var data = sheet.getDataRange().getValues();
  var key   = payload.key;
  var value = payload.value;

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(typeof value === 'string' ? value : JSON.stringify(value));
      return { status: 'ok', message: 'تم حفظ الإعداد' };
    }
  }
  // إضافة مفتاح جديد
  sheet.appendRow([key, typeof value === 'string' ? value : JSON.stringify(value)]);
  return { status: 'ok', message: 'تم إضافة الإعداد' };
}

// ============================================================
// النسخ الاحتياطي — استيراد كامل
// ============================================================
function importAll(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // حذف البيانات الموجودة وإعادة الكتابة
  var sheetsMap = {
    institutes: { name: SHEET_NAMES.INSTITUTES, headers: ['id','name','type','notes','created_at'] },
    classes:    { name: SHEET_NAMES.CLASSES,    headers: ['id','institute_id','grade','class_number','teacher_id','status','created_at'] },
    teachers:   { name: SHEET_NAMES.TEACHERS,   headers: ['id','name','institute_id','teacher_type','cadre_grade','legal_quota','actual_hours','math_hours','deficit_hours','extra_hours','supervision','notes','created_at'] },
    visits:     { name: SHEET_NAMES.VISITS,     headers: ['id','institute_id','visit_date','visitor_name','sufficient_classes','needs_convoy_classes','notes','snapshot_json','created_at'] }
  };

  var imported = 0;

  Object.keys(sheetsMap).forEach(function(key) {
    if (!payload[key]) return;
    var cfg   = sheetsMap[key];
    var sheet = ss.getSheetByName(cfg.name);
    if (!sheet) {
      sheet = ss.insertSheet(cfg.name);
      sheet.appendRow(cfg.headers);
      sheet.setFrozenRows(1);
    } else {
      // احتفظ بسطر الرؤوس
      if (sheet.getLastRow() > 1) sheet.deleteRows(2, sheet.getLastRow() - 1);
    }

    var rows = payload[key];
    rows.forEach(function(row) {
      var rowData = cfg.headers.map(function(h) {
        var val = row[h];
        if (val === undefined || val === null) return '';
        if (typeof val === 'object') return JSON.stringify(val);
        return val;
      });
      sheet.appendRow(rowData);
      imported++;
    });
  });

  return { status: 'ok', message: 'تم الاستيراد بنجاح. سجلات مُستوردة: ' + imported };
}
