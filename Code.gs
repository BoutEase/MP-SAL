// ==================================================
// MP-SAL Payroll Backend — Google Apps Script
// Mahitha Prasad Fashion Label
// ==================================================

function getSpreadsheetId() {
  return PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
}

function getSheet(name) {
  return SpreadsheetApp.openById(getSpreadsheetId()).getSheetByName(name);
}

function jsonOut(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---- Entry Points ----

function doGet(e) {
  return jsonOut({ status: 'ok', service: 'MP-SAL Payroll API v1' });
}

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    const { action, data, pin, role } = req;

    if (action !== 'verifyPin' && action !== 'initSheets') {
      if (!checkPin(pin, role)) {
        return jsonOut({ success: false, error: 'Unauthorized' });
      }
    }

    const handlers = {
      verifyPin:       () => verifyPin(data),
      initSheets:      () => initSheets(),
      getEmployees:    () => getEmployees(),
      saveEmployee:    () => saveEmployee(data),
      saveEmployees:   () => saveEmployees(data.records),
      deleteEmployee:  () => deleteEmployee(data.emp_id),
      getHolidays:     () => getHolidays(data && data.year),
      saveHolidays:    () => saveHolidays(data),
      getAdvances:     () => getAdvances(data),
      saveAdvance:     () => saveAdvance(data),
      approveAdvance:  () => approveAdvance(data.advance_id),
      fixAdvanceEmpId: () => fixAdvanceEmpId(data),
      approveAdvances: () => approveAdvances(data.advance_ids),
      deleteAdvance:   () => deleteAdvance(data.advance_id),
      getPayroll:      () => getPayroll(data && data.month),
      savePayroll:          () => savePayroll(data),
      patchAttendanceJson:  () => patchAttendanceJson(data),
      reopenPayrollRecord:  () => reopenPayrollRecord(data),
      finalizePayroll:      () => finalizePayroll(data),
      getPayments:     () => getPayments(data && data.month),
      savePayment:     () => savePayment(data),
      deletePayment:   () => deletePayment(data.payment_id),
      getBonusPool:         () => getBonusPool(),
      confirmBonusPaid:     () => confirmBonusPaid(data),
      saveBonusBalances:    () => saveBonusBalances(data),
      getFinalizedMonths:   () => getFinalizedMonths(),
      getPaymentRun:        () => getPaymentRun(data.month),
      saveBulkPayments:     () => saveBulkPayments(data.records),
      getSettings:          () => getSettings(),
      saveSettings:         () => saveSettings(data),
      generateLoginCodes:   () => generateLoginCodes(),
      getEmployeePayslip:   () => getEmployeePayslip(data),
      getEmployeeAdvanceHistory: () => getEmployeeAdvanceHistory(data),
    };

    const handler = handlers[action];
    if (!handler) return jsonOut({ success: false, error: 'Unknown action: ' + action });
    return jsonOut(handler());

  } catch (err) {
    return jsonOut({ success: false, error: err.toString() });
  }
}

// ---- Auth ----

function checkPin(pin, role) {
  const s = getSettingsMap();
  if (role === 'admin')   return String(pin) === String(s.ADMIN_PIN);
  if (role === 'manager') return String(pin) === String(s.MANAGER_PIN);
  if (role === 'employee') {
    var sheet = getSheet('Employees');
    var rows = sheet.getDataRange().getValues().slice(1);
    return rows.some(function(r) { return r[8] === 'Active' && String(r[7]) === String(pin); });
  }
  return false;
}

function verifyPin(data) {
  if (data.role === 'employee') {
    var sheet = getSheet('Employees');
    var rows = sheet.getDataRange().getValues().slice(1);
    var emp = rows.find(function(r) { return r[8] === 'Active' && String(r[7]) === String(data.pin); });
    if (!emp) return { success: false };
    return { success: true, role: 'employee', emp_id: emp[0], emp_name: emp[1] };
  }
  return { success: checkPin(data.pin, data.role), role: data.role };
}

function getEmployeePayslip(data) {
  var empId = data.emp_id;
  var month = data.month;

  // 'latest' → find the most recent finalized month for this employee
  if (month === 'latest') {
    var allRows = getPayroll().data
      .filter(function(p) { return p.emp_id === empId && p.status === 'finalized'; })
      .sort(function(a, b) { return b.month > a.month ? 1 : -1; });
    month = allRows.length ? allRows[0].month : null;
  }

  var payRows = month ? getPayroll(month).data.filter(function(p) { return p.emp_id === empId; }) : [];
  var pay = payRows[0] || null;
  var advances = [];
  if (pay && pay.adv_start_date && pay.adv_end_date) {
    advances = getAdvances({
      emp_id: empId,
      status: 'Approved',
      start_date: pay.adv_start_date,
      end_date: pay.adv_end_date
    }).data;
  }

  // Get latest bonus balance from Bonus sheet
  var bonusBalance = 0;
  try {
    var bonusSheet = getSheet('Bonus');
    var bonusRows = bonusSheet.getDataRange().getValues().slice(1)
      .filter(function(r) { return r[0] && String(r[1]) === String(empId); });
    if (bonusRows.length) {
      bonusRows.sort(function(a, b) { return String(a[3]) > String(b[3]) ? 1 : -1; });
      bonusBalance = Number(bonusRows[bonusRows.length - 1][7]) || 0;
    }
  } catch(e) {}

  return { success: true, data: { payroll: pay, advances: advances, bonus_balance: bonusBalance, month: month } };
}

function getEmployeeAdvanceHistory(data) {
  var empId = data.emp_id;
  // Only finalized payslips — each defines a settled advance period
  var finalized = getPayroll().data
    .filter(function(p) { return p.emp_id === empId && p.status === 'finalized' && p.adv_start_date && p.adv_end_date; })
    .sort(function(a, b) { return b.month > a.month ? 1 : -1; });

  var periods = finalized.map(function(p) {
    var advs = getAdvances({
      emp_id: empId,
      status: 'Approved',
      start_date: p.adv_start_date,
      end_date: p.adv_end_date
    }).data.sort(function(a, b) { return a.date > b.date ? 1 : -1; });
    var total = advs.reduce(function(s, a) { return s + Number(a.amount); }, 0);
    return {
      month: p.month,
      adv_start_date: p.adv_start_date,
      adv_end_date: p.adv_end_date,
      advances: advs,
      total: total,
      total_from_payroll: Number(p.total_advances) || 0
    };
  });

  return { success: true, data: periods };
}

// ---- Settings ----

function getDefaultSettings() {
  return {
    ADMIN_PIN: '182612',
    MANAGER_PIN: '123456',
    SHORTFALL_THRESHOLD_HOURS: '8.5',
    ADD_LUNCH_TO_OT: 'true',
    LUNCH_DURATION_HOURS: '1',
    COMPANY_NAME: 'Mahitha Prasad',
    BONUS_POOL_TARGET: '24000',
    EMPLOYEE_CONTRIBUTION: '500',
    COMPANY_CONTRIBUTION: '500',
    BONUS_ELIGIBILITY_MIN_DAYS: '6',
    OT_MIN_HOURS: '10',
    OT_ROUND_MINUTES: '15',
    TEAMS: 'Emb,Stitching'
  };
}

function getSettingsMap() {
  try {
    const sheet = getSheet('Settings');
    if (!sheet) return getDefaultSettings();
    const rows = sheet.getDataRange().getValues();
    const map = {};
    rows.forEach(function(r) { if (r[0]) map[String(r[0])] = String(r[1]); });
    return Object.assign({}, getDefaultSettings(), map);
  } catch(e) {
    return getDefaultSettings();
  }
}

function getSettings() {
  return { success: true, data: getSettingsMap() };
}

function saveSettings(data) {
  const sheet = getSheet('Settings');
  const rows = sheet.getDataRange().getValues();
  const keyIndex = {};
  rows.forEach(function(r, i) { if (r[0]) keyIndex[r[0]] = i + 1; });

  Object.keys(data).forEach(function(k) {
    if (keyIndex[k]) {
      sheet.getRange(keyIndex[k], 2).setValue(data[k]);
    } else {
      sheet.appendRow([k, data[k]]);
    }
  });
  return { success: true };
}

function generateLoginCodes() {
  var sheet = getSheet('Employees');
  var rows = sheet.getDataRange().getValues();
  var used = new Set();
  var updates = [];

  // Collect how many active employees need codes
  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    var code;
    do {
      code = String(Math.floor(100000 + Math.random() * 900000));
    } while (used.has(code));
    used.add(code);
    updates.push({ row: i + 1, code: code, name: rows[i][1] });
    sheet.getRange(i + 1, 8).setValue(code);
  }

  return { success: true, updated: updates.length, codes: updates.map(function(u) { return { name: u.name, code: u.code }; }) };
}

// ---- Sheet Init ----

function initSheets() {
  const id = getSpreadsheetId();
  if (!id) return { success: false, error: 'SPREADSHEET_ID not set in Script Properties. Go to Project Settings > Script Properties and add SPREADSHEET_ID.' };

  var ss;
  try {
    ss = SpreadsheetApp.openById(id);
  } catch(e) {
    return { success: false, error: 'Cannot open spreadsheet: ' + e.toString() };
  }

  const defs = {
    Employees: ['emp_id','name','team','designation','petpooja_id','petpooja_name','weekly_salary','login_code','status','joining_date','company_room','bonus_scheme','payment_batch'],
    Holidays:  ['date','name'],
    Payroll:   ['payroll_id','emp_id','month','full_days','half_days','absent_days','week_off_days','holiday_absent_days','ot_weekday_min','ot_sunday_min','ot_holiday_min','shortfall_min','weekly_salary','daily_rate','hourly_rate','TD','gross_pay','ot_earnings','shortfall_deduction','bonus_eligible','bonus_cut','total_advances','net_pay','status','finalized_date','notes','adv_start_date','adv_end_date','attendance_json'],
    Advances:  ['advance_id','emp_id','emp_name','date','amount','status','created_by','created_at'],
    Bonus:     ['bonus_id','emp_id','emp_name','month','eligible','contribution','payout','balance_after','notes'],
    Payments:  ['payment_id','emp_id','emp_name','month','date_paid','amount','mode','notes'],
    Settings:  ['key','value']
  };

  var created = [];
  Object.keys(defs).forEach(function(name) {
    var s = ss.getSheetByName(name);
    if (!s) {
      s = ss.insertSheet(name);
      s.appendRow(defs[name]);
      s.getRange(1, 1, 1, defs[name].length)
        .setFontWeight('bold')
        .setBackground('#1e3a5f')
        .setFontColor('#ffffff');
      s.setFrozenRows(1);
      created.push(name);
    }
  });

  const settingsSheet = ss.getSheetByName('Settings');
  if (settingsSheet.getLastRow() <= 1) {
    const defaults = getDefaultSettings();
    Object.keys(defaults).forEach(function(k) {
      settingsSheet.appendRow([k, defaults[k]]);
    });
  }

  return { success: true, created: created, message: 'Sheets ready.' };
}

// ---- Employees ----

function getEmployees() {
  const sheet = getSheet('Employees');
  const rows = sheet.getDataRange().getValues().slice(1);
  return {
    success: true,
    data: rows.filter(function(r) { return r[0]; }).map(function(r) {
      return {
        emp_id: r[0], name: r[1], team: r[2], designation: r[3],
        petpooja_id: String(r[4]), petpooja_name: r[5],
        weekly_salary: Number(r[6]), login_code: r[7],
        status: r[8], joining_date: r[9], company_room: r[10] || '',
        bonus_scheme: r[11] || 'No',
        payment_batch: r[12] || 'Direct'
      };
    })
  };
}

function saveEmployee(data) {
  const sheet = getSheet('Employees');
  const rows = sheet.getDataRange().getValues();

  if (!data.emp_id) {
    if (data.petpooja_id && parseInt(data.petpooja_id)) {
      // Derive emp_id from Petpooja ID so it stays consistent with advance records
      data.emp_id = 'MP' + String(parseInt(data.petpooja_id)).padStart(3, '0');
    } else {
      const nums = rows.slice(1).filter(function(r) { return r[0]; })
        .map(function(r) { return parseInt(String(r[0]).replace(/\D/g, '')) || 0; });
      const next = (Math.max.apply(null, [0].concat(nums)) + 1).toString().padStart(3, '0');
      data.emp_id = 'MP' + next;
    }
  }

  const row = [
    data.emp_id, data.name, data.team, data.designation,
    data.petpooja_id || '', data.petpooja_name || data.name,
    data.weekly_salary, data.login_code || '', data.status || 'Active',
    data.joining_date || Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd'),
    data.company_room || '',
    data.bonus_scheme || 'No',
    data.payment_batch || 'Direct'
  ];

  // First try exact emp_id match
  var idx = rows.findIndex(function(r, i) { return i > 0 && r[0] === data.emp_id; });

  if (idx > 0) {
    sheet.getRange(idx + 1, 1, 1, row.length).setValues([row]);
  } else {
    // Fall back to name match — handles migration where old emp_id was auto-sequential
    var normName = String(data.name || '').toLowerCase().trim();
    var nameIdx  = rows.findIndex(function(r, i) {
      return i > 0 && r[0] && String(r[1] || '').toLowerCase().trim() === normName;
    });
    if (nameIdx > 0) {
      sheet.getRange(nameIdx + 1, 1, 1, row.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }
  }
  return { success: true, emp_id: data.emp_id };
}

function saveEmployees(records) {
  if (!Array.isArray(records)) return { success: false, error: 'Expected array' };
  var results = records.map(function(data) { return saveEmployee(data); });
  return { success: true, saved: results.length };
}

function deleteEmployee(empId) {
  const sheet = getSheet('Employees');
  const rows = sheet.getDataRange().getValues();
  const idx = rows.findIndex(function(r, i) { return i > 0 && r[0] === empId; });
  if (idx > 0) sheet.getRange(idx + 1, 9).setValue('Inactive');
  return { success: true };
}

// ---- Holidays ----

function getHolidays(year) {
  const sheet = getSheet('Holidays');
  const rows = sheet.getDataRange().getValues().slice(1).filter(function(r) { return r[0]; });
  const data = year
    ? rows.filter(function(r) { return String(r[0]).startsWith(String(year)); })
    : rows;
  return {
    success: true,
    data: data.map(function(r) { return { date: r[0], name: r[1] }; })
  };
}

function saveHolidays(holidays) {
  if (!Array.isArray(holidays) || !holidays.length) return { success: false, error: 'No holidays provided' };
  const year = String(holidays[0].date).substring(0, 4);
  const sheet = getSheet('Holidays');
  const all = sheet.getDataRange().getValues();
  const header = all[0];
  const others = all.slice(1).filter(function(r) { return r[0] && !String(r[0]).startsWith(year); });

  sheet.clearContents();
  sheet.appendRow(header);
  others.forEach(function(r) { sheet.appendRow(r); });
  holidays.forEach(function(h) { sheet.appendRow([h.date, h.name]); });

  return { success: true };
}

// ---- Advances ----
// Schema: advance_id | emp_id | emp_name | date | amount | status | created_by | created_at

function fmtDate(val) {
  if (!val) return '';
  if (val instanceof Date) return Utilities.formatDate(val, 'Asia/Kolkata', 'yyyy-MM-dd');
  return String(val).substring(0, 10);
}

function getAdvances(filter) {
  const sheet = getSheet('Advances');
  var rows = sheet.getDataRange().getValues().slice(1).filter(function(r) { return r[0]; })
    .map(function(r) {
      return {
        advance_id: r[0], emp_id: String(r[1] || '').trim(), emp_name: r[2],
        date: fmtDate(r[3]),
        amount: Number(r[4]),
        status: r[5] || 'Pending',
        created_by: r[6], created_at: r[7]
      };
    });

  if (filter) {
    if (filter.emp_id)   rows = rows.filter(function(a) { return a.emp_id === String(filter.emp_id).trim(); });
    if (filter.emp_name) rows = rows.filter(function(a) { return a.emp_name.toLowerCase().includes(String(filter.emp_name).toLowerCase()); });
    if (filter.status)   rows = rows.filter(function(a) { return a.status.toLowerCase() === filter.status.toLowerCase(); });
    if (filter.start_date && filter.end_date) {
      rows = rows.filter(function(a) { return a.date >= filter.start_date && a.date <= filter.end_date; });
    }
  }
  return { success: true, data: rows };
}

function saveAdvance(data) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = getSheet('Advances');
    const rows = sheet.getDataRange().getValues();
    // Force date to string so Google Sheets won't auto-convert to Date type
    var dateStr = String(data.date).substring(0, 10);
    var amount = Math.round(Number(data.amount) || 0);

    if (!data.advance_id) {
      // Dedup: carry-forward records are unique per employee per month — never insert twice
      if (String(data.created_by) === 'Previous Month Bal') {
        var monthPrefix = dateStr.substring(0, 7);
        var existing = rows.slice(1).find(function(r) {
          return r[0] && String(r[1]).trim() === String(data.emp_id).trim() &&
                 fmtDate(r[3]).substring(0, 7) === monthPrefix &&
                 String(r[6]) === 'Previous Month Bal';
        });
        if (existing) return { success: true, advance_id: existing[0], duplicate: true };
      }
      const nums = rows.slice(1).filter(function(r) { return r[0]; })
        .map(function(r) { return parseInt(String(r[0]).replace(/\D/g, '')) || 0; });
      const next = (Math.max.apply(null, [0].concat(nums)) + 1).toString().padStart(5, '0');
      data.advance_id = 'ADV' + next;
      data.created_at = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd HH:mm');
      var initStatus = data.status || 'Pending';
      sheet.appendRow([
        data.advance_id, data.emp_id, data.emp_name, dateStr,
        amount, initStatus, data.created_by || 'Manager', data.created_at
      ]);
    } else {
      // Edit: update date and amount, reset to Pending so admin re-approves
      const idx = rows.findIndex(function(r, i) { return i > 0 && String(r[0]) === String(data.advance_id); });
      if (idx > 0) {
        sheet.getRange(idx + 1, 4, 1, 3).setValues([[dateStr, amount, 'Pending']]);
      } else {
        // advance_id not found in sheet — fall back to inserting as a new record
        var nums2 = rows.slice(1).filter(function(r) { return r[0]; })
          .map(function(r) { return parseInt(String(r[0]).replace(/\D/g, '')) || 0; });
        var next2 = (Math.max.apply(null, [0].concat(nums2)) + 1).toString().padStart(5, '0');
        data.advance_id = 'ADV' + next2;
        data.created_at = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd HH:mm');
        sheet.appendRow([
          data.advance_id, data.emp_id, data.emp_name, dateStr,
          amount, data.status || 'Pending', data.created_by || 'Manager', data.created_at
        ]);
      }
    }
    return { success: true, advance_id: data.advance_id };
  } finally {
    lock.releaseLock();
  }
}

function approveAdvance(advanceId) {
  const sheet = getSheet('Advances');
  const rows = sheet.getDataRange().getValues();
  const idx = rows.findIndex(function(r, i) { return i > 0 && String(r[0]) === String(advanceId); });
  if (idx > 0) {
    sheet.getRange(idx + 1, 6).setValue('Approved');
    return { success: true };
  }
  return { success: false, error: 'Advance not found: ' + advanceId };
}

function approveAdvances(advanceIds) {
  if (!Array.isArray(advanceIds)) advanceIds = [advanceIds];
  const sheet = getSheet('Advances');
  const rows = sheet.getDataRange().getValues();
  const idSet = {};
  advanceIds.forEach(function(id) { idSet[String(id)] = true; });
  var approved = 0, notFound = [];
  for (var i = 1; i < rows.length; i++) {
    if (idSet[String(rows[i][0])]) {
      sheet.getRange(i + 1, 6).setValue('Approved');
      approved++;
      delete idSet[String(rows[i][0])];
    }
  }
  Object.keys(idSet).forEach(function(id) { notFound.push(id); });
  return { success: true, approved: approved, not_found: notFound };
}

function fixAdvanceEmpId(data) {
  // Re-links a list of advances to a new emp_id (corrects mismatched records)
  var advanceIds = data.advance_ids;
  var newEmpId   = String(data.emp_id).trim();
  var newEmpName = data.emp_name || '';
  if (!Array.isArray(advanceIds) || !newEmpId) return { success: false, error: 'advance_ids and emp_id required' };
  const sheet = getSheet('Advances');
  const rows  = sheet.getDataRange().getValues();
  const idSet = {};
  advanceIds.forEach(function(id) { idSet[String(id)] = true; });
  var fixed = 0;
  for (var i = 1; i < rows.length; i++) {
    if (idSet[String(rows[i][0])]) {
      sheet.getRange(i + 1, 2, 1, 2).setValues([[newEmpId, newEmpName]]);
      fixed++;
    }
  }
  return { success: true, fixed: fixed };
}

function deleteAdvance(advanceId) {
  const sheet = getSheet('Advances');
  const rows = sheet.getDataRange().getValues();
  const idx = rows.findIndex(function(r, i) { return i > 0 && r[0] === advanceId; });
  if (idx > 0) sheet.deleteRow(idx + 1);
  return { success: true };
}

// ---- Payroll ----

function getPayroll(month) {
  const sheet = getSheet('Payroll');
  var rows = sheet.getDataRange().getValues().slice(1).filter(function(r) { return r[0]; })
    .map(function(r) {
      var m = r[2];
      if (m instanceof Date) m = Utilities.formatDate(m, 'Asia/Kolkata', 'yyyy-MM');
      else m = String(m).substring(0, 7);
      return {
        payroll_id: r[0], emp_id: r[1], month: m,
        full_days: r[3], half_days: r[4], absent_days: r[5],
        week_off_days: r[6], holiday_absent_days: r[7],
        ot_weekday_min: r[8], ot_sunday_min: r[9], ot_holiday_min: r[10],
        shortfall_min: r[11], weekly_salary: r[12], daily_rate: r[13],
        hourly_rate: r[14], TD: r[15], gross_pay: r[16], ot_earnings: r[17],
        shortfall_deduction: r[18], bonus_eligible: r[19], bonus_cut: r[20],
        total_advances: r[21], net_pay: r[22], status: r[23],
        finalized_date: r[24], notes: r[25],
        adv_start_date: r[26] ? (r[26] instanceof Date ? Utilities.formatDate(r[26], 'Asia/Kolkata', 'yyyy-MM-dd') : String(r[26]).substring(0, 10)) : '',
        adv_end_date:   r[27] ? (r[27] instanceof Date ? Utilities.formatDate(r[27], 'Asia/Kolkata', 'yyyy-MM-dd') : String(r[27]).substring(0, 10)) : '',
        attendance_json: r[28] || ''
      };
    });
  if (month) rows = rows.filter(function(p) { return p.month === month; });
  return { success: true, data: rows };
}

function savePayroll(records) {
  if (!Array.isArray(records)) records = [records];
  const sheet = getSheet('Payroll');
  const existing = sheet.getDataRange().getValues();

  records.forEach(function(rec) {
    if (!rec.payroll_id) {
      rec.payroll_id = 'PAY' + (rec.month || '').replace('-', '') + rec.emp_id;
    }

    const row = [
      rec.payroll_id, rec.emp_id, rec.month,
      rec.full_days, rec.half_days, rec.absent_days,
      rec.week_off_days, rec.holiday_absent_days,
      rec.ot_weekday_min, rec.ot_sunday_min, rec.ot_holiday_min,
      rec.shortfall_min, rec.weekly_salary, rec.daily_rate,
      rec.hourly_rate, rec.TD, rec.gross_pay, rec.ot_earnings,
      rec.shortfall_deduction, rec.bonus_eligible, rec.bonus_cut,
      rec.total_advances, rec.net_pay, rec.status || 'draft',
      rec.finalized_date || '', rec.notes || '',
      rec.adv_start_date || '', rec.adv_end_date || '',
      rec.attendance_json || ''
    ];

    const idx = existing.findIndex(function(r, i) {
      if (i === 0) return false;
      var rowMonth = r[2] instanceof Date ?
        Utilities.formatDate(r[2], 'Asia/Kolkata', 'yyyy-MM') :
        String(r[2]).substring(0, 7);
      return r[1] === rec.emp_id && rowMonth === rec.month;
    });

    if (idx > 0) {
      if (existing[idx][23] === 'finalized') return;
      sheet.getRange(idx + 1, 1, 1, row.length).setValues([row]);
      existing[idx] = row;
    } else {
      sheet.appendRow(row);
      existing.push(row);
    }
  });

  return { success: true };
}

// Patch attendance_json on existing rows (including finalized) without touching other columns.
function patchAttendanceJson(data) {
  var records = Array.isArray(data) ? data : (data.records || []);
  var sheet = getSheet('Payroll');
  var rows = sheet.getDataRange().getValues();
  var patched = 0;
  records.forEach(function(rec) {
    if (!rec.emp_id || !rec.month || rec.attendance_json === undefined) return;
    var idx = -1;
    for (var i = 1; i < rows.length; i++) {
      var rowMonth = rows[i][2] instanceof Date
        ? Utilities.formatDate(rows[i][2], 'Asia/Kolkata', 'yyyy-MM')
        : String(rows[i][2]).substring(0, 7);
      if (String(rows[i][1]) === String(rec.emp_id) && rowMonth === rec.month) { idx = i; break; }
    }
    if (idx < 0) return;
    sheet.getRange(idx + 1, 29).setValue(rec.attendance_json);
    rows[idx][28] = rec.attendance_json;
    patched++;
  });
  return { success: true, patched: patched };
}

function reopenPayrollRecord(data) {
  var emp_id = String(data.emp_id);
  var month  = data.month;
  var sheet  = getSheet('Payroll');
  var rows   = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    var rowMonth = rows[i][2] instanceof Date
      ? Utilities.formatDate(rows[i][2], 'Asia/Kolkata', 'yyyy-MM')
      : String(rows[i][2]).substring(0, 7);
    if (String(rows[i][1]) === emp_id && rowMonth === month) {
      sheet.getRange(i + 1, 24).setValue('draft');
      sheet.getRange(i + 1, 25).setValue('');
      return { success: true };
    }
  }
  return { success: false, error: 'Record not found' };
}

function finalizePayroll(data) {
  const month = data.month;
  const advEndDate = data.adv_end_date || '';

  const payrollSheet = getSheet('Payroll');
  const bonusSheet = getSheet('Bonus');
  const settings = getSettingsMap();

  const bonusTarget = parseInt(settings.BONUS_POOL_TARGET);
  const companyCont = parseInt(settings.COMPANY_CONTRIBUTION);
  const today = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd');

  const bonusRows = bonusSheet.getDataRange().getValues().slice(1).filter(function(r) { return r[0]; });
  const balanceMap = {};
  bonusRows.forEach(function(r) { balanceMap[r[1]] = Number(r[7]); });

  // Build emp_id -> name lookup
  const empSheet = getSheet('Employees');
  const empRows = empSheet.getDataRange().getValues().slice(1);
  const empNameMap = {};
  empRows.forEach(function(r) { if (r[0]) empNameMap[String(r[0])] = r[1]; });

  const payrollRows = payrollSheet.getDataRange().getValues();
  var count = 0;

  for (var i = 1; i < payrollRows.length; i++) {
    var rowMonth = payrollRows[i][2] instanceof Date
      ? Utilities.formatDate(payrollRows[i][2], 'Asia/Kolkata', 'yyyy-MM')
      : String(payrollRows[i][2]).substring(0, 7);
    if (rowMonth !== month) continue;
    if (payrollRows[i][23] === 'finalized') continue;

    payrollSheet.getRange(i + 1, 24).setValue('finalized');
    payrollSheet.getRange(i + 1, 25).setValue(today);
    if (advEndDate) payrollSheet.getRange(i + 1, 28).setValue(advEndDate);
    count++;

    const empId = payrollRows[i][1];
    const eligible = payrollRows[i][19];
    const empCont = Number(payrollRows[i][20]);

    if (eligible && empCont > 0) {
      // Check if a bonus row already exists for this emp_id + month (handles reopen & re-finalize)
      const existingBonusRows = bonusSheet.getDataRange().getValues().slice(1).filter(function(r) { return r[0]; });
      const alreadyExists = existingBonusRows.some(function(r) {
        var bMonth = r[3] instanceof Date
          ? Utilities.formatDate(r[3], 'Asia/Kolkata', 'yyyy-MM')
          : String(r[3]).substring(0, 7);
        return String(r[1]) === String(empId) && bMonth === month;
      });
      if (alreadyExists) continue;

      const prevBal = balanceMap[empId] || 0;
      const added = empCont + companyCont;
      let newBal = prevBal + added;
      let payout = 0;

      if (newBal >= bonusTarget) {
        payout = newBal;
        newBal = 0;
      }
      balanceMap[empId] = newBal;

      const bonusNums = existingBonusRows
        .map(function(r) { return parseInt(String(r[0]).replace(/\D/g, '')) || 0; });
      const nextBon = (Math.max.apply(null, [0].concat(bonusNums)) + 1).toString().padStart(5, '0');

      var empName = empNameMap[String(empId)] || '';
      bonusSheet.appendRow(['BON' + nextBon, empId, empName, month, 'Y', added, payout, newBal, '']);
    }
  }

  return { success: true, finalized: count };
}

// ---- Payments ----

function getPayments(month) {
  const sheet = getSheet('Payments');
  var rows = sheet.getDataRange().getValues().slice(1).filter(function(r) { return r[0]; })
    .map(function(r) {
      var m = r[3] instanceof Date ? Utilities.formatDate(r[3], 'Asia/Kolkata', 'yyyy-MM') : String(r[3]).substring(0, 7);
      var d = r[4] instanceof Date ? Utilities.formatDate(r[4], 'Asia/Kolkata', 'yyyy-MM-dd') : String(r[4]).substring(0, 10);
      return {
        payment_id: r[0], emp_id: r[1], emp_name: r[2],
        month: m, date_paid: d, amount: Number(r[5]),
        mode: r[6], notes: r[7]
      };
    });
  if (month) rows = rows.filter(function(p) { return p.month === month; });
  return { success: true, data: rows };
}

function savePayment(data) {
  const sheet = getSheet('Payments');
  const rows = sheet.getDataRange().getValues();

  if (!data.payment_id) {
    const nums = rows.slice(1).filter(function(r) { return r[0]; })
      .map(function(r) { return parseInt(String(r[0]).replace(/\D/g, '')) || 0; });
    const next = (Math.max.apply(null, [0].concat(nums)) + 1).toString().padStart(5, '0');
    data.payment_id = 'PMT' + next;
    sheet.appendRow([
      data.payment_id, data.emp_id, data.emp_name,
      data.month, data.date_paid, data.amount, data.mode, data.notes || ''
    ]);
  } else {
    const idx = rows.findIndex(function(r, i) { return i > 0 && r[0] === data.payment_id; });
    if (idx > 0) {
      sheet.getRange(idx + 1, 1, 1, 8).setValues([[
        data.payment_id, data.emp_id, data.emp_name,
        data.month, data.date_paid, data.amount, data.mode, data.notes || ''
      ]]);
    }
  }
  return { success: true, payment_id: data.payment_id };
}

function deletePayment(paymentId) {
  const sheet = getSheet('Payments');
  const rows = sheet.getDataRange().getValues();
  const idx = rows.findIndex(function(r, i) { return i > 0 && r[0] === paymentId; });
  if (idx > 0) sheet.deleteRow(idx + 1);
  return { success: true };
}

function getFinalizedMonths() {
  var rows = getSheet('Payroll').getDataRange().getValues().slice(1).filter(function(r) { return r[0] && r[23] === 'finalized'; });
  var months = {};
  rows.forEach(function(r) {
    var m = r[2] instanceof Date ? Utilities.formatDate(r[2], 'Asia/Kolkata', 'yyyy-MM') : String(r[2]).substring(0, 7);
    months[m] = true;
  });
  return { success: true, data: Object.keys(months).sort().reverse() };
}

function getPaymentRun(month) {
  var payroll = getPayroll(month).data.filter(function(p) { return p.status === 'finalized'; });
  var payments = getPayments(month).data;
  var empSheet = getSheet('Employees');
  var empRows = empSheet.getDataRange().getValues().slice(1).filter(function(r) { return r[0]; });
  var empMap = {};
  empRows.forEach(function(r) { empMap[r[0]] = r[1]; });

  var paidMap = {};
  payments.forEach(function(p) {
    paidMap[p.emp_id] = (paidMap[p.emp_id] || 0) + p.amount;
  });

  var run = payroll.map(function(p) {
    var paid = paidMap[p.emp_id] || 0;
    return {
      emp_id: p.emp_id,
      emp_name: empMap[p.emp_id] || p.emp_id,
      net_pay: Number(p.net_pay) || 0,
      already_paid: paid,
      remaining: (Number(p.net_pay) || 0) - paid
    };
  }).sort(function(a, b) { return a.emp_name > b.emp_name ? 1 : -1; });

  return { success: true, data: run };
}

function saveBulkPayments(records) {
  if (!Array.isArray(records) || !records.length) return { success: false, error: 'No records' };
  var sheet = getSheet('Payments');
  var rows = sheet.getDataRange().getValues();
  var nums = rows.slice(1).filter(function(r) { return r[0]; })
    .map(function(r) { return parseInt(String(r[0]).replace(/\D/g, '')) || 0; });
  var next = Math.max.apply(null, [0].concat(nums));
  records.forEach(function(rec) {
    next++;
    sheet.appendRow([
      'PMT' + String(next).padStart(5, '0'),
      rec.emp_id, rec.emp_name, rec.month,
      rec.date_paid, rec.amount, rec.mode, rec.notes || ''
    ]);
  });
  return { success: true, saved: records.length };
}

// ---- Bonus Pool ----

function getBonusPool() {
  const bonusSheet = getSheet('Bonus');
  const empSheet = getSheet('Employees');

  const empMap = {};
  empSheet.getDataRange().getValues().slice(1).filter(function(r) { return r[0]; })
    .forEach(function(r) { empMap[r[0]] = r[1]; });

  const rows = bonusSheet.getDataRange().getValues().slice(1).filter(function(r) { return r[0]; });

  const latest = {};
  rows.forEach(function(r) {
    const empId = r[1];
    if (!latest[empId] || r[3] > latest[empId].month) {
      latest[empId] = { month: r[3], balance: Number(r[7]), last_payout: Number(r[6]) };
    }
  });

  const data = Object.keys(latest).map(function(empId) {
    return {
      emp_id: empId,
      emp_name: empMap[empId] || empId,
      balance: latest[empId].balance,
      last_updated_month: latest[empId].month,
      last_payout: latest[empId].last_payout
    };
  });

  // Pending payouts: rows where payout > 0 and notes does not start with 'paid:'
  const pending = rows
    .filter(function(r) { return Number(r[6]) > 0 && !String(r[8] || '').toLowerCase().startsWith('paid:'); })
    .map(function(r) {
      return { bonus_id: r[0], emp_id: r[1], emp_name: r[2], month: r[3], payout: Number(r[6]) };
    });

  return { success: true, data: data, pending: pending };
}

function confirmBonusPaid(data) {
  var bonusId = data.bonus_id;
  var empId   = String(data.emp_id);
  var today   = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd');

  // Mark bonus row as paid in notes column
  var bonusSheet = getSheet('Bonus');
  var bonusRows  = bonusSheet.getDataRange().getValues();
  for (var i = 1; i < bonusRows.length; i++) {
    if (String(bonusRows[i][0]) === String(bonusId)) {
      bonusSheet.getRange(i + 1, 9).setValue('paid:' + today);
      break;
    }
  }

  // Flip bonus_scheme to No on the employee
  var empSheet = getSheet('Employees');
  var empRows  = empSheet.getDataRange().getValues();
  for (var j = 1; j < empRows.length; j++) {
    if (String(empRows[j][0]) === empId) {
      empSheet.getRange(j + 1, 12).setValue('No');
      break;
    }
  }

  return { success: true };
}

// Upload initial bonus balances from Excel (emp_name + balance columns)
function saveBonusBalances(data) {
  var records = data.records || [];
  if (!records.length) return { success: false, error: 'No records provided' };

  var empSheet = getSheet('Employees');
  var empRows = empSheet.getDataRange().getValues().slice(1).filter(function(r) { return r[0]; });
  var norm = function(s) { return String(s || '').toLowerCase().replace(/\s+/g, ' ').trim(); };
  var empByName = {};
  empRows.forEach(function(r) {
    empByName[norm(r[1])] = r[0]; // name -> emp_id
    empByName[norm(r[5])] = r[0]; // petpooja_name -> emp_id
  });

  var bonusSheet = getSheet('Bonus');
  var bonusRows = bonusSheet.getDataRange().getValues().slice(1).filter(function(r) { return r[0]; });
  var bonusNums = bonusRows.map(function(r) { return parseInt(String(r[0]).replace(/\D/g, '')) || 0; });
  var nextNum = Math.max.apply(null, [0].concat(bonusNums));

  var saved = 0, skipped = 0;
  records.forEach(function(rec) {
    var empId = rec.emp_id || empByName[norm(rec.emp_name)] || null;
    if (!empId) { skipped++; return; }
    var empNameVal = rec.emp_name || '';
    var balance = Number(rec.balance) || 0;
    nextNum++;
    bonusSheet.appendRow([
      'BON' + String(nextNum).padStart(5, '0'),
      empId, empNameVal, 'initial', 'Y', 0, 0, balance, 'uploaded'
    ]);
    saved++;
  });

  return { success: true, saved: saved, skipped: skipped };
}
