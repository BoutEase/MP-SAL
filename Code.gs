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
      deleteEmployee:  () => deleteEmployee(data.emp_id),
      getHolidays:     () => getHolidays(data && data.year),
      saveHolidays:    () => saveHolidays(data),
      getAdvances:     () => getAdvances(data),
      saveAdvance:     () => saveAdvance(data),
      approveAdvance:  () => approveAdvance(data.advance_id),
      deleteAdvance:   () => deleteAdvance(data.advance_id),
      getPayroll:      () => getPayroll(data && data.month),
      savePayroll:     () => savePayroll(data),
      finalizePayroll: () => finalizePayroll(data),
      getPayments:     () => getPayments(data && data.month),
      savePayment:     () => savePayment(data),
      deletePayment:   () => deletePayment(data.payment_id),
      getBonusPool:    () => getBonusPool(),
      getSettings:     () => getSettings(),
      saveSettings:    () => saveSettings(data),
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
  return false;
}

function verifyPin(data) {
  return { success: checkPin(data.pin, data.role), role: data.role };
}

// ---- Settings ----

function getDefaultSettings() {
  return {
    ADMIN_PIN: '1234',
    MANAGER_PIN: '5678',
    SHORTFALL_THRESHOLD_HOURS: '8.5',
    ADD_LUNCH_TO_OT: 'true',
    LUNCH_DURATION_HOURS: '1',
    COMPANY_NAME: 'Mahitha Prasad',
    BONUS_POOL_TARGET: '24000',
    EMPLOYEE_CONTRIBUTION: '500',
    COMPANY_CONTRIBUTION: '500',
    BONUS_ELIGIBILITY_MIN_DAYS: '6',
    OT_MIN_HOURS: '10',
    OT_ROUND_MINUTES: '15'
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
    Employees: ['emp_id','name','team','designation','petpooja_id','petpooja_name','weekly_salary','login_code','status','joining_date'],
    Holidays:  ['date','name'],
    Payroll:   ['payroll_id','emp_id','month','full_days','half_days','absent_days','week_off_days','holiday_absent_days','ot_weekday_min','ot_sunday_min','ot_holiday_min','shortfall_min','weekly_salary','daily_rate','hourly_rate','TD','gross_pay','ot_earnings','shortfall_deduction','bonus_eligible','bonus_cut','total_advances','net_pay','status','finalized_date','notes','adv_start_date','adv_end_date'],
    Advances:  ['advance_id','emp_id','emp_name','date','amount','status','created_by','created_at'],
    Bonus:     ['bonus_id','emp_id','month','eligible','contribution','payout','balance_after','notes'],
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
        status: r[8], joining_date: r[9]
      };
    })
  };
}

function saveEmployee(data) {
  const sheet = getSheet('Employees');
  const rows = sheet.getDataRange().getValues();

  if (!data.emp_id) {
    const nums = rows.slice(1).filter(function(r) { return r[0]; })
      .map(function(r) { return parseInt(String(r[0]).replace(/\D/g, '')) || 0; });
    const next = (Math.max.apply(null, [0].concat(nums)) + 1).toString().padStart(3, '0');
    data.emp_id = 'MP' + next;
  }

  const row = [
    data.emp_id, data.name, data.team, data.designation,
    data.petpooja_id || '', data.petpooja_name || data.name,
    data.weekly_salary, data.login_code || '', data.status || 'Active',
    data.joining_date || Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd')
  ];

  const idx = rows.findIndex(function(r, i) { return i > 0 && r[0] === data.emp_id; });
  if (idx > 0) {
    sheet.getRange(idx + 1, 1, 1, row.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }
  return { success: true, emp_id: data.emp_id };
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

function getAdvances(filter) {
  const sheet = getSheet('Advances');
  var rows = sheet.getDataRange().getValues().slice(1).filter(function(r) { return r[0]; })
    .map(function(r) {
      return {
        advance_id: r[0], emp_id: r[1], emp_name: r[2],
        date: String(r[3]).substring(0, 10),
        amount: Number(r[4]),
        status: r[5] || 'Pending',
        created_by: r[6], created_at: r[7]
      };
    });

  if (filter) {
    if (filter.emp_id)   rows = rows.filter(function(a) { return a.emp_id === filter.emp_id; });
    if (filter.status)   rows = rows.filter(function(a) { return a.status === filter.status; });
    if (filter.start_date && filter.end_date) {
      rows = rows.filter(function(a) { return a.date >= filter.start_date && a.date <= filter.end_date; });
    }
  }
  return { success: true, data: rows };
}

function saveAdvance(data) {
  const sheet = getSheet('Advances');
  const rows = sheet.getDataRange().getValues();

  if (!data.advance_id) {
    const nums = rows.slice(1).filter(function(r) { return r[0]; })
      .map(function(r) { return parseInt(String(r[0]).replace(/\D/g, '')) || 0; });
    const next = (Math.max.apply(null, [0].concat(nums)) + 1).toString().padStart(5, '0');
    data.advance_id = 'ADV' + next;
    data.created_at = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd HH:mm');
    sheet.appendRow([
      data.advance_id, data.emp_id, data.emp_name, data.date,
      data.amount, 'Pending', data.created_by || 'Manager', data.created_at
    ]);
  } else {
    // Edit: only update date and amount; keep status/created_by/created_at unchanged
    const idx = rows.findIndex(function(r, i) { return i > 0 && r[0] === data.advance_id; });
    if (idx > 0) {
      sheet.getRange(idx + 1, 4, 1, 2).setValues([[data.date, data.amount]]);
    }
  }
  return { success: true, advance_id: data.advance_id };
}

function approveAdvance(advanceId) {
  const sheet = getSheet('Advances');
  const rows = sheet.getDataRange().getValues();
  const idx = rows.findIndex(function(r, i) { return i > 0 && r[0] === advanceId; });
  if (idx > 0) sheet.getRange(idx + 1, 6).setValue('Approved');
  return { success: true };
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
      return {
        payroll_id: r[0], emp_id: r[1], month: r[2],
        full_days: r[3], half_days: r[4], absent_days: r[5],
        week_off_days: r[6], holiday_absent_days: r[7],
        ot_weekday_min: r[8], ot_sunday_min: r[9], ot_holiday_min: r[10],
        shortfall_min: r[11], weekly_salary: r[12], daily_rate: r[13],
        hourly_rate: r[14], TD: r[15], gross_pay: r[16], ot_earnings: r[17],
        shortfall_deduction: r[18], bonus_eligible: r[19], bonus_cut: r[20],
        total_advances: r[21], net_pay: r[22], status: r[23],
        finalized_date: r[24], notes: r[25],
        adv_start_date: r[26] || '', adv_end_date: r[27] || ''
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
      rec.adv_start_date || '', rec.adv_end_date || ''
    ];

    const idx = existing.findIndex(function(r, i) {
      return i > 0 && r[1] === rec.emp_id && r[2] === rec.month;
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
  bonusRows.forEach(function(r) { balanceMap[r[1]] = Number(r[6]); });

  const payrollRows = payrollSheet.getDataRange().getValues();
  var count = 0;

  for (var i = 1; i < payrollRows.length; i++) {
    if (payrollRows[i][2] !== month) continue;
    if (payrollRows[i][23] === 'finalized') continue;

    payrollSheet.getRange(i + 1, 24).setValue('finalized');
    payrollSheet.getRange(i + 1, 25).setValue(today);
    if (advEndDate) payrollSheet.getRange(i + 1, 28).setValue(advEndDate);
    count++;

    const empId = payrollRows[i][1];
    const eligible = payrollRows[i][19];
    const empCont = Number(payrollRows[i][20]);

    if (eligible && empCont > 0) {
      const prevBal = balanceMap[empId] || 0;
      const added = empCont + companyCont;
      let newBal = prevBal + added;
      let payout = 0;

      if (newBal >= bonusTarget) {
        payout = newBal;
        newBal = 0;
      }
      balanceMap[empId] = newBal;

      const bonusNums = bonusSheet.getDataRange().getValues().slice(1)
        .filter(function(r) { return r[0]; })
        .map(function(r) { return parseInt(String(r[0]).replace(/\D/g, '')) || 0; });
      const nextBon = (Math.max.apply(null, [0].concat(bonusNums)) + 1).toString().padStart(5, '0');

      bonusSheet.appendRow(['BON' + nextBon, empId, month, 'Y', added, payout, newBal, '']);
    }
  }

  return { success: true, finalized: count };
}

// ---- Payments ----

function getPayments(month) {
  const sheet = getSheet('Payments');
  var rows = sheet.getDataRange().getValues().slice(1).filter(function(r) { return r[0]; })
    .map(function(r) {
      return {
        payment_id: r[0], emp_id: r[1], emp_name: r[2],
        month: r[3], date_paid: r[4], amount: Number(r[5]),
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
    if (!latest[empId] || r[2] > latest[empId].month) {
      latest[empId] = { month: r[2], balance: Number(r[6]), last_payout: Number(r[5]) };
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

  return { success: true, data: data };
}
