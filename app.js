// ============================================================
// MP-SAL Payroll — Frontend App
// ============================================================

// ---- State ----
const State = {
  role: null, pin: null, gasUrl: null,
  employees: [], holidays: [], advances: [], payroll: [],
  payments: [], bonusPool: [], settings: {},
  currentMonth: (function() {
    const d = new Date();
    return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
  })(),
  parsedAttendance: null,
  unmatchedAttendance: []
};

// ---- API ----
const API = {
  async call(action, data) {
    const body = { action, data, pin: State.pin, role: State.role };
    const res = await fetch(State.gasUrl, { method: 'POST', body: JSON.stringify(body) });
    const json = await res.json();
    if (!json.success) throw new Error(json.error || 'API error');
    return json;
  }
};

// ---- Toast ----
function toast(msg, type) {
  const el = document.getElementById('toast');
  el.textContent = msg;
  el.style.background = type === 'error' ? 'var(--danger)' : type === 'success' ? 'var(--success)' : 'var(--text)';
  el.classList.add('show');
  setTimeout(() => el.classList.remove('show'), 3000);
}

// ---- App init ----
const App = {
  init() {
    State.gasUrl = localStorage.getItem('mpsal_url');
    const theme = localStorage.getItem('mpsal_theme') || 'light';
    document.documentElement.setAttribute('data-theme', theme);
    if (!State.gasUrl) {
      show('screen-setup');
    } else {
      show('screen-login');
    }
  },
  toggleTheme() {
    const cur = document.documentElement.getAttribute('data-theme');
    const next = cur === 'dark' ? 'light' : 'dark';
    document.documentElement.setAttribute('data-theme', next);
    localStorage.setItem('mpsal_theme', next);
  },
  changeUrl() {
    localStorage.removeItem('mpsal_url');
    State.gasUrl = null;
    show('screen-setup');
  },
  async setup() {
    const url = document.getElementById('setup-url').value.trim();
    if (!url.startsWith('https://')) { toast('Enter a valid URL', 'error'); return; }
    const status = document.getElementById('setup-status');
    status.textContent = 'Connecting…';
    State.gasUrl = url;
    State.role = 'admin'; State.pin = '';
    try {
      const r = await fetch(url, { method: 'POST', body: JSON.stringify({ action: 'initSheets', data: {} }) });
      const j = await r.json();
      if (!j.success) throw new Error(j.error);
      localStorage.setItem('mpsal_url', url);
      status.textContent = 'Connected! ' + (j.message || '');
      setTimeout(() => { State.role = null; State.pin = null; show('screen-login'); }, 1200);
    } catch(e) {
      status.textContent = 'Error: ' + e.message;
    }
  }
};

// ---- Auth ----
const Auth = {
  selectedRole: 'admin',
  selectRole(role) {
    this.selectedRole = role;
    document.getElementById('tab-admin').classList.toggle('active', role === 'admin');
    document.getElementById('tab-manager').classList.toggle('active', role === 'manager');
    document.getElementById('pin-input').value = '';
    document.getElementById('pin-input').focus();
  },
  async login() {
    const pin = document.getElementById('pin-input').value.trim();
    if (!pin) { toast('Enter PIN', 'error'); return; }
    State.role = this.selectedRole;
    State.pin = pin;
    try {
      const r = await API.call('verifyPin', { pin, role: this.selectedRole });
      if (!r.success) { toast('Wrong PIN', 'error'); State.pin = null; return; }
      await loadCoreData();
      if (State.role === 'admin') AdminApp.init();
      else ManagerApp.init();
    } catch(e) {
      toast(e.message, 'error');
    }
  },
  logout() {
    State.role = null; State.pin = null;
    document.getElementById('pin-input').value = '';
    show('screen-login');
  }
};

async function loadCoreData() {
  const [empRes, settRes] = await Promise.all([
    API.call('getEmployees', {}),
    API.call('getSettings', {})
  ]);
  State.employees = empRes.data || [];
  State.settings = settRes.data || {};
}

// ---- Utility ----
function show(id) {
  ['screen-setup','screen-login','screen-admin','screen-manager'].forEach(s => {
    const el = document.getElementById(s);
    if (el) el.classList.toggle('hidden', s !== id);
  });
}

function fmt(n) { return '₹' + Math.round(n).toLocaleString('en-IN'); }
function fmtHrs(mins) {
  const h = Math.floor(Math.abs(mins) / 60), m = Math.abs(mins) % 60;
  return (mins < 0 ? '-' : '') + h + 'h' + (m ? ' ' + m + 'm' : '');
}
function currentYear() { return State.currentMonth.substring(0, 4); }

function el(tag, attrs, ...children) {
  const e = document.createElement(tag);
  Object.entries(attrs || {}).forEach(([k, v]) => {
    if (k === 'class') e.className = v;
    else if (k === 'html') e.innerHTML = v;
    else if (k.startsWith('on')) e.addEventListener(k.slice(2), v);
    else e.setAttribute(k, v);
  });
  children.forEach(c => { if (c != null) e.append(typeof c === 'string' ? document.createTextNode(c) : c); });
  return e;
}

// ============================================================
// PETPOOJA PARSER
// ============================================================
const Parser = {
  parse(fileData) {
    const wb = XLSX.read(fileData, { type: 'array' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    const employees = [];

    for (let i = 0; i < rows.length; i++) {
      const cell = String(rows[i][0] || '').trim();
      if (!cell.includes('Employee Name:')) continue;

      const nameM  = cell.match(/Employee Name:\s*(.+?)\s+Department:/);
      const deptM  = cell.match(/Department:\s*(.+?)\s+Designation:/);
      const idM    = cell.match(/Employee ID:\s*(\d+)/);
      const fromM  = cell.match(/From:\s*(\d{4}-\d{2}-\d{2})/);

      const emp = {
        name: nameM ? nameM[1].trim() : '',
        department: deptM ? deptM[1].trim() : '',
        petpooja_id: idM ? String(parseInt(idM[1])) : '',
        month: fromM ? fromM[1].substring(0, 7) : '',
        days: []
      };

      // Scan next 15 rows for Day / Status / Total Hrs
      let dayRow = null, statusRow = null, hrsRow = null;
      for (let j = i + 1; j < Math.min(i + 18, rows.length); j++) {
        const lbl = String(rows[j][0] || '').trim();
        if (lbl === 'Day')       dayRow    = rows[j];
        if (lbl === 'Status')    statusRow = rows[j];
        if (lbl === 'Total Hrs') { hrsRow  = rows[j]; break; }
      }

      if (dayRow && statusRow && hrsRow) {
        for (let d = 1; d < dayRow.length; d++) {
          const header  = String(dayRow[d] || '').trim();
          const dm      = header.match(/(\d+)\s*\((\w+)\)/);
          if (!dm) continue;

          const dayNum  = parseInt(dm[1]);
          const dayName = dm[2]; // Sunday, Monday, …
          const status  = String(statusRow[d] || '').trim();
          const hrsStr  = String(hrsRow[d] || '').trim();

          let totalHours = 0;
          if (hrsStr && hrsStr !== '-') {
            const hm = hrsStr.match(/(\d+)h(?:\s*(\d+)m)?/);
            if (hm) totalHours = parseInt(hm[1]) + (parseInt(hm[2] || 0)) / 60;
          }

          const month = emp.month;
          emp.days.push({
            day: dayNum, dayName, status, totalHours,
            date: month + '-' + String(dayNum).padStart(2, '0')
          });
        }
        employees.push(emp);
      }
    }
    return employees;
  }
};

// ============================================================
// SALARY CALCULATOR
// ============================================================
const Calc = {
  run(empAttendance, dbEmployee, advancesForMonth, holidays, settings) {
    const weekly  = dbEmployee.weekly_salary;
    const daily   = weekly / 6;
    const hourly  = daily / 9;

    const shortfallThresh = parseFloat(settings.SHORTFALL_THRESHOLD_HOURS || 8.5);
    const addLunch        = settings.ADD_LUNCH_TO_OT === 'true';
    const lunchHrs        = parseFloat(settings.LUNCH_DURATION_HOURS || 1);
    const otMinHours      = parseFloat(settings.OT_MIN_HOURS || 10);
    const otRound         = parseInt(settings.OT_ROUND_MINUTES || 15);
    const bonusMinDays    = parseInt(settings.BONUS_ELIGIBILITY_MIN_DAYS || 6);
    const empContrib      = parseInt(settings.EMPLOYEE_CONTRIBUTION || 500);

    const holidayDates = new Set(holidays.map(h => h.date));

    let fullDays = 0, halfDays = 0, absentDays = 0, weekOffDays = 0, holidayAbsent = 0;
    let otWeekdayMin = 0, otSundayMin = 0, otHolidayMin = 0, shortfallMin = 0;

    for (const day of empAttendance.days) {
      const isSunday  = day.dayName === 'Sunday';
      const isHoliday = holidayDates.has(day.date);

      if (isSunday) {
        weekOffDays++;
        if (day.totalHours > 0) {
          const otHrs = day.totalHours + (addLunch ? lunchHrs : 0);
          otSundayMin += Math.round(otHrs * 60);
        }
      } else if (isHoliday) {
        if (day.status === 'FD') {
          fullDays++;
          // Present on holiday: counts as FD + all hours as holiday OT
          const otHrs = day.totalHours + (addLunch ? lunchHrs : 0);
          otHolidayMin += Math.round(otHrs * 60);
        } else if (day.status === 'HD') {
          halfDays++;
          shortfallMin += Math.round((9 - day.totalHours) * 60);
          const otHrs = day.totalHours + (addLunch ? lunchHrs : 0);
          otHolidayMin += Math.round(otHrs * 60);
        } else {
          // Absent on holiday = paid
          holidayAbsent++;
        }
      } else if (day.status === 'FD') {
        fullDays++;
        if (day.totalHours >= otMinHours) {
          const rawOt = day.totalHours - 9;
          const rounded = Math.floor(rawOt * 60 / otRound) * otRound;
          otWeekdayMin += rounded;
        } else if (day.totalHours > 0 && day.totalHours < shortfallThresh) {
          shortfallMin += Math.round((9 - day.totalHours) * 60);
        }
      } else if (day.status === 'HD') {
        halfDays++;
        shortfallMin += Math.round((9 - day.totalHours) * 60);
      } else if (day.status === 'Absent') {
        absentDays++;
      }
    }

    const TD        = fullDays + halfDays + holidayAbsent;
    const gross     = TD * daily;
    const allOtMin  = otWeekdayMin + otSundayMin + otHolidayMin;
    const otEarn    = (allOtMin / 60) * hourly;
    const shortfall = (shortfallMin / 60) * hourly;

    const workDays      = fullDays + halfDays;
    const bonusEligible = workDays >= bonusMinDays;
    const bonusCut      = bonusEligible ? empContrib : 0;

    const totalAdv = advancesForMonth.reduce((s, a) => s + Number(a.amount), 0);
    const net      = gross + otEarn - shortfall - bonusCut - totalAdv;

    return {
      emp_id: dbEmployee.emp_id, month: empAttendance.month,
      full_days: fullDays, half_days: halfDays, absent_days: absentDays,
      week_off_days: weekOffDays, holiday_absent_days: holidayAbsent,
      ot_weekday_min: otWeekdayMin, ot_sunday_min: otSundayMin, ot_holiday_min: otHolidayMin,
      shortfall_min: shortfallMin,
      weekly_salary: weekly, daily_rate: Math.round(daily * 100) / 100,
      hourly_rate: Math.round(hourly * 100) / 100,
      TD, gross_pay: Math.round(gross), ot_earnings: Math.round(otEarn),
      shortfall_deduction: Math.round(shortfall),
      bonus_eligible: bonusEligible, bonus_cut: bonusCut,
      total_advances: totalAdv, net_pay: Math.round(net),
      status: 'draft'
    };
  },

  matchEmployee(parsed, dbEmployees) {
    // Try petpooja_id first, then name
    return dbEmployees.find(e => e.status === 'Active' && String(e.petpooja_id) === String(parsed.petpooja_id))
        || dbEmployees.find(e => e.status === 'Active' && e.petpooja_name?.toLowerCase() === parsed.name.toLowerCase())
        || dbEmployees.find(e => e.status === 'Active' && e.name.toLowerCase() === parsed.name.toLowerCase());
  }
};

// ============================================================
// ADMIN APP
// ============================================================
const AdminApp = {
  activeView: 'payroll',

  init() {
    show('screen-admin');
    document.getElementById('screen-admin').innerHTML = this.renderShell();
    this.switchView('payroll');
    Views.Payroll.load();
  },

  renderShell() {
    return `
    <header style="background:var(--surface);border-bottom:1px solid var(--border);padding:0 16px;display:flex;align-items:center;gap:12px;height:52px;position:sticky;top:0;z-index:100">
      <span style="font-weight:700;font-size:16px">MP&#x2009;SAL</span>
      <span style="flex:1"></span>
      <button class="theme-toggle" onclick="App.toggleTheme()" style="border:1px solid var(--border);background:none;padding:5px 10px;border-radius:6px;cursor:pointer;font-size:13px;color:var(--muted)">Theme</button>
      <button class="theme-toggle" onclick="Auth.logout()" style="border:1px solid var(--border);background:none;padding:5px 10px;border-radius:6px;cursor:pointer;font-size:13px;color:var(--danger)">Logout</button>
    </header>
    <nav style="display:flex;overflow-x:auto;border-bottom:1px solid var(--border);background:var(--surface);padding:0 8px">
      ${['payroll','employees','advances','bonus','payments','settings'].map(v =>
        `<button id="nav-${v}" onclick="AdminApp.switchView('${v}')"
          style="padding:12px 16px;border:none;background:none;cursor:pointer;font-size:13px;font-weight:500;white-space:nowrap;border-bottom:2px solid transparent;color:var(--muted)"
          >${v.charAt(0).toUpperCase()+v.slice(1)}</button>`
      ).join('')}
    </nav>
    <main style="padding:16px;max-width:1200px;margin:0 auto">
      <div id="view-payroll"   class="view"></div>
      <div id="view-employees" class="view hidden"></div>
      <div id="view-advances"  class="view hidden"></div>
      <div id="view-bonus"     class="view hidden"></div>
      <div id="view-payments"  class="view hidden"></div>
      <div id="view-settings"  class="view hidden"></div>
    </main>`;
  },

  switchView(name) {
    this.activeView = name;
    document.querySelectorAll('.view').forEach(v => v.classList.add('hidden'));
    document.getElementById('view-' + name)?.classList.remove('hidden');
    document.querySelectorAll('nav button').forEach(b => {
      b.style.borderBottomColor = 'transparent';
      b.style.color = 'var(--muted)';
    });
    const active = document.getElementById('nav-' + name);
    if (active) { active.style.borderBottomColor = 'var(--primary)'; active.style.color = 'var(--primary)'; }

    const loaders = { employees: Views.Employees, advances: Views.Advances,
                      bonus: Views.Bonus, payments: Views.Payments, settings: Views.Settings };
    if (loaders[name]) loaders[name].load();
  }
};

// ============================================================
// VIEWS — PAYROLL
// ============================================================
const Views = {};
Views.Payroll = {
  async load() {
    const c = document.getElementById('view-payroll');
    c.innerHTML = `
    <div style="display:flex;align-items:center;gap:10px;margin-bottom:16px;flex-wrap:wrap">
      <h2 style="font-size:18px;font-weight:700">Payroll</h2>
      <input type="month" id="payroll-month" value="${State.currentMonth}"
        style="width:auto;padding:6px 10px" onchange="Views.Payroll.onMonthChange(this.value)" />
      <span style="flex:1"></span>
      <label class="btn btn-secondary btn-sm" style="cursor:pointer">
        Upload Petpooja Excel
        <input type="file" accept=".xlsx,.xls" style="display:none" onchange="Views.Payroll.onUpload(event)" />
      </label>
      <button class="btn btn-secondary btn-sm" onclick="Views.Payroll.saveDraft()">Save Draft</button>
      <button class="btn btn-primary btn-sm" onclick="Views.Payroll.finalize()">Finalise Month</button>
    </div>
    <div id="payroll-unmatched"></div>
    <div id="payroll-table-wrap"><p style="color:var(--muted);font-size:14px">Upload a Petpooja Excel to calculate payroll, or load saved draft below.</p></div>
    <div style="margin-top:12px"><button class="btn btn-secondary btn-sm" onclick="Views.Payroll.loadDraft()">Load Saved Draft</button></div>`;
  },

  onMonthChange(m) { State.currentMonth = m; },

  async onUpload(e) {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = async (ev) => {
      try {
        const parsed = Parser.parse(new Uint8Array(ev.target.result));
        if (!parsed.length) { toast('No employee data found in file', 'error'); return; }

        const month = parsed[0].month;
        document.getElementById('payroll-month').value = month;
        State.currentMonth = month;

        // Load holidays and advances
        const [holRes, advRes] = await Promise.all([
          API.call('getHolidays', { year: month.substring(0,4) }),
          API.call('getAdvances', { month })
        ]);
        State.holidays = holRes.data || [];
        State.advances = advRes.data || [];

        // Match and calculate
        const matched = [], unmatched = [];
        for (const emp of parsed) {
          const dbEmp = Calc.matchEmployee(emp, State.employees);
          if (dbEmp) {
            const advances = State.advances.filter(a => a.emp_id === dbEmp.emp_id);
            matched.push(Calc.run(emp, dbEmp, advances, State.holidays, State.settings));
          } else {
            unmatched.push(emp);
          }
        }
        State.payroll = matched;
        State.unmatchedAttendance = unmatched;

        this.renderUnmatched(unmatched);
        this.renderTable(matched);
        toast(`Calculated payroll for ${matched.length} employees`, 'success');
      } catch(err) { toast('Parse error: ' + err.message, 'error'); }
    };
    reader.readAsArrayBuffer(file);
  },

  renderUnmatched(list) {
    const c = document.getElementById('payroll-unmatched');
    if (!list.length) { c.innerHTML = ''; return; }
    c.innerHTML = `<div class="card" style="margin-bottom:12px;border-color:var(--warning);background:var(--warning-light)">
      <strong style="font-size:14px">⚠ ${list.length} employee(s) not matched — map them to continue:</strong>
      <table style="width:100%;margin-top:8px;font-size:13px;border-collapse:collapse">
        <thead><tr><th style="text-align:left;padding:4px 8px">Petpooja Name (ID)</th><th style="padding:4px 8px">Map to Employee</th><th></th></tr></thead>
        <tbody>${list.map((u, i) => `
          <tr>
            <td style="padding:4px 8px">${u.name} <span style="color:var(--muted)">(ID:${u.petpooja_id})</span></td>
            <td style="padding:4px 8px">
              <select id="map-${i}" style="padding:4px">
                <option value="">— select —</option>
                ${State.employees.filter(e => e.status==='Active').map(e =>
                  `<option value="${e.emp_id}">${e.name}</option>`).join('')}
              </select>
            </td>
            <td style="padding:4px 8px">
              <button class="btn btn-primary btn-sm" onclick="Views.Payroll.resolveUnmatched(${i})">Apply</button>
            </td>
          </tr>`).join('')}
      </tbody></table></div>`;
  },

  async resolveUnmatched(i) {
    const empId = document.getElementById('map-' + i).value;
    if (!empId) { toast('Select an employee', 'error'); return; }
    const dbEmp = State.employees.find(e => e.emp_id === empId);
    const att   = State.unmatchedAttendance[i];
    const advances = State.advances.filter(a => a.emp_id === empId);
    const rec   = Calc.run(att, dbEmp, advances, State.holidays, State.settings);
    State.payroll.push(rec);
    State.unmatchedAttendance.splice(i, 1);
    this.renderUnmatched(State.unmatchedAttendance);
    this.renderTable(State.payroll);
    toast('Mapped ' + dbEmp.name, 'success');
  },

  renderTable(payroll) {
    const wrap = document.getElementById('payroll-table-wrap');
    if (!payroll.length) { wrap.innerHTML = '<p style="color:var(--muted)">No payroll data.</p>'; return; }

    const empMap = {};
    State.employees.forEach(e => empMap[e.emp_id] = e);

    const rows = payroll.map((p, idx) => {
      const emp = empMap[p.emp_id] || { name: p.emp_id, team: '' };
      return `
      <tr id="prow-${idx}" style="border-bottom:1px solid var(--border)">
        <td style="padding:10px 8px;font-weight:500">${emp.name}<br><span style="font-size:11px;color:var(--muted)">${emp.team}</span></td>
        <td style="padding:10px 8px;text-align:center">${p.full_days}</td>
        <td style="padding:10px 8px;text-align:center">${p.half_days}</td>
        <td style="padding:10px 8px;text-align:center"><strong>${p.TD}</strong></td>
        <td style="padding:10px 8px;text-align:right">${fmt(p.gross_pay)}</td>
        <td style="padding:10px 8px;text-align:right;color:var(--success)">${p.ot_earnings > 0 ? '+'+fmt(p.ot_earnings) : '—'}</td>
        <td style="padding:10px 8px;text-align:right;color:var(--danger)">${p.shortfall_deduction > 0 ? '-'+fmt(p.shortfall_deduction) : '—'}</td>
        <td style="padding:10px 8px;text-align:right;color:var(--warning)">${p.total_advances > 0 ? '-'+fmt(p.total_advances) : '—'}</td>
        <td style="padding:10px 8px;text-align:right;color:var(--warning)">${p.bonus_cut > 0 ? '-'+fmt(p.bonus_cut) : '—'}</td>
        <td style="padding:10px 8px;text-align:right;font-weight:700">${fmt(p.net_pay)}</td>
        <td style="padding:10px 8px;text-align:center"><span class="badge badge-${p.status==='finalized'?'success':'warning'}">${p.status}</span></td>
        <td style="padding:10px 8px;text-align:center">
          <button class="btn btn-secondary btn-sm" onclick="Views.Payroll.editRow(${idx})" ${p.status==='finalized'?'disabled':''}>Edit</button>
        </td>
      </tr>
      <tr id="prow-edit-${idx}" class="hidden" style="background:var(--surface2)">
        <td colspan="12" style="padding:16px 12px">
          ${this.renderEditPanel(p, idx)}
        </td>
      </tr>`;
    }).join('');

    wrap.innerHTML = `
    <div style="overflow-x:auto">
    <table style="width:100%;border-collapse:collapse;font-size:13px">
      <thead style="background:var(--surface2)">
        <tr>
          <th style="padding:10px 8px;text-align:left">Employee</th>
          <th style="padding:10px 8px">FD</th><th style="padding:10px 8px">HD</th>
          <th style="padding:10px 8px">TD</th>
          <th style="padding:10px 8px;text-align:right">Gross</th>
          <th style="padding:10px 8px;text-align:right">OT</th>
          <th style="padding:10px 8px;text-align:right">Shortfall</th>
          <th style="padding:10px 8px;text-align:right">Advance</th>
          <th style="padding:10px 8px;text-align:right">Bonus↓</th>
          <th style="padding:10px 8px;text-align:right">Net Pay</th>
          <th style="padding:10px 8px">Status</th>
          <th style="padding:10px 8px"></th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
    </div>
    <div style="margin-top:12px;font-size:13px;color:var(--muted)">
      ${payroll.length} employees &nbsp;|&nbsp;
      Total net: <strong>${fmt(payroll.reduce((s,p)=>s+p.net_pay,0))}</strong>
    </div>`;
  },

  renderEditPanel(p, idx) {
    return `
    <div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(160px,1fr));gap:12px">
      <div class="field"><label>Full Days</label><input type="number" id="edit-fd-${idx}" value="${p.full_days}" min="0" /></div>
      <div class="field"><label>Half Days</label><input type="number" id="edit-hd-${idx}" value="${p.half_days}" min="0" /></div>
      <div class="field"><label>Holiday (paid absent)</label><input type="number" id="edit-hol-${idx}" value="${p.holiday_absent_days}" min="0" /></div>
      <div class="field"><label>OT Weekday (min)</label><input type="number" id="edit-otw-${idx}" value="${p.ot_weekday_min}" min="0" /></div>
      <div class="field"><label>OT Sunday (min)</label><input type="number" id="edit-ots-${idx}" value="${p.ot_sunday_min}" min="0" /></div>
      <div class="field"><label>Shortfall (min)</label><input type="number" id="edit-sf-${idx}" value="${p.shortfall_min}" min="0" /></div>
      <div class="field"><label>Advances</label><input type="number" id="edit-adv-${idx}" value="${p.total_advances}" min="0" /></div>
      <div class="field"><label>Bonus cut</label><input type="number" id="edit-bon-${idx}" value="${p.bonus_cut}" min="0" /></div>
      <div class="field"><label>Notes</label><input type="text" id="edit-note-${idx}" value="${p.notes||''}" /></div>
    </div>
    <div style="margin-top:12px;display:flex;gap:8px">
      <button class="btn btn-primary btn-sm" onclick="Views.Payroll.applyEdit(${idx})">Recalculate &amp; Apply</button>
      <button class="btn btn-secondary btn-sm" onclick="Views.Payroll.closeEdit(${idx})">Cancel</button>
    </div>`;
  },

  editRow(idx) {
    const editRow = document.getElementById('prow-edit-' + idx);
    const isOpen  = !editRow.classList.contains('hidden');
    // Close all
    document.querySelectorAll('[id^="prow-edit-"]').forEach(r => r.classList.add('hidden'));
    if (!isOpen) editRow.classList.remove('hidden');
  },

  closeEdit(idx) { document.getElementById('prow-edit-' + idx)?.classList.add('hidden'); },

  applyEdit(idx) {
    const p = State.payroll[idx];
    const g = id => parseFloat(document.getElementById(id)?.value || 0);
    p.full_days           = g(`edit-fd-${idx}`);
    p.half_days           = g(`edit-hd-${idx}`);
    p.holiday_absent_days = g(`edit-hol-${idx}`);
    p.ot_weekday_min      = g(`edit-otw-${idx}`);
    p.ot_sunday_min       = g(`edit-ots-${idx}`);
    p.shortfall_min       = g(`edit-sf-${idx}`);
    p.total_advances      = g(`edit-adv-${idx}`);
    p.bonus_cut           = g(`edit-bon-${idx}`);
    p.notes               = document.getElementById(`edit-note-${idx}`)?.value || '';

    const emp = State.employees.find(e => e.emp_id === p.emp_id);
    if (emp) {
      const daily  = emp.weekly_salary / 6;
      const hourly = daily / 9;
      p.TD              = p.full_days + p.half_days + p.holiday_absent_days;
      p.gross_pay       = Math.round(p.TD * daily);
      p.ot_earnings     = Math.round((p.ot_weekday_min + p.ot_sunday_min + p.ot_holiday_min) / 60 * hourly);
      p.shortfall_deduction = Math.round(p.shortfall_min / 60 * hourly);
      p.net_pay         = p.gross_pay + p.ot_earnings - p.shortfall_deduction - p.bonus_cut - p.total_advances;
    }
    this.renderTable(State.payroll);
    toast('Row updated', 'success');
  },

  async saveDraft() {
    if (!State.payroll.length) { toast('Nothing to save', 'error'); return; }
    try {
      await API.call('savePayroll', State.payroll);
      toast('Draft saved', 'success');
    } catch(e) { toast(e.message, 'error'); }
  },

  async loadDraft() {
    try {
      const [holRes, advRes, payRes] = await Promise.all([
        API.call('getHolidays', { year: currentYear() }),
        API.call('getAdvances', { month: State.currentMonth }),
        API.call('getPayroll', { month: State.currentMonth })
      ]);
      State.holidays = holRes.data || [];
      State.advances = advRes.data || [];
      State.payroll  = payRes.data || [];
      this.renderTable(State.payroll);
      toast(`Loaded ${State.payroll.length} records`, 'success');
    } catch(e) { toast(e.message, 'error'); }
  },

  async finalize() {
    const month = State.currentMonth;
    if (!confirm(`Finalise payroll for ${month}? This cannot be undone.`)) return;
    try {
      // Save current state first
      await API.call('savePayroll', State.payroll);
      const r = await API.call('finalizePayroll', { month });
      toast(`Finalised ${r.finalized} records`, 'success');
      await this.loadDraft();
    } catch(e) { toast(e.message, 'error'); }
  }
};

// ============================================================
// VIEWS — EMPLOYEES
// ============================================================
Views.Employees = {
  async load() {
    const c = document.getElementById('view-employees');
    c.innerHTML = `<div style="display:flex;align-items:center;gap:10px;margin-bottom:16px;flex-wrap:wrap">
      <h2 style="font-size:18px;font-weight:700">Employees</h2>
      <input type="search" id="emp-search" placeholder="Search…" style="max-width:200px" oninput="Views.Employees.filter()" />
      <select id="emp-team" style="width:auto" onchange="Views.Employees.filter()">
        <option value="">All Teams</option><option>Emb</option><option>Stitching</option>
      </select>
      <span style="flex:1"></span>
      <button class="btn btn-primary btn-sm" onclick="Views.Employees.openForm()">+ Add Employee</button>
    </div>
    <div id="emp-form-wrap"></div>
    <div id="emp-list"></div>`;
    this.render();
  },

  filter() {
    const q    = document.getElementById('emp-search')?.value.toLowerCase() || '';
    const team = document.getElementById('emp-team')?.value || '';
    this._filtered = State.employees.filter(e =>
      e.status === 'Active' &&
      (!q || e.name.toLowerCase().includes(q) || e.emp_id.toLowerCase().includes(q)) &&
      (!team || e.team === team)
    );
    this.renderList(this._filtered);
  },

  render() {
    this._filtered = State.employees.filter(e => e.status === 'Active');
    this.renderList(this._filtered);
  },

  renderList(list) {
    const c = document.getElementById('emp-list');
    if (!list.length) { c.innerHTML = '<p style="color:var(--muted)">No employees found.</p>'; return; }
    c.innerHTML = `<div style="overflow-x:auto"><table style="width:100%;border-collapse:collapse;font-size:13px">
      <thead style="background:var(--surface2)">
        <tr>
          <th style="padding:10px 8px;text-align:left">ID</th>
          <th style="padding:10px 8px;text-align:left">Name</th>
          <th style="padding:10px 8px">Team</th>
          <th style="padding:10px 8px">Petpooja ID</th>
          <th style="padding:10px 8px;text-align:right">Weekly ₹</th>
          <th style="padding:10px 8px"></th>
        </tr>
      </thead>
      <tbody>${list.map(e => `
        <tr style="border-bottom:1px solid var(--border)">
          <td style="padding:10px 8px;color:var(--muted)">${e.emp_id}</td>
          <td style="padding:10px 8px;font-weight:500">${e.name}<br><span style="font-size:11px;color:var(--muted)">${e.designation||''}</span></td>
          <td style="padding:10px 8px;text-align:center"><span class="badge badge-primary">${e.team||'—'}</span></td>
          <td style="padding:10px 8px;text-align:center">${e.petpooja_id||'—'}</td>
          <td style="padding:10px 8px;text-align:right">${fmt(e.weekly_salary)}</td>
          <td style="padding:10px 8px;text-align:right;display:flex;gap:4px;justify-content:flex-end">
            <button class="btn btn-secondary btn-sm" onclick="Views.Employees.openForm('${e.emp_id}')">Edit</button>
            <button class="btn btn-danger btn-sm" onclick="Views.Employees.remove('${e.emp_id}')">✕</button>
          </td>
        </tr>`).join('')}
      </tbody></table></div>
      <p style="font-size:13px;color:var(--muted);margin-top:8px">${list.length} active employee(s)</p>`;
  },

  openForm(empId) {
    const emp = empId ? State.employees.find(e => e.emp_id === empId) : {};
    const c = document.getElementById('emp-form-wrap');
    c.innerHTML = `<div class="card" style="margin-bottom:16px">
      <h3 style="font-size:15px;margin-bottom:14px">${empId ? 'Edit' : 'Add'} Employee</h3>
      <div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(180px,1fr));gap:12px">
        <div class="field"><label>Name *</label><input id="ef-name" value="${emp.name||''}" /></div>
        <div class="field"><label>Team</label>
          <select id="ef-team"><option value="">—</option>
            <option ${emp.team==='Emb'?'selected':''}>Emb</option>
            <option ${emp.team==='Stitching'?'selected':''}>Stitching</option>
          </select></div>
        <div class="field"><label>Designation</label><input id="ef-desig" value="${emp.designation||''}" /></div>
        <div class="field"><label>Petpooja ID</label><input id="ef-pid" value="${emp.petpooja_id||''}" /></div>
        <div class="field"><label>Petpooja Name (if different)</label><input id="ef-pname" value="${emp.petpooja_name||''}" /></div>
        <div class="field"><label>Weekly Salary (₹) *</label><input id="ef-salary" type="number" value="${emp.weekly_salary||''}" /></div>
        <div class="field"><label>Joining Date</label><input id="ef-join" type="date" value="${emp.joining_date||''}" /></div>
      </div>
      <div style="display:flex;gap:8px;margin-top:4px">
        <button class="btn btn-primary btn-sm" onclick="Views.Employees.save('${empId||''}')">Save</button>
        <button class="btn btn-secondary btn-sm" onclick="Views.Employees.closeForm()">Cancel</button>
      </div>
    </div>`;
  },

  closeForm() { document.getElementById('emp-form-wrap').innerHTML = ''; },

  async save(empId) {
    const name   = document.getElementById('ef-name')?.value.trim();
    const salary = parseFloat(document.getElementById('ef-salary')?.value);
    if (!name || !salary) { toast('Name and salary required', 'error'); return; }
    const data = {
      emp_id: empId || null, name,
      team:         document.getElementById('ef-team')?.value,
      designation:  document.getElementById('ef-desig')?.value,
      petpooja_id:  document.getElementById('ef-pid')?.value.trim(),
      petpooja_name:document.getElementById('ef-pname')?.value.trim() || name,
      weekly_salary: salary,
      joining_date: document.getElementById('ef-join')?.value,
      status: 'Active'
    };
    try {
      const r = await API.call('saveEmployee', data);
      const res2 = await API.call('getEmployees', {});
      State.employees = res2.data || [];
      this.closeForm(); this.render();
      toast('Saved ' + name, 'success');
    } catch(e) { toast(e.message, 'error'); }
  },

  async remove(empId) {
    if (!confirm('Mark employee as Inactive?')) return;
    try {
      await API.call('deleteEmployee', { emp_id: empId });
      const r = await API.call('getEmployees', {});
      State.employees = r.data || [];
      this.render(); toast('Employee deactivated', 'success');
    } catch(e) { toast(e.message, 'error'); }
  }
};

// ============================================================
// VIEWS — ADVANCES
// ============================================================
Views.Advances = {
  async load() {
    const c = document.getElementById('view-advances');
    c.innerHTML = `<div style="display:flex;align-items:center;gap:10px;margin-bottom:16px;flex-wrap:wrap">
      <h2 style="font-size:18px;font-weight:700">Advances</h2>
      <input type="month" id="adv-month" value="${State.currentMonth}" style="width:auto;padding:6px 10px"
        onchange="State.currentMonth=this.value;Views.Advances.fetch()" />
      <span style="flex:1"></span>
      <button class="btn btn-primary btn-sm" onclick="Views.Advances.openForm()">+ Add Advance</button>
    </div>
    <div id="adv-form-wrap"></div>
    <div id="adv-list"></div>`;
    await this.fetch();
  },

  async fetch() {
    try {
      const r = await API.call('getAdvances', { month: State.currentMonth });
      State.advances = r.data || [];
      this.renderList();
    } catch(e) { toast(e.message, 'error'); }
  },

  renderList() {
    const c = document.getElementById('adv-list');
    const list = State.advances;
    if (!list.length) { c.innerHTML = '<p style="color:var(--muted)">No advances for this month.</p>'; return; }

    // Group by employee
    const byEmp = {};
    list.forEach(a => { (byEmp[a.emp_id] = byEmp[a.emp_id] || []).push(a); });

    c.innerHTML = `<div style="overflow-x:auto"><table style="width:100%;border-collapse:collapse;font-size:13px">
      <thead style="background:var(--surface2)">
        <tr>
          <th style="padding:10px 8px;text-align:left">Date</th>
          <th style="padding:10px 8px;text-align:left">Employee</th>
          <th style="padding:10px 8px;text-align:right">Amount</th>
          <th style="padding:10px 8px">Mode</th>
          <th style="padding:10px 8px">Notes</th>
          <th style="padding:10px 8px">By</th>
          <th style="padding:10px 8px"></th>
        </tr>
      </thead>
      <tbody>${list.sort((a,b)=>a.date>b.date?-1:1).map(a => `
        <tr style="border-bottom:1px solid var(--border)">
          <td style="padding:10px 8px">${a.date}</td>
          <td style="padding:10px 8px;font-weight:500">${a.emp_name}</td>
          <td style="padding:10px 8px;text-align:right;font-weight:600">${fmt(a.amount)}</td>
          <td style="padding:10px 8px;text-align:center"><span class="badge badge-primary">${a.mode}</span></td>
          <td style="padding:10px 8px;color:var(--muted)">${a.notes||'—'}</td>
          <td style="padding:10px 8px;color:var(--muted)">${a.created_by||'—'}</td>
          <td style="padding:10px 8px;display:flex;gap:4px">
            <button class="btn btn-secondary btn-sm" onclick="Views.Advances.openForm('${a.advance_id}')">Edit</button>
            <button class="btn btn-danger btn-sm" onclick="Views.Advances.remove('${a.advance_id}')">✕</button>
          </td>
        </tr>`).join('')}
      </tbody></table></div>
      <p style="font-size:13px;color:var(--muted);margin-top:8px">
        Total advances: <strong>${fmt(list.reduce((s,a)=>s+a.amount,0))}</strong>
      </p>`;
  },

  openForm(advId) {
    const adv = advId ? State.advances.find(a => a.advance_id === advId) : {};
    const c = document.getElementById('adv-form-wrap');
    c.innerHTML = `<div class="card" style="margin-bottom:16px">
      <h3 style="font-size:15px;margin-bottom:14px">${advId ? 'Edit' : 'Add'} Advance</h3>
      <div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(180px,1fr));gap:12px">
        <div class="field"><label>Employee *</label>
          <select id="af-emp">
            <option value="">— select —</option>
            ${State.employees.filter(e=>e.status==='Active').map(e =>
              `<option value="${e.emp_id}" ${adv.emp_id===e.emp_id?'selected':''}>${e.name}</option>`).join('')}
          </select></div>
        <div class="field"><label>Date *</label><input type="date" id="af-date" value="${adv.date||new Date().toISOString().split('T')[0]}" /></div>
        <div class="field"><label>Amount (₹) *</label><input type="number" id="af-amt" value="${adv.amount||''}" /></div>
        <div class="field"><label>Mode</label>
          <select id="af-mode">
            <option ${adv.mode==='Cash'?'selected':''}>Cash</option>
            <option ${adv.mode==='GPay'?'selected':''}>GPay</option>
          </select></div>
        <div class="field"><label>Deduct from Month</label>
          <input type="month" id="af-month" value="${adv.month||State.currentMonth}" /></div>
        <div class="field"><label>Notes</label><input id="af-notes" value="${adv.notes||''}" /></div>
      </div>
      <div style="display:flex;gap:8px;margin-top:4px">
        <button class="btn btn-primary btn-sm" onclick="Views.Advances.save('${advId||''}')">Save</button>
        <button class="btn btn-secondary btn-sm" onclick="Views.Advances.closeForm()">Cancel</button>
      </div>
    </div>`;
  },

  closeForm() { document.getElementById('adv-form-wrap').innerHTML = ''; },

  async save(advId) {
    const empId = document.getElementById('af-emp')?.value;
    const amt   = parseFloat(document.getElementById('af-amt')?.value);
    if (!empId || !amt) { toast('Employee and amount required', 'error'); return; }
    const emp  = State.employees.find(e => e.emp_id === empId);
    const data = {
      advance_id: advId || null, emp_id: empId, emp_name: emp?.name || empId,
      date:       document.getElementById('af-date')?.value,
      amount:     amt,
      mode:       document.getElementById('af-mode')?.value,
      month:      document.getElementById('af-month')?.value,
      notes:      document.getElementById('af-notes')?.value,
      created_by: 'Admin'
    };
    try {
      await API.call('saveAdvance', data);
      this.closeForm(); await this.fetch();
      toast('Advance saved', 'success');
    } catch(e) { toast(e.message, 'error'); }
  },

  async remove(advId) {
    if (!confirm('Delete this advance?')) return;
    try {
      await API.call('deleteAdvance', { advance_id: advId });
      await this.fetch(); toast('Deleted', 'success');
    } catch(e) { toast(e.message, 'error'); }
  }
};

// ============================================================
// VIEWS — BONUS POOL
// ============================================================
Views.Bonus = {
  async load() {
    const c = document.getElementById('view-bonus');
    c.innerHTML = `<h2 style="font-size:18px;font-weight:700;margin-bottom:16px">Bonus Pool</h2><div id="bonus-list"><div class="spinner"></div></div>`;
    try {
      const r = await API.call('getBonusPool', {});
      const list = r.data || [];
      const target = parseInt(State.settings.BONUS_POOL_TARGET || 24000);
      if (!list.length) { document.getElementById('bonus-list').innerHTML = '<p style="color:var(--muted)">No bonus pool data yet.</p>'; return; }
      document.getElementById('bonus-list').innerHTML = `
        <div style="overflow-x:auto"><table style="width:100%;border-collapse:collapse;font-size:13px">
          <thead style="background:var(--surface2)">
            <tr><th style="padding:10px 8px;text-align:left">Employee</th>
            <th style="padding:10px 8px;text-align:right">Balance</th>
            <th style="padding:10px 8px">Progress</th>
            <th style="padding:10px 8px;text-align:right">Last Payout</th>
            <th style="padding:10px 8px">Last Updated</th></tr>
          </thead>
          <tbody>${list.sort((a,b)=>b.balance-a.balance).map(b => {
            const pct = Math.min(100, Math.round(b.balance/target*100));
            return `<tr style="border-bottom:1px solid var(--border)">
              <td style="padding:10px 8px;font-weight:500">${b.emp_name}</td>
              <td style="padding:10px 8px;text-align:right;font-weight:700">${fmt(b.balance)}</td>
              <td style="padding:10px 8px;min-width:120px">
                <div style="background:var(--border);border-radius:99px;height:8px">
                  <div style="background:${pct>=100?'var(--success)':'var(--primary)'};width:${pct}%;height:8px;border-radius:99px"></div>
                </div>
                <span style="font-size:11px;color:var(--muted)">${pct}% of ${fmt(target)}</span>
              </td>
              <td style="padding:10px 8px;text-align:right">${b.last_payout>0?fmt(b.last_payout):'—'}</td>
              <td style="padding:10px 8px;color:var(--muted)">${b.last_updated_month||'—'}</td>
            </tr>`;
          }).join('')}
          </tbody></table></div>`;
    } catch(e) { toast(e.message, 'error'); }
  }
};

// ============================================================
// VIEWS — PAYMENTS
// ============================================================
Views.Payments = {
  async load() {
    const c = document.getElementById('view-payments');
    c.innerHTML = `<div style="display:flex;align-items:center;gap:10px;margin-bottom:16px;flex-wrap:wrap">
      <h2 style="font-size:18px;font-weight:700">Payments</h2>
      <input type="month" id="pay-month" value="${State.currentMonth}" style="width:auto;padding:6px 10px"
        onchange="State.currentMonth=this.value;Views.Payments.fetch()" />
      <span style="flex:1"></span>
      <button class="btn btn-primary btn-sm" onclick="Views.Payments.openForm()">+ Record Payment</button>
    </div>
    <div id="pay-form-wrap"></div>
    <div id="pay-summary" style="margin-bottom:16px"></div>
    <div id="pay-list"></div>`;
    await this.fetch();
  },

  async fetch() {
    try {
      const [payRes, payrollRes] = await Promise.all([
        API.call('getPayments', { month: State.currentMonth }),
        API.call('getPayroll', { month: State.currentMonth })
      ]);
      State.payments = payRes.data || [];
      this.renderSummary(payRes.data || [], payrollRes.data || []);
      this.renderList(payRes.data || []);
    } catch(e) { toast(e.message, 'error'); }
  },

  renderSummary(payments, payroll) {
    const c = document.getElementById('pay-summary');
    const totalNet = payroll.reduce((s,p)=>s+Number(p.net_pay),0);
    const totalPaid = payments.reduce((s,p)=>s+Number(p.amount),0);
    const balance = totalNet - totalPaid;
    c.innerHTML = `<div style="display:flex;gap:12px;flex-wrap:wrap">
      <div class="card" style="flex:1;min-width:140px;text-align:center">
        <div style="font-size:12px;color:var(--muted)">Total Payroll</div>
        <div style="font-size:20px;font-weight:700">${fmt(totalNet)}</div>
      </div>
      <div class="card" style="flex:1;min-width:140px;text-align:center">
        <div style="font-size:12px;color:var(--muted)">Total Paid</div>
        <div style="font-size:20px;font-weight:700;color:var(--success)">${fmt(totalPaid)}</div>
      </div>
      <div class="card" style="flex:1;min-width:140px;text-align:center">
        <div style="font-size:12px;color:var(--muted)">Remaining</div>
        <div style="font-size:20px;font-weight:700;color:${balance>0?'var(--danger)':'var(--success)'}">${fmt(Math.abs(balance))}</div>
      </div>
    </div>`;
  },

  renderList(list) {
    const c = document.getElementById('pay-list');
    if (!list.length) { c.innerHTML = '<p style="color:var(--muted);margin-top:8px">No payments recorded yet.</p>'; return; }
    c.innerHTML = `<div style="overflow-x:auto;margin-top:12px"><table style="width:100%;border-collapse:collapse;font-size:13px">
      <thead style="background:var(--surface2)">
        <tr><th style="padding:10px 8px;text-align:left">Date</th><th style="padding:10px 8px;text-align:left">Employee</th>
        <th style="padding:10px 8px;text-align:right">Amount</th><th style="padding:10px 8px">Mode</th>
        <th style="padding:10px 8px">Notes</th><th></th></tr>
      </thead>
      <tbody>${list.sort((a,b)=>a.date_paid>b.date_paid?-1:1).map(p=>`
        <tr style="border-bottom:1px solid var(--border)">
          <td style="padding:10px 8px">${p.date_paid}</td>
          <td style="padding:10px 8px;font-weight:500">${p.emp_name}</td>
          <td style="padding:10px 8px;text-align:right;font-weight:600">${fmt(p.amount)}</td>
          <td style="padding:10px 8px;text-align:center"><span class="badge badge-primary">${p.mode}</span></td>
          <td style="padding:10px 8px;color:var(--muted)">${p.notes||'—'}</td>
          <td style="padding:10px 8px;display:flex;gap:4px">
            <button class="btn btn-danger btn-sm" onclick="Views.Payments.remove('${p.payment_id}')">✕</button>
          </td>
        </tr>`).join('')}
      </tbody></table></div>`;
  },

  openForm() {
    const c = document.getElementById('pay-form-wrap');
    c.innerHTML = `<div class="card" style="margin-bottom:16px">
      <h3 style="font-size:15px;margin-bottom:14px">Record Payment</h3>
      <div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(180px,1fr));gap:12px">
        <div class="field"><label>Employee *</label>
          <select id="pf-emp">
            <option value="">— select —</option>
            ${State.employees.filter(e=>e.status==='Active').map(e=>
              `<option value="${e.emp_id}">${e.name}</option>`).join('')}
          </select></div>
        <div class="field"><label>Date Paid *</label><input type="date" id="pf-date" value="${new Date().toISOString().split('T')[0]}" /></div>
        <div class="field"><label>Amount (₹) *</label><input type="number" id="pf-amt" /></div>
        <div class="field"><label>Mode</label>
          <select id="pf-mode"><option>Cash</option><option>GPay</option></select></div>
        <div class="field"><label>Notes</label><input id="pf-notes" /></div>
      </div>
      <div style="display:flex;gap:8px;margin-top:4px">
        <button class="btn btn-primary btn-sm" onclick="Views.Payments.save()">Save</button>
        <button class="btn btn-secondary btn-sm" onclick="document.getElementById('pay-form-wrap').innerHTML=''">Cancel</button>
      </div>
    </div>`;
  },

  async save() {
    const empId = document.getElementById('pf-emp')?.value;
    const amt   = parseFloat(document.getElementById('pf-amt')?.value);
    if (!empId || !amt) { toast('Employee and amount required', 'error'); return; }
    const emp = State.employees.find(e => e.emp_id === empId);
    try {
      await API.call('savePayment', {
        emp_id: empId, emp_name: emp?.name || empId,
        month: State.currentMonth,
        date_paid: document.getElementById('pf-date')?.value,
        amount: amt, mode: document.getElementById('pf-mode')?.value,
        notes: document.getElementById('pf-notes')?.value
      });
      document.getElementById('pay-form-wrap').innerHTML = '';
      await this.fetch(); toast('Payment recorded', 'success');
    } catch(e) { toast(e.message, 'error'); }
  },

  async remove(paymentId) {
    if (!confirm('Delete this payment?')) return;
    try {
      await API.call('deletePayment', { payment_id: paymentId });
      await this.fetch(); toast('Deleted', 'success');
    } catch(e) { toast(e.message, 'error'); }
  }
};

// ============================================================
// VIEWS — SETTINGS
// ============================================================
Views.Settings = {
  async load() {
    const c = document.getElementById('view-settings');
    c.innerHTML = `<h2 style="font-size:18px;font-weight:700;margin-bottom:16px">Settings</h2>
    <div style="display:grid;gap:16px;max-width:700px">
      <div class="card">
        <h3 style="font-size:15px;margin-bottom:14px">General</h3>
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px">
          <div class="field"><label>Company Name</label><input id="s-company" value="${State.settings.COMPANY_NAME||''}" /></div>
          <div class="field"><label>Admin PIN</label><input type="password" id="s-adminpin" value="${State.settings.ADMIN_PIN||''}" /></div>
          <div class="field"><label>Manager PIN</label><input type="password" id="s-manpin" value="${State.settings.MANAGER_PIN||''}" /></div>
        </div>
      </div>
      <div class="card">
        <h3 style="font-size:15px;margin-bottom:14px">Calculation Rules</h3>
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px">
          <div class="field"><label>Shortfall threshold (hours)</label><input type="number" step="0.25" id="s-thresh" value="${State.settings.SHORTFALL_THRESHOLD_HOURS||8.5}" /></div>
          <div class="field"><label>OT minimum hours</label><input type="number" id="s-otmin" value="${State.settings.OT_MIN_HOURS||10}" /></div>
          <div class="field"><label>OT rounding (minutes)</label><input type="number" id="s-otround" value="${State.settings.OT_ROUND_MINUTES||15}" /></div>
          <div class="field"><label>Add lunch hour to Sunday/Holiday OT?</label>
            <select id="s-lunch">
              <option value="true" ${State.settings.ADD_LUNCH_TO_OT==='true'?'selected':''}>Yes</option>
              <option value="false" ${State.settings.ADD_LUNCH_TO_OT!=='true'?'selected':''}>No</option>
            </select></div>
          <div class="field"><label>Lunch duration (hours)</label><input type="number" step="0.5" id="s-lunchhrs" value="${State.settings.LUNCH_DURATION_HOURS||1}" /></div>
        </div>
      </div>
      <div class="card">
        <h3 style="font-size:15px;margin-bottom:14px">Bonus Pool</h3>
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px">
          <div class="field"><label>Pool target (₹)</label><input type="number" id="s-bontarget" value="${State.settings.BONUS_POOL_TARGET||24000}" /></div>
          <div class="field"><label>Employee contribution / month (₹)</label><input type="number" id="s-empcon" value="${State.settings.EMPLOYEE_CONTRIBUTION||500}" /></div>
          <div class="field"><label>Company contribution / month (₹)</label><input type="number" id="s-comcon" value="${State.settings.COMPANY_CONTRIBUTION||500}" /></div>
          <div class="field"><label>Min days for bonus eligibility</label><input type="number" id="s-bondays" value="${State.settings.BONUS_ELIGIBILITY_MIN_DAYS||6}" /></div>
        </div>
      </div>
      <div class="card">
        <h3 style="font-size:15px;margin-bottom:14px">Holidays ${new Date().getFullYear()}</h3>
        <div id="holiday-list"></div>
        <button class="btn btn-secondary btn-sm" style="margin-top:8px" onclick="Views.Settings.addHolidayRow()">+ Add Holiday</button>
      </div>
      <button class="btn btn-primary" onclick="Views.Settings.save()">Save All Settings</button>
    </div>`;
    await this.loadHolidays();
  },

  async loadHolidays() {
    try {
      const r = await API.call('getHolidays', { year: currentYear() });
      State.holidays = r.data || [];
      this.renderHolidays();
    } catch(e) {}
  },

  renderHolidays() {
    const c = document.getElementById('holiday-list');
    const rows = State.holidays.length ? State.holidays : Array(8).fill({ date: '', name: '' });
    c.innerHTML = rows.map((h, i) => `
      <div style="display:flex;gap:8px;margin-bottom:8px" id="hol-row-${i}">
        <input type="date" id="hol-date-${i}" value="${h.date||''}" style="flex:0 0 160px" />
        <input type="text" id="hol-name-${i}" value="${h.name||''}" placeholder="Holiday name" style="flex:1" />
        <button class="btn btn-danger btn-sm" onclick="document.getElementById('hol-row-${i}').remove()">✕</button>
      </div>`).join('');
  },

  addHolidayRow() {
    const c = document.getElementById('holiday-list');
    const i = Date.now();
    const div = document.createElement('div');
    div.id = 'hol-row-' + i;
    div.style.cssText = 'display:flex;gap:8px;margin-bottom:8px';
    div.innerHTML = `<input type="date" id="hol-date-${i}" style="flex:0 0 160px" />
      <input type="text" id="hol-name-${i}" placeholder="Holiday name" style="flex:1" />
      <button class="btn btn-danger btn-sm" onclick="this.parentElement.remove()">✕</button>`;
    c.appendChild(div);
  },

  async save() {
    const settings = {
      COMPANY_NAME:                document.getElementById('s-company')?.value,
      ADMIN_PIN:                   document.getElementById('s-adminpin')?.value,
      MANAGER_PIN:                 document.getElementById('s-manpin')?.value,
      SHORTFALL_THRESHOLD_HOURS:   document.getElementById('s-thresh')?.value,
      OT_MIN_HOURS:                document.getElementById('s-otmin')?.value,
      OT_ROUND_MINUTES:            document.getElementById('s-otround')?.value,
      ADD_LUNCH_TO_OT:             document.getElementById('s-lunch')?.value,
      LUNCH_DURATION_HOURS:        document.getElementById('s-lunchhrs')?.value,
      BONUS_POOL_TARGET:           document.getElementById('s-bontarget')?.value,
      EMPLOYEE_CONTRIBUTION:       document.getElementById('s-empcon')?.value,
      COMPANY_CONTRIBUTION:        document.getElementById('s-comcon')?.value,
      BONUS_ELIGIBILITY_MIN_DAYS:  document.getElementById('s-bondays')?.value,
    };

    // Collect holidays
    const holidays = [];
    document.querySelectorAll('[id^="hol-row-"]').forEach(row => {
      const i    = row.id.replace('hol-row-', '');
      const date = document.getElementById('hol-date-' + i)?.value;
      const name = document.getElementById('hol-name-' + i)?.value;
      if (date && name) holidays.push({ date, name });
    });

    try {
      await Promise.all([
        API.call('saveSettings', settings),
        holidays.length ? API.call('saveHolidays', holidays) : Promise.resolve()
      ]);
      const r = await API.call('getSettings', {});
      State.settings = r.data || {};
      toast('Settings saved', 'success');
    } catch(e) { toast(e.message, 'error'); }
  }
};

// ============================================================
// MANAGER APP
// ============================================================
const ManagerApp = {
  init() {
    show('screen-manager');
    const c = document.getElementById('screen-manager');
    c.innerHTML = `
    <header style="background:var(--surface);border-bottom:1px solid var(--border);padding:0 16px;display:flex;align-items:center;gap:12px;height:52px">
      <span style="font-weight:700;font-size:16px">MP&#x2009;SAL &mdash; Manager</span>
      <span style="flex:1"></span>
      <button onclick="App.toggleTheme()" style="border:1px solid var(--border);background:none;padding:5px 10px;border-radius:6px;cursor:pointer;font-size:13px;color:var(--muted)">Theme</button>
      <button onclick="Auth.logout()" style="border:1px solid var(--border);background:none;padding:5px 10px;border-radius:6px;cursor:pointer;font-size:13px;color:var(--danger)">Logout</button>
    </header>
    <main style="padding:16px;max-width:600px;margin:0 auto">
      <h2 style="font-size:18px;font-weight:700;margin-bottom:16px">Submit Advance</h2>
      <div class="card" style="margin-bottom:16px">
        <div class="field"><label>Employee *</label>
          <select id="m-emp">
            <option value="">— select employee —</option>
            ${State.employees.filter(e=>e.status==='Active').sort((a,b)=>a.name.localeCompare(b.name)).map(e=>
              `<option value="${e.emp_id}">${e.name}</option>`).join('')}
          </select></div>
        <div class="field"><label>Date *</label><input type="date" id="m-date" value="${new Date().toISOString().split('T')[0]}" /></div>
        <div class="field"><label>Amount (₹) *</label><input type="number" id="m-amt" placeholder="0" /></div>
        <div class="field"><label>Mode</label>
          <select id="m-mode"><option>Cash</option><option>GPay</option></select></div>
        <div class="field"><label>Deduct from Month</label>
          <input type="month" id="m-month" value="${State.currentMonth}" /></div>
        <div class="field"><label>Notes</label><input id="m-notes" placeholder="Optional" /></div>
        <button class="btn btn-primary" style="width:100%" onclick="ManagerApp.submit()">Submit Advance</button>
      </div>
      <h3 style="font-size:15px;font-weight:600;margin-bottom:10px">Today's Advances</h3>
      <div id="m-recent"></div>
    </main>`;
    this.loadRecent();
  },

  async submit() {
    const empId = document.getElementById('m-emp')?.value;
    const amt   = parseFloat(document.getElementById('m-amt')?.value);
    if (!empId || !amt) { toast('Employee and amount required', 'error'); return; }
    const emp = State.employees.find(e => e.emp_id === empId);
    try {
      await API.call('saveAdvance', {
        emp_id: empId, emp_name: emp?.name || empId,
        date:   document.getElementById('m-date')?.value,
        amount: amt,
        mode:   document.getElementById('m-mode')?.value,
        month:  document.getElementById('m-month')?.value,
        notes:  document.getElementById('m-notes')?.value,
        created_by: 'Manager'
      });
      document.getElementById('m-amt').value = '';
      document.getElementById('m-notes').value = '';
      await this.loadRecent();
      toast('Advance submitted for ' + emp?.name, 'success');
    } catch(e) { toast(e.message, 'error'); }
  },

  async loadRecent() {
    const today = new Date().toISOString().split('T')[0];
    try {
      const r = await API.call('getAdvances', { month: State.currentMonth });
      const todayList = (r.data || []).filter(a => a.date === today && a.created_by === 'Manager');
      const c = document.getElementById('m-recent');
      if (!todayList.length) { c.innerHTML = '<p style="color:var(--muted);font-size:14px">None yet today.</p>'; return; }
      c.innerHTML = todayList.map(a => `
        <div class="card" style="margin-bottom:8px;display:flex;align-items:center;gap:12px">
          <div style="flex:1"><strong>${a.emp_name}</strong><br><span style="font-size:12px;color:var(--muted)">${a.mode} · ${a.notes||'—'}</span></div>
          <div style="font-size:18px;font-weight:700">${fmt(a.amount)}</div>
        </div>`).join('');
    } catch(e) {}
  }
};

// ---- Boot ----
App.init();
