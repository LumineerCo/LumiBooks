// ============================================================
//  LumiBooks v2.0 — Stock.gs
//  Stock/Inventory Management
//  © 2026 LumineerCo. All rights reserved.
// ============================================================

const SK = { ID:1,NAME:2,SKU:3,CAT:4,UNIT:5,QTY:6,COST:7,SELL:8,REORDER:9,LOC:10,DESC:11,CREATED:12,UPDATED:13 };

function addStockItem(token, data) {
  const ses = validateSession(token);
  if (!ses) return { ok:false, msg:'Session expired.' };
  if (ses.plan === CONFIG.PLANS.FREE) {
    const tab = getUserTab(ses.sheetId, CONFIG.USER_TABS.STOCK);
    if (tab.getLastRow() - 1 >= CONFIG.FREE_LIMITS.STOCK_ITEMS)
      return { ok:false, msg:'Free plan limit (50 items). Upgrade for unlimited.', upgrade:true };
  }
  if (!data.name) return { ok:false, msg:'Item name required.' };
  try {
    const id = generateId('SK');
    getUserTab(ses.sheetId, CONFIG.USER_TABS.STOCK).appendRow([
      id, sanitize(data.name), sanitize(data.sku||''), sanitize(data.category||'General'),
      sanitize(data.unit||'pcs'), parseFloat(data.quantity||0), parseFloat(data.cost_price||0),
      parseFloat(data.selling_price||0), parseFloat(data.reorder_level||0),
      sanitize(data.location||''), sanitize(data.description||''), now(), now(),
    ]);
    return { ok:true, msg:'Item added.', id };
  } catch(e) { return { ok:false, msg:'Failed to add item.' }; }
}

function getStockItems(token, opts) {
  const ses = validateSession(token);
  if (!ses) return { ok:false, msg:'Session expired.' };
  opts = opts || {};
  const page = parseInt(opts.page||1), size = parseInt(opts.size||25);
  const search = (opts.search||'').toLowerCase();
  const catF   = opts.category || '';
  try {
    const data = getUserTab(ses.sheetId, CONFIG.USER_TABS.STOCK).getDataRange().getValues();
    let rows = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      const r = data[i];
      const item = { idx:i, id:r[0], name:r[1], sku:r[2], category:r[3], unit:r[4],
        quantity:parseFloat(r[5])||0, costPrice:parseFloat(r[6])||0, sellingPrice:parseFloat(r[7])||0,
        reorderLevel:parseFloat(r[8])||0, location:r[9], description:r[10],
        isLow: (parseFloat(r[5])||0) <= (parseFloat(r[8])||0) };
      if (catF && item.category !== catF) continue;
      if (search && ![item.name,item.sku,item.category].join(' ').toLowerCase().includes(search)) continue;
      rows.push(item);
    }
    return { ok:true, total:rows.length, totalPages:Math.ceil(rows.length/size)||1, page,
      rows:rows.slice((page-1)*size, page*size),
      lowCount:rows.filter(r=>r.isLow).length };
  } catch(e) { return { ok:false, msg:'Failed to fetch stock.' }; }
}

function updateStockItem(token, itemId, data) {
  const ses = validateSession(token);
  if (!ses) return { ok:false, msg:'Session expired.' };
  try {
    const tab  = getUserTab(ses.sheetId, CONFIG.USER_TABS.STOCK);
    const rows = tab.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] !== itemId) continue;
      if (data.name)          tab.getRange(i+1, SK.NAME  ).setValue(sanitize(data.name));
      if (data.sku)           tab.getRange(i+1, SK.SKU   ).setValue(sanitize(data.sku));
      if (data.category)      tab.getRange(i+1, SK.CAT   ).setValue(sanitize(data.category));
      if (data.unit)          tab.getRange(i+1, SK.UNIT  ).setValue(sanitize(data.unit));
      if (data.quantity !== undefined) tab.getRange(i+1, SK.QTY).setValue(parseFloat(data.quantity)||0);
      if (data.cost_price !== undefined)    tab.getRange(i+1, SK.COST  ).setValue(parseFloat(data.cost_price)||0);
      if (data.selling_price !== undefined) tab.getRange(i+1, SK.SELL  ).setValue(parseFloat(data.selling_price)||0);
      if (data.reorder_level !== undefined) tab.getRange(i+1, SK.REORDER).setValue(parseFloat(data.reorder_level)||0);
      if (data.location)      tab.getRange(i+1, SK.LOC   ).setValue(sanitize(data.location));
      if (data.description)   tab.getRange(i+1, SK.DESC  ).setValue(sanitize(data.description));
      tab.getRange(i+1, SK.UPDATED).setValue(now());
      return { ok:true, msg:'Item updated.' };
    }
    return { ok:false, msg:'Item not found.' };
  } catch(e) { return { ok:false, msg:'Update failed.' }; }
}

function deleteStockItem(token, itemId) {
  const ses = validateSession(token);
  if (!ses) return { ok:false, msg:'Session expired.' };
  try {
    const tab  = getUserTab(ses.sheetId, CONFIG.USER_TABS.STOCK);
    const rows = tab.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] !== itemId) continue;
      tab.deleteRow(i+1);
      return { ok:true, msg:'Item deleted.' };
    }
    return { ok:false, msg:'Not found.' };
  } catch(e) { return { ok:false, msg:'Delete failed.' }; }
}

function adjustStock(token, itemId, qty, reason) {
  const ses = validateSession(token);
  if (!ses) return { ok:false, msg:'Session expired.' };
  try {
    const tab  = getUserTab(ses.sheetId, CONFIG.USER_TABS.STOCK);
    const rows = tab.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] !== itemId) continue;
      const current = parseFloat(rows[i][SK.QTY - 1]) || 0;
      const newQty  = current + parseFloat(qty);
      if (newQty < 0) return { ok:false, msg:'Insufficient stock.' };
      tab.getRange(i+1, SK.QTY    ).setValue(newQty);
      tab.getRange(i+1, SK.UPDATED).setValue(now());
      return { ok:true, msg:'Stock adjusted.', newQty };
    }
    return { ok:false, msg:'Item not found.' };
  } catch(e) { return { ok:false, msg:'Adjustment failed.' }; }
}

function getLowStockItems(token) {
  const ses = validateSession(token);
  if (!ses) return { ok:false, msg:'Session expired.' };
  try {
    const data = getUserTab(ses.sheetId, CONFIG.USER_TABS.STOCK).getDataRange().getValues();
    const rows = [];
    for (let i = 1; i < data.length; i++) {
      const qty = parseFloat(data[i][SK.QTY-1])||0, reorder = parseFloat(data[i][SK.REORDER-1])||0;
      if (qty <= reorder) rows.push({ id:data[i][0], name:data[i][1], quantity:qty, reorderLevel:reorder, unit:data[i][4] });
    }
    return { ok:true, rows };
  } catch(e) { return { ok:false, msg:'Failed.' }; }
}

// ============================================================
//  Customers.gs
// ============================================================

const CU = { ID:1,NAME:2,PHONE:3,EMAIL:4,GSTIN:5,ADDR:6,CITY:7,STATE:8,NOTES:9,CREDIT:10,BALANCE:11,CREATED:12 };

function addCustomer(token, data) {
  const ses = validateSession(token);
  if (!ses) return { ok:false, msg:'Session expired.' };
  if (ses.plan === CONFIG.PLANS.FREE) {
    const tab = getUserTab(ses.sheetId, CONFIG.USER_TABS.CUSTOMERS);
    if (tab.getLastRow() - 1 >= CONFIG.FREE_LIMITS.CUSTOMERS)
      return { ok:false, msg:'Free plan limit (25 customers). Upgrade for unlimited.', upgrade:true };
  }
  if (!data.name) return { ok:false, msg:'Customer name required.' };
  try {
    const id = generateId('CU');
    getUserTab(ses.sheetId, CONFIG.USER_TABS.CUSTOMERS).appendRow([
      id, sanitize(data.name), sanitize(data.phone||''), data.email ? data.email.trim().toLowerCase() : '',
      sanitize(data.gstin||''), sanitize(data.address||''), sanitize(data.city||''), sanitize(data.state||''),
      sanitize(data.notes||''), parseFloat(data.credit_limit||0), parseFloat(data.opening_balance||0), now(),
    ]);
    return { ok:true, msg:'Customer added.', id };
  } catch(e) { return { ok:false, msg:'Failed to add customer.' }; }
}

function getCustomers(token, opts) {
  const ses = validateSession(token);
  if (!ses) return { ok:false, msg:'Session expired.' };
  opts = opts || {};
  const page = parseInt(opts.page||1), size = parseInt(opts.size||25);
  const search = (opts.search||'').toLowerCase();
  try {
    const data = getUserTab(ses.sheetId, CONFIG.USER_TABS.CUSTOMERS).getDataRange().getValues();
    let rows = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      const r = data[i];
      const c = { idx:i, id:r[0], name:r[1], phone:r[2], email:r[3], gstin:r[4], address:r[5], city:r[6], state:r[7], notes:r[8] };
      if (search && ![c.name,c.phone,c.email,c.city].join(' ').toLowerCase().includes(search)) continue;
      rows.push(c);
    }
    return { ok:true, total:rows.length, totalPages:Math.ceil(rows.length/size)||1, page,
      rows:rows.slice((page-1)*size, page*size) };
  } catch(e) { return { ok:false, msg:'Failed.' }; }
}

function updateCustomer(token, custId, data) {
  const ses = validateSession(token);
  if (!ses) return { ok:false, msg:'Session expired.' };
  try {
    const tab = getUserTab(ses.sheetId, CONFIG.USER_TABS.CUSTOMERS);
    const rows = tab.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] !== custId) continue;
      if (data.name)    tab.getRange(i+1, CU.NAME  ).setValue(sanitize(data.name));
      if (data.phone)   tab.getRange(i+1, CU.PHONE ).setValue(sanitize(data.phone));
      if (data.email)   tab.getRange(i+1, CU.EMAIL ).setValue(data.email.trim().toLowerCase());
      if (data.gstin)   tab.getRange(i+1, CU.GSTIN ).setValue(sanitize(data.gstin));
      if (data.address) tab.getRange(i+1, CU.ADDR  ).setValue(sanitize(data.address));
      if (data.city)    tab.getRange(i+1, CU.CITY  ).setValue(sanitize(data.city));
      if (data.state)   tab.getRange(i+1, CU.STATE ).setValue(sanitize(data.state));
      if (data.notes)   tab.getRange(i+1, CU.NOTES ).setValue(sanitize(data.notes));
      return { ok:true, msg:'Customer updated.' };
    }
    return { ok:false, msg:'Customer not found.' };
  } catch(e) { return { ok:false, msg:'Update failed.' }; }
}

function deleteCustomer(token, custId) {
  const ses = validateSession(token);
  if (!ses) return { ok:false, msg:'Session expired.' };
  try {
    const tab = getUserTab(ses.sheetId, CONFIG.USER_TABS.CUSTOMERS);
    const rows = tab.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] !== custId) continue;
      tab.deleteRow(i+1);
      return { ok:true, msg:'Customer deleted.' };
    }
    return { ok:false, msg:'Not found.' };
  } catch(e) { return { ok:false, msg:'Delete failed.' }; }
}

function getCustomerStatement(token, custId) {
  const ses = validateSession(token);
  if (!ses) return { ok:false, msg:'Session expired.' };
  try {
    const cpData = getUserTab(ses.sheetId, CONFIG.USER_TABS.CUSTOMER_PAYMENTS).getDataRange().getValues();
    const estData = getUserTab(ses.sheetId, CONFIG.USER_TABS.ESTIMATES).getDataRange().getValues();
    const payments = [], bills = [];
    let totalBilled = 0, totalPaid = 0;
    for (let i = 1; i < cpData.length; i++) {
      if (cpData[i][1] !== custId) continue;
      const amt = parseFloat(cpData[i][3])||0;
      totalPaid += amt;
      payments.push({ date:cpData[i][2], amount:amt, mode:cpData[i][5], ref:cpData[i][4] });
    }
    for (let i = 1; i < estData.length; i++) {
      if (estData[i][EST.CUST_GSTIN - 1] === custId || estData[i][EST.CUST_NAME - 1]) {
        const amt = parseFloat(estData[i][EST.TOTAL - 1])||0;
        totalBilled += amt;
        bills.push({ billId:estData[i][0], date:estData[i][1], total:amt, status:estData[i][EST.STATUS-1] });
      }
    }
    return { ok:true, payments, bills, totalBilled, totalPaid, balance: totalBilled - totalPaid };
  } catch(e) { return { ok:false, msg:'Failed to generate statement.' }; }
}

// ============================================================
//  Staff.gs
// ============================================================

const SF = { ID:1,NAME:2,PHONE:3,EMAIL:4,ROLE:5,DEPT:6,SALARY:7,SAL_TYPE:8,JOIN:9,STATUS:10,BANK:11,IFSC:12,NOTES:13,CREATED:14 };
const AT = { DATE:1,STAFF_ID:2,STATUS:3,IN:4,OUT:5,HOURS:6,NOTES:7,CREATED:8 };
const SP = { ID:1,STAFF_ID:2,DATE:3,AMT:4,TYPE:5,MODE:6,MONTH:7,NOTES:8,CREATED:9 };

function addStaff(token, data) {
  const ses = validateSession(token);
  if (!ses) return { ok:false, msg:'Session expired.' };
  if (ses.plan === CONFIG.PLANS.FREE) {
    const tab = getUserTab(ses.sheetId, CONFIG.USER_TABS.STAFF);
    if (tab.getLastRow() - 1 >= CONFIG.FREE_LIMITS.STAFF)
      return { ok:false, msg:'Free plan limit (3 staff). Upgrade for unlimited.', upgrade:true };
  }
  if (!data.name) return { ok:false, msg:'Staff name required.' };
  try {
    const id = generateId('SF');
    getUserTab(ses.sheetId, CONFIG.USER_TABS.STAFF).appendRow([
      id, sanitize(data.name), sanitize(data.phone||''), data.email ? data.email.trim().toLowerCase() : '',
      sanitize(data.role||'Staff'), sanitize(data.department||'General'),
      parseFloat(data.salary||0), sanitize(data.salary_type||'monthly'),
      sanitize(data.join_date||todayStr()), 'active',
      sanitize(data.bank_acc||''), sanitize(data.ifsc||''), sanitize(data.notes||''), now(),
    ]);
    return { ok:true, msg:'Staff added.', id };
  } catch(e) { return { ok:false, msg:'Failed to add staff.' }; }
}

function getStaff(token, opts) {
  const ses = validateSession(token);
  if (!ses) return { ok:false, msg:'Session expired.' };
  opts = opts || {};
  const search = (opts.search||'').toLowerCase();
  try {
    const data = getUserTab(ses.sheetId, CONFIG.USER_TABS.STAFF).getDataRange().getValues();
    const rows = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      const r = data[i];
      const s = { idx:i, id:r[0], name:r[1], phone:r[2], email:r[3], role:r[4], department:r[5],
        salary:parseFloat(r[6])||0, salaryType:r[7], joinDate:r[8], status:r[9] };
      if (search && ![s.name,s.role,s.department].join(' ').toLowerCase().includes(search)) continue;
      rows.push(s);
    }
    return { ok:true, rows };
  } catch(e) { return { ok:false, msg:'Failed.' }; }
}

function updateStaff(token, staffId, data) {
  const ses = validateSession(token);
  if (!ses) return { ok:false, msg:'Session expired.' };
  try {
    const tab = getUserTab(ses.sheetId, CONFIG.USER_TABS.STAFF);
    const rows = tab.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] !== staffId) continue;
      const set = (col, val) => tab.getRange(i+1, col).setValue(val);
      if (data.name)        set(SF.NAME,    sanitize(data.name));
      if (data.phone)       set(SF.PHONE,   sanitize(data.phone));
      if (data.email)       set(SF.EMAIL,   data.email.trim().toLowerCase());
      if (data.role)        set(SF.ROLE,    sanitize(data.role));
      if (data.department)  set(SF.DEPT,    sanitize(data.department));
      if (data.salary !== undefined) set(SF.SALARY, parseFloat(data.salary)||0);
      if (data.salary_type) set(SF.SAL_TYPE, sanitize(data.salary_type));
      if (data.status)      set(SF.STATUS,  sanitize(data.status));
      if (data.bank_acc)    set(SF.BANK,    sanitize(data.bank_acc));
      if (data.ifsc)        set(SF.IFSC,    sanitize(data.ifsc));
      return { ok:true, msg:'Staff updated.' };
    }
    return { ok:false, msg:'Staff not found.' };
  } catch(e) { return { ok:false, msg:'Update failed.' }; }
}

function deleteStaff(token, staffId) {
  const ses = validateSession(token);
  if (!ses) return { ok:false, msg:'Session expired.' };
  try {
    const tab = getUserTab(ses.sheetId, CONFIG.USER_TABS.STAFF);
    const rows = tab.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] !== staffId) continue;
      tab.deleteRow(i+1);
      return { ok:true, msg:'Staff deleted.' };
    }
    return { ok:false, msg:'Not found.' };
  } catch(e) { return { ok:false, msg:'Delete failed.' }; }
}

function markAttendance(token, data) {
  const ses = validateSession(token);
  if (!ses) return { ok:false, msg:'Session expired.' };
  if (!data.date || !data.staff_id) return { ok:false, msg:'Date and staff ID required.' };
  try {
    const tab = getUserTab(ses.sheetId, CONFIG.USER_TABS.ATTENDANCE);
    const rows = tab.getDataRange().getValues();
    // Check existing entry
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][AT.DATE-1] === data.date && rows[i][AT.STAFF_ID-1] === data.staff_id) {
        tab.getRange(i+1, AT.STATUS).setValue(sanitize(data.status||'present'));
        if (data.check_in)  tab.getRange(i+1, AT.IN ).setValue(sanitize(data.check_in));
        if (data.check_out) tab.getRange(i+1, AT.OUT).setValue(sanitize(data.check_out));
        return { ok:true, msg:'Attendance updated.' };
      }
    }
    tab.appendRow([
      sanitize(data.date), sanitize(data.staff_id), sanitize(data.status||'present'),
      sanitize(data.check_in||''), sanitize(data.check_out||''), sanitize(data.total_hours||''),
      sanitize(data.notes||''), now(),
    ]);
    return { ok:true, msg:'Attendance marked.' };
  } catch(e) { return { ok:false, msg:'Failed.' }; }
}

function getAttendance(token, opts) {
  const ses = validateSession(token);
  if (!ses) return { ok:false, msg:'Session expired.' };
  opts = opts || {};
  const staffId = opts.staff_id || '';
  const month   = opts.month   || _monthKey(new Date()); // YYYY-MM
  try {
    const data = getUserTab(ses.sheetId, CONFIG.USER_TABS.ATTENDANCE).getDataRange().getValues();
    const rows = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      const dateStr = String(data[i][AT.DATE-1]).slice(0,10);
      if (month && !dateStr.startsWith(month)) continue;
      if (staffId && data[i][AT.STAFF_ID-1] !== staffId) continue;
      rows.push({ date:dateStr, staffId:data[i][1], status:data[i][2], checkIn:data[i][3], checkOut:data[i][4], hours:data[i][5] });
    }
    return { ok:true, rows };
  } catch(e) { return { ok:false, msg:'Failed.' }; }
}

function addStaffPayment(token, data) {
  const ses = validateSession(token);
  if (!ses) return { ok:false, msg:'Session expired.' };
  if (!data.staff_id || !data.amount) return { ok:false, msg:'Staff ID and amount required.' };
  try {
    const id = generateId('SP');
    getUserTab(ses.sheetId, CONFIG.USER_TABS.STAFF_PAYMENTS).appendRow([
      id, sanitize(data.staff_id), sanitize(data.date||todayStr()),
      parseFloat(data.amount), sanitize(data.pay_type||'salary'), sanitize(data.mode||'cash'),
      sanitize(data.month_year||_monthKey(new Date())), sanitize(data.notes||''), now(),
    ]);
    return { ok:true, msg:'Payment recorded.', id };
  } catch(e) { return { ok:false, msg:'Failed.' }; }
}

function getStaffPayments(token, opts) {
  const ses = validateSession(token);
  if (!ses) return { ok:false, msg:'Session expired.' };
  opts = opts || {};
  const staffId = opts.staff_id || '';
  const month   = opts.month   || '';
  try {
    const data = getUserTab(ses.sheetId, CONFIG.USER_TABS.STAFF_PAYMENTS).getDataRange().getValues();
    const rows = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      if (staffId && data[i][SP.STAFF_ID-1] !== staffId) continue;
      if (month   && data[i][SP.MONTH-1]    !== month  ) continue;
      rows.push({ id:data[i][0], staffId:data[i][1], date:data[i][2], amount:parseFloat(data[i][3])||0,
        payType:data[i][4], mode:data[i][5], month:data[i][6], notes:data[i][7] });
    }
    rows.sort((a,b) => new Date(b.date) - new Date(a.date));
    return { ok:true, rows };
  } catch(e) { return { ok:false, msg:'Failed.' }; }
}

// ============================================================
//  Payments.gs — Customer Payment Tracking + Plan Payments
// ============================================================

const CP = { ID:1,CUST_ID:2,DATE:3,AMT:4,BILL_REF:5,MODE:6,NOTES:7,CREATED:8 };

function addCustomerPayment(token, data) {
  const ses = validateSession(token);
  if (!ses) return { ok:false, msg:'Session expired.' };
  if (!data.amount) return { ok:false, msg:'Amount required.' };
  try {
    const id = generateId('CP');
    getUserTab(ses.sheetId, CONFIG.USER_TABS.CUSTOMER_PAYMENTS).appendRow([
      id, sanitize(data.cust_id||''), sanitize(data.date||todayStr()),
      parseFloat(data.amount), sanitize(data.bill_ref||''), sanitize(data.mode||'cash'),
      sanitize(data.notes||''), now(),
    ]);
    return { ok:true, msg:'Payment recorded.', id };
  } catch(e) { return { ok:false, msg:'Failed.' }; }
}

function getCustomerPayments(token, opts) {
  const ses = validateSession(token);
  if (!ses) return { ok:false, msg:'Session expired.' };
  opts = opts || {};
  const custId = opts.cust_id || '';
  const from   = opts.from    || '';
  const to     = opts.to      || '';
  try {
    const data = getUserTab(ses.sheetId, CONFIG.USER_TABS.CUSTOMER_PAYMENTS).getDataRange().getValues();
    let rows = [], total = 0;
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      const dateStr = String(data[i][CP.DATE-1]).slice(0,10);
      if (custId && data[i][CP.CUST_ID-1] !== custId) continue;
      if (from && dateStr < from) continue;
      if (to   && dateStr > to  ) continue;
      const amt = parseFloat(data[i][CP.AMT-1])||0;
      total += amt;
      rows.push({ id:data[i][0], custId:data[i][1], date:dateStr, amount:amt, billRef:data[i][4], mode:data[i][5], notes:data[i][6] });
    }
    rows.sort((a,b) => new Date(b.date) - new Date(a.date));
    return { ok:true, rows, total };
  } catch(e) { return { ok:false, msg:'Failed.' }; }
}

function deleteCustomerPayment(token, payId) {
  const ses = validateSession(token);
  if (!ses) return { ok:false, msg:'Session expired.' };
  try {
    const tab = getUserTab(ses.sheetId, CONFIG.USER_TABS.CUSTOMER_PAYMENTS);
    const rows = tab.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] !== payId) continue;
      tab.deleteRow(i+1);
      return { ok:true };
    }
    return { ok:false, msg:'Not found.' };
  } catch(e) { return { ok:false, msg:'Delete failed.' }; }
}

function submitPaymentRequest(token, data) {
  const ses = validateSession(token);
  if (!ses) return { ok:false, msg:'Session expired.' };
  const amount = parseFloat(data.amount || 0);
  const utr    = sanitize(data.utr_ref || '').trim();
  if (!utr || !amount) return { ok:false, msg:'Amount and UTR reference are required.' };
  try {
    const txId = 'PAY' + Date.now().toString(36).toUpperCase();
    getMasterTab(CONFIG.TABS.PAYMENTS).appendRow([txId, ses.userId, todayStr(), amount, utr, 'pending', '', data.plan_days||30]);
    return { ok:true, msg:'Payment submitted. Admin will verify within 24 hours.', txId };
  } catch(e) { return { ok:false, msg:'Submission failed.' }; }
}

function getPaymentDetails(token) {
  const ses = validateSession(token);
  if (!ses) return { ok:false, msg:'Session expired.' };
  try {
    const s = _getMasterSettings();
    return { ok:true, upiId: s['upi_id'] || CONFIG.UPI_ID, qrUrl: s['qr_url'] || '',
      prices: CONFIG.PRICES };
  } catch(e) { return { ok:false, msg:'Failed.' }; }
}

// ============================================================
//  Reports.gs
// ============================================================

function getReport(token, opts) {
  const ses = validateSession(token);
  if (!ses) return { ok:false, msg:'Session expired.' };

  opts = opts || {};
  const period = opts.period || 'month'; // week|month|quarter|year|custom
  const { from, to } = _getDateRange(period, opts.from, opts.to, ses.plan);

  try {
    const txData = getUserTab(ses.sheetId, CONFIG.USER_TABS.TRANSACTIONS).getDataRange().getValues();
    const estData = getUserTab(ses.sheetId, CONFIG.USER_TABS.ESTIMATES).getDataRange().getValues();

    let totalCredit=0, totalDebit=0, totalGst=0;
    const byCategory = {}, byMonth = {};
    const txRows = [];

    for (let i = 1; i < txData.length; i++) {
      const r = txData[i];
      if (!r[0]) continue;
      const dateStr = String(r[TX_COLS.DATE-1]).slice(0,10);
      if (dateStr < from || dateStr > to) continue;
      const type = r[TX_COLS.TYPE-1], amt = parseFloat(r[TX_COLS.AMOUNT-1])||0;
      const cat  = r[TX_COLS.CATEGORY-1]||'Other';
      const gstA = parseFloat(r[TX_COLS.GST_AMOUNT-1])||0;
      type === 'Credit' ? totalCredit += amt : totalDebit += amt;
      totalGst += gstA;
      byCategory[cat] = (byCategory[cat]||0) + amt;
      const mk = dateStr.slice(0,7);
      if (!byMonth[mk]) byMonth[mk] = { credit:0, debit:0 };
      type === 'Credit' ? byMonth[mk].credit += amt : byMonth[mk].debit += amt;
      txRows.push({ date:dateStr, type, category:cat, amount:amt, desc:r[TX_COLS.DESC-1] });
    }

    let totalBilled=0, billCount=0;
    for (let i = 1; i < estData.length; i++) {
      if (!estData[i][0]) continue;
      const dateStr = String(estData[i][EST.DATE-1]).slice(0,10);
      if (dateStr < from || dateStr > to) continue;
      totalBilled += parseFloat(estData[i][EST.TOTAL-1])||0;
      billCount++;
    }

    return {
      ok:true, from, to, period,
      totalCredit, totalDebit, netProfit: totalCredit - totalDebit,
      totalGst, totalBilled, billCount,
      byCategory, byMonth,
      txCount: txRows.length,
      transactions: opts.includeDetails ? txRows : [],
    };
  } catch(e) {
    Logger.log('getReport: ' + e);
    return { ok:false, msg:'Failed to generate report.' };
  }
}

function getStockReport(token, opts) {
  const ses = validateSession(token);
  if (!ses) return { ok:false, msg:'Session expired.' };
  try {
    const data = getUserTab(ses.sheetId, CONFIG.USER_TABS.STOCK).getDataRange().getValues();
    let totalItems=0, totalValue=0, lowStockItems=[], outOfStock=[];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      totalItems++;
      const qty = parseFloat(data[i][SK.QTY-1])||0, cost = parseFloat(data[i][SK.COST-1])||0, reorder = parseFloat(data[i][SK.REORDER-1])||0;
      totalValue += qty * cost;
      if (qty === 0) outOfStock.push(data[i][1]);
      else if (qty <= reorder) lowStockItems.push({ name:data[i][1], qty, reorder });
    }
    return { ok:true, totalItems, totalValue:parseFloat(totalValue.toFixed(2)), lowStockItems, outOfStock };
  } catch(e) { return { ok:false, msg:'Failed.' }; }
}

function getCustomerReport(token, opts) {
  const ses = validateSession(token);
  if (!ses) return { ok:false, msg:'Session expired.' };
  try {
    const cuData = getUserTab(ses.sheetId, CONFIG.USER_TABS.CUSTOMERS).getDataRange().getValues();
    const cpData = getUserTab(ses.sheetId, CONFIG.USER_TABS.CUSTOMER_PAYMENTS).getDataRange().getValues();
    const payMap = {};
    for (let i = 1; i < cpData.length; i++) {
      const cid = cpData[i][CP.CUST_ID-1];
      payMap[cid] = (payMap[cid]||0) + (parseFloat(cpData[i][CP.AMT-1])||0);
    }
    const rows = [];
    for (let i = 1; i < cuData.length; i++) {
      if (!cuData[i][0]) continue;
      rows.push({ id:cuData[i][0], name:cuData[i][1], city:cuData[i][6], totalPaid: payMap[cuData[i][0]]||0 });
    }
    return { ok:true, rows, total: rows.length };
  } catch(e) { return { ok:false, msg:'Failed.' }; }
}

function getStaffReport(token, opts) {
  const ses = validateSession(token);
  if (!ses) return { ok:false, msg:'Session expired.' };
  opts = opts || {};
  const month = opts.month || _monthKey(new Date());
  try {
    const sfData = getUserTab(ses.sheetId, CONFIG.USER_TABS.STAFF).getDataRange().getValues();
    const atData = getUserTab(ses.sheetId, CONFIG.USER_TABS.ATTENDANCE).getDataRange().getValues();
    const spData = getUserTab(ses.sheetId, CONFIG.USER_TABS.STAFF_PAYMENTS).getDataRange().getValues();

    const attendMap = {}, paidMap = {};
    for (let i = 1; i < atData.length; i++) {
      if (!String(atData[i][0]).startsWith(month)) continue;
      const sid = atData[i][AT.STAFF_ID-1];
      if (!attendMap[sid]) attendMap[sid] = { present:0, absent:0, leave:0 };
      const st = (atData[i][AT.STATUS-1]||'').toLowerCase();
      if (st === 'present' || st === 'p') attendMap[sid].present++;
      else if (st === 'absent' || st === 'a') attendMap[sid].absent++;
      else attendMap[sid].leave++;
    }
    for (let i = 1; i < spData.length; i++) {
      if (spData[i][SP.MONTH-1] !== month) continue;
      const sid = spData[i][SP.STAFF_ID-1];
      paidMap[sid] = (paidMap[sid]||0) + (parseFloat(spData[i][SP.AMT-1])||0);
    }

    const rows = [];
    for (let i = 1; i < sfData.length; i++) {
      if (!sfData[i][0]) continue;
      const sid = sfData[i][SF.ID-1];
      const att = attendMap[sid] || { present:0, absent:0, leave:0 };
      rows.push({ id:sid, name:sfData[i][1], role:sfData[i][4], salary:parseFloat(sfData[i][6])||0,
        ...att, totalPaid: paidMap[sid]||0 });
    }
    return { ok:true, rows, month };
  } catch(e) { return { ok:false, msg:'Failed.' }; }
}

function _getDateRange(period, from, to, plan) {
  const now_  = new Date();
  const today = todayStr();
  switch (period) {
    case 'week': {
      const d = new Date(now_); d.setDate(d.getDate() - 7);
      return { from: d.toISOString().slice(0,10), to: today };
    }
    case 'month': {
      const d = new Date(now_.getFullYear(), now_.getMonth(), 1);
      return { from: d.toISOString().slice(0,10), to: today };
    }
    case 'quarter': {
      const d = new Date(now_); d.setMonth(d.getMonth() - 3);
      return { from: d.toISOString().slice(0,10), to: today };
    }
    case 'year': {
      return { from: now_.getFullYear() + '-01-01', to: today };
    }
    case 'custom': {
      const maxDays = _getLimits(plan).REPORTS_DAYS;
      const f = from || today, t = to || today;
      const diff = (new Date(t) - new Date(f)) / 86400000;
      if (diff > maxDays) return { from: new Date(new Date(t) - maxDays*86400000).toISOString().slice(0,10), to: t };
      return { from: f, to: t };
    }
    default: return { from: now_.getFullYear() + '-01-01', to: today };
  }
}

// ============================================================
//  Referral.gs
// ============================================================

function getReferralInfo(token) {
  const ses = validateSession(token);
  if (!ses) return { ok:false, msg:'Session expired.' };
  try {
    const usersTab = getMasterTab(CONFIG.TABS.USERS);
    const data     = usersTab.getDataRange().getValues();
    let refCode = '', premExpiry = '';
    for (let i = 1; i < data.length; i++) {
      if (data[i][USER_COLS.ID - 1] !== ses.userId) continue;
      refCode    = data[i][USER_COLS.REFERRAL_CODE - 1] || '';
      premExpiry = data[i][USER_COLS.PREM_EXPIRY   - 1] || '';
      break;
    }

    const refTab  = getMasterTab(CONFIG.TABS.REFERRALS);
    const refData = refTab.getDataRange().getValues();
    let totalReferrals = 0, totalDays = 0;
    for (let i = 1; i < refData.length; i++) {
      if (refData[i][REF_COLS.REFERRER_ID - 1] !== ses.userId) continue;
      if (refData[i][REF_COLS.STATUS - 1] === 'completed') {
        totalReferrals++;
        totalDays += parseInt(refData[i][REF_COLS.DAYS - 1]) || 0;
      }
    }

    return { ok:true, refCode, premExpiry, totalReferrals, totalDaysEarned: totalDays,
      shareUrl: CONFIG.WEBSITE + '?ref=' + refCode };
  } catch(e) { return { ok:false, msg:'Failed to load referral info.' }; }
}
