// ============================================================
//  LumiBooks v2.0 — Auth.gs
//  Authentication, sessions, profile management.
//  © 2026 LumineerCo. All rights reserved.
// ============================================================

// ─────────────────────────────────────────────────────────────
//  REGISTER
// ─────────────────────────────────────────────────────────────

function registerUser(data) {
  try {
    const name  = sanitize(data.name  || '').trim();
    const phone = sanitize(data.phone || '').trim();
    const email = sanitizeEmail(data.email || '');
    const pass  = data.password || '';
    const refCode = (data.referralCode || '').trim().toUpperCase();

    if (!name || !email || !pass || pass.length < 6)
      return { ok:false, msg:'Invalid input. Password must be ≥ 6 characters.' };
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email))
      return { ok:false, msg:'Invalid email format.' };

    const usersTab = getMasterTab(CONFIG.TABS.USERS);
    if (_findUserByEmail(email, usersTab)) return { ok:false, msg:'Email already registered.' };

    const userId    = 'U' + Date.now().toString(36).toUpperCase();
    const refGenerated = generateReferralCode();

    // Create personal Google Sheet
    const userSheet = SpreadsheetApp.create('LumiBooks_' + userId);
    _initUserSheet(userSheet);
    const sheetId = userSheet.getId();

    // Create Drive folder and move sheet
    const folderId = _createUserDriveFolder(userId, name);
    if (folderId) _moveFileToFolder(sheetId, folderId);

    // Handle referral
    let referredBy = '';
    if (refCode) {
      const refResult = _validateAndUseReferral(refCode, userId);
      if (refResult.valid) referredBy = refResult.referrerId;
    }

    usersTab.appendRow([
      userId, name, phone, email,
      hashPassword(pass),
      CONFIG.PLANS.FREE, sheetId, folderId || '',
      now(), 'active',
      refGenerated, referredBy, ''
    ]);

    return { ok:true, msg:'Account created! Please sign in.' };
  } catch(e) {
    Logger.log('registerUser error: ' + e);
    return { ok:false, msg:'Registration failed. Please try again.' };
  }
}

// ─────────────────────────────────────────────────────────────
//  LOGIN
// ─────────────────────────────────────────────────────────────

function loginUser(email, password) {
  try {
    email = (email || "").toLowerCase().trim();
    if (!email || !password) return { ok: false, msg: "Credentials required." };

    const usersTab = getMasterTab(CONFIG.TABS.USERS);
    const user     = _findUserByEmail(email, usersTab);
    if (!user) return { ok:false, msg:'Invalid email or password.' };
    if (user[USER_COLS.STATUS - 1] !== 'active') return { ok:false, msg:'Account suspended. Contact support.' };
    if (user[USER_COLS.PASS_HASH - 1] !== hashPassword(password)) return { ok:false, msg:'Invalid email or password.' };

    const userId  = user[USER_COLS.ID       - 1];
    const sheetId = user[USER_COLS.SHEET_ID - 1];
    const plan    = _getEffectivePlan(user);
    const expires = new Date(Date.now() + CONFIG.SESSION_TTL_HOURS * 3600000).toISOString();
    const token   = generateToken();

    getMasterTab(CONFIG.TABS.SESSIONS).appendRow([token, userId, sheetId, plan, expires, now(), 'user']);

    return {
      ok: true, token, plan,
      name    : user[USER_COLS.NAME  - 1],
      userId,
      refCode : user[USER_COLS.REFERRAL_CODE - 1] || '',
    };
  } catch(e) {
    Logger.log('loginUser error: ' + e);
    return { ok:false, msg:'Login failed. Please try again.' };
  }
}

// ─────────────────────────────────────────────────────────────
//  LOGOUT
// ─────────────────────────────────────────────────────────────

function logoutUser(token) {
  try { _deleteSession(token); return { ok:true }; } catch(e) { return { ok:false }; }
}

// ─────────────────────────────────────────────────────────────
//  VALIDATE SESSION
// ─────────────────────────────────────────────────────────────

function validateSession(token) {
  if (!token) return null;
  try {
    const sesTab = getMasterTab(CONFIG.TABS.SESSIONS);
    const data   = sesTab.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][SES_COLS.TOKEN - 1] !== token) continue;
      const exp = new Date(data[i][SES_COLS.EXPIRES - 1]);
      if (exp < new Date()) { sesTab.deleteRow(i + 1); return null; }
      return {
        userId  : data[i][SES_COLS.USER_ID  - 1],
        sheetId : data[i][SES_COLS.SHEET_ID - 1],
        plan    : data[i][SES_COLS.PLAN     - 1],
        type    : data[i][SES_COLS.TYPE     - 1] || 'user',
        rowIdx  : i + 1,
      };
    }
    return null;
  } catch(e) { Logger.log('validateSession: ' + e); return null; }
}

// ─────────────────────────────────────────────────────────────
//  CHANGE PASSWORD
// ─────────────────────────────────────────────────────────────

function changePassword(token, oldPass, newPass) {
  const ses = validateSession(token);
  if (!ses) return { ok:false, msg:'Session expired.' };
  if (!newPass || newPass.length < 6) return { ok:false, msg:'New password must be ≥ 6 characters.' };
  try {
    const tab  = getMasterTab(CONFIG.TABS.USERS);
    const data = tab.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][USER_COLS.ID - 1] !== ses.userId) continue;
      if (data[i][USER_COLS.PASS_HASH - 1] !== hashPassword(oldPass))
        return { ok:false, msg:'Current password is incorrect.' };
      tab.getRange(i+1, USER_COLS.PASS_HASH).setValue(hashPassword(newPass));
      return { ok:true, msg:'Password updated successfully.' };
    }
    return { ok:false, msg:'User not found.' };
  } catch(e) { return { ok:false, msg:'Error updating password.' }; }
}

// ─────────────────────────────────────────────────────────────
//  PROFILE — GET
// ─────────────────────────────────────────────────────────────

function getProfile(token) {
  const ses = validateSession(token);
  if (!ses) return { ok:false, msg:'Session expired.' };
  const tab  = getMasterTab(CONFIG.TABS.USERS);
  const data = tab.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][USER_COLS.ID - 1] !== ses.userId) continue;
    const r = data[i];
    return {
      ok      : true,
      name    : r[USER_COLS.NAME         - 1],
      phone   : r[USER_COLS.PHONE        - 1],
      email   : r[USER_COLS.EMAIL        - 1],
      plan    : _getEffectivePlan(r),
      regDate : r[USER_COLS.REG_DATE     - 1],
      refCode : r[USER_COLS.REFERRAL_CODE- 1],
      premExpiry: r[USER_COLS.PREM_EXPIRY- 1],
      biz     : _getBusinessSettings(ses.sheetId),
    };
  }
  return { ok:false, msg:'User not found.' };
}

// ─────────────────────────────────────────────────────────────
//  PROFILE — UPDATE
// ─────────────────────────────────────────────────────────────

function updateProfile(token, data) {
  const ses = validateSession(token);
  if (!ses) return { ok:false, msg:'Session expired.' };
  try {
    const tab  = getMasterTab(CONFIG.TABS.USERS);
    const rows = tab.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][USER_COLS.ID - 1] !== ses.userId) continue;
      if (data.name)  tab.getRange(i+1, USER_COLS.NAME ).setValue(sanitize(data.name));
      if (data.phone) tab.getRange(i+1, USER_COLS.PHONE).setValue(sanitize(data.phone));
      break;
    }
    if (data.biz) _saveBusinessSettings(ses.sheetId, data.biz);
    return { ok:true, msg:'Profile updated successfully.' };
  } catch(e) { return { ok:false, msg:'Update failed.' }; }
}

// ─────────────────────────────────────────────────────────────
//  BUSINESS SETTINGS (user's own sheet)
// ─────────────────────────────────────────────────────────────

function _getBusinessSettings(sheetId) {
  try {
    const tab  = getUserTab(sheetId, CONFIG.USER_TABS.SETTINGS);
    const data = tab.getDataRange().getValues();
    const obj  = {};
    for (let i = 1; i < data.length; i++) obj[data[i][0]] = data[i][1];
    return obj;
  } catch(e) { return {}; }
}

function _saveBusinessSettings(sheetId, biz) {
  const tab  = getUserTab(sheetId, CONFIG.USER_TABS.SETTINGS);
  const data = tab.getDataRange().getValues();
  Object.keys(biz).forEach(key => {
    const sk = sanitize(String(key));
    // Preserve @ in email values
    const sv = key.toLowerCase().includes('email')
      ? String(biz[key]).trim().toLowerCase()
      : sanitize(String(biz[key]));
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === sk) { tab.getRange(i+1, 2).setValue(sv); found = true; break; }
    }
    if (!found) tab.appendRow([sk, sv]);
  });
}

// ─────────────────────────────────────────────────────────────
//  REFERRAL HELPERS
// ─────────────────────────────────────────────────────────────

function _validateAndUseReferral(code, newUserId) {
  try {
    const tab  = getMasterTab(CONFIG.TABS.USERS);
    const data = tab.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][USER_COLS.REFERRAL_CODE - 1] !== code) continue;
      const referrerId = data[i][USER_COLS.ID - 1];
      if (referrerId === newUserId) return { valid:false };
      // Grant premium days to referrer
      const days = CONFIG.REFERRAL.DEFAULT_DAYS;
      _extendPremium(referrerId, days);
      // Log referral
      getMasterTab(CONFIG.TABS.REFERRALS).appendRow([code, referrerId, newUserId, now(), days, 'completed']);
      return { valid:true, referrerId };
    }
    return { valid:false };
  } catch(e) { return { valid:false }; }
}

function _extendPremium(userId, days) {
  const tab  = getMasterTab(CONFIG.TABS.USERS);
  const data = tab.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][USER_COLS.ID - 1] !== userId) continue;
    const current  = data[i][USER_COLS.PREM_EXPIRY - 1];
    const base     = (current && new Date(current) > new Date()) ? new Date(current) : new Date();
    const newExpiry = new Date(base.getTime() + days * 86400000).toISOString();
    tab.getRange(i+1, USER_COLS.PLAN).setValue(CONFIG.PLANS.PREMIUM);
    tab.getRange(i+1, USER_COLS.PREM_EXPIRY).setValue(newExpiry);
    // Update active sessions
    const sesTab  = getMasterTab(CONFIG.TABS.SESSIONS);
    const sesData = sesTab.getDataRange().getValues();
    for (let j = 1; j < sesData.length; j++) {
      if (sesData[j][SES_COLS.USER_ID - 1] === userId) {
        sesTab.getRange(j+1, SES_COLS.PLAN).setValue(CONFIG.PLANS.PREMIUM);
      }
    }
    return;
  }
}

// ─────────────────────────────────────────────────────────────
//  INTERNAL HELPERS
// ─────────────────────────────────────────────────────────────

function _findUserByEmail(email, usersTab) {
  const data = usersTab.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][USER_COLS.EMAIL - 1] === email) return data[i];
  }
  return null;
}

function _deleteSession(token) {
  const tab  = getMasterTab(CONFIG.TABS.SESSIONS);
  const data = tab.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][SES_COLS.TOKEN - 1] === token) { tab.deleteRow(i + 1); return; }
  }
}

function _initUserSheet(ss) {
  // Transactions
  let t = ss.getActiveSheet();
  t.setName(CONFIG.USER_TABS.TRANSACTIONS);
  t.appendRow(['date','type','category','amount','description','reference_no','gst_rate','gst_amount','total_with_gst','created_at']);
  t.setFrozenRows(1);

  // Estimates
  let e = ss.insertSheet(CONFIG.USER_TABS.ESTIMATES);
  e.appendRow(['bill_id','date','due_date','cust_name','cust_phone','cust_addr','cust_gstin',
    'items_json','subtotal','discount','gst_type','gst_pct','gst_amount','total',
    'notes','status','template','locked_at','gst_enabled','created_at']);
  e.setFrozenRows(1);

  // Stock
  let sk = ss.insertSheet(CONFIG.USER_TABS.STOCK);
  sk.appendRow(['item_id','name','sku','category','unit','quantity','cost_price','selling_price',
    'reorder_level','location','description','created_at','updated_at']);
  sk.setFrozenRows(1);

  // Customers
  let cu = ss.insertSheet(CONFIG.USER_TABS.CUSTOMERS);
  cu.appendRow(['cust_id','name','phone','email','gstin','address','city','state','notes',
    'credit_limit','opening_balance','created_at']);
  cu.setFrozenRows(1);

  // Staff
  let sf = ss.insertSheet(CONFIG.USER_TABS.STAFF);
  sf.appendRow(['staff_id','name','phone','email','role','department','salary','salary_type',
    'join_date','status','bank_acc','ifsc','notes','created_at']);
  sf.setFrozenRows(1);

  // Attendance
  let at = ss.insertSheet(CONFIG.USER_TABS.ATTENDANCE);
  at.appendRow(['date','staff_id','status','check_in','check_out','total_hours','notes','created_at']);
  at.setFrozenRows(1);

  // StaffPayments
  let sp = ss.insertSheet(CONFIG.USER_TABS.STAFF_PAYMENTS);
  sp.appendRow(['pay_id','staff_id','date','amount','pay_type','mode','month_year','notes','created_at']);
  sp.setFrozenRows(1);

  // CustomerPayments
  let cp = ss.insertSheet(CONFIG.USER_TABS.CUSTOMER_PAYMENTS);
  cp.appendRow(['pay_id','cust_id','date','amount','bill_ref','mode','notes','created_at']);
  cp.setFrozenRows(1);

  // Categories
  let ca = ss.insertSheet(CONFIG.USER_TABS.CATEGORIES);
  ca.appendRow(['cat_id','name','type','color','icon','is_custom','created_at']);
  ca.setFrozenRows(1);
  // Seed default categories
  CONFIG.DEFAULT_INCOME_CATS.forEach(n =>
    ca.appendRow([generateId('CAT'), n, 'income', '#22c55e', '📈', 'false', now()]));
  CONFIG.DEFAULT_EXPENSE_CATS.forEach(n =>
    ca.appendRow([generateId('CAT'), n, 'expense', '#ef4444', '📉', 'false', now()]));

  // Settings
  let se = ss.insertSheet(CONFIG.USER_TABS.SETTINGS);
  se.appendRow(['key','value']);
  se.setFrozenRows(1);
  [['business_name','My Business'],['business_address',''],['business_phone',''],
   ['business_email',''],['currency','₹'],['gstin',''],['state_code',''],
   ['currency_code','INR'],['theme','dark']].forEach(r => se.appendRow(r));

  // CustomFields
  let cf = ss.insertSheet(CONFIG.USER_TABS.CUSTOM_FIELDS);
  cf.appendRow(['field_id','module','field_name','field_type','required','options_json','created_at']);
  cf.setFrozenRows(1);
}

// ─────────────────────────────────────────────────────────────
//  PUBLIC — CONTACT FORM (no auth required)
// ─────────────────────────────────────────────────────────────

function submitContactForm(data) {
  try {
    if (!data.name || !data.email || !data.message)
      return { ok:false, msg:'Name, email and message are required.' };
    const id = generateId('CF');
    getMasterTab(CONFIG.TABS.CONTACT).appendRow([
      id, sanitize(data.name), sanitizeEmail(data.email),
      sanitize(data.subject || 'General Enquiry'),
      sanitize(data.message), now(), 'new', ''
    ]);
    return { ok:true, msg:'Message sent! We\'ll respond within 24 hours.' };
  } catch(e) { return { ok:false, msg:'Failed to send. Please email us directly.' }; }
}
