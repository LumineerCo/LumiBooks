// ============================================================
//  LumiBooks v2.0 — Config.gs
//  All constants stay SERVER-SIDE. Never exposed to client.
//  © 2026 LumineerCo. All rights reserved.
// ============================================================

const CONFIG = {
  APP_NAME      : "LumiBooks",
  VERSION       : "2.0.0",
  BRAND         : "LumineerCo",
  TAGLINE       : "Smart Business Management — Without the Chaos",
  SUPPORT_EMAIL : "lumineerco@yahoo.com",
  WEBSITE       : "https://lumineerco.com",

  // ── Set after running setupMasterSheet() ──
  MASTER_SHEET_ID : "1hic2rciLXNQjMwFgRso-B0GJH1wOoIzUaK_NjE2UF5Y",

  // ── Google Drive folder structure: LumiBooks/Users/<userId>/ ──
  ROOT_FOLDER_NAME  : "LumiBooks",
  USERS_FOLDER_NAME : "Users",

  // ── Master sheet tab names ──
  TABS : {
    USERS    : "Users",
    PLANS    : "Plans",
    PAYMENTS : "ManualPayments",
    SETTINGS : "AdminSettings",
    SESSIONS : "Sessions",
    REFERRALS: "Referrals",
    CONTACT  : "ContactForms",
  },

  // ── Per-user sheet tab names ──
  USER_TABS : {
    TRANSACTIONS      : "Transactions",
    ESTIMATES         : "Estimates",
    STOCK             : "Stock",
    CUSTOMERS         : "Customers",
    STAFF             : "Staff",
    ATTENDANCE        : "Attendance",
    STAFF_PAYMENTS    : "StaffPayments",
    CUSTOMER_PAYMENTS : "CustomerPayments",
    CATEGORIES        : "Categories",
    SETTINGS          : "Settings",
    CUSTOM_FIELDS     : "CustomFields",
  },

  SESSION_TTL_HOURS   : 24,
  ADMIN_SESSION_HOURS : 8,

  PLANS : { FREE: "free", PREMIUM: "premium" },

  FREE_LIMITS : {
    TRANSACTIONS_PER_MONTH : 50,
    ESTIMATES_TOTAL        : 20,
    STOCK_ITEMS            : 50,
    CUSTOMERS              : 25,
    STAFF                  : 3,
    REPORTS_DAYS           : 30,
    EXPORTS                : ["print"],
    BILL_TEMPLATES         : ["classic", "modern"],
    BILL_EDIT_MINUTES      : 60,
  },

  PREMIUM_LIMITS : {
    TRANSACTIONS_PER_MONTH : Infinity,
    ESTIMATES_TOTAL        : Infinity,
    STOCK_ITEMS            : Infinity,
    CUSTOMERS              : Infinity,
    STAFF                  : Infinity,
    REPORTS_DAYS           : 1095,
    EXPORTS                : ["print", "pdf", "doc"],
    BILL_TEMPLATES         : ["classic", "modern", "professional", "executive", "gst"],
    BILL_EDIT_MINUTES      : 1440,
  },

  PRICES : { MONTHLY: 299, QUARTERLY: 799, YEARLY: 2499 },

  ADMIN_EMAIL : "lumineerco@yahoo.com",
  UPI_ID      : "lumineerco@ibl",

  // ── Change SALT before going live! ──
  SALT : "LumiBooks@LumineerCo#v2.0!ProductionKey2026@Secure",

  REFERRAL : { MIN_DAYS: 5, MAX_DAYS: 15, DEFAULT_DAYS: 7 },

  GST_RATES : [0, 5, 12, 18, 28],

  DEFAULT_INCOME_CATS  : ["Sales","Service","Consulting","Rental Income","Commission","Interest","Refund","Other Income"],
  DEFAULT_EXPENSE_CATS : ["Purchase","Utility","Salary","Rent","Transport","Marketing","Maintenance","Tax","Insurance","Office Supplies","Other Expense"],
};

// ─────────────────────────────────────────────────────────────
//  COLUMN INDICES (1-based)
// ─────────────────────────────────────────────────────────────

const USER_COLS = {
  ID           : 1, NAME      : 2, PHONE     : 3, EMAIL    : 4,
  PASS_HASH    : 5, PLAN      : 6, SHEET_ID  : 7, FOLDER_ID: 8,
  REG_DATE     : 9, STATUS    : 10,REFERRAL_CODE:11,REFERRED_BY:12,
  PREM_EXPIRY  :13,
};
const SES_COLS  = { TOKEN:1, USER_ID:2, SHEET_ID:3, PLAN:4, EXPIRES:5, CREATED:6, TYPE:7 };
const REF_COLS  = { CODE:1, REFERRER_ID:2, REFEREE_ID:3, DATE:4, DAYS:5, STATUS:6 };
const CONT_COLS = { ID:1, NAME:2, EMAIL:3, SUBJECT:4, MESSAGE:5, DATE:6, STATUS:7, REPLY:8 };

// ─────────────────────────────────────────────────────────────
//  SHEET ACCESSORS
// ─────────────────────────────────────────────────────────────

function getMasterSheet()            { return SpreadsheetApp.openById(CONFIG.MASTER_SHEET_ID); }
function getMasterTab(t)             { return getMasterSheet().getSheetByName(t); }
function getUserSpreadsheet(sid)     { return SpreadsheetApp.openById(sid); }
function getUserTab(sid, t)          { return getUserSpreadsheet(sid).getSheetByName(t); }

// ─────────────────────────────────────────────────────────────
//  SECURITY HELPERS
// ─────────────────────────────────────────────────────────────

function hashPassword(p) {
  const raw = CONFIG.SALT + p + CONFIG.SALT;
  const b   = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw, Utilities.Charset.UTF_8);
  return b.map(x => ('0' + (x & 0xff).toString(16)).slice(-2)).join('');
}

function generateToken()       { return Utilities.getUuid() + '-' + Utilities.getUuid() + '-' + Date.now().toString(36); }
function generateId(pfx)       { return (pfx||'ID') + Date.now().toString(36).toUpperCase() + Math.random().toString(36).slice(2,5).toUpperCase(); }
function now()                 { return new Date().toISOString(); }
function todayStr()            { return new Date().toISOString().slice(0,10); }

function generateReferralCode() {
  const c = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let s = 'LB';
  for (let i = 0; i < 6; i++) s += c[Math.floor(Math.random() * c.length)];
  return s;
}

function sanitize(v) {
  if (typeof v !== 'string') return v;
  return v.replace(/^[=+\-@|`]/, '').trim();
}
function sanitizeEmail(v) {
  if (typeof v !== 'string') return '';
  return v.replace(/[=+\-|`]/g, '').trim().toLowerCase();
}

// ─────────────────────────────────────────────────────────────
//  GOOGLE DRIVE FOLDER MANAGEMENT
// ─────────────────────────────────────────────────────────────

function _getRootFolder() {
  const it = DriveApp.getFoldersByName(CONFIG.ROOT_FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(CONFIG.ROOT_FOLDER_NAME);
}

function _getOrCreateSubFolder(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function _createUserDriveFolder(userId, safeName) {
  try {
    const root   = _getRootFolder();
    const users  = _getOrCreateSubFolder(root, CONFIG.USERS_FOLDER_NAME);
    const folder = users.createFolder('User_' + userId + '_' + safeName.replace(/[^a-zA-Z0-9_]/g,'_'));
    return folder.getId();
  } catch(e) { Logger.log('_createUserDriveFolder: ' + e); return null; }
}

function _moveFileToFolder(fileId, folderId) {
  try {
    if (!folderId) return;
    DriveApp.getFileById(fileId).moveTo(DriveApp.getFolderById(folderId));
  } catch(e) { Logger.log('_moveFileToFolder: ' + e); }
}

// ─────────────────────────────────────────────────────────────
//  PLAN UTILITIES
// ─────────────────────────────────────────────────────────────

function _getEffectivePlan(userRow) {
  const plan   = userRow[USER_COLS.PLAN        - 1];
  const expiry = userRow[USER_COLS.PREM_EXPIRY - 1];
  if (plan === CONFIG.PLANS.PREMIUM && expiry) {
    if (new Date(expiry) < new Date()) return CONFIG.PLANS.FREE;
  }
  return plan || CONFIG.PLANS.FREE;
}

function _getLimits(plan) {
  return plan === CONFIG.PLANS.PREMIUM ? CONFIG.PREMIUM_LIMITS : CONFIG.FREE_LIMITS;
}

// ─────────────────────────────────────────────────────────────
//  ADMIN CREDENTIAL HELPERS (stored as hashes in AdminSettings)
// ─────────────────────────────────────────────────────────────

function _getAdminCreds() {
  try {
    const s = _getMasterSettings();
    return { uHash: s['_au'] || '', pHash: s['_ap'] || '' };
  } catch(e) { return { uHash:'', pHash:'' }; }
}

// ─────────────────────────────────────────────────────────────
//  MASTER SETTINGS HELPERS
// ─────────────────────────────────────────────────────────────

function _getMasterSettings() {
  const d = getMasterTab(CONFIG.TABS.SETTINGS).getDataRange().getValues();
  const o = {};
  for (let i = 1; i < d.length; i++) o[d[i][0]] = d[i][1];
  return o;
}

function _upsertMasterSetting(key, val) {
  const tab = getMasterTab(CONFIG.TABS.SETTINGS);
  const d   = tab.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) {
    if (d[i][0] === key) { tab.getRange(i+1,2).setValue(val); return; }
  }
  tab.appendRow([key, val]);
}

// ─────────────────────────────────────────────────────────────
//  SETUP — Run ONCE to initialize all master sheet tabs
// ─────────────────────────────────────────────────────────────

function setupMasterSheet() {
  const ss = getMasterSheet();

  _ensureTab(ss, CONFIG.TABS.USERS,
    ['user_id','name','phone','email','password_hash','plan_tier','user_sheet_id','folder_id',
     'registration_date','status','referral_code','referred_by','premium_expiry']);
  _ensureTab(ss, CONFIG.TABS.PLANS, ['plan_name','features_json','price']);
  _ensureTab(ss, CONFIG.TABS.PAYMENTS, ['tx_id','user_id','date','amount','utr_ref','status','remarks','plan_days']);
  _ensureTab(ss, CONFIG.TABS.SETTINGS, ['key','value']);
  _ensureTab(ss, CONFIG.TABS.SESSIONS, ['token','user_id','sheet_id','plan','expires_at','created_at','type']);
  _ensureTab(ss, CONFIG.TABS.REFERRALS,['code','referrer_id','referee_id','date','days_granted','status']);
  _ensureTab(ss, CONFIG.TABS.CONTACT,  ['id','name','email','subject','message','date','status','admin_reply']);

  // Seed Plans
  const pt = getMasterTab(CONFIG.TABS.PLANS);
  pt.clearContents();
  pt.appendRow(['plan_name','features_json','price']);
  pt.appendRow(['free',    JSON.stringify({tx_limit:50, exports:['print'], templates:['classic','modern']}), 0]);
  pt.appendRow(['premium', JSON.stringify({tx_limit:-1, exports:['print','pdf','doc'],
    templates:['classic','modern','professional','executive','gst']}), 299]);

  // Seed default settings
  const defaults = [
    ['upi_id', CONFIG.UPI_ID],
    ['qr_url', ''],
    ['app_status', 'active'],
    ['maintenance_msg', ''],
    ['whatsapp_footer', 'Powered by LumiBooks — LumineerCo'],
  ];
  defaults.forEach(r => _upsertMasterSetting(r[0], r[1]));

  // Initialize admin credentials (stored as hashes — credentials never in plaintext in output)
  _setupAdminSecure();

  Logger.log('✅ LumiBooks v2.0 Master Sheet setup complete.');
}

function _setupAdminSecure() {
  // Admin credentials are initialized here and stored only as SHA-256 hashes.
  // After setup, the _setupAdminSecure function call can be removed from setupMasterSheet
  // for additional security. The plaintext values are never written to any sheet.
  const uRaw = 'BabluKSahu';
  const pRaw = 'Bablu@4151';
  _upsertMasterSetting('_au', hashPassword(uRaw));
  _upsertMasterSetting('_ap', hashPassword(pRaw));
  Logger.log('Admin credentials initialized (hashed).');
}

function _ensureTab(ss, name, headers) {
  let t = ss.getSheetByName(name);
  if (!t) { t = ss.insertSheet(name); t.appendRow(headers); t.setFrozenRows(1); }
  return t;
}

// ─────────────────────────────────────────────────────────────
//  CLEANUP TRIGGER — set up via Time Triggers
// ─────────────────────────────────────────────────────────────

function cleanupExpiredSessions() {
  const tab  = getMasterTab(CONFIG.TABS.SESSIONS);
  const data = tab.getDataRange().getValues();
  const now_ = new Date();
  const del  = [];
  for (let i = 1; i < data.length; i++) {
    if (new Date(data[i][SES_COLS.EXPIRES - 1]) < now_) del.push(i + 1);
  }
  for (let i = del.length - 1; i >= 0; i--) tab.deleteRow(del[i]);
  Logger.log('Cleaned ' + del.length + ' expired sessions.');
}

function checkPremiumExpiry() {
  const tab  = getMasterTab(CONFIG.TABS.USERS);
  const data = tab.getDataRange().getValues();
  const now_ = new Date();
  for (let i = 1; i < data.length; i++) {
    if (data[i][USER_COLS.PLAN - 1] !== CONFIG.PLANS.PREMIUM) continue;
    const expiry = data[i][USER_COLS.PREM_EXPIRY - 1];
    if (expiry && new Date(expiry) < now_) {
      tab.getRange(i+1, USER_COLS.PLAN).setValue(CONFIG.PLANS.FREE);
      Logger.log('Downgraded user: ' + data[i][USER_COLS.EMAIL - 1]);
    }
  }
}
