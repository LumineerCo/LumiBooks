// ============================================================
//  LumiBooks v2.0 — Admin.gs
//  Super Advanced Admin Panel — secure, feature-rich.
//  Admin credentials stored as hashes ONLY. Never in plaintext.
//  © 2026 LumineerCo. All rights reserved.
// ============================================================

// ─────────────────────────────────────────────────────────────
//  ADMIN SESSION TOKEN PREFIX — distinguishes from user tokens
// ─────────────────────────────────────────────────────────────
const ADMIN_TOKEN_PREFIX = 'ADMSESS_';

// ─────────────────────────────────────────────────────────────
//  ADMIN LOGIN — username + password (separate from user auth)
// ─────────────────────────────────────────────────────────────

function adminLogin(username, password) {
  try {
    if (!username || !password) return { ok:false, msg:'Credentials required.' };

    const creds = _getAdminCreds();
    if (!creds.uHash || !creds.pHash) return { ok:false, msg:'Admin not configured. Run setupMasterSheet().' };

    if (hashPassword(username) !== creds.uHash || hashPassword(password) !== creds.pHash)
      return { ok:false, msg:'Invalid credentials.' };

    const token   = ADMIN_TOKEN_PREFIX + generateToken();
    const expires = new Date(Date.now() + CONFIG.ADMIN_SESSION_HOURS * 3600000).toISOString();
    getMasterTab(CONFIG.TABS.SESSIONS).appendRow([token, 'ADMIN', '', 'admin', expires, now(), 'admin']);

    return { ok:true, token, expires };
  } catch(e) {
    Logger.log('adminLogin: ' + e);
    return { ok:false, msg:'Login failed.' };
  }
}

function adminLogout(token) {
  try { _deleteSession(token); return { ok:true }; } catch(e) { return { ok:false }; }
}

// ─────────────────────────────────────────────────────────────
//  ADMIN SESSION VALIDATE
// ─────────────────────────────────────────────────────────────

function _validateAdmin(token) {
  if (!token || !token.startsWith(ADMIN_TOKEN_PREFIX)) return false;
  const sesTab = getMasterTab(CONFIG.TABS.SESSIONS);
  const data   = sesTab.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][SES_COLS.TOKEN - 1] !== token) continue;
    if (data[i][SES_COLS.TYPE  - 1] !== 'admin') return false;
    const exp = new Date(data[i][SES_COLS.EXPIRES - 1]);
    if (exp < new Date()) { sesTab.deleteRow(i + 1); return false; }
    return true;
  }
  return false;
}

// ─────────────────────────────────────────────────────────────
//  ADMIN — DASHBOARD STATS
// ─────────────────────────────────────────────────────────────

function adminGetDashboard(token) {
  if (!_validateAdmin(token)) return { ok:false, msg:'Access denied.' };
  try {
    const usersData = getMasterTab(CONFIG.TABS.USERS).getDataRange().getValues();
    const payData   = getMasterTab(CONFIG.TABS.PAYMENTS).getDataRange().getValues();
    const refData   = getMasterTab(CONFIG.TABS.REFERRALS).getDataRange().getValues();
    const sesData   = getMasterTab(CONFIG.TABS.SESSIONS).getDataRange().getValues();

    let totalUsers=0, activeUsers=0, premiumUsers=0, suspendedUsers=0;
    const now_ = new Date();

    for (let i = 1; i < usersData.length; i++) {
      const r = usersData[i];
      if (!r[USER_COLS.ID - 1]) continue;
      totalUsers++;
      const status = r[USER_COLS.STATUS - 1];
      const plan   = _getEffectivePlan(r);
      if (status === 'active')    activeUsers++;
      if (status === 'suspended') suspendedUsers++;
      if (plan === CONFIG.PLANS.PREMIUM) premiumUsers++;
    }

    let pendingPayments=0, approvedRevenue=0;
    for (let i = 1; i < payData.length; i++) {
      const r = payData[i];
      if (!r[0]) continue;
      if (r[5] === 'pending')  pendingPayments++;
      if (r[5] === 'approved') approvedRevenue += parseFloat(r[3])||0;
    }

    let totalReferrals=0;
    for (let i = 1; i < refData.length; i++) {
      if (refData[i][REF_COLS.STATUS - 1] === 'completed') totalReferrals++;
    }

    let activeSessions=0;
    for (let i = 1; i < sesData.length; i++) {
      if (!sesData[i][SES_COLS.TOKEN - 1]) continue;
      if (new Date(sesData[i][SES_COLS.EXPIRES - 1]) > now_ &&
          sesData[i][SES_COLS.TYPE - 1] !== 'admin') activeSessions++;
    }

    // New users this month
    const curMonth = _monthKey(new Date());
    let newUsersThisMonth = 0;
    for (let i = 1; i < usersData.length; i++) {
      const reg = String(usersData[i][USER_COLS.REG_DATE - 1]).slice(0,7);
      if (reg === curMonth) newUsersThisMonth++;
    }

    return {
      ok:true,
      totalUsers, activeUsers, premiumUsers, suspendedUsers,
      pendingPayments, approvedRevenue, totalReferrals, activeSessions, newUsersThisMonth,
    };
  } catch(e) {
    Logger.log('adminGetDashboard: ' + e);
    return { ok:false, msg:'Failed to load dashboard.' };
  }
}

// ─────────────────────────────────────────────────────────────
//  ADMIN — USER MANAGEMENT
// ─────────────────────────────────────────────────────────────

function adminGetUsers(token, opts) {
  if (!_validateAdmin(token)) return { ok:false, msg:'Access denied.' };
  opts = opts || {};
  const page   = parseInt(opts.page || 1);
  const size   = parseInt(opts.size || 25);
  const search = (opts.search || '').toLowerCase();
  const planF  = opts.plan   || '';
  const statF  = opts.status || '';

  try {
    const data = getMasterTab(CONFIG.TABS.USERS).getDataRange().getValues();
    let rows = [];
    for (let i = 1; i < data.length; i++) {
      const r = data[i];
      if (!r[USER_COLS.ID - 1]) continue;
      const plan = _getEffectivePlan(r);
      if (planF && plan !== planF) continue;
      if (statF && r[USER_COLS.STATUS - 1] !== statF) continue;
      if (search && ![r[1],r[3],r[4]].join(' ').toLowerCase().includes(search)) continue;
      rows.push({
        userId   : r[USER_COLS.ID          - 1],
        name     : r[USER_COLS.NAME        - 1],
        email    : r[USER_COLS.EMAIL       - 1],
        phone    : r[USER_COLS.PHONE       - 1],
        plan,
        regDate  : String(r[USER_COLS.REG_DATE   - 1]).slice(0,10),
        status   : r[USER_COLS.STATUS      - 1],
        refCode  : r[USER_COLS.REFERRAL_CODE - 1],
        premExpiry: r[USER_COLS.PREM_EXPIRY - 1],
      });
    }
    return { ok:true, total:rows.length, totalPages:Math.ceil(rows.length/size)||1, page,
      rows: rows.slice((page-1)*size, page*size) };
  } catch(e) { return { ok:false, msg:'Failed to fetch users.' }; }
}

function adminSetUserStatus(token, userId, status) {
  if (!_validateAdmin(token)) return { ok:false, msg:'Access denied.' };
  if (!['active','suspended'].includes(status)) return { ok:false, msg:'Invalid status.' };
  try {
    const tab  = getMasterTab(CONFIG.TABS.USERS);
    const data = tab.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][USER_COLS.ID - 1] !== userId) continue;
      tab.getRange(i+1, USER_COLS.STATUS).setValue(status);
      return { ok:true, msg:`User ${status}.` };
    }
    return { ok:false, msg:'User not found.' };
  } catch(e) { return { ok:false, msg:'Operation failed.' }; }
}

function adminSetUserPlan(token, userId, plan, days) {
  if (!_validateAdmin(token)) return { ok:false, msg:'Access denied.' };
  if (![CONFIG.PLANS.FREE, CONFIG.PLANS.PREMIUM].includes(plan)) return { ok:false, msg:'Invalid plan.' };
  try {
    const tab  = getMasterTab(CONFIG.TABS.USERS);
    const data = tab.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][USER_COLS.ID - 1] !== userId) continue;
      tab.getRange(i+1, USER_COLS.PLAN).setValue(plan);
      if (plan === CONFIG.PLANS.PREMIUM && days) {
        const d = parseInt(days) || 30;
        const expiry = new Date(Date.now() + d * 86400000).toISOString();
        tab.getRange(i+1, USER_COLS.PREM_EXPIRY).setValue(expiry);
      } else if (plan === CONFIG.PLANS.FREE) {
        tab.getRange(i+1, USER_COLS.PREM_EXPIRY).setValue('');
      }
      // Sync active sessions
      _syncUserSessions(userId, plan);
      return { ok:true, msg:`User plan updated to ${plan}${days ? ' for ' + days + ' days' : ''}.` };
    }
    return { ok:false, msg:'User not found.' };
  } catch(e) { return { ok:false, msg:'Update failed.' }; }
}

function adminDeleteUser(token, userId) {
  if (!_validateAdmin(token)) return { ok:false, msg:'Access denied.' };
  try {
    const tab  = getMasterTab(CONFIG.TABS.USERS);
    const data = tab.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][USER_COLS.ID - 1] !== userId) continue;
      // Optionally archive user's sheet instead of deleting
      tab.getRange(i+1, USER_COLS.STATUS).setValue('deleted');
      // Purge sessions
      const sesTab  = getMasterTab(CONFIG.TABS.SESSIONS);
      const sesData = sesTab.getDataRange().getValues();
      const del = [];
      for (let j = 1; j < sesData.length; j++) {
        if (sesData[j][SES_COLS.USER_ID - 1] === userId) del.push(j+1);
      }
      for (let j = del.length-1; j >= 0; j--) sesTab.deleteRow(del[j]);
      return { ok:true, msg:'User account deactivated.' };
    }
    return { ok:false, msg:'User not found.' };
  } catch(e) { return { ok:false, msg:'Operation failed.' }; }
}

// ─────────────────────────────────────────────────────────────
//  ADMIN — PAYMENT MANAGEMENT
// ─────────────────────────────────────────────────────────────

function adminGetPendingPayments(token) {
  if (!_validateAdmin(token)) return { ok:false, msg:'Access denied.' };
  try {
    const data = getMasterTab(CONFIG.TABS.PAYMENTS).getDataRange().getValues();
    const rows = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][5] !== 'pending') continue;
      rows.push({ idx:i, txId:data[i][0], userId:data[i][1], date:data[i][2],
        amount:data[i][3], utr:data[i][4], status:data[i][5], planDays:data[i][7]||30 });
    }
    return { ok:true, rows };
  } catch(e) { return { ok:false, msg:'Failed.' }; }
}

function adminGetAllPayments(token, opts) {
  if (!_validateAdmin(token)) return { ok:false, msg:'Access denied.' };
  opts = opts || {};
  try {
    const data = getMasterTab(CONFIG.TABS.PAYMENTS).getDataRange().getValues();
    const rows = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      const status = opts.status || '';
      if (status && data[i][5] !== status) continue;
      rows.push({ idx:i, txId:data[i][0], userId:data[i][1], date:data[i][2],
        amount:data[i][3], utr:data[i][4], status:data[i][5], remarks:data[i][6], planDays:data[i][7]||30 });
    }
    rows.sort((a,b) => new Date(b.date) - new Date(a.date));
    return { ok:true, rows, total:rows.length };
  } catch(e) { return { ok:false, msg:'Failed.' }; }
}

function adminApprovePayment(token, txId, remarks, days) {
  if (!_validateAdmin(token)) return { ok:false, msg:'Access denied.' };
  try {
    const payTab  = getMasterTab(CONFIG.TABS.PAYMENTS);
    const payData = payTab.getDataRange().getValues();
    let userId = null, planDays = parseInt(days) || 30;

    for (let i = 1; i < payData.length; i++) {
      if (payData[i][0] !== txId) continue;
      payTab.getRange(i+1, 6).setValue('approved');
      payTab.getRange(i+1, 7).setValue(sanitize(remarks || 'Approved by admin'));
      userId = payData[i][1];
      planDays = parseInt(payData[i][7]) || planDays;
      break;
    }
    if (!userId) return { ok:false, msg:'Transaction not found.' };

    // Upgrade user
    const expiry = new Date(Date.now() + planDays * 86400000).toISOString();
    const usTab  = getMasterTab(CONFIG.TABS.USERS);
    const usData = usTab.getDataRange().getValues();
    for (let i = 1; i < usData.length; i++) {
      if (usData[i][USER_COLS.ID - 1] !== userId) continue;
      usTab.getRange(i+1, USER_COLS.PLAN).setValue(CONFIG.PLANS.PREMIUM);
      usTab.getRange(i+1, USER_COLS.PREM_EXPIRY).setValue(expiry);
      break;
    }
    _syncUserSessions(userId, CONFIG.PLANS.PREMIUM);
    return { ok:true, msg:`Payment approved. User upgraded to Premium for ${planDays} days.` };
  } catch(e) {
    Logger.log('adminApprovePayment: ' + e);
    return { ok:false, msg:'Approval failed.' };
  }
}

function adminRejectPayment(token, txId, remarks) {
  if (!_validateAdmin(token)) return { ok:false, msg:'Access denied.' };
  try {
    const tab  = getMasterTab(CONFIG.TABS.PAYMENTS);
    const data = tab.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] !== txId) continue;
      tab.getRange(i+1, 6).setValue('rejected');
      tab.getRange(i+1, 7).setValue(sanitize(remarks || 'Rejected by admin'));
      return { ok:true, msg:'Payment rejected.' };
    }
    return { ok:false, msg:'Transaction not found.' };
  } catch(e) { return { ok:false, msg:'Rejection failed.' }; }
}

// ─────────────────────────────────────────────────────────────
//  ADMIN — REFERRALS
// ─────────────────────────────────────────────────────────────

function adminGetReferrals(token) {
  if (!_validateAdmin(token)) return { ok:false, msg:'Access denied.' };
  try {
    const data = getMasterTab(CONFIG.TABS.REFERRALS).getDataRange().getValues();
    const rows = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      rows.push({ code:data[i][0], referrerId:data[i][1], refereeId:data[i][2],
        date:data[i][3], days:data[i][4], status:data[i][5] });
    }
    return { ok:true, rows, total:rows.length };
  } catch(e) { return { ok:false, msg:'Failed.' }; }
}

// ─────────────────────────────────────────────────────────────
//  ADMIN — CONTACT FORMS
// ─────────────────────────────────────────────────────────────

function adminGetContactForms(token) {
  if (!_validateAdmin(token)) return { ok:false, msg:'Access denied.' };
  try {
    const data = getMasterTab(CONFIG.TABS.CONTACT).getDataRange().getValues();
    const rows = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      rows.push({ id:data[i][0], name:data[i][1], email:data[i][2], subject:data[i][3],
        message:data[i][4], date:data[i][5], status:data[i][6], reply:data[i][7] });
    }
    rows.sort((a,b) => new Date(b.date) - new Date(a.date));
    return { ok:true, rows };
  } catch(e) { return { ok:false, msg:'Failed.' }; }
}

function adminReplyContact(token, formId, reply) {
  if (!_validateAdmin(token)) return { ok:false, msg:'Access denied.' };
  try {
    const tab  = getMasterTab(CONFIG.TABS.CONTACT);
    const data = tab.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][CONT_COLS.ID - 1] !== formId) continue;
      tab.getRange(i+1, CONT_COLS.STATUS).setValue('replied');
      tab.getRange(i+1, CONT_COLS.REPLY ).setValue(sanitize(reply || ''));
      // Send email reply
      try {
        MailApp.sendEmail({
          to      : data[i][CONT_COLS.EMAIL - 1],
          subject : 'Re: ' + data[i][CONT_COLS.SUBJECT - 1] + ' — LumiBooks Support',
          body    : `Hi ${data[i][CONT_COLS.NAME - 1]},\n\n${reply}\n\nBest regards,\nLumiBooks Support Team\nlumineerco@yahoo.com`,
        });
      } catch(mailErr) { Logger.log('Mail send failed: ' + mailErr); }
      return { ok:true, msg:'Reply sent.' };
    }
    return { ok:false, msg:'Form not found.' };
  } catch(e) { return { ok:false, msg:'Failed.' }; }
}

// ─────────────────────────────────────────────────────────────
//  ADMIN — SETTINGS
// ─────────────────────────────────────────────────────────────

function adminGetSettings(token) {
  if (!_validateAdmin(token)) return { ok:false, msg:'Access denied.' };
  try {
    const s = _getMasterSettings();
    // Never expose admin credential hashes
    delete s['_au']; delete s['_ap'];
    return { ok:true, settings: s };
  } catch(e) { return { ok:false, msg:'Failed.' }; }
}

function adminUpdateSettings(token, data) {
  if (!_validateAdmin(token)) return { ok:false, msg:'Access denied.' };
  // Block updating admin credential fields via this endpoint
  const blocked = ['_au', '_ap'];
  try {
    Object.keys(data).forEach(key => {
      if (blocked.includes(key)) return;
      _upsertMasterSetting(sanitize(key), sanitize(String(data[key])));
    });
    return { ok:true, msg:'Settings updated.' };
  } catch(e) { return { ok:false, msg:'Update failed.' }; }
}

// ─────────────────────────────────────────────────────────────
//  ADMIN — ANALYTICS
// ─────────────────────────────────────────────────────────────

function adminGetGrowthStats(token) {
  if (!_validateAdmin(token)) return { ok:false, msg:'Access denied.' };
  try {
    const data = getMasterTab(CONFIG.TABS.USERS).getDataRange().getValues();
    const byMonth = {};
    for (let i = 1; i < data.length; i++) {
      if (!data[i][USER_COLS.ID - 1]) continue;
      const mk = String(data[i][USER_COLS.REG_DATE - 1]).slice(0,7);
      if (!byMonth[mk]) byMonth[mk] = { total:0, premium:0 };
      byMonth[mk].total++;
      if (_getEffectivePlan(data[i]) === CONFIG.PLANS.PREMIUM) byMonth[mk].premium++;
    }
    return { ok:true, byMonth };
  } catch(e) { return { ok:false, msg:'Failed.' }; }
}

// ─────────────────────────────────────────────────────────────
//  ADMIN — BROADCAST MESSAGE (stored in AdminSettings for client to fetch)
// ─────────────────────────────────────────────────────────────

function adminBroadcast(token, msg) {
  if (!_validateAdmin(token)) return { ok:false, msg:'Access denied.' };
  try {
    _upsertMasterSetting('broadcast_msg', sanitize(msg || ''));
    _upsertMasterSetting('broadcast_date', now());
    return { ok:true, msg:'Broadcast set.' };
  } catch(e) { return { ok:false, msg:'Failed.' }; }
}

function getBroadcast() {
  // Public — shown to all logged-in users
  try {
    const s = _getMasterSettings();
    return { ok:true, msg: s['broadcast_msg'] || '', date: s['broadcast_date'] || '' };
  } catch(e) { return { ok:false, msg:'' }; }
}

// ─────────────────────────────────────────────────────────────
//  INTERNAL HELPERS
// ─────────────────────────────────────────────────────────────

function _syncUserSessions(userId, plan) {
  const sesTab  = getMasterTab(CONFIG.TABS.SESSIONS);
  const sesData = sesTab.getDataRange().getValues();
  for (let i = 1; i < sesData.length; i++) {
    if (sesData[i][SES_COLS.USER_ID - 1] === userId && sesData[i][SES_COLS.TYPE - 1] !== 'admin')
      sesTab.getRange(i+1, SES_COLS.PLAN).setValue(plan);
  }
}
