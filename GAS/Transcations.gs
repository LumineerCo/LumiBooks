// ============================================================
//  LumiBooks v2.0 — Transactions.gs
//  All transaction operations — server-side only.
//  © 2026 LumineerCo. All rights reserved.
// ============================================================

const TX_COLS = {
  DATE:1, TYPE:2, CATEGORY:3, AMOUNT:4, DESC:5,
  REF_NO:6, GST_RATE:7, GST_AMOUNT:8, TOTAL_GST:9, CREATED:10,
};

// ─────────────────────────────────────────────────────────────
//  ADD TRANSACTION
// ─────────────────────────────────────────────────────────────

function addTransaction(token, txData) {
  const ses = validateSession(token);
  if (!ses) return { ok:false, msg:'Session expired.' };

  if (ses.plan === CONFIG.PLANS.FREE) {
    const count = _countTxThisMonth(ses.sheetId);
    if (count >= CONFIG.FREE_LIMITS.TRANSACTIONS_PER_MONTH)
      return { ok:false, limit:true,
        msg:`Free limit reached (${CONFIG.FREE_LIMITS.TRANSACTIONS_PER_MONTH}/month). Upgrade to Premium.` };
  }

  try {
    const amount = parseFloat(txData.amount);
    if (isNaN(amount) || amount <= 0) return { ok:false, msg:'Invalid amount.' };

    const gstRate   = parseFloat(txData.gst_rate || 0);
    const gstAmt    = parseFloat(((amount * gstRate) / 100).toFixed(2));
    const totalGst  = parseFloat((amount + gstAmt).toFixed(2));
    const type      = txData.type === 'Credit' ? 'Credit' : 'Debit';

    getUserTab(ses.sheetId, CONFIG.USER_TABS.TRANSACTIONS).appendRow([
      sanitize(txData.date || todayStr()),
      type,
      sanitize(txData.category || 'Other'),
      amount,
      sanitize(txData.description || ''),
      sanitize(txData.reference_no || ''),
      gstRate, gstAmt, totalGst,
      now(),
    ]);
    return { ok:true, msg:'Transaction added.' };
  } catch(e) {
    Logger.log('addTransaction: ' + e);
    return { ok:false, msg:'Failed to save transaction.' };
  }
}

// ─────────────────────────────────────────────────────────────
//  GET TRANSACTIONS (paginated + filtered)
// ─────────────────────────────────────────────────────────────

function getTransactions(token, opts) {
  const ses = validateSession(token);
  if (!ses) return { ok:false, msg:'Session expired.' };

  opts = opts || {};
  const page     = parseInt(opts.page || 1);
  const pageSize = parseInt(opts.size || 25);
  const search   = (opts.search || '').toLowerCase().trim();
  const typeF    = opts.type   || '';
  const catF     = opts.cat    || '';
  const dateFrom = opts.from   || '';
  const dateTo   = opts.to     || '';

  try {
    const data = getUserTab(ses.sheetId, CONFIG.USER_TABS.TRANSACTIONS).getDataRange().getValues();
    let rows = [];

    for (let i = 1; i < data.length; i++) {
      const r = data[i];
      if (!r[TX_COLS.DATE - 1]) continue;
      const dateStr = String(r[TX_COLS.DATE - 1]).slice(0, 10);
      if (dateFrom && dateStr < dateFrom) continue;
      if (dateTo   && dateStr > dateTo  ) continue;
      const tx = {
        idx      : i,
        date     : dateStr,
        type     : r[TX_COLS.TYPE     - 1],
        category : r[TX_COLS.CATEGORY - 1],
        amount   : parseFloat(r[TX_COLS.AMOUNT    - 1]) || 0,
        desc     : r[TX_COLS.DESC     - 1],
        ref      : r[TX_COLS.REF_NO   - 1],
        gstRate  : parseFloat(r[TX_COLS.GST_RATE  - 1]) || 0,
        gstAmt   : parseFloat(r[TX_COLS.GST_AMOUNT- 1]) || 0,
        totalGst : parseFloat(r[TX_COLS.TOTAL_GST - 1]) || 0,
      };
      if (typeF && tx.type !== typeF) continue;
      if (catF  && tx.category !== catF) continue;
      if (search) {
        const h = [tx.date, tx.type, tx.category, tx.desc, tx.ref, String(tx.amount)].join(' ').toLowerCase();
        if (!h.includes(search)) continue;
      }
      rows.push(tx);
    }

    rows.sort((a, b) => new Date(b.date) - new Date(a.date));
    const total = rows.length;
    let credit = 0, debit = 0;
    rows.forEach(r => r.type === 'Credit' ? credit += r.amount : debit += r.amount);

    return {
      ok         : true,
      rows       : rows.slice((page - 1) * pageSize, page * pageSize),
      total, totalPages: Math.ceil(total / pageSize) || 1, page, credit, debit,
    };
  } catch(e) {
    Logger.log('getTransactions: ' + e);
    return { ok:false, msg:'Failed to fetch transactions.' };
  }
}

// ─────────────────────────────────────────────────────────────
//  UPDATE TRANSACTION
// ─────────────────────────────────────────────────────────────

function updateTransaction(token, rowIdx, txData) {
  const ses = validateSession(token);
  if (!ses) return { ok:false, msg:'Session expired.' };
  try {
    const tab    = getUserTab(ses.sheetId, CONFIG.USER_TABS.TRANSACTIONS);
    const rowNum = parseInt(rowIdx) + 1;
    if (rowNum < 2) return { ok:false, msg:'Invalid row.' };
    const amount  = parseFloat(txData.amount);
    if (isNaN(amount) || amount <= 0) return { ok:false, msg:'Invalid amount.' };
    const gstRate = parseFloat(txData.gst_rate || 0);
    const gstAmt  = parseFloat(((amount * gstRate) / 100).toFixed(2));
    tab.getRange(rowNum, TX_COLS.DATE    ).setValue(sanitize(txData.date || ''));
    tab.getRange(rowNum, TX_COLS.TYPE    ).setValue(txData.type === 'Credit' ? 'Credit' : 'Debit');
    tab.getRange(rowNum, TX_COLS.CATEGORY).setValue(sanitize(txData.category || 'Other'));
    tab.getRange(rowNum, TX_COLS.AMOUNT  ).setValue(amount);
    tab.getRange(rowNum, TX_COLS.DESC    ).setValue(sanitize(txData.description || ''));
    tab.getRange(rowNum, TX_COLS.REF_NO  ).setValue(sanitize(txData.reference_no || ''));
    tab.getRange(rowNum, TX_COLS.GST_RATE).setValue(gstRate);
    tab.getRange(rowNum, TX_COLS.GST_AMOUNT).setValue(gstAmt);
    tab.getRange(rowNum, TX_COLS.TOTAL_GST).setValue(parseFloat((amount + gstAmt).toFixed(2)));
    return { ok:true, msg:'Transaction updated.' };
  } catch(e) { return { ok:false, msg:'Update failed.' }; }
}

// ─────────────────────────────────────────────────────────────
//  DELETE TRANSACTION
// ─────────────────────────────────────────────────────────────

function deleteTransaction(token, rowIdx) {
  const ses = validateSession(token);
  if (!ses) return { ok:false, msg:'Session expired.' };
  try {
    const tab    = getUserTab(ses.sheetId, CONFIG.USER_TABS.TRANSACTIONS);
    const rowNum = parseInt(rowIdx) + 1;
    if (rowNum < 2) return { ok:false, msg:'Invalid row.' };
    tab.deleteRow(rowNum);
    return { ok:true, msg:'Transaction deleted.' };
  } catch(e) { return { ok:false, msg:'Delete failed.' }; }
}

// ─────────────────────────────────────────────────────────────
//  DASHBOARD SUMMARY
// ─────────────────────────────────────────────────────────────

function getDashboardSummary(token) {
  const ses = validateSession(token);
  if (!ses) return { ok:false, msg:'Session expired.' };

  try {
    const data = getUserTab(ses.sheetId, CONFIG.USER_TABS.TRANSACTIONS).getDataRange().getValues();
    const now_ = new Date();
    const curMonth = _monthKey(now_);
    const prevMonth = _monthKey(new Date(now_.getFullYear(), now_.getMonth() - 1, 1));
    let totalCredit = 0, totalDebit = 0, monthCredit = 0, monthDebit = 0, prevCredit = 0, prevDebit = 0;
    const monthMap = {}, catMap = {};

    for (let i = 1; i < data.length; i++) {
      const r = data[i];
      if (!r[0]) continue;
      const type = r[TX_COLS.TYPE   - 1];
      const amt  = parseFloat(r[TX_COLS.AMOUNT - 1]) || 0;
      const cat  = r[TX_COLS.CATEGORY - 1] || 'Other';
      const mk   = String(r[TX_COLS.DATE - 1]).slice(0, 7);

      type === 'Credit' ? totalCredit += amt : totalDebit += amt;

      if (!monthMap[mk]) monthMap[mk] = { credit:0, debit:0 };
      type === 'Credit' ? monthMap[mk].credit += amt : monthMap[mk].debit += amt;

      catMap[cat] = (catMap[cat] || 0) + amt;

      if (mk === curMonth)  type === 'Credit' ? monthCredit += amt : monthDebit += amt;
      if (mk === prevMonth) type === 'Credit' ? prevCredit  += amt : prevDebit  += amt;
    }

    const recent = [];
    for (let i = data.length - 1; i >= 1 && recent.length < 10; i--) {
      if (!data[i][0]) continue;
      recent.push({
        date: String(data[i][TX_COLS.DATE - 1]).slice(0,10),
        type: data[i][TX_COLS.TYPE - 1],
        category: data[i][TX_COLS.CATEGORY - 1],
        amount: parseFloat(data[i][TX_COLS.AMOUNT - 1]) || 0,
        desc: data[i][TX_COLS.DESC - 1],
      });
    }

    // Low stock count
    let lowStockCount = 0;
    try {
      const sd = getUserTab(ses.sheetId, CONFIG.USER_TABS.STOCK).getDataRange().getValues();
      for (let i = 1; i < sd.length; i++) {
        const qty     = parseFloat(sd[i][5]) || 0;
        const reorder = parseFloat(sd[i][8]) || 0;
        if (qty <= reorder) lowStockCount++;
      }
    } catch(e) {}

    return {
      ok: true, totalCredit, totalDebit,
      netProfit: totalCredit - totalDebit,
      monthCredit, monthDebit,
      monthProfit: monthCredit - monthDebit,
      prevCredit, prevDebit,
      recent, monthChart: monthMap, catChart: catMap,
      plan: ses.plan,
      txThisMonth: _countTxThisMonth(ses.sheetId),
      txLimit: ses.plan === CONFIG.PLANS.FREE ? CONFIG.FREE_LIMITS.TRANSACTIONS_PER_MONTH : -1,
      lowStockCount,
    };
  } catch(e) {
    Logger.log('getDashboardSummary: ' + e);
    return { ok:false, msg:'Failed to load dashboard.' };
  }
}

// ─────────────────────────────────────────────────────────────
//  GLOBAL SEARCH
// ─────────────────────────────────────────────────────────────

function globalSearch(token, query) {
  const ses = validateSession(token);
  if (!ses) return { ok:false, msg:'Session expired.' };
  if (!query || query.trim().length < 2) return { ok:true, results:[] };

  const q = query.toLowerCase().trim();
  const results = [];

  try {
    const txD = getUserTab(ses.sheetId, CONFIG.USER_TABS.TRANSACTIONS).getDataRange().getValues();
    for (let i = 1; i < txD.length; i++) {
      if (!txD[i][0]) continue;
      if (txD[i].join(' ').toLowerCase().includes(q)) {
        results.push({ type:'transaction', date: String(txD[i][0]).slice(0,10),
          txType:txD[i][1], category:txD[i][2], amount:txD[i][3], desc:txD[i][4] });
      }
    }
    const estD = getUserTab(ses.sheetId, CONFIG.USER_TABS.ESTIMATES).getDataRange().getValues();
    for (let i = 1; i < estD.length; i++) {
      if (!estD[i][0]) continue;
      if (estD[i].join(' ').toLowerCase().includes(q)) {
        results.push({ type:'estimate', billId:estD[i][0], date:estD[i][1], customer:estD[i][3], totalAmount:estD[i][13] });
      }
    }
    const cuD = getUserTab(ses.sheetId, CONFIG.USER_TABS.CUSTOMERS).getDataRange().getValues();
    for (let i = 1; i < cuD.length; i++) {
      if (!cuD[i][0]) continue;
      if (cuD[i].join(' ').toLowerCase().includes(q)) {
        results.push({ type:'customer', custId:cuD[i][0], name:cuD[i][1], phone:cuD[i][2], email:cuD[i][3] });
      }
    }
    return { ok:true, results: results.slice(0, 40) };
  } catch(e) { return { ok:false, msg:'Search failed.' }; }
}

// ─────────────────────────────────────────────────────────────
//  CATEGORIES CRUD
// ─────────────────────────────────────────────────────────────

function getCategories(token) {
  const ses = validateSession(token);
  if (!ses) return { ok:false, msg:'Session expired.' };
  try {
    const data = getUserTab(ses.sheetId, CONFIG.USER_TABS.CATEGORIES).getDataRange().getValues();
    const rows = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      rows.push({ id:data[i][0], name:data[i][1], type:data[i][2], color:data[i][3], icon:data[i][4], isCustom:data[i][5] === 'true' });
    }
    return { ok:true, rows };
  } catch(e) { return { ok:false, msg:'Failed to fetch categories.' }; }
}

function addCategory(token, data) {
  const ses = validateSession(token);
  if (!ses) return { ok:false, msg:'Session expired.' };
  if (!data.name || !data.type) return { ok:false, msg:'Name and type required.' };
  if (!['income','expense'].includes(data.type)) return { ok:false, msg:'Invalid type.' };
  try {
    const id = generateId('CAT');
    getUserTab(ses.sheetId, CONFIG.USER_TABS.CATEGORIES).appendRow([
      id, sanitize(data.name), data.type,
      sanitize(data.color || '#6b7280'),
      sanitize(data.icon  || '📌'),
      'true', now()
    ]);
    return { ok:true, msg:'Category added.', id };
  } catch(e) { return { ok:false, msg:'Failed to add category.' }; }
}

function updateCategory(token, catId, data) {
  const ses = validateSession(token);
  if (!ses) return { ok:false, msg:'Session expired.' };
  try {
    const tab  = getUserTab(ses.sheetId, CONFIG.USER_TABS.CATEGORIES);
    const rows = tab.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] !== catId) continue;
      if (data.name)  tab.getRange(i+1,2).setValue(sanitize(data.name));
      if (data.color) tab.getRange(i+1,4).setValue(sanitize(data.color));
      if (data.icon)  tab.getRange(i+1,5).setValue(sanitize(data.icon));
      return { ok:true, msg:'Category updated.' };
    }
    return { ok:false, msg:'Category not found.' };
  } catch(e) { return { ok:false, msg:'Update failed.' }; }
}

function deleteCategory(token, catId) {
  const ses = validateSession(token);
  if (!ses) return { ok:false, msg:'Session expired.' };
  try {
    const tab  = getUserTab(ses.sheetId, CONFIG.USER_TABS.CATEGORIES);
    const rows = tab.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] !== catId) continue;
      if (rows[i][5] !== 'true') return { ok:false, msg:'Cannot delete default category.' };
      tab.deleteRow(i + 1);
      return { ok:true, msg:'Category deleted.' };
    }
    return { ok:false, msg:'Not found.' };
  } catch(e) { return { ok:false, msg:'Delete failed.' }; }
}

// ─────────────────────────────────────────────────────────────
//  INTERNAL HELPERS
// ─────────────────────────────────────────────────────────────

function _monthKey(d) { return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0'); }

function _countTxThisMonth(sheetId) {
  const curMonth = _monthKey(new Date());
  const data = getUserTab(sheetId, CONFIG.USER_TABS.TRANSACTIONS).getDataRange().getValues();
  let count = 0;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][TX_COLS.DATE - 1]).slice(0, 7) === curMonth) count++;
  }
  return count;
}
