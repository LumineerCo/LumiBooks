// ============================================================
//  LumiBooks v2.0 — Estimates.gs
//  Bill/Estimate CRUD — server-side only.
//  Templates: classic, modern, professional, executive, gst
//  © 2026 LumineerCo. All rights reserved.
// ============================================================

const EST = {
  BILL_ID:1, DATE:2, DUE:3, CUST_NAME:4, CUST_PHONE:5, CUST_ADDR:6, CUST_GSTIN:7,
  ITEMS:8, SUBTOTAL:9, DISCOUNT:10, GST_TYPE:11, GST_PCT:12, GST_AMT:13, TOTAL:14,
  NOTES:15, STATUS:16, TEMPLATE:17, LOCKED_AT:18, GST_ENABLED:19, CREATED:20,
};

// ─────────────────────────────────────────────────────────────
//  SAVE ESTIMATE
// ─────────────────────────────────────────────────────────────

function saveEstimate(token, estData) {
  const ses = validateSession(token);
  if (!ses) return { ok:false, msg:'Session expired.' };

  // Template gate
  const tpl = estData.template || 'classic';
  const limits = _getLimits(ses.plan);
  if (!limits.BILL_TEMPLATES.includes(tpl))
    return { ok:false, msg:'This template requires Premium.', upgrade:true };

  try {
    const items = estData.items || [];
    if (!items.length) return { ok:false, msg:'At least one item is required.' };

    let subtotal = 0;
    const validated = items.map(it => {
      const qty   = parseFloat(it.qty)   || 0;
      const price = parseFloat(it.price) || 0;
      const hsn   = sanitize(it.hsn   || '');
      const gstR  = parseFloat(it.gst_rate || 0);
      const amt   = parseFloat((qty * price).toFixed(2));
      const gstA  = parseFloat((amt * gstR / 100).toFixed(2));
      subtotal   += amt;
      return { desc: sanitize(it.desc || ''), qty, price, amount: amt, hsn, gst_rate: gstR, gst_amount: gstA };
    });

    subtotal       = parseFloat(subtotal.toFixed(2));
    const discount = parseFloat(estData.discount || 0);
    const gstEnabled = estData.gst_enabled ? 'true' : 'false';
    const gstType  = sanitize(estData.gst_type  || 'CGST+SGST');
    const gstPct   = parseFloat(estData.gst_pct || 0);
    const afterD   = parseFloat((subtotal - discount).toFixed(2));
    const gstAmt   = gstEnabled === 'true' ? parseFloat((afterD * gstPct / 100).toFixed(2)) : 0;
    const total    = parseFloat((afterD + gstAmt).toFixed(2));

    const billId   = 'BILL' + Date.now().toString(36).toUpperCase();
    const lockTime = now();

    getUserTab(ses.sheetId, CONFIG.USER_TABS.ESTIMATES).appendRow([
      billId,
      sanitize(estData.date     || todayStr()),
      sanitize(estData.due_date || ''),
      sanitize(estData.cust_name  || ''),
      sanitize(estData.cust_phone || ''),
      sanitize(estData.cust_addr  || ''),
      sanitize(estData.cust_gstin || ''),
      JSON.stringify(validated),
      subtotal, discount, gstType, gstPct, gstAmt, total,
      sanitize(estData.notes || ''),
      'draft', tpl, lockTime, gstEnabled, now(),
    ]);

    return { ok:true, msg:'Bill saved.', billId };
  } catch(e) {
    Logger.log('saveEstimate: ' + e);
    return { ok:false, msg:'Failed to save bill.' };
  }
}

// ─────────────────────────────────────────────────────────────
//  GET ESTIMATE
// ─────────────────────────────────────────────────────────────

function getEstimate(token, billId) {
  const ses = validateSession(token);
  if (!ses) return { ok:false, msg:'Session expired.' };
  try {
    const tab  = getUserTab(ses.sheetId, CONFIG.USER_TABS.ESTIMATES);
    const data = tab.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][EST.BILL_ID - 1] !== billId) continue;
      const r = data[i];
      const lockedAt = r[EST.LOCKED_AT - 1];
      return {
        ok          : true,
        billId      : r[EST.BILL_ID   - 1],
        date        : String(r[EST.DATE      - 1]).slice(0,10),
        due_date    : r[EST.DUE       - 1],
        cust_name   : r[EST.CUST_NAME - 1],
        cust_phone  : r[EST.CUST_PHONE- 1],
        cust_addr   : r[EST.CUST_ADDR - 1],
        cust_gstin  : r[EST.CUST_GSTIN- 1],
        items       : JSON.parse(r[EST.ITEMS   - 1] || '[]'),
        subtotal    : r[EST.SUBTOTAL  - 1],
        discount    : r[EST.DISCOUNT  - 1],
        gst_type    : r[EST.GST_TYPE  - 1],
        gst_pct     : r[EST.GST_PCT   - 1],
        gst_amount  : r[EST.GST_AMT   - 1],
        total       : r[EST.TOTAL     - 1],
        notes       : r[EST.NOTES     - 1],
        status      : r[EST.STATUS    - 1],
        template    : r[EST.TEMPLATE  - 1] || 'classic',
        lockedAt,
        isLocked    : _isBillLocked(lockedAt, ses.plan),
        gst_enabled : r[EST.GST_ENABLED - 1] === 'true',
        biz         : _getBusinessSettings(ses.sheetId),
        rowIdx      : i,
      };
    }
    return { ok:false, msg:'Bill not found.' };
  } catch(e) { return { ok:false, msg:'Failed to fetch bill.' }; }
}

// ─────────────────────────────────────────────────────────────
//  UPDATE ESTIMATE (with lock + premium template check)
// ─────────────────────────────────────────────────────────────

function updateEstimate(token, billId, estData) {
  const ses = validateSession(token);
  if (!ses) return { ok:false, msg:'Session expired.' };

  const existing = getEstimate(token, billId);
  if (!existing.ok) return existing;
  if (existing.isLocked) return { ok:false, msg:'Bill is locked. Edit window has expired.', locked:true };

  const tpl    = estData.template || existing.template || 'classic';
  const limits = _getLimits(ses.plan);
  if (!limits.BILL_TEMPLATES.includes(tpl))
    return { ok:false, msg:'This template requires Premium.', upgrade:true };

  try {
    const tab  = getUserTab(ses.sheetId, CONFIG.USER_TABS.ESTIMATES);
    const data = tab.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][EST.BILL_ID - 1] !== billId) continue;

      const items = estData.items || [];
      let subtotal = 0;
      const validated = items.map(it => {
        const qty   = parseFloat(it.qty)   || 0;
        const price = parseFloat(it.price) || 0;
        const amt   = parseFloat((qty * price).toFixed(2));
        subtotal   += amt;
        return { desc: sanitize(it.desc || ''), qty, price, amount: amt,
          hsn: sanitize(it.hsn || ''), gst_rate: parseFloat(it.gst_rate || 0),
          gst_amount: parseFloat((amt * (it.gst_rate || 0) / 100).toFixed(2)) };
      });
      subtotal = parseFloat(subtotal.toFixed(2));
      const discount = parseFloat(estData.discount || 0);
      const gstEnabled = estData.gst_enabled ? 'true' : 'false';
      const gstPct  = parseFloat(estData.gst_pct || 0);
      const afterD  = parseFloat((subtotal - discount).toFixed(2));
      const gstAmt  = gstEnabled === 'true' ? parseFloat((afterD * gstPct / 100).toFixed(2)) : 0;
      const total   = parseFloat((afterD + gstAmt).toFixed(2));

      tab.getRange(i+1, EST.DATE      ).setValue(sanitize(estData.date || data[i][EST.DATE - 1]));
      tab.getRange(i+1, EST.DUE       ).setValue(sanitize(estData.due_date || ''));
      tab.getRange(i+1, EST.CUST_NAME ).setValue(sanitize(estData.cust_name || ''));
      tab.getRange(i+1, EST.CUST_PHONE).setValue(sanitize(estData.cust_phone || ''));
      tab.getRange(i+1, EST.CUST_ADDR ).setValue(sanitize(estData.cust_addr || ''));
      tab.getRange(i+1, EST.CUST_GSTIN).setValue(sanitize(estData.cust_gstin || ''));
      tab.getRange(i+1, EST.ITEMS     ).setValue(JSON.stringify(validated));
      tab.getRange(i+1, EST.SUBTOTAL  ).setValue(subtotal);
      tab.getRange(i+1, EST.DISCOUNT  ).setValue(discount);
      tab.getRange(i+1, EST.GST_TYPE  ).setValue(sanitize(estData.gst_type || 'CGST+SGST'));
      tab.getRange(i+1, EST.GST_PCT   ).setValue(gstPct);
      tab.getRange(i+1, EST.GST_AMT   ).setValue(gstAmt);
      tab.getRange(i+1, EST.TOTAL     ).setValue(total);
      tab.getRange(i+1, EST.NOTES     ).setValue(sanitize(estData.notes || ''));
      tab.getRange(i+1, EST.TEMPLATE  ).setValue(tpl);
      tab.getRange(i+1, EST.GST_ENABLED).setValue(gstEnabled);
      return { ok:true, msg:'Bill updated.' };
    }
    return { ok:false, msg:'Bill not found.' };
  } catch(e) { return { ok:false, msg:'Update failed.' }; }
}

// ─────────────────────────────────────────────────────────────
//  LIST ESTIMATES (paginated)
// ─────────────────────────────────────────────────────────────

function listEstimates(token, opts) {
  const ses = validateSession(token);
  if (!ses) return { ok:false, msg:'Session expired.' };
  opts = opts || {};
  const page = parseInt(opts.page || 1), size = parseInt(opts.size || 15);
  const search = (opts.search || '').toLowerCase().trim();
  const status = opts.status || '';
  try {
    const data = getUserTab(ses.sheetId, CONFIG.USER_TABS.ESTIMATES).getDataRange().getValues();
    let rows = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      const r = data[i];
      const est = {
        idx: i, billId: r[EST.BILL_ID - 1], date: String(r[EST.DATE - 1]).slice(0,10),
        custName: r[EST.CUST_NAME - 1], total: r[EST.TOTAL - 1],
        status: r[EST.STATUS - 1], template: r[EST.TEMPLATE - 1] || 'classic',
      };
      if (status && est.status !== status) continue;
      if (search) {
        const h = [est.billId, est.custName, est.total, est.status].join(' ').toLowerCase();
        if (!h.includes(search)) continue;
      }
      rows.push(est);
    }
    rows.sort((a, b) => new Date(b.date) - new Date(a.date));
    return {
      ok:true, total:rows.length, totalPages:Math.ceil(rows.length/size)||1, page,
      rows: rows.slice((page-1)*size, page*size),
    };
  } catch(e) { return { ok:false, msg:'Failed to list bills.' }; }
}

// ─────────────────────────────────────────────────────────────
//  UPDATE STATUS
// ─────────────────────────────────────────────────────────────

function updateEstimateStatus(token, billId, status) {
  const ses = validateSession(token);
  if (!ses) return { ok:false, msg:'Session expired.' };
  const allowed = ['draft','sent','paid','cancelled'];
  if (!allowed.includes(status)) return { ok:false, msg:'Invalid status.' };
  try {
    const tab  = getUserTab(ses.sheetId, CONFIG.USER_TABS.ESTIMATES);
    const data = tab.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][EST.BILL_ID - 1] !== billId) continue;
      tab.getRange(i+1, EST.STATUS).setValue(status);
      return { ok:true, msg:'Status updated.' };
    }
    return { ok:false, msg:'Bill not found.' };
  } catch(e) { return { ok:false, msg:'Update failed.' }; }
}

// ─────────────────────────────────────────────────────────────
//  DELETE
// ─────────────────────────────────────────────────────────────

function deleteEstimate(token, billId) {
  const ses = validateSession(token);
  if (!ses) return { ok:false, msg:'Session expired.' };
  try {
    const tab  = getUserTab(ses.sheetId, CONFIG.USER_TABS.ESTIMATES);
    const data = tab.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][EST.BILL_ID - 1] !== billId) continue;
      tab.deleteRow(i + 1);
      return { ok:true, msg:'Bill deleted.' };
    }
    return { ok:false, msg:'Bill not found.' };
  } catch(e) { return { ok:false, msg:'Delete failed.' }; }
}

// ─────────────────────────────────────────────────────────────
//  WHATSAPP MESSAGE BUILDER (FIXED)
// ─────────────────────────────────────────────────────────────

function buildWhatsAppMessage(token, billId) {
  const ses = validateSession(token);
  if (!ses) return { ok:false, msg:'Session expired.' };

  const est = getEstimate(token, billId);
  if (!est.ok) return est;

  const biz  = est.biz || {};
  const curr = biz.currency || '₹';
  const bizName = biz.business_name || 'Invoice';

  const fmt = n => curr + parseFloat(n).toLocaleString('en-IN', { minimumFractionDigits:2, maximumFractionDigits:2 });

  const itemLines = (est.items || []).map((it, idx) =>
    `  ${idx+1}. ${it.desc} × ${it.qty} @ ${fmt(it.price)} = *${fmt(it.amount)}*`
  ).join('\n');

  const lines = [
    `🧾 *${bizName}*`,
    biz.business_address ? `📍 ${biz.business_address}` : '',
    (biz.business_phone || biz.business_email)
      ? `📞 ${[biz.business_phone, biz.business_email].filter(Boolean).join(' | ')}` : '',
    biz.gstin ? `GSTIN: ${biz.gstin}` : '',
    '',
    `*Bill No:* ${est.billId}`,
    `*Date:* ${est.date}`,
    est.due_date ? `*Due Date:* ${est.due_date}` : '',
    est.cust_name ? `*To:* ${est.cust_name}` : '',
    est.cust_gstin ? `*GSTIN:* ${est.cust_gstin}` : '',
    '',
    '*Items:*',
    itemLines,
    '',
    `Subtotal: ${fmt(est.subtotal)}`,
    est.discount > 0 ? `Discount: -${fmt(est.discount)}` : '',
    est.gst_enabled && est.gst_amount > 0 ? `GST (${est.gst_pct}%): ${fmt(est.gst_amount)}` : '',
    `*Total: ${fmt(est.total)}*`,
    est.notes ? `\n📝 _${est.notes}_` : '',
    '',
    `_Powered by LumiBooks — LumineerCo_`,
  ].filter(l => l !== '' && l !== undefined);

  const message = lines.join('\n');

  // Clean phone: remove non-digits, ensure 10-digit Indian number
  let phone = '';
  if (est.cust_phone) {
    const cleaned = String(est.cust_phone).replace(/\D/g, '');
    if (cleaned.length === 10) phone = '91' + cleaned;
    else if (cleaned.length === 12 && cleaned.startsWith('91')) phone = cleaned;
    else if (cleaned.length > 0) phone = cleaned;
  }

  return { ok:true, message, phone, billId };
}

// ─────────────────────────────────────────────────────────────
//  LOCK HELPER
// ─────────────────────────────────────────────────────────────

function _isBillLocked(lockedAt, plan) {
  if (!lockedAt) return false;
  const limits  = _getLimits(plan || CONFIG.PLANS.FREE);
  const minutes = limits.BILL_EDIT_MINUTES || 60;
  return (new Date() - new Date(lockedAt)) > minutes * 60000;
}
