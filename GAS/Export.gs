// ============================================================
//  LumiBooks v2.0 — Export.gs
//  PDF & DOC generation. Premium only (server-side enforced).
//  5 Templates: classic, modern, professional, executive, gst
//  © 2026 LumineerCo. All rights reserved.
// ============================================================

function generatePDF(token, billId) {
  const ses = validateSession(token);
  if (!ses) return { ok:false, msg:'Session expired.' };
  if (ses.plan !== CONFIG.PLANS.PREMIUM) return { ok:false, msg:'PDF export is a Premium feature.', upgrade:true };
  try {
    const est = getEstimate(token, billId);
    if (!est.ok) return est;
    const doc  = _buildDoc(est);
    const pdf  = doc.getAs('application/pdf');
    DriveApp.getFileById(doc.getId()).setTrashed(true);
    return { ok:true, base64: Utilities.base64Encode(pdf.getBytes()),
      filename: billId + '.pdf', mimeType: 'application/pdf' };
  } catch(e) {
    Logger.log('generatePDF: ' + e);
    return { ok:false, msg:'PDF generation failed.' };
  }
}

function generateDOC(token, billId) {
  const ses = validateSession(token);
  if (!ses) return { ok:false, msg:'Session expired.' };
  if (ses.plan !== CONFIG.PLANS.PREMIUM) return { ok:false, msg:'DOC export is a Premium feature.', upgrade:true };
  try {
    const est  = getEstimate(token, billId);
    if (!est.ok) return est;
    const doc  = _buildDoc(est);
    const blob = doc.getAs('application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    DriveApp.getFileById(doc.getId()).setTrashed(true);
    return { ok:true, base64: Utilities.base64Encode(blob.getBytes()),
      filename: billId + '.docx',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' };
  } catch(e) {
    Logger.log('generateDOC: ' + e);
    return { ok:false, msg:'DOC generation failed.' };
  }
}

// ─────────────────────────────────────────────────────────────
//  INTERNAL — Build Google Doc from bill data
// ─────────────────────────────────────────────────────────────

function _buildDoc(est) {
  const tpl = est.template || 'classic';
  const doc  = DocumentApp.create('LumiBooks_Export_' + est.billId);
  const body = doc.getBody();
  body.setMarginTop(36).setMarginBottom(36).setMarginLeft(54).setMarginRight(54);

  const biz  = est.biz || {};
  const curr = biz.currency || '₹';
  const fmt  = n => curr + parseFloat(n).toLocaleString('en-IN', { minimumFractionDigits:2, maximumFractionDigits:2 });

  // ── HEADER: Business Name, Address, Phone/Email (fixed layout) ──
  const bizNameP = body.appendParagraph(biz.business_name || 'My Business');
  bizNameP.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  bizNameP.editAsText().setFontSize(20).setBold(true).setForegroundColor('#1a1a2e');

  if (biz.business_address) {
    const addrP = body.appendParagraph(biz.business_address);
    addrP.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    addrP.editAsText().setFontSize(10).setForegroundColor('#555555');
    addrP.setSpacingAfter(0);
  }

  const contactLine = [biz.business_phone, biz.business_email].filter(Boolean).join(' | ');
  if (contactLine) {
    const cP = body.appendParagraph(contactLine);
    cP.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    cP.editAsText().setFontSize(10).setForegroundColor('#555555');
    cP.setSpacingAfter(0);
  }

  if (biz.gstin) {
    const gP = body.appendParagraph('GSTIN: ' + biz.gstin);
    gP.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    gP.editAsText().setFontSize(9).setForegroundColor('#888888');
  }

  body.appendHorizontalRule();

  // ── BILL TITLE ──
  const docTitle = tpl === 'gst' ? 'TAX INVOICE' : (tpl === 'executive' ? 'INVOICE' : 'BILL / ESTIMATE');
  const titleP = body.appendParagraph(docTitle);
  titleP.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  titleP.editAsText().setFontSize(14).setBold(true).setForegroundColor('#c8a96e');

  // ── BILL META ──
  const metaRows = [
    ['Bill No', est.billId],
    ['Date', est.date],
  ];
  if (est.due_date) metaRows.push(['Due Date', est.due_date]);
  if (tpl === 'gst' && biz.state_code) metaRows.push(['State', biz.state_code]);

  const metaTable = body.appendTable(metaRows.map(r => [r[0] + ':', r[1]]));
  metaTable.setBorderWidth(0);
  for (let i = 0; i < metaRows.length; i++) {
    metaTable.getRow(i).getCell(0).editAsText().setBold(true).setFontSize(9);
    metaTable.getRow(i).getCell(1).editAsText().setFontSize(9);
  }

  body.appendParagraph('');

  // ── CUSTOMER INFO ──
  if (est.cust_name) {
    const toP = body.appendParagraph('BILL TO:');
    toP.editAsText().setBold(true).setFontSize(10).setForegroundColor('#333333');
    const custP = body.appendParagraph(est.cust_name);
    custP.editAsText().setFontSize(11).setBold(true);
    custP.setSpacingAfter(0);
    const custDetails = [est.cust_phone, est.cust_addr].filter(Boolean);
    if (custDetails.length) {
      const d = body.appendParagraph(custDetails.join(' | '));
      d.editAsText().setFontSize(9).setForegroundColor('#555555');
      d.setSpacingAfter(0);
    }
    if (est.cust_gstin) {
      const g = body.appendParagraph('GSTIN: ' + est.cust_gstin);
      g.editAsText().setFontSize(9).setForegroundColor('#888888');
    }
    body.appendParagraph('');
  }

  // ── ITEMS TABLE ──
  const gstEnabled = est.gst_enabled;
  const headers = gstEnabled
    ? ['#', 'Description', 'HSN', 'Qty', 'Unit Price', 'GST%', 'GST Amt', 'Amount']
    : ['#', 'Description', 'Qty', 'Unit Price', 'Amount'];

  const tableData = [headers];
  (est.items || []).forEach((it, idx) => {
    if (gstEnabled) {
      tableData.push([String(idx+1), it.desc||'', it.hsn||'', String(it.qty),
        fmt(it.price), (it.gst_rate||0)+'%', fmt(it.gst_amount||0), fmt(it.amount)]);
    } else {
      tableData.push([String(idx+1), it.desc||'', String(it.qty), fmt(it.price), fmt(it.amount)]);
    }
  });

  const table = body.appendTable(tableData);
  const hRow  = table.getRow(0);
  for (let c = 0; c < headers.length; c++) {
    hRow.getCell(c).editAsText().setBold(true).setFontSize(9);
    hRow.getCell(c).setBackgroundColor('#f0f0f0');
  }
  for (let r = 1; r < tableData.length; r++) {
    for (let c = 0; c < headers.length; c++) {
      table.getRow(r).getCell(c).editAsText().setFontSize(9);
    }
  }

  body.appendParagraph('');

  // ── TOTALS ──
  const addTotalRow = (label, value, bold, color) => {
    const p = body.appendParagraph(label + ': ' + value);
    p.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    p.editAsText().setBold(bold || false).setFontSize(bold ? 11 : 10)
      .setForegroundColor(color || '#333333');
    p.setSpacingAfter(2);
  };

  addTotalRow('Subtotal', fmt(est.subtotal));
  if (est.discount && est.discount > 0) addTotalRow('Discount (-)', fmt(est.discount));
  if (gstEnabled && est.gst_amount > 0) {
    if (est.gst_type === 'CGST+SGST') {
      addTotalRow('CGST (' + (est.gst_pct/2) + '%)', fmt(est.gst_amount/2));
      addTotalRow('SGST (' + (est.gst_pct/2) + '%)', fmt(est.gst_amount/2));
    } else {
      addTotalRow('IGST (' + est.gst_pct + '%)', fmt(est.gst_amount));
    }
  }
  addTotalRow('TOTAL', fmt(est.total), true, '#1a1a2e');

  // ── NOTES ──
  if (est.notes) {
    body.appendParagraph('');
    const nP = body.appendParagraph('Notes: ' + est.notes);
    nP.editAsText().setItalic(true).setFontSize(9).setForegroundColor('#666666');
  }

  // ── FOOTER ──
  body.appendHorizontalRule();
  const footer = body.appendParagraph('This is a computer-generated document. | Powered by LumiBooks — LumineerCo');
  footer.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  footer.editAsText().setFontSize(8).setItalic(true).setForegroundColor('#999999');

  doc.saveAndClose();
  return doc;
}
