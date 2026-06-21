// ============================================================
//  LumiBooks v2.0 — Code.gs
//  Main GAS entry point. doGet() router. All logic in modules.
//  © 2026 LumineerCo. All rights reserved.
// ============================================================

function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'home';

  const PUBLIC = ['home','login','register','pricing','about','contact','terms','privacy'];
  const AUTH   = ['dashboard','transactions','estimates','billing','stock','customers',
                  'staff','attendance','payments','reports','profile','upgrade','referral'];
  const ADMIN  = ['admin'];

  let tpl  = 'landing';
  let title = 'LumiBooks — Smart Business Management';

  if (PUBLIC.includes(page)) {
    tpl   = 'landing';
    title = 'LumiBooks — Smart Business Management';
  } else if (AUTH.includes(page)) {
    tpl   = 'app';
    title = 'LumiBooks — ' + _cap(page);
  } else if (ADMIN.includes(page)) {
    tpl   = 'admin';
    title = 'LumiBooks Admin';
  } else {
    tpl   = 'landing';
  }

  const t     = HtmlService.createTemplateFromFile(tpl);
  t.appName   = CONFIG.APP_NAME;
  t.brand     = CONFIG.BRAND;
  t.version   = CONFIG.VERSION;
  t.page      = page;
  t.title     = title;
  t.upiId     = CONFIG.UPI_ID;
  t.prices    = JSON.stringify(CONFIG.PRICES);
  t.gstRates  = JSON.stringify(CONFIG.GST_RATES);

  return t.evaluate()
    .setTitle(title)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport','width=device-width, initial-scale=1.0');
}

function doPost(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok:false, msg:'Use google.script.run' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function include(fn) {
  return HtmlService.createHtmlOutputFromFile(fn).getContent();
}

function _cap(s) { return s.charAt(0).toUpperCase() + s.slice(1); }

// ─────────────────────────────────────────────────────────────
//  PUBLIC API MANIFEST
//  (All functions are callable via google.script.run)
// ─────────────────────────────────────────────────────────────
//
// ── Auth (Auth.gs) ──────────────────────────────────────────
// registerUser(data)
// loginUser(email, pass)
// logoutUser(token)
// changePassword(token, old, new)
// getProfile(token)
// updateProfile(token, data)
// submitContactForm(data)          ← public, no auth
//
// ── Transactions (Transactions.gs) ──────────────────────────
// addTransaction(token, data)
// getTransactions(token, opts)
// updateTransaction(token, idx, data)
// deleteTransaction(token, idx)
// getDashboardSummary(token)
// globalSearch(token, query)
// getCategories(token)
// addCategory(token, data)
// updateCategory(token, id, data)
// deleteCategory(token, id)
//
// ── Estimates / Bills (Estimates.gs) ────────────────────────
// saveEstimate(token, data)
// getEstimate(token, billId)
// listEstimates(token, opts)
// updateEstimate(token, billId, data)
// updateEstimateStatus(token, billId, status)
// deleteEstimate(token, billId)
// buildWhatsAppMessage(token, billId)
//
// ── Export (Export.gs) ──────────────────────────────────────
// generatePDF(token, billId)
// generateDOC(token, billId)
//
// ── Stock (Stock.gs) ────────────────────────────────────────
// addStockItem(token, data)
// getStockItems(token, opts)
// updateStockItem(token, id, data)
// deleteStockItem(token, id)
// adjustStock(token, id, qty, reason)
// getLowStockItems(token)
//
// ── Customers (Customers.gs) ────────────────────────────────
// addCustomer(token, data)
// getCustomers(token, opts)
// updateCustomer(token, id, data)
// deleteCustomer(token, id)
// getCustomerStatement(token, custId)
//
// ── Staff (Staff.gs) ────────────────────────────────────────
// addStaff(token, data)
// getStaff(token, opts)
// updateStaff(token, id, data)
// deleteStaff(token, id)
// markAttendance(token, data)
// getAttendance(token, opts)
// addStaffPayment(token, data)
// getStaffPayments(token, opts)
//
// ── Payments (Payments.gs) ──────────────────────────────────
// addCustomerPayment(token, data)
// getCustomerPayments(token, opts)
// deleteCustomerPayment(token, id)
// submitPaymentRequest(token, data)
// getPaymentDetails(token)
//
// ── Reports (Reports.gs) ────────────────────────────────────
// getReport(token, opts)
// getStockReport(token, opts)
// getCustomerReport(token, opts)
// getStaffReport(token, opts)
//
// ── Referral (Referral.gs) ──────────────────────────────────
// getReferralInfo(token)
//
// ── Admin (Admin.gs) ────────────────────────────────────────
// adminLogin(username, password)
// adminLogout(token)
// adminGetDashboard(token)
// adminGetUsers(token, opts)
// adminSetUserStatus(token, uid, status)
// adminSetUserPlan(token, uid, plan, days)
// adminGetPendingPayments(token)
// adminApprovePayment(token, txId, remarks, days)
// adminRejectPayment(token, txId, remarks)
// adminGetReferrals(token)
// adminGetContactForms(token)
// adminReplyContact(token, formId, reply)
// adminGetSettings(token)
// adminUpdateSettings(token, data)
// adminGetAllTransactions(token, opts)
// adminBroadcast(token, msg)
