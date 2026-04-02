// ============================================================
//  LumiBooks — js/api.js
//  All GAS fetch calls go through here. No business logic.
//  Set GAS_URL before using.
// ============================================================

const LB = (() => {
  // ── Set your deployed GAS Web App URL here ──────────────────
  const GAS_URL = "https://script.google.com/macros/s/AKfycbwnTtykwBHX2UKgd4jH83uguFGgi18X_XqhhVzLAdnL6n14Ip9X5L-Ppb9Y61g9Twr3XQ/exec";

  const TOKEN_KEY = "lb_token";
  const NAME_KEY  = "lb_name";
  const PLAN_KEY  = "lb_plan";
  const EMAIL_KEY = "lb_email";

  // ── Core fetch ───────────────────────────────────────────────
  // GAS requires text/plain to avoid CORS preflight
  async function call(action, data = {}, token = null) {
    try {
      const body = JSON.stringify({
        action,
        token : token || localStorage.getItem(TOKEN_KEY),
        data,
      });

      const res = await fetch(GAS_URL, {
        method      : "POST",
        // text/plain avoids CORS preflight on GAS
        headers     : { "Content-Type": "text/plain;charset=utf-8" },
        body,
        redirect    : "follow",
      });

      if (!res.ok) return { ok: false, msg: "Network error " + res.status };
      const json = await res.json().catch(() => ({ ok: false, msg: "Bad response from server" }));

      // Session expired — clear storage
      if (json.auth === false) {
        LB.session.clear();
        window.location.href = "/";
        return null;
      }
      return json;

    } catch (e) {
      console.error("API error:", e);
      return { ok: false, msg: "Connection failed. Check your internet." };
    }
  }

  // ── Session helpers ──────────────────────────────────────────
  const session = {
    save(d)  { localStorage.setItem(TOKEN_KEY, d.token); localStorage.setItem(NAME_KEY, d.name); localStorage.setItem(PLAN_KEY, d.plan); localStorage.setItem(EMAIL_KEY, d.email || ""); },
    clear()  { [TOKEN_KEY, NAME_KEY, PLAN_KEY, EMAIL_KEY].forEach(k => localStorage.removeItem(k)); },
    token()  { return localStorage.getItem(TOKEN_KEY); },
    name()   { return localStorage.getItem(NAME_KEY) || "User"; },
    plan()   { return localStorage.getItem(PLAN_KEY) || "free"; },
    email()  { return localStorage.getItem(EMAIL_KEY) || ""; },
    active() { return !!localStorage.getItem(TOKEN_KEY); },
  };

  // ── Plan helpers (client UI only — server enforces truth) ────
  const plans = {
    canWhatsApp : () => ["standard","premium"].includes(session.plan()),
    canDOC      : () => ["standard","premium"].includes(session.plan()),
    canPDF      : () => session.plan() === "premium",
    isPremium   : () => session.plan() === "premium",
    isStandard  : () => session.plan() === "standard",
    label       : () => ({ free:"Free", standard:"Standard", premium:"Premium" })[session.plan()] || "Free",
  };

  // ── Auth ────────────────────────────────────────────────────
  const auth = {
    register : (d)          => call("register",  d, ""),
    login    : (email, pw)  => call("login",     { email, password: pw }, ""),
    logout   : ()           => call("logout",    {}, session.token()),
    profile  : ()           => call("getProfile",     {}),
    updateProfile : (d)     => call("updateProfile",  d),
    changePassword: (o, n)  => call("changePassword", { oldPassword:o, newPassword:n }),
  };

  // ── Transactions ─────────────────────────────────────────────
  const tx = {
    list   : (opts={})      => call("getTransactions",  opts),
    add    : (d)            => call("addTransaction",   d),
    update : (rowIdx, d)    => call("updateTransaction",{ ...d, rowIdx }),
    del    : (rowIdx)       => call("deleteTransaction",{ rowIdx }),
    dashboard: ()           => call("getDashboard",     {}),
    search : (query)        => call("globalSearch",     { query }),
  };

  // ── Estimates / Bills ─────────────────────────────────────────
  const est = {
    save   : (d)            => call("saveEstimate",         d),
    get    : (billId)       => call("getEstimate",          { billId }),
    list   : (opts={})      => call("listEstimates",        opts),
    status : (billId, s)    => call("updateEstimateStatus", { billId, status:s }),
    del    : (billId)       => call("deleteEstimate",       { billId }),
    whatsapp: (billId)      => call("buildWhatsApp",        { billId }),
  };

  // ── Export ───────────────────────────────────────────────────
  const exp = {
    pdf : (billId) => call("generatePDF", { billId }),
    doc : (billId) => call("generateDOC", { billId }),
  };

  // ── Admin / Payments ─────────────────────────────────────────
  const admin = {
    payments   : ()                   => call("adminGetPayments",   {}),
    approve    : (txId, remarks)      => call("adminApprovePayment",{ txId, remarks }),
    reject     : (txId, remarks)      => call("adminRejectPayment", { txId, remarks }),
    users      : ()                   => call("adminGetUsers",      {}),
    setStatus  : (userId, status)     => call("adminSetUserStatus", { userId, status }),
    submitPay  : (d)                  => call("submitPayment",      d),
    payInfo    : ()                   => call("getPaymentDetails",  {}),
    plans      : ()                   => call("getPlans",           {}, ""),
  };

  return { call, session, plans, auth, tx, est, exp, admin };
})();
