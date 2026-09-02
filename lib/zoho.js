// Shared Zoho Books helpers used by both /api/data.js (JSON preview for the
// browser interface) and /api/report.js (the .xlsx export). Keeping this in
// /lib (not /api) means Vercel does NOT treat it as its own route.

// ---- EDIT THESE TWO MAPS TO MATCH YOUR ORGANIZATION -----------------------

// Zoho's payment_mode value -> which template column it should land in.
// Check your own Zoho Books org's exact mode names under
// Settings > Sales > Customer Payments (or hit /api/report?...&debug=1).
export const PAYMENT_MODE_TO_COLUMN = {
  cash: 'Cash',
  check: 'Chq',
  cheque: 'Chq',
  creditcard: 'CC',
  'credit card': 'CC',
  banktransfer: 'DT',
  'bank transfer': 'DT',
  myfees: 'MY FEES',
  'my fees': 'MY FEES',
};

// Zoho "deposit to" account name -> short label shown in the Bank column.
// Add every account your school actually deposits into.
export const ACCOUNT_TO_BANK_LABEL = {
  'Cash in Hand': '-',
  'Seylan Bank C/A - 30013197592001': 'SEYLAN',
  'Online Payment Clearing / MyFees Control': '-',
};

export const REPORT_COLUMNS = ['Cash', 'Chq', 'CC', 'DT', 'MY FEES'];

// ---------------------------------------------------------------------------

// --- Zoho auth -------------------------------------------------------------

export async function getAccessToken() {
  const domain = process.env.ZOHO_ACCOUNTS_DOMAIN || 'https://accounts.zoho.com';
  const params = new URLSearchParams({
    grant_type: 'refresh_token',
    client_id: process.env.ZOHO_CLIENT_ID,
    client_secret: process.env.ZOHO_CLIENT_SECRET,
    refresh_token: process.env.ZOHO_REFRESH_TOKEN,
  });

  const resp = await fetch(`${domain}/oauth/v2/token`, { method: 'POST', body: params });
  const data = await resp.json();
  if (!data.access_token) {
    throw new Error('Failed to refresh access token: ' + JSON.stringify(data));
  }
  return data.access_token;
}

// --- Zoho data fetching ------------------------------------------------------

export async function fetchAllPayments(accessToken, from, to) {
  const apiDomain = process.env.ZOHO_API_DOMAIN || 'https://www.zohoapis.com';
  const orgId = process.env.ZOHO_ORGANIZATION_ID;

  let page = 1;
  let all = [];

  while (true) {
    const url =
      `${apiDomain}/books/v3/customerpayments?organization_id=${orgId}` +
      `&date_start=${from}&date_end=${to}&page=${page}&per_page=200&sort_column=date`;

    const resp = await fetch(url, {
      headers: { Authorization: `Zoho-oauthtoken ${accessToken}` },
    });
    const data = await resp.json();

    if (data.code !== 0) {
      throw new Error('Zoho API error fetching payments: ' + JSON.stringify(data));
    }

    all = all.concat(data.customerpayments || []);

    if (!data.page_context || !data.page_context.has_more_page) break;
    page++;
  }

  return all;
}

// Admission Number lives on the Customer/Contact record's custom fields,
// not on the payment itself — so we fetch each distinct customer once and
// cache the result.
export async function fetchAdmissionNumbers(accessToken, payments) {
  const apiDomain = process.env.ZOHO_API_DOMAIN || 'https://www.zohoapis.com';
  const orgId = process.env.ZOHO_ORGANIZATION_ID;

  const uniqueCustomerIds = [...new Set(payments.map((p) => p.customer_id).filter(Boolean))];
  const result = {};

  for (const customerId of uniqueCustomerIds) {
    const url = `${apiDomain}/books/v3/contacts/${customerId}?organization_id=${orgId}`;
    const resp = await fetch(url, {
      headers: { Authorization: `Zoho-oauthtoken ${accessToken}` },
    });
    const data = await resp.json();

    if (data.code !== 0 || !data.contact) {
      result[customerId] = '';
      continue;
    }

    result[customerId] = extractAdmissionNumber(data.contact);
  }

  return result;
}

// Zoho has returned custom fields under slightly different shapes across API
// versions/response contexts. This checks every shape we've seen so a label
// spelling difference doesn't silently produce a blank column. If this still
// comes up empty, hit /api/report?...&debug=1 and check the exact label Zoho
// sends back for your org, then adjust ADMISSION_FIELD_LABEL below.
const ADMISSION_FIELD_LABEL = 'admission number';

function normalizeLabel(s) {
  return (s || '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

export function extractAdmissionNumber(contact) {
  const target = normalizeLabel(ADMISSION_FIELD_LABEL);

  // Shape 1: custom_fields: [{ label, value }] or [{ customfield_name, value }]
  if (Array.isArray(contact.custom_fields)) {
    const match = contact.custom_fields.find((f) => {
      const label = f.label || f.customfield_name || f.field_name || f.placeholder || '';
      return normalizeLabel(label) === target;
    });
    if (match && match.value) return match.value;
  }

  // Shape 2: custom_field_hash: { cf_admission_number: '...', cf_admission_number_unformatted: '...' }
  if (contact.custom_field_hash && typeof contact.custom_field_hash === 'object') {
    for (const [k, v] of Object.entries(contact.custom_field_hash)) {
      if (normalizeLabel(k) === `cf${target}` || normalizeLabel(k) === target) {
        if (v) return v;
      }
    }
  }

  return '';
}

// --- Helpers to map Zoho values onto the template's columns ----------------

export function columnForPaymentMode(mode) {
  if (!mode) return null;
  const key = mode.toLowerCase().trim();
  return PAYMENT_MODE_TO_COLUMN[key] || null;
}

export function bankLabelForAccount(accountName) {
  if (!accountName) return '-';
  return ACCOUNT_TO_BANK_LABEL[accountName] || accountName;
}

// --- Shared row-shaping used by both the JSON preview and the .xlsx export -

// Turns raw Zoho payments + the admission-number lookup into the flat row
// shape the report needs, plus running totals per column. Both /api/data.js
// (preview table in the browser) and /api/report.js (Excel export) call this
// so the two never drift apart.
export function buildReportRows(payments, admissionNumbers) {
  const totals = { Cash: 0, Chq: 0, CC: 0, DT: 0, 'MY FEES': 0 };

  const rows = payments.map((payment) => {
    const column = columnForPaymentMode(payment.payment_mode);
    const amount = Number(payment.amount) || 0;
    if (column) totals[column] += amount;

    return {
      date: payment.date || '',
      receiptNo: payment.payment_number || payment.reference_number || '',
      cashier: payment.created_by || '',
      admissionNumber: admissionNumbers[payment.customer_id] || '',
      studentName: payment.customer_name || '',
      remark: payment.reference_number || '',
      details: payment.description || payment.notes || '',
      paymentMode: payment.payment_mode || '',
      column,
      amount,
      cash: column === 'Cash' ? amount : null,
      chq: column === 'Chq' ? amount : null,
      cc: column === 'CC' ? amount : null,
      dt: column === 'DT' ? amount : null,
      myFees: column === 'MY FEES' ? amount : null,
      bank: bankLabelForAccount(payment.account_name),
      status: payment.status || '',
    };
  });

  return { rows, totals };
}
