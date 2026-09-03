// Shared Zoho Books access logic. Nothing here is a route itself
// (the _lib folder is excluded from Vercel's automatic API routing).

// ---- EDIT THESE TWO MAPS TO MATCH YOUR ORGANIZATION -----------------------

const PAYMENT_MODE_TO_COLUMN = {
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

const ACCOUNT_TO_BANK_LABEL = {
  'Cash in Hand': '-',
  'Seylan Bank C/A - 30013197592001': 'SEYLAN',
  'Online Payment Clearing / MyFees Control': '-',
};

// If the debug view (?debug=1 on /api/report) showed a different label for
// your Admission Number custom field, update this to match exactly.
const ADMISSION_FIELD_LABEL = 'admission number';

// ---------------------------------------------------------------------------

export function columnForPaymentMode(mode) {
  if (!mode) return null;
  const key = mode.toLowerCase().trim();
  return PAYMENT_MODE_TO_COLUMN[key] || null;
}

export function bankLabelForAccount(accountName) {
  if (!accountName) return '-';
  return ACCOUNT_TO_BANK_LABEL[accountName] || accountName;
}

function normalizeLabel(s) {
  return (s || '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

export function extractAdmissionNumber(contact) {
  const target = normalizeLabel(ADMISSION_FIELD_LABEL);

  if (Array.isArray(contact.custom_fields)) {
    const match = contact.custom_fields.find((f) => {
      const label = f.label || f.customfield_name || f.field_name || f.placeholder || '';
      return normalizeLabel(label) === target;
    });
    if (match && match.value) return match.value;
  }

  if (contact.custom_field_hash && typeof contact.custom_field_hash === 'object') {
    for (const [k, v] of Object.entries(contact.custom_field_hash)) {
      if (normalizeLabel(k) === `cf${target}` || normalizeLabel(k) === target) {
        if (v) return v;
      }
    }
  }

  return '';
}

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

export async function fetchAllSalesReceipts(accessToken, from, to) {
  const apiDomain = process.env.ZOHO_API_DOMAIN || 'https://www.zohoapis.com';
  const orgId = process.env.ZOHO_ORGANIZATION_ID;

  let page = 1;
  let all = [];

  while (true) {
    const url =
      `${apiDomain}/books/v3/salesreceipts?organization_id=${orgId}` +
      `&date_start=${from}&date_end=${to}&page=${page}&per_page=200&sort_column=date`;

    const resp = await fetch(url, {
      headers: { Authorization: `Zoho-oauthtoken ${accessToken}` },
    });
    const data = await resp.json();

    if (data.code !== 0) {
      throw new Error('Zoho API error fetching sales receipts: ' + JSON.stringify(data));
    }

    all = all.concat(data.salesreceipts || []);

    if (!data.page_context || !data.page_context.has_more_page) break;
    page++;
  }

  return all;
}

// Admission Number lives on the Customer/Contact record's custom fields,
// not on the payment/receipt itself — fetch each distinct customer once and
// cache. Accepts any list of records that have a .customer_id, so it works
// for both payments and sales receipts (and can be called with the combined
// set to fetch each customer only once across both tables).
export async function fetchAdmissionNumbers(accessToken, records) {
  const apiDomain = process.env.ZOHO_API_DOMAIN || 'https://www.zohoapis.com';
  const orgId = process.env.ZOHO_ORGANIZATION_ID;

  const uniqueCustomerIds = [...new Set(records.map((r) => r.customer_id).filter(Boolean))];
  const result = {};

  for (const customerId of uniqueCustomerIds) {
    const url = `${apiDomain}/books/v3/contacts/${customerId}?organization_id=${orgId}`;
    const resp = await fetch(url, {
      headers: { Authorization: `Zoho-oauthtoken ${accessToken}` },
    });
    const data = await resp.json();

    result[customerId] = data.code === 0 && data.contact ? extractAdmissionNumber(data.contact) : '';
  }

  return result;
}

// Builds the flat row objects used by both the on-screen preview and the
// Excel export, so the two never drift out of sync. Returns two separate
// tables — payments and sales receipts — as your daily collection combines
// both, but you asked for them kept apart rather than merged into one list.
export async function buildReportRows(from, to) {
  const accessToken = await getAccessToken();
  const [payments, salesReceipts] = await Promise.all([
    fetchAllPayments(accessToken, from, to),
    fetchAllSalesReceipts(accessToken, from, to),
  ]);

  // Look up Admission Number once per customer across BOTH tables combined,
  // so a student appearing in both isn't fetched twice.
  const admissionNumbers = await fetchAdmissionNumbers(accessToken, [...payments, ...salesReceipts]);

  const totals = { Cash: 0, Chq: 0, CC: 0, DT: 0, 'MY FEES': 0 };
  const rows = payments.map((payment) => {
    const column = columnForPaymentMode(payment.payment_mode);
    const amount = Number(payment.amount) || 0;
    if (column) totals[column] += amount;

    return {
      date: payment.date || '',
      receiptNo: payment.payment_number || payment.reference_number || '',
      admissionNumber: admissionNumbers[payment.customer_id] || '',
      studentName: payment.customer_name || '',
      reference: payment.reference_number || '',
      details: payment.description || payment.notes || '',
      cash: column === 'Cash' ? amount : null,
      chq: column === 'Chq' ? amount : null,
      cc: column === 'CC' ? amount : null,
      dt: column === 'DT' ? amount : null,
      myFees: column === 'MY FEES' ? amount : null,
      bank: bankLabelForAccount(payment.account_name),
    };
  });

  // Sales Receipt columns match your sample export, minus every column that
  // came back empty in it (Supply Date, Reference#, Terms & Conditions, etc).
  let salesReceiptTotal = 0;
  const salesReceiptRows = salesReceipts.map((sr) => {
    const total = Number(sr.total ?? sr.sub_total) || 0;
    salesReceiptTotal += total;

    return {
      receiptNumber: sr.salesreceipt_number || sr.receipt_number || sr.salesreceipt_id || '',
      date: sr.date || '',
      paymentMode: sr.payment_mode || '',
      customerName: sr.customer_name || '',
      admissionNumber: admissionNumbers[sr.customer_id] || '',
      depositTo: bankLabelForAccount(sr.account_name),
      subTotal: Number(sr.sub_total) || 0,
      total,
      notes: sr.notes || '',
      location: sr.location_name || '',
    };
  });

  return {
    rows,
    totals,
    count: payments.length,
    salesReceiptRows,
    salesReceiptTotal,
    salesReceiptCount: salesReceipts.length,
  };
}
