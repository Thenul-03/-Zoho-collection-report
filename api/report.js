import ExcelJS from 'exceljs';

// ---- EDIT THESE TWO MAPS TO MATCH YOUR ORGANIZATION -----------------------

// Zoho's payment_mode value -> which template column it should land in.
// Check your own Zoho Books org's exact mode names under
// Settings > Sales > Customer Payments (or just log `payment_mode` once and see).
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

// Zoho "deposit to" account name -> short label shown in the Bank column.
// Add every account your school actually deposits into.
const ACCOUNT_TO_BANK_LABEL = {
  'Cash in Hand': '-',
  'Seylan Bank C/A - 30013197592001': 'SEYLAN',
  'Online Payment Clearing / MyFees Control': '-',
};

// ---------------------------------------------------------------------------

export default async function handler(req, res) {
  const { key, from, to } = req.query;

  if (!process.env.REPORT_ACCESS_KEY || key !== process.env.REPORT_ACCESS_KEY) {
    res.status(401).send('Unauthorized. Pass the correct ?key= value.');
    return;
  }
  if (!from || !to) {
    res.status(400).send('Missing date range. Use ?from=YYYY-MM-DD&to=YYYY-MM-DD');
    return;
  }

  try {
    const accessToken = await getAccessToken();
    const payments = await fetchAllPayments(accessToken, from, to);
    const admissionNumbers = await fetchAdmissionNumbers(accessToken, payments);
    const buffer = await buildWorkbook(payments, admissionNumbers, from, to);

    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );
    res.setHeader(
      'Content-Disposition',
      `attachment; filename="Daily_Collection_${from}_to_${to}.xlsx"`
    );
    res.status(200).send(buffer);
  } catch (err) {
    console.error(err);
    res.status(500).send('Error generating report: ' + err.message);
  }
}

// --- Zoho auth -------------------------------------------------------------

async function getAccessToken() {
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

async function fetchAllPayments(accessToken, from, to) {
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
async function fetchAdmissionNumbers(accessToken, payments) {
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

    const cf = (data.contact.custom_fields || []).find(
      (f) => (f.label || f.customfield_name || '').toLowerCase() === 'admission number'
    );
    result[customerId] = cf ? cf.value : '';
  }

  return result;
}

// --- Helpers to map Zoho values onto the template's columns ----------------

function columnForPaymentMode(mode) {
  if (!mode) return null;
  const key = mode.toLowerCase().trim();
  return PAYMENT_MODE_TO_COLUMN[key] || null;
}

function bankLabelForAccount(accountName) {
  if (!accountName) return '-';
  return ACCOUNT_TO_BANK_LABEL[accountName] || accountName;
}

// --- Build the workbook ------------------------------------------------------

async function buildWorkbook(payments, admissionNumbers, from, to) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Daily Collection');

  const columns = [
    'Date & Time',
    'Receipt No',
    'Cashier',
    'Admission Number',
    'Student Name',
    'Remark',
    'Details',
    'Cash',
    'Chq',
    'CC',
    'DT',
    'MY FEES',
    'Bank',
  ];

  // Letterhead
  sheet.mergeCells('A1:M1');
  sheet.getCell('A1').value = 'WYCHERLEY INTERNATIONAL SCHOOL (PVT) LTD';
  sheet.getCell('A1').font = { bold: true, size: 14 };

  sheet.mergeCells('A2:M2');
  sheet.getCell('A2').value = 'NO. 232, BAUDDHALOKA MAWATHA, COLOMBO 07';

  sheet.mergeCells('A3:M3');
  sheet.getCell('A3').value = `DAILY CASH COLLECTION REPORT - ${from} to ${to}`;
  sheet.getCell('A3').font = { bold: true, color: { argb: 'FFCC0000' } };

  const headerRowIndex = 5;
  const headerRow = sheet.getRow(headerRowIndex);
  columns.forEach((title, i) => {
    const cell = headerRow.getCell(i + 1);
    cell.value = title;
    cell.font = { bold: true };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFD9CBA0' },
    };
    cell.border = {
      top: { style: 'thin' }, left: { style: 'thin' },
      bottom: { style: 'thin' }, right: { style: 'thin' },
    };
  });

  let rowIndex = headerRowIndex + 1;
  const totals = { Cash: 0, Chq: 0, CC: 0, DT: 0, 'MY FEES': 0 };

  for (const payment of payments) {
    const row = sheet.getRow(rowIndex);
    const column = columnForPaymentMode(payment.payment_mode);
    const amount = Number(payment.amount) || 0;

    row.getCell(1).value = payment.date || '';
    row.getCell(2).value = payment.payment_number || payment.reference_number || '';
    row.getCell(3).value = payment.created_by || '';
    row.getCell(4).value = admissionNumbers[payment.customer_id] || '*';
    row.getCell(5).value = payment.customer_name || '*';
    row.getCell(6).value = payment.reference_number || '*';
    row.getCell(7).value = payment.description || payment.notes || '*';
    row.getCell(8).value = column === 'Cash' ? amount : '*';
    row.getCell(9).value = column === 'Chq' ? amount : '*';
    row.getCell(10).value = column === 'CC' ? amount : '*';
    row.getCell(11).value = column === 'DT' ? amount : '*';
    row.getCell(12).value = column === 'MY FEES' ? amount : '*';
    row.getCell(13).value = bankLabelForAccount(payment.account_name);

    if (column) totals[column] += amount;

    for (let c = 1; c <= 13; c++) {
      row.getCell(c).border = {
        top: { style: 'thin' }, left: { style: 'thin' },
        bottom: { style: 'thin' }, right: { style: 'thin' },
      };
    }

    rowIndex++;
  }

  // Totals row
  const totalRow = sheet.getRow(rowIndex);
  totalRow.getCell(7).value = 'TOTAL';
  totalRow.getCell(7).font = { bold: true };
  totalRow.getCell(8).value = totals.Cash || '*****';
  totalRow.getCell(9).value = totals.Chq || '*****';
  totalRow.getCell(10).value = totals.CC || '*****';
  totalRow.getCell(11).value = totals.DT || '*****';
  totalRow.getCell(12).value = totals['MY FEES'] || '*****';
  for (let c = 7; c <= 12; c++) {
    totalRow.getCell(c).font = { bold: true };
  }

  // Reasonable column widths
  sheet.columns = [
    { width: 18 }, { width: 12 }, { width: 10 }, { width: 16 }, { width: 22 },
    { width: 14 }, { width: 30 }, { width: 12 }, { width: 12 }, { width: 12 },
    { width: 12 }, { width: 12 }, { width: 12 },
  ];

  return workbook.xlsx.writeBuffer();
}
