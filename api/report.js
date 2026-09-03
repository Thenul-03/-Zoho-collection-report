import ExcelJS from 'exceljs';
import { getAccessToken, fetchAllPayments, buildReportRows } from './_lib/zoho.js';

export default async function handler(req, res) {
  const { key, from, to, debug } = req.query;

  if (!process.env.REPORT_ACCESS_KEY || key !== process.env.REPORT_ACCESS_KEY) {
    res.status(401).send('Unauthorized. Pass the correct ?key= value.');
    return;
  }
  if (!from || !to) {
    res.status(400).send('Missing date range. Use ?from=YYYY-MM-DD&to=YYYY-MM-DD');
    return;
  }

  try {
    // Debug mode: add &debug=1 to see the RAW contact JSON for the first
    // payment in range, to find the exact Admission Number field name/label.
    if (debug === '1') {
      const accessToken = await getAccessToken();
      const payments = await fetchAllPayments(accessToken, from, to);
      const firstPayment = payments[0];
      if (!firstPayment) {
        res.status(200).json({ message: 'No payments found in this date range.' });
        return;
      }
      const apiDomain = process.env.ZOHO_API_DOMAIN || 'https://www.zohoapis.com';
      const orgId = process.env.ZOHO_ORGANIZATION_ID;
      const url = `${apiDomain}/books/v3/contacts/${firstPayment.customer_id}?organization_id=${orgId}`;
      const resp = await fetch(url, { headers: { Authorization: `Zoho-oauthtoken ${accessToken}` } });
      const contactData = await resp.json();
      res.status(200).json({ sample_payment: firstPayment, sample_contact: contactData });
      return;
    }

    const { rows, totals } = await buildReportRows(from, to);
    const buffer = await buildWorkbook(rows, totals, from, to);

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

async function buildWorkbook(rows, totals, from, to) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Daily Collection');

  const columns = [
    'Date & Time', 'Receipt No', 'Cashier', 'Admission Number', 'Student Name',
    'Remark', 'Details', 'Cash', 'Chq', 'CC', 'DT', 'MY FEES', 'Bank',
  ];

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
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9CBA0' } };
    cell.border = {
      top: { style: 'thin' }, left: { style: 'thin' },
      bottom: { style: 'thin' }, right: { style: 'thin' },
    };
  });

  let rowIndex = headerRowIndex + 1;
  for (const r of rows) {
    const row = sheet.getRow(rowIndex);
    row.getCell(1).value = r.date;
    row.getCell(2).value = r.receiptNo;
    row.getCell(3).value = '';
    row.getCell(4).value = r.admissionNumber || '*';
    row.getCell(5).value = r.studentName || '*';
    row.getCell(6).value = r.reference || '*';
    row.getCell(7).value = r.details || '*';
    row.getCell(8).value = r.cash ?? '*';
    row.getCell(9).value = r.chq ?? '*';
    row.getCell(10).value = r.cc ?? '*';
    row.getCell(11).value = r.dt ?? '*';
    row.getCell(12).value = r.myFees ?? '*';
    row.getCell(13).value = r.bank;

    for (let c = 1; c <= 13; c++) {
      row.getCell(c).border = {
        top: { style: 'thin' }, left: { style: 'thin' },
        bottom: { style: 'thin' }, right: { style: 'thin' },
      };
    }
    rowIndex++;
  }

  const totalRow = sheet.getRow(rowIndex);
  totalRow.getCell(7).value = 'TOTAL';
  totalRow.getCell(7).font = { bold: true };
  totalRow.getCell(8).value = totals.Cash || '*****';
  totalRow.getCell(9).value = totals.Chq || '*****';
  totalRow.getCell(10).value = totals.CC || '*****';
  totalRow.getCell(11).value = totals.DT || '*****';
  totalRow.getCell(12).value = totals['MY FEES'] || '*****';
  for (let c = 7; c <= 12; c++) totalRow.getCell(c).font = { bold: true };

  sheet.columns = [
    { width: 18 }, { width: 12 }, { width: 10 }, { width: 16 }, { width: 22 },
    { width: 14 }, { width: 30 }, { width: 12 }, { width: 12 }, { width: 12 },
    { width: 12 }, { width: 12 }, { width: 12 },
  ];

  return workbook.xlsx.writeBuffer();
}
