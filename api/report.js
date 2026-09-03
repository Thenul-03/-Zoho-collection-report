import ExcelJS from 'exceljs';
import { getAccessToken, fetchAllPayments, buildReportRows } from './_lib/zoho.js';

export default async function handler(req, res) {
  const { key, from, to, debug, format = 'excel' } = req.query;

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

    const { rows, totals, salesReceiptRows, salesReceiptTotals } = await buildReportRows(from, to);
    if (format === 'pdf') {
      const buffer = await buildPdf(rows, totals, salesReceiptRows, salesReceiptTotals, from, to);
      res.setHeader('Content-Type', 'application/pdf');
      res.setHeader('Content-Disposition', `attachment; filename="Daily_Collection_${from}_to_${to}.pdf"`);
      res.status(200).send(buffer);
      return;
    }

    const buffer = await buildWorkbook(rows, totals, salesReceiptRows, salesReceiptTotals, from, to);

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

async function buildWorkbook(rows, totals, salesReceiptRows, salesReceiptTotals, from, to) {
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
    row.getCell(4).value = r.admissionNumber || '';
    row.getCell(5).value = r.studentName || '';
    row.getCell(6).value = r.reference || '';
    row.getCell(7).value = r.details || '';
    row.getCell(8).value = r.cash ?? 0;
    row.getCell(9).value = r.chq ?? 0;
    row.getCell(10).value = r.cc ?? 0;
    row.getCell(11).value = r.dt ?? 0;
    row.getCell(12).value = r.myFees ?? 0;
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
  totalRow.getCell(8).value = totals.Cash;
  totalRow.getCell(9).value = totals.Chq;
  totalRow.getCell(10).value = totals.CC;
  totalRow.getCell(11).value = totals.DT;
  totalRow.getCell(12).value = totals['MY FEES'];
  for (let c = 7; c <= 12; c++) totalRow.getCell(c).font = { bold: true };

  rowIndex += 3;
  sheet.getCell(`A${rowIndex}`).value = 'SALES RECEIPTS';
  sheet.getCell(`A${rowIndex}`).font = { bold: true, size: 12 };
  const salesColumns = [
    'Receipt Number', 'Receipt Date', 'Payment Mode', 'Customer Name',
    'Admission Number', 'Deposit To', 'SubTotal', 'Total', 'Notes',
  ];
  salesColumns.forEach((title, index) => {
    const cell = sheet.getCell(rowIndex + 1, index + 1);
    cell.value = title;
    cell.font = { bold: true };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9CBA0' } };
    cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
  });
  salesReceiptRows.forEach((receipt, offset) => {
    const row = sheet.getRow(rowIndex + 2 + offset);
    const values = [receipt.receiptNumber, receipt.receiptDate, receipt.paymentMode, receipt.customerName,
      receipt.admissionNumber, receipt.depositTo, receipt.subTotal, receipt.total, receipt.notes];
    values.forEach((value, index) => {
      row.getCell(index + 1).value = value || (index >= 6 ? 0 : '');
      row.getCell(index + 1).border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
    });
  });
  const salesTotalRow = sheet.getRow(rowIndex + 2 + salesReceiptRows.length);
  salesTotalRow.getCell(6).value = 'TOTAL';
  salesTotalRow.getCell(6).font = { bold: true };
  salesTotalRow.getCell(7).value = salesReceiptTotals.SubTotal;
  salesTotalRow.getCell(8).value = salesReceiptTotals.Total;
  for (let c = 6; c <= 8; c++) salesTotalRow.getCell(c).font = { bold: true };

  addSummaryAndSignatures(sheet, rowIndex + 5 + salesReceiptRows.length, rows);

  sheet.columns = [
    { width: 18 }, { width: 16 }, { width: 14 }, { width: 24 }, { width: 18 },
    { width: 30 }, { width: 14 }, { width: 14 }, { width: 30 }, { width: 12 },
    { width: 12 }, { width: 12 }, { width: 12 },
  ];

  return workbook.xlsx.writeBuffer();
}

function getBankSummary(rows) {
  const summary = new Map();
  for (const row of rows) {
    const bank = row.bank || '-';
    const current = summary.get(bank) || { Cash: 0, Chq: 0, CC: 0, DT: 0, 'MY FEES': 0 };
    if (row.cash !== null) current.Cash += row.cash || 0;
    if (row.chq !== null) current.Chq += row.chq || 0;
    if (row.cc !== null) current.CC += row.cc || 0;
    if (row.dt !== null) current.DT += row.dt || 0;
    if (row.myFees !== null) current['MY FEES'] += row.myFees || 0;
    summary.set(bank, current);
  }
  return [...summary.entries()].map(([bank, values]) => ({ bank, values }));
}

function addSummaryAndSignatures(sheet, startRow, rows) {
  const headers = ['Bank', 'Cash', 'Chq', 'CC', 'DT', 'MY FEES', 'Total'];
  sheet.getCell(`A${startRow}`).value = 'BANK-WISE SUMMARY';
  sheet.getCell(`A${startRow}`).font = { bold: true, size: 12 };
  headers.forEach((header, index) => {
    const cell = sheet.getCell(startRow + 1, index + 1);
    cell.value = header;
    cell.font = { bold: true };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9CBA0' } };
  });

  getBankSummary(rows).forEach(({ bank, values }, offset) => {
    const row = sheet.getRow(startRow + 2 + offset);
    const amounts = [values.Cash, values.Chq, values.CC, values.DT, values['MY FEES']];
    row.getCell(1).value = bank;
    amounts.forEach((amount, index) => { row.getCell(index + 2).value = amount; });
    row.getCell(7).value = amounts.reduce((sum, amount) => sum + amount, 0);
  });

  const signatureRow = startRow + 4 + getBankSummary(rows).length;
  ['REPORT GENERATED BY', 'REPORT CHECKED BY', 'AUTHORIZED BY'].forEach((label, index) => {
    const column = 1 + index * 5;
    sheet.mergeCells(signatureRow, column, signatureRow, column + 2);
    const cell = sheet.getCell(signatureRow, column);
    cell.value = label;
    cell.font = { bold: true };
    cell.alignment = { horizontal: 'center' };
    cell.border = { top: { style: 'dotted' } };
  });
}

async function buildPdf(rows, totals, salesReceiptRows, salesReceiptTotals, from, to) {
  const { default: PDFDocument } = await import('pdfkit');
  const document = new PDFDocument({ layout: 'landscape', size: 'A4', margin: 24 });
  const chunks = [];
  document.on('data', (chunk) => chunks.push(chunk));
  const finished = new Promise((resolve) => document.on('end', resolve));

  document.fontSize(14).font('Helvetica-Bold').text('WYCHERLEY INTERNATIONAL SCHOOL (PVT) LTD', { align: 'center' });
  document.fontSize(10).font('Helvetica').text('NO. 232, BAUDDHALOKA MAWATHA, COLOMBO 07', { align: 'center' });
  document.fontSize(11).font('Helvetica-Bold').text(`DAILY CASH COLLECTION REPORT - ${from} to ${to}`, { align: 'center' });
  document.moveDown();

  const columns = ['Date', 'Receipt No', 'Admission No.', 'Student Name', 'Details', 'Cash', 'Chq', 'CC', 'DT', 'MY FEES', 'Bank'];
  const widths = [55, 72, 65, 92, 112, 42, 42, 42, 42, 52, 65];
  let y = document.y;
  const drawRow = (values, bold = false) => {
    let x = document.page.margins.left;
    document.font(bold ? 'Helvetica-Bold' : 'Helvetica').fontSize(7);
    values.forEach((value, index) => {
      document.rect(x, y, widths[index], 16).stroke();
      document.text(String(value ?? ''), x + 2, y + 4, { width: widths[index] - 4, height: 12, ellipsis: true });
      x += widths[index];
    });
    y += 16;
    if (y > 540) { document.addPage(); y = 30; }
  };
  drawRow(columns, true);
  rows.forEach((row) => drawRow([row.date, row.receiptNo, row.admissionNumber, row.studentName, row.details, row.cash ?? 0, row.chq ?? 0, row.cc ?? 0, row.dt ?? 0, row.myFees ?? 0, row.bank]));
  drawRow(['', '', '', '', 'TOTAL', totals.Cash, totals.Chq, totals.CC, totals.DT, totals['MY FEES'], ''], true);

  if (y > 390) { document.addPage(); y = 30; }
  y += 24;
  document.font('Helvetica-Bold').fontSize(10).text('SALES RECEIPTS', document.page.margins.left, y);
  y += 18;
  const salesWidths = [75, 55, 58, 100, 75, 90, 55, 55, 107];
  const salesHeaders = ['Receipt Number', 'Receipt Date', 'Payment Mode', 'Customer Name', 'Admission Number', 'Deposit To', 'SubTotal', 'Total', 'Notes'];
  const drawSalesRow = (values, bold = false) => {
    let x = document.page.margins.left;
    document.font(bold ? 'Helvetica-Bold' : 'Helvetica').fontSize(7);
    values.forEach((value, index) => {
      document.rect(x, y, salesWidths[index], 18).stroke();
      document.text(String(value ?? (index >= 6 ? 0 : '')), x + 2, y + 5, { width: salesWidths[index] - 4, height: 12, ellipsis: true });
      x += salesWidths[index];
    });
    y += 18;
  };
  drawSalesRow(salesHeaders, true);
  salesReceiptRows.forEach((receipt) => drawSalesRow([
    receipt.receiptNumber, receipt.receiptDate, receipt.paymentMode, receipt.customerName,
    receipt.admissionNumber, receipt.depositTo, receipt.subTotal, receipt.total, receipt.notes,
  ]));
  drawSalesRow(['', '', '', '', '', 'TOTAL', salesReceiptTotals.SubTotal, salesReceiptTotals.Total, ''], true);

  if (y > 430) { document.addPage(); y = 30; }
  y += 24;
  document.font('Helvetica-Bold').fontSize(10).text('BANK-WISE SUMMARY', document.page.margins.left, y);
  y += 18;
  const summaryWidths = [145, 55, 55, 55, 55, 70, 70];
  const summaryHeaders = ['Bank', 'Cash', 'Chq', 'CC', 'DT', 'MY FEES', 'Total'];
  const drawSummaryRow = (values, bold = false) => {
    let x = document.page.margins.left;
    document.font(bold ? 'Helvetica-Bold' : 'Helvetica').fontSize(8);
    values.forEach((value, index) => {
      document.rect(x, y, summaryWidths[index], 18).stroke();
      document.text(String(value ?? 0), x + 3, y + 5, { width: summaryWidths[index] - 6, height: 12, ellipsis: true });
      x += summaryWidths[index];
    });
    y += 18;
  };
  drawSummaryRow(summaryHeaders, true);
  getBankSummary(rows).forEach(({ bank, values }) => {
    const amounts = [values.Cash, values.Chq, values.CC, values.DT, values['MY FEES']];
    drawSummaryRow([bank, ...amounts, amounts.reduce((sum, amount) => sum + amount, 0)]);
  });

  y += 42;
  const signatureWidth = 180;
  const signatureGap = 54;
  const signatureLabels = ['REPORT GENERATED BY', 'REPORT CHECKED BY', 'AUTHORIZED BY'];
  signatureLabels.forEach((label, index) => {
    const x = document.page.margins.left + index * (signatureWidth + signatureGap);
    document.moveTo(x, y).lineTo(x + signatureWidth, y).stroke();
    document.font('Helvetica-Bold').fontSize(8).text(label, x, y + 8, { width: signatureWidth, align: 'center' });
  });
  document.end();
  await finished;
  return Buffer.concat(chunks);
}
