import ExcelJS from 'exceljs';
import { getAccessToken, fetchAllPayments, buildReportRows } from './_lib/zoho.js';

export default async function handler(req, res) {
  const { key, from, to, debug, format = 'xlsx' } = req.query;

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

    const { rows, totals, salesReceiptRows, salesReceiptTotal } = await buildReportRows(from, to);

    if (format === 'pdf') {
      const buffer = await buildPdf(rows, totals, salesReceiptRows, salesReceiptTotal, from, to);
      res.setHeader('Content-Type', 'application/pdf');
      res.setHeader('Content-Disposition', `attachment; filename="Daily_Collection_${from}_to_${to}.pdf"`);
      res.status(200).send(buffer);
      return;
    }

    const buffer = await buildWorkbook(rows, totals, salesReceiptRows, salesReceiptTotal, from, to);

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

async function buildPdf(rows, totals, salesReceiptRows, salesReceiptTotal, from, to) {
  const { default: PDFDocument } = await import('pdfkit');
  return new Promise((resolve, reject) => {
    const doc = new PDFDocument({ layout: 'landscape', size: 'A4', margin: 28 });
    const chunks = [];
    doc.on('data', (chunk) => chunks.push(chunk));
    doc.on('end', () => resolve(Buffer.concat(chunks)));
    doc.on('error', reject);

    const pageWidth = doc.page.width - doc.page.margins.left - doc.page.margins.right;
    const formatAmount = (amount) => Number(amount || 0).toLocaleString();
    const drawHeader = (title) => {
      doc.fontSize(14).font('Helvetica-Bold').text('WYCHERLEY INTERNATIONAL SCHOOL (PVT) LTD');
      doc.fontSize(9).font('Helvetica').text('NO. 232, BAUDDHALOKA MAWATHA, COLOMBO 07');
      doc.fontSize(11).font('Helvetica-Bold').text(`${title} - ${from} to ${to}`);
      doc.moveDown(0.8);
    };
    const drawTable = (headers, data, widths) => {
      const rowHeight = 18;
      const drawRow = (values, header = false) => {
        if (doc.y + rowHeight > doc.page.height - doc.page.margins.bottom) doc.addPage();
        const y = doc.y;
        let x = doc.page.margins.left;
        doc.font(header ? 'Helvetica-Bold' : 'Helvetica').fontSize(7);
        values.forEach((value, index) => {
          const width = widths[index];
          doc.rect(x, y, width, rowHeight).stroke('#999');
          doc.text(String(value ?? ''), x + 3, y + 5, { width: width - 6, height: rowHeight - 4, ellipsis: true });
          x += width;
        });
        doc.y = y + rowHeight;
      };
      drawRow(headers, true);
      data.forEach((row) => drawRow(row));
    };

    drawHeader('DAILY CASH COLLECTION REPORT');
    doc.fontSize(11).font('Helvetica-Bold').text('PAYMENT SUMMARY');
    const modes = ['Cash', 'Chq', 'CC', 'DT', 'MY FEES'];
    const paymentCounts = Object.fromEntries(modes.map((mode) => [mode, 0]));
    rows.forEach((row) => {
      const mode = row.cash !== null ? 'Cash' : row.chq !== null ? 'Chq' : row.cc !== null ? 'CC' : row.dt !== null ? 'DT' : row.myFees !== null ? 'MY FEES' : null;
      if (mode) paymentCounts[mode]++;
    });
    drawTable(
      ['Payment Mode', 'Receipts', 'Amount'],
      modes.map((mode) => [mode, paymentCounts[mode], formatAmount(totals[mode])])
        .concat([['TOTAL', rows.length, formatAmount(Object.values(totals).reduce((sum, amount) => sum + Number(amount || 0), 0))]]),
      [pageWidth * 0.35, pageWidth * 0.2, pageWidth * 0.25]
    );
    doc.moveDown(1);

    doc.fontSize(11).font('Helvetica-Bold').text('PAYMENTS RECEIVED');
    drawTable(
      ['Date', 'Receipt No', 'Admission No.', 'Student Name', 'Details', 'Cash', 'Chq', 'CC', 'DT', 'MY FEES', 'Bank'],
      rows.map((row) => [row.date, row.receiptNo, row.admissionNumber || '*', row.studentName || '*', row.details || '*', formatAmount(row.cash), formatAmount(row.chq), formatAmount(row.cc), formatAmount(row.dt), formatAmount(row.myFees), row.bank]),
      [55, 65, 62, 95, 120, 48, 48, 48, 48, 55, 58]
    );

    doc.addPage();
    drawHeader('SALES RECEIPTS');
    drawTable(
      ['Receipt Number', 'Receipt Date', 'Payment Mode', 'Customer Name', 'Admission No.', 'Deposit To', 'SubTotal', 'Total', 'Notes'],
      salesReceiptRows.map((row) => [row.receiptNumber, row.date, row.paymentMode, row.customerName, row.admissionNumber || '*', row.depositTo, formatAmount(row.subTotal), formatAmount(row.total), row.notes]),
      [75, 62, 70, 105, 65, 75, 60, 60, 160]
    );
    doc.font('Helvetica-Bold').fontSize(8).text(`TOTAL SALES RECEIPTS: ${salesReceiptRows.length}    ${formatAmount(salesReceiptTotal)}`);
    doc.end();
  });
}

async function buildWorkbook(rows, totals, salesReceiptRows, salesReceiptTotal, from, to) {
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

  // --- Second, separate table: Sales Receipts -------------------------------
  rowIndex += 3; // leave a couple of blank rows between the two tables

  sheet.mergeCells(`A${rowIndex}:I${rowIndex}`);
  sheet.getCell(`A${rowIndex}`).value = 'SALES RECEIPTS';
  sheet.getCell(`A${rowIndex}`).font = { bold: true, size: 13 };
  rowIndex += 2;

  const srColumns = [
    'Receipt Number', 'Receipt Date', 'Payment Mode', 'Customer Name',
    'Admission Number', 'Deposit To', 'SubTotal', 'Total', 'Notes',
  ];
  const srHeaderRowIndex = rowIndex;
  const srHeaderRow = sheet.getRow(srHeaderRowIndex);
  srColumns.forEach((title, i) => {
    const cell = srHeaderRow.getCell(i + 1);
    cell.value = title;
    cell.font = { bold: true };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9CBA0' } };
    cell.border = {
      top: { style: 'thin' }, left: { style: 'thin' },
      bottom: { style: 'thin' }, right: { style: 'thin' },
    };
  });
  rowIndex++;

  for (const sr of salesReceiptRows) {
    const row = sheet.getRow(rowIndex);
    row.getCell(1).value = sr.receiptNumber;
    row.getCell(2).value = sr.date;
    row.getCell(3).value = sr.paymentMode;
    row.getCell(4).value = sr.customerName;
    row.getCell(5).value = sr.admissionNumber || '*';
    row.getCell(6).value = sr.depositTo;
    row.getCell(7).value = sr.subTotal;
    row.getCell(8).value = sr.total;
    row.getCell(9).value = sr.notes;

    for (let c = 1; c <= 9; c++) {
      row.getCell(c).border = {
        top: { style: 'thin' }, left: { style: 'thin' },
        bottom: { style: 'thin' }, right: { style: 'thin' },
      };
    }
    rowIndex++;
  }

  const srTotalRow = sheet.getRow(rowIndex);
  srTotalRow.getCell(6).value = 'TOTAL';
  srTotalRow.getCell(6).font = { bold: true };
  srTotalRow.getCell(8).value = salesReceiptTotal || 0;
  srTotalRow.getCell(8).font = { bold: true };

  sheet.columns = [
    { width: 18 }, { width: 12 }, { width: 10 }, { width: 16 }, { width: 22 },
    { width: 14 }, { width: 30 }, { width: 12 }, { width: 12 }, { width: 12 },
    { width: 12 }, { width: 12 }, { width: 12 },
  ];

  return workbook.xlsx.writeBuffer();
}
