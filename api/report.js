import ExcelJS from 'exceljs';
import { getAccessToken, fetchAllPayments, buildReportRows } from './_lib/zoho.js';

const PAYMENT_COLUMNS = ['Cash', 'Chq', 'CC', 'DT', 'MY FEES'];
const createSummaryTotals = () => ({ Cash: 0, Chq: 0, CC: 0, DT: 0, 'MY FEES': 0, 'Sales Receipts': 0 });

function paymentColumnForRow(row) {
  if (row.column) return row.column;
  if (row.cash !== null) return 'Cash';
  if (row.chq !== null) return 'Chq';
  if (row.cc !== null) return 'CC';
  if (row.dt !== null) return 'DT';
  if (row.myFees !== null) return 'MY FEES';
  return null;
}

function buildBankSummaries(rows, salesReceiptRows) {
  const createSummary = createSummaryTotals;
  const paymentSummary = new Map();
  const combinedSummary = new Map();
  const add = (map, bank, column, amount) => {
    const summary = map.get(bank) || createSummary();
    summary[column] += Number(amount) || 0;
    map.set(bank, summary);
  };

  rows.forEach((row) => {
    const column = paymentColumnForRow(row);
    if (column) {
      add(paymentSummary, row.bank || '-', column, row[column === 'Cash' ? 'cash' : column === 'Chq' ? 'chq' : column === 'CC' ? 'cc' : column === 'DT' ? 'dt' : 'myFees']);
      add(combinedSummary, row.bank || '-', column, row[column === 'Cash' ? 'cash' : column === 'Chq' ? 'chq' : column === 'CC' ? 'cc' : column === 'DT' ? 'dt' : 'myFees']);
    }
  });
  salesReceiptRows.forEach((row) => add(combinedSummary, row.depositTo || '-', 'Sales Receipts', row.total));
  return { paymentSummary, combinedSummary };
}

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

    // Debug mode for the Sales Receipts permission issue: add &debug=salesreceipts
    // to see the RAW response Zoho sends back for that endpoint, with the
    // access token's actual granted scopes for comparison.
    if (debug === 'salesreceipts') {
      const accessToken = await getAccessToken();
      const apiDomain = process.env.ZOHO_API_DOMAIN || 'https://www.zohoapis.com';
      const orgId = process.env.ZOHO_ORGANIZATION_ID;

      const srUrl = `${apiDomain}/books/v3/salesreceipts?organization_id=${orgId}&date_start=${from}&date_end=${to}&per_page=1`;
      const srResp = await fetch(srUrl, { headers: { Authorization: `Zoho-oauthtoken ${accessToken}` } });
      const srData = await srResp.json();

      const accountsDomain = process.env.ZOHO_ACCOUNTS_DOMAIN || 'https://accounts.zoho.com';
      const tokenInfoResp = await fetch(`${accountsDomain}/oauth/v2/tokeninfo?access_token=${accessToken}`);
      const tokenInfo = await tokenInfoResp.json().catch(() => ({ note: 'tokeninfo endpoint unavailable' }));

      res.status(200).json({
        organization_id_used: orgId,
        sales_receipts_endpoint_status: srResp.status,
        sales_receipts_raw_response: srData,
        token_info: tokenInfo,
      });
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
      doc.x = doc.page.margins.left;
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
        doc.x = doc.page.margins.left;
      };
      drawRow(headers, true);
      data.forEach((row) => drawRow(row));
    };

    const { paymentSummary, combinedSummary } = buildBankSummaries(rows, salesReceiptRows);
    const summaryRows = (summaryMap, includeSales) => [...summaryMap.entries()].map(([bank, values]) => {
      const total = PAYMENT_COLUMNS.reduce((sum, column) => sum + values[column], 0) + (includeSales ? values['Sales Receipts'] : 0);
      return [bank, ...PAYMENT_COLUMNS.map((column) => formatAmount(values[column])), ...(includeSales ? [formatAmount(values['Sales Receipts'])] : []), formatAmount(total)];
    });
    const summaryHeaders = (includeSales) => ['Bank', ...PAYMENT_COLUMNS, ...(includeSales ? ['Sales Receipts'] : []), 'Total'];
    const summaryWidths = (includeSales) => includeSales ? [100, 60, 60, 60, 60, 70, 90, 70] : [120, 75, 75, 75, 75, 90, 85];
    const summaryTotal = (summaryMap, includeSales) => {
      const values = [...summaryMap.values()].reduce((result, value) => {
        [...PAYMENT_COLUMNS, ...(includeSales ? ['Sales Receipts'] : [])].forEach((column) => { result[column] += value[column]; });
        return result;
      }, createSummaryTotals(includeSales));
      const total = [...PAYMENT_COLUMNS, ...(includeSales ? ['Sales Receipts'] : [])].reduce((sum, column) => sum + values[column], 0);
      return [['TOTAL', ...PAYMENT_COLUMNS.map((column) => formatAmount(values[column])), ...(includeSales ? [formatAmount(values['Sales Receipts'])] : []), formatAmount(total)]];
    };
    const paymentTotal = PAYMENT_COLUMNS.reduce((sum, column) => sum + Number(totals[column] || 0), 0);

    drawHeader('DAILY CASH COLLECTION REPORT');
    doc.fontSize(11).font('Helvetica-Bold').text('1. PAYMENT RECEIPTS');
    drawTable(
      ['Date', 'Receipt No', 'Admission No.', 'Student Name', 'Details', 'Cash', 'Chq', 'CC', 'DT', 'MY FEES', 'Bank'],
      rows.map((row) => [row.date, row.receiptNo, row.admissionNumber || '*', row.studentName || '*', row.details || '*', formatAmount(row.cash), formatAmount(row.chq), formatAmount(row.cc), formatAmount(row.dt), formatAmount(row.myFees), row.bank]),
      [55, 65, 62, 95, 120, 48, 48, 48, 48, 55, 58]
    );
    doc.moveDown(1);
    doc.fontSize(11).font('Helvetica-Bold').text('2. PAYMENT RECEIPTS BANK-WISE SUMMARY');
    drawTable(
      summaryHeaders(false), summaryRows(paymentSummary, false).concat(summaryTotal(paymentSummary, false)), summaryWidths(false)
    );

    doc.addPage();
    drawHeader('3. SALES RECEIPTS');
    drawTable(
      ['Receipt Number', 'Receipt Date', 'Payment Mode', 'Customer Name', 'Admission No.', 'Deposit To', 'SubTotal', 'Total', 'Notes'],
      salesReceiptRows.map((row) => [row.receiptNumber, row.date, row.paymentMode, row.customerName, row.admissionNumber || '*', row.depositTo, formatAmount(row.subTotal), formatAmount(row.total), row.notes]),
      [75, 62, 70, 105, 65, 75, 60, 60, 160]
    );
    doc.font('Helvetica-Bold').fontSize(8).text(`TOTAL SALES RECEIPTS: ${salesReceiptRows.length}    ${formatAmount(salesReceiptTotal)}`);
    doc.moveDown(1);
    doc.fontSize(11).font('Helvetica-Bold').text('4. COMBINED PAYMENT AND SALES RECEIPTS SUMMARY');
    drawTable(
      ['Description', 'Amount'],
      [['Total of Payment Receipts', formatAmount(paymentTotal)], ['Total of Sales Receipts', formatAmount(salesReceiptTotal)], ['Total Daily Collection', formatAmount(paymentTotal + Number(salesReceiptTotal || 0))]],
      [pageWidth * 0.58, pageWidth * 0.22]
    );
    doc.moveDown(1.5);
    doc.fontSize(11).font('Helvetica-Bold').text('5. APPROVAL');
    doc.moveDown(1.5);
    const signatureWidth = pageWidth / 3 - 18;
    const signatureLabelY = doc.y;
    const signatureLineY = signatureLabelY + 18;
    ['REPORT GENERATED BY', 'REPORT CHECKED BY', 'AUTHORIZED BY'].forEach((label, index) => {
      const x = doc.page.margins.left + index * (pageWidth / 3);
      doc.fontSize(9).font('Helvetica-Bold').text(label, x, signatureLabelY, { width: signatureWidth, align: 'center' });
      doc.moveTo(x + 8, signatureLineY).dash(2, { space: 2 }).lineTo(x + signatureWidth - 8, signatureLineY).stroke('#555').undash();
    });
    doc.y = signatureLineY + 10;
    doc.end();
  });
}

async function buildWorkbook(rows, totals, salesReceiptRows, salesReceiptTotal, from, to) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Daily Collection');
  const { paymentSummary, combinedSummary } = buildBankSummaries(rows, salesReceiptRows);
  const beige = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9CBA0' } };
  const border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
  const money = (value) => Number(value || 0);
  const styleTitle = (row, title, width) => {
    sheet.mergeCells(row, 1, row, width);
    const cell = sheet.getCell(row, 1);
    cell.value = title;
    cell.font = { bold: true, size: 12 };
  };
  const addTable = (startRow, headers, data, widths) => {
    const header = sheet.getRow(startRow);
    headers.forEach((title, index) => {
      const cell = header.getCell(index + 1);
      cell.value = title;
      cell.font = { bold: true };
      cell.fill = beige;
      cell.border = border;
    });
    data.forEach((values, rowOffset) => {
      const row = sheet.getRow(startRow + rowOffset + 1);
      values.forEach((value, index) => {
        const cell = row.getCell(index + 1);
        cell.value = value;
        cell.border = border;
        if (index > 0 && typeof value === 'number') cell.numFmt = '#,##0';
      });
    });
    widths.forEach((width, index) => { sheet.getColumn(index + 1).width = width; });
    return startRow + data.length + 1;
  };
  const summaryData = (summaryMap, includeSales) => [...summaryMap.entries()].map(([bank, values]) => {
    const total = PAYMENT_COLUMNS.reduce((sum, column) => sum + values[column], 0) + (includeSales ? values['Sales Receipts'] : 0);
    return [bank, ...PAYMENT_COLUMNS.map((column) => money(values[column])), ...(includeSales ? [money(values['Sales Receipts'])] : []), money(total)];
  });
  const summaryTotal = (summaryMap, includeSales) => {
    const values = [...summaryMap.values()].reduce((result, value) => {
      [...PAYMENT_COLUMNS, ...(includeSales ? ['Sales Receipts'] : [])].forEach((column) => { result[column] += value[column]; });
      return result;
    }, createSummaryTotals());
    return [['TOTAL', ...PAYMENT_COLUMNS.map((column) => money(values[column])), ...(includeSales ? [money(values['Sales Receipts'])] : []), [...PAYMENT_COLUMNS, ...(includeSales ? ['Sales Receipts'] : [])].reduce((sum, column) => sum + values[column], 0)]];
  };
  const paymentTotal = PAYMENT_COLUMNS.reduce((sum, column) => sum + money(totals[column]), 0);

  styleTitle(1, 'WYCHERLEY INTERNATIONAL SCHOOL (PVT) LTD', 13);
  sheet.getCell('A2').value = 'NO. 232, BAUDDHALOKA MAWATHA, COLOMBO 07';
  styleTitle(3, `DAILY CASH COLLECTION REPORT - ${from} to ${to}`, 13);
  let rowIndex = 5;
  styleTitle(rowIndex++, '1. PAYMENT RECEIPTS', 13);
  rowIndex = addTable(rowIndex, ['Date & Time', 'Receipt No', 'Cashier', 'Admission Number', 'Student Name', 'Remark', 'Details', 'Cash', 'Chq', 'CC', 'DT', 'MY FEES', 'Bank'], rows.map((r) => [r.date, r.receiptNo, '', r.admissionNumber || '*', r.studentName || '*', r.reference || '*', r.details || '*', r.cash ?? '*', r.chq ?? '*', r.cc ?? '*', r.dt ?? '*', r.myFees ?? '*', r.bank]), [18, 14, 10, 16, 22, 14, 30, 12, 12, 12, 12, 12, 18]);
  styleTitle(rowIndex + 1, '2. PAYMENT RECEIPTS BANK-WISE SUMMARY', 7);
  rowIndex += 2;
  rowIndex = addTable(rowIndex, ['Bank', ...PAYMENT_COLUMNS, 'Total'], summaryData(paymentSummary, false).concat(summaryTotal(paymentSummary, false)), [28, 14, 14, 14, 14, 16, 16]);
  styleTitle(rowIndex + 2, '3. SALES RECEIPTS', 9);
  rowIndex += 3;
  rowIndex = addTable(rowIndex, ['Receipt Number', 'Receipt Date', 'Payment Mode', 'Customer Name', 'Admission Number', 'Deposit To', 'SubTotal', 'Total', 'Notes'], salesReceiptRows.map((sr) => [sr.receiptNumber, sr.date, sr.paymentMode, sr.customerName, sr.admissionNumber || '*', sr.depositTo, money(sr.subTotal), money(sr.total), sr.notes]), [20, 16, 16, 24, 18, 22, 14, 14, 30]);
  styleTitle(rowIndex + 1, '4. COMBINED PAYMENT AND SALES RECEIPTS SUMMARY', 8);
  rowIndex += 2;
  rowIndex = addTable(rowIndex, ['Description', 'Amount'], [
    ['Total of Payment Receipts', paymentTotal],
    ['Total of Sales Receipts', money(salesReceiptTotal)],
    ['Total Daily Collection', paymentTotal + money(salesReceiptTotal)],
  ], [38, 18]);
  styleTitle(rowIndex + 2, '5. APPROVAL', 13);
  rowIndex += 3;
  sheet.mergeCells(rowIndex, 1, rowIndex, 4); sheet.getCell(rowIndex, 1).value = 'REPORT GENERATED BY';
  sheet.mergeCells(rowIndex, 6, rowIndex, 9); sheet.getCell(rowIndex, 6).value = 'REPORT CHECKED BY';
  sheet.mergeCells(rowIndex, 11, rowIndex, 13); sheet.getCell(rowIndex, 11).value = 'AUTHORIZED BY';
  [1, 6, 11].forEach((column) => { sheet.getCell(rowIndex, column).font = { bold: true }; sheet.getCell(rowIndex, column).alignment = { horizontal: 'center' }; });
  const signatureLine = { style: 'dotted', color: { argb: 'FF555555' } };
  [[1, 4], [6, 9], [11, 13]].forEach(([start, end]) => {
    for (let column = start; column <= end; column++) {
      sheet.getCell(rowIndex + 2, column).border = { bottom: signatureLine };
    }
  });
  sheet.getRow(rowIndex + 2).height = 28;
  return workbook.xlsx.writeBuffer();
}

async function buildWorkbookLegacy(rows, totals, salesReceiptRows, salesReceiptTotal, from, to) {
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
