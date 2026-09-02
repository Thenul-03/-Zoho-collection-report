// JSON preview endpoint for the browser interface (public/index.html).
// This is the same underlying Zoho Books connection used by /api/report.js
// (and the same OAuth/refresh-token pattern as the cheque-writer tool) —
// it just returns rows as JSON instead of building an .xlsx file, so the
// page can show a live "Zoho Connected" table before anyone downloads
// anything.

import { getAccessToken, fetchAllPayments, fetchAdmissionNumbers, buildReportRows } from '../lib/zoho.js';

export default async function handler(req, res) {
  const { key, from, to } = req.query;

  if (!process.env.REPORT_ACCESS_KEY || key !== process.env.REPORT_ACCESS_KEY) {
    res.status(401).json({ connected: false, error: 'Unauthorized. Pass the correct ?key= value.' });
    return;
  }
  if (!from || !to) {
    res.status(400).json({ connected: false, error: 'Missing date range. Use ?from=YYYY-MM-DD&to=YYYY-MM-DD' });
    return;
  }

  let accessToken;
  try {
    accessToken = await getAccessToken();
  } catch (err) {
    // Distinguish "Zoho isn't reachable / refresh token is bad" from other
    // errors so the UI can show the red "Zoho Disconnected" badge, matching
    // the cheque-writer's connection indicator.
    res.status(200).json({ connected: false, error: 'Could not reach Zoho: ' + err.message });
    return;
  }

  try {
    const payments = await fetchAllPayments(accessToken, from, to);
    const admissionNumbers = await fetchAdmissionNumbers(accessToken, payments);
    const { rows, totals } = buildReportRows(payments, admissionNumbers);

    res.status(200).json({ connected: true, from, to, rows, totals, count: rows.length });
  } catch (err) {
    console.error(err);
    res.status(200).json({ connected: true, error: 'Fetched Zoho token but the report query failed: ' + err.message });
  }
}
