import { buildReportRows } from './_lib/zoho.js';

export default async function handler(req, res) {
  const { key, from, to } = req.query;

  if (!process.env.REPORT_ACCESS_KEY || key !== process.env.REPORT_ACCESS_KEY) {
    res.status(401).json({ error: 'Unauthorized' });
    return;
  }
  if (!from || !to) {
    res.status(400).json({ error: 'Missing from/to date params (YYYY-MM-DD)' });
    return;
  }

  try {
    const data = await buildReportRows(from, to);
    res.status(200).json(data);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
}
