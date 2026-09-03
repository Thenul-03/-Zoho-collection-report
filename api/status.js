import { getAccessToken } from './_lib/zoho.js';

export default async function handler(req, res) {
  const { key } = req.query;
  if (!process.env.REPORT_ACCESS_KEY || key !== process.env.REPORT_ACCESS_KEY) {
    res.status(401).json({ connected: false, error: 'Unauthorized' });
    return;
  }

  try {
    await getAccessToken();
    res.status(200).json({ connected: true });
  } catch (err) {
    res.status(200).json({ connected: false, error: err.message });
  }
}
