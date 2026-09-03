// Starts a fresh Zoho OAuth consent flow with the scopes required by the report.
export default function handler(req, res) {
  const required = ['ZOHO_CLIENT_ID', 'ZOHO_ACCOUNTS_DOMAIN'];
  const missing = required.filter((name) => !process.env[name]);
  if (missing.length) {
    res.status(500).send(`Missing environment variable(s): ${missing.join(', ')}`);
    return;
  }

  const redirectUri = `https://${req.headers.host}/api/callback`;
  const domain = process.env.ZOHO_ACCOUNTS_DOMAIN;
  const params = new URLSearchParams({
    scope: 'ZohoBooks.customerpayments.READ,ZohoBooks.salesreceipts.READ,ZohoBooks.contacts.READ,ZohoBooks.settings.READ',
    client_id: process.env.ZOHO_CLIENT_ID,
    response_type: 'code',
    access_type: 'offline',
    redirect_uri: redirectUri,
    prompt: 'consent',
  });

  res.redirect(`${domain}/oauth/v2/auth?${params.toString()}`);
}