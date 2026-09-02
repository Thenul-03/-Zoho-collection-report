// One-time setup endpoint.
// Zoho redirects here after you approve access, with a `code` query param.
// This exchanges that code for tokens and shows you the refresh_token ONCE,
// directly in your browser — nothing is stored or logged anywhere.
//
// After you copy the refresh_token into your Vercel Environment Variables
// (as ZOHO_REFRESH_TOKEN), you can leave this file in place; it's harmless
// to keep around, but you won't need to visit it again unless you
// re-authorize (e.g. after revoking access).

export default async function handler(req, res) {
  const { code } = req.query;

  if (!code) {
    res.status(400).send('Missing ?code= — this endpoint is only meant to be hit by Zoho\'s OAuth redirect.');
    return;
  }

  const accountsDomain = process.env.ZOHO_ACCOUNTS_DOMAIN || 'https://accounts.zoho.com';
  const redirectUri = `https://${req.headers.host}/api/callback`;

  try {
    const params = new URLSearchParams({
      grant_type: 'authorization_code',
      client_id: process.env.ZOHO_CLIENT_ID,
      client_secret: process.env.ZOHO_CLIENT_SECRET,
      redirect_uri: redirectUri,
      code,
    });

    const resp = await fetch(`${accountsDomain}/oauth/v2/token`, {
      method: 'POST',
      body: params,
    });
    const data = await resp.json();

    if (!data.refresh_token) {
      res.status(400).send(
        `<pre>Token exchange failed. This usually means the code already expired (they last ~1-2 minutes) — go back and re-authorize.\n\nResponse from Zoho:\n${JSON.stringify(data, null, 2)}</pre>`
      );
      return;
    }

    res.setHeader('Content-Type', 'text/html');
    res.status(200).send(`
      <html>
        <body style="font-family: sans-serif; max-width: 700px; margin: 40px auto; line-height: 1.5;">
          <h2>✅ Authorization successful</h2>
          <p>Copy the value below into your Vercel project's Environment Variables as
             <code>ZOHO_REFRESH_TOKEN</code>, then redeploy.</p>
          <p style="background:#f4f4f4; padding: 12px; border-radius: 6px; word-break: break-all;">
            <strong>${data.refresh_token}</strong>
          </p>
          <p>This value will not be shown again by Zoho unless you re-authorize with
             <code>prompt=consent</code>. Store it somewhere safe (a password manager, not a chat).</p>
        </body>
      </html>
    `);
  } catch (err) {
    res.status(500).send('Error exchanging code: ' + err.message);
  }
}
