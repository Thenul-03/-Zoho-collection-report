# Zoho Books → Daily Collection Excel Report (Vercel)

Generates the school's Daily Cash Collection report, pivoted into
Cash / Chq / CC / DT / MY FEES / Bank columns, straight from Zoho Books
customer payments — no manual export/reformat step.

## 1. Deploy this project to Vercel

- Push this folder to a GitHub repo (or use `vercel` CLI directly from this folder).
- In Vercel, "Add New Project" → import the repo → Deploy.
- Note your project's URL, e.g. `https://your-project.vercel.app`.

## 2. Update your Zoho API Console client

Go back to https://api-console.zoho.com/, open the client you created, and
set (or add) this **Authorized Redirect URI**:

```
https://your-project.vercel.app/api/callback
```

(You can keep `http://localhost:8080` too if you want a local fallback —
Zoho allows multiple redirect URIs on one client.)

## 3. Set Environment Variables in Vercel

Project → Settings → Environment Variables. Add:

| Name | Value |
|---|---|
| `ZOHO_CLIENT_ID` | from API Console |
| `ZOHO_CLIENT_SECRET` | from API Console |
| `ZOHO_ORGANIZATION_ID` | Zoho Books → Settings → Organization Profile |
| `ZOHO_ACCOUNTS_DOMAIN` | `https://accounts.zoho.com` (use `.in` if your org is on the India data center) |
| `ZOHO_API_DOMAIN` | `https://www.zohoapis.com` (use `.in` to match, if applicable) |
| `REPORT_ACCESS_KEY` | any long random string you make up — this protects the report endpoint from public access |
| `ZOHO_REFRESH_TOKEN` | leave blank for now — you'll fill this in during step 4 |

Redeploy after saving these (Vercel does this automatically on env var changes,
or trigger a redeploy manually).

## 4. Authorize once, to get the refresh token

Build this URL with your real Client ID and visit it in your browser
(swap `accounts.zoho.com` for `.in` if needed):

```
https://accounts.zoho.com/oauth/v2/auth?scope=ZohoBooks.customerpayments.READ,ZohoBooks.contacts.READ,ZohoBooks.settings.READ&client_id=YOUR_CLIENT_ID&response_type=code&access_type=offline&redirect_uri=https://your-project.vercel.app/api/callback&prompt=consent
```

Log in, click Accept. You'll land on `/api/callback`, which shows your
`refresh_token` directly in the page — copy it into the `ZOHO_REFRESH_TOKEN`
environment variable in Vercel, and redeploy.

That's the only manual step. From here on, the deployed app runs
independently — no one needs to touch localhost or a terminal again.

## 5. Generate a report — the easy way (preview page)

Just visit your site's homepage:

```
https://your-project.vercel.app/
```

Pick a From/To date and paste in your `REPORT_ACCESS_KEY` once (it's
remembered in that browser after the first time). Click **Run Report** and
the page pulls the payments straight from Zoho Books and shows them in a
table right there — same live-connection pattern as the cheque-writer tool's
"Recent Payments" list, with a green **Zoho Connected** / red **Zoho
Disconnected** badge in the top right so you always know whether the
connection to Zoho is working before you trust the numbers.

Once you're happy with the preview, click **Download Excel** to get the same
`.xlsx` file the old button used to generate directly.

Under the hood this preview is powered by a new endpoint, `/api/data.js`,
which shares all of its Zoho-fetching logic with `/api/report.js` via
`lib/zoho.js` — so the numbers on screen and the numbers in the Excel file
can never drift apart; add a payment mode or account mapping once and both
update.

### Or generate a report directly by URL (no preview)

```
https://your-project.vercel.app/api/report?key=YOUR_REPORT_ACCESS_KEY&from=2026-09-01&to=2026-09-01
```

### Or fetch the raw JSON (e.g. to build another view)

```
https://your-project.vercel.app/api/data?key=YOUR_REPORT_ACCESS_KEY&from=2026-09-01&to=2026-09-01
```

## 5b. If Admission Number comes back blank

Visit the same URL with `&debug=1` added, e.g.:

```
https://your-project.vercel.app/api/report?key=YOUR_REPORT_ACCESS_KEY&from=2026-09-01&to=2026-09-01&debug=1
```

Instead of downloading a file, this shows raw JSON: one sample payment and
its full Contact record exactly as Zoho returns it. Look inside
`sample_contact.contact.custom_fields` (or `custom_field_hash`) for how your
Admission Number field is actually labeled — Zoho's API sometimes uses a
slightly different label than what's shown in the UI. If it doesn't match
`admission number` (case/spacing-insensitive), open `api/report.js` and
update the `ADMISSION_FIELD_LABEL` constant near `extractAdmissionNumber` to
match exactly what you see.

## 6. Adjust the mappings for your organization

Open `lib/zoho.js` and edit the two maps near the top:

- `PAYMENT_MODE_TO_COLUMN` — maps Zoho's `payment_mode` values (e.g.
  `cash`, `creditcard`) to the Cash/Chq/CC/DT/MY FEES columns. Check your
  own org's exact mode spelling if amounts aren't landing where expected.
- `ACCOUNT_TO_BANK_LABEL` — maps each "Deposit To" account name in Zoho to
  the short label you want in the Bank column (e.g. `SEYLAN`, `-`).

## Notes / limitations

- **Admission Number** must exist as a custom field on the Customer/Contact
  record in Zoho Books, labeled exactly `Admission Number` (case-insensitive
  match is handled, but the wording must match). If it's blank in the
  output, double check the field label and that the contact actually has
  a value set.
- The report endpoint is protected only by the `REPORT_ACCESS_KEY` query
  parameter — treat that key like a password. For stronger protection,
  consider adding Vercel's built-in password protection (Pro plans) or
  swapping in proper authentication if this will be used by multiple staff.
- Large date ranges fetch one extra API call per unique student (to look up
  Admission Number), so keep ranges reasonable (a day or a month) rather
  than pulling a full year in one request.
