# Motive odometer emails to shops

Mon/Wed/Fri 6:45 AM Pacific (GitHub Actions): pulls the latest odometer for every truck in the Motive
"Active Vehicles" group, reads the Ryder Truck Mapping file from Dropbox, and emails each shop ONLY the
trucks the mapping file assigns to it (via Microsoft 365 / Graph, from the Administration mailbox).

- Settings: `config.json`. Sender and recipient addresses: the `RECIPIENTS_JSON` secret (kept out of the repo).
- Manual runs: Actions tab → "Shop odometer emails" → Run workflow → mode `dry-run` (no email),
  `test` (all emails to `test_to` only), or `send`.
- Secrets (Settings → Secrets and variables → Actions): `MOTIVE_API_KEY`, `DROPBOX_APP_KEY`,
  `DROPBOX_REFRESH_TOKEN`, `MS_TENANT_ID`, `MS_CLIENT_ID`, `MS_CLIENT_SECRET`, `RECIPIENTS_JSON`.
