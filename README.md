# CalSync — Calendar Free/Busy Sync for Consultants

Automatically syncs your availability across multiple Microsoft 365 and Google Calendar
tenants. When you have a meeting in one calendar, it sends individual "Unavailable – Your Name"
busy block invites to your other identities — so clients always see you as unavailable,
without sharing event details.

Designed for consultants who work across multiple locked-down client tenants where
delegated access and calendar sharing are not available.

---

## How it works

1. Reads all your calendars via private ICS feed URLs (no API access needed for client tenants)
2. For each event, sends individual calendar invites to your other configured identities
3. Each invitee only sees themselves on the invite — no other addresses visible
4. Diff-based: only creates/updates/cancels when something actually changed
5. Runs on a schedule via GitHub Actions (free) — no server needed

---

## Quick start

### 1. Clone the repo
```bash
git clone https://github.com/YOUR_USERNAME/calsync.git
cd calsync
python3 -m venv venv && source venv/bin/activate  # Windows: venv\Scripts\activate
pip install -r requirements.txt
```

### 2. Configure your calendars
```bash
cp config.example.yaml config.yaml
```
Edit `config.yaml` with your details. See comments in the file for guidance.

**Getting your ICS URLs:**
- **Microsoft 365:** Outlook web → Settings → Calendar → Shared calendars → Publish a calendar → Copy ICS link
- **Google:** calendar.google.com → Settings → your calendar → Secret address in iCal format

### 3. Set up Google API credentials
1. Go to [console.cloud.google.com](https://console.cloud.google.com)
2. Create a project → Enable **Google Calendar API**
3. APIs & Services → Credentials → Create OAuth 2.0 Client ID → Desktop app
4. Download JSON → save as `google_credentials.json` in this folder

### 4. Set up Microsoft Azure app registration
1. Go to [portal.azure.com](https://portal.azure.com) — sign in with a personal Microsoft account
2. App registrations → New registration
   - Name: `CalSync`
   - Supported account types: **Multitenant and personal Microsoft accounts**
3. Authentication → Add platform → Mobile and desktop
   - Tick: `https://login.microsoftonline.com/common/oauth2/nativeclient`
   - Enable **Allow public client flows**
4. API permissions → Add → Microsoft Graph → Delegated → `Calendars.ReadWrite`
5. Copy the **Application (client) ID** into `config.yaml`

### 5. First run — triggers authentication
```bash
python sync.py
```
- Google: browser window opens for OAuth consent
- M365: device code prompt appears — visit the URL shown and enter the code

### 6. Schedule with GitHub Actions (free, recommended)
1. Push to a **private** GitHub repo
2. Add these secrets under Settings → Secrets → Actions:

| Secret | How to get it |
|---|---|
| `GOOGLE_CREDENTIALS_JSON` | `python -c "import base64; print(base64.b64encode(open('google_credentials.json','rb').read()).decode())"` |
| `GOOGLE_TOKEN_B64` | `python -c "import base64; print(base64.b64encode(open('.tokens/google_token.pickle','rb').read()).decode())"` |
| `M365_TOKEN_CACHE_B64` | `python -c "import base64; print(base64.b64encode(open('.tokens/m365_token_cache.json','rb').read()).decode())"` |
| `CONFIG_YAML` | `python -c "import base64; print(base64.b64encode(open('config.yaml','rb').read()).decode())"` |
| `SYNC_STATE` | `{}` |
| `GH_PAT` | GitHub → Settings → Developer settings → Personal access tokens → Classic → repo scope |

3. The workflow runs at 07:00 and 19:00 UTC daily. Trigger manually from the Actions tab anytime.

---

## Adding a new client

Add an entry to `config.yaml`:
```yaml
- name:    "Client B"
  email:   "you@clientb.com"
  domain:  "clientb.com"
  ics_url: "https://..."
```
Then update the `CONFIG_YAML` GitHub secret with the new base64-encoded config.

---

## Adjusting sync frequency

Edit `.github/workflows/sync.yml`:
```yaml
schedule:
  - cron: '0 */4 * * *'  # Every 4 hours
```

---

## Troubleshooting

| Issue | Fix |
|---|---|
| Google auth loop | Delete `.tokens/google_token.pickle` and re-run |
| M365 auth loop | Delete `.tokens/m365_token_cache.json` and re-run |
| ICS fetch 400/404 | Regenerate the ICS URL in your calendar settings |
| Sync state error | Set `SYNC_STATE` secret to `{}` |
| Events not cancelling | Set `SYNC_STATE` secret to `{}` to force full re-sync |

---

## Security

- ICS URLs are secret — treat them like passwords
- Use a **private** GitHub repo
- Tokens in `.tokens/` are gitignored — never commit them
- The sync only ever writes opaque "Unavailable" blocks — no event details are shared

---

## Limitations

- M365 tokens expire after ~90 days of inactivity — re-run locally to refresh
- Client tenants must allow ICS URL generation (most do by default)
- Requires one-time OAuth setup for Google and M365

---

## Contributing

PRs welcome. Key areas for improvement:
- Support for Apple iCalendar / iCloud
- Web UI for config management
- Automatic ICS URL refresh
