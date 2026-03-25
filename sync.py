#!/usr/bin/env python3
"""
Calendar Free/Busy Sync Tool
-----------------------------
- Reads ICS feeds from all configured calendars (M365 + Google tenants)
- For each event, invites all OTHER configured identities (excluding the source domain)
- Diff-based: only creates/updates/cancels when something actually changed
- Designed to run as a cron job (locally or via GitHub Actions)
"""

import os
import sys
import json
import logging
import hashlib
import pickle
import uuid
from datetime import datetime, timedelta, timezone

import yaml
import requests
from icalendar import Calendar
import msal
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

# ── Logging ───────────────────────────────────────────────────────────────────

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)]
)
log = logging.getLogger(__name__)

# ── Constants ─────────────────────────────────────────────────────────────────

CONFIG_PATH   = os.path.join(os.path.dirname(__file__), "config.yaml")
TOKEN_DIR     = os.path.join(os.path.dirname(__file__), ".tokens")
STATE_FILE    = os.path.join(os.path.dirname(__file__), ".sync_state.json")
SYNC_TAG      = "CALSYNCTOOL"
GOOGLE_SCOPES = ["https://www.googleapis.com/auth/calendar"]
SYNC_DAYS     = 30


# ── Config ────────────────────────────────────────────────────────────────────

def load_config() -> dict:
    with open(CONFIG_PATH) as f:
        return yaml.safe_load(f)


# ── Time helpers ──────────────────────────────────────────────────────────────

def now_utc() -> datetime:
    return datetime.now(timezone.utc)

def window_end() -> datetime:
    return now_utc() + timedelta(days=SYNC_DAYS)

def to_utc(dt) -> datetime:
    if not isinstance(dt, datetime):
        dt = datetime(dt.year, dt.month, dt.day, 0, 0, 0)
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt.astimezone(timezone.utc)


# ── ICS Fetching ──────────────────────────────────────────────────────────────

def fetch_ics_events(url: str, source_domain: str) -> list[dict]:
    """
    Fetch and parse an ICS URL.
    Returns list of {start, end, uid, source_domain} within the sync window.
    Skips events previously written by this tool.
    """
    log.info(f"  Fetching: {url[:70]}...")
    resp = requests.get(url, timeout=30)
    resp.raise_for_status()
    cal = Calendar.from_ical(resp.content)

    start_window = now_utc()
    end_window   = window_end()
    events = []

    for component in cal.walk():
        if component.name != "VEVENT":
            continue

        # Skip blocks previously written by this tool
        if SYNC_TAG in str(component.get("DESCRIPTION", "")):
            continue

        dtstart = component.get("DTSTART")
        dtend   = component.get("DTEND")
        if not dtstart or not dtend:
            continue

        start = to_utc(dtstart.dt)
        end   = to_utc(dtend.dt)

        if end <= start_window or start >= end_window:
            continue

        events.append({
            "start":         start,
            "end":           end,
            "uid":           str(component.get("UID", uuid.uuid4())),
            "source_domain": source_domain,
        })

    log.info(f"    → {len(events)} events in window")
    return events


def collect_all_events(cfg: dict) -> list[dict]:
    all_events = []
    for cal in cfg["calendars"]:
        if not cal.get("sync_to_others", True):
            log.info(f"  Skipping (sync_to_others: false): {cal.get('name', cal['domain'])}")
            continue
        try:
            events = fetch_ics_events(cal["ics_url"], cal["domain"])
            all_events.extend(events)
        except Exception as e:
            log.warning(f"  Failed to fetch '{cal.get('name', cal['domain'])}': {e}")
    return all_events


# ── Diffing ───────────────────────────────────────────────────────────────────

def event_fingerprint(event: dict, invitees: list[str]) -> str:
    key = f"{event['start'].isoformat()}|{event['end'].isoformat()}|{','.join(sorted(invitees))}"
    return hashlib.sha256(key.encode()).hexdigest()

def load_state() -> dict:
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE) as f:
            return json.load(f)
    return {}

def save_state(state: dict):
    with open(STATE_FILE, "w") as f:
        json.dump(state, f, indent=2)


# ── Invitee Resolution ────────────────────────────────────────────────────────

def resolve_invitees(source_domain: str, cfg: dict) -> list[str]:
    """
    All configured email addresses EXCEPT:
    - Those in the source domain (event already lives there)
    - Those marked with invite: false (read-only sources, never invited)
    Each invitee gets their OWN individual invite (no attendee list visible to others).
    """
    return [
        cal["email"]
        for cal in cfg["calendars"]
        if cal["domain"].lower() != source_domain.lower()
        and cal.get("invite", True)
    ]


# ── Google Calendar ───────────────────────────────────────────────────────────

def get_google_service(credentials_path: str):
    os.makedirs(TOKEN_DIR, exist_ok=True)
    token_path = os.path.join(TOKEN_DIR, "google_token.pickle")
    creds = None

    if os.path.exists(token_path):
        with open(token_path, "rb") as f:
            creds = pickle.load(f)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(credentials_path, GOOGLE_SCOPES)
            creds = flow.run_local_server(port=0)
        with open(token_path, "wb") as f:
            pickle.dump(creds, f)

    return build("calendar", "v3", credentials=creds)


def find_google_event(service, calendar_id: str, sync_id: str) -> str | None:
    results = service.events().list(
        calendarId=calendar_id,
        privateExtendedProperty=f"syncEventId={sync_id}",
        timeMin=now_utc().isoformat(),
        timeMax=window_end().isoformat(),
    ).execute()
    items = results.get("items", [])
    return items[0]["id"] if items else None


def upsert_google_event(service, calendar_id: str, event: dict,
                        invitees: list[str], display_name: str, sync_id: str):
    body = {
        "summary":      f"Unavailable – {display_name}",
        "description":  f"[{SYNC_TAG}] Auto-synced busy block. Do not edit manually.",
        "start":        {"dateTime": event["start"].isoformat()},
        "end":          {"dateTime": event["end"].isoformat()},
        "transparency": "opaque",
        "visibility":   "private",
        "attendees":    [{"email": addr} for addr in invitees],
        "guestsCanSeeOtherGuests": False,
        "guestsCanInviteOthers":   False,
        "extendedProperties": {
            "private": {"syncTag": SYNC_TAG, "syncEventId": sync_id}
        },
    }
    existing_id = find_google_event(service, calendar_id, sync_id)
    if existing_id:
        service.events().update(calendarId=calendar_id, eventId=existing_id,
                                body=body, sendUpdates="all").execute()
        log.info(f"    [Google] Updated {sync_id[:16]}...")
    else:
        service.events().insert(calendarId=calendar_id, body=body,
                                sendUpdates="all").execute()
        log.info(f"    [Google] Created {sync_id[:16]}...")


def cancel_google_event(service, calendar_id: str, sync_id: str):
    existing_id = find_google_event(service, calendar_id, sync_id)
    if existing_id:
        service.events().delete(calendarId=calendar_id, eventId=existing_id,
                                sendUpdates="all").execute()
        log.info(f"    [Google] Cancelled {sync_id[:16]}...")


# ── Microsoft 365 ─────────────────────────────────────────────────────────────

def get_m365_token(cfg: dict) -> str:
    os.makedirs(TOKEN_DIR, exist_ok=True)
    cache_path = os.path.join(TOKEN_DIR, "m365_token_cache.json")

    cache = msal.SerializableTokenCache()
    if os.path.exists(cache_path):
        cache.deserialize(open(cache_path).read())

    app = msal.PublicClientApplication(
        cfg["m365"]["client_id"],
        authority=f"https://login.microsoftonline.com/{cfg['m365']['tenant_id']}",
        token_cache=cache,
    )

    scopes   = ["Calendars.ReadWrite"]
    accounts = app.get_accounts()
    result   = app.acquire_token_silent(scopes, account=accounts[0]) if accounts else None

    if not result:
        flow = app.initiate_device_flow(scopes=scopes)
        if "error" in flow:
            raise RuntimeError(f"M365 device flow failed: {flow.get('error')}: {flow.get('error_description')}")
        # Print auth instructions — 'message' contains the URL and code to enter
        message = flow.get("message") or (
            f"Go to {flow.get('verification_uri', 'https://microsoft.com/devicelogin')} "
            f"and enter code: {flow.get('user_code', '(see above)')}"
        )
        print(f"\n{'='*60}")
        print("M365 Authentication required:")
        print(message)
        print(f"{'='*60}\n")
        result = app.acquire_token_by_device_flow(flow)

    if "access_token" not in result:
        raise RuntimeError(f"M365 auth failed: {result.get('error')} - {result.get('error_description')}")

    with open(cache_path, "w") as f:
        f.write(cache.serialize())

    return result["access_token"]


def graph_headers(token: str) -> dict:
    return {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

# Extended property ID for storing our sync metadata in M365
MAPI_SYNC_TAG_ID      = "String {00020329-0000-0000-C000-000000000046} Name syncTag"
MAPI_SYNC_EVENT_ID_ID = "String {00020329-0000-0000-C000-000000000046} Name syncEventId"


def find_m365_event(token: str, sync_id: str) -> str | None:
    url = (
        "https://graph.microsoft.com/v1.0/me/events"
        f"?$filter=singleValueExtendedProperties/any(ep:ep/id eq '{MAPI_SYNC_EVENT_ID_ID}'"
        f" and ep/value eq '{sync_id}')"
        f"&$expand=singleValueExtendedProperties($filter=id eq '{MAPI_SYNC_EVENT_ID_ID}')"
    )
    resp  = requests.get(url, headers=graph_headers(token))
    items = resp.json().get("value", [])
    return items[0]["id"] if items else None


def upsert_m365_event(token: str, event: dict, invitees: list[str],
                      display_name: str, sync_id: str):
    body = {
        "subject": f"Unavailable – {display_name}",
        "body":    {"contentType": "text",
                    "content": f"[{SYNC_TAG}] Auto-synced busy block. Do not edit manually."},
        "start":   {"dateTime": event["start"].strftime("%Y-%m-%dT%H:%M:%S"), "timeZone": "UTC"},
        "end":     {"dateTime": event["end"].strftime("%Y-%m-%dT%H:%M:%S"),   "timeZone": "UTC"},
        "showAs":  "busy",
        "sensitivity": "private",
        "attendees": [
            {"emailAddress": {"address": addr}, "type": "required"}
            for addr in invitees
        ],
        "singleValueExtendedProperties": [
            {"id": MAPI_SYNC_TAG_ID,      "value": SYNC_TAG},
            {"id": MAPI_SYNC_EVENT_ID_ID, "value": sync_id},
        ],
    }
    existing_id = find_m365_event(token, sync_id)
    if existing_id:
        requests.patch(
            f"https://graph.microsoft.com/v1.0/me/events/{existing_id}",
            headers=graph_headers(token), json=body
        )
        log.info(f"    [M365] Updated {sync_id[:16]}...")
    else:
        requests.post(
            "https://graph.microsoft.com/v1.0/me/events",
            headers=graph_headers(token), json=body
        )
        log.info(f"    [M365] Created {sync_id[:16]}...")


def cancel_m365_event(token: str, sync_id: str):
    existing_id = find_m365_event(token, sync_id)
    if existing_id:
        requests.post(
            f"https://graph.microsoft.com/v1.0/me/events/{existing_id}/cancel",
            headers=graph_headers(token),
            json={"comment": "This availability block has been removed."}
        )
        log.info(f"    [M365] Cancelled {sync_id[:16]}...")


# ── Main ──────────────────────────────────────────────────────────────────────

def run_sync():
    log.info("=" * 60)
    log.info("Calendar Sync starting")
    log.info(f"Window: now → {window_end().strftime('%Y-%m-%d %H:%M UTC')}")
    log.info("=" * 60)

    cfg          = load_config()
    state        = load_state()
    display_name = cfg["display_name"]
    new_state    = {}

    # Step 1: Collect events
    log.info("\n[1/4] Fetching ICS feeds...")
    all_events = collect_all_events(cfg)
    log.info(f"  Total: {len(all_events)} events")

    # Step 2: Authenticate
    log.info("\n[2/4] Authenticating...")
    google_service = None
    m365_token     = None
    google_cal_cfg = next((c for c in cfg["calendars"] if c.get("is_primary_google")), None)
    m365_cal_cfg   = next((c for c in cfg["calendars"] if c.get("is_primary_m365")),   None)

    if google_cal_cfg:
        log.info("  Authenticating Google...")
        google_service = get_google_service(cfg["google"]["credentials_file"])

    if m365_cal_cfg:
        log.info("  Authenticating M365...")
        m365_token = get_m365_token(cfg)

    # Step 3: Diff and sync
    log.info("\n[3/4] Syncing changed events...")
    processed_ids = set()

    for event in all_events:
        invitees      = resolve_invitees(event["source_domain"], cfg)
        sync_id       = f"{SYNC_TAG}-{event['uid']}"
        fingerprint   = event_fingerprint(event, invitees)
        processed_ids.add(sync_id)

        if state.get(sync_id) == fingerprint:
            new_state[sync_id] = fingerprint
            continue  # No change, skip

        log.info(
            f"  → {event['start'].strftime('%b %d %H:%M')}–{event['end'].strftime('%H:%M')} "
            f"from {event['source_domain']} → individual invites to: {invitees}"
        )

        # Send a separate individual invite to each invitee (no attendee list visible to others)
        for invitee in invitees:
            per_invite_id = f"{sync_id}-{invitee.replace('@', '_').replace('.', '_')}"

            if google_service and google_cal_cfg:
                try:
                    upsert_google_event(
                        google_service,
                        google_cal_cfg.get("calendar_id", "primary"),
                        event, [invitee], display_name, per_invite_id
                    )
                except Exception as e:
                    log.warning(f"    Google failed for {invitee}: {e}")

            if m365_token:
                try:
                    upsert_m365_event(m365_token, event, [invitee], display_name, per_invite_id)
                except Exception as e:
                    log.warning(f"    M365 failed for {invitee}: {e}")

        new_state[sync_id] = fingerprint

    # Step 4: Cancel removed events
    # Build the full set of active per-invitee IDs from this run
    log.info("\n[4/4] Cancelling removed events...")
    active_per_invite_ids = set()
    for event in all_events:
        s_id = f"{SYNC_TAG}-{event['uid']}"
        invitees = resolve_invitees(event["source_domain"], cfg)
        for invitee in invitees:
            active_per_invite_ids.add(
                f"{s_id}-{invitee.replace('@', '_').replace('.', '_')}"
            )

    cancelled = 0
    for per_invite_id in set(state.keys()) - active_per_invite_ids:
        log.info(f"  Removing {per_invite_id[:40]}...")
        if google_service and google_cal_cfg:
            try:
                cancel_google_event(
                    google_service,
                    google_cal_cfg.get("calendar_id", "primary"),
                    per_invite_id
                )
            except Exception as e:
                log.warning(f"    Google cancel failed: {e}")
        if m365_token:
            try:
                cancel_m365_event(m365_token, per_invite_id)
            except Exception as e:
                log.warning(f"    M365 cancel failed: {e}")
        cancelled += 1

    save_state(new_state)
    log.info(
        f"\n✓ Done. {len(all_events)} events processed, {cancelled} cancelled."
    )


if __name__ == "__main__":
    run_sync()
