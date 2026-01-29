#!/usr/bin/env python3
"""
Office365 Email Scanner - Interactive menu-driven CLI
Pulls emails from Office365 mailboxes via Microsoft Graph API.
"""

import os
import sys
import json
import requests
from datetime import datetime, timedelta
from dotenv import load_dotenv
import pytz

# ── Constants ────────────────────────────────────────────────────────────────

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
ENV_FILE = os.path.join(SCRIPT_DIR, '..', '..', 'Office365API', '.env')
USERS_FILE = os.path.join(SCRIPT_DIR, 'crc_users.txt')
DOMAINS_FILE = os.path.join(SCRIPT_DIR, 'search_domains.txt')
FOLDERS_FILE = os.path.join(SCRIPT_DIR, 'email_folders.txt')

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
TOKEN_URL = "https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token"
SCOPE = "https://graph.microsoft.com/.default"

EST = pytz.timezone('America/New_York')
GREEN = '\033[32m'
YELLOW = '\033[33m'
CYAN = '\033[36m'
RESET = '\033[0m'


# ── File Readers ─────────────────────────────────────────────────────────────

def read_lines_from_file(filepath):
    """Read non-empty, stripped lines from a text file."""
    if not os.path.exists(filepath):
        print(f"Error: File not found: {filepath}")
        sys.exit(1)
    with open(filepath, 'r') as f:
        return [line.strip() for line in f if line.strip()]


# ── Interactive Menus ────────────────────────────────────────────────────────

def menu_select(title, items, allow_all=True, allow_skip=False):
    """
    Display a numbered menu and return selected items.
    User can pick individual numbers (comma-separated), 'A' for all, or 'X' to skip.
    Returns None if skipped.
    """
    print(f"\n{CYAN}── {title} ──{RESET}")
    for i, item in enumerate(items, 1):
        print(f"  {i}. {item}")
    if allow_all:
        print(f"  A. ALL")
    if allow_skip:
        print(f"  X. SKIP (no filter)")

    while True:
        choice = input(f"\n{GREEN}Select (e.g. 1,3 or A{' or X' if allow_skip else ''}): {RESET}").strip()
        if not choice:
            continue

        if allow_skip and choice.upper() == 'X':
            return None

        if allow_all and choice.upper() == 'A':
            return list(items)

        try:
            indices = [int(x.strip()) for x in choice.split(',')]
            selected = []
            for idx in indices:
                if 1 <= idx <= len(items):
                    selected.append(items[idx - 1])
                else:
                    print(f"  Invalid number: {idx}. Valid range: 1-{len(items)}")
                    selected = []
                    break
            if selected:
                return selected
        except ValueError:
            print("  Invalid input. Enter numbers separated by commas, or A for all.")


def menu_select_domains(title, items):
    """
    Display a numbered menu for domain selection with option to type custom domain.
    User can pick numbers (comma-separated), 'A' for all, 'X' to skip, or 'T' to type.
    Returns None if skipped.
    """
    print(f"\n{CYAN}── {title} ──{RESET}")
    for i, item in enumerate(items, 1):
        print(f"  {i}. {item}")
    print(f"  A. ALL")
    print(f"  T. TYPE custom domain/email")
    print(f"  X. SKIP (no filter)")

    while True:
        choice = input(f"\n{GREEN}Select (e.g. 1,3 or A or T or X): {RESET}").strip()
        if not choice:
            continue

        if choice.upper() == 'X':
            return None

        if choice.upper() == 'A':
            return list(items)

        if choice.upper() == 'T':
            custom = input(f"{GREEN}Enter domain or email (e.g. gmail.com or user@example.com): {RESET}").strip()
            if custom:
                return [custom]
            print("  No input provided.")
            continue

        try:
            indices = [int(x.strip()) for x in choice.split(',')]
            selected = []
            for idx in indices:
                if 1 <= idx <= len(items):
                    selected.append(items[idx - 1])
                else:
                    print(f"  Invalid number: {idx}. Valid range: 1-{len(items)}")
                    selected = []
                    break
            if selected:
                return selected
        except ValueError:
            print("  Invalid input. Enter numbers, A for all, T to type, or X to skip.")


def menu_date(prompt, default):
    """Prompt for a date with a default value. Returns YYYY-MM-DD string."""
    default_str = default.strftime('%Y-%m-%d')
    while True:
        value = input(f"{GREEN}{prompt} [{default_str}]: {RESET}").strip()
        if not value:
            return default_str
        try:
            datetime.strptime(value, '%Y-%m-%d')
            return value
        except ValueError:
            print("  Invalid date format. Use YYYY-MM-DD.")


# ── Authentication ───────────────────────────────────────────────────────────

def get_access_token():
    """
    Get an OAuth2 access token using client credentials flow.
    Uses cached token from .env if less than 1 hour old.
    """
    load_dotenv(ENV_FILE, override=True)

    tenant_id = os.getenv("TENANT_ID")
    client_id = os.getenv("CLIENT_ID")
    client_secret = os.getenv("CLIENT_SECRET")

    if not all([tenant_id, client_id, client_secret]):
        print(f"Error: Missing credentials in {ENV_FILE}")
        print("Required: TENANT_ID, CLIENT_ID, CLIENT_SECRET")
        sys.exit(1)

    # Check cached token
    cached_token = os.getenv("ACCESS_TOKEN")
    timestamp_str = os.getenv("TOKEN_TIMESTAMP")

    if cached_token and timestamp_str:
        try:
            ts = datetime.fromisoformat(timestamp_str)
            if ts.tzinfo is None:
                ts = pytz.UTC.localize(ts)
            age = datetime.now(pytz.UTC) - ts.astimezone(pytz.UTC)
            if age < timedelta(hours=1):
                remaining = timedelta(hours=1) - age
                mins, secs = divmod(int(remaining.total_seconds()), 60)
                print(f"{GREEN}Using cached token (expires in {mins}m {secs}s){RESET}")
                return cached_token
        except ValueError:
            pass

    # Request new token
    print(f"{GREEN}Requesting new access token...{RESET}")
    url = TOKEN_URL.format(tenant=tenant_id)
    payload = {
        'client_id': client_id,
        'scope': SCOPE,
        'client_secret': client_secret,
        'grant_type': 'client_credentials'
    }

    try:
        resp = requests.post(url, data=payload,
                             headers={'Content-Type': 'application/x-www-form-urlencoded'})
        resp.raise_for_status()
        token = resp.json().get('access_token')
        if not token:
            print("Error: No access_token in response")
            sys.exit(1)

        # Cache token in .env
        _save_token_to_env(token)
        print(f"{GREEN}New token obtained and cached.{RESET}")
        return token

    except requests.exceptions.RequestException as e:
        print(f"Error obtaining token: {e}")
        if hasattr(e, 'response') and e.response is not None:
            try:
                print(json.dumps(e.response.json(), indent=2))
            except Exception:
                pass
        sys.exit(1)


def _save_token_to_env(token):
    """Save access token and timestamp to .env file."""
    now = datetime.now(pytz.UTC).isoformat()
    env_path = os.path.abspath(ENV_FILE)

    lines = []
    if os.path.exists(env_path):
        with open(env_path, 'r') as f:
            lines = f.readlines()

    token_found = False
    ts_found = False
    for i, line in enumerate(lines):
        if line.startswith('ACCESS_TOKEN='):
            lines[i] = f'ACCESS_TOKEN={token}\n'
            token_found = True
        elif line.startswith('TOKEN_TIMESTAMP='):
            lines[i] = f'TOKEN_TIMESTAMP={now}\n'
            ts_found = True

    if not token_found:
        lines.append(f'ACCESS_TOKEN={token}\n')
    if not ts_found:
        lines.append(f'TOKEN_TIMESTAMP={now}\n')

    with open(env_path, 'w') as f:
        f.writelines(lines)


# ── Graph API ────────────────────────────────────────────────────────────────

def get_folder_name(folder_id, token, user_email):
    """Resolve a folder ID to its display name. Results are cached."""
    if not hasattr(get_folder_name, '_cache'):
        get_folder_name._cache = {}
    if folder_id in get_folder_name._cache:
        return get_folder_name._cache[folder_id]

    url = f"{GRAPH_BASE}/users/{user_email}/mailFolders/{folder_id}"
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
    try:
        resp = requests.get(url, headers=headers)
        resp.raise_for_status()
        name = resp.json().get('displayName', folder_id)
    except Exception:
        name = folder_id

    get_folder_name._cache[folder_id] = name
    return name


def _get_folder_id(token, user_email, folder_name):
    """Resolve a folder display name to its ID."""
    # Map well-known folder names to their API aliases
    well_known_folders = {
        'inbox': 'inbox',
        'drafts': 'drafts',
        'sent items': 'sentitems',
        'sent': 'sentitems',
        'deleted items': 'deleteditems',
        'junk email': 'junkemail',
        'junk': 'junkemail',
        'archive': 'archive',
        'outbox': 'outbox',
    }

    # Try well-known folder alias first
    folder_lower = folder_name.lower()
    if folder_lower in well_known_folders:
        alias = well_known_folders[folder_lower]
        url = f"{GRAPH_BASE}/users/{user_email}/mailFolders/{alias}"
        headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
        try:
            resp = requests.get(url, headers=headers)
            if resp.status_code == 200:
                return resp.json().get('id')
        except Exception:
            pass

    # Fall back to searching all folders
    url = f"{GRAPH_BASE}/users/{user_email}/mailFolders"
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
    params = {"$top": 100}  # Get more folders to avoid pagination issues
    try:
        resp = requests.get(url, headers=headers, params=params)
        resp.raise_for_status()
        for folder in resp.json().get('value', []):
            if folder.get('displayName', '').lower() == folder_lower:
                return folder.get('id')
    except Exception as e:
        print(f"  Error resolving folder '{folder_name}': {e}")
    return None


def fetch_emails(token, user_email, search_domains, mindate, maxdate, folders):
    """
    Fetch emails from Graph API using $search for sender domains.
    Graph API does not allow $search combined with $filter or $orderby,
    so date filtering and sorting are done client-side.

    Returns a list of message dicts sorted ascending by date.
    """
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
        "Prefer": 'outlook.body-content-type="text"'
    }

    # Parse date bounds for client-side filtering
    min_dt = EST.localize(datetime.strptime(mindate, '%Y-%m-%d'))
    min_utc = min_dt.astimezone(pytz.UTC)

    max_dt = EST.localize(datetime.strptime(maxdate + ' 23:59:59', '%Y-%m-%d %H:%M:%S'))
    max_utc = max_dt.astimezone(pytz.UTC)

    # Build search query with both sender and date range
    query_parts = []
    
    # 1. Sender filter
    if search_domains:
        from_terms = [f'from:{domain}' for domain in search_domains]
        query_parts.append(f"({' OR '.join(from_terms)})")
    
    # 2. Date filter (Using KQL syntax for date range)
    # KQL date format: YYYY-MM-DD
    query_parts.append(f"received:{mindate}..{maxdate}")

    final_search_query = " AND ".join(query_parts)
    
    # Resolve folder names to IDs
    print("\n  Resolving folder IDs...", flush=True)
    folder_ids = {}
    for fname in folders:
        fid = _get_folder_id(token, user_email, fname)
        if fid:
            folder_ids[fid] = fname
            print(f"    {fname} -> OK", flush=True)
        else:
            print(f"    {fname} -> not found, skipping", flush=True)

    if not folder_ids:
        print("  No valid folders found.")
        return []

    all_messages = []

    # Query each folder separately
    for folder_id, folder_name in folder_ids.items():
        print(f"\n  Searching in {folder_name}...", flush=True)

        # Query messages from this specific folder
        url = f"{GRAPH_BASE}/users/{user_email}/mailFolders/{folder_id}/messages"
        params = {
            "$select": "subject,from,toRecipients,receivedDateTime,parentFolderId",
            "$top": 100,
            "$search": f'"{final_search_query}"'
        }

        page = 0
        folder_messages = []
        while url:
            try:
                # Debug: show full URL with params on first page
                if page == 0:
                    prepared = requests.Request("GET", url, headers=headers, params=params).prepare()
                    print(f"    DEBUG URL: {prepared.url}", flush=True)

                resp = requests.get(url, headers=headers, params=params if page == 0 else None)

                # Debug: show response on error
                if resp.status_code >= 400:
                    print(f"\n    DEBUG Response Status: {resp.status_code}")
                    try:
                        print(f"    DEBUG Response Body: {json.dumps(resp.json(), indent=2)}")
                    except Exception:
                        print(f"    DEBUG Response Text: {resp.text[:500]}")

                resp.raise_for_status()
                data = resp.json()

                messages = data.get('value', [])
                folder_messages.extend(messages)
                print(f"    Page {page + 1}: {len(messages)} messages (folder total: {len(folder_messages)})", flush=True)

                # Stop at 1000 messages per folder for safety
                if len(folder_messages) >= 1000:
                    folder_messages = folder_messages[:1000]
                    print("    (limited to 1000 messages)")
                    url = None
                    break

                url = data.get('@odata.nextLink')
                page += 1

                if page > 100:
                    print("    (page limit reached)")
                    break
            except requests.exceptions.HTTPError as e:
                print(f"    HTTP error: {e.response.status_code}")
                break
            except Exception as e:
                print(f"    Error: {e}")
                break

        all_messages.extend(folder_messages)

    print(f"\n  Total messages from all folders: {len(all_messages)}", flush=True)

    # Client-side filtering: date range
    # Although KQL filters by date, we double check to be precise with timezones if needed
    # but strictly speaking, KQL received:start..end includes the whole end day.
    # The previous logic had a precise UTC range. We'll keep the client-side check just in case
    # KQL returns slightly broader results.
    filtered = []
    for msg in all_messages:
        raw_dt = msg.get('receivedDateTime', '')
        try:
            dt = datetime.fromisoformat(raw_dt.replace('Z', '+00:00'))
            if dt.tzinfo is None:
                dt = pytz.UTC.localize(dt)
            if not (min_utc <= dt <= max_utc):
                continue
        except Exception:
            continue
        filtered.append(msg)

    # Sort ascending by date
    def parse_dt(m):
        try:
            return datetime.fromisoformat(m.get('receivedDateTime', '').replace('Z', '+00:00'))
        except Exception:
            return datetime.min.replace(tzinfo=pytz.UTC)

    filtered.sort(key=parse_dt)
    
    if len(filtered) < len(all_messages):
        print(f"\n  Filtered by precise date/time: {len(all_messages)} -> {len(filtered)} messages")

    return filtered


# ── Display ──────────────────────────────────────────────────────────────────

def display_results(messages, token, user_email):
    """Display messages in a formatted table."""
    if not messages:
        print("\nNo messages found.")
        return

    # Pre-fetch folder names
    folder_ids = {msg.get('parentFolderId') for msg in messages if msg.get('parentFolderId')}
    for fid in folder_ids:
        get_folder_name(fid, token, user_email)

    # Print table
    print(f"\n{'─' * 160}")
    print(f"{'#':<5} {'Date/Time':<22} {'From Address':<40} {'To Address':<35} {'Folder':<15} {'Subject':<40}")
    print(f"{'─' * 160}")

    for idx, msg in enumerate(messages, 1):
        # Date
        raw_dt = msg.get('receivedDateTime', '')
        try:
            dt = datetime.fromisoformat(raw_dt.replace('Z', '+00:00'))
            if dt.tzinfo is None:
                dt = pytz.UTC.localize(dt)
            dt_str = dt.astimezone(EST).strftime('%Y-%m-%d %H:%M:%S')
        except Exception:
            dt_str = 'N/A'

        # From
        from_info = msg.get('from', {}).get('emailAddress', {})
        from_addr = from_info.get('address', 'N/A')

        # To (first recipient)
        to_list = msg.get('toRecipients', [])
        to_addr = to_list[0].get('emailAddress', {}).get('address', 'N/A') if to_list else 'N/A'

        # Folder
        fid = msg.get('parentFolderId', '')
        folder = get_folder_name(fid, token, user_email) if fid else 'N/A'

        # Subject
        subject = msg.get('subject', '(No Subject)')

        # Truncate for display
        from_addr = (from_addr[:37] + '...') if len(from_addr) > 40 else from_addr
        to_addr = (to_addr[:32] + '...') if len(to_addr) > 35 else to_addr
        folder = (folder[:12] + '...') if len(folder) > 15 else folder
        subject = (subject[:37] + '...') if len(subject) > 40 else subject

        print(f"{idx:<5} {dt_str:<22} {from_addr:<40} {to_addr:<35} {folder:<15} {subject:<40}")

    print(f"{'─' * 160}")
    print(f"Total: {len(messages)} messages")


# ── Main ─────────────────────────────────────────────────────────────────────

def main():
    print(f"\n{CYAN}╔══════════════════════════════════════╗{RESET}")
    print(f"{CYAN}║    Office365 Email Scanner CLI       ║{RESET}")
    print(f"{CYAN}╚══════════════════════════════════════╝{RESET}")

    # Load menu data from files
    users = read_lines_from_file(USERS_FILE)
    domains = read_lines_from_file(DOMAINS_FILE)
    folders = read_lines_from_file(FOLDERS_FILE)

    # 1. Select users
    selected_users = menu_select("Select User Mailbox", users)

    # 2. Select from domains (X to skip, T to type custom)
    selected_domains = menu_select_domains("Select From (sender domain/address)", domains)

    # 3. Select min date
    default_min = datetime.now() - timedelta(days=10)
    mindate = menu_date("Min date (earliest)", default_min)

    # 4. Select max date
    default_max = datetime.now()
    maxdate = menu_date("Max date (latest)", default_max)

    # 5. Select folders
    selected_folders = menu_select("Select Mail Folders", folders)

    # Confirm selections
    print(f"\n{CYAN}── Summary ──{RESET}")
    print(f"  Users:   {', '.join(selected_users)}")
    print(f"  From:    {', '.join(selected_domains) if selected_domains else '(all senders)'}")
    print(f"  Dates:   {mindate} to {maxdate}")
    print(f"  Folders: {', '.join(selected_folders)}")
    print(f"  Sort:    ascending (earliest first)")

    proceed = input(f"\n{GREEN}Proceed? (Y/n): {RESET}").strip()
    if proceed.lower() == 'n':
        print("Cancelled.")
        sys.exit(0)

    # Authenticate
    token = get_access_token()

    # Fetch and display for each selected user
    for user_email in selected_users:
        print(f"\n{CYAN}{'═' * 60}{RESET}")
        print(f"{CYAN}  Mailbox: {user_email}{RESET}")
        print(f"{CYAN}{'═' * 60}{RESET}")

        messages = fetch_emails(token, user_email, selected_domains,
                                mindate, maxdate, selected_folders)
        display_results(messages, token, user_email)


if __name__ == "__main__":
    main()
