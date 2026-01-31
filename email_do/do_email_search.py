import requests
import json
import os
import configparser
from datetime import datetime, timedelta
from dotenv import load_dotenv, set_key, find_dotenv

# ============================================================================
# CONFIGURATION
# ============================================================================

# Load credentials from .env file
load_dotenv()

TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")

# Token management
ACCESS_TOKEN = os.getenv("ACCESS_TOKEN", "")
TOKEN_GENERATED_AT = os.getenv("TOKEN_GENERATED_AT", "")

# Safety margin - refresh token 5 minutes before expiry (in seconds)
TOKEN_REFRESH_MARGIN = 300

# Load search parameters from do.config
config = configparser.ConfigParser()
config.read(os.path.join(os.path.dirname(os.path.abspath(__file__)), "do.config"))

USERS = [u.strip() for u in config.get("mailbox", "users", fallback="").split(",") if u.strip()]
START_DATE = config.get("dates", "start_date", fallback="2025-12-01") + "T00:00:00Z"
END_DATE = config.get("dates", "end_date", fallback="2026-01-30") + "T23:59:59Z"
TOP = config.getint("messages", "top", fallback=500)
FOLDERS = [f.strip() for f in config.get("folders", "folders", fallback="SentItems").split(",") if f.strip()]
FILTER_DOMAINS = [d.strip().lower() for d in config.get("domains", "filter_domains", fallback="").split(",") if d.strip()]

# ============================================================================
# ACCESS TOKEN GENERATION
# ============================================================================

def is_token_expired():
    """Check if the current token is expired or about to expire"""
    
    if not ACCESS_TOKEN or not TOKEN_GENERATED_AT:
        print("No existing token found")
        return True
    
    try:
        # Parse the timestamp
        generated_at = datetime.fromisoformat(TOKEN_GENERATED_AT)
        
        # Tokens typically expire in 3600 seconds (1 hour)
        # We'll refresh 5 minutes before expiry to be safe
        expires_at = generated_at + timedelta(seconds=3600 - TOKEN_REFRESH_MARGIN)
        
        now = datetime.now()
        
        if now >= expires_at:
            time_since = (now - generated_at).total_seconds()
            print(f"Token expired (generated {int(time_since)} seconds ago)")
            return True
        else:
            time_remaining = (expires_at - now).total_seconds()
            print(f"✓ Existing token is valid (expires in {int(time_remaining)} seconds)")
            return False
            
    except Exception as e:
        print(f"Error checking token expiry: {e}")
        return True


def get_new_access_token():
    """Generate a new access token using client credentials"""
    
    print("Generating new access token...")
    
    token_url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    
    token_data = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials"
    }
    
    try:
        response = requests.post(token_url, data=token_data)
        response.raise_for_status()
        
        token_response = response.json()
        access_token = token_response.get("access_token")
        expires_in = token_response.get("expires_in", 3600)
        
        if not access_token:
            raise Exception("No access token in response")
        
        print(f"✓ New access token generated (expires in {expires_in} seconds)")
        
        # Save token to .env file
        save_token_to_env(access_token)
        
        return access_token
        
    except requests.exceptions.HTTPError as e:
        print(f"HTTP Error generating token: {e}")
        print(f"Response: {response.text}")
        raise
    except Exception as e:
        print(f"Error generating token: {e}")
        raise


def save_token_to_env(access_token):
    """Save the access token and timestamp to .env file"""
    
    env_file = find_dotenv()
    
    if not env_file:
        # Create .env file if it doesn't exist
        env_file = ".env"
        with open(env_file, 'w') as f:
            f.write("")
    
    # Update ACCESS_TOKEN in .env
    set_key(env_file, "ACCESS_TOKEN", access_token)
    
    # Save the current timestamp in ISO format
    timestamp = datetime.now().isoformat()
    set_key(env_file, "TOKEN_GENERATED_AT", timestamp)
    
    print(f"✓ Access token saved to {env_file}")
    print(f"✓ Token timestamp: {timestamp}")


def get_access_token():
    """Get access token - generate new one only if expired"""
    
    if is_token_expired():
        return get_new_access_token()
    else:
        return ACCESS_TOKEN


# ============================================================================
# MICROSOFT GRAPH API FUNCTIONS
# ============================================================================

def get_all_messages(access_token, user_email, folder):
    """Fetch all messages using pagination"""

    all_messages = []

    # Build initial URL
    base_url = f"https://graph.microsoft.com/v1.0/users/{user_email}/mailFolders/{folder}/messages"

    # Use $filter for date range (more reliable than $search)
    url = (
        f"{base_url}?"
        f"$select=subject,from,toRecipients,sentDateTime,parentFolderId&"
        f"$top={TOP}&"
        f"$filter=sentDateTime ge {START_DATE} and sentDateTime le {END_DATE}&"
        f"$orderby=sentDateTime asc"
    )
    
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/json"
    }
    
    page_count = 0
    
    while url:
        page_count += 1
        print(f"Fetching page {page_count}...")
        
        try:
            response = requests.get(url, headers=headers)
            response.raise_for_status()
            
            data = response.json()
            messages = data.get("value", [])
            
            print(f"  Retrieved {len(messages)} messages")
            
            # Add messages to our collection
            all_messages.extend(messages)
            
            # Check for next page
            url = data.get("@odata.nextLink")
            
            if url:
                print(f"  More results available, fetching next page...")
            else:
                print(f"  No more pages")
                
        except requests.exceptions.HTTPError as e:
            print(f"HTTP Error: {e}")
            print(f"Response: {response.text}")
            
            # If we get 401 Unauthorized, token might be invalid
            if response.status_code == 401:
                print("\n⚠ Token appears invalid. Generating new token...")
                new_token = get_new_access_token()
                headers["Authorization"] = f"Bearer {new_token}"
                print("Retrying request with new token...")
                continue
            
            break
        except Exception as e:
            print(f"Error: {e}")
            break
    
    print(f"\n{'='*60}")
    print(f"Total messages retrieved: {len(all_messages)}")
    print(f"{'='*60}\n")
    
    return all_messages


def filter_by_domains(messages, domains):
    """Filter messages by recipient domain(s)"""
    if not domains:
        return messages

    filtered = []

    for msg in messages:
        recipients = msg.get("toRecipients", [])

        for recipient in recipients:
            email = recipient.get("emailAddress", {}).get("address", "").lower()
            if any(domain in email for domain in domains):
                filtered.append(msg)
                break  # Found a match, move to next message

    return filtered


def get_contact(msg):
    """Return recipient(s) for SentItems, sender for all other folders."""
    folder = msg.get("_folder", "")
    if folder == "SentItems":
        recipients = msg.get("toRecipients", [])
        return "; ".join(r.get("emailAddress", {}).get("address", "") for r in recipients)
    else:
        return msg.get("from", {}).get("emailAddress", {}).get("address", "")


def print_message_table(messages):
    """Print messages in a formatted table"""
    if not messages:
        return

    # Build rows
    rows = []
    for i, msg in enumerate(messages, 1):
        num = str(i)
        user = msg.get("_user_mailbox", "").split("@")[0]
        folder = msg.get("_folder", "")
        subject = msg.get("subject", "(No subject)")
        sent_time = msg.get("sentDateTime", "")[:16].replace("T", " ")
        contact = get_contact(msg)
        rows.append((num, user, folder, subject, sent_time, contact))

    # Calculate column widths
    headers = ("#", "User", "Folder", "Subject", "Sent DateTime", "Recipient/Sender")
    widths = [len(h) for h in headers]
    for row in rows:
        for i, val in enumerate(row):
            widths[i] = max(widths[i], len(val))

    # Cap subject and recipient widths for readability
    widths[3] = min(widths[3], 50)
    widths[5] = min(widths[5], 50)

    def truncate(s, w):
        return s if len(s) <= w else s[:w - 3] + "..."

    # Print table
    fmt = "  ".join(f"{{:<{w}}}" for w in widths)
    sep = "  ".join("─" * w for w in widths)

    print(f"\n{fmt.format(*headers)}")
    print(sep)
    for row in rows:
        print(fmt.format(*(truncate(val, w) for val, w in zip(row, widths))))
    print(sep)
    print(f"{len(rows)} message(s)\n")


def save_to_json(messages, filename="messages.json"):
    """Save messages to JSON file"""
    with open(filename, 'w', encoding='utf-8') as f:
        json.dump(messages, f, indent=2, ensure_ascii=False)
    
    print(f"Messages saved to {filename}")


def save_to_csv(messages, filename="messages.csv"):
    """Save messages to CSV file"""
    import csv

    with open(filename, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)

        # Write header
        writer.writerow(['#', 'User', 'Folder', 'Subject', 'Sent DateTime', 'Recipient/Sender'])

        # Write data
        for i, msg in enumerate(messages, 1):
            user = msg.get("_user_mailbox", "").split("@")[0]
            folder = msg.get("_folder", "")
            subject = msg.get("subject", "(No subject)")
            sent_time = msg.get("sentDateTime", "")
            contact = get_contact(msg)

            writer.writerow([i, user, folder, subject, sent_time, contact])

    print(f"Messages saved to {filename}")


# ============================================================================
# RUN THE SCRIPT
# ============================================================================

if __name__ == "__main__":
    print("="*60)
    print("Microsoft Graph API - Email Retrieval with Token Caching")
    print("="*60)
    print()

    # Verify required environment variables
    if not all([TENANT_ID, CLIENT_ID, CLIENT_SECRET]):
        print("ERROR: Missing required environment variables!")
        print("Please ensure your .env file contains:")
        print("  TENANT_ID=your_tenant_id")
        print("  CLIENT_ID=your_client_id")
        print("  CLIENT_SECRET=your_client_secret")
        exit(1)

    if not USERS:
        print("ERROR: No users configured in do.config")
        exit(1)

    print(f"Users: {', '.join(USERS)}")
    print(f"Folders: {', '.join(FOLDERS)}")
    print(f"Date range: {START_DATE} to {END_DATE}")
    print(f"Results per page: {TOP}")
    if FILTER_DOMAINS:
        print(f"Filtering by domains: {', '.join(FILTER_DOMAINS)}")
    print()

    try:
        # Get access token (will reuse if still valid)
        access_token = get_access_token()

        all_final_messages = []

        for user_email in USERS:
            for folder in FOLDERS:
                print(f"\n{'─'*60}")
                print(f"Searching: {user_email} / {folder}")
                print(f"{'─'*60}")

                messages = get_all_messages(access_token, user_email, folder)

                # Filter by domains if specified
                if FILTER_DOMAINS:
                    print(f"Filtering for domains: {', '.join(FILTER_DOMAINS)}")
                    messages = filter_by_domains(messages, FILTER_DOMAINS)
                    print(f"Messages matching domains: {len(messages)}")

                # Tag each message with the user mailbox and folder
                for msg in messages:
                    msg["_user_mailbox"] = user_email
                    msg["_folder"] = folder

                all_final_messages.extend(messages)

        print(f"\n{'='*60}")
        print(f"Total messages across all users/folders: {len(all_final_messages)}")
        print(f"{'='*60}")

        if all_final_messages:
            print()
            print_message_table(all_final_messages)
            save_to_json(all_final_messages)
            save_to_csv(all_final_messages)
        else:
            print("No messages found matching criteria")

        print("\n✓ Script completed successfully!")

    except Exception as e:
        print(f"\n✗ Script failed: {e}")
        exit(1)