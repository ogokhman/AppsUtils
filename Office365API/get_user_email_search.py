import requests
import json
import time
import os
import sys 
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    print("Warning: python-dotenv not installed. Using environment variables or defaults.")
from datetime import datetime, timedelta
import pytz

# ANSI color codes for terminal output
GREEN = '\033[92m'
RESET = '\033[0m'

# EST timezone
est = pytz.timezone('America/New_York')

# =====================================================================
# Configuration: Load environment variables from the .env file (if available)
# =====================================================================

# Read the credentials from the environment (loaded from .env if available)
TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
SCOPE = "https://graph.microsoft.com/.default"

def generate_access_token():
    """
    Generates a fresh Office 365 API access token using client credentials flow,
    or returns cached token if still valid (less than 1 hour old).
    
    @return: The access token string or None if generation fails.
    """
    # Check for existing token and timestamp in .env
    cached_token = os.getenv("ACCESS_TOKEN")
    token_timestamp_str = os.getenv("TOKEN_TIMESTAMP")
    
    if cached_token and token_timestamp_str:
        try:
            # Parse the timestamp and convert to EST
            token_timestamp = datetime.fromisoformat(token_timestamp_str)
            # Make timezone-aware if needed (assume UTC if no timezone info)
            if token_timestamp.tzinfo is None:
                token_timestamp = pytz.UTC.localize(token_timestamp)
            token_timestamp_est = token_timestamp.astimezone(est)
            token_timestamp_formatted = token_timestamp_est.strftime('%Y-%m-%d %H:%M:%S')
            
            current_time = datetime.now(pytz.UTC).astimezone(est)
            time_diff = current_time - token_timestamp_est
            
            # Check if token is still valid (less than 1 hour old)
            if time_diff < timedelta(hours=1):
                expires_in_seconds = int((token_timestamp_est + timedelta(hours=1) - current_time).total_seconds())
                hours = expires_in_seconds // 3600
                minutes = (expires_in_seconds % 3600) // 60
                seconds = expires_in_seconds % 60
                time_str = f"{hours:02d}:{minutes:02d}:{seconds:02d}"
                print(f"{GREEN}Saved Token: {token_timestamp_formatted} EST, is less than 1 hour old{RESET}")
                print(f"{GREEN}Using cached token. Expires in {time_str} (HH:MM:SS).{RESET}")
                return cached_token
            else:
                print(f"{GREEN}Saved Token: {token_timestamp_formatted} EST, is more than 1 hour old{RESET}")
                print(f"{GREEN}Cached token expired. Generating new token...{RESET}")
        except ValueError:
            print(f"{GREEN}Invalid timestamp format in .env. Generating new token...{RESET}")
    else:
        print(f"{GREEN}No saved token found in .env. Generating new token...{RESET}")
    
    # Generate new token
    token_url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    
    payload = {
        'client_id': CLIENT_ID,
        'scope': SCOPE,
        'client_secret': CLIENT_SECRET,
        'grant_type': 'client_credentials'
    }
    
    headers = {
        'Content-Type': 'application/x-www-form-urlencoded'
    }
    
    try:
        response = requests.post(token_url, data=payload, headers=headers)
        response.raise_for_status()
        
        token_data = response.json()
        access_token = token_data.get('access_token')
        
        if access_token:
            expires_in = token_data.get('expires_in', 3600)
            current_time_est = datetime.now(pytz.UTC).astimezone(est).strftime('%Y-%m-%d %H:%M:%S')
            print(f"{GREEN}New Token was generated at: {current_time_est} EST{RESET}")
            print(f"{GREEN}Token generated successfully. Expires in {expires_in} seconds.{RESET}")
            
            # Save token and timestamp to .env file
            update_env_file(access_token)
            
            return access_token
        else:
            print("Error: No access token in response")
            print("Response:", json.dumps(token_data, indent=2))
            return None
            
    except requests.exceptions.RequestException as e:
        print(f"Error generating token: {e}")
        if hasattr(e, 'response') and e.response is not None:
            try:
                error_details = e.response.json()
                print("Error details:", json.dumps(error_details, indent=2))
            except:
                print("Could not parse error response")
        return None

def update_env_file(access_token):
    """
    Updates the .env file with the new access token and current timestamp.
    
    @param access_token: The new access token to save.
    """
    env_file_path = os.path.join(os.path.dirname(__file__), '.env')
    current_time = datetime.now(pytz.UTC).astimezone(est).isoformat()
    
    try:
        # Read existing .env file
        env_lines = []
        if os.path.exists(env_file_path):
            with open(env_file_path, 'r') as f:
                env_lines = f.readlines()
        
        # Update or add ACCESS_TOKEN and TOKEN_TIMESTAMP
        token_updated = False
        timestamp_updated = False
        
        for i, line in enumerate(env_lines):
            if line.startswith('ACCESS_TOKEN='):
                env_lines[i] = f'ACCESS_TOKEN={access_token}\n'
                token_updated = True
            elif line.startswith('TOKEN_TIMESTAMP='):
                env_lines[i] = f'TOKEN_TIMESTAMP={current_time}\n'
                timestamp_updated = True
        
        # Add missing entries
        if not token_updated:
            env_lines.append(f'ACCESS_TOKEN={access_token}\n')
        if not timestamp_updated:
            env_lines.append(f'TOKEN_TIMESTAMP={current_time}\n')
        
        # Write back to .env file
        with open(env_file_path, 'w') as f:
            f.writelines(env_lines)
        
        print(f"{GREEN}Token and timestamp saved to .env file{RESET}")
        
    except Exception as e:
        print(f"Warning: Could not update .env file: {e}")
        print("Token will not be cached for next run.")

def get_folder_name(folder_id, access_token, target_user_identifier=None, suppress_warnings=False):
    """
    Get the folder name for a given folder ID.
    
    @param folder_id: The ID of the folder to look up.
    @param access_token: The access token for authentication.
    @param target_user_identifier: The user's email/ID for mailbox access.
    @param suppress_warnings: If True, suppress warning messages.
    @return: The folder name or the folder ID if not found.
    """
    if not folder_id:
        return "N/A"
    
    # Cache for folder names to avoid repeated API calls
    if not hasattr(get_folder_name, '_cache'):
        get_folder_name._cache = {}
    
    if folder_id in get_folder_name._cache:
        return get_folder_name._cache[folder_id]
    
    try:
        # Make API request to get folder information
        # Use the target user's mailbox if specified, otherwise use /me/
        if target_user_identifier and target_user_identifier != 'me':
            folder_url = f"https://graph.microsoft.com/v1.0/users/{target_user_identifier}/mailFolders/{folder_id}"
        else:
            folder_url = f"https://graph.microsoft.com/v1.0/me/mailFolders/{folder_id}"
        
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Accept": "application/json"
        }
        
        response = requests.get(folder_url, headers=headers)
        response.raise_for_status()
        
        folder_data = response.json()
        folder_name = folder_data.get('displayName', folder_id)
        
        # Cache the result
        get_folder_name._cache[folder_id] = folder_name
        return folder_name
        
    except Exception as e:
        if not suppress_warnings:
            print(f"Warning: Could not fetch folder name for {folder_id}: {e}")
        # Cache the folder ID as fallback
        get_folder_name._cache[folder_id] = folder_id
        return folder_id

def get_folder_id(folder_name, access_token, user_identifier, suppress_warnings=False):
    """
    Get the folder ID by name (reverse of get_folder_name)
    """
    # Check cache first
    cache_key = f"{user_identifier}_{folder_name}"
    if hasattr(get_folder_id, '_cache'):
        if cache_key in get_folder_id._cache:
            return get_folder_id._cache[cache_key]
    else:
        get_folder_id._cache = {}
    
    try:
        # Get all folders
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Accept": "application/json"
        }
        
        API_ENDPOINT = f"https://graph.microsoft.com/v1.0/users/{user_identifier}/mailFolders"
        response = requests.get(API_ENDPOINT, headers=headers)
        response.raise_for_status()
        
        folders_data = response.json()
        folders = folders_data.get('value', [])
        
        # Find the folder by name
        for folder in folders:
            if folder.get('displayName', '').lower() == folder_name.lower():
                folder_id = folder.get('id')
                get_folder_id._cache[cache_key] = folder_id
                return folder_id
        
        if not suppress_warnings:
            print(f"Folder '{folder_name}' not found")
        return None
        
    except Exception as e:
        if not suppress_warnings:
            print(f"Error getting folder ID for '{folder_name}': {e}")
        return None

def get_user_email(target_user_identifier, num_messages=None, earliestdate=None, maxdate=None, fromEmailAddress=None, folder=None, sort=None):
    """
    Sends an authenticated GET request to Microsoft Graph API
    to retrieve the specified user's email messages using SEARCH API.
    
    @param target_user_identifier: The email or ID of the user to fetch.
    @param num_messages: Number of messages to retrieve (default: 3).
    @param folder: Specific folder(s) to query (comma-separated, e.g., 'Inbox', 'SentItems,Drafts').
                  If not specified, defaults to 'Inbox' only.
    @param sort: Sort order - 'latest' for newest first, 'earliest' for oldest first.
    """
    # Generate fresh access token
    ACCESS_TOKEN = generate_access_token()
    
    if not ACCESS_TOKEN:
        print("Error: Failed to generate access token. Cannot proceed with API request.")
        return
    #print(ACCESS_TOKEN)
    # Conceptual API endpoint for fetching the user's profile.
    # The identifier provided via the command line is used to construct the API URL.
    # Replace 'https://api.example.com/v1.0/users/' with your real API base URL if needed.
    # Parse folder list - if no folder specified, default to Inbox only
    folders_to_query = []
    print(f"DEBUG: folder parameter = {folder}", flush=True)
    if folder is not None and folder.strip():
        # Split by comma and clean up
        folders_to_query = [f.strip() for f in folder.strip().split(',') if f.strip()]
        print(f"Querying folders: {folders_to_query}")
    else:
        folders_to_query = ['Inbox']  # Default to Inbox only
        print("No folder specified, querying Inbox only")
    
    print(f"DEBUG: folders_to_query = {folders_to_query}", flush=True)

    effective_top = num_messages
    # If the user did not specify a number of messages, don't send $top to get all messages (requires pagination)
    if num_messages is None:
        effective_top = None
    elif num_messages == 0:
        # When num_messages is 0, don't send $top argument to Microsoft Graph
        effective_top = None

    # Build search query instead of filter
    search_terms = []
    
    # Note: Graph Search API for messages doesn't support date ranges properly
    # We'll use search for from addresses only, and rely on $filter for dates
    
    # Add from addresses to search query
    from_email_addresses = []
    if fromEmailAddress and fromEmailAddress.strip():
        # Split by comma and clean up
        from_email_addresses = [addr.strip().replace("'", "''") 
                               for addr in fromEmailAddress.strip().split(',') 
                               if addr.strip()]
        print("FROM EMAIL ADDRESSES: ", from_email_addresses)
        
        # Add each from address to search query with OR
        from_terms = []
        for addr in from_email_addresses:
            # Support both email address and name search
            from_terms.append(f'from:"{addr}"')
        if from_terms:
            search_terms.append(f"({' OR '.join(from_terms)})")

    # Combine search terms (only from addresses, no dates)
    search_query = " AND ".join(search_terms) if search_terms else ""
    
    # Determine sort order
    sort_order = "receivedDateTime desc" if sort != "earliest" else "receivedDateTime asc"
    print(f"Sort order: {sort_order}", flush=True)
    
    # Set up the authorization header
    headers = {
        "Authorization": f"Bearer {ACCESS_TOKEN}",
        "Accept": "application/json"
    }

    all_messages = []
    
    # Query each folder separately
    for folder_name in folders_to_query:
        print(f"\nQuerying folder: {folder_name}")
        print("-" * 40)
        
        # Check if we can use search API for this specific folder
        folder_id = get_folder_id(folder_name, ACCESS_TOKEN, target_user_identifier, suppress_warnings=True)
        
        # Determine which API to use
        # Search API only works without date filters, so if dates are specified, use Filter API
        use_search_for_this_folder = bool(folder_id) and bool(search_query) and not earliestdate and not maxdate
        
        if use_search_for_this_folder:
            # Use search API approach for this folder (no date filters)
            print(f"Using Search API for {folder_name}")
            API_ENDPOINT = f"https://graph.microsoft.com/v1.0/users/{target_user_identifier}/messages"
            print(f"***API ENDPOINT: {API_ENDPOINT}", flush=True)
            
            # Prepare search parameters
            params = {
                "$select": "subject,from,toRecipients,receivedDateTime,parentFolderId",
                "$orderby": sort_order
            }
            
            # Add folder filter when using search
            params["$filter"] = f"parentFolderId eq '{folder_id}'"
            
            # Add search query (no date filters in search)
            params["$search"] = f'"{search_query}"'
            print(f"***SEARCH QUERY: {search_query}", flush=True)
            
            # Make single API call for search API
            try:
                # Make the GET request
                prepared = requests.Request("GET", API_ENDPOINT, headers=headers, params=params).prepare()
                print(f"***FULL URL: {prepared.url}", flush=True)
                
                response = requests.get(API_ENDPOINT, headers=headers, params=params)
                
                # Raise an exception for bad status codes (4xx or 5xx)
                response.raise_for_status()

                # Parse the JSON response
                user_data = response.json()
                messages = user_data.get('value', [])
                
                # Implement pagination if needed
                next_link = user_data.get('@odata.nextLink')
                page_count = 1
                
                while next_link and effective_top is None:
                    print(f"Fetching page {page_count + 1} (got {len(messages)} messages so far)...")
                    response = requests.get(next_link, headers=headers)
                    response.raise_for_status()
                    page_data = response.json()
                    messages.extend(page_data.get('value', []))
                    next_link = page_data.get('@odata.nextLink')
                    page_count += 1
                    
                    # Safety check to prevent infinite loops
                    if page_count > 100:
                        print("Warning: Reached maximum page limit (100). Stopping pagination.")
                        break

                print(f"Retrieved {len(messages)} messages from {folder_name}")
                all_messages.extend(messages)
                
            except requests.exceptions.HTTPError as e:
                if e.response.status_code == 404:
                    print(f"Folder '{folder_name}' not found or inaccessible. Skipping...")
                else:
                    print(f"HTTP Error occurred while querying {folder_name}: {e}")
            except Exception as e:
                print(f"Error occurred while querying {folder_name}: {e}")
        else:
            # Use filter API approach for this folder
            if not folder_id:
                print(f"Using Filter API for {folder_name} (could not get folder ID)")
            else:
                print(f"Using Filter API for {folder_name} (no From filter specified)")
            
            # Handle multiple From addresses with Filter API
            if from_email_addresses and len(from_email_addresses) > 1:
                print(f"Making {len(from_email_addresses)} separate API calls for each From address...")
                folder_messages = []
                
                for addr in from_email_addresses:
                    print(f"  Querying for: {addr}")
                    API_ENDPOINT = f"https://graph.microsoft.com/v1.0/users/{target_user_identifier}/mailFolders/{folder_name}/messages"
                    print(f"***API ENDPOINT: {API_ENDPOINT}", flush=True)
                    
                    # Prepare filter parameters
                    params = {
                        "$select": "subject,from,toRecipients,receivedDateTime,parentFolderId",
                        "$orderby": sort_order
                    }
                    
                    # Build filters
                    filters = []
                    
                    # Add date filters
                    if earliestdate:
                        filters.append(f"receivedDateTime ge {earliestdate}")
                    if maxdate:
                        filters.append(f"receivedDateTime le {maxdate}")
                    
                    # Add from address filter
                    addr_clean = addr.replace("'", "''")
                    filters.append(f"from/emailAddress/address eq '{addr_clean}'")
                    
                    if filters:
                        params["$filter"] = " and ".join(filters)
                    
                    # Add $top if specified
                    if effective_top is not None:
                        params["$top"] = effective_top
                    # Note: No automatic limit - let API handle pagination naturally
                    
                    try:
                        # Make the GET request
                        prepared = requests.Request("GET", API_ENDPOINT, headers=headers, params=params).prepare()
                        print(f"***FULL URL: {prepared.url}", flush=True)
                        
                        response = requests.get(API_ENDPOINT, headers=headers, params=params)
                        
                        # Raise an exception for bad status codes (4xx or 5xx)
                        response.raise_for_status()

                        # Parse the JSON response
                        user_data = response.json()
                        messages = user_data.get('value', [])
                        
                        # Implement pagination if needed
                        next_link = user_data.get('@odata.nextLink')
                        page_count = 1
                        
                        while next_link and effective_top is None:
                            print(f"Fetching page {page_count + 1} (got {len(messages)} messages so far)...")
                            response = requests.get(next_link, headers=headers)
                            response.raise_for_status()
                            page_data = response.json()
                            messages.extend(page_data.get('value', []))
                            next_link = page_data.get('@odata.nextLink')
                            page_count += 1
                            
                            # Safety check to prevent infinite loops
                            if page_count > 100:
                                print("Warning: Reached maximum page limit (100). Stopping pagination.")
                                break

                        print(f"  Retrieved {len(messages)} messages for {addr}")
                        folder_messages.extend(messages)
                        
                    except requests.exceptions.HTTPError as e:
                        if e.response.status_code == 404:
                            print(f"Folder '{folder_name}' not found or inaccessible. Skipping...")
                        else:
                            print(f"HTTP Error occurred while querying {folder_name}: {e}")
                    except Exception as e:
                        print(f"Error occurred while querying {folder_name}: {e}")
                
                print(f"Total messages from {folder_name}: {len(folder_messages)}")
                all_messages.extend(folder_messages)
            else:
                # Single From address or no From filter - use original logic
                API_ENDPOINT = f"https://graph.microsoft.com/v1.0/users/{target_user_identifier}/mailFolders/{folder_name}/messages"
                print(f"***API ENDPOINT: {API_ENDPOINT}", flush=True)
                
                # Prepare filter parameters
                params = {
                    "$select": "subject,from,toRecipients,receivedDateTime,parentFolderId",
                    "$orderby": sort_order
                }
                
                # Build all filters
                filters = []
                
                # Add date filters
                if earliestdate:
                    filters.append(f"receivedDateTime ge {earliestdate}")
                if maxdate:
                    filters.append(f"receivedDateTime le {maxdate}")
                
                # Add from address filter (if any)
                if from_email_addresses:
                    first_addr = from_email_addresses[0].replace("'", "''")
                    filters.append(f"from/emailAddress/address eq '{first_addr}'")
                
                if filters:
                    params["$filter"] = " and ".join(filters)
                
                # Add $top if specified
                if effective_top is not None:
                    params["$top"] = effective_top
                # Note: No automatic limit - let API handle pagination naturally
                
                try:
                    # Make the GET request
                    prepared = requests.Request("GET", API_ENDPOINT, headers=headers, params=params).prepare()
                    print(f"***FULL URL: {prepared.url}", flush=True)
                    
                    response = requests.get(API_ENDPOINT, headers=headers, params=params)
                    
                    # Raise an exception for bad status codes (4xx or 5xx)
                    response.raise_for_status()

                    # Parse the JSON response
                    user_data = response.json()
                    messages = user_data.get('value', [])
                    
                    # Implement pagination if needed
                    next_link = user_data.get('@odata.nextLink')
                    page_count = 1
                    
                    while next_link and effective_top is None:
                        print(f"Fetching page {page_count + 1} (got {len(messages)} messages so far)...")
                        response = requests.get(next_link, headers=headers)
                        response.raise_for_status()
                        page_data = response.json()
                        messages.extend(page_data.get('value', []))
                        next_link = page_data.get('@odata.nextLink')
                        page_count += 1
                        
                        # Safety check to prevent infinite loops
                        if page_count > 100:
                            print("Warning: Reached maximum page limit (100). Stopping pagination.")
                            break

                    print(f"Retrieved {len(messages)} messages from {folder_name}")
                    all_messages.extend(messages)
                    
                except requests.exceptions.HTTPError as e:
                    if e.response.status_code == 404:
                        print(f"Folder '{folder_name}' not found or inaccessible. Skipping...")
                    else:
                        print(f"HTTP Error occurred while querying {folder_name}: {e}")
                except Exception as e:
                    print(f"Error occurred while querying {folder_name}: {e}")
 
    # Display total messages retrieved from API
    print(f"\nTotal messages retrieved from API: {len(all_messages)}")
    messages = all_messages

    # No client-side filtering - all filtering done server-side
    print("Note: All filtering done server-side", flush=True)

    # Check if we have messages in the response
    if not messages:
        print("No messages found or unexpected response format.")
        return
    
    # Get all unique folder IDs for batch lookup
    unique_folder_ids = set()
    for msg in messages:
        folder_id = msg.get('parentFolderId')
        if folder_id:
            unique_folder_ids.add(folder_id)
    
    # Pre-fetch all folder names
    for folder_id in unique_folder_ids:
        get_folder_name(folder_id, ACCESS_TOKEN, target_user_identifier, suppress_warnings=True)
    
    # Display messages in table format
    print(f"\nFound {len(messages)} recent messages for '{target_user_identifier}':")
    print("-" * 160)
    print(f"{'#':<4} {'Date/Time':<30} {'From Address':<50} {'To Address':<30} {'Folder':<15} {'Subject':<40}")
    print("-" * 160)

    for idx, message in enumerate(messages, start=1):
        # Extract from information
        from_data = message.get('from', {})
        from_email = from_data.get('emailAddress', {}).get('address', 'N/A')
        from_name = from_data.get('emailAddress', {}).get('name', 'N/A')
        
        # Extract to information (handle multiple recipients)
        to_recipients = message.get('toRecipients', [])
        if to_recipients:
            to_email = to_recipients[0].get('emailAddress', {}).get('address', 'N/A')
            to_name = to_recipients[0].get('emailAddress', {}).get('name', 'N/A')
        else:
            to_email = 'N/A'
            to_name = 'N/A'
        
        # Extract subject
        subject = message.get('subject', '(No Subject)')
        
        # Extract folder information
        folder_id = message.get('parentFolderId')
        folder_name = get_folder_name(folder_id, ACCESS_TOKEN, target_user_identifier)

        received_dt_raw = message.get('receivedDateTime')
        try:
            received_dt = datetime.fromisoformat(received_dt_raw.replace('Z', '+00:00')) if received_dt_raw else None
            if received_dt is not None:
                if received_dt.tzinfo is None:
                    received_dt = pytz.UTC.localize(received_dt)
                dt_est = received_dt.astimezone(est)
                dt_str = dt_est.strftime('%Y-%m-%d %H:%M:%S')
            else:
                dt_str = 'N/A'
        except Exception as e:
            print(f"Could not parse datetime: {received_dt_raw} - Error: {e}")
            dt_str = 'N/A'
            
        # Truncate long fields for better table display (but keep From Address full)
        to_email = to_email[:27] + '...' if len(to_email) > 30 else to_email
        from_name = from_name[:17] + '...' if len(from_name) > 20 else from_name
        to_name = to_name[:17] + '...' if len(to_name) > 20 else to_name
        folder_name = folder_name[:12] + '...' if len(folder_name) > 15 else folder_name
        subject = subject[:37] + '...' if len(subject) > 40 else subject

        print(f"{idx:<4} {dt_str:<30} {from_email:<50} {to_email:<30} {folder_name:<15} {subject:<40}")

    print("-" * 160)


if __name__ == "__main__": 
    # Parse named command line arguments
    # Format: user=xxx@christoffersonrobb.com from=yyyy@zzz.com count=n mindate=yyyy-mm-dd maxdate=yyyy-mm-dd folder=Inbox,SentItems
    
    # Initialize defaults
    target_identifier = None
    earliestdate = None
    maxdate = None
    from_addr = None
    num_messages = None
    folder = None
    sort = 'latest'  # Default sort order
    auth_username = None
    auth_password = None
    
    print("Command line arguments:", sys.argv)
    
    # Parse arguments
    for arg in sys.argv[1:]:
        if "=" in arg:
            key, value = arg.split("=", 1)
            key = key.lower()
            value = value.strip()
            
            if key == "user":
                target_identifier = value
            elif key == "username":
                auth_username = value
            elif key == "password":
                auth_password = value
            elif key == "from":
                from_addr = value
            elif key == "count":
                try:
                    num_messages = int(value)
                except ValueError:
                    print(f"Invalid count value: {value}. Using default.")
                    num_messages = None
            elif key == "mindate":
                earliestdate = value
            elif key == "maxdate":
                maxdate = value
            elif key == "folder":
                folder = value
            elif key == "sort":
                if value.lower() in ['earliest', 'latest']:
                    sort = value.lower()
                else:
                    print(f"Invalid sort value: {value}. Using default (latest).")
                    sort = 'latest'
    
    # Validate required user parameter
    if not target_identifier:
        print("Usage: python get_user_email_search.py user=xxx@christoffersonrobb.com username=xxx password=xxx [from=yyyy@zzz.com] [count=n] [mindate=yyyy-mm-dd] [maxdate=yyyy-mm-dd] [folder=Inbox,SentItems] [sort=latest|earliest]")
        print("Examples:")
        print("  python get_user_email_search.py user=jane.doe@example.com username=oleg password=ThisIsMyLife")
        print("  python get_user_email_search.py user=jane.doe@example.com username=oleg password=ThisIsMyLife count=5")
        print("  python get_user_email_search.py user=jane.doe@example.com username=oleg password=ThisIsMyLife from=sender@contoso.com count=10")
        print("  python get_user_email_search.py user=jane.doe@example.com username=oleg password=ThisIsMyLife mindate=2025-01-01 maxdate=2025-01-31")
        sys.exit(1)
    
    # Validate authentication credentials
    if not auth_username or not auth_password:
        print("Error: Username and password are required for security.")
        print("Please provide username and password parameters.")
        sys.exit(1)
    
    # Check credentials (simple validation)
    expected_username = "oleg"
    expected_password = "Hello2026"
    
    if auth_username != expected_username or auth_password != expected_password:
        print("Error: Wrong credentials")
        sys.exit(1)
    
    print("TARGET IDENTIFIER: ", target_identifier)
    
    # Parse and convert dates if provided
    if earliestdate:
        try:
            # Allow 'YYYY-MM-DD', 'YYYY-MM-DD HH:MM', or 'YYYY-MM-DD HH:MM:SS'
            if " " not in earliestdate:
                v_to_parse = earliestdate + " 00:00:00"
            else:
                v_to_parse = earliestdate
            # Try parsing with seconds first, then without
            try:
                dt_naive = datetime.strptime(v_to_parse, "%Y-%m-%d %H:%M:%S")
            except ValueError:
                dt_naive = datetime.strptime(v_to_parse, "%Y-%m-%d %H:%M")
            dt_est = est.localize(dt_naive)
            dt_utc = dt_est.astimezone(pytz.UTC)
            earliestdate = dt_utc.strftime("%Y-%m-%dT%H:%M:%SZ")
        except ValueError:
            print("Invalid mindate format. Expected 'YYYY-MM-DD' or 'YYYY-MM-DD HH:MM'. Ignoring.")
            earliestdate = None
    
    if maxdate:
        try:
            # Allow 'YYYY-MM-DD', 'YYYY-MM-DD HH:MM', or 'YYYY-MM-DD HH:MM:SS'
            if " " not in maxdate:
                v_to_parse = maxdate + " 23:59:59"
            else:
                v_to_parse = maxdate
            # Try parsing with seconds first, then without
            try:
                dt_naive = datetime.strptime(v_to_parse, "%Y-%m-%d %H:%M:%S")
            except ValueError:
                dt_naive = datetime.strptime(v_to_parse, "%Y-%m-%d %H:%M")
            dt_est = est.localize(dt_naive)
            dt_utc = dt_est.astimezone(pytz.UTC)
            maxdate = dt_utc.strftime("%Y-%m-%dT%H:%M:%SZ")
        except ValueError:
            print("Invalid maxdate format. Expected 'YYYY-MM-DD' or 'YYYY-MM-DD HH:MM'. Ignoring.")
            maxdate = None
    
    # Validate num_messages bounds (if specified)
    if num_messages is not None:
        if num_messages < 0:
            num_messages = 3
        elif num_messages > 100:
            print("Maximum number of messages is 100. Using: 100")
            num_messages = 100
    # Default maxdate to "now" in UTC if not supplied
    if maxdate is None:
        now_utc = datetime.now(pytz.UTC)
        maxdate = now_utc.strftime("%Y-%m-%dT%H:%M:%SZ")
    
    # Note: When no dates are supplied, we only use maxdate (today)
    # This allows retrieving all messages up to today, not just from today

    if num_messages == 0:
        print(f"Retrieving all recent messages for '{target_identifier}'...")
    elif num_messages is None:
        print(f"Retrieving recent messages for '{target_identifier}'...")
    else:
        print(f"Retrieving {num_messages} recent messages for '{target_identifier}'...")

    with open('debug.txt', 'w') as f:
        f.write(f"GET USER EMAIL PARAMS: target={target_identifier}, num={num_messages}, earliest={earliestdate}, latest={maxdate}, from={from_addr}, folder={folder}\n")
    
    get_user_email(target_identifier, num_messages, earliestdate, maxdate, fromEmailAddress=from_addr, folder=folder, sort=sort)
 