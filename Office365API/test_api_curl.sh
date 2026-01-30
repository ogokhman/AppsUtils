#!/bin/bash

# Script to test Office 365 API requests using curl
# Reads credentials from .env file

# Colors for output
GREEN='\033[0;32m'
BLUE='\033[0;34m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
NC='\033[0m' # No Color

# Load environment variables from .env file
if [ -f .env ]; then
    export $(cat .env | grep -v '^#' | xargs)
else
    echo -e "${RED}Error: .env file not found${NC}"
    exit 1
fi

# Check if required variables are set
if [ -z "$TENANT_ID" ] || [ -z "$CLIENT_ID" ] || [ -z "$CLIENT_SECRET" ]; then
    echo -e "${RED}Error: Missing required environment variables (TENANT_ID, CLIENT_ID, CLIENT_SECRET)${NC}"
    exit 1
fi

echo -e "${BLUE}=== Office 365 API Testing Tool ===${NC}\n"

# Global variable for access token
ACCESS_TOKEN=""

# Check for debug flag
DEBUG_MODE=0
for arg in "$@"; do
    if [ "$arg" = "--debug=1" ]; then
        DEBUG_MODE=1
    fi
done

# Function to get access token
get_token() {
    echo -e "${YELLOW}Getting access token...${NC}"
    
    TOKEN_RESPONSE=$(curl -s -X POST "https://login.microsoftonline.com/$TENANT_ID/oauth2/v2.0/token" \
        -H "Content-Type: application/x-www-form-urlencoded" \
        -d "client_id=$CLIENT_ID&scope=https://graph.microsoft.com/.default&client_secret=$CLIENT_SECRET&grant_type=client_credentials")
    
    ACCESS_TOKEN=$(echo $TOKEN_RESPONSE | grep -o '"access_token":"[^"]*' | cut -d'"' -f4)
    
    if [ -z "$ACCESS_TOKEN" ]; then
        echo -e "${RED}Error: Failed to get access token${NC}"
        echo "Response: $TOKEN_RESPONSE"
        exit 1
    fi
    
    echo -e "${GREEN}✓ Access token obtained successfully${NC}\n"
    echo -e "${BLUE}Token (first 50 chars):${NC} ${ACCESS_TOKEN:0:50}...\n"
}

# Function to show curl command examples
show_examples() {
    echo -e "${BLUE}=== Example curl commands ===${NC}\n"
    
    echo -e "${YELLOW}1. Get 10 latest emails from Inbox:${NC}"
    echo "curl -X GET 'https://graph.microsoft.com/v1.0/users/USER_EMAIL/mailFolders/Inbox/messages?\$select=subject,from,toRecipients,receivedDateTime,parentFolderId&\$orderby=receivedDateTime desc&\$top=10' \\"
    echo "  -H 'Authorization: Bearer YOUR_ACCESS_TOKEN' \\"
    echo "  -H 'Accept: application/json'"
    echo ""
    
    echo -e "${YELLOW}2. Get emails with date filter:${NC}"
    echo "curl -X GET 'https://graph.microsoft.com/v1.0/users/USER_EMAIL/mailFolders/Inbox/messages?\$select=subject,from,toRecipients,receivedDateTime,parentFolderId&\$orderby=receivedDateTime desc&\$top=10&\$filter=receivedDateTime ge 2026-01-01T00:00:00Z and receivedDateTime le 2026-01-27T23:59:59Z' \\"
    echo "  -H 'Authorization: Bearer YOUR_ACCESS_TOKEN' \\"
    echo "  -H 'Accept: application/json'"
    echo ""
    
    echo -e "${YELLOW}3. Get emails from Sent Items folder:${NC}"
    echo "curl -X GET 'https://graph.microsoft.com/v1.0/users/USER_EMAIL/mailFolders/Sent%20Items/messages?\$select=subject,from,toRecipients,receivedDateTime,parentFolderId&\$orderby=receivedDateTime desc&\$top=10' \\"
    echo "  -H 'Authorization: Bearer YOUR_ACCESS_TOKEN' \\"
    echo "  -H 'Accept: application/json'"
    echo ""
}

# Function to test actual API request
test_request() {
    # Predefined users
    USERS=("oleg@christoffersonrobb.com" "apapritz@christoffersonrobb.com" "malik@christoffersonrobb.com")
    MAILBOXES=("Inbox" "Archive" "Sent Items")
    
    # Mapping of display names to API folder names
    declare -A FOLDER_MAP
    FOLDER_MAP["Inbox"]="Inbox"
    FOLDER_MAP["Archive"]="Archive"
    FOLDER_MAP["Sent Items"]="SentItems"
    
    # Ask user to select from predefined list
    echo -e "\n${BLUE}=== Select Users ===${NC}"
    for i in "${!USERS[@]}"; do
        echo "$((i+1)). ${USERS[$i]}"
    done
    echo "$((${#USERS[@]}+1)). ALL users"
    echo ""
    read -p "Select users (comma or space-separated, e.g., 1,3 or 1 3): " USER_CHOICE
    
    SELECTED_USERS=()
    
    # Check if user wants ALL users (just the number)
    if [[ "$USER_CHOICE" == "$((${#USERS[@]}+1))" ]]; then
        SELECTED_USERS=("${USERS[@]}")
    else
        # Replace commas with spaces and process each number
        USER_CHOICE=${USER_CHOICE//,/ }
        for choice in $USER_CHOICE; do
            # Trim whitespace
            choice=$(echo "$choice" | sed 's/^[[:space:]]*//;s/[[:space:]]*$//')
            # Check if it's a valid number
            if [[ "$choice" =~ ^[0-9]+$ ]] && [ "$choice" -ge 1 ] && [ "$choice" -le "${#USERS[@]}" ]; then
                SELECTED_USERS+=("${USERS[$((choice-1))]}")
            fi
        done
        
        if [ ${#SELECTED_USERS[@]} -eq 0 ]; then
            echo -e "${RED}Invalid selection${NC}"
            return 1
        fi
    fi
    
    # Ask user to select mailboxes
    echo -e "\n${BLUE}=== Select Mailboxes ===${NC}"
    for i in "${!MAILBOXES[@]}"; do
        echo "$((i+1)). ${MAILBOXES[$i]}"
    done
    echo "$((${#MAILBOXES[@]}+1)). ALL mailboxes"
    echo ""
    read -p "Select mailboxes (comma or space-separated, e.g., 1,3 or 1 3): " MAILBOX_CHOICE
    
    SELECTED_MAILBOXES=()
    
    # Check if user wants ALL mailboxes (just the number)
    if [[ "$MAILBOX_CHOICE" == "$((${#MAILBOXES[@]}+1))" ]]; then
        SELECTED_MAILBOXES=("${MAILBOXES[@]}")
    else
        # Replace commas with spaces and process each number
        MAILBOX_CHOICE=${MAILBOX_CHOICE//,/ }
        for choice in $MAILBOX_CHOICE; do
            # Trim whitespace
            choice=$(echo "$choice" | sed 's/^[[:space:]]*//;s/[[:space:]]*$//')
            # Check if it's a valid number
            if [[ "$choice" =~ ^[0-9]+$ ]] && [ "$choice" -ge 1 ] && [ "$choice" -le "${#MAILBOXES[@]}" ]; then
                SELECTED_MAILBOXES+=("${MAILBOXES[$((choice-1))]}")
            fi
        done
        
        if [ ${#SELECTED_MAILBOXES[@]} -eq 0 ]; then
            echo -e "${RED}Invalid selection${NC}"
            return 1
        fi
    fi
    
    read -p "Enter number of messages to retrieve (default: 10, max: 500): " NUM_MESSAGES
    NUM_MESSAGES=${NUM_MESSAGES:-10}
    
    # Validate NUM_MESSAGES doesn't exceed 500
    if [ "$NUM_MESSAGES" -gt 500 ]; then
        echo -e "${YELLOW}Warning: Maximum is 500 messages. Setting to 500.${NC}"
        NUM_MESSAGES=500
    fi
    
    # Calculate default dates (today and today - 3 days)
    END_DATE=$(date +%Y-%m-%d)
    START_DATE=$(date -d "3 days ago" +%Y-%m-%d)
    
    read -p "Enter start date (default: $START_DATE, format: YYYY-MM-DD): " USER_START_DATE
    START_DATE=${USER_START_DATE:-$START_DATE}
    
    read -p "Enter end date (default: $END_DATE, format: YYYY-MM-DD): " USER_END_DATE
    END_DATE=${USER_END_DATE:-$END_DATE}
    
    read -p "Filter by sender email (optional, press Enter to skip): " SENDER_FILTER
    
    echo -e "\n${YELLOW}Making API request...${NC}"
    
    # Process each selected user
    for USER_EMAIL in "${SELECTED_USERS[@]}"; do
        echo -e "\n${BLUE}=== Processing: $USER_EMAIL ===${NC}"
        echo "Selected folders: ${SELECTED_MAILBOXES[*]}"
        
        # Initialize response collection
        RESPONSE_COUNT=0
        
        # Collect all responses
        RESPONSE_COUNT=0
        
        for FOLDER in "${SELECTED_MAILBOXES[@]}"; do
            echo -e "${YELLOW}Fetching from $FOLDER...${NC}"
            
            # Get the actual API folder name from the mapping
            API_FOLDER="${FOLDER_MAP[$FOLDER]}"
            
            # Build URL using $search (supports more flexible queries, but no orderby)
            # Build search query based on filters
            SEARCH_QUERY=""
            
            if [ -n "$SENDER_FILTER" ]; then
                SEARCH_QUERY="from:${SENDER_FILTER}"
            fi
            
            # Add date range to search query
            SEARCH_QUERY="received>=${START_DATE} AND received<=${END_DATE}"
            if [ -n "$SENDER_FILTER" ]; then
                SEARCH_QUERY="from:${SENDER_FILTER} AND ${SEARCH_QUERY}"
            fi
            
            SEARCH_ENCODED="${SEARCH_QUERY// /%20}"
            SEARCH_ENCODED="${SEARCH_ENCODED//\"/%22}"
            SEARCH_ENCODED="${SEARCH_ENCODED//>/%3E}"
            SEARCH_ENCODED="${SEARCH_ENCODED//=/%3D}"
            SEARCH_ENCODED="${SEARCH_ENCODED//</%3C}"
            
            API_URL="https://graph.microsoft.com/v1.0/users/$USER_EMAIL/mailFolders/$API_FOLDER/messages?\$select=subject,from,toRecipients,receivedDateTime&\$top=$NUM_MESSAGES&\$search=\"${SEARCH_ENCODED}\""
            
            # Debug: Check token
            if [ -z "$ACCESS_TOKEN" ]; then
                echo -e "${RED}Error: ACCESS_TOKEN is empty!${NC}"
                continue
            fi
            
            # Save URL and token for debug_api.sh
            echo "$API_URL" > /tmp/last_api_url.txt
            echo "$ACCESS_TOKEN" > /tmp/last_access_token.txt
            
            # Show URL and full curl command for testing
            echo -e "${BLUE}=== API URL ===${NC}"
            echo "$API_URL"
            echo ""
            echo -e "${BLUE}=== Full curl command (copy & paste to test) ===${NC}"
            echo "curl -X GET '$API_URL' \\"
            echo "  -H 'Authorization: Bearer $ACCESS_TOKEN' \\"
            echo "  -H 'Accept: application/json'"
            echo ""
            echo -e "${YELLOW}Note: The access token is included in the command above${NC}"
            echo -e "${YELLOW}URL and token saved to /tmp for use with debug_api.sh${NC}"
            echo ""
            
            # Use temp file to separate verbose output from response
            TEMP_RESPONSE="/tmp/curl_response_$$.txt"
            HTTP_CODE=$(curl -s -w "%{http_code}" -X GET "$API_URL" \
                -H "Authorization: Bearer $ACCESS_TOKEN" \
                -H "Accept: application/json" \
                -o "$TEMP_RESPONSE")
            
            RESPONSE_BODY=$(cat "$TEMP_RESPONSE")
            rm -f "$TEMP_RESPONSE"
            
            # Debug output for troubleshooting
            echo -e "${YELLOW}HTTP Status: $HTTP_CODE${NC}"
            
            if [ "$DEBUG_MODE" -ge 2 ]; then
                echo -e "${BLUE}Response Body Length: ${#RESPONSE_BODY}${NC}"
                echo -e "${BLUE}Response (first 500 chars):${NC}"
                echo "$RESPONSE_BODY" | head -c 500
                echo ""
                # Also save to file for inspection
                echo "$RESPONSE_BODY" > /tmp/api_response_$FOLDER.json
                echo "Response saved to /tmp/api_response_$FOLDER.json"
            fi
            
            # Check HTTP status
            if [ "$HTTP_CODE" != "200" ]; then
                echo -e "${RED}Error: HTTP $HTTP_CODE${NC}"
                echo "$RESPONSE_BODY"
                continue
            fi
            
            # Check if response is empty
            if [ -z "$RESPONSE_BODY" ]; then
                echo -e "${YELLOW}Empty response from API for $FOLDER${NC}"
                continue
            fi
            
            # Check if response contains error
            if echo "$RESPONSE_BODY" | grep -q '"error"'; then
                echo -e "${RED}Error in API response for $FOLDER:${NC}"
                echo "$RESPONSE_BODY" | python3 -m json.tool 2>/dev/null || echo "$RESPONSE_BODY"
                continue
            fi
            
            # Tag each message with folder name and store response
            echo "$RESPONSE_BODY" | python3 -c "
import json
import sys
try:
    data = json.loads(sys.stdin.read())
    messages = data.get('value', [])
    for msg in messages:
        msg['_folderName'] = '$FOLDER'
    with open('/tmp/response_$RESPONSE_COUNT.json', 'w') as f:
        json.dump({'value': messages}, f)
except json.JSONDecodeError as e:
    print(f'Error parsing JSON: {e}', file=sys.stderr)
    with open('/tmp/response_$RESPONSE_COUNT.json', 'w') as f:
        json.dump({'value': []}, f)
"
            RESPONSE_COUNT=$((RESPONSE_COUNT + 1))
        done
        
        # Merge all responses
        ALL_MESSAGES=$(python3 -c "
import json
import sys
import os

all_messages = []
for i in range($RESPONSE_COUNT):
    filename = f'/tmp/response_{i}.json'
    if os.path.exists(filename):
        with open(filename, 'r') as f:
            try:
                data = json.load(f)
                messages = data.get('value', [])
                all_messages.extend(messages)
            except:
                pass
        os.remove(filename)  # Clean up

# Sort by receivedDateTime desc
all_messages.sort(key=lambda x: x.get('receivedDateTime', ''), reverse=True)

# Limit to NUM_MESSAGES total
all_messages = all_messages[:$NUM_MESSAGES]

result = {'value': all_messages}
print(json.dumps(result))
")
        
        # Now display all collected messages
        echo "$ALL_MESSAGES" | python3 -c "
import json
import sys
import re
from datetime import datetime

user_email = '$USER_EMAIL'
user_name = user_email.split('@')[0] if '@' in user_email else user_email

try:
    data = json.loads(sys.stdin.read())
    messages = data.get('value', [])
    
    print(f'Total messages fetched: {len(messages)}')

    sender_filter = '$SENDER_FILTER'.strip().lower()
    if sender_filter:
        print(f'Filtering by sender: {sender_filter}')
        filtered = []
        for msg in messages:
            email_info = (msg.get('from') or {}).get('emailAddress') or {}
            from_addr = (email_info.get('address') or '').lower()
            from_name = (email_info.get('name') or '').lower()
            if sender_filter in from_addr or sender_filter in from_name:
                filtered.append(msg)
        messages = filtered
        print(f'Found {len(messages)} messages from sender\n')

    if not messages:
        print('No messages found.')
        sys.exit(0)

    print(f'Found {len(messages)} messages:\n')
    print(f\"{'#':<4} {'User':<12} {'Folder':<15} {'Date/Time':<25} {'From/To':<40} {'Subject':<40}\")
    print('-' * 148)

    for idx, msg in enumerate(messages, 1):
        email_info = (msg.get('from') or {}).get('emailAddress') or {}
        from_email = email_info.get('address') or 'N/A'
        from_name = email_info.get('name') or ''
        
        # Get the actual folder name from the message
        folder_name = msg.get('_folderName', 'Unknown')

        from_name = re.sub(r'<\/[^>]*>', '', from_name).strip()
        from_name = from_name.replace('<', '').replace('>', '')

        # For Sent Items, display recipients instead of sender
        if folder_name == 'Sent Items':
            to_recipients = msg.get('toRecipients', [])
            if to_recipients:
                to_list = []
                for recipient in to_recipients:
                    email_info = recipient.get('emailAddress') or {}
                    to_email = email_info.get('address') or 'N/A'
                    to_name = email_info.get('name') or ''
                    if to_name and to_email != 'N/A':
                        to_list.append(f'{to_name} <{to_email}>')
                    elif to_name:
                        to_list.append(to_name)
                    else:
                        to_list.append(to_email)
                from_display = '; '.join(to_list[:2])  # Limit to 2 recipients
                if len(to_list) > 2:
                    from_display += f'; +{len(to_list) - 2} more'
            else:
                from_display = 'N/A'
        else:
            if from_name and from_email != 'N/A':
                from_display = f'{from_name} <{from_email}>'
            elif from_name:
                from_display = from_name
            else:
                from_display = from_email

        subject = msg.get('subject') or '(No Subject)'
        subject = subject[:37] + '...' if len(subject) > 40 else subject

        received = msg.get('receivedDateTime') or ''
        try:
            dt = datetime.fromisoformat(received.replace('Z', '+00:00'))
            date_str = dt.strftime('%Y-%m-%d %H:%M:%S')
        except:
            date_str = 'N/A'
        
        # Get the actual folder name from the message
        folder_name = msg.get('_folderName', 'Unknown')

        print(f'{idx:<4} {user_name:<12} {folder_name:<15} {date_str:<25} {from_display:<40} {subject:<40}')

except Exception as e:
    print(f'Error: {e}', file=sys.stderr)
    import traceback
    traceback.print_exc()
    sys.exit(1)
"
        
        # Show raw JSON only in debug mode
        if [ "$DEBUG_MODE" -eq 1 ]; then
            echo ""
            echo -e "${BLUE}=== RAW JSON RESPONSE ===${NC}"
            echo "$ALL_MESSAGES" | python3 -m json.tool 2>/dev/null || echo "$ALL_MESSAGES"
        fi
    done
}

# Main menu
while true; do
    echo -e "\n${BLUE}=== Menu ===${NC}"
    echo "1. Office 365 API Request"
    echo "2. Get Access Token (optional)"
    echo "3. Show Curl Command Examples"
    echo "Q. Exit"
    echo ""
    read -p "Select option (1, 2, 3, or Q): " OPTION
    
    case $OPTION in
        1)
            if [ -z "$ACCESS_TOKEN" ]; then
                echo -e "${YELLOW}No access token found. Getting token first...${NC}\n"
                get_token
            fi
            test_request
            ;;
        2)
            get_token
            ;;
        3)
            show_examples
            ;;
        Q|q)
            echo -e "${GREEN}Goodbye!${NC}"
            exit 0
            ;;
        *)
            echo -e "${RED}Invalid option${NC}"
            ;;
    esac
done
