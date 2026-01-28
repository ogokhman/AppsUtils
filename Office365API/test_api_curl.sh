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
    
    echo -e "${YELLOW}3. Get emails from SentItems folder:${NC}"
    echo "curl -X GET 'https://graph.microsoft.com/v1.0/users/USER_EMAIL/mailFolders/SentItems/messages?\$select=subject,from,toRecipients,receivedDateTime,parentFolderId&\$orderby=receivedDateTime desc&\$top=10' \\"
    echo "  -H 'Authorization: Bearer YOUR_ACCESS_TOKEN' \\"
    echo "  -H 'Accept: application/json'"
    echo ""
}

# Function to test actual API request
test_request() {
    read -p "Enter user email address: " USER_EMAIL
    read -p "Enter number of messages to retrieve (default: 10): " NUM_MESSAGES
    NUM_MESSAGES=${NUM_MESSAGES:-10}
    read -p "Enter folder name (default: Inbox): " FOLDER
    FOLDER=${FOLDER:-Inbox}
    read -p "Filter by sender email (optional, press Enter to skip): " SENDER_FILTER
    
    echo -e "\n${YELLOW}Making API request...${NC}"
    
    API_URL="https://graph.microsoft.com/v1.0/users/$USER_EMAIL/mailFolders/$FOLDER/messages?\$select=subject,from,receivedDateTime&\$orderby=receivedDateTime%20desc&\$top=$NUM_MESSAGES"
    
    echo -e "${BLUE}Request URL:${NC}"
    echo "$API_URL"
    echo ""
    
    # Debug: Check token
    if [ -z "$ACCESS_TOKEN" ]; then
        echo -e "${RED}Error: ACCESS_TOKEN is empty!${NC}"
        return 1
    fi
    echo -e "${YELLOW}Debug: Token present (${#ACCESS_TOKEN} chars)${NC}"
    
    RESPONSE=$(curl -s -X GET "$API_URL" \
        -H "Authorization: Bearer $ACCESS_TOKEN" \
        -H "Accept: application/json")
    
    # Check if response is empty
    if [ -z "$RESPONSE" ]; then
        echo -e "${RED}Error: Empty response from API${NC}"
        return 1
    fi
    
    # Check if response contains error
    if echo "$RESPONSE" | grep -q '"error"'; then
        echo -e "${RED}Error in API response:${NC}"
        echo "$RESPONSE" | python3 -m json.tool 2>/dev/null || echo "$RESPONSE"
        return 1
    fi
    
    echo -e "${GREEN}✓ Request successful${NC}\n"
    
    # Parse and display messages with Python
    echo "$RESPONSE" | python3 -c "
import json
import sys
from datetime import datetime

try:
    data = json.loads(sys.stdin.read())
    messages = data.get('value', [])
    
    # Filter by sender if specified
    sender_filter = '$SENDER_FILTER'.strip().lower()
    if sender_filter:
        print(f'Filtering by sender: {sender_filter}')
        filtered = []
        for msg in messages:
            from_data = msg.get('from', {})
            email_info = from_data.get('emailAddress', {})
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
    print(f\"{'#':<4} {'Date/Time':<25} {'From':<40} {'Subject':<60}\")
    print('-' * 130)
    
    for idx, msg in enumerate(messages, 1):
        # Extract from information
        from_data = msg.get('from', {})
        from_email = from_data.get('emailAddress', {}).get('address', 'N/A')
        from_name = from_data.get('emailAddress', {}).get('name', '')
        from_display = f'{from_name} <{from_email}>' if from_name else from_email
        from_display = from_display[:37] + '...' if len(from_display) > 40 else from_display
        
        # Extract subject
        subject = msg.get('subject', '(No Subject)')
        subject = subject[:57] + '...' if len(subject) > 60 else subject
        
        # Format date
        received = msg.get('receivedDateTime', '')
        try:
            dt = datetime.fromisoformat(received.replace('Z', '+00:00'))
            date_str = dt.strftime('%Y-%m-%d %H:%M:%S')
        except:
            date_str = 'N/A'
        
        print(f'{idx:<4} {date_str:<25} {from_display:<40} {subject:<60}')
    
except Exception as e:
    print(f'Error: {e}', file=sys.stderr)
    import traceback
    traceback.print_exc()
    sys.exit(1)
"
    
    if [ $? -eq 0 ]; then
        echo ""
        read -p "Show full JSON response? (y/n): " SHOW_FULL
        if [ "$SHOW_FULL" = "y" ]; then
            echo "$RESPONSE" | python3 -m json.tool | less
        fi
    fi
}

# Main menu
while true; do
    echo -e "\n${BLUE}=== Menu ===${NC}"
    echo "1. Get access token"
    echo "2. Show curl command examples"
    echo "3. Test API request (interactive)"
    echo "4. Exit"
    echo ""
    read -p "Select option (1-4): " OPTION
    
    case $OPTION in
        1)
            get_token
            ;;
        2)
            show_examples
            ;;
        3)
            if [ -z "$ACCESS_TOKEN" ]; then
                echo -e "${YELLOW}No access token found. Getting token first...${NC}\n"
                get_token
            fi
            test_request
            ;;
        4)
            echo -e "${GREEN}Goodbye!${NC}"
            exit 0
            ;;
        *)
            echo -e "${RED}Invalid option${NC}"
            ;;
    esac
done
