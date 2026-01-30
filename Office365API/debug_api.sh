#!/bin/bash

# Quick debug script to see raw API response

# Load .env
if [ -f .env ]; then
    export $(cat .env | grep -v '^#' | xargs)
fi

# Get token
echo "Getting token..."
TOKEN_RESPONSE=$(curl -s -X POST "https://login.microsoftonline.com/$TENANT_ID/oauth2/v2.0/token" \
    -H "Content-Type: application/x-www-form-urlencoded" \
    -d "client_id=$CLIENT_ID&scope=https://graph.microsoft.com/.default&client_secret=$CLIENT_SECRET&grant_type=client_credentials")

ACCESS_TOKEN=$(echo $TOKEN_RESPONSE | grep -o '"access_token":"[^"]*' | cut -d'"' -f4)

if [ -z "$ACCESS_TOKEN" ]; then
    echo "Failed to get token"
    echo "Response: $TOKEN_RESPONSE"
    exit 1
fi

echo "Token obtained: ${ACCESS_TOKEN:0:50}..."
echo ""

# Check if we have a saved URL from test_api_curl.sh
if [ -f /tmp/last_api_url.txt ] && [ -f /tmp/last_access_token.txt ]; then
    echo "Found saved URL from test_api_curl.sh"
    read -p "Use saved URL? (y/n, default=y): " USE_SAVED
    USE_SAVED=${USE_SAVED:-y}
    
    if [[ "$USE_SAVED" == "y" || "$USE_SAVED" == "Y" ]]; then
        API_URL=$(cat /tmp/last_api_url.txt)
        ACCESS_TOKEN=$(cat /tmp/last_access_token.txt)
        echo "Using saved URL and token"
        echo ""
        echo "Making request..."
        echo "URL: $API_URL"
        echo ""
    else
        # Manual mode - build URL from scratch
        # Get user email
        read -p "Enter user email: " USER_EMAIL
        
        # Get folder
        echo ""
        echo "Select folder:"
        echo "1. Inbox"
        echo "2. SentItems"
        echo "3. Archive"
        read -p "Choose (1-3, default=1): " FOLDER_CHOICE
        FOLDER_CHOICE=${FOLDER_CHOICE:-1}
        
        case $FOLDER_CHOICE in
            1) FOLDER="Inbox" ;;
            2) FOLDER="SentItems" ;;
            3) FOLDER="Archive" ;;
            *) FOLDER="Inbox" ;;
        esac
        
        # Get number of messages
        read -p "Number of messages (default=5): " NUM_MESSAGES
        NUM_MESSAGES=${NUM_MESSAGES:-5}
        
        # Get dates
        END_DATE=$(date +%Y-%m-%d)
        START_DATE=$(date -d "3 days ago" +%Y-%m-%d)
        read -p "Start date (default=$START_DATE, format YYYY-MM-DD): " USER_START_DATE
        START_DATE=${USER_START_DATE:-$START_DATE}
        read -p "End date (default=$END_DATE, format YYYY-MM-DD): " USER_END_DATE
        END_DATE=${USER_END_DATE:-$END_DATE}
        
        # Get optional sender filter
        read -p "Filter by sender (optional, press Enter to skip): " SENDER_FILTER
        
        # Build search query
        SEARCH_QUERY="received>=${START_DATE} AND received<=${END_DATE}"
        if [ -n "$SENDER_FILTER" ]; then
            SEARCH_QUERY="from:${SENDER_FILTER} AND ${SEARCH_QUERY}"
        fi
        
        # URL encode the search query
        SEARCH_ENCODED="${SEARCH_QUERY// /%20}"
        SEARCH_ENCODED="${SEARCH_ENCODED//\"/%22}"
        SEARCH_ENCODED="${SEARCH_ENCODED//>/%3E}"
        SEARCH_ENCODED="${SEARCH_ENCODED//=/%3D}"
        SEARCH_ENCODED="${SEARCH_ENCODED//</%3C}"
        
        # Make API request
        echo ""
        echo "Making request..."
        API_URL="https://graph.microsoft.com/v1.0/users/$USER_EMAIL/mailFolders/$FOLDER/messages?\$select=subject,from,toRecipients,receivedDateTime&\$top=$NUM_MESSAGES&\$search=\"${SEARCH_ENCODED}\""
        echo "URL: $API_URL"
        echo ""
    fi
else
    # No saved URL - manual mode
    # Get user email
    read -p "Enter user email: " USER_EMAIL
    
    # Get folder
    echo ""
    echo "Select folder:"
    echo "1. Inbox"
    echo "2. SentItems"
    echo "3. Archive"
    read -p "Choose (1-3, default=1): " FOLDER_CHOICE
    FOLDER_CHOICE=${FOLDER_CHOICE:-1}
    
    case $FOLDER_CHOICE in
        1) FOLDER="Inbox" ;;
        2) FOLDER="SentItems" ;;
        3) FOLDER="Archive" ;;
        *) FOLDER="Inbox" ;;
    esac
    
    # Get number of messages
    read -p "Number of messages (default=5): " NUM_MESSAGES
    NUM_MESSAGES=${NUM_MESSAGES:-5}
    
    # Get dates
    END_DATE=$(date +%Y-%m-%d)
    START_DATE=$(date -d "3 days ago" +%Y-%m-%d)
    read -p "Start date (default=$START_DATE, format YYYY-MM-DD): " USER_START_DATE
    START_DATE=${USER_START_DATE:-$START_DATE}
    read -p "End date (default=$END_DATE, format YYYY-MM-DD): " USER_END_DATE
    END_DATE=${USER_END_DATE:-$END_DATE}
    
    # Get optional sender filter
    read -p "Filter by sender (optional, press Enter to skip): " SENDER_FILTER
    
    # Build search query
    SEARCH_QUERY="received>=${START_DATE} AND received<=${END_DATE}"
    if [ -n "$SENDER_FILTER" ]; then
        SEARCH_QUERY="from:${SENDER_FILTER} AND ${SEARCH_QUERY}"
    fi
    
    # URL encode the search query
    SEARCH_ENCODED="${SEARCH_QUERY// /%20}"
    SEARCH_ENCODED="${SEARCH_ENCODED//\"/%22}"
    SEARCH_ENCODED="${SEARCH_ENCODED//>/%3E}"
    SEARCH_ENCODED="${SEARCH_ENCODED//=/%3D}"
    SEARCH_ENCODED="${SEARCH_ENCODED//</%3C}"
    
    # Make API request
    echo ""
    echo "Making request..."
    API_URL="https://graph.microsoft.com/v1.0/users/$USER_EMAIL/mailFolders/$FOLDER/messages?\$select=subject,from,toRecipients,receivedDateTime&\$top=$NUM_MESSAGES&\$search=\"${SEARCH_ENCODED}\""
    echo "URL: $API_URL"
    echo ""
fi

RESPONSE=$(curl -s -X GET "$API_URL" \
    -H "Authorization: Bearer $ACCESS_TOKEN" \
    -H "Accept: application/json")

echo "=== RAW RESPONSE ==="
echo "$RESPONSE"
echo ""
echo "=== END RESPONSE ==="
echo ""

# Try to format as JSON
echo "=== FORMATTED (if valid JSON) ==="
echo "$RESPONSE" | python3 -m json.tool 2>&1
