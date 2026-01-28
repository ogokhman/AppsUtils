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

# Get user email
read -p "Enter user email: " USER_EMAIL

# Make API request
echo ""
echo "Making request..."
API_URL="https://graph.microsoft.com/v1.0/users/$USER_EMAIL/mailFolders/Inbox/messages?\$select=subject,from,receivedDateTime&\$top=5"
echo "URL: $API_URL"
echo ""

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
