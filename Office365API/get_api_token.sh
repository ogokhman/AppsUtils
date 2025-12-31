#!/bin/bash
# Load credentials from .env file or environment variables
source .env 2>/dev/null || true

tenantId="${TENANT_ID}";
clientId="${CLIENT_ID}";
clientSecret="${CLIENT_SECRET}";
scope="https://graph.microsoft.com/.default";   


curl -X POST https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token \
  -H "Content-Type: application/x-www-form-urlencoded" \
  -d "client_id=$clientId" \
  -d "scope=$scope "\
  -d "client_secret=$clientSecret" \
  -d "grant_type=client_credentials"
