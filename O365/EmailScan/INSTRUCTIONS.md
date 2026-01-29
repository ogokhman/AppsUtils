# Project: Office365 Email Scanner CLI

## Goal
Build a command-line Python app to pull emails from Office365 using Microsoft Graph API.

## Project Location
- App directory: `O365/EmailScan/`
- Reuse Azure AD credentials from: `Office365API/.env`

## Authentication
- OAuth 2.0 Client Credentials flow
- Token endpoint: https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token
- Scope: https://graph.microsoft.com/.default
- Cache token with 1-hour expiry

## CLI Arguments
| Flag         | Required | Default | Description                              |
|--------------|----------|---------|------------------------------------------|
| `--user`     | Yes      |         | Target mailbox email address             |
| `--from`     | Yes      |         | Filter by sender (comma-separated)       |
| `--count`    | Yes      | 10      | Number of messages (0 = all)             |
| `--mindate`  | Yes      |         | Start date filter (YYYY-MM-DD)           |
| `--maxdate`  | Yes      |         | End date filter (YYYY-MM-DD)             |
| `--folder`   | Yes      | Inbox   | Mailbox folder (comma-separated)         |
| `--sort`     | Yes      | latest  | Sort order: earliest           |

## Graph API Endpoints
- Messages: `GET /v1.0/users/{user}/mailFolders/{folder}/messages`
- Folder lookup: `GET /v1.0/users/{user}/mailFolders/{id}`
- Use `$filter` for date ranges
- Use `$search` for sender filtering
- Use `$select` to limit fields: subject, from, toRecipients, receivedDateTime, parentFolderId
- Use `$orderby` for sort order
- Handle pagination via `@odata.nextLink`

## Output Format
- Formatted table with columns: #, Date/Time, From, To, Folder, Subject
- Dates displayed in EST timezone
- Truncate long fields for clean display

## Dependencies
- requests
- python-dotenv
- pytz

## Files to Create
- `EmailScan/email_scan.py` - Main CLI script
- `EmailScan/requirements.txt` - Python dependencies
- `EmailScan/.env.example` - Credentials template

## Notes
   Instead of sending parameters on command line, I want to present a menu for users to choose from

   1. select users: read users from crc_users.txt. Each user will be on a separate line and user can select them individually or ALL (letter A). Example,
         1. oleg@christoffersonrobb.com
	 2. apapritz@christoffersonrobb.com
	 3. malik@christoffersonrobb.com
	 A. ALL
	 
   2. select from. Similar idea. The from list will be read from a file search_domains.txt
   3. select mindate, i.e. the earlist date. Set default to today -10
   4. select maxdate, set Default to 'today' 
   5. select folders. Same idea as users. Read folders from email_folders.txt. Dont forget to add A

   sort will always be in ascending order

- I don't want to parse messages on the client side  that match the 'from' domain. I want to use graph's 'search' and not filter

