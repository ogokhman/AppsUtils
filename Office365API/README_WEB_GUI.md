# Office 365 Email Viewer - Web GUI

A simple web interface for fetching Office 365 email messages using the Microsoft Graph API.

## Features

- **Web Interface**: User-friendly HTML form to enter email addresses
- **Token Caching**: Automatically handles API token generation and caching
- **Real-time Results**: Displays messages in a clean table format
- **Error Handling**: Shows helpful error messages for troubleshooting
- **Responsive Design**: Works on desktop and mobile devices

## Installation

1. Install required dependencies:
```bash
pip install -r requirements.txt
```

2. Ensure your `.env` file is configured with your Office 365 credentials:
```
CLIENT_ID=your_client_id
TENANT_ID=your_tenant_id
CLIENT_SECRET=your_client_secret
```

## Running the Web GUI

1. Start the web server:
```bash
python web_gui.py
```

2. Open your web browser and navigate to:
```
http://localhost:5000
```

3. Enter a user email address and click "Get Recent Messages"

## Usage

1. **Enter Email**: Type the target user's email address in the input field
2. **Submit**: Click the "Get Recent Messages" button
3. **View Results**: The page will display up to 3 recent messages in a table format
4. **Token Management**: The system automatically handles token generation and caching

## Output Format

The results display the following information for each message:
- **From Address**: Sender's email address
- **From Name**: Sender's display name
- **To Address**: Recipient's email address
- **To Name**: Recipient's display name
- **Subject**: Email subject line

## Security Notes

- **Credentials**: Your Office 365 credentials are stored in the `.env` file and are not exposed in the web interface
- **Token Caching**: API tokens are cached locally to reduce API calls
- **Local Only**: By default, the server runs on localhost only

## Troubleshooting

- **No Messages Found**: Ensure the user has emails in their mailbox and the API permissions are correct
- **Token Errors**: Check your Office 365 app registration and permissions
- **Connection Issues**: Verify your internet connection and firewall settings

## API Permissions Required

Your Office 365 app needs the following permissions:
- `User.Read.All`
- `Mail.Read`

## File Structure

```
Office365API/
├── web_gui.py              # Flask web application
├── get_user_email.py       # Core email fetching script
├── templates/
│   └── index.html         # Web interface HTML
├── .env                   # Credentials and token cache
├── requirements.txt       # Python dependencies
└── README_WEB_GUI.md      # This file
```
