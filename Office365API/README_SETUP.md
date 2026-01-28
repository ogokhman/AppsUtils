# Office 365 API Setup Guide

This guide will help you set up the necessary credentials to use the Office 365 API tools in this project.

## Prerequisites

- An active Microsoft 365 subscription
- Administrative access to Azure Active Directory (now called Microsoft Entra ID)
- Access to the Azure Portal

## Step 1: Access Azure Portal

1. Go to [https://portal.azure.com](https://portal.azure.com)
2. Sign in with your Microsoft 365 administrator account

## Step 2: Find Your Tenant ID

1. In the Azure Portal search bar, type **"Azure Active Directory"** or **"Microsoft Entra ID"**
2. Click on the service to open it
3. On the **Overview** page, you'll see **Tenant ID** (also called Directory ID)
4. **Copy this value** - this is your `TENANT_ID`

## Step 3: Register a New Application

1. In Azure Active Directory, click **App registrations** in the left sidebar
2. Click **+ New registration** at the top
3. Fill in the registration form:
   - **Name**: Choose a descriptive name (e.g., "Office365 Email API")
   - **Supported account types**: Select "Accounts in this organizational directory only"
   - **Redirect URI**: Leave blank for now
4. Click **Register**

## Step 4: Get Your Client ID

1. After registration, you'll be taken to the app's Overview page
2. Find **Application (client) ID** on this page
3. **Copy this value** - this is your `CLIENT_ID`

## Step 5: Create a Client Secret

1. In your app registration, click **Certificates & secrets** in the left sidebar
2. Under "Client secrets", click **+ New client secret**
3. Add a description (e.g., "API Access Secret")
4. Choose an expiration period:
   - **Recommended**: 24 months (maximum security balance)
   - Note: You'll need to create a new secret before it expires
5. Click **Add**
6. **IMPORTANT**: Immediately copy the **Value** (not the Secret ID)
   - This is your `CLIENT_SECRET`
   - ⚠️ **You won't be able to see this value again!**
   - If you lose it, you'll need to create a new secret

## Step 6: Configure API Permissions

Your application needs permission to access Microsoft Graph API for reading emails.

1. In your app registration, click **API permissions** in the left sidebar
2. Click **+ Add a permission**
3. Select **Microsoft Graph**
4. Select **Application permissions** (not Delegated permissions)
5. Add the following permissions:
   - `Mail.Read` - Read mail in all mailboxes
   - `Mail.ReadWrite` - Read and write mail in all mailboxes (optional, only if you need write access)
   - `User.Read.All` - Read all users' full profiles
6. Click **Add permissions**

## Step 7: Grant Admin Consent

⚠️ **This step requires administrator privileges**

1. After adding permissions, you'll see them listed with a status of "Not granted"
2. Click the **✓ Grant admin consent for [Your Organization]** button
3. Confirm by clicking **Yes**
4. Wait for the status to change to "Granted for [Your Organization]"

Without this step, the API calls will fail with permission errors.

## Step 8: Update Your .env File

1. Copy the `.env.example` file to `.env`:
   ```bash
   cp .env.example .env
   ```

2. Edit the `.env` file and replace the placeholder values:
   ```bash
   nano .env
   ```

3. Update these lines with your actual values:
   ```
   TENANT_ID=your-actual-tenant-id-here
   CLIENT_ID=your-actual-client-id-here
   CLIENT_SECRET=your-actual-client-secret-here
   ```

4. Save the file (in nano: Ctrl+O, Enter, Ctrl+X)

5. The `ACCESS_TOKEN` and `TOKEN_TIMESTAMP` fields will be automatically populated when you run the scripts

## Step 9: Test Your Configuration

Run the test script to verify everything is configured correctly:

```bash
./test_api_curl.sh
```

Or test with Python:

```bash
python get_user_email.py user=your-email@yourdomain.com count=5
```

## Troubleshooting

### Error: "Invalid client secret"
- Your `CLIENT_SECRET` may have expired or was copied incorrectly
- Create a new client secret in Azure Portal (Step 5)

### Error: "Insufficient privileges"
- Admin consent was not granted (Step 7)
- Or you don't have the required API permissions (Step 6)

### Error: "Application not found"
- Your `CLIENT_ID` or `TENANT_ID` is incorrect
- Verify the values in Azure Portal (Steps 2 and 4)

### Error: "AADSTS700016: Application not found in the directory"
- The `TENANT_ID` doesn't match the tenant where the app is registered
- Verify you're using the correct tenant

## Security Best Practices

1. **Never commit your .env file to version control**
   - The `.gitignore` file should already exclude it
   - Double-check before committing

2. **Rotate client secrets regularly**
   - Set a calendar reminder before expiration
   - Microsoft recommends rotating every 6-12 months

3. **Use the principle of least privilege**
   - Only grant the minimum permissions needed
   - Review permissions periodically

4. **Store credentials securely**
   - Keep the `.env` file permissions restricted: `chmod 600 .env`
   - Use Azure Key Vault for production environments

## Additional Resources

- [Microsoft Graph API Documentation](https://docs.microsoft.com/en-us/graph/)
- [App Registration Documentation](https://docs.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app)
- [Microsoft Graph Permissions Reference](https://docs.microsoft.com/en-us/graph/permissions-reference)
- [Azure AD Application Permissions](https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-permissions-and-consent)

## Support

If you encounter issues not covered in this guide:
1. Check the Azure Portal audit logs for detailed error messages
2. Review Microsoft Graph API error codes
3. Verify your Microsoft 365 subscription includes API access
