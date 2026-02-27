# Setting Up Excel Online Oracle via MS Graph

The `--oracle excel` fuzz mode uses the MS Graph API to upload XLSX files to OneDrive and evaluate formulas with Excel Online. This requires a Microsoft 365 tenant with SharePoint Online / OneDrive for Business.

## Why Your Current Azure Tenant Doesn't Work

Azure AD (for Azure subscriptions) and Microsoft 365 are separate products. An Azure-only tenant does not include SharePoint Online or OneDrive for Business, which the Graph API needs to store and open Excel files. You'll get this error:

```
Tenant does not have a SPO license.
```

## Setup Steps

### 1. Get a Microsoft 365 Business Basic subscription

- Go to https://www.microsoft.com/en-us/microsoft-365/business/microsoft-365-business-basic
- Sign up for the free 1-month trial ($6/user/month after that)
- This creates a new tenant (e.g. `yourname.onmicrosoft.com`) with OneDrive + SharePoint Online
- Note: `admin.microsoft.com` only works for Microsoft 365 tenants, not Azure-only tenants

### 2. Register an app in the new tenant

- Go to https://portal.azure.com and switch to your new Microsoft 365 tenant
- Navigate to **Azure Active Directory** > **App registrations** > **New registration**
- Name: `werkbook-fuzz` (or whatever you like)
- Supported account types: **Accounts in this organizational directory only**
- Redirect URI: leave blank (not needed for device code flow)
- Click **Register**

### 3. Configure the app as a public client

- In your app registration, go to **Authentication**
- Under **Advanced settings**, set **Allow public client flows** to **Yes**
- Click **Save**
- This is required for device code flow (no client secret needed)

### 4. Add API permissions

- Go to **API permissions** > **Add a permission**
- Choose **Microsoft Graph** > **Delegated permissions**
- Add these permissions:
  - `Files.ReadWrite`
  - `User.Read`
  - `offline_access`
- Click **Grant admin consent** for your organization

### 5. Note your new IDs

From the app registration **Overview** page, copy:
- **Application (client) ID** — this is your `MSGRAPH_CLIENT_ID`
- **Directory (tenant) ID** — this is your `MSGRAPH_TENANT_ID`

### 6. Update your `.env`

```
MSGRAPH_TENANT_ID=<new-tenant-id>
MSGRAPH_CLIENT_ID=<new-client-id>
```

You can remove `MSGRAPH_CLIENT_SECRET` if you configured the app as a public client.

### 7. Run the setup command

```bash
go run ./cmd/msgraph-setup/
```

This will open a browser for device code auth. Sign in with your `@yourname.onmicrosoft.com` account.

### 8. Run the fuzzer with Excel Online

```bash
go run ./cmd/fuzz/ --oracle excel -rounds 5 -v
```

## Fallback: LibreOffice

If you don't want to set up Microsoft 365, LibreOffice works as a local oracle with no account needed:

```bash
# Install LibreOffice (macOS)
brew install --cask libreoffice

# Run fuzzer
go run ./cmd/fuzz/ --oracle libreoffice -rounds 5 -v
```
