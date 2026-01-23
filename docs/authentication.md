# Authentication Guide

> **Last Updated:** 2026-01-23  
> **Repository:** https://github.com/michelfritzschjoin/join-ms-365-mcp-server

## Overview

The Join Microsoft 365 MCP Server uses OAuth 2.1 with PKCE for secure authentication with Microsoft Graph API.

## Authentication Methods

### 1. Device Code Flow (Recommended)

The default and recommended method for interactive authentication:

```bash
ms-365-mcp-server --login
```

1. The server displays a URL and code
2. Open the URL in your browser
3. Enter the code and sign in with your Microsoft account
4. Grant the requested permissions
5. Return to the terminal - authentication is complete

### 2. Bring Your Own Token (BYOT)

For programmatic access with pre-existing tokens:

```env
MS365_MCP_OAUTH_TOKEN=your-access-token
```

### 3. Azure Key Vault Integration

For enterprise deployments with centralized secret management:

```env
MS365_MCP_KEYVAULT_URL=https://your-vault.vault.azure.net/
```

Required secrets in Key Vault:

- `ms365-mcp-client-id` (required)
- `ms365-mcp-tenant-id` (optional)
- `ms365-mcp-client-secret` (optional)
- `ms365-mcp-cloud-type` (optional)

## Azure AD App Registration

### Using Default App

The server includes a built-in Azure AD app registration that works for most scenarios. No configuration needed.

### Custom App Registration

For enterprise deployments or custom requirements:

#### Step 1: Create App Registration

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** → **App registrations**
3. Click **New registration**
4. Configure:
   - **Name:** Microsoft 365 MCP Server
   - **Supported account types:** Choose based on your needs
   - **Redirect URI:** `http://localhost` (Mobile and desktop applications)

#### Step 2: Configure Authentication

1. Go to **Authentication**
2. Under **Advanced settings**:
   - Enable **Allow public client flows** = Yes
3. Save changes

#### Step 3: Configure API Permissions

Add the following Microsoft Graph permissions:

**Delegated Permissions (User Consent):**

| Permission            | Description          | Required For        |
| --------------------- | -------------------- | ------------------- |
| `User.Read`           | Read user profile    | Basic functionality |
| `Mail.Read`           | Read mail            | Email tools         |
| `Mail.ReadWrite`      | Read/write mail      | Send email          |
| `Mail.Send`           | Send mail            | Send email          |
| `Calendars.Read`      | Read calendars       | Calendar tools      |
| `Calendars.ReadWrite` | Read/write calendars | Create events       |
| `Contacts.Read`       | Read contacts        | Contact tools       |
| `Contacts.ReadWrite`  | Read/write contacts  | Manage contacts     |
| `Files.Read`          | Read files           | OneDrive read       |
| `Files.ReadWrite`     | Read/write files     | OneDrive write      |
| `Tasks.Read`          | Read tasks           | To-Do tools         |
| `Tasks.ReadWrite`     | Read/write tasks     | Manage tasks        |
| `Notes.Read`          | Read OneNote         | OneNote tools       |
| `Notes.ReadWrite`     | Read/write OneNote   | Manage OneNote      |

**Organization Mode Additional Permissions:**

| Permission              | Description           | Required For     |
| ----------------------- | --------------------- | ---------------- |
| `Team.ReadBasic.All`    | Read teams            | Teams tools      |
| `Channel.ReadBasic.All` | Read channels         | Teams tools      |
| `Chat.Read`             | Read chats            | Teams chat       |
| `Sites.Read.All`        | Read SharePoint       | SharePoint tools |
| `Sites.ReadWrite.All`   | Read/write SharePoint | SharePoint write |
| `User.ReadBasic.All`    | Read user directory   | User lookup      |
| `Group.Read.All`        | Read groups           | Group operations |

#### Step 4: Configure in Server

```env
MS365_MCP_CLIENT_ID=your-client-id
MS365_MCP_TENANT_ID=your-tenant-id
```

## Account Management

### List Cached Accounts

```bash
ms-365-mcp-server --list-accounts
```

### Switch Between Accounts

```bash
ms-365-mcp-server --select-account <account-id>
```

### Remove Account

```bash
ms-365-mcp-server --remove-account <account-id>
```

### Logout

```bash
ms-365-mcp-server --logout
```

## Token Storage

### Secure Storage Locations

Tokens are stored securely using platform-specific mechanisms:

| Platform | Storage Method                          |
| -------- | --------------------------------------- |
| Windows  | Windows Credential Manager (via keytar) |
| macOS    | Keychain                                |
| Linux    | libsecret / GNOME Keyring               |

### Fallback Storage

If secure storage is unavailable, tokens are stored in an encrypted file:

- Location: `~/.ms365-mcp/token-cache.json`
- Encryption: AES-256 with machine-specific key

## Cloud Environments

### Global (Default)

```env
MS365_MCP_CLOUD_TYPE=global
```

Endpoints:

- Login: `https://login.microsoftonline.com`
- Graph: `https://graph.microsoft.com`

### China (21Vianet)

```env
MS365_MCP_CLOUD_TYPE=china
```

Endpoints:

- Login: `https://login.partner.microsoftonline.cn`
- Graph: `https://microsoftgraph.chinacloudapi.cn`

## Security Best Practices

### Token Security

- ✅ Tokens are stored in secure system keychain
- ✅ Tokens are never logged
- ✅ Tokens expire and are automatically refreshed
- ✅ HTTPS is required for all API calls

### Permission Scopes

- ✅ Request minimum required scopes
- ✅ Use `--preset` to limit tool exposure
- ✅ Enable `--read-only` for read-only access
- ✅ Review permissions in Azure Portal regularly

### Enterprise Recommendations

1. Use custom Azure AD app registration
2. Configure Conditional Access policies
3. Enable MFA for all users
4. Use Azure Key Vault for secrets
5. Monitor access logs in Azure AD

## Troubleshooting

### "No accounts found"

```bash
# Re-authenticate
ms-365-mcp-server --login
```

### "Token expired"

Tokens are automatically refreshed. If issues persist:

```bash
ms-365-mcp-server --logout
ms-365-mcp-server --login
```

### "Insufficient privileges"

1. Check required permissions in Azure Portal
2. Ensure admin consent is granted (for org permissions)
3. Re-authenticate after permission changes

### "AADSTS error codes"

| Code         | Description      | Solution               |
| ------------ | ---------------- | ---------------------- |
| AADSTS50076  | MFA required     | Complete MFA challenge |
| AADSTS65001  | Consent required | Grant permissions      |
| AADSTS700016 | App not found    | Check Client ID        |
| AADSTS90002  | Tenant not found | Check Tenant ID        |

---

_For more details, see [Microsoft Identity Platform documentation](https://docs.microsoft.com/en-us/azure/active-directory/develop/)._
