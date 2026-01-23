# Security Guide

> **Last Updated:** 2026-01-23  
> **Repository:** https://github.com/michelfritzschjoin/join-ms-365-mcp-server

## Overview

The Join Microsoft 365 MCP Server implements comprehensive security measures following ISO 27001 and DSGVO/GDPR standards.

## Security Architecture

```
┌─────────────────┐     ┌─────────────────┐     ┌─────────────────┐
│   MCP Client    │────►│   MCP Server    │────►│  Microsoft      │
│  (AI Assistant) │ TLS │  (Resource)     │ TLS │  Graph API      │
└─────────────────┘     └─────────────────┘     └─────────────────┘
                              │
                              ▼
                        ┌─────────────────┐
                        │  Secure Token   │
                        │    Storage      │
                        └─────────────────┘
```

## Authentication Security

### OAuth 2.1 with PKCE

All authentication flows use:

- **PKCE (Proof Key for Code Exchange):** Prevents authorization code interception
- **State Parameter:** Prevents CSRF attacks
- **Short-lived Access Tokens:** Minimize exposure window
- **Secure Refresh Token Storage:** Platform-specific secure storage

### Token Security

| Aspect     | Implementation                                                                  |
| ---------- | ------------------------------------------------------------------------------- |
| Storage    | Platform keychain (Windows Credential Manager, macOS Keychain, Linux libsecret) |
| Encryption | AES-256 for fallback file storage                                               |
| Logging    | Tokens are NEVER logged                                                         |
| Transport  | TLS 1.2+ required                                                               |

## Data Protection (DSGVO/GDPR)

### Data Minimization

The server follows the principle of data minimization:

- Only requests necessary Microsoft Graph scopes
- Does not cache personal data beyond tokens
- Implements pagination to limit data exposure

### Data Processing Principles

```typescript
// Article 5 DSGVO Compliance
const DATA_PROCESSING = {
  lawful: true, // Clear legal basis for processing
  fair: true, // Transparent processing
  purposeLimited: true, // Data used only for stated purpose
  minimized: true, // Collect only necessary data
  accurate: true, // Data is current
  storageLimited: true, // Retention only as needed
  secure: true, // Appropriate security measures
};
```

### User Rights (Articles 15-22 DSGVO)

| Right                   | Implementation                  |
| ----------------------- | ------------------------------- |
| Access (Art. 15)        | User can view cached accounts   |
| Rectification (Art. 16) | N/A - no persistent user data   |
| Erasure (Art. 17)       | `--logout` / `--remove-account` |
| Portability (Art. 20)   | N/A - no persistent user data   |

## ISO 27001 Compliance

### A.9 Access Control

```typescript
// Principle of least privilege
const ACCESS_CONTROL = {
  minimumScopes: true, // Request only needed permissions
  scopeFiltering: true, // --preset limits exposed tools
  readOnlyMode: true, // --read-only for read-only access
  tokenValidation: true, // Validate every request
};
```

### A.10 Cryptography

| Requirement    | Implementation     |
| -------------- | ------------------ |
| Key Length     | AES-256, RSA-2048+ |
| Hash Algorithm | SHA-256            |
| TLS Version    | 1.2 or higher      |
| Token Entropy  | 32 bytes minimum   |

### A.12 Operations Security

```typescript
// Audit logging (without sensitive data)
interface AuditLog {
  timestamp: string; // ISO 8601
  action: string; // What happened
  toolName: string; // Which tool was called
  success: boolean; // Result
  // NEVER includes: tokens, passwords, PII
}
```

## Security Best Practices

### For Operators

1. **Use Custom Azure AD App**
   - Register your own app for production
   - Configure appropriate permissions
   - Enable Conditional Access

2. **Enable Read-Only Mode**

   ```bash
   ms-365-mcp-server --read-only
   ```

3. **Use Presets to Limit Tools**

   ```bash
   ms-365-mcp-server --preset=mail,calendar
   ```

4. **Use Azure Key Vault for Secrets**

   ```env
   MS365_MCP_KEYVAULT_URL=https://your-vault.vault.azure.net/
   ```

5. **Enable Verbose Logging for Audit**
   ```bash
   ms-365-mcp-server -v
   ```

### For AI Assistants

1. **Validate User Intent**
   - Confirm before sending emails
   - Confirm before modifying data

2. **Handle Data Carefully**
   - Don't store sensitive data
   - Don't expose tokens in responses

3. **Follow Least Privilege**
   - Only request necessary data
   - Use pagination appropriately

## Input Validation

All inputs are validated using Zod schemas:

```typescript
// Example: Mail message validation
const mailMessageSchema = z.object({
  subject: z.string().max(255),
  body: z.object({
    content: z.string().max(100000),
    contentType: z.enum(['text', 'html']),
  }),
  toRecipients: z.array(
    z.object({
      emailAddress: z.object({
        address: z.string().email(),
      }),
    })
  ),
});
```

## Error Handling

### Security Error Responses

Errors are designed to be helpful without exposing internals:

```typescript
// ✅ Good: Helpful without exposure
"Authentication required. Please use the 'login' tool.";

// ❌ Bad: Exposes implementation details
'JWT validation failed: invalid signature at line 42';
```

### Error Logging

```typescript
// Security-aware logging
logger.error('Authentication failed', {
  action: 'verify-token',
  success: false,
  // NO token, NO user PII
});
```

## Rate Limiting

The server implements rate limiting for HTTP mode:

| Endpoint       | Limit                |
| -------------- | -------------------- |
| General        | 100 requests/minute  |
| Authentication | 10 requests/minute   |
| High-volume    | 1000 requests/minute |

## Security Headers

For HTTP mode, security headers are applied:

```typescript
const SECURITY_HEADERS = {
  'X-Content-Type-Options': 'nosniff',
  'X-Frame-Options': 'DENY',
  'X-XSS-Protection': '1; mode=block',
  'Content-Security-Policy': "default-src 'self'",
  'Strict-Transport-Security': 'max-age=31536000; includeSubDomains',
};
```

## Vulnerability Reporting

If you discover a security vulnerability, please report it responsibly:

1. **Email:** security@yourcompany.com
2. **Do not** create public GitHub issues for security vulnerabilities
3. Allow reasonable time for response and fix

## Security Checklist

### Before Deployment

- [ ] Use HTTPS in production
- [ ] Configure custom Azure AD app
- [ ] Enable appropriate logging
- [ ] Set up rate limiting
- [ ] Review granted permissions
- [ ] Enable MFA for admin accounts

### Ongoing

- [ ] Monitor audit logs
- [ ] Update dependencies regularly
- [ ] Review access patterns
- [ ] Rotate secrets periodically
- [ ] Test security configurations

---

_For Microsoft-specific security documentation, see [Microsoft Graph Security documentation](https://docs.microsoft.com/en-us/graph/security-concept-overview)._
