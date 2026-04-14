# SharePoint Authentication Migration Plan

## Date: April 2026

## Problem

This connector currently supports three authentication methods, two of which are deprecated by Microsoft:

### 1. Cookie-Based Authentication (SharePoint Online "Basic Auth") - DEPRECATED
- **How it works**: Uses `SharePointOnlineCredentials` from the CSOM library to obtain FedAuth/rtFA cookies via SAML token exchange with `login.microsoftonline.com/extSTS.srf`.
- **Key files**:
  - `Helpers/SharePointHelper.cs` - `GetLatestCookies()` and `GetCookies()` methods
  - `Authentication/MsOnlineClaimsHelper.cs` - WSTrust SAML token acquisition
  - `Authentication/ClaimsWebClient.cs` - WebClient with cookie injection
  - `Authentication/WcfClientContracts.cs` - WCF contracts for WSTrust
- **Why deprecated**: Microsoft disabled legacy authentication for SharePoint Online. The `SharePointOnlineCredentials.GetAuthenticationCookie()` API no longer works for most tenants.

### 2. Azure ACS OAuth (SharePoint Online "Modern Auth") - ALSO DEPRECATED
- **How it works**: Uses OAuth 2.0 client credentials flow against Azure Access Control Service (ACS) at `accounts.accesscontrol.windows.net/{realm}/tokens/OAuth/2`.
- **Key files**:
  - `Helpers/SharePointHelper.cs` - `GetOAuthResponse()` method
  - `Helpers/SharePointOAuthTokenResult.cs` - Token response model
- **Why deprecated**: Azure ACS was retired by Microsoft (April 2023). While some existing app principals may still work temporarily, no new ACS app registrations can be created and the service is being decommissioned.

### 3. NTLM (On-Premises) - Still Supported
- **How it works**: Uses `CredentialCache` with NTLM scheme for on-premises SharePoint.
- **Key files**: `Helpers/SharePointHelper.cs` - `GetCredentialCache()` method
- **Status**: Still valid for on-premises SharePoint Server deployments.

## Solution: Migrate to Microsoft Entra ID (Azure AD) OAuth 2.0

Replace the Azure ACS OAuth flow with the modern Microsoft Identity Platform (Entra ID) OAuth 2.0 client credentials grant.

### What Changes

| Aspect | Old (Azure ACS) | New (Entra ID) |
|--------|-----------------|----------------|
| Token endpoint | `accounts.accesscontrol.windows.net/{realm}/tokens/OAuth/2` | `login.microsoftonline.com/{tenantId}/oauth2/v2.0/token` |
| Scope/Resource | `00000003-0000-0ff1-ce00-000000000000/{host}@{realm}` | `https://{tenant}.sharepoint.com/.default` |
| Client ID format | `{clientId}@{realm}` | `{clientId}` (plain GUID) |
| Grant type | `client_credentials` | `client_credentials` (same) |
| App registration | SharePoint ACS app principal (`/_layouts/15/appregnew.aspx`) | Azure Portal > App registrations (Entra ID) |

### Files Modified

1. **`Helpers/SharePointHelper.cs`** - Updated `GetOAuthResponse()` to call Entra ID token endpoint with `scope` instead of `resource`, plain `client_id` instead of `{clientId}@{realm}`.
2. **`Web/SharePointConfiguration.aspx.cs`** - Fixed Modern Auth validation to use the correct error message key (`nourlmodernauthvaluesspecified` instead of `nourlusernamepasswordspecified`).
3. **`EmbeddedLangFiles/EPiServer.SharePoint_*.xml`** (all 12 language files: DA, DE, EN, ES, FI, FR, IT, JA, NL, NO, SV, ZH) - Updated "Realm" label to "Tenant ID", "Basic" to "Basic (On-Premises Only)", and modernized the validation error message.
4. **`.gitignore`** - Updated with standard .NET glob patterns.
5. **`README.md`** - Added step-by-step authentication setup guide for customers.

**Not modified** (by design):
- **`Configuration/SharePointSettings.cs`** - The existing `Realm` property is reused to store the tenant ID. This preserves backward compatibility with existing DDS data so customers who already have a value stored do not need to reconfigure.
- **`Web/SharePointConfiguration.aspx`** - No structural changes needed; the field is bound to the `Realm` property via language file labels which now display "Tenant ID".

### Customer Migration Steps

When customers upgrade to this version, they need to:

1. **Register an app in Microsoft Entra ID** (Azure Portal > App registrations):
   - Create a new app registration
   - Note the Application (client) ID and Directory (tenant) ID
   - Create a client secret under Certificates & secrets
   - Under API permissions, add `SharePoint > Application permissions > Sites.Read.All` (or appropriate level)
   - Grant admin consent for the permissions

2. **Update the connector configuration**:
   - Select "Modern" authentication type
   - Enter the Application (client) ID as "Client ID"
   - Enter the client secret as "Client Secret"
   - Enter the Directory (tenant) ID as "Tenant ID"
   - Click Connect to test

### Cookie-Based Auth Removal

The cookie-based authentication code (`MsOnlineClaimsHelper`, `ClaimsWebClient`, `WcfClientContracts`, `GetLatestCookies`, `GetCookies`) is left in place for now but is effectively dead code for SharePoint Online customers. It can be removed in a future cleanup pass. The "Basic" auth radio button will remain functional only for on-premises NTLM connections.
