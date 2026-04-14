<a href="https://github.com/episerver/sharepoint-connector"><img src="https://ux.episerver.com/images/logo.png" title="Sharepoint Connector" alt="Sharepoint Connector"></a>

## Sharepoint Connector
Optimizely Connect for SharePoint makes it easy to transfer documents from your SharePoint instance to Optimizely assets.

## Table of Contents

- [System requirements](#system-requirements)
- [Pre-installation set-up](#pre-installation-set-up)
- [Installation](#installation)
- [Troubleshooting](#troubleshooting)
- [Modular set-up](#modular-set-up)
- [Contributing](#contributing)

---

## System requirements

* Visual Studio 2017 or higher - [Download](https://visualstudio.microsoft.com/downloads/)

See also the general [Optimizely system requirements](https://world.optimizely.com/documentation/system-requirements/) on Optimizely World.

---

## Authentication

The connector supports two authentication modes:

| Mode | Use case |
|------|----------|
| **Modern (OAuth 2.0)** | SharePoint Online (recommended) |
| **Basic (NTLM)** | On-premises SharePoint Server only |

> **Note:** Cookie-based authentication and Azure ACS app principals have been deprecated by Microsoft. SharePoint Online customers must use Modern authentication with a Microsoft Entra ID (Azure AD) app registration.

### Setting up Modern Authentication (SharePoint Online)

#### Step 1 - Register an application in Microsoft Entra ID

1. Sign in to the [Azure Portal](https://portal.azure.com).
2. Navigate to **Microsoft Entra ID** > **App registrations** > **New registration**.
3. Enter a name for the application (e.g. "Optimizely SharePoint Connector").
4. Under **Supported account types**, select **Accounts in this organizational directory only**.
5. Leave **Redirect URI** blank (not required for client credentials flow).
6. Click **Register**.
7. On the app overview page, copy the **Application (client) ID** and **Directory (tenant) ID** -- you will need these later.

#### Step 2 - Create a credential (client secret or certificate)

The connector supports two credential types. Certificate-based auth is recommended as it is required for `Sites.Selected` permissions and avoids secret expiry.

**Option A -- Client Secret (simpler setup):**

1. In your app registration, go to **Certificates & secrets** > **Client secrets** > **New client secret**.
2. Enter a description and choose an expiry period.
3. Click **Add**.
4. Copy the **Value** of the secret immediately -- it will not be shown again.

**Option B -- Certificate (recommended):**

1. Create a self-signed certificate (or use a CA-issued one):
   ```powershell
   $cert = New-SelfSignedCertificate -Subject "CN=Optimizely SharePoint Connector" -CertStoreLocation "Cert:\LocalMachine\My" -KeyExportPolicy Exportable -KeySpec Signature -KeyLength 2048 -HashAlgorithm SHA256 -NotAfter (Get-Date).AddYears(2)
   Export-Certificate -Cert $cert -FilePath "SharePointConnector.cer"
   $cert.Thumbprint | Set-Clipboard
   Write-Host "Thumbprint copied to clipboard: $($cert.Thumbprint)"
   ```
2. In your app registration, go to **Certificates & secrets** > **Certificates** > **Upload certificate**.
3. Upload the `.cer` file exported above.
4. Note the **Thumbprint** -- you will need it when configuring the connector.
5. Ensure the certificate (with private key) is installed in the **Windows Certificate Store** (Local Machine\Personal) on the server running the CMS site. The IIS application pool identity must have read access to the certificate's private key.

#### Step 3 - Grant SharePoint API permissions

1. In your app registration, go to **API permissions** > **Add a permission**.
2. Select **SharePoint** (not Microsoft Graph) > **Application permissions**.
3. Select the permissions your connector needs:
   - `Sites.Read.All` -- read access to all site collections (minimum required)
   - `Sites.ReadWrite.All` -- if write access is also needed
4. Click **Add permissions**.
5. Click **Grant admin consent for [your organization]** and confirm.
6. Verify that all permissions show a green checkmark under the **Status** column.

> **Important:** You must select permissions under **SharePoint**, not Microsoft Graph. Both APIs have similarly named permissions, but only SharePoint Application permissions work with this connector.
>
> **Note on `Sites.Selected`:** This permission requires **certificate-based authentication**. If you use a client secret with `Sites.Selected`, you will get an "Unsupported app only token" error. To use `Sites.Selected`, select "Certificate" as the credential type in Step 2 (Option B) and Step 4, and grant per-site access (see Step 3c below).

#### Step 3b - Ensure app-only access is enabled

SharePoint Online has a tenant-level setting that can block app-only access. Verify it is enabled using SharePoint Online Management Shell (or [Azure Cloud Shell](https://shell.azure.com)):

```powershell
# Install module if needed
Install-Module -Name Microsoft.Online.SharePoint.PowerShell

# Connect as SharePoint admin
Connect-SPOService -Url https://{tenant}-admin.sharepoint.com

# Check the setting
Get-SPOTenant | Select-Object DisableCustomAppAuthentication

# If it returns True, enable app-only access
Set-SPOTenant -DisableCustomAppAuthentication $false
```

#### Step 3c - Grant per-site access (only if using `Sites.Selected`)

If you chose `Sites.Selected` in Step 3, you must grant the app access to each SharePoint site individually using Microsoft Graph PowerShell:

```powershell
# Install the module if needed
Install-Module Microsoft.Graph.Sites

# Connect with admin permissions
Connect-MgGraph -Scopes "Sites.FullControl.All"

# Find your site
$site = Get-MgSite -Search "your-site-name"

# Grant the app read access to the site
New-MgSiteSPermission -SiteId $site.Id -BodyParameter @{
    roles = @("read")  # use "write" if write access is needed
    grantedToIdentities = @(@{
        application = @{
            id = "<your-app-client-id>"
            displayName = "Optimizely SharePoint Connector"
        }
    })
}
```

> **Note:** Replace `<your-app-client-id>` with the Application (client) ID from Step 1. Repeat the command for each site the connector needs to access.

#### Step 4 - Configure the connector

1. In Optimizely CMS, go to **Admin** > **Connect For SharePoint**.
2. Enter your SharePoint site URL (e.g. `https://contoso.sharepoint.com/sites/media`).
3. Under **Connection type**, select **Online**.
4. Under **Authentication type**, select **Modern**.
5. Under **Credential Type**, select either **Client Secret** or **Certificate**.
6. Fill in the following fields:
   - **Client ID** -- the Application (client) ID from Step 1
   - **Client Secret** (if using client secret) -- the secret value from Step 2
   - **Certificate Thumbprint** (if using certificate) -- the thumbprint from Step 2
   - **Tenant ID** -- the Directory (tenant) ID from Step 1
7. Click **Test connection** to verify the credentials.
8. Select the SharePoint lists and views you want to sync.
9. Choose a destination folder for the imported assets.
10. Click **Save**.

### Setting up Basic Authentication (On-Premises only)

1. In Optimizely CMS, go to **Admin** > **Connect For SharePoint**.
2. Enter your SharePoint Server URL.
3. Under **Connection type**, select **On-premises**.
4. Under **Authentication type**, select **Basic (On-Premises Only)**.
5. Enter the **Username** and **Password** of a domain account with access to the SharePoint site.
6. Click **Test connection** to verify.
7. Select the lists/views and destination folder, then click **Save**.

### Troubleshooting authentication

| Problem | Solution |
|---------|----------|
| "Unauthorized" error with Modern auth | Verify the Client ID, Client Secret, and Tenant ID are correct. Ensure admin consent has been granted for the API permissions (green checkmark in Status column). |
| "Unauthorized" after token succeeds | App-only access may be disabled at the tenant level. Check with `Get-SPOTenant \| Select-Object DisableCustomAppAuthentication`. If `True`, run `Set-SPOTenant -DisableCustomAppAuthentication $false` (see Step 3b). |
| "Unsupported app only token" | You may be using a client secret with `Sites.Selected` permission. The SharePoint REST API requires certificate-based authentication for `Sites.Selected`. Switch the credential type to "Certificate" in the connector settings, or use `Sites.Read.All` with a client secret instead. |
| Token request fails (400 Bad Request) | Verify the Tenant ID is correct (GUID from Entra ID Overview). Confirm the client secret has not expired. Ensure the SharePoint URL starts with `https://`. |
| No `roles` claim in token | Admin consent was not applied. Go to **API permissions** and click **Grant admin consent for [your organization]**. Get a new token after consenting. |
| "Access denied" to specific lists | The app registration may need higher-level permissions (e.g. `Sites.ReadWrite.All`). Grant the required permissions and re-consent. |
| On-premises connection fails | Ensure the account has access to the SharePoint site and that NTLM authentication is enabled on the server. |

---

