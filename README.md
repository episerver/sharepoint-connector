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

#### Step 2 - Create a client secret

1. In your app registration, go to **Certificates & secrets** > **Client secrets** > **New client secret**.
2. Enter a description and choose an expiry period.
3. Click **Add**.
4. Copy the **Value** of the secret immediately -- it will not be shown again.

#### Step 3 - Grant SharePoint API permissions

1. In your app registration, go to **API permissions** > **Add a permission**.
2. Select **SharePoint** > **Application permissions**.
3. Select the permissions your connector needs:
   - `Sites.Read.All` -- read access to all site collections (minimum required)
   - `Sites.ReadWrite.All` -- if write access is also needed
4. Click **Add permissions**.
5. Click **Grant admin consent for [your organization]** and confirm.

#### Step 4 - Configure the connector

1. In Optimizely CMS, go to **Admin** > **Connect For SharePoint**.
2. Enter your SharePoint site URL (e.g. `https://contoso.sharepoint.com/sites/media`).
3. Under **Connection type**, select **Online**.
4. Under **Authentication type**, select **Modern**.
5. Fill in the following fields:
   - **Client ID** -- the Application (client) ID from Step 1
   - **Client Secret** -- the secret value from Step 2
   - **Tenant ID** -- the Directory (tenant) ID from Step 1
6. Click **Test connection** to verify the credentials.
7. Select the SharePoint lists and views you want to sync.
8. Choose a destination folder for the imported assets.
9. Click **Save**.

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
| "Unauthorized" error with Modern auth | Verify the Client ID, Client Secret, and Tenant ID are correct. Ensure admin consent has been granted for the API permissions. |
| Token request fails | Confirm the client secret has not expired. Create a new secret if needed. |
| "Access denied" to specific lists | The app registration may need higher-level permissions (e.g. `Sites.ReadWrite.All`). Grant the required permissions and re-consent. |
| On-premises connection fails | Ensure the account has access to the SharePoint site and that NTLM authentication is enabled on the server. |

---

