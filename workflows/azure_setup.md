# Azure App Registration Setup Guide
## For Mownur Services — Microsoft 365 Integration

**Who this is for:** A non-technical tax professional who has never used Azure before.
**What you'll end up with:** Three credentials — `AZURE_CLIENT_ID`, `AZURE_CLIENT_SECRET`, and `AZURE_TENANT_ID` — that let your automation tools securely access OneDrive, SharePoint, and Outlook on behalf of Mownur Services.
**Time required:** 20–30 minutes.
**What you need before you start:** An account with Global Administrator or Application Administrator role in your Microsoft 365 organization.

---

## What Is Azure App Registration?

Think of it like creating an employee badge for your automation software. Just like an employee needs a badge to enter the building and access certain rooms, your automation tools need an "app registration" to access Microsoft 365 services. You decide exactly which "rooms" (services) it can enter.

---

## Step 1: Open the Azure Portal

1. Open your web browser (Chrome or Edge recommended).
2. Go to: **https://portal.azure.com**
3. Sign in with your Microsoft 365 admin account (the same email and password you use for Outlook or Teams at Mownur Services).
4. You'll land on the Azure home page. It has a dark blue/gray navigation bar at the top and a bunch of tiles in the center. This is normal — it looks complicated but you'll only use a small part of it.

---

## Step 2: Find "App Registrations"

1. Look at the top of the page. You'll see a **search bar** that says *"Search resources, services, and docs"*.
2. Click on that search bar and type: **App registrations**
3. You'll see a dropdown appear. Under the section labeled **Services**, click on **App registrations**.
4. You're now on the App registrations page. It will likely say *"No results"* or show a list if others exist. Either is fine.

---

## Step 3: Create a New App Registration

1. Near the top-left of the page, click the button that says **+ New registration**.
2. You'll see a form with three sections. Fill them out as follows:

**Name:**
- Type: `Mownur Services Automation`
- This is just a label so you can find it later. It doesn't affect how it works.

**Supported account types:**
- Select the first option: **Accounts in this organizational directory only (Mownur Services only — Single tenant)**
- This means only your organization can use this app. This is the most secure option.

**Redirect URI (optional):**
- Leave this section completely blank. Do not fill anything in here.

3. Scroll down and click the blue **Register** button at the bottom.

---

## Step 4: Collect Your First Two Credentials

After clicking Register, you'll be taken to a page that shows your new app. This page has the two first credentials you need.

1. Look for a section called **Essentials** near the top. You'll see two fields:

**Application (client) ID:**
- This is your `AZURE_CLIENT_ID`.
- It looks like: `a1b2c3d4-e5f6-7890-abcd-ef1234567890`
- Copy this value and save it somewhere safe (a notepad, a password manager, or a secure note).

**Directory (tenant) ID:**
- This is your `AZURE_TENANT_ID`.
- It also looks like a string of letters, numbers, and dashes.
- Copy this value and save it in the same place.

> **Important:** Do not share these values publicly. They are like the first part of a lock combination.

---

## Step 5: Create a Client Secret

The client secret is like a password for your app. You'll generate it here.

1. On the left-hand sidebar of the page, look for a section called **Manage**. Click on **Certificates & secrets**.
2. You'll see three tabs: *Certificates*, *Client secrets*, and *Federated credentials*. Make sure you're on the **Client secrets** tab.
3. Click the button: **+ New client secret**
4. A panel slides in from the right with two fields:

**Description:**
- Type: `Mownur Automation Key`

**Expires:**
- Select: **24 months** (this means you'll need to renew it in 2 years — set a calendar reminder for 23 months from today).

5. Click the **Add** button at the bottom of the panel.

6. The page will now show a new entry in the table. **You must copy the secret value RIGHT NOW.** Look at the **Value** column (not the Secret ID column) and copy the long string of characters.

- This is your `AZURE_CLIENT_SECRET`.
- It looks like: `abc~defGHIjkl12MNOPqrstu.VWXYZabcdef`
- **This value will be hidden forever after you leave this page.** If you don't copy it now, you'll have to delete it and create a new one.

> Save this in the same secure location as your Client ID and Tenant ID.

---

## Step 6: Grant API Permissions

This is where you tell Microsoft which "rooms" your app is allowed to enter. You'll grant four specific permissions.

1. In the left-hand sidebar, click **API permissions** (still under the **Manage** section).
2. You'll see a table with one permission already listed: `Microsoft Graph — User.Read — Delegated`. You'll keep this one and add more.
3. Click **+ Add a permission**.
4. A panel slides in. Click on **Microsoft Graph** (the large tile at the top).
5. You'll see two options: **Delegated permissions** and **Application permissions**. Click **Application permissions**.

> **Why Application permissions?** This means the app runs on its own, without needing someone to be logged in. This is what you want for automated background tasks.

6. A search box appears. Add each of the following permissions one at a time:

---

### Permission 1: Files.ReadWrite.All

1. In the search box, type: `Files.ReadWrite`
2. A section called **Files** will expand. Check the box next to **Files.ReadWrite.All**.
   - Description shown: *Read and write files in all site collections*

---

### Permission 2: Sites.ReadWrite.All

1. Clear the search box and type: `Sites.ReadWrite`
2. A section called **Sites** will expand. Check the box next to **Sites.ReadWrite.All**.
   - Description shown: *Read and write items in all site collections*

---

### Permission 3: Mail.Send

1. Clear the search box and type: `Mail.Send`
2. A section called **Mail** will expand. Check the box next to **Mail.Send**.
   - Description shown: *Send mail as any user*

---

### Permission 4: User.Read (already added, confirm it's there)

1. Clear the search box and type: `User.Read`
2. Under **User**, check the box next to **User.Read** if it is not already checked.

---

7. After checking all four permissions, click the **Add permissions** button at the bottom of the panel.

---

## Step 7: Grant Admin Consent

Adding permissions is not enough — an administrator must "approve" them for the whole organization. This step does that.

1. Back on the API permissions page, you'll see a button near the top that says:
   **Grant admin consent for [Your Organization Name]**
2. Click that button.
3. A dialog box will appear asking: *"Do you want to grant consent for the requested permissions for all accounts in [org]?"*
4. Click **Yes**.
5. After a moment, the **Status** column next to each permission will change to a green checkmark that says **Granted for [Your Organization]**.

> If you don't see this button, your account may not have admin rights. Contact whoever manages your Microsoft 365 account.

---

## Step 8: Verify Everything

Before finishing, confirm you have all three values saved:

| Credential | Where to find it | Looks like |
|---|---|---|
| `AZURE_CLIENT_ID` | App registration > Overview > Application (client) ID | `a1b2c3d4-...` |
| `AZURE_TENANT_ID` | App registration > Overview > Directory (tenant) ID | `f9e8d7c6-...` |
| `AZURE_CLIENT_SECRET` | You copied this in Step 5 (not visible again) | `abc~defGHI...` |

Also confirm on the API permissions page that these four permissions show green checkmarks:
- `Files.ReadWrite.All` — Application
- `Sites.ReadWrite.All` — Application
- `Mail.Send` — Application
- `User.Read` — Application (or Delegated)

---

## Step 9: Add Credentials to Your .env File

Open the `.env` file in your project folder (ask your developer if you can't find it) and add these three lines:

```
AZURE_CLIENT_ID=paste-your-client-id-here
AZURE_TENANT_ID=paste-your-tenant-id-here
AZURE_CLIENT_SECRET=paste-your-client-secret-here
```

Replace the placeholder text with your actual values. Do not add quotes around the values.

---

## Troubleshooting

**"You don't have permission to register applications"**
Your account doesn't have the right role. Ask your Microsoft 365 Global Administrator to either give you the *Application Administrator* role or complete this setup themselves.

**"Admin consent button is grayed out"**
Same issue — you need admin rights. Contact your Microsoft 365 admin.

**"I left the page and now I can't see my client secret"**
Go back to Certificates & secrets, click the three dots next to the old secret, click **Delete**, then create a new one following Step 5 again.

**"I need to find this app registration again later"**
Go to portal.azure.com > search "App registrations" > click "All applications" tab > find "Mownur Services Automation".

---

## Security Reminders

- Never paste your `AZURE_CLIENT_SECRET` into a Teams message, email, or document.
- Store it only in your `.env` file or a password manager.
- Set a calendar reminder to renew the secret in 23 months.
- If you ever think the secret was exposed, go back to Certificates & secrets, delete the old one, and create a new one immediately.
