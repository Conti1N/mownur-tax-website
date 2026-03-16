# Power Automate Flow Setup Guide
## Auto-Update Client Status When TaxWise Acknowledgement Emails Arrive

**Who this is for:** A non-technical tax professional who has never used Power Automate before.
**What this flow does:** Every time an IRS or state acknowledgement email arrives in Outlook from TaxWise, this flow automatically finds the matching client in your Excel tracker and changes their status to "Filed".
**Time required:** 30–45 minutes.
**What you need before you start:**
- A Microsoft 365 account (Outlook, OneDrive, Excel)
- Your client tracker Excel file saved in OneDrive (not on your desktop — it must be in OneDrive)
- A sample TaxWise acknowledgement email in your inbox to test with

---

## What Is Power Automate?

Think of Power Automate like a set of automatic rules for your computer. You describe a trigger ("when this happens") and an action ("do this"). Once you set it up, it runs in the background every time that trigger fires — no clicking required.

For this flow: **Trigger** = a new email arrives matching TaxWise acknowledgement patterns → **Action** = find the client in Excel and update their status.

---

## Before You Start: Prepare Your Excel File

Your Excel client tracker must be in OneDrive and formatted as an official Excel Table (not just rows of data). Here's how to check and fix that:

1. Open your Excel file from OneDrive (open it in the browser at office.com, or in the desktop app with the file synced from OneDrive).
2. Click on any cell in your data.
3. In the top menu, click **Insert** → **Table**.
4. A dialog will ask "Where is the data for your table?" — confirm the range looks correct and check the box **My table has headers**.
5. Click **OK**. Your data will now have a striped/colored style with dropdown arrows in the header row.
6. The table needs a name. Click anywhere inside the table, then in the top menu look for the **Table Design** tab (it may appear only when you're inside the table). Click it and find **Table Name** on the left. Change it from something like "Table1" to: `ClientTracker`
7. Make sure your table has at least these columns (the exact names matter for later steps):
   - `Client Name` (or First Name / Last Name columns)
   - `Status` (this is the column the flow will update)
   - Some identifier for SSN last 4 digits if you use that in TaxWise emails
8. Save and close the file.

---

## Step 1: Open Power Automate

1. Open your web browser.
2. Go to: **https://make.powerautomate.com**
3. Sign in with your Microsoft 365 account.
4. You'll land on the Power Automate home page. The left sidebar has options like *Home*, *My flows*, *Create*, *Templates*, etc.

---

## Step 2: Start a New Flow

1. In the left sidebar, click **Create**.
2. You'll see several options. Click **Automated cloud flow**.
   - "Automated" means it starts automatically when something happens (your email trigger).
   - "Cloud flow" means it runs in Microsoft's cloud, not on your computer.
3. A dialog box appears asking you to name your flow and choose a trigger.

**Flow name:**
Type: `TaxWise Ack — Update Client Status`

**Choose your flow's trigger:**
- In the search box, type: `When a new email arrives`
- You'll see a result: **When a new email arrives (V3)** with the Outlook logo.
- Click on it to select it.
- Click **Create**.

---

## Step 3: Configure the Email Trigger

You're now in the flow editor. It looks like a diagram with boxes connected by arrows. You'll see one box at the top labeled **When a new email arrives (V3)**. Click on it to expand its settings.

You'll see several fields. Fill them out as follows:

**Folder:**
- Click the folder icon or the dropdown.
- Select **Inbox**.
- (If TaxWise acknowledgements go to a specific subfolder, navigate to that folder instead.)

**Include Attachments:**
- Set to **No** (acknowledgement emails usually don't need attachments for this flow).

**Only with Attachments:**
- Set to **No**.

Now click **Show advanced options** (a small link at the bottom of this box). More fields will appear:

**From:**
- Type the email address that TaxWise or the IRS sends acknowledgements from.
- If you're not sure, open a sample acknowledgement email in Outlook and look at the "From" address. It might look like `noreply@taxwise.com` or `irs-acks@irs.gov`.
- Paste that address here. This ensures the flow only runs for these specific emails, not every email you receive.

**Subject Filter:**
- Type a keyword that appears in every acknowledgement email subject line.
- Common examples: `Acknowledgement`, `Accepted`, `e-File Confirmation`
- Check a sample email's subject line and use a word that always appears there.
- Example: `Acknowledgement`

**Importance:**
- Leave as **Any**.

Leave all other fields at their defaults. Click somewhere outside the box to collapse it.

---

## Step 4: Add a "Parse" Step to Read the Email Content

Now you need to tell the flow how to extract the client's name from the email body. This uses a built-in action called **Initialize variable** to hold the name you extract.

### Step 4a: Add a Compose action to work with the email body

1. Below the trigger box, click the **+** button (it says "New step" or shows a plus sign).
2. In the search bar that appears, type: `Compose`
3. Click on **Compose** (under Data Operation).
4. In the **Inputs** field, click inside it.
5. A blue panel called **Dynamic content** will appear on the right side of the screen. This shows variables from previous steps.
6. Find and click **Body** (this is the full text of the email).
7. Rename this step by clicking the three dots (...) at the top right of the box → **Rename** → type: `Email Body`

### Step 4b: Initialize a variable for Client Name

1. Click **+ New step**.
2. Search for: `Initialize variable`
3. Click **Initialize variable**.
4. Fill in:
   - **Name:** `ClientName`
   - **Type:** String
   - **Value:** Leave blank for now.
5. Rename this step: `Store Client Name`

---

## Step 5: Extract the Client Name from the Email

TaxWise acknowledgement emails typically include the client's name in a consistent format. For example: *"Return for: John Smith has been accepted."*

You'll use a **Set variable** action to extract this.

> **Note:** The exact extraction method depends on how your TaxWise emails are formatted. Open a sample acknowledgement email and look for a consistent pattern around the client name. Common formats:
> - `Return for: FirstName LastName`
> - `Taxpayer: LastName, FirstName`
> - `Client: FirstName LastName`

### Option A: If the name appears after a consistent label like "Return for:"

1. Click **+ New step**.
2. Search for: `Set variable`
3. Click **Set variable**.
4. **Name:** Select `ClientName` from the dropdown.
5. **Value:** Click in the field, then in the Dynamic content panel click **Expression** (a tab at the top of the panel).
6. In the expression editor, type a formula to extract the name. Here is a template — replace `"Return for: "` with whatever label precedes the name in your emails:

```
trim(first(skip(split(outputs('Email_Body'), 'Return for: '), 1)))
```

What this does: splits the email text at the phrase "Return for: ", takes everything after it, then trims extra spaces.

> If this feels too technical, skip to **Option B** below.

### Option B: Simpler — Use the email Subject line directly

Many TaxWise acknowledgements include the client name in the subject line. For example: *"E-File Acknowledgement — John Smith"*

1. Click **+ New step** → search **Set variable** → click it.
2. **Name:** Select `ClientName`.
3. **Value:** Click in the field, then in Dynamic content find **Subject** (from the trigger step).
4. This sets the full subject line as the name. You may want to clean it up with an expression later, but for now this gives you something to test with.

---

## Step 6: Find the Client Row in Excel

Now tell the flow to look up the client in your Excel tracker.

1. Click **+ New step**.
2. Search for: `List rows present in a table`
3. Click **List rows present in a table** (Excel Online — Business).
4. Fill in the fields:

**Location:**
- Click the dropdown and select **OneDrive for Business**.

**Document Library:**
- Select **OneDrive**.

**File:**
- Click the folder icon. Navigate through your OneDrive folders to find your client tracker Excel file.
- Click on it to select it.

**Table:**
- Click the dropdown. It should automatically detect the tables in your Excel file.
- Select **ClientTracker** (the table you named in the preparation step).

**Filter Query (Advanced Options):**
- Click **Show advanced options**.
- In the **Filter Query** field, type a formula to find the matching row. Replace `Client Name` with your actual column header name:

```
`Client Name` eq '@{variables('ClientName')}'
```

> This tells Excel: "Find rows where the Client Name column matches the name we extracted from the email."

Rename this step: `Find Client in Excel`

---

## Step 7: Update the Client's Status to "Filed"

Now update the row you found.

1. Click **+ New step**.
2. Search for: `Update a row`
3. Click **Update a row** (Excel Online — Business).
4. Fill in:

**Location:** OneDrive for Business
**Document Library:** OneDrive
**File:** Same Excel file as above
**Table:** ClientTracker

**Key Column:**
- Select the column that uniquely identifies each row. This is usually something like `Client Name` or an ID column.

**Key Value:**
- Click in the field.
- In Dynamic content, look under **Find Client in Excel** for the column you chose above (e.g., `Client Name`).
- Click it.

**Status:**
- In the additional column fields that appear below, find **Status**.
- Type: `Filed`

Rename this step: `Update Status to Filed`

---

## Step 8: Add a Teams Notification (Optional but Recommended)

This step sends you a Teams message confirming the update happened.

1. Click **+ New step**.
2. Search for: `Post a message in a chat or channel`
3. Click **Post a message in a chat or channel** (Microsoft Teams).
4. Fill in:

**Post as:** Flow bot
**Post in:** Chat with Flow bot (or choose a specific Teams channel)
**Recipient:** Your own email address
**Message:**
Click in the field and type:

```
Status updated to "Filed" for client:
```

Then in Dynamic content, click on `ClientName` to insert the variable.

So the full message will say: *Status updated to "Filed" for client: John Smith*

Rename this step: `Notify — Client Filed`

---

## Step 9: Save and Test the Flow

### Save the flow

1. Click the **Save** button at the top right of the page.
2. If there are errors, they'll be highlighted in red. Common ones:
   - Missing connection: Click the step with an error and sign in to connect your Outlook, Excel, or Teams account.
   - Missing required field: Fill in any field shown in red.

### Test the flow manually

1. Click **Test** (top right, near the Save button).
2. Select **Manually** → click **Test**.
3. Now go to your Outlook inbox and forward a real TaxWise acknowledgement email to yourself, or find one already in your inbox.
4. The test will run. You'll see each step turn green (success) or red (failure) in real time.

### If a step fails:

- Click on the failed step (red box) to see the error message.
- Common issues and fixes:

| Error | Likely Cause | Fix |
|---|---|---|
| "Unauthorized" | Not connected to Outlook/Excel | Click the step, sign in |
| "Table not found" | Excel table not named correctly | Go back to Excel, confirm table name is `ClientTracker` |
| "No rows found" | Name extraction didn't match | Check the Filter Query formula — try printing the ClientName variable with a Compose step to see what it contains |
| "File not found" | Excel file path changed | Re-select the file in the affected step |

---

## Step 10: Enable the Flow for Live Use

Once the test passes successfully:

1. Go back to the flow overview page (click the back arrow or the flow name in the breadcrumb at the top).
2. Confirm the flow shows **On** status (toggle switch near the top right).
3. If it says **Off**, click the toggle to turn it on.

The flow is now live. Every time a TaxWise acknowledgement email arrives in your inbox, it will automatically update the client's status in Excel.

---

## Maintenance and Monitoring

### Checking if the flow is running

1. Go to make.powerautomate.com.
2. Click **My flows** in the left sidebar.
3. Click on **TaxWise Ack — Update Client Status**.
4. Scroll down to see **Run history**. Each run shows the date, whether it succeeded or failed, and how long it took.

### If runs start failing after working before

Common causes:
- Your Outlook or Excel connection expired. Click on the failed run, find the red step, and re-authenticate.
- The Excel file was moved or renamed in OneDrive. Update the file path in the affected steps.
- TaxWise changed the format of their acknowledgement emails. Review the email text and update your extraction expression in Step 5.

### Adding more status values later

If you want the flow to also catch state acknowledgements or rejection emails and set different statuses (like "Rejected"), you can add a **Condition** step after Step 4 that branches based on keywords in the subject line, routing to different **Update a row** steps with different status values.

---

## Quick Reference: The Full Flow Sequence

```
TRIGGER: New email from TaxWise in Inbox matching subject "Acknowledgement"
    ↓
ACTION: Compose — capture email body
    ↓
ACTION: Initialize variable — ClientName (empty string)
    ↓
ACTION: Set variable — extract client name from email body
    ↓
ACTION: List rows — search Excel ClientTracker for matching client
    ↓
ACTION: Update a row — set Status = "Filed"
    ↓
ACTION: Post Teams message — "Status updated to Filed for [ClientName]"
```

Total steps: 6 actions + 1 trigger = 7 steps in the flow.
