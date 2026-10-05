# Batch Mailer

Send personal emails to a list, from your own mailbox: Outlook / Microsoft 365, Gmail, Zoho and more.
Replies come back to your normal inbox.

---

## Install

No admin rights and no Python needed. It takes 1–2 minutes and needs internet the first time.

### Windows
1. Click **Start**, type **PowerShell**, and press **Enter**.
2. Copy this line, paste it into the PowerShell window (**Ctrl+V**), and press **Enter**:
   ```powershell
   irm https://raw.githubusercontent.com/mkarasneh07/batch-mailer/main/install.ps1 | iex
   ```
3. Wait until it says **Done**. Batch Mailer opens in your browser, and a **Batch Mailer** icon appears on your desktop and in the Start menu.

### Mac
1. Open **Terminal** (press **Cmd+Space**, type *Terminal*, press **Enter**).
2. Paste this line and press **Enter**:
   ```bash
   curl -LsSf https://raw.githubusercontent.com/mkarasneh07/batch-mailer/main/install.sh | bash
   ```
3. Next time, double-click **Batch Mailer** on your Desktop.

### Update
Run the same line again. Your connected email, sent history and do-not-email list are kept.

---

## How to use it

1. Open **Batch Mailer** from the desktop icon. It opens in your web browser.
2. **First time only:** type your email address and click **Continue**. The tool recognizes your email provider:
   - **Microsoft 365 / Outlook on Windows:** click **Use Outlook on this computer**. No password, no sign-in.
   - **Google, Zoho, Yahoo and others:** enter your name and password. If it asks for an *app password*, click the button it shows; it opens the right page.
3. **Your list:** upload an Excel or CSV file with one row per person. See [`sample_leads.csv`](sample_leads.csv) for the format.
4. **Your message:** write it once. Type `{{FirstName}}` to add each person's first name. Any column works the same way, and `{{FirstName|there}}` adds a backup word for empty cells.
5. **Check it:** look at the preview, then click **Send this as a test to myself**.
6. **Send:** click **Start sending**. Keep the browser tab open until it says *Finished*.

**Someone replied "no"?** Add their address to the **Do-not-email list** at the bottom of the page. They'll never be emailed again.

**Stopped halfway?** Open the tool and send again. People already emailed are skipped automatically.

---

## Outlook and Microsoft 365

Batch Mailer sends through the **classic Outlook app** on Windows (it has a **File** menu at the top left).
- If Outlook shows a **"New Outlook"** switch at the top right, turn it off. The new Outlook can't be used by other programs.
- Keep Outlook open while sending. Emails leave from the Outbox and appear in Sent Items.
- If Outlook asks whether to allow a program to send email, click **Allow**.

<details>
<summary>No classic Outlook (for example on a Mac)? Microsoft sign-in setup for IT</summary>

Done once, by an admin of the company's Microsoft 365.
1. Go to https://entra.microsoft.com → search **App registrations** → **New registration**.
2. Name: *Batch Mailer*. Account types: *Accounts in this organizational directory only*.
   Redirect URI: **Public client/native (mobile & desktop)**, value `http://localhost`. Click **Register**.
3. Copy the **Application (client) ID** and **Directory (tenant) ID** from the Overview page.
4. **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated** → tick `Mail.Send` → **Add**, then **Grant admin consent**.
5. In Batch Mailer: left panel → **Admin settings** → paste the client ID, replace `common` with the tenant ID → **Save admin settings**.
</details>

## Gmail and Google Workspace

The tool walks you through creating an **app password**. Your Google account needs 2-Step Verification turned on. If your company's Google admin has turned app passwords off, the admin needs to allow them.

---

## Your data stays on your computer

Nothing is sent anywhere except the emails themselves. Passwords are kept in Windows Credential Manager or the Mac Keychain, never in a file.

| What | File |
|---|---|
| Which mailbox is connected | `account.json` |
| Every email sent | `sent_log.csv` |
| Do-not-email list | `do_not_contact.txt` |
| Last message written | `last_template.json` |

These live in `%LOCALAPPDATA%\BatchMailer` on Windows and `~/BatchMailer` on Mac.

## Uninstall

Delete the **Batch Mailer** folder above and the **Batch Mailer** shortcut. That's everything.

## Good sending habits

- For people who don't know you, keep it to 30–50 emails a day per mailbox (left panel → *Most emails per day*).
- Write short, plain emails without attachments, and always include a line like "Reply 'no' and I won't email again."
