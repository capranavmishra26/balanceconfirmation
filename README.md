# Vendor Balance Confirmation

Sends balance confirmation letters to vendors through **Outlook**, using the SAP open-item dumps for
**Advance to Vendor** (G/L 2220004) and **Creditors** (1200000–1200020).

For each vendor the app:
1. Nets the advance (debit) and creditor (credit) balances into **one net balance**.
2. Creates a PDF letter in the company format (*Confirmation of Outstanding Balance as at …*).
3. Emails the PDF through the Outlook desktop app, then logs what was sent.

## No-Python version (index.html)
Copy **`index.html`** and **`Outlook_Connector.bat`** into one folder on any Windows laptop.
- Double-click `index.html` to use the app in Edge/Chrome (offline). In the preview, **Email this vendor** offers
  *Letter + text* (Windows share menu with the PDF), *Text only* (opens the mail app) and *Outlook draft* (.eml with
  To/subject/letter filled in). **Download letters in one go** gives one combined PDF or a ZIP.
- To send in bulk straight from the page, double-click `Outlook_Connector.bat` instead; the app opens already connected
  to Outlook. Letters and `sent_log.csv` are saved in `output\` next to it.
- To change the app, edit `web/app.html` / `web/connector.ps1` and run `powershell -ExecutionPolicy Bypass -File web\build.ps1`.

## Requirements
- Windows with **Microsoft Outlook desktop**, signed in to the mailbox you send from
- Python 3.10+ (`setup.bat` installs it through `winget` if it's missing)

## Install
```bat
git clone https://github.com/<your-account>/balance-confirmation-app.git
cd balance-confirmation-app
setup.bat
```
(You can also use **Code → Download ZIP** on GitHub, extract it, and double-click `setup.bat`.)

## Run
Double-click **`run.bat`**. The app opens in your browser at http://localhost:8501.

## Monthly use
1. **Export from SAP** two line-item dumps with the standard layout (columns include `Vendor`,
   `Vendor Account: Name 1`, `Company Code Currency Value`):
   - Advance to Vendor open items
   - Creditor open items
2. **Upload** both dumps, plus the **vendor master** (Excel with `Vendor Code`, `Email`, and optionally `CC`;
   separate multiple emails with `;`).
   The first time, click **Vendor master template**, fill in the emails, and keep that file for future months.
3. In the sidebar, set **Balance as at**, **Letter date**, the minimum balance, and the Outlook account to send from.
4. **Review** the table. Untick vendors you don't want to send to; you can correct emails right in the table.
5. **Preview** the email and download the PDF letter for any vendor.
6. **Test first:** put your own address in *Test mode*, send a few, and check them.
7. Clear *Test mode*, tick the confirmation box, and click **Send emails**.

Letters are saved to `output/<as-at date>/run_<timestamp>/`. Every send is appended to `output/sent_log.csv`.

## Letter wording
| Net balance | Confirmation line |
|---|---|
| Credit (we owe the vendor) | The balance **receivable from us** of ₹ X as at … |
| Debit (advance due from vendor) | The balance **payable by us to you** of ₹ X as at … |

To change the company name, address, salutation, wording, or the email subject/body, edit **`config.json`**.
Placeholders you can use in the email: `{vendor_name}` `{vendor_code}` `{as_at}` `{amount}` `{nature}` `{company_name}`.

## Data safety
`.gitignore` excludes all Excel/CSV files and the `output/` folder, so SAP data, vendor emails and letters
are never pushed to GitHub. Keep the repository **private**.
