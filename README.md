## Setup

1. Create a 'config.json' file with your Google Sheet ID and path to your credentials file:

```json
{
    "sheet_id": "your_google_sheet_id",
    "credentials_file": "credentials.json"
}
```
## How it Works

1. Setup a Google form
2. Connect the form to a Google Sheets workbook
3. Integrate the data from Google Sheets workbook to a Microsoft workbook
4. Setup a template of an invoice with the relevant placeholders
5. Take the data from the Microsoft workbook and use it to fill up invoice template
6. Save the invoice template into an invoice folder
