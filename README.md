# Circle K Club sign-up

A one-page sign-up form for the club roster. Each sign-up becomes one row in a Google Sheet, one answer per column: role, last name, first name, date of birth, grade, email, phone number.

- `index.html` is the form. GitHub Pages serves it.
- `Code.gs` is the Google Apps Script that lives in the roster spreadsheet and writes the rows.

## Connecting the form to the sheet

1. Open the roster spreadsheet, then **Extensions → Apps Script**.
2. Replace the contents of `Code.gs` there with this repo's `Code.gs`.
3. **Deploy → New deployment → Web app**. Execute as **Me**, access **Anyone**.
4. Copy the web app address (it ends in `/exec`) into `ENDPOINT` near the top of the script in `index.html`.

After changing `Code.gs` later, use **Deploy → Manage deployments → Edit → New version** so the address stays the same.
