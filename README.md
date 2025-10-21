# Invoice Generator

A professional invoice generator that creates formatted invoices in Google Sheets with automated Canadian holiday detection and configurable billing data.

## Features

- **Professional Layout**: Clean, formatted invoice template with company branding
- **Google Sheets Integration**: Direct creation and formatting in Google Sheets
- **JSON Configuration**: Easy-to-modify invoice data via `invoice_config.json`
- **Canadian Holiday Detection**: Automatically identifies statutory holidays in billing periods
- **Dynamic Dates**: Auto-generates invoice dates and descriptions
- **Batch Processing**: Optimized API calls for fast execution

## Quick Start

1. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

2. **Configure credentials**: Place `invoicepacksize-google-creds.json` in the project directory

3. **Run the generator**:
   ```bash
   python invoice_optimized.py
   ```

## Configuration

Edit `invoice_config.json` to customize:

```json
{
  "company": {
    "name": "Motyer Corp",
    "service_description": "Software Engineering Consulting"
  },
  "invoice": {
    "number": 13,
    "po_number": "PO-2024-001",
    "tax_rate": 0.0
  },
  "client": {
    "name": "Packsize LLC"
  },
  "line_items": [
    {
      "description": "Week of April 13",
      "quantity": 40,
      "unit_price": 28.85
    }
  ]
}
```

## Output

The script generates:
- A new Google Sheet with professional invoice formatting
- Automatic Canadian holiday detection in notes section
- Calculated totals and tax amounts
- Publicly accessible invoice URL

## Scheduling

### Windows Task Scheduler
```bash
schtasks /create /tn "Invoice Generator" /tr "C:\Python313\python.exe C:\Users\rlmot\OneDrive\Documents\scripts\invoice\invoice_optimized.py" /sc weekly /d SUN /mo 2 /st 09:00
```

### GitHub Actions (Cloud)
Create `.github/workflows/invoice.yml`:
```yaml
name: Generate Invoice
on:
  schedule:
    - cron: '0 9 */14 * *'  # Every 2 weeks at 9 AM
jobs:
  generate:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v2
      - name: Setup Python
        uses: actions/setup-python@v2
        with:
          python-version: '3.11'
      - name: Install dependencies
        run: pip install -r requirements.txt
      - name: Generate Invoice
        run: python invoice_optimized.py
        env:
          GOOGLE_CREDENTIALS: ${{ secrets.GOOGLE_CREDENTIALS }}
```

## Dependencies

- `gspread` - Google Sheets API
- `google-auth` - Authentication
- `holidays` - Canadian holiday detection

## License

MIT License