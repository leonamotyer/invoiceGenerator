# Invoice Generator for Google Sheets

This script creates professional invoices in Google Sheets using the same structure and approach as the rest_test.py script, but specifically designed for invoice generation.

## Features

- **Professional Invoice Layout**: Matches the formatting from the provided invoice image
- **Google Sheets Integration**: Uses the invoicepacksize Google credentials
- **Configurable Invoice Data**: Easy to modify company info, client details, and line items
- **Automatic Calculations**: Handles subtotals, taxes, and totals
- **Professional Styling**: Orange company name, proper formatting, and currency display
- **Canadian Holiday Detection**: Automatically detects and displays Canadian statutory holidays in the notes section
- **Dynamic Date Descriptions**: Invoice descriptions automatically update based on system date

## Setup

1. **Install Dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

2. **Credentials**: Ensure `invoicepacksize-google-creds.json` is in the same directory as the script

3. **Run the Script**:
   ```bash
   python invoice.py
   ```

## Configuration

The invoice data is configured in the `INVOICE_CONFIG` dictionary at the top of the script:

```python
INVOICE_CONFIG = {
    'company_name': 'Motyer Corp',
    'address': '8140 76th ave NW',
    'city_province': 'Edmonton AB',
    'phone': '(250) 258-1143',
    'client_name': 'Packsize LLC',
    'currency': 'CAD',
    'currency_symbol': '$',
    'tax_rate': 0.0,  # Set to 0.13 for 13% GST
    'invoice_number': 13,
    'invoice_date': '4/24/2026',
    'line_items': [
        {
            'description': 'Week of April 13',
            'quantity': 40,
            'unit_price': 28.85,
            'total_price': 1154.00
        },
        # Add more line items as needed
    ]
}
```

## Usage Examples

### Basic Usage
```python
python invoice.py
```

### Programmatic Usage
```python
from invoice import update_invoice_config, add_line_item, main

# Update invoice details
update_invoice_config(
    invoice_number=14,
    client_name="New Client LLC",
    invoice_date="5/1/2026"
)

# Add a new line item
add_line_item("Week of April 27", 40, 28.85)

# Generate the invoice
main()
```

## Invoice Layout

The script creates an invoice with the following structure:

1. **Company Header** (Rows 3-6): Company name, address, city/province, phone
2. **Invoice Title** (Rows 8-9): "Invoice" title and submission date
3. **Invoice Details** (Rows 11-12): Client info, payable to, invoice number
4. **Line Items Table** (Rows 18+): Description, quantity, unit price, total price
5. **Summary** (Rows 24-26): Notes, subtotal, tax (if applicable), total

## Features

- **Automatic Spreadsheet Creation**: Creates a new Google Sheet for each invoice
- **Professional Formatting**: Orange company name, bold headers, proper alignment
- **Currency Formatting**: Proper currency symbols and decimal places
- **Tax Calculation**: Optional tax calculation (set `tax_rate` to enable)
- **Column Width Optimization**: Automatically sets appropriate column widths
- **Public Access**: Makes the spreadsheet publicly viewable (optional)

## Customization

### Adding Line Items
```python
add_line_item("Service Description", quantity, unit_price)
```

### Updating Configuration
```python
update_invoice_config(
    company_name="Your Company",
    client_name="Client Name",
    tax_rate=0.13  # 13% tax
)
```

### Using Existing Spreadsheet
To use an existing spreadsheet instead of creating a new one, modify the `connect_to_google_sheets()` function:

```python
# Replace the create() call with:
spreadsheet = gc.open_by_key('YOUR_SPREADSHEET_ID')
```

## Output

The script will:
1. Create a new Google Sheet with the invoice
2. Print the spreadsheet URL
3. Display invoice totals and summary
4. Make the sheet publicly viewable (optional)

## Error Handling

The script includes comprehensive error handling for:
- Missing credentials file
- Invalid credentials format
- Google Sheets API errors
- Network connectivity issues

## Canadian Holiday Detection

The invoice generator automatically detects Canadian statutory holidays within the invoice period (11 days ago to 4 days ago) and includes them in the notes section. This helps document paid days off without affecting the hours calculation.

### How It Works

1. **Date Range**: Checks the period from 11 days ago to 4 days ago
2. **Holiday Detection**: Uses the `holidays` package to identify Canadian statutory holidays
3. **Notes Section**: Displays holidays in the invoice notes with clear labeling
4. **Hours Unaffected**: Holiday detection is for documentation only - hours remain the same

### Example Output

When holidays are detected, the notes section will show:
```
📝 Notes

🇨🇦 Canadian Statutory Holidays (Paid Days Off):
• December 25: Christmas Day
• December 26: Boxing Day

Note: Holiday hours are included in regular billing.
```

## Dependencies

- `gspread`: Google Sheets API wrapper
- `google-auth`: Google authentication
- `gspread-formatting`: Advanced formatting for Google Sheets
- `holidays`: Canadian statutory holiday detection

See `requirements.txt` for specific versions.
# invoiceGenerator
