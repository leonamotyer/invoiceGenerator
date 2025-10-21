# Expert-Level Invoice Generator - Optimized Architecture
# Uses template-based approach with minimal API calls and JSON configuration
import json
import datetime
import gspread
import holidays
from google.oauth2.service_account import Credentials
from typing import Dict, List, Any
from dataclasses import dataclass

# Layout constants (only the ones actually used)
MAX_ROWS = 30
MAX_COLS = 6

# Default formatting values
DEFAULT_ORANGE_COLOR = {'red': 0.706, 'green': 0.373, 'blue': 0.024}
DEFAULT_FONT_FAMILY = 'Roboto'
DEFAULT_COMPANY_NAME_SIZE = 20
DEFAULT_INVOICE_TITLE_SIZE = 20
DEFAULT_TOTAL_SIZE = 20
DEFAULT_REGULAR_TEXT_SIZE = 10
DEFAULT_LABEL_TEXT_SIZE = 12

@dataclass
class InvoiceData:
    """Structured invoice data"""
    company_name: str
    service_description: str
    client_name: str
    invoice_number: int
    po_number: str
    invoice_date: str
    line_items: List[Dict[str, Any]]
    currency_symbol: str = '$'
    tax_rate: float = 0.0

class InvoiceConfigLoader:
    """Load and process invoice configuration from JSON"""
    
    def __init__(self, config_file: str = 'invoice_config.json'):
        self.config_file = config_file
        self.config = self._load_config()
    
    def _load_config(self) -> Dict[str, Any]:
        """Load configuration from JSON file"""
        try:
            with open(self.config_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except FileNotFoundError:
            print(f"Error: {self.config_file} not found!")
            print("Please create invoice_config.json with your invoice data.")
            exit(1)
        except json.JSONDecodeError as e:
            print(f"Error parsing {self.config_file}: {e}")
            exit(1)
    
    def get_invoice_data(self) -> InvoiceData:
        """Convert JSON config to InvoiceData object"""
        # Calculate invoice date
        invoice_date = datetime.datetime.now() + datetime.timedelta(days=self.config['invoice'].get('date_offset_days', 0))
        
        # Process line items with dynamic dates
        processed_line_items = []
        for item in self.config['line_items']:
            # Calculate week date
            week_date = datetime.datetime.now() + datetime.timedelta(days=item['week_offset'])
            week_date_str = week_date.strftime('%B %d')
            
            # Replace placeholder in description
            description = item['description'].replace('{week1_date}', week_date_str).replace('{week2_date}', week_date_str)
            
            processed_line_items.append({
                'description': description,
                'quantity': item['quantity'],
                'unit_price': item['unit_price']
            })
        
        return InvoiceData(
            company_name=self.config['company']['name'],
            service_description=self.config['company']['service_description'],
            client_name=self.config['client']['name'],
            invoice_number=self.config['invoice']['number'],
            po_number=self.config['invoice']['po_number'],
            invoice_date=invoice_date.strftime('%m/%d/%Y'),
            line_items=processed_line_items,
            currency_symbol=self.config['invoice']['currency_symbol'],
            tax_rate=self.config['invoice']['tax_rate']
        )
    
    def get_formatting_config(self) -> Dict[str, Any]:
        """Get formatting configuration"""
        return self.config.get('formatting', {})
    
    def get_notes_config(self) -> Dict[str, Any]:
        """Get notes configuration"""
        return self.config.get('notes', {})

class InvoiceTemplate:
    """Template-based invoice generator with minimal API calls"""
    
    def __init__(self, spreadsheet_id: str, credentials_file: str, config_loader: InvoiceConfigLoader = None):
        self.spreadsheet_id = spreadsheet_id
        self.credentials_file = credentials_file
        self.config_loader = config_loader or InvoiceConfigLoader()
        self.gc = None
        self.spreadsheet = None
        
    def connect(self) -> None:
        """Single connection setup"""
        SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
        creds = Credentials.from_service_account_file(self.credentials_file, scopes=SCOPES)
        self.gc = gspread.authorize(creds)
        self.spreadsheet = self.gc.open_by_key(self.spreadsheet_id)
        
    def create_invoice(self, data: InvoiceData = None) -> str:
        """Create invoice with only 2 API calls total"""
        try:
            if data is None:
                data = self.config_loader.get_invoice_data()
            
            worksheet_title = f"Invoice {data.invoice_number}"
            
            # Get or create worksheet
            try:
                worksheet = self.spreadsheet.worksheet(worksheet_title)
                worksheet.clear()
            except gspread.WorksheetNotFound:
                worksheet = self.spreadsheet.add_worksheet(worksheet_title, rows=30, cols=6)
            
            # CALL 1: Batch ALL data, formatting, and formulas in one request
            self._apply_everything_in_one_batch(worksheet, data)
            
            # CALL 2: Batch merging and column widths in one request
            self._apply_merging_and_widths_in_one_batch(worksheet)
            
            print("Invoice created successfully!")
            return worksheet.url
            
        except Exception as e:
            if "429" in str(e) or "RATE_LIMIT_EXCEEDED" in str(e):
                print("\n⚠️  Rate limit exceeded. Please wait a minute and try again.")
                print("Google Sheets API allows 60 requests per minute per user.")
            else:
                print(f"\n❌ Error creating invoice: {e}")
            raise
        
    
    def _apply_everything_in_one_batch(self, worksheet, data: InvoiceData) -> None:
        """Apply ALL data, formatting, and formulas in one batch request"""
        formatting_config = self.config_loader.get_formatting_config()
        orange_color = formatting_config.get('company_name_color', DEFAULT_ORANGE_COLOR)
        font_family = formatting_config.get('font_family', DEFAULT_FONT_FAMILY)
        company_name_size = formatting_config.get('company_name_size', DEFAULT_COMPANY_NAME_SIZE)
        invoice_title_size = formatting_config.get('invoice_title_size', DEFAULT_INVOICE_TITLE_SIZE)
        total_size = formatting_config.get('total_size', DEFAULT_TOTAL_SIZE)
        regular_text_size = formatting_config.get('regular_text_size', DEFAULT_REGULAR_TEXT_SIZE)
        label_text_size = formatting_config.get('label_text_size', DEFAULT_LABEL_TEXT_SIZE)
        
        # Build notes content
        notes_content = self._build_notes_content(data)
        
        # Calculate total on code side
        total_amount = sum(item['quantity'] * item['unit_price'] for item in data.line_items)
        
        # Prepare all requests
        requests = []
        
        # 1. Hide gridlines
        requests.append({
            'updateSheetProperties': {
                'properties': {
                    'sheetId': worksheet.id,
                    'gridProperties': {'hideGridlines': True}
                },
                'fields': 'gridProperties.hideGridlines'
            }
        })
        
        # 2. Update all cell values
        requests.append({
            'updateCells': {
                'range': {'sheetId': worksheet.id, 'startRowIndex': 0, 'endRowIndex': 30, 'startColumnIndex': 0, 'endColumnIndex': 6},
                'rows': [
                    # Row 1-2: Header background
                    {'values': [{'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}]},
                    {'values': [{'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}]},
                    # Row 3: Company name
                    {'values': [{'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': data.company_name}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}]},
                    # Row 4: Service Description
                    {'values': [{'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': data.service_description}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}]},
                    # Row 5: Empty (to maintain row count)
                    {'values': [{'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}]},
                    # Row 6: Empty (removed phone)
                    {'values': [{'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}]},
                    # Row 7: Empty
                    {'values': [{'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}]},
                    # Row 8: Invoice title
                    {'values': [{'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': 'INVOICE'}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}]},
                    # Row 9: Invoice date
                    {'values': [{'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': data.invoice_date}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}]},
                    # Row 10: Empty
                    {'values': [{'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}]},
                    # Row 11: Labels
                    {'values': [{'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': 'Invoice For:'}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': 'Payable To:'}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': 'Invoice #:'}}]},
                    # Row 12: Values
                    {'values': [{'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': data.client_name}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': data.company_name}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': str(data.invoice_number)}}]},
                     # Row 13: PO Number label only
                     {'values': [{'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': 'PO #:'}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}]},
                     # Row 14: PO Number value (using existing empty row)
                     {'values': [{'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': data.po_number}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}]},
                    # Row 15: Empty
                    {'values': [{'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}]},
                    # Row 16: Empty
                    {'values': [{'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}]},
                    # Row 17: Empty
                    {'values': [{'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}]},
                    # Row 18: Table headers
                    {'values': [{'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': 'Description'}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': 'Qty'}}, {'userEnteredValue': {'stringValue': 'Unit Price'}}, {'userEnteredValue': {'stringValue': 'Amount'}}]},
                ] + [
                    # Dynamic line items (rows 19, 20, etc.)
                    {'values': [
                        {'userEnteredValue': {'stringValue': ''}},
                        {'userEnteredValue': {'stringValue': item['description']}},
                        {'userEnteredValue': {'stringValue': ''}},
                        {'userEnteredValue': {'numberValue': item['quantity']}},
                        {'userEnteredValue': {'numberValue': item['unit_price']}},
                        {'userEnteredValue': {'numberValue': item['quantity'] * item['unit_price']}}
                    ]} for i, item in enumerate(data.line_items)
                ] + [
                    # Row 19+ (after line items): Empty rows
                    {'values': [{'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}]},
                    {'values': [{'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}]},
                    {'values': [{'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}]},
                    # Row 24: Notes section
                    {'values': [{'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': notes_content}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}]},
                    # Row 25: Total label
                    {'values': [{'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': 'Total'}}, {'userEnteredValue': {'stringValue': ''}}]},
                    # Row 26: Empty
                    {'values': [{'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}]},
                    # Row 27: Total value (calculated on code side, placed in merged E27:F27)
                    {'values': [{'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'numberValue': total_amount}}, {'userEnteredValue': {'stringValue': ''}}]},
                    # Row 28-30: Empty rows
                    {'values': [{'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}]},
                    {'values': [{'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}]},
                    {'values': [{'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}, {'userEnteredValue': {'stringValue': ''}}]},
                ],
                'fields': 'userEnteredValue'
            }
        })
        
        # 3. Apply all formatting
        # Company name formatting
        requests.append({
            'repeatCell': {
                'range': {'sheetId': worksheet.id, 'startRowIndex': 2, 'endRowIndex': 3, 'startColumnIndex': 1, 'endColumnIndex': 2},
                'cell': {
                    'userEnteredFormat': {
                        'textFormat': {'bold': True, 'fontSize': company_name_size, 'fontFamily': font_family, 'foregroundColor': {'red': orange_color['red'], 'green': orange_color['green'], 'blue': orange_color['blue']}},
                        'horizontalAlignment': 'LEFT'
                    }
                },
                'fields': 'userEnteredFormat'
            }
        })
        
        # Contact info formatting
        requests.append({
            'repeatCell': {
                'range': {'sheetId': worksheet.id, 'startRowIndex': 3, 'endRowIndex': 6, 'startColumnIndex': 1, 'endColumnIndex': 2},
                'cell': {
                    'userEnteredFormat': {
                        'textFormat': {'fontSize': regular_text_size, 'fontFamily': font_family},
                        'horizontalAlignment': 'LEFT'
                    }
                },
                'fields': 'userEnteredFormat'
            }
        })
        
        # Invoice title formatting
        requests.append({
            'repeatCell': {
                'range': {'sheetId': worksheet.id, 'startRowIndex': 7, 'endRowIndex': 8, 'startColumnIndex': 1, 'endColumnIndex': 2},
                'cell': {
                    'userEnteredFormat': {
                        'textFormat': {'bold': True, 'fontSize': invoice_title_size, 'fontFamily': font_family, 'foregroundColor': {'red': orange_color['red'], 'green': orange_color['green'], 'blue': orange_color['blue']}},
                        'horizontalAlignment': 'LEFT'
                    }
                },
                'fields': 'userEnteredFormat'
            }
        })
        
        # Labels formatting (Payable To, Invoice #) - Bold, size 12
        requests.append({
            'repeatCell': {
                'range': {'sheetId': worksheet.id, 'startRowIndex': 10, 'endRowIndex': 11, 'startColumnIndex': 1, 'endColumnIndex': 2},
                'cell': {
                    'userEnteredFormat': {
                        'textFormat': {'bold': True, 'fontSize': label_text_size, 'fontFamily': font_family},
                        'horizontalAlignment': 'LEFT'
                    }
                },
                'fields': 'userEnteredFormat'
            }
        })
        
        requests.append({
            'repeatCell': {
                'range': {'sheetId': worksheet.id, 'startRowIndex': 10, 'endRowIndex': 11, 'startColumnIndex': 3, 'endColumnIndex': 4},
                'cell': {
                    'userEnteredFormat': {
                        'textFormat': {'bold': True, 'fontSize': label_text_size, 'fontFamily': font_family},
                        'horizontalAlignment': 'LEFT'
                    }
                },
                'fields': 'userEnteredFormat'
            }
        })
        
        requests.append({
            'repeatCell': {
                'range': {'sheetId': worksheet.id, 'startRowIndex': 10, 'endRowIndex': 11, 'startColumnIndex': 5, 'endColumnIndex': 6},
                'cell': {
                    'userEnteredFormat': {
                        'textFormat': {'bold': True, 'fontSize': label_text_size, 'fontFamily': font_family},
                        'horizontalAlignment': 'LEFT'
                    }
                },
                'fields': 'userEnteredFormat'
            }
        })
        
        # Values formatting (client name, invoice number) - Regular, size 10
        requests.append({
            'repeatCell': {
                'range': {'sheetId': worksheet.id, 'startRowIndex': 11, 'endRowIndex': 12, 'startColumnIndex': 1, 'endColumnIndex': 2},
                'cell': {
                    'userEnteredFormat': {
                        'textFormat': {'fontSize': regular_text_size, 'fontFamily': font_family},
                        'horizontalAlignment': 'LEFT'
                    }
                },
                'fields': 'userEnteredFormat'
            }
        })
        
        requests.append({
            'repeatCell': {
                'range': {'sheetId': worksheet.id, 'startRowIndex': 11, 'endRowIndex': 12, 'startColumnIndex': 3, 'endColumnIndex': 4},
                'cell': {
                    'userEnteredFormat': {
                        'textFormat': {'fontSize': regular_text_size, 'fontFamily': font_family},
                        'horizontalAlignment': 'LEFT'
                    }
                },
                'fields': 'userEnteredFormat'
            }
        })
        
        requests.append({
            'repeatCell': {
                'range': {'sheetId': worksheet.id, 'startRowIndex': 11, 'endRowIndex': 12, 'startColumnIndex': 5, 'endColumnIndex': 6},
                'cell': {
                    'userEnteredFormat': {
                        'textFormat': {'fontSize': regular_text_size, 'fontFamily': font_family},
                        'horizontalAlignment': 'LEFT'
                    }
                },
                'fields': 'userEnteredFormat'
            }
        })
        
        # PO Number formatting (same as invoice number) - Rows 13-14
        requests.append({
            'repeatCell': {
                'range': {'sheetId': worksheet.id, 'startRowIndex': 12, 'endRowIndex': 14, 'startColumnIndex': 1, 'endColumnIndex': 2},
                'cell': {
                    'userEnteredFormat': {
                        'textFormat': {'bold': True, 'fontSize': label_text_size, 'fontFamily': font_family},
                        'horizontalAlignment': 'LEFT'
                    }
                },
                'fields': 'userEnteredFormat'
            }
        })
        
        requests.append({
            'repeatCell': {
                'range': {'sheetId': worksheet.id, 'startRowIndex': 12, 'endRowIndex': 13, 'startColumnIndex': 2, 'endColumnIndex': 3},
                'cell': {
                    'userEnteredFormat': {
                        'textFormat': {'fontSize': regular_text_size, 'fontFamily': font_family},
                        'horizontalAlignment': 'LEFT'
                    }
                },
                'fields': 'userEnteredFormat'
            }
        })
        
        # PO Number value formatting (row 14, column B)
        requests.append({
            'repeatCell': {
                'range': {'sheetId': worksheet.id, 'startRowIndex': 13, 'endRowIndex': 14, 'startColumnIndex': 1, 'endColumnIndex': 2},
                'cell': {
                    'userEnteredFormat': {
                        'textFormat': {'fontSize': regular_text_size, 'fontFamily': font_family},
                        'horizontalAlignment': 'LEFT'
                    }
                },
                'fields': 'userEnteredFormat'
            }
        })
        
        # Table headers formatting
        requests.append({
            'repeatCell': {
                'range': {'sheetId': worksheet.id, 'startRowIndex': 17, 'endRowIndex': 18, 'startColumnIndex': 1, 'endColumnIndex': 3},
                'cell': {
                    'userEnteredFormat': {
                        'textFormat': {'bold': True, 'fontSize': label_text_size, 'fontFamily': font_family},
                        'horizontalAlignment': 'LEFT',
                        'verticalAlignment': 'MIDDLE',
                        'padding': {'top': 10, 'bottom': 10, 'left': 0, 'right': 0}
                    }
                },
                'fields': 'userEnteredFormat'
            }
        })
        
        requests.append({
            'repeatCell': {
                'range': {'sheetId': worksheet.id, 'startRowIndex': 17, 'endRowIndex': 18, 'startColumnIndex': 3, 'endColumnIndex': 6},
                'cell': {
                    'userEnteredFormat': {
                        'textFormat': {'bold': True, 'fontSize': label_text_size, 'fontFamily': font_family},
                        'horizontalAlignment': 'RIGHT',
                        'verticalAlignment': 'MIDDLE',
                        'padding': {'top': 10, 'bottom': 10, 'left': 0, 'right': 0}
                    }
                },
                'fields': 'userEnteredFormat'
            }
        })
        
        # Line items formatting
        for i in range(len(data.line_items)):
            row = 18 + i
            # Currency formatting for unit price and amount
            requests.append({
                'repeatCell': {
                    'range': {'sheetId': worksheet.id, 'startRowIndex': row, 'endRowIndex': row+1, 'startColumnIndex': 4, 'endColumnIndex': 6},
                    'cell': {
                        'userEnteredFormat': {
                            'textFormat': {'fontSize': regular_text_size, 'fontFamily': font_family},
                            'horizontalAlignment': 'RIGHT',
                            'numberFormat': {'type': 'CURRENCY', 'pattern': '$#,##0.00'}
                        }
                    },
                    'fields': 'userEnteredFormat'
                }
            })
            
            # Description formatting
            requests.append({
                'repeatCell': {
                    'range': {'sheetId': worksheet.id, 'startRowIndex': row, 'endRowIndex': row+1, 'startColumnIndex': 1, 'endColumnIndex': 2},
                    'cell': {
                        'userEnteredFormat': {
                            'textFormat': {'fontSize': regular_text_size, 'fontFamily': font_family},
                            'horizontalAlignment': 'LEFT'
                        }
                    },
                    'fields': 'userEnteredFormat'
                }
            })
            
            # Quantity formatting
            requests.append({
                'repeatCell': {
                    'range': {'sheetId': worksheet.id, 'startRowIndex': row, 'endRowIndex': row+1, 'startColumnIndex': 3, 'endColumnIndex': 4},
                    'cell': {
                        'userEnteredFormat': {
                            'textFormat': {'fontSize': regular_text_size, 'fontFamily': font_family},
                            'horizontalAlignment': 'RIGHT'
                        }
                    },
                    'fields': 'userEnteredFormat'
                }
            })
        
        # Alternating row backgrounds
        for i in range(5):  # Rows 18-22
            row = 17 + i
            bg_color = {'red': 1.0, 'green': 1.0, 'blue': 1.0} if i % 2 == 0 else {'red': 0.9, 'green': 0.9, 'blue': 0.9}
            requests.append({
                'repeatCell': {
                    'range': {'sheetId': worksheet.id, 'startRowIndex': row, 'endRowIndex': row+1, 'startColumnIndex': 1, 'endColumnIndex': 6},
                    'cell': {
                        'userEnteredFormat': {
                            'backgroundColor': bg_color
                        }
                    },
                    'fields': 'userEnteredFormat.backgroundColor'
                }
            })
        
        # Borders
        requests.append({
            'repeatCell': {
                'range': {'sheetId': worksheet.id, 'startRowIndex': 16, 'endRowIndex': 17, 'startColumnIndex': 1, 'endColumnIndex': 6},
                'cell': {
                    'userEnteredFormat': {
                        'borders': {'top': {'style': 'SOLID', 'width': 1}}
                    }
                },
                'fields': 'userEnteredFormat.borders'
            }
        })
        
        requests.append({
            'repeatCell': {
                'range': {'sheetId': worksheet.id, 'startRowIndex': 22, 'endRowIndex': 23, 'startColumnIndex': 1, 'endColumnIndex': 6},
                'cell': {
                    'userEnteredFormat': {
                        'borders': {'bottom': {'style': 'SOLID', 'width': 1}}
                    }
                },
                'fields': 'userEnteredFormat.borders'
            }
        })
        
        # Total formatting
        requests.append({
            'repeatCell': {
                'range': {'sheetId': worksheet.id, 'startRowIndex': 24, 'endRowIndex': 25, 'startColumnIndex': 4, 'endColumnIndex': 5},
                'cell': {
                    'userEnteredFormat': {
                        'textFormat': {'bold': True, 'fontSize': label_text_size, 'fontFamily': font_family},
                        'horizontalAlignment': 'LEFT'
                    }
                },
                'fields': 'userEnteredFormat'
            }
        })
        
        requests.append({
            'repeatCell': {
                'range': {'sheetId': worksheet.id, 'startRowIndex': 26, 'endRowIndex': 27, 'startColumnIndex': 4, 'endColumnIndex': 6},
                'cell': {
                    'userEnteredFormat': {
                        'textFormat': {'bold': True, 'fontSize': total_size, 'fontFamily': font_family, 'foregroundColor': {'red': orange_color['red'], 'green': orange_color['green'], 'blue': orange_color['blue']}},
                        'horizontalAlignment': 'RIGHT',
                        'numberFormat': {'type': 'CURRENCY', 'pattern': '$#,##0.00'}
                    }
                },
                'fields': 'userEnteredFormat'
            }
        })
        
        # Notes formatting
        requests.append({
            'repeatCell': {
                'range': {'sheetId': worksheet.id, 'startRowIndex': 23, 'endRowIndex': 30, 'startColumnIndex': 1, 'endColumnIndex': 4},
                'cell': {
                    'userEnteredFormat': {
                        'textFormat': {'fontSize': regular_text_size, 'fontFamily': font_family},
                        'horizontalAlignment': 'LEFT',
                        'verticalAlignment': 'TOP',
                        'wrapStrategy': 'WRAP'
                    }
                },
                'fields': 'userEnteredFormat'
            }
        })
        
        # Header background
        requests.append({
            'repeatCell': {
                'range': {'sheetId': worksheet.id, 'startRowIndex': 0, 'endRowIndex': 2, 'startColumnIndex': 0, 'endColumnIndex': 6},
                'cell': {
                    'userEnteredFormat': {
                        'backgroundColor': {'red': orange_color['red'], 'green': orange_color['green'], 'blue': orange_color['blue']}
                    }
                },
                'fields': 'userEnteredFormat.backgroundColor'
            }
        })
        
        # Execute all requests in one batch
        worksheet.spreadsheet.batch_update({'requests': requests})
    
    def _create_text_format_request(self, worksheet, start_row: int, end_row: int, start_col: int, end_col: int, 
                                  font_size: int, font_family: str, bold: bool = False, 
                                  color: Dict[str, float] = None, alignment: str = 'LEFT') -> Dict:
        """Helper method to create text formatting requests"""
        text_format = {
            'fontSize': font_size,
            'fontFamily': font_family,
            'bold': bold
        }
        if color:
            text_format['foregroundColor'] = color
            
        return {
            'repeatCell': {
                'range': {'sheetId': worksheet.id, 'startRowIndex': start_row, 'endRowIndex': end_row, 
                         'startColumnIndex': start_col, 'endColumnIndex': end_col},
                'cell': {
                    'userEnteredFormat': {
                        'textFormat': text_format,
                        'horizontalAlignment': alignment
                    }
                },
                'fields': 'userEnteredFormat'
            }
        }
    
    def _create_background_request(self, worksheet, start_row: int, end_row: int, start_col: int, end_col: int, 
                                 color: Dict[str, float]) -> Dict:
        """Helper method to create background color requests"""
        return {
            'repeatCell': {
                'range': {'sheetId': worksheet.id, 'startRowIndex': start_row, 'endRowIndex': end_row, 
                         'startColumnIndex': start_col, 'endColumnIndex': end_col},
                'cell': {
                    'userEnteredFormat': {
                        'backgroundColor': color
                    }
                },
                'fields': 'userEnteredFormat.backgroundColor'
            }
        }
    
    def _apply_merging_and_widths_in_one_batch(self, worksheet) -> None:
        """Apply merging and column widths in one batch request"""
        requests = []
        
        # Merging requests
        merge_requests = [
            # Header merge (rows 1-2, columns A-F)
            {
                'mergeCells': {
                    'range': {'sheetId': worksheet.id, 'startRowIndex': 0, 'endRowIndex': 2, 'startColumnIndex': 0, 'endColumnIndex': 6},
                    'mergeType': 'MERGE_ALL'
                }
            },
            # Company name merge (row 3, columns B-F)
            {
                'mergeCells': {
                    'range': {'sheetId': worksheet.id, 'startRowIndex': 2, 'endRowIndex': 3, 'startColumnIndex': 1, 'endColumnIndex': 6},
                    'mergeType': 'MERGE_ALL'
                }
            },
            # Address merge (row 4, columns B-F)
            {
                'mergeCells': {
                    'range': {'sheetId': worksheet.id, 'startRowIndex': 3, 'endRowIndex': 4, 'startColumnIndex': 1, 'endColumnIndex': 6},
                    'mergeType': 'MERGE_ALL'
                }
            },
            # City/Province merge (row 5, columns B-F)
            {
                'mergeCells': {
                    'range': {'sheetId': worksheet.id, 'startRowIndex': 4, 'endRowIndex': 5, 'startColumnIndex': 1, 'endColumnIndex': 6},
                    'mergeType': 'MERGE_ALL'
                }
            },
            # Phone merge (row 6, columns B-F)
            {
                'mergeCells': {
                    'range': {'sheetId': worksheet.id, 'startRowIndex': 5, 'endRowIndex': 6, 'startColumnIndex': 1, 'endColumnIndex': 6},
                    'mergeType': 'MERGE_ALL'
                }
            },
            # Invoice title merge (row 8, columns B-F)
            {
                'mergeCells': {
                    'range': {'sheetId': worksheet.id, 'startRowIndex': 7, 'endRowIndex': 8, 'startColumnIndex': 1, 'endColumnIndex': 6},
                    'mergeType': 'MERGE_ALL'
                }
            },
            # Invoice date merge (row 9, columns B-F)
            {
                'mergeCells': {
                    'range': {'sheetId': worksheet.id, 'startRowIndex': 8, 'endRowIndex': 9, 'startColumnIndex': 1, 'endColumnIndex': 6},
                    'mergeType': 'MERGE_ALL'
                }
            },
            # Notes merge (rows 24-30, columns A-D)
            {
                'mergeCells': {
                    'range': {'sheetId': worksheet.id, 'startRowIndex': 23, 'endRowIndex': 30, 'startColumnIndex': 0, 'endColumnIndex': 4},
                    'mergeType': 'MERGE_ALL'
                }
            },
            # Total label merge (row 25, columns E-F)
            {
                'mergeCells': {
                    'range': {'sheetId': worksheet.id, 'startRowIndex': 24, 'endRowIndex': 25, 'startColumnIndex': 4, 'endColumnIndex': 6},
                    'mergeType': 'MERGE_ALL'
                }
            },
            # Total value merge (row 27, columns E-F)
            {
                'mergeCells': {
                    'range': {'sheetId': worksheet.id, 'startRowIndex': 26, 'endRowIndex': 27, 'startColumnIndex': 4, 'endColumnIndex': 6},
                    'mergeType': 'MERGE_ALL'
                }
            }
        ]
        
        requests.extend(merge_requests)
        
        # Column width requests
        width_requests = [
            {'updateDimensionProperties': {'range': {'sheetId': worksheet.id, 'dimension': 'COLUMNS', 'startIndex': 0, 'endIndex': 1}, 'properties': {'pixelSize': 30}, 'fields': 'pixelSize'}},
            {'updateDimensionProperties': {'range': {'sheetId': worksheet.id, 'dimension': 'COLUMNS', 'startIndex': 1, 'endIndex': 2}, 'properties': {'pixelSize': 200}, 'fields': 'pixelSize'}},
            {'updateDimensionProperties': {'range': {'sheetId': worksheet.id, 'dimension': 'COLUMNS', 'startIndex': 2, 'endIndex': 3}, 'properties': {'pixelSize': 50}, 'fields': 'pixelSize'}},
            {'updateDimensionProperties': {'range': {'sheetId': worksheet.id, 'dimension': 'COLUMNS', 'startIndex': 3, 'endIndex': 4}, 'properties': {'pixelSize': 60}, 'fields': 'pixelSize'}},
            {'updateDimensionProperties': {'range': {'sheetId': worksheet.id, 'dimension': 'COLUMNS', 'startIndex': 4, 'endIndex': 5}, 'properties': {'pixelSize': 80}, 'fields': 'pixelSize'}},
            {'updateDimensionProperties': {'range': {'sheetId': worksheet.id, 'dimension': 'COLUMNS', 'startIndex': 5, 'endIndex': 6}, 'properties': {'pixelSize': 100}, 'fields': 'pixelSize'}},
        ]
        
        requests.extend(width_requests)
        
        # Execute all requests in one batch
        worksheet.spreadsheet.batch_update({'requests': requests})
    
    
    def _build_notes_content(self, data: InvoiceData) -> str:
        """Build notes content"""
        notes_config = self.config_loader.get_notes_config()
        content = "Notes:"
        
        # Check for holidays in the last two weeks
        holidays_in_period = self._get_canadian_holidays_in_period()
        
        # Add holiday information if applicable
        if notes_config.get('include_holidays', True) and holidays_in_period:
            content += "\n\nCanadian Statutory Holidays (Paid Days Off):"
            for holiday in holidays_in_period:
                content += f"\n• {holiday['date']}: {holiday['name']}"
        
        # Add custom notes - automatically include holiday names if holidays detected
        custom_notes = notes_config.get('custom_notes', '')
        if holidays_in_period and not custom_notes:
            # Auto-generate custom notes with holiday names
            holiday_names = [holiday['name'] for holiday in holidays_in_period]
            if len(holiday_names) == 1:
                custom_notes = f"Note: Holiday hours for {holiday_names[0]} are included in regular billing."
            else:
                holiday_list = ", ".join(holiday_names[:-1]) + f" and {holiday_names[-1]}"
                custom_notes = f"Note: Holiday hours for {holiday_list} are included in regular billing."
        
        if custom_notes:
            content += f"\n\n{custom_notes}"
        
        return content
    
    def _get_canadian_holidays_in_period(self) -> List[Dict[str, str]]:
        """Get Canadian holidays in the last two weeks"""
        end_date = datetime.datetime.now()
        start_date = datetime.datetime.now() - datetime.timedelta(days=14)
        
        ca_holidays = holidays.Canada()
        holidays_in_period = []
        current_date = start_date
        
        while current_date <= end_date:
            if current_date.date() in ca_holidays:
                holiday_name = ca_holidays.get(current_date.date())
                holidays_in_period.append({
                    'date': current_date.strftime('%B %d'),
                    'name': holiday_name
                })
            current_date += datetime.timedelta(days=1)
        
        return holidays_in_period

# Usage example
def main():
    """Main function using the optimized approach with JSON configuration"""
    # Load configuration
    with open('invoicepacksize-google-creds.json', 'r') as f:
        creds_data = json.load(f)
    
    # Create invoice using JSON configuration
    template = InvoiceTemplate(creds_data['spreadsheet_id'], 'invoicepacksize-google-creds.json')
    template.connect()
    url = template.create_invoice()  # No need to pass data - it loads from JSON
    
    print(f"Invoice created successfully!")
    print(f"URL: {url}")

if __name__ == "__main__":
    main()
