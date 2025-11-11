"""
Quotation UI Module - Updated with Excel Export
Contains the main PyQt5 GUI for the quotation system with Excel export functionality.
"""

import sys
import os
import re
from datetime import datetime
from PyQt5.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QLabel, QComboBox, QSpinBox, QPushButton,
                             QTableWidget, QTableWidgetItem, QGroupBox, QCheckBox,
                             QLineEdit, QMessageBox, QFileDialog, QHeaderView,
                             QGridLayout, QTextEdit, QDateEdit, QTabWidget, QListWidget, QSpacerItem, QSizePolicy,
                             QDialog, QProgressBar, QApplication)
from PyQt5.QtCore import Qt, QDate
from PyQt5.QtGui import QFont, QColor, QIcon

from utils.price_calculator import PriceCalculator
from utils.excel_exporter import ExcelQuotationExporter
from utils.excel_importer import ExcelItemImporter
from utils.filter_utils import get_filter_price
from utils.product_utils import extract_product_flags_and_filter, convert_dimension_to_inches, find_matching_product, extract_slot_number_from_model, get_product_type_flags
from utils.quote_utils import build_quote_item


class ExcelUploadProgressDialog(QDialog):
    """Dialog showing progress for Excel file upload with scrollable error display"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle('Uploading Excel File')
        self.setModal(True)
        self.setMinimumWidth(600)
        self.setMinimumHeight(90)  # Compact initial height
        self.resize(600, 90)  # Set initial compact size
        
        layout = QVBoxLayout()
        layout.setSpacing(5)  # Reduced spacing for tighter layout
        
        # Status label
        self.status_label = QLabel('Initializing...')
        self.status_label.setWordWrap(True)
        layout.addWidget(self.status_label)
        
        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)  # Show percentage
        layout.addWidget(self.progress_bar)
        
        # Scrollable error/warning area (initially hidden)
        self.error_text = QTextEdit()
        self.error_text.setReadOnly(True)
        self.error_text.setPlaceholderText('No errors or warnings')
        self.error_text.setMinimumHeight(150)
        self.error_text.setVisible(False)
        layout.addWidget(self.error_text)  # No stretch initially - will expand when shown
        
        # Close button (initially hidden)
        self.close_button = QPushButton('Close')
        self.close_button.setStyleSheet('''
            QPushButton {
                background-color: #4CAF50;
                color: white;
                font-weight: bold;
                padding: 8px 16px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        ''')
        self.close_button.clicked.connect(self.accept)
        self.close_button.setVisible(False)
        layout.addWidget(self.close_button)
        
        self.setLayout(layout)
    
    def update_progress(self, value, status_text=''):
        """Update progress bar and status text"""
        self.progress_bar.setValue(value)
        if status_text:
            self.status_label.setText(status_text)
        # Force UI updates to ensure progress is visible
        self.update()  # Update the dialog
        QApplication.processEvents()  # Process pending events
    
    def show_results(self, added_count, title_count, invalid_items, warnings):
        """Show final results with errors/warnings in scrollable area"""
        # Hide progress bar
        self.progress_bar.setVisible(False)
        
        # Update status
        status_text = f'Upload Complete!\n\n'
        status_text += f'✓ Successfully imported {added_count} item(s)'
        if title_count > 0:
            status_text += f' and {title_count} title(s)'
        
        if invalid_items or warnings:
            status_text += f'\n\n⚠ {len(invalid_items)} item(s) with errors'
            if warnings:
                status_text += f' and {len(warnings)} warning(s)'
        
        self.status_label.setText(status_text)
        
        # Show errors/warnings if any
        if invalid_items or warnings:
            error_text = ''
            
            if invalid_items:
                error_text += '=== ERRORS ===\n\n'
                for idx, invalid in enumerate(invalid_items, 1):
                    error_text += f'{idx}. Model: {invalid["model"]}\n'
                    error_text += f'   Error: {invalid["error"]}\n\n'
            
            if warnings:
                error_text += '=== WARNINGS ===\n\n'
                for idx, warning in enumerate(warnings, 1):
                    error_text += f'{idx}. {warning}\n\n'
            
            self.error_text.setText(error_text)
            self.error_text.setVisible(True)
            # Expand dialog to accommodate error area
            self.setMinimumHeight(400)
            self.resize(600, 400)
        else:
            # Keep compact if no errors
            self.setMinimumHeight(130)
            self.resize(600, 130)
        
        # Show close button
        self.close_button.setVisible(True)
        QApplication.processEvents()  # Ensure UI updates


class QuotationApp(QMainWindow):
    """Main application window for the quotation system"""
    
    def __init__(self):
        super().__init__()
        self.quote_items = []
        self.price_loader = None
        self.excel_exporter = ExcelQuotationExporter()
        self.font_size_multiplier = 1.0  # Default font size multiplier
        self.original_fonts = {}  # Store original font sizes for widgets
        self.table_item_base_font_size = 13  # Base font size for table items
        self.init_ui()
        self.load_price_list()
    
    def init_ui(self):
        """Initialize the user interface"""
        self.setWindowTitle('Quotation System')
        self.setGeometry(100, 100, 1500, 900)
        
        # Set window icon
        self.set_window_icon()
        
        # Main widget and layout
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout()
        main_widget.setLayout(main_layout)
        
        # Title bar with text size control
        title_bar = QHBoxLayout()
        title = QLabel('ระบบจัดการใบเสนอราคา / QUOTATION MANAGEMENT SYSTEM')
        title.setFont(QFont('Arial', 16, QFont.Bold))
        title.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        title_bar.addWidget(title)
        
        # Add stretch to push adjuster to the right
        title_bar.addStretch()
        
        # Text size adjustment controls (smaller)
        text_size_container = QWidget()
        text_size_layout = QHBoxLayout()
        text_size_layout.setContentsMargins(0, 0, 0, 0)
        text_size_layout.setSpacing(0)
        
        # Decrease button (-)
        self.text_size_decrease_button = QPushButton('−')
        self.text_size_decrease_button.setToolTip('Decrease text size')
        self.text_size_decrease_button.setStyleSheet('''
            QPushButton {
                background-color: #7B68EE;
                color: white;
                font-weight: bold;
                font-size: 14px;
                padding: 4px 8px;
                border-top-left-radius: 4px;
                border-bottom-left-radius: 4px;
                border-top-right-radius: 0px;
                border-bottom-right-radius: 0px;
                border: 1px solid #6A5ACD;
                min-width: 28px;
                max-width: 28px;
            }
            QPushButton:hover {
                background-color: #6A5ACD;
            }
            QPushButton:pressed {
                background-color: #5A4FCF;
            }
            QPushButton:disabled {
                background-color: #CCCCCC;
                color: #888888;
            }
        ''')
        self.text_size_decrease_button.clicked.connect(self.decrease_text_size)
        text_size_layout.addWidget(self.text_size_decrease_button)
        
        # Size label showing current percentage
        self.text_size_label = QLabel('100%')
        self.text_size_label.setToolTip('Current text size')
        self.text_size_label.setAlignment(Qt.AlignCenter)
        self.text_size_label.setStyleSheet('''
            QLabel {
                background-color: #7B68EE;
                color: white;
                font-weight: bold;
                font-size: 10px;
                padding: 4px 6px;
                border: none;
                min-width: 35px;
                max-width: 35px;
            }
        ''')
        text_size_layout.addWidget(self.text_size_label)
        
        # Increase button (+)
        self.text_size_increase_button = QPushButton('+')
        self.text_size_increase_button.setToolTip('Increase text size')
        self.text_size_increase_button.setStyleSheet('''
            QPushButton {
                background-color: #7B68EE;
                color: white;
                font-weight: bold;
                font-size: 14px;
                padding: 4px 8px;
                border-top-left-radius: 0px;
                border-bottom-left-radius: 0px;
                border-top-right-radius: 4px;
                border-bottom-right-radius: 4px;
                border: 1px solid #6A5ACD;
                min-width: 28px;
                max-width: 28px;
            }
            QPushButton:hover {
                background-color: #6A5ACD;
            }
            QPushButton:pressed {
                background-color: #5A4FCF;
            }
            QPushButton:disabled {
                background-color: #CCCCCC;
                color: #888888;
            }
        ''')
        self.text_size_increase_button.clicked.connect(self.increase_text_size)
        text_size_layout.addWidget(self.text_size_increase_button)
        
        text_size_container.setLayout(text_size_layout)
        title_bar.addWidget(text_size_container)
        
        title_widget = QWidget()
        title_widget.setLayout(title_bar)
        main_layout.addWidget(title_widget)
        
        # Create tab widget for better organization
        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs)
        
        # Tab 1: Main quotation
        main_tab = QWidget()
        main_tab_layout = QVBoxLayout()
        main_tab.setLayout(main_tab_layout)
        
        main_tab_layout.addWidget(self.create_quote_info_section())
        main_tab_layout.addWidget(self.create_title_section())
        main_tab_layout.addWidget(self.create_product_selection_section())
        main_tab_layout.addWidget(self.create_items_table_section())
        
        self.tabs.addTab(main_tab, "ใบเสนอราคา / Quotation")
        
        # Tab 2: Additional Information
        additional_tab = QWidget()
        additional_tab_layout = QVBoxLayout()
        additional_tab.setLayout(additional_tab_layout)
        
        additional_tab_layout.addWidget(self.create_additional_info_section())
        additional_tab_layout.addStretch()
        
        self.tabs.addTab(additional_tab, "ข้อมูลเพิ่มเติม / Additional Info")
        
        # Action buttons at the bottom
        main_layout.addWidget(self.create_action_buttons())
        
        # Set up product dropdown for floating positioning
        self.product_dropdown.setParent(main_widget)
        self.product_dropdown.setWindowFlags(Qt.Widget)
        self.product_dropdown.setAttribute(Qt.WA_ShowWithoutActivating)
        
        self.statusBar().showMessage('Ready')
        
        # Store original fonts and apply initial font size
        self.store_original_fonts()
        self.update_text_size_controls()
    
    def set_window_icon(self):
        """Set the window icon for the application"""
        try:
            if getattr(sys, 'frozen', False):
                icon_path = os.path.join(sys._MEIPASS, 'assets', 'icon.ico')
            else:
                icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..', 'assets', 'icon.ico')
            
            if os.path.exists(icon_path):
                self.setWindowIcon(QIcon(icon_path))
        except Exception:
            pass  # Continue without icon if loading fails
    
    def create_quote_info_section(self):
        """Create the quote information input section"""
        group = QGroupBox('ข้อมูลใบเสนอราคา / Quote Information')
        layout = QGridLayout()
        
        # Row 1: TO and Quote Number
        layout.addWidget(QLabel('ถึง / TO:'), 0, 0)
        self.to_input = QLineEdit()
        self.to_input.setPlaceholderText('ชื่อผู้รับ / Recipient Name')
        layout.addWidget(self.to_input, 0, 1, 1, 3)  # Span to column 3 to match Fax input end
        
        # Add horizontal spacer to push the next group to the right (only in this row)
        spacer1 = QSpacerItem(20, 0, QSizePolicy.Expanding, QSizePolicy.Minimum)
        layout.addItem(spacer1, 0, 4)
        
        layout.addWidget(QLabel('เลขที่ / NO.:'), 0, 5)
        self.quote_number = QLineEdit()
        self.quote_number.setText(f"{datetime.now().strftime('%y-%m')}{datetime.now().day:03d}")
        layout.addWidget(self.quote_number, 0, 6)
        
        # Row 2: Company and Date
        layout.addWidget(QLabel('บริษัท / COMPANY:'), 1, 0)
        self.company_input = QLineEdit()
        self.company_input.setPlaceholderText('ชื่อบริษัท / Company Name')
        layout.addWidget(self.company_input, 1, 1, 1, 3)  # Span to column 3 to match Fax input end
        
        # Add horizontal spacer to push the next group to the right (only in this row)
        spacer2 = QSpacerItem(20, 0, QSizePolicy.Expanding, QSizePolicy.Minimum)
        layout.addItem(spacer2, 1, 4)
        
        layout.addWidget(QLabel('วันที่ / DATE:'), 1, 5)
        self.quote_date = QDateEdit()
        self.quote_date.setDate(QDate.currentDate())
        self.quote_date.setCalendarPopup(True)
        self.quote_date.setDisplayFormat('yyyy-MM-dd')
        layout.addWidget(self.quote_date, 1, 6)
        
        # Row 3: Tel, Fax, and Project
        layout.addWidget(QLabel('โทร / TEL:'), 2, 0)
        self.tel_input = QLineEdit()
        self.tel_input.setPlaceholderText('เบอร์โทรศัพท์ / Phone Number')
        layout.addWidget(self.tel_input, 2, 1)
        
        layout.addWidget(QLabel('Fax:'), 2, 2)
        self.fax_input = QLineEdit()
        self.fax_input.setPlaceholderText('เบอร์แฟกซ์ / Fax Number')
        self.fax_input.setMaximumWidth(150)  # Limit the width to fit in the same row
        layout.addWidget(self.fax_input, 2, 3)
        
        # Add horizontal spacer to push the next group to the right (only in this row)
        spacer3 = QSpacerItem(20, 0, QSizePolicy.Expanding, QSizePolicy.Minimum)
        layout.addItem(spacer3, 2, 4)
        
        layout.addWidget(QLabel('งาน / PROJECT:'), 2, 5)
        self.project_input = QLineEdit()
        self.project_input.setPlaceholderText('ชื่อโครงการ / Project Name')
        layout.addWidget(self.project_input, 2, 6)
        
        # Set column stretch only for input columns within rows
        layout.setColumnStretch(1, 2)  # TO/Company/Tel input area (spans columns 1-3)
        layout.setColumnStretch(6, 2)  # Quote number/Date/Project input area
        # Don't stretch column 3 (Fax input end) or column 5 (label column) globally
        # Set minimum width for label column to prevent excessive spacing
        layout.setColumnMinimumWidth(5, 0)  # Label column (all three labels in same column)
        
        group.setLayout(layout)
        return group
    
    def create_additional_info_section(self):
        """Create additional information section for footer details"""
        group = QGroupBox('ข้อมูลเพิ่มเติม / Additional Information')
        layout = QGridLayout()
        
        # Row 1: Remarks
        layout.addWidget(QLabel('หมายเหตุ / REMARKS:'), 0, 0)
        self.remarks_input = QTextEdit()
        self.remarks_input.setPlainText('1. ยืนราคา 60 วัน')
        self.remarks_input.setMaximumHeight(80)
        layout.addWidget(self.remarks_input, 0, 1, 1, 3)
        
        # Row 2: Payment Term
        layout.addWidget(QLabel('เงื่อนไขการชำระเงิน / PAYMENT TERM:'), 1, 0)
        self.payment_term_combo = QComboBox()
        self.payment_term_combo.addItems([
            'เครดิต 30 วัน / Credit 30 days',
            'เครดิต 60 วัน / Credit 60 days',
            'เงินสด / Cash',
            'เช็ค / Cheque',
            '50% มัดจำ, 50% ก่อนส่งของ / 50% deposit, 50% before delivery'
        ])
        self.payment_term_combo.setEditable(True)
        layout.addWidget(self.payment_term_combo, 1, 1, 1, 3)
        
        # Row 3: Delivery Place
        layout.addWidget(QLabel('สถานที่ส่งของ / DELIVERY PLACE:'), 2, 0)
        self.delivery_place_input = QLineEdit()
        self.delivery_place_input.setPlaceholderText('สถานที่ส่งสินค้า / Delivery Location')
        layout.addWidget(self.delivery_place_input, 2, 1, 1, 3)
        
        # Row 4: Delivery Date
        layout.addWidget(QLabel('วันส่งของ / DELIVERY DATE:'), 3, 0)
        self.delivery_date_input = QLineEdit()
        self.delivery_date_input.setPlaceholderText('e.g., หลังได้รับใบสั่งซื้อ 15 วัน / 15 days after PO')
        layout.addWidget(self.delivery_date_input, 3, 1, 1, 3)
        
        # Row 5: Quoted By and Purchased By
        layout.addWidget(QLabel('เสนอราคาโดย / QUOTED BY:'), 4, 0)
        self.quoted_by_input = QLineEdit()
        self.quoted_by_input.setPlaceholderText('ชื่อผู้เสนอราคา / Salesperson Name')
        layout.addWidget(self.quoted_by_input, 4, 1)
        
        layout.addWidget(QLabel('ผู้สั่งซื้อ / PURCHASED BY:'), 4, 2)
        self.purchased_by_input = QLineEdit()
        self.purchased_by_input.setPlaceholderText('(ลายเซ็นผู้สั่งซื้อ / Buyer Signature)')
        layout.addWidget(self.purchased_by_input, 4, 3)
        
        # Set column stretch
        layout.setColumnStretch(1, 2)
        layout.setColumnStretch(3, 2)
        
        group.setLayout(layout)
        return group
    
    def create_product_selection_section(self):
        """Create the product selection panel"""
        group = QGroupBox('Product Selection')
        main_layout = QVBoxLayout()  # Main vertical layout
        
        # First row: Product Type, Finish, Unit, Width, Height, Quantity, Discount, etc.
        first_row = QHBoxLayout()
        
        # Product Type
        prod_layout = QVBoxLayout()
        prod_layout.addWidget(QLabel('Product Type:'))
        self.product_input = QLineEdit()
        self.product_input.setPlaceholderText('Type to search product models...')
        self.product_input.textChanged.connect(self.on_product_text_changed)
        self.product_input.returnPressed.connect(self.on_product_selected)
        self.product_input.focusOutEvent = self.on_product_input_focus_out
        self.product_input.keyPressEvent = self.on_product_input_key_press
        prod_layout.addWidget(self.product_input)
        first_row.addLayout(prod_layout)
        
        # Dropdown list for matching products - positioned absolutely to float over content
        self.product_dropdown = QListWidget()
        self.product_dropdown.setMaximumHeight(150)
        self.product_dropdown.setVisible(False)
        self.product_dropdown.itemClicked.connect(self.on_dropdown_item_selected)
        self.product_dropdown.setStyleSheet("""
            QListWidget {
                border: 1px solid #ccc;
                background-color: white;
                selection-background-color: #0078d4;
            }
            QListWidget::item {
                padding: 5px;
                border-bottom: 1px solid #eee;
            }
            QListWidget::item:hover {
                background-color: #f0f0f0;
            }
            QListWidget::item:selected {
                background-color: #0078d4;
                color: white;
            }
        """)
        
        # Finish
        finish_layout = QVBoxLayout()
        finish_layout.addWidget(QLabel('Finish:'))
        self.finish_combo = QComboBox()
        # Finish options will be populated dynamically based on selected product
        self.finish_combo.currentTextChanged.connect(self.on_selection_changed)
        finish_layout.addWidget(self.finish_combo)
        first_row.addLayout(finish_layout)
        
        # Powder Coating Color (initially hidden)
        self.powder_color_layout = QVBoxLayout()
        self.powder_color_label = QLabel('Powder Coating Color:')
        self.powder_color_layout.addWidget(self.powder_color_label)
        self.powder_color_combo = QComboBox()
        self.powder_color_combo.addItems([
            'ขาวนวล',
            'ขาวเงา',
            'ขาวด้าน',
            'ขาวบริสุทธิ์',
            'ดำเงา',
            'ดำด้าน',
            'บรอนส์'
        ])
        self.powder_color_combo.currentTextChanged.connect(self.on_selection_changed)
        self.powder_color_layout.addWidget(self.powder_color_combo)
        first_row.addLayout(self.powder_color_layout)
        
        # Initially hide powder color selection
        self.show_hide_widgets(self.powder_color_layout, False)
        
        # Special Color Name (initially hidden)
        self.special_color_layout = QVBoxLayout()
        self.special_color_label = QLabel('Special Color Name:')
        self.special_color_layout.addWidget(self.special_color_label)
        self.special_color_input = QLineEdit()
        self.special_color_input.setPlaceholderText('e.g., Custom Blue, RAL 5005, etc.')
        self.special_color_input.textChanged.connect(self.on_selection_changed)
        self.special_color_layout.addWidget(self.special_color_input)
        first_row.addLayout(self.special_color_layout)
        
        # Special Color Multiplier (initially hidden)
        self.special_color_multiplier_layout = QVBoxLayout()
        self.special_color_multiplier_label = QLabel('Special Color Multiplier:')
        self.special_color_multiplier_layout.addWidget(self.special_color_multiplier_label)
        self.special_color_multiplier_spin = QSpinBox()
        self.special_color_multiplier_spin.setMinimum(1)
        self.special_color_multiplier_spin.setMaximum(9999)
        self.special_color_multiplier_spin.setValue(100)  # Default to 1.00 (100/100)
        self.special_color_multiplier_spin.setSuffix('%')
        self.special_color_multiplier_spin.valueChanged.connect(self.on_selection_changed)
        self.special_color_multiplier_layout.addWidget(self.special_color_multiplier_spin)
        first_row.addLayout(self.special_color_multiplier_layout)
        
        # Initially hide special color inputs
        self.show_hide_widgets(self.special_color_layout, False)
        self.show_hide_widgets(self.special_color_multiplier_layout, False)
        
        # Unit Selection
        unit_layout = QVBoxLayout()
        unit_layout.addWidget(QLabel('Unit:'))
        self.unit_combo = QComboBox()
        self.unit_combo.addItems(['Inches', 'Millimeters'])
        self.unit_combo.currentTextChanged.connect(self.on_unit_changed)
        unit_layout.addWidget(self.unit_combo)
        first_row.addLayout(unit_layout)
        
        # Width
        self.width_layout = QVBoxLayout()
        self.width_label = QLabel('Width (inches):')
        self.width_layout.addWidget(self.width_label)
        self.width_spin = QSpinBox()
        self.width_spin.setMinimum(1)
        self.width_spin.setMaximum(9999)
        self.width_spin.setValue(4)
        self.width_spin.valueChanged.connect(self.update_price_display)
        self.width_layout.addWidget(self.width_spin)
        first_row.addLayout(self.width_layout)
        
        # Height
        self.height_layout = QVBoxLayout()
        self.height_label = QLabel('Height (inches):')
        self.height_layout.addWidget(self.height_label)
        self.height_spin = QSpinBox()
        self.height_spin.setMinimum(1)
        self.height_spin.setMaximum(9999)
        self.height_spin.setValue(4)
        self.height_spin.valueChanged.connect(self.update_price_display)
        self.height_layout.addWidget(self.height_spin)
        first_row.addLayout(self.height_layout)
        
        # Other Table Size (initially hidden)
        self.other_table_layout = QVBoxLayout()
        self.other_table_label = QLabel('Size (inches):')
        self.other_table_layout.addWidget(self.other_table_label)
        self.other_table_spin = QSpinBox()
        self.other_table_spin.setMinimum(1)
        self.other_table_spin.setMaximum(9999)
        self.other_table_spin.setValue(4)
        self.other_table_spin.valueChanged.connect(self.update_price_display)
        self.other_table_layout.addWidget(self.other_table_spin)
        first_row.addLayout(self.other_table_layout)
        
        # Initially hide other table layout
        self.show_hide_widgets(self.other_table_layout, False)
        
        # Quantity
        qty_layout = QVBoxLayout()
        qty_layout.addWidget(QLabel('Quantity:'))
        self.quantity_spin = QSpinBox()
        self.quantity_spin.setMinimum(1)
        self.quantity_spin.setMaximum(9999)
        self.quantity_spin.setValue(1)
        self.quantity_spin.valueChanged.connect(self.update_price_display)
        qty_layout.addWidget(self.quantity_spin)
        first_row.addLayout(qty_layout)
        
        # Discount
        discount_layout = QVBoxLayout()
        discount_layout.addWidget(QLabel('Discount (%):'))
        self.discount_spin = QSpinBox()
        self.discount_spin.setMinimum(0)
        self.discount_spin.setMaximum(100)
        self.discount_spin.setValue(0)
        self.discount_spin.setSuffix('%')
        self.discount_spin.valueChanged.connect(self.update_price_display)
        discount_layout.addWidget(self.discount_spin)
        first_row.addLayout(discount_layout)
        
        # Unit Price Display
        price_layout = QVBoxLayout()
        price_layout.addWidget(QLabel('Unit Price:'))
        self.unit_price_label = QLabel('฿ 0.00')
        self.unit_price_label.setFont(QFont('Arial', 12, QFont.Bold))
        self.unit_price_label.setStyleSheet('color: #2E7D32; padding: 5px;')
        self.unit_price_label.setMinimumWidth(120)  # Fixed width to prevent shifting
        price_layout.addWidget(self.unit_price_label)
        first_row.addLayout(price_layout)
        
        # Rounded Size Display
        rounded_size_layout = QVBoxLayout()
        rounded_size_layout.addWidget(QLabel('Rounded Size:'))
        self.rounded_size_label = QLabel('')
        self.rounded_size_label.setFont(QFont('Arial', 12, QFont.Bold))
        self.rounded_size_label.setStyleSheet('color: #FF5722; padding: 5px;')
        self.rounded_size_label.setMinimumWidth(120)  # Fixed width to prevent shifting
        rounded_size_layout.addWidget(self.rounded_size_label)
        first_row.addLayout(rounded_size_layout)
        
        # Total Price Display
        total_layout = QVBoxLayout()
        total_layout.addWidget(QLabel('Total:'))
        self.total_price_label = QLabel('฿ 0.00')
        self.total_price_label.setFont(QFont('Arial', 12, QFont.Bold))
        self.total_price_label.setStyleSheet('color: #1565C0; padding: 5px;')
        self.total_price_label.setMinimumWidth(120)  # Fixed width to prevent shifting
        total_layout.addWidget(self.total_price_label)
        first_row.addLayout(total_layout)
        
        # Add Button
        add_layout = QVBoxLayout()
        add_layout.addWidget(QLabel(''))  # Spacer
        self.add_button = QPushButton('Add to Quote')
        self.add_button.setStyleSheet('background-color: #4CAF50; color: white; padding: 10px; font-weight: bold;')
        self.add_button.clicked.connect(self.add_item_to_quote)
        add_layout.addWidget(self.add_button)
        first_row.addLayout(add_layout)
        
        # Add first row to main layout
        main_layout.addLayout(first_row)
        
        # Second row: Detail (on its own row)
        second_row = QHBoxLayout()
        detail_layout = QVBoxLayout()
        detail_layout.addWidget(QLabel('Detail:'))
        self.detail_input = QLineEdit()
        self.detail_input.setPlaceholderText('Enter detail for Excel export...')
        detail_layout.addWidget(self.detail_input)
        second_row.addLayout(detail_layout)
        
        second_row.addStretch()  # Push detail to the left
        
        # Add second row to main layout
        main_layout.addLayout(second_row)
        
        group.setLayout(main_layout)
        
        # Store reference to product input for dropdown positioning
        self.product_input_widget = self.product_input
        
        return group
    
    def create_title_section(self):
        """Create the title input section"""
        group = QGroupBox('Title / หัวข้อ')
        layout = QHBoxLayout()
        
        # Title input
        title_layout = QVBoxLayout()
        title_layout.addWidget(QLabel('Title:'))
        self.title_input = QLineEdit()
        self.title_input.setPlaceholderText('Enter title to add between products')
        title_layout.addWidget(self.title_input)
        layout.addLayout(title_layout)
        
        # Add Title Button
        add_title_layout = QVBoxLayout()
        add_title_layout.addWidget(QLabel(''))  # Spacer to align with "Title:" label
        self.add_title_button = QPushButton('Add Title')
        self.add_title_button.setStyleSheet('''
            QPushButton {
                background-color: #FF9800;
                color: white;
                padding: 5px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #F57C00;
            }
        ''')
        self.add_title_button.clicked.connect(self.add_title_to_quote)
        add_title_layout.addWidget(self.add_title_button)
        layout.addLayout(add_title_layout)
        
        layout.addStretch()
        
        group.setLayout(layout)
        return group
    
    def create_items_table_section(self):
        """Create the quote items table"""
        group = QGroupBox('Quote Items')
        layout = QVBoxLayout()
        
        # Table
        self.items_table = QTableWidget()
        self.items_table.setColumnCount(9)
        self.items_table.setHorizontalHeaderLabels([
            'Item', 'Product', 'Detail', 'Finish', 'Size', 'Qty', 'Unit Price', 'Discount', 'Total'
        ])
        
        # Set column widths
        header = self.items_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        header.setSectionResizeMode(2, QHeaderView.Stretch)
        header.setSectionResizeMode(3, QHeaderView.Stretch)
        header.setSectionResizeMode(4, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(5, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(6, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(7, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(8, QHeaderView.ResizeToContents)
        
        # Connect selection change to update move button states
        self.items_table.itemSelectionChanged.connect(self.update_move_button_states)
        
        layout.addWidget(self.items_table)
        
        # Buttons for table operations
        button_layout = QHBoxLayout()
        
        # Upload Excel button (smaller)
        self.upload_excel_button = QPushButton('Upload Excel')
        self.upload_excel_button.setStyleSheet('''
            QPushButton {
                background-color: #9C27B0;
                color: white;
                padding: 5px 10px;
                font-weight: bold;
                font-size: 11px;
            }
            QPushButton:hover {
                background-color: #7B1FA2;
            }
        ''')
        self.upload_excel_button.clicked.connect(self.upload_excel_file)
        button_layout.addWidget(self.upload_excel_button)
        
        self.move_up_button = QPushButton('Move Up')
        self.move_up_button.setStyleSheet('''
            QPushButton {
                background-color: #2196F3;
                color: white;
                font-weight: bold;
                padding-left: 15px;
                padding-right: 15px;
            }
            QPushButton:hover {
                background-color: #1976D2;
            }
            QPushButton:disabled {
                background-color: #CCCCCC;
                color: #888888;
            }
        ''')
        self.move_up_button.clicked.connect(self.move_item_up)
        button_layout.addWidget(self.move_up_button)
        
        self.move_down_button = QPushButton('Move Down')
        self.move_down_button.setStyleSheet('''
            QPushButton {
                background-color: #2196F3;
                color: white;
                font-weight: bold;
                padding-left: 15px;
                padding-right: 15px;
            }
            QPushButton:hover {
                background-color: #1976D2;
            }
            QPushButton:disabled {
                background-color: #CCCCCC;
                color: #888888;
            }
        ''')
        self.move_down_button.clicked.connect(self.move_item_down)
        button_layout.addWidget(self.move_down_button)
        
        self.remove_button = QPushButton('Remove Selected Item')
        self.remove_button.clicked.connect(self.remove_selected_item)
        button_layout.addWidget(self.remove_button)
        
        self.clear_button = QPushButton('Clear All Items')
        self.clear_button.clicked.connect(self.clear_all_items)
        button_layout.addWidget(self.clear_button)
        
        # Set the same minimum height for all buttons to ensure they're aligned
        button_height = 30  # Standard button height
        self.move_up_button.setMinimumHeight(button_height)
        self.move_down_button.setMinimumHeight(button_height)
        self.remove_button.setMinimumHeight(button_height)
        self.clear_button.setMinimumHeight(button_height)
        self.upload_excel_button.setMinimumHeight(button_height)
        
        button_layout.addStretch()
        
        # Grand Total
        self.grand_total_label = QLabel('Grand Total: ฿ 0.00')
        self.grand_total_label.setFont(QFont('Arial', 14, QFont.Bold))
        self.grand_total_label.setStyleSheet('color: #C62828;')
        button_layout.addWidget(self.grand_total_label)
        
        layout.addLayout(button_layout)
        
        group.setLayout(layout)
        
        # Initialize button states (disabled initially since no items/selection)
        # This will be called automatically when selection changes or items are added/removed
        # via refresh_items_table() and itemSelectionChanged signal
        
        return group
    
    def create_action_buttons(self):
        """Create action buttons for save, print, excel export, etc."""
        widget = QWidget()
        layout = QHBoxLayout()
        
        self.new_button = QPushButton('New Quote')
        self.new_button.setStyleSheet('''
            QPushButton {
                background-color: #4CAF50;
                color: white;
                font-weight: bold;
                padding: 8px 16px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        ''')
        self.new_button.clicked.connect(self.new_quote)
        layout.addWidget(self.new_button)
        
        # Excel export button
        self.excel_button = QPushButton('Generate Excel Quotation')
        self.excel_button.setStyleSheet('''
            QPushButton {
                background-color: #1565C0;
                color: white;
                font-weight: bold;
                padding: 8px 16px;
            }
            QPushButton:hover {
                background-color: #0D47A1;
            }
        ''')
        self.excel_button.clicked.connect(self.generate_excel_quotation)
        layout.addWidget(self.excel_button)
        
        layout.addStretch()
        
        self.exit_button = QPushButton('Exit')
        self.exit_button.setStyleSheet('''
            QPushButton {
                background-color: #f44336;
                color: white;
                font-weight: bold;
                padding: 8px 16px;
            }
            QPushButton:hover {
                background-color: #da190b;
            }
        ''')
        self.exit_button.clicked.connect(self.close)
        layout.addWidget(self.exit_button)
        
        widget.setLayout(layout)
        return widget
    
    def generate_excel_quotation(self):
        """Generate Excel quotation using the template format"""
        if not self.quote_items:
            QMessageBox.warning(self, 'Warning', 'No items in quote to export')
            return
        
        # Prepare quote data
        quote_data = {
            'to': self.to_input.text(),
            'company': self.company_input.text(),
            'tel': self.tel_input.text(),
            'fax': self.fax_input.text(),
            'quote_no': self.quote_number.text(),
            'date': self.quote_date.date().toString('yyyy-MM-dd'),
            'project': self.project_input.text(),
            'remarks': self.remarks_input.toPlainText(),
            'payment_term': self.payment_term_combo.currentText().split('/')[0].strip(),
            'delivery_place': self.delivery_place_input.text(),
            'delivery_date': self.delivery_date_input.text(),
            'quoted_by_name': self.quoted_by_input.text(),
            'purchased_by': self.purchased_by_input.text(),
        }
        
        # Prepare items with proper Thai finish names
        items_for_excel = []
        for item in self.quote_items:
            excel_item = item.copy()
            # Convert finish to Thai using the exporter's method
            excel_item['finish'] = self.excel_exporter.get_thai_finishing(item.get('finish', ''))
            # Remove warning messages from detail field for Excel export
            detail = excel_item.get('detail', '')
            if detail:
                # Remove warning messages (lines starting with "⚠ Warning:")
                detail_lines = detail.split('\n')
                cleaned_lines = [line for line in detail_lines if not line.strip().startswith('⚠ Warning:')]
                excel_item['detail'] = '\n'.join(cleaned_lines).strip()
            items_for_excel.append(excel_item)
        
        # Get save file path
        file_name, _ = QFileDialog.getSaveFileName(
            self, 'Save Excel Quotation',
            f"Quotation_{self.quote_number.text()}.xlsx",
            'Excel Files (*.xlsx);;All Files (*)'
        )
        
        if file_name:
            try:
                # Generate the Excel file
                success = self.excel_exporter.create_excel_quotation(
                    quote_data, 
                    items_for_excel, 
                    file_name
                )
                
                if success:
                    self.statusBar().showMessage(f'Excel quotation saved to {file_name}')
                    QMessageBox.information(
                        self, 'Success', 
                        f'Excel quotation generated successfully!\n\nSaved to: {file_name}'
                    )
                else:
                    QMessageBox.warning(self, 'Warning', 'Failed to generate Excel quotation')
                    
            except Exception as e:
                QMessageBox.critical(self, 'Error', f'Failed to generate Excel quotation: {str(e)}')
    
    def load_price_list(self):
        """Load the SQLite price database"""
        # Handle both development and bundled executable paths
        if getattr(sys, 'frozen', False):
            # Running as compiled executable - database is in the root of extracted files
            application_path = sys._MEIPASS
            db_file = os.path.join(application_path, 'prices.db')
        else:
            # Running as script
            application_path = os.path.dirname(os.path.abspath(__file__))
            db_file = os.path.join(application_path, '..', '..', 'prices.db')
        
        if not os.path.exists(db_file):
            QMessageBox.critical(self, 'Error', 
                               f'Price database not found: {db_file}\n\n'
                               f'Please run getsql.py first to create the database from the Excel file.')
            return
        
        try:
            self.price_loader = PriceCalculator(db_file)
            
            # Store available models for searching
            base_models = self.price_loader.get_available_models()
            # Add "(WD)" variants for products that have damper option
            self.available_models = []
            for model in base_models:
                # Add the base model
                self.available_models.append(model)
                # Check if this product has damper option for any finish
                # We check by trying to get available finishes - if it has damper, add WD variant
                finishes = self.price_loader.get_available_finishes(model)
                has_damper = False
                for finish in finishes:
                    if self.price_loader.has_damper_option(model, finish):
                        has_damper = True
                        break
                if has_damper:
                    # Add WD variant
                    self.available_models.append(f"{model}(WD)")
            
            if self.available_models:
                self.statusBar().showMessage(f'Price database loaded successfully ({len(self.available_models)} models found)')
            else:
                QMessageBox.warning(self, 'Warning', 'No products found in the database!')
            
            self.update_price_display()
        except Exception as e:
            QMessageBox.critical(self, 'Error', f'Failed to load price database: {str(e)}')
    
    
    def on_product_changed(self):
        """Handle product type change"""
        if not self.price_loader:
            return
        
        # Get available finish options for the selected product
        product_with_wd = self.product_input.text().strip()
        product, _, _, _ = extract_product_flags_and_filter(product_with_wd)
        if product:
            available_finishes = self.price_loader.get_available_finishes(product)
            
            # Update finish combo box with available options
            self.finish_combo.clear()
            if available_finishes:
                self.finish_combo.addItems(available_finishes)
                # Select the first available finish
                self.finish_combo.setCurrentIndex(0)
            else:
                # No finishes available for this product
                self.finish_combo.addItem('No finishes available')
            
            # Get product type flags using consolidated helper
            has_no_dimensions, has_price_per_foot, is_other_table = get_product_type_flags(self.price_loader, product)
            
            if has_no_dimensions:
                # For products with no dimensions, show only height field
                self.show_hide_widgets(self.width_layout, False)
                self.show_hide_widgets(self.height_layout, True)
                self.show_hide_widgets(self.other_table_layout, False)
                # Update label to indicate height is required
                unit = self.unit_combo.currentText()
                unit_text = 'mm' if unit == 'Millimeters' else 'inches'
                self.height_label.setText(f'Height ({unit_text}) *')
            elif has_price_per_foot:
                # For price_per_foot products, show width and height fields (required)
                self.show_hide_widgets(self.width_layout, True)
                self.show_hide_widgets(self.height_layout, True)
                self.show_hide_widgets(self.other_table_layout, False)
                # Update labels to indicate they are required
                self.width_label.setText('Width (inches) *')
                self.height_label.setText('Height (inches) *')
            elif is_other_table:
                # Hide width and height fields, show other table size field
                self.show_hide_widgets(self.width_layout, False)
                self.show_hide_widgets(self.height_layout, False)
                self.show_hide_widgets(self.other_table_layout, True)
            else:
                # Show width and height fields, hide other table size field
                self.show_hide_widgets(self.width_layout, True)
                self.show_hide_widgets(self.height_layout, True)
                self.show_hide_widgets(self.other_table_layout, False)
                # Reset labels if they were modified
                self.width_label.setText('Width (inches):')
                self.height_label.setText('Height (inches):')
            
            # Show/hide finish-specific options
            finish = self.finish_combo.currentText()
            self.show_hide_widgets(self.powder_color_layout, finish == 'Powder Coated')
            self.show_hide_widgets(self.special_color_layout, finish == 'Special Color')
            self.show_hide_widgets(self.special_color_multiplier_layout, finish == 'Special Color')
        
        self.update_price_display()
    
    def show_hide_widgets(self, layout, show):
        """Helper method to show or hide widgets in a layout"""
        for i in range(layout.count()):
            widget = layout.itemAt(i).widget()
            if widget:
                widget.setVisible(show)
    
    def position_dropdown(self):
        """Position the dropdown below the product input field"""
        if not hasattr(self, 'product_input_widget'):
            return
        
        # Get the position of the product input field relative to the main widget
        input_pos = self.product_input_widget.mapTo(self.product_dropdown.parent(), self.product_input_widget.rect().bottomLeft())
        
        # Set the dropdown position and size
        self.product_dropdown.move(input_pos.x(), input_pos.y())
        self.product_dropdown.resize(self.product_input_widget.width(), 150)
        self.product_dropdown.raise_()  # Bring to front
    
    def on_product_text_changed(self):
        """Handle product text input changes for search functionality"""
        if not hasattr(self, 'available_models') or not self.available_models:
            return
        
        search_text = self.product_input.text().strip().lower()
        
        # Clear and hide dropdown if no search text
        if not search_text:
            self.product_dropdown.clear()
            self.product_dropdown.setVisible(False)
            return
        
        # Find matching models
        matching_models = [model for model in self.available_models 
                          if search_text in model.lower()]
        
        # Update dropdown with matching models
        self.product_dropdown.clear()
        if matching_models:
            for model in matching_models[:10]:  # Limit to 10 results for better UX
                self.product_dropdown.addItem(model)
            
            # Position the dropdown below the product input
            self.position_dropdown()
            self.product_dropdown.setVisible(True)
        else:
            self.product_dropdown.setVisible(False)
    
    def on_product_selected(self):
        """Handle when user presses Enter or selects a product"""
        if not self.price_loader:
            return
        
        product_with_wd = self.product_input.text().strip()
        if not product_with_wd:
            return
        
        # Validate that the product exists in available models
        if hasattr(self, 'available_models') and product_with_wd not in self.available_models:
            # Try to find a close match
            match_result = find_matching_product(product_with_wd, self.available_models)
            if match_result:
                # Use the matched product
                matched_product, matched_has_wd = match_result
                product_with_wd = matched_product if not matched_has_wd else f"{matched_product}(WD)"
                self.product_input.setText(product_with_wd)
            else:
                # No match found, show error
                QMessageBox.warning(self, 'Product Not Found', 
                                  f'Product "{product_with_wd}" not found. Please check the spelling or try a different search term.')
                return
        
        # Proceed with the product selection logic
        self.on_product_changed()
    
    def on_dropdown_item_selected(self, item):
        """Handle when user clicks on a dropdown item"""
        selected_product = item.text()
        self.product_input.setText(selected_product)
        self.product_dropdown.setVisible(False)
        self.on_product_selected()
    
    def on_product_input_focus_out(self, event):
        """Handle when focus leaves the product input"""
        # Hide dropdown when focus is lost
        self.product_dropdown.setVisible(False)
        # Call the original focusOutEvent
        QLineEdit.focusOutEvent(self.product_input, event)
    
    def on_product_input_key_press(self, event):
        """Handle keyboard navigation for product input"""
        if self.product_dropdown.isVisible() and self.product_dropdown.count() > 0:
            if event.key() == Qt.Key_Down:
                # Move to first item or next item
                current_row = self.product_dropdown.currentRow()
                if current_row < self.product_dropdown.count() - 1:
                    self.product_dropdown.setCurrentRow(current_row + 1)
                else:
                    self.product_dropdown.setCurrentRow(0)
                return
            elif event.key() == Qt.Key_Up:
                # Move to previous item or last item
                current_row = self.product_dropdown.currentRow()
                if current_row > 0:
                    self.product_dropdown.setCurrentRow(current_row - 1)
                else:
                    self.product_dropdown.setCurrentRow(self.product_dropdown.count() - 1)
                return
            elif event.key() == Qt.Key_Return or event.key() == Qt.Key_Enter:
                # Select the highlighted item
                current_item = self.product_dropdown.currentItem()
                if current_item:
                    self.on_dropdown_item_selected(current_item)
                return
            elif event.key() == Qt.Key_Escape:
                # Hide dropdown
                self.product_dropdown.setVisible(False)
                return
        
        # Call the original keyPressEvent for other keys
        QLineEdit.keyPressEvent(self.product_input, event)
    
    def on_selection_changed(self):
        """Handle selection changes"""
        if not self.price_loader:
            return
        
        product_with_wd = self.product_input.text().strip()
        product, _, _, _ = extract_product_flags_and_filter(product_with_wd)
        finish = self.finish_combo.currentText()
        
        if product and finish:
            # Show/hide finish-specific options
            self.show_hide_widgets(self.powder_color_layout, finish == 'Powder Coated')
            self.show_hide_widgets(self.special_color_layout, finish == 'Special Color')
            self.show_hide_widgets(self.special_color_multiplier_layout, finish == 'Special Color')
        
        self.update_price_display()
    
    def on_unit_changed(self):
        """Handle unit selection change"""
        unit = self.unit_combo.currentText()
        
        if unit == 'Millimeters':
            # Update labels
            self.width_label.setText('Width (mm):')
            self.height_label.setText('Height (mm):')
            self.other_table_label.setText('Size (mm):')
            
            # Update spin box ranges for mm
            self.width_spin.setMinimum(1)
            self.width_spin.setMaximum(9999)
            self.width_spin.setValue(100)  # 100mm ≈ 4 inches
            
            self.height_spin.setMinimum(1)
            self.height_spin.setMaximum(9999)
            self.height_spin.setValue(100)  # 100mm ≈ 4 inches
            
            self.other_table_spin.setMinimum(1)
            self.other_table_spin.setMaximum(9999)
            self.other_table_spin.setValue(100)  # 100mm ≈ 4 inches
        else:
            # Update labels
            self.width_label.setText('Width (inches):')
            self.height_label.setText('Height (inches):')
            self.other_table_label.setText('Size (inches):')
            
            # Update spin box ranges for inches
            self.width_spin.setMinimum(1)
            self.width_spin.setMaximum(9999)
            self.width_spin.setValue(4)
            
            self.height_spin.setMinimum(1)
            self.height_spin.setMaximum(9999)
            self.height_spin.setValue(4)
            
            self.other_table_spin.setMinimum(1)
            self.other_table_spin.setMaximum(9999)
            self.other_table_spin.setValue(4)
        
        self.update_price_display()
    
    def update_price_display(self):
        """Update the price display based on current selections"""
        if not self.price_loader:
            return
        
        product_with_wd = self.product_input.text().strip()
        product, with_damper, _, _ = extract_product_flags_and_filter(product_with_wd)
        
        # If no product is selected, show default values
        if not product:
            self.unit_price_label.setText('฿ 0.00')
            self.total_price_label.setText('฿ 0.00')
            self.rounded_size_label.setText('')
            return
        
        finish = self.finish_combo.currentText()
        
        # Process finish to include color/multiplier information for price calculation
        if finish == 'Powder Coated':
            powder_color = self.powder_color_combo.currentText()
            finish = f"Powder Coated - {powder_color}"
        elif finish == 'Special Color':
            # Add special color name to finish
            color_name = self.special_color_input.text().strip()
            if color_name:
                finish = f"Special Color - {color_name}"
            else:
                # No color name entered, use default for display
                finish = "Special Color - Custom"
        quantity = self.quantity_spin.value()
        discount = self.discount_spin.value()
        unit = self.unit_combo.currentText()
        
        # Get special color multiplier if Special Color is selected
        special_color_multiplier = 1.0  # Default multiplier
        if self.finish_combo.currentText() == 'Special Color':
            special_color_multiplier = self.special_color_multiplier_spin.value() / 100.0  # Convert percentage to decimal
        
        # Get product type flags using consolidated helper
        if product:
            has_no_dimensions, has_price_per_foot, is_other_table = get_product_type_flags(self.price_loader, product)
        else:
            has_no_dimensions = has_price_per_foot = is_other_table = False
        
        if has_no_dimensions:
            # Handle products with no height/width - extract price_id first
            price_id = self.price_loader.get_price_id_for_no_dimensions(product)
            if price_id is None:
                self.unit_price_label.setText('N/A')
                self.total_price_label.setText('฿ 0.00')
                self.rounded_size_label.setText('N/A')
                return
            
            # Use price_id with appropriate function based on product type
            if has_price_per_foot:
                # For price_per_foot products with no dimensions, use get_price_for_price_per_foot with price_id
                # Note: height is still required for price_per_foot calculation
                height = self.height_spin.value()
                unit = self.unit_combo.currentText()
                height_inches = convert_dimension_to_inches(height, unit)
                unit_price = self.price_loader.get_price_for_price_per_foot(product, finish, 0, height_inches, with_damper, special_color_multiplier, price_id=price_id)
                self.rounded_size_label.setText('N/A')
            elif is_other_table:
                # For other table products with no dimensions, use find_rounded_other_table_size with price_id
                rounded_size = self.price_loader.find_rounded_other_table_size(product, finish, 0, price_id=price_id)
                if not rounded_size:
                    self.unit_price_label.setText('N/A')
                    self.total_price_label.setText('฿ 0.00')
                    self.rounded_size_label.setText('N/A')
                    return
                self.rounded_size_label.setText(rounded_size)
                # Get price using the rounded size
                unit_price = self.price_loader.get_price_for_other_table(product, finish, rounded_size, with_damper, special_color_multiplier)

        elif has_price_per_foot:
            # Handle price_per_foot products - require width and height
            width = self.width_spin.value()
            height = self.height_spin.value()
            
            # Convert to inches if needed
            width_inches = convert_dimension_to_inches(width, unit)
            height_inches = convert_dimension_to_inches(height, unit)
            
            # Find the rounded height that matches database (for other tables, size is stored in height column)
            rounded_height = self.price_loader.find_rounded_price_per_foot_width(product, height_inches)
            
            # Fallback: if lookup fails, try swapping width and height (user might have swapped them)
            if not rounded_height:
                rounded_height = self.price_loader.find_rounded_price_per_foot_width(product, width_inches)
                if rounded_height:
                    # Swap worked, swap the values and show warning
                    width_inches, height_inches = height_inches, width_inches
                    QMessageBox.warning(
                        self,
                        'Dimensions Swapped',
                        f'⚠ Warning: Width and height appear to be swapped.\nUsing {width_inches}" x {height_inches}" instead.'
                    )
                else:
                    self.unit_price_label.setText('N/A')
                    self.total_price_label.setText('฿ 0.00')
                    self.rounded_size_label.setText('N/A')
                    return
            
            # Display the rounded width and height
            self.rounded_size_label.setText(f'{width_inches}" x {rounded_height}"')
            
            # Get price using price_per_foot formula: (width / 12) × price_per_foot
            unit_price = self.price_loader.get_price_for_price_per_foot(product, finish, rounded_height, width_inches, with_damper, special_color_multiplier)
        elif is_other_table:
            # Handle other table products
            size = self.other_table_spin.value()
            
            # Convert to inches if needed
            size_inches = convert_dimension_to_inches(size, unit)
            
            # Find the rounded up size for pricing
            rounded_size = self.price_loader.find_rounded_other_table_size(product, finish, size_inches)
            if not rounded_size:
                self.unit_price_label.setText('N/A')
                self.total_price_label.setText('฿ 0.00')
                self.rounded_size_label.setText('N/A')
                return
            
            # Display the rounded size
            self.rounded_size_label.setText(rounded_size)
            
            # Get price using the rounded size
            unit_price = self.price_loader.get_price_for_other_table(product, finish, rounded_size, with_damper, special_color_multiplier)
        else:
            # Handle width/height-based products
            width = self.width_spin.value()
            height = self.height_spin.value()
            
            # Validate dimensions: width should be greater than height
            if height > width:
                self.unit_price_label.setText('N/A')
                self.total_price_label.setText('฿ 0.00')
                self.rounded_size_label.setText('N/A')
                return
            
            # Convert to inches if needed
            width_inches = convert_dimension_to_inches(width, unit)
            height_inches = convert_dimension_to_inches(height, unit)
            
            # Find the rounded up size for pricing
            rounded_size = self.price_loader.find_rounded_default_table_size(product, finish, width_inches, height_inches)
            
            # If no rounded size found, try to get price directly (for exceeded dimensions)
            if not rounded_size:
                # Try to get price directly with the original dimensions
                unit_price = self.price_loader.get_price_for_default_table(product, finish, f'{width_inches}" x {height_inches}"', with_damper, special_color_multiplier)
                if unit_price == 0:
                    self.unit_price_label.setText('N/A')
                    self.total_price_label.setText('฿ 0.00')
                    self.rounded_size_label.setText('N/A')
                    return
                else:
                    # Display the original size for exceeded dimensions (width x height format)
                    self.rounded_size_label.setText(f'{width_inches}" x {height_inches}"')
            else:
                # Display the rounded size (already in width x height format)
                self.rounded_size_label.setText(rounded_size)
                
                # Get price using the rounded size (format: width x height)
                unit_price = self.price_loader.get_price_for_default_table(product, finish, rounded_size, with_damper, special_color_multiplier)
        
        # Apply discount
        if unit_price is None:
            self.unit_price_label.setText('N/A')
            self.total_price_label.setText('฿ 0.00')
            self.rounded_size_label.setText('N/A')
            return
        
        discount_amount = unit_price * (discount / 100)
        discounted_unit_price = unit_price - discount_amount
        total_price = discounted_unit_price * quantity
        
        # Display the original unit price (without discount)
        self.unit_price_label.setText(f'฿ {unit_price:,.2f}')
        # Display the total price (with discount applied)
        if discount > 0:
            self.total_price_label.setText(f'฿ {total_price:,.2f} (Discounted)')
        else:
            self.total_price_label.setText(f'฿ {total_price:,.2f}')
    
    def add_item_to_quote(self):
        """Add the selected item to the quote"""
        if not self.price_loader:
            return
        
        product_input = self.product_input.text().strip()
        # Extract product name, WD flag, INS flag, and filter type from input
        product, with_damper, has_ins, filter_type = extract_product_flags_and_filter(product_input)
        finish = self.finish_combo.currentText()
        
        # Add powder coating color to finish if Powder Coated is selected
        if finish == 'Powder Coated':
            powder_color = self.powder_color_combo.currentText()
            finish = f"Powder Coated - {powder_color}"
        elif finish == 'Special Color':
            # Add special color name to finish
            color_name = self.special_color_input.text().strip()
            if color_name:
                finish = f"Special Color - {color_name}"
            else:
                QMessageBox.warning(self, 'Missing Color Name', 'Please enter a color name for special color.')
                return
        quantity = self.quantity_spin.value()
        discount = self.discount_spin.value()
        unit = self.unit_combo.currentText()
        
        # Get special color multiplier if Special Color is selected
        special_color_multiplier = 1.0  # Default multiplier
        if self.finish_combo.currentText() == 'Special Color':
            special_color_multiplier = self.special_color_multiplier_spin.value() / 100.0  # Convert percentage to decimal
        
        # Get product type flags using consolidated helper
        has_no_dimensions, has_price_per_foot, is_other_table = get_product_type_flags(self.price_loader, product)
        
        # Get dimensions from UI
        width = None
        height = None
        size = None
        width_unit = unit.lower()
        height_unit = unit.lower()
        size_unit = unit.lower()
        
        # Extract slot number for no-dimension products
        slot_number = None
        if has_no_dimensions:
            slot_number = extract_slot_number_from_model(product_input)
        
        if has_no_dimensions:
            # For no-dimension products, height might still be needed for price_per_foot
            if has_price_per_foot:
                height = self.height_spin.value()
        elif has_price_per_foot or not is_other_table:
            width = self.width_spin.value()
            height = self.height_spin.value()
            
            # Validate dimensions: width should be greater than height (for default products)
            if not has_price_per_foot and height > width:
                QMessageBox.warning(self, 'Invalid Dimensions', 
                                  'Width must be greater than height. Please adjust the dimensions.')
                return
        elif is_other_table:
            size = self.other_table_spin.value()
        
        # Use shared function to build quote item
        item, error = build_quote_item(
            price_loader=self.price_loader,
            product=product,
            finish=finish,
            quantity=quantity,
            has_wd=with_damper,
            has_price_per_foot=has_price_per_foot,
            is_other_table=is_other_table,
            width=width,
            height=height,
            size=size,
            width_unit=width_unit,
            height_unit=height_unit,
            size_unit=size_unit,
            filter_type=filter_type,
            discount=discount,
            special_color_multiplier=special_color_multiplier,
            detail=self.detail_input.text().strip(),
            has_ins=has_ins,
            has_no_dimensions=has_no_dimensions,
            slot_number=slot_number
        )
        
        if error:
            QMessageBox.warning(self, 'Warning', error)
            return
        
        self.quote_items.append(item)
        self.refresh_items_table()
        
        self.statusBar().showMessage(f'Added {item["product_code"]} {item["size"]} to quote')
    
    def add_title_to_quote(self):
        """Add a title item to the quote"""
        title = self.title_input.text().strip()
        
        if not title:
            QMessageBox.warning(self, 'Warning', 'Please enter a title')
            return
        
        # Create title item
        item = {
            'is_title': True,
            'title': title,
            'product_code': '',  # No product code for titles
            'size': '',
            'finish': '',
            'quantity': 0,
            'unit_price': 0,
            'discount': 0,
            'discounted_unit_price': 0,
            'total': 0,
            'rounded_size': None,
            'detail': ''  # No detail for titles
        }
        
        self.quote_items.append(item)
        self.refresh_items_table()
        self.title_input.clear()  # Clear the input after adding
        
        self.statusBar().showMessage(f'Added title: {title}')
    
    def refresh_items_table(self):
        """Refresh the items table display"""
        self.items_table.setRowCount(len(self.quote_items))
        
        grand_total = 0
        item_counter = 1
        
        # Use the base font size for table items (always use the stored base, not the current font)
        adjusted_font_size = max(1, int(self.table_item_base_font_size * self.font_size_multiplier))
        
        for row, item in enumerate(self.quote_items):
            if item.get('is_title', False):
                # Title row - no ID number, show title in product column
                self.items_table.setItem(row, 0, QTableWidgetItem(''))  # No ID for titles
                self.items_table.setItem(row, 1, QTableWidgetItem(item['title']))
                self.items_table.setItem(row, 2, QTableWidgetItem(''))  # No detail for titles
                self.items_table.setItem(row, 3, QTableWidgetItem(''))  # No finish
                self.items_table.setItem(row, 4, QTableWidgetItem(''))  # No size
                self.items_table.setItem(row, 5, QTableWidgetItem(''))  # No quantity
                self.items_table.setItem(row, 6, QTableWidgetItem(''))  # No unit price
                self.items_table.setItem(row, 7, QTableWidgetItem(''))  # No discount
                self.items_table.setItem(row, 8, QTableWidgetItem(''))  # No total
                
                # Style the title row differently
                for col in range(9):
                    cell = self.items_table.item(row, col)
                    if cell:
                        cell.setBackground(QColor(240, 240, 240))  # Light gray background
                        # Make title text bold and apply font size
                        font = cell.font()
                        font.setPointSize(adjusted_font_size)
                        if col == 1:  # Product column where title is displayed
                            font.setBold(True)
                        cell.setFont(font)
            elif item.get('is_invalid', False):
                # Invalid item row - show error information
                self.items_table.setItem(row, 0, QTableWidgetItem(str(item_counter)))
                product_text = item['product_code']
                if item.get('error_message'):
                    product_text += f" (ERROR: {item['error_message']})"
                self.items_table.setItem(row, 1, QTableWidgetItem(product_text))
                self.items_table.setItem(row, 2, QTableWidgetItem(item.get('detail', '')))
                self.items_table.setItem(row, 3, QTableWidgetItem(item['finish']))
                self.items_table.setItem(row, 4, QTableWidgetItem(item['size']))
                self.items_table.setItem(row, 5, QTableWidgetItem(str(item['quantity'])))
                self.items_table.setItem(row, 6, QTableWidgetItem('N/A'))
                self.items_table.setItem(row, 7, QTableWidgetItem('N/A'))
                self.items_table.setItem(row, 8, QTableWidgetItem('N/A'))
                
                # Style invalid items with red background
                for col in range(9):
                    cell = self.items_table.item(row, col)
                    if cell:
                        cell.setBackground(QColor(255, 200, 200))  # Light red background
                        cell.setForeground(QColor(180, 0, 0))  # Dark red text
                        font = cell.font()
                        font.setPointSize(adjusted_font_size)
                        cell.setFont(font)
                
                item_counter += 1
            elif item.get('warning_message'):
                # Warning item row - show warning information (similar to errors but with yellow background)
                self.items_table.setItem(row, 0, QTableWidgetItem(str(item_counter)))
                product_text = item['product_code']
                product_text += f" (WARNING: {item.get('warning_message')})"
                self.items_table.setItem(row, 1, QTableWidgetItem(product_text))
                self.items_table.setItem(row, 2, QTableWidgetItem(item.get('detail', '')))  # Detail column
                self.items_table.setItem(row, 3, QTableWidgetItem(item['finish']))
                self.items_table.setItem(row, 4, QTableWidgetItem(item['size']))
                self.items_table.setItem(row, 5, QTableWidgetItem(str(item['quantity'])))
                
                # Show original unit price
                self.items_table.setItem(row, 6, QTableWidgetItem(f"฿ {item['unit_price']:,.2f}"))
                
                # Show discount percentage
                discount_percent = item.get('discount', 0) * 100
                if discount_percent > 0:
                    self.items_table.setItem(row, 7, QTableWidgetItem(f"{discount_percent:.0f}%"))
                else:
                    self.items_table.setItem(row, 7, QTableWidgetItem("0%"))
                
                # Show total (after discount)
                self.items_table.setItem(row, 8, QTableWidgetItem(f"฿ {item['total']:,.2f}"))
                
                # Style warning items with yellow background
                for col in range(9):
                    cell = self.items_table.item(row, col)
                    if cell:
                        cell.setBackground(QColor(255, 255, 200))  # Light yellow background
                        font = cell.font()
                        font.setPointSize(adjusted_font_size)
                        cell.setFont(font)
                
                # Only add to grand total if not invalid
                grand_total += item['total']
                item_counter += 1
            else:
                # Regular product row
                self.items_table.setItem(row, 0, QTableWidgetItem(str(item_counter)))
                self.items_table.setItem(row, 1, QTableWidgetItem(item['product_code']))
                self.items_table.setItem(row, 2, QTableWidgetItem(item.get('detail', '')))  # Detail column
                self.items_table.setItem(row, 3, QTableWidgetItem(item['finish']))
                self.items_table.setItem(row, 4, QTableWidgetItem(item['size']))
                self.items_table.setItem(row, 5, QTableWidgetItem(str(item['quantity'])))
                
                # Show original unit price
                self.items_table.setItem(row, 6, QTableWidgetItem(f"฿ {item['unit_price']:,.2f}"))
                
                # Show discount percentage
                discount_percent = item.get('discount', 0) * 100
                if discount_percent > 0:
                    self.items_table.setItem(row, 7, QTableWidgetItem(f"{discount_percent:.0f}%"))
                else:
                    self.items_table.setItem(row, 7, QTableWidgetItem("0%"))
                
                # Show total (after discount)
                self.items_table.setItem(row, 8, QTableWidgetItem(f"฿ {item['total']:,.2f}"))
                
                # Apply font size and styling to all cells in this row
                for col in range(9):
                    cell = self.items_table.item(row, col)
                    if cell:
                        font = cell.font()
                        font.setPointSize(adjusted_font_size)
                        cell.setFont(font)
                
                # Only add to grand total if not invalid
                if not item.get('is_invalid', False):
                    grand_total += item['total']
                item_counter += 1
        
        self.grand_total_label.setText(f'Grand Total: ฿ {grand_total:,.2f}')
        
        # Update move button states based on selection
        self.update_move_button_states()
    
    def update_move_button_states(self):
        """Update the enabled state of move up/down buttons based on current selection"""
        current_row = self.items_table.currentRow()
        total_rows = len(self.quote_items)
        
        # Enable/disable move up button (disabled if first row or no selection)
        self.move_up_button.setEnabled(current_row > 0 and current_row < total_rows)
        
        # Enable/disable move down button (disabled if last row or no selection)
        self.move_down_button.setEnabled(current_row >= 0 and current_row < total_rows - 1)
    
    def move_item_up(self):
        """Move the selected item up one position"""
        current_row = self.items_table.currentRow()
        if current_row > 0 and current_row < len(self.quote_items):
            # Swap items
            self.quote_items[current_row], self.quote_items[current_row - 1] = \
                self.quote_items[current_row - 1], self.quote_items[current_row]
            
            # Refresh table and maintain selection on the moved item
            self.refresh_items_table()
            self.items_table.selectRow(current_row - 1)
            self.update_move_button_states()
            self.statusBar().showMessage('Item moved up')
    
    def move_item_down(self):
        """Move the selected item down one position"""
        current_row = self.items_table.currentRow()
        if current_row >= 0 and current_row < len(self.quote_items) - 1:
            # Swap items
            self.quote_items[current_row], self.quote_items[current_row + 1] = \
                self.quote_items[current_row + 1], self.quote_items[current_row]
            
            # Refresh table and maintain selection on the moved item
            self.refresh_items_table()
            self.items_table.selectRow(current_row + 1)
            self.update_move_button_states()
            self.statusBar().showMessage('Item moved down')
    
    def remove_selected_item(self):
        """Remove the selected item from the quote"""
        current_row = self.items_table.currentRow()
        if current_row >= 0:
            del self.quote_items[current_row]
            self.refresh_items_table()
            self.update_move_button_states()
            self.statusBar().showMessage('Item removed')
    
    def clear_all_items(self):
        """Clear all items from the quote"""
        reply = QMessageBox.question(self, 'Confirm', 
                                    'Are you sure you want to clear all items?',
                                    QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.quote_items = []
            self.refresh_items_table()
            self.update_move_button_states()
            self.statusBar().showMessage('All items cleared')
    
    def new_quote(self):
        """Start a new quote"""
        if self.quote_items:
            reply = QMessageBox.question(self, 'Confirm', 
                                        'Start a new quote? Current items will be cleared.',
                                        QMessageBox.Yes | QMessageBox.No)
            if reply == QMessageBox.No:
                return
        
        self.quote_items = []
        self.to_input.clear()
        self.company_input.clear()
        self.tel_input.clear()
        self.fax_input.clear()
        self.project_input.clear()
        self.detail_input.clear()
        self.quote_number.setText(f"{datetime.now().strftime('%y-%m')}{datetime.now().day:03d}")
        self.refresh_items_table()
        self.statusBar().showMessage('New quote started')
    
    def store_original_fonts(self):
        """Store original font sizes for all widgets"""
        # Find all widgets recursively
        all_widgets = self.findChildren(QWidget)
        
        for widget in all_widgets:
            widget_id = id(widget)
            if widget_id not in self.original_fonts:
                # Get current font
                current_font = widget.font()
                self.original_fonts[widget_id] = {
                    'font': QFont(current_font),  # Create a copy
                    'widget': widget
                }
        
        # Also store fonts for special widgets
        if hasattr(self, 'statusBar'):
            widget_id = id(self.statusBar())
            if widget_id not in self.original_fonts:
                status_font = self.statusBar().font()
                self.original_fonts[widget_id] = {
                    'font': QFont(status_font),
                    'widget': self.statusBar()
                }
    
    def apply_font_size(self, multiplier):
        """Apply font size multiplier to all widgets"""
        # Re-store fonts to catch any newly created widgets
        self.store_original_fonts()
        
        # Find all widgets and apply font size
        all_widgets = self.findChildren(QWidget)
        
        for widget in all_widgets:
            widget_id = id(widget)
            if widget_id in self.original_fonts:
                original_font = self.original_fonts[widget_id]['font']
                new_size = max(1, int(original_font.pointSize() * multiplier))
                new_font = QFont(original_font)
                new_font.setPointSize(new_size)
                widget.setFont(new_font)
        
        # Update status bar
        if hasattr(self, 'statusBar'):
            status_bar_id = id(self.statusBar())
            if status_bar_id in self.original_fonts:
                original_font = self.original_fonts[status_bar_id]['font']
                new_status_size = max(1, int(original_font.pointSize() * multiplier))
                new_font = QFont(original_font)
                new_font.setPointSize(new_status_size)
                self.statusBar().setFont(new_font)
        
        # Update table headers and items if table exists
        if hasattr(self, 'items_table'):
            header = self.items_table.horizontalHeader()
            header_font = header.font()
            new_header_size = max(1, int(header_font.pointSize() * multiplier))
            header_font.setPointSize(new_header_size)
            header.setFont(header_font)
            
            # Update font for existing table items (use base font size, not current)
            base_item_size = self.table_item_base_font_size
            new_item_size = max(1, int(base_item_size * multiplier))
            for row in range(self.items_table.rowCount()):
                for col in range(self.items_table.columnCount()):
                    item = self.items_table.item(row, col)
                    if item:
                        item_font = item.font()
                        item_font.setPointSize(new_item_size)
                        item.setFont(item_font)
    
    def increase_text_size(self):
        """Increase text size by 10%"""
        min_size = 0.5  # Minimum 50%
        max_size = 2.0  # Maximum 200%
        step = 0.1  # 10% increments
        
        new_multiplier = min(max_size, self.font_size_multiplier + step)
        
        if new_multiplier != self.font_size_multiplier:
            self.font_size_multiplier = new_multiplier
            self.apply_font_size(new_multiplier)
            self.update_text_size_controls()
            self.statusBar().showMessage(f'Text size increased to {new_multiplier:.0%}', 2000)
    
    def decrease_text_size(self):
        """Decrease text size by 10%"""
        min_size = 0.5  # Minimum 50%
        max_size = 2.0  # Maximum 200%
        step = 0.1  # 10% increments
        
        new_multiplier = max(min_size, self.font_size_multiplier - step)
        
        if new_multiplier != self.font_size_multiplier:
            self.font_size_multiplier = new_multiplier
            self.apply_font_size(new_multiplier)
            self.update_text_size_controls()
            self.statusBar().showMessage(f'Text size decreased to {new_multiplier:.0%}', 2000)
    
    def update_text_size_controls(self):
        """Update the text size label and enable/disable buttons based on limits"""
        # Update label
        self.text_size_label.setText(f'{int(self.font_size_multiplier * 100)}%')
        
        # Enable/disable buttons based on limits
        min_size = 0.5
        max_size = 2.0
        
        self.text_size_decrease_button.setEnabled(self.font_size_multiplier > min_size)
        self.text_size_increase_button.setEnabled(self.font_size_multiplier < max_size)
    
    def upload_excel_file(self):
        """Handle Excel file upload and extract items with progress dialog"""
        if not self.price_loader:
            QMessageBox.warning(self, 'Warning', 'Price database not loaded. Please wait for the database to load.')
            return
        
        # Open file dialog
        file_name, _ = QFileDialog.getOpenFileName(
            self, 'Select Excel File', '', 'Excel Files (*.xlsx *.xls);;All Files (*)'
        )
        
        if not file_name:
            return
        
        # Create and show progress dialog
        progress_dialog = ExcelUploadProgressDialog(self)
        progress_dialog.show()
        progress_dialog.update_progress(0, 'Opening Excel file...')
        
        try:
            # Create Excel importer
            progress_dialog.update_progress(5, 'Initializing importer...')
            importer = ExcelItemImporter(self.price_loader, self.available_models)
            
            # Parse the Excel file with progress callback
            # The parsing will take 5% to 90% of progress
            def parse_progress_callback(percent, status):
                # Map parser's 0-100% to our 5-90% range
                mapped_percent = 5 + int((percent / 100) * 85)
                progress_dialog.update_progress(mapped_percent, status)
            
            items = importer.parse_excel_file(file_name, progress_callback=parse_progress_callback)
            
            if not items:
                progress_dialog.close()
                QMessageBox.warning(self, 'Warning', 'No items found in the Excel file. Please check the file format.')
                return
            
            # Add items to quote with progress updates (90% to 98%)
            total_items = len(items)
            added_count = 0
            title_count = 0
            invalid_items = []  # Store items with errors
            warnings = []  # Store warnings
            
            progress_dialog.update_progress(90, f'Processing {total_items} item(s)...')
            
            for idx, item in enumerate(items):
                # Update progress (90% to 98%)
                # Use idx + 1 to ensure progress advances even for single item
                if total_items > 0:
                    progress = 90 + int(((idx + 1) / total_items) * 8)
                    progress = min(progress, 98)  # Cap at 98% before finalization
                else:
                    progress = 98
                progress_dialog.update_progress(progress, f'Processing item {idx + 1} of {total_items}...')
                
                if item.get('is_title', False):
                    self.quote_items.append(item)
                    title_count += 1
                else:
                    # Validate and add product item
                    result = importer.add_item_from_excel(item)
                    if result['success']:
                        self.quote_items.append(result['item'])
                        added_count += 1
                    else:
                        # Add as invalid item
                        invalid_item = importer._create_invalid_item(item, result['error'])
                        self.quote_items.append(invalid_item)
                        invalid_items.append({
                            'model': item.get('model', 'Unknown'),
                            'error': result['error']
                        })
            
            # Refresh the table
            progress_dialog.update_progress(99, 'Finalizing...')
            self.refresh_items_table()
            
            # Show results in dialog
            progress_dialog.update_progress(100, 'Complete!')
            # Process events to ensure 100% is visible before showing results
            QApplication.processEvents()
            progress_dialog.show_results(added_count, title_count, invalid_items, warnings)
            
            # Update status bar
            self.statusBar().showMessage(f'Imported {added_count} items, {title_count} titles, {len(invalid_items)} invalid items from Excel file')
            
        except Exception as e:
            progress_dialog.close()
            QMessageBox.critical(self, 'Error', f'Failed to parse Excel file: {str(e)}')
    