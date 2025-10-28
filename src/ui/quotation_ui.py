"""
Quotation UI Module - Updated with Excel Export
Contains the main PyQt5 GUI for the quotation system with Excel export functionality.
"""

import sys
import os
from datetime import datetime
from PyQt5.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QLabel, QComboBox, QSpinBox, QPushButton,
                             QTableWidget, QTableWidgetItem, QGroupBox, QCheckBox,
                             QLineEdit, QMessageBox, QFileDialog, QHeaderView,
                             QGridLayout, QTextEdit, QDateEdit, QTabWidget, QListWidget)
from PyQt5.QtCore import Qt, QDate
from PyQt5.QtGui import QFont, QColor, QIcon

from database.price_loader import PriceListLoader
from excel_exporter import ExcelQuotationExporter


class QuotationApp(QMainWindow):
    """Main application window for the quotation system"""
    
    def __init__(self):
        super().__init__()
        self.quote_items = []
        self.price_loader = None
        self.excel_exporter = ExcelQuotationExporter()
        self.init_ui()
        self.load_price_list()
    
    def init_ui(self):
        """Initialize the user interface"""
        self.setWindowTitle('Quotation System - HRG & WSG Products')
        self.setGeometry(100, 100, 1500, 900)
        
        # Set window icon
        self.set_window_icon()
        
        # Main widget and layout
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout()
        main_widget.setLayout(main_layout)
        
        # Title
        title = QLabel('ระบบจัดการใบเสนอราคา / QUOTATION MANAGEMENT SYSTEM')
        title.setFont(QFont('Arial', 16, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title)
        
        # Create tab widget for better organization
        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs)
        
        # Tab 1: Main quotation
        main_tab = QWidget()
        main_tab_layout = QVBoxLayout()
        main_tab.setLayout(main_tab_layout)
        
        main_tab_layout.addWidget(self.create_quote_info_section())
        main_tab_layout.addWidget(self.create_product_selection_section())
        main_tab_layout.addWidget(self.create_title_section())
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
        layout.addWidget(self.to_input, 0, 1, 1, 2)
        
        layout.addWidget(QLabel('เลขที่ / NO.:'), 0, 3)
        self.quote_number = QLineEdit()
        self.quote_number.setText(f"{datetime.now().strftime('%y-%m')}{datetime.now().day:03d}")
        layout.addWidget(self.quote_number, 0, 4)
        
        # Row 2: Company and Date
        layout.addWidget(QLabel('บริษัท / COMPANY:'), 1, 0)
        self.company_input = QLineEdit()
        self.company_input.setPlaceholderText('ชื่อบริษัท / Company Name')
        layout.addWidget(self.company_input, 1, 1, 1, 2)
        
        layout.addWidget(QLabel('วันที่ / DATE:'), 1, 3)
        self.quote_date = QDateEdit()
        self.quote_date.setDate(QDate.currentDate())
        self.quote_date.setCalendarPopup(True)
        self.quote_date.setDisplayFormat('yyyy-MM-dd')
        layout.addWidget(self.quote_date, 1, 4)
        
        # Row 3: Tel, Fax, and Project
        layout.addWidget(QLabel('โทร / TEL:'), 2, 0)
        self.tel_input = QLineEdit()
        self.tel_input.setPlaceholderText('เบอร์โทรศัพท์ / Phone Number')
        layout.addWidget(self.tel_input, 2, 1)
        
        layout.addWidget(QLabel('Fax:'), 2, 2)
        self.fax_input = QLineEdit()
        self.fax_input.setPlaceholderText('เบอร์แฟกซ์ / Fax Number')
        layout.addWidget(self.fax_input, 2, 3)
        
        layout.addWidget(QLabel('งาน / PROJECT:'), 2, 4)
        self.project_input = QLineEdit()
        self.project_input.setPlaceholderText('ชื่อโครงการ / Project Name')
        layout.addWidget(self.project_input, 2, 5)
        
        # Set column stretch
        layout.setColumnStretch(1, 2)
        layout.setColumnStretch(3, 1)
        layout.setColumnStretch(5, 2)
        
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
        layout = QHBoxLayout()
        
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
        layout.addLayout(prod_layout)
        
        # Finish
        finish_layout = QVBoxLayout()
        finish_layout.addWidget(QLabel('Finish:'))
        self.finish_combo = QComboBox()
        # Finish options will be populated dynamically based on selected product
        self.finish_combo.currentTextChanged.connect(self.on_selection_changed)
        finish_layout.addWidget(self.finish_combo)
        layout.addLayout(finish_layout)
        
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
        layout.addLayout(self.powder_color_layout)
        
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
        layout.addLayout(self.special_color_layout)
        
        # Initially hide special color input
        self.show_hide_widgets(self.special_color_layout, False)
        
        # Unit Selection
        unit_layout = QVBoxLayout()
        unit_layout.addWidget(QLabel('Unit:'))
        self.unit_combo = QComboBox()
        self.unit_combo.addItems(['Inches', 'Millimeters'])
        self.unit_combo.currentTextChanged.connect(self.on_unit_changed)
        unit_layout.addWidget(self.unit_combo)
        layout.addLayout(unit_layout)
        
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
        layout.addLayout(self.width_layout)
        
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
        layout.addLayout(self.height_layout)
        
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
        layout.addLayout(self.other_table_layout)
        
        # Initially hide other table layout
        self.show_hide_widgets(self.other_table_layout, False)
        
        # With Damper
        self.damper_layout = QVBoxLayout()
        self.damper_layout.addWidget(QLabel('Options:'))
        self.damper_check = QCheckBox('With Damper (WD)')
        self.damper_check.stateChanged.connect(self.update_price_display)
        self.damper_layout.addWidget(self.damper_check)
        layout.addLayout(self.damper_layout)
        
        # Quantity
        qty_layout = QVBoxLayout()
        qty_layout.addWidget(QLabel('Quantity:'))
        self.quantity_spin = QSpinBox()
        self.quantity_spin.setMinimum(1)
        self.quantity_spin.setMaximum(9999)
        self.quantity_spin.setValue(1)
        self.quantity_spin.valueChanged.connect(self.update_price_display)
        qty_layout.addWidget(self.quantity_spin)
        layout.addLayout(qty_layout)
        
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
        layout.addLayout(discount_layout)
        
        # Unit Price Display
        price_layout = QVBoxLayout()
        price_layout.addWidget(QLabel('Unit Price:'))
        self.unit_price_label = QLabel('฿ 0.00')
        self.unit_price_label.setFont(QFont('Arial', 12, QFont.Bold))
        self.unit_price_label.setStyleSheet('color: #2E7D32; padding: 5px;')
        self.unit_price_label.setMinimumWidth(120)  # Fixed width to prevent shifting
        price_layout.addWidget(self.unit_price_label)
        layout.addLayout(price_layout)
        
        # Rounded Size Display
        rounded_size_layout = QVBoxLayout()
        rounded_size_layout.addWidget(QLabel('Rounded Size:'))
        self.rounded_size_label = QLabel('')
        self.rounded_size_label.setFont(QFont('Arial', 12, QFont.Bold))
        self.rounded_size_label.setStyleSheet('color: #FF5722; padding: 5px;')
        self.rounded_size_label.setMinimumWidth(120)  # Fixed width to prevent shifting
        rounded_size_layout.addWidget(self.rounded_size_label)
        layout.addLayout(rounded_size_layout)
        
        # Total Price Display
        total_layout = QVBoxLayout()
        total_layout.addWidget(QLabel('Total:'))
        self.total_price_label = QLabel('฿ 0.00')
        self.total_price_label.setFont(QFont('Arial', 12, QFont.Bold))
        self.total_price_label.setStyleSheet('color: #1565C0; padding: 5px;')
        self.total_price_label.setMinimumWidth(120)  # Fixed width to prevent shifting
        total_layout.addWidget(self.total_price_label)
        layout.addLayout(total_layout)
        
        # Add Button
        add_layout = QVBoxLayout()
        add_layout.addWidget(QLabel(''))  # Spacer
        self.add_button = QPushButton('Add to Quote')
        self.add_button.setStyleSheet('background-color: #4CAF50; color: white; padding: 10px; font-weight: bold;')
        self.add_button.clicked.connect(self.add_item_to_quote)
        add_layout.addWidget(self.add_button)
        layout.addLayout(add_layout)
        
        group.setLayout(layout)
        
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
        self.items_table.setColumnCount(8)
        self.items_table.setHorizontalHeaderLabels([
            'Item', 'Product', 'Size', 'Finish', 'Qty', 'Unit Price', 'Discount', 'Total'
        ])
        
        # Set column widths
        header = self.items_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.Stretch)
        header.setSectionResizeMode(4, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(5, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(6, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(7, QHeaderView.ResizeToContents)
        
        layout.addWidget(self.items_table)
        
        # Buttons for table operations
        button_layout = QHBoxLayout()
        
        self.remove_button = QPushButton('Remove Selected Item')
        self.remove_button.clicked.connect(self.remove_selected_item)
        button_layout.addWidget(self.remove_button)
        
        self.clear_button = QPushButton('Clear All Items')
        self.clear_button.clicked.connect(self.clear_all_items)
        button_layout.addWidget(self.clear_button)
        
        button_layout.addStretch()
        
        # Grand Total
        self.grand_total_label = QLabel('Grand Total: ฿ 0.00')
        self.grand_total_label.setFont(QFont('Arial', 14, QFont.Bold))
        self.grand_total_label.setStyleSheet('color: #C62828;')
        button_layout.addWidget(self.grand_total_label)
        
        layout.addLayout(button_layout)
        
        group.setLayout(layout)
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
            self.price_loader = PriceListLoader(db_file)
            
            # Store available models for searching
            self.available_models = self.price_loader.get_available_models()
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
        product = self.product_input.text().strip()
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
            
            # Check if this product uses other table format instead of width/height
            is_other_table = self.price_loader.is_other_table(product)
            
            if is_other_table:
                # Hide width and height fields, show other table size field
                self.show_hide_widgets(self.width_layout, False)
                self.show_hide_widgets(self.height_layout, False)
                self.show_hide_widgets(self.other_table_layout, True)
            else:
                # Show width and height fields, hide other table size field
                self.show_hide_widgets(self.width_layout, True)
                self.show_hide_widgets(self.height_layout, True)
                self.show_hide_widgets(self.other_table_layout, False)
            
            # Show/hide finish-specific options
            finish = self.finish_combo.currentText()
            self.show_hide_widgets(self.powder_color_layout, finish == 'Powder Coated')
            self.show_hide_widgets(self.special_color_layout, finish == 'Special Color')
            
            # Check if damper option is available for this product
            if finish and finish != 'No finishes available':
                has_damper = self.price_loader.has_damper_option(product, finish)
                self.show_hide_widgets(self.damper_layout, has_damper)
                if not has_damper:
                    self.damper_check.setChecked(False)
            else:
                # No valid finish selected, hide damper option
                self.damper_check.setChecked(False)
                self.show_hide_widgets(self.damper_layout, False)
        
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
        
        product = self.product_input.text().strip()
        if not product:
            return
        
        # Validate that the product exists in available models
        if hasattr(self, 'available_models') and product not in self.available_models:
            # Try to find a close match
            matching_models = [model for model in self.available_models 
                              if product.lower() in model.lower()]
            if matching_models:
                # Use the first match
                product = matching_models[0]
                self.product_input.setText(product)
            else:
                # No match found, show error
                QMessageBox.warning(self, 'Product Not Found', 
                                  f'Product "{product}" not found. Please check the spelling or try a different search term.')
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
        
        product = self.product_input.text().strip()
        finish = self.finish_combo.currentText()
        
        if product and finish:
            # Show/hide finish-specific options
            self.show_hide_widgets(self.powder_color_layout, finish == 'Powder Coated')
            self.show_hide_widgets(self.special_color_layout, finish == 'Special Color')
            
            # Check if damper option is available for this product/finish combination
            has_damper = self.price_loader.has_damper_option(product, finish)
            self.show_hide_widgets(self.damper_layout, has_damper)
            if not has_damper:
                self.damper_check.setChecked(False)
        
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
        
        product = self.product_input.text().strip()
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
        
        with_damper = self.damper_check.isChecked()
        quantity = self.quantity_spin.value()
        discount = self.discount_spin.value()
        unit = self.unit_combo.currentText()
        
        # Check if this product uses other table format
        is_other_table = self.price_loader.is_other_table(product)
        
        if is_other_table:
            # Handle other table products
            size = self.other_table_spin.value()
            
            # Convert to inches if needed (mm to inches: divide by 25.4)
            if unit == 'Millimeters':
                size_inches = size / 25
            else:
                size_inches = size
            
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
            unit_price = self.price_loader.get_price_for_other_table(product, finish, rounded_size, with_damper)
        else:
            # Handle width/height-based products
            width = self.width_spin.value()
            height = self.height_spin.value()
            
            # Validate dimensions: width should not be greater than height
            if width > height:
                self.unit_price_label.setText('Invalid')
                self.total_price_label.setText('฿ 0.00')
                self.rounded_size_label.setText('Invalid')
                return
            
            # Convert to inches if needed (mm to inches: divide by 25.4)
            if unit == 'Millimeters':
                width_inches = width / 25
                height_inches = height / 25
            else:
                width_inches = width
                height_inches = height
            
            # Find the rounded up size for pricing
            rounded_size = self.price_loader.find_rounded_default_table_size(product, finish, width_inches, height_inches)
            
            # If no rounded size found, try to get price directly (for exceeded dimensions)
            if not rounded_size:
                # Try to get price directly with the original dimensions
                unit_price = self.price_loader.get_price_for_default_table(product, finish, f'{width_inches}" x {height_inches}"', with_damper)
                if unit_price == 0:
                    self.unit_price_label.setText('N/A')
                    self.total_price_label.setText('฿ 0.00')
                    self.rounded_size_label.setText('N/A')
                    return
                else:
                    # Display the original size for exceeded dimensions
                    self.rounded_size_label.setText(f'{width_inches}" x {height_inches}"')
            else:
                # Display the rounded size
                self.rounded_size_label.setText(rounded_size)
                
                # Get price using the rounded size
                unit_price = self.price_loader.get_price_for_default_table(product, finish, rounded_size, with_damper)
        
        # Apply discount
        if unit_price is None:
            self.unit_price_label.setText('N/A')
            self.total_price_label.setText('฿ 0.00')
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
        
        product = self.product_input.text().strip()
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
        
        with_damper = self.damper_check.isChecked()
        quantity = self.quantity_spin.value()
        discount = self.discount_spin.value()
        unit = self.unit_combo.currentText()
        
        # Check if this product uses other table format
        is_other_table = self.price_loader.is_other_table(product)
        
        if is_other_table:
            # Handle other table products
            size = self.other_table_spin.value()
            
            # Convert to inches if needed
            if unit == 'Millimeters':
                size_inches = size / 25
            else:
                size_inches = size
            
            # Find the rounded up size for pricing
            rounded_size = self.price_loader.find_rounded_other_table_size(product, finish, size_inches)
            
            # If no rounded size found, try to get price directly (for exceeded dimensions)
            if not rounded_size:
                # Try to get price directly with the original size
                unit_price = self.price_loader.get_price_for_other_table(product, finish, f'{size_inches}" diameter', with_damper)
                if unit_price == 0:
                    QMessageBox.warning(self, 'Warning', 'Size not available in price list')
                    return
                else:
                    # Use the original size for exceeded dimensions
                    rounded_size = f'{size_inches}" diameter'
            else:
                # Get price using the rounded size
                unit_price = self.price_loader.get_price_for_other_table(product, finish, rounded_size, with_damper)
            
            if unit_price == 0:
                QMessageBox.warning(self, 'Warning', 'Price not available for this configuration')
                return
            
            # Apply discount
            discount_amount = unit_price * (discount / 100)
            discounted_unit_price = unit_price - discount_amount
            total_price = discounted_unit_price * quantity
            
            # Create product title
            product_code = product
            if with_damper:
                product_code += "(WD)"
            
            # Store the original size entered by user with proper unit
            if unit == 'Millimeters':
                original_size = f"{size}mm"
            else:
                original_size = f'{size}"'
            
            item = {
                'product_code': product_code,
                'size': original_size,  # Original size entered by user
                'finish': finish,
                'quantity': quantity,
                'unit_price': unit_price,  # Original unit price
                'discount': discount / 100,  # Store as decimal (0.1 for 10%)
                'discounted_unit_price': discounted_unit_price,  # Price after discount
                'total': total_price,
                'rounded_size': rounded_size  # Rounded size for pricing
            }
        else:
            # Handle width/height-based products
            width = self.width_spin.value()
            height = self.height_spin.value()
            
            # Validate dimensions: width should not be greater than height
            if width > height:
                QMessageBox.warning(self, 'Invalid Dimensions', 
                                  'Width cannot be greater than height. Please adjust the dimensions.')
                return
            
            # Convert to inches if needed
            if unit == 'Millimeters':
                width_inches = width / 25
                height_inches = height / 25
            else:
                width_inches = width
                height_inches = height
            
            # Find the rounded up size for pricing
            rounded_size = self.price_loader.find_rounded_default_table_size(product, finish, width_inches, height_inches)
            
            # If no rounded size found, try to get price directly (for exceeded dimensions)
            if not rounded_size:
                # Try to get price directly with the original dimensions
                unit_price = self.price_loader.get_price_for_default_table(product, finish, f'{width_inches}" x {height_inches}"', with_damper)
                if unit_price == 0:
                    QMessageBox.warning(self, 'Warning', 'Size not available in price list')
                    return
                else:
                    # Use the original size for exceeded dimensions
                    rounded_size = f'{width_inches}" x {height_inches}"'
            else:
                # Get price using the rounded size
                unit_price = self.price_loader.get_price_for_default_table(product, finish, rounded_size, with_damper)
            
            if unit_price == 0:
                QMessageBox.warning(self, 'Warning', 'Price not available for this configuration')
                return
            
            # Apply discount
            discount_amount = unit_price * (discount / 100)
            discounted_unit_price = unit_price - discount_amount
            total_price = discounted_unit_price * quantity
            
            # Create product title
            product_code = product
            if with_damper:
                product_code += "(WD)"
            
            # Store the original size entered by user with proper unit
            if unit == 'Millimeters':
                original_size = f"{width}mm x {height}mm"
            else:
                original_size = f'{width}" x {height}"'
            
            item = {
                'product_code': product_code,
                'size': original_size,  # Original size entered by user
                'finish': finish,
                'quantity': quantity,
                'unit_price': unit_price,  # Original unit price
                'discount': discount / 100,  # Store as decimal (0.1 for 10%)
                'discounted_unit_price': discounted_unit_price,  # Price after discount
                'total': total_price,
                'rounded_size': rounded_size  # Rounded size for pricing
            }
        
        self.quote_items.append(item)
        self.refresh_items_table()
        
        self.statusBar().showMessage(f'Added {product_code} {item["size"]} to quote')
    
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
            'rounded_size': None
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
        
        for row, item in enumerate(self.quote_items):
            if item.get('is_title', False):
                # Title row - no ID number, show title in product column
                self.items_table.setItem(row, 0, QTableWidgetItem(''))  # No ID for titles
                self.items_table.setItem(row, 1, QTableWidgetItem(item['title']))
                self.items_table.setItem(row, 2, QTableWidgetItem(''))  # No size
                self.items_table.setItem(row, 3, QTableWidgetItem(''))  # No finish
                self.items_table.setItem(row, 4, QTableWidgetItem(''))  # No quantity
                self.items_table.setItem(row, 5, QTableWidgetItem(''))  # No unit price
                self.items_table.setItem(row, 6, QTableWidgetItem(''))  # No discount
                self.items_table.setItem(row, 7, QTableWidgetItem(''))  # No total
                
                # Style the title row differently
                for col in range(8):
                    cell = self.items_table.item(row, col)
                    if cell:
                        cell.setBackground(QColor(240, 240, 240))  # Light gray background
                        # Make title text bold
                        if col == 1:  # Product column where title is displayed
                            font = cell.font()
                            font.setBold(True)
                            cell.setFont(font)
            else:
                # Regular product row
                self.items_table.setItem(row, 0, QTableWidgetItem(str(item_counter)))
                self.items_table.setItem(row, 1, QTableWidgetItem(item['product_code']))
                self.items_table.setItem(row, 2, QTableWidgetItem(item['size']))
                self.items_table.setItem(row, 3, QTableWidgetItem(item['finish']))
                self.items_table.setItem(row, 4, QTableWidgetItem(str(item['quantity'])))
                
                # Show original unit price
                self.items_table.setItem(row, 5, QTableWidgetItem(f"฿ {item['unit_price']:,.2f}"))
                
                # Show discount percentage
                discount_percent = item.get('discount', 0) * 100
                if discount_percent > 0:
                    self.items_table.setItem(row, 6, QTableWidgetItem(f"{discount_percent:.0f}%"))
                else:
                    self.items_table.setItem(row, 6, QTableWidgetItem("0%"))
                
                # Show total (after discount)
                self.items_table.setItem(row, 7, QTableWidgetItem(f"฿ {item['total']:,.2f}"))
                
                grand_total += item['total']
                item_counter += 1
        
        self.grand_total_label.setText(f'Grand Total: ฿ {grand_total:,.2f}')
    
    def remove_selected_item(self):
        """Remove the selected item from the quote"""
        current_row = self.items_table.currentRow()
        if current_row >= 0:
            del self.quote_items[current_row]
            self.refresh_items_table()
            self.statusBar().showMessage('Item removed')
    
    def clear_all_items(self):
        """Clear all items from the quote"""
        reply = QMessageBox.question(self, 'Confirm', 
                                    'Are you sure you want to clear all items?',
                                    QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.quote_items = []
            self.refresh_items_table()
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
        self.quote_number.setText(f"{datetime.now().strftime('%y-%m')}{datetime.now().day:03d}")
        self.refresh_items_table()
        self.statusBar().showMessage('New quote started')
    
    