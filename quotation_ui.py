"""
Quotation UI Module
Contains the main PyQt5 GUI for the quotation system.
"""

import sys
import os
from datetime import datetime
from PyQt5.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QLabel, QComboBox, QSpinBox, QPushButton,
                             QTableWidget, QTableWidgetItem, QGroupBox, QCheckBox,
                             QLineEdit, QMessageBox, QFileDialog, QHeaderView)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont
from PyQt5.QtPrintSupport import QPrinter, QPrintDialog

from price_loader import PriceListLoader


class QuotationApp(QMainWindow):
    """Main application window for the quotation system"""
    
    def __init__(self):
        super().__init__()
        self.quote_items = []
        self.price_loader = None
        self.init_ui()
        self.load_price_list()
    
    def init_ui(self):
        """Initialize the user interface"""
        self.setWindowTitle('Quotation System - HRG & WSG Products')
        self.setGeometry(100, 100, 1400, 800)
        
        # Main widget and layout
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout()
        main_widget.setLayout(main_layout)
        
        # Title
        title = QLabel('QUOTATION MANAGEMENT SYSTEM')
        title.setFont(QFont('Arial', 16, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title)
        
        # Create sections
        main_layout.addWidget(self.create_quote_info_section())
        main_layout.addWidget(self.create_product_selection_section())
        main_layout.addWidget(self.create_items_table_section())
        main_layout.addWidget(self.create_action_buttons())
        
        self.statusBar().showMessage('Ready')
    
    def create_quote_info_section(self):
        """Create the quote information input section"""
        group = QGroupBox('Quote Information')
        layout = QHBoxLayout()
        
        # Customer Name
        layout.addWidget(QLabel('Customer:'))
        self.customer_input = QLineEdit()
        self.customer_input.setPlaceholderText('Customer Name')
        layout.addWidget(self.customer_input)
        
        # Quote Number
        layout.addWidget(QLabel('Quote #:'))
        self.quote_number = QLineEdit()
        self.quote_number.setText(f"Q{datetime.now().strftime('%Y%m%d')}-001")
        layout.addWidget(self.quote_number)
        
        # Date
        layout.addWidget(QLabel('Date:'))
        self.quote_date = QLineEdit()
        self.quote_date.setText(datetime.now().strftime('%Y-%m-%d'))
        self.quote_date.setReadOnly(True)
        layout.addWidget(self.quote_date)
        
        group.setLayout(layout)
        return group
    
    def create_product_selection_section(self):
        """Create the product selection panel"""
        group = QGroupBox('Product Selection')
        layout = QHBoxLayout()
        
        # Product Type
        prod_layout = QVBoxLayout()
        prod_layout.addWidget(QLabel('Product Type:'))
        self.product_combo = QComboBox()
        # Products will be loaded from the database
        self.product_combo.currentTextChanged.connect(self.on_product_changed)
        prod_layout.addWidget(self.product_combo)
        layout.addLayout(prod_layout)
        
        # Finish
        finish_layout = QVBoxLayout()
        finish_layout.addWidget(QLabel('Finish:'))
        self.finish_combo = QComboBox()
        self.finish_combo.addItems(['Anodized Aluminum', 'White Powder Coated'])
        self.finish_combo.currentTextChanged.connect(self.on_selection_changed)
        finish_layout.addWidget(self.finish_combo)
        layout.addLayout(finish_layout)
        
        # Unit Selection
        unit_layout = QVBoxLayout()
        unit_layout.addWidget(QLabel('Unit:'))
        self.unit_combo = QComboBox()
        self.unit_combo.addItems(['Inches', 'Millimeters'])
        self.unit_combo.currentTextChanged.connect(self.on_unit_changed)
        unit_layout.addWidget(self.unit_combo)
        layout.addLayout(unit_layout)
        
        # Width
        width_layout = QVBoxLayout()
        self.width_label = QLabel('Width (inches):')
        width_layout.addWidget(self.width_label)
        self.width_spin = QSpinBox()
        self.width_spin.setMinimum(1)
        self.width_spin.setMaximum(100)
        self.width_spin.setValue(4)
        self.width_spin.valueChanged.connect(self.update_price_display)
        width_layout.addWidget(self.width_spin)
        layout.addLayout(width_layout)
        
        # Height
        height_layout = QVBoxLayout()
        self.height_label = QLabel('Height (inches):')
        height_layout.addWidget(self.height_label)
        self.height_spin = QSpinBox()
        self.height_spin.setMinimum(1)
        self.height_spin.setMaximum(100)
        self.height_spin.setValue(4)
        self.height_spin.valueChanged.connect(self.update_price_display)
        height_layout.addWidget(self.height_spin)
        layout.addLayout(height_layout)
        
        # With Damper
        damper_layout = QVBoxLayout()
        damper_layout.addWidget(QLabel('Options:'))
        self.damper_check = QCheckBox('With Damper (WD)')
        self.damper_check.stateChanged.connect(self.update_price_display)
        damper_layout.addWidget(self.damper_check)
        layout.addLayout(damper_layout)
        
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
        
        # Unit Price Display
        price_layout = QVBoxLayout()
        price_layout.addWidget(QLabel('Unit Price:'))
        self.unit_price_label = QLabel('฿ 0.00')
        self.unit_price_label.setFont(QFont('Arial', 12, QFont.Bold))
        self.unit_price_label.setStyleSheet('color: #2E7D32; padding: 5px;')
        price_layout.addWidget(self.unit_price_label)
        layout.addLayout(price_layout)
        
        # Total Price Display
        total_layout = QVBoxLayout()
        total_layout.addWidget(QLabel('Total:'))
        self.total_price_label = QLabel('฿ 0.00')
        self.total_price_label.setFont(QFont('Arial', 12, QFont.Bold))
        self.total_price_label.setStyleSheet('color: #1565C0; padding: 5px;')
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
        return group
    
    def create_items_table_section(self):
        """Create the quote items table"""
        group = QGroupBox('Quote Items')
        layout = QVBoxLayout()
        
        # Table
        self.items_table = QTableWidget()
        self.items_table.setColumnCount(7)
        self.items_table.setHorizontalHeaderLabels([
            'Item', 'Product', 'Size', 'Finish', 'Qty', 'Unit Price', 'Total'
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
        """Create action buttons for save, print, etc."""
        widget = QWidget()
        layout = QHBoxLayout()
        
        self.new_button = QPushButton('New Quote')
        self.new_button.clicked.connect(self.new_quote)
        layout.addWidget(self.new_button)
        
        self.save_button = QPushButton('Save Quote')
        self.save_button.clicked.connect(self.save_quote)
        layout.addWidget(self.save_button)
        
        self.print_button = QPushButton('Print Quote')
        self.print_button.clicked.connect(self.print_quote)
        layout.addWidget(self.print_button)
        
        layout.addStretch()
        
        self.exit_button = QPushButton('Exit')
        self.exit_button.clicked.connect(self.close)
        layout.addWidget(self.exit_button)
        
        widget.setLayout(layout)
        return widget
    
    def load_price_list(self):
        """Load the SQLite price database"""
        # Handle both development and bundled executable paths
        if getattr(sys, 'frozen', False):
            # Running as compiled executable
            application_path = sys._MEIPASS
        else:
            # Running as script
            application_path = os.path.dirname(os.path.abspath(__file__))
        
        db_file = os.path.join(application_path, 'prices.db')
        
        if not os.path.exists(db_file):
            QMessageBox.critical(self, 'Error', 
                               f'Price database not found: {db_file}\n\n'
                               f'Please run getsql.py first to create the database from the Excel file.')
            return
        
        try:
            self.price_loader = PriceListLoader(db_file)
            
            # Populate product combo with available models from database
            available_models = self.price_loader.get_available_models()
            if available_models:
                self.product_combo.clear()
                self.product_combo.addItems(available_models)
            else:
                QMessageBox.warning(self, 'Warning', 'No products found in the database!')
            
            self.update_price_display()
            self.statusBar().showMessage(f'Price database loaded successfully ({len(available_models)} models found)')
        except Exception as e:
            QMessageBox.critical(self, 'Error', f'Failed to load price database: {str(e)}')
    
    def on_product_changed(self):
        """Handle product type change"""
        self.update_price_display()
    
    def on_selection_changed(self):
        """Handle selection changes"""
        self.update_price_display()
    
    def on_unit_changed(self):
        """Handle unit selection change"""
        unit = self.unit_combo.currentText()
        
        if unit == 'Millimeters':
            # Update labels
            self.width_label.setText('Width (mm):')
            self.height_label.setText('Height (mm):')
            
            # Update spin box ranges for mm (1-2500mm ≈ 1-100 inches)
            self.width_spin.setMinimum(1)
            self.width_spin.setMaximum(2500)
            self.width_spin.setValue(100)  # 100mm ≈ 4 inches
            
            self.height_spin.setMinimum(1)
            self.height_spin.setMaximum(2500)
            self.height_spin.setValue(100)  # 100mm ≈ 4 inches
        else:
            # Update labels
            self.width_label.setText('Width (inches):')
            self.height_label.setText('Height (inches):')
            
            # Update spin box ranges for inches
            self.width_spin.setMinimum(1)
            self.width_spin.setMaximum(100)
            self.width_spin.setValue(4)
            
            self.height_spin.setMinimum(1)
            self.height_spin.setMaximum(100)
            self.height_spin.setValue(4)
        
        self.update_price_display()
    
    def update_price_display(self):
        """Update the price display based on current selections"""
        if not self.price_loader:
            return
        
        product = self.product_combo.currentText()
        finish = self.finish_combo.currentText()
        width = self.width_spin.value()
        height = self.height_spin.value()
        with_damper = self.damper_check.isChecked()
        quantity = self.quantity_spin.value()
        unit = self.unit_combo.currentText()
        
        # Convert to inches if needed (mm to inches: divide by 25.4)
        if unit == 'Millimeters':
            width_inches = width / 25
            height_inches = height / 25
        else:
            width_inches = width
            height_inches = height
        
        # Find the rounded up size for pricing
        rounded_size = self.price_loader.find_rounded_size(product, finish, width_inches, height_inches)
        if not rounded_size:
            self.unit_price_label.setText('Size not available')
            self.total_price_label.setText('฿ 0.00')
            return
        
        # Get price using the rounded size
        unit_price = self.price_loader.get_price(product, finish, rounded_size, with_damper)
        total_price = unit_price * quantity
        
        # Display the price with a note about rounding
        self.unit_price_label.setText(f'฿ {unit_price:,.2f} (rounded from {rounded_size})')
        self.total_price_label.setText(f'฿ {total_price:,.2f}')
    
    def add_item_to_quote(self):
        """Add the selected item to the quote"""
        if not self.price_loader:
            return
        
        product = self.product_combo.currentText()
        finish = self.finish_combo.currentText()
        width = self.width_spin.value()
        height = self.height_spin.value()
        with_damper = self.damper_check.isChecked()
        quantity = self.quantity_spin.value()
        unit = self.unit_combo.currentText()
        
        # Convert to inches if needed (mm to inches: divide by 25.4)
        if unit == 'Millimeters':
            width_inches = width / 25.4
            height_inches = height / 25.4
        else:
            width_inches = width
            height_inches = height
        
        # Find the rounded up size for pricing
        rounded_size = self.price_loader.find_rounded_size(product, finish, width_inches, height_inches)
        if not rounded_size:
            QMessageBox.warning(self, 'Warning', 'Size not available in price list')
            return
        
        # Get price using the rounded size
        unit_price = self.price_loader.get_price(product, finish, rounded_size, with_damper)
        
        if unit_price == 0:
            QMessageBox.warning(self, 'Warning', 'Price not available for this configuration')
            return
        
        total_price = unit_price * quantity
        
        # Create product description
        product_code = product
        if with_damper:
            product_code += "(WD)"
        
        # Store the original size entered by user with proper unit
        if unit == 'Millimeters':
            original_size = f"{width}mm x {height}mm"
        else:
            original_size = f"{width}\" x {height}\""
        
        item = {
            'product_code': product_code,
            'size': original_size,  # Store original size for display
            'finish': finish,
            'quantity': quantity,
            'unit_price': unit_price,
            'total': total_price,
            'rounded_size': rounded_size  # Store rounded size for reference
        }
        
        self.quote_items.append(item)
        self.refresh_items_table()
        
        self.statusBar().showMessage(f'Added {product_code} {original_size} to quote (priced as {rounded_size})')
    
    def refresh_items_table(self):
        """Refresh the items table display"""
        self.items_table.setRowCount(len(self.quote_items))
        
        grand_total = 0
        for row, item in enumerate(self.quote_items):
            self.items_table.setItem(row, 0, QTableWidgetItem(str(row + 1)))
            self.items_table.setItem(row, 1, QTableWidgetItem(item['product_code']))
            self.items_table.setItem(row, 2, QTableWidgetItem(item['size']))
            self.items_table.setItem(row, 3, QTableWidgetItem(item['finish']))
            self.items_table.setItem(row, 4, QTableWidgetItem(str(item['quantity'])))
            self.items_table.setItem(row, 5, QTableWidgetItem(f"฿ {item['unit_price']:,.2f}"))
            self.items_table.setItem(row, 6, QTableWidgetItem(f"฿ {item['total']:,.2f}"))
            
            grand_total += item['total']
        
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
        self.customer_input.clear()
        self.quote_number.setText(f"Q{datetime.now().strftime('%Y%m%d')}-001")
        self.refresh_items_table()
        self.statusBar().showMessage('New quote started')
    
    def save_quote(self):
        """Save the quote to a file"""
        if not self.quote_items:
            QMessageBox.warning(self, 'Warning', 'No items in quote to save')
            return
        
        file_name, _ = QFileDialog.getSaveFileName(
            self, 'Save Quote', 
            f"Quote_{self.quote_number.text()}.txt",
            'Text Files (*.txt);;All Files (*)'
        )
        
        if file_name:
            try:
                with open(file_name, 'w', encoding='utf-8') as f:
                    f.write('=' * 80 + '\n')
                    f.write('QUOTATION\n')
                    f.write('=' * 80 + '\n\n')
                    f.write(f"Quote Number: {self.quote_number.text()}\n")
                    f.write(f"Date: {self.quote_date.text()}\n")
                    f.write(f"Customer: {self.customer_input.text()}\n\n")
                    f.write('-' * 80 + '\n')
                    f.write(f"{'Item':<6} {'Product':<15} {'Size':<15} {'Finish':<25} {'Qty':<6} {'Unit Price':<12} {'Total':<12}\n")
                    f.write('-' * 80 + '\n')
                    
                    grand_total = 0
                    for idx, item in enumerate(self.quote_items, 1):
                        f.write(f"{idx:<6} {item['product_code']:<15} {item['size']:<15} "
                               f"{item['finish']:<25} {item['quantity']:<6} "
                               f"฿{item['unit_price']:>10,.2f}  ฿{item['total']:>10,.2f}\n")
                        grand_total += item['total']
                    
                    f.write('-' * 80 + '\n')
                    f.write(f"{'GRAND TOTAL:':<70} ฿{grand_total:>10,.2f}\n")
                    f.write('=' * 80 + '\n')
                
                self.statusBar().showMessage(f'Quote saved to {file_name}')
                QMessageBox.information(self, 'Success', 'Quote saved successfully')
            except Exception as e:
                QMessageBox.critical(self, 'Error', f'Failed to save quote: {str(e)}')
    
    def print_quote(self):
        """Print the quote"""
        if not self.quote_items:
            QMessageBox.warning(self, 'Warning', 'No items in quote to print')
            return
        
        printer = QPrinter(QPrinter.HighResolution)
        dialog = QPrintDialog(printer, self)
        
        if dialog.exec_() == QPrintDialog.Accepted:
            self.statusBar().showMessage('Printing...')
            # TODO: Implement actual printing functionality
            QMessageBox.information(self, 'Print', 'Print functionality not yet implemented')

