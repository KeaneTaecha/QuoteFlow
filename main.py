"""
Quotation System - Main Entry Point
Simple entry point for the quotation application.
"""

import sys
from PyQt5.QtWidgets import QApplication
from quotation_ui import QuotationApp


def main():
    """Main application entry point"""
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    
    window = QuotationApp()
    window.show()
    
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()

