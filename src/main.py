"""
Quotation System - Main Entry Point
Simple entry point for the quotation application.
"""

import sys
import os

# PyInstaller compatibility
if getattr(sys, 'frozen', False):
    # Running as a bundled executable
    application_path = sys._MEIPASS
    sys.path.insert(0, os.path.join(application_path, 'src'))
else:
    # Running as a script
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from PyQt5.QtWidgets import QApplication
from ui.quotation_ui import QuotationApp


def main():
    """Main application entry point"""
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    
    window = QuotationApp()
    window.show()
    
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()

