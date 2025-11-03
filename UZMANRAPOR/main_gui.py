
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from PySide6.QtWidgets import QApplication
from app.gui import MainWindow
def main():
    app = QApplication(sys.argv)
    win = MainWindow(); win.show()
    sys.exit(app.exec())
if __name__ == "__main__":
    main()
