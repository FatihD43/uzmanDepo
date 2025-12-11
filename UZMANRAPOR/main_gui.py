
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from PySide6.QtWidgets import QApplication, QDialog
from app.gui import MainWindow
from app.login_dialog import LoginDialog
from app.itema_tab import ItemaAyarTab

def main():
    app = QApplication(sys.argv)
    login = LoginDialog()
    if login.exec() != QDialog.Accepted or login.user is None:
        return
    win = MainWindow(user=login.user)
    win.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
