from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QLabel,
    QLineEdit,
    QDialogButtonBox,
    QMessageBox,
    QFormLayout,
    QPushButton,
)

from app import storage
from app import auth

def _bootstrap_user_db() -> None:
    """SQL kullanıcı tablosunun varlığını garanti eder."""

    ensure_fn = getattr(storage, "ensure_user_db", None)
    if callable(ensure_fn):
        ensure_fn()
class LoginDialog(QDialog):
    """Kullanıcı adı / şifre girişi için basit dialog."""

    def __init__(self, parent=None):
        super().__init__(parent)
        _bootstrap_user_db()
        self.setWindowTitle("UzmanRapor • Kullanıcı Girişi")
        self.setModal(True)
        self._user = None

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("Lütfen kullanıcı adınızı ve şifrenizi girin."))

        form = QFormLayout()
        self.ed_username = QLineEdit(storage.get_username_default())
        self.ed_password = QLineEdit()
        self.ed_password.setEchoMode(QLineEdit.Password)

        form.addRow("Kullanıcı adı", self.ed_username)
        form.addRow("Şifre", self.ed_password)
        layout.addLayout(form)

        self.lbl_hint = QLabel("Size verilen şifre UzmanRapor'a özeldir")
        self.lbl_hint.setWordWrap(True)
        self.lbl_hint.setAlignment(Qt.AlignLeft | Qt.AlignTop)
        self.lbl_hint.setStyleSheet("color: #555; font-size: 11px;")
        layout.addWidget(self.lbl_hint)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        self.btn_login: QPushButton = buttons.button(QDialogButtonBox.Ok)
        self.btn_login.setText("Giriş")
        buttons.button(QDialogButtonBox.Cancel).setText("Vazgeç")
        layout.addWidget(buttons)

        buttons.accepted.connect(self._try_login)
        buttons.rejected.connect(self.reject)

        self.ed_password.returnPressed.connect(self._try_login)
        self.ed_username.returnPressed.connect(self._focus_password)

        self.resize(420, 200)
        self.ed_username.setFocus()

    @property
    def user(self):
        return self._user

    def _focus_password(self):
        self.ed_password.setFocus()
        self.ed_password.selectAll()

    def _try_login(self):
        username = self.ed_username.text().strip()
        password = self.ed_password.text()

        user = auth.authenticate(username, password)
        if not user:
            QMessageBox.warning(self, "Giriş başarısız", "Kullanıcı adı veya şifre hatalı.")
            self.ed_password.selectAll()
            self.ed_password.setFocus()
            return

        if not user.has_permission("read"):
            QMessageBox.warning(self, "Yetki bulunamadı", "Bu kullanıcı okuma yetkisine sahip değil.")
            return

        self._user = user
        storage.set_username_default(user.username)
        self.accept()