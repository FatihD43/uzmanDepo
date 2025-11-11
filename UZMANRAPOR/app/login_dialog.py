from __future__ import annotations

import json
import secrets
from pathlib import Path

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
    """Ensure the users.json file exists even on deployments without storage.ensure_user_db."""

    ensure_fn = getattr(storage, "ensure_user_db", None)
    if callable(ensure_fn):
        ensure_fn()
        return

    # Backwards-compatible fallback: create the default admin user if the helper is missing.
    user_db_path = getattr(storage, "USERS_DB_PATH", None)
    if user_db_path is None:
        app_dir = Path.home() / ".uzman_rapor"
        user_db_path = app_dir / "users.json"
    else:
        user_db_path = Path(user_db_path)

    if user_db_path.exists():
        return

    user_db_path.parent.mkdir(parents=True, exist_ok=True)

    hash_fn = getattr(storage, "hash_password", None)
    if callable(hash_fn):
        salt = secrets.token_hex(16)
        password_hash = hash_fn("admin", salt)
    else:
        # Minimal fallback – store the password as plain text if hashing helper is absent.
        salt = ""
        password_hash = "admin"

    payload = {
        "users": [
            {
                "username": "admin",
                "salt": salt,
                "password_hash": password_hash,
                "permissions": ["admin", "read", "write"],
            }
        ]
    }

    try:
        with open(user_db_path, "w", encoding="utf-8") as fh:
            json.dump(payload, fh, ensure_ascii=False, indent=2)
    except Exception:
        # Silently ignore; login flow will surface an authentication error later if this fails.
        pass

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

        self.lbl_hint = QLabel(
            "Varsayılan kullanıcı: <b>admin</b> / <b>admin</b><br>"
            "Kullanıcı listesi: ~/.uzman_rapor/users.json"
        )
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