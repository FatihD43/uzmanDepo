from __future__ import annotations

import secrets
from typing import Iterable

from PySide6.QtWidgets import (
    QCheckBox,
    QDialog,
    QDialogButtonBox,
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from app import storage


def _normalize_permissions(perms: Iterable[str] | str) -> list[str]:
    if perms is None:
        return []
    if isinstance(perms, str):
        items = perms.split(",")
    else:
        items = perms
    normalized: list[str] = []
    for p in items:
        token = str(p or "").strip()
        if token and token not in normalized:
            normalized.append(token)
    return normalized


class _AddUserDialog(QDialog):
    def __init__(self, existing_usernames: set[str], parent=None):
        super().__init__(parent)
        self.setWindowTitle("Yeni kullanıcı")
        self._existing = {u.lower() for u in existing_usernames}

        layout = QVBoxLayout(self)
        form = QFormLayout()
        self.ed_username = QLineEdit()
        self.ed_password = QLineEdit()
        self.ed_password.setEchoMode(QLineEdit.Password)
        self.ed_permissions = QLineEdit("read,write")
        self.chk_active = QCheckBox("Aktif")
        self.chk_active.setChecked(True)

        form.addRow("Kullanıcı adı", self.ed_username)
        form.addRow("Parola", self.ed_password)
        form.addRow("Yetkiler (virgülle)", self.ed_permissions)
        form.addRow(" ", self.chk_active)
        layout.addLayout(form)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("Ekle")
        buttons.button(QDialogButtonBox.Cancel).setText("Vazgeç")
        buttons.accepted.connect(self._validate_and_accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _validate_and_accept(self):
        username = self.ed_username.text().strip()
        password = self.ed_password.text()
        perms = _normalize_permissions(self.ed_permissions.text())

        if not username:
            QMessageBox.warning(self, "Eksik bilgi", "Kullanıcı adı zorunludur.")
            return
        if username.lower() in self._existing:
            QMessageBox.warning(self, "Tekrarlı kullanıcı", "Bu kullanıcı adı zaten mevcut.")
            return
        if not password:
            QMessageBox.warning(self, "Eksik bilgi", "Parola giriniz.")
            return
        if not perms:
            QMessageBox.warning(self, "Eksik bilgi", "En az 1 yetki giriniz.")
            return

        self._result = {
            "username": username,
            "password": password,
            "permissions": perms,
            "is_active": self.chk_active.isChecked(),
        }
        self.accept()

    @property
    def result(self) -> dict | None:
        return getattr(self, "_result", None)


class _ResetPasswordDialog(QDialog):
    def __init__(self, username: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Parola sıfırla — {username}")
        layout = QVBoxLayout(self)
        form = QFormLayout()
        self.ed_password = QLineEdit()
        self.ed_password.setEchoMode(QLineEdit.Password)
        form.addRow("Yeni parola", self.ed_password)
        layout.addLayout(form)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("Kaydet")
        buttons.button(QDialogButtonBox.Cancel).setText("Vazgeç")
        buttons.accepted.connect(self._validate_and_accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _validate_and_accept(self):
        pwd = self.ed_password.text()
        if not pwd:
            QMessageBox.warning(self, "Eksik bilgi", "Yeni parolayı giriniz.")
            return
        self._password = pwd
        self.accept()

    @property
    def password(self) -> str | None:
        return getattr(self, "_password", None)


class _EditPermissionsDialog(QDialog):
    def __init__(self, username: str, current_perms: list[str], parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Yetkiler — {username}")
        layout = QVBoxLayout(self)
        form = QFormLayout()
        self.ed_permissions = QLineEdit(", ".join(current_perms))
        form.addRow("Yetkiler", self.ed_permissions)
        layout.addLayout(form)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("Güncelle")
        buttons.button(QDialogButtonBox.Cancel).setText("Vazgeç")
        buttons.accepted.connect(self._validate_and_accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _validate_and_accept(self):
        perms = _normalize_permissions(self.ed_permissions.text())
        if not perms:
            QMessageBox.warning(self, "Eksik bilgi", "En az 1 yetki giriniz.")
            return
        self._permissions = perms
        self.accept()

    @property
    def permissions(self) -> list[str] | None:
        return getattr(self, "_permissions", None)


class UserManagementWidget(QWidget):
    """Basit kullanıcı yönetim ekranı (admin için)."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._users: list[dict] = []

        layout = QVBoxLayout(self)

        info = QLabel(
            "Bu ekran yalnızca admin kullanıcılar içindir. Kullanıcı ekleme, parola sıfırlama ve yetki düzenleme işlemleri SQL üzerindeki AppUsers tablosuna kaydedilir."
        )
        info.setWordWrap(True)
        layout.addWidget(info)

        btn_row = QHBoxLayout()
        self.btn_refresh = QPushButton("Listeyi Yenile")
        self.btn_add = QPushButton("Yeni Kullanıcı")
        self.btn_reset = QPushButton("Parola Sıfırla")
        self.btn_edit_perms = QPushButton("Yetkileri Güncelle")

        btn_row.addWidget(self.btn_refresh)
        btn_row.addStretch(1)
        btn_row.addWidget(self.btn_add)
        btn_row.addWidget(self.btn_reset)
        btn_row.addWidget(self.btn_edit_perms)
        layout.addLayout(btn_row)

        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels([
            "Kullanıcı",
            "Yetkiler",
            "Aktif",
            "Oluşturma",
        ])
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.SingleSelection)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.verticalHeader().setVisible(False)
        layout.addWidget(self.table)

        self.btn_refresh.clicked.connect(self._load_users)
        self.btn_add.clicked.connect(self._add_user)
        self.btn_reset.clicked.connect(self._reset_password)
        self.btn_edit_perms.clicked.connect(self._edit_permissions)

        self._load_users()

    def _load_users(self):
        self._users = storage.load_users()
        self.table.setRowCount(len(self._users))
        for row_idx, user in enumerate(self._users):
            username = str(user.get("username", ""))
            perms = ", ".join(_normalize_permissions(user.get("permissions", [])))
            active = "Evet" if user.get("is_active", True) else "Hayır"
            created = str(user.get("created_at", ""))

            self.table.setItem(row_idx, 0, QTableWidgetItem(username))
            self.table.setItem(row_idx, 1, QTableWidgetItem(perms))
            self.table.setItem(row_idx, 2, QTableWidgetItem(active))
            self.table.setItem(row_idx, 3, QTableWidgetItem(created))
        self.table.resizeColumnsToContents()

    def _selected_username(self) -> str | None:
        row = self.table.currentRow()
        if row < 0 or row >= len(self._users):
            return None
        return str(self._users[row].get("username", ""))

    def _add_user(self):
        existing = {u.get("username", "").lower() for u in self._users}
        dlg = _AddUserDialog(existing, self)
        if dlg.exec() != QDialog.Accepted or not dlg.result:
            return

        salt = secrets.token_hex(16)
        password_hash = storage.hash_password(dlg.result["password"], salt)
        new_user = {
            "username": dlg.result["username"],
            "salt": salt,
            "password_hash": password_hash,
            "permissions": dlg.result.get("permissions", []),
            "is_active": dlg.result.get("is_active", True),
        }
        self._users.append(new_user)
        storage.save_users(self._users)
        QMessageBox.information(self, "Kullanıcı eklendi", f"{new_user['username']} başarıyla eklendi.")
        self._load_users()

    def _reset_password(self):
        username = self._selected_username()
        if not username:
            QMessageBox.warning(self, "Seçim yok", "Lütfen bir kullanıcı seçiniz.")
            return

        dlg = _ResetPasswordDialog(username, self)
        if dlg.exec() != QDialog.Accepted or not dlg.password:
            return

        salt = secrets.token_hex(16)
        pwd_hash = storage.hash_password(dlg.password, salt)
        updated = False
        for u in self._users:
            if str(u.get("username", "")) == username:
                u["salt"] = salt
                u["password_hash"] = pwd_hash
                updated = True
                break
        if not updated:
            QMessageBox.warning(self, "Bulunamadı", "Kullanıcı listede bulunamadı.")
            return

        storage.save_users(self._users)
        QMessageBox.information(self, "Parola güncellendi", "Yeni parola kaydedildi.")
        self._load_users()

    def _edit_permissions(self):
        username = self._selected_username()
        if not username:
            QMessageBox.warning(self, "Seçim yok", "Lütfen bir kullanıcı seçiniz.")
            return

        current = []
        for u in self._users:
            if str(u.get("username", "")) == username:
                current = _normalize_permissions(u.get("permissions", []))
                break

        dlg = _EditPermissionsDialog(username, current, self)
        if dlg.exec() != QDialog.Accepted or not dlg.permissions:
            return

        for u in self._users:
            if str(u.get("username", "")) == username:
                u["permissions"] = dlg.permissions
                break

        storage.save_users(self._users)
        QMessageBox.information(self, "Yetkiler güncellendi", "Kullanıcı yetkileri kaydedildi.")
        self._load_users()