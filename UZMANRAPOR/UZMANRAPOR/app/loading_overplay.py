from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt, QSize, QEvent, QTimer
from PySide6.QtGui import QPixmap
from PySide6.QtWidgets import QWidget, QLabel, QProgressBar, QVBoxLayout


class LoadingOverlay(QWidget):
    """
    MainWindow üzerinde tam kaplama (overlay).
    assets/acilis.png'yi pencere boyutuna sığdırır, altta indeterminate bar gösterir.
    """

    def __init__(self, parent: QWidget, image_path: Path) -> None:
        super().__init__(parent)
        self._image_path = image_path

        self.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        self.setStyleSheet("background: black;")

        self._img = QLabel(self)
        self._img.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._img.setStyleSheet("background: transparent;")

        self._bar = QProgressBar(self)
        self._bar.setRange(0, 0)  # hareketli "loading"
        self._bar.setTextVisible(False)
        self._bar.setFixedHeight(10)
        self._bar.setStyleSheet("""
            QProgressBar {
                background: rgba(255,255,255,30);
                border: 1px solid rgba(255,255,255,40);
                border-radius: 5px;
            }
            QProgressBar::chunk {
                background-color: #00c2ff;
                border-radius: 5px;
            }
        """)

        lay = QVBoxLayout(self)
        lay.setContentsMargins(24, 24, 24, 24)
        lay.setSpacing(16)
        lay.addWidget(self._img, stretch=1)
        lay.addWidget(self._bar, stretch=0)

        self.hide()
        parent.installEventFilter(self)

        # İlk show’dan sonra bir tur daha güncelle (layout otursun diye)
        QTimer.singleShot(0, self._update_pixmap)

    def show_overlay(self) -> None:
        self.setGeometry(self.parentWidget().rect())
        self.raise_()
        self.show()
        self._update_pixmap()

    def hide_overlay(self) -> None:
        self.hide()

    def eventFilter(self, obj, event):
        if obj is self.parentWidget() and event.type() == QEvent.Type.Resize:
            self.setGeometry(self.parentWidget().rect())
            self._update_pixmap()
        return False

    def _update_pixmap(self) -> None:
        if not self._image_path.exists():
            self._img.clear()
            return

        pix = QPixmap(str(self._image_path))
        if pix.isNull():
            self._img.clear()
            return

        target: QSize = self._img.size() if self._img.width() > 0 else self.size()
        scaled = pix.scaled(
            target,
            Qt.AspectRatioMode.KeepAspectRatio,
            Qt.TransformationMode.SmoothTransformation,
        )
        self._img.setPixmap(scaled)
