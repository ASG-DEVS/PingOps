"""
PingOps v1.1
Developer: ASG-DEVS (Abdullah Saad Alghamdi)

A lightweight multi-IP monitoring tool for Network & NOC engineers.
Features:
- Multiple simultaneous ping windows
- Real-time UP/DOWN/FLAPPING detection
- Continuous ping
- Names/labels per IP
- Excel export
- Clean PyQt5 dark UI
"""

import sys
import subprocess
import threading
import platform
import pandas as pd
from pathlib import Path

from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QLabel, QPlainTextEdit, QScrollArea, QFrame, QMessageBox,
    QDialog, QFileDialog
)
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import Qt, pyqtSignal, QObject


# ---------------------------------------------------------
# Resource Path (supports PyInstaller EXE)
# ---------------------------------------------------------
def resource_path(relative_path: str) -> str:
    """Return absolute path to resource for both script & PyInstaller."""
    base_path = getattr(sys, "_MEIPASS", Path(__file__).parent)
    return str(Path(base_path) / relative_path)


# ---------------------------------------------------------
# Ping Thread Signal
# ---------------------------------------------------------
class PingSignals(QObject):
    status_signal = pyqtSignal(str, str)  # ip, reply


# ---------------------------------------------------------
# Single IP Widget
# ---------------------------------------------------------
class PingWidget(QWidget):
    def __init__(self, main_window, ip: str, name: str = ""):
        super().__init__()
        self.main_window = main_window
        self.ip = ip
        self.name = name

        self.thread_running = False
        self.thread = None
        self.signals = PingSignals()
        self.last_status = None  # For FLAPPING detection

        # ---------------- UI Layout ----------------
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(8, 8, 8, 8)
        main_layout.setSpacing(4)

        # Row 1: IP + Name
        ip_row = QHBoxLayout()
        self.ip_label = QLabel(self.ip)
        self.ip_label.setStyleSheet("color: #FFFFFF; font-size: 13px; font-weight: bold;")
        ip_row.addWidget(self.ip_label)

        self.name_label = QLabel(self.name)
        self.name_label.setStyleSheet("color: #CCCCCC; font-size: 12px;")
        ip_row.addStretch()
        ip_row.addWidget(self.name_label)
        main_layout.addLayout(ip_row)

        # Row 2: Status box + Reply
        status_row = QHBoxLayout()

        self.status_box = QFrame()
        self.status_box.setFixedSize(130, 32)
        self.status_box.setStyleSheet("""
            QFrame {
                background-color: #333333;
                border-radius: 10px;
                border: 1px solid #555555;
            }
        """)

        status_inner = QHBoxLayout(self.status_box)
        status_inner.setContentsMargins(8, 0, 8, 0)

        self.icon_label = QLabel("•")
        self.icon_label.setStyleSheet("color: #AAAAAA; font-size: 18px;")

        self.status_text = QLabel("Unknown")
        self.status_text.setStyleSheet("color: #DDDDDD; font-size: 11px;")

        status_inner.addWidget(self.icon_label)
        status_inner.addWidget(self.status_text)

        status_row.addWidget(self.status_box)

        self.reply_label = QLabel("Waiting...")
        self.reply_label.setWordWrap(True)
        self.reply_label.setStyleSheet("color: #BBBBBB; font-size: 11px;")
        status_row.addWidget(self.reply_label, 1)

        main_layout.addLayout(status_row)

        # Row 3: Buttons
        btn_row = QHBoxLayout()
        self.start_btn = QPushButton("Start")
        self.stop_btn = QPushButton("Stop")
        self.delete_btn = QPushButton("Delete")

        for b in (self.start_btn, self.stop_btn, self.delete_btn):
            b.setCursor(Qt.PointingHandCursor)

        btn_row.addWidget(self.start_btn)
        btn_row.addWidget(self.stop_btn)
        btn_row.addWidget(self.delete_btn)
        btn_row.addStretch()
        main_layout.addLayout(btn_row)

        self.setLayout(main_layout)

        # Connect signals
        self.signals.status_signal.connect(self.update_status)
        self.start_btn.clicked.connect(self.start_ping)
        self.stop_btn.clicked.connect(self.stop_ping)
        self.delete_btn.clicked.connect(self.request_delete)

        # Button Style
        self.setStyleSheet("""
            QPushButton {
                background-color: #3A3A3A;
                color: white;
                padding: 4px 10px;
                border-radius: 6px;
                border: 1px solid #555555;
            }
            QPushButton:hover {
                background-color: #505050;
            }
        """)

    # ---------------------------------------------------------
    # Delete widget
    # ---------------------------------------------------------
    def request_delete(self):
        if QMessageBox.question(
            self, "Confirm delete",
            f"Delete ping window for {self.ip}?",
            QMessageBox.Yes | QMessageBox.No
        ) == QMessageBox.Yes:
            self.main_window.remove_widget(self)

    # ---------------------------------------------------------
    # Update Visual Status
    # ---------------------------------------------------------
    def update_status(self, ip: str, reply: str):
        if "Reply from" in reply or "bytes=" in reply:
            status = "UP"
            grad = ("#225522", "#2E7D32")
            icon_char, icon_color = "✔", "#A5D6A7"

        elif "Request timed out" in reply:
            status = "DOWN"
            grad = ("#552222", "#C62828")
            icon_char, icon_color = "✖", "#EF9A9A"

        else:
            status = "DOWN"
            grad = ("#552222", "#C62828")
            icon_char, icon_color = "✖", "#EF9A9A"

        # FLAPPING detection
        if self.last_status and self.last_status != status:
            status = "FLAPPING"
            grad = ("#665500", "#FBC02D")
            icon_char, icon_color = "⚠", "#FFE082"

        self.last_status = status

        # Apply Styles
        self.status_box.setStyleSheet(f"""
            QFrame {{
                background-color: qlineargradient(
                    x1: 0, y1: 0, x2: 1, y2: 0,
                    stop: 0 {grad[0]}, stop: 1 {grad[1]}
                );
                border-radius: 10px;
                border: 1px solid #666666;
            }}
        """)

        self.icon_label.setText(icon_char)
        self.icon_label.setStyleSheet(f"color: {icon_color}; font-size: 16px;")
        self.status_text.setText(status)
        self.reply_label.setText(reply)

        # Save result in main window
        self.main_window.ping_results[self.ip] = (status, self.name)

    # ---------------------------------------------------------
    # Ping Thread
    # ---------------------------------------------------------
    def start_ping(self):
        if not self.thread_running:
            self.thread_running = True
            self.thread = threading.Thread(target=self.run_ping, daemon=True)
            self.thread.start()

    def stop_ping(self):
        self.thread_running = False

    def run_ping(self):
        is_win = platform.system().lower() == "windows"
        param = "-n" if is_win else "-c"

        startupinfo = None
        if is_win:
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW

        while self.thread_running:
            result = subprocess.run(
                ["ping", param, "1", self.ip],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                startupinfo=startupinfo
            )

            if result.returncode == 0:
                line = next(
                    (l for l in result.stdout.splitlines()
                     if "Reply from" in l or "bytes=" in l),
                    "Reply received"
                )
            else:
                line = "Request timed out."

            self.signals.status_signal.emit(self.ip, line)


# ---------------------------------------------------------
# Help / About Dialog
# ---------------------------------------------------------
class HelpDialog(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("About PingOps")
        self.resize(460, 360)
        self.setStyleSheet("background-color: #1E1E1E; color: white;")

        layout = QVBoxLayout()

        header = QLabel("PingOps v1.1")
        header.setStyleSheet("font-size: 20px; font-weight: bold; color: #4CAF50;")
        layout.addWidget(header)

        text = QLabel()
        text.setWordWrap(True)
        text.setTextFormat(Qt.RichText)
        text.setOpenExternalLinks(True)
        text.setStyleSheet("font-size: 13px; color: #DDDDDD;")
        text.setText("""
<b>English:</b><br>
PingOps is a lightweight professional IP monitoring tool designed for NOC engineers.<br>
• Monitor multiple IPs<br>
• Real-time UP / DOWN / FLAPPING detection<br>
• Export results to Excel<br><br>
<b>Developer:</b> Abdullah Al-Ghamdi<br>
<a href="https://www.linkedin.com/in/abdullah-saad-alghamdi-553b67167/" style="color:#4CAF50;">LinkedIn Profile</a><br><br>
<hr>
<b>العربية:</b><br>
PingOps أداة احترافية لمراقبة الشبكات ومهندسي الـ NOC.<br>
• مراقبة عدة IPs<br>
• كشف UP / DOWN / FLAPPING<br>
• تصدير النتائج لملف Excel<br><br>
شكراً لاستخدامك PingOps ❤️
        """)
        layout.addWidget(text)

        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.accept)
        close_btn.setCursor(Qt.PointingHandCursor)
        close_btn.setStyleSheet("""
            QPushButton {
                background-color: #3A3A3A;
                color: white;
                padding: 6px 12px;
                border-radius: 6px;
                border: 1px solid #555555;
            }
            QPushButton:hover {
                background-color: #505050;
            }
        """)
        layout.addWidget(close_btn, alignment=Qt.AlignRight)

        self.setLayout(layout)


# ---------------------------------------------------------
# Main Window
# ---------------------------------------------------------
class MainWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("PingOps Dashboard v1.1")
        self.resize(900, 700)
        self.setWindowIcon(QIcon(resource_path("pingops_final.ico")))
        self.setStyleSheet("background-color: #121212; color: white;")

        self.widgets = []
        self.existing_ips = set()
        self.ping_results = {}

        root = QVBoxLayout()
        root.setContentsMargins(0, 0, 0, 0)

        # Top Bar
        top = QFrame()
        top.setStyleSheet("""
            QFrame {
                background-color: #202020;
                border-bottom: 1px solid #333333;
            }
        """)
        top_layout = QHBoxLayout(top)
        top_layout.setContentsMargins(12, 6, 12, 6)

        icon = QLabel()
        icon.setPixmap(QIcon(resource_path("pingops_final.ico")).pixmap(22, 22))

        title = QLabel("PingOps Dashboard v1.1")
        title.setStyleSheet("font-size: 14px; font-weight: bold;")

        top_layout.addWidget(icon)
        top_layout.addWidget(title)
        top_layout.addStretch()

        root.addWidget(top)

        # Content Layout
        content = QVBoxLayout()
        content.setContentsMargins(10, 10, 10, 10)

        # Two input text boxes
        row = QHBoxLayout()

        self.ip_box = QPlainTextEdit()
        self.ip_box.setPlaceholderText("Paste IP list (one IP per line)")
        self.ip_box.setStyleSheet("""
            QPlainTextEdit {
                background-color: #1F1F1F;
                color: white;
                border: 1px solid #333333;
                border-radius: 6px;
            }
        """)

        self.name_box = QPlainTextEdit()
        self.name_box.setPlaceholderText("Optional: Names matching IP order")
        self.name_box.setStyleSheet("""
            QPlainTextEdit {
                background-color: #1F1F1F;
                color: white;
                border: 1px solid #333333;
                border-radius: 6px;
            }
        """)

        row.addWidget(self.ip_box, 1)
        row.addWidget(self.name_box, 1)
        content.addLayout(row)

        # Toolbar
        tools = QFrame()
        tools.setStyleSheet("""
            QFrame {
                background-color: #2B2B1F;
                border: 1px solid #4A4A32;
                border-radius: 8px;
            }
        """)
        t_layout = QHBoxLayout(tools)
        t_layout.setContentsMargins(10, 6, 10, 6)

        self.btn_generate = QPushButton("Generate")
        self.btn_start = QPushButton("Start All")
        self.btn_stop = QPushButton("Stop All")
        self.btn_delete = QPushButton("Delete All")
        self.btn_export = QPushButton("Export")
        self.btn_help = QPushButton("?")
        self.btn_help.setFixedWidth(32)

        for b in (self.btn_generate, self.btn_start, self.btn_stop, self.btn_delete, self.btn_export, self.btn_help):
            b.setCursor(Qt.PointingHandCursor)

        t_layout.addWidget(self.btn_generate)
        t_layout.addWidget(self.btn_start)
        t_layout.addWidget(self.btn_stop)
        t_layout.addWidget(self.btn_delete)
        t_layout.addWidget(self.btn_export)
        t_layout.addStretch()
        t_layout.addWidget(self.btn_help)

        tools.setStyleSheet(tools.styleSheet() + """
            QPushButton {
                background-color: #3A3A2A;
                color: white;
                padding: 4px 12px;
                border-radius: 6px;
                border: 1px solid #55553A;
            }
            QPushButton:hover {
                background-color: #545438;
            }
        """)

        content.addWidget(tools)

        # Scroll area for widgets
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll_content = QWidget()
        self.container = QVBoxLayout(scroll_content)
        self.container.setContentsMargins(4, 4, 4, 4)
        scroll.setWidget(scroll_content)

        content.addWidget(scroll)

        root.addLayout(content)
        self.setLayout(root)

        # Signals
        self.btn_generate.clicked.connect(self.generate_windows)
        self.btn_start.clicked.connect(self.start_all)
        self.btn_stop.clicked.connect(self.stop_all)
        self.btn_delete.clicked.connect(self.delete_all)
        self.btn_export.clicked.connect(self.export_results)
        self.btn_help.clicked.connect(self.show_help)

    # ---------------------------------------------------------
    # Window Actions
    # ---------------------------------------------------------
    def generate_windows(self):
        ips = self.ip_box.toPlainText().splitlines()
        names = self.name_box.toPlainText().splitlines()

        for idx, raw in enumerate(ips):
            ip = raw.strip()
            if not ip:
                continue
            if ip in self.existing_ips:
                continue

            name = names[idx].strip() if idx < len(names) else ""

            widget = PingWidget(self, ip, name)
            self.widgets.append(widget)
            self.existing_ips.add(ip)
            self.ping_results[ip] = ("Unknown", name)

            self.container.addWidget(widget)

    def remove_widget(self, widget: PingWidget):
        widget.thread_running = False
        if widget.ip in self.existing_ips:
            self.existing_ips.remove(widget.ip)
        if widget.ip in self.ping_results:
            del self.ping_results[widget.ip]

        if widget in self.widgets:
            self.widgets.remove(widget)

        widget.setParent(None)
        widget.deleteLater()

    def start_all(self):
        for w in self.widgets:
            w.start_ping()

    def stop_all(self):
        for w in self.widgets:
            w.stop_ping()

    def delete_all(self):
        if not self.widgets:
            return

        if QMessageBox.question(
            self, "Confirm delete",
            "Delete all ping windows?",
            QMessageBox.Yes | QMessageBox.No
        ) == QMessageBox.Yes:
            for w in list(self.widgets):
                self.remove_widget(w)

    # ---------------------------------------------------------
    # Export
    # ---------------------------------------------------------
    def export_results(self):
        if not self.ping_results:
            QMessageBox.information(self, "Export", "No results to export.")
            return

        path, _ = QFileDialog.getSaveFileName(
            self, "Save Results", "Ping_Results.xlsx", "Excel Files (*.xlsx)"
        )
        if not path:
            return

        if not path.lower().endswith(".xlsx"):
            path += ".xlsx"

        df = pd.DataFrame({
            "Name": [v[1] for v in self.ping_results.values()],
            "IP Address": list(self.ping_results.keys()),
            "Status": [v[0] for v in self.ping_results.values()]
        })

        try:
            df.to_excel(path, index=False)
            QMessageBox.information(self, "Export", f"Saved:\n{path}")
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))

    # ---------------------------------------------------------
    # Help
    # ---------------------------------------------------------
    def show_help(self):
        dlg = HelpDialog()
        dlg.exec_()


# ---------------------------------------------------------
# Run App
# ---------------------------------------------------------
if __name__ == "__main__":
    app = QApplication(sys.argv)

    icon_path = Path(__file__).parent / "pingops_final.ico"
    app.setWindowIcon(QIcon(str(icon_path)))

    win = MainWindow()
    win.show()
    sys.exit(app.exec_())
