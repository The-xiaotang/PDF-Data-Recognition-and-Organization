APP_QSS = """
QMainWindow {
    background-color: #0f172a;
    color: #e5e7eb;
    font-family: "Microsoft YaHei UI", "Segoe UI", sans-serif;
    font-size: 13px;
}
QTabWidget::pane {
    border: 1px solid #1f2937;
    border-radius: 8px;
    background: #111827;
}
QTabBar::tab {
    background: #1f2937;
    color: #cbd5e1;
    padding: 8px 16px;
    border-top-left-radius: 6px;
    border-top-right-radius: 6px;
    margin: 2px;
}
QTabBar::tab:selected {
    background: #111827;
    color: #ffffff;
}
QFrame#card {
    background: #0b1220;
    border: 1px solid #1f2937;
    border-radius: 10px;
}
QLabel#headerTitle {
    color: #ffffff;
    font-size: 18px;
    font-weight: 600;
}
QLabel#filePathLabel {
    color: #93c5fd;
}
QPushButton {
    background: #1f2937;
    border: 1px solid #374151;
    color: #e5e7eb;
    padding: 8px 14px;
    border-radius: 6px;
}
QPushButton:hover { background: #374151; }
QPushButton#primaryButton {
    background: #2563eb;
    border: 1px solid #1d4ed8;
    color: #ffffff;
}
QPushButton#primaryButton:hover { background: #1d4ed8; }
QProgressBar {
    background: #0b1220;
    border: 1px solid #1f2937;
    border-radius: 6px;
    text-visible: false;
}
QProgressBar::chunk { background-color: #22c55e; border-radius: 6px; }
QFrame#dropArea {
    background: #0b1220;
    border: 2px dashed #334155;
    border-radius: 10px;
}
QLabel#dropLabel { color: #cbd5e1; }
QTextEdit#logArea {
    background: #0b1220;
    border: 1px solid #1f2937;
    color: #93c5fd;
    border-radius: 8px;
}
"""
