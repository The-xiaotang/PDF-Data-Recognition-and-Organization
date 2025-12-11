import os
from PySide6.QtWidgets import (QWidget, QVBoxLayout, QPushButton, QLabel,
                               QFileDialog, QMessageBox, QProgressBar,
                               QFrame, QHBoxLayout, QTextEdit, QComboBox)
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QStyle
from workers import PdfToWordWorker, WordToPdfWorker, PdfToExcelWorker, PdfToRearrangementWorker, PdfToMutationWorker

class FileDropArea(QFrame):
    def __init__(self, mode, on_file_selected):
        super().__init__()
        self.mode = mode
        self.on_file_selected = on_file_selected
        self.setObjectName("dropArea")
        self.setAcceptDrops(True)
        layout = QVBoxLayout()
        self.label = QLabel("拖拽文件到此处或点击下方浏览")
        self.label.setObjectName("dropLabel")
        self.label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.label)
        self.setLayout(layout)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            url = event.mimeData().urls()[0]
            path = url.toLocalFile()
            if self._is_valid_file(path):
                event.acceptProposedAction()
                return
        event.ignore()

    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            path = event.mimeData().urls()[0].toLocalFile()
            if self._is_valid_file(path):
                self.on_file_selected(path)

    def _is_valid_file(self, path):
        ext = os.path.splitext(path)[1].lower()
        if self.mode in ["pdf2word", "pdf2excel", "pdf2rearrangement", "pdf2mutation"]:
            return ext == ".pdf"
        return ext in [".doc", ".docx"]

class ConversionTab(QWidget):
    def __init__(self, mode="pdf2word"):
        super().__init__()
        self.mode = mode
        self.layout = QVBoxLayout()
        self.file_path = ""
        
        # File Selection
        header = QLabel("模式")
        header.setObjectName("headerTitle")
        header.setText(self._mode_title())

        card = QFrame()
        card.setObjectName("card")
        card_layout = QVBoxLayout()

        self.file_label = QLabel("未选择文件")
        self.file_label.setObjectName("filePathLabel")
        self.file_label.setWordWrap(True)

        self.drop_area = FileDropArea(self.mode, self._on_file_selected)

        self.select_btn = QPushButton("浏览…")
        self.select_btn.setIcon(self.style().standardIcon(QStyle.SP_DialogOpenButton))
        self.select_btn.clicked.connect(self.select_file)
        
        # Convert Button
        btn_text = "Convert"

        '''
        if mode == 'pdf2excel':
            btn_text = "基础信息" # Basic Info
        elif mode == 'pdf2rearrangement':
            btn_text = "重排结果" # Rearrangement Result
        elif mode == 'pdf2mutation':
            btn_text = "突变数据" # Mutation Data
        '''

        self.convert_btn = QPushButton(btn_text)
        self.convert_btn.setObjectName("primaryButton")
        self.convert_btn.clicked.connect(self.convert_file)
        self.convert_btn.setEnabled(False)
        
        # Progress Bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 0)
        self.progress_bar.hide()
        
        footer = QHBoxLayout()
        footer.addWidget(self.select_btn)
        footer.addStretch(1)
        footer.addWidget(self.convert_btn)

        self.log_area = QTextEdit()
        self.log_area.setObjectName("logArea")
        self.log_area.setReadOnly(True)

        card_layout.addWidget(self.file_label)
        card_layout.addWidget(self.drop_area)
        card_layout.addLayout(footer)
        card_layout.addWidget(self.progress_bar)
        card.setLayout(card_layout)

        self.layout.addWidget(header)
        self.layout.addWidget(card)
        self.layout.addWidget(self.log_area)
        self.layout.addStretch()
        self.setLayout(self.layout)
        
        self.worker = None

    def select_file(self):
        if self.mode in ["pdf2word", "pdf2excel", "pdf2rearrangement", "pdf2mutation"]:
            file_filter = "PDF Files (*.pdf)"
        else:
            file_filter = "Word Files (*.docx *.doc)"
            
        fname, _ = QFileDialog.getOpenFileName(self, "Select File", "", file_filter)
        if fname:
            self._on_file_selected(fname)

    def _on_file_selected(self, path):
        self.file_path = path
        self.file_label.setText(path)
        self.convert_btn.setEnabled(True)
        self.log_area.append(f"已选择文件: {os.path.basename(path)}")

    def convert_file(self):
        if not self.file_path:
            return
            
        # Determine output path
        if self.mode == "pdf2word":
            default_out = os.path.splitext(self.file_path)[0] + ".docx"
            file_filter = "Word Files (*.docx)"
        elif self.mode == "word2pdf":
            default_out = os.path.splitext(self.file_path)[0] + ".pdf"
            file_filter = "PDF Files (*.pdf)"
        elif self.mode == "pdf2excel":
            # Set default filename to "基础信息.xlsx" in the same directory as the source file
            source_dir = os.path.dirname(self.file_path)
            default_out = os.path.join(source_dir, "基础信息.xlsx")
            file_filter = "Excel Files (*.xlsx)"
        elif self.mode == "pdf2rearrangement":
            # Set default filename to "重排结果.xlsx"
            source_dir = os.path.dirname(self.file_path)
            default_out = os.path.join(source_dir, "重排结果.xlsx")
            file_filter = "Excel Files (*.xlsx)"
        elif self.mode == "pdf2mutation":
            # Set default filename to "突变数据.xlsx"
            source_dir = os.path.dirname(self.file_path)
            default_out = os.path.join(source_dir, "突变数据.xlsx")
            file_filter = "Excel Files (*.xlsx)"
            
        out_fname, _ = QFileDialog.getSaveFileName(self, "Save Result", default_out, file_filter)
        
        if not out_fname:
            return

        self.progress_bar.show()
        self.convert_btn.setEnabled(False)
        self.select_btn.setEnabled(False)
        
        if self.mode == "pdf2word":
            self.worker = PdfToWordWorker(self.file_path, out_fname)
        elif self.mode == "word2pdf":
            self.worker = WordToPdfWorker(self.file_path, out_fname)
        elif self.mode == "pdf2excel":
            self.worker = PdfToExcelWorker(self.file_path, out_fname)
        elif self.mode == "pdf2rearrangement":
            self.worker = PdfToRearrangementWorker(self.file_path, out_fname)
        elif self.mode == "pdf2mutation":
            self.worker = PdfToMutationWorker(self.file_path, out_fname)
            
        self.worker.finished.connect(self.on_finished)
        self.worker.error.connect(self.on_error)
        self.worker.start()

    def on_finished(self, msg):
        self.progress_bar.hide()
        self.convert_btn.setEnabled(True)
        self.select_btn.setEnabled(True)
        QMessageBox.information(self, "Success", msg)
        self.log_area.append(msg)

    def on_error(self, err):
        self.progress_bar.hide()
        self.convert_btn.setEnabled(True)
        self.select_btn.setEnabled(True)
        QMessageBox.critical(self, "Error", f"Conversion failed:\n{err}")
        self.log_area.append(str(err))

    def _mode_title(self):
        if self.mode == 'pdf2word':
            return "PDF 转 Word"
        if self.mode == 'word2pdf':
            return "Word 转 PDF"
        if self.mode == 'pdf2excel':
            return "PDF 转 Excel（基础信息）"
        if self.mode == 'pdf2rearrangement':
            return "PDF 转 Excel（重排结果）"
        if self.mode == 'pdf2mutation':
            return "PDF 转 Excel（突变数据）"
        return "模式"

class CombinedConversionTab(QWidget):
    def __init__(self, group):
        super().__init__()
        self.group = group
        self.layout = QVBoxLayout()
        self.selector = QComboBox()
        self.selector.setEditable(False)
        self.inner = None

        if group == 'doc':
            self.selector.addItems(["PDF 转 Word", "Word 转 PDF"]) 
        else:
            self.selector.addItems(["基础信息", "重排结果", "突变数据"]) 

        self.selector.currentIndexChanged.connect(self._on_mode_change)
        self.layout.addWidget(self.selector)
        self.setLayout(self.layout)
        self._on_mode_change(0)

    def _on_mode_change(self, idx):
        if self.inner:
            self.layout.removeWidget(self.inner)
            self.inner.deleteLater()
            self.inner = None
        if self.group == 'doc':
            mode = 'pdf2word' if idx == 0 else 'word2pdf'
        else:
            mode = ['pdf2excel', 'pdf2rearrangement', 'pdf2mutation'][idx]
        self.inner = ConversionTab(mode)
        self.layout.addWidget(self.inner)
