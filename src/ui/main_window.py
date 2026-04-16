"""Main GUI layout using PySide6. Implements the full Project Wizard multi-view."""

import sys
import os
import copy
from pathlib import Path

from PySide6.QtCore import Qt, Slot, QSize, QStandardPaths, QSettings, QTimer
from PySide6.QtGui import QIcon, QFont, QImage, QPixmap, QShortcut, QKeySequence, QColor
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
    QPushButton, QLabel, QFileDialog, QTableWidget, 
    QTableWidgetItem, QProgressBar, QHeaderView, QMessageBox,
    QStackedWidget, QLineEdit, QComboBox, QSpinBox, QDialog,
    QTabWidget, QScrollArea, QCheckBox, QGridLayout, QSlider,
    QListWidget
)

import numpy as np

from src.core.project import ProjectManager
from src.ui.workers import PDFGraderWorker, AutoTuneWorker
from src.models.schemas import GradingResult
from src.data.excel import ExcelManager
from src.utils.paths import get_resource_path, get_config_path

# --- Shared UI Styling ---
APP_STYLESHEET = """
    /* Premium Night Slate Design System */
    QMainWindow, QDialog, QStackedWidget { 
        background-color: #020617; 
    }
    QScrollArea {
        background-color: transparent;
        border: none;
    }
    QWidget { 
        color: #F8FAFC; 
        font-family: 'Inter', 'Segoe UI', 'Outfit', sans-serif; 
        font-size: 10pt; 
    }
    
    /* Typography & Labels */
    QLabel { background: transparent; color: #94A3B8; }
    QLabel[cssClass="title"] { 
        font-size: 24pt; 
        font-weight: 800; 
        color: #F8FAFC; 
        margin-bottom: 5px;
    }
    QLabel[cssClass="subtitle"] { 
        font-size: 11pt; 
        color: #64748B; 
    }
    QLabel#header_title { 
        font-size: 16pt; 
        font-weight: 700; 
        color: #38BDF8; 
    }
    
    /* Card/Surface Style */
    QWidget#card {
        background-color: #0F172A;
        border: 1px solid #1E293B;
        border-radius: 16px;
    }
    
    /* Premium Buttons */
    QPushButton {
        background-color: #1E293B;
        border: 1px solid #334155;
        border-radius: 10px;
        padding: 10px 20px;
        font-weight: 600;
        color: #F1F5F9;
    }
    QPushButton:hover { 
        background-color: #334155; 
        border-color: #475569;
    }
    QPushButton:pressed { 
        background-color: #0F172A; 
    }
    QPushButton:disabled { 
        color: #475569; 
        background-color: #020617; 
        border: 1px solid #0F172A; 
    }
    
    QPushButton[cssClass="primary"] { 
        background-color: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #0EA5E9, stop:1 #2563EB);
        color: #FFFFFF; 
        border: none;
    }
    QPushButton[cssClass="primary"]:hover { 
        background-color: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #38BDF8, stop:1 #3B82F6);
    }
    
    QPushButton[cssClass="success"] { 
        background-color: #059669; 
        color: #FFFFFF; 
        border: none;
    }
    QPushButton[cssClass="success"]:hover { background-color: #10B981; }
    
    QPushButton[cssClass="warning"] { 
        background-color: #D97706; 
        color: #FFFFFF; 
        border: none;
    }
    QPushButton[cssClass="warning"]:hover { background-color: #F59E0B; }
    
    /* Inputs */
    QLineEdit, QComboBox, QSpinBox {
        background-color: #020617;
        border: 1px solid #334155;
        border-radius: 10px;
        padding: 12px;
        color: #F8FAFC;
    }
    QLineEdit:focus, QComboBox:focus, QSpinBox:focus { 
        border-color: #38BDF8; 
        background-color: #0F172A;
    }
    
    /* Fix for Dropdown visibility */
    QComboBox QAbstractItemView {
        background-color: #0F172A;
        border: 1px solid #334155;
        selection-background-color: #1E293B;
        selection-color: #38BDF8;
        outline: none;
    }
    
    /* Modern Table Styling */
    QTableWidget {
        background-color: #0F172A;
        gridline-color: #1E293B;
        border: 1px solid #1E293B;
        border-radius: 12px;
        alternate-background-color: #131C2F;
        selection-background-color: rgba(56, 189, 248, 0.2);
        selection-color: #38BDF8;
        outline: none;
    }
    QHeaderView::section {
        background-color: #1E293B;
        color: #94A3B8;
        font-weight: 700;
        padding: 15px;
        border: none;
        border-bottom: 1px solid #334155;
    }
    
    /* Professional Scrollbars */
    QScrollBar:vertical {
        border: none;
        background: #020617;
        width: 10px;
        margin: 0px;
    }
    QScrollBar::handle:vertical {
        background: #334155;
        min-height: 20px;
        border-radius: 5px;
    }
    QScrollBar::handle:vertical:hover {
        background: #475569;
    }
    QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
        height: 0px;
    }
    
    /* Progress Bar */
    QProgressBar {
        background-color: #1E293B;
        border: none;
        border-radius: 6px;
        text-align: center;
        color: transparent;
        height: 8px;
    }
    QProgressBar::chunk {
        background-color: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #38BDF8, stop:1 #818CF8);
        border-radius: 6px;
    }
    
    /* Modal Lists */
    QListWidget {
        background-color: #0F172A;
        border: 1px solid #1E293B;
        border-radius: 10px;
        padding: 5px;
    }
    QListWidget::item {
        padding: 12px;
        border-radius: 6px;
        margin: 2px;
    }
    QListWidget::item:selected {
        background-color: #1E293B;
        color: #38BDF8;
        font-weight: bold;
    }
    
    /* Window Controls */
    QPushButton#min_btn, QPushButton#close_btn {
        background-color: transparent;
        border: none;
        border-radius: 0px;
        padding: 0px;
        font-size: 16pt;
        font-family: 'Arial', sans-serif;
        font-weight: bold;
        color: #FFFFFF;
        text-align: center;
    }
    QPushButton#min_btn:hover {
        background-color: #334155;
    }
    QPushButton#close_btn:hover {
        background-color: #EF4444;
    }
"""

def cv2_to_qpixmap(cv_img) -> QPixmap:
    """Helper to convert OpenCV crops into PySide natively."""
    if cv_img is None: return QPixmap()
    
    # Numpy slices aren't contiguous in memory, which QImage requires.
    if not cv_img.flags['C_CONTIGUOUS']:
        cv_img = np.ascontiguousarray(cv_img)
        
    h, w, ch = cv_img.shape
    bytes_per_line = ch * w
    # cv2 loads BGR
    img = QImage(cv_img.data, w, h, bytes_per_line, QImage.Format_BGR888)
    return QPixmap.fromImage(img)


class ReviewModal(QDialog):
    def __init__(self, parent, res: GradingResult, pdf_path: str, question_count: int, config_path: str):
        super().__init__(parent)
        self.res = res
        self.pdf_path = pdf_path
        self.question_count = question_count
        self.config_path = config_path
        
        self.setWindowTitle(f"Manual Review — Page {res.page_number}")
        
        # Match the parent window size for a seamless experience
        if parent:
            self.resize(parent.size())
        else:
            self.setMinimumSize(1000, 700)
        
        self._ensure_images_loaded()
        
        main_layout = QHBoxLayout(self)
        
        # --- LEFT PANE ---
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)
        
        lbl_list = QLabel("Detected Errors:")
        lbl_list.setStyleSheet("font-weight: bold; color: #FFFFFF;")
        left_layout.addWidget(lbl_list)
        
        from PySide6.QtWidgets import QListWidget
        self.list_widget = QListWidget()
        self.list_widget.currentRowChanged.connect(self._on_list_selected)
        left_layout.addWidget(self.list_widget)
        
        left_layout.addSpacing(10)
        
        # Action Buttons
        self.btn_save = QPushButton("Accept Changes & Continue")
        self.btn_save.setProperty("cssClass", "primary")
        self.btn_save.setEnabled(False) # Wait for input
        self.btn_save.clicked.connect(self.accept)
        left_layout.addWidget(self.btn_save)
        
        self.btn_ignore = QPushButton("Ignore for Now")
        self.btn_ignore.setProperty("cssClass", "warning")
        self.btn_ignore.clicked.connect(self.reject)
        left_layout.addWidget(self.btn_ignore)
        
        main_layout.addWidget(left_widget, stretch=1)
        
        # --- RIGHT PANE (Scrollable) ---
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setFrameShape(QScrollArea.NoFrame)
        self.scroll.setStyleSheet("background-color: #1A1A1A; border-radius: 8px;")
        
        self.stack = QStackedWidget()
        self.scroll.setWidget(self.stack)
        
        main_layout.addWidget(self.scroll, stretch=2)
        
        # Track inputs for applying later
        self.id_input: QLineEdit | None = None
        self.ver_input: QComboBox | None = None
        self.q_inputs: dict[int, QComboBox] = {}
        
        self.error_keys = [] # list of (type, data) e.g., ("id", None) or ("question", q_num)
        
        self._populate_errors()

    def _ensure_images_loaded(self):
        """Loads images from disk cache if paths exist, otherwise falls back to PDF extraction."""
        # 1. Try Disk Cache First
        if self.res.id_crop_path:
            # We assume paths are relative to the project directory
            proj_dir = Path(self.pdf_path).parent.parent
            crops_dir = proj_dir / "crops"
            
            if (crops_dir / self.res.id_crop_path).exists():
                self.res._id_crop = cv2.imread(str(crops_dir / self.res.id_crop_path))
                
            if self.res.version_crop_path and (crops_dir / self.res.version_crop_path).exists():
                self.res._version_crop = cv2.imread(str(crops_dir / self.res.version_crop_path))
                
            if self.res.signature_crop_path and (crops_dir / self.res.signature_crop_path).exists():
                self.res._signature_crop = cv2.imread(str(crops_dir / self.res.signature_crop_path))
                
            for q_num, p_name in self.res.question_crop_paths.items():
                if (crops_dir / p_name).exists():
                    self.res._question_crops[q_num] = cv2.imread(str(crops_dir / p_name))
            
            # If we successfully loaded from disk, we can skip the expensive PDF extraction
            if self.res._id_crop is not None:
                return

        # 2. Fallback to PDF extraction (Slow)
        if self.res._id_crop is None and self.pdf_path:
            try:
                import fitz
                from src.core.grader import OMRGrader
                doc = fitz.open(self.pdf_path)
                page = doc.load_page(self.res.page_number - 1)
                # Ensure we use 200 DPI for fallback extraction as well!
                pix = page.get_pixmap(matrix=fitz.Matrix(200/72, 200/72), alpha=False)
                img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
                if pix.n == 3: img = img[:, :, ::-1].copy()
                doc.close()
                
                gr = OMRGrader(self.config_path)
                crops = gr.extract_crops_only(img, self.question_count)
                
                self.res._id_crop = crops.get("id_crop")
                self.res._version_crop = crops.get("version_crop")
                self.res._signature_crop = crops.get("signature_crop")
                self.res._question_crops = crops.get("question_crops")
            except Exception as e:
                print(f"Modal Image Load Error: {e}")

    def _populate_errors(self):
        # 1. ID Section (Always active for auditing)
        self.list_widget.addItem("👤 Student Identity")
        self.error_keys.append(("id", None))
        self._build_id_panel()
            
        # 2. Version Section (Always active for auditing/cheating correction)
        self.list_widget.addItem("🔖 Exam Version")
        self.error_keys.append(("version", None))
        self._build_version_panel()
            
        # 3. Question Items (Show those needing review OR already reviewed)
        all_q = set(self.res.question_errors) | set(self.res.manually_reviewed_questions)
        for q_num in sorted(list(all_q)):
            prefix = "⚠️" if q_num in self.res.question_errors else "✅"
            self.list_widget.addItem(f"{prefix} Question {q_num}")
            self.error_keys.append(("question", q_num))
            self._build_question_panel(q_num)
            
        if self.list_widget.count() > 0:
            self.list_widget.setCurrentRow(0)

    def _build_id_panel(self):
        w = QWidget()
        l = QVBoxLayout(w)
        lbl_title = QLabel("Identity Verification")
        lbl_title.setProperty("cssClass", "title")
        l.addWidget(lbl_title)
        
        lbl_sig = QLabel()
        if self.res._signature_crop is not None:
            lbl_sig.setPixmap(cv2_to_qpixmap(self.res._signature_crop))
        l.addWidget(QLabel("Student's Handwritten Name:"))
        l.addWidget(lbl_sig)
        
        l.addSpacing(15)
        
        l.addSpacing(15)
        l.addWidget(QLabel("ID Box Extract:"))
        
        # Ensure the image can be scrolled if too large
        self.lbl_id = QLabel()
        if self.res._id_crop is not None:
            pix = cv2_to_qpixmap(self.res._id_crop)
            self.lbl_id.setPixmap(pix)
        
        l.addWidget(self.lbl_id)
        
        l.addSpacing(15)
        l.addWidget(QLabel("Override Extracted ID:"))
        self.id_input = QLineEdit(self.res.student_id)
        self.id_input.setPlaceholderText("Enter corrected ID...")
        self.id_input.textChanged.connect(self._validate_review)
        l.addWidget(self.id_input)
        
        l.addStretch()
        self.stack.addWidget(w)

    def _validate_review(self):
        # Enable Accept button if anything has been changed from initial state
        has_id_change = self.id_input and self.id_input.text() != self.res.student_id
        has_ver_change = self.ver_input and self.ver_input.currentIndex() != 0
        
        has_q_change = False
        for q_num, combo in self.q_inputs.items():
            if combo.currentIndex() != 0: # Anything other than "(Keep Original Read)"
                has_q_change = True
                break
                
        self.btn_save.setEnabled(has_id_change or has_ver_change or has_q_change)

    def _build_version_panel(self):
        w = QWidget()
        l = QVBoxLayout(w)
        lbl_title = QLabel("Version Verification")
        lbl_title.setProperty("cssClass", "title")
        l.addWidget(lbl_title)
        lbl_v = QLabel()
        if self.res._version_crop is not None:
            lbl_v.setPixmap(cv2_to_qpixmap(self.res._version_crop))
        l.addWidget(lbl_v)
        
        l.addSpacing(15)
        l.addWidget(QLabel("Select Correct Version:"))
        self.ver_input = QComboBox()
        self.ver_input.addItems(["(Keep Original Read)", "A", "B", "C", "D", "E", "F"])
        
        # Set to current version if it's already set
        if self.res.version in ["A", "B", "C", "D", "E", "F"]:
            self.ver_input.setCurrentText(self.res.version)
            
        self.ver_input.currentIndexChanged.connect(self._validate_review)
        l.addWidget(self.ver_input)
        
        l.addStretch()
        self.stack.addWidget(w)

    def _build_question_panel(self, q_num: int):
        w = QWidget()
        l = QVBoxLayout(w)
        lbl_title = QLabel(f"Question {q_num} Resolution")
        lbl_title.setProperty("cssClass", "title")
        l.addWidget(lbl_title)
        
        lbl_img = QLabel()
        if q_num in self.res._question_crops:
            lbl_img.setPixmap(cv2_to_qpixmap(self.res._question_crops[q_num]))
        l.addWidget(lbl_img)
        
        l.addSpacing(15)
        l.addWidget(QLabel("Select Intended Answer:"))
        
        combo = QComboBox()
        combo.addItems(["(Keep Original Read)", "BLANK", "A", "B", "C", "D", "E", "F"])
        
        # Pre-select based on current answer state
        ans = self.res.answers.get(q_num, [])
        if not ans:
            # If it was explicitly marked as reviewed but answers are empty, it's BLANK
            if q_num in self.res.manually_reviewed_questions:
                combo.setCurrentText("BLANK")
        elif len(ans) == 1:
            val = ans[0]
            if val in ["A", "B", "C", "D", "E", "F"]:
                combo.setCurrentText(val)
        
        combo.currentIndexChanged.connect(self._validate_review)
        self.q_inputs[q_num] = combo
        l.addWidget(combo)
        
        l.addStretch()
        self.stack.addWidget(w)

    def _on_list_selected(self, index: int):
        if 0 <= index < self.stack.count():
            self.stack.setCurrentIndex(index)

    def apply_corrections(self) -> GradingResult:
        self.res.is_manual_fix = True # Mark as human-reviewed
        self.res.student_id = self.id_input.text()
        
        # Smart Check: If the doctor didn't remove the '*', it's still an error
        if "*" in self.res.student_id:
            self.res.id_error = True
        else:
            self.res.id_error = False
        
        if self.ver_input:
            v_val = self.ver_input.currentText()
            if v_val != "(Keep Original Read)":
                self.res.version = v_val
                self.res.version_error = False
                
        # Questions
        for q, combo in self.q_inputs.items():
            val = combo.currentText()
            if val != "(Keep Original Read)":
                # Track as reviewed
                if q not in self.res.manually_reviewed_questions:
                    self.res.manually_reviewed_questions.append(q)
                
                # Clear from active errors
                if q in self.res.question_errors:
                    self.res.question_errors.remove(q)
                
                if val == "BLANK":
                    self.res.answers[q] = []
                else:
                    self.res.answers[q] = [val]
                
        return self.res


class AnswerKeyDialog(QDialog):
    def __init__(self, parent, config_path, current_keys, num_questions, last_path):
        super().__init__(parent)
        self.setWindowTitle("Configure Answer Keys")
        self.setMinimumSize(700, 500)
        self.config_path = config_path
        self.num_questions = num_questions
        self.keys = current_keys # Dict[str, AnswerKey]
        self.last_path = last_path
        
        layout = QVBoxLayout(self)
        
        self.tabs = QTabWidget()
        self.version_widgets = {} # version -> list of [checkboxes]
        
        for ver in ["A", "B", "C", "D", "E", "F"]:
            tab = QWidget()
            t_layout = QVBoxLayout(tab)
            
            # Scan button
            btn_scan = QPushButton(f"Scan Model {ver} from PDF")
            btn_scan.clicked.connect(lambda _, v=ver: self._scan_key(v))
            t_layout.addWidget(btn_scan)
            
            scroll = QScrollArea()
            scroll.setWidgetResizable(True)
            scroll_content = QWidget()
            grid = QGridLayout(scroll_content)
            
            self.version_widgets[ver] = {}
            
            for q in range(1, num_questions + 1):
                grid.addWidget(QLabel(f"Q{q}:"), q-1, 0)
                q_boxes = {}
                for i, letter in enumerate(["A", "B", "C", "D", "E", "F"]):
                    cb = QCheckBox(letter)
                    # Load existing if available
                    if ver in self.keys and q in self.keys[ver].answers:
                        if letter in self.keys[ver].answers[q]:
                            cb.setChecked(True)
                    
                    grid.addWidget(cb, q-1, i + 1)
                    q_boxes[letter] = cb
                self.version_widgets[ver][q] = q_boxes
            
            scroll.setWidget(scroll_content)
            t_layout.addWidget(scroll)
            self.tabs.addTab(tab, f"Version {ver}")
            
        layout.addWidget(self.tabs)
        
        # Apply the shared stylesheet to ensure correct dark mode colors
        self.setStyleSheet(parent.styleSheet() if parent else "")
        self.tabs.setStyleSheet("""
            QTabWidget::pane { border: 1px solid #334155; top: -1px; background: #0F172A; }
            QTabBar::tab { background: #1E293B; color: #94A3B8; padding: 10px 20px; border: 1px solid #334155; border-bottom: none; border-top-left-radius: 8px; border-top-right-radius: 8px; }
            QTabBar::tab:selected { background: #0F172A; color: #38BDF8; font-weight: bold; }
            QCheckBox { color: #F8FAFC; spacing: 8px; background: transparent; }
            QCheckBox::indicator { 
                width: 20px; height: 20px; 
                background-color: #0F172A; 
                border: 1px solid #334155; 
                border-radius: 4px; 
            }
            QCheckBox::indicator:hover { border-color: #38BDF8; background-color: #1E293B; }
            QCheckBox::indicator:checked { 
                background-color: #38BDF8; 
                border-color: #0EA5E9;
            }
            QScrollArea, QScrollArea > QWidget > QWidget { 
                background-color: #0F172A; 
                border: none;
            }
            QLabel { color: #F1F5F9; font-weight: 500; }
        """)
        
        btns = QHBoxLayout()
        btn_ok = QPushButton("Save Keys")
        btn_ok.setProperty("cssClass", "primary")
        btn_ok.clicked.connect(self.accept)
        btns.addWidget(btn_ok)
        layout.addLayout(btns)

    def _scan_key(self, version):
        file, _ = QFileDialog.getOpenFileName(self, f"Select Scan for Version {version}", self.last_path, "PDF (*.pdf)")
        if not file: return
        
        self.last_path = str(Path(file).parent)
        try:
            from src.core.grader import OMRGrader
            import fitz
            grader = OMRGrader(self.config_path)
            doc = fitz.open(file)
            page = doc.load_page(0)
            pix = page.get_pixmap(matrix=fitz.Matrix(2.0, 2.0), alpha=False)
            img_data = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
            if pix.n == 3: img_data = img_data[:, :, ::-1].copy()
            
            res = grader.process_array(img_data, page_num=1, expected_questions=self.num_questions)
            doc.close()
            
            if res:
                # Update checkboxes
                for q, boxes in self.version_widgets[version].items():
                    correct_letters = res.answers.get(q, [])
                    for L, box in boxes.items():
                        box.setChecked(L in correct_letters)
                QMessageBox.information(self, "Success", f"Version {version} populated from scan.")
        except Exception as e:
            QMessageBox.critical(self, "Scan Error", str(e))

    def get_updated_keys(self):
        new_keys = {}
        from src.models.schemas import AnswerKey
        for ver, questions in self.version_widgets.items():
            ans_dict = {}
            for q, boxes in questions.items():
                selected = [L for L, box in boxes.items() if box.isChecked()]
                if selected:
                    ans_dict[q] = selected
            if ans_dict:
                new_keys[ver] = AnswerKey(version=ver, answers=ans_dict)
        return new_keys


class MainWindow(QMainWindow):
    def __init__(self, config_path: str | None = None):
        super().__init__()
        
        # Enable Frameless Mode
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.Window)
        self.setAcceptDrops(True)
        self._drag_pos = None
        
        # Fallback to local config.json relative to EXE if not provided
        self.config_path = config_path or get_config_path()
        self.pm: ProjectManager | None = None
        self.worker: PDFGraderWorker | None = None
        
        # Application Branding
        self.setWindowTitle("EUI OMR Engine - Academic Dashboard")
        self.setWindowIcon(QIcon(get_resource_path("assets/logo.ico")))
        self.setMinimumSize(1100, 750)
        self.setStyleSheet(APP_STYLESHEET)
        
        # History System
        self.undo_stack: list[list[GradingResult]] = []
        self.redo_stack: list[list[GradingResult]] = []
        
        # Shortcuts
        QShortcut(QKeySequence("Ctrl+Z"), self, self._undo)
        QShortcut(QKeySequence("Ctrl+Y"), self, self._redo)
        QShortcut(QKeySequence("Ctrl+Shift+Z"), self, self._redo)
        
        # Main Container
        self.central = QWidget()
        self.setCentralWidget(self.central)
        self.layout = QVBoxLayout(self.central)
        self.layout.setContentsMargins(0, 0, 0, 0)
        self.layout.setSpacing(0)
        
        # --- GLOBAL HEADER ---
        self.header = QWidget()
        self.header.setFixedHeight(70)
        self.header.setStyleSheet("background-color: #1E293B; border-bottom: 1px solid #334155;")
        h_layout = QHBoxLayout(self.header)
        h_layout.setContentsMargins(30, 0, 30, 0)
        
        # Logo
        self.logo_lbl = QLabel()
        logo_pix = QPixmap(get_resource_path("assets/logo_premium.png"))
        if not logo_pix.isNull():
            self.logo_lbl.setPixmap(logo_pix.scaled(40, 40, Qt.KeepAspectRatio, Qt.SmoothTransformation))
        h_layout.addWidget(self.logo_lbl)
        
        self.title_lbl = QLabel("EUI OMR DASHBOARD")
        self.title_lbl.setObjectName("header_title")
        h_layout.addWidget(self.title_lbl)
        h_layout.addStretch()
        
        # Undo/Redo Buttons
        self.btn_undo = QPushButton("↩️ Undo")
        self.btn_undo.setToolTip("Undo Last Action (Ctrl+Z)")
        self.btn_undo.setFixedWidth(80)
        self.btn_undo.clicked.connect(self._undo)
        self.btn_undo.setEnabled(False)
        self.btn_undo.setStyleSheet("""
            QPushButton { background-color: #334155; color: white; border: 1px solid #475569; padding: 5px; }
            QPushButton:hover { background-color: #475569; }
            QPushButton:disabled { background-color: #1E293B; color: #475569; }
        """)
        h_layout.addWidget(self.btn_undo)
        
        self.btn_redo = QPushButton("Redo ↪️")
        self.btn_redo.setToolTip("Redo Last Action (Ctrl+Y)")
        self.btn_redo.setFixedWidth(80)
        self.btn_redo.clicked.connect(self._redo)
        self.btn_redo.setEnabled(False)
        self.btn_redo.setStyleSheet("""
            QPushButton { background-color: #334155; color: white; border: 1px solid #475569; padding: 5px; }
            QPushButton:hover { background-color: #475569; }
            QPushButton:disabled { background-color: #1E293B; color: #475569; }
        """)
        h_layout.addWidget(self.btn_redo)
        
        h_layout.addSpacing(10)
        
        self.project_name_lbl = QLabel("No Project Loaded")
        self.project_name_lbl.setStyleSheet("color: #94A3B8; font-weight: 500;")
        h_layout.addWidget(self.project_name_lbl)
        
        h_layout.addStretch()
        
        self.btn_min = QPushButton("−")
        self.btn_min.setObjectName("min_btn")
        self.btn_min.setFixedSize(45, 45)
        self.btn_min.clicked.connect(self.showMinimized)
        h_layout.addWidget(self.btn_min)
        
        self.btn_close = QPushButton("X")
        self.btn_close.setObjectName("close_btn")
        self.btn_close.setFixedSize(45, 45)
        self.btn_close.clicked.connect(self.close)
        h_layout.addWidget(self.btn_close)
        
        self.layout.addWidget(self.header)
        
        # --- MAIN STACK ---
        self.stack = QStackedWidget()
        self.layout.addWidget(self.stack)
        
        # Persistent Settings Memory (Windows Registry / Config)
        self.settings = QSettings("EUI", "OMREngine")
        self.docs_path = QStandardPaths.writableLocation(QStandardPaths.DocumentsLocation)
        
        self.last_project_path = self.settings.value("last_project_path", self.docs_path)
        self.last_excel_path = self.settings.value("last_excel_path", self.docs_path)
        self.last_pdf_path = self.settings.value("last_pdf_path", self.docs_path)
        self.recent_projects = self.settings.value("recent_projects", [])
        if isinstance(self.recent_projects, str): # Handle QSettings single item quirk
            self.recent_projects = [self.recent_projects]
            
        # Clean up invalid recent projects
        valid_recents = [p for p in self.recent_projects if (Path(p) / "project_state.json").exists()]
        if len(valid_recents) != len(self.recent_projects):
            self.recent_projects = valid_recents
            self.settings.setValue("recent_projects", self.recent_projects)
        
        # Build views
        self.build_welcome_view()
        self.build_setup_view()
        self.build_processing_view()
        
        self.stack.setCurrentIndex(0)

    def update_header_project(self, name: str):
        self.project_name_lbl.setText(f"Project: {name}")

    # ---------------------------------------------------------
    # VIEW 1: WELCOME / LOAD PROJECT
    # ---------------------------------------------------------
    # --- Window Dragging Logic (Frameless Support) ---
    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            # Check if click is on the header area
            if self.header.underMouse():
                self._drag_pos = event.globalPosition().toPoint()
                event.accept()

    def mouseMoveEvent(self, event):
        if event.buttons() == Qt.LeftButton and self._drag_pos:
            delta = event.globalPosition().toPoint() - self._drag_pos
            self.move(self.pos() + delta)
            self._drag_pos = event.globalPosition().toPoint()
            event.accept()

    def mouseReleaseEvent(self, event):
        self._drag_pos = None

    def dragEnterEvent(self, event):
        if self.stack.currentIndex() == 0 and event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if urls and urls[0].isLocalFile():
                event.acceptProposedAction()
                return
        event.ignore()

    def dropEvent(self, event):
        if self.stack.currentIndex() == 0 and event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if urls and urls[0].isLocalFile():
                folder_path = urls[0].toLocalFile()
                if os.path.isdir(folder_path):
                    self._do_load_project(folder_path)
                else:
                    QMessageBox.warning(self, "Invalid Drop", "Please drop a valid project folder, not a file.")

    def build_welcome_view(self):
        w = QWidget()
        l = QVBoxLayout(w)
        l.setContentsMargins(100, 50, 100, 50)
        l.setAlignment(Qt.AlignCenter)
        
        # Hero Image/Logo
        l.setSpacing(30)
        
        l.addStretch()
        title = QLabel("Welcome to EUI OMR Engine")
        title.setProperty("cssClass", "title")
        l.addWidget(title, alignment=Qt.AlignCenter)
        
        desc = QLabel("Streamlined optical mark recognition for academic excellence.")
        desc.setProperty("cssClass", "subtitle")
        l.addWidget(desc, alignment=Qt.AlignCenter)
        
        lbl_drag = QLabel("<small><i>✨ Tip: You can seamlessly drag and drop a project folder anywhere here to load it.</i></small>")
        lbl_drag.setStyleSheet("color: #475569; margin-top: 5px; margin-bottom: 10px;")
        l.addWidget(lbl_drag, alignment=Qt.AlignCenter)
        
        actions = QHBoxLayout()
        actions.setSpacing(20)
        
        card_new = QWidget()
        card_new.setObjectName("card")
        cl = QVBoxLayout(card_new)
        cl.setContentsMargins(30,30,30,30)
        cl.addWidget(QLabel("<b>New Project</b>"), alignment=Qt.AlignCenter)
        cl.addWidget(QLabel("Start a fresh grading session from PDF scans."), alignment=Qt.AlignCenter)
        btn_new = QPushButton("Create New Project")
        btn_new.setProperty("cssClass", "primary")
        btn_new.clicked.connect(self._create_project)
        cl.addWidget(btn_new)
        actions.addWidget(card_new)
        
        card_load = QWidget()
        card_load.setObjectName("card")
        cl2 = QVBoxLayout(card_load)
        cl2.setContentsMargins(30,30,30,30)
        cl2.addWidget(QLabel("<b>Open Project</b>"), alignment=Qt.AlignCenter)
        cl2.addWidget(QLabel("Resume an existing grading session."), alignment=Qt.AlignCenter)
        btn_load = QPushButton("Load Saved State")
        btn_load.clicked.connect(self._load_project)
        cl2.addWidget(btn_load)
        actions.addWidget(card_load)
        
        l.addLayout(actions)
        
        # CARD: Recent Projects
        if self.recent_projects:
            card_recent = QWidget()
            card_recent.setObjectName("card")
            card_recent.setFixedWidth(640)
            rl = QVBoxLayout(card_recent)
            rl.setContentsMargins(20, 20, 20, 20)
            
            rl.addWidget(QLabel("<b>🕒 Recent Projects</b>"))
            self.recent_list = QListWidget()
            self.recent_list.setFixedHeight(120)
            self.recent_list.addItems([str(Path(p).name) for p in self.recent_projects])
            self.recent_list.itemDoubleClicked.connect(self._on_recent_double_clicked)
            rl.addWidget(self.recent_list)
            rl.addWidget(QLabel("Double-click to open quick-load."), alignment=Qt.AlignRight)
            
            l.addWidget(card_recent, alignment=Qt.AlignCenter)
        
        l.addStretch()
        self.stack.addWidget(w)

    def _add_to_recent(self, path: str):
        if path in self.recent_projects:
            self.recent_projects.remove(path)
        self.recent_projects.insert(0, path)
        self.recent_projects = self.recent_projects[:5] # Keep last 5
        self.settings.setValue("recent_projects", self.recent_projects)

    def _on_recent_double_clicked(self, item):
        idx = self.recent_list.row(item)
        if 0 <= idx < len(self.recent_projects):
            path = self.recent_projects[idx]
            if (Path(path) / "project_state.json").exists():
                self._do_load_project(path)
            else:
                QMessageBox.warning(self, "Missing Project", "This project is invalid or no longer exists. Removing from recent list.")
                self.recent_projects.remove(path)
                self.settings.setValue("recent_projects", self.recent_projects)
                self.recent_list.takeItem(idx)

    def _create_project(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Empty Folder for New Project", self.last_project_path)
        if folder:
            self._do_create_project(folder)

    def _do_create_project(self, folder: str):
        # Safety Check: Is this already a project folder?
        state_file = Path(folder) / "project_state.json"
        if state_file.exists():
            reply = QMessageBox.question(
                self, "Project Already Exists",
                "This folder already contains an existing project.\n\n"
                "• Click 'Yes' to overwrite (DELETE everything and start fresh).\n"
                "• Click 'No' to load the existing project instead.\n"
                "• Click 'Cancel' to abort.",
                QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel,
                QMessageBox.Cancel
            )
            if reply == QMessageBox.Cancel:
                return
            elif reply == QMessageBox.No:
                self._do_load_project(folder)
                return
            # If Yes, we proceed to overwrite with create_project() below

        self._reset_internal_state()
        self.last_project_path = folder
        self.settings.setValue("last_project_path", folder)
        self.pm = ProjectManager(folder)
        self.pm.create_project(Path(folder).name)
        self.update_header_project(self.pm.project_name)
        self._add_to_recent(folder)
        self._transition_to_setup()

    def _reset_internal_state(self):
        """Clears all UI variables and state for a fresh session."""
        self.results_data = []
        self.table.setRowCount(0)
        self.lbl_excel.setText("No Excel File Selected")
        self.lbl_setup_pdf.setText("No PDF Selected")
        self.lbl_pdf.setText("")
        self.stat_total.setText("Total: 0")
        self.stat_success.setText("Resolved: 0")
        self.stat_errors.setText("Review Needed: 0")
        self.active_pdf_path = ""
        self.btn_resume.setVisible(False)
        self.btn_start.setText("Start Grading & Evaluation →")
        self.btn_start.setProperty("cssClass", "primary")
        self.btn_start.style().unpolish(self.btn_start)
        self.btn_start.style().polish(self.btn_start)

    def _load_project(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Existing Project Folder", self.last_project_path)
        if folder:
            self._do_load_project(folder)

    def _do_load_project(self, folder: str):
        state_file = Path(folder) / "project_state.json"
        if not state_file.exists():
            QMessageBox.warning(self, "Invalid Project Folder", "No saved project state found in this folder.\n\nPlease select a valid project folder or create a 'New Project' instead.")
            return

        self.last_project_path = folder
        self.settings.setValue("last_project_path", folder)
        self.pm = ProjectManager(folder)
        if self.pm.logger:
            self.pm.logger.info(f"User loaded project folder: {folder}")
        self.update_header_project(self.pm.project_name)
        self._add_to_recent(folder)
        self._transition_to_setup()

    # ---------------------------------------------------------
    # VIEW 2: SETUP EXCEL AND ANSWER KEYS
    # ---------------------------------------------------------
    def build_setup_view(self):
        self.v_setup = QWidget()
        l = QVBoxLayout(self.v_setup)
        l.setContentsMargins(50, 30, 50, 30)
        l.setSpacing(20)
        
        title = QLabel("Session Setup")
        title.setProperty("cssClass", "title")
        l.addWidget(title)
        
        # Card 1: Roster Configuration
        card_roster = QWidget()
        card_roster.setObjectName("card")
        rl = QVBoxLayout(card_roster)
        rl.setContentsMargins(25, 25, 25, 25)
        
        # Group 1: Data Sources (Card 1)
        card_roster = QWidget()
        card_roster.setObjectName("card")
        rl = QVBoxLayout(card_roster)
        rl.setContentsMargins(25, 20, 25, 20)
        
        rl.addWidget(QLabel("<b>1. Data Sources & Roster</b>"))
        h_excel = QHBoxLayout()
        self.btn_excel = QPushButton("Import Student Roster (Excel)")
        self.btn_excel.clicked.connect(self._import_excel)
        h_excel.addWidget(self.btn_excel)
        self.lbl_excel = QLabel("No Excel File Selected")
        self.lbl_excel.setStyleSheet("color: #38BDF8;")
        h_excel.addWidget(self.lbl_excel)
        h_excel.addStretch()
        rl.addLayout(h_excel)

        rl.addSpacing(10)
        self.btn_keys = QPushButton("Manage & Scan Answer Keys (Forms A-F)")
        self.btn_keys.setFixedHeight(40)
        self.btn_keys.setFixedWidth(300)
        self.btn_keys.clicked.connect(self._manage_keys)
        rl.addWidget(self.btn_keys)
        
        l.addWidget(card_roster)
        l.addSpacing(10)

        # Group 2: Configuration (Card 2)
        card_exam = QWidget()
        card_exam.setObjectName("card")
        el = QVBoxLayout(card_exam)
        el.setContentsMargins(25, 20, 25, 20)
        
        el.addWidget(QLabel("<b>2. Data Mapping & Exam Size</b>"))
        h_set = QHBoxLayout()
        h_set.addWidget(QLabel("Number of Questions (1-60):"))
        self.spin_q = QSpinBox()
        self.spin_q.setRange(1, 60); self.spin_q.setValue(60)
        h_set.addWidget(self.spin_q)
        h_set.addSpacing(30)
        h_set.addWidget(QLabel("ID Column:"))
        self.cb_id = QComboBox()
        h_set.addWidget(self.cb_id)
        h_set.addSpacing(20)
        h_set.addWidget(QLabel("Grade Column:"))
        self.cb_out = QComboBox()
        h_set.addWidget(self.cb_out)
        h_set.addStretch()
        el.addLayout(h_set)
        
        l.addWidget(card_exam)
        l.addSpacing(10)

        # Group 3: Engine Calibration (The Final Step)
        card_calib = QWidget()
        card_calib.setObjectName("card")
        cl = QVBoxLayout(card_calib)
        cl.setContentsMargins(25, 20, 25, 20)
        
        cl.addWidget(QLabel("<b>3. Engine Calibration (Optimize Detection)</b>"))
        
        h_pdf = QHBoxLayout()
        self.btn_select_pdf = QPushButton("📂 1. Select Student Scans (PDF)")
        self.btn_select_pdf.setFixedWidth(250)
        self.btn_select_pdf.setFixedHeight(40)
        self.btn_select_pdf.clicked.connect(self._select_pdf)
        h_pdf.addWidget(self.btn_select_pdf)
        
        self.lbl_setup_pdf = QLabel("No PDF Selected")
        self.lbl_setup_pdf.setStyleSheet("color: #38BDF8; font-weight: bold;")
        h_pdf.addWidget(self.lbl_setup_pdf)
        h_pdf.addStretch()
        cl.addLayout(h_pdf)
        
        cl.addSpacing(10)
        
        h_sens = QHBoxLayout()
        self.btn_auto_sens = QPushButton("✨ 2. Intelligent Auto-Calibration")
        self.btn_auto_sens.setToolTip("Analyzes the first page to find the perfect sensitivity automatically.")
        self.btn_auto_sens.setFixedWidth(250)
        self.btn_auto_sens.setFixedHeight(40)
        self.btn_auto_sens.clicked.connect(self._on_auto_optimize)
        self.btn_auto_sens.setEnabled(False) 
        h_sens.addWidget(self.btn_auto_sens)
        
        self.slider_sens = QSlider(Qt.Horizontal)
        self.slider_sens.setRange(40, 95); self.slider_sens.setValue(75)
        self.slider_sens.setTickPosition(QSlider.TicksBelow); self.slider_sens.setTickInterval(5)
        self.slider_sens.valueChanged.connect(self._on_sensitivity_changed)
        h_sens.addWidget(self.slider_sens)
        
        self.lbl_sens_val = QLabel("75% (Recommended)")
        self.lbl_sens_val.setFixedWidth(150)
        h_sens.addWidget(self.lbl_sens_val)
        cl.addLayout(h_sens)
        
        cl.addWidget(QLabel("<small><i>Tip: Use Auto-Calibration if marks are faint or background is noisy.</i></small>"))
        
        l.addWidget(card_calib)
        l.addStretch()
        
        # Bottom Buttons
        btns = QHBoxLayout()
        btn_back = QPushButton("← Global Settings")
        btn_back.clicked.connect(lambda: self.stack.setCurrentIndex(0))
        btns.addWidget(btn_back)
        btns.addStretch()
        
        self.btn_resume = QPushButton("📊 View Results Dashboard →")
        self.btn_resume.setProperty("cssClass", "success")
        self.btn_resume.setFixedHeight(50)
        self.btn_resume.setFixedWidth(250)
        self.btn_resume.clicked.connect(self._resume_to_dashboard)
        self.btn_resume.setVisible(False)
        btns.addWidget(self.btn_resume)

        self.btn_start = QPushButton("Start Fresh Batch →")
        self.btn_start.setFixedHeight(50)
        self.btn_start.setFixedWidth(200)
        self.btn_start.clicked.connect(self._save_setup_and_continue)
        self.btn_start.setEnabled(True)
        btns.addWidget(self.btn_start)
        l.addLayout(btns)
        
        self.stack.addWidget(self.v_setup)

    def _transition_to_setup(self):
        # Prepopulate loaded data
        if self.pm.excel_roster_path:
            self.lbl_excel.setText(Path(self.pm.excel_roster_path).name)
            self._populate_excel_columns(self.pm.excel_roster_path)
            self.cb_id.setCurrentText(self.pm.student_id_col)
            self.cb_out.setCurrentText(self.pm.grade_output_col)
            
        if self.pm.student_pdf_path:
            self.active_pdf_path = self.pm.student_pdf_path
            self.lbl_setup_pdf.setText(f"Target: {Path(self.active_pdf_path).name}")
            self.btn_auto_sens.setEnabled(True)
            
        self._validate_setup_ready()
            
        # RESTORE PROJECT STATE
        self.results_data = copy.deepcopy(self.pm.last_results)
        self.results_data.sort(key=lambda x: x.page_number) # Ensure sorted order on load
        
        self.table.setRowCount(0)
        for res in self.results_data:
            self._update_table_row(res)
            
        # Show Resume button if we have data
        has_data = len(self.results_data) > 0
        self.btn_resume.setVisible(has_data)
        self._update_stats()
        
        # If we have data, change text and color of start button to reflect it's a 'Reset'
        if has_data:
            self.btn_start.setText("Re-run All Scans ↺")
            self.btn_start.setProperty("cssClass", "") # Make it gray/secondary
            self.btn_start.setStyleSheet("") # Reset any custom style
        else:
            self.btn_start.setText("Start Grading & Evaluation →")
            self.btn_start.setProperty("cssClass", "primary")

        self.btn_start.style().unpolish(self.btn_start)
        self.btn_start.style().polish(self.btn_start)

        self.spin_q.setValue(self.pm.question_count)
        self.slider_sens.setValue(self.pm.mark_sensitivity)
        self._on_sensitivity_changed(self.pm.mark_sensitivity)
        self.stack.setCurrentIndex(1)
        
    def _import_excel(self):
        file, _ = QFileDialog.getOpenFileName(self, "Select Roster", self.last_excel_path, "Excel (*.xlsx *.xls)")
        if file:
            p = str(Path(file).parent)
            self.last_excel_path = p
            self.settings.setValue("last_excel_path", p)
            safe_path = self.pm.import_excel_file(file)
            self.lbl_excel.setText(safe_path.name)
            self._populate_excel_columns(safe_path)
            self._validate_setup_ready()

    def _populate_excel_columns(self, path):
        em = ExcelManager(path)
        cols = em.get_columns()
        self.cb_id.clear()
        self.cb_id.addItems(cols)
        self.cb_out.clear()
        self.cb_out.addItems(cols)

    def _validate_setup_ready(self):
        pass # Explicitly making sure logic doesn't override btn_start to False

    def _resume_to_dashboard(self):
        self.lbl_pdf.setText("Results Loaded ✅")
        self.stack.setCurrentIndex(2)

    def _save_setup_and_continue(self):
        if not self.pm or not self.pm.excel_roster_path or not self.pm.student_pdf_path:
            QMessageBox.warning(self, "Setup Incomplete", "Please ensure both an Excel roster and Student PDF scans are selected before starting.")
            return
            
        self.pm.student_id_col = self.cb_id.currentText()
        self.pm.grade_output_col = self.cb_out.currentText()
        self.pm.question_count = self.spin_q.value()
        self.pm.mark_sensitivity = self.slider_sens.value()
        self.pm.save_state()
        
        # Move to Processing View
        self.stack.setCurrentIndex(2)
        
        # AUTO-START: Since everything is configured, start grading immediately
        QTimer.singleShot(100, self._start_grading)

    # ---------------------------------------------------------
    # VIEW 3: GRADING / PROCESSING
    # ---------------------------------------------------------
    def build_processing_view(self):
        self.v_proc = QWidget()
        l = QVBoxLayout(self.v_proc)
        l.setContentsMargins(30, 30, 30, 30)
        l.setSpacing(20)
        
        # Dashboard Top Actions
        card_actions = QWidget()
        card_actions.setObjectName("card")
        al = QHBoxLayout(card_actions)
        al.setContentsMargins(20, 20, 20, 20)
        
        al.addStretch()
        self.lbl_pdf = QLabel("") # Start empty
        al.addWidget(self.lbl_pdf)
        l.addWidget(card_actions)
        
        # Main Dashboard Area
        dashboard_content = QHBoxLayout()
        
        # Left Side: Table
        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["Page #", "Extracted ID", "Model", "Condition"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.itemChanged.connect(self._on_table_item_changed)
        self.table.itemSelectionChanged.connect(self._on_selection_changed)
        dashboard_content.addWidget(self.table, stretch=3)
        
        # Right Side: Status Sidebar
        sidebar = QWidget()
        sidebar.setFixedWidth(250)
        sidebar.setObjectName("card")
        sl = QVBoxLayout(sidebar)
        sl.addWidget(QLabel("<b>Grading Summary</b>"))
        self.stat_total = QLabel("Total: 0")
        self.stat_success = QLabel("Resolved: 0")
        self.stat_errors = QLabel("Review Needed: 0")
        self.stat_errors.setStyleSheet("color: #F59E0B;")
        sl.addWidget(self.stat_total)
        sl.addWidget(self.stat_success)
        sl.addWidget(self.stat_errors)
        sl.addStretch()
        
        self.btn_export = QPushButton("EXPORT FINAL MARKS")
        self.btn_export.setProperty("cssClass", "primary")
        self.btn_export.setFixedHeight(50)
        self.btn_export.setEnabled(False)
        self.btn_export.clicked.connect(self._export_to_excel)
        sl.addWidget(self.btn_export)
        
        btn_back = QPushButton("Settings")
        btn_back.clicked.connect(lambda: self.stack.setCurrentIndex(1))
        sl.addWidget(btn_back)
        
        dashboard_content.addWidget(sidebar)
        l.addLayout(dashboard_content, stretch=1)
        
        self.stack.addWidget(self.v_proc)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.hide()
        l.addWidget(self.progress_bar)
        
        self.active_pdf_path = ""
        self.results_data: list[GradingResult] = [] 
        self.table.cellDoubleClicked.connect(self._on_table_double_clicked)
        
        # Load existing results if any
        if self.pm and self.pm.last_results:
            self.results_data = sorted(self.pm.last_results, key=lambda x: x.page_number)
            for res in self.results_data:
                self._update_table_row(res)
            self.btn_export.setEnabled(True)

    def _on_sensitivity_changed(self, val):
        rec = " (Recommended)" if val == 75 else ""
        self.lbl_sens_val.setText(f"{val}%{rec}")

    def _select_pdf(self):
        file, _ = QFileDialog.getOpenFileName(self, "Select Student PDF Scans", self.last_pdf_path, "PDF Documents (*.pdf)")
        if file:
            p = str(Path(file).parent)
            self.last_pdf_path = p
            self.settings.setValue("last_pdf_path", self.last_pdf_path)
            
            # Use Mirroring logic
            safe_pdf_path = self.pm.import_pdf_file(file)
            self.active_pdf_path = str(safe_pdf_path)
            
            self.lbl_setup_pdf.setText(f"Selected: {safe_pdf_path.name}")
            self.btn_auto_sens.setEnabled(True)
            self._validate_setup_ready()

    def _on_auto_optimize(self):
        if not self.active_pdf_path:
            return
            
        self.btn_auto_sens.setText("✨ Calibrating...")
        self.btn_auto_sens.setEnabled(False)
        
        self.tune_worker = AutoTuneWorker(
            self.active_pdf_path, 
            self.config_path, 
            self.spin_q.value()
        )
        self.tune_worker.finished.connect(self._on_auto_tune_finished)
        self.tune_worker.error.connect(lambda e: QMessageBox.critical(self, "Calibration Error", e))
        self.tune_worker.start()

    def _on_auto_tune_finished(self, best_sens):
        self.slider_sens.setValue(best_sens)
        self.btn_auto_sens.setText("✨ Smart Calibration")
        self.btn_auto_sens.setEnabled(True)
        QMessageBox.information(self, "Calibration Complete", 
                               f"Optimized sensitivity to {best_sens}% based on the sample page.")

    def _start_grading(self):
        if not self.active_pdf_path: return
        
        if self.pm and self.pm.logger:
            self.pm.logger.info(f"Starting fresh grading batch for PDF: {self.active_pdf_path}")
        
        self.lbl_pdf.setText("Processing Student Scans (Please Wait)...")
        
        # Wipe protection
        if self.results_data:
            reply = QMessageBox.warning(
                self, "Reset Current Session?",
                "You have existing results in this session. Re-running will delete all manual edits and start fresh.\n\nDo you want to continue?",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No
            )
            if reply == QMessageBox.No: return

        # Clear previous session data safely
        self.results_data = []
        self.table.setRowCount(0)
        self._update_stats()
        
        self.btn_select_pdf.setEnabled(False)
        self.btn_start.setEnabled(False)
        self.progress_bar.setValue(0)
        self.progress_bar.show()
        
        self.worker = PDFGraderWorker(
            self.active_pdf_path, 
            self.config_path, 
            expected_questions=self.pm.question_count,
            sensitivity=self.pm.mark_sensitivity,
            crops_dir=self.pm.crops_dir
        )
        self.worker.progress_updated.connect(self._on_progress)
        self.worker.page_processed.connect(self._on_page_processed)
        self.worker.review_required.connect(self._on_review_required)
        self.worker.error_occurred.connect(self._on_error)
        self.worker.finished.connect(self._on_finished)
        self.worker.start()

    @Slot(int, int)
    def _on_progress(self, current, total):
        self.progress_bar.setMaximum(total)
        self.progress_bar.setValue(current)

    @Slot(object)
    def _on_review_required(self, res: GradingResult):
        # Everything goes through _on_page_processed.
        pass

    @Slot(object)
    def _on_page_processed(self, res: GradingResult):
        # 1. Thread-safe data update (ensure no duplicates by page number)
        self.results_data = [r for r in self.results_data if r.page_number != res.page_number]
        self.results_data.append(res)
        self.results_data.sort(key=lambda x: x.page_number)
        
        # 2. Persistence via PM
        self.pm.add_or_update_result(res) 
        
        # 3. UI Update
        self._update_table_row(res)
        
        if self.pm and self.pm.logger:
            self.pm.logger.debug(f"UI synchronized page {res.page_number}")

    def _update_table_row(self, res: GradingResult, existing_row: int = -1):
        if existing_row == -1:
            row_idx = self.table.rowCount()
            self.table.insertRow(row_idx)
        else:
            row_idx = existing_row
            
        # Page Number (Non-editable)
        item_page = QTableWidgetItem(f"Page {res.page_number}")
        item_page.setFlags(item_page.flags() & ~Qt.ItemIsEditable)
        self.table.setItem(row_idx, 0, item_page)
        
        # Student ID (Editable)
        item_id = QTableWidgetItem(res.student_id)
        self.table.setItem(row_idx, 1, item_id)
        
        # Version (Editable)
        item_ver = QTableWidgetItem(res.version)
        self.table.setItem(row_idx, 2, item_ver)
        
        # Condition (Non-editable)
        if res.is_manual_fix:
            status = QTableWidgetItem("🛠️ Reviewed")
            status.setForeground(QColor("#38BDF8")) # Cyan/Sky Blue for Review
        elif res.id_error or res.version_error or len(res.question_errors) > 0:
            status = QTableWidgetItem("⚠️ Needs Review")
            status.setForeground(Qt.red)
        else:
            status = QTableWidgetItem("✅ Success")
            status.setForeground(Qt.green)
        status.setFlags(status.flags() & ~Qt.ItemIsEditable)
        self.table.setItem(row_idx, 3, status)
        
        self.table.scrollToBottom()
        self._update_stats()

    def _on_selection_changed(self):
        """No longer used for live preview, but kept as a slot if needed for other UI updates."""
        pass

    def _set_pixmap(self, label, ndarray):
        if ndarray is None:
            label.setText("No Image")
            return
        
        h, w = ndarray.shape[:2]
        # Handle BGR -> RGB for QImage
        qimg = QImage(ndarray.data, w, h, w*3, QImage.Format_RGB888).rgbSwapped()
        pix = QPixmap.fromImage(qimg)
        
        # Scale to fit label comfortably
        scaled = pix.scaled(label.width()-10, label.height()-10, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        label.setPixmap(scaled)

    @Slot(QTableWidgetItem)
    def _on_table_item_changed(self, item: QTableWidgetItem):
        col = item.column()
        row = item.row()
        
        # Only handle edits on ID (1) and Version (2)
        if col not in [1, 2] or row >= len(self.results_data):
            return
            
        res = self.results_data[row]
        new_text = item.text().strip()
        
        # Detect if anything actually changed to avoid infinite save loops
        changed = False
        if col == 1 and res.student_id != new_text:
            res.student_id = new_text
            res.id_error = False 
            res.is_manual_fix = True # Mark as audit-trail fix
            changed = True
        elif col == 2 and res.version != new_text:
            res.version = new_text
            res.version_error = False
            res.is_manual_fix = True
            changed = True
            
        if changed:
            # Update persistence
            self.pm.add_or_update_result(res)
            self._update_stats()
            
            # Refresh the row UI (specifically the condition column) 
            # while blocking signals to prevent recursion
            self.table.blockSignals(True)
            self._update_table_row(res, existing_row=row)
            self.table.blockSignals(False)

    def _update_stats(self):
        total = len(self.results_data)
        errors = sum(1 for r in self.results_data if (r.id_error or r.version_error or r.question_errors))
        success = total - errors
        
        self.stat_total.setText(f"Total: {total}")
        self.stat_success.setText(f"Resolved: {success}")
        self.stat_errors.setText(f"Review Needed: {errors}")
        self.btn_export.setEnabled(total > 0 and errors == 0)
        
    def _on_table_double_clicked(self, row, col):
        """Opens the advanced ReviewModal for the selected student."""
        if row >= len(self.results_data):
            return
            
        res = self.results_data[row]
        from src.utils.paths import get_config_path
        
        # 1. Snapshot for undo
        self._push_history()
        
        # 2. Clone the object so the dialog handles a "scratchpad" version
        import copy
        res_clone = copy.deepcopy(res)
        
        dlg = ReviewModal(
            self, 
            res_clone, 
            self.active_pdf_path, 
            self.pm.question_count,
            get_config_path()
        )
        if dlg.exec():
            # 3. User accepted! Apply corrections to the clone and update main list
            corrected_res = dlg.apply_corrections()
            self.results_data[row] = corrected_res
            self.pm.add_or_update_result(corrected_res) # Persistence
            self._update_table_row(corrected_res, existing_row=row)
            self.redo_stack.clear()
        else:
            # 4. User cancelled!
            self.undo_stack.pop()
            if not self.undo_stack: self.btn_undo.setEnabled(False)

    def _on_auto_optimize(self):
        """Analyzes multiple sample pages of the PDF to find the real optimal sensitivity."""
        if not self.active_pdf_path: return
        
        self.btn_auto_sens.setEnabled(False)
        self.btn_auto_sens.setText("✨ Analyzing Samples (1-3)...")
        
        # Real Engine Calibration using the Worker
        self.tune_worker = AutoTuneWorker(
            self.active_pdf_path, 
            get_config_path(), 
            self.spin_q.value()
        )
        self.tune_worker.finished.connect(self._finish_auto_optimize)
        self.tune_worker.error.connect(self._on_tune_error)
        self.tune_worker.start()

    def _on_tune_error(self, message):
        self.btn_auto_sens.setEnabled(True)
        self.btn_auto_sens.setText("✨ Intelligent Auto-Calibration")
        QMessageBox.warning(self, "Calibration Error", f"Failed to auto-tune: {message}")

    def _finish_auto_optimize(self, best_sens):
        # Now we apply the REAL calculated value from the engine analysis
        self.slider_sens.setValue(best_sens)
        self.btn_auto_sens.setEnabled(True)
        self.btn_auto_sens.setText("✨ Intelligent Auto-Calibration")
        
        QMessageBox.information(
            self, "Auto-Calibration Complete", 
            f"Optimization Finished!\n\nSensitivity has been tuned to {best_sens}% based on the analysis of scan background and clarity."
        )

    def _manage_keys(self):
        if not self.pm: return
        dlg = AnswerKeyDialog(self, self.config_path, self.pm.answer_keys, self.spin_q.value(), self.last_pdf_path)
        if dlg.exec():
            # Update keys
            for ver, key in dlg.get_updated_keys().items():
                self.pm.set_answer_key(ver, key.answers)
            QMessageBox.information(self, "Success", "Answer keys updated successfully.")

    def _push_history(self):
        """Saves current state to undo stack using Deep Copy."""
        # Using deepcopy ensures images and lists are fully cloned and isolated
        snapshot = [copy.deepcopy(r) for r in self.results_data]
        self.undo_stack.append(snapshot)
        if len(self.undo_stack) > 50: self.undo_stack.pop(0)
        self.btn_undo.setEnabled(True)

    def _undo(self):
        if not self.undo_stack: return
        
        # Save current state to redo (Deep Copy)
        current_state = [copy.deepcopy(r) for r in self.results_data]
        self.redo_stack.append(current_state)
        self.btn_redo.setEnabled(True)
        
        # Restore last state
        self.results_data = self.undo_stack.pop()
        self.btn_undo.setEnabled(len(self.undo_stack) > 0)
        
        # Sync with Persistence
        self.pm.last_results = self.results_data
        self.pm.save_state()
        
        # Refresh UI
        self._refresh_table()

    def _redo(self):
        if not self.redo_stack: return
        
        # Push current to undo
        self._push_history()
        
        # Restore from redo
        self.results_data = self.redo_stack.pop()
        self.btn_redo.setEnabled(len(self.redo_stack) > 0)
        
        # Sync with Persistence
        self.pm.last_results = self.results_data
        self.pm.save_state()
        
        # Refresh UI
        self._refresh_table()

    def _refresh_table(self):
        """Full table rebuild (used after undo/redo)."""
        self.results_data.sort(key=lambda x: x.page_number)
        self.table.setRowCount(0)
        for res in self.results_data:
            self._update_table_row(res)

    @Slot(str, str)
    def _on_error(self, title: str, message: str):
        QMessageBox.critical(self, title, message)

    @Slot()
    def _on_finished(self):
        self.progress_bar.hide()
        self._update_stats()
        self.btn_select_pdf.setEnabled(True)
        self.btn_start.setEnabled(True)
        self.lbl_pdf.setText("Grading Complete ✅")

    def _export_to_excel(self):
        if not self.pm or not self.pm.excel_roster_path or not Path(self.pm.excel_roster_path).exists():
            QMessageBox.critical(self, "Export Error", "No valid Excel roster selected. Please configure it in Settings.")
            return

        from src.data.excel import ExcelManager
        em = ExcelManager(self.pm.excel_roster_path)
        try:
            dicts = [r.to_dict() for r in self.results_data]
            em.export_grades(
                results_data=dicts,
                student_id_col=self.pm.student_id_col,
                grade_col=self.pm.grade_output_col,
                question_count=self.pm.question_count,
                answer_keys=self.pm.answer_keys
            )
            
            msg = f"Grades exported to {Path(self.pm.excel_roster_path).name}\n\nWould you like to open the file now?"
            reply = QMessageBox.question(self, "Export Successful", msg, QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
            
            if reply == QMessageBox.Yes:
                os.startfile(self.pm.excel_roster_path)
                
        except Exception as e:
            QMessageBox.critical(self, "Export Error", str(e))


