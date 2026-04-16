"""Main GUI layout using PySide6. Implements the full Project Wizard multi-view."""

import sys
import os
import copy
from pathlib import Path

from PySide6.QtCore import Qt, Slot, QSize, QStandardPaths, QSettings, QTimer, Signal
from PySide6.QtGui import QIcon, QFont, QImage, QPixmap, QShortcut, QKeySequence, QColor
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
    QPushButton, QLabel, QFileDialog, QTableWidget, 
    QTableWidgetItem, QProgressBar, QHeaderView, QMessageBox,
    QStackedWidget, QLineEdit, QComboBox, QSpinBox, QDialog,
    QTabWidget, QScrollArea, QCheckBox, QGridLayout, QSlider,
    QListWidget, QSpacerItem, QSizePolicy, QGraphicsOpacityEffect
)

import cv2
import numpy as np

from src.core.project import ProjectManager
from src.ui.workers import PDFGraderWorker, AutoTuneWorker
from src.models.schemas import GradingResult
from src.data.excel import ExcelManager
from src.ui.style import COLORS, APP_QSS, STYLE_TOKENS
from src.utils.paths import get_resource_path, get_config_path
from PySide6.QtCore import (
    QPropertyAnimation, QEasingCurve, QPoint, 
    QParallelAnimationGroup, QSequentialAnimationGroup
)

# --- Animation Utilities ---

class SlidingStackedWidget(QStackedWidget):
    animation_finished = Signal(int)

    def __init__(self, parent=None):
        super(SlidingStackedWidget, self).__init__(parent)
        self.m_duration = 400
        self.m_easing = QEasingCurve.OutQuart
        self.m_active = False

    def slideInIdx(self, idx):
        if idx == self.currentIndex():
            return
        
        # Calculate direction
        old_idx = self.currentIndex()
        if idx > old_idx:
            direction = "right_to_left"
        else:
            direction = "left_to_right"

        self.slideInWidget(self.widget(idx), direction)

    def slideInWidget(self, new_widget, direction):
        if self.m_active:
            return

        self.m_active = True
        width = self.frameRect().width()
        height = self.frameRect().height()

        offset_x = width if direction == "right_to_left" else -width
        
        curr_widget = self.currentWidget()
        next_widget = new_widget

        next_widget.setGeometry(0, 0, width, height)
        next_widget.move(offset_x, 0)
        next_widget.show()
        next_widget.raise_()

        next_widget.raise_()
        
        # Add opacity effect for incoming widget
        self.eff = QGraphicsOpacityEffect(next_widget)
        next_widget.setGraphicsEffect(self.eff)

        # Animations
        self.anim_next = QPropertyAnimation(next_widget, b"pos")
        self.anim_next.setDuration(self.m_duration)
        self.anim_next.setEasingCurve(self.m_easing)
        self.anim_next.setStartValue(QPoint(offset_x, 0))
        self.anim_next.setEndValue(QPoint(0, 0))

        self.anim_curr = QPropertyAnimation(curr_widget, b"pos")
        self.anim_curr.setDuration(self.m_duration)
        self.anim_curr.setEasingCurve(self.m_easing)
        self.anim_curr.setStartValue(QPoint(0, 0))
        self.anim_curr.setEndValue(QPoint(-offset_x, 0))
        
        self.anim_fade = QPropertyAnimation(self.eff, b"opacity")
        self.anim_fade.setDuration(self.m_duration)
        self.anim_fade.setEasingCurve(self.m_easing)
        self.anim_fade.setStartValue(0.0)
        self.anim_fade.setEndValue(1.0)

        self.anims = QParallelAnimationGroup()
        self.anims.addAnimation(self.anim_next)
        self.anims.addAnimation(self.anim_curr)
        self.anims.addAnimation(self.anim_fade)
        self.anims.finished.connect(lambda: self._on_finished(idx_to_set=self.indexOf(new_widget), target=next_widget))
        self.anims.start()

    def _on_finished(self, idx_to_set, target):
        self.setCurrentIndex(idx_to_set)
        target.setGraphicsEffect(None) # Cleanup
        self.m_active = False
        self.animation_finished.emit(idx_to_set)

# --- Shared UI Styling ---

class TitleBar(QWidget):
    def __init__(self, parent):
        super().__init__(parent)
        self.setObjectName("title_bar")
        self.setFixedHeight(40)
        
        layout = QHBoxLayout(self)
        layout.setContentsMargins(15, 0, 5, 0)
        layout.setSpacing(5)
        
        # Logo and Title
        self.logo = QLabel()
        logo_pix = QPixmap(get_resource_path("assets/logo_premium.png"))
        if not logo_pix.isNull():
            self.logo.setPixmap(logo_pix.scaled(20, 20, Qt.KeepAspectRatio, Qt.SmoothTransformation))
        layout.addWidget(self.logo)
        
        self.title = QLabel("EUI OMR | Egypt University of Informatics")
        self.title.setStyleSheet(f"color: {COLORS['text_primary']}; font-weight: 600; font-size: 12px;")
        layout.addWidget(self.title)
        
        layout.addStretch()
        
        # Undo/Redo Controls
        self.btn_undo = QPushButton("↩")
        self.btn_undo.setObjectName("win_control")
        self.btn_undo.setFixedSize(40, 40)
        self.btn_undo.setToolTip("Undo (Ctrl+Z)")
        self.btn_undo.clicked.connect(parent._undo)
        self.btn_undo.setEnabled(False)
        layout.addWidget(self.btn_undo)
        
        self.btn_redo = QPushButton("↪")
        self.btn_redo.setObjectName("win_control")
        self.btn_redo.setFixedSize(40, 40)
        self.btn_redo.setToolTip("Redo (Ctrl+Y)")
        self.btn_redo.clicked.connect(parent._redo)
        self.btn_redo.setEnabled(False)
        layout.addWidget(self.btn_redo)
        
        layout.addSpacing(10)
        
        # Window Controls
        self.btn_min = QPushButton("−")
        self.btn_min.setObjectName("win_control")
        self.btn_min.setFixedSize(40, 40)
        self.btn_min.clicked.connect(parent.showMinimized)
        layout.addWidget(self.btn_min)
        
        self.btn_close = QPushButton("✕")
        self.btn_close.setObjectName("win_control")
        self.btn_close.setProperty("cssId", "close_btn") # For special hover
        self.btn_close.setFixedSize(40, 40)
        self.btn_close.clicked.connect(parent.close)
        # Custom QSS for close button hover
        self.btn_close.setStyleSheet("QPushButton:hover { background-color: #EF4444; color: white; }")
        layout.addWidget(self.btn_close)

    def mousePressEvent(self, event):
        self.parent().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        self.parent().mouseMoveEvent(event)


class SidebarItem(QPushButton):
    def __init__(self, icon_text, label, index, parent=None):
        super().__init__(parent)
        self.setObjectName("nav_item")
        self.setCheckable(True)
        self.index = index
        self.label_text = label
        self.icon_text = icon_text # Emojis or placeholders for now
        self.setFixedHeight(44)
        self.setCursor(Qt.PointingHandCursor)
        self.setText(f" {icon_text}   {label}")

    def setCollapsed(self, collapsed):
        if collapsed:
            self.setText(f" {self.icon_text}")
            self.setToolTip(self.label_text)
        else:
            self.setText(f" {self.icon_text}   {self.label_text}")
            self.setToolTip("")


class Sidebar(QWidget):
    nav_clicked = Signal(int)

    def __init__(self, parent):
        super().__init__(parent)
        self.setObjectName("sidebar")
        self.setFixedWidth(240)
        self.is_collapsed = False
        
        self.layout = QVBoxLayout(self)
        self.layout.setContentsMargins(12, 20, 12, 20)
        self.layout.setSpacing(8)
        
        # Toggle Button at top
        self.btn_toggle = QPushButton(" ☰ ")
        self.btn_toggle.setObjectName("nav_item")
        self.btn_toggle.setFixedSize(40, 40)
        self.btn_toggle.clicked.connect(self.toggle)
        self.layout.addWidget(self.btn_toggle)
        self.layout.addSpacing(20)
        
        # Nav Items
        self.items = []
        self._add_item("🏠", "Home", 0)
        self._add_item("⚙️", "Session Setup", 1)
        self._add_item("📊", "Dashboard", 2)
        
        self.layout.addStretch()
        
        # Bottom info
        self.lbl_uni = QLabel("Egypt University of Informatics")
        self.lbl_uni.setStyleSheet(f"color: {COLORS['text_dim']}; font-size: 10px;")
        self.lbl_uni.setAlignment(Qt.AlignCenter)
        self.lbl_uni.setWordWrap(True)
        self.layout.addWidget(self.lbl_uni)

        self.lbl_version = QLabel("v2.1.0")
        self.lbl_version.setAlignment(Qt.AlignCenter)
        self.layout.addWidget(self.lbl_version)

    def _add_item(self, icon, label, index):
        item = SidebarItem(icon, label, index, self)
        item.clicked.connect(lambda: self.nav_clicked.emit(index))
        self.layout.addWidget(item)
        self.items.append(item)
        if index == 0:
            item.setChecked(True)
            item.setProperty("active", "true")

    def toggle(self):
        new_width = 70 if not self.is_collapsed else 240
        self.anim = QPropertyAnimation(self, b"minimumWidth")
        self.anim.setDuration(250)
        self.anim.setStartValue(self.width())
        self.anim.setEndValue(new_width)
        self.anim.setEasingCurve(QEasingCurve.OutCubic)
        
        self.anim2 = QPropertyAnimation(self, b"maximumWidth")
        self.anim2.setDuration(250)
        self.anim2.setStartValue(self.width())
        self.anim2.setEndValue(new_width)
        self.anim2.setEasingCurve(QEasingCurve.OutCubic)
        
        self.anim.start()
        self.anim2.start()
        
        self.is_collapsed = not self.is_collapsed
        for item in self.items:
            item.setCollapsed(self.is_collapsed)
        
        if self.is_collapsed:
            self.lbl_version.setText("v2.1")
            self.lbl_uni.setText("EUI")
        else:
            self.lbl_version.setText("v2.1.0")
            self.lbl_uni.setText("Egypt University of Informatics")

    def set_active(self, index):
        for i, item in enumerate(self.items):
            active = (i == index)
            item.setChecked(active)
            item.setProperty("active", "true" if active else "false")
            item.style().unpolish(item)
            item.style().polish(item)

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
            
        # 1. Initialize Transient In-memory crops (Loaded only for this modal session)
        self.temp_id_crop = None
        self.temp_ver_crop = None
        self.temp_sig_crop = None
        self.temp_q_crops = {}
        self.error_keys = [] 
        
        # 2. Load images into transients BEFORE building panels
        self._ensure_images_loaded()
        
        main_layout = QHBoxLayout(self)
        main_layout.setContentsMargins(24, 24, 24, 24)
        main_layout.setSpacing(24)
        
        # --- LEFT PANE ---
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(16)
        
        lbl_list = QLabel("Detected Errors")
        lbl_list.setProperty("cssClass", "header")
        left_layout.addWidget(lbl_list)
        
        self.list_widget = QListWidget()
        self.list_widget.currentRowChanged.connect(self._on_list_selected)
        left_layout.addWidget(self.list_widget)
        
        # Action Buttons
        button_group = QVBoxLayout()
        button_group.setSpacing(8)
        
        self.btn_save = QPushButton("Accept Changes")
        self.btn_save.setProperty("cssClass", "primary")
        self.btn_save.setFixedHeight(40)
        self.btn_save.setEnabled(False) 
        self.btn_save.clicked.connect(self.accept)
        button_group.addWidget(self.btn_save)
        
        self.btn_ignore = QPushButton("Discard Changes")
        self.btn_ignore.setFixedHeight(40)
        self.btn_ignore.clicked.connect(self.reject)
        button_group.addWidget(self.btn_ignore)
        left_layout.addLayout(button_group)
        
        main_layout.addWidget(left_widget, stretch=1)
        
        # --- RIGHT PANE (Scrollable) ---
        container_right = QWidget()
        container_right.setObjectName("card")
        right_layout = QVBoxLayout(container_right)
        right_layout.setContentsMargins(0, 0, 0, 0)
        
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setFrameShape(QScrollArea.NoFrame)
        
        self.stack = QStackedWidget()
        self.scroll.setWidget(self.stack)
        right_layout.addWidget(self.scroll)
        
        main_layout.addWidget(container_right, stretch=3)
        
        self.id_input: QLineEdit | None = None
        self.ver_input: QComboBox | None = None
        self.q_inputs: dict[int, QComboBox] = {}
        
        self._populate_errors()

    def _ensure_images_loaded(self):
        """Loads images from disk cache if paths exist, otherwise falls back to PDF extraction."""
        # 1. Try Disk Cache First
        if self.res.id_crop_path:
            # We assume paths are relative to the project directory
            proj_dir = Path(self.pdf_path).parent.parent
            crops_dir = proj_dir / "crops"
            
            if (crops_dir / self.res.id_crop_path).exists():
                self.temp_id_crop = cv2.imread(str(crops_dir / self.res.id_crop_path))
                
            if self.res.version_crop_path and (crops_dir / self.res.version_crop_path).exists():
                self.temp_ver_crop = cv2.imread(str(crops_dir / self.res.version_crop_path))
                
            if self.res.signature_crop_path and (crops_dir / self.res.signature_crop_path).exists():
                self.temp_sig_crop = cv2.imread(str(crops_dir / self.res.signature_crop_path))
                
            for q_num, p_name in self.res.question_crop_paths.items():
                if (crops_dir / p_name).exists():
                    self.temp_q_crops[q_num] = cv2.imread(str(crops_dir / p_name))
            
            # If we successfully loaded from disk, we can skip the expensive PDF extraction
            if self.temp_id_crop is not None:
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
                
                self.temp_id_crop = crops.get("id_crop")
                self.temp_ver_crop = crops.get("version_crop")
                self.temp_sig_crop = crops.get("signature_crop")
                self.temp_q_crops = crops.get("question_crops")
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
        l.setContentsMargins(24, 24, 24, 24)
        l.setSpacing(16)
        
        lbl_title = QLabel("Identity Verification")
        lbl_title.setProperty("cssClass", "header")
        l.addWidget(lbl_title)
        
        l.addWidget(QLabel("Student's Handwritten Name:"))
        lbl_sig = QLabel()
        if self.temp_sig_crop is not None:
            pix = cv2_to_qpixmap(self.temp_sig_crop)
            scaled = pix.scaled(pix.width() * 0.75, pix.height() * 0.75, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            lbl_sig.setPixmap(scaled)
        l.addWidget(lbl_sig)
        
        l.addWidget(QLabel("ID Box Extract:"))
        self.lbl_id = QLabel()
        if self.temp_id_crop is not None:
            pix = cv2_to_qpixmap(self.temp_id_crop)
            scaled = pix.scaled(pix.width() * 0.75, pix.height() * 0.75, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            self.lbl_id.setPixmap(scaled)
        l.addWidget(self.lbl_id)
        
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
        if self.temp_ver_crop is not None:
            pix = cv2_to_qpixmap(self.temp_ver_crop)
            scaled = pix.scaled(pix.width() * 0.75, pix.height() * 0.75, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            lbl_v.setPixmap(scaled)
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
        if q_num in self.temp_q_crops:
            pix = cv2_to_qpixmap(self.temp_q_crops[q_num])
            scaled = pix.scaled(pix.width() * 0.75, pix.height() * 0.75, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            lbl_img.setPixmap(scaled)
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
        
        btns = QHBoxLayout()
        btns.setContentsMargins(0, 10, 0, 0)
        btn_ok = QPushButton("Apply and Save Keys")
        btn_ok.setProperty("cssClass", "primary")
        btn_ok.setFixedHeight(40)
        btn_ok.clicked.connect(self.accept)
        btns.addWidget(btn_ok)
        layout.addLayout(btns)
        
        # Apply the shared stylesheet
        self.setStyleSheet(APP_QSS)
        self.tabs.setStyleSheet(f"""
            QTabWidget::pane {{ border: 1px solid {COLORS['border']}; top: -1px; background: {COLORS['surface']}; }}
            QTabBar::tab {{ background: {COLORS['surface']}; color: {COLORS['text_secondary']}; padding: 10px 20px; border: 1px solid {COLORS['border']}; border-bottom: none; border-top-left-radius: 8px; border-top-right-radius: 8px; }}
            QTabBar::tab:selected {{ background: {COLORS['background']}; color: {COLORS['accent']}; font-weight: bold; }}
            QCheckBox {{ color: {COLORS['text_primary']}; spacing: 8px; background: transparent; }}
            QCheckBox::indicator {{ 
                width: 18px; height: 18px; 
                background-color: {COLORS['background']}; 
                border: 1px solid {COLORS['border']}; 
                border-radius: 4px; 
            }}
            QCheckBox::indicator:hover {{ border-color: {COLORS['accent']}; }}
            QCheckBox::indicator:checked {{ 
                background-color: {COLORS['accent']}; 
                border-color: {COLORS['accent']};
            }}
            QScrollArea, QScrollArea > QWidget > QWidget {{ 
                background-color: {COLORS['surface']}; 
                border: none;
            }}
        """)

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
        self.setMinimumSize(1200, 800)
        self.setStyleSheet(APP_QSS)
        
        # History System
        self.undo_stack: list[list[GradingResult]] = []
        self.redo_stack: list[list[GradingResult]] = []
        
        # Shortcuts
        QShortcut(QKeySequence("Ctrl+Z"), self, self._undo)
        QShortcut(QKeySequence("Ctrl+Y"), self, self._redo)
        
        # --- NEW STRUCTURAL LAYOUT ---
        self.central = QWidget()
        self.setCentralWidget(self.central)
        self.main_layout = QVBoxLayout(self.central)
        self.main_layout.setContentsMargins(0, 0, 0, 0)
        self.main_layout.setSpacing(0)
        
        # 1. Custom Title Bar
        self.title_bar = TitleBar(self)
        self.btn_undo = self.title_bar.btn_undo
        self.btn_redo = self.title_bar.btn_redo
        self.main_layout.addWidget(self.title_bar)
        
        # 2. Horizontal Container for Sidebar + Content
        self.body_container = QWidget()
        self.body_layout = QHBoxLayout(self.body_container)
        self.body_layout.setContentsMargins(0, 0, 0, 0)
        self.body_layout.setSpacing(0)
        
        # 3. Sidebar
        self.sidebar = Sidebar(self)
        self.sidebar.nav_clicked.connect(self._on_nav_clicked)
        self.body_layout.addWidget(self.sidebar)
        
        # 4. Main Stacked Content
        self.stack = SlidingStackedWidget()
        self.body_layout.addWidget(self.stack)
        
        self.main_layout.addWidget(self.body_container)
        
        # Persistent Settings Memory
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

    def _on_nav_clicked(self, index):
        if not self.pm and index > 0:
            QMessageBox.information(self, "No Project", "Please create or load a project first.")
            self.sidebar.set_active(0)
            return
        self.stack.slideInIdx(index)
        self.sidebar.set_active(index)
        # Immediately trigger entrance if it's the first time or if we want faster feedback
        # but better to wait for slide to finish for the "cascading" look.

    def update_header_project(self, name: str):
        # Update sidebar version label or sidebar title instead of header
        pass
    # --- Window Dragging Logic (Frameless Support) ---
    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            # Check if click is on the title bar area
            if self.title_bar.underMouse():
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
        l.setContentsMargins(80, 40, 80, 40)
        l.setSpacing(40)
        
        # Hero Section
        hero = QWidget()
        hero_layout = QVBoxLayout(hero)
        hero_layout.setSpacing(10)
        
        title = QLabel("Welcome to EUI OMR")
        title.setProperty("cssClass", "title")
        title.setAlignment(Qt.AlignCenter)
        hero_layout.addWidget(title)
        
        desc = QLabel("Industrial-grade optical mark recognition for academic assessments.")
        desc.setProperty("cssClass", "body")
        desc.setAlignment(Qt.AlignCenter)
        hero_layout.addWidget(desc)
        
        l.addWidget(hero)
        
        # Action Cards
        actions_container = QWidget()
        actions_layout = QHBoxLayout(actions_container)
        actions_layout.setSpacing(24)
        
        # New Project Card
        card_new = QWidget()
        card_new.setObjectName("card")
        nl = QVBoxLayout(card_new)
        nl.setContentsMargins(32, 32, 32, 32)
        nl.setSpacing(16)
        
        icon_new = QLabel("📄")
        icon_new.setStyleSheet("font-size: 32px;")
        icon_new.setAlignment(Qt.AlignCenter)
        nl.addWidget(icon_new)
        
        nl.addWidget(QLabel("<b>New Project</b>"), alignment=Qt.AlignCenter)
        nl.addWidget(QLabel("Initialize a fresh grading session\nfrom PDF scans."), alignment=Qt.AlignCenter)
        
        btn_new = QPushButton("Create Project")
        btn_new.setProperty("cssClass", "primary")
        btn_new.setCursor(Qt.PointingHandCursor)
        btn_new.clicked.connect(self._create_project)
        nl.addWidget(btn_new)
        actions_layout.addWidget(card_new)
        
        # Load Project Card
        card_load = QWidget()
        card_load.setObjectName("card")
        ll = QVBoxLayout(card_load)
        ll.setContentsMargins(32, 32, 32, 32)
        ll.setSpacing(16)
        
        icon_load = QLabel("📂")
        icon_load.setStyleSheet("font-size: 32px;")
        icon_load.setAlignment(Qt.AlignCenter)
        ll.addWidget(icon_load)
        
        ll.addWidget(QLabel("<b>Open Project</b>"), alignment=Qt.AlignCenter)
        ll.addWidget(QLabel("Resume an existing grading session\nfrom a saved state."), alignment=Qt.AlignCenter)
        
        btn_load = QPushButton("Load State")
        btn_load.setCursor(Qt.PointingHandCursor)
        btn_load.clicked.connect(self._load_project)
        ll.addWidget(btn_load)
        actions_layout.addWidget(card_load)
        
        l.addWidget(actions_container)
        
        # Recent Projects
        if self.recent_projects:
            recent_container = QWidget()
            recent_container.setObjectName("card")
            recent_container.setFixedWidth(600)
            rl = QVBoxLayout(recent_container)
            rl.setContentsMargins(24, 24, 24, 24)
            
            rl.addWidget(QLabel("<b>🕒 Recent Projects</b>"))
            self.recent_list = QListWidget()
            self.recent_list.setFixedHeight(140)
            self.recent_list.addItems([str(Path(p).name) for p in self.recent_projects])
            self.recent_list.itemDoubleClicked.connect(self._on_recent_double_clicked)
            rl.addWidget(self.recent_list)
            
            hint = QLabel("Double-click to quick-load project state.")
            hint.setStyleSheet(f"color: {COLORS['text_secondary']}; font-size: 11px;")
            rl.addWidget(hint, alignment=Qt.AlignRight)
            
            l.addWidget(recent_container, alignment=Qt.AlignCenter)
        
        l.addStretch()
        
        tip = QLabel("✨ Tip: Drag and drop a project folder anywhere to load it.")
        tip.setStyleSheet(f"color: {COLORS['text_secondary']}; font-style: italic;")
        l.addWidget(tip, alignment=Qt.AlignCenter)
        
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

        
    def build_setup_view(self):
        self.v_setup = QWidget()
        self.v_setup.setObjectName("setup_container")
        
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(self.v_setup)
        
        # Main Layout
        l = QVBoxLayout(self.v_setup)
        l.setContentsMargins(40, 20, 40, 30)
        l.setSpacing(16)
        
        header_sec = QWidget()
        hl = QVBoxLayout(header_sec)
        hl.setContentsMargins(0, 0, 0, 0)
        hl.setSpacing(4)
        
        title = QLabel("Session Setup")
        title.setProperty("cssClass", "title")
        hl.addWidget(title)
        
        desc = QLabel("Configure your roster, answer keys, and scan source.")
        desc.setProperty("cssClass", "body")
        hl.addWidget(desc)
        l.addWidget(header_sec)
        
        # 1. Roster & Data Mapping
        card_roster = QWidget()
        card_roster.setObjectName("card")
        rl = QVBoxLayout(card_roster)
        rl.setContentsMargins(24, 16, 24, 16)
        rl.setSpacing(12)
        
        st1 = QLabel("1. Roster & Data Mapping")
        st1.setProperty("cssClass", "header")
        rl.addWidget(st1)
        
        h_excel = QHBoxLayout()
        h_excel.setSpacing(16)
        self.btn_excel = QPushButton("Import Student Roster")
        self.btn_excel.setCursor(Qt.PointingHandCursor)
        self.btn_excel.clicked.connect(self._import_excel)
        h_excel.addWidget(self.btn_excel)
        
        self.lbl_excel = QLabel("No file selected")
        self.lbl_excel.setStyleSheet(f"color: {COLORS['accent']}; font-weight: 600;")
        h_excel.addWidget(self.lbl_excel)
        h_excel.addStretch()
        rl.addLayout(h_excel)
        
        grid = QGridLayout()
        grid.setSpacing(16)
        grid.addWidget(QLabel("ID Column:"), 0, 0)
        self.cb_id = QComboBox()
        self.cb_id.setMinimumWidth(200)
        grid.addWidget(self.cb_id, 0, 1)
        
        grid.addItem(QSpacerItem(40, 0, QSizePolicy.Expanding, QSizePolicy.Minimum), 0, 2)
        
        grid.addWidget(QLabel("Grade Output:"), 0, 3)
        self.cb_out = QComboBox()
        self.cb_out.setMinimumWidth(200)
        grid.addWidget(self.cb_out, 0, 4)
        rl.addLayout(grid)
        l.addWidget(card_roster)
        
        # 2. Exam Configuration
        card_exam = QWidget()
        card_exam.setObjectName("card")
        el = QVBoxLayout(card_exam)
        el.setContentsMargins(24, 16, 24, 16)
        el.setSpacing(12)
        
        st2 = QLabel("2. Exam Configuration")
        st2.setProperty("cssClass", "header")
        el.addWidget(st2)
        
        h_config = QHBoxLayout()
        h_config.setSpacing(16)
        h_config.addWidget(QLabel("Number of Questions:"))
        self.spin_q = QSpinBox()
        self.spin_q.setRange(1, 100); self.spin_q.setValue(60)
        self.spin_q.setFixedWidth(80)
        h_config.addWidget(self.spin_q)
        
        h_config.addSpacing(32)
        
        self.btn_keys = QPushButton("Manage Answer Keys")
        self.btn_keys.setCursor(Qt.PointingHandCursor)
        self.btn_keys.clicked.connect(self._manage_keys)
        h_config.addWidget(self.btn_keys)
        h_config.addStretch()
        el.addLayout(h_config)
        l.addWidget(card_exam)
        
        # 3. Scan Source
        card_scan = QWidget()
        card_scan.setObjectName("card")
        sl = QVBoxLayout(card_scan)
        sl.setContentsMargins(24, 16, 24, 16)
        sl.setSpacing(12)
        
        st3 = QLabel("3. Scan Source (PDF)")
        st3.setProperty("cssClass", "header")
        sl.addWidget(st3)
        
        h_pdf = QHBoxLayout()
        h_pdf.setSpacing(16)
        self.btn_select_pdf = QPushButton("Select Student Scans")
        self.btn_select_pdf.setCursor(Qt.PointingHandCursor)
        self.btn_select_pdf.clicked.connect(self._select_pdf)
        h_pdf.addWidget(self.btn_select_pdf)
        
        self.lbl_setup_pdf = QLabel("No PDF selected")
        self.lbl_setup_pdf.setStyleSheet(f"color: {COLORS['accent']}; font-weight: 600;")
        h_pdf.addWidget(self.lbl_setup_pdf)
        h_pdf.addStretch()
        sl.addLayout(h_pdf)
        
        hint = QLabel("Engine will automatically calibrate scan quality before processing.")
        hint.setStyleSheet(f"color: {COLORS['text_secondary']}; font-style: italic;")
        sl.addWidget(hint)
        l.addWidget(card_scan)
        
        l.addStretch()
        
        # Footer Actions
        footer = QWidget()
        fl = QHBoxLayout(footer)
        fl.setContentsMargins(0, 0, 0, 0)
        
        self.btn_resume = QPushButton("View Current Dashboard")
        self.btn_resume.setProperty("cssClass", "special")
        self.btn_resume.setCursor(Qt.PointingHandCursor)
        self.btn_resume.clicked.connect(self._resume_to_dashboard)
        self.btn_resume.setVisible(False)
        self.btn_resume.setFixedWidth(200)
        fl.addWidget(self.btn_resume, alignment=Qt.AlignVCenter)
        
        fl.addStretch()
        
        self.btn_start = QPushButton("Start Grading Process")
        self.btn_start.setProperty("cssClass", "primary")
        self.btn_start.setFixedHeight(48)
        self.btn_start.setFixedWidth(220)
        self.btn_start.setCursor(Qt.PointingHandCursor)
        self.btn_start.clicked.connect(self._save_setup_and_continue)
        fl.addWidget(self.btn_start)
        l.addWidget(footer)
        
        self.stack.addWidget(scroll)

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
            self._validate_setup_ready()
            
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
        self.stack.slideInIdx(1)
        self.sidebar.set_active(1)
        
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
        self.stack.slideInIdx(2)
        self.sidebar.set_active(2)

    def _save_setup_and_continue(self):
        if not self.pm or not self.pm.excel_roster_path or not self.pm.student_pdf_path:
            QMessageBox.warning(self, "Setup Incomplete", "Please ensure both an Excel roster and Student PDF scans are selected before starting.")
            return
            
        self.pm.student_id_col = self.cb_id.currentText()
        self.pm.grade_output_col = self.cb_out.currentText()
        self.pm.question_count = self.spin_q.value()
        self.pm.save_state()
        
        # Navigate to setup
        self.stack.slideInIdx(2)
        self.sidebar.set_active(2, easing=QEasingCurve.OutCubic)
        
        # Phase 1: Silent Automated Calibration (5 Random Pages)
        QTimer.singleShot(100, self._start_silent_calibration)

    def _start_silent_calibration(self):
        """Phase 1: Automatically calibrate engine sensitivity before starting the batch."""
        import random
        import fitz
        
        self.lbl_pdf.setText("Phase 1/2: Optimizing for scan quality (5 random samples)...")
        self.btn_export.setEnabled(False)
        
        try:
            doc = fitz.open(self.active_pdf_path)
            num_pages = len(doc)
            num_samples = min(5, num_pages)
            
            # Select 5 random indices distributed across the PDF
            indices = random.sample(range(num_pages), num_samples)
            doc.close()
            
            self.tune_worker = AutoTuneWorker(
                self.active_pdf_path, 
                self.config_path, 
                self.spin_q.value(),
                sample_indices=indices # We'll update the worker to accept specific indices
            )
            self.tune_worker.finished.connect(self._on_silent_tune_finished)
            self.tune_worker.error.connect(lambda e: self._on_error("Calibration Error", e))
            self.tune_worker.start()
        except Exception as e:
            self._on_error("Calibration Error", f"Failed to access PDF for calibration: {e}")

    def _on_silent_tune_finished(self, best_sens):
        """Phase 2: Transition seamlessly to actual grading."""
        self.pm.mark_sensitivity = best_sens
        self.pm.save_state()
        
        if self.pm.logger:
            self.pm.logger.info(f"Silent auto-calibration complete. Best sensitivity: {best_sens}")
            
        self.lbl_pdf.setText("Phase 2/2: Processing Student Sheets...")
        self._start_grading()

    # ---------------------------------------------------------
    # VIEW 3: GRADING / PROCESSING
    # ---------------------------------------------------------
    def build_processing_view(self):
        self.v_proc = QWidget()
        l = QVBoxLayout(self.v_proc)
        l.setContentsMargins(32, 32, 32, 32)
        l.setSpacing(24)
        
        # Header Info
        header_area = QWidget()
        hl = QHBoxLayout(header_area)
        hl.setContentsMargins(0, 0, 0, 0)
        
        info_sec = QVBoxLayout()
        self.lbl_pdf = QLabel("Ready to process scans")
        self.lbl_pdf.setObjectName("header_title")
        self.lbl_pdf.setStyleSheet(f"font-size: 18px; color: {COLORS['text_primary']};")
        info_sec.addWidget(self.lbl_pdf)
        
        self.lbl_session = QLabel("Please Review the Conditions before Exporting") # Placeholder or session info
        self.lbl_session.setStyleSheet(f"color: {COLORS['text_secondary']}; font-size: 12px;")
        info_sec.addWidget(self.lbl_session)
        hl.addLayout(info_sec)
        
        hl.addStretch()
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setFixedWidth(300)
        self.progress_bar.hide()
        hl.addWidget(self.progress_bar)
        l.addWidget(header_area)
        
        # Dashboard Content
        dashboard_content = QHBoxLayout()
        dashboard_content.setSpacing(24)
        
        # Table Container
        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["Page #", "Extracted ID", "Model", "Condition"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.horizontalHeader().setDefaultAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.table.setAlternatingRowColors(True)
        self.table.setStyleSheet(f"QTableWidget {{ alternate-background-color: {COLORS['background']}; }}")
        self.table.itemChanged.connect(self._on_table_item_changed)
        self.table.cellDoubleClicked.connect(self._on_table_double_clicked)
        dashboard_content.addWidget(self.table, stretch=1)
        
        # Sidebar Summary
        sidebar_summary = QWidget()
        sidebar_summary.setObjectName("card")
        sidebar_summary.setFixedWidth(280)
        sl = QVBoxLayout(sidebar_summary)
        sl.setContentsMargins(24, 24, 24, 24)
        sl.setSpacing(20)
        
        sl.addWidget(QLabel("<b>Grading Summary</b>"))
        
        stats_box = QVBoxLayout()
        stats_box.setSpacing(12)
        self.stat_total = QLabel("Total Sheets: 0")
        self.stat_success = QLabel("Resolved: 0")
        self.stat_errors = QLabel("Review Needed: 0")
        
        for lbl in [self.stat_total, self.stat_success, self.stat_errors]:
            lbl.setStyleSheet(f"color: {COLORS['text_primary']}; font-size: 14px;")
            stats_box.addWidget(lbl)
        sl.addLayout(stats_box)
        
        sl.addStretch()
        
        self.btn_export = QPushButton("Export Final Marks")
        self.btn_export.setProperty("cssClass", "primary")
        self.btn_export.setFixedHeight(44)
        self.btn_export.setCursor(Qt.PointingHandCursor)
        self.btn_export.setEnabled(False)
        self.btn_export.clicked.connect(self._export_to_excel)
        sl.addWidget(self.btn_export)
        
        btn_settings = QPushButton("Back to Setup")
        btn_settings.setProperty("cssClass", "ghost")
        btn_settings.setCursor(Qt.PointingHandCursor)
        btn_settings.clicked.connect(lambda: self.stack.slideInIdx(1))
        sl.addWidget(btn_settings)
        
        dashboard_content.addWidget(sidebar_summary)
        l.addLayout(dashboard_content)
        
        self.stack.addWidget(self.v_proc)
        
        self.results_data: list[GradingResult] = [] 
        
        # Load existing results if any
        if self.pm and self.pm.last_results:
            self.results_data = sorted(self.pm.last_results, key=lambda x: x.page_number)
            for res in self.results_data:
                self._update_table_row(res)
            self.btn_export.setEnabled(True)


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
            self._validate_setup_ready()


    def _stop_worker(self):
        """Safely stops and disposes of current workers to prevent resource leaks."""
        if hasattr(self, 'worker') and self.worker and self.worker.isRunning():
            try:
                self.worker.terminate()
                self.worker.wait(500)
            except: pass
        self.worker = None

    def _start_grading(self):
        if not self.active_pdf_path: return
        
        self._stop_worker()
        
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
            status.setForeground(QColor(COLORS['accent']))
        elif res.id_error or res.version_error or len(res.question_errors) > 0:
            status = QTableWidgetItem("⚠️ Needs Review")
            status.setForeground(QColor(COLORS['error']))
        else:
            status = QTableWidgetItem("✅ Success")
            status.setForeground(QColor(COLORS['success']))
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
        
        self.stat_total.setText(f"Total Sheets: {total}")
        self.stat_success.setText(f"Resolved: {success}")
        self.stat_errors.setText(f"Review Needed: {errors}")
        
        # Error text highlighting
        if errors > 0:
            self.stat_errors.setStyleSheet(f"color: {COLORS['error']}; font-weight: 600;")
        else:
            self.stat_errors.setStyleSheet(f"color: {COLORS['success']};")
            
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


