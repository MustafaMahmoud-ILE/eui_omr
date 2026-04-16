"""Modern SaaS Design System for EUI OMR (Inspired by Linear, Vercel, Stripe)."""

# Color Palette
COLORS = {
    "background": "#121212",
    "surface": "#1E1E1E",
    "surface_light": "#252525",
    "accent": "#3B82F6",
    "accent_hover": "#60A5FA",
    "accent_pressed": "#2563EB",
    "text_primary": "#FFFFFF",
    "text_secondary": "#B0B0B0",
    "text_dim": "#71717A",
    "border": "#2A2A2A",
    "border_focus": "#3B82F6",
    "error": "#EF4444",
    "success": "#10B981",
    "warning": "#F59E0B"
}

# Design Tokens
STYLE_TOKENS = {
    "radius": "8px",
    "font_family": "'Inter', 'Segoe UI', system-ui, sans-serif",
    "spacing_xs": "4px",
    "spacing_sm": "8px",
    "spacing_md": "16px",
    "spacing_lg": "24px",
    "spacing_xl": "32px",
}

APP_QSS = f"""
    /* Global Styles */
    QMainWindow, QDialog, QStackedWidget {{ 
        background-color: {COLORS['background']}; 
        color: {COLORS['text_primary']};
        font-family: {STYLE_TOKENS['font_family']};
    }}

    QWidget {{
        color: {COLORS['text_primary']};
        font-size: 13px;
    }}

    /* Scroll Area */
    QScrollArea, QScrollArea > QWidget > QWidget {{
        background-color: {COLORS['background']};
        border: none;
    }}
    QScrollArea QWidget {{ background-color: transparent; }}

    /* Typography */
    QLabel {{ color: {COLORS['text_secondary']}; }}
    QLabel[cssClass="title"] {{
        font-size: 24px;
        font-weight: 800;
        color: {COLORS['text_primary']};
    }}
    QLabel[cssClass="header"] {{
        font-size: 16px;
        font-weight: 600;
        color: {COLORS['text_primary']};
    }}
    QLabel[cssClass="body"] {{
        font-size: 14px;
        color: {COLORS['text_secondary']};
    }}

    /* Cards */
    QWidget#card {{
        background-color: {COLORS['surface']};
        border: 1px solid {COLORS['border']};
        border-radius: {STYLE_TOKENS['radius']};
    }}

    /* Buttons */
    QPushButton {{
        background-color: {COLORS['surface']};
        border: 1px solid {COLORS['border']};
        border-radius: 6px;
        padding: 8px 16px;
        font-weight: 500;
        color: {COLORS['text_primary']};
    }}
    QPushButton:hover {{
        background-color: {COLORS['surface_light']};
        border-color: {COLORS['text_secondary']};
    }}
    QPushButton:pressed {{
        background-color: {COLORS['background']};
    }}
    QPushButton:disabled {{
        color: #4A4A4A;
        background-color: transparent;
        border: 1px solid #333333;
    }}

    QPushButton[cssClass="primary"] {{
        background-color: {COLORS['accent']};
        color: white;
        border: none;
    }}
    QPushButton[cssClass="primary"]:hover {{
        background-color: {COLORS['accent_hover']};
    }}
    QPushButton[cssClass="primary"]:pressed {{
        background-color: {COLORS['accent_pressed']};
    }}
    QPushButton[cssClass="primary"]:disabled {{
        background-color: #2D2D2D;
        color: #555555;
        border: 1px solid #333333;
    }}

    QPushButton[cssClass="ghost"] {{
        background-color: transparent;
        border: none;
        color: {COLORS['text_secondary']};
    }}
    QPushButton[cssClass="ghost"]:hover {{
        background-color: rgba(59, 130, 246, 0.1);
        color: {COLORS['accent']};
    }}

    QPushButton[cssClass="special"] {{
        background-color: transparent;
        border: 1px solid {COLORS['accent']};
        color: {COLORS['accent']};
        font-weight: 600;
        border-radius: 6px;
    }}
    QPushButton[cssClass="special"]:hover {{
        background-color: {COLORS['accent']};
        color: white;
    }}

    /* Inputs */
    QLineEdit, QComboBox, QSpinBox {{
        background-color: {COLORS['background']};
        border: 1px solid {COLORS['border']};
        border-radius: 6px;
        padding: 8px 12px;
        color: {COLORS['text_primary']};
    }}
    QLineEdit:focus, QComboBox:focus, QSpinBox:focus {{
        border-color: {COLORS['border_focus']};
    }}

    QComboBox::drop-down {{
        border: none;
        width: 30px;
    }}
    QComboBox::down-arrow {{
        image: none;
        width: 0px;
        height: 0px;
        border-left: 5px solid transparent;
        border-right: 5px solid transparent;
        border-top: 6px solid {COLORS['text_secondary']};
        subcontrol-origin: content;
        subcontrol-position: center;
        margin-right: 12px;
    }}
    QComboBox QAbstractItemView {{
        background-color: {COLORS['surface']};
        border: 1px solid {COLORS['border']};
        selection-background-color: {COLORS['accent']};
        outline: none;
        padding: 4px;
    }}


    /* Table */
    QTableWidget {{
        background-color: {COLORS['surface']};
        gridline-color: {COLORS['border']};
        border: 1px solid {COLORS['border']};
        border-radius: 8px;
        outline: none;
    }}
    QHeaderView::section {{
        background-color: #1A1A1A;
        color: {COLORS['text_secondary']};
        font-weight: 600;
        padding: 12px;
        border: none;
        border-bottom: 1px solid {COLORS['border']};
        border-right: 1px solid {COLORS['border']};
    }}
    QHeaderView::section:vertical {{
        padding: 4px 8px;
        font-size: 10px;
        font-weight: normal;
        border-right: 1px solid {COLORS['border']};
        border-bottom: 1px solid {COLORS['border']};
    }}
    QTableCornerButton::section {{
        background-color: #1A1A1A;
        border: none;
        border-bottom: 1px solid {COLORS['border']};
        border-right: 1px solid {COLORS['border']};
    }}
    QTableWidget::item {{
        padding: 12px;
    }}
    QTableWidget::item:selected {{
        background-color: rgba(59, 130, 246, 0.1);
        color: {COLORS['accent']};
    }}

    /* List Widget (Recent Projects) */
    QListWidget {{
        background-color: transparent;
        border: none;
        outline: none;
    }}
    QListWidget::item {{
        background-color: {COLORS['surface_light']};
        border: 1px solid {COLORS['border']};
        border-radius: 8px;
        padding: 12px 16px;
        margin-bottom: 8px;
        color: {COLORS['text_primary']};
    }}
    QListWidget::item:hover {{
        background-color: {COLORS['surface']};
        border-color: {COLORS['accent']};
    }}
    QListWidget::item:selected {{
        background-color: rgba(59, 130, 246, 0.1);
        color: {COLORS['accent']};
        border-color: {COLORS['accent']};
    }}

    /* Progress Bar */
    QProgressBar {{
        background-color: {COLORS['surface']};
        border: none;
        border-radius: 4px;
        height: 6px;
        text-align: center;
        color: transparent;
    }}
    QProgressBar::chunk {{
        background-color: {COLORS['accent']};
        border-radius: 4px;
    }}

    /* Sidebar Navigation */
    QWidget#sidebar {{
        background-color: {COLORS['surface']};
        border-right: 1px solid {COLORS['border']};
    }}
    QPushButton#nav_item {{
        background-color: transparent;
        border: none;
        border-radius: 6px;
        padding: 10px;
        text-align: left;
        color: {COLORS['text_secondary']};
        font-weight: 500;
    }}
    QPushButton#nav_item:hover {{
        background-color: {COLORS['surface_light']};
        color: {COLORS['text_primary']};
    }}
    QPushButton#nav_item[active="true"] {{
        background-color: rgba(59, 130, 246, 0.1);
        color: {COLORS['accent']};
    }}

    /* Title Bar */
    QWidget#title_bar {{
        background-color: {COLORS['surface']};
        border-bottom: 1px solid {COLORS['border']};
    }}
    QPushButton#win_control {{
        background-color: transparent;
        border: none;
        border-radius: 0px;
        color: {COLORS['text_secondary']};
    }}
    QPushButton#win_control:hover {{
        background-color: {COLORS['surface_light']};
    }}
    QPushButton#close_btn:hover {{
        background-color: {COLORS['error']};
        color: white;
    }}

    /* Modern SaaS ScrollBars */
    QScrollBar:vertical {{
        background: transparent;
        width: 10px;
        margin: 0px;
    }}
    QScrollBar::handle:vertical {{
        background: #333333;
        min-height: 20px;
        border-radius: 5px;
        margin: 2px;
    }}
    QScrollBar::handle:vertical:hover {{
        background: {COLORS['text_secondary']};
    }}
    QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
        height: 0px;
    }}
    QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {{
        background: none;
    }}

    QScrollBar:horizontal {{
        background: transparent;
        height: 10px;
        margin: 0px;
    }}
    QScrollBar::handle:horizontal {{
        background: #333333;
        min-width: 20px;
        border-radius: 5px;
        margin: 2px;
    }}
    QScrollBar::handle:horizontal:hover {{
        background: {COLORS['text_secondary']};
    }}
    QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{
        width: 0px;
    }}
    QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal {{
        background: none;
    }}
"""
