# Standard library module for interacting with the Python interpreter (used here for sys.exit())
import sys

# Data manipulation library; used to read Excel files into DataFrames and perform row/column operations
import pandas as pd

# PyQt5 widget classes for building the graphical user interface
from PyQt5.QtWidgets import (
    QApplication,       # Manages the GUI application's control flow and main settings
    QMainWindow,        # Provides a main application window with menu bar, status bar, etc.
    QWidget,            # Base class for all UI objects; used as a container for layouts
    QVBoxLayout,        # Arranges widgets vertically in a column
    QHBoxLayout,        # Arranges widgets horizontally in a row
    QLineEdit,          # Single-line text input field for the Excel sheet name
    QPushButton,        # Command buttons for SHOW, START UPDATE, CANCEL UPDATE, and SELECT ALL
    QTableWidget,       # Displays the device inventory in a tabular format with rows and columns
    QTableWidgetItem,   # Represents an individual cell within the QTableWidget
    QHeaderView,        # Manages the table's column and row headers (titles, resizing, etc.)
    QCheckBox,          # Toggleable checkbox for selecting/deselecting individual devices
    QLabel,             # Displays static text such as the app title and section headers
    QFrame,             # Container widget used for decorative separators and panel borders
    QAbstractItemView,  # Base class providing selection and interaction behaviors for item views
    QDialog,            # Base class for modal dialog windows (error dialogs, credentials popup)
    QMessageBox,        # Convenience class for displaying standard message boxes (warnings, info)
    QSizePolicy,        # Controls how a widget resizes relative to its layout container
    QPlainTextEdit,      # Multi-line plain text input area for displaying or editing large amounts of unformatted text
    QScrollArea,        # Scrollable container — used in PushCommandsDialog for the horizontal device tab bar
    QStackedWidget      # Container that shows one child widget at a time — used for phase switching in PushCommandsDialog
)

# PyQt5 core classes for non-GUI functionality and application fundamentals
from PyQt5.QtCore import (
    Qt,          # Contains enumerations for alignment, focus policies, keyboard modifiers, etc.
    QSize,       # Represents a two-dimensional size (width and height) for widgets and windows
    QTimer,      # Provides repetitive and single-shot timers for simulating async operations
    QThread,     # Enables multi-threading to keep UI responsive during long operations
    QObject,     # Base class for all Qt objects; required for signal/slot mechanism
    pyqtSignal   # Decorator for defining custom signals that can be emitted between threads
)

# PyQt5 GUI classes for visual customization and appearance
from PyQt5.QtGui import (
    QFont,       # Defines font family, size, weight, and style for text rendering
    QColor,      # Represents RGB color values used in stylesheets and custom painting
    QPalette,    # Manages color groups (active, inactive, disabled) for widget states
    QTextCursor  # Controls selection, insertion, deletion, and navigation within QPlainTextEdit or QTextEdit
)

# Custom module containing Excel file operations and validation functions
import excel_and_data_ops

# Custom module containing CLI operations to push commands to devices and retrieve data
import device_cli_ops

# Standard library module for writing structured log messages (INFO, ERROR, etc.)
import logging

# Logging handler that auto-rotates the log file when it reaches a defined size limit
from logging.handlers import RotatingFileHandler

###################################################################################################################################

# SETUP LOGGING
# -------------
logger = logging.getLogger(__name__) # use the module's name as the name in the logs
logger.setLevel(logging.INFO) # set the logging level
log_file_path = 'execution_logs.log' # define the logging file path

# Use RotatingFileHandler
# log_file_path = the absolute path for the log file
# maxBytes: 5 * 1024
# backupCount=0: When the file is full, delete it and start a new one
handler = RotatingFileHandler(
    log_file_path, maxBytes=5*1024*1024, backupCount=0
)

# Set the format of the log messages
formatter = logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s")
handler.setFormatter(formatter)

# Add the formatted HANDLER to the logger
logger.addHandler(handler)

###################################################################################################################################

# ── PALETTE — USMC: Navy Blue / Scarlet Red / White ───────────────────────────

BG_BASE        = "#0a0f2e"      # deep navy
BG_PANEL       = "#0d1435"      # slightly lighter navy panel
BG_ROW         = "#0d1435"
BG_ROW_ALT     = "#111840"      # alternating row — subtle lighter navy
BORDER         = "#1e2d6e"      # navy border
BORDER_LIGHT   = "#263580"      # lighter navy border

ACCENT         = "#cc0000"      # USMC scarlet
ACCENT_MID     = "#e00000"      # brighter scarlet
ACCENT_LIGHT   = "#ff3333"      # light scarlet for hover
ACCENT_SHINE   = "#ff6666"      # scarlet highlight
ACCENT_DIM     = "#2a0a0a"      # very dark scarlet tint

SILVER         = "#d0d8e8"      # metallic silver-white
SILVER_BRIGHT  = "#eef2fa"      # bright silver
SILVER_DIM     = "#8898b8"      # muted silver

TEXT_PRIMARY   = "#eef2fa"      # bright silver-white — max readability
TEXT_SECONDARY = "#b0bcd8"      # mid silver
TEXT_MUTED     = "#6878a8"      # muted blue-silver
TEXT_HEADER    = "#d0d8e8"      # silver for headers

GOLD           = "#c8a84a"      # brass/gold — title details, SHOW button base, and premium accents
GOLD_LIGHT     = "#e8c870"      # light brass/gold — hover states for gold buttons and shimmer effects

DANGER         = "#cc0000"      # scarlet red — CANCEL button enabled state and destructive actions
DANGER_LIGHT   = "#ff3333"      # bright scarlet — CANCEL button hover and warning highlights
DANGER_DIM     = "#2a0808"      # dark scarlet tint — subtle danger backgrounds and warning overlays

BTN_DISABLED   = "#1a2050"      # dark navy — background for disabled START button
BTN_DIS_TEXT   = "#3a4878"      # muted navy-silver — text color for disabled buttons

SUCCESS_GREEN  = "#33cc66"      # clear green — result state label in PushCommandsDialog
 
GREEN          = "#1a8a40"      # base dark green — PUSH and PUSH MORE button base
GREEN_MID      = "#22aa50"      # mid green — gradient midpoint
GREEN_LIGHT    = "#2acc60"      # lighter green — hover state
GREEN_SHINE    = "#44ee80"      # shine highlight — top of hover gradient

###################################################################################################################################

STYLESHEET = f"""

    QMainWindow {{
        background-color: {BG_BASE};
    }}
    QWidget#centralWidget {{
        background: qlineargradient(
            x1:0, y1:0, x2:1, y2:1,
            stop:0   #0c1138,
            stop:0.3 #0a0f2e,
            stop:0.7 #0d1230,
            stop:1   #080d28
        );
    }}

    /* ── Title ── */
    QLabel#appTitle {{
        color: {SILVER_BRIGHT};
        font-size: 13px;
        font-family: 'Courier New', monospace;
        letter-spacing: 4px;
        font-weight: bold;
        font-style: italic;
    }}
    QLabel#appSubtitle {{
        color: {GOLD};
        font-size: 9px;
        font-family: 'Courier New', monospace;
        letter-spacing: 5px;
        font-style: oblique; 
    }}
    QLabel#sectionLabel {{
        color: {SILVER};
        font-size: 12px;
        font-family: 'Courier New', monospace;
        letter-spacing: 3px;
        font-weight: bold;
    }}

    /* ── Sheet input ── */
    QLineEdit#sheetInput {{
        background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
            stop:0 #161e50, stop:1 {BG_PANEL});
        color: {TEXT_PRIMARY};
        border: 1px solid {BORDER_LIGHT};
        border-radius: 6px;
        padding: 8px 14px;
        font-size: 13px;
        font-family: 'Courier New', monospace;
        selection-background-color: {ACCENT_DIM};
        selection-color: {ACCENT_SHINE};
    }}
    QLineEdit#sheetInput:focus {{
        border: 1px solid {ACCENT};
        background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
            stop:0 #1a2260, stop:1 {BG_PANEL});
    }}

   /* ── Show button ── */
    QPushButton#btnShow {{
        background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
            stop:0 {GOLD_LIGHT}, stop:0.4 #d4a840, stop:1 {GOLD});
        color: #0a0f2e;
        border: 1px solid {GOLD};
        border-bottom: 2px solid #8a6818;
        border-radius: 6px;
        padding: 8px 18px;
        font-size: 11px;
        font-family: 'Courier New', monospace;
        font-weight: bold;
        letter-spacing: 1px;
    }}
    QPushButton#btnShow:hover {{
        background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
            stop:0 #f8e898, stop:0.4 {GOLD_LIGHT}, stop:1 #d4a840);
    }}
    QPushButton#btnShow:pressed {{
        background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
            stop:0 {GOLD}, stop:1 #8a6818);
        border-bottom: 1px solid #6a5010;
    }}

    QPushButton#btnShow:disabled {{
        background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
            stop:0 #6a5a2a, stop:0.4 #5a4a20, stop:1 #4a3a18);
        color: #1a1a0a;  /* Very dark text for contrast */
        border: 1px solid #5a4a20;
        border-bottom: 2px solid #3a2a10;
    }}

    /* ── Table panel ── */
    QFrame#tablePanel {{
        background: #141e50;
        border: 1px solid {BORDER_LIGHT};
        border-top: 2px solid {ACCENT};
        border-radius: 8px;
    }}

    /* ── Table ── */
QTableWidget {{
    background-color: #0d1435;  /* solid background for empty space */
    color: #eef2fa;
    gridline-color: #1e2d6e;
    border: none;
    font-size: 12px;
    font-family: 'Courier New', monospace;
    selection-background-color: #2a0a0a;
    selection-color: #ff6666;
    outline: none;
    }}

    QTableWidget QAbstractItemView {{
        background-color: #0d1435;
    }}

    QTableWidget::item {{
        padding: 4px 10px;
        border: none;
        color: {TEXT_PRIMARY};
    }}
    QTableWidget::item:alternate {{
        background-color: {BG_ROW_ALT};
    }}
    QTableWidget::item:selected {{
        background-color: {ACCENT_DIM};
        color: {ACCENT_SHINE};
    }}

    /* ── Table header ── */
    QHeaderView::section {{
        background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
            stop:0 #1e2d70, stop:0.5 #192560, stop:1 #141e50);
        color: {TEXT_HEADER};
        border: none;
        border-bottom: 2px solid {ACCENT};
        border-right: 1px solid {BORDER};
        padding: 7px 10px;
        font-size: 10px;
        font-family: 'Courier New', monospace;
        letter-spacing: 2px;
        font-weight: bold;
    }}
    QHeaderView::section:hover {{
        background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
            stop:0 #2a3a88, stop:1 #1e2d70);
    }}
    QHeaderView::section:last {{ border-right: none; }}

    /* ── Select All button ── */
    QPushButton#btnSelectAll {{
        background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
            stop:0 {GOLD_LIGHT}, stop:1 {GOLD});
        color: #0a0f2e;
        border: 1px solid {GOLD};
        border-radius: 3px;
        font-size: 8px;
        font-family: 'Courier New', monospace;
        font-weight: bold;
        letter-spacing: 1px;
        padding: 2px 4px;
    }}
    QPushButton#btnSelectAll:hover {{
        background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
            stop:0 #f8e090, stop:1 {GOLD_LIGHT});
    }}
    QPushButton#btnSelectAll:checked {{
        background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
            stop:0 {ACCENT_MID}, stop:1 {ACCENT});
        border-color: {ACCENT_MID};
        color: {SILVER_BRIGHT};
    }}

    /* ── Checkbox ── */
    QCheckBox {{ spacing: 0px; }}
    QCheckBox::indicator {{
        width: 15px;
        height: 15px;
        border: 1px solid {BORDER_LIGHT};
        border-radius: 3px;
        background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
            stop:0 #1a2260, stop:1 {BG_PANEL});
    }}
    QCheckBox::indicator:checked {{
        background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
            stop:0 {ACCENT_LIGHT}, stop:1 {ACCENT});
        border-color: {ACCENT_MID};
    }}
    QCheckBox::indicator:hover {{ border-color: {ACCENT_MID}; }}

    /* ── Start button ── */
    QPushButton#btnStart {{
        background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
            stop:0 #4a8a50, stop:0.4 #306838, stop:1 #1e4824);
        color: {SILVER_BRIGHT};
        border: 1px solid #3a7840;
        border-bottom: 2px solid #142e18;
        border-radius: 6px;
        font-size: 12px;
        font-family: 'Courier New', monospace;
        font-weight: bold;
        letter-spacing: 2px;
    }}
    QPushButton#btnStart:hover {{
        background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
            stop:0 #5aaa60, stop:0.4 #4a8a50, stop:1 #306838);
    }}
    QPushButton#btnStart:pressed {{
        background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
            stop:0 #1e4824, stop:1 #0e2810);
        border-bottom: 1px solid #0e2810;
    }}
    QPushButton#btnStart:disabled {{
        background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
            stop:0 #2a4a30, stop:0.4 #1a3820, stop:1 #0e2814);;
        color: #6a8a70;
        border: 1px solid {BORDER};
        border-bottom: 2px solid {BORDER};
    }}

    /* ── Cancel button ── */
    QPushButton#btnCancel {{
        background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
            stop:0 #1a2260, stop:1 #0d1435);
        color: #8898b8;
        border: 1px solid #263580;
        border-bottom: 2px solid #1e2d6e;
        border-radius: 6px;
        font-size: 12px;
        font-family: 'Courier New', monospace;
        letter-spacing: 2px;
    }}
    QPushButton#btnCancel:enabled {{
        background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
            stop:0 #ff3333, stop:0.4 #e00000, stop:1 #cc0000);
        color: #eef2fa;
        border: 1px solid #e00000;
        border-bottom: 2px solid #8a0000;
    }}
    QPushButton#btnCancel:enabled:hover {{
        background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
            stop:0 #ff5555, stop:0.4 #ff3333, stop:1 #e00000);
    }}
    QPushButton#btnCancel:enabled:pressed {{
        background: #cc0000;
        border-bottom: 1px solid #660000;
    }}

    QPushButton#btnCancel:disabled {{
        background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
            stop:0 #5a2020, stop:0.4 #4a1818, stop:1 #3a1010);
        color: #aa8888;
        border: 1px solid #4a1818;
        border-bottom: 2px solid #2a0808;
    }}

    /* ── Vertical scrollbar ── */
    QScrollBar:vertical {{
        background: {BG_PANEL};
        width: 8px;
        margin: 0;
        border-radius: 4px;
    }}
    QScrollBar::handle:vertical {{
        background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
            stop:0 {ACCENT_LIGHT}, stop:0.5 {ACCENT_MID}, stop:1 {ACCENT});
        border-radius: 4px;
        min-height: 28px;
    }}
    QScrollBar::handle:vertical:hover {{
        background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
            stop:0 #ff5555, stop:1 {ACCENT_MID});
    }}
    QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}
    QScrollBar::add-page:vertical,  QScrollBar::sub-page:vertical  {{ background: none; }}

    /* ── Horizontal scrollbar ── */
    QScrollBar:horizontal {{
        background: {BG_PANEL};
        height: 8px;
        margin: 0;
        border-radius: 4px;
    }}
    QScrollBar::handle:horizontal {{
        background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
            stop:0 {ACCENT_LIGHT}, stop:0.5 {ACCENT_MID}, stop:1 {ACCENT});
        border-radius: 4px;
        min-width: 28px;
    }}
    QScrollBar::handle:horizontal:hover {{
        background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
            stop:0 #ff5555, stop:1 {ACCENT_MID});
    }}
    QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{ width: 0; }}
    QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal {{ background: none; }}
"""

###################################################################################################################################

# FUDNAMENTAL VARIABLES 
# ---------------------

# Get the absolute path of the EXCEL file
excel_file = 'LAN_Network_Devices.xlsx'

#------------------------------
logger.info("STARTING PROGRAM")
#------------------------------

###################################################################################################################################

# DIALOG BOX FOR ERROR MESSAGES
# ------------------------------
class ErrorDialog(QDialog):

    """
    Modal dialog window for displaying error messages to the user.
    
    Presents an error message with a scarlet-themed appearance matching
    the application's USMC palette. The user can close the dialog using
    the X button or by pressing Escape. No OK button is present — the
    user dismisses it when ready.
    """

    ### <=== STYLESHEET AND INITIALIZATION ===> ###
    def __init__(self, title, message, parent=None):

        """
        PURPOSE
        -------
        Initialize the error dialog window with a title and message.
        
        Creates a modal dialog that dynamically sizes itself based on the
        message content. The dialog features a scarlet-themed appearance
        matching the application's USMC palette, with a warning icon,
        separator line, and instructional hint text.
        
        The dialog has no OK button — the user dismisses it by clicking
        the X button or pressing Escape.
        

        ARGUMENTS
        ---------
        title (str): The dialog window title, also displayed as a header alongside the warning icon.
                       
        message (str): The error message to display. Long messages will wrap automatically and the dialog 
        will expand vertically to accommodate them.

        parent (QWidget, optional): Parent widget for modal behavior. Defaults to None.
        

        RETURN VALUE
        ------------
        None
        """

        super().__init__(parent)
        self.setWindowTitle(title)

        # No fixed size — let dialog auto-size based on content
        self.setMinimumWidth(500)      # Set minimum width only
        self.setModal(True)
        
        # Remove question mark from title bar (Windows)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        
        # Apply styling to match main window theme (error/danger variant)
        self.setStyleSheet(f"""
            QDialog {{
                background-color: {BG_PANEL};
                border: 2px solid {DANGER};
            }}
            QLabel {{
                color: {SILVER_BRIGHT};
                font-size: 12px;
                font-family: 'Courier New', monospace;
                letter-spacing: 0px;
            }}
            QLabel#errorTitle {{
                color: {DANGER_LIGHT};
                font-size: 14px;
                font-family: 'Courier New', monospace;
                font-weight: bold;
                letter-spacing: 1px;
            }}
        """)
        
        # Method call that builds and arranges all the visual components inside the error dialog window.
        self._setup_ui(title, message)
        
        # Auto-size the dialog to fit the content
        self.adjustSize()


    ### <=== LAYOUT MANAGER ===> ###
    def _setup_ui(self, title, message):

        """
        PURPOSE
        -------
        Create and arrange the dialog's UI components.
        
        Builds the visual structure of the error dialog using a vertical
        box layout. The dialog consists of:
            - Header row with warning icon (⚠) and title text
            - Scarlet separator line
            - Word-wrapped error message that expands vertically as needed
            - Stretchable spacer
            - Hint text at the bottom ("Close this window to continue")
        

        ARGUMENTS
        ---------
        title (str):   The dialog window title, displayed in uppercase alongside the warning icon in the header row.

        message (str): The error message to display. Word wrap is enable and the label can expand vertically to accommodate
        multi-line content.
        

        RETURN VALUES
        -------------
        None — modifies the dialog's layout in place.

        """

        # Main vertical layout container for the entire dialog
        layout = QVBoxLayout(self)
        layout.setSpacing(16)  # Space between each UI element
        layout.setContentsMargins(24, 20, 24, 20)  # Padding: left, top, right, bottom
        
        # === HEADER ROW: Warning Icon + Title ===
        header_row = QHBoxLayout()
        header_row.setSpacing(10)  # Space between icon and title
        
        error_icon = QLabel("⚠")
        error_icon.setStyleSheet(f"""
            color: {DANGER_LIGHT};
            font-size: 24px;
        """)
        header_row.addWidget(error_icon)
        
        title_label = QLabel(title.upper())  # Display title in ALL CAPS
        title_label.setObjectName("errorTitle")  # ID for stylesheet targeting
        header_row.addWidget(title_label)
        header_row.addStretch()  # Pushes everything left
        
        layout.addLayout(header_row)  # Add header to main layout
        
        # === VISUAL SEPARATOR ===
        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)  # Horizontal line
        sep.setStyleSheet(f"background-color: {DANGER}; max-height: 1px; border: none;")
        layout.addWidget(sep)
        
        # === MAIN ERROR MESSAGE ===
        message_label = QLabel(message)
        message_label.setWordWrap(True)  # Text wraps to fit width
        message_label.setAlignment(Qt.AlignTop | Qt.AlignLeft)  # Stick to top-left
        message_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)  # Grow vertically
        layout.addWidget(message_label)
        
        layout.addStretch()  # Flexible space — pushes hint to bottom
        
        # === BOTTOM HINT TEXT ===
        hint_label = QLabel("Close this window to continue")
        hint_label.setStyleSheet(f"""
            color: {TEXT_MUTED};
            font-size: 10px;
            font-style: italic;
        """)
        hint_label.setAlignment(Qt.AlignCenter)  # Centered horizontally
        layout.addWidget(hint_label)

# DIALOG BOX FOR WARNING MESSAGES
# --------------------------------
class WarningDialog(QDialog):
 
    """
    Modal dialog for displaying non-fatal warning messages to the user.
 
    Styled with a gold/amber theme to visually distinguish warnings from
    the scarlet error dialogs. The workflow continues after dismissal —
    no action is required from the user beyond closing the window.
    """
 
    ### <=== STYLESHEET AND INITIALIZATION ===> ###
    def __init__(self, title, message, parent=None):
 
        """
        PURPOSE
        -------
        Initialize the warning dialog with a title and message.
 
        Uses the same layout structure as ErrorDialog but with a gold accent
        to communicate non-fatal severity. Auto-sizes to fit the message content.
 
 
        ARGUMENTS
        ---------
        title   (str):                  The dialog window title, displayed in the header row.
        message (str):                  The warning message to display. Supports multi-line content.
        parent  (QWidget, optional):    Parent widget for modal behavior. Defaults to None.
 
 
        RETURN VALUE
        ------------
        None
        """
 
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setMinimumWidth(500)
        self.setModal(True)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
 
        self.setStyleSheet(f"""
            QDialog {{
                background-color: {BG_PANEL};
                border: 2px solid {GOLD};
            }}
            QLabel {{
                color: {SILVER_BRIGHT};
                font-size: 12px;
                font-family: 'Courier New', monospace;
                letter-spacing: 0px;
            }}
            QLabel#warnTitle {{
                color: {GOLD_LIGHT};
                font-size: 14px;
                font-family: 'Courier New', monospace;
                font-weight: bold;
                letter-spacing: 1px;
            }}
        """)
 
        self._setup_ui(title, message)
        self.adjustSize()
 
 
    ### <=== LAYOUT MANAGER ===> ###
    def _setup_ui(self, title, message):
 
        """
        PURPOSE
        -------
        Build and arrange all child widgets inside the dialog.
 
        Mirrors the ErrorDialog layout: icon + title header, gold separator,
        word-wrapped message body, and a bottom hint label.
 
 
        ARGUMENTS
        ---------
        title   (str): Displayed in uppercase alongside the warning icon.
        message (str): Warning body text. Word wrap is enabled.
 
 
        RETURN VALUE
        ------------
        None — modifies the dialog layout in place.
        """
 
        layout = QVBoxLayout(self)
        layout.setSpacing(16)
        layout.setContentsMargins(24, 20, 24, 20)
 
        # === HEADER ROW: Icon + Title ===
        header_row = QHBoxLayout()
        header_row.setSpacing(10)
 
        warn_icon = QLabel("⚠")
        warn_icon.setStyleSheet(f"color: {GOLD_LIGHT}; font-size: 24px;")
        header_row.addWidget(warn_icon)
 
        title_label = QLabel(title.upper())
        title_label.setObjectName("warnTitle")
        header_row.addWidget(title_label)
        header_row.addStretch()
 
        layout.addLayout(header_row)
 
        # === GOLD SEPARATOR ===
        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        sep.setStyleSheet(f"background-color: {GOLD}; max-height: 1px; border: none;")
        layout.addWidget(sep)
 
        # === WARNING MESSAGE ===
        message_label = QLabel(message)
        message_label.setWordWrap(True)
        message_label.setAlignment(Qt.AlignTop | Qt.AlignLeft)
        message_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        layout.addWidget(message_label)
 
        layout.addStretch()
 
        # === BOTTOM HINT ===
        hint_label = QLabel("Close this window to continue")
        hint_label.setStyleSheet(f"""
            color: {TEXT_MUTED};
            font-size: 10px;
            font-style: italic;
        """)
        hint_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(hint_label)


# DIALOG BOX FOR CREDENTIALS
# --------------------------
class Credentials_and_mode_Dialog(QDialog):

    """
    Modal dialog window for collecting device authentication credentials and mode of operation.
    
    Prompts the user to enter a username and password that will be used
    to authenticate against all Cisco devices listed in the Excel sheet.
    The password field includes a visibility toggle (eye icon) button.

    It also lets the user choos between UPDATER or COMMAND PUSHER modes.
    """

    ### <=== STYLESHEET AND INITIALIZATION ===> ###
    def __init__(self, parent=None):

        """
        PURPOSE
        -------
        Initialize the credentials and mode dialog window with input fields, mode selection,
        and a submit button.
        
        Creates a modal dialog that collects device authentication credentials
        (username and password) and the desired operation mode (UPDATER or
        COMMAND PUSHER) from the user. The dialog features the USMC
        navy/scarlet/gold theme with a visibility toggle for the password
        field and mutually exclusive mode buttons. The SUBMIT button remains
        disabled until both credentials are supplied and a mode is chosen.
        

        ARGUMENTS
        ---------
        parent (QWidget, optional): Parent widget for modal behavior. Defaults to None.
        

        RETURN VALUE
        ------------
        None
        """

        super().__init__(parent)
        self.setWindowTitle("Device Credentials")
        self.setFixedSize(470, 360)
        self.setModal(True) # Blocks interaction with main window until dialog is closed
        
        # Remove question mark from title bar (Windows)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        
        # Apply styling to match main window theme
        self.setStyleSheet(f"""
                           
            /* Dialog background */                
            QDialog {{
                background-color: {BG_PANEL};
            }}

            /* Standard labels (USERNAME, PASSWORD) */
            QLabel {{
                color: {SILVER};
                font-size: 11px;
                font-family: 'Courier New', monospace;
                font-weight: bold;
                letter-spacing: 1px;
            }}

            /* Instructional text at top of dialog */
            QLabel#infoLabel {{
                color: {SILVER_BRIGHT};
                font-size: 15px;
                font-family: 'Courier New', monospace;
                letter-spacing: 0px;
                font-weight: normal;
            }}

            /* Username and password input fields */
            QLineEdit {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #161e50, stop:1 {BG_PANEL});
                color: {TEXT_PRIMARY};
                border: 1px solid {BORDER_LIGHT};
                border-radius: 4px;
                padding: 6px 10px;
                font-size: 11px;
                font-family: 'Courier New', monospace;
                selection-background-color: {ACCENT_DIM};
            }}

            /* Input field when focused/typing */
            QLineEdit:focus {{
                border: 1px solid {ACCENT};
            }}

            /* SUBMIT button (gold) */
            QPushButton {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 {GOLD_LIGHT}, stop:0.4 #d4a840, stop:1 {GOLD});
                color: #0a0f2e;
                border: 1px solid {GOLD};
                border-bottom: 2px solid #8a6818;
                border-radius: 4px;
                padding: 6px 16px;
                font-size: 11px;
                font-family: 'Courier New', monospace;
                font-weight: bold;
                letter-spacing: 1px;
            }}

            /* SUBMIT button hover */
            QPushButton:hover {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #f8e898, stop:0.4 {GOLD_LIGHT}, stop:1 #d4a840);
            }}

            /* SUBMIT button pressed */
            QPushButton:pressed {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 {GOLD}, stop:1 #8a6818);
                border-bottom: 1px solid #6a5010;
            }}

            /* SUBMIT button disabled */
            QPushButton:disabled {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #6a5a2a, stop:0.4 #5a4a20, stop:1 #4a3a18);
                color: #2a2a0a;
                border: 1px solid #5a4a20;
                border-bottom: 2px solid #3a2a10;
            }}

            /* Eye icon button for password visibility */
            QPushButton#btnToggle {{
                background: {BORDER_LIGHT};
                color: {SILVER};
                border: 1px solid {BORDER};
                border-bottom: 2px solid {BORDER};
                padding: 6px 10px;
                font-size: 9px;
            }}

            /* Eye icon button hover */
            QPushButton#btnToggle:hover {{
                background: {BORDER};
                color: {SILVER_BRIGHT};
            }}

            /* Mode buttons — unselected (dark navy) */
            QPushButton#btnModeUpdater,
            QPushButton#btnModePusher {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #1a2260, stop:1 {BG_PANEL});
                color: {SILVER_DIM};
                border: 1px solid {BORDER_LIGHT};
                border-bottom: 2px solid {BORDER};
                border-radius: 4px;
                padding: 6px 10px;
                font-size: 10px;
                font-family: 'Courier New', monospace;
                font-weight: bold;
                letter-spacing: 1px;
            }}
 
            /* Mode buttons — hover */
            QPushButton#btnModeUpdater:hover,
            QPushButton#btnModePusher:hover {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #2a3a88, stop:1 #1e2d70);
                color: {SILVER};
            }}

            /* UPDATER selected — gold */
            QPushButton#btnModeUpdater:checked {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 {GOLD_LIGHT}, stop:0.4 #d4a840, stop:1 {GOLD});
                color: #0a0f2e;
                border: 1px solid {GOLD};
                border-bottom: 2px solid #8a6818;
            }}
 
            /* COMMAND PUSHER selected — silver/blue */
            QPushButton#btnModePusher:checked {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 {SILVER_BRIGHT}, stop:0.4 {SILVER}, stop:1 #a0b0cc);
                color: #0a0f2e;
                border: 1px solid {SILVER};
                border-bottom: 2px solid #7080a0;
            }}

        """)
        
        self._mode   = None  # Stores the selected operation mode
        self._setup_ui() # Build and arrange all dialog widgets
        self._username = "" # Stores validated username after submission
        self._password = "" # Stores validated password after submission


    ### <=== LAYOUT MANAGER ===> ###
    def _setup_ui(self):

        """
        PURPOSE
        -------
        Create and arrange the dialog's UI components.
        
        Builds the visual structure of the credentials dialog using a vertical
        box layout. The dialog now consists of:
            - Instructional info label at the top
            - Username input field with label
            - Password input field with label and visibility toggle button (👁)
            - OPERATION MODE row with two mutually exclusive checkable buttons:
                🔧 UPDATER and 📡 COMMAND PUSHER
            - SUBMIT button at the bottom (disabled until both credentials
                are entered AND a mode is selected)
        

        ARGUMENTS
        ---------
        None
        

        RETURN VALUE
        ------------
        None — modifies the dialog's layout in place.
        """

        layout = QVBoxLayout(self)  # Main vertical container
        layout.setSpacing(12)  # Space between major sections
        layout.setContentsMargins(24, 20, 24, 20)  # Padding: left, top, right, bottom
        
        # === INSTRUCTIONAL TEXT ===
        info = QLabel("Enter valid credentials for the devices in the EXCEL sheet:")
        info.setObjectName("infoLabel")  # ID for stylesheet targeting
        info.setWordWrap(True)  # Allow text to wrap if needed
        layout.addWidget(info)
        
        # === USERNAME FIELD ===
        user_layout = QVBoxLayout()  # Vertical stack for label + input
        user_label = QLabel("USERNAME")
        user_layout.addWidget(user_label)
        
        self.user_input = QLineEdit()
        self.user_input.setPlaceholderText("Enter username...")
        self.user_input.setFixedHeight(32)  # Match eye button height
        self.user_input.textChanged.connect(self._update_submit_state)  # Re-evaluate gate, ensure that some text is in it
        user_layout.addWidget(self.user_input)
        layout.addLayout(user_layout)
        
        # === PASSWORD FIELD WITH VISIBILITY TOGGLE ===
        pass_layout = QVBoxLayout()  # Vertical stack for label + input row
        pass_label = QLabel("PASSWORD")
        pass_layout.addWidget(pass_label)
        
        pass_row = QHBoxLayout()  # Horizontal row for input + eye button
        pass_row.setSpacing(6)  # Space between input and button
        
        self.pass_input = QLineEdit()
        self.pass_input.setPlaceholderText("Enter password...")
        self.pass_input.setEchoMode(QLineEdit.Password)  # Mask input with bullets
        self.pass_input.setFixedHeight(32)  # Match eye button height
        self.pass_input.textChanged.connect(self._update_submit_state)  # Re-evaluate gate, ensure that some text is in it
        pass_row.addWidget(self.pass_input)
        
        self.btn_toggle = QPushButton("👁")  # Eye icon for visibility toggle
        self.btn_toggle.setObjectName("btnToggle")  # ID for stylesheet targeting
        self.btn_toggle.setFixedSize(32, 32)  # Square button
        self.btn_toggle.setCursor(Qt.PointingHandCursor)  # Hand cursor on hover
        self.btn_toggle.setCheckable(True)  # Allows toggled state styling
        self.btn_toggle.clicked.connect(self._toggle_password_visibility)
        pass_row.addWidget(self.btn_toggle)
        
        pass_layout.addLayout(pass_row)
        layout.addLayout(pass_layout)
        
        # === MODE SELECTION ===
        # Allow the user to choose between firmware-update and command‑push modes.
        mode_layout = QVBoxLayout()
        mode_label = QLabel("OPERATION MODE")
        mode_layout.addWidget(mode_label)

        mode_row = QHBoxLayout()
        mode_row.setSpacing(8)  # space between the two buttons

        # UPDATER button – runs the full IOS upgrade pipeline
        self.btn_mode_updater = QPushButton("🔧 UPDATER")
        self.btn_mode_updater.setObjectName("btnModeUpdater")  # CSS selector
        self.btn_mode_updater.setFixedHeight(32)               # compact button
        self.btn_mode_updater.setCheckable(True)               # stays pressed when selected
        self.btn_mode_updater.setCursor(Qt.PointingHandCursor) # hand cursor on hover
        self.btn_mode_updater.clicked.connect(lambda: self._on_mode_selected('updater'))

        # COMMAND PUSHER button – pushes a single config command set
        self.btn_mode_pusher = QPushButton("📡 COMMAND PUSHER")
        self.btn_mode_pusher.setObjectName("btnModePusher")  # CSS selector
        self.btn_mode_pusher.setFixedHeight(32)              # compact button
        self.btn_mode_pusher.setCheckable(True)              # stays pressed when selected
        self.btn_mode_pusher.setCursor(Qt.PointingHandCursor)# hand cursor on hover
        self.btn_mode_pusher.clicked.connect(lambda: self._on_mode_selected('command_pusher'))

        mode_row.addWidget(self.btn_mode_updater)
        mode_row.addWidget(self.btn_mode_pusher)
        mode_layout.addLayout(mode_row)
        layout.addLayout(mode_layout)

        layout.addSpacing(4)  # extra breathing room before the submit button
  
        # === SUBMIT BUTTON ===
        self.btn_submit = QPushButton("SUBMIT")
        self.btn_submit.setFixedHeight(29)
        self.btn_submit.setCursor(Qt.PointingHandCursor)  # Hand cursor on hover
        self.btn_submit.setEnabled(False)  # Gated — requires credentials + mode
        self.btn_submit.clicked.connect(self._on_submit)
        layout.addWidget(self.btn_submit)


    ### <=== PASSWORD VISIBILITY ===> ###
    def _toggle_password_visibility(self):

        """
        PURPOSE
        -------
        Toggle password field between plain text and masked (bullet) display.
        
        When the eye button is checked (pressed), the password is shown in
        plain text and the button icon changes to a lock (🔒). When unchecked,
        the password is masked with bullets and the button shows an eye (👁).
        

        ARGUMENTS
        ---------
        None
        

        RETURN VALUE
        ------------
        None
        """

        if self.btn_toggle.isChecked():  # Button is pressed down
            self.pass_input.setEchoMode(QLineEdit.Normal)  # Show plain text
            self.btn_toggle.setText("🔒")  # Lock icon = password visible
        else:  # Button is not pressed
            self.pass_input.setEchoMode(QLineEdit.Password)  # Mask with bullets
            self.btn_toggle.setText("👁")  # Eye icon = password hidden


    ### <=== MODE SELECTION ===> ###
    def _on_mode_selected(self, mode):
 
        """
        PURPOSE
        -------
        Handle a mode button click. Ensures mutual exclusivity between the two
        mode buttons (only one can be checked at a time), stores the selection,
        and re-evaluates the Submit gate.
 
        
        ARGUMENTS
        ---------
        mode (str): 'updater' or 'command_pusher'
 
        
        RETURN VALUE
        ------------
        None
        """
 
        self._mode = mode
 
        # Enforce mutual exclusivity manually — QButtonGroup would work too but
        # this keeps things explicit and avoids extra imports
        self.btn_mode_updater.setChecked(mode == 'updater')
        self.btn_mode_pusher.setChecked(mode == 'command_pusher')
 
        self._update_submit_state()
 
 
    ### <=== SUBMIT GATE ===> ###
    def _update_submit_state(self):
 
        """
        PURPOSE
        -------
        Enable the Submit button only when all three conditions are met:
            1. Username field is non-empty
            2. Password field is non-empty
            3. An operation mode has been selected
 
        Called on every textChanged signal from both input fields and after
        every mode button click.
 
        
        ARGUMENTS
        ---------
        None
 
        
        RETURN VALUE
        ------------
        None
        """
 
        has_user = bool(self.user_input.text().strip())
        has_pass = bool(self.pass_input.text().strip())
        has_mode = self._mode is not None
        self.btn_submit.setEnabled(has_user and has_pass and has_mode)


    ### <=== CREDENTIAL VARIABLES MANAGEMENT ===> ###
    # Validates user input upon submission
    def _on_submit(self):

        """
        PURPOSE
        -------
        Validate and store the entered credentials.
        
        Checks that both username and password fields contain non-whitespace
        text. If valid, stores the values and accepts the dialog. If either
        field is empty, displays a temporary warning message that does not
        block the dialog's close button.
        

        ARGUMENTS
        ---------
        None
        

        RETURN VALUE
        ------------
        None — stores credentials and calls accept() on success, or shows warning and returns early.
        """

        username = self.user_input.text().strip()  # Remove leading/trailing whitespace
        password = self.pass_input.text().strip()
        
        if not username or not password:  # Either field is empty
                # Create a frameless, non-modal warning popup
                warning = QLabel("⚠ Both username and password are required", self)
                warning.setStyleSheet(f"""
                    background-color: {DANGER_DIM};
                    color: {DANGER_LIGHT};
                    border: 1px solid {DANGER};
                    border-radius: 4px;
                    padding: 8px 16px;
                    font-family: 'Courier New', monospace;
                    font-size: 15px;
                    font-weight: bold;
                """)
                warning.setWindowFlags(Qt.ToolTip | Qt.FramelessWindowHint)  # Floating, no borders
                warning.setAlignment(Qt.AlignCenter)
                warning.adjustSize()  # Fit content exactly
                
                # Position near the submit button (centered above it)
                btn_pos = self.btn_submit.mapToGlobal(self.btn_submit.rect().center())
                warning.move(btn_pos.x() - warning.width() // 2, btn_pos.y() - 120)
                warning.show()
                
                # Auto-close after 5 seconds
                QTimer.singleShot(5000, warning.deleteLater)
                return
        
        self._username = username  # Store validated username
        self._password = password  # Store validated password
        self.accept()  # Close dialog with Accepted result

    # Returns credentials as variables
    def get_credentials(self):

        """
        PURPOSE
        -------
        Return the stored credentials after dialog acceptance.
        

        ARGUMENTS
        ---------
        None
        

        RETURN VALUE
        ------------
        tuple: (username, password) as strings.
        """

        return self._username, self._password
    

    ### <=== MODE VARIABLE MANAGEMENT ===> ###
    # Returns the selected operation mode
    def get_mode(self):
 
        """
        PURPOSE
        -------
        Return the selected operation mode after dialog acceptance.
 
        
        ARGUMENTS
        ---------
        None
 
        
        RETURN VALUE
        ------------
        str: 'updater' or 'command_pusher'
        """
 
        return self._mode


# DIALOG BOX FOR TRANSFER MODE SELECTION
# --------------------------------------
class TransferModeDialog(QDialog):

    """
    Modal dialog that prompts the user to choose between Sequential and
    Threaded SCP transfer modes before starting the firmware update.
    """

    ### <=== STYLESHEET AND INITIALIZATION ===> ###
    def __init__(self, parent=None):

        """
        PURPOSE
        -------
        Initialize the dialog with warning text and two selection buttons.
        The dialog is modal and blocks interaction with the main window until
        the user makes a choice or closes it. The selected mode is stored in
        self.selected_mode as either 'sequential' or 'threaded'.
        

        ARGUMENTS
        ---------
        parent (QWidget, optional):
        Parent widget for modal behavior. Defaults to None.
        

        RETURN VALUE
        ------------
        None
        """

        super().__init__(parent)
        self.setWindowTitle("Transfer Mode")
        self.setFixedSize(520, 280)  # Fixed size for consistent appearance
        self.setModal(True)  # Blocks main window interaction
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)  # Remove ? button

        self.selected_mode = None  # Will be set to 'sequential' or 'threaded' on button click

        # Apply styling to match main window theme
        self.setStyleSheet(f"""
                           
            /* Dialog background with gold border */                
            QDialog {{
                background-color: {BG_PANEL};
                border: 2px solid {GOLD};
            }}

            /* Standard labels (explanation text) */
            QLabel {{
                color: {SILVER_BRIGHT};
                font-size: 12px;
                font-family: 'Courier New', monospace;
            }}

            /* Title label at the top of the dialog */
            QLabel#titleLabel {{
                color: {GOLD_LIGHT};
                font-size: 14px;
                font-weight: bold;
                letter-spacing: 2px;
            }}

            /* Default gold button style (used by THREADED button) */
            QPushButton {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 {GOLD_LIGHT}, stop:0.4 #d4a840, stop:1 {GOLD});
                color: #0a0f2e;
                border: 1px solid {GOLD};
                border-bottom: 2px solid #8a6818;
                border-radius: 4px;
                padding: 8px 16px;
                font-size: 12px;
                font-family: 'Courier New', monospace;
                font-weight: bold;
            }}

            /* Gold button hover */
            QPushButton:hover {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #f8e898, stop:0.4 {GOLD_LIGHT}, stop:1 #d4a840);
            }}

            /* Sequential button - styled green to match START button */
            QPushButton#btnSequential {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #4a8a50, stop:0.4 #306838, stop:1 #1e4824);
                color: {SILVER_BRIGHT};
                border: 1px solid #3a7840;
                border-bottom: 2px solid #142e18;
            }}

            /* Sequential button hover */
            QPushButton#btnSequential:hover {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #5aaa60, stop:0.4 #4a8a50, stop:1 #306838);
            }}
        """)

        self._setup_ui()  # Build and arrange all dialog widgets


    ### <=== LAYOUT MANAGER ===> ###
    def _setup_ui(self):

        """
        PURPOSE
        -------
        Create and arrange the dialog's UI components.
        
        Builds the visual structure using a vertical box layout. The dialog consists of:
            - Title label at the top (⚠️ SELECT TRANSFER MODE)
            - Explanatory text describing Sequential and Threaded modes
            - Stretchable spacer
            - Horizontal row with SEQUENTIAL (green) and THREADED (gold) buttons
        

        ARGUMENTS
        ---------
        None
        

        RETURN VALUE
        ------------
        None — modifies the dialog's layout in place.
        """

        layout = QVBoxLayout(self)  # Main vertical container
        layout.setSpacing(14)  # Space between major elements
        layout.setContentsMargins(24, 20, 24, 20)  # Padding: left, top, right, bottom

        # === TITLE ===
        title = QLabel("⚠️ SELECT TRANSFER MODE")
        title.setObjectName("titleLabel")  # ID for stylesheet targeting
        title.setAlignment(Qt.AlignCenter)  # Centered horizontally
        layout.addWidget(title)

        # === EXPLANATION TEXT ===
        info = QLabel(
            "Choose how IOS image files are transferred to the devices.\n\n"
            "SEQUENTIAL (Recommended):\n"
            "• Transfers files one device at a time.\n"
            "• Safer, less network strain, easier to monitor.\n\n"
            "THREADED:\n"
            "• Transfers to all devices simultaneously.\n"
            "• Faster overall, but may saturate network or cause\n"
            "  overlapping terminal output (if running CLI)."
        )
        info.setWordWrap(True)  # Allow text to wrap within the dialog width
        layout.addWidget(info)

        layout.addStretch()  # Pushes buttons to the bottom

        # === BUTTON ROW ===
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(12)  # Space between the two buttons

        self.btn_sequential = QPushButton("🔒 SEQUENTIAL")
        self.btn_sequential.setObjectName("btnSequential")  # Green styling
        self.btn_sequential.setCursor(Qt.PointingHandCursor)  # Hand cursor on hover
        self.btn_sequential.clicked.connect(self._on_sequential)

        self.btn_threaded = QPushButton("⚡ THREADED")
        self.btn_threaded.setCursor(Qt.PointingHandCursor)  # Hand cursor on hover
        self.btn_threaded.clicked.connect(self._on_threaded)

        btn_layout.addWidget(self.btn_sequential)
        btn_layout.addWidget(self.btn_threaded)
        layout.addLayout(btn_layout)


    ### <=== BUTTON HANDLERS ===> ###
    def _on_sequential(self):

        """
        PURPOSE
        -------
        Handle Sequential button click. Sets selected_mode and closes dialog.

        
        ARGUMENTS
        ---------
        None
        

        RETURN VALUE
        ------------
        None
        """

        self.selected_mode = "sequential"
        self.accept()  # Closes dialog with QDialog.Accepted result

    def _on_threaded(self):

        """
        PURPOSE
        -------
        Handle Threaded button click. Sets selected_mode and closes dialog.

        
        ARGUMENTS
        ---------
        None
        

        RETURN VALUE
        ------------
        None
        """

        self.selected_mode = "threaded"
        self.accept()  # Closes dialog with QDialog.Accepted result


# DIALOG BOX FOR UPDATE SUMMARY
# -----------------------------
class SummaryDialog(QDialog):

    """
    Modal dialog that displays the results of the firmware update process.
    Features a gold-themed header matching the USMC palette and a close button.
    """

    ### <=== STYLESHEET AND INITIALIZATION ===> ###
    def __init__(self, title, message, parent=None):

        """
        PURPOSE
        -------
        Initialize the summary dialog with a title and message.
        
        Creates a modal dialog that dynamically sizes to fit the content.
        The dialog uses a gold border and header to distinguish it from
        error dialogs while staying within the USMC theme.
        

        ARGUMENTS
        ---------
        title (str): The dialog window title, also displayed as a header alongside a checkmark icon.
                       
        message (str): The summary message to display. Long messages will wrap automatically
        and the dialog will expand vertically to accommodate them.

        parent (QWidget, optional): Parent widget for modal behavior. Defaults to None.
        

        RETURN VALUE
        ------------
        None
        """
        
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setMinimumWidth(480)      # Set minimum width only — height auto-adjusts
        self.setModal(True)            # Blocks interaction with parent window
        
        # Remove question mark from title bar (Windows)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        
        # Apply styling to match main window theme (gold/success variant)
        self.setStyleSheet(f"""
                           
            /* Dialog background with gold border */                
            QDialog {{
                background-color: {BG_PANEL};
                border: 2px solid {GOLD};
            }}

            /* Standard labels (message text) */
            QLabel {{
                color: {SILVER_BRIGHT};
                font-size: 12px;
                font-family: 'Courier New', monospace;
            }}

            /* Title label at the top of the dialog */
            QLabel#summaryTitle {{
                color: {GOLD_LIGHT};
                font-size: 14px;
                font-weight: bold;
                letter-spacing: 1px;
            }}

            /* Close button (gold) */
            QPushButton {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 {GOLD_LIGHT}, stop:0.4 #d4a840, stop:1 {GOLD});
                color: #0a0f2e;
                border: 1px solid {GOLD};
                border-bottom: 2px solid #8a6818;
                border-radius: 4px;
                padding: 8px 18px;
                font-size: 11px;
                font-family: 'Courier New', monospace;
                font-weight: bold;
            }}

            /* Close button hover */
            QPushButton:hover {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #f8e898, stop:0.4 {GOLD_LIGHT}, stop:1 #d4a840);
            }}

            /* Close button pressed */
            QPushButton:pressed {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 {GOLD}, stop:1 #8a6818);
                border-bottom: 1px solid #6a5010;
            }}
        """)
        
        self._setup_ui(title, message)  # Build and arrange all dialog widgets
        
        # Auto-size the dialog to fit the content
        self.adjustSize()


    ### <=== LAYOUT MANAGER ===> ###
    def _setup_ui(self, title, message):

        """
        PURPOSE
        -------
        Create and arrange the dialog's UI components.
        
        Builds the visual structure using a vertical box layout. The dialog consists of:
            - Header row with checkmark icon (✅) and title text
            - Gold separator line
            - Word-wrapped summary message that expands vertically as needed
            - CLOSE button at the bottom
        

        ARGUMENTS
        ---------
        title (str): The dialog window title, displayed alongside a checkmark icon in the header row.
        
        message (str): The summary message to display. Word wrap is enabled and the label can
        expand vertically to accommodate multi-line content.
        

        RETURN VALUE
        ------------
        None — modifies the dialog's layout in place.
        """

        # Main vertical layout container for the entire dialog
        layout = QVBoxLayout(self)
        layout.setSpacing(16)  # Space between each UI element
        layout.setContentsMargins(24, 20, 24, 20)  # Padding: left, top, right, bottom
        
        # === HEADER ROW: Checkmark Icon + Title ===
        header_row = QHBoxLayout()
        header_row.setSpacing(10)  # Space between icon and title
        
        check_icon = QLabel("✅")  # Checkmark icon for success/summary
        check_icon.setStyleSheet(f"""
            color: {GOLD_LIGHT};
            font-size: 24px;
        """)
        header_row.addWidget(check_icon)
        
        title_label = QLabel(title.upper())  # Display title in ALL CAPS
        title_label.setObjectName("summaryTitle")  # ID for stylesheet targeting
        header_row.addWidget(title_label)
        header_row.addStretch()  # Pushes everything left
        
        layout.addLayout(header_row)  # Add header to main layout
        
        # === VISUAL SEPARATOR ===
        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)  # Horizontal line
        sep.setStyleSheet(f"background-color: {GOLD}; max-height: 1px; border: none;")
        layout.addWidget(sep)
        
        # === MAIN SUMMARY MESSAGE ===
        message_label = QLabel(message)
        message_label.setWordWrap(True)  # Text wraps to fit width
        message_label.setAlignment(Qt.AlignTop | Qt.AlignLeft)  # Stick to top-left
        message_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)  # Grow vertically
        layout.addWidget(message_label)
        
        layout.addStretch()  # Flexible space — pushes button to bottom
        
        # === CLOSE BUTTON ===
        btn_close = QPushButton("CLOSE")
        btn_close.setFixedHeight(34)
        btn_close.setCursor(Qt.PointingHandCursor)  # Hand cursor on hover
        btn_close.clicked.connect(self.accept)  # Close dialog with Accepted result
        layout.addWidget(btn_close)

###################################################################################################################################

# PUSH COMMANDS DIALOG
# --------------------
class PushCommandsDialog(QDialog):
 
    """
    Modal dialog for entering and pushing exec-mode CLI commands to selected devices.

    Lifecycle managed via QStackedWidget (three pages):
        Phase 0 — Input   : user types commands; PUSH button gates on content.
        Phase 1 — Pushing : 'Pushing...' label shown; worker runs on background thread.
        Phase 2 — Output  : per-device tab bar + read-only output area + PUSH MORE button.

    Closing without pushing (X button) is allowed in phases 0 and 2.
    Closing is blocked in phase 1 while the worker is running.
    """

    MAX_LINES    = 20  # Hard ceiling on the number of command lines
    _PHASE_INPUT   = 0  # QStackedWidget page index — input phase
    _PHASE_PUSHING = 1  # QStackedWidget page index — pushing phase
    _PHASE_OUTPUT  = 2  # QStackedWidget page index — output phase
 
    ### <=== STYLESHEET AND INITIALIZATION ===> ###
    def __init__(self, parent=None):
 
        """
        PURPOSE
        -------
        Initialize the dialog, apply the stylesheet, and build all three phase
        pages inside a QStackedWidget. Opens on phase 0 (input).


        ARGUMENTS
        ---------
        parent (QWidget, optional): Parent widget for modal behavior. Defaults to None.


        RETURN VALUE
        ------------
        None
        """
 
        # Basic settings
        super().__init__(parent)
        self.setWindowTitle("Push Commands")
        self.setFixedSize(700, 520)
        self.setModal(True)

        # Remove the '?' help button from the title bar on Windows
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint) 

        # Runtime state
        self._pushing      = False  # True while worker is running — blocks closeEvent
        self._results      = {}     # {df_index: result_str} — populated by _on_push_finished
        self._tab_buttons  = {}     # {df_index: QPushButton} — device tab buttons
        self._active_index = None   # df_index of the currently selected device tab

        # Worker/thread references — populated in _on_push_clicked
        self._push_thread = None
        self._push_worker = None
 
        # Some style here...
        self.setStyleSheet(f"""
 
            QDialog {{
                background-color: {BG_PANEL};
                border: 2px solid {ACCENT};
            }}

            /* Phase 0 — 'Enter commands:' label */
            QLabel#cmdLabel {{
                color: {GOLD_LIGHT};
                font-size: 12px;
                font-weight: bold;
                font-family: 'Courier New', monospace;
                letter-spacing: 1px;
            }}

            /* Phase 1 — 'Pushing...' centered label */
            QLabel#statusLabel {{
                color: {SILVER_BRIGHT};
                font-size: 16px;
                font-family: 'Courier New', monospace;
                font-weight: bold;
                letter-spacing: 2px;
            }}

            /* Phase 2 — 'COMMAND OUTPUT' header */
            QLabel#outputHeader {{
                color: {GOLD_LIGHT};
                font-size: 13px;
                font-weight: bold;
                font-family: 'Courier New', monospace;
                letter-spacing: 2px;
            }}

            /* Command input editor (phase 0) and output area (phase 2) */
            QPlainTextEdit {{
                background-color: {BG_BASE};
                color: {SILVER_BRIGHT};
                font-size: 12px;
                font-family: 'Courier New', monospace;
                border: 1px solid {BORDER_LIGHT};
                border-radius: 4px;
                padding: 6px;
                selection-background-color: {ACCENT};
            }}

            /* Tab scroll area — transparent so panel background shows through */
            QScrollArea {{
                background-color: transparent;
                border: none;
            }}
            QScrollArea > QWidget > QWidget {{
                background-color: transparent;
            }}

            /* Device tab button — inactive */
            QPushButton#tabBtn {{
                background-color: {BG_BASE};
                color: {SILVER_DIM};
                border: 1px solid {BORDER};
                border-radius: 3px;
                padding: 4px 14px;
                font-size: 11px;
                font-family: 'Courier New', monospace;
                font-weight: bold;
            }}
            QPushButton#tabBtn:hover {{
                background-color: {BORDER_LIGHT};
                color: {SILVER_BRIGHT};
            }}

            /* Device tab button — active (gold) */
            QPushButton#tabBtnActive {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 {GOLD_LIGHT}, stop:1 {GOLD});
                color: {BG_BASE};
                border: 1px solid {GOLD};
                border-bottom: 2px solid #7a6020;
                border-radius: 3px;
                padding: 4px 14px;
                font-size: 11px;
                font-family: 'Courier New', monospace;
                font-weight: bold;
            }}

            /* PUSH button — green, enabled */
            QPushButton#pushBtn {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 {GREEN_LIGHT}, stop:0.4 {GREEN_MID}, stop:1 {GREEN});
                color: {SILVER_BRIGHT};
                border: 1px solid {GREEN};
                border-bottom: 2px solid #0d4a22;
                border-radius: 4px;
                padding: 8px 28px;
                font-size: 11px;
                font-family: 'Courier New', monospace;
                font-weight: bold;
                letter-spacing: 1px;
            }}
            QPushButton#pushBtn:hover {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 {GREEN_SHINE}, stop:0.4 {GREEN_LIGHT}, stop:1 {GREEN_MID});
            }}
            QPushButton#pushBtn:pressed {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 {GREEN}, stop:1 #0d4a22);
                border-bottom: 1px solid #082a14;
            }}
            QPushButton#pushBtn:disabled {{
                background: {BTN_DISABLED};
                color: {BTN_DIS_TEXT};
                border: 1px solid {BORDER};
                border-bottom: 2px solid {BORDER};
            }}

            /* PUSH MORE button — same green as PUSH */
            QPushButton#pushMoreBtn {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 {GREEN_LIGHT}, stop:0.4 {GREEN_MID}, stop:1 {GREEN});
                color: {SILVER_BRIGHT};
                border: 1px solid {GREEN};
                border-bottom: 2px solid #0d4a22;
                border-radius: 4px;
                padding: 8px 28px;
                font-size: 11px;
                font-family: 'Courier New', monospace;
                font-weight: bold;
                letter-spacing: 1px;
            }}
            QPushButton#pushMoreBtn:hover {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 {GREEN_SHINE}, stop:0.4 {GREEN_LIGHT}, stop:1 {GREEN_MID});
            }}
            QPushButton#pushMoreBtn:pressed {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 {GREEN}, stop:1 #0d4a22);
                border-bottom: 1px solid #082a14;
            }}
        """)
 
        # Build all three phase pages and wire the stack
        self._setup_ui()
 
 
    ### <=== CLOSE EVENT GUARD ===> ###
    def closeEvent(self, event):

        """
        PURPOSE
        -------
        Block closing while the worker is running. Allowed in phases 0 and 2.


        ARGUMENTS
        ---------
        event (QCloseEvent): The Qt close event.


        RETURN VALUE
        ------------
        None
        """

        if self._pushing:
            event.ignore()   # Push in progress — swallow the close request
        else:
            event.accept()   # Safe to close in phases 0 and 2
 
 
    ### <=== LAYOUT MANAGER ===> ###
    def _setup_ui(self):

        """
        PURPOSE
        -------
        Build all three phase pages inside a QStackedWidget and attach it
        to the dialog. Phase transitions are handled by setCurrentIndex().


        ARGUMENTS
        ---------
        None


        RETURN VALUE
        ------------
        None
        """

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)  # Stack fills the dialog edge-to-edge
        root.setSpacing(0)

        self._stack = QStackedWidget()  # One page visible at a time — controls phase transitions
        root.addWidget(self._stack)

        # Each builder returns a QWidget page; index order must match _PHASE_* constants
        self._stack.addWidget(self._build_phase_input())    # index 0 — _PHASE_INPUT
        self._stack.addWidget(self._build_phase_pushing())  # index 1 — _PHASE_PUSHING
        self._stack.addWidget(self._build_phase_output())   # index 2 — _PHASE_OUTPUT

        self._stack.setCurrentIndex(self._PHASE_INPUT)  # Dialog opens on the input phase
 
 
    ### <=== PHASE 0: INPUT PAGE ===> ###
    def _build_phase_input(self):

        """
        PURPOSE
        -------
        Build and return the input phase page widget containing the
        'Enter commands:' label, the QPlainTextEdit editor, and the
        centered green PUSH button.


        ARGUMENTS
        ---------
        None


        RETURN VALUE
        ------------
        QWidget: The fully constructed input phase page.
        """

        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(14)

        # Section label — tells the user what the editor expects
        self._lbl_cmd = QLabel("Enter commands:")
        self._lbl_cmd.setObjectName("cmdLabel")
        layout.addWidget(self._lbl_cmd)

        # Multi-line command editor — no word wrap so each line stays as one command
        self._editor = QPlainTextEdit()
        self._editor.setLineWrapMode(QPlainTextEdit.NoWrap)              # One command per line, no soft wrap
        self._editor.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)  # H-scroll when a line is too long
        self._editor.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)    # V-scroll when lines exceed visible height
        self._editor.setTabChangesFocus(True)                             # Tab moves focus to PUSH instead of inserting a tab character
        self._editor.setPlaceholderText("  one command per line  |  max 20 lines")
        self._editor.textChanged.connect(self._on_text_changed)          # Live line-cap and button gating on every keystroke
        layout.addWidget(self._editor)

        # PUSH button — starts disabled; _on_text_changed enables it when content exists
        self._btn_push = QPushButton("PUSH")
        self._btn_push.setObjectName("pushBtn")
        self._btn_push.setFixedHeight(36)
        self._btn_push.setEnabled(False)
        self._btn_push.setCursor(Qt.PointingHandCursor)
        self._btn_push.clicked.connect(self._on_push_clicked)

        btn_row = QHBoxLayout()          # Horizontal row used only to center the button
        btn_row.addStretch()             # Left stretch — pushes button to center
        btn_row.addWidget(self._btn_push)
        btn_row.addStretch()             # Right stretch — mirrors left
        layout.addLayout(btn_row)

        return page


    ### <=== PHASE 1: PUSHING PAGE ===> ###
    def _build_phase_pushing(self):

        """
        PURPOSE
        -------
        Build and return the pushing phase page widget showing a vertically
        centered 'Pushing...' label while the worker runs.


        ARGUMENTS
        ---------
        None


        RETURN VALUE
        ------------
        QWidget: The fully constructed pushing phase page.
        """

        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(24, 20, 24, 20)

        # Status label — displayed centered while push_commands_all runs in background
        self._lbl_status = QLabel("Pushing...")
        self._lbl_status.setObjectName("statusLabel")
        self._lbl_status.setAlignment(Qt.AlignCenter)

        layout.addStretch()                 # Top stretch — pushes label down to vertical center
        layout.addWidget(self._lbl_status)
        layout.addStretch()                 # Bottom stretch — mirrors top

        return page


    ### <=== PHASE 2: OUTPUT PAGE ===> ###
    def _build_phase_output(self):

        """
        PURPOSE
        -------
        Build and return the output phase page widget containing:
            - 'COMMAND OUTPUT' header label
            - Scarlet horizontal separator
            - Horizontally-scrolling device tab bar (QScrollArea)
            - Read-only QPlainTextEdit for device output
            - Centered green PUSH MORE button

        Device tab buttons are populated later by _build_device_tabs() once
        set_push_context() has been called with the actual device DataFrame.


        ARGUMENTS
        ---------
        None


        RETURN VALUE
        ------------
        QWidget: The fully constructed output phase page.
        """

        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(24, 16, 24, 20)
        layout.setSpacing(10)

        # Section header
        lbl_header = QLabel("COMMAND OUTPUT")
        lbl_header.setObjectName("outputHeader")
        layout.addWidget(lbl_header)

        # Visual separator — scarlet line below the header
        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        sep.setStyleSheet(f"background-color: {ACCENT}; max-height: 1px; border: none;")
        layout.addWidget(sep)

        # Device tab bar — inner widget holds the buttons in a horizontal layout
        self._tab_inner = QWidget()
        self._tab_layout = QHBoxLayout(self._tab_inner)
        self._tab_layout.setContentsMargins(2, 4, 2, 4)
        self._tab_layout.setSpacing(6)
        self._tab_layout.addStretch()  # Trailing stretch — keeps buttons anchored left as new ones are inserted before it

        # QScrollArea wraps the inner widget — horizontal scroll only, fixed height for one row of buttons
        self._tab_scroll = QScrollArea()
        self._tab_scroll.setWidget(self._tab_inner)
        self._tab_scroll.setWidgetResizable(True)                                    # Inner widget resizes with content
        self._tab_scroll.setFixedHeight(46)                                          # Tall enough for one button row plus scrollbar
        self._tab_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)          # H-scroll appears when buttons overflow
        self._tab_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)           # Never scroll vertically
        self._tab_scroll.setFrameShape(QFrame.NoFrame)                               # No visible border around the scroll area
        layout.addWidget(self._tab_scroll)

        # Output text area — read-only; content is set by _show_device_output on tab clicks
        self._output_area = QPlainTextEdit()
        self._output_area.setReadOnly(True)
        self._output_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)         # H-scroll for long output lines
        self._output_area.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)           # V-scroll when output exceeds visible height
        self._output_area.setLineWrapMode(QPlainTextEdit.NoWrap)                     # No soft wrap — output lines stay as-is
        layout.addWidget(self._output_area)

        # PUSH MORE button — returns to phase 0 with a cleared editor
        self._btn_push_more = QPushButton("PUSH MORE")
        self._btn_push_more.setObjectName("pushMoreBtn")
        self._btn_push_more.setFixedHeight(36)
        self._btn_push_more.setCursor(Qt.PointingHandCursor)
        self._btn_push_more.clicked.connect(self._on_push_more_clicked)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        btn_row.addWidget(self._btn_push_more)
        btn_row.addStretch()
        layout.addLayout(btn_row)

        return page


     ### <=== DEVICE TAB BUILDER ===> ###
    
   
    ### <=== DEVICE TAB BUILDER ===> ###
    def _build_device_tabs(self):

        """
        PURPOSE
        -------
        Populate the device tab bar with one button per device in _selected_df.
        Called from set_push_context() and again on each PUSH MORE cycle.
        Clears any previously built tabs before rebuilding.


        ARGUMENTS
        ---------
        None


        RETURN VALUE
        ------------
        None
        """

        # Reset state — wipe the previous run's buttons and active selection
        self._tab_buttons.clear()
        self._active_index = None

        # Drain existing button widgets — keep the trailing stretch (always the last item)
        while self._tab_layout.count() > 1:
            item = self._tab_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()  # Schedule for safe deletion on the next event loop cycle

        # Build one tab button per device, inserted before the trailing stretch
        for idx, row in self._selected_df.iterrows():
            hostname = row['Hostname']
            btn = QPushButton(hostname)
            btn.setObjectName("tabBtn")          # Inactive style by default
            btn.setFixedHeight(30)
            btn.setCursor(Qt.PointingHandCursor)
            btn.clicked.connect(lambda checked, i=idx: self._select_device_tab(i))  # Capture idx by value — avoids late-binding closure bug
            self._tab_buttons[idx] = btn
            self._tab_layout.insertWidget(self._tab_layout.count() - 1, btn)        # Insert before the trailing stretch


    ### <=== DEVICE TAB SELECTOR ===> ###
    def _select_device_tab(self, index):

        """
        PURPOSE
        -------
        Highlight the selected device tab (gold) and dim all others,
        then populate the output area with that device's result.


        ARGUMENTS
        ---------
        index: DataFrame index of the device whose tab was clicked.


        RETURN VALUE
        ------------
        None
        """

        self._active_index = index  # Track which device is currently displayed

        for idx, btn in self._tab_buttons.items():
            # Active tab gets gold style; all others get the dimmed navy style
            btn.setObjectName("tabBtnActive" if idx == index else "tabBtn")
            # Qt does not re-evaluate the stylesheet automatically when objectName changes —
            # unpolish + polish forces it to re-read and re-apply the updated selector
            btn.style().unpolish(btn)
            btn.style().polish(btn)

        self._show_device_output(index)  # Refresh the output area for the selected device


    ### <=== DEVICE OUTPUT DISPLAY ===> ###
    def _show_device_output(self, index):

        """
        PURPOSE
        -------
        Populate the output area with the stored result for the given device.


        ARGUMENTS
        ---------
        index: DataFrame index of the device to display.


        RETURN VALUE
        ------------
        None
        """

        result   = self._results.get(index, "No result recorded")   # Fall back gracefully if index is somehow missing
        hostname = self._selected_df.at[index, 'Hostname']           # .at[] for fast scalar lookup by label
        ip       = self._selected_df.at[index, 'OOBM IP Address']

        self._output_area.setPlainText(
            f"Device  :  {hostname}  ({ip})\n"
            f"{'─' * 40}\n"
            f"Result  :  {result}\n"
        )


    ### <=== TEXT CHANGE HANDLER ===> ###
    def _on_text_changed(self):

        """
        PURPOSE
        -------
        Enforce the 20-line hard limit and gate the PUSH button on every
        keystroke or paste event.


        ARGUMENTS
        ---------
        None


        RETURN VALUE
        ------------
        None
        """

        lines = self._editor.toPlainText().splitlines()

        if len(lines) > self.MAX_LINES:
            # blockSignals prevents this setPlainText from re-triggering _on_text_changed recursively
            self._editor.blockSignals(True)
            self._editor.setPlainText("\n".join(lines[:self.MAX_LINES]))  # Truncate to first 20 lines
            cursor = self._editor.textCursor()
            cursor.movePosition(QTextCursor.End)    # Move cursor to end so typing continues from line 20
            self._editor.setTextCursor(cursor)
            self._editor.blockSignals(False)        # Re-enable signals after the forced overwrite
            lines = lines[:self.MAX_LINES]          # Use truncated list for the button gate below

        # Enable PUSH only when at least one line has non-whitespace content
        self._btn_push.setEnabled(any(line.strip() for line in lines))


    ### <=== PUSH CLICKED HANDLER ===> ###
    def _on_push_clicked(self):

        """
        PURPOSE
        -------
        Collect the command list, switch to phase 1 (Pushing...), and start
        the PushCommandsWorker on a background QThread.


        ARGUMENTS
        ---------
        None


        RETURN VALUE
        ------------
        None
        """

        commands = self.get_commands()  # Grab clean command list before switching away from the input phase

        self._pushing = True                               # Block closeEvent for the duration of the push
        self._stack.setCurrentIndex(self._PHASE_PUSHING)  # Switch to the "Pushing..." page

        # Standard Qt worker-on-thread pattern
        self._push_thread = QThread()
        self._push_worker = PushCommandsWorker(
            self._selected_df,
            self._username,
            self._password,
            commands
        )
        self._push_worker.moveToThread(self._push_thread)
        self._push_thread.started.connect(self._push_worker.run)           # Thread start triggers worker.run
        self._push_worker.finished.connect(self._on_push_finished)         # Results flow back to the dialog
        self._push_worker.error.connect(self._on_push_error)               # Unhandled worker exceptions surface here
        self._push_worker.finished.connect(self._push_thread.quit)         # Stop the thread when done
        self._push_worker.error.connect(self._push_thread.quit)
        self._push_thread.finished.connect(self._push_worker.deleteLater)  # Qt cleans up objects after thread exits
        self._push_thread.finished.connect(self._push_thread.deleteLater)
        self._push_thread.start()


    ### <=== PUSH FINISHED HANDLER ===> ###
    def _on_push_finished(self, results):

        """
        PURPOSE
        -------
        Store per-device results, auto-select the first device tab, and
        switch to phase 2 (output).


        ARGUMENTS
        ---------
        results (list): List of (index, result_str) tuples from push_commands_all.


        RETURN VALUE
        ------------
        None
        """

        self._pushing = False                                                   # Unblock closeEvent
        self._results = {index: result for index, result in results}            # Index by df_index for fast tab lookup

        self._build_device_tabs()  # Rebuild tabs — clears any buttons from a previous PUSH MORE cycle

        if self._tab_buttons:
            first_index = next(iter(self._tab_buttons))  # iter() respects dict insertion order — first device in
            self._select_device_tab(first_index)         # Auto-select and display first device on entry to phase 2

        self._stack.setCurrentIndex(self._PHASE_OUTPUT)  # Transition to the output phase


    ### <=== PUSH ERROR HANDLER ===> ###
    def _on_push_error(self, error_msg):

        """
        PURPOSE
        -------
        Called when the worker emits an unhandled exception. Populates all
        device results with the error message and transitions to phase 2.


        ARGUMENTS
        ---------
        error_msg (str): Exception string from the worker.


        RETURN VALUE
        ------------
        None
        """

        self._pushing = False

        # Worker-level errors mean nothing ran at all — mark every device with the same error
        # (contrast with push_commands return codes, which are per-device and surface normally via results)
        self._results = {
            idx: f"WORKER ERROR: {error_msg}"
            for idx in self._selected_df.index
        }

        self._build_device_tabs()

        if self._tab_buttons:
            first_index = next(iter(self._tab_buttons))
            self._select_device_tab(first_index)

        self._stack.setCurrentIndex(self._PHASE_OUTPUT)


    ### <=== PUSH MORE HANDLER ===> ###
    def _on_push_more_clicked(self):

        """
        PURPOSE
        -------
        Return to phase 0 with a cleared editor and a disabled PUSH button,
        ready for a new set of commands against the same device selection.


        ARGUMENTS
        ---------
        None


        RETURN VALUE
        ------------
        None
        """

        self._editor.clear()               # PUSH MORE always starts with a fresh editor
        self._btn_push.setEnabled(False)   # Gate resets — user must type content before pushing again
        self._stack.setCurrentIndex(self._PHASE_INPUT)


    ### <=== CREDENTIALS / DEVICE INJECTION ===> ###
    def set_push_context(self, selected_df, username, password):

        """
        PURPOSE
        -------
        Inject the device DataFrame and credentials before exec_() is called.
        Pre-builds the device tab bar so it is ready when phase 2 is shown.


        ARGUMENTS
        ---------
        selected_df (pd.DataFrame): Devices selected by the user in the table.
        username    (str):          Device SSH username.
        password    (str):          Device SSH password.


        RETURN VALUE
        ------------
        None
        """

        self._selected_df = selected_df
        self._username    = username
        self._password    = password
        self._build_device_tabs()  # Pre-build tabs here so phase 2 renders instantly without a rebuild on entry


    ### <=== PUBLIC GETTER ===> ###
    def get_commands(self):

        """
        PURPOSE
        -------
        Return the validated list of commands from the editor, blank lines removed.


        ARGUMENTS
        ---------
        None


        RETURN VALUE
        ------------
        list[str]: Non-empty command lines in order.
        """

        # splitlines handles all line endings; the filter drops blank and whitespace-only lines
        return [line for line in self._editor.toPlainText().splitlines() if line.strip()]

###################################################################################################################################

# MAIN WINDOW
# -----------
class MainWindow(QMainWindow):

    
    ### <=== LAYOUT MANAGER ===> ###
    def __init__(self):
        
        """
        PURPOSE
        -------
        Initialize the main application window and construct its UI layout.
        
        Sets up the window properties, creates all child widgets, arranges them
        in a vertical layout hierarchy, and establishes signal/slot connections.
        The window consists of:
            - Title block with app name and series subtitle
            - Scarlet/gold shimmer separator
            - Top row with Excel sheet input and SHOW button
            - Device table panel with column headers and SELECT ALL button
            - Bottom row with CANCEL UPDATE and START UPDATE buttons
        
        Initializes internal state variables for tracking checkbox selection,
        update workflow status, and credential storage. The table starts empty
        and buttons are appropriately disabled until devices are loaded.
        

        ARGUMENTS
        ---------
        None
        

        RETURN VALUE
        ------------
        None
        """
        
        super().__init__()
        self.setWindowTitle("Cisco Tactical Controller — CrayonEaters Series")
        self.setFixedSize(QSize(860, 610))  # Fixed window dimensions
        self.setStyleSheet(STYLESHEET)

        # Internal state tracking
        self._all_checked           = False  # SELECT ALL toggle state
        self._update_running        = False  # Update workflow in progress flag
        self._selected_for_update   = set()  # Row indices selected when START clicked
        self._mode                  = None   # Active operation mode: 'updater' or 'command_pusher'

        # Credential storage for device authentication
        self._device_username = ""
        self._device_password = ""

        # DataFrame storage
        self._eligible_df = None
        self._valid_devices_df = None

        # Central widget — container for all UI elements
        central = QWidget()
        central.setObjectName("centralWidget")
        self.setCentralWidget(central)

        # Root vertical layout
        root = QVBoxLayout(central)
        root.setContentsMargins(24, 18, 24, 20)  # left, top, right, bottom
        root.setSpacing(14)  # Space between major sections

        # === TITLE BLOCK ===
        # Outer horizontal row: title/subtitle stack on the left, mode indicator on the right
        title_row = QHBoxLayout()
        title_row.setSpacing(0)

        title_text_block = QVBoxLayout()
        title_text_block.setSpacing(2)  # Tight spacing between title and subtitle

        title = QLabel("CISCO TACTICAL CONTROLLER")
        title.setObjectName("appTitle")  # ID for stylesheet targeting

        subtitle = QLabel("CRAYONEATERS SERIES")
        subtitle.setObjectName("appSubtitle")  # ID for stylesheet targeting

        title_text_block.addWidget(title)
        title_text_block.addWidget(subtitle)
 
        title_row.addLayout(title_text_block)
        title_row.addStretch()  # Pushes mode label to the right

        # Mode indicator — starts empty, updated when a mode is selected
        self._mode_label = QLabel("")
        self._mode_label.setObjectName("modeIndicator")
        self._mode_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self._mode_label.setStyleSheet(f"""
            color: {TEXT_MUTED};
            font-size: 9px;
            font-family: 'Courier New', monospace;
            letter-spacing: 2px;
        """)
        title_row.addWidget(self._mode_label)
 
        root.addLayout(title_row)

        # Scarlet/gold shimmer separator (decorative line)
        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)  # Horizontal line
        sep.setStyleSheet(
            f"background: qlineargradient(x1:0, y1:0, x2:1, y2:0, "
            f"stop:0 transparent, stop:0.2 {GOLD_LIGHT}, stop:0.5 {ACCENT}, "
            f"stop:0.8 {GOLD_LIGHT}, stop:1 transparent); "
            f"max-height: 1px; border: none;"
        )
        root.addWidget(sep)

        # === TOP ROW: Sheet Input + Show Button ===
        top_row = QHBoxLayout()
        top_row.setSpacing(10)  # Space between input and button

        self.sheet_input = QLineEdit()
        self.sheet_input.setObjectName("sheetInput")
        self.sheet_input.setPlaceholderText("Enter Excel sheet name...")
        self.sheet_input.setFixedHeight(38)
        self.sheet_input.textChanged.connect(self._update_show_button_state)  # Enable button when text entered

        self.btn_show = QPushButton("SHOW ELIGIBLE DEVICES")
        self.btn_show.setObjectName("btnShow")
        self.btn_show.setFixedHeight(38)
        self.btn_show.setFixedWidth(234)
        self.btn_show.setCursor(Qt.PointingHandCursor)  # Hand cursor on hover
        self.btn_show.setEnabled(False)  # Disabled until text entered
        self.btn_show.clicked.connect(self._on_show_devices)
        self._update_show_button_state()  # Set initial disabled state

        top_row.addWidget(self.sheet_input)
        top_row.addWidget(self.btn_show)
        root.addLayout(top_row)

        # === TABLE PANEL ===
        table_panel = QFrame()
        table_panel.setObjectName("tablePanel")
        panel_layout = QVBoxLayout(table_panel)
        panel_layout.setContentsMargins(0, 0, 0, 0)  # No internal padding
        panel_layout.setSpacing(0)

        # Header bar with "ELIGIBLE DEVICES" label
        header_bar = QHBoxLayout()
        header_bar.setContentsMargins(14, 8, 14, 0)
        header_bar.setSpacing(0)
        lbl = QLabel("ELIGIBLE DEVICES")
        lbl.setObjectName("sectionLabel")
        header_bar.addWidget(lbl)
        header_bar.addStretch()  # Push label to the left
        panel_layout.addLayout(header_bar)

        # Separator line below header
        div = QFrame()
        div.setFrameShape(QFrame.HLine)  # Horizontal line
        div.setStyleSheet(
            f"background: qlineargradient(x1:0, y1:0, x2:1, y2:0, "
            f"stop:0 {BORDER}, stop:0.5 {ACCENT}, stop:1 {BORDER}); "
            f"max-height: 1px; border: none;"
        )
        panel_layout.addWidget(div)

        # === DEVICE TABLE ===
        self.table = QTableWidget()
        self.table.setObjectName("deviceTable")
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels([
            "", "HOSTNAME", "IP ADDRESS", "CURRENT VER", "TARGET VER"
        ])
        self.table.setAlternatingRowColors(True)  # Zebra striping
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)  # Full row selection
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)  # Read-only
        self.table.setShowGrid(False)  # Hide grid lines
        self.table.verticalHeader().setVisible(False)  # Hide row numbers
        self.table.horizontalHeader().setHighlightSections(False)  # No highlight on click
        self.table.verticalHeader().setDefaultSectionSize(34)  # Row height

        # Column sizing — user can drag to resize
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Fixed)
        self.table.setColumnWidth(0, 52)   # Checkbox column
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Interactive)
        self.table.setColumnWidth(1, 300)  # Hostname
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Interactive)
        self.table.setColumnWidth(2, 160)  # IP Address
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Interactive)
        self.table.setColumnWidth(3, 140)  # Current Version
        self.table.horizontalHeader().setSectionResizeMode(4, QHeaderView.Interactive)
        self.table.setColumnWidth(4, 149)  # Target Version

        # Hide the 'Current Version' and 'Target Version' columns initially
        self.table.setColumnHidden(3, True)
        self.table.setColumnHidden(4, True)

        # Scrollbar policies
        self.table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        self.table.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)

        self.table.horizontalHeader().sectionClicked.connect(self._on_header_clicked)  # Header click handler

        panel_layout.addWidget(self.table)

        # SELECT ALL floating button inside col-0 header
        hdr = self.table.horizontalHeader()
        self.btn_select_all = QPushButton("ALL", hdr)  # Parent is header viewport
        self.btn_select_all.setObjectName("btnSelectAll")
        self.btn_select_all.setCheckable(True)  # Can be toggled on/off
        self.btn_select_all.setCursor(Qt.PointingHandCursor)
        self.btn_select_all.setFixedSize(36, 20)
        self.btn_select_all.clicked.connect(self._toggle_select_all)

        root.addWidget(table_panel, stretch=1)  # Table panel expands to fill space

        # === BOTTOM CONTAINER — populated dynamically after mode selection ===
        # Sits in the layout permanently to reserve vertical space; children are
        # added/removed by _build_bottom_buttons() / _clear_bottom_buttons().
        self._bottom_container = QWidget()
        self._bottom_layout    = QHBoxLayout(self._bottom_container)
        self._bottom_layout.setSpacing(12)
        self._bottom_layout.setContentsMargins(0, 0, 0, 0)
        self._bottom_container.setFixedHeight(44)
        root.addWidget(self._bottom_container)
 
        # Initialize empty table
        self.table.setRowCount(0)
        self._update_button_states()  # Set initial button states
        
        # Connect geometry change signal for SELECT ALL button repositioning
        self.table.horizontalHeader().geometriesChanged.connect(self._reposition_select_all)

        # === FOR CANCEL FUNCTIONALITY ===
        self._selected_device_indices = None  # Original DataFrame indices of devices being updated

    # Manage the columns width to adapt at the beginning and fill the space
    def showEvent(self, event):

        """
        PURPOSE
        -------
        Called when the window is shown for the first time; adjust column widths evenly.
        
        
        ARGUMENTS
        ---------
        None
        

        RETURN VALUE
        ------------
        None
        """

        super().showEvent(event)
        # Only set initial widths once – after that the user controls them
        if not hasattr(self, '_initial_columns_set'):
            self._initial_columns_set = True
            self._distribute_columns_evenly()
    
    def _distribute_columns_evenly(self):

        """
        PURPOSE
        -------
        Set Hostname and IP columns to equal halves of the table’s available width.
        
        
        ARGUMENTS
        ---------
        None
        

        RETURN VALUE
        ------------
        None
        """

        total_width = self.table.viewport().width() - self.table.columnWidth(0)  # minus checkbox
        if total_width <= 0:
            return
        half = total_width // 2
        self.table.setColumnWidth(1, half)
        self.table.setColumnWidth(2, half)

    def _fit_columns_for_updater(self):

        """
        PURPOSE
        -------
        Adjust interactive column widths so all four data columns (Hostname,
        IP Address, Current Ver, Target Ver) fit within the table viewport
        without a horizontal scrollbar. Called after columns 3 & 4 are
        unhidden in Updater mode.


        ARGUMENTS
        ---------
        None
        

        RETURN VALUE
        ------------
        None

        """

        # Total usable width minus the fixed checkbox column and scrollbar margin
        table_width = self.table.viewport().width()
        checkbox_width = self.table.columnWidth(0)
        available = table_width - checkbox_width - 4   # 4 px for safety

        if available <= 0:
            return

        # Allocate proportions (sums to 1.0)
        ratios = {
            1: 0.35,   # Hostname
            2: 0.30,   # IP Address
            3: 0.175,  # Current Ver
            4: 0.175,  # Target Ver
        }

        for col, ratio in ratios.items():
            width = max(60, int(available * ratio))  # enforce minimum 60 px
            self.table.setColumnWidth(col, width)

    ##############################################################

    ### <=== SELECT ALL BUTTON ===> ###
    # Center it in COL 0
    def _reposition_select_all(self):

        """
        PURPOSE
        -------
        Repositions the SELECT ALL button so it stays centered inside the
        col-0 header cell at all times.
        
        The button is a floating QPushButton parented to the header viewport
        rather than placed via a layout — so its position must be calculated
        and set manually in pixel coordinates whenever the header geometry
        changes (e.g. on first render, or when the user resizes a column).
        
        The horizontal center is computed by taking the left edge of col-0
        in viewport coordinates, then offsetting by half the difference
        between the column width and the button width.
        
        The vertical center is computed by taking half the difference between
        the header height and the button height.
        
        raise_() ensures the button renders on top of the header cell and is
        not obscured by it.
        

        ARGUMENTS
        ---------
        None
        

        RETURN VALUE
        ------------
        None
        """

        hdr   = self.table.horizontalHeader()
        x     = self.table.columnViewportPosition(0)  # left edge of col-0 in viewport px
        col_w = self.table.columnWidth(0)
        btn_w = self.btn_select_all.width()
        btn_h = self.btn_select_all.height()
        hdr_h = hdr.height()
        self.btn_select_all.move(
            x + (col_w - btn_w) // 2,   # horizontally centered within col-0
            (hdr_h - btn_h) // 2        # vertically centered within the header row
        )
        self.btn_select_all.raise_()  # Bring button to front

    # Capture the select all button click
    def _on_header_clicked(self, col):

        """
        PURPOSE
        -------
        Slot connected to the horizontalHeader().sectionClicked signal.
        
        sectionClicked emits the index of whichever column header the user
        clicked. This function intercepts that signal and delegates to
        _toggle_select_all() only when col-0 is clicked — the checkbox
        column. Clicks on any other column header are ignored, leaving
        their default sort/resize behaviour unaffected.
        

        ARGUMENTS
        ---------
        col (int): Zero-based index of the column header that was clicked.
        

        RETURN VALUE
        ------------
        None
        """

        if col == 0:  # Checkbox column clicked
            self._toggle_select_all()

    # Selects all devices after clicking on the select all button
    def _toggle_select_all(self):

        """
        PURPOSE
        -------
        Toggles the checked state of every device checkbox in the table.
        
        Flips the internal _all_checked flag on each call, then applies
        that state to every row's QCheckBox and syncs the SELECT ALL
        button's checked appearance. Also calls _update_button_states()
        so the START UPDATE and CANCEL UPDATE buttons reflect the new
        selection state immediately.
        

        ARGUMENTS
        ---------
        None
        

        RETURN VALUE
        ------------
        None
        """

        self._all_checked = not self._all_checked  # Flip toggle state
        self.btn_select_all.setChecked(self._all_checked)  # Sync button appearance
        for row in range(self.table.rowCount()):
            chk = self._get_checkbox(row)
            if chk:
                chk.setChecked(self._all_checked)  # Apply to each checkbox
        self._update_button_states()  # Refresh START/CANCEL button states
    

    ### <=== BOTTOM BUTTON MANAGEMENT ===> ###
    def _build_bottom_buttons(self, mode):
 
        """
        PURPOSE
        -------
        Populate self._bottom_container with the appropriate action buttons
        for the selected operation mode. Always clears existing buttons first.
 
        Updater:        CANCEL UPDATE (1/3) + START UPDATE (2/3)
        Command Pusher: PUSH COMMANDS (full width)
 
        Also updates the mode indicator label in the title row.
 
 
        ARGUMENTS
        ---------
        mode (str): 'updater' or 'command_pusher'
 
 
        RETURN VALUE
        ------------
        None
        """
 
        self._clear_bottom_buttons()  # Remove any previously built widgets
 
        if mode == 'updater':
 
            self.btn_cancel = QPushButton("CANCEL UPDATE")
            self.btn_cancel.setObjectName("btnCancel")
            self.btn_cancel.setFixedHeight(44)
            self.btn_cancel.setCursor(Qt.PointingHandCursor)
            self.btn_cancel.setEnabled(False)  # Only active during transfer
            self.btn_cancel.clicked.connect(self._on_cancel_clicked)
 
            self.btn_start = QPushButton("START UPDATE")
            self.btn_start.setObjectName("btnStart")
            self.btn_start.setFixedHeight(44)
            self.btn_start.setCursor(Qt.PointingHandCursor)
            self.btn_start.setEnabled(False)  # Enabled once devices are selected
            self.btn_start.clicked.connect(self._on_start)
 
            self._bottom_layout.addWidget(self.btn_cancel, stretch=1)
            self._bottom_layout.addWidget(self.btn_start,  stretch=2)

            # Do not hide the 'Current Version' and 'Target Version' columns
            self.table.setColumnHidden(3, False)
            self.table.setColumnHidden(4, False)

            # Fit columns to the viewport so all are visible without scrolling
            QTimer.singleShot(0, self._fit_columns_for_updater)

            # Gold indicator for Updater mode
            self._mode_label.setText("UPDATER MODE")
            self._mode_label.setStyleSheet(f"""
                color: {GOLD};
                font-size: 9px;
                font-family: 'Courier New', monospace;
                letter-spacing: 2px;
                font-weight: bold;
            """)
 
        elif mode == 'command_pusher':
 
            self.btn_push = QPushButton("PUSH COMMANDS")
            self.btn_push.setObjectName("btnStart")  # Reuse green styling
            self.btn_push.setFixedHeight(44)
            self.btn_push.setCursor(Qt.PointingHandCursor)
            self.btn_push.setEnabled(False)  # Enabled once devices are selected
            self.btn_push.clicked.connect(self._on_push_commands)
 
            self._bottom_layout.addWidget(self.btn_push)

            # Hide the 'Current Version' and 'Target Version' columns
            self.table.setColumnHidden(3, True)
            self.table.setColumnHidden(4, True)

            # Ensure the 2 columns take all the space
            self._distribute_columns_evenly()

            # Silver indicator for Command Pusher mode
            self._mode_label.setText("COMMAND PUSHER")
            self._mode_label.setStyleSheet(f"""
                color: {SILVER};
                font-size: 9px;
                font-family: 'Courier New', monospace;
                letter-spacing: 2px;
                font-weight: bold;
            """)
 
        self._update_button_states()  # Sync enable/disable state immediately
    
    def _clear_bottom_buttons(self):
 
        """
        PURPOSE
        -------
        Remove all widgets from self._bottom_container and clear the
        corresponding instance attribute references and the mode label.
        Called before every _build_bottom_buttons() call and at the
        start of each new scan to ensure a clean slate.
 
 
        ARGUMENTS
        ---------
        None
 
 
        RETURN VALUE
        ------------
        None
        """
 
        # Drain the layout — takeAt(0) is the correct Qt pattern for this
        while self._bottom_layout.count():
            item = self._bottom_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.disconnect()  # Sever all signal connections before deferred deletion
                widget.deleteLater()
 
        # Clear instance references so hasattr guards work correctly
        for attr in ('btn_start', 'btn_cancel', 'btn_push'):
            if hasattr(self, attr):
                delattr(self, attr)
 
        # Reset the mode indicator
        self._mode_label.setText("")
        self._mode_label.setStyleSheet(f"""
            color: {TEXT_MUTED};
            font-size: 9px;
            font-family: 'Courier New', monospace;
            letter-spacing: 2px;
        """)

        # Hide the columns specific for 'Updater' mode
        self.table.setColumnHidden(3, True)
        self.table.setColumnHidden(4, True)

        # Since we are hiding columns 3 and 4, ensure that columns 1 and 2 take all the leftover space evenly
        self._distribute_columns_evenly()


    ### <=== COMMAND PUSHER ACTION (STUB) ===> ###
    def _on_push_commands(self):
 
        """
        PURPOSE
        -------
        Handler for the PUSH COMMANDS button. Collects the selected device
        indices from the table, builds a DataFrame from them, injects the
        device credentials and DataFrame into PushCommandsDialog, and opens
        the dialog. The worker is started from within the dialog itself once
        the user clicks PUSH.
 
 
        ARGUMENTS
        ---------
        None
 
 
        RETURN VALUE
        ------------
        None
        """
 
        # Collect checked device indices from the table
        selected_indices = []
        for row in range(self.table.rowCount()):
            cb = self._get_checkbox(row)
            if cb and cb.isChecked():
                item = self.table.item(row, 1)  # Hostname column carries the index
                if item is not None:
                    selected_indices.append(item.data(Qt.UserRole))
 
        if not selected_indices:
            return  # Nothing selected — should not happen if button gating is correct
        
        if self._valid_devices_df is None:
            return  # No scan has completed yet — should not happen in normal flow, but it prevents potential future bugs
 
        selected_df = self._valid_devices_df.loc[selected_indices]
 
        dlg = PushCommandsDialog(self)
        dlg.set_push_context(selected_df, self._device_username, self._device_password)
        dlg.exec_()  # Blocking — worker runs inside the dialog on its own thread

    ##############################################################

    ### <=== START UPDATE BUTTON ===> ###
    def _on_start(self):

        """
        PURPOSE
        -------
        Triggers the firmware update workflow for selected devices.
        
        Captures the indices of all currently checked device rows from the table,
        maps them back to the original valid_devices_df indices, prompts the user
        to select a transfer mode (Sequential or Threaded), and starts a background
        worker to execute the multi-stage update process.
        
        The UI is locked during the operation, and progress is displayed in the
        table loading area. Upon completion, a summary dialog is shown.
        """

        # Capture selected indices from table
        selected_indices = []

        # Guard — btn_start only exists in Updater mode
        if not hasattr(self, 'btn_start'):
            return
        
        for row in range(self.table.rowCount()):
            chk = self._get_checkbox(row)
            if chk and chk.isChecked():
                item = self.table.item(row, 1)  # Hostname column
                if item is not None:
                    original_idx = item.data(Qt.UserRole)
                    if original_idx is not None:
                        selected_indices.append(original_idx)

        if not selected_indices:
            return  # Should not happen
        
        # Store for cancellation purposes
        self._selected_device_indices = selected_indices

        # Ask user for transfer mode — skip dialog if only 1 device selected (always sequential)
        if len(selected_indices) == 1:
            transfer_mode = "sequential"
        else:
            mode_dialog = TransferModeDialog(self)
            if mode_dialog.exec_() != QDialog.Accepted:
                return  # User cancelled (closed dialog)
            transfer_mode = mode_dialog.selected_mode  # "sequential" or "threaded"

        # Lock UI
        self.btn_show.setEnabled(False)
        self.sheet_input.setEnabled(False)
        self.btn_start.setEnabled(False)
        self.btn_select_all.setEnabled(False)
        self._update_running = True

        # Show loading state
        self.table.setRowCount(0)  # remove all device rows & checkboxes
        self._show_loading_in_table("Preparing update...")

        # Create worker and thread (pass transfer_mode)
        self._update_thread = QThread()
        self._update_worker = UpdateWorker(
            excel_file,
            self.sheet_input.text().strip(),
            self._valid_devices_df,
            selected_indices,
            self._device_username,
            self._device_password,
            transfer_mode
        )
        self._update_worker.moveToThread(self._update_thread)

        # Wire signals
        self._update_thread.started.connect(self._update_worker.run)
        self._update_worker.progress.connect(self._on_update_progress)
        self._update_worker.finished.connect(self._on_update_finished)
        self._update_worker.error.connect(self._on_update_error)
        self._update_worker.warning.connect(self._on_update_warning)
        self._update_worker.finished.connect(self._update_thread.quit)
        self._update_worker.error.connect(self._update_thread.quit)
        self._update_thread.finished.connect(self._update_worker.deleteLater)
        self._update_thread.finished.connect(self._update_thread.deleteLater)

        self._update_thread.start()

    def _on_update_finished(self, summary):

        """
        PURPOSE
        -------
        Called when the UpdateWorker completes successfully.
        
        Hides the loading message in the table, restores all UI elements
        to their idle state, and displays a summary dialog with the
        update results (successes, failures, etc.).
        

        ARGUMENTS
        ---------
        summary (str):
        Human-readable summary of the update process generated by UpdateWorker.
        

        RETURN VALUE
        ------------
        None
        """

        self._hide_loading_in_table()  # Remove loading overlay from table
        self._restore_ui_after_update()  # Re-enable all buttons and inputs
        SummaryDialog("Update Complete", summary, self).exec_()

    def _on_update_error(self, error_msg):

        """
        PURPOSE
        -------
        Called when the UpdateWorker encounters a fatal error.
        
        Hides the loading message, restores UI to its idle state, and
        displays an error dialog with the failure details.
        

        ARGUMENTS
        ---------
        error_msg (str):
        Error message describing what went wrong during the update.
        

        RETURN VALUE
        ------------
        None
        """

        self._hide_loading_in_table()  # Remove loading overlay from table
        self._restore_ui_after_update()  # Re-enable all buttons and inputs
        ErrorDialog("Update Failed", error_msg, self).exec_()  # Show error dialog

    def _on_update_warning(self, warning_msg):
 
        """
        PURPOSE
        -------
        Display a non-fatal warning to the user when the Excel tracker could not
        be updated during the update process. The update workflow continues
        uninterrupted — this is informational only.
 
 
        ARGUMENTS
        ---------
        warning_msg (str):
        Warning message describing which tracker write failed and why.
 
 
        RETURN VALUE
        ------------
        None
        """
 
        dlg = WarningDialog("Excel Tracker Warning", warning_msg, self)
        dlg.setAttribute(Qt.WA_DeleteOnClose)  # Qt cleans up the dialog object when user closes it
        dlg.open() # Non-blocking window-modal — no nested event loop

    def _on_update_progress(self, message):

        """
        PURPOSE
        -------
        Update the loading message and enable the Cancel button only during
        the file transfer stage of the update process.


        ARGUMENTS
        ---------
        message (str):
        The progress message emitted by the UpdateWorker.
        

        RETURN VALUE
        ------------
        None
        """

        # Keep the existing loading text in the table
        self._update_loading_message(message)

        # Enable Cancel button only during the SCP transfer phase
        if hasattr(self, 'btn_cancel'):
            if "Transferring IOS files" in message:
                self.btn_cancel.setEnabled(True)   # user can cancel transfers
            else:
                self.btn_cancel.setEnabled(False)  # no cancel during install, reload, etc.

    def _restore_ui_after_update(self):

        """
        PURPOSE
        -------
        Restore all UI elements to their idle state after an update run
        (whether successful or failed).
        
        Clears the update running flag, re-enables the Show button, input
        field, SELECT ALL button, and disables the Cancel button. Refreshes
        the START button state based on current device selection.
        

        ARGUMENTS
        ---------
        None
        

        RETURN VALUE
        ------------
        None
        """

        self._update_running = False  # Clear the update-in-progress flag
        self.btn_show.setEnabled(True)  # Re-enable SHOW button
        self.sheet_input.setEnabled(True)  # Re-enable Excel sheet input
        self.btn_select_all.setEnabled(True)  # Re-enable SELECT ALL button
        if hasattr(self, 'btn_cancel'):
            self.btn_cancel.setEnabled(False)  # Disable CANCEL UPDATE button
        self._update_show_button_state()  # Refresh SHOW button based on input text
        self._update_button_states()  # Re-evaluate START button based on selection

    ##############################################################
    
    ### <=== CANCEL BUTTON ===> ###
    def _on_cancel_clicked(self):

        """
        PURPOSE
        -------
        Cancel the active SCP file transfers on all selected devices.

        Retrieves the list of devices originally selected for this update run
        (stored in _selected_device_indices), builds a DataFrame from them,
        and executes the real cancellation routine (clear VTY lines + delete
        partially transferred files) via device_cli_ops.cancel_active_transfers_all.

        The Cancel button is disabled immediately to prevent duplicate
        cancellation requests. The UpdateWorker will detect the global
        cancel_event and abort the process after the transfer stage.
        

        ARGUMENTS
        ---------
        None
        

        RETURN VALUE
        ------------
        None
        """

        # Guard — btn_cancel only exists in Updater mode
        if not hasattr(self, 'btn_cancel'):
            return

        # Disable immediately — gives instant visual feedback before the 
        # blocking CLI call begins (which can take several seconds)
        self.btn_cancel.setEnabled(False)

        if not self._update_running:  # Update not in progress
            return

        # Ensure we have a valid list of devices
        if not self._selected_device_indices:
            return

        # Build a DataFrame for the selected devices
        selected_df = self._valid_devices_df.loc[self._selected_device_indices]

        # Execute the cancellation (sets global cancel_event, clears VTY, deletes files)
        device_cli_ops.cancel_active_transfers_all(
            selected_df,
            self._device_username,
            self._device_password
        )

    ##############################################################

    ### <=== CHECKBOX RETRIEVAL ===> ###
    def _get_checkbox(self, row):

        """
        PURPOSE 
        -------
        Retrieves the QCheckBox widget embedded in a specific table row.
        
        Each row's first column (col-0) contains a container QWidget that
        holds a QHBoxLayout with a centered QCheckBox. This method navigates
        that widget hierarchy to extract the checkbox instance.
        
        The container widget is fetched via cellWidget(row, 0). If it exists,
        findChild() searches its children recursively for a QCheckBox object.
        This approach is necessary because the checkbox is not a direct
        QTableWidgetItem but rather a custom widget embedded in the cell.
        

        ARGUMENTS
        ---------
        row (int): Zero-based index of the table row to inspect.
            
        
        RETURN VALUE
        ------------
        QCheckBox or None:
        The checkbox widget if found and the row contains a valid container;
        otherwise None (e.g., empty row, malformed cell, or row index out of bounds).
        """

        container = self.table.cellWidget(row, 0)  # Get the widget in col-0
        if container:
            return container.findChild(QCheckBox)  # Search for checkbox inside
        return None

    ##############################################################

    ### <=== BUTTON STATE MANAGEMENT ===> ###
    def _update_button_states(self):

        """
        PURPOSE
        -------
        Evaluates current UI conditions and enables/disables the action buttons
        (START UPDATE and PUSH COMMANDS) accordingly.
        
        Both action buttons require the same three conditions to be enabled:
            1. The table contains at least one device row (has_devices)
            2. At least one checkbox is currently checked (has_selection)
            3. No update or push operation is currently in progress (not _update_running)
            
        If any condition fails, both buttons remain disabled (dark green).
        
        The CANCEL UPDATE button is managed separately — it is only enabled
        during an active transfer stage (file transfer) and is forcefully disabled
        here when no operation is running, returning it to its dark-red disabled state.
        
        This method is called after any action that modifies device selection,
        table content, or the update running state, including:
            - Toggling individual checkboxes
            - Clicking SELECT ALL
            - Loading device data
            - Completing or cancelling an operation


        ARGUMENTS
        ---------
        None
        

        RETURN VALUE
        ------------
        None
        """

        row_count = self.table.rowCount()
        has_devices = row_count > 0  # At least one device in table
        has_selection = False  # At least one checkbox checked
        
        if has_devices:
            for row in range(row_count):
                chk = self._get_checkbox(row)
                if chk and chk.isChecked():
                    has_selection = True
                    break  # Found at least one — stop searching
        
        # Action buttons require devices AND at least one selected
        enabled = has_devices and has_selection and not self._update_running

        if hasattr(self, 'btn_start'):
            self.btn_start.setEnabled(enabled) # UPDATER mode start
 
        if hasattr(self, 'btn_push'):
            self.btn_push.setEnabled(enabled) # COMMAND PUSHER mode start

        # Cancel button only enabled when update is running
        if not self._update_running:
            if hasattr(self, 'btn_cancel'):
                self.btn_cancel.setEnabled(False)  # Force disable when no operation

    ### <=== SHOW BUTTON STATE MANAGEMENT ===> ###
    def _update_show_button_state(self):

        """
        PURPOSE
        -------
        Controls the enabled state of the SHOW ELIGIBLE DEVICES button.
        
        The SHOW button is enabled only when the sheet input QLineEdit contains
        non-whitespace text. An empty field or one containing only spaces/tabs
        results in the button being disabled and rendered in muted gold/brown
        tones via the :disabled stylesheet state.
        
        This method is connected to the textChanged signal of self.sheet_input,
        meaning it fires on every keystroke in the input field. It is also
        called manually during initialization and after cancellation/completion
        of an update to restore the correct button state based on current
        input content.
        
        The strip() call ensures that purely whitespace input (spaces, tabs,
        newlines) is treated as empty, preventing the user from triggering
        a search with meaningless input.
        

        ARGUMENTS
        ---------
        None
        

        RETURN VALUE
        ------------
        None
        """

        has_text = bool(self.sheet_input.text().strip())  # True if non-whitespace exists
        self.btn_show.setEnabled(has_text)  # Enable only when text present

    ##############################################################

    ### <=== SHOW DEVICES ===> ###
    def _on_show_devices(self):

        """
        PURPOSE
        -------
        Prompt for credentials and operation mode, then load eligible devices
        from the Excel sheet.

        Retrieves the sheet name from the input field, then displays a modal
        credentials dialog for the user to enter device authentication
        credentials (username and password) and select an operation mode
        (UPDATER or COMMAND PUSHER). The password is masked and includes a
        visibility toggle.

        After the dialog is accepted, the credentials and the selected mode are
        stored as instance variables. The bottom action buttons are cleared to
        provide a fresh start before the scan begins.

        Before loading devices, verifies that the Excel file is not currently
        open in another application. If it is locked, displays an error dialog
        and aborts the operation.

        The method then proceeds to load and display the device inventory from
        the specified Excel sheet using a background thread. The selected
        operation mode is passed to the worker so that the subsequent workflow
        can adapt accordingly.
        

        ARGUMENTS
        ---------
        None
        

        RETURN VALUE
        ------------
        None
        """

        sheet_name = self.sheet_input.text().strip()  # Get and clean sheet name
        # if not sheet_name:
        #     return

        # Show Credentials Dialog
        dialog = Credentials_and_mode_Dialog(self)
        if dialog.exec_() != QDialog.Accepted:  # User cancelled
            return

        self._device_username, self._device_password = dialog.get_credentials()  # Store credentials
        self._mode = dialog.get_mode()  # Store the selected operation mode

        # Reset the bottom area before the scan so re-runs start clean
        self._clear_bottom_buttons()

        # Lock UI while working
        self.btn_show.setEnabled(False)
        self.btn_show.setText("SCANNING...")  # Visual feedback
        self.sheet_input.setEnabled(False)

        # Disable buttons that should not be clickable during the scan
        self.btn_select_all.setEnabled(False)

        self.table.setRowCount(0)  # Clear existing table

        # Show loading message in table
        self._show_loading_in_table("Scanning devices...")

        # Create worker and thread for background processing
        self._show_thread = QThread()
        self._show_worker = ShowDevicesWorker(
            excel_file,
            sheet_name,
            self._device_username,
            self._device_password,
            self._mode
        )
        self._show_worker.moveToThread(self._show_thread)  # Move worker to thread

        # Wire signals between thread, worker, and GUI
        self._show_thread.started.connect(self._show_worker.run)  # Start work when thread starts
        self._show_worker.finished.connect(self._on_show_devices_done)  # Success handler
        self._show_worker.error.connect(self._on_show_devices_error)  # Error handler
        self._show_worker.progress.connect(self._update_loading_message)  # Progress updates
        self._show_worker.finished.connect(self._show_thread.quit)  # Stop thread on completion
        self._show_worker.error.connect(self._show_thread.quit)  # Stop thread on error
        self._show_thread.finished.connect(self._show_worker.deleteLater)  # Cleanup worker
        self._show_thread.finished.connect(self._show_thread.deleteLater)  # Cleanup thread

        self._show_thread.start()  # Begin background execution

    def _on_show_devices_done(self, eligible_df, valid_devices_df):

        """
        PURPOSE
        -------
        Handle successful device scan completion.
        
        Hides the loading message, populates the table with eligible devices,
        and restores UI elements to their normal state.
        

        ARGUMENTS
        ---------
        eligible_df (pd.DataFrame): DataFrame containing eligible devices with columns: Hostname,
        OOBM IP Address, Current IOS Version, Recommended IOS Version.

        valid_devices_df (pd.DataFrame): DataFrame containing all valid devices.


        RETURN VALUE
        ------------
        None
        """

        # Hide loading message
        self._hide_loading_in_table()

        # Back on the GUI thread — safe to touch widgets here
        self._eligible_df = eligible_df  # store for use by START UPDATE later
        self._valid_devices_df = valid_devices_df # store for use by START UPDATE later
        
        # Populate table with results
        self._populate_table(eligible_df)

        # Restore UI
        self.btn_show.setText("SHOW ELIGIBLE DEVICES")
        self.sheet_input.setEnabled(True)
        self.btn_select_all.setEnabled(True)
        self._update_show_button_state()  # Refresh based on input content
        
        # Build the correct bottom buttons for the active mode
        self._build_bottom_buttons(self._mode)

    def _on_show_devices_error(self, message):

        """
        PURPOSE
        -------
        Handle device scan failure.
        
        Hides the loading message, restores UI elements, and displays
        an error dialog with the failure reason.
        

        ARGUMENTS
        ---------
        message (str): Error message describing what went wrong during the scan.


        RETURN VALUE
        ------------
        None
        """
        
        # Hide loading message
        self._hide_loading_in_table()

        # Restore UI
        self.btn_show.setText("SHOW ELIGIBLE DEVICES")
        self.btn_select_all.setEnabled(True)
        self.sheet_input.setEnabled(True)
        self._update_show_button_state()  # Refresh based on input content
        
        # Show error dialog
        ErrorDialog("Scan Failed", message, self).exec_()

    def _populate_table(self, eligible_df):
        
        """
        PURPOSE
        -------
        Populate the table widget with device data from the eligible DataFrame.
        
        Replaces the demo data loading with real device information retrieved
        from the Excel scan. Each row receives a checkbox in column 0 and
        device details in subsequent columns.
        
        The original DataFrame index is stored as UserRole data on the Hostname
        cell so that action button can correctly map back to the source
        valid_devices_df when initiating firmware updates or pushing commands.
        

        ARGUMENTS
        ---------
        eligible_df (pd.DataFrame): Filtered DataFrame containing only devices eligible for update.
        

        RETURN VALUE
        ------------
        None
        """
        
        self.table.setRowCount(len(eligible_df))  # Set table size

        for row_idx, (_, row) in enumerate(eligible_df.iterrows()):
            # Create checkbox for device selection
            chk = QCheckBox()
            chk.stateChanged.connect(self._update_button_states)  # Refresh button states on toggle
            container = QWidget()
            lay = QHBoxLayout(container)
            lay.addWidget(chk)
            lay.setAlignment(Qt.AlignCenter)  # Center checkbox in cell
            lay.setContentsMargins(0, 0, 0, 0)
            container.setStyleSheet("background: transparent;")
            self.table.setCellWidget(row_idx, 0, container)  # Place in col-0

            # Populate data columns
            for col_idx, value in enumerate([
                row['Hostname'],
                row['OOBM IP Address'],
                str(row['Current IOS Version']),
                str(row['Recommended IOS Version'])
            ], start=1):
                item = QTableWidgetItem(value)
                item.setTextAlignment(Qt.AlignVCenter | Qt.AlignLeft)
                self.table.setItem(row_idx, col_idx, item)
            
            # 'eligible_df' is a filtered subset of 'valid_devices_df'. After all the eligibility checks, 
            # only some rows survive. The DataFrame index preserves the original row numbers from 'valid_devices_df':
            #     eligible_df index:  3,  7, 11
            #     row_idx:            0,  1,  2
            # When START UPDATE runs, it needs to look up devices back in 'valid_devices_df' using that 
            # stored index. If we stored 'row_idx' instead of the real index, we'd be looking up rows 
            # 0, 1, 2 in 'valid_devices_df' — the wrong devices entirely.
            
            # Store the DataFrame index on the row so START UPDATE can retrieve it
            self.table.item(row_idx, 1).setData(Qt.UserRole, _)  # _ is the original DataFrame index

    ##############################################################

    ### <=== TABLE LOADING STATE ===> ###
    def _show_loading_in_table(self, message="Loading devices..."):

        """
        PURPOSE
        -------
        Display a centered loading message in the empty table.
        
        Creates a three-row table structure with spacer rows above and below
        a centered message row. The message spans all five columns and displays
        a gold-colored loading indicator with the provided text.
        
        The table is disabled during loading to prevent user interaction.
        The loading label is stored as an instance variable (_loading_label)
        for subsequent updates via _update_loading_message().
        

        ARGUMENTS
        ---------
        message (str): The loading message to display. Defaults to "Loading devices...".

        
        RETURN VALUE
        ------------
        None
        """

        # Create 3 rows: top spacer, message, bottom spacer
        self.table.setRowCount(3)
        
        # Hide vertical headers (row numbers)
        self.table.verticalHeader().setVisible(False)
        
        # Set row heights to position the message in the upper portion of the table
        self.table.setRowHeight(0, 120)  # Top spacer — pushes message down
        self.table.setRowHeight(1, 60)   # Message row — contains the loading text
        self.table.setRowHeight(2, 200)  # Bottom spacer — fills remaining space
        
        # Make spacer rows empty and non-selectable
        for spacer_row in [0, 2]:
            for col in range(5):
                empty_item = QTableWidgetItem("")
                empty_item.setFlags(Qt.NoItemFlags)  # Cannot be selected or edited
                self.table.setItem(spacer_row, col, empty_item)
        
        # Span the message across all 5 columns (col-0 through col-4)
        self.table.setSpan(1, 0, 1, 5)
        
        # Create container widget to hold the loading label
        container = QWidget()
        layout = QHBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setAlignment(Qt.AlignCenter)  # Center horizontally in the cell
        
        # Loading label - store as instance variable for easy updates during progress
        self._loading_label = QLabel(f"⏳ {message}")
        self._loading_label.setStyleSheet(f"""
            color: {GOLD_LIGHT};
            font-size: 16px;
            font-family: 'Courier New', monospace;
            font-weight: bold;
            letter-spacing: 2px;
        """)
        self._loading_label.setAlignment(Qt.AlignCenter)
        self._loading_label.setObjectName("loadingLabel")  # Give it a name for findChild
        layout.addWidget(self._loading_label)
        
        self.table.setCellWidget(1, 0, container)  # Place container in message row
        
        # Disable interactions during loading
        self.table.setEnabled(False)
        
    def _hide_loading_in_table(self):

        """
        PURPOSE
        -------
        Remove the loading message and restore table interactivity.
        
        Clears the table, removes all cell spans, and re-enables user
        interaction. Called when the device scan completes (success or error).
        

        ARGUMENTS
        ---------
        None
        

        RETURN VALUE
        ------------
        None
        """

        self.table.setEnabled(True)  # Re-enable user interaction
        self.table.setRowCount(0)  # Clear all rows

        # Clear any cell spans (resets the 5-column span from loading)
        self.table.clearSpans()

        self.table.verticalHeader().setVisible(False)  # Already hidden by default

    def _update_loading_message(self, message):

        """
        PURPOSE
        -------
        Update the loading message text without recreating the table structure.
        
        Connected to the ShowDevicesWorker.progress signal. Updates the text
        of the existing loading label to provide feedback on the current
        scanning stage (e.g., "Checking Excel file...", "Reading inventory...").
        

        ARGUMENTS
        ---------
        message (str):
        The new loading message to display.
        
        
        RETURN VALUE
        ------------
        None
        """
        
        if hasattr(self, '_loading_label') and self._loading_label:
            self._loading_label.setText(f"⏳ {message}")  # Update with hourglass icon

###################################################################################################################################

# WORKER: SHOW ELIGIBLE DEVICES
# -----------------------------
class ShowDevicesWorker(QObject):

    # Signals back to the GUI thread
    finished    = pyqtSignal(object, object) # (eligible_df, valid_devices_df)
    error       = pyqtSignal(str) # emits an error message string on failure
    progress    = pyqtSignal(str) # emits progress messages

    def __init__(self, excel_file, sheet_name, username, password, mode):
        super().__init__()
        self.excel_file  = excel_file
        self.sheet_name  = sheet_name
        self.username    = username
        self.password    = password
        self.mode        = mode  # 'updater' or 'command_pusher'

    def run(self):
        try:
            

            #-------------------------------------------
            logger.info("SHOW ELIGIBLE DEVICES clicked")
            #-------------------------------------------


            ### <=== CHECK IF EXCEL FILE OPEN ===> ###
            self.progress.emit("Checking Excel file...")
            QThread.msleep(100)

            result = excel_and_data_ops.check_excel_file_not_open(self.excel_file)
            if result == "ERROR_EXCEL_FILE_OPEN":
                self.error.emit(f"❌ EXCEL_FILE_OPEN\n\nThe Excel file {self.excel_file} is currently open by\nanother process.\n\nPlease close it and try again.")
                return


            ### <=== CHECK EXCEL SHEET NAME ===> ###
            all_devices_df = excel_and_data_ops.create_devices_dataframe(
                self.excel_file, self.sheet_name
            )
            
            # Check for error codes returned as strings
            if isinstance(all_devices_df, str):
                # Pass the raw error code to the GUI for custom handling
                if all_devices_df == "FILE_NOT_FOUND_ERROR":
                    self.error.emit(f"❌ FILE NOT FOUND\n\nThe Excel file could not be located")
                elif all_devices_df == "COLUMN_NOT_FOUND_ERROR":
                    self.error.emit(f"❌ MISSING COLUMN\n\nThe Excel sheet '{self.sheet_name}' is missing the required 'OOBM IP Address' column")
                elif all_devices_df == "SHEET_NOT_FOUND_ERROR":
                    self.error.emit(f"❌ SHEET NOT FOUND\n\nThe worksheet '{self.sheet_name}' does not exist in the Excel file.\n\nPlease verify the sheet name and try again.")
                elif all_devices_df == "UNEXPECTED_ERROR":
                    self.error.emit(f"❌ UNEXPECTED ERROR\n\nAn error occurred while reading the Excel file.\n\nSheet: {self.sheet_name}\n\nPlease verify the file is not corrupted and the sheet name is correct.")
                else:
                    self.error.emit(f"❌ ERROR\n\nCould not read Excel file: {all_devices_df}")
                return
            
            # Ensure the returned value is a DataFrame 
            if not isinstance(all_devices_df, pd.DataFrame):
                self.error.emit(f"Could not read Excel file: {all_devices_df}")
                return
            
            # If the DataFrame is missing columns, the return value is NONE
            valid_devices_df = excel_and_data_ops.valid_devices_dataframe(all_devices_df)
            if valid_devices_df is None:
                self.error.emit("❌ MISSING COLUMNS\n\nThe Excel sheet is missing one or more required tracking columns.\n\nPlease ensure the sheet contains all expected columns or use the provided template.")
                return


            ### <=== ZERO THE EXCEL TRACKER ===> ###
            self.progress.emit("Zeroing the EXCEL tracker...")
            QThread.msleep(100)
            if not self._update_tracker(valid_devices_df):
                return


            ### <=== CHECK DEVICE STATUS FOR ONLINE AND AUTHENTICATION ===> ###
            self.progress.emit("Checking ONLINE status and authentication...")
            QThread.msleep(100)

            result = excel_and_data_ops.populate_status_and_auth_status_column(
                valid_devices_df, self.username, self.password
            )

            # Handle errors...
            if result is not None:
                if result == "NO_DEVICES_ERROR":
                    self.error.emit(f"❌ NO DEVICES FOUND\n\nThe Excel sheet contains no devices to process.\n\nPlease verify the sheet has valid device entries.")
                elif result == "CONNECTION_TIMEOUT_ERROR":
                    self.error.emit(f"❌ CONNECTION FAILED\n\nAll devices are offline or unreachable.\n\nPlease verify:\n• Devices are powered on\n• OOBM IP addresses are correct\n• Network connectivity to devices")
                elif result == "AUTH_FAILED_ERROR":
                    self.error.emit(f"❌ AUTHENTICATION FAILED\n\nAll online devices rejected the provided credentials.\n\nPlease verify:\n• Username is correct\n• Password is correct\n• Account has appropriate privileges")
                elif result == "UNEXPECTED_ERROR":
                    self.error.emit(f"❌ UNEXPECTED ERROR\n\nAn error occurred while scanning device status.\n\nPlease check the logs for details.")
                else:
                    self.error.emit(f"❌ ERROR\n\nDevice scan failed: {result}")
                return


            ### <=== COMMAND PUSHER: EARLY EXIT ===> ###
            # In Command Pusher mode the pipeline stops here — no RESTCONF check,
            # no version retrieval, no flash check. Devices shown are ONLINE + AUTH_OK only.
            if self.mode == 'command_pusher':
                online_authok_df = valid_devices_df[
                    (valid_devices_df['Status']      == 'ONLINE') &
                    (valid_devices_df['Auth Status'] == 'AUTH_OK')
                ].copy()
                #------------------------------------------------------------------------------------------
                logger.info(f"Command Pusher mode: {len(online_authok_df)} ONLINE/AUTH_OK device(s) ready")
                #------------------------------------------------------------------------------------------
                self.finished.emit(online_authok_df, valid_devices_df)
                return


            ### <=== CHECK DEVICE RESTCONF STATUS ===> ###
            self.progress.emit("Checking RESTCONF Status on devices...")
            QThread.msleep(100)

            result = excel_and_data_ops.populate_restconf_status_column(
                valid_devices_df, self.username, self.password, 30, 7
            )
            
            if result is not None:                   
                if result == "RESTCONF_TIMEOUT_ERROR":
                    self.error.emit(
                        f"❌ RESTCONF FAILED\n\n"
                        f"All eligible devices failed to establish RESTCONF connectivity.\n\n"
                        f"Possible causes:\n"
                        f"• RESTCONF not enabled on devices\n"
                        f"• HTTPS/RESTCONF configuration missing\n"
                        f"• Firewall blocking port 443\n\n"
                        f"The update cannot proceed without RESTCONF access."
                    )
                    return
                    
                elif result == "UNEXPECTED_ERROR":
                    self.error.emit(
                        f"❌ UNEXPECTED ERROR\n\n"
                        f"An error occurred during the RESTCONF status check.\n\n"
                        f"Please check the logs for details."
                    )
                    return
                    
                else:
                    self.error.emit(f"❌ ERROR\n\nRESTCONF check failed: {result}")
                    return


            ### <=== CHECK SCP SERVER STATUS ===> ###
            self.progress.emit("Checking SCP server status on devices...")
            QThread.msleep(100)
 
            result = excel_and_data_ops.populate_scp_status_column(
                valid_devices_df, self.username, self.password
            )
 
            if result is not None:
                if result == "ALL_DEVICES_FAILED_ERROR":
                    self.error.emit(
                        f"❌ SCP DISABLED ON ALL DEVICES\n\n"
                        f"SCP server is not enabled on any device.\n\n"
                        f"Please enable SCP on the devices before running an update:\n"
                        f"  ip scp server enable"
                    )
                    return
 
                elif result == "UNEXPECTED_ERROR":
                    self.error.emit(
                        f"❌ UNEXPECTED ERROR\n\n"
                        f"An error occurred while checking SCP status.\n\n"
                        f"Please check the logs for details."
                    )
                    return
 
                else:
                    self.error.emit(f"❌ ERROR\n\nSCP status check failed: {result}")
                    return
                

            ### <=== RETRIEVE CURRENT IOS VERSION ===> ###
            self.progress.emit("Retrieving the current IOS version...")
            QThread.msleep(100)

            result = excel_and_data_ops.populate_current_version_column(
                valid_devices_df, self.username, self.password
            )
            
            # Handle errors...
            if result is not None:
                if result == "NO_OPERATIVE_DEVICES_ERROR":
                    self.error.emit(
                        f"❌ NO OPERATIVE DEVICES\n\n"
                        f"No devices with operative RESTCONF available for version retrieval.\n\n"
                        f"This may be because:\n"
                        f"• All devices are offline\n"
                        f"• Authentication failed\n"
                        f"• RESTCONF is not operative\n\n"
                        f"The operation cannot proceed without at least one operative device."
                    )
                    return
                    
                elif result == "VERSION_RETRIEVAL_FAILED_ERROR":
                    self.error.emit(
                        f"❌ VERSION RETRIEVAL FAILED\n\n"
                        f"All operative devices failed to return their current IOS version.\n\n"
                        f"Possible causes:\n"
                        f"• RESTCONF endpoint not responding properly\n"
                        f"• Device model doesn't support the required YANG model\n"
                        f"• Network instability during version query\n\n"
                        f"Please check device connectivity and try again."
                    )
                    return
                    
                elif result == "UNEXPECTED_ERROR":
                    self.error.emit(
                        f"❌ UNEXPECTED ERROR\n\n"
                        f"An error occurred while retrieving IOS versions.\n\n"
                        f"Please check the logs for details."
                    )
                    return
                    
                else:
                    self.error.emit(f"❌ ERROR\n\nVersion retrieval failed: {result}")
                    return


            ### <=== DETERMINE IF AN UPDATE IS NEEDED ===> ###
            self.progress.emit("Determining which devices need an update...")
            QThread.msleep(100)

            result = excel_and_data_ops.populate_needs_update_column(valid_devices_df)
            
            if result is not None:
                if result == "NO_DEVICES_NEED_UPDATE":
                    self.error.emit(
                        f"❌ NO DEVICES ELIGIBLE\n\n"
                        f"All the devices are up to date.\n\n"
                        f"Please check the recommended version of the devices."
                    )
                    return
   
                elif result == "MISSING_VERSION_DATA_ERROR":
                    self.error.emit(
                        f"❌ MISSING VERSION DATA\n\n"
                        f"No devices have both current and recommended IOS version data.\n\n"
                        f"Possible causes:\n"
                        f"• Current version retrieval failed for all devices\n"
                        f"• Recommended version column is empty\n"
                        f"• Version data is incomplete\n\n"
                        f"Please verify the Excel sheet has valid version information."
                    )
                    return
                    
                elif result == "UNEXPECTED_ERROR":
                    self.error.emit(
                        f"❌ UNEXPECTED ERROR\n\n"
                        f"An error occurred while checking for needed updates.\n\n"
                        f"Please check the logs for details."
                    )
                    return
                    
                else:
                    self.error.emit(f"❌ ERROR\n\nUpdate check failed: {result}")
                    return
            

            ### <=== VERIFY THE IOS IMAGE FILE PATH IS EXISTENT ===> ###
            self.progress.emit("Verifying if the IOS image file path exists...")
            QThread.msleep(100)

            result = excel_and_data_ops.valid_devices_df_with_image_path(valid_devices_df)
            
            if result is not None:
                if result == "IOS_REPOSITORY_NOT_FOUND_ERROR":
                    self.error.emit(
                        f"❌ IOS REPOSITORY NOT FOUND\n\n"
                        f"The IOS image repository folder could not be located.\n\n"
                        f"Expected location:\n"
                        f"  .../ios_repository/\n\n"
                        f"Please verify that:\n"
                        f"• The 'ios_repository' folder exists in the application directory\n"
                        f"• The folder contains model-specific subfolders with IOS images"
                    )
                    return
                    
                elif result == "NO_IMAGE_FOUND_ERROR":
                    self.error.emit(
                        f"❌ NO IOS IMAGES FOUND\n\n"
                        f"No matching IOS images were found for any device needing an update.\n\n"
                        f"Please verify that:\n"
                        f"• The 'ios_repository' folder contains subfolders for each device model\n"
                        f"• Each model folder contains the required IOS image file\n"
                        f"• The image filename includes the recommended version string\n\n"
                        f"Check the logs for specific devices and missing versions."
                    )
                    return
                    
                elif result == "UNEXPECTED_ERROR":
                    self.error.emit(
                        f"❌ UNEXPECTED ERROR\n\n"
                        f"An error occurred while locating IOS image files.\n\n"
                        f"Please check the logs for details."
                    )
                    return
                    
                else:
                    self.error.emit(f"❌ ERROR\n\nImage path lookup failed: {result}")
                    return


            ### <=== GET THE SIZE OF THE IOS IMAGE FILE ===> ###
            self.progress.emit("Retrieving size of the IOS image file...")
            QThread.msleep(100)

            result = excel_and_data_ops.get_image_files_size(valid_devices_df)
            
            if result is not None:
                    
                if result == "NO_VALID_IMAGE_FILES_ERROR":
                    self.error.emit(
                        f"❌ NO VALID IMAGE FILES\n\n"
                        f"None of the IOS image files for devices needing update could be read.\n\n"
                        f"Possible causes:\n"
                        f"• Image files are missing or inaccessible\n"
                        f"• File permissions prevent reading\n"
                        f"• Files were moved or deleted after path lookup\n\n"
                        f"Please verify the IOS image files exist and are readable."
                    )
                    return
                    
                elif result == "UNEXPECTED_ERROR":
                    self.error.emit(
                        f"❌ UNEXPECTED ERROR\n\n"
                        f"An error occurred while retrieving IOS image file sizes.\n\n"
                        f"Please check the logs for details."
                    )
                    return
                    
                else:
                    self.error.emit(f"❌ ERROR\n\nImage size retrieval failed: {result}")
                    return


            ### <=== GET AND SET THE FLASH FREE SPACE ===> ###
            self.progress.emit("Checking device's FLASH memory free space...")
            QThread.msleep(100)

            result = excel_and_data_ops.populate_flash_free_space_column(
                valid_devices_df, self.username, self.password
            )
            
            if result is not None:
                if result == "ALL_DEVICES_PARSE_ERROR":
                    self.error.emit(
                        f"❌ CLI PARSING FAILED - ALL ELIGIBLE DEVICES\n\n"
                        f"Connected to devices but could not parse flash free space from CLI output.\n\n"
                        f"Possible causes:\n"
                        f"• 'show file systems' command output format unexpected\n"
                        f"• Devices running unsupported IOS version\n"
                        f"• Flash filesystem naming differs from expected\n\n"
                        f"Please check device compatibility."
                    )
                    return
                    
                elif result == "ALL_DEVICES_FAILED_ERROR":
                    self.error.emit(
                        f"❌ FLASH SPACE RETRIEVAL FAILED\n\n"
                        f"All eligible devices failed to return flash free space information.\n\n"
                        f"Please check the logs for specific failure reasons per device."
                    )
                    return
                    
                elif result == "UNEXPECTED_ERROR":
                    self.error.emit(
                        f"❌ UNEXPECTED ERROR\n\n"
                        f"An error occurred while retrieving flash free space.\n\n"
                        f"Please check the logs for details."
                    )
                    return
                    
                else:
                    self.error.emit(f"❌ ERROR\n\nFlash space retrieval failed: {result}")
                    return


            ### <=== DETERMINE IF DEVICES HAVE ENOUGH FLASH FREE SPACE ===> ###
            self.progress.emit("Determining if devices have enough FLASH memory free space for the update...")
            QThread.msleep(100)

            # No need to handle errors, because no errors can occur here.
            excel_and_data_ops.populate_enough_space_column(valid_devices_df)


            ### <=== UPDATE THE EXCEL TRACKER ===> ###
            self.progress.emit("Updating the EXCEL tracker...")
            QThread.msleep(100)
            if not self._update_tracker(valid_devices_df):
                return


            ### <=== GET ELIGIBLE DEVICES FOR UPDATE ===> ###
            self.progress.emit("Getting eligible devices...")
            eligible_devices_df = excel_and_data_ops.get_eligible_devices_df(valid_devices_df)
            QThread.msleep(100)

            if isinstance(eligible_devices_df, str) and eligible_devices_df == "ELIGIBLE_DEVICES_EMPTY":
                self.error.emit(
                    f"❌ NO ELIGIBLE DEVICES\n\n"
                    f"No devices meet all the update requirements.\n\n"
                    f"Possible reasons:\n"
                    f"• Device is OFFLINE or AUTH_BAD\n"
                    f"• RESTCONF is NOT_OPERATIVE\n"
                    f"• SCP is not enabled\n"
                    f"• Current version could not be retrieved\n"
                    f"• Not enough flash space\n"
                    f"• No update needed\n"
                    f"• IOS image file not found\n\n"
                    f"Please check the Excel tracker for per-device details."
                )
                return
            
            self.finished.emit(eligible_devices_df, valid_devices_df)


        except Exception as e:
            self.error.emit(str(e))

    def _update_tracker(self, valid_devices_df):
 
        """
        PURPOSE
        -------
        Calls update_excel_tracker and handles every possible error code by
        emitting the appropriate error signal. The caller checks the return
        value and does 'return' if False, halting the scan.
 
 
        ARGUMENTS
        ---------
        valid_devices_df (pd.DataFrame): Current state of the devices DataFrame to be written to the Excel file.
 
 
        RETURN VALUE
        ------------
        True  — Excel file was updated successfully.
        False — A failure occurred; error signal already emitted.
        """
 
        result = excel_and_data_ops.update_excel_tracker(
            self.excel_file, self.sheet_name, valid_devices_df
        )
 
        if result == "SUCCESS":
            return True
 
        if result == "MISSING_COLUMN_EXCEL_ERROR":
            self.error.emit(
                f"❌ MISSING COLUMNS IN EXCEL\n\n"
                f"The Excel sheet '{self.sheet_name}' is missing one or more required tracking columns.\n\n"
                f"Please ensure the sheet contains all expected columns:\n"
                f"• Hostname\n"
                f"• Current IOS Version\n"
                f"• Status\n"
                f"• Auth Status\n"
                f"• Enough Flash Space\n"
                f"• Needs Update\n"
                f"• Update IOS File Present\n"
                f"• Transfer Result\n"
                f"• Install Status\n"
                f"• Update Result\n"
                f"• Cleaned Inactive\n\n"
                f"Use the provided template or add the missing columns."
            )
 
        elif result == "PERMISSION_DENIED_ERROR":
            self.error.emit(
                f"❌ PERMISSION DENIED\n\n"
                f"Cannot save changes to the Excel file.\n\n"
                f"Please ensure:\n"
                f"• The Excel file is CLOSED (not open in Excel)\n"
                f"• You have write permissions to the file\n"
                f"• The file is not marked as read-only\n\n"
                f"Close Excel and try again."
            )
 
        elif result == "UNEXPECTED_ERROR":
            self.error.emit(
                f"❌ UNEXPECTED ERROR\n\n"
                f"An error occurred while updating the Excel tracker.\n\n"
                f"Please check the logs for details."
            )
 
        else:
            self.error.emit(
                f"❌ ERROR\n\n"
                f"Excel tracker update failed: {result}"
            )
 
        return False
    

# WORKER: PUSH COMMANDS
# ---------------------
class PushCommandsWorker(QObject):
 
    """
    Background worker that pushes a list of exec-mode CLI commands to all
    selected devices concurrently. Runs on a separate QThread to keep the
    GUI responsive during execution.
    """
 
    finished = pyqtSignal(list) # Emitted with list[(index, result)] when all devices processed
    error    = pyqtSignal(str) # Emitted on an unhandled exception
 
    def __init__(self, selected_df, username, password, commands):
 
        """
        PURPOSE
        -------
        Initialize the worker with the device DataFrame, credentials, and
        the list of commands to push.
 
 
        ARGUMENTS
        ---------
        selected_df (pd.DataFrame): Devices selected by the user in the table.
        username    (str):          Device SSH username.
        password    (str):          Device SSH password.
        commands    (list[str]):    Ordered list of exec-mode CLI commands.
 
 
        RETURN VALUE
        ------------
        None
        """
 
        super().__init__()
        self.selected_df = selected_df
        self.username    = username
        self.password    = password
        self.commands    = commands
 
    def run(self):
 
        """
        PURPOSE
        -------
        Execute the command push across all selected devices and emit finished
        with the results list when complete. Any unhandled exception emits the
        error signal instead.


        ARGUMENTS
        ---------
        None


        RETURN VALUE
        ------------
        None
        """
 
        try:
            results = device_cli_ops.push_commands_all(
                self.selected_df,
                self.username,
                self.password,
                self.commands
            )
            self.finished.emit(results)
 
        except Exception as e:
            self.error.emit(str(e))


# WORKER: START UPDATE PROCESS
# ----------------------------
class UpdateWorker(QObject):

    """
    Background worker that orchestrates the multi-stage firmware update
    workflow for the devices selected by the user. Runs on a separate thread
    to keep the GUI responsive.
    
    Stages:
        1. SCP file transfer of IOS images.
        2. Install activation and reload.
        3. Post-reload verification and commit.
        4. Cleanup of inactive packages.
    
    Progress, completion, and error signals are emitted to the main window.
    """

    finished = pyqtSignal(str)   # Emits final summary message on success
    error    = pyqtSignal(str)   # Emits error message on fatal failure
    progress = pyqtSignal(str)   # Emits stage descriptions
    warning  = pyqtSignal(str)   # Emits non-fatal warning (e.g. Excel tracker write failure)

    def __init__(self, excel_file, sheet_name, valid_devices_df, selected_indices, username, password, transfer_mode):

        """
        PURPOSE
        -------
        Initialize the worker with all data required for the update process.
        The worker operates on a copy of the valid devices DataFrame to avoid
        side effects on the original stored in the main window.
        

        ARGUMENTS
        ---------
        excel_file (str): Absolute path to the Excel (.xlsx) file.
        
        sheet_name (str): Name of the worksheet to update.
        
        valid_devices_df (pd.DataFrame): Full device DataFrame containing all valid devices and their status columns.
        
        selected_indices (list): List of original DataFrame indices selected by the user for update.
        
        username (str): Device SSH/API username.
        
        password (str): Device SSH/API password.
        
        transfer_mode (str): 'sequential' or 'threaded' — chosen by the user before starting the update.
        

        RETURN VALUE
        ------------
        None
        """

        super().__init__()
        self.excel_file = excel_file
        self.sheet_name = sheet_name
        self.valid_devices_df = valid_devices_df.copy()  # Work on a copy to avoid side effects
        self.selected_indices = selected_indices
        self.username = username
        self.password = password
        self.transfer_mode = transfer_mode    # "sequential" or "threaded"

    def run(self):

        """
        PURPOSE
        -------
        Execute the update workflow sequentially, updating the Excel tracker
        after each major stage. Emits progress signals and a final summary.
        
        Stages:
            - PART 3: SCP file transfer of IOS images.
            - PART 4: Install activation and reload.
            - PART 5: Post-reload verification and commit.
            - PART 6: Cleanup of inactive packages.
        
        If no devices achieve a successful transfer, the remaining stages
        are skipped and a summary is generated immediately.
        

        ARGUMENTS
        ---------
        None
        

        RETURN VALUE
        ------------
        None — emits either finished(str) or error(str) signal.
        """

        try:
            
            # Reset cancellation flag from any previous run
            device_cli_ops.cancel_event.clear()

            ### <=== PREPARE SELECTED DEVICES DATAFRAME ===> ###
            selected_df = self.valid_devices_df.loc[self.selected_indices].copy()
            total_selected = len(selected_df)
            #----------------------------------------------------------------
            logger.info(f"START UPDATE: {total_selected} device(s) selected")
            #----------------------------------------------------------------


            ### <=== IOS FILE TRANSFER ===> ###
            self.progress.emit(f"Transferring IOS files to {total_selected} device(s) via SCP...")

            QThread.msleep(100) # Brief pause to let GUI update
            result = excel_and_data_ops.populate_transfer_status_column(
                self.valid_devices_df,
                selected_df,
                self.username,
                self.password,
                self.transfer_mode # Pass the user's choice from the dialog
            )

            self._update_tracker() # Save progress to Excel

            # Check if the user pressed Cancel during the transfer
            if device_cli_ops.cancel_event.is_set():
                self.error.emit("Transfer cancelled by user.")
                return

            # Handle fatal transfer errors
            if result is not None:
                if result == "ALL_TRANSFERS_FAILED_ERROR":
                    self.error.emit(
                        f"❌ ALL TRANSFERS FAILED\n\n"
                        f"All {total_selected} selected device(s) failed the IOS file transfer.\n\n"
                        f"Possible causes:\n"
                        f"• SCP/SSH connectivity issues\n"
                        f"• Insufficient flash space\n"
                        f"• Authentication failures\n"
                        f"• Image file not found on disk\n\n"
                        f"Please check the logs for per-device details."
                    )
                    return
                elif result == "UNEXPECTED_ERROR":
                    self.error.emit(
                        f"❌ UNEXPECTED ERROR\n\n"
                        f"An error occurred during the file transfer stage.\n\n"
                        f"Please check the logs for details."
                    )
                    return
                else:
                    self.error.emit(f"❌ ERROR\n\nFile transfer failed: {result}")
                    return


            ### <=== INSTALL ACTIVATION AND RELOAD ===> ###
            install_eligible_df = excel_and_data_ops.get_install_eligible_devices_df(self.valid_devices_df)
            if install_eligible_df.empty:
                # This should not happen if transfers succeeded — treat as fatal
                self.error.emit(
                    f"❌ NO DEVICES ELIGIBLE FOR INSTALL\n\n"
                    f"All selected devices failed the file transfer stage.\n\n"
                    f"Please check the Excel tracker for per-device transfer results."
                )
                return

            else:
                self.progress.emit("Pushing boot commands and triggering reload...")
                QThread.msleep(100)
                result = excel_and_data_ops.populate_install_status_column(
                self.valid_devices_df,
                install_eligible_df,
                self.username,
                self.password,
                skip_confirmation=True
            )
            self._update_tracker()

            if result is not None:
                if result == "ALL_INSTALLS_FAILED_ERROR":
                    self.error.emit(
                        f"❌ ALL INSTALLS FAILED\n\n"
                        f"All devices that received the IOS file failed to trigger the install.\n\n"
                        f"Possible causes:\n"
                        f"• 'install add' or 'install activate' command failed\n"
                        f"• Device connectivity lost during install\n"
                        f"• Authentication failure\n\n"
                        f"Check the Excel tracker for per-device Install Status details."
                    )
                elif result == "UNEXPECTED_ERROR":
                    self.error.emit(
                        f"❌ UNEXPECTED ERROR\n\n"
                        f"An unexpected error occurred during the install process.\n\n"
                        f"Please check the logs for details."
                    )
                else:
                    self.error.emit(f"❌ ERROR\n\nInstall activation failed: {result}")
                return
            

            ### <=== POST-INSTALL: VERSION VERIFICATION AND COMMIT ===> ###
            self.progress.emit("Waiting for reload and verifying installed version...")

            QThread.msleep(100)
            result = excel_and_data_ops.populate_post_install_columns(
                self.valid_devices_df,
                install_eligible_df,
                self.username,
                self.password
            )
            self._update_tracker()

            if result is not None:
                if result == "ALL_DEVICES_TIMEOUT_ERROR":
                    self.error.emit(
                        f"❌ ALL DEVICES TIMED OUT\n\n"
                        f"No devices came back online after the reload.\n\n"
                        f"Possible causes:\n"
                        f"• Devices stuck in boot loop\n"
                        f"• Network connectivity lost\n"
                        f"• Reload taking longer than expected\n\n"
                        f"Check the Excel tracker for per-device status."
                    )
                elif result == "UNEXPECTED_ERROR":
                    self.error.emit(
                        f"❌ UNEXPECTED ERROR\n\n"
                        f"An error occurred during post-install verification.\n\n"
                        f"Please check the logs for details."
                    )
                else:
                    self.error.emit(f"❌ ERROR\n\nPost-install failed: {result}")
                
                return


            ### <=== CLEANUP INACTIVE PACKAGES ===> ###
            self.progress.emit("Cleaning inactive packages from devices...")

            QThread.msleep(100)
            result = excel_and_data_ops.populate_cleaned_inactive_column(
                self.valid_devices_df,
                self.username,
                self.password
            )
            self._update_tracker()

            if result is not None:
                if result == "ALL_CLEANUP_FAILED_ERROR":
                    self.error.emit(
                        f"⚠️ CLEANUP FAILED FOR ALL DEVICES\n\n"
                        f"The inactive package removal could not be completed on any device.\n"
                        f"This does not affect the IOS upgrade itself, but old packages\n"
                        f"may remain on the devices.\n\n"
                        f"Please check the Excel tracker for per‑device details."
                    )
                elif result == "UNEXPECTED_ERROR":
                    self.error.emit(
                        f"❌ UNEXPECTED ERROR\n\n"
                        f"An unexpected error occurred during cleanup.\n\n"
                        f"Please check the logs for details."
                    )
                else:
                    self.error.emit(f"❌ ERROR\n\nCleanup failed: {result}")
                return


            ### <=== FINAL SUMMARY ===> ###
            summary = self._generate_summary()
            self.finished.emit(summary)

        except Exception as e:
            logger.exception("UpdateWorker encountered an unexpected error")
            self.error.emit(str(e))

    def _update_tracker(self):
 
        """
        PURPOSE
        -------
        Save the current state of valid_devices_df to the Excel file.
 
        Called after each major stage of the update process to persist
        results in case of a later failure or cancellation. On failure,
        emits a warning signal so the user is informed without interrupting
        the update workflow.
 
 
        ARGUMENTS
        ---------
        None
 
 
        RETURN VALUE
        ------------
        None
        """
 
        result = excel_and_data_ops.update_excel_tracker(
            self.excel_file,
            self.sheet_name,
            self.valid_devices_df
        )
 
        if result == "SUCCESS":
            return
 
        if result == "MISSING_COLUMN_EXCEL_ERROR":
            self.warning.emit(
                f"⚠️ EXCEL TRACKER NOT UPDATED\n\n"
                f"The Excel sheet '{self.sheet_name}' is missing one or more required tracking columns.\n\n"
                f"The update process will continue, but results will not be saved to the Excel file.\n\n"
                f"Please add the missing columns to the sheet after the operation completes."
            )
 
        elif result == "PERMISSION_DENIED_ERROR":
            self.warning.emit(
                f"⚠️ EXCEL TRACKER NOT UPDATED\n\n"
                f"Cannot write to the Excel file — permission denied.\n\n"
                f"The update process will continue, but results will not be saved to the Excel file.\n\n"
                f"Please ensure the file is closed and not read-only, then update it manually."
            )
 
        elif result == "UNEXPECTED_ERROR":
            self.warning.emit(
                f"⚠️ EXCEL TRACKER NOT UPDATED\n\n"
                f"An unexpected error occurred while saving to the Excel file.\n\n"
                f"The update process will continue. Please check the logs for details."
            )
 
        else:
            self.warning.emit(
                f"⚠️ EXCEL TRACKER NOT UPDATED\n\n"
                f"Excel tracker write failed: {result}\n\n"
                f"The update process will continue."
            )

    def _generate_summary(self):

        """
        PURPOSE
        -------
        Analyze the 'Update Result' column to produce a human-readable summary
        of the update process.
        
        Only devices that were part of this run (selected_indices) are counted.
        The summary reports SUCCESS, FAILED, COMMIT_FAILED, and UNKNOWN counts.
        

        ARGUMENTS
        ---------
        None
        

        RETURN VALUE
        ------------
        str: Multi-line summary string suitable for display in a QMessageBox.
        """

        df = self.valid_devices_df
        # Only consider devices that were part of this run (selected indices)
        selected_df = df.loc[self.selected_indices]
        total = len(selected_df)
        success = (selected_df['Update Result'] == 'SUCCESS').sum()  # Count successful updates
        failed = (selected_df['Update Result'] == 'FAILED').sum()  # Count version mismatches
        commit_failed = (selected_df['Update Result'] == 'COMMIT_FAILED').sum()  # Count commit failures
        unknown = total - success - failed - commit_failed  # Remaining devices (transfer failed, etc.)

        summary_lines = [
            f"Update process completed.",
            f"",
            f"Selected devices: {total}",
            f"  ✅ SUCCESS       : {success}",
            f"  ❌ FAILED        : {failed}",
            f"  ⚠️ COMMIT_FAILED : {commit_failed}",
            f"  ❓ UNKNOWN       : {unknown}",
            f"",
            f"Detailed results are available in the Excel tracker."
        ]
        return "\n".join(summary_lines)

###################################################################################################################################

# Ensure this code only executes on the main program file, and not imported as an external module
if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
