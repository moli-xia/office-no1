"""PySide6 图形界面：预设管理 + 参数配置 + 批量处理。

布局为三栏工作流：① 选择预设 → ② 调整参数 → ③ 添加文件并排版。
"""

import json
import os
import re
import shutil
import sys
import tempfile
import time

from PySide6.QtCore import QSettings, Qt, QThread, Signal
from PySide6.QtGui import QColor, QIcon, QKeySequence, QPainter, QPalette, QShortcut
from PySide6.QtWidgets import (
    QAbstractItemView, QApplication, QCheckBox, QComboBox, QDoubleSpinBox, QFileDialog,
    QFormLayout, QFrame, QGridLayout, QHBoxLayout, QHeaderView, QInputDialog, QLabel, QLineEdit, QListWidget,
    QMainWindow, QMenu, QMessageBox, QPlainTextEdit, QProgressBar, QPushButton, QRadioButton,
    QScrollArea, QSizePolicy, QSpinBox, QSplitter, QTabWidget, QToolButton, QTreeWidget,
    QTreeWidgetItem, QVBoxLayout, QWidget,
)

from . import engine, fonts
from .presets import BUILTIN_PRESETS, merge_preset

APP_TITLE = "文档格式排版助手"
ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ASSET_DIR = os.path.join(ROOT_DIR, "app", "assets")
# PyInstaller 单文件模式下 __file__ 指向临时解压目录（每次运行都会变化），
# 只读资源（QSS 用的 SVG）从解压目录取，可写的用户预设必须放在 exe 旁边
if getattr(sys, "frozen", False):
    APP_DATA_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    APP_DATA_DIR = ROOT_DIR
USER_PRESET_DIR = os.path.join(APP_DATA_DIR, "presets_user")
OUTPUT_SUFFIX_DEFAULT = "_已排版"
DOC_EXTS = (".docx", ".doc", ".wps")

SIZE_OPTIONS = [("初号", 42.0), ("小初", 36.0), ("一号", 26.0), ("小一", 24.0),
                ("二号", 22.0), ("小二", 18.0), ("三号", 16.0), ("小三", 15.0),
                ("四号", 14.0), ("小四", 12.0), ("五号", 10.5), ("小五", 9.0),
                ("六号", 7.5), ("小六", 6.5)]
SIZE_NAMES = {v: n for n, v in SIZE_OPTIONS}

CHINESE_FONTS = ["宋体", "黑体", "仿宋", "仿宋_GB2312", "楷体", "楷体_GB2312", "微软雅黑",
                 "等线", "华文中宋", "华文仿宋", "华文楷体", "方正小标宋简体", "隶书",
                 "幼圆", "思源黑体 CN", "思源宋体 CN"]
WESTERN_FONTS = ["Times New Roman", "Arial", "Calibri", "Cambria", "Georgia", "Garamond",
                 "Segoe UI", "Verdana", "Helvetica"]

PAGE_SIZE_LABELS = {"A4": "A4（210 × 297 mm）", "B5": "B5（182 × 257 mm）",
                    "Letter": "Letter（8.5 × 11 in）", "16K": "16K（184 × 260 mm）"}
ORIENTATION_LABELS = {"portrait": "纵向", "landscape": "横向"}
ALIGN_LABELS = {"justify": "两端对齐", "left": "左对齐", "center": "居中", "right": "右对齐"}
LS_TYPE_LABELS = {"multiple": "多倍行距", "exact": "固定值", "minimum": "最小值"}
LS_DEFAULTS = {"multiple": 1.5, "exact": 28.0, "minimum": 20.0}

# 文件队列状态：文本、颜色
STATUS_STYLES = {
    "pending": ("待处理", "#64748b"),
    "waiting": ("等待中", "#64748b"),
    "running": ("处理中", "#0c8276"),
    "done": ("已完成", "#087f6b"),
    "failed": ("失败", "#c2413b"),
}

INVALID_NAME_RE = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
RESERVED_NAMES = {"CON", "PRN", "AUX", "NUL", *[f"COM{i}" for i in range(1, 10)],
                  *[f"LPT{i}" for i in range(1, 10)]}

QSS = """
* { font-family: 'Microsoft YaHei UI','Microsoft YaHei','PingFang SC',sans-serif; font-size: 13px; }
QMainWindow, QWidget#central { background: #eef1f5; }
QStatusBar { background: #eef1f5; color: #475569; }
QLabel { color: #0f172a; background: transparent; }
QLabel#appTitle { font-size: 18px; font-weight: 700; color: #0f172a; }
QLabel#appSub { color: #64748b; font-size: 12px; }
QLabel#shortcutHint { color: #64748b; font-size: 12px; }
QLabel#hint { color: #64748b; font-size: 11px; }
QLabel#sectionTitle { font-size: 12px; font-weight: 600; color: #0c8276; }
QLabel#matrixHead { font-weight: 600; color: #334155; }

QFrame#card { background: #ffffff; border: 1px solid #e2e8f0; border-radius: 10px; }
QLabel#cardTitle { font-size: 14px; font-weight: 600; color: #0f172a; }
QLabel#cardStep { color: #ffffff; background: #0c8276; border-radius: 11px;
                  min-width: 22px; max-width: 22px; min-height: 22px; max-height: 22px;
                  font-weight: 700; font-size: 12px; }
QFrame#summary { background: #f0f7f6; border: 1px solid #cfe5e2; border-radius: 8px; }
QLabel#summaryText { color: #1f3d3a; }
QFrame#line { background: #e2e8f0; max-height: 1px; min-height: 1px; border: none; }
QFrame#fontBanner { background: #fff7e6; border: 1px solid #f5d38f; border-radius: 8px; }
QFrame#fontBanner QLabel { color: #7a4b00; }
QFrame#fontBanner QPushButton { padding: 5px 12px; }

QTabWidget::pane { border: none; border-top: 1px solid #e2e8f0; background: #ffffff; }
QTabWidget::tab-bar { left: 6px; }
QTabBar::tab { background: transparent; color: #475569; padding: 8px 14px; margin-right: 2px;
               border-bottom: 2px solid transparent; }
QTabBar::tab:selected { color: #0c8276; font-weight: 600; border-bottom: 2px solid #0c8276; }
QTabBar::tab:hover:!selected { color: #0c8276; }

QComboBox, QDoubleSpinBox, QSpinBox, QLineEdit { background: #ffffff; border: 1px solid #cbd5e1;
               border-radius: 6px; padding: 4px 8px; min-height: 22px; color: #0f172a;
               selection-background-color: #0c8276; selection-color: #ffffff; }
QComboBox:focus, QDoubleSpinBox:focus, QSpinBox:focus, QLineEdit:focus { border-color: #0c8276; }
QComboBox:disabled, QDoubleSpinBox:disabled, QSpinBox:disabled, QLineEdit:disabled {
               background: #f8fafc; color: #94a3b8; }
QComboBox QLineEdit { border: none; padding: 0; margin: 0; background: transparent; min-height: 0; }
QComboBox QAbstractItemView { background: #ffffff; color: #0f172a; border: 1px solid #cbd5e1;
               selection-background-color: #0c8276; selection-color: #ffffff; outline: none; }
QComboBox QAbstractItemView::item { min-height: 26px; padding: 3px 8px; margin: 2px 4px;
               border-radius: 4px; color: #0f172a; background: #ffffff; }
QComboBox QAbstractItemView::item:hover, QComboBox QAbstractItemView::item:selected {
               background: #0c8276; color: #ffffff; }
QComboBox::drop-down { subcontrol-origin: padding; subcontrol-position: center right;
               width: 24px; background: #f1f5f9; border-left: 1px solid #cbd5e1;
               border-top-right-radius: 6px; border-bottom-right-radius: 6px; }
QComboBox::drop-down:hover { background: #d4eeea; }
QComboBox::down-arrow { image: url(ASSET_DIR/chevron-down.svg); width: 12px; height: 12px; }
QSpinBox::up-button, QDoubleSpinBox::up-button { subcontrol-origin: border;
               subcontrol-position: top right; width: 20px; background: #f1f5f9;
               border-left: 1px solid #cbd5e1; border-bottom: 1px solid #cbd5e1;
               border-top-right-radius: 6px; }
QSpinBox::down-button, QDoubleSpinBox::down-button { subcontrol-origin: border;
               subcontrol-position: bottom right; width: 20px; background: #f1f5f9;
               border-left: 1px solid #cbd5e1; border-bottom-right-radius: 6px; }
QSpinBox::up-button:hover, QDoubleSpinBox::up-button:hover,
QSpinBox::down-button:hover, QDoubleSpinBox::down-button:hover { background: #d4eeea; }
QSpinBox::up-button:pressed, QDoubleSpinBox::up-button:pressed,
QSpinBox::down-button:pressed, QDoubleSpinBox::down-button:pressed { background: #0c8276; }
QSpinBox::up-arrow, QDoubleSpinBox::up-arrow { image: url(ASSET_DIR/chevron-up.svg); width: 10px; height: 10px; }
QSpinBox::down-arrow, QDoubleSpinBox::down-arrow { image: url(ASSET_DIR/chevron-down.svg); width: 10px; height: 10px; }

QCheckBox, QRadioButton { color: #0f172a; background: transparent; spacing: 6px; }
QCheckBox:disabled, QRadioButton:disabled { color: #94a3b8; }
QCheckBox::indicator, QRadioButton::indicator { width: 16px; height: 16px;
               border: 1px solid #94a3b8; background: #ffffff; }
QCheckBox::indicator { border-radius: 4px; }
QRadioButton::indicator { border-radius: 8px; }
QCheckBox::indicator:hover, QRadioButton::indicator:hover { border-color: #0c8276; }
QCheckBox::indicator:checked { background: #0c8276; border-color: #0c8276;
               image: url(ASSET_DIR/check.svg); }
QRadioButton::indicator:checked { background: #0c8276; border-color: #0c8276; }
QCheckBox::indicator:disabled, QRadioButton::indicator:disabled { border-color: #cbd5e1; background: #f1f5f9; }

QListWidget { background: #ffffff; border: 1px solid #e2e8f0; border-radius: 8px; padding: 4px; outline: none; }
QListWidget::item { border-radius: 6px; padding: 8px 10px; margin: 2px 0; color: #0f172a; }
QListWidget::item:selected { background: #0c8276; color: white; }
QListWidget::item:hover:!selected { background: #eef2f7; }

QTreeWidget { background: #ffffff; border: 1px solid #e2e8f0; border-radius: 8px; outline: none;
              alternate-background-color: #f8fafc; }
QTreeWidget::item { padding: 5px 4px; color: #0f172a; }
QTreeWidget::item:selected { background: #d4eeea; color: #0f172a; }
QTreeWidget::item:hover:!selected { background: #eef2f7; }
QHeaderView::section { background: #f1f5f9; color: #475569; border: none;
                       border-bottom: 1px solid #e2e8f0; border-right: 1px solid #e2e8f0;
                       padding: 6px 8px; font-weight: 600; }
QHeaderView::section:last { border-right: none; }

QPushButton, QToolButton { background: #e8edf3; color: #0f172a; border: none; border-radius: 6px; padding: 7px 14px; }
QPushButton:hover, QToolButton:hover { background: #d5dce6; }
QPushButton:pressed, QToolButton:pressed { background: #c4ccd8; }
QPushButton:disabled, QToolButton:disabled { background: #eef1f5; color: #a0aec0; }
QToolButton::menu-indicator { image: none; }
QPushButton[compact="true"], QToolButton[compact="true"] { padding: 7px 9px; }
QPushButton#primary { background: #0c8276; color: white; font-weight: 600; padding: 10px 22px; font-size: 14px; }
QPushButton#primary:hover { background: #09685f; }
QPushButton#primary:disabled { background: #a9c9c5; color: #f1f5f9; }
QPushButton#stop { background: #fdecec; color: #b42318; }
QPushButton#stop:hover { background: #f9d6d6; }
QPushButton#stop:disabled { background: #eef1f5; color: #a0aec0; }
QPushButton#link { background: transparent; color: #0c8276; padding: 2px 4px; }
QPushButton#link:hover { color: #09685f; text-decoration: underline; }

QMenu { background: #ffffff; border: 1px solid #e2e8f0; border-radius: 6px; padding: 4px; }
QMenu::item { padding: 6px 28px 6px 12px; border-radius: 4px; color: #0f172a; }
QMenu::item:selected { background: #0c8276; color: #ffffff; }
QMenu::item:disabled { color: #94a3b8; }
QMenu::separator { height: 1px; background: #e2e8f0; margin: 4px 6px; }

QPlainTextEdit#log { background: #f6f8fa; color: #475569; border: 1px solid #e2e8f0; border-radius: 8px;
                     padding: 6px; font-family: 'Cascadia Mono','Consolas',monospace; font-size: 12px; }
QProgressBar { background: #e2e8f0; border: none; border-radius: 7px; min-height: 14px; max-height: 14px;
               text-align: center; color: #334155; font-size: 11px; }
QProgressBar::chunk { background: #0c8276; border-radius: 7px; }
QScrollArea { background: transparent; border: none; }
QScrollArea > QWidget > QWidget { background: transparent; }
QSplitter::handle { background: transparent; }
QScrollBar:vertical { background: transparent; width: 9px; margin: 2px; }
QScrollBar::handle:vertical { background: #cbd5e1; border-radius: 4px; min-height: 30px; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }
QScrollBar:horizontal { background: transparent; height: 9px; margin: 2px; }
QScrollBar::handle:horizontal { background: #cbd5e1; border-radius: 4px; min-width: 30px; }
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal { width: 0; }
QToolTip { background: #0f172a; color: #ffffff; border: none; padding: 5px 8px; }
"""


# ---------------------------------------------------------------------------
# 小部件工具
# ---------------------------------------------------------------------------

def size_label(pt):
    pt = float(pt)
    name = SIZE_NAMES.get(pt)
    return f"{name}（{pt:g} 磅）" if name else f"{pt:g} 磅"


def make_size_combo():
    c = QComboBox()
    for name, v in SIZE_OPTIONS:
        c.addItem(f"{name}（{v:g} 磅）", v)
    compact_combo(c)
    return c


def set_size_combo(c, pt):
    pt = float(pt)
    idx = c.findData(pt)
    if idx < 0:
        c.addItem(f"{pt:g} 磅", pt)
        idx = c.count() - 1
    c.setCurrentIndex(idx)


def make_font_combo(items):
    c = QComboBox()
    c.setEditable(True)
    c.setInsertPolicy(QComboBox.NoInsert)
    c.addItems(items)
    compact_combo(c)
    # 从下拉列表选中后光标回到行首，长字体名在窄栏里不会被左侧裁掉
    c.currentIndexChanged.connect(lambda _i, c=c: c.lineEdit().setCursorPosition(0))
    return c


def compact_combo(c, chars=8):
    """让下拉框的最小宽度不随最长选项变化，窄栏（标题矩阵）里不会撑出横向滚动条。"""
    c.setSizeAdjustPolicy(QComboBox.AdjustToMinimumContentsLengthWithIcon)
    c.setMinimumContentsLength(chars)
    c.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
    return c


def set_font_combo(c, value, fallback_items):
    c.blockSignals(True)
    c.clear()
    c.addItems(fallback_items)
    c.blockSignals(False)
    if value:
        c.setCurrentText(value)
    c.lineEdit().setCursorPosition(0)


def set_data_combo(c, data, default=None):
    idx = c.findData(data)
    if idx < 0:
        idx = c.findData(default) if default is not None else 0
    c.setCurrentIndex(max(idx, 0))


def make_spin(lo, hi, suffix="", decimals=1, step=1.0, width=118):
    s = QDoubleSpinBox()
    s.setRange(lo, hi)
    s.setDecimals(decimals)
    s.setSingleStep(step)
    if suffix:
        s.setSuffix(suffix)
    if width:
        s.setFixedWidth(width)
    else:
        s.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
    return s


def hint_label(text):
    lab = QLabel(text)
    lab.setObjectName("hint")
    lab.setWordWrap(True)
    return lab


def hline():
    f = QFrame()
    f.setObjectName("line")
    f.setFrameShape(QFrame.HLine)
    return f


def make_card(step, title):
    """带步骤序号标题的白色卡片，返回 (frame, body_layout)。"""
    card = QFrame()
    card.setObjectName("card")
    v = QVBoxLayout(card)
    v.setContentsMargins(14, 12, 14, 14)
    v.setSpacing(10)
    head = QHBoxLayout()
    head.setSpacing(8)
    badge = QLabel(str(step))
    badge.setObjectName("cardStep")
    badge.setAlignment(Qt.AlignCenter)
    lab = QLabel(title)
    lab.setObjectName("cardTitle")
    head.addWidget(badge)
    head.addWidget(lab)
    head.addStretch(1)
    v.addLayout(head)
    return card, v, head


def make_form(parent=None):
    form = QFormLayout(parent) if parent is not None else QFormLayout()
    form.setLabelAlignment(Qt.AlignRight | Qt.AlignVCenter)
    form.setFormAlignment(Qt.AlignLeft | Qt.AlignTop)
    form.setHorizontalSpacing(12)
    form.setVerticalSpacing(10)
    form.setFieldGrowthPolicy(QFormLayout.AllNonFixedFieldsGrow)
    return form


def wrap_scroll(widget):
    scroll = QScrollArea()
    scroll.setWidgetResizable(True)
    scroll.setFrameShape(QFrame.NoFrame)
    scroll.setWidget(widget)
    return scroll


def collect_doc_files(folder):
    return [os.path.join(folder, fn) for fn in sorted(os.listdir(folder))
            if fn.lower().endswith(DOC_EXTS) and not fn.startswith("~$")]


def paths_from_mime(mime):
    paths = []
    for url in mime.urls():
        p = url.toLocalFile()
        if not p:
            continue
        if os.path.isdir(p):
            paths.extend(collect_doc_files(p))
        elif p.lower().endswith(DOC_EXTS) and not os.path.basename(p).startswith("~$"):
            paths.append(p)
    return paths


def valid_file_name(name):
    stem = name.split(".")[0].upper()
    return not (INVALID_NAME_RE.search(name) or name.endswith((".", " ")) or stem in RESERVED_NAMES)


def make_settings():
    """用户设置：默认走系统位置；设置环境变量 DOCFMT_SETTINGS 可改为指定 ini（测试 / 便携使用）。"""
    path = os.environ.get("DOCFMT_SETTINGS")
    if path:
        return QSettings(path, QSettings.IniFormat)
    return QSettings("office-no1", "DocFormatter")


def open_path(path):
    try:
        os.startfile(path)  # noqa: S606 —— 仅 Windows
    except (OSError, AttributeError) as error:
        QMessageBox.warning(None, "无法打开", f"{path}\n{error}")


# ---------------------------------------------------------------------------
# 文件队列
# ---------------------------------------------------------------------------

class FileTree(QTreeWidget):
    """待处理文件表：状态 / 文件名 / 结果；支持拖拽文件与文件夹、右键菜单、Delete 删除。"""

    COL_STATUS, COL_NAME, COL_INFO = 0, 1, 2
    removeRequested = Signal()
    openRequested = Signal(QTreeWidgetItem)
    openFolderRequested = Signal(QTreeWidgetItem)

    def __init__(self, on_add):
        super().__init__()
        self.on_add = on_add
        self.setColumnCount(3)
        self.setHeaderLabels(["状态", "文件", "结果"])
        self.setRootIsDecorated(False)
        self.setUniformRowHeights(True)
        self.setAlternatingRowColors(True)
        self.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.setAcceptDrops(True)
        self.setDragDropMode(QAbstractItemView.NoDragDrop)
        header = self.header()
        header.setHighlightSections(False)
        header.setStretchLastSection(False)
        header.setMinimumSectionSize(40)
        header.setSectionResizeMode(self.COL_STATUS, QHeaderView.Fixed)
        header.setSectionResizeMode(self.COL_NAME, QHeaderView.Stretch)
        header.setSectionResizeMode(self.COL_INFO, QHeaderView.Interactive)
        self.setColumnWidth(self.COL_STATUS, 66)
        self.setColumnWidth(self.COL_INFO, 130)
        self.setContextMenuPolicy(Qt.CustomContextMenu)
        self.customContextMenuRequested.connect(self._context_menu)

    # 便捷访问
    def count(self):
        return self.topLevelItemCount()

    def item(self, i):
        return self.topLevelItem(i)

    def paths(self):
        return [self.item(i).data(self.COL_NAME, Qt.UserRole) for i in range(self.count())]

    def add_path(self, path):
        it = QTreeWidgetItem(["", os.path.basename(path), ""])
        it.setData(self.COL_NAME, Qt.UserRole, path)
        it.setToolTip(self.COL_NAME, path)
        self.addTopLevelItem(it)
        self.set_status(it, "pending")
        return it

    def find_item(self, path):
        for i in range(self.count()):
            it = self.item(i)
            if it.data(self.COL_NAME, Qt.UserRole) == path:
                return it
        return None

    @staticmethod
    def set_status(item, status, info=None):
        text, color = STATUS_STYLES[status]
        item.setText(FileTree.COL_STATUS, text)
        item.setForeground(FileTree.COL_STATUS, QColor(color))
        item.setData(FileTree.COL_STATUS, Qt.UserRole, status)
        if info is not None:
            item.setText(FileTree.COL_INFO, info)
            item.setToolTip(FileTree.COL_INFO, info)

    # 空列表提示
    def paintEvent(self, event):
        super().paintEvent(event)
        if self.count() == 0:
            painter = QPainter(self.viewport())
            painter.setPen(QColor("#94a3b8"))
            painter.drawText(self.viewport().rect(), Qt.AlignCenter,
                             "将 .docx / .doc / .wps 文件或文件夹\n拖到这里，或点击下方“添加文件”")

    # 拖拽
    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()
        else:
            e.ignore()

    def dragMoveEvent(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()
        else:
            e.ignore()

    def dropEvent(self, e):
        if e.mimeData().hasUrls():
            paths = paths_from_mime(e.mimeData())
            if paths:
                self.on_add(paths)
            e.acceptProposedAction()
        else:
            e.ignore()

    def keyPressEvent(self, e):
        if e.key() in (Qt.Key_Delete, Qt.Key_Backspace):
            self.removeRequested.emit()
        else:
            super().keyPressEvent(e)

    def _context_menu(self, pos):
        item = self.itemAt(pos)
        menu = QMenu(self)
        if item is not None:
            act_open = menu.addAction("打开输出文件")
            act_open.setEnabled(item.data(self.COL_STATUS, Qt.UserRole) == "done")
            act_open.triggered.connect(lambda: self.openRequested.emit(item))
            act_src = menu.addAction("打开源文件")
            act_src.triggered.connect(lambda: open_path(item.data(self.COL_NAME, Qt.UserRole)))
            act_folder = menu.addAction("打开所在文件夹")
            act_folder.triggered.connect(lambda: self.openFolderRequested.emit(item))
            menu.addSeparator()
            act_rm = menu.addAction("从列表移除")
            act_rm.setEnabled(self.acceptDrops())  # 处理中时不允许改队列
            act_rm.triggered.connect(self.removeRequested.emit)
        else:
            act = menu.addAction("暂无文件，拖入或点击“添加文件”")
            act.setEnabled(False)
        menu.exec(self.viewport().mapToGlobal(pos))


# ---------------------------------------------------------------------------
# 后台处理线程
# ---------------------------------------------------------------------------

class FormatWorker(QThread):
    logSig = Signal(str)
    fileStarted = Signal(str)
    fileDone = Signal(str, bool, str)
    progressSig = Signal(int, int)
    finishedAll = Signal(str)
    outputReady = Signal(str, str)

    def __init__(self, files, preset, output_mode, output_dir, suffix):
        super().__init__()
        self.files = files
        self.preset = preset
        self.output_mode = output_mode
        self.output_dir = output_dir
        self.suffix = suffix

    def _build_output(self, f, ext=".docx"):
        stem = os.path.splitext(os.path.basename(f))[0]
        name = f"{stem}{self.suffix}{ext}"
        folder = self.output_dir if self.output_mode == "dir" else os.path.dirname(f)
        os.makedirs(folder, exist_ok=True)
        candidate = os.path.join(folder, name)
        index = 2
        while os.path.exists(candidate):
            candidate = os.path.join(folder, f"{stem}{self.suffix} ({index}){ext}")
            index += 1
        return candidate

    @staticmethod
    def _temp_docx(tag):
        return os.path.join(tempfile.gettempdir(), f"docfmt_{tag}_{int(time.time() * 1000)}.docx")

    def _process_one(self, f):
        """处理单个文件，返回 (ok, output_path_or_None, message)。"""
        is_wps = f.lower().endswith(".wps")
        src, temps = f, []
        try:
            if f.lower().endswith((".doc", ".wps")):
                ext = os.path.splitext(f)[1].lower()
                converted = self._temp_docx("in")
                temps.append(converted)
                self.logSig.emit(f"    检测到 {ext}，正在调用本机 Word / WPS 转换为 .docx …")
                if not engine.convert_doc_to_docx(src, converted):
                    return False, None, "转换失败：未找到 Word/WPS 或文件受保护，请先另存为 .docx"
                src = converted

            out = self._build_output(f, ".wps" if is_wps else ".docx")
            if os.path.abspath(out) == os.path.abspath(f):
                raise ValueError("输出路径与源文件相同，请修改后缀或输出目录")

            if not is_wps:
                res = engine.apply_preset(src, self.preset, out)
                return (res["ok"], res.get("output"), res["message"] if res["ok"] else res["error"])

            # .wps：先在临时 .docx 上排版，再用 WPS 存回 .wps；WPS 不可用时退回 .docx
            result_docx = self._temp_docx("out")
            temps.append(result_docx)
            res = engine.apply_preset(src, self.preset, result_docx)
            if not res["ok"]:
                return False, None, res["error"]
            self.logSig.emit("    正在用 WPS 保存回 .wps 格式 …")
            if engine.convert_docx_to_wps(result_docx, out):
                return True, out, res["message"]
            fallback = self._build_output(f, ".docx")
            shutil.move(result_docx, fallback)
            self.logSig.emit("    ⚠ 未能保存为 .wps（本机 WPS 不可用），已改为输出 .docx")
            return True, fallback, res["message"] + "（已输出 .docx）"
        finally:
            for t in temps:
                if os.path.exists(t):
                    try:
                        os.remove(t)
                    except OSError:
                        pass

    def run(self):
        total = len(self.files)
        ok_count = 0
        processed = 0
        for i, f in enumerate(self.files, 1):
            if self.isInterruptionRequested():
                break
            self.fileStarted.emit(f)
            self.logSig.emit(f"▶ [{i}/{total}] 开始处理：{os.path.basename(f)}")
            try:
                ok, output, message = self._process_one(f)
            except Exception as e:  # 保证线程内不崩溃
                ok, output, message = False, None, f"{type(e).__name__}: {e}"
            if ok:
                ok_count += 1
                self.outputReady.emit(f, output)
                self.logSig.emit(f"    ✔ 完成：{os.path.basename(output)}（{message}）")
            else:
                self.logSig.emit(f"    ✘ 失败：{message}")
            self.fileDone.emit(f, ok, message)
            processed = i
            self.progressSig.emit(i, total)
        label = "已停止" if self.isInterruptionRequested() else "处理完成"
        self.finishedAll.emit(f"{label}：成功 {ok_count}，失败 {processed - ok_count}，未处理 {total - processed}。")


# ---------------------------------------------------------------------------
# 主窗口
# ---------------------------------------------------------------------------

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.presets = []
        self.worker = None
        self._updating = False
        self._close_after_finish = False
        self._ls_values = dict(LS_DEFAULTS)
        self.outputs = {}
        self.settings = make_settings()
        self.setWindowTitle(APP_TITLE)
        self.setAcceptDrops(True)
        self.resize(1320, 840)
        self.setMinimumSize(1180, 700)  # 三栏最小宽度之和 + 边距，保证右栏不会被挤出窗口

        self._ui_ready = False
        self._build_ui()
        self._ui_ready = True
        self._load_presets()
        self._restore_settings()
        self._connect_summary_signals()
        self._update_summary()
        self._update_queue()
        self._refresh_font_banner()

        QShortcut(QKeySequence("Ctrl+O"), self, activated=self._add_files_dialog)
        QShortcut(QKeySequence("Ctrl+Return"), self, activated=self._start)
        QShortcut(QKeySequence("Ctrl+Enter"), self, activated=self._start)
        QShortcut(QKeySequence("Ctrl+S"), self, activated=self._save_current_preset)

    # ---------------- UI 构建 ----------------

    def _build_ui(self):
        central = QWidget()
        central.setObjectName("central")
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(14, 10, 14, 8)
        root.setSpacing(10)

        # 顶部标题栏
        head = QHBoxLayout()
        head.setSpacing(12)
        title = QLabel(APP_TITLE)
        title.setObjectName("appTitle")
        sub = QLabel("选择规范 → 调整细节 → 批量排版 Word / WPS 文档，输出为新文件，不覆盖原稿")
        sub.setObjectName("appSub")
        head.addWidget(title)
        head.addWidget(sub)
        head.addStretch(1)
        keys = QLabel("Ctrl+O 添加文件 · Ctrl+S 保存预设 · Ctrl+Enter 开始排版 · Delete 移除所选")
        keys.setObjectName("shortcutHint")
        head.addWidget(keys)
        root.addLayout(head)

        # 字体缺失提示条（默认隐藏，启动检测后按需显示）
        self.fontBanner = QFrame()
        self.fontBanner.setObjectName("fontBanner")
        fb = QHBoxLayout(self.fontBanner)
        fb.setContentsMargins(12, 8, 12, 8)
        fb.setSpacing(8)
        self.fontBannerLabel = QLabel()
        self.fontBannerLabel.setWordWrap(True)
        fb.addWidget(self.fontBannerLabel, 1)
        b_install = QPushButton("安装字体")
        b_install.setObjectName("primary")
        b_install.clicked.connect(self._install_missing_fonts)
        b_folder = QPushButton("打开字体文件夹")
        b_folder.clicked.connect(lambda: open_path(fonts.FONT_DIR))
        b_ignore = QPushButton("忽略")
        b_ignore.clicked.connect(self.fontBanner.hide)
        for b in (b_install, b_folder, b_ignore):
            fb.addWidget(b)
        self.fontBanner.hide()
        root.addWidget(self.fontBanner)

        # 三栏工作区
        self.splitter = QSplitter(Qt.Horizontal)
        self.splitter.setChildrenCollapsible(False)
        self.splitter.setHandleWidth(10)
        self.splitter.addWidget(self._build_preset_pane())
        self.splitter.addWidget(self._build_params_pane())
        self.splitter.addWidget(self._build_files_pane())
        self.splitter.setStretchFactor(0, 0)
        self.splitter.setStretchFactor(1, 3)
        self.splitter.setStretchFactor(2, 2)
        self.splitter.setSizes([270, 630, 400])
        root.addWidget(self.splitter, 1)

        self.statusBar().showMessage("就绪。选择左侧预设 → 添加文件 → 开始排版。")

    # ---- ① 预设 ----
    def _build_preset_pane(self):
        card, v, head = make_card(1, "选择预设")
        card.setMinimumWidth(240)
        card.setMaximumWidth(360)
        self.presetPanel = card

        self.presetList = QListWidget()
        self.presetList.setWordWrap(True)
        self.presetList.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.presetList.currentRowChanged.connect(self._on_preset_selected)
        self.presetList.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Expanding)
        v.addWidget(self.presetList, 1)

        summary = QFrame()
        summary.setObjectName("summary")
        sv = QVBoxLayout(summary)
        sv.setContentsMargins(12, 10, 12, 10)
        sv.setSpacing(4)
        st = QLabel("当前规范")
        st.setObjectName("sectionTitle")
        self.summaryLabel = QLabel()
        self.summaryLabel.setObjectName("summaryText")
        self.summaryLabel.setWordWrap(True)
        self.summaryLabel.setTextFormat(Qt.RichText)
        sv.addWidget(st)
        sv.addWidget(self.summaryLabel)
        v.addWidget(summary)

        row = QHBoxLayout()
        row.setSpacing(6)
        self.btnReset = QPushButton("恢复")
        self.btnReset.setToolTip("放弃当前修改，恢复为所选预设的参数")
        self.btnReset.clicked.connect(lambda: self._on_preset_selected(self.presetList.currentRow()))
        btn_save = QPushButton("保存为预设")
        btn_save.setToolTip("把当前参数保存为自定义预设（Ctrl+S）")
        btn_save.clicked.connect(self._save_current_preset)
        more = QToolButton()
        more.setText("更多 ▾")
        more.setPopupMode(QToolButton.InstantPopup)
        menu = QMenu(more)
        menu.addAction("导入预设文件…", self._import_preset)
        menu.addAction("导出当前参数…", self._export_preset)
        menu.addSeparator()
        self.actDelete = menu.addAction("删除所选预设", self._delete_preset)
        menu.addAction("打开预设文件夹", self._open_preset_dir)
        menu.addSeparator()
        menu.addAction("检查并安装预设字体…", self._check_fonts_interactive)
        more.setMenu(menu)
        for b in (self.btnReset, btn_save, more):
            b.setProperty("compact", True)
        row.addWidget(self.btnReset)
        row.addWidget(btn_save, 1)
        row.addWidget(more)
        v.addLayout(row)
        v.addWidget(hint_label("内置预设不可删除；调整中间参数后可另存为新预设。"))
        return card

    # ---- ② 参数 ----
    def _build_params_pane(self):
        card, v, head = make_card(2, "调整参数")
        card.setMinimumWidth(480)
        self.paramsPanel = card
        self.tabs = QTabWidget()
        self.tabs.setDocumentMode(True)
        self.tabs.addTab(wrap_scroll(self._build_page_tab()), "页面")
        self.tabs.addTab(wrap_scroll(self._build_body_tab()), "正文")
        self.tabs.addTab(wrap_scroll(self._build_heading_tab()), "标题")
        self.tabs.addTab(wrap_scroll(self._build_pagenum_tab()), "页码 / 页眉页脚")
        v.addWidget(self.tabs, 1)
        return card

    def _build_page_tab(self):
        w = QWidget()
        form = make_form(w)
        form.setContentsMargins(12, 16, 16, 12)

        paper = QHBoxLayout()
        paper.setSpacing(8)
        self.pageSize = QComboBox()
        for k, lab in PAGE_SIZE_LABELS.items():
            self.pageSize.addItem(lab, k)
        self.pageSize.setMinimumWidth(190)
        self.pageOrient = QComboBox()
        for k, lab in ORIENTATION_LABELS.items():
            self.pageOrient.addItem(lab, k)
        self.pageOrient.setFixedWidth(90)
        paper.addWidget(self.pageSize, 1)
        paper.addWidget(QLabel("方向"))
        paper.addWidget(self.pageOrient)
        form.addRow("纸张：", paper)

        margins = QGridLayout()
        margins.setHorizontalSpacing(8)
        margins.setVerticalSpacing(8)
        self.margins = {}
        for index, (key, label) in enumerate((("top", "上"), ("bottom", "下"),
                                              ("left", "左"), ("right", "右"))):
            s = make_spin(0.0, 10.0, " cm", decimals=2, step=0.1)
            self.margins[key] = s
            r, c = index // 2, (index % 2) * 2
            margins.addWidget(QLabel(label), r, c, Qt.AlignRight | Qt.AlignVCenter)
            margins.addWidget(s, r, c + 1)
        margins.setColumnMinimumWidth(2, 24)
        margins.setColumnStretch(4, 1)
        form.addRow("页边距：", margins)

        row = QHBoxLayout()
        row.setSpacing(8)
        self.headerDist = make_spin(0.0, 10.0, " cm", decimals=2, step=0.1)
        self.footerDist = make_spin(0.0, 10.0, " cm", decimals=2, step=0.1)
        self.footerDist.setToolTip("页脚距页面下边缘的距离。公文：2.5 cm 时页码一字线上边距版心下边缘约 7 mm")
        row.addWidget(self.headerDist)
        row.addWidget(QLabel("页脚距边界"))
        row.addWidget(self.footerDist)
        row.addStretch(1)
        form.addRow("页眉距边界：", row)

        form.addRow("", hint_label("纸张与页边距会应用到文档中的所有分节；"
                                   "行距在“正文”页设置，排版时会禁用文档网格保证行距真实生效。"))
        return w

    def _build_body_tab(self):
        w = QWidget()
        form = make_form(w)
        form.setContentsMargins(12, 16, 16, 12)

        self.bodyEast = make_font_combo(CHINESE_FONTS)
        form.addRow("中文字体：", self.bodyEast)
        self.bodyWest = make_font_combo(WESTERN_FONTS)
        form.addRow("西文字体：", self.bodyWest)

        row = QHBoxLayout()
        row.setSpacing(8)
        self.bodySize = make_size_combo()
        self.bodyAlign = QComboBox()
        for k, lab in ALIGN_LABELS.items():
            self.bodyAlign.addItem(lab, k)
        row.addWidget(self.bodySize, 1)
        row.addWidget(QLabel("对齐"))
        row.addWidget(self.bodyAlign, 1)
        form.addRow("字号：", row)

        row = QHBoxLayout()
        row.setSpacing(8)
        self.lsType = QComboBox()
        for k, lab in LS_TYPE_LABELS.items():
            self.lsType.addItem(lab, k)
        self.lsValue = make_spin(0.1, 500.0, decimals=2, step=0.5)
        self.lsType.currentIndexChanged.connect(self._on_ls_type_changed)
        self.lsValue.valueChanged.connect(self._remember_ls_value)
        self._on_ls_type_changed()
        row.addWidget(self.lsType, 1)
        row.addWidget(self.lsValue)
        row.addStretch(1)
        form.addRow("行距：", row)

        row = QHBoxLayout()
        row.setSpacing(8)
        self.spaceBefore = make_spin(0.0, 200.0, " 磅")
        self.spaceAfter = make_spin(0.0, 200.0, " 磅")
        row.addWidget(self.spaceBefore)
        row.addWidget(QLabel("段后"))
        row.addWidget(self.spaceAfter)
        row.addStretch(1)
        form.addRow("段前：", row)

        row = QHBoxLayout()
        row.setSpacing(8)
        self.firstIndent = make_spin(0.0, 10.0, " 字符", decimals=1, step=0.5)
        row.addWidget(self.firstIndent)
        row.addWidget(hint_label("按字符缩进，换字号不跑偏；编号列表段落保留原缩进"))
        row.addStretch(1)
        form.addRow("首行缩进：", row)

        self.fmtTables = QCheckBox("同时统一表格内文字（字体、字号、行距；不加首行缩进）")
        form.addRow("", self.fmtTables)
        form.addRow("", hint_label("正文只统一字体与字号，不改动您手动设置的加粗、颜色等局部强调；"
                                   "标题按内置标题样式或大纲级别识别，包括基于标题样式的自定义样式。"))
        return w

    def _build_heading_tab(self):
        """标题参数矩阵：行 = 属性，列 = 1/2/3 级标题，一屏看清三级差异。"""
        w = QWidget()
        v = QVBoxLayout(w)
        v.setContentsMargins(12, 14, 16, 12)
        v.setSpacing(10)
        grid = QGridLayout()
        grid.setHorizontalSpacing(10)
        grid.setVerticalSpacing(8)
        grid.setColumnStretch(0, 0)
        for col, text in enumerate(("一级标题", "二级标题", "三级标题", "四级标题"), 1):
            lab = QLabel(text)
            lab.setObjectName("matrixHead")
            lab.setAlignment(Qt.AlignCenter)
            grid.addWidget(lab, 0, col)
            grid.setColumnStretch(col, 1)

        rows = (("中文字体", "east"), ("西文字体", "west"), ("字号", "size"), ("加粗", "bold"),
                ("对齐", "align"), ("段前", "before"), ("段后", "after"), ("首行缩进", "indent"))
        self.heading_widgets = [{}, {}, {}, {}]
        for r, (label, key) in enumerate(rows, 1):
            lab = QLabel(label + "：")
            lab.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
            grid.addWidget(lab, r, 0)
            for col, d in enumerate(self.heading_widgets, 1):
                if key == "east":
                    wdg = compact_combo(make_font_combo(CHINESE_FONTS), 4)
                elif key == "west":
                    wdg = compact_combo(make_font_combo(WESTERN_FONTS), 4)
                elif key == "size":
                    wdg = compact_combo(make_size_combo(), 4)
                elif key == "bold":
                    wdg = QCheckBox("加粗")
                elif key == "align":
                    wdg = compact_combo(QComboBox(), 4)
                    for k, lab_ in ALIGN_LABELS.items():
                        wdg.addItem(lab_, k)
                elif key in ("before", "after"):
                    wdg = make_spin(0, 200, " 磅", width=0)
                else:
                    wdg = make_spin(0, 10, " 字符", decimals=1, step=0.5, width=0)
                d[key] = wdg
                grid.addWidget(wdg, r, col)
        v.addLayout(grid)
        v.addWidget(hint_label("标题行距沿用“正文”页的行距设置；5 级及以上标题保持原样。"
                               "公文预设中：一级 = 公文标题（2 号小标宋）、二级 = “一、”（黑体）、"
                               "三级 = “（一）”（楷体）、四级 = “1.”（仿宋）。"))
        v.addStretch(1)
        return w

    def _build_pagenum_tab(self):
        w = QWidget()
        v = QVBoxLayout(w)
        v.setContentsMargins(12, 14, 16, 12)
        v.setSpacing(10)

        t1 = QLabel("页码")
        t1.setObjectName("sectionTitle")
        v.addWidget(t1)
        self.pnEnabled = QCheckBox("添加 / 替换页码")
        self.pnEnabled.setToolTip("勾选后会重建页眉页脚；不勾选且无页眉页脚文字时，原文档页眉页脚原样保留")
        v.addWidget(self.pnEnabled)

        self.pnForm = QWidget()
        f1 = make_form(self.pnForm)
        f1.setContentsMargins(18, 0, 0, 0)
        row = QHBoxLayout()
        row.setSpacing(8)
        self.pnPos = compact_combo(QComboBox())
        for k, lab in engine.POSITION_LABELS.items():
            self.pnPos.addItem(lab, k)
        self.pnFmt = compact_combo(QComboBox())
        for k, lab in engine.FORMAT_LABELS.items():
            self.pnFmt.addItem(lab, k)
        row.addWidget(self.pnPos, 1)
        row.addWidget(QLabel("样式"))
        row.addWidget(self.pnFmt, 1)
        f1.addRow("位置：", row)
        self.pnEast = make_font_combo(CHINESE_FONTS)
        f1.addRow("中文字体：", self.pnEast)
        self.pnWest = make_font_combo(WESTERN_FONTS)
        f1.addRow("西文字体：", self.pnWest)
        row = QHBoxLayout()
        row.setSpacing(8)
        self.pnSize = make_size_combo()
        self.pnStart = QSpinBox()
        self.pnStart.setRange(1, 9999)
        self.pnStart.setFixedWidth(90)
        row.addWidget(self.pnSize, 1)
        row.addWidget(QLabel("起始编号"))
        row.addWidget(self.pnStart)
        f1.addRow("字号：", row)
        self.pnHideFirst = QCheckBox("首页不显示页码（封面）")
        f1.addRow("", self.pnHideFirst)
        v.addWidget(self.pnForm)
        self.pnEnabled.toggled.connect(self.pnForm.setEnabled)

        v.addWidget(hline())
        t2 = QLabel("页眉 / 页脚文字")
        t2.setObjectName("sectionTitle")
        v.addWidget(t2)
        f2 = make_form()
        f2.setContentsMargins(0, 0, 0, 0)
        self.hdEnabled = QCheckBox("页眉文字")
        self.hdText = QLineEdit()
        self.hdText.setPlaceholderText("例如：××公司内部文件（居中，字体字号与页码相同）")
        self.hdEnabled.toggled.connect(self.hdText.setEnabled)
        f2.addRow(self.hdEnabled, self.hdText)
        self.ftEnabled = QCheckBox("页脚文字")
        self.ftText = QLineEdit()
        self.ftText.setPlaceholderText("例如：×× 部门制表（与页码位置对齐）")
        self.ftEnabled.toggled.connect(self.ftText.setEnabled)
        f2.addRow(self.ftEnabled, self.ftText)
        v.addLayout(f2)
        v.addWidget(hint_label("页码与页眉页脚写入文档第一节，其余分节自动链接到前一节，全文连续编号。"
                               "公文国标要求单页码居右、双页码居左，本工具简化为统一位置。"))
        v.addStretch(1)
        return w

    # ---- ③ 文件与输出 ----
    def _build_files_pane(self):
        card, v, head = make_card(3, "添加文件并排版")
        card.setMinimumWidth(360)
        self.queueLabel = QLabel("")
        self.queueLabel.setObjectName("hint")
        head.addWidget(self.queueLabel)

        self.filesList = FileTree(self.add_files)
        self.filesList.setMinimumHeight(140)
        self.filesList.itemDoubleClicked.connect(lambda item, _c: self._open_result(item))
        self.filesList.openRequested.connect(self._open_result)
        self.filesList.openFolderRequested.connect(self._open_item_folder)
        self.filesList.removeRequested.connect(self._remove_selected)
        self.filesList.model().rowsInserted.connect(self._update_queue)
        self.filesList.model().rowsRemoved.connect(self._update_queue)
        self.filesList.model().modelReset.connect(self._update_queue)
        v.addWidget(self.filesList, 3)

        row = QHBoxLayout()
        row.setSpacing(6)
        b1 = QPushButton("添加文件…")
        b1.clicked.connect(self._add_files_dialog)
        b2 = QPushButton("添加文件夹…")
        b2.clicked.connect(self._add_folder_dialog)
        b3 = QPushButton("移除")
        b3.setToolTip("从列表移除所选文件（Delete）")
        b3.clicked.connect(self._remove_selected)
        b4 = QPushButton("清空")
        b4.clicked.connect(self._clear_files)
        for b in (b1, b2, b3, b4):
            b.setProperty("compact", True)
            row.addWidget(b)
        row.addStretch(1)
        v.addLayout(row)
        self.queueButtons = [b1, b2, b3, b4]

        v.addWidget(hline())
        out_title = QLabel("输出位置")
        out_title.setObjectName("sectionTitle")
        v.addWidget(out_title)
        self.outputPanel = QWidget()
        og = QGridLayout(self.outputPanel)
        og.setContentsMargins(0, 0, 0, 0)
        og.setHorizontalSpacing(8)
        og.setVerticalSpacing(6)
        self.radioSuffix = QRadioButton("原目录，加后缀")
        self.radioSuffix.setToolTip("保存到源文件所在目录，文件名追加后缀")
        self.radioSuffix.setChecked(True)
        self.suffixEdit = QLineEdit(OUTPUT_SUFFIX_DEFAULT)
        self.suffixEdit.setFixedWidth(120)
        self.radioDir = QRadioButton("指定目录")
        self.outdirEdit = QLineEdit()
        self.outdirEdit.setPlaceholderText("选择输出目录…")
        bBrowse = QPushButton("浏览…")
        bBrowse.clicked.connect(self._pick_outdir)
        og.addWidget(self.radioSuffix, 0, 0)
        og.addWidget(self.suffixEdit, 0, 1)
        og.addWidget(self.radioDir, 1, 0)
        og.addWidget(self.outdirEdit, 1, 1, 1, 2)
        og.addWidget(bBrowse, 1, 3)
        og.setColumnStretch(2, 1)
        self.radioDir.toggled.connect(self._on_output_mode_changed)
        v.addWidget(self.outputPanel)

        v.addWidget(hline())
        row = QHBoxLayout()
        row.setSpacing(8)
        self.btnStart = QPushButton("开始排版")
        self.btnStart.setObjectName("primary")
        self.btnStart.setMinimumWidth(130)
        self.btnStart.clicked.connect(self._start)
        self.btnCancel = QPushButton("停止")
        self.btnCancel.setObjectName("stop")
        self.btnCancel.setEnabled(False)
        self.btnCancel.clicked.connect(self._cancel)
        self.btnOpenOut = QPushButton("打开输出目录")
        self.btnOpenOut.clicked.connect(self._open_output_dir)
        row.addWidget(self.btnStart)
        row.addWidget(self.btnCancel)
        row.addStretch(1)
        row.addWidget(self.btnOpenOut)
        v.addLayout(row)

        self.progress = QProgressBar()
        self.progress.setRange(0, 1)
        self.progress.setValue(0)
        self.progress.setFormat("%v / %m")
        self.progress.setTextVisible(False)
        v.addWidget(self.progress)

        log_head = QHBoxLayout()
        lt = QLabel("处理日志")
        lt.setObjectName("sectionTitle")
        btn_clear_log = QPushButton("清空日志")
        btn_clear_log.setObjectName("link")
        btn_clear_log.setCursor(Qt.PointingHandCursor)
        log_head.addWidget(lt)
        log_head.addStretch(1)
        log_head.addWidget(btn_clear_log)
        v.addLayout(log_head)
        self.log = QPlainTextEdit()
        self.log.setObjectName("log")
        self.log.setReadOnly(True)
        self.log.setMinimumHeight(90)
        self.log.setMaximumBlockCount(2000)
        self.log.setPlaceholderText("处理日志将显示在这里…")
        btn_clear_log.clicked.connect(self.log.clear)
        v.addWidget(self.log, 2)

        self.btnStart.setEnabled(False)
        self.btnOpenOut.setEnabled(False)
        self._on_output_mode_changed()
        return card

    # ---------------- 设置持久化 ----------------

    def _restore_settings(self):
        s = self.settings
        geo = s.value("geometry")
        if geo is not None:
            self.restoreGeometry(geo)
        state = s.value("splitter")
        if state is not None:
            self.splitter.restoreState(state)
        self.suffixEdit.setText(s.value("suffix", OUTPUT_SUFFIX_DEFAULT) or OUTPUT_SUFFIX_DEFAULT)
        self.outdirEdit.setText(s.value("outdir", "") or "")
        if s.value("output_mode", "suffix") == "dir":
            self.radioDir.setChecked(True)
        last = s.value("preset", "")
        row = next((i for i, p in enumerate(self.presets) if p["name"] == last), None)
        if row is None:
            # 默认选中党政机关公文预设（不能只匹配“公文”——“标准办公文档”也含“公文”二字）
            row = next((i for i, p in enumerate(self.presets) if "党政机关公文" in p["name"]), 0)
        self.presetList.setCurrentRow(row)  # 列表刚重建，currentRow 为 -1，一定会触发载入

    def _save_settings(self):
        s = self.settings
        s.setValue("geometry", self.saveGeometry())
        s.setValue("splitter", self.splitter.saveState())
        s.setValue("suffix", self.suffixEdit.text())
        s.setValue("outdir", self.outdirEdit.text())
        s.setValue("output_mode", "dir" if self.radioDir.isChecked() else "suffix")
        row = self.presetList.currentRow()
        if 0 <= row < len(self.presets):
            s.setValue("preset", self.presets[row]["name"])

    # ---------------- 预设 <-> 界面 ----------------

    def _load_presets(self):
        self.presetList.blockSignals(True)
        self.presetList.clear()
        self.presets = [merge_preset(p) for p in BUILTIN_PRESETS]
        for p in self.presets:
            self.presetList.addItem(p["name"].replace("（GB/T 9704 参照）", "\nGB/T 9704 参照"))
        for p in self._load_user_presets():
            self.presets.append(p)
            self.presetList.addItem(f"★ {p['name']}")
        self.presetList.blockSignals(False)

    @staticmethod
    def _load_user_presets():
        out = []
        if os.path.isdir(USER_PRESET_DIR):
            for fn in sorted(os.listdir(USER_PRESET_DIR)):
                if fn.endswith(".json"):
                    try:
                        with open(os.path.join(USER_PRESET_DIR, fn), encoding="utf-8") as f:
                            data = json.load(f)
                        preset = merge_preset(data)
                        preset["name"] = str(data.get("name") or os.path.splitext(fn)[0])
                        out.append(preset)
                    except (OSError, ValueError):
                        continue
        return out

    def _current_preset(self):
        row = self.presetList.currentRow()
        return self.presets[row] if 0 <= row < len(self.presets) else None

    def _on_preset_selected(self, row):
        if 0 <= row < len(self.presets):
            self.apply_preset_to_ui(self.presets[row])
            self.actDelete.setEnabled(row >= len(BUILTIN_PRESETS))
            self.statusBar().showMessage(f"已载入预设：{self.presets[row]['name']}")

    def _connect_summary_signals(self):
        for widget in self.tabs.findChildren(QWidget):
            if isinstance(widget, QComboBox):
                widget.currentTextChanged.connect(self._update_summary)
            elif isinstance(widget, (QSpinBox, QDoubleSpinBox)):
                widget.valueChanged.connect(self._update_summary)
            elif isinstance(widget, QCheckBox):
                widget.toggled.connect(self._update_summary)
            elif isinstance(widget, QLineEdit) and widget.parent() is not None \
                    and not isinstance(widget.parent(), (QComboBox, QSpinBox, QDoubleSpinBox)):
                widget.textChanged.connect(self._update_summary)

    def _on_ls_type_changed(self):
        t = self.lsType.currentData()
        if t is None or not hasattr(self, "lsValue"):
            return
        self.lsValue.blockSignals(True)
        if t == "multiple":
            self.lsValue.setSuffix(" 倍")
            self.lsValue.setRange(0.1, 10.0)
            self.lsValue.setSingleStep(0.25)
        else:
            self.lsValue.setSuffix(" 磅")
            self.lsValue.setRange(1.0, 500.0)
            self.lsValue.setSingleStep(1.0)
        # 每种行距类型记住各自的数值，切换后不会出现“28 磅 → 10 倍”的怪值
        self.lsValue.setValue(self._ls_values.get(t, LS_DEFAULTS[t]))
        self.lsValue.blockSignals(False)
        self._update_summary()

    def _remember_ls_value(self, value):
        t = self.lsType.currentData()
        if t:
            self._ls_values[t] = float(value)

    def collect_preset(self):
        p = merge_preset({})
        cur = self._current_preset()
        p["name"] = cur["name"] if cur else "自定义"
        p["page"] = {
            "size": self.pageSize.currentData(),
            "orientation": self.pageOrient.currentData(),
            "margin_top_cm": self.margins["top"].value(),
            "margin_bottom_cm": self.margins["bottom"].value(),
            "margin_left_cm": self.margins["left"].value(),
            "margin_right_cm": self.margins["right"].value(),
            "header_distance_cm": self.headerDist.value(),
            "footer_distance_cm": self.footerDist.value(),
        }
        p["body"] = {
            "font_east": self.bodyEast.currentText().strip() or "宋体",
            "font_west": self.bodyWest.currentText().strip() or "Times New Roman",
            "size_pt": float(self.bodySize.currentData()),
            "align": self.bodyAlign.currentData(),
            "line_spacing_type": self.lsType.currentData(),
            "line_spacing_value": float(self.lsValue.value()),
            "space_before_pt": float(self.spaceBefore.value()),
            "space_after_pt": float(self.spaceAfter.value()),
            "first_line_indent_chars": float(self.firstIndent.value()),
            "format_tables": self.fmtTables.isChecked(),
        }
        p["headings"] = []
        for lvl, d in enumerate(self.heading_widgets, 1):
            p["headings"].append({
                "level": lvl,
                "font_east": d["east"].currentText().strip() or "黑体",
                "font_west": d["west"].currentText().strip() or "Times New Roman",
                "size_pt": float(d["size"].currentData()),
                "bold": d["bold"].isChecked(),
                "align": d["align"].currentData(),
                "space_before_pt": float(d["before"].value()),
                "space_after_pt": float(d["after"].value()),
                "first_line_indent_chars": float(d["indent"].value()),
            })
        p["page_number"] = {
            "enabled": self.pnEnabled.isChecked(),
            "position": self.pnPos.currentData(),
            "format": self.pnFmt.currentData(),
            "font_east": self.pnEast.currentText().strip() or "宋体",
            "font_west": self.pnWest.currentText().strip() or "Times New Roman",
            "size_pt": float(self.pnSize.currentData()),
            "hide_on_first_page": self.pnHideFirst.isChecked(),
            "start_at": self.pnStart.value(),
        }
        p["header"] = {"enabled": self.hdEnabled.isChecked(), "text": self.hdText.text()}
        p["footer"] = {"enabled": self.ftEnabled.isChecked(), "text": self.ftText.text()}
        return p

    def apply_preset_to_ui(self, p):
        self._updating = True
        try:
            pg = p["page"]
            set_data_combo(self.pageSize, pg["size"], "A4")
            set_data_combo(self.pageOrient, pg["orientation"], "portrait")
            for key in ("top", "bottom", "left", "right"):
                self.margins[key].setValue(float(pg[f"margin_{key}_cm"]))
            self.headerDist.setValue(float(pg.get("header_distance_cm", 1.5)))
            self.footerDist.setValue(float(pg.get("footer_distance_cm", 1.75)))

            b = p["body"]
            set_font_combo(self.bodyEast, b["font_east"], CHINESE_FONTS)
            set_font_combo(self.bodyWest, b["font_west"], WESTERN_FONTS)
            set_size_combo(self.bodySize, b["size_pt"])
            set_data_combo(self.bodyAlign, b["align"], "justify")
            self._ls_values = dict(LS_DEFAULTS)
            self._ls_values[b["line_spacing_type"]] = float(b["line_spacing_value"])
            set_data_combo(self.lsType, b["line_spacing_type"], "multiple")
            self.lsValue.setValue(float(b["line_spacing_value"]))
            self.spaceBefore.setValue(float(b["space_before_pt"]))
            self.spaceAfter.setValue(float(b["space_after_pt"]))
            self.firstIndent.setValue(float(b["first_line_indent_chars"]))
            self.fmtTables.setChecked(bool(b.get("format_tables", False)))

            for d, h in zip(self.heading_widgets, p["headings"]):
                set_font_combo(d["east"], h["font_east"], CHINESE_FONTS)
                set_font_combo(d["west"], h["font_west"], WESTERN_FONTS)
                set_size_combo(d["size"], h["size_pt"])
                d["bold"].setChecked(bool(h["bold"]))
                set_data_combo(d["align"], h["align"], "left")
                d["before"].setValue(float(h["space_before_pt"]))
                d["after"].setValue(float(h["space_after_pt"]))
                d["indent"].setValue(float(h.get("first_line_indent_chars", 0)))

            pn = p["page_number"]
            self.pnEnabled.setChecked(bool(pn["enabled"]))
            self.pnForm.setEnabled(bool(pn["enabled"]))
            set_data_combo(self.pnPos, pn["position"], "footer_center")
            set_data_combo(self.pnFmt, pn["format"], "plain")
            set_font_combo(self.pnEast, pn["font_east"], CHINESE_FONTS)
            set_font_combo(self.pnWest, pn["font_west"], WESTERN_FONTS)
            set_size_combo(self.pnSize, pn["size_pt"])
            self.pnHideFirst.setChecked(bool(pn.get("hide_on_first_page", False)))
            self.pnStart.setValue(int(pn.get("start_at", 1)))
            self.hdEnabled.setChecked(bool(p["header"]["enabled"]))
            self.hdText.setText(p["header"].get("text", ""))
            self.hdText.setEnabled(self.hdEnabled.isChecked())
            self.ftEnabled.setChecked(bool(p["footer"]["enabled"]))
            self.ftText.setText(p["footer"].get("text", ""))
            self.ftText.setEnabled(self.ftEnabled.isChecked())
        finally:
            self._updating = False
        self._update_summary()

    # ---------------- 规范摘要 ----------------

    @staticmethod
    def _normalized(preset):
        """去掉名称并把浮点数取整到 3 位，用于判断参数是否被修改。"""
        def walk(v):
            if isinstance(v, dict):
                return {k: walk(x) for k, x in v.items() if k != "name"}
            if isinstance(v, list):
                return [walk(x) for x in v]
            if isinstance(v, float):
                return round(v, 3)
            if isinstance(v, str):
                return v.strip()
            return v
        return walk(preset)

    def _is_dirty(self, current=None):
        base = self._current_preset()
        if base is None:
            return False
        current = current or self.collect_preset()
        return self._normalized(current) != self._normalized(base)

    def _update_summary(self, *args):
        if self._updating or not self._ui_ready:
            return
        p = self.collect_preset()
        b, pn = p["body"], p["page_number"]
        unit = "倍" if b["line_spacing_type"] == "multiple" else "磅"
        ls_kind = "" if b["line_spacing_type"] == "multiple" else LS_TYPE_LABELS[b["line_spacing_type"]]
        para_text = f"{ls_kind}行距 {b['line_spacing_value']:g} {unit}，缩进 {b['first_line_indent_chars']:g} 字符"
        h1 = p["headings"][0]
        if pn["enabled"]:
            pn_text = (f"{engine.POSITION_LABELS[pn['position']].replace('页面', '').replace('（公文）', '')}"
                       f" · {size_label(pn['size_pt'])}")
        else:
            pn_text = "不添加（保留原页眉页脚）"
        dirty = self._is_dirty(p)
        state = ("<span style='color:#b45309'>已修改，未保存</span>" if dirty
                 else "<span style='color:#64748b'>与预设一致</span>")
        rows = [
            ("预设", p["name"]),
            ("状态", state),
            ("纸张", f"{p['page']['size']} {ORIENTATION_LABELS[p['page']['orientation']]}，边距 "
                    f"{p['page']['margin_top_cm']:g}/{p['page']['margin_bottom_cm']:g}/"
                    f"{p['page']['margin_left_cm']:g}/{p['page']['margin_right_cm']:g} cm"),
            ("正文", f"{b['font_east']} / {b['font_west']}，{size_label(b['size_pt'])}"),
            ("段落", para_text),
            ("标题", f"一级 {h1['font_east']} {size_label(h1['size_pt'])}"
                    f"{'，加粗' if h1['bold'] else ''}，{ALIGN_LABELS[h1['align']]}"),
            ("页码", pn_text),
        ]
        html = "<br>".join(f"<span style='color:#64748b'>{k}</span>&nbsp; {v}" for k, v in rows)
        self.summaryLabel.setText(f"<div style='line-height:145%'>{html}</div>")
        self.btnReset.setEnabled(dirty)

    # ---------------- 预设增删 / 导入导出 ----------------

    def _save_current_preset(self):
        if self.worker and self.worker.isRunning():
            return
        cur = self._current_preset()
        default = cur["name"] if cur else ""
        name, ok = QInputDialog.getText(self, "保存为预设", "请输入预设名称：", text=default)
        if not ok or not name.strip():
            return
        name = name.strip()
        if not valid_file_name(name):
            QMessageBox.warning(self, "名称无效", "请使用不含 \\ / : * ? \" < > | 等特殊字符的普通名称。")
            return
        if any(p["name"] == name for p in self.presets[:len(BUILTIN_PRESETS)]):
            QMessageBox.warning(self, "名称冲突", "该名称已被内置预设使用，请换一个名称。")
            return
        preset = self.collect_preset()
        preset["name"] = name
        if not self._write_user_preset(preset, confirm_overwrite=True):
            return
        self._select_preset_by_name(name)
        self.log.appendPlainText(f"已保存预设：{name}")
        self.statusBar().showMessage(f"预设“{name}”已保存到 presets_user 目录。")

    def _write_user_preset(self, preset, confirm_overwrite):
        path = os.path.join(USER_PRESET_DIR, f"{preset['name']}.json")
        if confirm_overwrite and os.path.exists(path) and QMessageBox.question(
                self, "覆盖预设", f"预设“{preset['name']}”已存在，是否替换？") != QMessageBox.Yes:
            return False
        try:
            os.makedirs(USER_PRESET_DIR, exist_ok=True)
            with open(path, "w", encoding="utf-8") as f:
                json.dump(preset, f, ensure_ascii=False, indent=2)
        except OSError as error:
            QMessageBox.warning(self, "保存失败", str(error))
            return False
        self._load_presets()
        return True

    def _select_preset_by_name(self, name):
        for i, p in enumerate(self.presets):
            if p["name"] == name:
                self.presetList.setCurrentRow(i)
                return

    def _delete_preset(self):
        row = self.presetList.currentRow()
        if row < len(BUILTIN_PRESETS) or row >= len(self.presets):
            QMessageBox.information(self, "提示", "内置预设不可删除。")
            return
        name = self.presets[row]["name"]
        if QMessageBox.question(self, "删除预设", f"确定删除预设“{name}”吗？") != QMessageBox.Yes:
            return
        path = os.path.join(USER_PRESET_DIR, f"{name}.json")
        try:
            if os.path.exists(path):
                os.remove(path)
        except OSError as error:
            QMessageBox.warning(self, "删除失败", str(error))
            return
        self._load_presets()
        self.presetList.setCurrentRow(min(row, len(self.presets) - 1))

    def _export_preset(self):
        preset = self.collect_preset()
        default = os.path.join(self.settings.value("export_dir", "") or "", f"{preset['name']}.json")
        path, _ = QFileDialog.getSaveFileName(self, "导出预设", default, "预设文件 (*.json)")
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(preset, f, ensure_ascii=False, indent=2)
        except OSError as error:
            QMessageBox.warning(self, "导出失败", str(error))
            return
        self.settings.setValue("export_dir", os.path.dirname(path))
        self.statusBar().showMessage(f"已导出预设：{path}")

    def _import_preset(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "导入预设文件", "", "预设文件 (*.json)")
        if not paths:
            return
        imported = []
        for path in paths:
            try:
                with open(path, encoding="utf-8") as f:
                    data = json.load(f)
                if not isinstance(data, dict) or "body" not in data:
                    raise ValueError("不是有效的预设文件")
            except (OSError, ValueError) as error:
                QMessageBox.warning(self, "导入失败", f"{os.path.basename(path)}：{error}")
                continue
            preset = merge_preset(data)
            name = str(data.get("name") or os.path.splitext(os.path.basename(path))[0]).strip()
            if not valid_file_name(name) or any(p["name"] == name for p in self.presets[:len(BUILTIN_PRESETS)]):
                name = f"导入 {os.path.splitext(os.path.basename(path))[0]}"
            preset["name"] = name
            if self._write_user_preset(preset, confirm_overwrite=True):
                imported.append(name)
        if imported:
            self._select_preset_by_name(imported[-1])
            self.statusBar().showMessage(f"已导入 {len(imported)} 个预设。")

    def _open_preset_dir(self):
        os.makedirs(USER_PRESET_DIR, exist_ok=True)
        open_path(USER_PRESET_DIR)

    # ---------------- 字体检测与安装 ----------------

    def _refresh_font_banner(self):
        try:
            missing = fonts.missing_fonts()
        except Exception:
            missing = []
        self._missing_fonts = missing
        if not missing:
            self.fontBanner.hide()
            return []
        names = "、".join(f"{f['names'][0]}（{f['usage']}）" for f in missing)
        self.fontBannerLabel.setText(
            f"本机未安装预设所需字体：{names}。未安装时 Word / WPS 会用其它字体替代显示，"
            "建议安装（仅安装到当前用户，无需管理员权限）。")
        self.fontBanner.show()
        return missing

    def _install_missing_fonts(self):
        missing = getattr(self, "_missing_fonts", None) or fonts.missing_fonts()
        if not missing:
            self.fontBanner.hide()
            return
        ok, failed = fonts.install_fonts(missing)
        lines = []
        if ok:
            lines.append("已安装：" + "、".join(f["names"][0] for f in ok))
        for f, err in failed:
            lines.append(f"安装失败：{f['names'][0]} —— {err}")
        if failed:
            lines.append("\n可点击“打开字体文件夹”，右键字体文件选择“安装”手动完成。")
            QMessageBox.warning(self, "字体安装", "\n".join(lines))
        else:
            lines.append("\n已生效；若 Word / WPS 正在运行，需重新打开文档才能看到新字体。")
            QMessageBox.information(self, "字体安装", "\n".join(lines))
        self.log.appendPlainText("\n".join(lines))
        self._refresh_font_banner()

    def _check_fonts_interactive(self):
        missing = self._refresh_font_banner()
        if not missing:
            names = "、".join(f["names"][0] for f in fonts.BUNDLED_FONTS)
            QMessageBox.information(self, "字体检查", f"预设所需字体已全部安装：{names}。")
            return
        names = "\n".join(f"  • {f['names'][0]}（{f['usage']}）" for f in missing)
        if QMessageBox.question(self, "字体检查",
                                f"以下字体未安装：\n{names}\n\n是否现在安装到当前用户？") == QMessageBox.Yes:
            self._install_missing_fonts()

    # ---------------- 文件队列 ----------------

    def _is_busy(self):
        return bool(self.worker and self.worker.isRunning())

    def add_files(self, paths):
        if self._is_busy():
            return
        existing = {os.path.normcase(p) for p in self.filesList.paths()}
        added = skipped = 0
        for p in paths:
            ap = os.path.abspath(p)
            key = os.path.normcase(ap)
            if (not os.path.isfile(ap) or not ap.lower().endswith(DOC_EXTS)
                    or os.path.basename(ap).startswith("~$")):
                skipped += 1
                continue
            if key in existing:
                continue
            self.filesList.add_path(ap)
            existing.add(key)
            added += 1
        msg = f"已添加 {added} 个文件，共 {self.filesList.count()} 个。"
        if skipped:
            msg += f" 忽略 {skipped} 个非文档文件。"
        self.statusBar().showMessage(msg)

    def _add_files_dialog(self):
        if self._is_busy():
            return
        start = self.settings.value("open_dir", "") or ""
        files, _ = QFileDialog.getOpenFileNames(
            self, "选择 Word / WPS 文档", start,
            "Word / WPS 文档 (*.docx *.doc *.wps);;所有文件 (*)")
        if files:
            self.settings.setValue("open_dir", os.path.dirname(files[0]))
            self.add_files(files)

    def _add_folder_dialog(self):
        if self._is_busy():
            return
        d = QFileDialog.getExistingDirectory(self, "选择文件夹", self.settings.value("open_dir", "") or "")
        if not d:
            return
        self.settings.setValue("open_dir", d)
        paths = collect_doc_files(d)
        if paths:
            self.add_files(paths)
        else:
            QMessageBox.information(self, "提示", "该文件夹下没有 .docx / .doc / .wps 文件。")

    def _remove_selected(self):
        if self._is_busy():
            return
        for it in self.filesList.selectedItems():
            self.filesList.takeTopLevelItem(self.filesList.indexOfTopLevelItem(it))
        self._update_queue()

    def _clear_files(self):
        if self._is_busy():
            return
        self.filesList.clear()
        self.outputs.clear()
        self.btnOpenOut.setEnabled(False)
        self.progress.setRange(0, 1)
        self.progress.setValue(0)
        self._update_queue()

    def _update_queue(self, *args):
        count = self.filesList.count()
        done = sum(1 for i in range(count)
                   if self.filesList.item(i).data(FileTree.COL_STATUS, Qt.UserRole) == "done")
        if count == 0:
            text = "尚未添加文件"
        elif done:
            text = f"{count} 个文档 · {done} 个已完成"
        else:
            text = f"{count} 个文档"
        self.queueLabel.setText(text)
        if hasattr(self, "btnStart"):
            self.btnStart.setEnabled(count > 0 and not self._is_busy())

    def _on_output_mode_changed(self, *args):
        use_dir = self.radioDir.isChecked()
        self.suffixEdit.setEnabled(not use_dir)
        self.outdirEdit.setEnabled(use_dir)

    def _pick_outdir(self):
        d = QFileDialog.getExistingDirectory(self, "选择输出目录", self.outdirEdit.text().strip())
        if d:
            self.outdirEdit.setText(d)
            self.radioDir.setChecked(True)

    # 整个窗口都接受拖拽，不必精确拖到列表上
    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls() and not self._is_busy():
            e.acceptProposedAction()
        else:
            e.ignore()

    def dropEvent(self, e):
        paths = paths_from_mime(e.mimeData())
        if paths:
            self.add_files(paths)
        e.acceptProposedAction()

    # ---------------- 处理 ----------------

    def _validate(self, preset):
        pg = preset["page"]
        width, height = engine.PAGE_SIZES_CM[pg["size"]]
        if pg["orientation"] == "landscape":
            width, height = height, width
        if (pg["margin_left_cm"] + pg["margin_right_cm"] >= width
                or pg["margin_top_cm"] + pg["margin_bottom_cm"] >= height):
            return "页边距之和必须小于纸张宽度和高度，请检查“页面”页的页边距。"
        if self.radioDir.isChecked():
            outdir = self.outdirEdit.text().strip()
            if not outdir:
                return "请先选择输出目录。"
            if os.path.isfile(outdir):
                return "输出目录不是文件夹。"
        else:
            suffix = self.suffixEdit.text().strip()
            if INVALID_NAME_RE.search(suffix):
                return "文件名后缀不能含有路径分隔符或特殊字符。"
        if preset["header"]["enabled"] and not preset["header"]["text"].strip():
            return "已勾选页眉文字，但内容为空。"
        if preset["footer"]["enabled"] and not preset["footer"]["text"].strip():
            return "已勾选页脚文字，但内容为空。"
        return None

    def _start(self):
        if self._is_busy():
            return
        if self.filesList.count() == 0:
            QMessageBox.information(self, "提示", "请先添加要处理的文档。")
            return
        preset = self.collect_preset()
        problem = self._validate(preset)
        if problem:
            QMessageBox.warning(self, "无法开始", problem)
            return
        mode = "dir" if self.radioDir.isChecked() else "suffix"
        outdir = self.outdirEdit.text().strip()
        suffix = self.suffixEdit.text().strip() or OUTPUT_SUFFIX_DEFAULT
        files = self.filesList.paths()

        self._set_busy(True)
        self.outputs.clear()
        for i in range(self.filesList.count()):
            FileTree.set_status(self.filesList.item(i), "waiting", "")
        self.progress.setRange(0, len(files))
        self.progress.setValue(0)
        self.progress.setTextVisible(True)
        self.log.appendPlainText(f"—— 使用预设「{preset['name']}」开始处理 {len(files)} 个文件 ——")
        self.worker = FormatWorker(files, preset, mode, outdir, suffix)
        self.worker.logSig.connect(self.log.appendPlainText)
        self.worker.progressSig.connect(lambda v, t: self.progress.setValue(v))
        self.worker.fileStarted.connect(self._file_started)
        self.worker.fileDone.connect(self._file_done)
        self.worker.outputReady.connect(self._output_ready)
        self.worker.finishedAll.connect(self._on_all_done)
        self.worker.finished.connect(self._on_worker_finished)
        self.worker.start()

    def _on_worker_finished(self):
        self._set_busy(False)
        self._update_queue()
        if self._close_after_finish:
            self.close()

    def _on_all_done(self, summary):
        self.statusBar().showMessage(summary)
        self.log.appendPlainText(summary)

    def _set_busy(self, busy):
        self.paramsPanel.setEnabled(not busy)
        self.presetPanel.setEnabled(not busy)
        self.outputPanel.setEnabled(not busy)
        self.filesList.setAcceptDrops(not busy)
        for button in self.queueButtons:
            button.setEnabled(not busy)
        self.btnCancel.setEnabled(busy)
        self.btnCancel.setText("停止")
        self.btnStart.setEnabled(not busy and self.filesList.count() > 0)
        self.btnStart.setText("正在排版…" if busy else "开始排版")

    def _cancel(self):
        if self._is_busy():
            self.worker.requestInterruption()
            self.btnCancel.setEnabled(False)
            self.btnCancel.setText("停止中…")
            self.statusBar().showMessage("正在停止：当前文档完成后停止，已完成的文件保留。")

    def _file_started(self, path):
        item = self.filesList.find_item(path)
        if item is not None:
            FileTree.set_status(item, "running", "")
            self.filesList.scrollToItem(item)

    def _file_done(self, path, success, message):
        item = self.filesList.find_item(path)
        if item is not None:
            FileTree.set_status(item, "done" if success else "failed", message)

    def _output_ready(self, source, output):
        self.outputs[source] = output
        self.btnOpenOut.setEnabled(True)

    def _open_result(self, item):
        path = self.outputs.get(item.data(FileTree.COL_NAME, Qt.UserRole))
        if path and os.path.isfile(path):
            open_path(path)
        elif item.data(FileTree.COL_STATUS, Qt.UserRole) != "done":
            self.statusBar().showMessage("该文件尚未处理完成；处理完成后双击可打开输出文档。")

    def _open_item_folder(self, item):
        src = item.data(FileTree.COL_NAME, Qt.UserRole)
        out = self.outputs.get(src)
        open_path(os.path.dirname(out if out and os.path.isfile(out) else src))

    def _open_output_dir(self):
        if self.outputs:
            open_path(os.path.dirname(list(self.outputs.values())[-1]))
        elif self.radioDir.isChecked() and os.path.isdir(self.outdirEdit.text().strip()):
            open_path(self.outdirEdit.text().strip())
        elif self.filesList.count():
            open_path(os.path.dirname(self.filesList.paths()[0]))

    def closeEvent(self, event):
        if self._is_busy():
            self._close_after_finish = True
            self._cancel()
            event.ignore()
            self.statusBar().showMessage("正在安全停止，当前文档处理完成后将自动关闭窗口。")
            return
        self._save_settings()
        event.accept()


def build_stylesheet():
    return QSS.replace("ASSET_DIR", ASSET_DIR.replace("\\", "/"))


def apply_app_style(app):
    app.setStyle("Fusion")
    # 强制浅色调色板：避免系统深色模式下，未覆盖 QSS 的部件（如下拉弹窗、对话框）黑底黑字
    palette = QPalette()
    for role, color in (
        (QPalette.Window, "#eef1f5"), (QPalette.WindowText, "#0f172a"),
        (QPalette.Base, "#ffffff"), (QPalette.AlternateBase, "#f8fafc"),
        (QPalette.Text, "#0f172a"), (QPalette.PlaceholderText, "#94a3b8"),
        (QPalette.Button, "#ffffff"), (QPalette.ButtonText, "#0f172a"),
        (QPalette.Highlight, "#0c8276"), (QPalette.HighlightedText, "#ffffff"),
        (QPalette.ToolTipBase, "#0f172a"), (QPalette.ToolTipText, "#ffffff"),
        (QPalette.BrightText, "#ffffff"),
    ):
        palette.setColor(role, QColor(color))
    app.setPalette(palette)
    app.setStyleSheet(build_stylesheet())
    icon_path = os.path.join(ASSET_DIR, "icon.ico")
    if not os.path.exists(icon_path):
        icon_path = os.path.join(ASSET_DIR, "icon.png")
    if os.path.exists(icon_path):
        app.setWindowIcon(QIcon(icon_path))


def main():
    app = QApplication(sys.argv)
    apply_app_style(app)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())
