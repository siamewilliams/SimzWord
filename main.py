#!/usr/bin/env python3
"""SimzWord - a professional word processor built with PySide6."""

import sys
import os
import re
import math
import json
import tempfile
import zipfile
import hashlib
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape as xml_escape, quoteattr
from datetime import datetime

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QTextEdit, QToolButton, QFontComboBox, QComboBox,
    QFileDialog, QMessageBox, QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QCheckBox, QTabWidget, QWidget, QColorDialog, QStatusBar, QSlider,
    QFrame, QGridLayout, QMenu, QStyle, QDialogButtonBox, QScrollArea, QSizePolicy,
    QToolBar, QListWidget, QListWidgetItem, QInputDialog, QSpinBox, QDoubleSpinBox,
    QDockWidget, QPlainTextEdit, QFontDialog, QRadioButton, QButtonGroup,
    QAbstractItemView, QTextBrowser, QGraphicsScene, QGraphicsView
)
from PySide6.QtGui import (
    QAction, QActionGroup, QFont, QTextCharFormat, QTextCursor, QTextListFormat,
    QColor, QIcon, QPixmap, QPainter, QPen, QTextDocument, QTextBlockFormat,
    QTextTableFormat, QTextFrameFormat, QTextLength, QTextImageFormat, QKeySequence,
    QTextFormat, QPageSize, QPageLayout, QRegion, QLinearGradient, QBrush,
    QFontMetrics, QPalette, QPolygonF, QTextTableCellFormat
)
from PySide6.QtCore import (
    Qt, QSize, QSignalBlocker, QRect, QRectF, QPoint, QPointF, QMarginsF, QSizeF,
    QTimer, QSettings, QRegularExpression, QEvent
)
from PySide6.QtPrintSupport import QPrinter, QPrintDialog, QPrintPreviewDialog

APP_NAME = "SimzWord"
APP_VERSION = "1.0.0"
ORG_NAME = "SimzWord"
COMMENT_PROP = int(QTextFormat.Property.UserProperty) + 1
TOC_PROP = int(QTextFormat.Property.UserProperty) + 2
TOC_END_PROP = int(QTextFormat.Property.UserProperty) + 3
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

PAGE_SIZES = {
    "Letter": (8.5, 11.0),
    "A4": (8.27, 11.69),
    "Legal": (8.5, 14.0),
    "A5": (5.83, 8.27),
    "Tabloid": (11.0, 17.0),
    "Custom": (8.5, 11.0),
}

STYLE_DEFS = {
    "Normal": dict(level=0, size=11, bold=False, italic=False, color="#111111", before=0, after=6, line=100, indent=0),
    "Title": dict(level=0, size=28, bold=True, italic=False, color="#1F3864", before=0, after=10, line=100, indent=0),
    "Subtitle": dict(level=0, size=14, bold=False, italic=True, color="#555555", before=0, after=12, line=100, indent=0),
    "Heading 1": dict(level=1, size=17, bold=True, italic=False, color="#1F4E79", before=14, after=6, line=100, indent=0),
    "Heading 2": dict(level=2, size=14, bold=True, italic=False, color="#2E74B5", before=11, after=4, line=100, indent=0),
    "Heading 3": dict(level=3, size=12, bold=True, italic=False, color="#1F4E79", before=8, after=3, line=100, indent=0),
    "Quote": dict(level=0, size=11, bold=False, italic=True, color="#444444", before=4, after=8, line=100, indent=1),
    "Intense Quote": dict(level=0, size=12, bold=True, italic=True, color="#1F4E79", before=6, after=10, line=100, indent=1),
    "List Paragraph": dict(level=0, size=11, bold=False, italic=False, color="#111111", before=0, after=3, line=100, indent=0),
}

FALLBACK_WORDS = set("""
the be to of and a in that have i it for not on with he as you do at this
but his by from they we say her she or an will my one all would there their
what so up out if about who get which go me when make can like time no just
him know take people into year your good some could them see other than then
now look only come its over think also back after use two how our work first
well way even new want because any these give day most us is was are were been
has had did does doing done would should could may might must shall will can
document text word words page pages format font style table picture insert
simzword office professional writing editor paragraph line space color view
print file save open new edit tools help home design layout review references
""".split())


def load_word_file(name):
    path = os.path.join(BASE_DIR, name)
    try:
        with open(path, "r", encoding="utf-8", errors="ignore") as handle:
            return {line.strip().lower() for line in handle if line.strip()}
    except Exception:
        return set()


WORD_SET = load_word_file("simzword_words.txt") or set(FALLBACK_WORDS)
COMMON_WORDS = [w for w in load_word_file("simzword_common.txt") if w.isalpha()]
if not COMMON_WORDS:
    COMMON_WORDS = sorted(FALLBACK_WORDS)

SYNONYMS = {
    "good": ["excellent", "fine", "great", "positive", "satisfactory", "superior", "valuable"],
    "bad": ["poor", "inferior", "unacceptable", "negative", "awful", "harmful"],
    "big": ["large", "huge", "enormous", "substantial", "massive", "considerable"],
    "small": ["little", "tiny", "minor", "compact", "slight", "modest"],
    "happy": ["glad", "pleased", "delighted", "cheerful", "content", "joyful"],
    "sad": ["unhappy", "sorrowful", "gloomy", "downcast", "miserable"],
    "fast": ["quick", "rapid", "swift", "speedy", "brisk"],
    "slow": ["sluggish", "gradual", "unhurried", "delayed", "leisurely"],
    "important": ["significant", "essential", "critical", "vital", "major"],
    "difficult": ["hard", "challenging", "demanding", "tough", "complex"],
    "easy": ["simple", "effortless", "straightforward", "uncomplicated"],
    "beautiful": ["attractive", "lovely", "gorgeous", "elegant", "stunning"],
    "smart": ["intelligent", "clever", "bright", "brilliant", "sharp"],
    "start": ["begin", "commence", "initiate", "launch", "open"],
    "end": ["finish", "conclude", "complete", "terminate", "close"],
    "show": ["display", "demonstrate", "present", "reveal", "illustrate"],
    "make": ["create", "build", "produce", "construct", "generate"],
    "use": ["utilize", "employ", "apply", "operate", "exercise"],
    "help": ["assist", "support", "aid", "facilitate", "guide"],
    "change": ["alter", "modify", "adjust", "transform", "revise"],
    "think": ["believe", "consider", "reflect", "suppose", "contemplate"],
    "say": ["state", "declare", "mention", "express", "remark"],
    "get": ["obtain", "acquire", "receive", "gain", "secure"],
    "keep": ["retain", "maintain", "preserve", "hold", "sustain"],
    "problem": ["issue", "difficulty", "challenge", "obstacle", "concern"],
    "answer": ["response", "reply", "solution", "reaction", "explanation"],
    "idea": ["concept", "notion", "thought", "plan", "suggestion"],
    "work": ["labor", "effort", "task", "job", "activity"],
    "very": ["extremely", "highly", "exceedingly", "particularly", "especially"],
    "also": ["additionally", "furthermore", "likewise", "moreover", "too"],
    "but": ["however", "nevertheless", "yet", "although", "conversely"],
    "because": ["since", "as", "owing to", "due to", "for the reason that"],
    "quick": ["fast", "rapid", "swift", "prompt", "speedy"],
    "clean": ["tidy", "spotless", "pure", "neat", "clear"],
    "strong": ["powerful", "robust", "sturdy", "intense", "forceful"],
    "weak": ["fragile", "feeble", "faint", "delicate", "unstable"],
    "old": ["ancient", "aged", "former", "previous", "vintage"],
    "new": ["fresh", "modern", "recent", "novel", "current"],
    "begin": ["start", "commence", "initiate", "launch", "open"],
    "finish": ["end", "complete", "conclude", "finalize", "close"],
    "buy": ["purchase", "acquire", "obtain", "procure"],
    "sell": ["market", "vend", "offer", "trade", "retail"],
    "speak": ["talk", "communicate", "express", "converse", "address"],
    "build": ["construct", "create", "assemble", "develop", "establish"],
}


def make_logo_pixmap(size=64):
    pm = QPixmap(size, size)
    pm.fill(Qt.GlobalColor.transparent)
    painter = QPainter(pm)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing)
    gradient = QLinearGradient(0, 0, size, size)
    gradient.setColorAt(0.0, QColor("#2B7CD3"))
    gradient.setColorAt(1.0, QColor("#6F42C1"))
    painter.setBrush(QBrush(gradient))
    painter.setPen(Qt.PenStyle.NoPen)
    painter.drawRoundedRect(QRectF(2, 2, size - 4, size - 4), size * 0.24, size * 0.24)
    painter.setPen(QColor("#FFFFFF"))
    font = QFont("Segoe UI", int(size * 0.52), QFont.Weight.Bold)
    painter.setFont(font)
    painter.drawText(pm.rect(), Qt.AlignmentFlag.AlignCenter, "S")
    painter.end()
    return pm


def make_text_icon(text, style="", color="#333333", size=24):
    pm = QPixmap(size, size)
    pm.fill(Qt.GlobalColor.transparent)
    painter = QPainter(pm)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing)
    font = QFont("Times New Roman", int(size * 0.58))
    if style == "bold":
        font.setBold(True)
    elif style == "italic":
        font.setItalic(True)
    elif style == "underline":
        font.setUnderline(True)
    elif style == "strike":
        font.setStrikeOut(True)
    painter.setFont(font)
    painter.setPen(QColor(color))
    painter.drawText(pm.rect(), Qt.AlignmentFlag.AlignCenter, text)
    if style == "strike":
        painter.drawLine(4, size // 2, size - 4, size // 2)
    painter.end()
    return QIcon(pm)


def make_align_icon(align, size=24, color="#444444"):
    pm = QPixmap(size, size)
    pm.fill(Qt.GlobalColor.transparent)
    painter = QPainter(pm)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing)
    pen = QPen(QColor(color))
    pen.setWidth(2)
    painter.setPen(pen)
    lengths = [16, 11, 16, 11]
    for i, length in enumerate(lengths):
        y = 6 + i * 4
        if align == "left":
            x = 4
        elif align == "right":
            x = size - 4 - length
        elif align == "center":
            x = (size - length) // 2
        else:
            x, length = 4, 16
        painter.drawLine(int(x), y, int(x + length), y)
    painter.end()
    return QIcon(pm)


def make_list_icon(numbered=False, size=24, color="#444444"):
    pm = QPixmap(size, size)
    pm.fill(Qt.GlobalColor.transparent)
    painter = QPainter(pm)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing)
    pen = QPen(QColor(color))
    pen.setWidth(2)
    painter.setPen(pen)
    for i in range(3):
        y = 7 + i * 6
        if numbered:
            painter.drawText(QRect(0, y - 6, 8, 12), Qt.AlignmentFlag.AlignCenter, str(i + 1))
        else:
            painter.setBrush(QColor(color))
            painter.drawEllipse(QPoint(4, y - 1), 1, 1)
        painter.drawLine(11, y - 1, size - 1, y - 1)
    painter.end()
    return QIcon(pm)


def make_color_icon(color, letter="A", size=22):
    pm = QPixmap(size, size)
    pm.fill(Qt.GlobalColor.transparent)
    painter = QPainter(pm)
    font = QFont("Segoe UI", 10, QFont.Weight.Bold)
    painter.setFont(font)
    painter.setPen(QColor("#444444"))
    painter.drawText(QRect(0, 0, size, size - 8), Qt.AlignmentFlag.AlignCenter, letter)
    painter.fillRect(2, size - 6, size - 4, 4, color)
    painter.end()
    return QIcon(pm)


def make_glyph_icon(glyph, color="#444444", size=24):
    pm = QPixmap(size, size)
    pm.fill(Qt.GlobalColor.transparent)
    painter = QPainter(pm)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing)
    painter.setPen(QColor(color))
    painter.setFont(QFont("Segoe UI Symbol", int(size * 0.55)))
    painter.drawText(pm.rect(), Qt.AlignmentFlag.AlignCenter, glyph)
    painter.end()
    return QIcon(pm)


def make_shape_pixmap(kind, color="#2B7CD3", width=220, height=140):
    pm = QPixmap(width, height)
    pm.fill(Qt.GlobalColor.transparent)
    painter = QPainter(pm)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing)
    painter.setPen(QPen(QColor(color).darker(120), 3))
    painter.setBrush(QColor(color))
    margin = 14
    rect = QRectF(margin, margin, width - margin * 2, height - margin * 2)
    if kind == "Rectangle":
        painter.drawRect(rect)
    elif kind == "Rounded Rectangle":
        painter.drawRoundedRect(rect, 24, 24)
    elif kind == "Ellipse":
        painter.drawEllipse(rect)
    elif kind == "Triangle":
        points = [QPointF(width / 2, margin), QPointF(width - margin, height - margin), QPointF(margin, height - margin)]
        painter.drawPolygon(QPolygonF(points))
    elif kind == "Arrow":
        points = [
            QPointF(margin, height / 2 - 18), QPointF(width * 0.62, height / 2 - 18),
            QPointF(width * 0.62, margin), QPointF(width - margin, height / 2),
            QPointF(width * 0.62, height - margin), QPointF(width * 0.62, height / 2 + 18),
            QPointF(margin, height / 2 + 18),
        ]
        painter.drawPolygon(QPolygonF(points))
    elif kind == "Star":
        cx, cy = width / 2, height / 2
        outer, inner = min(width, height) * 0.42, min(width, height) * 0.18
        points = []
        for i in range(10):
            angle = math.radians(-90 + i * 36)
            radius = outer if i % 2 == 0 else inner
            points.append(QPointF(cx + radius * math.cos(angle), cy + radius * math.sin(angle)))
        painter.drawPolygon(QPolygonF(points))
    elif kind == "Line":
        painter.setBrush(Qt.BrushStyle.NoBrush)
        painter.drawLine(QPointF(margin, height - margin), QPointF(width - margin, margin))
    painter.end()
    return pm


class TableDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Insert Table")
        self.setMinimumWidth(320)
        layout = QVBoxLayout(self)
        grid = QGridLayout()
        grid.addWidget(QLabel("Rows:"), 0, 0)
        self.rows = QSpinBox()
        self.rows.setRange(1, 100)
        self.rows.setValue(3)
        grid.addWidget(self.rows, 0, 1)
        grid.addWidget(QLabel("Columns:"), 1, 0)
        self.cols = QSpinBox()
        self.cols.setRange(1, 30)
        self.cols.setValue(3)
        grid.addWidget(self.cols, 1, 1)
        self.header = QCheckBox("Use first row as a header")
        self.header.setChecked(True)
        grid.addWidget(self.header, 2, 0, 1, 2)
        self.style_combo = QComboBox()
        self.style_combo.addItem("Grid", "grid")
        self.style_combo.addItem("Plain", "plain")
        grid.addWidget(QLabel("Style:"), 3, 0)
        grid.addWidget(self.style_combo, 3, 1)
        layout.addLayout(grid)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)


class ImageGalleryDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Insert Picture")
        self.setMinimumWidth(540)
        self.selected_path = None
        self.generated_pix = None
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("Choose a sample image or browse your computer:"))
        grid = QGridLayout()
        colors = [("#FF6B6B", "#4ECDC4"), ("#45B7D1", "#96CEB4"), ("#FFEEAD", "#FF6B6B"), ("#A2D5F2", "#07689F")]
        titles = ["Gradient", "Chart", "Geometry", "Logo"]
        for i in range(4):
            pm = QPixmap(120, 80)
            painter = QPainter(pm)
            if i == 0:
                gradient = QLinearGradient(0, 0, 120, 80)
                gradient.setColorAt(0, QColor(colors[i][0]))
                gradient.setColorAt(1, QColor(colors[i][1]))
                painter.fillRect(0, 0, 120, 80, QBrush(gradient))
            elif i == 1:
                painter.fillRect(0, 0, 120, 80, QColor("#F8F9FA"))
                painter.setBrush(QColor("#4ECDC4"))
                painter.drawRect(20, 50, 15, 20)
                painter.drawRect(45, 30, 15, 40)
                painter.drawRect(70, 10, 15, 60)
            elif i == 2:
                painter.fillRect(0, 0, 120, 80, QColor("#2B7CD3"))
                painter.setBrush(QColor("#FFFFFF"))
                painter.drawEllipse(40, 20, 40, 40)
            else:
                painter.fillRect(0, 0, 120, 80, QColor("#333333"))
                painter.setPen(QColor("#FFFFFF"))
                painter.setFont(QFont("Segoe UI", 16, QFont.Weight.Bold))
                painter.drawText(pm.rect(), Qt.AlignmentFlag.AlignCenter, "LOGO")
            painter.end()
            button = QPushButton()
            button.setIcon(QIcon(pm))
            button.setIconSize(QSize(120, 80))
            button.setFixedSize(130, 90)
            button.setToolTip(titles[i])
            button.clicked.connect(lambda checked=False, pix=pm: self.select_default(pix))
            grid.addWidget(button, 0, i)
        layout.addLayout(grid)
        browse = QPushButton("Browse Computer...")
        browse.clicked.connect(self.browse)
        layout.addWidget(browse)
        cancel = QPushButton("Cancel")
        cancel.clicked.connect(self.reject)
        layout.addWidget(cancel)

    def select_default(self, pixmap):
        self.generated_pix = pixmap
        self.accept()

    def browse(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select Image", "", "Images (*.png *.jpg *.jpeg *.bmp *.gif)")
        if path:
            self.selected_path = path
            self.accept()


class FindReplaceDialog(QDialog):
    def __init__(self, window, parent=None):
        super().__init__(parent or window)
        self.win = window
        self.setWindowTitle("Find and Replace")
        self.setMinimumWidth(440)
        self.setModal(False)
        layout = QVBoxLayout(self)
        row1 = QHBoxLayout()
        row1.addWidget(QLabel("Find:"))
        self.find_edit = QLineEdit()
        row1.addWidget(self.find_edit)
        layout.addLayout(row1)
        row2 = QHBoxLayout()
        row2.addWidget(QLabel("Replace:"))
        self.replace_edit = QLineEdit()
        row2.addWidget(self.replace_edit)
        layout.addLayout(row2)
        row3 = QHBoxLayout()
        self.case_cb = QCheckBox("Match case")
        self.whole_cb = QCheckBox("Whole words")
        self.regex_cb = QCheckBox("Use regular expression")
        row3.addWidget(self.case_cb)
        row3.addWidget(self.whole_cb)
        row3.addWidget(self.regex_cb)
        row3.addStretch()
        layout.addLayout(row3)
        buttons = QHBoxLayout()
        self.find_btn = QPushButton("Find Next")
        self.replace_btn = QPushButton("Replace")
        self.replace_all_btn = QPushButton("Replace All")
        self.count_btn = QPushButton("Count")
        for button in (self.find_btn, self.replace_btn, self.replace_all_btn, self.count_btn):
            buttons.addWidget(button)
        layout.addLayout(buttons)
        self.result_label = QLabel("")
        self.result_label.setStyleSheet("color: #666666;")
        layout.addWidget(self.result_label)
        self.find_btn.clicked.connect(lambda: self.find_next())
        self.replace_btn.clicked.connect(self.replace_one)
        self.replace_all_btn.clicked.connect(self.replace_all)
        self.count_btn.clicked.connect(self.count_matches)
        self.find_edit.returnPressed.connect(lambda: self.find_next())
        self.replace_edit.returnPressed.connect(self.replace_one)

    @property
    def editor(self):
        return self.win.editor

    def _flags(self):
        flags = QTextDocument.FindFlag(0)
        if self.case_cb.isChecked():
            flags |= QTextDocument.FindFlag.FindCaseSensitively
        if self.whole_cb.isChecked():
            flags |= QTextDocument.FindFlag.FindWholeWords
        return flags

    def _regex(self):
        text = self.find_edit.text()
        if not text:
            return None
        pattern = text
        if self.whole_cb.isChecked():
            pattern = r"\b(?:" + pattern + r")\b"
        options = QRegularExpression.PatternOption.NoPatternOption
        if not self.case_cb.isChecked():
            options |= QRegularExpression.PatternOption.CaseInsensitiveOption
        regex = QRegularExpression(pattern, options)
        if not regex.isValid():
            QMessageBox.warning(self, "Invalid Expression", regex.errorString())
            return None
        return regex

    def find_next(self, from_start=False):
        editor = self.editor
        if editor is None:
            return False
        if not self.find_edit.text():
            return False
        if self.regex_cb.isChecked():
            regex = self._regex()
            if regex is None:
                return False
            if from_start:
                cursor = editor.textCursor()
                cursor.movePosition(QTextCursor.MoveOperation.Start)
                editor.setTextCursor(cursor)
            current = editor.textCursor()
            position = 0 if from_start else (current.selectionEnd() if current.hasSelection() else current.position())
            found = editor.document().find(regex, position)
            if found.isNull() and position > 0:
                found = editor.document().find(regex, 0)
            if found.isNull():
                self.result_label.setText("No matches found.")
                return False
            editor.setTextCursor(found)
            editor.ensureCursorVisible()
            self.result_label.setText("Match found.")
            return True
        if from_start:
            cursor = editor.textCursor()
            cursor.movePosition(QTextCursor.MoveOperation.Start)
            editor.setTextCursor(cursor)
        found = editor.find(self.find_edit.text(), self._flags())
        if not found:
            cursor = editor.textCursor()
            cursor.movePosition(QTextCursor.MoveOperation.Start)
            editor.setTextCursor(cursor)
            found = editor.find(self.find_edit.text(), self._flags())
        if found:
            editor.ensureCursorVisible()
            self.result_label.setText("Match found.")
        else:
            self.result_label.setText("No matches found.")
        return found

    def _selection_matches(self):
        editor = self.editor
        cursor = editor.textCursor()
        if not cursor.hasSelection():
            return False
        selected = cursor.selectedText()
        if self.regex_cb.isChecked():
            regex = self._regex()
            if regex is None:
                return False
            match = regex.match(selected)
            return match.hasMatch() and match.capturedLength() == len(selected)
        if self.case_cb.isChecked():
            return selected == self.find_edit.text()
        return selected.lower() == self.find_edit.text().lower()

    def replace_one(self):
        editor = self.editor
        if editor is None:
            return
        if not self._selection_matches():
            self.find_next()
            return
        cursor = editor.textCursor()
        cursor.insertText(self.replace_edit.text())
        editor.setTextCursor(cursor)
        self.find_next()

    def replace_all(self):
        editor = self.editor
        if editor is None or not self.find_edit.text():
            return
        replacement = self.replace_edit.text()
        count = 0
        cursor = editor.textCursor()
        cursor.beginEditBlock()
        cursor.movePosition(QTextCursor.MoveOperation.Start)
        editor.setTextCursor(cursor)
        if self.regex_cb.isChecked():
            regex = self._regex()
            if regex is None:
                cursor.endEditBlock()
                return
            last_position = -1
            while count < 10000:
                current = editor.textCursor()
                found = editor.document().find(regex, current.position())
                if found.isNull():
                    break
                if found.position() <= last_position:
                    move = QTextCursor(editor.document())
                    move.setPosition(min(found.position() + 1, editor.document().characterCount() - 1))
                    editor.setTextCursor(move)
                    continue
                last_position = found.position()
                editor.setTextCursor(found)
                found.insertText(replacement)
                count += 1
        else:
            flags = self._flags()
            last_position = -1
            while count < 10000 and editor.find(self.find_edit.text(), flags):
                found = editor.textCursor()
                if found.selectionStart() <= last_position:
                    move = QTextCursor(editor.document())
                    move.setPosition(min(found.selectionEnd() + 1, editor.document().characterCount() - 1))
                    editor.setTextCursor(move)
                    continue
                last_position = found.selectionStart()
                found.insertText(replacement)
                editor.setTextCursor(found)
                count += 1
        cursor.endEditBlock()
        self.result_label.setText(f"Replaced {count} occurrence(s).")

    def count_matches(self):
        editor = self.editor
        if editor is None or not self.find_edit.text():
            return
        original = editor.textCursor()
        count = 0
        if self.regex_cb.isChecked():
            regex = self._regex()
            if regex is None:
                return
            position = 0
            while True:
                found = editor.document().find(regex, position)
                if found.isNull():
                    break
                count += 1
                position = found.selectionEnd() if found.selectionEnd() > found.selectionStart() else found.selectionEnd() + 1
                if position >= editor.document().characterCount():
                    break
        else:
            flags = self._flags()
            cursor = editor.textCursor()
            cursor.movePosition(QTextCursor.MoveOperation.Start)
            editor.setTextCursor(cursor)
            while editor.find(self.find_edit.text(), flags):
                count += 1
        editor.setTextCursor(original)
        self.result_label.setText(f"{count} match(es) found.")


class PageSetupDialog(QDialog):
    def __init__(self, parent, page_width, page_height, margin):
        super().__init__(parent)
        self.setWindowTitle("Page Setup")
        self.setMinimumWidth(360)
        layout = QVBoxLayout(self)
        grid = QGridLayout()
        grid.addWidget(QLabel("Paper size:"), 0, 0)
        self.size_combo = QComboBox()
        for name, dims in PAGE_SIZES.items():
            self.size_combo.addItem(name, dims)
        self.size_combo.addItem("Custom", None)
        grid.addWidget(self.size_combo, 0, 1)
        grid.addWidget(QLabel("Orientation:"), 1, 0)
        self.orientation = QComboBox()
        self.orientation.addItem("Portrait", "portrait")
        self.orientation.addItem("Landscape", "landscape")
        grid.addWidget(self.orientation, 1, 1)
        grid.addWidget(QLabel("Top/bottom/left/right margin (in):"), 2, 0)
        self.margin_spin = QDoubleSpinBox()
        self.margin_spin.setRange(0.1, 3.0)
        self.margin_spin.setSingleStep(0.1)
        self.margin_spin.setDecimals(2)
        self.margin_spin.setValue(max(0.1, margin / 96.0))
        grid.addWidget(self.margin_spin, 2, 1)
        grid.addWidget(QLabel("Custom width (in):"), 3, 0)
        self.width_spin = QDoubleSpinBox()
        self.width_spin.setRange(3.0, 30.0)
        self.width_spin.setSingleStep(0.1)
        self.width_spin.setDecimals(2)
        self.width_spin.setValue(page_width / 96.0)
        grid.addWidget(self.width_spin, 3, 1)
        grid.addWidget(QLabel("Custom height (in):"), 4, 0)
        self.height_spin = QDoubleSpinBox()
        self.height_spin.setRange(3.0, 40.0)
        self.height_spin.setSingleStep(0.1)
        self.height_spin.setDecimals(2)
        self.height_spin.setValue(page_height / 96.0)
        grid.addWidget(self.height_spin, 4, 1)
        layout.addLayout(grid)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        self.size_combo.currentIndexChanged.connect(self._sync_custom)
        self._sync_custom()
        self._match_current(page_width, page_height)

    def _match_current(self, page_width, page_height):
        best = None
        best_diff = 1e9
        for name, dims in PAGE_SIZES.items():
            if name == "Custom":
                continue
            diff = abs(dims[0] * 96 - page_width) + abs(dims[1] * 96 - page_height)
            if diff < best_diff:
                best_diff = diff
                best = name
        if best and best_diff < 10:
            self.size_combo.setCurrentText(best)
        else:
            self.size_combo.setCurrentText("Custom")
        if page_width > page_height:
            self.orientation.setCurrentIndex(1)

    def _sync_custom(self):
        custom = self.size_combo.currentText() == "Custom"
        self.width_spin.setEnabled(custom)
        self.height_spin.setEnabled(custom)

    def result_data(self):
        width_in, height_in = self.size_combo.currentData() or (self.width_spin.value(), self.height_spin.value())
        if self.orientation.currentData() == "landscape":
            width_in, height_in = max(width_in, height_in), min(width_in, height_in)
        else:
            width_in, height_in = min(width_in, height_in), max(width_in, height_in)
        return {
            "width": int(width_in * 96),
            "height": int(height_in * 96),
            "margin": int(self.margin_spin.value() * 96),
        }


class HeaderFooterDialog(QDialog):
    def __init__(self, parent, header_text, footer_text, show_header, show_footer, show_numbers, number_format):
        super().__init__(parent)
        self.setWindowTitle("Header and Footer")
        self.setMinimumWidth(430)
        layout = QVBoxLayout(self)
        self.header_cb = QCheckBox("Show header")
        self.header_cb.setChecked(show_header)
        layout.addWidget(self.header_cb)
        self.header_edit = QLineEdit(header_text)
        self.header_edit.setPlaceholderText("Header text")
        layout.addWidget(self.header_edit)
        self.footer_cb = QCheckBox("Show footer")
        self.footer_cb.setChecked(show_footer)
        layout.addWidget(self.footer_cb)
        self.footer_edit = QLineEdit(footer_text)
        self.footer_edit.setPlaceholderText("Footer text")
        layout.addWidget(self.footer_edit)
        self.numbers_cb = QCheckBox("Show page numbers")
        self.numbers_cb.setChecked(show_numbers)
        layout.addWidget(self.numbers_cb)
        row = QHBoxLayout()
        row.addWidget(QLabel("Page number format:"))
        self.format_combo = QComboBox()
        for item in ("Page X of Y", "Page X", "X of Y", "X"):
            self.format_combo.addItem(item)
        self.format_combo.setCurrentText(number_format)
        row.addWidget(self.format_combo, 1)
        layout.addLayout(row)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def result_data(self):
        return {
            "header_text": self.header_edit.text(),
            "footer_text": self.footer_edit.text(),
            "show_header": self.header_cb.isChecked(),
            "show_footer": self.footer_cb.isChecked(),
            "show_numbers": self.numbers_cb.isChecked(),
            "number_format": self.format_combo.currentText(),
        }


class WatermarkDialog(QDialog):
    def __init__(self, parent, text="", size=60, color=QColor(180, 180, 180), opacity=18):
        super().__init__(parent)
        self.setWindowTitle("Watermark")
        self.setMinimumWidth(380)
        self.color = QColor(color)
        layout = QVBoxLayout(self)
        self.text_edit = QLineEdit(text)
        self.text_edit.setPlaceholderText("Watermark text")
        layout.addWidget(self.text_edit)
        row = QHBoxLayout()
        row.addWidget(QLabel("Font size:"))
        self.size_spin = QSpinBox()
        self.size_spin.setRange(16, 160)
        self.size_spin.setValue(size)
        row.addWidget(self.size_spin)
        row.addWidget(QLabel("Opacity:"))
        self.opacity_slider = QSlider(Qt.Orientation.Horizontal)
        self.opacity_slider.setRange(5, 60)
        self.opacity_slider.setValue(opacity)
        row.addWidget(self.opacity_slider, 1)
        layout.addLayout(row)
        color_row = QHBoxLayout()
        color_row.addWidget(QLabel("Color:"))
        self.color_button = QPushButton("Choose Color")
        self.color_button.clicked.connect(self.choose_color)
        color_row.addWidget(self.color_button)
        color_row.addStretch()
        layout.addLayout(color_row)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        self._update_color_button()

    def choose_color(self):
        color = QColorDialog.getColor(self.color, self, "Watermark Color")
        if color.isValid():
            self.color = color
            self._update_color_button()

    def _update_color_button(self):
        self.color_button.setStyleSheet(
            f"background: {self.color.name()}; color: {'#000' if self.color.lightness() > 128 else '#fff'};"
        )

    def result_data(self):
        return {
            "text": self.text_edit.text(),
            "size": self.size_spin.value(),
            "color": self.color.name(),
            "opacity": self.opacity_slider.value(),
        }


class ParagraphDialog(QDialog):
    def __init__(self, parent, alignment, left_indent, first_indent, before, after, line_height):
        super().__init__(parent)
        self.setWindowTitle("Paragraph")
        self.setMinimumWidth(380)
        layout = QVBoxLayout(self)
        grid = QGridLayout()
        grid.addWidget(QLabel("Alignment:"), 0, 0)
        self.align_combo = QComboBox()
        for label, value in (("Left", "left"), ("Center", "center"), ("Right", "right"), ("Justify", "justify")):
            self.align_combo.addItem(label, value)
        self.align_combo.setCurrentIndex(max(0, ["left", "center", "right", "justify"].index(alignment) if alignment in ("left", "center", "right", "justify") else 0))
        grid.addWidget(self.align_combo, 0, 1)
        grid.addWidget(QLabel("Left indent (in):"), 1, 0)
        self.left_spin = QDoubleSpinBox()
        self.left_spin.setRange(0, 5)
        self.left_spin.setSingleStep(0.1)
        self.left_spin.setDecimals(2)
        self.left_spin.setValue(left_indent / 96.0)
        grid.addWidget(self.left_spin, 1, 1)
        grid.addWidget(QLabel("First line indent (in):"), 2, 0)
        self.first_spin = QDoubleSpinBox()
        self.first_spin.setRange(-2, 2)
        self.first_spin.setSingleStep(0.1)
        self.first_spin.setDecimals(2)
        self.first_spin.setValue(first_indent / 96.0)
        grid.addWidget(self.first_spin, 2, 1)
        grid.addWidget(QLabel("Space before (pt):"), 3, 0)
        self.before_spin = QDoubleSpinBox()
        self.before_spin.setRange(0, 100)
        self.before_spin.setValue(before)
        grid.addWidget(self.before_spin, 3, 1)
        grid.addWidget(QLabel("Space after (pt):"), 4, 0)
        self.after_spin = QDoubleSpinBox()
        self.after_spin.setRange(0, 100)
        self.after_spin.setValue(after)
        grid.addWidget(self.after_spin, 4, 1)
        grid.addWidget(QLabel("Line spacing:"), 5, 0)
        self.line_combo = QComboBox()
        for label, value in (("Single", 100), ("1.15", 115), ("1.5", 150), ("Double", 200)):
            self.line_combo.addItem(label, value)
        self.line_combo.setCurrentIndex(max(0, [100, 115, 150, 200].index(line_height) if line_height in (100, 115, 150, 200) else 0))
        grid.addWidget(self.line_combo, 5, 1)
        layout.addLayout(grid)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def result_data(self):
        return {
            "alignment": self.align_combo.currentData(),
            "left_indent": int(self.left_spin.value() * 96),
            "first_indent": int(self.first_spin.value() * 96),
            "before": self.before_spin.value(),
            "after": self.after_spin.value(),
            "line_height": self.line_combo.currentData(),
        }


class WordCountDialog(QDialog):
    def __init__(self, parent, text, pages):
        super().__init__(parent)
        self.setWindowTitle("Word Count")
        self.setMinimumWidth(340)
        layout = QVBoxLayout(self)
        words = len(re.findall(r"\S+", text))
        chars_no_spaces = len(re.sub(r"\s", "", text))
        paragraphs = len([b for b in text.split("\n") if b.strip()])
        lines = len(text.splitlines()) or 1
        reading = max(1, math.ceil(words / 200)) if words else 0
        data = [
            ("Pages", str(pages)),
            ("Words", str(words)),
            ("Characters (no spaces)", str(chars_no_spaces)),
            ("Characters (with spaces)", str(len(text))),
            ("Paragraphs", str(paragraphs)),
            ("Lines", str(lines)),
            ("Reading time", f"{reading} min"),
        ]
        grid = QGridLayout()
        for row, (label, value) in enumerate(data):
            grid.addWidget(QLabel(label + ":"), row, 0)
            value_label = QLabel(value)
            value_label.setStyleSheet("font-weight: bold;")
            grid.addWidget(value_label, row, 1)
        layout.addLayout(grid)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        buttons.rejected.connect(self.reject)
        buttons.clicked.connect(self.accept)
        layout.addWidget(buttons)


def edit_distance(a, b, limit=2):
    if abs(len(a) - len(b)) > limit:
        return limit + 1
    previous = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        current = [i]
        minimum = current[0]
        for j, cb in enumerate(b, 1):
            cost = 0 if ca == cb else 1
            value = min(previous[j] + 1, current[j - 1] + 1, previous[j - 1] + cost)
            current.append(value)
            minimum = min(minimum, value)
        if minimum > limit:
            return limit + 1
        previous = current
    return previous[-1]


class SpellingDialog(QDialog):
    def __init__(self, window, parent=None):
        super().__init__(parent or window)
        self.win = window
        self.setWindowTitle("Spelling")
        self.setMinimumSize(680, 420)
        self.ignored = set()
        self.suggestions_cache = {}
        layout = QVBoxLayout(self)
        info = QLabel("Words not found in the SimzWord dictionary:")
        layout.addWidget(info)
        body = QHBoxLayout()
        self.word_list = QListWidget()
        self.word_list.setFixedWidth(210)
        body.addWidget(self.word_list)
        right = QVBoxLayout()
        right.addWidget(QLabel("Suggestions:"))
        self.suggestion_list = QListWidget()
        right.addWidget(self.suggestion_list, 1)
        replace_row = QHBoxLayout()
        replace_row.addWidget(QLabel("Change to:"))
        self.replacement_edit = QLineEdit()
        replace_row.addWidget(self.replacement_edit, 1)
        right.addLayout(replace_row)
        buttons = QHBoxLayout()
        self.change_btn = QPushButton("Change All")
        self.ignore_btn = QPushButton("Ignore")
        self.ignore_all_btn = QPushButton("Ignore All")
        self.add_btn = QPushButton("Add to Dictionary")
        self.recheck_btn = QPushButton("Recheck")
        for button in (self.change_btn, self.ignore_btn, self.ignore_all_btn, self.add_btn, self.recheck_btn):
            buttons.addWidget(button)
        right.addLayout(buttons)
        body.addLayout(right, 1)
        layout.addLayout(body, 1)
        self.status_label = QLabel("")
        layout.addWidget(self.status_label)
        close_row = QHBoxLayout()
        close_row.addStretch()
        close = QPushButton("Close")
        close.clicked.connect(self.accept)
        close_row.addWidget(close)
        layout.addLayout(close_row)
        self.word_list.currentTextChanged.connect(self._on_word_selected)
        self.suggestion_list.itemDoubleClicked.connect(lambda item: self.replacement_edit.setText(item.text()))
        self.change_btn.clicked.connect(self.change_word)
        self.ignore_btn.clicked.connect(self.ignore_word)
        self.ignore_all_btn.clicked.connect(self.ignore_word)
        self.add_btn.clicked.connect(self.add_word)
        self.recheck_btn.clicked.connect(self.rescan)
        self.rescan()

    @property
    def editor(self):
        return self.win.editor

    def _known(self, word):
        lower = word.lower()
        if lower in WORD_SET or lower in self.win.custom_words or lower in self.ignored:
            return True
        if lower.endswith("'s") and lower[:-2] in WORD_SET:
            return True
        if "'" in lower or "\u2019" in lower:
            parts = re.split(r"['\u2019]", lower)
            if all(part in WORD_SET or part in {"t", "s", "re", "ve", "ll", "d", "m"} for part in parts):
                return True
        return False

    def _scan(self):
        editor = self.editor
        if editor is None:
            return []
        found = []
        seen = set()
        for match in re.finditer(r"[A-Za-z]+(?:['\u2019][A-Za-z]+)?", editor.toPlainText()):
            word = match.group(0)
            if len(word) < 2:
                continue
            if word[0].isupper():
                continue
            if self._known(word):
                continue
            key = word.lower()
            if key in seen:
                continue
            seen.add(key)
            found.append(word)
        return found

    def rescan(self):
        words = self._scan()
        self.word_list.blockSignals(True)
        self.word_list.clear()
        self.word_list.addItems(words)
        self.word_list.blockSignals(False)
        if words:
            self.word_list.setCurrentRow(0)
            self.status_label.setText(f"{len(words)} possible spelling issue(s).")
        else:
            self.status_label.setText("No spelling issues found.")
            self.suggestion_list.clear()
            self.replacement_edit.clear()

    def _suggestions(self, word):
        lower = word.lower()
        if lower in self.suggestions_cache:
            return self.suggestions_cache[lower]
        results = []
        first = lower[:1]
        for candidate in COMMON_WORDS:
            if len(candidate) < 2:
                continue
            if abs(len(candidate) - len(lower)) > 2:
                continue
            if candidate[:1] != first and edit_distance(candidate[:2], lower[:2], 2) > 1:
                continue
            distance = edit_distance(lower, candidate, 2)
            if distance <= 2:
                results.append((distance, candidate))
        results.sort(key=lambda item: (item[0], len(item[1]), item[1]))
        suggestions = [item[1] for item in results[:10]]
        self.suggestions_cache[lower] = suggestions
        return suggestions

    def _on_word_selected(self, word):
        if not word:
            return
        suggestions = self._suggestions(word)
        self.suggestion_list.clear()
        self.suggestion_list.addItems(suggestions)
        if suggestions:
            self.suggestion_list.setCurrentRow(0)
            self.replacement_edit.setText(suggestions[0])
        else:
            self.replacement_edit.setText("")
        self._select_in_editor(word)

    def _select_in_editor(self, word):
        editor = self.editor
        if editor is None:
            return
        pattern = QRegularExpression(r"\b" + re.escape(word) + r"\b")
        found = editor.document().find(pattern, 0)
        if not found.isNull():
            editor.setTextCursor(found)
            editor.ensureCursorVisible()

    def change_word(self):
        word = self.word_list.currentItem().text() if self.word_list.currentItem() else ""
        replacement = self.replacement_edit.text()
        editor = self.editor
        if not word or not replacement or editor is None:
            return
        pattern = QRegularExpression(r"\b" + re.escape(word) + r"\b")
        position = 0
        count = 0
        cursor = editor.textCursor()
        cursor.beginEditBlock()
        while count < 5000:
            found = editor.document().find(pattern, position)
            if found.isNull():
                break
            found.insertText(replacement)
            position = found.position()
            count += 1
        cursor.endEditBlock()
        self.status_label.setText(f"Replaced {count} occurrence(s) of '{word}'.")
        self.rescan()

    def ignore_word(self):
        item = self.word_list.currentItem()
        if item is None:
            return
        self.ignored.add(item.text().lower())
        self.word_list.takeItem(self.word_list.row(item))
        if self.word_list.count() == 0:
            self.status_label.setText("No spelling issues found.")

    def add_word(self):
        item = self.word_list.currentItem()
        if item is None:
            return
        word = item.text()
        self.win.add_custom_word(word)
        self.ignored.add(word.lower())
        self.word_list.takeItem(self.word_list.row(item))
        self.status_label.setText(f"Added '{word}' to the dictionary.")


class ThesaurusDialog(QDialog):
    def __init__(self, window, parent=None, selected_text=""):
        super().__init__(parent or window)
        self.win = window
        self.setWindowTitle("Thesaurus")
        self.setMinimumWidth(430)
        layout = QVBoxLayout(self)
        row = QHBoxLayout()
        row.addWidget(QLabel("Word:"))
        self.word_edit = QLineEdit(selected_text)
        row.addWidget(self.word_edit, 1)
        search = QPushButton("Search")
        search.clicked.connect(self.search)
        row.addWidget(search)
        layout.addLayout(row)
        self.results = QListWidget()
        layout.addWidget(self.results, 1)
        self.status = QLabel("")
        layout.addWidget(self.status)
        buttons = QHBoxLayout()
        insert = QPushButton("Insert Synonym")
        insert.clicked.connect(self.insert_synonym)
        close = QPushButton("Close")
        close.clicked.connect(self.accept)
        buttons.addStretch()
        buttons.addWidget(insert)
        buttons.addWidget(close)
        layout.addLayout(buttons)
        self.results.itemDoubleClicked.connect(lambda item: self.insert_synonym())
        if selected_text:
            self.search()

    def search(self):
        word = self.word_edit.text().strip().lower()
        self.results.clear()
        if not word:
            return
        synonyms = SYNONYMS.get(word, [])
        if synonyms:
            self.results.addItems(synonyms)
            self.status.setText(f"{len(synonyms)} synonym(s) found.")
        else:
            self.status.setText("No synonyms found in the built-in thesaurus.")

    def insert_synonym(self):
        item = self.results.currentItem()
        if item is None or self.win.editor is None:
            return
        editor = self.win.editor
        cursor = editor.textCursor()
        if cursor.hasSelection():
            cursor.insertText(item.text())
        else:
            editor.insertPlainText(item.text())
        self.accept()


class ChartDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Insert Chart")
        self.setMinimumWidth(420)
        self.pixmap = None
        layout = QVBoxLayout(self)
        grid = QGridLayout()
        grid.addWidget(QLabel("Chart title:"), 0, 0)
        self.title_edit = QLineEdit("Monthly Sales")
        grid.addWidget(self.title_edit, 0, 1)
        grid.addWidget(QLabel("Values:"), 1, 0)
        self.values_edit = QLineEdit("12, 19, 8, 24, 17, 30")
        grid.addWidget(self.values_edit, 1, 1)
        grid.addWidget(QLabel("Labels:"), 2, 0)
        self.labels_edit = QLineEdit("Jan, Feb, Mar, Apr, May, Jun")
        grid.addWidget(self.labels_edit, 2, 1)
        grid.addWidget(QLabel("Chart type:"), 3, 0)
        self.type_combo = QComboBox()
        for name in ("Column", "Line", "Pie"):
            self.type_combo.addItem(name)
        grid.addWidget(self.type_combo, 3, 1)
        layout.addLayout(grid)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.build_chart)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _parse_values(self):
        raw = [part.strip() for part in self.values_edit.text().split(",") if part.strip()]
        values = []
        for part in raw:
            try:
                values.append(float(part))
            except ValueError:
                values.append(0.0)
        return values

    def build_chart(self):
        values = self._parse_values()
        if not values:
            QMessageBox.warning(self, "Insert Chart", "Please enter at least one numeric value.")
            return
        labels = [part.strip() for part in self.labels_edit.text().split(",") if part.strip()]
        while len(labels) < len(values):
            labels.append(str(len(labels) + 1))
        width, height = 600, 340
        pixmap = QPixmap(width, height)
        pixmap.fill(QColor("#FFFFFF"))
        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.setPen(QColor("#222222"))
        painter.setFont(QFont("Segoe UI", 13, QFont.Weight.Bold))
        painter.drawText(QRectF(0, 12, width, 28), Qt.AlignmentFlag.AlignCenter, self.title_edit.text() or "Chart")
        palette = [QColor("#2B7CD3"), QColor("#4ECDC4"), QColor("#FF6B6B"), QColor("#F7B731"), QColor("#6F42C1"), QColor("#20BF6B")]
        chart_type = self.type_combo.currentText()
        if chart_type == "Pie":
            total = sum(max(0.0, value) for value in values) or 1.0
            rect = QRectF(70, 70, 220, 220)
            start = 90 * 16
            legend_x = 330
            for index, value in enumerate(values):
                span = int(-(max(0.0, value) / total) * 360 * 16)
                painter.setBrush(palette[index % len(palette)])
                painter.setPen(QPen(QColor("#FFFFFF"), 2))
                painter.drawPie(rect, start, span)
                start += span
                painter.setPen(QColor("#333333"))
                painter.setBrush(palette[index % len(palette)])
                painter.drawRect(QRectF(legend_x, 90 + index * 26, 14, 14))
                painter.drawText(QRectF(legend_x + 22, 86 + index * 26, 220, 22), Qt.AlignmentFlag.AlignLeft, f"{labels[index]}  ({value:g})")
        else:
            plot = QRectF(70, 60, width - 120, height - 130)
            max_value = max(values) or 1.0
            painter.setPen(QPen(QColor("#CCCCCC"), 1))
            for step in range(5):
                y = plot.bottom() - step * plot.height() / 4
                painter.drawLine(QPointF(plot.left(), y), QPointF(plot.right(), y))
                painter.setPen(QColor("#666666"))
                painter.drawText(QRectF(20, y - 10, 44, 20), Qt.AlignmentFlag.AlignRight, f"{max_value * step / 4:.0f}")
                painter.setPen(QPen(QColor("#CCCCCC"), 1))
            painter.setPen(QPen(QColor("#555555"), 2))
            painter.drawLine(plot.bottomLeft(), plot.bottomRight())
            painter.drawLine(plot.bottomLeft(), plot.topLeft())
            count = len(values)
            if chart_type == "Column":
                gap = plot.width() / count
                bar_width = gap * 0.6
                for index, value in enumerate(values):
                    bar_height = (max(0.0, value) / max_value) * plot.height()
                    rect = QRectF(plot.left() + index * gap + (gap - bar_width) / 2, plot.bottom() - bar_height, bar_width, bar_height)
                    painter.setBrush(palette[index % len(palette)])
                    painter.setPen(QPen(QColor("#FFFFFF"), 1))
                    painter.drawRect(rect)
                    painter.setPen(QColor("#333333"))
                    painter.drawText(QRectF(rect.left() - 8, rect.top() - 20, rect.width() + 16, 18), Qt.AlignmentFlag.AlignCenter, f"{value:g}")
            else:
                points = []
                for index, value in enumerate(values):
                    x = plot.left() + (plot.width() / max(1, count - 1)) * index
                    y = plot.bottom() - (max(0.0, value) / max_value) * plot.height()
                    points.append(QPointF(x, y))
                painter.setPen(QPen(palette[0], 3))
                for index in range(len(points) - 1):
                    painter.drawLine(points[index], points[index + 1])
                painter.setBrush(palette[0])
                for point in points:
                    painter.drawEllipse(point, 5, 5)
            for index, label in enumerate(labels[:count]):
                x = plot.left() + (plot.width() / max(1, count)) * (index + 0.5)
                painter.setPen(QColor("#444444"))
                painter.drawText(QRectF(x - 40, plot.bottom() + 8, 80, 20), Qt.AlignmentFlag.AlignCenter, label)
        painter.end()
        self.pixmap = pixmap
        self.accept()


class AboutDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"About {APP_NAME}")
        self.setMinimumWidth(520)
        layout = QVBoxLayout(self)
        top = QHBoxLayout()
        logo = QLabel()
        logo.setPixmap(make_logo_pixmap(72))
        top.addWidget(logo)
        text = QVBoxLayout()
        title = QLabel(APP_NAME)
        title.setStyleSheet("font-size: 24px; font-weight: bold; color: #2B7CD3;")
        text.addWidget(title)
        version = QLabel(f"Version {APP_VERSION} - Professional Word Processor")
        version.setStyleSheet("color: #666666;")
        text.addWidget(version)
        text.addStretch()
        top.addLayout(text, 1)
        layout.addLayout(top)
        description = QLabel(
            "SimzWord is a modern, feature-rich word processor for everyday writing, "
            "professional documents, reports, letters and more. It combines a familiar "
            "ribbon interface with Word .docx support, a native .simz project format, "
            "editable PDF text import, rich formatting, tables, pictures, charts, "
            "comments, review tools and print-ready output."
        )
        description.setWordWrap(True)
        layout.addWidget(description)
        creator = QLabel(
            "Created by <b>William Siame</b>, a software developer and machine learning "
            "engineer from Zambia.<br>Email: "
            '<a href="mailto:siamewilliam10@gmail.com">siamewilliam10@gmail.com</a>'
        )
        creator.setTextFormat(Qt.TextFormat.RichText)
        creator.setOpenExternalLinks(True)
        creator.setWordWrap(True)
        layout.addWidget(creator)
        features = QLabel(
            "Highlights\n"
            "  -  Multiple document tabs, .docx support and native .simz projects\n"
            "  -  Rich text, styles, bullets, numbering and alignment\n"
            "  -  Tables, pictures, shapes, charts and text boxes\n"
            "  -  Headers, footers, page numbers, watermarks and borders\n"
            "  -  Comments, spelling, thesaurus, word count and revision log\n"
            "  -  Table of contents, bookmarks, citations and captions\n"
            "  -  Find and replace with regular expressions\n"
            "  -  Open/edit text-based PDFs; export PDF, DOCX, HTML, Markdown and text\n"
            "  -  Light and dark themes, focus mode and full screen\n"
            "  -  Autosave and crash recovery"
        )
        features.setStyleSheet("font-family: 'Consolas', monospace; color: #333333;")
        layout.addWidget(features)
        copyright_label = QLabel("Copyright (c) 2026 SimzWord. Built with PySide6.")
        copyright_label.setStyleSheet("color: #777777;")
        layout.addWidget(copyright_label)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok)
        buttons.accepted.connect(self.accept)
        layout.addWidget(buttons)


class PreferencesDialog(QDialog):
    def __init__(self, parent, settings):
        super().__init__(parent)
        self.setWindowTitle("SimzWord Options")
        self.setMinimumWidth(430)
        layout = QVBoxLayout(self)
        form = QGridLayout()
        self.autosave_cb = QCheckBox("Enable autosave and crash recovery")
        self.autosave_cb.setChecked(settings.value("autosave/enabled", True, bool))
        form.addWidget(self.autosave_cb, 0, 0, 1, 2)
        form.addWidget(QLabel("Autosave interval (seconds):"), 1, 0)
        self.interval_spin = QSpinBox()
        self.interval_spin.setRange(15, 600)
        self.interval_spin.setValue(settings.value("autosave/interval", 60, int))
        form.addWidget(self.interval_spin, 1, 1)
        form.addWidget(QLabel("Theme:"), 2, 0)
        self.theme_combo = QComboBox()
        self.theme_combo.addItem("Light", "light")
        self.theme_combo.addItem("Dark", "dark")
        self.theme_combo.setCurrentIndex(1 if settings.value("ui/dark", False, bool) else 0)
        form.addWidget(self.theme_combo, 2, 1)
        self.ruler_cb = QCheckBox("Show horizontal ruler")
        self.ruler_cb.setChecked(settings.value("ui/ruler", True, bool))
        form.addWidget(self.ruler_cb, 3, 0, 1, 2)
        self.status_cb = QCheckBox("Show status bar")
        self.status_cb.setChecked(settings.value("ui/statusbar", True, bool))
        form.addWidget(self.status_cb, 4, 0, 1, 2)
        self.nav_cb = QCheckBox("Show navigation pane at startup")
        self.nav_cb.setChecked(settings.value("ui/navigation", True, bool))
        form.addWidget(self.nav_cb, 5, 0, 1, 2)
        form.addWidget(QLabel("Default font:"), 6, 0)
        self.font_combo = QFontComboBox()
        self.font_combo.setCurrentFont(QFont(settings.value("editor/font_family", "Calibri", str)))
        form.addWidget(self.font_combo, 6, 1)
        form.addWidget(QLabel("Default font size:"), 7, 0)
        self.size_spin = QSpinBox()
        self.size_spin.setRange(6, 72)
        self.size_spin.setValue(settings.value("editor/font_size", 11, int))
        form.addWidget(self.size_spin, 7, 1)
        layout.addLayout(form)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def result_data(self):
        return {
            "autosave_enabled": self.autosave_cb.isChecked(),
            "autosave_interval": self.interval_spin.value(),
            "dark": self.theme_combo.currentData() == "dark",
            "ruler": self.ruler_cb.isChecked(),
            "statusbar": self.status_cb.isChecked(),
            "navigation": self.nav_cb.isChecked(),
            "font_family": self.font_combo.currentFont().family(),
            "font_size": self.size_spin.value(),
        }


class ShortcutsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Keyboard Shortcuts")
        self.setMinimumSize(560, 520)
        layout = QVBoxLayout(self)
        text = QPlainTextEdit()
        text.setReadOnly(True)
        text.setPlainText(
            "FILE\n"
            "  Ctrl+N            New document tab\n"
            "  Ctrl+O            Open .docx, .pdf, .simz or text document\n"
            "  Ctrl+S            Save (default .docx; .simz for native projects)\n"
            "  Ctrl+Shift+S      Save As\n"
            "  Ctrl+W            Close document tab\n"
            "  Ctrl+P            Print\n"
            "  Ctrl+Q            Exit SimzWord\n\n"
            "EDIT\n"
            "  Ctrl+Z / Ctrl+Y   Undo / Redo\n"
            "  Ctrl+X / C / V    Cut / Copy / Paste\n"
            "  Ctrl+Shift+V      Paste as plain text\n"
            "  Ctrl+A            Select all\n"
            "  Ctrl+F            Find\n"
            "  Ctrl+H            Find and Replace\n"
            "  F3                Find next\n"
            "  Ctrl+G            Go to line\n\n"
            "FORMATTING\n"
            "  Ctrl+B / I / U    Bold / Italic / Underline\n"
            "  Ctrl+Shift+X      Strikethrough\n"
            "  Ctrl+L / E / R / J  Align left / center / right / justify\n"
            "  Ctrl+K            Insert hyperlink\n"
            "  Ctrl+Enter        Insert page break\n"
            "  Ctrl+Alt+M        New comment\n"
            "  Ctrl+Alt+F        Insert footnote\n\n"
            "VIEW\n"
            "  Ctrl++ / Ctrl+-   Zoom in / out\n"
            "  Ctrl+0            Reset zoom\n"
            "  F11               Full screen\n\n"
            "REVIEW\n"
            "  F7                Spelling\n"
            "  Shift+F7          Thesaurus\n"
        )
        layout.addWidget(text)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        buttons.rejected.connect(self.reject)
        buttons.clicked.connect(self.accept)
        layout.addWidget(buttons)


class SymbolDialog(QDialog):
    CATEGORIES = {
        "Common": "Â©Â®â„¢Â°Â±Ã—Ã·ÂµÂ§Â¶â€¢â€¦â€°â‚¬Â£Â¥Â¢â†â†’â†‘â†“â†”âˆžâ‰ˆâ‰ â‰¤â‰¥Â½Â¼Â¾â…“â…”Î±Î²Î³Î´Ï€Î©âˆ‘âˆšâˆ«âˆ‚",
        "Currency": "$â‚¬Â£Â¥Â¢â‚¹â‚½â‚©â‚ªâ‚«â‚´â‚¦â‚±â‚²â‚µâ‚¡",
        "Math": "Â±Ã—Ã·â‰ â‰ˆâ‰¤â‰¥âˆžâˆ‘âˆšâˆ«âˆ‚âˆ†âˆ‡âˆˆâˆ‰âˆ‹âˆŒâŠ‚âŠƒâŠ†âŠ‡âˆªâˆ©âˆ§âˆ¨Â¬âˆ€âˆƒâˆ…",
        "Arrows": "â†â†‘â†’â†“â†”â†•â†–â†—â†˜â†™â‡â‡‘â‡’â‡“â‡”â‡•â‡–â‡—â‡˜â‡™",
        "Greek": "Î±Î²Î³Î´ÎµÎ¶Î·Î¸Î¹ÎºÎ»Î¼Î½Î¾Î¿Ï€ÏÏƒÏ„Ï…Ï†Ï‡ÏˆÏ‰Î‘Î’Î“Î”Î•Î–Î—Î˜Î™ÎšÎ›ÎœÎÎžÎŸÎ Î¡Î£Î¤Î¥Î¦Î§Î¨Î©",
        "Punctuation": "â€œâ€â€˜â€™Â«Â»â€¹â€ºâ€“â€”â€¦â€¢â€ â€¡Â§Â¶Â©Â®â„¢",
    }

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Insert Symbol")
        self.setMinimumWidth(540)
        self.selected = None
        layout = QVBoxLayout(self)
        tabs = QTabWidget()
        for name, characters in self.CATEGORIES.items():
            page = QWidget()
            grid = QGridLayout(page)
            for index, character in enumerate(characters):
                button = QPushButton(character)
                button.setFixedSize(38, 34)
                button.setFont(QFont("Segoe UI Symbol", 12))
                button.clicked.connect(lambda checked=False, ch=character: self.choose(ch))
                grid.addWidget(button, index // 12, index % 12)
            tabs.addTab(page, name)
        layout.addWidget(tabs)
        close = QPushButton("Cancel")
        close.clicked.connect(self.reject)
        layout.addWidget(close)

    def choose(self, character):
        self.selected = character
        self.accept()


class BookmarkManagerDialog(QDialog):
    def __init__(self, window, parent=None):
        super().__init__(parent or window)
        self.win = window
        self.setWindowTitle("Bookmarks")
        self.setMinimumWidth(380)
        layout = QVBoxLayout(self)
        self.list_widget = QListWidget()
        layout.addWidget(self.list_widget, 1)
        buttons = QHBoxLayout()
        add = QPushButton("Add")
        delete = QPushButton("Delete")
        go = QPushButton("Go To")
        close = QPushButton("Close")
        add.clicked.connect(self.add_bookmark)
        delete.clicked.connect(self.delete_bookmark)
        go.clicked.connect(self.go_to_bookmark)
        close.clicked.connect(self.accept)
        for button in (add, delete, go):
            buttons.addWidget(button)
        buttons.addStretch()
        buttons.addWidget(close)
        layout.addLayout(buttons)
        self.list_widget.itemDoubleClicked.connect(lambda item: self.go_to_bookmark())
        self.refresh()

    def refresh(self):
        self.list_widget.clear()
        tab = self.win._current_tab()
        if tab is None:
            return
        for name, position in tab.bookmarks.items():
            item = QListWidgetItem(name)
            item.setData(Qt.ItemDataRole.UserRole, position)
            self.list_widget.addItem(item)

    def add_bookmark(self):
        tab = self.win._current_tab()
        if tab is None:
            return
        name, ok = QInputDialog.getText(self, "Add Bookmark", "Bookmark name:")
        if not ok or not name.strip():
            return
        tab.bookmarks[name.strip()] = tab.editor.textCursor().position()
        self.refresh()
        self.win._refresh_navigation()

    def delete_bookmark(self):
        tab = self.win._current_tab()
        item = self.list_widget.currentItem()
        if tab is None or item is None:
            return
        tab.bookmarks.pop(item.text(), None)
        self.refresh()
        self.win._refresh_navigation()

    def go_to_bookmark(self):
        tab = self.win._current_tab()
        item = self.list_widget.currentItem()
        if tab is None or item is None:
            return
        position = item.data(Qt.ItemDataRole.UserRole)
        cursor = tab.editor.textCursor()
        cursor.setPosition(min(int(position), tab.editor.document().characterCount() - 1))
        tab.editor.setTextCursor(cursor)
        tab.editor.ensureCursorVisible()
        tab.editor.setFocus()

class PaperWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.base_page_w = 816
        self.base_page_h = 1056
        self.base_margin = 96
        self.page_w = self.base_page_w
        self.page_h = self.base_page_h
        self.margin = self.base_margin
        self.gap = 0
        self.pages = 1
        self.page_color = QColor("#FFFFFF")
        self.watermark = ""
        self.watermark_size = 60
        self.watermark_color = QColor(180, 180, 180)
        self.watermark_opacity = 18
        self.header_text = ""
        self.footer_text = ""
        self.show_header = False
        self.show_footer = False
        self.show_page_numbers = True
        self.page_number_format = "Page X of Y"
        self.show_border = False
        self.dark = False
        self.on_resize = None
        self.setFixedWidth(self.page_w + 40)
        self.editor = QTextEdit(self)
        self.editor.setAcceptRichText(True)
        self.editor.setTabStopDistance(40)
        self.editor.setAutoFormatting(QTextEdit.AutoFormattingFlag.AutoAll)
        self.editor.setFrameShape(QFrame.Shape.NoFrame)
        self.editor.setStyleSheet("QTextEdit { background: transparent; border: none; }")
        self.editor.document().setDocumentMargin(self.margin)
        self.editor.document().setPageSize(QSizeF(self.page_w, self.page_h))
        self.editor.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.editor.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        default_font = QFont("Calibri", 11)
        self.editor.document().setDefaultFont(default_font)
        self.editor.setFont(default_font)
        self.editor.document().documentLayout().documentSizeChanged.connect(self.update_size)
        self.editor.setGeometry(20, 0, self.page_w, self.page_h)
        QTimer.singleShot(10, self.update_size)

    def set_page(self, width, height, margin):
        self.base_page_w = max(200, int(width))
        self.base_page_h = max(200, int(height))
        self.base_margin = max(10, int(margin))
        self.page_w = self.base_page_w
        self.page_h = self.base_page_h
        self.margin = self.base_margin
        self.editor.document().setDocumentMargin(self.margin)
        self.editor.document().setPageSize(QSizeF(self.page_w, self.page_h))
        self.update_size()

    def update_size(self, *args):
        if not hasattr(self, "editor"):
            return
        document = self.editor.document()
        document.setPageSize(QSizeF(self.page_w, self.page_h))
        pages = max(1, int(document.pageCount()))
        self.pages = pages
        total_height = pages * self.page_h + max(0, pages - 1) * self.gap
        self.setFixedWidth(self.page_w + 40)
        self.setFixedHeight(int(total_height))
        self.editor.setGeometry(20, 0, self.page_w, int(total_height))
        region = QRegion()
        y = 0
        for _ in range(pages):
            region += QRegion(0, int(y), self.page_w, self.page_h)
            y += self.page_h + self.gap
        if not region.isEmpty():
            self.editor.viewport().setMask(region)
        if callable(self.on_resize):
            self.on_resize()
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        canvas = QColor("#202124") if self.dark else QColor("#E6E6E6")
        painter.fillRect(0, 0, self.width(), self.height(), canvas)
        page_number_font = QFont("Segoe UI", 9)
        header_font = QFont("Segoe UI", 9)
        y = 0
        for index in range(self.pages):
            shadow = QColor(0, 0, 0, 70 if self.dark else 35)
            painter.fillRect(20 + 4, y + 6, self.page_w, self.page_h, shadow)
            painter.fillRect(20, y, self.page_w, self.page_h, self.page_color)
            if self.show_border:
                pen = QPen(QColor("#9AA7B4"))
                pen.setWidth(2)
                painter.setPen(pen)
                painter.setBrush(Qt.BrushStyle.NoBrush)
                painter.drawRect(QRectF(20 + 12, y + 12, self.page_w - 24, self.page_h - 24))
            if self.watermark:
                painter.save()
                painter.setOpacity(self.watermark_opacity / 100.0)
                painter.translate(20 + self.page_w / 2, y + self.page_h / 2)
                painter.rotate(-35)
                watermark_font = QFont("Segoe UI", self.watermark_size, QFont.Weight.Bold)
                painter.setFont(watermark_font)
                painter.setPen(self.watermark_color)
                metrics = QFontMetrics(watermark_font)
                text_width = metrics.horizontalAdvance(self.watermark)
                painter.drawText(
                    QRectF(-text_width / 2, -metrics.height() / 2, text_width, metrics.height()),
                    Qt.AlignmentFlag.AlignCenter,
                    self.watermark,
                )
                painter.restore()
            if self.show_header and self.header_text:
                painter.setFont(header_font)
                painter.setPen(QColor("#666666"))
                painter.drawText(
                    QRectF(20 + self.margin, y + self.margin * 0.30, self.page_w - self.margin * 2, self.margin * 0.45),
                    Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter,
                    self.header_text,
                )
                painter.setPen(QColor("#D0D0D0"))
                painter.drawLine(
                    int(20 + self.margin), int(y + self.margin * 0.82),
                    int(20 + self.page_w - self.margin), int(y + self.margin * 0.82),
                )
            if self.show_footer and self.footer_text:
                painter.setFont(header_font)
                painter.setPen(QColor("#666666"))
                painter.drawText(
                    QRectF(20 + self.margin, y + self.page_h - self.margin * 0.75, self.page_w - self.margin * 2, self.margin * 0.45),
                    Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter,
                    self.footer_text,
                )
            if self.show_page_numbers:
                text = self.page_number_format.replace("X", str(index + 1)).replace("Y", str(self.pages))
                painter.setFont(page_number_font)
                painter.setPen(QColor("#999999"))
                painter.drawText(
                    QRectF(20, y + self.page_h - self.margin * 0.62, self.page_w, self.margin * 0.4),
                    Qt.AlignmentFlag.AlignCenter,
                    text,
                )
            y += self.page_h + self.gap
        if self.gap == 0:
            painter.setPen(QPen(QColor("#3A3A3F") if self.dark else QColor("#C8C8C8")))
            for index in range(1, self.pages):
                boundary = index * self.page_h
                painter.drawLine(20, int(boundary), 20 + self.page_w, int(boundary))


class HorizontalRuler(QWidget):
    def __init__(self):
        super().__init__()
        self.setFixedHeight(22)
        self.page_width = 816
        self.margin = 96
        self.zoom = 1.0
        self.dark = False

    def set_page(self, page_width, margin, zoom=1.0):
        self.page_width = int(page_width)
        self.margin = int(margin)
        self.zoom = zoom
        self.update()

    def set_dark(self, dark):
        self.dark = dark
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        background = QColor("#2A2A2E") if self.dark else QColor("#F3F3F3")
        band = QColor("#3A3A3F") if self.dark else QColor("#FFFFFF")
        text_color = QColor("#BBBBBB") if self.dark else QColor("#888888")
        line_color = QColor("#888888") if self.dark else QColor("#999999")
        painter.fillRect(0, 0, self.width(), self.height(), background)
        start_x = max(0, (self.width() - self.page_width) // 2)
        painter.fillRect(start_x + self.margin, 0, max(0, self.page_width - self.margin * 2), self.height(), band)
        painter.setPen(line_color)
        painter.setFont(QFont("Segoe UI", 7))
        unit = max(8.0, 96.0 * self.zoom)
        step = max(4.0, unit / 8.0)
        count = int(max(0, self.page_width - self.margin * 2) / step)
        for index in range(count + 1):
            x = start_x + self.margin + index * step
            if index % 8 == 0:
                painter.drawLine(int(x), 11, int(x), 22)
                painter.setPen(text_color)
                painter.drawText(int(x) + 2, 10, str(index // 8))
                painter.setPen(line_color)
            elif index % 4 == 0:
                painter.drawLine(int(x), 15, int(x), 22)
            else:
                painter.drawLine(int(x), 18, int(x), 22)


class DocumentTab(QWidget):
    TOP_PAD = 24

    def __init__(self, window, name):
        super().__init__()
        self.window_ref = window
        self.file_path = None
        self.untitled_name = name
        self.bookmarks = {}
        self.comments = {}
        self.next_comment_id = 1
        self.citations = []
        self.footnote_count = 0
        self.figure_count = 0
        self.zoom_steps = 0
        self.recovery_path = None
        self.last_text = ""
        self.revision_log = []
        self._pending_old = None
        self.toc_position = None
        self.paper = PaperWidget()
        self.editor = self.paper.editor
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        self.ruler = HorizontalRuler()
        layout.addWidget(self.ruler)
        self.scene = QGraphicsScene(self)
        self.proxy = self.scene.addWidget(self.paper)
        self.proxy.setPos(0, self.TOP_PAD)
        self.scene.setSceneRect(0, 0, self.paper.width(), self.paper.height() + self.TOP_PAD * 2)
        self.view = QGraphicsView(self.scene)
        self.view.setFrameShape(QFrame.Shape.NoFrame)
        self.view.setAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignTop)
        self.view.setRenderHints(
            QPainter.RenderHint.Antialiasing
            | QPainter.RenderHint.TextAntialiasing
            | QPainter.RenderHint.SmoothPixmapTransform
        )
        self.view.setTransformationAnchor(QGraphicsView.ViewportAnchor.AnchorViewCenter)
        self.view.setResizeAnchor(QGraphicsView.ViewportAnchor.AnchorViewCenter)
        self.view.setDragMode(QGraphicsView.DragMode.NoDrag)
        layout.addWidget(self.view, 1)
        self.paper.on_resize = self._sync_scene
        self.editor.viewport().installEventFilter(self)
        self.editor.installEventFilter(self)
        self.ruler.set_page(self.paper.page_w, self.paper.margin, 1.0)
        self.set_dark(window.dark)

    def _sync_scene(self):
        self.scene.setSceneRect(0, 0, self.paper.width(), self.paper.height() + self.TOP_PAD * 2)

    def eventFilter(self, watched, event):
        try:
            if watched is self.editor.viewport() and event.type() == QEvent.Type.Wheel:
                bar = self.view.verticalScrollBar()
                delta = event.angleDelta().y()
                bar.setValue(bar.value() - delta)
                return True
            if watched is self.editor and event.type() == QEvent.Type.KeyPress:
                if event.key() in (Qt.Key.Key_Backspace, Qt.Key.Key_Delete):
                    self._clear_page_break_for_edit(event.key())
                    return False
                if event.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
                    cursor = self.editor.textCursor()
                    block = cursor.block()
                    if (
                        not cursor.hasSelection()
                        and block.isValid()
                        and block.text().strip()
                        and bool(
                            block.blockFormat().pageBreakPolicy()
                            & QTextFormat.PageBreakFlag.PageBreak_AlwaysBefore
                        )
                    ):
                        QTimer.singleShot(0, self._clear_inherited_page_break)
                    return False
        except RuntimeError:
            return False
        return super().eventFilter(watched, event)

    def _clear_inherited_page_break(self):
        try:
            cursor = self.editor.textCursor()
            block = cursor.block()
            previous = block.previous()
            if not previous.isValid() or not previous.text().strip():
                return
            if not bool(
                previous.blockFormat().pageBreakPolicy()
                & QTextFormat.PageBreakFlag.PageBreak_AlwaysBefore
            ):
                return
            block_format = block.blockFormat()
            if bool(
                block_format.pageBreakPolicy()
                & QTextFormat.PageBreakFlag.PageBreak_AlwaysBefore
            ):
                block_format.setPageBreakPolicy(QTextFormat.PageBreakFlag.PageBreak_Auto)
                block_cursor = QTextCursor(block)
                block_cursor.setBlockFormat(block_format)
        except RuntimeError:
            return

    def _clear_page_break_for_edit(self, key):
        cursor = self.editor.textCursor()
        if cursor.hasSelection():
            return
        block = cursor.block()
        if key == Qt.Key.Key_Backspace:
            if cursor.position() != block.position():
                return
            if not bool(block.blockFormat().pageBreakPolicy() & QTextFormat.PageBreakFlag.PageBreak_AlwaysBefore):
                return
            if not block.text().strip():
                following = block.next()
                while following.isValid() and not following.text().strip():
                    following = following.next()
                if following.isValid():
                    following_cursor = QTextCursor(following)
                    following_format = following_cursor.blockFormat()
                    if bool(following_format.pageBreakPolicy() & QTextFormat.PageBreakFlag.PageBreak_AlwaysBefore):
                        following_format.setPageBreakPolicy(QTextFormat.PageBreakFlag.PageBreak_Auto)
                        following_cursor.setBlockFormat(following_format)
            block_cursor = QTextCursor(block)
            block_format = block_cursor.blockFormat()
            block_format.setPageBreakPolicy(QTextFormat.PageBreakFlag.PageBreak_Auto)
            block_cursor.setBlockFormat(block_format)
        elif key == Qt.Key.Key_Delete:
            if cursor.position() != block.position() + block.length() - 1:
                return
            following = block.next()
            if not following.isValid():
                return
            following_cursor = QTextCursor(following)
            following_format = following_cursor.blockFormat()
            if not bool(following_format.pageBreakPolicy() & QTextFormat.PageBreakFlag.PageBreak_AlwaysBefore):
                return
            following_format.setPageBreakPolicy(QTextFormat.PageBreakFlag.PageBreak_Auto)
            following_cursor.setBlockFormat(following_format)

    def display_name(self):
        if self.file_path:
            return os.path.basename(self.file_path)
        return self.untitled_name

    def set_default_font(self, family, size):
        font = QFont(family, size)
        self.editor.document().setDefaultFont(font)
        self.editor.setFont(font)

    def apply_zoom(self, percent):
        percent = max(50, min(200, int(percent)))
        self.zoom_steps = percent
        scale = percent / 100.0
        self.view.resetTransform()
        self.view.scale(scale, scale)
        self.ruler.set_page(
            int(self.paper.page_w * scale),
            int(self.paper.margin * scale),
            scale,
        )

    def refresh_ruler(self):
        scale = max(0.5, min(2.0, self.zoom_steps / 100.0 if self.zoom_steps else 1.0))
        self.ruler.set_page(int(self.paper.page_w * scale), int(self.paper.margin * scale), scale)

    def set_dark(self, dark):
        self.paper.dark = dark
        self.paper.update()
        self.ruler.set_dark(dark)
        background = "#202124" if dark else "#E6E6E6"
        self.view.setBackgroundBrush(QColor(background))
        self.view.viewport().setStyleSheet(f"background: {background};")


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
PDF_CACHE_DIR = os.path.join(
    os.environ.get("LOCALAPPDATA") or tempfile.gettempdir(), "SimzWord", "pdf_cache"
)


def _docx_qn(tag):
    return f"{{{W_NS}}}{tag}"


def _docx_local(tag):
    return tag.rsplit("}", 1)[-1]


def _xml_text(value):
    return xml_escape(str(value))


def _xml_attr(value):
    return quoteattr(str(value))


def _px_to_twips(value):
    return max(0, int(round(float(value) * 15.0)))


def _px_to_twips_signed(value):
    return int(round(float(value) * 15.0))


def _docx_on(element):
    if element is None:
        return False
    return str(element.get(_docx_qn("val"), "1")).lower() not in ("0", "false", "off")


def _docx_val(element, default=None):
    if element is None:
        return default
    return element.get(_docx_qn("val"), default)


def _docx_run_properties(fmt):
    parts = []
    family = fmt.font().family()
    if family:
        parts.append(
            f"<w:rFonts w:ascii={_xml_attr(family)} w:hAnsi={_xml_attr(family)}/>"
        )
    if int(fmt.fontWeight()) >= int(QFont.Weight.Bold):
        parts.append("<w:b/>")
    if fmt.fontItalic():
        parts.append("<w:i/>")
    if fmt.fontUnderline():
        parts.append('<w:u w:val="single"/>')
    if fmt.fontStrikeOut():
        parts.append("<w:strike/>")
    size = float(fmt.fontPointSize())
    if size > 0:
        parts.append(f'<w:sz w:val="{max(2, int(round(size * 2)))}"/>')
    color = fmt.foreground().color()
    if color.isValid() and color.alpha() > 0:
        parts.append(f'<w:color w:val="{color.name().lstrip("#").upper()}"/>')
    background_brush = fmt.background()
    if background_brush.style() != Qt.BrushStyle.NoBrush:
        background = background_brush.color()
        if background.isValid() and background.name().upper() != "#FFFFFF":
            parts.append(
                f'<w:shd w:val="clear" w:color="auto" w:fill="{background.name().lstrip("#").upper()}"/>'
            )
    vertical = fmt.verticalAlignment()
    if vertical == QTextCharFormat.VerticalAlignment.AlignSuperScript:
        parts.append('<w:vertAlign w:val="superscript"/>')
    elif vertical == QTextCharFormat.VerticalAlignment.AlignSubScript:
        parts.append('<w:vertAlign w:val="subscript"/>')
    return f'<w:rPr>{"".join(parts)}</w:rPr>' if parts else ""


def _docx_run_xml(text, fmt):
    if not text:
        return ""
    properties = _docx_run_properties(fmt)
    pieces = []
    for part in re.split(r"(\t|\n)", text):
        if part == "\t":
            pieces.append("<w:tab/>")
        elif part == "\n":
            pieces.append("<w:br/>")
        elif part:
            pieces.append(f'<w:t xml:space="preserve">{_xml_text(part)}</w:t>')
    return f'<w:r>{properties}{"".join(pieces)}</w:r>'


def _docx_paragraph_properties(block, include_page_break=True):
    fmt = block.blockFormat()
    parts = []
    if (
        include_page_break
        and block.position() > 0
        and bool(fmt.pageBreakPolicy() & QTextFormat.PageBreakFlag.PageBreak_AlwaysBefore)
    ):
        parts.append("<w:pageBreakBefore/>")
    level = int(fmt.headingLevel())
    if level > 0:
        parts.append(f'<w:pStyle w:val="Heading{min(6, level)}"/>')
    alignment = fmt.alignment()
    if alignment & Qt.AlignmentFlag.AlignHCenter:
        parts.append('<w:jc w:val="center"/>')
    elif alignment & Qt.AlignmentFlag.AlignRight:
        parts.append('<w:jc w:val="right"/>')
    elif alignment & Qt.AlignmentFlag.AlignJustify:
        parts.append('<w:jc w:val="both"/>')
    left = _px_to_twips(fmt.leftMargin())
    right = _px_to_twips(fmt.rightMargin())
    first_line = _px_to_twips_signed(fmt.textIndent())
    if left or right or first_line:
        parts.append(
            f'<w:ind w:left="{left}" w:right="{right}" w:firstLine="{first_line}"/>'
        )
    before = _px_to_twips(fmt.topMargin())
    after = _px_to_twips(fmt.bottomMargin())
    if before or after:
        parts.append(f'<w:spacing w:before="{before}" w:after="{after}"/>')
    list_obj = block.textList()
    if list_obj is not None:
        list_format = list_obj.format()
        number_id = 2 if list_format.style() == QTextListFormat.Style.ListDecimal else 1
        list_level = max(0, min(8, int(list_format.indent())))
        parts.append(
            f'<w:numPr><w:ilvl w:val="{list_level}"/><w:numId w:val="{number_id}"/></w:numPr>'
        )
    return f'<w:pPr>{"".join(parts)}</w:pPr>' if parts else ""


def _docx_paragraph_xml(block, include_page_break=True):
    properties = _docx_paragraph_properties(block, include_page_break)
    runs = []
    iterator = block.begin()
    while not iterator.atEnd():
        fragment = iterator.fragment()
        if fragment.isValid():
            fmt = fragment.charFormat()
            text = "[Image]" if fmt.isImageFormat() else fragment.text()
            runs.append(_docx_run_xml(text, fmt))
        iterator += 1
    return f'<w:p>{properties}{"".join(runs)}</w:p>'


def _docx_table_xml(table, document, content_width):
    rows = table.rows()
    columns = max(1, table.columns())
    column_width = _px_to_twips(max(100.0, float(content_width) / columns))
    border = "".join(
        f'<w:{side} w:val="single" w:sz="4" w:space="0" w:color="B0B0B0"/>'
        for side in ("top", "left", "bottom", "right", "insideH", "insideV")
    )
    parts = [
        "<w:tbl><w:tblPr>",
        '<w:tblW w:w="0" w:type="auto"/>',
        f"<w:tblBorders>{border}</w:tblBorders>",
        "</w:tblPr><w:tblGrid>",
    ]
    parts.extend(f'<w:gridCol w:w="{column_width}"/>' for _ in range(columns))
    parts.append("</w:tblGrid>")
    for row in range(rows):
        parts.append("<w:tr>")
        for column in range(columns):
            cell = table.cellAt(row, column)
            parts.append(
                f'<w:tc><w:tcPr><w:tcW w:w="{column_width}" w:type="dxa"/></w:tcPr>'
            )
            paragraphs = []
            if cell.isValid():
                block = document.findBlock(cell.firstPosition())
                while block.isValid() and block.position() < cell.lastPosition():
                    paragraphs.append(_docx_paragraph_xml(block, include_page_break=False))
                    block = block.next()
            parts.append("".join(paragraphs) if paragraphs else "<w:p/>")
            parts.append("</w:tc>")
        parts.append("</w:tr>")
    parts.append("</w:tbl>")
    return "".join(parts)


def _docx_section_xml(paper):
    width = _px_to_twips(paper.page_w)
    height = _px_to_twips(paper.page_h)
    margin = _px_to_twips(paper.margin)
    orientation = ' w:orient="landscape"' if paper.page_w > paper.page_h else ""
    return (
        "<w:sectPr>"
        f'<w:pgSz w:w="{width}" w:h="{height}"{orientation}/>'
        f'<w:pgMar w:top="{margin}" w:right="{margin}" w:bottom="{margin}" '
        f'w:left="{margin}" w:header="720" w:footer="720" w:gutter="0"/>'
        "</w:sectPr>"
    )


def _docx_styles_xml():
    heading_sizes = (34, 28, 24, 22, 22, 22)
    styles = [
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
        '<w:name w:val="Normal"/><w:qFormat/></w:style>'
    ]
    for index, size in enumerate(heading_sizes, start=1):
        styles.append(
            f'<w:style w:type="paragraph" w:styleId="Heading{index}">'
            f'<w:name w:val="heading {index}"/><w:basedOn w:val="Normal"/>'
            '<w:next w:val="Normal"/><w:qFormat/>'
            f'<w:pPr><w:keepNext/><w:outlineLvl w:val="{index - 1}"/></w:pPr>'
            f'<w:rPr><w:b/><w:sz w:val="{size}"/></w:rPr></w:style>'
        )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:styles xmlns:w="{W_NS}"><w:docDefaults><w:rPrDefault><w:rPr>'
        '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>'
        '<w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>'
        + "".join(styles)
        + "</w:styles>"
    )


def _docx_numbering_xml():
    parts = []
    for abstract_id, number_format, text in (
        (0, "bullet", "•"),
        (1, "decimal", "%1."),
    ):
        levels = []
        for level in range(9):
            indent = 720 * (level + 1)
            levels.append(
                f'<w:lvl w:ilvl="{level}"><w:start w:val="1"/>'
                f'<w:numFmt w:val="{number_format}"/><w:lvlText w:val="{text}"/>'
                '<w:lvlJc w:val="left"/><w:pPr>'
                f'<w:ind w:left="{indent}" w:hanging="360"/></w:pPr>'
                '<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/></w:rPr></w:lvl>'
            )
        parts.append(
            f'<w:abstractNum w:abstractNumId="{abstract_id}">'
            '<w:multiLevelType w:val="hybridMultilevel"/>'
            + "".join(levels)
            + "</w:abstractNum>"
        )
    parts.append('<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>')
    parts.append('<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>')
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:numbering xmlns:w="{W_NS}">'
        + "".join(parts)
        + "</w:numbering>"
    )


def save_docx(tab, path):
    document = tab.editor.document()
    body = []
    block = document.begin()
    while block.isValid():
        cursor = QTextCursor(block)
        table = cursor.currentTable()
        if table is not None and block.position() == table.firstPosition():
            body.append(
                _docx_table_xml(
                    table, document, tab.paper.page_w - tab.paper.margin * 2
                )
            )
            last_position = table.lastPosition()
            block = document.findBlock(
                min(last_position + 1, document.characterCount() - 1)
            )
            while block.isValid() and block.position() <= last_position:
                block = block.next()
            continue
        body.append(_docx_paragraph_xml(block))
        block = block.next()
    body.append(_docx_section_xml(tab.paper))
    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{W_NS}" xmlns:r="{R_NS}"><w:body>'
        + "".join(body)
        + "</w:body></w:document>"
    )
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Types xmlns="{CT_NS}">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
        '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
        '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
        '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
        "</Types>"
    )
    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{PKG_REL_NS}">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
        "</Relationships>"
    )
    document_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{PKG_REL_NS}">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>'
        "</Relationships>"
    )
    now = datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")
    core = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        f"<dc:title>{_xml_text(os.path.basename(path))}</dc:title>"
        "<dc:creator>SimzWord</dc:creator><cp:lastModifiedBy>SimzWord</cp:lastModifiedBy>"
        f'<dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>'
        f'<dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>'
        "</cp:coreProperties>"
    )
    app = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
        'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
        "<Application>SimzWord</Application></Properties>"
    )
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types)
        archive.writestr("_rels/.rels", root_rels)
        archive.writestr("word/document.xml", document_xml)
        archive.writestr("word/styles.xml", _docx_styles_xml())
        archive.writestr("word/numbering.xml", _docx_numbering_xml())
        archive.writestr("word/_rels/document.xml.rels", document_rels)
        archive.writestr("docProps/core.xml", core)
        archive.writestr("docProps/app.xml", app)


def _docx_number_formats(archive):
    if "word/numbering.xml" not in archive.namelist():
        return {}
    root = ET.fromstring(archive.read("word/numbering.xml"))
    abstract_formats = {}
    for abstract in root.findall(_docx_qn("abstractNum")):
        abstract_id = abstract.get(_docx_qn("abstractNumId"))
        level = abstract.find(_docx_qn("lvl"))
        if level is not None:
            number_format = level.find(_docx_qn("numFmt"))
            abstract_formats[abstract_id] = _docx_val(number_format, "bullet")
    formats = {}
    for number in root.findall(_docx_qn("num")):
        number_id = number.get(_docx_qn("numId"))
        abstract = number.find(_docx_qn("abstractNumId"))
        formats[number_id] = abstract_formats.get(_docx_val(abstract), "bullet")
    return formats


def _docx_apply_paragraph_format(cursor, paragraph, number_formats):
    properties = paragraph.find(_docx_qn("pPr"))
    fmt = cursor.blockFormat()
    fmt.setHeadingLevel(0)
    fmt.setPageBreakPolicy(QTextFormat.PageBreakFlag.PageBreak_Auto)
    fmt.setAlignment(Qt.AlignmentFlag.AlignLeft)
    fmt.setLeftMargin(0)
    fmt.setRightMargin(0)
    fmt.setTextIndent(0)
    fmt.setTopMargin(0)
    fmt.setBottomMargin(0)
    if properties is not None:
        style = _docx_val(properties.find(_docx_qn("pStyle")))
        if style:
            match = re.search(r"heading\s*(\d+)", style, re.IGNORECASE)
            if match:
                fmt.setHeadingLevel(max(1, min(6, int(match.group(1)))))
        outline = properties.find(_docx_qn("outlineLvl"))
        if outline is not None:
            try:
                fmt.setHeadingLevel(max(1, min(6, int(_docx_val(outline, "0")) + 1)))
            except (TypeError, ValueError):
                pass
        alignment = _docx_val(properties.find(_docx_qn("jc")))
        if alignment == "center":
            fmt.setAlignment(Qt.AlignmentFlag.AlignHCenter)
        elif alignment == "right":
            fmt.setAlignment(Qt.AlignmentFlag.AlignRight)
        elif alignment in ("both", "distribute"):
            fmt.setAlignment(Qt.AlignmentFlag.AlignJustify)
        indentation = properties.find(_docx_qn("ind"))
        if indentation is not None:
            try:
                left = int(_docx_val(indentation, "0") or 0) / 15.0
            except (TypeError, ValueError):
                left = 0.0
            try:
                right = int(indentation.get(_docx_qn("right"), "0") or 0) / 15.0
            except (TypeError, ValueError):
                right = 0.0
            try:
                first_line = int(indentation.get(_docx_qn("firstLine"), "0") or 0) / 15.0
            except (TypeError, ValueError):
                first_line = 0.0
            try:
                hanging = int(indentation.get(_docx_qn("hanging"), "0") or 0) / 15.0
            except (TypeError, ValueError):
                hanging = 0.0
            if left:
                fmt.setLeftMargin(left)
            if right:
                fmt.setRightMargin(right)
            if first_line or hanging:
                fmt.setTextIndent(first_line - hanging)
        spacing = properties.find(_docx_qn("spacing"))
        if spacing is not None:
            for attribute, setter in (
                ("before", fmt.setTopMargin),
                ("after", fmt.setBottomMargin),
            ):
                try:
                    value = int(spacing.get(_docx_qn(attribute), "0") or 0) / 15.0
                except (TypeError, ValueError):
                    value = 0.0
                if value:
                    setter(value)
        if _docx_on(properties.find(_docx_qn("pageBreakBefore"))):
            fmt.setPageBreakPolicy(QTextFormat.PageBreakFlag.PageBreak_AlwaysBefore)
        cursor.setBlockFormat(fmt)
        number_properties = properties.find(_docx_qn("numPr"))
        if number_properties is not None:
            number_id = _docx_val(number_properties.find(_docx_qn("numId")))
            list_format = QTextListFormat()
            if number_formats.get(number_id) == "decimal":
                list_format.setStyle(QTextListFormat.Style.ListDecimal)
            else:
                list_format.setStyle(QTextListFormat.Style.ListDisc)
            try:
                level = int(_docx_val(number_properties.find(_docx_qn("ilvl")), "0") or 0)
            except (TypeError, ValueError):
                level = 0
            list_format.setIndent(max(1, min(9, level + 1)))
            cursor.createList(list_format)
    else:
        cursor.setBlockFormat(fmt)


def _docx_insert_run(run, cursor):
    properties = run.find(_docx_qn("rPr"))
    fmt = QTextCharFormat()
    if properties is not None:
        if _docx_on(properties.find(_docx_qn("b"))):
            fmt.setFontWeight(QFont.Weight.Bold)
        if _docx_on(properties.find(_docx_qn("i"))):
            fmt.setFontItalic(True)
        underline = properties.find(_docx_qn("u"))
        if underline is not None and _docx_val(underline, "single") not in ("none", "0"):
            fmt.setFontUnderline(True)
        if _docx_on(properties.find(_docx_qn("strike"))):
            fmt.setFontStrikeOut(True)
        fonts = properties.find(_docx_qn("rFonts"))
        if fonts is not None:
            family = fonts.get(_docx_qn("ascii")) or fonts.get(_docx_qn("hAnsi"))
            if family:
                fmt.setFontFamily(family)
        size = properties.find(_docx_qn("sz"))
        if size is not None:
            try:
                fmt.setFontPointSize(float(_docx_val(size, "0")) / 2.0)
            except (TypeError, ValueError):
                pass
        color = properties.find(_docx_qn("color"))
        if color is not None:
            value = _docx_val(color, "auto")
            if value and value != "auto":
                fmt.setForeground(QColor("#" + value))
        shading = properties.find(_docx_qn("shd"))
        if shading is not None:
            fill = shading.get(_docx_qn("fill"))
            if fill and fill not in ("auto", "FFFFFF"):
                fmt.setBackground(QColor("#" + fill))
        vertical = _docx_val(properties.find(_docx_qn("vertAlign")))
        if vertical == "superscript":
            fmt.setVerticalAlignment(QTextCharFormat.VerticalAlignment.AlignSuperScript)
        elif vertical == "subscript":
            fmt.setVerticalAlignment(QTextCharFormat.VerticalAlignment.AlignSubScript)
    for child in list(run):
        tag = _docx_local(child.tag)
        if tag == "t":
            cursor.insertText(child.text or "", fmt)
        elif tag == "tab":
            cursor.insertText("\t", fmt)
        elif tag in ("br", "cr"):
            if _docx_val(child, "") == "page":
                cursor.insertBlock()
                block_format = cursor.blockFormat()
                block_format.setPageBreakPolicy(
                    QTextFormat.PageBreakFlag.PageBreak_AlwaysBefore
                )
                cursor.setBlockFormat(block_format)
            else:
                cursor.insertText("\n", fmt)
        elif tag == "noBreakHyphen":
            cursor.insertText("\u2011", fmt)


def _docx_insert_runs(parent, cursor):
    for child in list(parent):
        tag = _docx_local(child.tag)
        if tag == "r":
            _docx_insert_run(child, cursor)
        elif tag in ("hyperlink", "smartTag", "ins", "sdt", "sdtContent"):
            _docx_insert_runs(child, cursor)


def _docx_insert_table(cursor, table_element):
    rows = table_element.findall(_docx_qn("tr"))
    if not rows:
        return
    column_count = max(
        1,
        max(len(row.findall(_docx_qn("tc"))) for row in rows),
    )
    table_format = QTextTableFormat()
    table_format.setBorder(1)
    table_format.setBorderBrush(QBrush(QColor("#B0B0B0")))
    table_format.setBorderStyle(QTextFrameFormat.BorderStyle.BorderStyle_Solid)
    table_format.setCellPadding(4)
    table = cursor.insertTable(len(rows), column_count, table_format)
    for row_index, row in enumerate(rows):
        for column_index, cell_element in enumerate(row.findall(_docx_qn("tc"))):
            if column_index >= column_count:
                break
            cell = table.cellAt(row_index, column_index)
            cell_cursor = cell.firstCursorPosition()
            first_paragraph = True
            for paragraph in cell_element.findall(_docx_qn("p")):
                if not first_paragraph:
                    cell_cursor.insertBlock()
                _docx_apply_paragraph_format(cell_cursor, paragraph, {})
                _docx_insert_runs(paragraph, cell_cursor)
                first_paragraph = False


def load_docx(tab, path):
    with zipfile.ZipFile(path) as archive:
        root = ET.fromstring(archive.read("word/document.xml"))
        number_formats = _docx_number_formats(archive)
    body = root.find(_docx_qn("body"))
    tab.editor.clear()
    cursor = tab.editor.textCursor()
    first = True
    page_info = None
    if body is not None:
        for child in list(body):
            tag = _docx_local(child.tag)
            if tag == "p":
                if not first:
                    cursor.insertBlock()
                _docx_apply_paragraph_format(cursor, child, number_formats)
                _docx_insert_runs(child, cursor)
                first = False
            elif tag == "tbl":
                if not first:
                    cursor.insertBlock()
                _docx_insert_table(cursor, child)
                cursor = tab.editor.textCursor()
                cursor.movePosition(QTextCursor.MoveOperation.End)
                tab.editor.setTextCursor(cursor)
                first = False
            elif tag == "sectPr":
                page_size = child.find(_docx_qn("pgSz"))
                margins = child.find(_docx_qn("pgMar"))
                if page_size is not None:
                    try:
                        width = int(page_size.get(_docx_qn("w"), "0")) / 15.0
                        height = int(page_size.get(_docx_qn("h"), "0")) / 15.0
                    except (TypeError, ValueError):
                        width = height = 0.0
                    if width > 0 and height > 0:
                        margin = tab.paper.base_margin
                        if margins is not None:
                            try:
                                margin = int(margins.get(_docx_qn("left"), "0")) / 15.0
                            except (TypeError, ValueError):
                                margin = tab.paper.base_margin
                        page_info = (width, height, margin)
    tab.editor.setTextCursor(cursor)
    if page_info is not None:
        tab.paper.set_page(*page_info)
    else:
        tab.paper.update_size()
    tab.editor.document().setModified(False)


def extract_pdf_pages(path):
    try:
        from PySide6.QtPdf import QPdfDocument
    except Exception as error:
        raise RuntimeError("Qt PDF support is not available.") from error
    document = QPdfDocument()
    document.load(path)
    if document.status() != QPdfDocument.Status.Ready:
        status = document.status()
        error = document.error()
        document.close()
        raise RuntimeError(f"Could not read this PDF ({status}, {error}).")
    pages = []
    for index in range(document.pageCount()):
        selection = document.getAllText(index)
        text = selection.text() if selection.isValid() else ""
        text = text.replace("\r\n", "\n").replace("\r", "\n")
        pages.append(text)
    page_size = document.pagePointSize(0) if document.pageCount() else None
    document.close()
    return pages, page_size


def _pdf_cache_path(path):
    digest = hashlib.sha1()
    with open(path, "rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return os.path.join(PDF_CACHE_DIR, digest.hexdigest() + ".simz")


class SimzWord(QMainWindow):
    def __init__(self):
        super().__init__()
        self.settings = QSettings(ORG_NAME, APP_NAME)
        self.dark = self.settings.value("ui/dark", False, bool)
        self._zoom = 100
        self._doc_counter = 0
        self._track_changes = False
        self._format_painter_fmt = None
        self._focus_mode = False
        self._focus_hidden_docks = []
        self.default_font_family = self.settings.value("editor/font_family", "Calibri", str)
        self.default_font_size = self.settings.value("editor/font_size", 11, int)
        self.custom_words = set(self.settings.value("spelling/custom", [], list) or [])
        self._recovery_dir = os.path.join(tempfile.gettempdir(), "SimzWord_recovery")
        try:
            os.makedirs(self._recovery_dir, exist_ok=True)
        except Exception:
            pass
        self._nav_timer = QTimer(self)
        self._nav_timer.setSingleShot(True)
        self._nav_timer.timeout.connect(self._refresh_navigation)
        self._diff_timer = QTimer(self)
        self._diff_timer.setSingleShot(True)
        self._diff_timer.timeout.connect(self._flush_revisions)
        self.setWindowTitle(f"Untitled - {APP_NAME}")
        self.setWindowIcon(QIcon(make_logo_pixmap(64)))
        self.resize(1320, 880)
        self.setMinimumSize(960, 620)
        self._create_actions()
        self._create_ui()
        self._create_menus()
        self._create_docks()
        self._create_statusbar()
        self._apply_theme()
        self._apply_ui_settings()
        self._new_document(initial=True)
        self._restore_window_state()
        self._setup_autosave()
        QTimer.singleShot(500, self._check_recovery)

    @property
    def editor(self):
        tab = self._current_tab()
        return tab.editor if tab else None

    def _current_tab(self):
        widget = self.tab_widget.currentWidget() if hasattr(self, "tab_widget") else None
        return widget if isinstance(widget, DocumentTab) else None

    def _tab_for_editor(self, editor):
        for index in range(self.tab_widget.count()):
            tab = self.tab_widget.widget(index)
            if isinstance(tab, DocumentTab) and tab.editor is editor:
                return tab
        return None

    def _act(self, text, slot=None, shortcut=None, checkable=False, icon=None, tip=None):
        action = QAction(text, self)
        if shortcut is not None:
            try:
                action.setShortcut(QKeySequence(shortcut))
            except Exception:
                pass
        if checkable:
            action.setCheckable(True)
        if icon is not None:
            action.setIcon(icon)
        if tip:
            action.setToolTip(tip)
            action.setStatusTip(tip)
        if slot is not None:
            action.triggered.connect(slot)
        return action

    def _create_actions(self):
        style = self.style()
        icon_new = style.standardIcon(QStyle.StandardPixmap.SP_FileIcon)
        icon_open = style.standardIcon(QStyle.StandardPixmap.SP_DialogOpenButton)
        icon_save = style.standardIcon(QStyle.StandardPixmap.SP_DialogSaveButton)
        icon_print = style.standardIcon(QStyle.StandardPixmap.SP_FileDialogDetailedView)
        icon_find = style.standardIcon(QStyle.StandardPixmap.SP_FileDialogContentsView)
        icon_undo = style.standardIcon(QStyle.StandardPixmap.SP_ArrowBack)
        icon_redo = style.standardIcon(QStyle.StandardPixmap.SP_ArrowForward)
        icon_info = style.standardIcon(QStyle.StandardPixmap.SP_MessageBoxInformation)

        self.act_new = self._act("New", self._new_document, QKeySequence.StandardKey.New, icon=icon_new, tip="Create a new document (Ctrl+N)")
        self.act_open = self._act("Open...", self._open_file, QKeySequence.StandardKey.Open, icon=icon_open, tip="Open a document (Ctrl+O)")
        self.act_save = self._act("Save", self._save_file, QKeySequence.StandardKey.Save, icon=icon_save, tip="Save the document (Ctrl+S)")
        self.act_save_as = self._act("Save As...", self._save_file_as, "Ctrl+Shift+S", icon=icon_save, tip="Save under a new name")
        self.act_save_all = self._act("Save All", self._save_all, "Ctrl+Alt+S", icon=icon_save, tip="Save every open document")
        self.act_close_tab = self._act("Close Document", self._close_current_tab, QKeySequence.StandardKey.Close, tip="Close the current document tab")
        self.act_export_docx = self._act("Export as Word Document...", self._export_docx)
        self.act_export_pdf = self._act("Export as PDF...", self._export_pdf)
        self.act_export_html = self._act("Export as HTML...", self._export_html)
        self.act_export_md = self._act("Export as Markdown...", self._export_md)
        self.act_export_txt = self._act("Export as Plain Text...", self._export_txt)
        self.act_print = self._act("Print...", self._print, QKeySequence.StandardKey.Print, icon=icon_print, tip="Print the document (Ctrl+P)")
        self.act_print_preview = self._act("Print Preview...", self._print_preview)
        self.act_page_setup = self._act("Page Setup...", self._page_setup)
        self.act_properties = self._act("Document Properties...", self._show_properties)
        self.act_exit = self._act("Exit", self.close, "Ctrl+Q")

        self.act_undo = self._act("Undo", lambda: self.editor and self.editor.undo(), QKeySequence.StandardKey.Undo, icon=icon_undo)
        self.act_redo = self._act("Redo", lambda: self.editor and self.editor.redo(), QKeySequence.StandardKey.Redo, icon=icon_redo)
        self.act_cut = self._act("Cut", lambda: self.editor and self.editor.cut(), QKeySequence.StandardKey.Cut)
        self.act_copy = self._act("Copy", lambda: self.editor and self.editor.copy(), QKeySequence.StandardKey.Copy)
        self.act_paste = self._act("Paste", lambda: self.editor and self.editor.paste(), QKeySequence.StandardKey.Paste)
        self.act_paste_special = self._act("Paste as Plain Text", self._paste_plain, "Ctrl+Shift+V")
        self.act_select_all = self._act("Select All", self._select_all, QKeySequence.StandardKey.SelectAll)
        self.act_find = self._act("Find...", self._show_find, QKeySequence.StandardKey.Find, icon=icon_find)
        self.act_replace = self._act("Find and Replace...", self._show_find_replace, "Ctrl+H", icon=icon_find)
        self.act_find_next = self._act("Find Next", self._find_next, "F3")
        self.act_go_to = self._act("Go To...", self._go_to, "Ctrl+G")
        self.act_clear_contents = self._act("Clear Contents", self._clear_contents)

        self.act_bold = self._act("Bold", lambda checked=False: self._merge_fmt("weight", QFont.Weight.Bold if self.editor and self.editor.textCursor().charFormat().fontWeight() < QFont.Weight.Bold else QFont.Weight.Normal), "Ctrl+B", checkable=True, icon=make_text_icon("B", "bold"), tip="Bold (Ctrl+B)")
        self.act_italic = self._act("Italic", lambda checked=False: self._merge_fmt("italic", not (self.editor and self.editor.textCursor().charFormat().fontItalic())), "Ctrl+I", checkable=True, icon=make_text_icon("I", "italic"), tip="Italic (Ctrl+I)")
        self.act_underline = self._act("Underline", lambda checked=False: self._merge_fmt("underline", not (self.editor and self.editor.textCursor().charFormat().fontUnderline())), "Ctrl+U", checkable=True, icon=make_text_icon("U", "underline"), tip="Underline (Ctrl+U)")
        self.act_strike = self._act("Strikethrough", lambda checked=False: self._merge_fmt("strike", not (self.editor and self.editor.textCursor().charFormat().fontStrikeOut())), "Ctrl+Shift+X", checkable=True, icon=make_text_icon("S", "strike"))
        self.act_superscript = self._act("Superscript", lambda: self._merge_fmt("v-align", QTextCharFormat.VerticalAlignment.AlignSuperScript), checkable=True, icon=make_text_icon("xÂ²"))
        self.act_subscript = self._act("Subscript", lambda: self._merge_fmt("v-align", QTextCharFormat.VerticalAlignment.AlignSubScript), checkable=True, icon=make_text_icon("xâ‚‚"))
        self.act_grow_font = self._act("Grow Font", lambda: self._change_font_size(1), icon=make_text_icon("A+"))
        self.act_shrink_font = self._act("Shrink Font", lambda: self._change_font_size(-1), icon=make_text_icon("A-"))
        self.act_text_color = self._act("Text Color", self._choose_text_color, icon=make_color_icon(QColor("#000000"), "A"))
        self.act_highlight = self._act("Highlight", self._choose_highlight_color, icon=make_color_icon(QColor("#FFFF00"), "H"))
        self.act_clear_format = self._act("Clear Formatting", self._clear_formatting, icon=make_glyph_icon("âŒ«"))
        self.act_format_painter = self._act("Format Painter", self._toggle_format_painter, checkable=True, icon=make_glyph_icon("ðŸ–Œ"))
        self.act_change_case = self._act("Change Case", self._cycle_case, icon=make_glyph_icon("Aa"))
        self.act_bullets = self._act("Bullets", lambda: self._apply_list(QTextListFormat.Style.ListDisc), icon=make_list_icon(False))
        self.act_numbering = self._act("Numbering", lambda: self._apply_list(QTextListFormat.Style.ListDecimal), icon=make_list_icon(True))
        self.act_indent_inc = self._act("Increase Indent", lambda: self._change_indent(1), icon=make_glyph_icon("â†’"))
        self.act_indent_dec = self._act("Decrease Indent", lambda: self._change_indent(-1), icon=make_glyph_icon("â†"))
        self.act_sort = self._act("Sort", lambda: self._sort_selection(True), icon=make_glyph_icon("â‡…"))
        self.act_shading = self._act("Shading", self._choose_shading, icon=make_glyph_icon("â–¤"))
        self.act_line_spacing = self._act("Line Spacing", lambda: self._apply_line_spacing(115), icon=make_glyph_icon("â‰¡"))
        self.act_font_dialog = self._act("Font...", self._font_dialog)
        self.act_paragraph_dialog = self._act("Paragraph...", self._paragraph_dialog)

        self.align_group = QActionGroup(self)
        self.align_group.setExclusive(True)
        self.act_align_left = self._act("Align Left", lambda: self._set_alignment(Qt.AlignmentFlag.AlignLeft), "Ctrl+L", checkable=True, icon=make_align_icon("left"))
        self.act_align_center = self._act("Center", lambda: self._set_alignment(Qt.AlignmentFlag.AlignHCenter), "Ctrl+E", checkable=True, icon=make_align_icon("center"))
        self.act_align_right = self._act("Align Right", lambda: self._set_alignment(Qt.AlignmentFlag.AlignRight), "Ctrl+R", checkable=True, icon=make_align_icon("right"))
        self.act_align_justify = self._act("Justify", lambda: self._set_alignment(Qt.AlignmentFlag.AlignJustify), "Ctrl+J", checkable=True, icon=make_align_icon("justify"))
        for action in (self.act_align_left, self.act_align_center, self.act_align_right, self.act_align_justify):
            self.align_group.addAction(action)
        self.act_align_left.setChecked(True)

        self.act_page_break = self._act("Page Break", self._insert_page_break, "Ctrl+Return")
        self.act_blank_page = self._act("Blank Page", self._insert_blank_page)
        self.act_cover_page = self._act("Cover Page", self._insert_cover_page)
        self.act_table = self._act("Table", self._insert_table, icon=make_glyph_icon("â–¦"))
        self.act_picture = self._act("Picture", self._insert_image, icon=make_glyph_icon("â–§"))
        self.act_shapes = self._act("Shapes", lambda: self._insert_shape("Rectangle"), icon=make_glyph_icon("â—¼"))
        self.act_chart = self._act("Chart", self._insert_chart, icon=make_glyph_icon("â–¥"))
        self.act_hyperlink = self._act("Hyperlink", self._insert_link, "Ctrl+K", icon=make_glyph_icon("â›“"))
        self.act_bookmark = self._act("Bookmark", self._insert_bookmark, icon=make_glyph_icon("ðŸ”–"))
        self.act_header_footer = self._act("Header and Footer", self._header_footer, icon=make_glyph_icon("â–¤"))
        self.act_page_number = self._act("Page Number", self._page_number_dialog, icon=make_glyph_icon("#"))
        self.act_date_time = self._act("Date and Time", self._insert_datetime, icon=make_glyph_icon("ðŸ•’"))
        self.act_symbol = self._act("Symbol...", self._insert_symbol, icon=make_glyph_icon("Î©"))
        self.act_horizontal_rule = self._act("Horizontal Line", self._insert_hr, icon=make_glyph_icon("â€”"))
        self.act_text_box = self._act("Text Box", self._insert_text_box, icon=make_glyph_icon("â–­"))
        self.act_comment = self._act("New Comment", self._add_comment, "Ctrl+Alt+M", icon=make_glyph_icon("ðŸ’¬"))
        self.act_footnote = self._act("Insert Footnote", self._insert_footnote, "Ctrl+Alt+F", icon=make_glyph_icon("Â¹"))
        self.act_citation = self._act("Insert Citation...", self._insert_citation)
        self.act_bibliography = self._act("Bibliography", self._insert_bibliography)
        self.act_caption = self._act("Insert Caption...", self._insert_caption)
        self.act_toc = self._act("Table of Contents", self._insert_toc)
        self.act_update_toc = self._act("Update Table of Contents", self._update_toc)

        self.act_row_above = self._act("Insert Row Above", lambda: self._table_insert_row(True))
        self.act_row_below = self._act("Insert Row Below", lambda: self._table_insert_row(False))
        self.act_col_left = self._act("Insert Column Left", lambda: self._table_insert_column(True))
        self.act_col_right = self._act("Insert Column Right", lambda: self._table_insert_column(False))
        self.act_delete_row = self._act("Delete Row", self._table_delete_row)
        self.act_delete_col = self._act("Delete Column", self._table_delete_column)
        self.act_delete_table = self._act("Delete Table", self._table_delete)
        self.act_merge_cells = self._act("Merge Cells", self._table_merge_cells)
        self.act_split_cell = self._act("Split Cell", self._table_split_cell)
        self.act_cell_properties = self._act("Cell Properties...", self._table_cell_properties)
        self.act_table_properties = self._act("Table Properties...", self._table_properties)

        self.act_spelling = self._act("Spelling...", self._spelling, "F7", icon=make_glyph_icon("âœ“"))
        self.act_thesaurus = self._act("Thesaurus...", self._thesaurus, "Shift+F7", icon=make_glyph_icon("T"))
        self.act_word_count = self._act("Word Count...", self._word_count, icon=make_glyph_icon("123"))
        self.act_delete_comment = self._act("Delete Comment", self._delete_selected_comment, icon=make_glyph_icon("âœ•"))
        self.act_toggle_comments = self._act("Show Comments Pane", self._toggle_comments_pane, checkable=True)
        self.act_track_changes = self._act("Track Changes", self._toggle_track_changes, checkable=True, icon=make_glyph_icon("âœŽ"))
        self.act_toggle_revisions = self._act("Show Revision Log", self._toggle_revisions_pane, checkable=True)
        self.act_language = self._act("Set Language...", self._set_language)

        self.act_zoom_in = self._act("Zoom In", lambda: self._zoom_delta(10), "Ctrl+=", icon=make_glyph_icon("ï¼‹"))
        self.act_zoom_out = self._act("Zoom Out", lambda: self._zoom_delta(-10), "Ctrl+-", icon=make_glyph_icon("ï¼"))
        self.act_zoom_reset = self._act("100%", lambda: self._set_zoom(100), "Ctrl+0")
        self.act_zoom_dialog = self._act("Zoom...", self._zoom_dialog)
        self.act_toggle_ruler = self._act("Ruler", self._toggle_ruler, checkable=True, icon=make_glyph_icon("â–­"))
        self.act_toggle_nav = self._act("Navigation Pane", self._toggle_navigation, checkable=True, icon=make_glyph_icon("â‰¡"))
        self.act_toggle_status = self._act("Status Bar", self._toggle_statusbar, checkable=True)
        self.act_dark = self._act("Dark Mode", self._toggle_dark, checkable=True, icon=make_glyph_icon("â—"))
        self.act_focus_mode = self._act("Focus Mode", self._toggle_focus_mode, checkable=True)
        self.act_full_screen = self._act("Full Screen", self._toggle_full_screen, "F11", checkable=True)

        self.act_preferences = self._act("Options...", self._preferences)
        self.act_about = self._act(f"About {APP_NAME}", self._about, icon=QIcon(make_logo_pixmap(32)))
        self.act_shortcuts = self._act("Keyboard Shortcuts", self._shortcuts, icon=icon_info)

    def _create_ui(self):
        self.quick_toolbar = QToolBar("Quick Access", self)
        self.quick_toolbar.setObjectName("quickToolbar")
        self.quick_toolbar.setMovable(False)
        self.quick_toolbar.setIconSize(QSize(18, 18))
        logo_label = QLabel()
        logo_label.setPixmap(make_logo_pixmap(24))
        logo_label.setContentsMargins(4, 0, 4, 0)
        brand_label = QLabel(APP_NAME)
        brand_label.setObjectName("brandLabel")
        self.quick_toolbar.addWidget(logo_label)
        self.quick_toolbar.addWidget(brand_label)
        self.quick_toolbar.addSeparator()
        for action in (self.act_new, self.act_open, self.act_save, self.act_undo, self.act_redo, self.act_print):
            self.quick_toolbar.addAction(action)
        spacer = QWidget()
        spacer.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        self.quick_toolbar.addWidget(spacer)
        self.doc_label = QLabel("")
        self.doc_label.setObjectName("docLabel")
        self.quick_toolbar.addWidget(self.doc_label)
        self.addToolBar(Qt.ToolBarArea.TopToolBarArea, self.quick_toolbar)

        central = QWidget()
        layout = QVBoxLayout(central)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        self.ribbon = QTabWidget()
        self.ribbon.setObjectName("ribbon")
        self.ribbon.setDocumentMode(True)
        self.ribbon.setFixedHeight(150)
        self._build_home_tab()
        self._build_insert_tab()
        self._build_design_tab()
        self._build_layout_tab()
        self._build_references_tab()
        self._build_review_tab()
        self._build_view_tab()
        self._build_help_tab()
        layout.addWidget(self.ribbon)
        self.tab_widget = QTabWidget()
        self.tab_widget.setObjectName("docTabs")
        self.tab_widget.setTabsClosable(True)
        self.tab_widget.setMovable(True)
        self.tab_widget.setDocumentMode(True)
        self.tab_widget.tabCloseRequested.connect(self._close_tab)
        self.tab_widget.currentChanged.connect(self._on_tab_changed)
        new_tab_button = QToolButton()
        new_tab_button.setText("+")
        new_tab_button.setToolTip("New document tab")
        new_tab_button.setAutoRaise(True)
        new_tab_button.clicked.connect(lambda: self._new_document())
        self.tab_widget.setCornerWidget(new_tab_button, Qt.Corner.TopRightCorner)
        layout.addWidget(self.tab_widget, 1)
        self.setCentralWidget(central)

    def _group(self, title, layout):
        widget = QWidget()
        outer = QVBoxLayout(widget)
        outer.setContentsMargins(8, 4, 8, 2)
        outer.setSpacing(1)
        row = QWidget()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(2)
        row.setLayout(layout)
        outer.addWidget(row, 1)
        label = QLabel(title)
        label.setObjectName("ribbonGroupLabel")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        outer.addWidget(label)
        separator = QFrame()
        separator.setFrameShape(QFrame.Shape.VLine)
        separator.setStyleSheet("color: #D0D0D0;")
        container = QWidget()
        horizontal = QHBoxLayout(container)
        horizontal.setContentsMargins(0, 0, 0, 0)
        horizontal.setSpacing(0)
        horizontal.addWidget(widget)
        horizontal.addWidget(separator)
        return container

    def _btn(self, action, large=False, text=None, width=None):
        button = QToolButton()
        button.setDefaultAction(action)
        button.setAutoRaise(True)
        if large:
            if action.icon().isNull():
                button.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextOnly)
            else:
                button.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextUnderIcon)
            button.setMinimumHeight(62)
            button.setMinimumWidth(56)
        else:
            if action.icon().isNull():
                button.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextOnly)
            else:
                button.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonIconOnly)
            button.setMinimumHeight(26)
            if width:
                button.setFixedWidth(width)
        if text is not None:
            button.setText(text)
        return button

    def _menu_button(self, action, menu, large=False, text=None, width=None):
        button = self._btn(action, large, text, width)
        button.setMenu(menu)
        button.setPopupMode(QToolButton.ToolButtonPopupMode.MenuButtonPopup)
        return button

    def _build_home_tab(self):
        tab = QWidget()
        layout = QHBoxLayout(tab)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        clipboard = QHBoxLayout()
        clipboard.addWidget(self._btn(self.act_paste, True, "Paste"))
        column = QVBoxLayout()
        column.setSpacing(1)
        column.addWidget(self._btn(self.act_cut, text="Cut"))
        column.addWidget(self._btn(self.act_copy, text="Copy"))
        clipboard.addLayout(column)
        column2 = QVBoxLayout()
        column2.setSpacing(1)
        column2.addWidget(self._btn(self.act_format_painter, text="Format Painter", width=110))
        column2.addWidget(self._btn(self.act_paste_special, text="Paste Text", width=110))
        clipboard.addLayout(column2)
        layout.addWidget(self._group("Clipboard", clipboard))
        font_layout = QVBoxLayout()
        font_layout.setSpacing(2)
        font_top = QHBoxLayout()
        font_top.setSpacing(2)
        self.font_combo = QFontComboBox()
        self.font_combo.setFixedWidth(150)
        self.font_combo.setCurrentFont(QFont(self.default_font_family))
        self.font_combo.currentFontChanged.connect(lambda font: self._merge_fmt("family", [font.family()]))
        font_top.addWidget(self.font_combo)
        self.size_combo = QComboBox()
        self.size_combo.setEditable(True)
        self.size_combo.setFixedWidth(52)
        for size in ("8", "9", "10", "11", "12", "14", "16", "18", "20", "24", "28", "32", "48", "72"):
            self.size_combo.addItem(size)
        self.size_combo.setCurrentText(str(self.default_font_size))
        self.size_combo.currentTextChanged.connect(self._on_size_changed)
        font_top.addWidget(self.size_combo)
        font_top.addWidget(self._btn(self.act_grow_font))
        font_top.addWidget(self._btn(self.act_shrink_font))
        font_layout.addLayout(font_top)
        font_bottom = QHBoxLayout()
        font_bottom.setSpacing(1)
        for action in (
            self.act_bold, self.act_italic, self.act_underline, self.act_strike,
            self.act_subscript, self.act_superscript, self.act_text_color,
            self.act_highlight, self.act_change_case, self.act_clear_format,
        ):
            font_bottom.addWidget(self._btn(action))
        font_layout.addLayout(font_bottom)
        layout.addWidget(self._group("Font", font_layout))
        paragraph_layout = QVBoxLayout()
        paragraph_layout.setSpacing(2)
        paragraph_top = QHBoxLayout()
        paragraph_top.setSpacing(1)
        paragraph_top.addWidget(self._menu_button(self.act_bullets, self._bullets_menu()))
        paragraph_top.addWidget(self._menu_button(self.act_numbering, self._numbering_menu()))
        paragraph_top.addWidget(self._btn(self.act_indent_dec))
        paragraph_top.addWidget(self._btn(self.act_indent_inc))
        paragraph_top.addWidget(self._btn(self.act_sort))
        paragraph_layout.addLayout(paragraph_top)
        paragraph_bottom = QHBoxLayout()
        paragraph_bottom.setSpacing(1)
        for action in (self.act_align_left, self.act_align_center, self.act_align_right, self.act_align_justify):
            paragraph_bottom.addWidget(self._btn(action))
        paragraph_bottom.addWidget(self._menu_button(self.act_line_spacing, self._line_spacing_menu()))
        paragraph_bottom.addWidget(self._btn(self.act_shading))
        paragraph_layout.addLayout(paragraph_bottom)
        layout.addWidget(self._group("Paragraph", paragraph_layout))
        style_layout = QHBoxLayout()
        self.style_list = QListWidget()
        self.style_list.setFlow(QListWidget.Flow.LeftToRight)
        self.style_list.setWrapping(True)
        self.style_list.setResizeMode(QListWidget.ResizeMode.Adjust)
        self.style_list.setGridSize(QSize(122, 28))
        self.style_list.setFixedSize(252, 64)
        self.style_list.setObjectName("styleGallery")
        for name, definition in STYLE_DEFS.items():
            item = QListWidgetItem(name)
            font = QFont("Calibri", 9 if name != "Title" else 11)
            font.setBold(definition["bold"])
            font.setItalic(definition["italic"])
            item.setFont(font)
            item.setData(Qt.ItemDataRole.UserRole, name)
            self.style_list.addItem(item)
        self.style_list.itemClicked.connect(lambda item: self._apply_style(item.data(Qt.ItemDataRole.UserRole)))
        style_layout.addWidget(self.style_list)
        layout.addWidget(self._group("Styles", style_layout))
        editing_layout = QHBoxLayout()
        editing_layout.setSpacing(2)
        editing_layout.addWidget(self._btn(self.act_find, text="Find"))
        editing_layout.addWidget(self._btn(self.act_replace, text="Replace"))
        editing_layout.addWidget(self._btn(self.act_select_all, text="Select All"))
        layout.addWidget(self._group("Editing", editing_layout))
        layout.addStretch()
        self.ribbon.addTab(tab, "Home")

    def _bullets_menu(self):
        menu = QMenu(self)
        for label, style in (
            ("Standard Bullets", QTextListFormat.Style.ListDisc),
            ("Open Circle", QTextListFormat.Style.ListCircle),
            ("Square", QTextListFormat.Style.ListSquare),
        ):
            action = menu.addAction(label)
            action.triggered.connect(lambda checked=False, value=style: self._apply_list(value))
        menu.addSeparator()
        remove = menu.addAction("No Bullets")
        remove.triggered.connect(self._remove_list)
        return menu

    def _numbering_menu(self):
        menu = QMenu(self)
        for label, style in (
            ("1, 2, 3", QTextListFormat.Style.ListDecimal),
            ("a, b, c", QTextListFormat.Style.ListLowerAlpha),
            ("A, B, C", QTextListFormat.Style.ListUpperAlpha),
            ("i, ii, iii", QTextListFormat.Style.ListLowerRoman),
            ("I, II, III", QTextListFormat.Style.ListUpperRoman),
        ):
            action = menu.addAction(label)
            action.triggered.connect(lambda checked=False, value=style: self._apply_list(value))
        menu.addSeparator()
        remove = menu.addAction("No Numbering")
        remove.triggered.connect(self._remove_list)
        return menu

    def _line_spacing_menu(self):
        menu = QMenu(self)
        for label, value in (("1.0", 100), ("1.15", 115), ("1.5", 150), ("2.0", 200)):
            action = menu.addAction(label)
            action.triggered.connect(lambda checked=False, v=value: self._apply_line_spacing(v))
        return menu

    def _case_menu(self):
        menu = QMenu(self)
        for label, mode in (
            ("Sentence case", "sentence"),
            ("lowercase", "lower"),
            ("UPPERCASE", "upper"),
            ("Capitalize Each Word", "title"),
            ("tOGGLE cASE", "toggle"),
        ):
            action = menu.addAction(label)
            action.triggered.connect(lambda checked=False, m=mode: self._change_case(m))
        return menu

    def _shapes_menu(self):
        menu = QMenu(self)
        for kind in ("Rectangle", "Rounded Rectangle", "Ellipse", "Triangle", "Arrow", "Star", "Line"):
            action = menu.addAction(kind)
            action.triggered.connect(lambda checked=False, k=kind: self._insert_shape(k))
        return menu

    def _margins_menu(self):
        menu = QMenu(self)
        for label, inches in (("Normal (1 inch)", 1.0), ("Narrow (0.5 inch)", 0.5), ("Wide (2 inches)", 2.0)):
            action = menu.addAction(label)
            action.triggered.connect(lambda checked=False, value=inches: self._set_margin_inches(value))
        menu.addSeparator()
        custom = menu.addAction("Custom Margins...")
        custom.triggered.connect(self._page_setup)
        return menu

    def _orientation_menu(self):
        menu = QMenu(self)
        portrait = menu.addAction("Portrait")
        portrait.triggered.connect(lambda: self._set_orientation("portrait"))
        landscape = menu.addAction("Landscape")
        landscape.triggered.connect(lambda: self._set_orientation("landscape"))
        return menu

    def _page_size_menu(self):
        menu = QMenu(self)
        for name in PAGE_SIZES:
            if name == "Custom":
                continue
            action = menu.addAction(name)
            action.triggered.connect(lambda checked=False, n=name: self._set_page_size(n))
        menu.addSeparator()
        custom = menu.addAction("More Paper Sizes...")
        custom.triggered.connect(self._page_setup)
        return menu

    def _page_color_menu(self):
        menu = QMenu(self)
        for label, color in (
            ("White", "#FFFFFF"),
            ("Warm Cream", "#FFF8E7"),
            ("Light Blue", "#EAF3FB"),
            ("Light Green", "#EDF7EE"),
            ("Light Gray", "#F2F2F2"),
            ("Light Yellow", "#FFF9DB"),
        ):
            action = menu.addAction(label)
            action.triggered.connect(lambda checked=False, c=color: self._set_page_color(QColor(c)))
        menu.addSeparator()
        more = menu.addAction("More Colors...")
        more.triggered.connect(self._choose_page_color)
        no_color = menu.addAction("No Color")
        no_color.triggered.connect(lambda: self._set_page_color(QColor("#FFFFFF")))
        return menu

    def _watermark_menu(self):
        menu = QMenu(self)
        add = menu.addAction("Add Watermark...")
        add.triggered.connect(self._add_watermark)
        remove = menu.addAction("Remove Watermark")
        remove.triggered.connect(self._remove_watermark)
        return menu

    def _build_insert_tab(self):
        tab = QWidget()
        layout = QHBoxLayout(tab)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        pages = QHBoxLayout()
        for action in (self.act_cover_page, self.act_blank_page, self.act_page_break):
            pages.addWidget(self._btn(action, True))
        layout.addWidget(self._group("Pages", pages))
        tables = QHBoxLayout()
        tables.addWidget(self._btn(self.act_table, True))
        layout.addWidget(self._group("Tables", tables))
        illustrations = QHBoxLayout()
        illustrations.addWidget(self._btn(self.act_picture, True))
        illustrations.addWidget(self._menu_button(self.act_shapes, self._shapes_menu(), True))
        illustrations.addWidget(self._btn(self.act_chart, True))
        layout.addWidget(self._group("Illustrations", illustrations))
        links = QHBoxLayout()
        links.addWidget(self._btn(self.act_hyperlink, True))
        links.addWidget(self._btn(self.act_bookmark, True))
        layout.addWidget(self._group("Links", links))
        header_footer = QHBoxLayout()
        header_footer.addWidget(self._btn(self.act_header_footer, True))
        header_footer.addWidget(self._btn(self.act_page_number, True))
        layout.addWidget(self._group("Header and Footer", header_footer))
        text_group = QHBoxLayout()
        for action in (self.act_text_box, self.act_date_time, self.act_symbol, self.act_horizontal_rule, self.act_footnote):
            text_group.addWidget(self._btn(action, True))
        layout.addWidget(self._group("Text", text_group))
        comments = QHBoxLayout()
        comments.addWidget(self._btn(self.act_comment, True))
        layout.addWidget(self._group("Comments", comments))
        layout.addStretch()
        self.ribbon.addTab(tab, "Insert")

    def _build_design_tab(self):
        tab = QWidget()
        layout = QHBoxLayout(tab)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        formatting = QHBoxLayout()
        page_color_action = self._act("Page Color", lambda: None, icon=make_color_icon(QColor("#FFF8E7"), "â–­"))
        formatting.addWidget(self._menu_button(page_color_action, self._page_color_menu(), True))
        watermark_action = self._act("Watermark", lambda: None, icon=make_glyph_icon("â‰ˆ"))
        formatting.addWidget(self._menu_button(watermark_action, self._watermark_menu(), True))
        border_action = self._act("Page Borders", self._toggle_page_border, checkable=True, icon=make_glyph_icon("â–£"))
        formatting.addWidget(self._btn(border_action, True))
        layout.addWidget(self._group("Document Formatting", formatting))
        spacing = QHBoxLayout()
        for label, before, after, line in (
            ("No Spacing", 0, 0, 100),
            ("Compact", 2, 4, 100),
            ("Open", 10, 12, 115),
        ):
            button = QPushButton(label)
            button.setMinimumHeight(50)
            button.clicked.connect(lambda checked=False, b=before, a=after, l=line: self._set_paragraph_spacing(b, a, l))
            spacing.addWidget(button)
        layout.addWidget(self._group("Paragraph Spacing", spacing))
        layout.addStretch()
        self.ribbon.addTab(tab, "Design")

    def _build_layout_tab(self):
        tab = QWidget()
        layout = QHBoxLayout(tab)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        page_setup = QHBoxLayout()
        margins_action = self._act("Margins", lambda: self._set_margin_inches(1.0), icon=make_glyph_icon("â–­"))
        page_setup.addWidget(self._menu_button(margins_action, self._margins_menu(), True))
        orientation_action = self._act("Orientation", lambda: self._set_orientation("portrait"), icon=make_glyph_icon("â–¯"))
        page_setup.addWidget(self._menu_button(orientation_action, self._orientation_menu(), True))
        size_action = self._act("Size", lambda: self._set_page_size("Letter"), icon=make_glyph_icon("â–¤"))
        page_setup.addWidget(self._menu_button(size_action, self._page_size_menu(), True))
        page_setup.addWidget(self._btn(self.act_page_setup, True))
        layout.addWidget(self._group("Page Setup", page_setup))
        paragraph = QHBoxLayout()
        paragraph.addWidget(self._btn(self.act_indent_dec, True))
        paragraph.addWidget(self._btn(self.act_indent_inc, True))
        paragraph.addWidget(self._menu_button(self.act_line_spacing, self._line_spacing_menu(), True))
        paragraph.addWidget(self._btn(self.act_paragraph_dialog, True))
        layout.addWidget(self._group("Paragraph", paragraph))
        layout.addStretch()
        self.ribbon.addTab(tab, "Layout")

    def _build_references_tab(self):
        tab = QWidget()
        layout = QHBoxLayout(tab)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        toc_layout = QHBoxLayout()
        toc_layout.addWidget(self._btn(self.act_toc, True))
        toc_layout.addWidget(self._btn(self.act_update_toc, True))
        layout.addWidget(self._group("Table of Contents", toc_layout))
        footnotes = QHBoxLayout()
        footnotes.addWidget(self._btn(self.act_footnote, True))
        layout.addWidget(self._group("Footnotes", footnotes))
        citations = QHBoxLayout()
        citations.addWidget(self._btn(self.act_citation, True))
        citations.addWidget(self._btn(self.act_bibliography, True))
        layout.addWidget(self._group("Citations and Bibliography", citations))
        captions = QHBoxLayout()
        captions.addWidget(self._btn(self.act_caption, True))
        layout.addWidget(self._group("Captions", captions))
        bookmarks = QHBoxLayout()
        bookmarks.addWidget(self._btn(self.act_bookmark, True))
        layout.addWidget(self._group("Bookmarks", bookmarks))
        layout.addStretch()
        self.ribbon.addTab(tab, "References")

    def _build_review_tab(self):
        tab = QWidget()
        layout = QHBoxLayout(tab)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        proofing = QHBoxLayout()
        proofing.addWidget(self._btn(self.act_spelling, True))
        proofing.addWidget(self._btn(self.act_thesaurus, True))
        proofing.addWidget(self._btn(self.act_word_count, True))
        layout.addWidget(self._group("Proofing", proofing))
        comments = QHBoxLayout()
        comments.addWidget(self._btn(self.act_comment, True))
        comments.addWidget(self._btn(self.act_delete_comment, True))
        comments.addWidget(self._btn(self.act_toggle_comments, True))
        layout.addWidget(self._group("Comments", comments))
        tracking = QHBoxLayout()
        tracking.addWidget(self._btn(self.act_track_changes, True))
        tracking.addWidget(self._btn(self.act_toggle_revisions, True))
        layout.addWidget(self._group("Tracking", tracking))
        language = QHBoxLayout()
        language.addWidget(self._btn(self.act_language, True))
        layout.addWidget(self._group("Language", language))
        layout.addStretch()
        self.ribbon.addTab(tab, "Review")

    def _build_view_tab(self):
        tab = QWidget()
        layout = QHBoxLayout(tab)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        views = QHBoxLayout()
        views.addWidget(self._btn(self.act_focus_mode, True))
        views.addWidget(self._btn(self.act_full_screen, True))
        layout.addWidget(self._group("Views", views))
        show = QHBoxLayout()
        for action in (self.act_toggle_ruler, self.act_toggle_nav, self.act_toggle_status, self.act_toggle_comments, self.act_toggle_revisions):
            show.addWidget(self._btn(action, True))
        layout.addWidget(self._group("Show", show))
        zoom = QHBoxLayout()
        zoom.addWidget(self._btn(self.act_zoom_in, True))
        zoom.addWidget(self._btn(self.act_zoom_out, True))
        zoom.addWidget(self._btn(self.act_zoom_reset, True))
        zoom.addWidget(self._btn(self.act_zoom_dialog, True))
        layout.addWidget(self._group("Zoom", zoom))
        theme = QHBoxLayout()
        theme.addWidget(self._btn(self.act_dark, True))
        layout.addWidget(self._group("Theme", theme))
        layout.addStretch()
        self.ribbon.addTab(tab, "View")

    def _build_help_tab(self):
        tab = QWidget()
        layout = QHBoxLayout(tab)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        help_layout = QHBoxLayout()
        help_layout.addWidget(self._btn(self.act_about, True))
        help_layout.addWidget(self._btn(self.act_shortcuts, True))
        layout.addWidget(self._group("Help", help_layout))
        layout.addStretch()
        self.ribbon.addTab(tab, "Help")

    def _create_menus(self):
        menu_bar = self.menuBar()
        file_menu = menu_bar.addMenu("&File")
        for action in (self.act_new, self.act_open, self.act_save, self.act_save_as, self.act_save_all, self.act_close_tab):
            file_menu.addAction(action)
        file_menu.addSeparator()
        self.recent_menu = file_menu.addMenu("Recent Documents")
        self.recent_menu.aboutToShow.connect(self._refresh_recent_menu)
        file_menu.addSeparator()
        export_menu = file_menu.addMenu("Export")
        export_menu.addAction(self.act_export_docx)
        export_menu.addAction(self.act_export_pdf)
        export_menu.addAction(self.act_export_html)
        export_menu.addAction(self.act_export_md)
        export_menu.addAction(self.act_export_txt)
        file_menu.addAction(self.act_print)
        file_menu.addAction(self.act_print_preview)
        file_menu.addSeparator()
        file_menu.addAction(self.act_page_setup)
        file_menu.addAction(self.act_properties)
        file_menu.addSeparator()
        file_menu.addAction(self.act_exit)
        edit_menu = menu_bar.addMenu("&Edit")
        for action in (self.act_undo, self.act_redo):
            edit_menu.addAction(action)
        edit_menu.addSeparator()
        for action in (self.act_cut, self.act_copy, self.act_paste, self.act_paste_special):
            edit_menu.addAction(action)
        edit_menu.addSeparator()
        edit_menu.addAction(self.act_select_all)
        edit_menu.addAction(self.act_clear_contents)
        edit_menu.addSeparator()
        edit_menu.addAction(self.act_find)
        edit_menu.addAction(self.act_replace)
        edit_menu.addAction(self.act_find_next)
        edit_menu.addAction(self.act_go_to)
        view_menu = menu_bar.addMenu("&View")
        view_menu.addAction(self.act_focus_mode)
        view_menu.addAction(self.act_full_screen)
        view_menu.addSeparator()
        for action in (self.act_toggle_ruler, self.act_toggle_nav, self.act_toggle_status, self.act_toggle_comments, self.act_toggle_revisions):
            view_menu.addAction(action)
        view_menu.addSeparator()
        for action in (self.act_zoom_in, self.act_zoom_out, self.act_zoom_reset, self.act_zoom_dialog):
            view_menu.addAction(action)
        view_menu.addSeparator()
        view_menu.addAction(self.act_dark)
        insert_menu = menu_bar.addMenu("&Insert")
        insert_menu.addAction(self.act_cover_page)
        insert_menu.addAction(self.act_blank_page)
        insert_menu.addAction(self.act_page_break)
        insert_menu.addSeparator()
        insert_menu.addAction(self.act_table)
        insert_menu.addAction(self.act_picture)
        insert_menu.addAction(self.act_shapes)
        insert_menu.addAction(self.act_chart)
        insert_menu.addSeparator()
        insert_menu.addAction(self.act_hyperlink)
        insert_menu.addAction(self.act_bookmark)
        insert_menu.addSeparator()
        insert_menu.addAction(self.act_header_footer)
        insert_menu.addAction(self.act_page_number)
        insert_menu.addSeparator()
        insert_menu.addAction(self.act_text_box)
        insert_menu.addAction(self.act_date_time)
        insert_menu.addAction(self.act_symbol)
        insert_menu.addAction(self.act_horizontal_rule)
        insert_menu.addAction(self.act_footnote)
        insert_menu.addSeparator()
        insert_menu.addAction(self.act_comment)
        format_menu = menu_bar.addMenu("F&ormat")
        format_menu.addAction(self.act_font_dialog)
        format_menu.addAction(self.act_paragraph_dialog)
        format_menu.addSeparator()
        format_menu.addAction(self.act_bold)
        format_menu.addAction(self.act_italic)
        format_menu.addAction(self.act_underline)
        format_menu.addAction(self.act_strike)
        format_menu.addAction(self.act_superscript)
        format_menu.addAction(self.act_subscript)
        format_menu.addSeparator()
        format_menu.addAction(self.act_text_color)
        format_menu.addAction(self.act_highlight)
        format_menu.addAction(self.act_change_case)
        format_menu.addAction(self.act_clear_format)
        format_menu.addAction(self.act_format_painter)
        format_menu.addSeparator()
        align_menu = format_menu.addMenu("Alignment")
        for action in (self.act_align_left, self.act_align_center, self.act_align_right, self.act_align_justify):
            align_menu.addAction(action)
        format_menu.addAction(self.act_bullets)
        format_menu.addAction(self.act_numbering)
        format_menu.addAction(self.act_indent_inc)
        format_menu.addAction(self.act_indent_dec)
        format_menu.addAction(self.act_line_spacing)
        format_menu.addAction(self.act_shading)
        format_menu.addAction(self.act_sort)
        table_menu = menu_bar.addMenu("&Table")
        table_menu.addAction(self.act_table)
        table_menu.addSeparator()
        for action in (self.act_row_above, self.act_row_below, self.act_col_left, self.act_col_right):
            table_menu.addAction(action)
        table_menu.addSeparator()
        for action in (self.act_delete_row, self.act_delete_col, self.act_delete_table):
            table_menu.addAction(action)
        table_menu.addSeparator()
        table_menu.addAction(self.act_merge_cells)
        table_menu.addAction(self.act_split_cell)
        table_menu.addSeparator()
        table_menu.addAction(self.act_cell_properties)
        table_menu.addAction(self.act_table_properties)
        tools_menu = menu_bar.addMenu("&Tools")
        tools_menu.addAction(self.act_spelling)
        tools_menu.addAction(self.act_thesaurus)
        tools_menu.addAction(self.act_word_count)
        tools_menu.addSeparator()
        tools_menu.addAction(self.act_track_changes)
        tools_menu.addAction(self.act_toggle_revisions)
        tools_menu.addSeparator()
        tools_menu.addAction(self.act_preferences)
        references_menu = menu_bar.addMenu("&References")
        references_menu.addAction(self.act_toc)
        references_menu.addAction(self.act_update_toc)
        references_menu.addSeparator()
        references_menu.addAction(self.act_footnote)
        references_menu.addAction(self.act_citation)
        references_menu.addAction(self.act_bibliography)
        references_menu.addAction(self.act_caption)
        references_menu.addSeparator()
        references_menu.addAction(self.act_bookmark)
        help_menu = menu_bar.addMenu("&Help")
        help_menu.addAction(self.act_about)
        help_menu.addAction(self.act_shortcuts)

    def _create_docks(self):
        self.nav_dock = QDockWidget("Navigation", self)
        self.nav_dock.setObjectName("navDock")
        self.nav_dock.setAllowedAreas(Qt.DockWidgetArea.LeftDockWidgetArea | Qt.DockWidgetArea.RightDockWidgetArea)
        nav_widget = QWidget()
        nav_layout = QVBoxLayout(nav_widget)
        nav_layout.setContentsMargins(6, 6, 6, 6)
        self.nav_tabs = QTabWidget()
        self.nav_headings = QListWidget()
        self.nav_headings.setToolTip("Document headings")
        self.nav_headings.itemClicked.connect(self._nav_heading_clicked)
        self.nav_bookmarks = QListWidget()
        self.nav_bookmarks.setToolTip("Document bookmarks")
        self.nav_bookmarks.itemClicked.connect(self._nav_bookmark_clicked)
        self.nav_tabs.addTab(self.nav_headings, "Headings")
        self.nav_tabs.addTab(self.nav_bookmarks, "Bookmarks")
        nav_layout.addWidget(self.nav_tabs, 1)
        nav_buttons = QHBoxLayout()
        add_bookmark = QPushButton("Add Bookmark")
        add_bookmark.clicked.connect(self._insert_bookmark)
        refresh = QPushButton("Refresh")
        refresh.clicked.connect(self._refresh_navigation)
        nav_buttons.addWidget(add_bookmark)
        nav_buttons.addWidget(refresh)
        nav_layout.addLayout(nav_buttons)
        self.nav_dock.setWidget(nav_widget)
        self.addDockWidget(Qt.DockWidgetArea.LeftDockWidgetArea, self.nav_dock)
        self.nav_dock.visibilityChanged.connect(lambda visible: self.act_toggle_nav.setChecked(visible))
        self.comments_dock = QDockWidget("Comments", self)
        self.comments_dock.setObjectName("commentsDock")
        comments_widget = QWidget()
        comments_layout = QVBoxLayout(comments_widget)
        comments_layout.setContentsMargins(6, 6, 6, 6)
        self.comments_list = QListWidget()
        self.comments_list.itemClicked.connect(self._comment_item_clicked)
        comments_layout.addWidget(self.comments_list, 1)
        comment_buttons = QHBoxLayout()
        new_comment = QPushButton("New")
        new_comment.clicked.connect(self._add_comment)
        go_comment = QPushButton("Go To")
        go_comment.clicked.connect(self._go_to_selected_comment)
        delete_comment = QPushButton("Delete")
        delete_comment.clicked.connect(self._delete_selected_comment)
        comment_buttons.addWidget(new_comment)
        comment_buttons.addWidget(go_comment)
        comment_buttons.addWidget(delete_comment)
        comments_layout.addLayout(comment_buttons)
        self.comments_dock.setWidget(comments_widget)
        self.addDockWidget(Qt.DockWidgetArea.RightDockWidgetArea, self.comments_dock)
        self.comments_dock.visibilityChanged.connect(lambda visible: self.act_toggle_comments.setChecked(visible))
        self.revisions_dock = QDockWidget("Revision Log", self)
        self.revisions_dock.setObjectName("revisionsDock")
        revisions_widget = QWidget()
        revisions_layout = QVBoxLayout(revisions_widget)
        revisions_layout.setContentsMargins(6, 6, 6, 6)
        self.revision_view = QPlainTextEdit()
        self.revision_view.setReadOnly(True)
        revisions_layout.addWidget(self.revision_view, 1)
        clear_button = QPushButton("Clear Log")
        clear_button.clicked.connect(self._clear_revision_log)
        revisions_layout.addWidget(clear_button)
        self.revisions_dock.setWidget(revisions_widget)
        self.addDockWidget(Qt.DockWidgetArea.RightDockWidgetArea, self.revisions_dock)
        self.revisions_dock.visibilityChanged.connect(lambda visible: self.act_toggle_revisions.setChecked(visible))
        self.tabifyDockWidget(self.comments_dock, self.revisions_dock)
        self.comments_dock.hide()
        self.revisions_dock.hide()

    def _create_statusbar(self):
        status = QStatusBar()
        status.setObjectName("statusBar")
        self.setStatusBar(status)
        self.lbl_page = QLabel("Page 1 of 1")
        self.lbl_words = QLabel("0 words")
        self.lbl_chars = QLabel("0 characters")
        self.lbl_linecol = QLabel("Ln 1, Col 1")
        self.lbl_language = QLabel("English (US)")
        for label in (self.lbl_page, self.lbl_words, self.lbl_chars, self.lbl_linecol, self.lbl_language):
            status.addWidget(label)
            status.addWidget(self._separator())
        self.lbl_track = QLabel("")
        self.lbl_track.setStyleSheet("color: #D35400; font-weight: bold;")
        status.addPermanentWidget(self.lbl_track)
        self.lbl_autosave = QLabel("Autosave ready")
        status.addPermanentWidget(self.lbl_autosave)
        status.addPermanentWidget(self._separator())
        zoom_label = QLabel("Zoom")
        status.addPermanentWidget(zoom_label)
        self.zoom_slider = QSlider(Qt.Orientation.Horizontal)
        self.zoom_slider.setRange(50, 200)
        self.zoom_slider.setValue(100)
        self.zoom_slider.setFixedWidth(130)
        self.zoom_slider.valueChanged.connect(self._set_zoom_from_slider)
        status.addPermanentWidget(self.zoom_slider)
        self.zoom_label = QLabel("100%")
        self.zoom_label.setMinimumWidth(42)
        status.addPermanentWidget(self.zoom_label)

    def _separator(self):
        line = QFrame()
        line.setFrameShape(QFrame.Shape.VLine)
        line.setStyleSheet("color: #D0D0D0;")
        return line

    def _apply_theme(self):
        if self.dark:
            background = "#1E1E1E"
            panel = "#252526"
            panel2 = "#2D2D30"
            text = "#E8E8E8"
            muted = "#A0A0A0"
            border = "#3A3A3C"
            hover = "#3A3D41"
            selection = "#094771"
            accent = "#4A9EFF"
        else:
            background = "#F3F3F3"
            panel = "#FFFFFF"
            panel2 = "#F7F7F7"
            text = "#222222"
            muted = "#666666"
            border = "#DCDCDC"
            hover = "#E5F0FA"
            selection = "#D6E8F7"
            accent = "#2B7CD3"
        css = f"""
        QMainWindow {{ background: {background}; }}
        QMenuBar {{ background: {panel}; border-bottom: 1px solid {border}; color: {text}; }}
        QMenuBar::item {{ padding: 6px 10px; background: transparent; }}
        QMenuBar::item:selected {{ background: {hover}; }}
        QMenu {{ background: {panel}; border: 1px solid {border}; color: {text}; }}
        QMenu::item {{ padding: 6px 24px; }}
        QMenu::item:selected {{ background: {selection}; }}
        QMenu::separator {{ height: 1px; background: {border}; margin: 4px 8px; }}
        QToolBar#quickToolbar {{ background: {panel}; border-bottom: 1px solid {border}; spacing: 4px; padding: 3px 8px; }}
        QLabel#brandLabel {{ font-size: 15px; font-weight: bold; color: {accent}; padding: 0 8px 0 2px; }}
        QLabel#docLabel {{ color: {muted}; padding-right: 8px; }}
        QTabWidget#ribbon::pane {{ border-top: 1px solid {border}; background: {panel2}; }}
        QTabBar::tab {{ background: transparent; padding: 8px 16px; border: none; color: {text}; font-size: 10pt; }}
        QTabBar::tab:selected {{ background: {panel}; border-bottom: 2px solid {accent}; color: {accent}; }}
        QTabBar::tab:hover {{ background: {hover}; }}
        QTabWidget#docTabs::pane {{ border: none; }}
        QTabWidget#docTabs QTabBar::tab {{ background: {panel2}; padding: 7px 14px; border: 1px solid {border}; border-bottom: none; margin-right: 2px; }}
        QTabWidget#docTabs QTabBar::tab:selected {{ background: {panel}; color: {accent}; }}
        QToolButton {{ border: 1px solid transparent; border-radius: 4px; padding: 3px; background: transparent; color: {text}; }}
        QToolButton:hover {{ background: {hover}; border-color: {border}; }}
        QToolButton:pressed {{ background: {selection}; }}
        QToolButton:checked {{ background: {selection}; border-color: {accent}; }}
        QPushButton {{ background: {panel}; border: 1px solid {border}; border-radius: 4px; padding: 5px 10px; color: {text}; }}
        QPushButton:hover {{ background: {hover}; border-color: {accent}; }}
        QPushButton:disabled {{ color: {muted}; }}
        QStatusBar {{ background: {panel}; border-top: 1px solid {border}; color: {text}; font-size: 9pt; }}
        QStatusBar QLabel {{ color: {text}; padding: 0 8px; }}
        QComboBox, QFontComboBox, QLineEdit, QSpinBox, QDoubleSpinBox, QPlainTextEdit, QListWidget, QTextBrowser {{
            background: {panel}; border: 1px solid {border}; border-radius: 4px; color: {text}; padding: 2px 4px;
            selection-background-color: {selection};
        }}
        QComboBox:hover, QFontComboBox:hover, QLineEdit:hover {{ border-color: {accent}; }}
        QListWidget::item {{ padding: 4px; }}
        QListWidget::item:selected {{ background: {selection}; color: {text}; }}
        QDockWidget {{ color: {text}; }}
        QDockWidget::title {{ background: {panel2}; padding: 6px; border-bottom: 1px solid {border}; }}
        QDialog {{ background: {background}; color: {text}; }}
        QLabel {{ color: {text}; }}
        QLabel#ribbonGroupLabel {{ color: {muted}; font-size: 9px; }}
        QSlider::groove:horizontal {{ height: 4px; background: {border}; border-radius: 2px; }}
        QSlider::handle:horizontal {{ background: {accent}; width: 12px; margin: -5px 0; border-radius: 6px; }}
        QScrollBar:vertical {{ background: transparent; width: 12px; }}
        QScrollBar::handle:vertical {{ background: {border}; min-height: 24px; border-radius: 6px; }}
        QScrollBar::add-line, QScrollBar::sub-line {{ height: 0; }}
        QToolTip {{ background: {panel}; color: {text}; border: 1px solid {border}; padding: 4px; }}
        QTextEdit {{ selection-background-color: {selection}; }}
        """
        app = QApplication.instance()
        if app is not None:
            app.setStyleSheet(css)
        for index in range(self.tab_widget.count()) if hasattr(self, "tab_widget") else range(0):
            tab = self.tab_widget.widget(index)
            if isinstance(tab, DocumentTab):
                tab.set_dark(self.dark)

    def _apply_ui_settings(self):
        self.act_toggle_ruler.setChecked(self.settings.value("ui/ruler", True, bool))
        self.act_toggle_status.setChecked(self.settings.value("ui/statusbar", True, bool))
        self.act_toggle_nav.setChecked(self.settings.value("ui/navigation", True, bool))
        self.act_dark.setChecked(self.dark)
        self._toggle_ruler(self.act_toggle_ruler.isChecked())
        self.statusBar().setVisible(self.act_toggle_status.isChecked())
        self.nav_dock.setVisible(self.act_toggle_nav.isChecked())
        self.act_toggle_comments.setChecked(False)
        self.act_toggle_revisions.setChecked(False)
        self.zoom_slider.setValue(self._zoom)
        self.zoom_label.setText(f"{self._zoom}%")

    def _restore_window_state(self):
        geometry = self.settings.value("window/geometry")
        if geometry is not None:
            try:
                self.restoreGeometry(geometry)
            except Exception:
                pass

    def _save_window_state(self):
        self.settings.setValue("window/geometry", self.saveGeometry())

    def _setup_autosave(self):
        old_timer = getattr(self, "autosave_timer", None)
        if old_timer is not None:
            old_timer.stop()
            old_timer.deleteLater()
        self.autosave_timer = QTimer(self)
        self.autosave_timer.timeout.connect(self._autosave)
        enabled = self.settings.value("autosave/enabled", True, bool)
        interval = self.settings.value("autosave/interval", 60, int)
        if enabled:
            self.autosave_timer.start(max(15, int(interval)) * 1000)
            self.lbl_autosave.setText("Autosave on")
        else:
            self.autosave_timer.stop()
            self.lbl_autosave.setText("Autosave off")

    # ------------------------------------------------------------------
    # Tabs, signals and status
    # ------------------------------------------------------------------
    def _new_document(self, initial=False):
        self._doc_counter += 1
        tab = DocumentTab(self, f"Document{self._doc_counter}")
        tab.set_default_font(self.default_font_family, self.default_font_size)
        index = self.tab_widget.addTab(tab, tab.display_name())
        self._attach_tab(tab)
        self.tab_widget.setCurrentIndex(index)
        tab.apply_zoom(self._zoom)
        tab.refresh_ruler()
        tab.ruler.setVisible(self.act_toggle_ruler.isChecked())
        tab.editor.setFocus()
        self._update_tab_title(tab)
        self._refresh_navigation()
        self._refresh_comments()
        self._refresh_revision_log()
        self._update_statusbar()
        return tab

    def _attach_tab(self, tab):
        editor = tab.editor
        editor.cursorPositionChanged.connect(self._on_cursor_changed)
        editor.currentCharFormatChanged.connect(self._on_char_format_changed)
        editor.textChanged.connect(self._on_text_changed)
        editor.document().modificationChanged.connect(
            lambda modified, t=tab: self._on_modification_changed(t, modified)
        )
        editor.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        editor.customContextMenuRequested.connect(self._show_context_menu)
        editor.viewport().installEventFilter(self)

    def _on_tab_changed(self, index):
        tab = self._current_tab()
        if tab is None:
            return
        tab.apply_zoom(self._zoom)
        tab.refresh_ruler()
        tab.ruler.setVisible(self.act_toggle_ruler.isChecked())
        tab.set_dark(self.dark)
        self._update_tab_title(tab)
        self._update_statusbar()
        self._update_format_ui()
        self._refresh_navigation()
        self._refresh_comments()
        self._refresh_revision_log()
        self._ensure_cursor_visible(tab)

    def _update_tab_title(self, tab=None):
        tab = tab or self._current_tab()
        if tab is None:
            return
        index = self.tab_widget.indexOf(tab)
        if index < 0:
            return
        name = tab.display_name()
        if tab.editor.document().isModified():
            name += " *"
        self.tab_widget.setTabText(index, name)
        self.tab_widget.setTabToolTip(index, tab.file_path or "Not saved yet")
        if tab is self._current_tab():
            self.setWindowTitle(f"{tab.display_name()} - {APP_NAME}")
            self.doc_label.setText(tab.display_name())

    def _on_modification_changed(self, tab, modified):
        self._update_tab_title(tab)

    def _on_cursor_changed(self):
        tab = self._tab_for_editor(self.sender())
        if tab is None or tab is not self._current_tab():
            return
        self._update_statusbar()
        self._update_format_ui()
        self._ensure_cursor_visible(tab)

    def _on_char_format_changed(self, fmt):
        tab = self._tab_for_editor(self.sender())
        if tab is not None and tab is self._current_tab():
            self._update_format_ui()

    def _on_text_changed(self):
        tab = self._tab_for_editor(self.sender())
        if tab is None:
            return
        self._update_tab_title(tab)
        if self._track_changes:
            if tab._pending_old is None:
                tab._pending_old = tab.last_text
            self._diff_timer.start(1200)
        self._nav_timer.start(500)
        if tab is self._current_tab():
            self._update_statusbar()
            self._refresh_comments()

    def _update_statusbar(self):
        tab = self._current_tab()
        if tab is None:
            return
        text = tab.editor.toPlainText()
        words = len(re.findall(r"\S+", text))
        chars = len(text)
        cursor = tab.editor.textCursor()
        line = cursor.blockNumber() + 1
        column = cursor.columnNumber() + 1
        pages = tab.paper.pages
        try:
            top = tab.editor.document().documentLayout().blockBoundingRect(cursor.block()).top()
            page_height = max(1.0, float(tab.paper.page_h))
            current_page = max(1, min(pages, int(top // page_height) + 1))
        except Exception:
            current_page = 1
        self.lbl_page.setText(f"Page {current_page} of {pages}")
        self.lbl_words.setText(f"{words} word{'s' if words != 1 else ''}")
        self.lbl_chars.setText(f"{chars} character{'s' if chars != 1 else ''}")
        self.lbl_linecol.setText(f"Ln {line}, Col {column}")
        self.lbl_language.setText("English (US)")
        self.lbl_track.setText("TRACKING" if self._track_changes else "")

    def _ensure_cursor_visible(self, tab):
        try:
            rect = tab.editor.cursorRect()
            top_left = tab.editor.mapTo(tab.paper, rect.topLeft())
            scene_rect = tab.proxy.mapRectToScene(
                QRectF(float(top_left.x()), float(top_left.y()), float(rect.width()), float(rect.height()))
            )
            tab.view.ensureVisible(scene_rect, 30, 60)
        except Exception:
            pass

    # ------------------------------------------------------------------
    # Native file format, opening and saving
    # ------------------------------------------------------------------
    def _serialize_tab(self, tab):
        paper = tab.paper
        return {
            "app": APP_NAME,
            "format_version": 1,
            "title": tab.display_name(),
            "html": tab.editor.toHtml(),
            "page": {
                "width": paper.base_page_w,
                "height": paper.base_page_h,
                "margin": paper.base_margin,
                "color": paper.page_color.name(),
                "watermark": paper.watermark,
                "watermark_size": paper.watermark_size,
                "watermark_color": paper.watermark_color.name(),
                "watermark_opacity": paper.watermark_opacity,
                "border": paper.show_border,
            },
            "header": {"text": paper.header_text, "show": paper.show_header},
            "footer": {"text": paper.footer_text, "show": paper.show_footer},
            "page_numbers": {
                "show": paper.show_page_numbers,
                "format": paper.page_number_format,
            },
            "bookmarks": dict(tab.bookmarks),
            "citations": list(tab.citations),
            "comments": {str(key): value for key, value in tab.comments.items()},
            "comment_ranges": [
                [comment_id, start, end]
                for comment_id in tab.comments
                for start, end, _ in self._comment_ranges(tab, comment_id)
            ],
            "toc_blocks": [
                block.position()
                for block in self._document_blocks(tab.editor.document())
                if bool(block.blockFormat().property(TOC_PROP))
            ],
        }

    def _document_blocks(self, document):
        block = document.begin()
        while block.isValid():
            yield block
            block = block.next()

    def _load_serialized(self, tab, data):
        tab.editor.setHtml(data.get("html", ""))
        page = data.get("page", {})
        paper = tab.paper
        paper.page_color = QColor(page.get("color", "#FFFFFF"))
        paper.watermark = page.get("watermark", "")
        paper.watermark_size = int(page.get("watermark_size", 60))
        paper.watermark_color = QColor(page.get("watermark_color", "#B4B4B4"))
        paper.watermark_opacity = int(page.get("watermark_opacity", 18))
        paper.show_border = bool(page.get("border", False))
        header = data.get("header", {})
        paper.header_text = header.get("text", "")
        paper.show_header = bool(header.get("show", False))
        footer = data.get("footer", {})
        paper.footer_text = footer.get("text", "")
        paper.show_footer = bool(footer.get("show", False))
        numbers = data.get("page_numbers", {})
        paper.show_page_numbers = bool(numbers.get("show", True))
        paper.page_number_format = numbers.get("format", "Page X of Y")
        paper.set_page(
            page.get("width", 816),
            page.get("height", 1056),
            page.get("margin", 96),
        )
        tab.bookmarks = {str(k): int(v) for k, v in data.get("bookmarks", {}).items()}
        tab.citations = list(data.get("citations", []))
        tab.comments = {}
        for key, value in data.get("comments", {}).items():
            try:
                tab.comments[int(key)] = value
            except (TypeError, ValueError):
                continue
        document = tab.editor.document()
        for entry in data.get("comment_ranges", []):
            try:
                comment_id, start, end = int(entry[0]), int(entry[1]), int(entry[2])
            except (TypeError, ValueError, IndexError):
                continue
            if comment_id not in tab.comments:
                continue
            cursor = QTextCursor(document)
            cursor.setPosition(start)
            cursor.setPosition(end, QTextCursor.MoveMode.KeepAnchor)
            char_format = QTextCharFormat()
            char_format.setProperty(COMMENT_PROP, comment_id)
            char_format.setToolTip(f"Comment: {tab.comments[comment_id].get('text', '')}")
            char_format.setBackground(QColor("#FFF2CC"))
            char_format.setUnderlineStyle(QTextCharFormat.UnderlineStyle.DotLine)
            char_format.setUnderlineColor(QColor("#D35400"))
            cursor.mergeCharFormat(char_format)
        for position in data.get("toc_blocks", []):
            try:
                block = document.findBlock(int(position))
            except (TypeError, ValueError):
                continue
            if block.isValid():
                block_cursor = QTextCursor(block)
                block_format = block_cursor.blockFormat()
                block_format.setProperty(TOC_PROP, True)
                block_cursor.setBlockFormat(block_format)
        tab.last_text = tab.editor.toPlainText()
        tab.editor.document().setModified(False)
        tab.refresh_ruler()
        self._refresh_comments()
        self._refresh_navigation()

    def _write_pdf_cache(self, tab, path):
        try:
            data = self._serialize_tab(tab)
            data["title"] = os.path.basename(path)
            cache_path = _pdf_cache_path(path)
            os.makedirs(os.path.dirname(cache_path), exist_ok=True)
            with open(cache_path, "w", encoding="utf-8") as handle:
                json.dump(data, handle, ensure_ascii=False, indent=1)
        except Exception:
            pass

    def _load_pdf_cache(self, path):
        try:
            cache_path = _pdf_cache_path(path)
            if not os.path.exists(cache_path):
                return None
            with open(cache_path, "r", encoding="utf-8") as handle:
                return json.load(handle)
        except Exception:
            return None

    def _load_pdf_into_tab(self, tab, path):
        cached = self._load_pdf_cache(path)
        if cached:
            self._load_serialized(tab, cached)
            return "cache"
        try:
            pages, page_size = extract_pdf_pages(path)
        except Exception as error:
            QMessageBox.critical(self, "Open PDF", f"Could not read this PDF:\n{error}")
            return None
        if not any(page.strip() for page in pages):
            QMessageBox.warning(
                self,
                "Open PDF",
                "This PDF does not contain selectable text.\n\n"
                "It may be a scanned or image-only PDF. SimzWord can only import "
                "PDF text, so this file cannot be edited as text.",
            )
            return None
        tab.editor.clear()
        cursor = tab.editor.textCursor()
        for index, page in enumerate(pages):
            if index:
                cursor.insertBlock()
                block_format = cursor.blockFormat()
                block_format.setPageBreakPolicy(
                    QTextFormat.PageBreakFlag.PageBreak_AlwaysBefore
                )
                cursor.setBlockFormat(block_format)
            text = page.strip("\n")
            if text:
                cursor.insertText(text)
        tab.editor.setTextCursor(cursor)
        if page_size is not None and page_size.width() > 0 and page_size.height() > 0:
            width = max(200, int(round(page_size.width() * 96.0 / 72.0)))
            height = max(200, int(round(page_size.height() * 96.0 / 72.0)))
            tab.paper.set_page(width, height, tab.paper.base_margin)
        else:
            tab.paper.update_size()
        tab.editor.document().setModified(False)
        return "text"

    def _write_pdf_to_path(self, tab, path):
        printer = QPrinter(QPrinter.PrinterMode.HighResolution)
        printer.setOutputFormat(QPrinter.OutputFormat.PdfFormat)
        printer.setOutputFileName(path)
        self._render_document_to_printer(tab, printer)
        self._write_pdf_cache(tab, path)

    def _open_file(self, checked=False):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Open Document",
            "",
            "All Supported (*.docx *.pdf *.simz *.html *.htm *.md *.markdown *.txt);;"
            "Word Documents (*.docx);;PDF Documents (*.pdf);;"
            "SimzWord Projects (*.simz);;All Files (*)",
        )
        if path:
            self._open_path(path)

    def _open_path(self, path):
        if not path or not os.path.exists(path):
            QMessageBox.warning(self, "Open Document", f"File not found:\n{path}")
            self._remove_recent(path)
            return
        extension = os.path.splitext(path)[1].lower()
        if extension == ".doc":
            QMessageBox.warning(
                self,
                "Open Document",
                "Legacy .doc files are not supported.\n\nOpen a .docx file instead.",
            )
            return
        tab = self._current_tab()
        if (
            tab is None
            or tab.file_path
            or tab.editor.toPlainText().strip()
            or tab.editor.document().isModified()
        ):
            tab = self._new_document()
        try:
            if extension == ".simz":
                with open(path, "r", encoding="utf-8") as handle:
                    data = json.load(handle)
                self._load_serialized(tab, data)
            elif extension == ".docx":
                load_docx(tab, path)
            elif extension == ".pdf":
                pdf_mode = self._load_pdf_into_tab(tab, path)
                if not pdf_mode:
                    return
            else:
                with open(path, "r", encoding="utf-8", errors="ignore") as handle:
                    content = handle.read()
                if extension in (".html", ".htm"):
                    tab.editor.setHtml(content)
                elif extension in (".md", ".markdown"):
                    tab.editor.document().setMarkdown(content)
                else:
                    tab.editor.setPlainText(content)
                tab.editor.document().setModified(False)
            tab.untitled_name = os.path.basename(path)
            tab.file_path = path
            tab.last_text = tab.editor.toPlainText()
            self._update_tab_title(tab)
            self._add_recent(path)
            self._refresh_navigation()
            self._refresh_comments()
            self._update_statusbar()
            if extension == ".pdf":
                if pdf_mode == "cache":
                    self.statusBar().showMessage(
                        f"Opened {os.path.basename(path)} as an editable SimzWord PDF project",
                        5000,
                    )
                else:
                    self.statusBar().showMessage(
                        "PDF imported as editable text; original layout was not preserved",
                        5000,
                    )
            else:
                self.statusBar().showMessage(f"Opened {os.path.basename(path)}", 3000)
        except Exception as error:
            QMessageBox.critical(self, "Open Error", f"Could not open this file:\n{error}")

    def _save_tab(self, tab, save_as=False):
        if tab is None:
            return False
        if save_as or not tab.file_path:
            suggested = tab.display_name()
            if not os.path.splitext(suggested)[1]:
                suggested += ".docx"
            path, selected_filter = QFileDialog.getSaveFileName(
                self,
                "Save Document",
                suggested,
                "Word Document (*.docx);;PDF Document (*.pdf);;"
                "SimzWord Project (*.simz);;HTML Document (*.html);;"
                "Markdown Document (*.md);;Plain Text (*.txt);;All Files (*)",
            )
            if not path:
                return False
            if not os.path.splitext(path)[1]:
                if "PDF" in selected_filter:
                    path += ".pdf"
                elif "SimzWord" in selected_filter:
                    path += ".simz"
                elif "Word" in selected_filter:
                    path += ".docx"
                elif "HTML" in selected_filter:
                    path += ".html"
                elif "Markdown" in selected_filter:
                    path += ".md"
                elif "Plain Text" in selected_filter:
                    path += ".txt"
                else:
                    path += ".docx"
        else:
            path = tab.file_path
        extension = os.path.splitext(path)[1].lower()
        try:
            if extension == ".simz":
                data = self._serialize_tab(tab)
                data["title"] = os.path.basename(path)
                with open(path, "w", encoding="utf-8") as handle:
                    json.dump(data, handle, ensure_ascii=False, indent=1)
            elif extension == ".docx":
                save_docx(tab, path)
            elif extension == ".pdf":
                self._write_pdf_to_path(tab, path)
            elif extension in (".html", ".htm"):
                with open(path, "w", encoding="utf-8") as handle:
                    handle.write(tab.editor.toHtml())
            elif extension in (".md", ".markdown"):
                with open(path, "w", encoding="utf-8") as handle:
                    handle.write(tab.editor.document().toMarkdown())
            elif extension == ".doc":
                raise ValueError(
                    "Legacy .doc files are not supported. Save as .docx instead."
                )
            else:
                with open(path, "w", encoding="utf-8") as handle:
                    handle.write(tab.editor.toPlainText())
            tab.file_path = path
            tab.untitled_name = os.path.basename(path)
            tab.editor.document().setModified(False)
            self._clear_recovery(tab)
            self._update_tab_title(tab)
            self._add_recent(path)
            self.statusBar().showMessage(f"Saved {os.path.basename(path)}", 3000)
            return True
        except Exception as error:
            QMessageBox.critical(self, "Save Error", f"Could not save this file:\n{error}")
            return False

    def _save_file(self, checked=False):
        self._save_tab(self._current_tab())

    def _save_file_as(self, checked=False):
        self._save_tab(self._current_tab(), save_as=True)

    def _save_all(self, checked=False):
        for index in range(self.tab_widget.count()):
            tab = self.tab_widget.widget(index)
            if not isinstance(tab, DocumentTab):
                continue
            if tab.editor.document().isModified() or not tab.file_path:
                if not self._save_tab(tab):
                    return False
        self.statusBar().showMessage("All documents saved", 3000)
        return True

    def _maybe_save_tab(self, tab):
        if tab is None or not tab.editor.document().isModified():
            return True
        result = QMessageBox.question(
            self,
            "Unsaved Changes",
            f"Save changes to '{tab.display_name()}'?",
            QMessageBox.StandardButton.Save
            | QMessageBox.StandardButton.Discard
            | QMessageBox.StandardButton.Cancel,
            QMessageBox.StandardButton.Save,
        )
        if result == QMessageBox.StandardButton.Save:
            return self._save_tab(tab)
        if result == QMessageBox.StandardButton.Cancel:
            return False
        return True

    def _close_tab(self, index):
        tab = self.tab_widget.widget(index)
        if not isinstance(tab, DocumentTab):
            return
        if not self._maybe_save_tab(tab):
            return
        self._clear_recovery(tab)
        self.tab_widget.removeTab(index)
        tab.deleteLater()
        if self.tab_widget.count() == 0:
            self._new_document()

    def _close_current_tab(self, checked=False):
        index = self.tab_widget.currentIndex()
        if index >= 0:
            self._close_tab(index)

    # ------------------------------------------------------------------
    # Export and printing
    # ------------------------------------------------------------------
    def _export_docx(self, checked=False):
        tab = self._current_tab()
        if tab is None:
            return
        suggested = os.path.splitext(tab.display_name())[0] + ".docx"
        path, _ = QFileDialog.getSaveFileName(
            self, "Export as Word Document", suggested, "Word Document (*.docx)"
        )
        if not path:
            return
        if not path.lower().endswith(".docx"):
            path += ".docx"
        try:
            save_docx(tab, path)
            self.statusBar().showMessage(f"Exported {os.path.basename(path)}", 3000)
        except Exception as error:
            QMessageBox.critical(
                self, "Export Error", f"Could not export Word document:\n{error}"
            )

    def _export_pdf(self, checked=False):
        tab = self._current_tab()
        if tab is None:
            return
        suggested = os.path.splitext(tab.display_name())[0] + ".pdf"
        path, _ = QFileDialog.getSaveFileName(self, "Export as PDF", suggested, "PDF Files (*.pdf)")
        if not path:
            return
        if not path.lower().endswith(".pdf"):
            path += ".pdf"
        try:
            self._write_pdf_to_path(tab, path)
            self.statusBar().showMessage(f"Exported {os.path.basename(path)}", 3000)
        except Exception as error:
            QMessageBox.critical(self, "Export Error", f"Could not export PDF:\n{error}")

    def _export_html(self, checked=False):
        tab = self._current_tab()
        if tab is None:
            return
        suggested = os.path.splitext(tab.display_name())[0] + ".html"
        path, _ = QFileDialog.getSaveFileName(self, "Export as HTML", suggested, "HTML Files (*.html)")
        if not path:
            return
        if not path.lower().endswith((".html", ".htm")):
            path += ".html"
        try:
            with open(path, "w", encoding="utf-8") as handle:
                handle.write(tab.editor.toHtml())
            self.statusBar().showMessage(f"Exported {os.path.basename(path)}", 3000)
        except Exception as error:
            QMessageBox.critical(self, "Export Error", f"Could not export HTML:\n{error}")

    def _export_md(self, checked=False):
        tab = self._current_tab()
        if tab is None:
            return
        suggested = os.path.splitext(tab.display_name())[0] + ".md"
        path, _ = QFileDialog.getSaveFileName(self, "Export as Markdown", suggested, "Markdown Files (*.md)")
        if not path:
            return
        if not path.lower().endswith((".md", ".markdown")):
            path += ".md"
        try:
            with open(path, "w", encoding="utf-8") as handle:
                handle.write(tab.editor.document().toMarkdown())
            self.statusBar().showMessage(f"Exported {os.path.basename(path)}", 3000)
        except Exception as error:
            QMessageBox.critical(self, "Export Error", f"Could not export Markdown:\n{error}")

    def _export_txt(self, checked=False):
        tab = self._current_tab()
        if tab is None:
            return
        suggested = os.path.splitext(tab.display_name())[0] + ".txt"
        path, _ = QFileDialog.getSaveFileName(self, "Export as Plain Text", suggested, "Text Files (*.txt)")
        if not path:
            return
        if not path.lower().endswith(".txt"):
            path += ".txt"
        try:
            with open(path, "w", encoding="utf-8") as handle:
                handle.write(tab.editor.toPlainText())
            self.statusBar().showMessage(f"Exported {os.path.basename(path)}", 3000)
        except Exception as error:
            QMessageBox.critical(self, "Export Error", f"Could not export text:\n{error}")

    def _configure_printer(self, tab, printer):
        paper = tab.paper
        printer.setPageSize(
            QPageSize(
                QSizeF(paper.base_page_w / 96.0, paper.base_page_h / 96.0),
                QPageSize.Unit.Inch,
                "SimzWord",
                QPageSize.SizeMatchPolicy.ExactMatch,
            )
        )
        printer.setPageMargins(QMarginsF(0, 0, 0, 0), QPageLayout.Unit.Inch)
        orientation = (
            QPageLayout.Orientation.Landscape
            if paper.base_page_w > paper.base_page_h
            else QPageLayout.Orientation.Portrait
        )
        printer.setPageOrientation(orientation)

    def _render_document_to_printer(self, tab, printer):
        self._configure_printer(tab, printer)
        paper = tab.paper
        document = tab.editor.document().clone()
        document.setDocumentMargin(0)
        content_width = max(100.0, paper.page_w - paper.margin * 2)
        content_height = max(100.0, paper.page_h - paper.margin * 2)
        document.setPageSize(QSizeF(content_width, content_height))
        page_count = max(1, document.pageCount())
        scale = printer.resolution() / 96.0
        full = printer.pageLayout().fullRectPixels(printer.resolution())
        painter = QPainter(printer)
        try:
            painter.setRenderHint(QPainter.RenderHint.Antialiasing)
            painter.setRenderHint(QPainter.RenderHint.TextAntialiasing)
            for page in range(page_count):
                if page > 0 and not printer.newPage():
                    break
                painter.save()
                painter.translate(full.left(), full.top())
                painter.scale(scale, scale)
                painter.fillRect(QRectF(0, 0, paper.page_w, paper.page_h), paper.page_color)
                if paper.show_border:
                    painter.setPen(QPen(QColor("#9AA7B4"), 2))
                    painter.setBrush(Qt.BrushStyle.NoBrush)
                    painter.drawRect(QRectF(12, 12, paper.page_w - 24, paper.page_h - 24))
                if paper.watermark:
                    painter.save()
                    painter.setOpacity(paper.watermark_opacity / 100.0)
                    painter.translate(paper.page_w / 2, paper.page_h / 2)
                    painter.rotate(-35)
                    watermark_font = QFont("Segoe UI", paper.watermark_size, QFont.Weight.Bold)
                    painter.setFont(watermark_font)
                    painter.setPen(paper.watermark_color)
                    metrics = QFontMetrics(watermark_font)
                    width = metrics.horizontalAdvance(paper.watermark)
                    painter.drawText(
                        QRectF(-width / 2, -metrics.height() / 2, width, metrics.height()),
                        Qt.AlignmentFlag.AlignCenter,
                        paper.watermark,
                    )
                    painter.restore()
                painter.save()
                painter.setClipRect(
                    QRectF(paper.margin, paper.margin, content_width, content_height)
                )
                painter.translate(paper.margin, paper.margin)
                document.drawContents(
                    painter,
                    QRectF(0, page * content_height, content_width, content_height),
                )
                painter.restore()
                if paper.show_header and paper.header_text:
                    painter.setFont(QFont("Segoe UI", 9))
                    painter.setPen(QColor("#666666"))
                    painter.drawText(
                        QRectF(
                            paper.margin,
                            paper.margin * 0.30,
                            paper.page_w - paper.margin * 2,
                            paper.margin * 0.45,
                        ),
                        Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter,
                        paper.header_text,
                    )
                    painter.setPen(QColor("#D0D0D0"))
                    painter.drawLine(
                        int(paper.margin),
                        int(paper.margin * 0.82),
                        int(paper.page_w - paper.margin),
                        int(paper.margin * 0.82),
                    )
                if paper.show_footer and paper.footer_text:
                    painter.setFont(QFont("Segoe UI", 9))
                    painter.setPen(QColor("#666666"))
                    painter.drawText(
                        QRectF(
                            paper.margin,
                            paper.page_h - paper.margin * 0.75,
                            paper.page_w - paper.margin * 2,
                            paper.margin * 0.45,
                        ),
                        Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter,
                        paper.footer_text,
                    )
                if paper.show_page_numbers:
                    text = paper.page_number_format.replace("X", str(page + 1)).replace(
                        "Y", str(page_count)
                    )
                    painter.setFont(QFont("Segoe UI", 9))
                    painter.setPen(QColor("#999999"))
                    painter.drawText(
                        QRectF(
                            0,
                            paper.page_h - paper.margin * 0.62,
                            paper.page_w,
                            paper.margin * 0.4,
                        ),
                        Qt.AlignmentFlag.AlignCenter,
                        text,
                    )
                painter.restore()
        finally:
            painter.end()

    def _print(self, checked=False):
        tab = self._current_tab()
        if tab is None:
            return
        printer = QPrinter(QPrinter.PrinterMode.HighResolution)
        self._configure_printer(tab, printer)
        dialog = QPrintDialog(printer, self)
        dialog.setWindowTitle("Print Document")
        if dialog.exec() == QDialog.DialogCode.Accepted:
            try:
                self._render_document_to_printer(tab, printer)
            except Exception as error:
                QMessageBox.critical(self, "Print Error", f"Could not print the document:\n{error}")

    def _print_preview(self, checked=False):
        tab = self._current_tab()
        if tab is None:
            return
        printer = QPrinter(QPrinter.PrinterMode.HighResolution)
        self._configure_printer(tab, printer)
        preview = QPrintPreviewDialog(printer, self)
        preview.setWindowTitle("Print Preview")
        preview.paintRequested.connect(lambda target: self._render_document_to_printer(tab, target))
        preview.exec()

    def _page_setup(self, checked=False):
        tab = self._current_tab()
        if tab is None:
            return
        dialog = PageSetupDialog(
            self, tab.paper.base_page_w, tab.paper.base_page_h, tab.paper.base_margin
        )
        if dialog.exec() == QDialog.DialogCode.Accepted:
            data = dialog.result_data()
            tab.paper.set_page(data["width"], data["height"], data["margin"])
            tab.refresh_ruler()
            self._update_statusbar()

    def _show_properties(self, checked=False):
        tab = self._current_tab()
        if tab is None:
            return
        text = tab.editor.toPlainText()
        words = len(re.findall(r"\S+", text))
        path = tab.file_path or "Not saved yet"
        created = "Unknown"
        modified = "Unknown"
        if tab.file_path and os.path.exists(tab.file_path):
            try:
                created = datetime.fromtimestamp(os.path.getctime(tab.file_path)).strftime(
                    "%Y-%m-%d %H:%M"
                )
                modified = datetime.fromtimestamp(os.path.getmtime(tab.file_path)).strftime(
                    "%Y-%m-%d %H:%M"
                )
            except Exception:
                pass
        info = (
            f"Name: {tab.display_name()}\n"
            f"Location: {path}\n"
            f"Pages: {tab.paper.pages}\n"
            f"Words: {words}\n"
            f"Characters: {len(text)}\n"
            f"Created: {created}\n"
            f"Modified: {modified}\n"
            f"Application: {APP_NAME} {APP_VERSION}"
        )
        QMessageBox.information(self, "Document Properties", info)

    # ------------------------------------------------------------------
    # Recent files, autosave and recovery
    # ------------------------------------------------------------------
    def _recent_files(self):
        recent = self.settings.value("recent_files", [])
        if isinstance(recent, str):
            recent = [recent]
        if not isinstance(recent, list):
            return []
        return [item for item in recent if isinstance(item, str) and item]

    def _add_recent(self, path):
        if not path:
            return
        path = os.path.abspath(path)
        normalized = os.path.normcase(path)
        recent = [
            item
            for item in self._recent_files()
            if os.path.normcase(os.path.abspath(item)) != normalized
        ]
        recent.insert(0, path)
        self.settings.setValue("recent_files", recent[:12])

    def _remove_recent(self, path):
        if not path:
            return
        normalized = os.path.normcase(os.path.abspath(path))
        recent = [
            item
            for item in self._recent_files()
            if os.path.normcase(os.path.abspath(item)) != normalized
        ]
        self.settings.setValue("recent_files", recent)

    def _refresh_recent_menu(self):
        self.recent_menu.clear()
        recent = self._recent_files()
        if not recent:
            empty = self.recent_menu.addAction("No recent documents")
            empty.setEnabled(False)
            return
        for path in recent:
            action = self.recent_menu.addAction(os.path.basename(path))
            action.setToolTip(path)
            action.setStatusTip(path)
            action.triggered.connect(lambda checked=False, p=path: self._open_path(p))
        self.recent_menu.addSeparator()
        clear = self.recent_menu.addAction("Clear Recent Documents")
        clear.triggered.connect(
            lambda: (self.settings.setValue("recent_files", []), self._refresh_recent_menu())
        )

    def _clear_recovery(self, tab):
        if tab is None or not tab.recovery_path:
            return
        try:
            os.remove(tab.recovery_path)
        except Exception:
            pass
        tab.recovery_path = None

    def _clear_recovery_dir(self):
        try:
            for name in os.listdir(self._recovery_dir):
                if name.endswith(".simz"):
                    try:
                        os.remove(os.path.join(self._recovery_dir, name))
                    except Exception:
                        pass
        except Exception:
            pass

    def _autosave(self):
        enabled = self.settings.value("autosave/enabled", True, bool)
        if not enabled:
            self.lbl_autosave.setText("Autosave off")
            return
        for index in range(self.tab_widget.count()):
            tab = self.tab_widget.widget(index)
            if not isinstance(tab, DocumentTab):
                continue
            if not tab.editor.document().isModified():
                continue
            if not tab.editor.toPlainText().strip() and not tab.file_path:
                continue
            try:
                if not tab.recovery_path:
                    safe_name = re.sub(r"[^A-Za-z0-9_.-]+", "_", tab.display_name())[:40]
                    safe_name = safe_name or "document"
                    tab.recovery_path = os.path.join(
                        self._recovery_dir, f"{safe_name}_{abs(hash(tab))}.simz"
                    )
                data = self._serialize_tab(tab)
                data["recovery_title"] = tab.display_name()
                data["recovery_time"] = datetime.now().isoformat(timespec="seconds")
                with open(tab.recovery_path, "w", encoding="utf-8") as handle:
                    json.dump(data, handle, ensure_ascii=False)
            except Exception:
                pass
        self.lbl_autosave.setText(f"Autosaved {datetime.now().strftime('%H:%M')}")

    def _check_recovery(self):
        if not os.path.isdir(self._recovery_dir):
            return
        files = [name for name in os.listdir(self._recovery_dir) if name.endswith(".simz")]
        if not files:
            return
        result = QMessageBox.question(
            self,
            "Recover Documents",
            f"{APP_NAME} found {len(files)} autosaved document(s) from a previous session.\n"
            "Would you like to recover them now?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes,
        )
        if result == QMessageBox.StandardButton.Yes:
            for name in files:
                path = os.path.join(self._recovery_dir, name)
                try:
                    with open(path, "r", encoding="utf-8") as handle:
                        data = json.load(handle)
                    tab = self._new_document()
                    self._load_serialized(tab, data)
                    title = data.get("recovery_title") or data.get("title") or "Recovered Document"
                    tab.untitled_name = f"Recovered {title}"
                    tab.file_path = None
                    tab.editor.document().setModified(True)
                    self._update_tab_title(tab)
                except Exception:
                    pass
                try:
                    os.remove(path)
                except Exception:
                    pass
            self.statusBar().showMessage("Recovered autosaved documents", 5000)
        else:
            self._clear_recovery_dir()

    def closeEvent(self, event):
        for index in range(self.tab_widget.count()):
            tab = self.tab_widget.widget(index)
            if isinstance(tab, DocumentTab) and not self._maybe_save_tab(tab):
                event.ignore()
                return
        self._clear_recovery_dir()
        self.settings.setValue("ui/dark", self.dark)
        self.settings.setValue("ui/ruler", self.act_toggle_ruler.isChecked())
        self.settings.setValue("ui/statusbar", self.act_toggle_status.isChecked())
        self.settings.setValue("ui/navigation", self.act_toggle_nav.isChecked())
        self.settings.setValue("editor/font_family", self.default_font_family)
        self.settings.setValue("editor/font_size", self.default_font_size)
        self.settings.setValue("spelling/custom", sorted(self.custom_words))
        self._save_window_state()
        event.accept()

    # ------------------------------------------------------------------
    # Formatting
    # ------------------------------------------------------------------
    def _merge_fmt(self, prop, value):
        if self.editor is None:
            return
        fmt = QTextCharFormat()
        if prop == "family":
            fmt.setFontFamilies(list(value))
        elif prop == "size":
            fmt.setFontPointSize(float(value))
        elif prop == "weight":
            fmt.setFontWeight(value)
        elif prop == "italic":
            fmt.setFontItalic(bool(value))
        elif prop == "underline":
            fmt.setFontUnderline(bool(value))
        elif prop == "strike":
            fmt.setFontStrikeOut(bool(value))
        elif prop == "v-align":
            fmt.setVerticalAlignment(value)
        elif prop == "color":
            fmt.setForeground(value)
        elif prop == "background":
            fmt.setBackground(value)
        cursor = self.editor.textCursor()
        if cursor.hasSelection():
            cursor.mergeCharFormat(fmt)
        self.editor.mergeCurrentCharFormat(fmt)
        self.editor.setFocus()

    def _on_size_changed(self, text):
        try:
            size = float(text)
        except (TypeError, ValueError):
            return
        if size > 0:
            self._merge_fmt("size", size)

    def _change_font_size(self, delta):
        if self.editor is None:
            return
        fmt = self.editor.textCursor().charFormat()
        size = fmt.fontPointSize() or self.editor.fontPointSize() or self.default_font_size
        size = max(1, size + delta)
        self._merge_fmt("size", size)
        with QSignalBlocker(self.size_combo):
            self.size_combo.setCurrentText(str(int(round(size))))

    def _choose_text_color(self, checked=False):
        if self.editor is None:
            return
        color = QColorDialog.getColor(
            self.editor.textCursor().charFormat().foreground().color(),
            self,
            "Text Color",
        )
        if color.isValid():
            self.act_text_color.setIcon(make_color_icon(color, "A"))
            self._merge_fmt("color", color)

    def _choose_highlight_color(self, checked=False):
        if self.editor is None:
            return
        color = QColorDialog.getColor(QColor("#FFFF00"), self, "Highlight Color")
        if color.isValid():
            self.act_highlight.setIcon(make_color_icon(color, "H"))
            self._merge_fmt("background", color)

    def _clear_formatting(self, checked=False):
        if self.editor is None:
            return
        fmt = QTextCharFormat()
        fmt.setFontFamilies([self.default_font_family])
        fmt.setFontPointSize(self.default_font_size)
        fmt.setFontWeight(QFont.Weight.Normal)
        fmt.setFontItalic(False)
        fmt.setFontUnderline(False)
        fmt.setFontStrikeOut(False)
        fmt.setVerticalAlignment(QTextCharFormat.VerticalAlignment.AlignNormal)
        fmt.setForeground(QColor("#111111"))
        fmt.setBackground(QColor(0, 0, 0, 0))
        cursor = self.editor.textCursor()
        if cursor.hasSelection():
            cursor.setCharFormat(fmt)
        self.editor.setCurrentCharFormat(fmt)
        self.editor.setFocus()

    def _toggle_format_painter(self, checked):
        if checked:
            if self.editor is None:
                self.act_format_painter.setChecked(False)
                return
            self._format_painter_fmt = self.editor.currentCharFormat()
            self.statusBar().showMessage("Format Painter: select text to apply the format", 5000)
        else:
            self._format_painter_fmt = None

    def _cycle_case(self, checked=False):
        cursor = self.editor.textCursor() if self.editor else None
        if cursor is None or not cursor.hasSelection():
            return
        text = cursor.selectedText()
        if text.isupper():
            new_text = text.lower()
        elif text.islower():
            new_text = text.title()
        else:
            new_text = text.upper()
        cursor.insertText(new_text)

    def _change_case(self, mode):
        if self.editor is None:
            return
        cursor = self.editor.textCursor()
        if not cursor.hasSelection():
            cursor.select(QTextCursor.SelectionType.WordUnderCursor)
        text = cursor.selectedText()
        if not text:
            return
        if mode == "upper":
            new_text = text.upper()
        elif mode == "lower":
            new_text = text.lower()
        elif mode == "title":
            new_text = text.title()
        elif mode == "sentence":
            new_text = text[:1].upper() + text[1:].lower()
        else:
            new_text = text.swapcase()
        cursor.insertText(new_text)

    def _selected_blocks(self, cursor):
        document = cursor.document()
        start = cursor.selectionStart()
        end = cursor.selectionEnd()
        block = document.findBlock(start)
        blocks = []
        while block.isValid():
            blocks.append(block)
            if block.position() + max(1, block.length() - 1) >= end:
                break
            block = block.next()
        if not blocks:
            blocks.append(cursor.block())
        return blocks

    def _apply_list(self, style):
        if self.editor is None:
            return
        cursor = self.editor.textCursor()
        current = cursor.currentList()
        cursor.beginEditBlock()
        if current is not None and current.format().style() == style:
            for block in self._selected_blocks(cursor):
                text_list = block.textList()
                if text_list is not None:
                    text_list.remove(block)
            cursor.endEditBlock()
            return
        fmt = QTextListFormat()
        fmt.setStyle(style)
        fmt.setIndent(1)
        cursor.createList(fmt)
        cursor.endEditBlock()

    def _remove_list(self, checked=False):
        if self.editor is None:
            return
        cursor = self.editor.textCursor()
        cursor.beginEditBlock()
        for block in self._selected_blocks(cursor):
            text_list = block.textList()
            if text_list is not None:
                text_list.remove(block)
            block_cursor = QTextCursor(block)
            block_format = block_cursor.blockFormat()
            block_format.setIndent(0)
            block_cursor.setBlockFormat(block_format)
        cursor.endEditBlock()

    def _change_indent(self, delta):
        if self.editor is None:
            return
        cursor = self.editor.textCursor()
        cursor.beginEditBlock()
        for block in self._selected_blocks(cursor):
            block_cursor = QTextCursor(block)
            block_format = block_cursor.blockFormat()
            block_format.setIndent(max(0, min(10, block_format.indent() + delta)))
            block_cursor.setBlockFormat(block_format)
        cursor.endEditBlock()

    def _sort_selection(self, ascending=True):
        if self.editor is None:
            return
        cursor = self.editor.textCursor()
        if not cursor.hasSelection():
            cursor.select(QTextCursor.SelectionType.BlockUnderCursor)
        text = cursor.selectedText()
        if not text:
            return
        parts = text.split("\u2029")
        trailing = bool(parts and parts[-1] == "")
        if trailing:
            parts = parts[:-1]
        parts.sort(key=str.lower, reverse=not ascending)
        new_text = "\u2029".join(parts)
        if trailing:
            new_text += "\u2029"
        cursor.insertText(new_text)

    def _choose_shading(self, checked=False):
        if self.editor is None:
            return
        color = QColorDialog.getColor(QColor("#FFF2CC"), self, "Paragraph Shading")
        if not color.isValid():
            return
        cursor = self.editor.textCursor()
        cursor.beginEditBlock()
        for block in self._selected_blocks(cursor):
            block_cursor = QTextCursor(block)
            block_format = block_cursor.blockFormat()
            block_format.setBackground(color)
            block_cursor.setBlockFormat(block_format)
        cursor.endEditBlock()

    def _apply_line_spacing(self, value):
        if self.editor is None:
            return
        cursor = self.editor.textCursor()
        cursor.beginEditBlock()
        for block in self._selected_blocks(cursor):
            block_cursor = QTextCursor(block)
            block_format = block_cursor.blockFormat()
            block_format.setLineHeight(
                float(value), QTextBlockFormat.LineHeightTypes.ProportionalHeight.value
            )
            block_cursor.setBlockFormat(block_format)
        cursor.endEditBlock()

    def _set_paragraph_spacing(self, before, after, line):
        if self.editor is None:
            return
        cursor = self.editor.textCursor()
        cursor.beginEditBlock()
        for block in self._selected_blocks(cursor):
            block_cursor = QTextCursor(block)
            block_format = block_cursor.blockFormat()
            block_format.setTopMargin(float(before))
            block_format.setBottomMargin(float(after))
            block_format.setLineHeight(
                float(line), QTextBlockFormat.LineHeightTypes.ProportionalHeight.value
            )
            block_cursor.setBlockFormat(block_format)
        cursor.endEditBlock()

    def _font_dialog(self, checked=False):
        if self.editor is None:
            return
        font, accepted = QFontDialog.getFont(self.editor.currentFont(), self, "Font")
        if not accepted:
            return
        fmt = QTextCharFormat()
        fmt.setFontFamilies([font.family()])
        fmt.setFontPointSize(font.pointSize() if font.pointSize() > 0 else self.default_font_size)
        fmt.setFontWeight(font.weight())
        fmt.setFontItalic(font.italic())
        fmt.setFontUnderline(font.underline())
        fmt.setFontStrikeOut(font.strikeOut())
        cursor = self.editor.textCursor()
        if cursor.hasSelection():
            cursor.mergeCharFormat(fmt)
        self.editor.mergeCurrentCharFormat(fmt)
        self.editor.setFocus()

    def _paragraph_dialog(self, checked=False):
        if self.editor is None:
            return
        cursor = self.editor.textCursor()
        block_format = cursor.blockFormat()
        alignment = self.editor.alignment()
        if alignment & Qt.AlignmentFlag.AlignHCenter:
            alignment_name = "center"
        elif alignment & Qt.AlignmentFlag.AlignRight:
            alignment_name = "right"
        elif alignment & Qt.AlignmentFlag.AlignJustify:
            alignment_name = "justify"
        else:
            alignment_name = "left"
        line_height = 100
        if block_format.lineHeightType() == QTextBlockFormat.LineHeightTypes.ProportionalHeight.value:
            line_height = int(block_format.lineHeight())
        dialog = ParagraphDialog(
            self,
            alignment_name,
            int(block_format.leftIndent()),
            int(block_format.firstLineIndent()),
            block_format.topMargin(),
            block_format.bottomMargin(),
            line_height,
        )
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        data = dialog.result_data()
        align_map = {
            "left": Qt.AlignmentFlag.AlignLeft,
            "center": Qt.AlignmentFlag.AlignHCenter,
            "right": Qt.AlignmentFlag.AlignRight,
            "justify": Qt.AlignmentFlag.AlignJustify,
        }
        cursor.beginEditBlock()
        for block in self._selected_blocks(cursor):
            block_cursor = QTextCursor(block)
            fmt = block_cursor.blockFormat()
            fmt.setAlignment(align_map.get(data["alignment"], Qt.AlignmentFlag.AlignLeft))
            fmt.setLeftIndent(float(data["left_indent"]))
            fmt.setFirstLineIndent(float(data["first_indent"]))
            fmt.setTopMargin(float(data["before"]))
            fmt.setBottomMargin(float(data["after"]))
            fmt.setLineHeight(
                float(data["line_height"]), QTextBlockFormat.LineHeightTypes.ProportionalHeight.value
            )
            block_cursor.setBlockFormat(fmt)
        cursor.endEditBlock()
        self._update_format_ui()

    def _set_alignment(self, align):
        if self.editor is None:
            return
        self.editor.setAlignment(align)
        self.editor.setFocus()
        self._update_format_ui()

    def _apply_style(self, name):
        if self.editor is None or name not in STYLE_DEFS:
            return
        definition = STYLE_DEFS[name]
        cursor = self.editor.textCursor()
        cursor.beginEditBlock()
        for block in self._selected_blocks(cursor):
            block_cursor = QTextCursor(block)
            block_format = block_cursor.blockFormat()
            block_format.setHeadingLevel(definition["level"])
            block_format.setTopMargin(float(definition["before"]))
            block_format.setBottomMargin(float(definition["after"]))
            block_format.setLineHeight(
                float(definition["line"]), QTextBlockFormat.LineHeightTypes.ProportionalHeight.value
            )
            block_format.setIndent(definition["indent"])
            block_cursor.setBlockFormat(block_format)
            char_format = QTextCharFormat()
            char_format.setFontFamilies([self.default_font_family])
            char_format.setFontPointSize(definition["size"])
            char_format.setFontWeight(
                QFont.Weight.Bold if definition["bold"] else QFont.Weight.Normal
            )
            char_format.setFontItalic(definition["italic"])
            char_format.setForeground(QColor(definition["color"]))
            if block.length() > 1:
                content_cursor = QTextCursor(block.document())
                content_cursor.setPosition(block.position())
                content_cursor.setPosition(
                    block.position() + block.length() - 1, QTextCursor.MoveMode.KeepAnchor
                )
                content_cursor.mergeCharFormat(char_format)
            else:
                block_cursor.setCharFormat(char_format)
        cursor.endEditBlock()
        self._update_format_ui()
        self._refresh_navigation()

    def _update_format_ui(self):
        if self.editor is None:
            return
        fmt = self.editor.currentCharFormat()
        self.act_bold.setChecked(fmt.fontWeight() >= QFont.Weight.Bold)
        self.act_italic.setChecked(fmt.fontItalic())
        self.act_underline.setChecked(fmt.fontUnderline())
        self.act_strike.setChecked(fmt.fontStrikeOut())
        self.act_superscript.setChecked(
            fmt.verticalAlignment() == QTextCharFormat.VerticalAlignment.AlignSuperScript
        )
        self.act_subscript.setChecked(
            fmt.verticalAlignment() == QTextCharFormat.VerticalAlignment.AlignSubScript
        )
        font = self.editor.currentFont()
        size = fmt.fontPointSize() or self.editor.fontPointSize() or self.default_font_size
        with QSignalBlocker(self.font_combo):
            self.font_combo.setCurrentFont(font)
        with QSignalBlocker(self.size_combo):
            self.size_combo.setCurrentText(str(int(round(size))))
        alignment = self.editor.alignment()
        self.act_align_left.setChecked(not alignment & Qt.AlignmentFlag.AlignHCenter and not alignment & Qt.AlignmentFlag.AlignRight and not alignment & Qt.AlignmentFlag.AlignJustify)
        self.act_align_center.setChecked(bool(alignment & Qt.AlignmentFlag.AlignHCenter))
        self.act_align_right.setChecked(bool(alignment & Qt.AlignmentFlag.AlignRight))
        self.act_align_justify.setChecked(bool(alignment & Qt.AlignmentFlag.AlignJustify))

    # ------------------------------------------------------------------
    # Page and layout
    # ------------------------------------------------------------------
    def _set_margin_inches(self, inches):
        tab = self._current_tab()
        if tab is None:
            return
        tab.paper.set_page(tab.paper.base_page_w, tab.paper.base_page_h, int(inches * 96))
        tab.refresh_ruler()
        self._update_statusbar()

    def _set_orientation(self, orientation):
        tab = self._current_tab()
        if tab is None:
            return
        width, height = tab.paper.base_page_w, tab.paper.base_page_h
        if orientation == "landscape" and width < height:
            width, height = height, width
        elif orientation == "portrait" and width > height:
            width, height = height, width
        tab.paper.set_page(width, height, tab.paper.base_margin)
        tab.refresh_ruler()
        self._update_statusbar()

    def _set_page_size(self, name):
        tab = self._current_tab()
        if tab is None or name not in PAGE_SIZES:
            return
        width_in, height_in = PAGE_SIZES[name]
        landscape = tab.paper.base_page_w > tab.paper.base_page_h
        if landscape:
            width_in, height_in = height_in, width_in
        tab.paper.set_page(int(width_in * 96), int(height_in * 96), tab.paper.base_margin)
        tab.refresh_ruler()
        self._update_statusbar()

    def _set_page_color(self, color):
        tab = self._current_tab()
        if tab is None:
            return
        tab.paper.page_color = QColor(color)
        tab.paper.update()

    def _choose_page_color(self, checked=False):
        tab = self._current_tab()
        if tab is None:
            return
        color = QColorDialog.getColor(tab.paper.page_color, self, "Page Color")
        if color.isValid():
            self._set_page_color(color)

    def _add_watermark(self, checked=False):
        tab = self._current_tab()
        if tab is None:
            return
        paper = tab.paper
        dialog = WatermarkDialog(
            self,
            paper.watermark,
            paper.watermark_size,
            paper.watermark_color,
            paper.watermark_opacity,
        )
        if dialog.exec() == QDialog.DialogCode.Accepted:
            data = dialog.result_data()
            paper.watermark = data["text"]
            paper.watermark_size = data["size"]
            paper.watermark_color = QColor(data["color"])
            paper.watermark_opacity = data["opacity"]
            paper.update()

    def _remove_watermark(self, checked=False):
        tab = self._current_tab()
        if tab is None:
            return
        tab.paper.watermark = ""
        tab.paper.update()

    def _toggle_page_border(self, checked):
        tab = self._current_tab()
        if tab is None:
            return
        tab.paper.show_border = bool(checked)
        tab.paper.update()

    # ------------------------------------------------------------------
    # Inserting pages, tables, illustrations and fields
    # ------------------------------------------------------------------
    def _insert_page_break(self, checked=False):
        tab = self._current_tab()
        if tab is None:
            return
        cursor = tab.editor.textCursor()
        cursor.beginEditBlock()
        cursor.insertBlock()
        block_format = cursor.blockFormat()
        block_format.setPageBreakPolicy(QTextFormat.PageBreakFlag.PageBreak_AlwaysBefore)
        cursor.setBlockFormat(block_format)
        cursor.endEditBlock()
        tab.editor.setTextCursor(cursor)
        tab.paper.update_size()
        self._update_statusbar()

    def _insert_blank_page(self, checked=False):
        tab = self._current_tab()
        if tab is None:
            return
        cursor = tab.editor.textCursor()
        at_block_start = cursor.position() == cursor.block().position()
        cursor.beginEditBlock()
        cursor.insertBlock()
        if at_block_start:
            blank_block = cursor.block().previous()
            following = cursor.block()
        else:
            blank_block = cursor.block()
            following = blank_block.next()
        if not blank_block.isValid():
            blank_block = cursor.block()
        blank_cursor = QTextCursor(blank_block)
        block_format = blank_cursor.blockFormat()
        if blank_block.previous().isValid():
            block_format.setPageBreakPolicy(QTextFormat.PageBreakFlag.PageBreak_AlwaysBefore)
        else:
            block_format.setPageBreakPolicy(QTextFormat.PageBreakFlag.PageBreak_Auto)
        blank_cursor.setBlockFormat(block_format)
        while following.isValid() and not following.text().strip():
            following = following.next()
        if following.isValid():
            following_cursor = QTextCursor(following)
            following_format = following_cursor.blockFormat()
            following_format.setPageBreakPolicy(QTextFormat.PageBreakFlag.PageBreak_AlwaysBefore)
            following_cursor.setBlockFormat(following_format)
        cursor.endEditBlock()
        tab.editor.setTextCursor(blank_cursor)
        tab.paper.update_size()
        self._update_statusbar()

    def _insert_cover_page(self, checked=False):
        tab = self._current_tab()
        if tab is None:
            return
        from html import escape

        title, accepted = QInputDialog.getText(self, "Cover Page", "Title:", text="Document Title")
        if not accepted:
            return
        subtitle, accepted = QInputDialog.getText(
            self, "Cover Page", "Subtitle:", text="Prepared with SimzWord"
        )
        if not accepted:
            return
        author, accepted = QInputDialog.getText(
            self, "Cover Page", "Author:", text=os.environ.get("USERNAME", "Author")
        )
        if not accepted:
            return
        date_text = datetime.now().strftime("%B %d, %Y")
        cover_html = (
            '<div style="text-align:center; margin-top:150px;">'
            f'<p style="font-size:32pt; font-weight:bold; color:#1F3864;">{escape(title)}</p>'
            f'<p style="font-size:15pt; font-style:italic; color:#555555;">{escape(subtitle)}</p>'
            '<p style="font-size:12pt; color:#666666; margin-top:160px;">'
            f"{escape(author)}</p>"
            f'<p style="font-size:11pt; color:#777777;">{escape(date_text)}</p>'
            "</div>"
        )
        cursor = tab.editor.textCursor()
        cursor.beginEditBlock()
        cursor.movePosition(QTextCursor.MoveOperation.Start)
        cursor.insertHtml(cover_html)
        cursor.insertBlock()
        block_format = cursor.blockFormat()
        block_format.setPageBreakPolicy(QTextFormat.PageBreakFlag.PageBreak_AlwaysBefore)
        cursor.setBlockFormat(block_format)
        cursor.endEditBlock()
        tab.paper.update_size()
        self._update_statusbar()

    def _insert_table(self, checked=False):
        tab = self._current_tab()
        if tab is None:
            return
        dialog = TableDialog(self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        rows = dialog.rows.value()
        columns = dialog.cols.value()
        style = dialog.style_combo.currentData()
        table_format = QTextTableFormat()
        table_format.setCellPadding(6)
        table_format.setCellSpacing(0)
        table_format.setWidth(QTextLength(QTextLength.Type.PercentageLength, 100))
        table_format.setAlignment(Qt.AlignmentFlag.AlignHCenter)
        if style == "grid":
            table_format.setBorder(1)
            table_format.setBorderBrush(QColor("#9AA7B4"))
        else:
            table_format.setBorder(0)
        cursor = tab.editor.textCursor()
        cursor.beginEditBlock()
        table = cursor.insertTable(rows, columns, table_format)
        if dialog.header.isChecked() and rows > 0:
            for column in range(columns):
                cell = table.cellAt(0, column)
                cell_format = cell.format()
                cell_format.setBackground(QColor("#EAF3FB"))
                cell.setFormat(cell_format)
                cell_cursor = cell.firstCursorPosition()
                char_format = QTextCharFormat()
                char_format.setFontWeight(QFont.Weight.Bold)
                cell_cursor.setCharFormat(char_format)
        cursor.endEditBlock()
        tab.editor.setTextCursor(table.cellAt(0, 0).firstCursorPosition())
        tab.paper.update_size()
        self._update_statusbar()

    def _insert_image(self, checked=False):
        tab = self._current_tab()
        if tab is None:
            return
        dialog = ImageGalleryDialog(self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        if dialog.selected_path:
            pixmap = QPixmap(dialog.selected_path)
        elif dialog.generated_pix is not None:
            pixmap = QPixmap(dialog.generated_pix)
        else:
            return
        if pixmap.isNull():
            QMessageBox.warning(self, "Insert Picture", "The selected image could not be loaded.")
            return
        image_path = os.path.join(
            tempfile.gettempdir(),
            f"simzword_image_{datetime.now().strftime('%Y%m%d%H%M%S%f')}.png",
        )
        if not pixmap.save(image_path, "PNG"):
            QMessageBox.warning(self, "Insert Picture", "The image could not be prepared for insertion.")
            return
        image_format = QTextImageFormat()
        image_format.setName(image_path)
        maximum_width = max(120.0, tab.paper.page_w - tab.paper.margin * 2)
        width = min(float(pixmap.width()), maximum_width)
        image_format.setWidth(width)
        image_format.setHeight(width * pixmap.height() / max(1, pixmap.width()))
        tab.editor.textCursor().insertImage(image_format)
        tab.paper.update_size()

    def _insert_shape(self, kind):
        tab = self._current_tab()
        if tab is None:
            return
        color = QColorDialog.getColor(QColor("#2B7CD3"), self, "Shape Color")
        if not color.isValid():
            return
        pixmap = QPixmap(280, 170)
        pixmap.fill(Qt.GlobalColor.transparent)
        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        pen = QPen(color, 4)
        painter.setPen(pen)
        painter.setBrush(QColor(color.red(), color.green(), color.blue(), 60))
        rect = QRectF(35, 25, 210, 120)
        if kind == "Rounded Rectangle":
            painter.drawRoundedRect(rect, 22, 22)
        elif kind == "Ellipse":
            painter.drawEllipse(rect)
        elif kind == "Triangle":
            painter.drawPolygon(
                QPolygonF(
                    [QPointF(140, 20), QPointF(255, 145), QPointF(25, 145)]
                )
            )
        elif kind == "Arrow":
            painter.drawPolygon(
                QPolygonF(
                    [
                        QPointF(25, 70),
                        QPointF(165, 70),
                        QPointF(165, 35),
                        QPointF(255, 85),
                        QPointF(165, 135),
                        QPointF(165, 100),
                        QPointF(25, 100),
                    ]
                )
            )
        elif kind == "Star":
            points = []
            for index in range(10):
                angle = math.pi / 2 + index * math.pi / 5
                radius = 62 if index % 2 == 0 else 27
                points.append(
                    QPointF(140 + radius * math.cos(angle), 85 - radius * math.sin(angle))
                )
            painter.drawPolygon(QPolygonF(points))
        elif kind == "Line":
            painter.drawLine(QPointF(30, 140), QPointF(250, 30))
        else:
            painter.drawRect(rect)
        painter.end()
        shape_path = os.path.join(
            tempfile.gettempdir(),
            f"simzword_shape_{datetime.now().strftime('%Y%m%d%H%M%S%f')}.png",
        )
        if not pixmap.save(shape_path, "PNG"):
            return
        image_format = QTextImageFormat()
        image_format.setName(shape_path)
        image_format.setWidth(220)
        image_format.setHeight(220 * pixmap.height() / max(1, pixmap.width()))
        tab.editor.textCursor().insertImage(image_format)
        tab.paper.update_size()

    def _insert_chart(self, checked=False):
        tab = self._current_tab()
        if tab is None:
            return
        dialog = ChartDialog(self)
        if dialog.exec() != QDialog.DialogCode.Accepted or dialog.pixmap is None:
            return
        chart_path = os.path.join(
            tempfile.gettempdir(),
            f"simzword_chart_{datetime.now().strftime('%Y%m%d%H%M%S%f')}.png",
        )
        if not dialog.pixmap.save(chart_path, "PNG"):
            return
        image_format = QTextImageFormat()
        image_format.setName(chart_path)
        width = min(560.0, max(180.0, tab.paper.page_w - tab.paper.margin * 2))
        image_format.setWidth(width)
        image_format.setHeight(width * dialog.pixmap.height() / max(1, dialog.pixmap.width()))
        tab.editor.textCursor().insertImage(image_format)
        tab.paper.update_size()

    def _insert_link(self, checked=False):
        tab = self._current_tab()
        if tab is None:
            return
        cursor = tab.editor.textCursor()
        selected = cursor.selectedText()
        text, accepted = QInputDialog.getText(
            self, "Insert Hyperlink", "Display text:", text=selected or "Link"
        )
        if not accepted or not text:
            return
        url, accepted = QInputDialog.getText(
            self, "Insert Hyperlink", "Address (URL):", text="https://"
        )
        if not accepted or not url:
            return
        char_format = QTextCharFormat()
        char_format.setAnchor(True)
        char_format.setAnchorHref(url)
        char_format.setForeground(QColor("#0563C1"))
        char_format.setFontUnderline(True)
        if cursor.hasSelection():
            cursor.mergeCharFormat(char_format)
        else:
            cursor.insertText(text, char_format)

    def _insert_bookmark(self, checked=False):
        tab = self._current_tab()
        if tab is None:
            return
        name, accepted = QInputDialog.getText(
            self,
            "Bookmark",
            "Bookmark name:",
            text=f"Bookmark{len(tab.bookmarks) + 1}",
        )
        if not accepted or not name.strip():
            return
        tab.bookmarks[name.strip()] = tab.editor.textCursor().position()
        self._refresh_navigation()
        self.statusBar().showMessage(f"Bookmark '{name.strip()}' added", 3000)

    def _header_footer(self, checked=False):
        tab = self._current_tab()
        if tab is None:
            return
        paper = tab.paper
        dialog = HeaderFooterDialog(
            self,
            paper.header_text,
            paper.footer_text,
            paper.show_header,
            paper.show_footer,
            paper.show_page_numbers,
            paper.page_number_format,
        )
        if dialog.exec() == QDialog.DialogCode.Accepted:
            data = dialog.result_data()
            paper.header_text = data["header_text"]
            paper.footer_text = data["footer_text"]
            paper.show_header = data["show_header"]
            paper.show_footer = data["show_footer"]
            paper.show_page_numbers = data["show_numbers"]
            paper.page_number_format = data["number_format"]
            paper.update()

    def _page_number_dialog(self, checked=False):
        tab = self._current_tab()
        if tab is None:
            return
        paper = tab.paper
        dialog = HeaderFooterDialog(
            self,
            paper.header_text,
            paper.footer_text,
            paper.show_header,
            paper.show_footer,
            True,
            paper.page_number_format,
        )
        if dialog.exec() == QDialog.DialogCode.Accepted:
            data = dialog.result_data()
            paper.header_text = data["header_text"]
            paper.footer_text = data["footer_text"]
            paper.show_header = data["show_header"]
            paper.show_footer = data["show_footer"]
            paper.show_page_numbers = data["show_numbers"]
            paper.page_number_format = data["number_format"]
            paper.update()

    def _insert_datetime(self, checked=False):
        if self.editor is None:
            return
        formats = [
            "%B %d, %Y",
            "%d/%m/%Y",
            "%A, %B %d, %Y",
            "%Y-%m-%d",
            "%I:%M %p",
            "%B %d, %Y %I:%M %p",
        ]
        item, accepted = QInputDialog.getItem(
            self, "Date and Time", "Format:", formats, 0, False
        )
        if accepted and item:
            self.editor.insertPlainText(datetime.now().strftime(item))

    def _insert_symbol(self, checked=False):
        if self.editor is None:
            return
        dialog = SymbolDialog(self)
        if dialog.exec() == QDialog.DialogCode.Accepted and dialog.selected:
            self.editor.insertPlainText(dialog.selected)

    def _insert_hr(self, checked=False):
        if self.editor is None:
            return
        cursor = self.editor.textCursor()
        cursor.beginEditBlock()
        cursor.insertHtml("<hr>")
        cursor.insertBlock()
        cursor.endEditBlock()
        self.editor.setTextCursor(cursor)
        tab = self._current_tab()
        if tab is not None:
            tab.paper.update_size()

    def _insert_text_box(self, checked=False):
        tab = self._current_tab()
        if tab is None:
            return
        text, accepted = QInputDialog.getMultiLineText(
            self, "Text Box", "Text:", "Type your text here."
        )
        if not accepted:
            return
        frame_format = QTextFrameFormat()
        frame_format.setBorder(1)
        frame_format.setBorderBrush(QColor("#9AA7B4"))
        frame_format.setPadding(10)
        frame_format.setBackground(QColor("#F2F7FB"))
        frame_format.setWidth(QTextLength(QTextLength.Type.PercentageLength, 70))
        cursor = tab.editor.textCursor()
        cursor.beginEditBlock()
        frame = cursor.insertFrame(frame_format)
        frame_cursor = frame.firstCursorPosition()
        frame_cursor.insertText(text)
        cursor.endEditBlock()
        tab.paper.update_size()

    def _add_comment(self, checked=False):
        tab = self._current_tab()
        if tab is None:
            return
        cursor = tab.editor.textCursor()
        if not cursor.hasSelection():
            QMessageBox.information(
                self, "New Comment", "Select the text you want to comment on first."
            )
            return
        text, accepted = QInputDialog.getMultiLineText(self, "New Comment", "Comment:")
        if not accepted or not text.strip():
            return
        comment_id = tab.next_comment_id
        tab.next_comment_id += 1
        excerpt = cursor.selectedText().replace("\u2029", " ")[:80]
        tab.comments[comment_id] = {
            "text": text.strip(),
            "author": os.environ.get("USERNAME", "You"),
            "date": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "excerpt": excerpt,
        }
        char_format = QTextCharFormat()
        char_format.setProperty(COMMENT_PROP, comment_id)
        char_format.setToolTip(f"Comment: {text.strip()}")
        char_format.setBackground(QColor("#FFF2CC"))
        char_format.setUnderlineStyle(QTextCharFormat.UnderlineStyle.DotLine)
        char_format.setUnderlineColor(QColor("#D35400"))
        cursor.mergeCharFormat(char_format)
        self._refresh_comments()
        self.act_toggle_comments.setChecked(True)
        self.comments_dock.show()
        self.comments_dock.raise_()

    def _insert_footnote(self, checked=False):
        tab = self._current_tab()
        if tab is None:
            return
        text, accepted = QInputDialog.getText(self, "Insert Footnote", "Footnote text:")
        if not accepted or not text.strip():
            return
        tab.footnote_count += 1
        number = tab.footnote_count
        cursor = tab.editor.textCursor()
        marker_format = QTextCharFormat()
        marker_format.setVerticalAlignment(QTextCharFormat.VerticalAlignment.AlignSuperScript)
        marker_format.setForeground(QColor("#2B7CD3"))
        cursor.insertText(str(number), marker_format)
        end_cursor = QTextCursor(tab.editor.document())
        end_cursor.movePosition(QTextCursor.MoveOperation.End)
        end_cursor.beginEditBlock()
        end_cursor.insertBlock()
        separator_format = QTextCharFormat()
        separator_format.setForeground(QColor("#AAAAAA"))
        end_cursor.insertText("â”€" * 40, separator_format)
        end_cursor.insertBlock()
        footnote_format = QTextCharFormat()
        footnote_format.setFontPointSize(max(8, self.default_font_size - 2))
        footnote_format.setForeground(QColor("#555555"))
        end_cursor.insertText(f"{number}. {text.strip()}", footnote_format)
        end_cursor.endEditBlock()
        tab.paper.update_size()

    def _insert_citation(self, checked=False):
        tab = self._current_tab()
        if tab is None:
            return
        author, accepted = QInputDialog.getText(self, "Insert Citation", "Author:")
        if not accepted or not author.strip():
            return
        year, accepted = QInputDialog.getText(self, "Insert Citation", "Year:")
        if not accepted or not year.strip():
            return
        title, accepted = QInputDialog.getText(self, "Insert Citation", "Title:")
        if not accepted:
            return
        citation = {
            "author": author.strip(),
            "year": year.strip(),
            "title": title.strip(),
        }
        tab.citations.append(citation)
        tab.editor.insertPlainText(f"({citation['author']}, {citation['year']})")

    def _insert_bibliography(self, checked=False):
        tab = self._current_tab()
        if tab is None:
            return
        if not tab.citations:
            QMessageBox.information(
                self, "Bibliography", "Add at least one citation before creating a bibliography."
            )
            return
        cursor = QTextCursor(tab.editor.document())
        cursor.movePosition(QTextCursor.MoveOperation.End)
        cursor.beginEditBlock()
        cursor.insertBlock()
        heading_format = QTextCharFormat()
        heading_format.setFontPointSize(15)
        heading_format.setFontWeight(QFont.Weight.Bold)
        heading_format.setForeground(QColor("#1F4E79"))
        cursor.insertText("References", heading_format)
        cursor.insertBlock()
        for citation in tab.citations:
            entry = f"{citation['author']} ({citation['year']}). {citation['title']}."
            cursor.insertText(entry)
            cursor.insertBlock()
        cursor.endEditBlock()
        tab.paper.update_size()

    def _insert_caption(self, checked=False):
        tab = self._current_tab()
        if tab is None:
            return
        text, accepted = QInputDialog.getText(self, "Insert Caption", "Caption text:")
        if not accepted or not text.strip():
            return
        tab.figure_count += 1
        caption_format = QTextCharFormat()
        caption_format.setFontItalic(True)
        caption_format.setForeground(QColor("#555555"))
        tab.editor.textCursor().insertText(
            f"Figure {tab.figure_count}: {text.strip()}", caption_format
        )

    def _page_for_block(self, tab, block):
        try:
            top = tab.editor.document().documentLayout().blockBoundingRect(block).top()
            page_height = max(1.0, float(tab.paper.page_h))
            return max(1, min(tab.paper.pages, int(top // page_height) + 1))
        except Exception:
            return 1

    def _insert_toc(self, checked=False):
        tab = self._current_tab()
        if tab is None:
            return
        document = tab.editor.document()
        headings = []
        block = document.begin()
        while block.isValid():
            level = block.blockFormat().headingLevel()
            text = block.text().strip()
            if level > 0 and text and not bool(block.blockFormat().property(TOC_PROP)):
                headings.append((level, text, block))
            block = block.next()
        cursor = tab.editor.textCursor()
        origin_block = cursor.block()
        origin_position = origin_block.position()
        origin_text = origin_block.text()
        origin_format = QTextBlockFormat(origin_block.blockFormat())
        cursor.beginEditBlock()
        if origin_text.strip() or bool(origin_block.blockFormat().property(TOC_PROP)):
            cursor.insertBlock()
        title_format = QTextCharFormat()
        title_format.setFontPointSize(16)
        title_format.setFontWeight(QFont.Weight.Bold)
        title_format.setForeground(QColor("#1F4E79"))
        cursor.insertText("Table of Contents", title_format)
        title_block = cursor.block()
        title_block_format = cursor.blockFormat()
        title_block_format.setHeadingLevel(0)
        title_block_format.setProperty(TOC_PROP, True)
        cursor.setBlockFormat(title_block_format)
        cursor.insertBlock()
        entry_block_format = cursor.blockFormat()
        entry_block_format.setHeadingLevel(0)
        entry_block_format.setProperty(TOC_PROP, True)
        cursor.setBlockFormat(entry_block_format)
        if not headings:
            cursor.insertText("No headings found. Apply a Heading style to build a table of contents.")
            cursor.insertBlock()
        else:
            for level, text, heading_block in headings:
                indent = "    " * max(0, level - 1)
                page = self._page_for_block(tab, heading_block)
                dots = "." * max(3, 72 - len(indent) - len(text))
                cursor.insertText(f"{indent}{text} {dots} {page}")
                entry_format = cursor.blockFormat()
                entry_format.setHeadingLevel(0)
                entry_format.setProperty(TOC_PROP, True)
                cursor.setBlockFormat(entry_format)
                cursor.insertBlock()
        final_format = cursor.blockFormat()
        final_format.setHeadingLevel(0)
        final_format.setProperty(TOC_PROP, False)
        final_format.setProperty(TOC_END_PROP, True)
        cursor.setBlockFormat(final_format)
        cursor.endEditBlock()
        if origin_text:
            block = document.begin()
            while block.isValid():
                if (block.position() >= origin_position and block.text() == origin_text
                        and not bool(block.blockFormat().property(TOC_PROP))):
                    block_cursor = QTextCursor(block)
                    block_cursor.setBlockFormat(origin_format)
                    break
                block = block.next()
            block = document.begin()
            title = None
            while block.isValid():
                if bool(block.blockFormat().property(TOC_PROP)) and block.text() == "Table of Contents":
                    title = block
                    break
                block = block.next()
            if title is not None:
                previous = title.previous()
                if (previous.isValid() and not previous.text().strip()
                        and not bool(previous.blockFormat().property(TOC_PROP))
                        and previous.blockFormat().headingLevel() > 0):
                    block_cursor = QTextCursor(previous)
                    block_cursor.select(QTextCursor.SelectionType.BlockUnderCursor)
                    block_cursor.removeSelectedText()
        tab.paper.update_size()

    def _update_toc(self, checked=False):
        tab = self._current_tab()
        if tab is None:
            return
        document = tab.editor.document()
        marked = []
        block = document.begin()
        while block.isValid():
            if (bool(block.blockFormat().property(TOC_PROP))
                    or bool(block.blockFormat().property(TOC_END_PROP))):
                marked.append(block)
            block = block.next()
        if not marked:
            self._insert_toc()
            return
        marked.sort(key=lambda item: item.position())
        first_block = marked[0]
        last_block = marked[-1]
        following = last_block.next()
        following_format = QTextBlockFormat(following.blockFormat()) if following.isValid() else None
        start = first_block.position()
        end = last_block.position() + last_block.length()
        cursor = QTextCursor(document)
        cursor.setPosition(start)
        cursor.setPosition(min(end, document.characterCount() - 1), QTextCursor.MoveMode.KeepAnchor)
        cursor.removeSelectedText()
        target = document.findBlock(min(start, max(0, document.characterCount() - 1)))
        if target.isValid():
            cursor = QTextCursor(target)
            if following_format is not None:
                cursor.setBlockFormat(following_format)
            else:
                target_format = cursor.blockFormat()
                target_format.setProperty(TOC_PROP, False)
                target_format.setProperty(TOC_END_PROP, False)
                cursor.setBlockFormat(target_format)
        insert_block = None
        if target.isValid():
            previous = target.previous()
            if (previous.isValid() and not previous.text().strip()
                    and not bool(previous.blockFormat().property(TOC_PROP))):
                insert_block = previous
            elif (not target.text().strip()
                    and not bool(target.blockFormat().property(TOC_PROP))):
                insert_block = target
        if insert_block is not None:
            cursor = QTextCursor(insert_block)
            cursor.movePosition(QTextCursor.MoveOperation.StartOfBlock)
            tab.editor.setTextCursor(cursor)
        elif target.isValid():
            cursor = QTextCursor(target)
            cursor.setPosition(target.position())
            tab.editor.setTextCursor(cursor)
        else:
            cursor = QTextCursor(document)
            cursor.movePosition(QTextCursor.MoveOperation.End)
            tab.editor.setTextCursor(cursor)
        self._insert_toc()
        self.statusBar().showMessage("Table of contents updated", 3000)

    # ------------------------------------------------------------------
    # Tables
    # ------------------------------------------------------------------
    def _table_context(self):
        tab = self._current_tab()
        if tab is None:
            return None
        cursor = tab.editor.textCursor()
        table = cursor.currentTable()
        if table is None:
            return None
        cell = table.cellAt(cursor)
        if not cell.isValid():
            return None
        return tab, cursor, table, cell

    def _table_insert_row(self, above):
        context = self._table_context()
        if context is None:
            return
        _, _, table, cell = context
        row = cell.row() if above else cell.row() + 1
        table.insertRows(row, 1)

    def _table_insert_column(self, left):
        context = self._table_context()
        if context is None:
            return
        _, _, table, cell = context
        column = cell.column() if left else cell.column() + 1
        table.insertColumns(column, 1)

    def _table_delete_row(self, checked=False):
        context = self._table_context()
        if context is None:
            return
        tab, _, table, cell = context
        if table.rows() <= 1:
            QMessageBox.information(self, "Delete Row", "A table must have at least one row.")
            return
        table.removeRows(cell.row(), 1)
        tab.paper.update_size()

    def _table_delete_column(self, checked=False):
        context = self._table_context()
        if context is None:
            return
        tab, _, table, cell = context
        if table.columns() <= 1:
            QMessageBox.information(self, "Delete Column", "A table must have at least one column.")
            return
        table.removeColumns(cell.column(), 1)
        tab.paper.update_size()

    def _table_delete(self, checked=False):
        context = self._table_context()
        if context is None:
            return
        tab, _, table, _ = context
        first = table.firstCursorPosition().position()
        last = table.lastCursorPosition().position()
        cursor = QTextCursor(tab.editor.document())
        cursor.setPosition(first)
        cursor.setPosition(last, QTextCursor.MoveMode.KeepAnchor)
        cursor.removeSelectedText()
        tab.paper.update_size()

    def _table_selection_cells(self, tab, cursor, table):
        document = tab.editor.document()
        start_cursor = QTextCursor(document)
        start_cursor.setPosition(cursor.selectionStart())
        end_cursor = QTextCursor(document)
        end_cursor.setPosition(max(cursor.selectionStart(), cursor.selectionEnd() - 1))
        start_cell = table.cellAt(start_cursor)
        end_cell = table.cellAt(end_cursor)
        if not start_cell.isValid() or not end_cell.isValid():
            return None
        return start_cell, end_cell

    def _table_merge_cells(self, checked=False):
        context = self._table_context()
        if context is None:
            return
        tab, cursor, table, cell = context
        cells = self._table_selection_cells(tab, cursor, table)
        if cells is None:
            return
        start_cell, end_cell = cells
        row_span = end_cell.row() - start_cell.row() + 1
        column_span = end_cell.column() - start_cell.column() + 1
        if row_span * column_span < 2:
            QMessageBox.information(
                self, "Merge Cells", "Select two or more adjacent cells to merge them."
            )
            return
        table.mergeCells(start_cell.row(), start_cell.column(), row_span, column_span)
        tab.paper.update_size()

    def _table_split_cell(self, checked=False):
        context = self._table_context()
        if context is None:
            return
        tab, _, table, cell = context
        table.splitCell(cell.row(), cell.column(), 1, 1)
        tab.paper.update_size()

    def _table_cell_properties(self, checked=False):
        context = self._table_context()
        if context is None:
            return
        tab, cursor, table, cell = context
        color = QColorDialog.getColor(QColor("#EAF3FB"), self, "Cell Background")
        if not color.isValid():
            return
        cells = self._table_selection_cells(tab, cursor, table)
        if cells is None:
            cells = (cell, cell)
        start_cell, end_cell = cells
        for row in range(start_cell.row(), end_cell.row() + 1):
            for column in range(start_cell.column(), end_cell.column() + 1):
                target = table.cellAt(row, column)
                cell_format = target.format()
                cell_format.setBackground(color)
                target.setFormat(cell_format)
        tab.paper.update_size()

    def _table_properties(self, checked=False):
        context = self._table_context()
        if context is None:
            return
        tab, _, table, _ = context
        border, accepted = QInputDialog.getInt(
            self,
            "Table Properties",
            "Border width (0 for none):",
            int(table.format().border()),
            0,
            20,
            1,
        )
        if not accepted:
            return
        padding, accepted = QInputDialog.getInt(
            self,
            "Table Properties",
            "Cell padding:",
            int(table.format().cellPadding()),
            0,
            30,
            1,
        )
        if not accepted:
            return
        table_format = table.format()
        table_format.setBorder(border)
        table_format.setCellPadding(padding)
        if border > 0:
            table_format.setBorderBrush(QColor("#9AA7B4"))
        table.setFormat(table_format)
        tab.paper.update_size()

    # ------------------------------------------------------------------
    # Navigation, comments and revision tracking
    # ------------------------------------------------------------------
    def _refresh_navigation(self, *args):
        if not hasattr(self, "nav_headings"):
            return
        self.nav_headings.clear()
        self.nav_bookmarks.clear()
        tab = self._current_tab()
        if tab is None:
            return
        document = tab.editor.document()
        block = document.begin()
        while block.isValid():
            level = block.blockFormat().headingLevel()
            if level > 0:
                text = block.text().strip() or "(empty heading)"
                item = QListWidgetItem("    " * max(0, level - 1) + text)
                item.setData(Qt.ItemDataRole.UserRole, block.position())
                item.setToolTip(f"Heading {level}")
                self.nav_headings.addItem(item)
            block = block.next()
        for name, position in sorted(tab.bookmarks.items(), key=lambda pair: pair[1]):
            item = QListWidgetItem(name)
            item.setData(Qt.ItemDataRole.UserRole, int(position))
            self.nav_bookmarks.addItem(item)

    def _goto_position(self, position):
        tab = self._current_tab()
        if tab is None:
            return
        document = tab.editor.document()
        block = document.findBlock(int(position))
        if not block.isValid():
            return
        cursor = QTextCursor(block)
        tab.editor.setTextCursor(cursor)
        tab.editor.setFocus()
        self._ensure_cursor_visible(tab)

    def _nav_heading_clicked(self, item):
        if item is not None:
            self._goto_position(item.data(Qt.ItemDataRole.UserRole))

    def _nav_bookmark_clicked(self, item):
        if item is not None:
            self._goto_position(item.data(Qt.ItemDataRole.UserRole))

    def _comment_ranges(self, tab, comment_id):
        ranges = []
        document = tab.editor.document()
        block = document.begin()
        while block.isValid():
            iterator = block.begin()
            while not iterator.atEnd():
                fragment = iterator.fragment()
                if fragment.isValid() and fragment.charFormat().property(COMMENT_PROP) == comment_id:
                    ranges.append(
                        (fragment.position(), fragment.position() + fragment.length(), fragment.charFormat())
                    )
                iterator += 1
            block = block.next()
        return ranges

    def _refresh_comments(self):
        if not hasattr(self, "comments_list"):
            return
        self.comments_list.clear()
        tab = self._current_tab()
        if tab is None:
            return
        for comment_id, data in sorted(tab.comments.items()):
            item = QListWidgetItem(
                f"{data.get('author', 'You')}  ({data.get('date', '')})\n"
                f"\"{data.get('excerpt', '')}\"\n{data.get('text', '')}"
            )
            item.setData(Qt.ItemDataRole.UserRole, int(comment_id))
            self.comments_list.addItem(item)

    def _goto_comment(self, comment_id):
        tab = self._current_tab()
        if tab is None:
            return
        ranges = self._comment_ranges(tab, int(comment_id))
        if not ranges:
            return
        start, end, _ = ranges[0]
        cursor = QTextCursor(tab.editor.document())
        cursor.setPosition(start)
        cursor.setPosition(end, QTextCursor.MoveMode.KeepAnchor)
        tab.editor.setTextCursor(cursor)
        tab.editor.setFocus()
        self._ensure_cursor_visible(tab)

    def _comment_item_clicked(self, item):
        if item is not None:
            self._goto_comment(item.data(Qt.ItemDataRole.UserRole))

    def _go_to_selected_comment(self, checked=False):
        item = self.comments_list.currentItem()
        if item is None:
            QMessageBox.information(self, "Comments", "Select a comment first.")
            return
        self._goto_comment(item.data(Qt.ItemDataRole.UserRole))

    def _comment_id_at_cursor(self):
        tab = self._current_tab()
        if tab is None:
            return None
        value = tab.editor.textCursor().charFormat().property(COMMENT_PROP)
        return int(value) if value is not None else None

    def _remove_comment(self, tab, comment_id):
        ranges = self._comment_ranges(tab, comment_id)
        document = tab.editor.document()
        for start, end, char_format in ranges:
            clean = char_format
            clean.clearProperty(COMMENT_PROP)
            clean.setToolTip("")
            clean.setBackground(QColor(0, 0, 0, 0))
            clean.setUnderlineStyle(QTextCharFormat.UnderlineStyle.NoUnderline)
            clean.setUnderlineColor(QColor(0, 0, 0, 0))
            cursor = QTextCursor(document)
            cursor.setPosition(start)
            cursor.setPosition(end, QTextCursor.MoveMode.KeepAnchor)
            cursor.setCharFormat(clean)
        tab.comments.pop(comment_id, None)
        self._refresh_comments()

    def _delete_selected_comment(self, checked=False):
        tab = self._current_tab()
        if tab is None:
            return
        item = self.comments_list.currentItem()
        comment_id = item.data(Qt.ItemDataRole.UserRole) if item is not None else None
        if comment_id is None:
            comment_id = self._comment_id_at_cursor()
        if comment_id is None:
            QMessageBox.information(self, "Comments", "Select a comment to delete.")
            return
        self._remove_comment(tab, int(comment_id))

    def _toggle_comments_pane(self, checked):
        self.comments_dock.setVisible(bool(checked))
        if checked:
            self.comments_dock.raise_()

    def _toggle_revisions_pane(self, checked):
        self.revisions_dock.setVisible(bool(checked))
        if checked:
            self.revisions_dock.raise_()

    def _refresh_revision_log(self):
        if not hasattr(self, "revision_view"):
            return
        tab = self._current_tab()
        if tab is None:
            self.revision_view.clear()
            return
        self.revision_view.setPlainText("\n".join(reversed(getattr(tab, "revision_log", []))))

    def _log_revision(self, tab):
        if tab._pending_old is None:
            return
        old_text = tab._pending_old
        new_text = tab.editor.toPlainText()
        tab._pending_old = None
        if old_text == new_text:
            tab.last_text = new_text
            return
        prefix = 0
        limit = min(len(old_text), len(new_text))
        while prefix < limit and old_text[prefix] == new_text[prefix]:
            prefix += 1
        suffix = 0
        while (
            suffix < min(len(old_text) - prefix, len(new_text) - prefix)
            and old_text[len(old_text) - 1 - suffix] == new_text[len(new_text) - 1 - suffix]
        ):
            suffix += 1
        removed = old_text[prefix : len(old_text) - suffix] if suffix else old_text[prefix:]
        added = new_text[prefix : len(new_text) - suffix] if suffix else new_text[prefix:]
        if not removed and not added:
            tab.last_text = new_text
            return

        def shorten(value):
            value = value.replace("\n", " ").replace("\r", " ")
            return value[:60] + "..." if len(value) > 60 else value

        entry = (
            f"{datetime.now().strftime('%H:%M:%S')}  "
            f"removed: {shorten(removed)!r}  added: {shorten(added)!r}"
        )
        tab.revision_log.append(entry)
        tab.revision_log = tab.revision_log[-500:]
        tab.last_text = new_text
        if tab is self._current_tab():
            self._refresh_revision_log()

    def _flush_revisions(self):
        for index in range(self.tab_widget.count()):
            tab = self.tab_widget.widget(index)
            if isinstance(tab, DocumentTab) and tab._pending_old is not None:
                self._log_revision(tab)

    def _clear_revision_log(self, checked=False):
        tab = self._current_tab()
        if tab is None:
            return
        tab.revision_log.clear()
        self._refresh_revision_log()

    def _toggle_track_changes(self, checked):
        self._track_changes = bool(checked)
        if checked:
            for index in range(self.tab_widget.count()):
                tab = self.tab_widget.widget(index)
                if isinstance(tab, DocumentTab):
                    tab.last_text = tab.editor.toPlainText()
                    tab._pending_old = None
            self.lbl_track.setText("TRACKING")
            self.act_toggle_revisions.setChecked(True)
            self.revisions_dock.show()
            self.revisions_dock.raise_()
            self.statusBar().showMessage(
                "Track Changes is on - edits are recorded in the revision log", 4000
            )
        else:
            self._flush_revisions()
            self.lbl_track.setText("")
            for index in range(self.tab_widget.count()):
                tab = self.tab_widget.widget(index)
                if isinstance(tab, DocumentTab):
                    tab.last_text = tab.editor.toPlainText()
                    tab._pending_old = None
            self.statusBar().showMessage("Track Changes is off", 3000)

    # ------------------------------------------------------------------
    # Review tools
    # ------------------------------------------------------------------
    def _spelling(self, checked=False):
        if self.editor is None:
            return
        dialog = SpellingDialog(self)
        dialog.exec()

    def _thesaurus(self, checked=False):
        if self.editor is None:
            return
        cursor = self.editor.textCursor()
        selected = cursor.selectedText()
        if not selected:
            selected = cursor.selectedText()
        dialog = ThesaurusDialog(self, selected_text=selected)
        dialog.exec()

    def _word_count(self, checked=False):
        tab = self._current_tab()
        if tab is None:
            return
        dialog = WordCountDialog(self, tab.editor.toPlainText(), tab.paper.pages)
        dialog.exec()

    def _set_language(self, checked=False):
        QMessageBox.information(
            self,
            "Set Language",
            "SimzWord currently checks spelling with its bundled English (US) dictionary.",
        )

    def add_custom_word(self, word):
        if not word:
            return
        self.custom_words.add(word.strip().lower())
        self.settings.setValue("spelling/custom", sorted(self.custom_words))

    def _add_custom_word(self, word):
        self.add_custom_word(word)

    # ------------------------------------------------------------------
    # Zoom and view
    # ------------------------------------------------------------------
    def _set_zoom(self, value):
        value = max(50, min(200, int(value)))
        self._zoom = value
        with QSignalBlocker(self.zoom_slider):
            self.zoom_slider.setValue(value)
        self.zoom_label.setText(f"{value}%")
        tab = self._current_tab()
        if tab is not None:
            tab.apply_zoom(value)

    def _set_zoom_from_slider(self, value):
        self._set_zoom(value)

    def _zoom_delta(self, delta):
        self._set_zoom(self._zoom + delta)

    def _zoom_dialog(self, checked=False):
        value, accepted = QInputDialog.getInt(
            self, "Zoom", "Zoom level (%):", self._zoom, 50, 200, 10
        )
        if accepted:
            self._set_zoom(value)

    def _toggle_ruler(self, checked):
        for index in range(self.tab_widget.count()):
            tab = self.tab_widget.widget(index)
            if isinstance(tab, DocumentTab):
                tab.ruler.setVisible(bool(checked))
        self.settings.setValue("ui/ruler", bool(checked))

    def _toggle_navigation(self, checked):
        self.nav_dock.setVisible(bool(checked))
        self.settings.setValue("ui/navigation", bool(checked))

    def _toggle_statusbar(self, checked):
        self.statusBar().setVisible(bool(checked))
        self.settings.setValue("ui/statusbar", bool(checked))

    def _toggle_dark(self, checked):
        self.dark = bool(checked)
        self.settings.setValue("ui/dark", self.dark)
        self._apply_theme()

    def _toggle_focus_mode(self, checked):
        self._focus_mode = bool(checked)
        if checked:
            self._focus_hidden_docks = [
                dock
                for dock in (self.nav_dock, self.comments_dock, self.revisions_dock)
                if dock.isVisible()
            ]
            for dock in (self.nav_dock, self.comments_dock, self.revisions_dock):
                dock.hide()
            self.quick_toolbar.hide()
            self.ribbon.hide()
            self.menuBar().hide()
            self.statusBar().hide()
            for index in range(self.tab_widget.count()):
                tab = self.tab_widget.widget(index)
                if isinstance(tab, DocumentTab):
                    tab.ruler.hide()
            self.statusBar().showMessage("Focus Mode: press Esc or click Focus Mode again", 4000)
        else:
            self.quick_toolbar.show()
            self.ribbon.show()
            self.menuBar().show()
            self.statusBar().setVisible(self.act_toggle_status.isChecked())
            for index in range(self.tab_widget.count()):
                tab = self.tab_widget.widget(index)
                if isinstance(tab, DocumentTab):
                    tab.ruler.setVisible(self.act_toggle_ruler.isChecked())
            for dock in self._focus_hidden_docks:
                dock.show()
            self._focus_hidden_docks = []

    def _toggle_full_screen(self, checked):
        if checked:
            self.showFullScreen()
        else:
            self.showNormal()

    # ------------------------------------------------------------------
    # Find, go to and editing helpers
    # ------------------------------------------------------------------
    def _open_find_dialog(self):
        dialog = getattr(self, "find_dialog", None)
        if dialog is None:
            self.find_dialog = FindReplaceDialog(self, self)
            dialog = self.find_dialog
        if self.editor is not None:
            cursor = self.editor.textCursor()
            if cursor.hasSelection():
                dialog.find_edit.setText(cursor.selectedText())
        dialog.show()
        dialog.raise_()
        dialog.activateWindow()
        dialog.find_edit.setFocus()

    def _show_find(self, checked=False):
        self._open_find_dialog()

    def _show_find_replace(self, checked=False):
        self._open_find_dialog()

    def _find_next(self, checked=False):
        dialog = getattr(self, "find_dialog", None)
        if dialog is None or not dialog.isVisible():
            self._open_find_dialog()
            dialog = self.find_dialog
        dialog.find_next()

    def _go_to(self, checked=False):
        tab = self._current_tab()
        if tab is None:
            return
        line, accepted = QInputDialog.getInt(
            self, "Go To", "Line number:", 1, 1, max(1, tab.editor.document().blockCount())
        )
        if not accepted:
            return
        block = tab.editor.document().findBlockByNumber(line - 1)
        if block.isValid():
            tab.editor.setTextCursor(QTextCursor(block))
            tab.editor.setFocus()
            self._ensure_cursor_visible(tab)

    def _clear_contents(self, checked=False):
        if self.editor is None:
            return
        cursor = self.editor.textCursor()
        if cursor.hasSelection():
            cursor.removeSelectedText()
            return
        result = QMessageBox.question(
            self,
            "Clear Contents",
            "Clear the entire document?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if result == QMessageBox.StandardButton.Yes:
            self.editor.clear()

    def _paste_plain(self, checked=False):
        if self.editor is None:
            return
        self.editor.insertPlainText(QApplication.clipboard().text())

    def _select_all(self, checked=False):
        if self.editor is not None:
            self.editor.selectAll()

    # ------------------------------------------------------------------
    # Preferences and help
    # ------------------------------------------------------------------
    def _preferences(self, checked=False):
        dialog = PreferencesDialog(self, self.settings)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        data = dialog.result_data()
        self.settings.setValue("autosave/enabled", data["autosave_enabled"])
        self.settings.setValue("autosave/interval", data["autosave_interval"])
        self.settings.setValue("ui/dark", data["dark"])
        self.settings.setValue("ui/ruler", data["ruler"])
        self.settings.setValue("ui/statusbar", data["statusbar"])
        self.settings.setValue("ui/navigation", data["navigation"])
        self.settings.setValue("editor/font_family", data["font_family"])
        self.settings.setValue("editor/font_size", data["font_size"])
        self.default_font_family = data["font_family"]
        self.default_font_size = data["font_size"]
        self.dark = data["dark"]
        self.act_dark.setChecked(self.dark)
        self._apply_theme()
        self.act_toggle_ruler.setChecked(data["ruler"])
        self._toggle_ruler(data["ruler"])
        self.act_toggle_status.setChecked(data["statusbar"])
        self.statusBar().setVisible(data["statusbar"])
        self.act_toggle_nav.setChecked(data["navigation"])
        self.nav_dock.setVisible(data["navigation"])
        self._setup_autosave()

    def _about(self, checked=False):
        AboutDialog(self).exec()

    def _shortcuts(self, checked=False):
        ShortcutsDialog(self).exec()

    # ------------------------------------------------------------------
    # Context menu and main-window event filter
    # ------------------------------------------------------------------
    def _show_context_menu(self, position):
        tab = self._current_tab()
        if tab is None:
            return
        menu = QMenu(self)
        menu.addAction(self.act_undo)
        menu.addAction(self.act_redo)
        menu.addSeparator()
        menu.addAction(self.act_cut)
        menu.addAction(self.act_copy)
        menu.addAction(self.act_paste)
        menu.addAction(self.act_paste_special)
        menu.addSeparator()
        insert_menu = menu.addMenu("Insert")
        for action in (
            self.act_picture,
            self.act_table,
            self.act_shapes,
            self.act_chart,
            self.act_hyperlink,
            self.act_symbol,
            self.act_date_time,
            self.act_horizontal_rule,
            self.act_comment,
        ):
            insert_menu.addAction(action)
        format_menu = menu.addMenu("Format")
        for action in (
            self.act_font_dialog,
            self.act_paragraph_dialog,
            self.act_bold,
            self.act_italic,
            self.act_underline,
            self.act_strike,
            self.act_text_color,
            self.act_highlight,
            self.act_bullets,
            self.act_numbering,
        ):
            format_menu.addAction(action)
        if tab.editor.textCursor().currentTable() is not None:
            table_menu = menu.addMenu("Table")
            for action in (
                self.act_row_above,
                self.act_row_below,
                self.act_col_left,
                self.act_col_right,
                self.act_delete_row,
                self.act_delete_col,
                self.act_delete_table,
                self.act_merge_cells,
                self.act_split_cell,
            ):
                table_menu.addAction(action)
        menu.addSeparator()
        menu.addAction(self.act_spelling)
        menu.addAction(self.act_thesaurus)
        menu.addAction(self.act_word_count)
        menu.exec(tab.editor.viewport().mapToGlobal(position))

    def eventFilter(self, watched, event):
        try:
            if (
                event.type() == QEvent.Type.MouseButtonRelease
                and self._format_painter_fmt is not None
            ):
                tab = self._current_tab()
                if tab is not None and watched is tab.editor.viewport():
                    cursor = tab.editor.textCursor()
                    if cursor.hasSelection():
                        cursor.mergeCharFormat(self._format_painter_fmt)
                        tab.editor.setTextCursor(cursor)
                        self._format_painter_fmt = None
                        self.act_format_painter.setChecked(False)
                        self.statusBar().showMessage("Format applied", 2000)
        except RuntimeError:
            return False
        return super().eventFilter(watched, event)


def main():
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setOrganizationName(ORG_NAME)
    app.setApplicationVersion(APP_VERSION)
    window = SimzWord()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()



