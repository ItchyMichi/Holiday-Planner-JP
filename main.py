import base64
from datetime import datetime, time
import math
import os
import configparser
import re
import webbrowser
import PyPDF2
import tiktoken
from pdfminer.high_level import extract_text
from email import policy
from email.parser import BytesParser
from bs4 import BeautifulSoup
import re
import json
from json.decoder import JSONDecodeError

import re
from google.oauth2 import service_account
from googleapiclient.discovery import build
from PyQt5.QtWidgets import QInputDialog, QMessageBox, QAbstractItemView, QTableWidget, QTableWidgetItem, QPushButton

import openai
import googlemaps
import requests
from googlemaps.exceptions import ApiError, HTTPError
from PyQt5.QtWidgets import QFileDialog, QMessageBox, QAction, QToolTip, QSizePolicy, QMenu, QListWidget, \
    QListWidgetItem, QTableWidget, QTableWidgetItem, QAbstractItemView
from PyQt5.QtWidgets import (
    QDialog, QTextEdit, QGroupBox, QComboBox, QCompleter,
    QPushButton, QGridLayout, QHBoxLayout, QVBoxLayout,
    QDialogButtonBox, QMessageBox, QFileDialog, QLabel
)

from PyQt5.QtWidgets import QStyledItemDelegate, QDateEdit, QTimeEdit, QSpinBox, QDoubleSpinBox, QCheckBox
from PyQt5.QtCore import QDate, QTime, Qt
from urllib.parse import urlparse, parse_qs
import sys
import traceback
import uuid
import json
import sqlite3
from PyQt5.QtCore import (
    Qt, QRectF, QDate, QTime, QDateTime, pyqtSignal, QPointF
)
from PyQt5.QtGui import QPen, QBrush, QColor, QFontMetrics, QKeySequence
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QFrame, QHBoxLayout, QVBoxLayout,
    QLabel, QGraphicsView, QGraphicsScene, QGraphicsRectItem,
    QGraphicsTextItem, QAction, QToolBar, QDialog, QCalendarWidget,
    QDialogButtonBox, QInputDialog, QColorDialog, QMessageBox,
    QGraphicsItem, QFormLayout, QDateTimeEdit, QSpinBox, QPushButton,
    QLineEdit, QDoubleSpinBox, QTabWidget, QFileDialog
)

# Constants
DEFAULT_SLOT_MINUTES    = 30
DEFAULT_VISIBLE_COLUMNS = 7
DEFAULT_CELL_WIDTH      = 120
DEFAULT_CELL_HEIGHT     = 20
START_HOUR              = 6
END_HOUR                = 23
TEXT_MARGIN             = 5
TIME_LABEL_WIDTH        = 60
HEADER_HEIGHT           = 20
SCRIPT_DIR   = os.path.dirname(os.path.abspath(__file__))
PLACES_FILE  = os.path.join(SCRIPT_DIR, "Saved Places.json")
PLANS_DIR    = os.path.join(SCRIPT_DIR, "plans")
# load config.ini
cfg = configparser.ConfigParser()
cfg.read('config.ini')

# OpenAI
openai.api_key = cfg['openai']['api_key']

# Google credentials for any Google service (e.g. Drive / Gmail if you expand later)
os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = cfg['google']['credentials_dir']
# Google Maps REST client (for place lookups / route optimization)
gmaps_client = googlemaps.Client(key=cfg['google']['maps_api_key'])



# Custom excepthook for full tracebacks
def my_excepthook(exc_type, exc_value, exc_tb):
    traceback.print_exception(exc_type, exc_value, exc_tb)
sys.excepthook = my_excepthook

from PyQt5.QtCore import Qt, QPoint
from PyQt5.QtWidgets import QGraphicsView

URL_RE = re.compile(r'https?://\S+')

from PyQt5.QtWidgets import QDialog, QVBoxLayout, QTextEdit, QPushButton, QFileDialog

class TextImportDialog(QDialog):
    def __init__(self, parent=None, locations=None):
        super().__init__(parent)
        self.setWindowTitle("Import Things to Do/Eat")
        layout = QVBoxLayout(self)

        # — new location selector —
        layout.addWidget(QLabel("Location (e.g. Tokyo, Japan):"))
        self.locationBox = QComboBox()
        self.locationBox.setEditable(True)
        if locations:
            self.locationBox.addItems(locations)
        self.locationBox.setPlaceholderText("Type or select a location…")
        layout.addWidget(self.locationBox)

        # existing text area
        self.textEdit = QTextEdit()
        self.textEdit.setPlaceholderText("Paste your text here, or click “Load File…”")
        layout.addWidget(self.textEdit)

        btn_load = QPushButton("Load .txt File…")
        btn_load.clicked.connect(self.load_file)
        layout.addWidget(btn_load)

        btn_box = QDialogButtonBox(QDialogButtonBox.Ok|QDialogButtonBox.Cancel)
        btn_box.accepted.connect(self.accept)
        btn_box.rejected.connect(self.reject)
        layout.addWidget(btn_box)

    def load_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Select Text File", "", "Text Files (*.txt)"
        )
        if path:
            with open(path, 'r', encoding='utf-8') as f:
                self.textEdit.setPlainText(f.read())

    def get_text(self):
        return self.textEdit.toPlainText().strip()

    def get_location(self):
        return self.locationBox.currentText().strip()



class PlaceDialog(QDialog):
    placeSelected = pyqtSignal(str)
    """ Let user search + pick one of self.starred_places, with duration & activity """
    def __init__(self, parent, starred_places):
        super().__init__(parent)
        self.selected = None
        self.setWindowTitle("Select Saved Place")
        self.starred_places = starred_places
        self._build_ui()

    def _build_ui(self):
        vbox = QVBoxLayout(self)

        # Search bar
        self.filterEdit = QLineEdit()
        self.filterEdit.setPlaceholderText("Filter by name…")
        self.filterEdit.textChanged.connect(self._on_filter)
        vbox.addWidget(self.filterEdit)

        # List of place names
        self.listWidget = QListWidget()
        for p in self.starred_places:
            item = QListWidgetItem(p['name'])
            self.listWidget.addItem(item)
        vbox.addWidget(self.listWidget)

        # Duration + Activity + Add button
        hbox = QHBoxLayout()

        self.durationCombo = QComboBox()
        # add a blank first entry:
        self.durationCombo.addItem("")
        self.durationCombo.addItems([
            "15minutes", "30minutes", "45minutes",
            "60minutes", "2hours", "3hours"
        ])
        hbox.addWidget(QLabel("Duration:"))
        hbox.addWidget(self.durationCombo)

        self.activityCombo = QComboBox()
        # blank default
        self.activityCombo.addItem("")
        self.activityCombo.addItems(["Lunch", "Dinner", "Other Food"])
        hbox.addWidget(QLabel("Activity:"))
        hbox.addWidget(self.activityCombo)

        self.addButton = QPushButton("➕ Add")
        self.addButton.clicked.connect(self._on_add_clicked)
        hbox.addWidget(self.addButton)

        vbox.addLayout(hbox)

        # OK / Cancel
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        vbox.addWidget(bb)

    def _on_filter(self, txt):
        txt = txt.lower()
        for i in range(self.listWidget.count()):
            item = self.listWidget.item(i)
            item.setHidden(txt not in item.text().lower())

    def _on_add_clicked(self):
        idx = self.listWidget.currentRow()
        if idx < 0:
            QMessageBox.warning(self, "No Selection", "Please select a place.")
            return

        p = self.starred_places[idx]
        self.selected = p
        dur = self.durationCombo.currentText()
        act = self.activityCombo.currentText().lower()

        # 2) build your snippet string here
        snippet = (
            f"{p['name']} "
            f"(cid:{p['cid']} lat:{p['lat']} lng:{p['lng']})"
            + (f"({dur})"    if dur else "")
            + (f"({act}),\n" if act else ",\n")
        )

        # 3) emit exactly that local variable
        self.placeSelected.emit(snippet)

        # 4) reset the combos back to blank
        self.durationCombo.setCurrentIndex(0)
        self.activityCombo.setCurrentIndex(0)
        self.listWidget.clearSelection()

    def accept(self):
        super().accept()

    def get_selected(self):
        return self.selected



class LinkableLineEdit(QLineEdit):
    def contextMenuEvent(self, ev):
        # 1) let Qt build the standard menu (Cut/Copy/Paste/etc)
        menu = self.createStandardContextMenu()

        # 2) see if the user has actually selected something that looks like a URL
        sel = self.selectedText().strip()
        if URL_RE.match(sel):
            act = QAction("Open Link", self)
            act.triggered.connect(lambda: webbrowser.open(sel))
            menu.addSeparator()
            menu.addAction(act)

        # 3) pop it up
        menu.exec_(ev.globalPos())

class ClickableLabel(QLabel):
    clicked = pyqtSignal()
    def mouseReleaseEvent(self, ev):
        if ev.button() == Qt.LeftButton:
            self.clicked.emit()
        super().mouseReleaseEvent(ev)

class ZoomableGraphicsView(QGraphicsView):
    def __init__(self, parent=None):
        super().__init__(parent)
        # zoom around the mouse
        self.setTransformationAnchor(QGraphicsView.AnchorUnderMouse)
        self.setResizeAnchor(QGraphicsView.AnchorUnderMouse)
        # --- initialize panning state! ---
        self._panning = False
        self._pan_start = QPoint()
        # allow vertical scrolling
        self.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        self.setMouseTracking(True)
        self.viewport().setMouseTracking(True)


    def wheelEvent(self, event):
        if event.modifiers() & Qt.ControlModifier:
            # Ctrl + wheel → zoom
            zoomIn, zoomOut = 1.25, 1/1.25
            factor = zoomIn if event.angleDelta().y() > 0 else zoomOut
            self.scale(factor, factor)
        else:
            # wheel alone → scroll vertically
            delta = event.angleDelta().y()
            v = self.verticalScrollBar()
            # subtract so wheel-down moves content up, wheel-up moves content down
            v.setValue(v.value() - delta)
        event.accept()

    def mousePressEvent(self, event):
        if event.button() == Qt.MiddleButton:
            # start panning
            self._panning = True
            self._pan_start = event.pos()
            self.setCursor(Qt.ClosedHandCursor)
            event.accept()
        else:
            super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        super().mouseMoveEvent(event)

        # map from widget coords → scene coords
        scene_pt = self.mapToScene(event.pos())
        x, y = scene_pt.x(), scene_pt.y()

        # only inside the grid (skip headers)
        if x >= TIME_LABEL_WIDTH and y >= HEADER_HEIGHT:
            # calculate column and row
            col = int((x - TIME_LABEL_WIDTH) / self.scene().cell_w)
            row = int((y - HEADER_HEIGHT) / self.scene().cell_h)

            # clamp in-bounds
            days = self.scene().start_date.daysTo(self.scene().end_date) + 1
            slots = ((END_HOUR - START_HOUR) * 60) // self.scene().slot_minutes
            if 0 <= col < days and 0 <= row < slots:
                date = self.scene().start_date.addDays(col)
                mins = START_HOUR*60 + row*self.scene().slot_minutes
                time = QTime(mins//60, mins%60)
                txt = f"{date.toString('ddd dd MMM yyyy')} @ {time.toString('HH:mm')}"
                QToolTip.showText(event.globalPos(), txt, self.viewport())
                return

        QToolTip.hideText()

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.MiddleButton:
            self._panning = False
            self.setCursor(Qt.ArrowCursor)
            event.accept()
        else:
            super().mouseReleaseEvent(event)



class DateDelegate(QStyledItemDelegate):
    def createEditor(self, parent, option, index):
        editor = QDateEdit(parent)
        editor.setCalendarPopup(True)
        editor.setDisplayFormat("yyyy-MM-dd")
        return editor
    def setEditorData(self, editor, index):
        text = index.model().data(index, Qt.EditRole)
        date = QDate.fromString(text, Qt.ISODate)
        editor.setDate(date if date.isValid() else QDate.currentDate())
    def setModelData(self, editor, model, index):
        model.setData(index, editor.date().toString(Qt.ISODate), Qt.EditRole)

class TimeDelegate(QStyledItemDelegate):
    def createEditor(self, parent, option, index):
        editor = InvertedTimeEdit(parent)
        editor.setDisplayFormat("HH:mm")
        editor.setTime(QTime.currentTime())
        editor.setKeyboardTracking(False)
        return editor
    def setEditorData(self, editor, index):
        text = index.model().data(index, Qt.EditRole)
        time = QTime.fromString(text, "HH:mm")
        editor.setTime(time if time.isValid() else QTime(0,0))
    def setModelData(self, editor, model, index):
        model.setData(index, editor.time().toString("HH:mm"), Qt.EditRole)

class IntDelegate(QStyledItemDelegate):
    def createEditor(self, parent, option, index):
        editor = QSpinBox(parent)
        editor.setRange(0, 10000)
        return editor
    def setEditorData(self, editor, index):
        editor.setValue(int(index.model().data(index, Qt.EditRole) or 0))
    def setModelData(self, editor, model, index):
        model.setData(index, editor.value(), Qt.EditRole)

from PyQt5.QtWidgets import QTimeEdit
from PyQt5.QtCore import QTime

class InvertedTimeEdit(QTimeEdit):
    def stepBy(self, steps: int):
        # invert up/down
        super().stepBy(+steps)
class DoubleDelegate(QStyledItemDelegate):
    def createEditor(self, parent, option, index):
        editor = QDoubleSpinBox(parent)
        editor.setRange(0.0, 1e6)
        editor.setDecimals(2)
        return editor
    def setEditorData(self, editor, index):
        editor.setValue(float(index.model().data(index, Qt.EditRole) or 0.0))
    def setModelData(self, editor, model, index):
        model.setData(index, editor.value(), Qt.EditRole)

class BoolDelegate(QStyledItemDelegate):
    def createEditor(self, parent, option, index):
        return QCheckBox(parent)
    def setEditorData(self, editor, index):
        editor.setChecked(str(index.model().data(index, Qt.EditRole)).lower() in ("1","true","yes"))
    def setModelData(self, editor, model, index):
        model.setData(index, editor.isChecked(), Qt.EditRole)


class DatabaseManager:
    def __init__(self, path="schedules.db"):
        self.conn = sqlite3.connect(path)
        self.conn.execute("PRAGMA foreign_keys = ON")
        self._create_tables()

    def _create_tables(self):
        c = self.conn.cursor()

        # schedules
        c.execute("""
        CREATE TABLE IF NOT EXISTS schedule (
            id INTEGER PRIMARY KEY,
            name TEXT UNIQUE,
            start_date TEXT,
            end_date TEXT,
            visible_columns INTEGER,
            slot_minutes INTEGER
        )""")

        # core events
        c.execute("""
        CREATE TABLE IF NOT EXISTS event (
            id INTEGER PRIMARY KEY,
            schedule_id INTEGER,
            title TEXT,
            description TEXT,
            group_id TEXT,
            start_dt TEXT,
            end_dt TEXT,
            duration INTEGER,
            link TEXT,
            city TEXT,
            region TEXT,
            event_type TEXT,
            cost REAL,
            color TEXT,
            x REAL, y REAL, w REAL, h REAL,
            FOREIGN KEY(schedule_id) REFERENCES schedule(id) ON DELETE CASCADE
        )""")

        # Location-specific tables
        c.execute("""
        CREATE TABLE IF NOT EXISTS eat_item (
            id          INTEGER PRIMARY KEY,
            schedule_id INTEGER NOT NULL,
            location    TEXT    NOT NULL,
            item        TEXT    NOT NULL,
            link        TEXT,
            description TEXT,
            media       TEXT,
            FOREIGN KEY(schedule_id) REFERENCES schedule(id) ON DELETE CASCADE
        )
        """)

        c.execute("""
        CREATE TABLE IF NOT EXISTS do_item (
            id          INTEGER PRIMARY KEY,
            schedule_id INTEGER NOT NULL,
            location    TEXT    NOT NULL,
            item        TEXT    NOT NULL,
            link        TEXT,
            description TEXT,
            media       TEXT,
            FOREIGN KEY(schedule_id) REFERENCES schedule(id) ON DELETE CASCADE
        )
        """)

        c.execute("""
        CREATE TABLE IF NOT EXISTS hotel (
            id              INTEGER PRIMARY KEY,
            schedule_id     INTEGER NOT NULL,
            location        TEXT    NOT NULL,
            item            TEXT    NOT NULL,
            link            TEXT,
            start_date      TEXT,        -- ISODate
            end_date        TEXT,        -- ISODate
            check_in_time   TEXT,        -- "HH:mm"
            check_out_time  TEXT,        -- "HH:mm"
            breakfast       REAL    DEFAULT 0.0,
            dinner          REAL    DEFAULT 0.0,
            half_board      REAL    DEFAULT 0.0,
            num_rooms       INTEGER DEFAULT 1,
            room_types      TEXT,
            prepaid         BOOLEAN DEFAULT 0,
            cost            REAL    DEFAULT 0.0,
            FOREIGN KEY(schedule_id) REFERENCES schedule(id) ON DELETE CASCADE
        )
        """)

        c.execute("""
        CREATE TABLE IF NOT EXISTS reservation (
            id            INTEGER PRIMARY KEY,
            schedule_id   INTEGER NOT NULL,
            location      TEXT    NOT NULL,
            item          TEXT    NOT NULL,
            link          TEXT,
            description   TEXT,
            media         TEXT,
            start_date    TEXT,        -- ISODate string
            end_date      TEXT,        -- ISODate string
            start_time    TEXT,        -- HH:MM string
            end_time      TEXT,        -- HH:MM string
            FOREIGN KEY(schedule_id) REFERENCES schedule(id) ON DELETE CASCADE
        )
        """)

        # Persisted list of filterable locations
        c.execute("""
        CREATE TABLE IF NOT EXISTS location (
            id          INTEGER PRIMARY KEY,
            schedule_id INTEGER NOT NULL,
            name        TEXT    NOT NULL,
            UNIQUE(schedule_id, name),
            FOREIGN KEY(schedule_id) REFERENCES schedule(id) ON DELETE CASCADE
        )
        """)

        c.execute("PRAGMA table_info(schedule)")
        existing_cols = {row[1] for row in c.fetchall()}

        if 'spreadsheet_id' not in existing_cols:
            c.execute("ALTER TABLE schedule ADD COLUMN spreadsheet_id TEXT")
        if 'sheet_tab' not in existing_cols:
            c.execute("ALTER TABLE schedule ADD COLUMN sheet_tab TEXT")

        self.conn.commit()

        # Indexes to speed up filtering by schedule_id + location
        c.execute("CREATE INDEX IF NOT EXISTS idx_eat_schedule_location ON eat_item(schedule_id, location)")
        c.execute("CREATE INDEX IF NOT EXISTS idx_do_schedule_location ON do_item(schedule_id, location)")
        c.execute("CREATE INDEX IF NOT EXISTS idx_hotel_schedule_location ON hotel(schedule_id, location)")
        c.execute("CREATE INDEX IF NOT EXISTS idx_reservation_schedule_location ON reservation(schedule_id, location)")

        self.conn.commit()

        # For existing DBs (one-time migration)
        c.execute("PRAGMA table_info(event)")
        cols = [row[1] for row in c.fetchall()]
        if "description" not in cols:
            c.execute("ALTER TABLE event ADD COLUMN description TEXT")
        self.conn.commit()

    def save_schedule(self, name, start_date, end_date, visible_columns, slot_minutes):
        c = self.conn.cursor()
        c.execute(
            "INSERT INTO schedule(name, start_date, end_date, visible_columns, slot_minutes) VALUES(?,?,?,?,?)",
            (
                name,
                start_date.toString(Qt.ISODate),
                end_date.toString(Qt.ISODate),
                visible_columns,
                slot_minutes
            )
        )
        self.conn.commit()
        return c.lastrowid

    def list_event_types(self, schedule_id):
        c = self.conn.cursor()
        c.execute("""
            SELECT DISTINCT event_type
              FROM event
             WHERE schedule_id = ?
               AND event_type IS NOT NULL
               AND event_type != ''
             ORDER BY event_type
        """, (schedule_id,))
        return [row[0] for row in c.fetchall()]

    def update_schedule(self, schedule_id, name, start_date, end_date, visible_columns, slot_minutes):
        c = self.conn.cursor()
        c.execute(
            "UPDATE schedule SET name=?, start_date=?, end_date=?, visible_columns=?, slot_minutes=? WHERE id=?",
            (
                name,
                start_date.toString(Qt.ISODate),
                end_date.toString(Qt.ISODate),
                visible_columns,
                slot_minutes,
                schedule_id
            )
        )
        self.conn.commit()

    def update_hotel(self, hotel_id, **fields):
        sets, vals = [], []
        for k, v in fields.items():
            sets.append(f"{k} = ?")
            if isinstance(v, QDate):
                vals.append(v.toString(Qt.ISODate))
            elif isinstance(v, QTime):
                vals.append(v.toString("HH:mm"))
            else:
                vals.append(v)
        vals.append(hotel_id)
        stmt = f"UPDATE hotel SET {','.join(sets)} WHERE id = ?"
        self.conn.cursor().execute(stmt, vals)
        self.conn.commit()

    def update_reservation(self, reservation_id, **fields):
        """
        Update one or more columns on the reservation row.
        Usage: db.update_reservation(42, start_date=QDate(...), end_time="15:30", ...)
        """
        sets, vals = [], []
        for col, val in fields.items():
            sets.append(f"{col} = ?")
            if isinstance(val, QDate):
                vals.append(val.toString(Qt.ISODate))
            elif isinstance(val, QTime):
                vals.append(val.toString("HH:mm"))
            else:
                vals.append(val)
        vals.append(reservation_id)
        stmt = f"UPDATE reservation SET {', '.join(sets)} WHERE id = ?"
        self.conn.cursor().execute(stmt, vals)
        self.conn.commit()
    def add_reservation(self, schedule_id, location, item, link, description, media,
                        start_date, end_date, start_time, end_time):
        c = self.conn.cursor()
        c.execute("""
          INSERT INTO reservation(
            schedule_id, location, item, link, description, media,
            start_date, end_date, start_time, end_time
          ) VALUES (?,?,?,?,?,?,?,?,?,?)
        """, (
            schedule_id, location, item, link, description, media,
            start_date.toString(Qt.ISODate),
            end_date.toString(Qt.ISODate),
            start_time.toString("HH:mm"),
            end_time.toString("HH:mm")
        ))
        self.conn.commit()
        return c.lastrowid

    def save_sheet_info(self, schedule_id: int, spreadsheet_id: str, sheet_tab: str):
        self.conn.execute(
            "UPDATE schedule SET spreadsheet_id=?, sheet_tab=? WHERE id=?",
            (spreadsheet_id, sheet_tab, schedule_id)
        )
        self.conn.commit()

    def list_reservations(self, schedule_id, location=None):
        sql = """
            SELECT id, location, item, link, description, media,
                   start_date, end_date, start_time, end_time
              FROM reservation
             WHERE schedule_id = ?
        """
        args = [schedule_id]
        if location and location != "All":
            sql += " AND location = ?"
            args.append(location)
        return self.conn.cursor().execute(sql, args).fetchall()

    def add_location(self, schedule_id, name):
        """
        Insert a location for a schedule, ignoring duplicates.
        """
        c = self.conn.cursor()
        c.execute(
            "INSERT OR IGNORE INTO location(schedule_id, name) VALUES (?, ?)",
            (schedule_id, name)
        )
        self.conn.commit()
        return c.lastrowid

    def list_locations(self, schedule_id):
        c = self.conn.cursor()
        c.execute(
            "SELECT name FROM location WHERE schedule_id = ? ORDER BY name",
            (schedule_id,)
        )
        return [r[0] for r in c.fetchall()]


    def delete_events_for_schedule(self, schedule_id):
        c = self.conn.cursor()
        c.execute("DELETE FROM event WHERE schedule_id=?", (schedule_id,))
        self.conn.commit()

    def insert_event(self,
                     schedule_id,
                     title,
                     description,
                     start_dt,
                     end_dt,
                     duration,
                     link,
                     city,
                     region,
                     event_type,
                     cost,
                     color,
                     x,
                     y,
                     w,
                     h,
                     group_id=None):
        c = self.conn.cursor()
        c.execute("""
            INSERT INTO event(
                schedule_id,
                title,
                description,
                start_dt,
                end_dt,
                duration,
                link,
                city,
                region,
                event_type,
                cost,
                color,
                x,
                y,
                w,
                h,
                group_id
            ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        """, (
            schedule_id,
            title,
            description,
            start_dt.toString(Qt.ISODate),
            end_dt.toString(Qt.ISODate),
            duration,
            link,
            city,
            region,
            event_type,
            cost,
            color,
            x,
            y,
            w,
            h,
            group_id
        ))
        self.conn.commit()
        return c.lastrowid

    def update_event(self, event_id, **fields):
        sets, vals = [], []
        for key, val in fields.items():
            sets.append(f"{key} = ?")
            if isinstance(val, QDateTime):
                vals.append(val.toString(Qt.ISODate))
            elif isinstance(val, QColor):
                vals.append(val.name())
            else:
                vals.append(val)
        vals.append(event_id)
        stmt = f"UPDATE event SET {', '.join(sets)} WHERE id = ?"
        self.conn.cursor().execute(stmt, vals)
        self.conn.commit()

    def list_schedules(self):
        return [r[0] for r in self.conn.cursor().execute("SELECT name FROM schedule")]

    def update_eat_item(self, eat_id, **fields):
        sets, vals = [], []
        for col, val in fields.items():
            sets.append(f"{col} = ?")
            vals.append(val)
        vals.append(eat_id)
        stmt = f"UPDATE eat_item SET {', '.join(sets)} WHERE id = ?"
        self.conn.cursor().execute(stmt, vals)
        self.conn.commit()

    def update_do_item(self, do_id, **fields):
        sets, vals = [], []
        for col, val in fields.items():
            sets.append(f"{col} = ?")
            vals.append(val)
        vals.append(do_id)
        stmt = f"UPDATE do_item SET {', '.join(sets)} WHERE id = ?"
        self.conn.cursor().execute(stmt, vals)
        self.conn.commit()

    def add_eat_item(self, schedule_id, location, item, link, description, media):
        c = self.conn.cursor()
        c.execute("""
          INSERT INTO eat_item(schedule_id, location, item, link, description, media)
          VALUES (?, ?, ?, ?, ?, ?)
        """, (schedule_id, location, item, link, description, media))
        self.conn.commit()
        return c.lastrowid

    def list_event_cities(self, schedule_id):
        """
        Return all distinct, non‐empty City values for a given schedule.
        """
        c = self.conn.cursor()
        c.execute("""
            SELECT DISTINCT city
              FROM event
             WHERE schedule_id = ?
               AND city IS NOT NULL
               AND city != ''
        """, (schedule_id,))
        return [row[0] for row in c.fetchall()]

    def list_eat_items(self, schedule_id, location=None):
        """
        Returns rows in the order your table headers expect:
        [Item, Location, Link, Description, Media]
        """
        q = """
        SELECT id,
               item,
               location,
               link,
               description,
               media
          FROM eat_item
         WHERE schedule_id = ?
        """
        args = [schedule_id]
        if location and location != "All":
            q += " AND location = ?"
            args.append(location)
        return self.conn.execute(q, args).fetchall()

    def add_do_item(self, schedule_id, location, item, link, description, media):
        """Insert a ‘Thing to Do’ for a given schedule and location."""
        c = self.conn.cursor()
        c.execute("""
            INSERT INTO do_item(
                schedule_id,
                location,
                item,
                link,
                description,
                media
            ) VALUES (?, ?, ?, ?, ?, ?)
        """, (
            schedule_id,
            location,
            item,
            link,
            description,
            media
        ))
        self.conn.commit()
        return c.lastrowid

    def list_do_items(self, schedule_id, location=None):
        """
        Returns rows in the order your table headers expect:
        [Item, Location, Link, Description, Media]
        """
        q = """
        SELECT id,
               item,
               location,
               link,
               description,
               media
          FROM do_item
         WHERE schedule_id = ?
        """
        args = [schedule_id]
        if location and location != "All":
            q += " AND location = ?"
            args.append(location)
        return self.conn.execute(q, args).fetchall()

    # -- Hotels --
    def add_hotel(self, schedule_id, location, item, link,
                  start_date: QDate, end_date: QDate,
                  check_in: QTime, check_out: QTime,
                  breakfast, dinner, half_board,
                  num_rooms, room_types, prepaid, cost):
        c = self.conn.cursor()
        c.execute("""
          INSERT INTO hotel(
            schedule_id, location, item, link,
            start_date, end_date, check_in_time, check_out_time,
            breakfast, dinner, half_board,
            num_rooms, room_types, prepaid, cost
          ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        """, (
            schedule_id, location, item, link,
            start_date.toString(Qt.ISODate),
            end_date.toString(Qt.ISODate),
            check_in.toString("HH:mm"),
            check_out.toString("HH:mm"),
            breakfast, dinner, half_board,
            num_rooms, room_types, int(prepaid), cost
        ))
        self.conn.commit()
        return c.lastrowid

    def list_hotels(self, schedule_id, location=None):
        q = """SELECT id, location, item, link,
        start_date, end_date, check_in_time, check_out_time,
                      breakfast, dinner, half_board,
                      num_rooms, room_types, prepaid, cost
               FROM hotel
               WHERE schedule_id=?"""
        args = [schedule_id]
        if location and location != "All":
            q += " AND location = ?"
            args.append(location)
        return self.conn.cursor().execute(q, args).fetchall()


    def load_schedule(self, name):
        c = self.conn.cursor()
        row = c.execute(
            "SELECT id, start_date, end_date, visible_columns, slot_minutes FROM schedule WHERE name=?", (name,)
        ).fetchone()
        if not row:
            raise KeyError(f"No schedule named {name!r}")
        sid, sd, ed, vis_cols, slot_min = row
        start = QDate.fromString(sd, Qt.ISODate)
        end   = QDate.fromString(ed, Qt.ISODate)
        evs = []
        for r in c.execute("""
                    SELECT id, title, description, group_id,
                           start_dt, end_dt, duration,
                           link, city, region, event_type,
                           cost, color, x, y, w, h
                      FROM event
                     WHERE schedule_id=?
                """, (sid,)):
            (eid, title, desc, gid,
             sdt, edt, duration,
             link, city, region, etype,
             cost, color, x, y, w, h) = r
            evs.append({
                'id': eid,
                'title': title,
                'description': desc,
                'group_id': gid,
                'start_dt': QDateTime.fromString(sdt, Qt.ISODate),
                'end_dt': QDateTime.fromString(edt, Qt.ISODate),
                'duration': duration,
                'link': link,
                'city': city,
                'region': region,
                'type': etype,
                'cost': cost,
                'color': QColor(color),
                'x': x,
                'y': y,
                'w': w,
                'h': h
            })
        return sid, start, end, vis_cols, slot_min, evs

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFontMetrics, QTextOption
from PyQt5.QtWidgets import QGraphicsRectItem, QGraphicsTextItem

class EventItem(QGraphicsRectItem):
    def __init__(self, rect, title="New Event", color=QColor("#FFA")):
        super().__init__(rect)
        self.group_id = None
        self.description = ""
        self.setBrush(QBrush(color))
        self.setFlag(QGraphicsItem.ItemIsMovable, True)
        self.setFlag(QGraphicsItem.ItemIsSelectable, True)
        self.setFlag(QGraphicsItem.ItemSendsGeometryChanges, True)

        self.link = ""
        self.city = ""
        self.region = ""
        self.event_type = ""
        self.cost = 0.0
        self.color = color

        # --- Ghost for drag-preview ---
        self._ghost = QGraphicsRectItem(self.rect())
        pen = QPen(Qt.DashLine)
        pen.setColor(Qt.darkGray)
        self._ghost.setPen(pen)
        self._ghost.setBrush(QBrush(Qt.transparent))
        self._ghost.hide()

        # --- Text item with wrapping + centering ---
        self.text = QGraphicsTextItem(parent=self)
        self.text.setDefaultTextColor(Qt.black)
        # enable word-wrap and horizontal centering
        doc = self.text.document()
        opt = doc.defaultTextOption()
        opt.setWrapMode(QTextOption.WordWrap)
        opt.setAlignment(Qt.AlignHCenter)
        doc.setDefaultTextOption(opt)

        self.linked_items = []
        self._last_pos = self.pos()
        self._syncing = False

        # initial title layout
        self._update_text(title)

    def contextMenuEvent(self, event):
        # 1) build the menu
        menu = QMenu()
        open_link = menu.addAction("Open Link")
        # you can add other actions here:
        # delete_act = menu.addAction("Delete Event")

        # 2) show it at the mouse
        selected = menu.exec_(event.screenPos())

        # 3) handle the choice
        if selected == open_link:
            url = self.link.strip()
            if not url:
                # no link to open
                return
            # ensure scheme
            if not re.match(r'https?://', url):
                url = "http://" + url
            webbrowser.open(url)
            # done!

    def set_title(self, title):
        """Public: update title and re-layout text."""
        self._update_text(title)

    def setRect(self, *args):  # override resize
        super().setRect(*args)
        # re-layout the text to new size
        self._update_text(self.text.toPlainText())

    def _update_text(self, title):
        rect = self.rect()
        avail_w = max(1, rect.width() - 2 * TEXT_MARGIN)
        avail_h = max(1, rect.height() - 2 * TEXT_MARGIN)

        # 1) disable wrapping
        self.text.setTextWidth(-1)

        # 2) set the raw text (at its unscaled size)
        self.text.setPlainText(title)
        self.text.setFont(self.text.font())  # ensure the font is the one you want

        # 3) measure its natural size
        br = self.text.boundingRect()  # unscaled bounding box

        # 4) compute the scale factor to fit both dims (never upscale, only shrink)
        sx = avail_w / br.width() if br.width() > 0 else 1.0
        sy = avail_h / br.height() if br.height() > 0 else 1.0
        scale = min(sx, sy, 1.0)

        # 5) apply the scale uniformly
        self.text.setScale(scale)

        # 6) re‐measure scaled size
        scaled_w = br.width() * scale
        scaled_h = br.height() * scale

        # 7) centre inside the rect
        x = (rect.width() - scaled_w) / 2
        y = (rect.height() - scaled_h) / 2
        self.text.setPos(x, y)

    # --- mouse & movement logic preserved ---
    def mousePressEvent(self, ev):
        self._last_pos = self.pos()
        if not self._ghost.scene():
            self.scene().addItem(self._ghost)
        self._ghost.setRect(self.rect())
        self._ghost.setPos(self.pos())
        self._ghost.show()
        super().mousePressEvent(ev)

    def mouseReleaseEvent(self, ev):
        self._ghost.hide()
        if self._ghost.scene():
            self.scene().removeItem(self._ghost)

        x, y = self.scene().snap_to_grid(self.pos())
        self.setPos(x, y)

        if getattr(self, "db_id", None) is not None:
            # 1) recompute the new start‐datetime from its grid position
            col = int((x - TIME_LABEL_WIDTH) // self.scene().cell_w)
            row = int((y - HEADER_HEIGHT) // self.scene().cell_h)
            day = self.scene().start_date.addDays(col)
            mins = START_HOUR * 60 + row * self.scene().slot_minutes
            new_start = QDateTime(day, QTime(mins // 60, mins % 60))

            # 2) persist x,y AND new start_dt
            self.scene().db.update_event(
                self.db_id,
                x=x,
                y=y,
                start_dt=new_start
            )
            if hasattr(self.scene, "main_window"):
                self.scene.main_window._reload_location_views()

        super().mouseReleaseEvent(ev)

    def itemChange(self, change, value):
        if (change == QGraphicsItem.ItemPositionChange
            and not self._syncing
            and self.scene() is not None):
            new_x, new_y = self.scene().snap_to_grid(value)
            new_pos = QPointF(new_x, new_y)
            delta = new_pos - self._last_pos
            for sib in self.linked_items:
                sib._syncing = True
                raw = sib.pos() + delta
                sx, sy = self.scene().snap_to_grid(raw)
                sib.setPos(sx, sy)
                sib._last_pos = QPointF(sx, sy)
                sib._syncing = False
            self._last_pos = new_pos
            return new_pos
        return super().itemChange(change, value)

class AIPlanDialog(QDialog):
    def __init__(self, parent, starred_places):
        super().__init__(parent)
        self.starred_places = starred_places
        self._build_ui()

    def _build_ui(self):
        self.setWindowTitle("🖋️ AI Plan Input")
        vbox = QVBoxLayout(self)

        # Context
        vbox.addWidget(QLabel("Trip context (city, region):"))
        self.ctxEdit = QLineEdit()
        self.ctxEdit.setPlaceholderText("e.g. Kyoto, Japan")
        vbox.addWidget(self.ctxEdit)

        # Free-form text area (plain QTextEdit)
        self.textEdit = QTextEdit()
        self.textEdit.setPlaceholderText("Enter or build your plan text here…")
        vbox.addWidget(self.textEdit)

        # Quick-insert panel
        quick = QGroupBox("Quick inserts")
        hq = QHBoxLayout(quick)

        # 1) Saved-place button
        loc_box = QVBoxLayout()
        loc_box.addWidget(QLabel("Add starred location:"))
        btn_loc = QPushButton("➕ Insert saved place")
        btn_loc.clicked.connect(self._insert_saved_place)
        loc_box.addWidget(btn_loc)
        hq.addLayout(loc_box)

        # 2) Your keywords + comma + time intervals


        vbox.addWidget(quick)

        # OK / Cancel
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        vbox.addWidget(bb)

    def _insert_saved_place(self):
        dlg = PlaceDialog(self, self.starred_places)

        # every time the user clicks “➕ Add” in that dialog,
        # insert immediately into our QTextEdit
        def _do_insert(snippet: str):
            cursor = self.textEdit.textCursor()
            cursor.insertText(snippet)
            self.textEdit.setTextCursor(cursor)

        dlg.placeSelected.connect(_do_insert)

        # run the dialog — Add can be clicked arbitrarily many times;
        # only OK/Cancel will close it.
        dlg.exec_()

    def _insert_with_comma(self, text: str):
        """Insert `<text>,\n` at cursor so each line ends with a comma."""
        cursor = self.textEdit.textCursor()
        cursor.insertText(f"{text},\n")
        self.textEdit.setTextCursor(cursor)

    def get_context(self) -> str:
        return self.ctxEdit.text().strip()

    def get_text(self) -> str:
        raw = self.textEdit.toPlainText().strip()
        # flatten into single block: replace comma+newline with comma+space
        return re.sub(r',\s*\n', ', ', raw)

class CalendarScene(QGraphicsScene):
    eventCreated = pyqtSignal(object)

    def __init__(self, start_date, end_date,
                 slot_minutes=DEFAULT_SLOT_MINUTES,
                 cell_width=DEFAULT_CELL_WIDTH,
                 cell_height=DEFAULT_CELL_HEIGHT):
        super().__init__()
        self.start_date = start_date; self.end_date = end_date
        self.slot_minutes = slot_minutes
        self.cell_w = cell_width; self.cell_h = cell_height
        self._drawing = False; self._start_pt = None
        self._rubber1 = self._rubber2 = None
        self.draw_background_grid()

    def draw_background_grid(self):
        days = self.start_date.daysTo(self.end_date) + 1
        slots = ((END_HOUR - START_HOUR) * 60) // self.slot_minutes
        total_w = TIME_LABEL_WIDTH + days * self.cell_w
        total_h = HEADER_HEIGHT + slots * self.cell_h
        self.setSceneRect(0, 0, total_w, total_h)
        pen = QPen(Qt.lightGray)
        for i in range(days + 1):
            x = TIME_LABEL_WIDTH + i * self.cell_w
            self.addLine(x, HEADER_HEIGHT, x, HEADER_HEIGHT + slots * self.cell_h, pen)
        for j in range(slots + 1):
            y = HEADER_HEIGHT + j * self.cell_h
            self.addLine(TIME_LABEL_WIDTH, y, TIME_LABEL_WIDTH + days * self.cell_w, y, pen)
        self.addLine(0, HEADER_HEIGHT, total_w, HEADER_HEIGHT, pen)
        self.addLine(TIME_LABEL_WIDTH, 0, TIME_LABEL_WIDTH, total_h, pen)
        for i in range(days):
            date = self.start_date.addDays(i)
            ti = QGraphicsTextItem(date.toString("ddd dd MMM"))
            ti.setDefaultTextColor(Qt.black)
            ti.setPos(TIME_LABEL_WIDTH + i * self.cell_w + TEXT_MARGIN,
                      (HEADER_HEIGHT - ti.boundingRect().height())/2)
            self.addItem(ti)

    def snap_to_grid(self, pos):
        days  = self.start_date.daysTo(self.end_date) + 1
        slots = ((END_HOUR - START_HOUR)*60) // self.slot_minutes

        # figure out which grid‐cell you’re *in*, not nearest
        col = math.floor((pos.x() - TIME_LABEL_WIDTH) / self.cell_w)
        row = math.floor((pos.y() - HEADER_HEIGHT) / self.cell_h)

        # clamp to [0..days-1] and [0..slots-1]
        col = max(0, min(col, days-1))
        row = max(0, min(row, slots-1))

        # build the snapped scene coords
        x = TIME_LABEL_WIDTH + col * self.cell_w
        y = HEADER_HEIGHT       + row * self.cell_h
        return x, y

    def mousePressEvent(self, ev):
        if (ev.button() == Qt.LeftButton and
            ev.modifiers() & Qt.ShiftModifier and
            not self.itemAt(ev.scenePos(), self.views()[0].transform())):
            self._drawing = True
            self._start_pt = self.snap_to_grid(ev.scenePos())
            pen = QPen(Qt.DashLine); pen.setColor(Qt.blue)
            brush = QBrush(QColor(100,100,255,50))
            self._rubber1 = self.addRect(QRectF(*self._start_pt, self.cell_w, self.cell_h), pen, brush)
            self._rubber2 = self.addRect(QRectF(*self._start_pt, self.cell_w, self.cell_h), pen, brush)
            self._rubber2.hide()
        else:
            super().mousePressEvent(ev)

    def mouseMoveEvent(self, ev):
        if self._drawing and self._rubber1:
            sx, sy = self._start_pt
            ex, ey = self.snap_to_grid(ev.scenePos())
            sc = (sx - TIME_LABEL_WIDTH)//self.cell_w
            ec = (ex - TIME_LABEL_WIDTH)//self.cell_w
            slots = ((END_HOUR - START_HOUR)*60)//self.slot_minutes
            top = HEADER_HEIGHT; bottom = HEADER_HEIGHT + slots*self.cell_h
            if ec == sc:
                self._rubber2.hide()
                w = self.cell_w; h = ey - sy + self.cell_h
                self._rubber1.setRect(0,0,w,h); self._rubber1.setPos(sx,sy)
            elif ec > sc:
                h1 = bottom - sy; x1 = TIME_LABEL_WIDTH + sc*self.cell_w
                self._rubber1.setRect(0,0,self.cell_w,h1); self._rubber1.setPos(x1,sy)
                h2 = ey - top + self.cell_h; x2 = TIME_LABEL_WIDTH + ec*self.cell_w
                self._rubber2.setRect(0,0,self.cell_w,h2); self._rubber2.setPos(x2,top)
                self._rubber2.show()
            else:
                h1 = sy - top + self.cell_h; x1 = TIME_LABEL_WIDTH + sc*self.cell_w
                self._rubber1.setRect(0,0,self.cell_w,h1); self._rubber1.setPos(x1,top)
                h2 = bottom - ey; x2 = TIME_LABEL_WIDTH + ec*self.cell_w
                self._rubber2.setRect(0,0,self.cell_w,h2); self._rubber2.setPos(x2,ey)
                self._rubber2.show()
        else:
            super().mouseMoveEvent(ev)

    def mouseReleaseEvent(self, ev):
        if self._drawing and self._rubber1:
            # clean up preview
            self.removeItem(self._rubber1)
            self.removeItem(self._rubber2)
            self._rubber1 = self._rubber2 = None
            self._drawing = False

            SX, SY = self._start_pt
            ex, ey = self.snap_to_grid(ev.scenePos())
            sc = (SX - TIME_LABEL_WIDTH) // self.cell_w
            ec = (ex - TIME_LABEL_WIDTH) // self.cell_w

            # compute scene bounds
            top = HEADER_HEIGHT
            bottom = HEADER_HEIGHT + (((END_HOUR - START_HOUR)*60)//self.slot_minutes)*self.cell_h

            def make_piece(x, y, w, h, gid=None):
                e = EventItem(QRectF(0,0,w,h), "New Event", QColor("#FFA"))
                e.setPos(x, y)
                e.group_id = gid
                self.addItem(e)
                self.eventCreated.emit(e)
                return e

            # single‐day
            if ec == sc:
                height = ey - SY + self.cell_h
                make_piece(SX, SY, self.cell_w, height)

            # multi‐day: one piece per column
            else:
                gid = str(uuid.uuid4())
                pieces = []
                # ensure we handle dragging left‐to‐right or right‐to‐left
                start_col, end_col = sorted((sc, ec))
                for col in range(start_col, end_col+1):
                    x = TIME_LABEL_WIDTH + col * self.cell_w

                    if col == sc:
                        # first day: from SY down to bottom
                        y = SY
                        h = bottom - SY
                    elif col == ec:
                        # last day: from top down to ey
                        y = top
                        h = ey - top + self.cell_h
                    else:
                        # full‐day middle spans
                        y = top
                        h = bottom - top

                    pieces.append(make_piece(x, y, self.cell_w, h, gid))

                # link them all together
                for p in pieces:
                    p.linked_items = [o for o in pieces if o is not p]

        else:
            super().mouseReleaseEvent(ev)

class DateRangeDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Select Date Range")
        layout = QVBoxLayout(self)
        self.cal_from = QCalendarWidget()
        self.cal_to   = QCalendarWidget()
        self.cal_to.setSelectedDate(self.cal_from.selectedDate().addDays(1))
        layout.addWidget(self.cal_from)
        layout.addWidget(self.cal_to)
        btns = QDialogButtonBox(QDialogButtonBox.Ok|QDialogButtonBox.Cancel)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)

    def dates(self):
        return self.cal_from.selectedDate(), self.cal_to.selectedDate()

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Holiday Planner")
        self.db = DatabaseManager()
        self.current_schedule_id = None
        self.current_event = None
        self._handling_selection = False

        self.gmaps = gmaps_client

        self.visible_columns    = DEFAULT_VISIBLE_COLUMNS
        self.slot_minutes       = DEFAULT_SLOT_MINUTES
        self.sidePanelWidth     = 300
        self.settingsPanelWidth = 200



        self._init_ui()
        today = QDate.currentDate()
        self.init_calendar(today, today.addDays(13))

        SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
        creds = service_account.Credentials.from_service_account_file(
            cfg['google']['credentials_dir'],
            scopes=SCOPES
        )
        self.sheets_service = build('sheets', 'v4', credentials=creds)

        # also build a Drive service with at least the drive.file scope
        DRIVE_SCOPES = ['https://www.googleapis.com/auth/drive.file']
        drive_creds = service_account.Credentials.from_service_account_file(
            cfg['google']['credentials_dir'],
            scopes=DRIVE_SCOPES
        )
        drive_service = build('drive', 'v3', credentials=drive_creds)

        def share_with_service_account(spreadsheet_id):
            sa_email = drive_creds.service_account_email
            drive_service.permissions().create(
                fileId=spreadsheet_id,
                body={
                    'type': 'user',
                    'role': 'writer',
                    'emailAddress': sa_email
                },
                fields='id'
            ).execute()

        # — load starred places from JSON file —

        try:
            self.starred_places = self.load_starred_places(PLACES_FILE)
        except Exception as e:
            QMessageBox.critical(self, "Load Error",
                                 f"Could not read Saved Places.json:\n{e}")
            self.starred_places = []

    def _init_ui(self):
        # — build the two overlays first —
        self.detailsPanel = QFrame()
        self.detailsPanel.setFrameShape(QFrame.StyledPanel)
        self.detailsPanel.setMinimumWidth(self.sidePanelWidth)
        self.detailsPanel.setStyleSheet("background-color: rgba(255,255,255,230);")
        # … your old detailsPanel layout code here …

        dp_layout = QVBoxLayout(self.detailsPanel)
        dp_layout.setContentsMargins(8, 8, 8, 8)

        form = QFormLayout()
        self.titleEdit = QLineEdit();
        form.addRow("Title:", self.titleEdit)
        self.startEdit = QDateTimeEdit();
        form.addRow("Start:", self.startEdit);
        self.startEdit.setCalendarPopup(True)
        self.endEdit = QDateTimeEdit();
        form.addRow("End:", self.endEdit);
        self.endEdit.setCalendarPopup(True)
        self.durationSpin = QSpinBox();
        form.addRow("Duration:", self.durationSpin)
        self.durationSpin.setSuffix(" min");
        self.durationSpin.setRange(1, 24 * 60)
        self.linkEdit = LinkableLineEdit()
        self.linkLabel = ClickableLabel("Link:")
        self.linkLabel.setCursor(Qt.PointingHandCursor)
        # when it’s clicked, grab whatever’s in the QLineEdit and open it
        self.linkLabel.clicked.connect(self._on_open_link)
        form.addRow(self.linkLabel, self.linkEdit)
        self.cityEdit = QLineEdit();
        form.addRow("City:", self.cityEdit)
        self.regionEdit = QLineEdit();
        self.descriptionEdit = QLineEdit()
        form.addRow("Description:", self.descriptionEdit)
        self.descriptionEdit.editingFinished.connect(self.update_event_description)
        form.addRow("Region:", self.regionEdit)
        self.typeEdit = QLineEdit();
        form.addRow("Type:", self.typeEdit)
        self.costEdit = QDoubleSpinBox();
        form.addRow("Cost:", self.costEdit);
        self.costEdit.setPrefix("£")
        self.colorBtn = QPushButton("Choose color…");
        form.addRow("Color:", self.colorBtn)

        dp_layout.addLayout(form)
        dp_layout.addStretch()

        self.settingsPanel = QFrame()
        self.settingsPanel.setFrameShape(QFrame.StyledPanel)
        self.settingsPanel.setMinimumWidth(self.settingsPanelWidth)
        self.settingsPanel.setStyleSheet("background-color: rgba(255,255,255,230);")
        # … your old settingsPanel layout code here …

        sp_form = QFormLayout()
        self.daysSpin = QSpinBox();
        self.daysSpin.setRange(1, 30);
        self.daysSpin.setValue(self.visible_columns)
        sp_form.addRow("Visible days:", self.daysSpin)
        self.slotSpin = QSpinBox();
        self.slotSpin.setRange(1, 240);
        self.slotSpin.setSuffix(" min");
        self.slotSpin.setValue(self.slot_minutes)
        sp_form.addRow("Slot interval:", self.slotSpin)

        applyBtn = QPushButton("Apply")
        applyBtn.clicked.connect(self.apply_settings)

        sp_layout = QVBoxLayout(self.settingsPanel)
        sp_layout.addLayout(sp_form)
        sp_layout.addWidget(applyBtn)
        sp_layout.addStretch()

        # — Tier 1: Global toolbar —
        self.tier1_toolbar = QToolBar("Global")
        self.addToolBar(Qt.TopToolBarArea, self.tier1_toolbar)
        for txt, slot in [
            ("New Schedule", self.on_new_schedule),
            ("Save Schedule", self.on_save_schedule),
            ("Load Schedule", self.on_load_schedule),
        ]:
            act = QAction(txt, self)
            act.triggered.connect(slot)
            self.tier1_toolbar.addAction(act)

        # — Tier 2: View tabs —
        self.tab_widget = QTabWidget()
        self.setCentralWidget(self.tab_widget)
        self.general_tab = QWidget()
        self.schedule_tab = QWidget()
        self.location_tab = QWidget()
        self.tab_widget.addTab(self.general_tab, "General")
        self.tab_widget.addTab(self.schedule_tab, "Schedule")
        self.tab_widget.addTab(self.location_tab, "Location")
        self.tab_widget.currentChanged.connect(self.on_tab_changed)

        # force the next toolbar onto a new row
        self.addToolBarBreak(Qt.TopToolBarArea)

        # — Tier 3: Context toolbar —
        self.context_toolbar = QToolBar("Context")
        self.context_toolbar.setToolButtonStyle(Qt.ToolButtonTextOnly)
        self.addToolBar(Qt.TopToolBarArea, self.context_toolbar)

        # Build each page
        self._init_general_tab()
        self._init_schedule_tab()
        self._init_location_tab()

        # Set initial context
        self.on_tab_changed(self.tab_widget.currentIndex())

        self.titleEdit.editingFinished.connect(self.update_event_title)
        self.linkEdit.editingFinished.connect(self.update_event_link)
        self.cityEdit.editingFinished.connect(self.update_event_city)
        self.regionEdit.editingFinished.connect(self.update_event_region)
        self.typeEdit.editingFinished.connect(self.update_event_type)
        self.costEdit.valueChanged.connect(self.update_event_cost)
        self.colorBtn.clicked.connect(self.choose_event_color)
        self.startEdit.dateTimeChanged.connect(self.update_event_time)
        self.endEdit.dateTimeChanged.connect(self.update_event_time)
        self.durationSpin.valueChanged.connect(self.update_event_duration)

        delete_act = QAction("Delete Event", self)
        delete_act.setShortcut(QKeySequence.Delete)
        delete_act.triggered.connect(self.on_delete_selected_events)
        # make sure MainWindow catches the shortcut even if the toolbar isn't focused
        self.addAction(delete_act)

    def _on_open_link(self):
        url = self.linkEdit.text().strip()
        if not url:
            QMessageBox.warning(self, "No link", "There’s no URL to open.")
            return

        # ensure it has a scheme
        if not re.match(r'https?://', url):
            url = 'http://' + url

        try:
            webbrowser.open(url)
        except Exception as e:
            QMessageBox.critical(self, "Couldn’t open link", f"Error: {e}")

    def on_import_text_items(self):
        if self.current_schedule_id is None:
            QMessageBox.warning(self, "No Schedule", "Please create or load a schedule first.")
            return

        # let the user pick (or type) a location
        locs = self.db.list_locations(self.current_schedule_id)
        dlg = TextImportDialog(self, locations=locs)
        if dlg.exec_() != QDialog.Accepted:
            return

        raw_text = dlg.get_text()
        context_loc = dlg.get_location().strip() or "Unknown Location"
        if not raw_text:
            QMessageBox.warning(self, "No Text", "Please paste or load some text.")
            return

        # summarize & extract via AI (your existing code)…
        summary = self._summarize_text_via_ai(raw_text)
        items = self._extract_items_via_ai(summary, context_location=context_loc)

        # insert each location into the 'location' table, then the detail tables:
        for obj in items.get("things_to_do", []):
            loc = obj["location"]
            self.db.add_location(self.current_schedule_id, loc)
            self.db.add_do_item(
                self.current_schedule_id,
                loc,
                obj["item"],
                obj.get("link", ""),
                obj.get("description", ""),
                obj.get("media", "")
            )
        for obj in items.get("things_to_eat", []):
            loc = obj["location"]
            self.db.add_location(self.current_schedule_id, loc)
            self.db.add_eat_item(
                self.current_schedule_id,
                loc,
                obj["item"],
                obj.get("link", ""),
                obj.get("description", ""),
                obj.get("media", "")
            )

        # refresh the location‐filtered tables & dropdown
        self._populate_location_dropdown()
        self._reload_location_views()

        QMessageBox.information(self, "Done", "Imported things to do and things to eat.")

    def _summarize_text_via_ai(self, text: str) -> str:
        """
        Use a cheaper model to produce an unconstrained summary that:
          • Mentions each vendor by name and the exact dish they tried (with price if stated).
          • Calls out any non-vendor local specialties (e.g. “Peaches are a Kagoshima specialty”).
          • Notes unique activities, sites, or points of interest.
          • Preserves street names, districts, and context needed for location.
        Returns the raw summary text—no JSON, no fences.
        """
        SYSTEM = {
            "role": "system",
            "content": (
                "You are a detail-oriented summarizer.  "
                "Read the user’s transcript and produce a clear, well-structured summary that captures:\n"
                "  1. Each vendor name and the exact dish they tried (with price if stated), as “Vendor – Dish (price)”.\n"
                "  2. Any mentions of local specialty foods not tied to a vendor.\n"
                "  3. Unique activities, sights, or points of interest visited.\n"
                "  4. All street names, districts, and any context needed for location.\n"
                "Do not limit the length—include every relevant detail."
            )
        }
        CONTEXT = {
            "role": "system",
            "content": (
                "Return a plain-text summary.  "
                "Use bullet points or numbered lists for clarity.  "
                "Do not add headings, JSON, or markdown fences."
            )
        }
        USER = {"role": "user", "content": text}

        resp = openai.ChatCompletion.create(
            model="gpt-3.5-turbo",
            temperature=0.3,
            messages=[SYSTEM, CONTEXT, USER]
        )
        summary = resp.choices[0].message.content.strip()
        print(f"Summary:\n{summary}\n")
        return summary

    def _extract_items_via_ai(self, text: str, context_location: str) -> dict:
        """
        Parse the user's summary into a JSON object.  Every 'things_to_eat' and
        'things_to_do' entry will share the same location (context_location), have
        a brief description sentence, and never appear as bare strings.
        """
        SYSTEM = {
            "role": "system",
            "content": (
                "You are a strict JSON parser.  The user has provided a summary of street food "
                "and activities in one place.  **All** entries—both things_to_eat and things_to_do—"
                f"must have their `location` set to exactly \"{context_location}\".  "
                "Fill every `description` field with a concise sentence describing the dish "
                "or activity.  For `things_to_eat`, `item` must be “Vendor – Dish (price if given)” "
                "or “Local Specialty – X”.  For `things_to_do`, `item` is just the name of the activity.  "
                "Return **only** valid JSON with this exact schema and no extra keys:\n\n"
                "{\n"
                "  \"summary\": string,\n"
                "  \"things_to_eat\": [\n"
                "    {\"location\":string, \"item\":string, \"description\":string, "
                "\"link\":string, \"media\":string}\n"
                "  ],\n"
                "  \"things_to_do\": [\n"
                "    {\"location\":string, \"item\":string, \"description\":string, "
                "\"link\":string, \"media\":string}\n"
                "  ]\n"
                "}"
            )
        }
        CONTEXT = {
            "role": "system",
            "content": (
                "• Always output one JSON object.  No markdown fences, no explanations.\n"
                "• If any `link` or `media` is not provided in the summary, use an empty string.\n"
                "• Do not output any array element as a bare string—wrap everything in an object."
            )
        }
        USER = {"role": "user", "content": text}

        resp = openai.ChatCompletion.create(
            model="gpt-4o",
            temperature=0,
            messages=[SYSTEM, CONTEXT, USER]
        )
        raw = resp.choices[0].message.content.strip()
        # strip ```json fences if present
        m = re.search(r'```(?:json)?\n(.+?)```', raw, re.S)
        payload = (m.group(1) if m else raw).strip()

        try:
            data = json.loads(payload)
        except JSONDecodeError as e:
            raise RuntimeError(f"AI returned invalid JSON:\n{payload}") from e

        # enforce the shared location
        for section in ("things_to_eat", "things_to_do"):
            for obj in data.get(section, []):
                obj["location"] = context_location

        return data

    def on_add_event_files(self):
        if self.current_schedule_id is None:
            QMessageBox.warning(
                self, "No Schedule",
                "Please create or load a schedule before importing bookings."
            )
            return

        paths, _ = QFileDialog.getOpenFileNames(
            self, "Select .eml or .txt files", "",
            "Email Files (*.eml);;Text Files (*.txt)"
        )
        if not paths:
            return

        for path in paths:
            ext = os.path.splitext(path)[1].lower()
            try:
                if ext == ".eml":
                    text = self.parse_email_file(path)
                else:
                    with open(path, "r", encoding="utf-8") as f:
                        text = f.read()
            except Exception as e:
                QMessageBox.critical(self, "Import Error",
                                     f"Failed to read {os.path.basename(path)}:\n{e}")
                continue

            # parse via AI (your existing code) to get `doc`
            doc = self._parse_booking_via_ai(path, text)
            if doc.get("type") == "hotel":
                self.db.add_hotel(
                    self.current_schedule_id,
                    doc["location"], doc["item"], doc["link"],
                    QDate.fromString(doc["start_date"], Qt.ISODate),
                    QDate.fromString(doc["end_date"],   Qt.ISODate),
                    QTime.fromString(doc["check_in_time"],  "HH:mm"),
                    QTime.fromString(doc["check_out_time"], "HH:mm"),
                    doc.get("breakfast",0),
                    doc.get("dinner",0),
                    doc.get("half_board",0),
                    doc.get("num_rooms",1),
                    doc.get("room_types",""),
                    doc.get("prepaid",False),
                    doc.get("cost",0.0)
                )
            else:  # reservation
                self.db.add_reservation(
                    self.current_schedule_id,
                    doc["location"], doc["item"], doc["link"],
                    doc.get("description",""), doc.get("media",""),
                    QDate.fromString(doc["start_date"], Qt.ISODate),
                    QDate.fromString(doc["end_date"],   Qt.ISODate),
                    QTime.fromString(doc["start_time"], "HH:mm"),
                    QTime.fromString(doc["end_time"],   "HH:mm")
                )

        # once everything’s in the hotel/reservation tables, re-sync the calendar:
        self.sync_bookings_to_schedule()
        QMessageBox.information(
            self,
            "Import Complete",
            "All events have been imported and the calendar is now up to date."
        )

    @staticmethod
    def extract_spreadsheet_id(url_or_id):
        """Pull the Spreadsheet ID out of a full URL or return it unchanged."""
        m = re.search(r'/spreadsheets/d/([a-zA-Z0-9-_]+)', url_or_id)
        return m.group(1) if m else url_or_id

    def on_sync_sheets(self):
        """Sync to Google Sheets, but only prompt once per schedule."""
        if self.current_schedule_id is None:
            QMessageBox.warning(self, "No Schedule", "Please load or save a schedule first.")
            return

        # 1) see if we've already bound a sheet ↔ schedule
        row = self.db.conn.execute(
            "SELECT spreadsheet_id, sheet_tab FROM schedule WHERE id=?",
            (self.current_schedule_id,)
        ).fetchone()

        if row and row[0] and row[1]:
            # already have one saved → just auto‐sync
            try:
                self._auto_sync_to_sheet()
                QMessageBox.information(self, "Sync Complete", "Data synced to Google Sheets successfully.")
            except Exception as e:
                QMessageBox.critical(self, "Sync Failed", f"Error syncing to Sheets:\n{e}")
            return

        # 2) otherwise prompt for a new binding
        raw, ok = QInputDialog.getText(
            self, "Spreadsheet", "Enter Google Sheets URL or ID:"
        )
        if not ok or not raw.strip():
            return
        sheet_name, ok2 = QInputDialog.getText(
            self, "Sheet Name", "Enter tab name:", text="Sheet1"
        )
        if not ok2 or not sheet_name.strip():
            return

        ssid = self.extract_spreadsheet_id(raw.strip())
        # 3) save that binding
        self.db.save_sheet_info(self.current_schedule_id, ssid, sheet_name)
        # 4) and immediately do your first push
        try:
            self._auto_sync_to_sheet()
            QMessageBox.information(self, "Bound & Synced", "Sheet linked and data pushed successfully.")
        except Exception as e:
            QMessageBox.critical(self, "Sync Failed", f"Error syncing to Sheets:\n{e}")

    def _auto_sync_to_sheet(self):
        # load schedule → get its saved sheet info
        row = self.db.conn.execute(
            "SELECT spreadsheet_id,sheet_tab FROM schedule WHERE id=?",
            (self.current_schedule_id,)
        ).fetchone()
        if not row or not row[0] or not row[1]:
            return  # nothing bound yet
        ssid, tab = row
        self.update_google_sheet(ssid, tab, start_cell="C4")

    def update_google_sheet_by_header(self,
                                      spreadsheet_id: str,
                                      sheet_name: str,
                                      data_dicts: list[dict],
                                      *,
                                      start_cell: str = "A1"):
        # parse “C4” →  col_letter="C", header_row=4
        import re
        m = re.match(r"([A-Z]+)(\d+)", start_cell)
        col_letter, header_row = m.group(1), int(m.group(2))
        # 1) Read the header row at that offset
        header_resp = self.sheets_service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=f"{sheet_name}!{col_letter}{header_row}:{col_letter}{header_row}"
        ).execute()
        headers = header_resp.get("values", [[]])[0]

        # 2) Build rows
        values = [[row.get(h, "") for h in headers] for row in data_dicts]

        # 3) Clear everything below your header
        self.sheets_service.spreadsheets().values().clear(
            spreadsheetId=spreadsheet_id,
            range=f"{sheet_name}!{col_letter}{header_row + 1}:{col_letter}1000"
        ).execute()

        # 4) Write new rows, starting one row below your header
        body = {"values": values}
        last_col = chr(ord(col_letter) + len(headers) - 1)
        last_row = header_row + len(values)
        write_range = f"{sheet_name}!{col_letter}{header_row + 1}:{last_col}{last_row}"
        self.sheets_service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=write_range,
            valueInputOption="RAW",
            body=body
        ).execute()

    import re

    def update_google_sheet(self, spreadsheet_id, sheet_name, start_cell="A1"):
        """Fetch your schedule, format it, then overwrite the given sheet
        starting at `start_cell` (e.g. "C4")."""
        # 1) load & sort the canonical events
        _, _, _, _, _, evs = self.db.load_schedule(self.current_schedule_name)
        evs_sorted = sorted(evs, key=lambda ev: (ev['start_dt'], ev['title']))

        # 2) build the 2D array (including header)
        headers = ["Date", "Time", "Activity", "Location", "Link", "Cost"]
        values = [headers]
        for ev in evs_sorted:
            sd, ed = ev['start_dt'], ev['end_dt']
            date_str = sd.toString("ddd, M/d/yy")
            time_str = sd.toString("HH:mm")
            activity = f"From {sd.toString('HH:mm')} to {ed.toString('HH:mm')} - {ev['title']}"

            # defaults from the event record
            location = ev['city'] + (f", {ev['region']}" if ev['region'] else "")
            link = ev['link']
            cost_val = ev['cost']

            # if this came from a hotel or reservation, override from that table
            gid = ev.get('group_id')
            if gid:
                if ev['type'] == "hotel":
                    rec = self.db.conn.execute(
                        "SELECT location, link, cost FROM hotel WHERE id=?",
                        (gid,)
                    ).fetchone()
                    if rec:
                        location, link, cost_val = rec[0], rec[1] or link, rec[2]
                elif ev['type'] == "reservation":
                    rec = self.db.conn.execute(
                        "SELECT location, link FROM reservation WHERE id=?",
                        (gid,)
                    ).fetchone()
                    if rec:
                        location, link = rec[0], rec[1] or link

            cost_str = f"£{cost_val:.2f}"
            values.append([date_str, time_str, activity, location, link, cost_str])

        # 3) parse start_cell into column letters and row number
        m = re.match(r"^([A-Z]+)(\d+)$", start_cell)
        if not m:
            start_col, start_row = "A", 1
        else:
            start_col, start_row = m.group(1), int(m.group(2))

        # 4) compute end column (assumes your sheet has <= 26 cols; adjust if you need AA, etc.)
        num_cols = len(headers)
        end_col = chr(ord(start_col) + num_cols - 1)
        end_row = start_row + len(values) - 1

        target_range = f"{sheet_name}!{start_col}{start_row}:{end_col}{end_row}"

        # 5) clear only that exact rectangle
        self.sheets_service.spreadsheets().values().clear(
            spreadsheetId=spreadsheet_id,
            range=target_range
        ).execute()

        # 6) push your new block into that same range
        body = {"values": values}
        self.sheets_service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=target_range,
            valueInputOption="RAW",
            body=body
        ).execute()

    def sync_bookings_to_schedule(self):
        sid = self.current_schedule_id
        if sid is None:
            return

        # 1) remove old hotel & reservation events
        self.db.conn.execute(
            "DELETE FROM event WHERE schedule_id = ?"
            "  AND event_type IN ('hotel','reservation')",
            (sid,)
        )
        self.db.conn.commit()

        # 2) clear them from the scene
        for item in list(self.scene.items()):
            if isinstance(item, EventItem) and item.event_type in ('hotel', 'reservation'):
                self.scene.removeItem(item)

        # 3) re-insert hotels (check-in / check-out)
        for (hid, loc, itm, link,
             sd, ed, cin, cout,
             *_) in self.db.list_hotels(sid):
            self._insert_event_from_booking(
                sid, hid, f"Check-In at {itm}", sd, cin, "hotel"
            )
            self._insert_event_from_booking(
                sid, hid, f"Check-Out of {itm}", ed, cout, "hotel"
            )

        # 4) re-insert reservations as full‐length events
        for (rid, loc, itm, link, desc, media,
             sd, ed, st, et) in self.db.list_reservations(sid):
            # treat each reservation as one event spanning start→end
            self._insert_event_from_booking(
                sid, rid,
                title=itm,
                date=sd,
                time=st,
                event_type="reservation",
                description=desc,
                media=media,
                end_date=ed,
                end_time=et
            )

        # 5) optional: let user know

    def on_reservation_cell_changed(self, row, col):
        # 1) pull the hidden reservation_id
        cell = self.resTable.item(row, 0)
        res_id = cell.data(Qt.UserRole)
        if res_id is None:
            return

        # 2) map header text → DB column
        header = self.resTable.horizontalHeaderItem(col).text()
        mapping = {
            "Location": "location",
            "Item": "item",
            "Link": "link",
            "Description": "description",
            "Media": "media",
            "Start Date": "start_date",
            "End Date": "end_date",
            "Start Time": "start_time",
            "End Time": "end_time",
        }
        field = mapping.get(header)
        if not field:
            return

        raw = self.resTable.item(row, col).text()

        # 3) parse into the right type or reject
        if field in ("start_date", "end_date"):
            qd = QDate.fromString(raw, Qt.ISODate)
            if not qd.isValid():
                QMessageBox.warning(self, "Invalid Date",
                                    f"'{raw}' is not a valid YYYY-MM-DD. Reverting.")
                self._reload_location_views()
                return
            val = qd

        elif field in ("start_time", "end_time"):
            qt = QTime.fromString(raw, "HH:mm")
            if not qt.isValid():
                QMessageBox.warning(self, "Invalid Time",
                                    f"'{raw}' is not a valid HH:MM. Reverting.")
                self._reload_location_views()
                return
            val = qt

        else:
            val = raw

        # 4) persist to DB
        try:
            self.db.update_reservation(res_id, **{field: val})
        except Exception as e:
            QMessageBox.critical(self, "Update Error",
                                 f"Failed to save change to {header}:\n{e}")
            return

        # 5) mirror that change back into the calendar
        self.sync_bookings_to_schedule()
        # 6) ensure all tables stay in sync
        self._reload_location_views()
    def _insert_event_from_booking(self, schedule_id, booking_id,
                                   title, date, time, event_type,
                                   description="", media="",
                                   end_date=None, end_time=None):
        """
        date, time, end_date, end_time may be either QDate/QTime or ISO-formatted strings.
        This method will coerce them, then insert into the DB.event table and
        drop an EventItem into the calendar view.
        """
        # 1) convert strings → QDate/QTime
        if isinstance(date, str):
            date = QDate.fromString(date, Qt.ISODate)
        if isinstance(time, str):
            time = QTime.fromString(time, "HH:mm")
        if end_date is not None and isinstance(end_date, str):
            end_date = QDate.fromString(end_date, Qt.ISODate)
        if end_time is not None and isinstance(end_time, str):
            end_time = QTime.fromString(end_time, "HH:mm")

        # 2) build QDateTime start_dt & end_dt, compute duration
        start_dt = QDateTime(date, time)
        if end_date and end_time:
            end_dt = QDateTime(end_date, end_time)
            duration = max(15, start_dt.secsTo(end_dt) // 60)
        else:
            duration = DEFAULT_SLOT_MINUTES
            end_dt  = start_dt.addSecs(duration * 60)

        # 3) insert into `event` table
        # choose color per type

        if event_type == "hotel":
            col_str = "#00ff00"
        elif event_type == "reservation":
            col_str = "#55ffff"
        else:
            col_str = "#FFA"

        x, y, w, h = self._geometry_for(start_dt, duration)
        db_id = self.db.insert_event(
            schedule_id,
            title,
            description,
            start_dt,
            end_dt,
            duration,
            "",     # link
            "", "", # city, region
            event_type,
            0.0,    # cost
            col_str, # color
            x, y, w, h,
            group_id=str(booking_id)
        )

        # 4) drop it into the scene
        rect = QRectF(0, 0, w, h)
        item = EventItem(rect, title, QColor(col_str))
        item.db_id      = db_id
        item.event_type = event_type
        item.setPos(x, HEADER_HEIGHT + y)
        self.scene.addItem(item)

    def _parse_booking_via_ai(self, file_name: str, text: str) -> dict:
        """
        Send file_name & text to GPT-4, extract only the JSON payload,
        and raise a RuntimeError if it isn’t valid JSON.
        """
        SYSTEM = {
            "role": "system",
            "content": (
                "You are a booking-parser. The user sends you a JSON with keys:\n"
                "  • file_name: string\n"
                "  • text: booking details (plain text)\n\n"
                "Extract values and return exactly one JSON matching either the hotel or reservation schema. "
                "If a value is missing, use only the default \"\", 0 or false.\n\n"
                "--- Hotel schema ---\n"
                "{ \"type\":\"hotel\",\"location\":string,\"item\":string,"
                "\"link\":string,\"start_date\":\"YYYY-MM-DD\",\"end_date\":\"YYYY-MM-DD\","
                "\"check_in_time\":\"HH:MM\",\"check_out_time\":\"HH:MM\","
                "\"breakfast\":boolean,\"dinner\":boolean,\"half_board\":boolean,"
                "\"num_rooms\":number,\"room_types\":string,"
                "\"prepaid\":boolean,\"cost\":number }\n\n"
                "--- Reservation schema ---\n"
                "{ \"type\":\"reservation\",\"location\":string,\"item\":string,"
                "\"link\":string,\"description\":string,\"media\":string,"
                "\"start_date\":\"YYYY-MM-DD\",\"end_date\":\"YYYY-MM-DD\","
                "\"start_time\":\"HH:MM\",\"end_time\":\"HH:MM\" }\n\n"
                "Return *only* the JSON object—no prose, no markdown fences."
            )
        }
        CONTEXT = {
            "role": "system",
            "content": "Fill each field with the actual value from `text`; do not leave defaults unless truly absent."
        }
        USER = {"role": "user", "content": json.dumps({"file_name": file_name, "text": text})}

        resp = openai.ChatCompletion.create(
            model="gpt-4o",
            temperature=0,
            messages=[SYSTEM, CONTEXT, USER]
        )
        raw = resp.choices[0].message.content

        # 1) Strip markdown fences if present
        m = re.search(r'```(?:json)?\n(.+?)```', raw, re.S)
        payload = m.group(1) if m else raw

        # 2) Trim whitespace
        payload = payload.strip()

        # 3) Parse
        try:
            return json.loads(payload)
        except JSONDecodeError as e:
            # Include the invalid payload so you can inspect why it failed
            raise RuntimeError(f"AI returned invalid JSON:\n{payload}") from e

    def _geometry_for(self, dt: QDateTime, duration_minutes: int):
        """Return (x, y, w, h) in scene coords for a QDateTime + duration."""
        col = self.scene.start_date.daysTo(dt.date())
        row = ((dt.time().hour()*60 + dt.time().minute())
               - START_HOUR*60) // self.scene.slot_minutes
        x = TIME_LABEL_WIDTH + col * self.scene.cell_w
        y = row * self.scene.cell_h
        w = self.scene.cell_w
        h = (duration_minutes / self.scene.slot_minutes) * self.scene.cell_h
        return x, y, w, h


    def pdf_to_chunks(self, path, model="gpt-4o", chunk_size=8000):
        # 1) Grab the full raw text from the PDF
        raw = extract_text(path) or ""
        raw = raw.strip()
        if not raw:
            raise RuntimeError(f"No text extracted from PDF {path}")

        # 2) Split on form-feed (page break) or just by length
        pages = raw.split("\f")  # pdfminer uses \f between pages

        # 3) Tokenize & chunk
        enc = tiktoken.encoding_for_model(model)
        chunks, current = [], ""
        for pg in pages:
            candidate = f"{current}\n\n{pg}".strip()
            if len(enc.encode(candidate)) > chunk_size:
                chunks.append(current)
                current = pg
            else:
                current = candidate
        if current:
            chunks.append(current)
        print(f"[DEBUG] PDF extracted text (first 500 chars):\n{raw[:500]}\n")
        print(f"[DEBUG] PDF chunks count: {len(chunks)}")
        return chunks

    def on_delete_selected_events(self):
        # 1) find selected EventItem(s)
        items = [it for it in self.scene.selectedItems()
                 if isinstance(it, EventItem)]
        if not items:
            return  # nothing to do

        # 2) ask for confirmation
        count = len(items)
        resp = QMessageBox.question(
            self,
            "Delete Events",
            f"Are you sure you want to delete the {count} selected event"
            + ("s?" if count > 1 else "?"),
            QMessageBox.Yes | QMessageBox.No
        )
        if resp != QMessageBox.Yes:
            return

        # 3) delete from DB *and* scene
        for ev in items:
            if getattr(ev, 'db_id', None) is not None:
                # remove from sqlite
                self.db.conn.execute(
                    "DELETE FROM event WHERE id = ?", (ev.db_id,)
                )
                self.db.conn.commit()
            # remove the graphic
            self.scene.removeItem(ev)

        # clear any leftover selection
        self.scene.clearSelection()
        self._reload_location_views()

    def _init_general_tab(self):
        layout = QVBoxLayout(self.general_tab)

        updateBtn = QPushButton("Update View")
        updateBtn.clicked.connect(self.update_general_view)
        layout.addWidget(updateBtn)

        syncBtn = QPushButton("Sync to Google Sheets…")
        syncBtn.clicked.connect(self.on_sync_sheets)
        layout.addWidget(syncBtn)

        self.generalTable = QTableWidget(0, 6)
        self.generalTable.setHorizontalHeaderLabels(
            ["Date", "Time", "Activity", "Location", "Link", "Cost"]
        )
        self.generalTable.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.generalTable.setSelectionBehavior(QAbstractItemView.SelectRows)
        layout.addWidget(self.generalTable)

    def update_general_view(self):
        if self.current_schedule_id is None:
            #QMessageBox.warning(self, "No Schedule", "Please load or save a schedule first.")
            return

        # Load canonical events
        _, _, _, _, _, evs = self.db.load_schedule(self.current_schedule_name)
        evs_sorted = sorted(evs, key=lambda ev: (ev['start_dt'], ev['title']))

        self.generalTable.setRowCount(len(evs_sorted))

        for row_idx, ev in enumerate(evs_sorted):
            # Base columns
            sd = ev['start_dt']
            ed = ev['end_dt']
            date_str = sd.toString("ddd, M/d/yy")
            time_str = sd.toString("HH:mm")
            activity = f"From {sd.toString('HH:mm')} to {ed.toString('HH:mm')} - {ev['title']}"

            # Defaults from the event record
            location = ev['city'] + (f", {ev['region']}" if ev['region'] else "")
            link = ev['link']
            cost_val = ev['cost']

            # If this was created from a booking, override from the booking table
            gid = ev.get('group_id')
            if gid:
                try:
                    if ev['type'] == 'hotel':
                        rec = self.db.conn.execute(
                            "SELECT location, link, cost FROM hotel WHERE id = ?",
                            (gid,)
                        ).fetchone()
                        if rec:
                            location, link, cost_val = rec[0], rec[1] or link, rec[2]
                    elif ev['type'] == 'reservation':
                        rec = self.db.conn.execute(
                            "SELECT location, link FROM reservation WHERE id = ?",
                            (gid,)
                        ).fetchone()
                        if rec:
                            location, link = rec[0], rec[1] or link
                except Exception:
                    # on any lookup error, just fall back to the event values
                    pass

            cost = f"£{cost_val:.2f}"

            # Populate the table row
            for col_idx, text in enumerate([date_str, time_str, activity, location, link, cost]):
                item = QTableWidgetItem(text)
                item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                self.generalTable.setItem(row_idx, col_idx, item)

    def _init_schedule_tab(self):
        # Container inside the Schedule tab
        self.schedule_central = QWidget()
        hb = QHBoxLayout(self.schedule_central)
        hb.setContentsMargins(0, 0, 0, 0)

        # — Create timeFrame & timeLayout here —
        self.timeFrame = QFrame()
        self.timeFrame.setFixedWidth(TIME_LABEL_WIDTH)
        self.timeLayout = QVBoxLayout(self.timeFrame)
        self.timeLayout.setContentsMargins(0, HEADER_HEIGHT, 0, 0)
        self.timeLayout.setSpacing(0)

        # — Create the calendar view here —
        self.view = ZoomableGraphicsView()
        self.view.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        self.view.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        self.view.setAlignment(Qt.AlignLeft | Qt.AlignTop)

        hb.addWidget(self.timeFrame)
        hb.addWidget(self.view)
        hb.addWidget(self.detailsPanel)
        hb.addWidget(self.settingsPanel)

        self.timeFrame.hide()
        self.detailsPanel.hide()
        self.settingsPanel.hide()

        self.detailsPanel.setFixedWidth(self.sidePanelWidth)
        self.settingsPanel.setFixedWidth(self.settingsPanelWidth)

        hb.addWidget(self.timeFrame)
        hb.addWidget(self.view)

        # Now parent the overlays
        self.detailsPanel.setParent(self.schedule_central)
        self.settingsPanel.setParent(self.schedule_central)
        self.detailsPanel.hide()
        self.settingsPanel.hide()
        self.timeFrame.hide()
        #hb.addWidget(self.detailsPanel)
        #hb.addWidget(self.settingsPanel)

        # Insert into the tab
        self.schedule_tab.setLayout(QVBoxLayout())
        self.schedule_tab.layout().addWidget(self.schedule_central)

        # Finally, initialize the calendar
        today = QDate.currentDate()
        self.init_calendar(today, today.addDays(13))

    def _init_location_tab(self):
        layout = QVBoxLayout(self.location_tab)

        # 1) Location filter dropdown
        self.locationCombo = QComboBox()
        self.locationCombo.addItem("All")
        self.locationCombo.currentTextChanged.connect(self._reload_location_views)
        layout.addWidget(QLabel("Filter by Location:"))
        layout.addWidget(self.locationCombo)

        layout.addWidget(QLabel("Convert to Euros?"))
        self.convertCurrencyCheck = QCheckBox("Use conversion")
        self.convertCurrencyCheck.setChecked(False)
        self.convertCurrencyCheck.stateChanged.connect(self._reload_location_views)
        layout.addWidget(self.convertCurrencyCheck)

        layout.addWidget(QLabel("Conversion Rate (£→€):"))
        self.conversionRateSpin = QDoubleSpinBox()
        self.conversionRateSpin.setRange(0.1, 10.0)
        self.conversionRateSpin.setSingleStep(0.01)
        self.conversionRateSpin.setValue(1.15)  # e.g. £1 = €1.15
        self.conversionRateSpin.valueChanged.connect(self._reload_location_views)
        layout.addWidget(self.conversionRateSpin)

        layout.addWidget(QLabel("Filter by Event Type:"))
        self.eventTypeCombo = QComboBox()
        self.eventTypeCombo.addItem("All")
        self.eventTypeCombo.currentTextChanged.connect(self._reload_location_views)
        layout.addWidget(self.eventTypeCombo)

        # 2) Tab widget with four pages
        self.locTabWidget = QTabWidget()
        # Things to Eat
        self.eatTable = QTableWidget(0, 5)
        self.eatTable.setHorizontalHeaderLabels(
            ["Item", "Location", "Link", "Description", "Media"])
        self.locTabWidget.addTab(self.eatTable, "Things to Eat")

        # Things to Do
        self.doTable = QTableWidget(0, 5)
        self.doTable.setHorizontalHeaderLabels(
            ["Item", "Location", "Link", "Description", "Media"])
        self.locTabWidget.addTab(self.doTable, "Things to Do")

        # Hotels
        headers = [
            "Location", "Item", "Link",
            "Start Date", "End Date", "Check-in", "Check-out",
            "Breakfast", "Dinner", "Half-board",
            "#Rooms", "Room Types", "Pre-paid", "Cost"
        ]
        self.hotelTable = QTableWidget(0, len(headers))
        self.hotelTable.setHorizontalHeaderLabels(headers)
        self.locTabWidget.addTab(self.hotelTable, "Hotels")

        # — Single Delete button for Hotels & Reservations —
        self.deleteBtn = QPushButton("🗑️ Delete Selected")
        self.deleteBtn.setToolTip("Delete selected rows from Hotels or Reservations")
        self.deleteBtn.clicked.connect(self.on_delete_selected_location_items)
        layout.addWidget(self.deleteBtn)

        self.hotelTable.cellChanged.connect(self.on_hotel_cell_changed)
        self.eatTable.cellChanged.connect(self.on_eat_cell_changed)
        self.doTable.cellChanged.connect(self.on_do_cell_changed)
        # Reservations

        self.resTable = QTableWidget(0, 10)
        self.resTable.setHorizontalHeaderLabels([
            "Location", "Item", "Link", "Description", "Media",
            "Start Date", "End Date", "Start Time", "End Time", ""
        ])
        self.resTable.cellChanged.connect(self.on_reservation_cell_changed)



        self.locTabWidget.addTab(self.resTable, "Reservations")

        # after self.hotelTable = QTableWidget(...)
        boolDelegate = BoolDelegate(self)
        dateDelegate = DateDelegate(self)
        timeDelegate = TimeDelegate(self)
        intDelegate = IntDelegate(self)
        doubleDelegate = DoubleDelegate(self)
        boolDelegate = BoolDelegate(self)

        # Hotels: columns 3=Start Date, 4=End Date, 5=Check-in, 6=Check-out,
        #         7=Breakfast, 8=Dinner, 9=Half-board, 10=#Rooms, 12=Pre-paid, 13=Cost
        self.hotelTable.setItemDelegateForColumn(3, dateDelegate)
        self.hotelTable.setItemDelegateForColumn(4, dateDelegate)
        self.hotelTable.setItemDelegateForColumn(5, timeDelegate)
        self.hotelTable.setItemDelegateForColumn(6, timeDelegate)
        self.hotelTable.setItemDelegateForColumn(7, boolDelegate)  # Breakfast → checkbox
        self.hotelTable.setItemDelegateForColumn(8, boolDelegate)  # Dinner   → checkbox
        self.hotelTable.setItemDelegateForColumn(9, boolDelegate)  # Half-board → checkbox
        self.hotelTable.setItemDelegateForColumn(10, intDelegate)  # #Rooms
        self.hotelTable.setItemDelegateForColumn(12, boolDelegate)  # Pre-paid → checkbox
        self.hotelTable.setItemDelegateForColumn(13, doubleDelegate)  # Cost

        # Reservations: columns 5=Start Date, 6=End Date, 7=Start Time, 8=End Time
        self.resTable.setItemDelegateForColumn(5, dateDelegate)
        self.resTable.setItemDelegateForColumn(6, dateDelegate)
        self.resTable.setItemDelegateForColumn(7, timeDelegate)
        self.resTable.setItemDelegateForColumn(8, timeDelegate)

        # Total Cost Tab
        self.totalCostTab = QWidget()
        tot_layout = QVBoxLayout(self.totalCostTab)

        # Table of individual costs
        self.totalCostTable = QTableWidget(0, 3, self.totalCostTab)
        self.totalCostTable.setHorizontalHeaderLabels(["Type", "Item", "Cost"])
        tot_layout.addWidget(self.totalCostTable)

        # Summary label at the bottom
        self.totalCostSummaryLabel = QLabel("Total: £0.00", self.totalCostTab)
        self.totalCostSummaryLabel.setAlignment(Qt.AlignCenter)
        tot_layout.addWidget(self.totalCostSummaryLabel)

        # Finally, add *that* tab
        self.locTabWidget.addTab(self.totalCostTab, "Total Cost")

        layout.addWidget(self.locTabWidget)

    def on_eat_cell_changed(self, row, col):
        # Fetch the PK stashed in UserRole on first column
        cell = self.eatTable.item(row, 0)
        eat_id = cell.data(Qt.UserRole)
        if eat_id is None:
            return

        # Map column header → eat_item column
        header = self.eatTable.horizontalHeaderItem(col).text()
        mapping = {
            "Item": "item",
            "Location": "location",
            "Link": "link",
            "Description": "description",
            "Media": "media",
        }
        field = mapping.get(header)
        if not field:
            return

        raw = self.eatTable.item(row, col).text()
        # No special parsing needed—all columns are text
        try:
            self.db.update_eat_item(eat_id, **{field: raw})
        except Exception as e:
            QMessageBox.critical(self, "Update Error",
                                 f"Failed to save change to {header}:\n{e}")
            # reload in case of failure
            self._reload_location_views()

    def on_do_cell_changed(self, row, col):
        cell = self.doTable.item(row, 0)
        do_id = cell.data(Qt.UserRole)
        if do_id is None:
            return

        header = self.doTable.horizontalHeaderItem(col).text()
        mapping = {
            "Item": "item",
            "Location": "location",
            "Link": "link",
            "Description": "description",
            "Media": "media",
        }
        field = mapping.get(header)
        if not field:
            return

        raw = self.doTable.item(row, col).text()
        try:
            self.db.update_do_item(do_id, **{field: raw})
        except Exception as e:
            QMessageBox.critical(self, "Update Error",
                                 f"Failed to save change to {header}:\n{e}")
            self._reload_location_views()

    def _populate_event_type_dropdown(self):
        self.eventTypeCombo.blockSignals(True)
        self.eventTypeCombo.clear()
        self.eventTypeCombo.addItem("All")
        if self.current_schedule_id is not None:
            for et in self.db.list_event_types(self.current_schedule_id):
                self.eventTypeCombo.addItem(et)
        self.eventTypeCombo.blockSignals(False)

    def on_delete_selected_location_items(self):
        # figure out which page of the location-tab we're on
        current = self.locTabWidget.currentWidget()

        if current is self.eatTable:
            table, sql, title = self.eatTable, "DELETE FROM eat_item WHERE id = ?", "thing to eat"
        elif current is self.doTable:
            table, sql, title = self.doTable, "DELETE FROM do_item WHERE id = ?", "thing to do"
        elif current is self.hotelTable:
            table, sql, title = self.hotelTable, "DELETE FROM hotel WHERE id = ?", "hotel"
        elif current is self.resTable:
            table, sql, title = self.resTable, "DELETE FROM reservation WHERE id = ?", "reservation"
        else:
            QMessageBox.warning(
                self, "Delete",
                "Switch to the Things to Eat, Things to Do, Hotels, or Reservations tab to delete."
            )
            return

        rows = table.selectionModel().selectedRows()
        if not rows:
            QMessageBox.warning(self, "No Selection", f"Please select one or more {title}s to delete.")
            return

        count = len(rows)
        resp = QMessageBox.question(
            self,
            f"Delete {title.capitalize()}{'s' if count > 1 else ''}",
            f"Are you sure you want to delete the {count} selected {title}{'s' if count > 1 else ''}?",
            QMessageBox.Yes | QMessageBox.No
        )
        if resp != QMessageBox.Yes:
            return

        cur = self.db.conn.cursor()
        # delete in reverse‐row order so the view doesn’t shift under us
        for idx in sorted(rows, key=lambda ix: ix.row(), reverse=True):
            pk = table.item(idx.row(), 0).data(Qt.UserRole)
            cur.execute(sql, (pk,))
        self.db.conn.commit()

        # refresh views
        self._reload_location_views()

    def on_hotel_cell_changed(self, row, col):
        # 1) Fetch the hotel_id you stashed in UserRole (see earlier fix)
        cell = self.hotelTable.item(row, 0)
        hotel_id = cell.data(Qt.UserRole)
        if hotel_id is None:
            return

        # 2) Map the header label to the actual hotel.* column
        header = self.hotelTable.horizontalHeaderItem(col).text()
        mapping = {
            "Location":       "location",
            "Item":           "item",
            "Link":           "link",
            "Start Date":     "start_date",
            "End Date":       "end_date",
            "Check-in":       "check_in_time",
            "Check-out":      "check_out_time",
            "Breakfast":      "breakfast",
            "Dinner":         "dinner",
            "Half-board":     "half_board",
            "#Rooms":         "num_rooms",
            "Room Types":     "room_types",
            "Pre-paid":       "prepaid",
            "Cost":           "cost",
        }
        field = mapping.get(header)
        if not field:
            # unknown column — nothing to do
            return

        raw = self.hotelTable.item(row, col).text()

        # 3) Convert the raw string into the right Python/Qt type
        if field in ("start_date", "end_date"):
            val = QDate.fromString(raw, Qt.ISODate)
        elif field in ("check_in_time", "check_out_time"):
            val = QTime.fromString(raw, "HH:mm")
        elif field in ("breakfast", "dinner", "half_board", "cost"):
            val = float(raw or 0)
        elif field == "num_rooms":
            val = int(raw or 0)
        elif field == "prepaid":
            # your table shows "0"/"1" or "True"/"False"
            val = bool(raw.lower() in ("1", "true", "yes"))
        else:
            val = raw

        # 4) Update the hotel row
        try:
            self.db.update_hotel(hotel_id, **{field: val})
            self.sync_bookings_to_schedule()
        except Exception as e:
            QMessageBox.critical(self, "Update Error",
                                 f"Failed to save changes to {header}:\n{e}")

    def parse_email_file(self, path: str) -> str:
        # Load the .eml
        with open(path, 'rb') as f:
            msg = BytesParser(policy=policy.default).parse(f)

        # Prefer plain-text
        part = msg.get_body(preferencelist=('plain',))
        if part:
            return part.get_content().strip()

        # Fallback to HTML → strip tags
        html = msg.get_body(preferencelist=('html',)).get_content()
        return BeautifulSoup(html, 'html.parser').get_text().strip()

    def on_tab_changed(self, index):
        self.context_toolbar.clear()

        if index == 0:  # General
            # no context actions by default
            self.update_general_view()

        elif index == 1:  # Schedule
            # ☰ Event Details
            details_act = QAction("☰", self)
            details_act.setToolTip("Event Details")
            details_act.triggered.connect(self.toggle_details_panel)
            self.context_toolbar.addAction(details_act)

            # ⚙ Schedule Settings
            settings_act = QAction("⚙", self)
            settings_act.setToolTip("Schedule Settings")
            settings_act.triggered.connect(self.toggle_settings_panel)
            self.context_toolbar.addAction(settings_act)

            # Import action
            imp_act = QAction("⤵ Import Plan JSON…", self)
            imp_act.triggered.connect(self.on_import_plan)
            self.context_toolbar.addAction(imp_act)
            # Export action
            exp_act = QAction("⤴ Export Plan JSON…", self)
            exp_act.triggered.connect(self.on_export_plan)
            self.context_toolbar.addAction(exp_act)

            route_act = QAction("🛣️ Add Route", self)
            route_act.triggered.connect(self.on_add_route)
            self.context_toolbar.addAction(route_act)

            # ── New “Build Description” button ──
            desc_act = QAction("✍️ Build Description", self)
            desc_act.setToolTip("Open the Description Builder")
            desc_act.triggered.connect(self.on_build_description)
            self.context_toolbar.addAction(desc_act)

            # ── New “Convert to Plan” button ──
            conv_act = QAction("🔄 Convert to Plan", self)
            conv_act.setToolTip("Turn an event’s description into a full plan")
            conv_act.triggered.connect(self.on_convert_to_plan)
            self.context_toolbar.addAction(conv_act)

            # ↓ new “Map” button:
            map_act = QAction("🗺️ Show Map Route", self)
            map_act.triggered.connect(self.on_create_map_link)
            self.context_toolbar.addAction(map_act)

            ai_act = QAction("🤖 AI Plan…", self)
            ai_act.triggered.connect(self.on_ai_plan)
            self.context_toolbar.addAction(ai_act)

            place_act = QAction("📍 Add Place", self)
            place_act.triggered.connect(self.on_add_place)
            self.context_toolbar.addAction(place_act)

            # in on_tab_changed, in the Schedule section:
            add_files_act = QAction("➕ Add Event", self)
            add_files_act.setToolTip("Import one or more .txt/.pdf files as reservations or hotel bookings")
            add_files_act.triggered.connect(self.on_add_event_files)
            self.context_toolbar.addAction(add_files_act)

            # ↻ Update Events
            update_act = QAction("↻ Update Events", self)
            update_act.setToolTip("Re-fetch travel times for selected directions links and resize events")
            update_act.triggered.connect(self.on_update_events)
            self.context_toolbar.addAction(update_act)

        else:  # Location
            loc_act = QAction("☰", self)
            loc_act.setToolTip("Location Filter")
            loc_act.triggered.connect(lambda: print("Filter by location"))
            self.context_toolbar.addAction(loc_act)

            import_action = QAction("📄 Import Things to Do and Eat…", self)
            import_action.setToolTip("Extract local things to do/eat from a text file or pasted text")
            import_action.triggered.connect(self.on_import_text_items)
            self.context_toolbar.addAction(import_action)

            check_act = QAction("Check Locations", self)
            check_act.setToolTip("Scan all events and add their City values here")
            check_act.triggered.connect(self.on_check_locations)
            self.context_toolbar.addAction(check_act)

            self._populate_location_dropdown()
            self._populate_event_type_dropdown()
            self._reload_location_views()

        # Highlight the active tab in blue
        for i in range(self.tab_widget.count()):
            col = QColor("blue") if i == index else QColor("black")
            self.tab_widget.tabBar().setTabTextColor(i, col)

    import json

    from urllib.parse import urlparse, parse_qs
    from PyQt5.QtWidgets import QInputDialog

    @staticmethod
    def expand_link(link: str) -> str:
        """Follows any HTTP redirect (maps.app.goo.gl, goo.gl, etc.) and returns the final URL."""
        print(f"[expand_link] input short link: {link}")
        try:
            resp = requests.head(link, allow_redirects=True, timeout=5)
            print(f"[expand_link] expanded link: {resp.url}")
            return resp.url
        except Exception:
            # if anything goes wrong, fall back on the original
            print(f"[expand_link] failed to expand link: {link}")
            return link

    def on_check_locations(self):
        if not self.current_schedule_id:
            QMessageBox.warning(self, "No Schedule", "Load or save a schedule first.")
            return

        # 1) Fetch every distinct city in events
        found = self.db.list_event_cities(self.current_schedule_id)
        if not found:
            QMessageBox.information(self, "No Cities", "No city values found in events.")
            return

        # 2) Insert each into the location table (duplicates auto-ignored)
        new_cities = []
        for city in found:
            rid = self.db.add_location(self.current_schedule_id, city)
            if rid:  # only non-zero if a new row was inserted
                new_cities.append(city)

        # 3) Update the combo with exactly those new ones
        for city in new_cities:
            self.locationCombo.addItem(city)

        QMessageBox.information(
            self,
            "Locations Added",
            f"Added {len(new_cities)} new location(s):\n" + ", ".join(new_cities)
        )

    def _drop_event_box(self, db_id, title, date, time, duration_minutes, event_type):
        # compute geometry
        col = self.scene.start_date.daysTo(date)
        row = (time.hour() * 60 + time.minute() - START_HOUR * 60) // self.scene.slot_minutes
        x = TIME_LABEL_WIDTH + col * self.scene.cell_w
        h = (duration_minutes / self.scene.slot_minutes) * self.scene.cell_h
        rect = QRectF(0, 0, self.scene.cell_w, h)
        item = EventItem(rect, title, QColor("#FFA"))
        item.link = ""  # or pull from DB if you like
        item.city = "";
        item.region = "";
        item.event_type = event_type
        item.cost = 0.0
        item.setPos(x, HEADER_HEIGHT + row * self.scene.cell_h)
        item.db_id = db_id
        self.scene.addItem(item)

    def on_build_description(self):
        # 1) Require exactly one event selected
        sel = [it for it in self.scene.selectedItems() if isinstance(it, EventItem)]
        if len(sel) != 1:
            QMessageBox.warning(self, "Select Event", "Select exactly one event to build a description for.")
            return
        ev = sel[0]

        # 2) Show the same dialog you use for AI‐plans
        dlg = AIPlanDialog(self, self.starred_places)
        # Pre-fill context from the event, if you like:
        dlg.ctxEdit.setText(f"{ev.city}, {ev.region}".strip())
        dlg.textEdit.setPlainText(ev.description)

        if dlg.exec_() != QDialog.Accepted:
            return

        # 3) Pull out the built text and save it in the event & DB
        new_desc = dlg.get_text()
        ev.description = new_desc
        if getattr(ev, 'db_id', None):
            self.db.update_event(ev.db_id, description=new_desc)
            self._reload_location_views()
        # 4) If the details‐panel is visible, update its QLineEdit
        self.descriptionEdit.setText(new_desc)

        QMessageBox.information(self, "Description Saved", "Event description updated.")

    def on_convert_to_plan(self):
        # 1) Require exactly one event selected
        sel = [it for it in self.scene.selectedItems() if isinstance(it, EventItem)]
        if len(sel) != 1:
            QMessageBox.warning(self, "Select Event", "Select exactly one event to convert to a plan.")
            return
        ev = sel[0]

        # 2) Must have a non‐empty description
        raw_text = (ev.description or "").strip()
        print(f"[DEBUG] ev.db_id = {ev.db_id!r}")
        print(f"[DEBUG] raw_text = {raw_text!r}")
        if not raw_text:
            QMessageBox.warning(self, "No Description", "This event has no description to convert.")
            return

        # 3) Compute the anchor’s datetime (so we know where to drop the plan)
        ax, ay = ev.pos().x(), ev.pos().y()
        col = int((ax - TIME_LABEL_WIDTH) // self.scene.cell_w)
        row = int((ay - HEADER_HEIGHT) // self.scene.cell_h)
        day = self.scene.start_date.addDays(col)
        mins = START_HOUR * 60 + row * self.scene.slot_minutes
        anchor_dt = QDateTime(day, QTime(mins // 60, mins % 60))
        print(f"[DEBUG] anchor position = ({ax:.1f}, {ay:.1f}) → anchor_dt = {anchor_dt.toString()}")
        # 4) Remove the original event from scene & DB
        if getattr(ev, 'db_id', None):
            self.db.conn.execute("DELETE FROM event WHERE id=?", (ev.db_id,))
            self.db.conn.commit()
        self.scene.removeItem(ev)
        print("[DEBUG] removed EventItem from scene")

        # 5) Call OpenAI to generate the plan JSON
        try:
            resp = openai.ChatCompletion.create(
                model="gpt-4",
                temperature=0,
                messages=[
                    # 1) your existing schema‐enforcing prompt
                    {"role": "system", "content": (
                        "You are an assistant that reads raw holiday planning text "
                        "and outputs a JSON object with this schema:\n"
                        "{\n"
                        "  \"plan_name\": string,\n"
                        "  \"default_slot_minutes\": int,\n"
                        "  \"events\": [\n"
                        "    { \"title\":string, \"description\":string, \"duration\":int, "
                        "\"spacing_after\":int, \"link\":string, \"city\":string, "
                        "\"region\":string, \"event_type\":string, \"cost\":float, "
                        "\"color\":string }\n"
                        "  ]\n"
                        "}\n"
                        "Return *only* valid JSON. Duration must be 5 or greater; for events "
                        "including lunch or dinner use 60, for other food places 15. Categorise "
                        "as the following types: Food (color = green), Travel (color = blue), "
                        "Activity (color = orange), Accommodation (color = purple), Other (color = grey)."
                    )},
                    # 2) new context + place_id instruction
                    {"role": "system", "content": (
                        f"Context for planning: Japan\n"
                        "Whenever you mention a place, if you can identify its Google Place ID, "
                        "append `(place_id:THE_ID)` immediately after the place name."
                    )},
                    {"role": "user", "content": raw_text},
                    # 3) the user’s freeform (or built‐up) text

                ]
            )
            base_json = resp.choices[0].message.content
            print(f"[DEBUG] OpenAI response: {base_json}")
            plan = json.loads(base_json)

        except Exception as e:
            QMessageBox.critical(self, "AI Error", f"Failed to generate plan:\n{e}")
            return

        # 6) Enrich with routes
        try:
            plan = self.enrich_plan_with_routes(plan, walking_threshold=20)
        except Exception as e:
            QMessageBox.critical(self, "Routing Error", f"Failed to enrich:\n{e}")
            return

        # 7) Auto‐save to PLANS_DIR with timestamp from anchor_dt
        timestamp = anchor_dt.toString("yyyy-MM-dd_HH-mm")
        filename = f"{timestamp}.json"
        os.makedirs(PLANS_DIR, exist_ok=True)
        path = os.path.join(PLANS_DIR, filename)
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(plan, f, indent=2)

        # 8) Drop each sub‐event onto the calendar, advancing a cursor
        cursor = anchor_dt
        for ev_def in plan["events"]:
            slots = ev_def["duration"] / self.scene.slot_minutes
            height = slots * self.scene.cell_h
            rect = QRectF(0, 0, self.scene.cell_w, height)
            item = EventItem(rect, ev_def["title"], QColor(ev_def.get("color", "#FFA")))
            item.description = ev_def.get("description", "")
            item.link = ev_def.get("link", "")
            item.city = ev_def.get("city", "")
            item.region = ev_def.get("region", "")
            item.event_type = ev_def.get("event_type", "")
            item.cost = ev_def.get("cost", 0.0)
            # position
            day_idx = self.scene.start_date.daysTo(cursor.date())
            x = TIME_LABEL_WIDTH + day_idx * self.scene.cell_w
            mins_from_start = (cursor.time().hour() * 60 + cursor.time().minute())
            y = HEADER_HEIGHT + ((mins_from_start - START_HOUR * 60) // self.scene.slot_minutes) * self.scene.cell_h
            item.setPos(x, y)
            self.scene.addItem(item)

            # persist to DB if loaded
            if self.current_schedule_id:
                sd = cursor
                ed = cursor.addSecs(int(ev_def["duration"]) * 60)
                item.db_id = self.db.insert_event(
                    self.current_schedule_id,
                    ev_def["title"], ev_def.get("description", ""),
                    sd, ed, ev_def["duration"],
                    ev_def.get("link", ""), ev_def.get("city", ""),
                    ev_def.get("region", ""), ev_def.get("event_type", ""),
                    ev_def.get("cost", 0.0), ev_def.get("color", "#FFA"),
                    x, y, self.scene.cell_w, height
                )

            # advance the cursor by duration + spacing_after
            advance = ev_def["duration"] + ev_def.get("spacing_after", 0)
            cursor = cursor.addSecs(int(advance) * 60)

        QMessageBox.information(self, "Plan Created",
                                f"Converted event into {len(plan['events'])} sub‐events.\nSaved plan to {path}")

    from datetime import datetime

    def on_update_events(self):
        items = [it for it in self.scene.selectedItems() if isinstance(it, EventItem)]
        if not items:
            QMessageBox.warning(self, "No Selection", "Select one or more events to update travel times.")
            return

        updated = 0
        for ev in items:
            raw = (ev.link or "").strip()
            if not raw or "/dir/" not in raw:
                continue

            # expand short links
            long_url = raw
            if raw.startswith("https://maps.app.goo.gl/"):
                long_url = self.expand_link(raw)

            p = urlparse(long_url)
            qs = parse_qs(p.query)
            origin = qs.get("origin", [None])[0]
            dest = qs.get("destination", [None])[0]

            if not origin or not dest:
                # try path segments
                parts = p.path.split("/")
                if "dir" in parts:
                    idx = parts.index("dir")
                    origin = origin or (parts[idx + 1] if len(parts) > idx + 1 else None)
                    dest = dest or (parts[idx + 2] if len(parts) > idx + 2 else None)
            if not origin or not dest:
                continue

            # choose mode based on the event type, not the URL
            mode = ev.event_type if ev.event_type in ("walking", "driving", "transit", "bus", "train") else "walking"

            # rebuild the link so it shows the correct travelmode
            new_url = (
                "https://www.google.com/maps/dir/?api=1"
                f"&origin={origin}"
                f"&destination={dest}"
                f"&travelmode={mode}"
            )
            ev.link = new_url
            if getattr(ev, "db_id", None):
                self.db.update_event(ev.db_id, link=new_url)

            # set up transit parameters if needed
            kwargs = {}
            if mode in ("transit", "bus", "train"):
                kwargs["transit_mode"] = ["bus", "train"]
                kwargs["departure_time"] = int(time.time())

            resp = self.gmaps.directions(origin, dest, mode=mode, **kwargs)
            if not resp:
                QMessageBox.warning(self, "Directions Error",
                                    f"No {mode} route found for:\n{ev.link}")
                continue

            leg = resp[0]["legs"][0]

            # if transit was requested, ensure we got a transit leg
            if mode in ("transit", "bus", "train"):
                if not any(step.get("travel_mode", "").upper() == "TRANSIT" for step in leg.get("steps", [])):
                    QMessageBox.warning(self, "Directions Error",
                                        f"No transit route found for:\n{ev.link}")
                    continue

            # calculate and apply new duration
            new_min = math.ceil(leg["duration"]["value"] / 60)
            slot = self.scene.slot_minutes
            old_min = int((ev.rect().height() / self.scene.cell_h) * slot)
            if new_min == old_min:
                continue

            new_h = (new_min / slot) * self.scene.cell_h
            r = ev.rect()
            ev.setRect(QRectF(r.x(), r.y(), r.width(), new_h))

            if getattr(ev, "db_id", None):
                x = ev.pos().x();
                y = ev.pos().y()
                col = int((x - TIME_LABEL_WIDTH) // self.scene.cell_w)
                row = int((y - HEADER_HEIGHT) // self.scene.cell_h)
                day = self.scene.start_date.addDays(col)
                mins = START_HOUR * 60 + row * slot
                start_dt = QDateTime(day, QTime(mins // 60, mins % 60))
                end_dt = start_dt.addSecs(new_min * 60)
                self.db.update_event(ev.db_id,
                                     duration=new_min,
                                     end_dt=end_dt,
                                     h=new_h)
                self._reload_location_views()

            updated += 1

        QMessageBox.information(self, "Update Complete",
                                f"Updated {updated} travel-time event" + ("s." if updated != 1 else "."))

    def get_event_end_timestamp(self, event_item):
        """Return the end time of ``event_item`` as a Unix timestamp."""
        col = int((event_item.pos().x() - TIME_LABEL_WIDTH) // self.scene.cell_w)
        row = int((event_item.pos().y() - HEADER_HEIGHT) // self.scene.cell_h)
        day = self.scene.start_date.addDays(col)
        mins = START_HOUR * 60 + row * self.scene.slot_minutes
        start_dt = QDateTime(day, QTime(mins // 60, mins % 60))
        end_dt = start_dt.addSecs(
            int(event_item.rect().height() / self.scene.cell_h)
            * self.scene.slot_minutes * 60
        )
        return end_dt.toSecsSinceEpoch()

    def get_event_coordinates(self, event_item):
        """Resolve ``event_item`` into ``(lat, lng)`` coordinates."""
        if hasattr(event_item, "lat") and hasattr(event_item, "lng"):
            return event_item.lat, event_item.lng

        name = event_item.text.toPlainText().strip()
        if event_item.city:
            name += f", {event_item.city}"
        if event_item.region:
            name += f", {event_item.region}"

        resp = self.gmaps.find_place(
            input=name,
            input_type="textquery",
            fields=["geometry"],
        )
        cands = resp.get("candidates", [])
        if cands:
            loc = cands[0]["geometry"]["location"]
            return loc["lat"], loc["lng"]

        geo = self.gmaps.geocode(name)
        if geo:
            loc = geo[0]["geometry"]["location"]
            return loc["lat"], loc["lng"]

        raise RuntimeError(f"Could not resolve coords for “{name}”")

    def on_add_route(self):
        # 1) require exactly two events selected
        #
        # ``get_event_coordinates`` and ``get_event_end_timestamp`` are helpers
        # that resolve EventItem details.
        #
        # 2) require exactly two events selected
        items = [it for it in self.scene.selectedItems() if isinstance(it, EventItem)]
        if len(items) != 2:
            QMessageBox.warning(self, "Selection Error",
                                "Select exactly two events to create a route.")
            return
        a, b = sorted(items, key=lambda ev: (ev.pos().x(), ev.pos().y()))

        # 3) resolve origin & destination
        try:
            lat1, lng1 = self.get_event_coordinates(a)
            lat2, lng2 = self.get_event_coordinates(b)
        except RuntimeError as e:
            QMessageBox.critical(self, "Routing Error", str(e))
            return

        # 1) figure out when event A ends and convert to a UNIX timestamp
        departure = self.get_event_end_timestamp(a)
        a_end = QDateTime.fromSecsSinceEpoch(departure)

        modes = {
            "walking": {"label": "Walking", "params": {}},
            "transit": {
                "label": "Bus",
                "params": {
                    "transit_mode": ["bus"],
                    "departure_time": departure
                }
            },
            "driving": {
                "label": "Driving",
                "params": {
                    "departure_time": departure
                }
            },
        }

        print("[DEBUG] Fetching routes for:", modes)
        print(f"[DEBUG] Departure time: {departure} ({datetime.fromtimestamp(departure)})")
        print("[DEBUG] Origin:", (lat1, lng1), "→ Destination:", (lat2, lng2))
        print("[DEBUG] Event A end time:", a_end.toString())
        print("[DEBUG] Event B text:", b.text.toPlainText())
        print("[DEBUG] Driving params:", modes["driving"]["params"])
        print("[DEBUG] Transit params:", modes["transit"]["params"])
        print("[DEBUG] Walking params:", modes["walking"]["params"])

        best = {}
        for mode, info in modes.items():
            try:
                routes = self.gmaps.directions(
                    (lat1, lng1),
                    (lat2, lng2),
                    mode=mode,
                    alternatives=True,
                    **info["params"]
                )
            except Exception:
                continue
            if not routes:
                continue
            fastest = min(routes, key=lambda r: r["legs"][0]["duration"]["value"])
            leg = fastest["legs"][0]
            best[mode] = {
                "route": fastest,
                "seconds": leg["duration"]["value"],
                "label": info["label"]
            }

        if not best:
            QMessageBox.warning(self, "No Routes", "No walking, bus or driving routes found.")
            return

        # 5) apply your priority rules
        chosen = None
        if "walking" in best and best["walking"]["seconds"] <= 10 * 60:
            chosen = best["walking"]
        else:
            if "transit" in best:
                chosen = best["transit"]
            if "driving" in best and "transit" in best:
                bus = best["transit"]["seconds"]
                drv = best["driving"]["seconds"]
                if drv * 1.3 <= bus:
                    chosen = best["driving"]
            if chosen is None and "driving" in best:
                chosen = best["driving"]

        # 6) drop it onto your calendar just like before
        label = chosen["label"]
        dur_min = math.ceil(chosen["seconds"] / 60)
        title = f"{label} → {b.text.toPlainText()}"

        route_url = (
            "https://www.google.com/maps/dir/?api=1"
            f"&origin={lat1},{lng1}"
            f"&destination={lat2},{lng2}"
            f"&travelmode={label.lower()}"
            f"&departure_time={departure}"
        )

        # compute insert‐point using the helper-derived end time
        new_col = self.scene.start_date.daysTo(a_end.date())
        new_row = ((a_end.time().hour() * 60 + a_end.time().minute()
                    - START_HOUR * 60) // self.scene.slot_minutes)
        x = TIME_LABEL_WIDTH + new_col * self.scene.cell_w
        h = (dur_min / self.scene.slot_minutes) * self.scene.cell_h

        rect = QRectF(0, 0, self.scene.cell_w, h)
        leg_item = EventItem(rect, title, QColor("#ADE1F9"))
        leg_item.link = route_url
        leg_item.setPos(x, HEADER_HEIGHT + new_row * self.scene.cell_h)
        self.scene.addItem(leg_item)

        if self.current_schedule_id:
            leg_item.db_id = self.db.insert_event(
                self.current_schedule_id,
                title, "",  # no description
                a_end,
                a_end.addSecs(dur_min * 60),
                dur_min,
                route_url, "", "",
                label.lower(), 0.0,
                "#ADE1F9",
                x, HEADER_HEIGHT + new_row * self.scene.cell_h,
                self.scene.cell_w, h
            )

    def on_add_place(self):
        # 1) Require exactly one event selected
        items = [it for it in self.scene.selectedItems() if isinstance(it, EventItem)]
        if len(items) != 1:
            QMessageBox.warning(self, "Select Event", "Select exactly one event to add a place to.")
            return
        ev = items[0]

        # 2) Show the place‐picker dialog
        dlg = PlaceDialog(self, self.starred_places)
        if dlg.exec_() != QDialog.Accepted:
            return
        place = dlg.get_selected()

        # 3) Update the EventItem & database
        ev.set_title(place['name'])
        ev.link = place['maps_url']
        if getattr(ev, 'db_id', None):
            self.db.update_event(ev.db_id,
                                 title=place['name'],
                                 link=place['maps_url'])
            self._reload_location_views()


    def on_create_map_link(self):
        # 1) Gather selected EventItems
        items = [it for it in self.scene.selectedItems() if isinstance(it, EventItem)]
        if not items:
            QMessageBox.warning(self, "No Selection",
                                "Select one or more events to generate a map link.")
            return

        # 2) Sort by calendar order (day → time)
        def sort_key(ev):
            col = int((ev.pos().x() - TIME_LABEL_WIDTH) // self.scene.cell_w)
            row = int((ev.pos().y() - HEADER_HEIGHT) // self.scene.cell_h)
            return (col, row)

        items.sort(key=sort_key)

        # 3) Extract “points” for each link
        points = []
        for ev in items:
            link = ev.link or ""
            p = urlparse(link)
            qs = parse_qs(p.query)

            if "/dir/" in p.path:
                # a route-link → grab both ends
                origin = qs.get("origin", [None])[0]
                dest = qs.get("destination", [None])[0]
                if origin: points.append(origin)
                if dest:   points.append(dest)

            elif "/place/" in p.path:
                # a place-link → q=place_id:XYZ
                qvals = qs.get("q", [])
                if qvals:
                    points.append(qvals[0])

            elif "cid" in qs:
                # legacy maps?cid=XYZ
                cid = qs["cid"][0]
                points.append(f"cid:{cid}")

            else:
                # fallback: just push the raw link
                points.append(link)

        # 4) Dedupe while preserving order
        seen = set();
        unique = []
        for pt in points:
            if pt and pt not in seen:
                seen.add(pt)
                unique.append(pt)

        # 5) Build the final URL
        if len(unique) == 1:
            # only one location → show its original link
            url = items[0].link

        elif len(unique) >= 2:
            origin = unique[0]
            destination = unique[-1]
            waypoints = unique[1:-1]

            url = (
                "https://www.google.com/maps/dir/?api=1"
                f"&origin={origin}"
                f"&destination={destination}"
                "&travelmode=walking"
            )
            if waypoints:
                # Google expects '|'-separated waypoints
                url += "&waypoints=" + "|".join(waypoints)

        else:
            QMessageBox.warning(self, "Invalid Links",
                                "Could not parse any valid map links from the selected events.")
            return

        # 6) Pop up an input dialog pre-filled with the URL so the user can copy it
        dlg = QInputDialog(self)
        dlg.setWindowTitle("Google Maps Link")
        dlg.setLabelText("Combined route / place link:")
        dlg.setTextValue(url)
        dlg.setOption(QInputDialog.NoButtons, False)
        dlg.setTextEchoMode(QLineEdit.Normal)

        # find the embedded QLineEdit so we can override its context menu
        edit = dlg.findChild(QLineEdit)
        if edit:
            edit.setContextMenuPolicy(Qt.CustomContextMenu)

            def on_ctx(pos):
                menu = edit.createStandardContextMenu()
                sel = edit.selectedText().strip()
                if re.match(r'https?://|www\.', sel):
                    act = menu.addAction("Open Link")

                    def _open():
                        link = sel
                        if not re.match(r'https?://', link):
                            link = 'http://' + link
                        webbrowser.open(link)

                    act.triggered.connect(_open)
                    menu.addSeparator()
                menu.exec_(edit.mapToGlobal(pos))

            edit.customContextMenuRequested.connect(on_ctx)

        if dlg.exec_() == QDialog.Accepted:
            # user clicked “OK” (you can grab dlg.textValue() here if you need)
            pass

    def load_starred_places(self, json_path):
        """Return list of dicts with name, cid, lat, lng, and a maps_url we can re-use."""
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)

        places = []
        for feat in data.get('features', []):
            props = feat.get('properties', {})
            loc = props.get('location', {})
            name = loc.get('name')
            url = props.get('google_maps_url', '')
            # pull out the CID query-param
            cid = None
            try:
                qs = parse_qs(urlparse(url).query)
                cid = qs.get('cid', [None])[0]
            except Exception:
                pass

            # geometry coordinates come as [lon, lat]
            coords = feat.get('geometry', {}).get('coordinates', [])
            if name and cid and len(coords) == 2:
                lon, lat = coords
                places.append({
                    'name': name,
                    'cid': cid,
                    'lat': lat,
                    'lng': lon,
                    'maps_url': f"https://www.google.com/maps?cid={cid}"
                })
        return places

    def on_ai_plan(self):
        # 1) Require an anchor cell selected
        sel = [it for it in self.scene.selectedItems() if isinstance(it, EventItem)]
        if not sel:
            QMessageBox.warning(self, "Select Anchor", "Select a single event cell to serve as start point.")
            return
        anchor = sel[0]

        # 2) Remove that anchor from scene & DB (if any)
        self.scene.removeItem(anchor)
        if getattr(anchor, 'db_id', None):
            self.db.conn.execute("DELETE FROM event WHERE id=?", (anchor.db_id,))
            self.db.conn.commit()

        # --- load your starred JSON once ---
        if not hasattr(self, 'starred_places'):
            path, _ = QFileDialog.getOpenFileName(
                self,
                "Select your Google Takeout starred-places JSON",
                "", "JSON Files (*.json)"
            )
            if not path:
                return
            try:
                self.starred_places = self.load_starred_places(path)
            except Exception as e:
                QMessageBox.critical(self, "Load error",
                                     f"Could not parse starred places:\n{e}")
                return

        # --- show the AIPlanDialog we built earlier ---
        dlg = AIPlanDialog(self, self.starred_places)
        if dlg.exec_() != QDialog.Accepted:
            return
        trip_ctx = dlg.get_context()
        user_text = dlg.get_text().strip()
        if not user_text:
            QMessageBox.warning(self, "Empty Plan",
                                "You didn’t enter any text.")
            return

        # --- now your existing OpenAI logic, replacing file-read with user_text ---
        trip_ctx = dlg.get_context()
        user_text = dlg.get_text().strip()

        try:
            resp = openai.ChatCompletion.create(
                model="gpt-4",
                temperature=0,
                messages=[
                    # 1) your existing schema‐enforcing prompt
                    {"role": "system", "content": (
                        "You are an assistant that reads raw holiday planning text "
                        "and outputs a JSON object with this schema:\n"
                        "{\n"
                        "  \"plan_name\": string,\n"
                        "  \"default_slot_minutes\": int,\n"
                        "  \"events\": [\n"
                        "    { \"title\":string, \"description\":string, \"duration\":int, "
                        "\"spacing_after\":int, \"link\":string, \"city\":string, "
                        "\"region\":string, \"event_type\":string, \"cost\":float, "
                        "\"color\":string }\n"
                        "  ]\n"
                        "}\n"
                        "Return *only* valid JSON. Duration must be 5 or greater; for events "
                        "including lunch or dinner use 60, for other food places 15. Categorise "
                        "as the following types: Food (color = green), Travel (color = blue), "
                        "Activity (color = orange), Accommodation (color = purple), Other (color = grey)."
                    )},
                    # 2) new context + place_id instruction
                    {"role": "system", "content": (
                        f"Context for planning: {trip_ctx}\n"
                        "Whenever you mention a place, if you can identify its Google Place ID, "
                        "append `(place_id:THE_ID)` immediately after the place name."
                    )},
                    # 3) the user’s freeform (or built‐up) text
                    {"role": "user", "content": user_text}
                ]
            )
            base_json = resp.choices[0].message.content
            plan = json.loads(base_json)
        except Exception as e:
            QMessageBox.critical(self, "AI Error", f"Failed to generate plan:\n{e}")
            return

        # 3) enrich it with walking/bus legs
        try:
            plan = self.enrich_plan_with_routes(plan, walking_threshold=20)
        except Exception as e:
            QMessageBox.critical(self, "Routing Error", f"Failed to enrich routes:\n{e}")
            return

        # 4) Compute starting QDateTime from the anchor’s position:
        ax, ay = anchor.pos().x(), anchor.pos().y()
        acol = int((ax - TIME_LABEL_WIDTH) // self.scene.cell_w)
        arow = int((ay - HEADER_HEIGHT) // self.scene.cell_h)
        aday = self.scene.start_date.addDays(acol)
        amin = START_HOUR * 60 + arow * self.scene.slot_minutes
        cursor_dt = QDateTime(aday, QTime(amin // 60, amin % 60))

        # 5) Auto‐save the enriched plan JSON into PLANS_DIR with a timestamped filename
        timestamp = cursor_dt.toString("yyyy-MM-dd_HH-mm")
        filename = f"{timestamp}.json"
        os.makedirs(PLANS_DIR, exist_ok=True)
        save_path = os.path.join(PLANS_DIR, filename)
        with open(save_path, 'w', encoding='utf-8') as f:
            json.dump(plan, f, indent=2)
        print(f"Enriched plan auto-saved to {save_path}")

        # 6) Now drop *all* events (visits + travel legs) onto the calendar:
        for ev in plan["events"]:
            slots = ev["duration"] / self.scene.slot_minutes
            height = slots * self.scene.cell_h
            rect = QRectF(0, 0, self.scene.cell_w, height)
            item = EventItem(rect, ev["title"], QColor(ev.get("color", "#FFA")))
            # copy metadata
            item.description = ev.get("description", "")
            item.link = ev.get("link", "")
            item.city = ev.get("city", "")
            item.region = ev.get("region", "")
            item.event_type = ev.get("event_type", "")
            item.cost = ev.get("cost", 0.0)

            # position at cursor_dt
            day_offset = self.scene.start_date.daysTo(cursor_dt.date())
            x = TIME_LABEL_WIDTH + day_offset * self.scene.cell_w
            mins_from_start = (
                                      cursor_dt.time().hour() * 60 +
                                      cursor_dt.time().minute()
                              ) - START_HOUR * 60
            y = HEADER_HEIGHT + (mins_from_start // self.scene.slot_minutes) * self.scene.cell_h
            item.setPos(x, y)
            self.scene.addItem(item)

            # advance the cursor
            advance = ev["duration"] + ev.get("spacing_after", 0)
            cursor_dt = cursor_dt.addSecs(int(advance) * 60)

            # insert into DB if needed
            if self.current_schedule_id:
                item.db_id = self.db.insert_event(
                    self.current_schedule_id,
                    ev["title"], ev.get("description", ""),
                    QDateTime(cursor_dt.date(),
                              cursor_dt.time().addSecs(-ev["duration"] * 60)),
                    cursor_dt,
                    ev["duration"],
                    ev.get("link", ""),
                    ev.get("city", ""),
                    ev.get("region", ""),
                    ev.get("event_type", ""),
                    ev.get("cost", 0.0),
                    ev.get("color", "#FFA"),
                    x, y, self.scene.cell_w, height
                )

    def enrich_plan_with_routes(self, plan, walking_threshold=20):
        evs = plan.get('events', [])
        print(f"[DEBUG enrich] plan keys: {list(plan.keys())}")
        print(f"[DEBUG enrich] events count: {len(evs)}")
        if not evs:
            raise RuntimeError("Enrich failed: plan['events'] is empty")
        gmaps = self.gmaps

        # build fast lookups
        starred_by_name = {
            p['name'].lower(): p
            for p in self.starred_places
        }

        pid_regex = re.compile(r'place_id:([A-Za-z0-9_-]+)')

        for ev in plan['events']:
            title = ev.get('title', '').strip()
            loc = None

            # --- 1) Already has lat/lng? skip ---
            if 'lat' in ev and 'lng' in ev:
                continue

            # --- 2) Check starred places by exact name match ---
            key = title.lower()
            if key in starred_by_name:
                star = starred_by_name[key]
                ev['lat'], ev['lng'] = star['lat'], star['lng']
                ev['link'] = star.get('maps_url', ev.get('link', ''))
                continue

            # --- 3) Check for an explicit place_id in the link or title ---
            pid = None
            # a) maybe link is already a place_id url
            link = ev.get('link', '')
            m = pid_regex.search(link)
            if m:
                pid = m.group(1)
            else:
                # b) maybe title itself contains “(place_id:XYZ)”
                m2 = pid_regex.search(title)
                if m2:
                    pid = m2.group(1)

            if pid:
                try:
                    resp = gmaps.place(place_id=pid, fields=['geometry'])
                    loc = resp['result']['geometry']['location']
                    ev['link'] = f"https://www.google.com/maps/place/?q=place_id:{pid}"
                    ev['lat'], ev['lng'] = loc['lat'], loc['lng']
                    continue
                except Exception:
                    # failed to fetch by place_id → fall through
                    pass

            # --- 4) Fallback: textquery lookup by name (+ optional city/region for context) ---
            query = title
            if ev.get('city'):
                query += ", " + ev['city']
            if ev.get('region'):
                query += ", " + ev['region']

            try:
                resp = gmaps.find_place(
                    input=query,
                    input_type="textquery",
                    fields=['geometry', 'place_id']
                )
                cands = resp.get('candidates', [])
                if cands:
                    best = cands[0]
                    loc = best['geometry']['location']
                    pid = best['place_id']
                    ev['link'] = f"https://www.google.com/maps/place/?q=place_id:{pid}"
                    ev['lat'], ev['lng'] = loc['lat'], loc['lng']
                else:
                    # last resort: geocode
                    geo = gmaps.geocode(query)
                    if geo:
                        loc = geo[0]['geometry']['location']
                        ev['lat'], ev['lng'] = loc['lat'], loc['lng']
            except ApiError as e:
                raise RuntimeError(f"Lookup error for “{title}”: {e}")

            if not loc:
                raise RuntimeError(f"Could not resolve location “{title}”")

        # 2) (Optional) optimize overall loop if 3+ places
        if len(plan['events']) >= 3:
            try:
                origin = (plan['events'][0]['lat'], plan['events'][0]['lng'])
                waypts = [f"{e['lat']},{e['lng']}" for e in plan['events'][1:]]
                resp = gmaps.directions(
                    origin=origin,
                    destination=origin,
                    mode="walking",
                    waypoints=waypts,
                    optimize_waypoints=True
                )
                if resp and 'waypoint_order' in resp[0]:
                    order = resp[0]['waypoint_order']
                    reordered = [plan['events'][0]] + [plan['events'][i + 1] for i in order]
                    plan['events'] = reordered
            except ApiError as e:
                raise RuntimeError(f"[Optimize] Directions API error: {e}")

        # 3) Insert travel legs
        enriched = []
        for a, b in zip(plan['events'], plan['events'][1:]):
            enriched.append(a)

            # Always get walking first
            try:
                walk_resp = gmaps.directions(
                    (a['lat'], a['lng']),
                    (b['lat'], b['lng']),
                    mode="walking"
                )
            except ApiError as e:
                raise RuntimeError(f"Walking directions failed for {a['title']}→{b['title']}: {e}")

            if not walk_resp or not walk_resp[0].get('legs'):
                raise RuntimeError(f"No walking route between {a['title']} and {b['title']}")

            walk_minutes = walk_resp[0]['legs'][0]['duration']['value'] / 60

            # Decide mode
            if walk_minutes <= walking_threshold:
                chosen, mode, dur = walk_resp, "walking", int(walk_minutes)
            else:
                # Try bus
                try:
                    bus_resp = gmaps.directions(
                        (a['lat'], a['lng']),
                        (b['lat'], b['lng']),
                        mode="transit",
                        transit_mode=["bus"]
                    )
                except ApiError:
                    bus_resp = []

                if bus_resp and bus_resp[0].get('legs'):
                    chosen, mode = bus_resp, "bus"
                    dur = int(bus_resp[0]['legs'][0]['duration']['value'] / 60)
                else:
                    # Try train
                    try:
                        train_resp = gmaps.directions(
                            (a['lat'], a['lng']),
                            (b['lat'], b['lng']),
                            mode="transit",
                            transit_mode=["train"]
                        )
                    except ApiError:
                        train_resp = []

                    if train_resp and train_resp[0].get('legs'):
                        chosen, mode = train_resp, "train"
                        dur = int(train_resp[0]['legs'][0]['duration']['value'] / 60)
                    else:
                        # Fallback to walking
                        chosen, mode, dur = walk_resp, "walking", int(walk_minutes)

            steps = chosen[0]['legs'][0].get('steps', [])
            description = "\n".join(step.get('html_instructions', '') for step in steps)

            enriched.append({
                "title": f"{mode.title()} to {b['title']}",
                "description": description,
                "duration": dur,
                "spacing_after": 0,
                "link": (
                    f"https://www.google.com/maps/dir/?api=1"
                    f"&origin={a['lat']},{a['lng']}"
                    f"&destination={b['lat']},{b['lng']}"
                    f"&travelmode={mode}"
                ),
                "city": "",
                "region": "",
                "event_type": mode,
                "cost": 0.0,
                "color": "#ADE1F9"
            })

        # Append last visit and replace
        enriched.append(plan['events'][-1])
        plan['events'] = enriched
        return plan

    def _populate_plan_into_calendar(self, plan):
        # start at first day, 6 AM
        cursor = QDateTime(self.scene.start_date, QTime(START_HOUR, 0))

        for ev in plan.get("events", []):
            # 1) Enrich the link via Google Maps if blank
            if not ev.get("link") and ev.get("title"):
                query = ev["title"]
                if ev.get("city"):
                    query += f", {ev['city']}"
                try:
                    res = self.gmaps.find_place(query, input_type="textquery")
                    pid = res["candidates"][0]["place_id"]
                    ev["link"] = f"https://www.google.com/maps/place/?q=place_id:{pid}"
                except Exception:
                    pass

            # 2) create the visual EventItem
            slots = ev["duration"] / self.scene.slot_minutes
            height = slots * self.scene.cell_h
            rect = QRectF(0, 0, self.scene.cell_w, height)
            item = EventItem(rect, ev["title"], QColor(ev.get("color", "#FFA")))
            item.description = ev.get("description", "")
            item.link = ev.get("link", "")
            item.city = ev.get("city", "")
            item.region = ev.get("region", "")
            item.event_type = ev.get("event_type", "")
            item.cost = ev.get("cost", 0.0)

            # 3) position it according to cursor
            day_idx = self.scene.start_date.daysTo(cursor.date())
            x = TIME_LABEL_WIDTH + day_idx * self.scene.cell_w

            mins_from_start = (cursor.time().hour() * 60 + cursor.time().minute())
            y = HEADER_HEIGHT + ((mins_from_start - START_HOUR * 60) // self.scene.slot_minutes) * self.scene.cell_h

            item.setPos(x, y)
            self.scene.addItem(item)

            # 4) advance the cursor
            advance = ev["duration"] + ev.get("spacing_after", 0)
            cursor = cursor.addSecs(int(advance) * 60)

    def on_export_plan(self):
        items = [it for it in self.scene.selectedItems() if isinstance(it, EventItem)]
        if not items:
            QMessageBox.warning(self, "No Selection", "Select one or more events to export.")
            return

        # build a list of (start_datetime, duration, item)
        records = []
        for ev in items:
            col = int((ev.pos().x() - TIME_LABEL_WIDTH) // self.scene.cell_w)
            row = int((ev.pos().y() - HEADER_HEIGHT) // self.scene.cell_h)
            day = self.scene.start_date.addDays(col)
            mins = START_HOUR * 60 + row * self.scene.slot_minutes
            start_dt = QDateTime(day, QTime(mins // 60, mins % 60))
            dur = int(ev.rect().height() / self.scene.cell_h) * self.scene.slot_minutes
            records.append((start_dt, dur, ev))

        # sort by start time
        records.sort(key=lambda r: r[0])

        plan = {
            "plan_name": "untitled",
            "default_slot_minutes": self.scene.slot_minutes,
            "events": []
        }

        # compute spacing AFTER each event
        for i, (start_dt, dur, ev) in enumerate(records):
            if i < len(records) - 1:
                next_start = records[i + 1][0]
                end_dt = start_dt.addSecs(dur * 60)
                spacing = end_dt.secsTo(next_start) // 60
            else:
                spacing = 0

            plan["events"].append({
                "title": ev.text.toPlainText(),
                "description": ev.description,
                "duration": dur,
                "spacing_after": spacing,
                "link": ev.link,
                "city": ev.city,
                "region": ev.region,
                "event_type": ev.event_type,
                "cost": ev.cost,
                "color": ev.color.name()
            })

        path, _ = QFileDialog.getSaveFileName(self, "Export Plan JSON", "", "JSON Files (*.json)")
        if path:
            with open(path, 'w') as f:
                json.dump(plan, f, indent=2)
            QMessageBox.information(self, "Exported", f"Plan exported to {path}")

    # --- IMPORT FUNCTION ---
    def on_import_plan(self):
        sel = self.scene.selectedItems()
        if not sel or not isinstance(sel[0], EventItem):
            QMessageBox.warning(self, "Select Anchor", "Select a single event cell to serve as start point.")
            return
        anchor = sel[0]

        # remove the anchor cell from scene (and DB if it has one)
        self.scene.removeItem(anchor)
        if getattr(anchor, 'db_id', None):
            self.db.conn.execute("DELETE FROM event WHERE id=?", (anchor.db_id,))
            self.db.conn.commit()

        path, _ = QFileDialog.getOpenFileName(self, "Import Plan JSON", "", "JSON Files (*.json)")
        if not path:
            return

        with open(path) as f:
            plan = json.load(f)

        # compute starting datetime from anchor’s former position
        ax, ay = anchor.pos().x(), anchor.pos().y()
        acol = int((ax - TIME_LABEL_WIDTH) // self.scene.cell_w)
        arow = int((ay - HEADER_HEIGHT) // self.scene.cell_h)
        aday = self.scene.start_date.addDays(acol)
        amin = START_HOUR * 60 + arow * self.scene.slot_minutes
        cursor_dt = QDateTime(aday, QTime(amin // 60, amin % 60))

        for ev in plan["events"]:
            slots = ev["duration"] / self.scene.slot_minutes
            height = slots * self.scene.cell_h
            rect = QRectF(0, 0, self.scene.cell_w, height)
            item = EventItem(rect, ev["title"], QColor(ev.get("color", "#FFA")))
            item.description = ev.get("description", "")
            item.link = ev.get("link", "")
            item.city = ev.get("city", "")
            item.region = ev.get("region", "")
            item.event_type = ev.get("event_type", "")
            item.cost = ev.get("cost", 0.0)

            # position the new item
            day_offset = self.scene.start_date.daysTo(cursor_dt.date())
            x = TIME_LABEL_WIDTH + day_offset * self.scene.cell_w
            mins_from_start = (cursor_dt.time().hour() * 60 + cursor_dt.time().minute()) - START_HOUR * 60
            y = HEADER_HEIGHT + (mins_from_start // self.scene.slot_minutes) * self.scene.cell_h
            item.setPos(x, y)
            self.scene.addItem(item)

            # advance cursor by duration + spacing_after
            cursor_dt = cursor_dt.addSecs((ev["duration"] + ev.get("spacing_after", 0)) * 60)

            # (optional) insert into DB if schedule is loaded
            if self.current_schedule_id:
                self.db.insert_event(
                    self.current_schedule_id,
                    ev['title'],  # title
                    ev.get("description", ""),  # description
                    cursor_dt,  # start_dt
                    cursor_dt.addSecs(ev["duration"] * 60),  # end_dt
                    ev["duration"],  # duration
                    ev.get("link", ""), ev.get("city", ""), ev.get("region", ""),
                    ev.get("event_type", ""), ev.get("cost", 0.0), ev.get("color", "#FFA"),
                    x, y, self.scene.cell_w, height
                )

    def toggle_details_panel(self):
        if self.detailsPanel.isVisible():
            self.detailsPanel.hide()
        else:
            p = self.schedule_central
            self.detailsPanel.setVisible(not self.detailsPanel.isVisible())
            self.detailsPanel.show()
            self.detailsPanel.raise_()

    def toggle_settings_panel(self):
        if self.settingsPanel.isVisible():
            self.settingsPanel.hide()
        else:
            p = self.schedule_central
            self.settingsPanel.setGeometry(
                p.width() - self.settingsPanelWidth,
                0,
                self.settingsPanelWidth,
                p.height()
            )
            self.settingsPanel.show()
            self.settingsPanel.raise_()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        # Keep overlays aligned when Schedule tab is visible
        if self.tab_widget.currentIndex() == 1:
            if self.detailsPanel.isVisible():
                p = self.schedule_central
                self.detailsPanel.setGeometry(0, 0, self.sidePanelWidth, p.height())
            if self.settingsPanel.isVisible():
                p = self.schedule_central
                self.settingsPanel.setGeometry(
                    p.width() - self.settingsPanelWidth, 0,
                    self.settingsPanelWidth, p.height()
                )

    def apply_settings(self):
        # 1) Pull the new values from the spin-boxes
        self.visible_columns = self.daysSpin.value()
        self.slot_minutes = self.slotSpin.value()

        # 2) Update the schedule record in SQLite
        if self.current_schedule_id is not None:
            # name & date‐range haven’t changed, but visible_columns & slot_minutes did:
            self.db.update_schedule(
                self.current_schedule_id,
                self.current_schedule_name,
                self.scene.start_date,
                self.scene.end_date,
                self.visible_columns,
                self.slot_minutes
            )

        # 3) Re-build the grid (this clears out every EventItem)
        sd, ed = self.scene.start_date, self.scene.end_date
        self.init_calendar(sd, ed)

        # 4) If we have a loaded schedule, re-populate & re-persist w/h
        if self.current_schedule_id is not None:
            # load the canonical event data (duration, start_dt, etc.)
            # load canonical data (duration, start_dt, etc.)
            _, _, _, _, _, evs = self.db.load_schedule(self.current_schedule_name)

            for ev in evs:
                # 1) compute size from duration + new slot interval
                slots = ev['duration'] / self.slot_minutes
                new_h = slots * self.scene.cell_h
                new_w = self.scene.cell_w

                # 2) compute column/day from start_dt
                col = self.scene.start_date.daysTo(ev['start_dt'].date())
                mins = ev['start_dt'].time().hour() * 60 + ev['start_dt'].time().minute()
                row = (mins - START_HOUR * 60) // self.slot_minutes

                x = TIME_LABEL_WIDTH + col * self.scene.cell_w
                y = HEADER_HEIGHT + row * self.scene.cell_h

                # 3) drop a fresh EventItem at the newly computed geometry
                rect = QRectF(0, 0, new_w, new_h)
                item = EventItem(rect, ev['title'], ev['color'])
                item.db_id = ev['id']
                item.description = ev['description']
                item.link = ev['link']
                item.city = ev['city']
                item.region = ev['region']
                item.event_type = ev['type']
                item.cost = ev['cost']
                item.group_id = ev['group_id']
                item.setPos(x, y)
                self.scene.addItem(item)


                self.db.update_event(
                    ev['id'],
                    x=x,
                    y=y,
                    w=new_w,
                    h=new_h
                )
                self._reload_location_views()

            # re-link multi-day pieces if you use group_id…
            # (your existing code for that goes here)

    def init_calendar(self, start, end):
        self.scene = CalendarScene(start, end,
                                   slot_minutes=self.slot_minutes,
                                   cell_width=DEFAULT_CELL_WIDTH,
                                   cell_height=DEFAULT_CELL_HEIGHT)
        self.scene.main_window = self
        self.scene.db = self.db
        self.view.setScene(self.scene)
        self.view.setScene(self.scene)
        # now this QGraphicsScene wrapper gets a `.db` attribute:
        self.view.scene().db = self.db
        # rebuild time labels
        while self.timeLayout.count():
            item = self.timeLayout.takeAt(0)
            w = item.widget()
            if w: w.deleteLater()
        slots = ((END_HOUR - START_HOUR) * 60) // self.slot_minutes
        for j in range(slots):
            mins = START_HOUR*60 + j*self.slot_minutes
            lbl = QLabel(f"{mins//60:02d}:{mins%60:02d}")
            lbl.setFixedHeight(self.scene.cell_h)
            lbl.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.timeLayout.addWidget(lbl)

        total_days = start.daysTo(end) + 1
        cols = min(self.visible_columns, total_days)
        visible_w = cols * self.scene.cell_w
        # — instead of forcing the view to exactly your content height —
        # let it expand in the layout so that if content > viewport you'll get scrollbars.
        self.view.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        # optionally enforce a minimum so you never get too tiny:
        min_h = HEADER_HEIGHT + 10 * self.scene.cell_h  # show at least 10 slots
        self.view.setMinimumHeight(min_h)
        self.timeFrame.setMinimumHeight(min_h)

        self.scene.selectionChanged.connect(self.on_event_selected)
        self.scene.eventCreated.connect(self.on_event_created)

    def on_new_schedule(self):
        dlg = DateRangeDialog(self)
        if dlg.exec_():
            d1, d2 = dlg.dates()
            if d1 > d2:
                QMessageBox.warning(self, "Invalid Range", "Start must be ≤ end.")
                return
            self.current_schedule_id = None
            self.init_calendar(d1, d2)

    def on_save_schedule(self):
        name, ok = QInputDialog.getText(self, "Save Schedule", "Schedule Name:")
        if not ok or not name.strip():
            return
        start, end = self.scene.start_date, self.scene.end_date
        if self.current_schedule_id is None:
            sid = self.db.save_schedule(
                name, start, end,
                self.visible_columns, self.slot_minutes
            )
            self.current_schedule_id = sid
        else:
            sid = self.current_schedule_id
            self.db.update_schedule(
                sid, name, start, end,
                self.visible_columns, self.slot_minutes
            )
            self.db.delete_events_for_schedule(sid)

        evs = []
        for it in self.scene.items():
            if not isinstance(it, EventItem):
                continue
            x, y = it.pos().x(), it.pos().y()
            col = int((x - TIME_LABEL_WIDTH) // self.scene.cell_w)
            row = int((y - HEADER_HEIGHT) // self.scene.cell_h)
            day = start.addDays(col)
            mins = START_HOUR*60 + row*self.scene.slot_minutes
            sd = QDateTime(day, QTime(mins//60, mins%60))
            slots = it.rect().height()/self.scene.cell_h
            dur = int(slots*self.scene.slot_minutes)
            ed = sd.addSecs(dur*60)
            evs.append({
                'item': it, 'title': it.text.toPlainText(),
                'description': it.description,
                'start_dt': sd, 'end_dt': ed, 'duration': dur,
                'link': it.link, 'city': it.city, 'region': it.region,
                'type': it.event_type, 'cost': it.cost,
                'color': it.color.name(), 'x': x, 'y': y,
                'w': it.rect().width(), 'h': it.rect().height(),
                'group_id': it.group_id
            })
        for ev in evs:
            dbid = self.db.insert_event(
                sid,
                ev['title'],  # title
                ev['description'],  # description
                ev['start_dt'],  # start_dt
                ev['end_dt'],  # end_dt
                ev['duration'],  # duration
                ev['link'], ev['city'], ev['region'],
                ev['type'], ev['cost'], ev['color'],
                ev['x'], ev['y'], ev['w'], ev['h'],
                ev['group_id']  # optional group_id
            )
            ev['item'].db_id = dbid

        QMessageBox.information(self, "Saved", f"Schedule '{name}' saved.")
        self.current_schedule_name = name
        self._populate_location_dropdown()

    def on_load_schedule(self):
        names = self.db.list_schedules()
        if not names:
            QMessageBox.information(self, "No Schedules", "No saved schedules.")
            return
        name, ok = QInputDialog.getItem(self, "Load Schedule", "Choose one:", names, editable=False)
        if not ok:
            return
        try:
            sid, start, end, vis_cols, slot_min, evs = self.db.load_schedule(name)
        except KeyError:
            QMessageBox.critical(self, "Error", f"Schedule not found: {name}")
            return
        self.current_schedule_id = sid
        self.visible_columns = vis_cols
        self.slot_minutes = slot_min
        self.daysSpin.setValue(vis_cols)
        self.slotSpin.setValue(slot_min)

        self.init_calendar(start, end)

        items = []
        for it in list(self.scene.items()):
            if isinstance(it, EventItem):
                self.scene.removeItem(it)

        for ev in evs:
            rect = QRectF(0,0,ev['w'],ev['h'])
            it = EventItem(rect, ev['title'], ev['color'])
            it.db_id = ev['id']
            it.description = ev['description']
            it.link = ev['link']; it.city = ev['city']; it.region = ev['region']
            it.event_type = ev['type']; it.cost = ev['cost']; it.group_id = ev['group_id']
            it.setPos(ev['x'], ev['y'])
            self.scene.addItem(it)
            items.append(it)

        groups = {}
        for it in items:
            if it.group_id:
                groups.setdefault(it.group_id, []).append(it)
        for grp in groups.values():
            if len(grp) > 1:
                for a in grp:
                    a.linked_items = [b for b in grp if b is not a]

        self.scene.clearSelection()
        if items:
            items[0].setSelected(True)
        self.current_schedule_name = name
        self.current_schedule_id = sid
        self._populate_location_dropdown()
        self._reload_location_views()

    def _populate_location_dropdown(self):
        self.locationCombo.blockSignals(True)
        self.locationCombo.clear()
        self.locationCombo.addItem("All")
        for loc in self.db.list_locations(self.current_schedule_id):
            self.locationCombo.addItem(loc)
        self.locationCombo.blockSignals(False)

    def _reload_location_views(self):
        sid = self.current_schedule_id
        loc = self.locationCombo.currentText()
        etype = self.eventTypeCombo.currentText().lower()
        use_conv = self.convertCurrencyCheck.isChecked()
        rate     = self.conversionRateSpin.value()

        # — things to eat & do always filtered by location —
        self._populate_table(self.eatTable, self.db.list_eat_items(sid, loc))
        self._populate_table(self.doTable, self.db.list_do_items(sid, loc))

        # — hotels only if “All” or “hotel” selected —
        hotel_rows = self.db.list_hotels(sid, loc) if etype in ("all", "hotel") else []
        self._populate_table(self.hotelTable, hotel_rows)

        # — reservations only if “All” or “reservation” —
        res_rows = self.db.list_reservations(sid, loc) if etype in ("all", "reservation") else []
        self._populate_table(self.resTable, res_rows)

        # — total-cost tab only —
        if self.locTabWidget.currentWidget() is self.totalCostTab:
            combined = []

            if etype in ("all", "hotel"):
                for (_id, _loc, item, _link, _desc, _media, *_rest, cost) in self.db.list_hotels(sid, loc):
                    combined.append(("Hotel", item, cost))

            if etype in ("all", "reservation"):
                for (_id, _loc, item, _link, _desc, _media, *_rest) in self.db.list_reservations(sid, loc):
                    # look up the cost in the event table
                    row = self.db.conn.execute(
                        "SELECT cost FROM event WHERE schedule_id=? AND group_id=? AND event_type='reservation'",
                        (sid, str(_id))
                    ).fetchone()
                    combined.append(("Reservation", item, row[0] if row else 0.0))

            # “other events” (not hotel/reservation)
            if etype == "all":
                q = """
                    SELECT title, cost
                      FROM event
                     WHERE schedule_id=?
                       AND event_type NOT IN ('hotel','reservation')
                       AND (group_id IS NULL OR group_id = '')
                """
                args = [sid]
            else:
                q = """
                    SELECT title, cost
                      FROM event
                     WHERE schedule_id=?
                       AND event_type = ?
                       AND (group_id IS NULL OR group_id = '')
                """
                args = [sid, self.eventTypeCombo.currentText()]

            for title, cost in self.db.conn.execute(q, args).fetchall():
                combined.append(("Event", title, cost))

            symbol = "€" if use_conv else "£"

            # 4) Populate the table (3 columns: Type, Item, Cost)
            self.totalCostTable.clearContents()
            self.totalCostTable.setColumnCount(3)
            self.totalCostTable.setHorizontalHeaderLabels(["Type", "Item", "Cost"])
            self.totalCostTable.setRowCount(len(combined))
            total = 0.0
            for i, (kind, name, cost_val) in enumerate(combined):
                display = cost_val * rate if use_conv else cost_val


                self.totalCostTable.setItem(i, 0, QTableWidgetItem(kind))
                self.totalCostTable.setItem(i, 1, QTableWidgetItem(name))
                self.totalCostTable.setItem(i, 2, QTableWidgetItem(f"{symbol}{display:.2f}"))

                total += cost_val

            grand = total * rate if use_conv else total
            self.totalCostSummaryLabel.setText(f"Total: {symbol}{grand:.2f}")

    def _populate_table(self, table: QTableWidget, data: list):
        """
        data is a list of tuples: (id, col1, col2, ..., colN).
        The first element is always the PK, which we stash into Qt.UserRole
        on the first visible column.
        """
        # prevent cellChanged from firing during bulk load
        table.blockSignals(True)

        table.clearContents()
        table.setRowCount(len(data))
        for r, row in enumerate(data):
            row_id = row[0]
            for c, val in enumerate(row[1:]):
                item = QTableWidgetItem(str(val))
                if c == 0:
                    # stash the real ID so on_cell_changed can find it
                    item.setData(Qt.UserRole, row_id)
                # if you want certain columns editable, tweak flags here...
                table.setItem(r, c, item)

        table.blockSignals(False)

    def on_event_created(self, ev: EventItem):
        if self.current_schedule_id is not None:
            x, y = ev.pos().x(), ev.pos().y()
            col = int((x - TIME_LABEL_WIDTH)//self.scene.cell_w)
            row = int((y - HEADER_HEIGHT)//self.scene.cell_h)
            day = self.scene.start_date.addDays(col)
            mins = START_HOUR*60 + row*self.scene.slot_minutes
            sd = QDateTime(day, QTime(mins//60, mins%60))
            slots = ev.rect().height()/self.scene.cell_h
            dur = int(slots*self.scene.slot_minutes)
            ed = sd.addSecs(dur*60)
            ev.db_id = self.db.insert_event(
                self.current_schedule_id,
                ev.text.toPlainText(),  # title
                ev.description,  # description
                sd,  # start_dt
                ed,  # end_dt
                dur,  # duration
                ev.link, ev.city, ev.region,
                ev.event_type, ev.cost, ev.color.name(),
                x, y, ev.rect().width(), ev.rect().height(),
                ev.group_id  # optional group_id
            )
        else:
            ev.db_id = None

    def on_event_selected(self):
        if self._handling_selection:
            return
        self._handling_selection = True
        items = self.view.scene().selectedItems()
        if items and isinstance(items[0], EventItem):
            ev = items[0]
            for sib in ev.linked_items:
                sib.setSelected(True)

        widgets = [
            self.titleEdit, self.linkEdit, self.cityEdit, self.regionEdit,
            self.typeEdit, self.descriptionEdit, self.startEdit, self.endEdit, self.durationSpin,
            self.costEdit
        ]
        for w in widgets: w.blockSignals(True)

        if not items or not isinstance(items[0], EventItem):
            self.current_event = None
            for w in widgets:
                if isinstance(w, QSpinBox) or isinstance(w, QDoubleSpinBox):
                    w.setValue(0)
                else:
                    w.clear()
            self.colorBtn.setStyleSheet("")
        else:
            ev = items[0]
            parts = [ev] + ev.linked_items
            starts, ends = [], []
            for p in parts:
                x,y = p.pos().x(), p.pos().y()
                col = int((x - TIME_LABEL_WIDTH)//self.scene.cell_w)
                row = int((y - HEADER_HEIGHT)//self.scene.cell_h)
                day = self.scene.start_date.addDays(col)
                mins = START_HOUR*60 + row*self.scene.slot_minutes
                sd = QDateTime(day, QTime(mins//60, mins%60))
                dur = int((p.rect().height()/self.scene.cell_h)*self.scene.slot_minutes)
                ed = sd.addSecs(dur*60)
                starts.append(sd); ends.append(ed)
            sd_min = min(starts)
            ed_max = max(ends)
            total_dur = sd_min.secsTo(ed_max)//60

            self.current_event = ev
            self.titleEdit.setText(ev.text.toPlainText())
            self.linkEdit.setText(ev.link)
            self.cityEdit.setText(ev.city)
            self.regionEdit.setText(ev.region)
            self.typeEdit.setText(ev.event_type)
            self.descriptionEdit.setText(ev.description)
            self.costEdit.setValue(ev.cost)
            self.colorBtn.setStyleSheet(f"background-color:{ev.color.name()}")
            self.startEdit.setDateTime(sd_min)
            self.endEdit.setDateTime(ed_max)
            self.durationSpin.setValue(total_dur)

        for w in widgets: w.blockSignals(False)
        self._handling_selection = False

    def update_event_title(self):
        ev = self.current_event
        if not ev: return
        new = self.titleEdit.text()
        for p in [ev] + ev.linked_items:
            p.set_title(new)
            if getattr(p, 'db_id', None):
                self.db.update_event(p.db_id, title=new)
                self._reload_location_views()

    def update_event_link(self):
        ev = self.current_event
        if not ev: return
        new = self.linkEdit.text()
        for p in [ev] + ev.linked_items:
            p.link = new
            if getattr(p, 'db_id', None):
                self.db.update_event(p.db_id, link=new)
                self._reload_location_views()

    def update_event_description(self):
        ev = self.current_event
        if not ev:
            return
        new = self.descriptionEdit.text()
        for part in [ev] + ev.linked_items:
            part.description = new
            if getattr(part, 'db_id', None):
                self.db.update_event(part.db_id, description=new)
                self._reload_location_views()

    def update_event_city(self):
        ev = self.current_event
        if not ev: return
        new = self.cityEdit.text()
        for p in [ev] + ev.linked_items:
            p.city = new
            if getattr(p, 'db_id', None):
                self.db.update_event(p.db_id, city=new)
                self._reload_location_views()

    def update_event_region(self):
        ev = self.current_event
        if not ev: return
        new = self.regionEdit.text()
        for p in [ev] + ev.linked_items:
            p.region = new
            if getattr(p, 'db_id', None):
                self.db.update_event(p.db_id, region=new)
                self._reload_location_views()

    def update_event_type(self):
        ev = self.current_event
        if not ev: return
        new = self.typeEdit.text()
        for p in [ev] + ev.linked_items:
            p.event_type = new
            if getattr(p, 'db_id', None):
                self.db.update_event(p.db_id, event_type=new)
                self._reload_location_views()

    def update_event_cost(self, v):
        ev = self.current_event
        if not ev: return
        for p in [ev] + ev.linked_items:
            p.cost = v
            if getattr(p, 'db_id', None):
                self.db.update_event(p.db_id, cost=v)
                self._reload_location_views()

    def choose_event_color(self):
        ev = self.current_event
        if not ev: return
        col = QColorDialog.getColor(ev.color, self, "Choose Event Color")
        if col.isValid():
            for p in [ev] + ev.linked_items:
                p.color = col
                p.setBrush(QBrush(col))
                if getattr(p, 'db_id', None):
                    self.db.update_event(p.db_id, color=col.name())
                    self._reload_location_views()
            self.colorBtn.setStyleSheet(f"background-color:{col.name()}")

    def update_event_time(self):
        ev = self.current_event
        if not ev: return
        dt = self.startEdit.dateTime()
        col = self.scene.start_date.daysTo(dt.date())
        row = (dt.time().hour()*60 + dt.time().minute() - START_HOUR*60)//self.scene.slot_minutes
        x = TIME_LABEL_WIDTH + col*self.scene.cell_w
        y = HEADER_HEIGHT + row*self.scene.cell_h
        ev.setPos(x, y)
        if getattr(ev, 'db_id', None):
            self.db.update_event(ev.db_id, start_dt=dt, x=x, y=y)
        new_dur = dt.secsTo(self.endEdit.dateTime())//60
        self.durationSpin.setValue(new_dur)
        if getattr(ev, 'db_id', None):
            self.db.update_event(ev.db_id, duration=new_dur, end_dt=self.endEdit.dateTime())
            self._reload_location_views()

    def update_event_duration(self, minutes):
        ev = self.current_event
        if not ev: return
        slots = minutes / self.scene.slot_minutes
        h = slots * self.scene.cell_h
        r = ev.rect()
        ev.setRect(QRectF(r.x(), r.y(), r.width(), h))

        # 2) persist *all* changed fields in one go
        self.db.update_event(
            ev.db_id,
            duration=minutes,
            end_dt=self.endEdit.dateTime(),
            h=h
        )
        self._reload_location_views()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = MainWindow()
    win.showMaximized()
    win.show()
    sys.exit(app.exec_())