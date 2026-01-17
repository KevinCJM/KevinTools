import sys
import tempfile
import types
import unittest
from pathlib import Path
from unittest import mock

from src import gui
from src import config


class GuiTests(unittest.TestCase):
    def tearDown(self) -> None:
        sys.modules.pop("PySide6", None)

    def test_require_pyside6_missing(self) -> None:
        sys.modules.pop("PySide6", None)
        with mock.patch("builtins.__import__") as mocked_import:
            def _raise(name, *args, **kwargs):
                if name == "PySide6":
                    raise ImportError("missing")
                return __import__(name, *args, **kwargs)

            mocked_import.side_effect = _raise
            with self.assertRaises(SystemExit):
                gui._require_pyside6()

    def test_require_pyside6_present(self) -> None:
        dummy = types.SimpleNamespace(
            QtCore=object(),
            QtGui=object(),
            QtWidgets=object(),
        )
        sys.modules["PySide6"] = dummy
        QtCore, QtGui, QtWidgets = gui._require_pyside6()
        self.assertIs(QtCore, dummy.QtCore)
        self.assertIs(QtGui, dummy.QtGui)
        self.assertIs(QtWidgets, dummy.QtWidgets)

    def test_main_flow(self) -> None:
        fixtures_dir = Path(__file__).resolve().parents[1] / "tests" / "fixtures"
        fixture_path = fixtures_dir / "TPL_BASIC.docx"
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = Path(tmpdir) / "out.json"
            state = {
                "window": None,
                "open_paths": [],
                "open_results": [False, True],
                "warnings": 0,
                "criticals": 0,
                "parse_enabled_initial": None,
                "parse_enabled_after": None,
                "allow_fallback_default": None,
                "busy_snapshot": {},
                "idle_snapshot": {},
                "status_after_success": None,
                "status_after_failure": None,
            }

            class DummySignal:
                def __init__(self):
                    self._callbacks = []

                def connect(self, callback):
                    self._callbacks.append(callback)

                def emit(self, *args, **kwargs):
                    for callback in list(self._callbacks):
                        try:
                            callback(*args, **kwargs)
                        except TypeError:
                            callback()

            class Signal:
                def __init__(self, *args, **kwargs):
                    self._name = None

                def __set_name__(self, owner, name):
                    self._name = f"__signal_{name}"

                def __get__(self, instance, owner):
                    if instance is None:
                        return self
                    signal = instance.__dict__.get(self._name)
                    if signal is None:
                        signal = DummySignal()
                        instance.__dict__[self._name] = signal
                    return signal

            def Slot(*args, **kwargs):
                def _wrap(func):
                    return func

                return _wrap

            class QObject:
                def moveToThread(self, thread):
                    return None

                def deleteLater(self):
                    return None

            class QThread:
                def __init__(self):
                    self.started = DummySignal()
                    self.finished = DummySignal()
                    self._running = False

                def start(self):
                    self._running = True
                    self.started.emit()
                    self._running = False
                    self.finished.emit()

                def quit(self):
                    return None

                def deleteLater(self):
                    return None

                def isRunning(self):
                    return self._running

            class QUrl:
                @staticmethod
                def fromLocalFile(path):
                    return path

            class QTimer:
                def __init__(self, *_):
                    self.timeout = DummySignal()

                def setSingleShot(self, *_):
                    return None

                def start(self, *_):
                    return None

            QtCore = types.SimpleNamespace(
                Signal=Signal,
                Slot=Slot,
                QObject=QObject,
                QThread=QThread,
                QUrl=QUrl,
                QTimer=QTimer,
                Qt=types.SimpleNamespace(Checked=2, Unchecked=0, UserRole=32),
            )

            class QDesktopServices:
                @staticmethod
                def openUrl(path):
                    state["open_paths"].append(path)
                    if state["open_results"]:
                        return state["open_results"].pop(0)
                    return True

            QtGui = types.SimpleNamespace(QDesktopServices=QDesktopServices)

            class QtWidgets:
                class QWidget:
                    def __init__(self):
                        self._enabled = True

                    def show(self):
                        state["window"] = self

                    def setWindowTitle(self, _):
                        return None

                    def resize(self, *_):
                        return None

                    def setLayout(self, _):
                        return None

                    def setEnabled(self, value):
                        self._enabled = value

                class QLineEdit(QWidget):
                    def __init__(self):
                        super().__init__()
                        self._text = ""
                        self.textChanged = DummySignal()

                    def setText(self, value):
                        self._text = value
                        self.textChanged.emit(value)

                    def text(self):
                        return self._text

                class QPushButton(QWidget):
                    def __init__(self, _=None):
                        super().__init__()
                        self.clicked = DummySignal()

                class QCheckBox(QWidget):
                    def __init__(self, _=None):
                        super().__init__()
                        self._checked = False

                    def setChecked(self, value):
                        self._checked = value

                    def isChecked(self):
                        return self._checked

                class QLabel(QWidget):
                    def __init__(self, text=None):
                        super().__init__()
                        self._text = text or ""

                    def setText(self, text):
                        self._text = text

                    def text(self):
                        return self._text

                class QComboBox(QWidget):
                    def __init__(self):
                        super().__init__()
                        self._items = []
                        self._current_index = 0
                        self.currentIndexChanged = DummySignal()
                        self._signals_blocked = False

                    def addItem(self, text, userData=None):
                        self._items.append((text, userData))

                    def clear(self):
                        self._items = []
                        self._current_index = 0

                    def blockSignals(self, blocked):
                        self._signals_blocked = bool(blocked)

                    def findData(self, value):
                        for idx, (_, data) in enumerate(self._items):
                            if data == value:
                                return idx
                        return -1

                    def setCurrentIndex(self, index):
                        if not self._items:
                            self._current_index = 0
                            return None
                        if index < 0:
                            index = 0
                        if index >= len(self._items):
                            index = len(self._items) - 1
                        self._current_index = index
                        if not self._signals_blocked:
                            self.currentIndexChanged.emit(index)

                    def count(self):
                        return len(self._items)

                    def itemData(self, index):
                        if 0 <= index < len(self._items):
                            return self._items[index][1]
                        return None

                    def currentData(self):
                        if not self._items:
                            return None
                        return self._items[self._current_index][1]

                    def currentText(self):
                        if not self._items:
                            return ""
                        return self._items[self._current_index][0]


                class QTableWidgetItem:
                    def __init__(self, text=""):
                        self._text = text
                        self._data = {}
                        self._check_state = None

                    def setText(self, text):
                        self._text = text

                    def text(self):
                        return self._text

                    def setData(self, role, value):
                        self._data[role] = value

                    def data(self, role):
                        return self._data.get(role)

                    def setCheckState(self, value):
                        self._check_state = value

                    def checkState(self):
                        return self._check_state

                class QTableWidget(QWidget):
                    def __init__(self):
                        super().__init__()
                        self._rows = []
                        self._cols = 0
                        self._widgets = {}
                        self.itemChanged = DummySignal()

                    def setColumnCount(self, count):
                        self._cols = count

                    def setHorizontalHeaderLabels(self, _):
                        return None

                    def horizontalHeader(self):
                        class Header:
                            def setStretchLastSection(self, _):
                                return None

                        return Header()

                    def horizontalHeaderItem(self, _):
                        return None

                    def setRowCount(self, count):
                        self._rows = [[None for _ in range(self._cols)] for _ in range(count)]

                    def rowCount(self):
                        return len(self._rows)

                    def insertRow(self, index):
                        self._rows.insert(index, [None for _ in range(self._cols)])

                    def removeRow(self, index):
                        if 0 <= index < len(self._rows):
                            self._rows.pop(index)

                    def setItem(self, row, col, item):
                        self._rows[row][col] = item

                    def item(self, row, col):
                        if 0 <= row < len(self._rows) and 0 <= col < self._cols:
                            return self._rows[row][col]
                        return None

                    def setCellWidget(self, row, col, widget):
                        self._widgets[(row, col)] = widget

                    def cellWidget(self, row, col):
                        return self._widgets.get((row, col))

                    def selectionModel(self):
                        class Model:
                            def selectedRows(self):
                                return []

                        return Model()

                class QProgressBar(QWidget):
                    def setRange(self, *_):
                        return None

                class QTextEdit(QWidget):
                    def __init__(self):
                        super().__init__()
                        self._text = ""

                    def append(self, text):
                        self._text += text + "\\n"

                    def setReadOnly(self, _):
                        return None

                    def setPlainText(self, text):
                        self._text = text

                    def clear(self):
                        self._text = ""

                class QGridLayout:
                    def addWidget(self, *_):
                        return None

                class QHBoxLayout:
                    def addWidget(self, *_):
                        return None

                    def addStretch(self, *_):
                        return None

                class QVBoxLayout:
                    def addLayout(self, *_):
                        return None

                    def addWidget(self, *_):
                        return None

                class QFileDialog:
                    open_return = ("", "")
                    save_return = ("", "")

                    @staticmethod
                    def getOpenFileName(*_):
                        return QtWidgets.QFileDialog.open_return

                    @staticmethod
                    def getSaveFileName(*_):
                        return QtWidgets.QFileDialog.save_return

                class QMessageBox:
                    @staticmethod
                    def warning(*_):
                        state["warnings"] += 1

                    @staticmethod
                    def critical(*_):
                        state["criticals"] += 1

                class QInputDialog:
                    @staticmethod
                    def getText(*_, **__):
                        return "", False

                class QApplication:
                    def __init__(self, _):
                        return None

                    def exec(self):
                        window = state["window"]
                        state["allow_fallback_default"] = window.allow_fallback_check.isChecked()
                        state["parse_enabled_initial"] = window.parse_button._enabled
                        window._choose_template()
                        state["parse_enabled_after"] = window.parse_button._enabled
                        window._choose_output()
                        window._update_actions()
                        window._set_busy(True)
                        state["busy_snapshot"] = {
                            "status": window.status_label.text(),
                            "parse_enabled": window.parse_button._enabled,
                            "template_enabled": window.template_edit._enabled,
                            "output_enabled": window.output_edit._enabled,
                            "open_enabled": window.open_button._enabled,
                        }
                        window._set_busy(False)
                        state["idle_snapshot"] = {
                            "status": window.status_label.text(),
                            "parse_enabled": window.parse_button._enabled,
                            "template_enabled": window.template_edit._enabled,
                            "output_enabled": window.output_edit._enabled,
                            "open_enabled": window.open_button._enabled,
                        }
                        window._open_output_dir()
                        window.output_edit.setText("")
                        window._open_output_dir()
                        window.template_edit.setText(str(fixture_path))
                        window.output_edit.setText(str(output_path))
                        window.allow_fallback_check.setChecked(True)
                        window._start_parse()
                        state["status_after_success"] = window.status_label.text()
                        window._handle_finished(False, "", "boom")
                        state["status_after_failure"] = window.status_label.text()
                        return 0

            QtWidgets.QFileDialog.open_return = (str(fixture_path), "")
            QtWidgets.QFileDialog.save_return = (str(output_path), "")

            with mock.patch.object(gui, "_require_pyside6", return_value=(QtCore, QtGui, QtWidgets)), \
                mock.patch.object(sys, "exit"):
                gui.main()

            self.assertTrue(output_path.exists())
            self.assertGreaterEqual(state["warnings"], 1)
            self.assertGreaterEqual(state["criticals"], 1)
            self.assertFalse(state["parse_enabled_initial"])
            self.assertTrue(state["parse_enabled_after"])
            self.assertTrue(state["allow_fallback_default"])
            self.assertEqual(state["open_paths"][0], str(output_path.parent))
            self.assertEqual(state["open_paths"][1], str(config.DEFAULT_OUTPUT_PATH.parent))
            self.assertEqual(state["busy_snapshot"]["status"], "解析中...")
            self.assertFalse(state["busy_snapshot"]["parse_enabled"])
            self.assertFalse(state["busy_snapshot"]["template_enabled"])
            self.assertFalse(state["busy_snapshot"]["output_enabled"])
            self.assertFalse(state["busy_snapshot"]["open_enabled"])
            self.assertEqual(state["idle_snapshot"]["status"], "就绪")
            self.assertTrue(state["idle_snapshot"]["parse_enabled"])
            self.assertTrue(state["idle_snapshot"]["template_enabled"])
            self.assertTrue(state["idle_snapshot"]["output_enabled"])
            self.assertTrue(state["idle_snapshot"]["open_enabled"])
            self.assertEqual(state["status_after_success"], "完成")
            self.assertEqual(state["status_after_failure"], "失败")


if __name__ == "__main__":
    unittest.main()
