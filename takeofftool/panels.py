from __future__ import annotations

from collections import defaultdict
from typing import Any

from PyQt5 import QtCore, QtGui, QtWidgets

from .viewer import PDFGraphicsView, HighlightItem, LineItem

# color options used for the combo boxes
color_options = {
    "Red": "#FF0000",
    "Green": "#00FF00",
    "Blue": "#0000FF",
    "Yellow": "#FFFF00",
    "Magenta": "#FF00FF",
    "Cyan": "#00FFFF",
}


class TakeoffPanel(QtWidgets.QWidget):
    """UI for per‑category take‑offs."""

    new_takeoff_signal = QtCore.pyqtSignal(dict)
    saveRequested = QtCore.pyqtSignal()
    importRequested = QtCore.pyqtSignal()
    totalsUpdated = QtCore.pyqtSignal()

    def __init__(self, parent=None, *, include_wire: bool = True):
        super().__init__(parent)
        self.include_wire = include_wire

        main_layout = QtWidgets.QVBoxLayout(self)
        main_layout.setContentsMargins(2, 2, 2, 2)
        main_layout.setSpacing(4)

        # control row ----------------------------------------------------
        ctrl = QtWidgets.QHBoxLayout()
        self.add_btn = QtWidgets.QPushButton("Add Takeoff")
        self.save_btn = QtWidgets.QPushButton("Save")
        ctrl.addWidget(self.add_btn)
        ctrl.addWidget(self.save_btn)
        main_layout.addLayout(ctrl)
        self.add_btn.clicked.connect(self.add_takeoff)
        self.save_btn.clicked.connect(self.saveRequested.emit)

        # totals ---------------------------------------------------------
        self.totals_label = QtWidgets.QLabel("Totals: Count=0; Hours=0.00")
        main_layout.addWidget(self.totals_label)

        # scroll area ----------------------------------------------------
        self.takeoff_items: list[dict] = []
        self.scroll = QtWidgets.QScrollArea()
        self.scroll.setWidgetResizable(True)
        main_layout.addWidget(self.scroll)
        container = QtWidgets.QWidget()
        self.scroll.setWidget(container)
        self.container_layout = QtWidgets.QVBoxLayout(container)
        self.container_layout.setAlignment(QtCore.Qt.AlignTop)
        self.pdf_view: PDFGraphicsView | None = None

    def setPdfView(self, view: PDFGraphicsView):
        self.pdf_view = view

    def add_takeoff(self):
        frame = QtWidgets.QFrame()
        frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        layout = QtWidgets.QVBoxLayout(frame)

        r1 = QtWidgets.QHBoxLayout()
        lbl = QtWidgets.QLabel(f"Takeoff {len(self.takeoff_items) + 1}")
        draw_btn = QtWidgets.QPushButton("Draw Highlight")
        color_cb = QtWidgets.QComboBox()
        color_cb.addItems(color_options.keys())
        r1.addWidget(lbl)
        r1.addWidget(draw_btn)
        r1.addWidget(color_cb)
        layout.addLayout(r1)

        r2 = QtWidgets.QHBoxLayout()
        count_lbl = QtWidgets.QLabel("Count: 0")
        name_f = QtWidgets.QLineEdit()
        name_f.setPlaceholderText("Name")
        labor_lbl = QtWidgets.QLabel("Labor:")
        labor_f = QtWidgets.QLineEdit()
        labor_f.setPlaceholderText("0.0")
        labor_f.setFixedWidth(60)
        labor_f.textChanged.connect(self.update_totals)
        r2.addWidget(count_lbl)
        r2.addWidget(name_f)
        r2.addWidget(labor_lbl)
        r2.addWidget(labor_f)

        if self.include_wire:
            type_cb = QtWidgets.QComboBox()
            type_cb.addItems(["NMD", "AC90", "SOW", "SJOW"])
            cable_cb = QtWidgets.QComboBox()
            cable_cb.addItems([
                "14-2",
                "14-3",
                "14-4",
                "12-2",
                "12-3",
                "12-4",
                "10-2",
                "10-3",
                "10-4",
                "8-2",
                "8-3",
                "8-4",
                "6-2",
                "6-3",
                "6-4",
                "4-2",
                "4-3",
                "4-4",
                "2-2",
                "2-3",
                "2-4",
                "1/0-3",
                "1/0-4",
                "2/0-3",
                "2/0-4",
                "3/0-3",
                "3/0-4",
                "4/0-3",
                "4/0-4",
                "250 MCM",
                "300 MCM",
                "350 MCM",
                "500 MCM",
            ])
            mat_cb = QtWidgets.QComboBox()
            mat_cb.addItems(["CU", "AL"])
            length_lbl = QtWidgets.QLabel("Length:")
            length_f = QtWidgets.QLineEdit()
            length_f.setPlaceholderText("0.0")
            length_f.setFixedWidth(60)
            length_f.textChanged.connect(self.update_totals)
            r2.addWidget(type_cb)
            r2.addWidget(cable_cb)
            r2.addWidget(mat_cb)
            r2.addWidget(length_lbl)
            r2.addWidget(length_f)

        del_btn = QtWidgets.QPushButton("Delete")
        notes_btn = QtWidgets.QPushButton("Show Notes")
        r2.addWidget(del_btn)
        r2.addWidget(notes_btn)
        layout.addLayout(r2)

        note_w = QtWidgets.QPlainTextEdit()
        note_w.setFixedHeight(100)
        note_w.setVisible(False)
        layout.addWidget(note_w)

        self.container_layout.addWidget(frame)

        item: dict[str, Any] = {
            "frame": frame,
            "label": lbl,
            "draw": draw_btn,
            "color": color_cb,
            "count": count_lbl,
            "name": name_f,
            "labor": labor_f,
            "notes": note_w,
            "highlights": [],
            "color_dropdown": color_cb,
            "name_field": name_f,
            "labor_field": labor_f,
            "notes_edit": note_w,
        }
        if self.include_wire:
            item.update(
                {
                    "wire_type": type_cb,
                    "wire_cable": cable_cb,
                    "wire_mat": mat_cb,
                    "wire_length": length_f,
                }
            )
        self.takeoff_items.append(item)

        draw_btn.clicked.connect(lambda _=False, it=item: self.new_takeoff_signal.emit(it))
        del_btn.clicked.connect(lambda _=False, it=item: self.delete_takeoff(it))
        notes_btn.clicked.connect(
            lambda _=False, b=notes_btn, n=note_w: (
                n.setVisible(not n.isVisible()),
                b.setText("Hide Notes" if n.isVisible() else "Show Notes"),
            )
        )

        self.update_totals()

    def delete_takeoff(self, item: dict):
        for h in item["highlights"]:
            if h.scene():
                h.scene().removeItem(h)
        item["frame"].setParent(None)
        self.takeoff_items.remove(item)
        self.update_totals()

    def update_count(self, takeoff_item: dict):
        valid = [h for h in takeoff_item["highlights"] if h.scene()]
        takeoff_item["count"].setText(f"Count: {len(valid)}")
        self.update_totals()

    def update_totals(self):
        total_count = 0
        total_hours = 0.0
        for it in self.takeoff_items:
            valid = [h for h in it["highlights"] if h.scene()]
            cnt = len(valid)
            it["count"].setText(f"Count: {cnt}")
            total_count += cnt
            try:
                lab_widget = it.get("labor_field", it.get("labor"))
                lab = float(lab_widget.text()) if lab_widget else 0.0
            except Exception:
                lab = 0.0
            total_hours += cnt * lab
        self.totals_label.setText(f"Totals: Count={total_count}; Hours={total_hours:.2f}")
        self.totalsUpdated.emit()

    def clearTakeoffs(self):
        while self.takeoff_items:
            self.delete_takeoff(self.takeoff_items[0])

    def get_wire_totals(self) -> defaultdict[tuple, float]:
        totals: defaultdict[tuple, float] = defaultdict(float)
        if not self.include_wire:
            return totals

        for it in self.takeoff_items:
            try:
                base_len = float(it["wire_length"].text())
            except Exception:
                continue
            if base_len <= 0:
                continue
            count = len([h for h in it["highlights"] if h.scene()])
            if count == 0:
                continue
            key = (
                it["wire_type"].currentText(),
                it["wire_cable"].currentText(),
                it["wire_mat"].currentText(),
            )
            totals[key] += base_len * count
        return totals
