# -*- coding: utf-8 -*-
"""PDF viewing widgets and graphics items."""

from __future__ import annotations

from PyQt5 import QtCore, QtGui, QtWidgets
import fitz  # PyMuPDF


class HighlightItem(QtWidgets.QGraphicsObject):
    """Movable/deletable rectangle used for highlights."""

    deleted = QtCore.pyqtSignal(object)

    def __init__(self, rect: QtCore.QRectF, color: QtGui.QColor, parent=None):
        super().__init__(parent)
        self._rect = rect
        self._color = color
        self.setFlags(QtWidgets.QGraphicsItem.ItemIsSelectable)

    # QGraphicsObject interface --------------------------------------------
    def boundingRect(self) -> QtCore.QRectF:
        return self._rect

    def paint(self, painter: QtGui.QPainter, _option, _widget=None):
        painter.setBrush(QtGui.QBrush(self._color))
        painter.setPen(QtCore.Qt.NoPen)
        painter.drawRect(self._rect)

    # helpers ---------------------------------------------------------------
    def setRect(self, rect: QtCore.QRectF):
        self.prepareGeometryChange()
        self._rect = rect
        self.update()

    def rect(self) -> QtCore.QRectF:
        """Return the item's rectangle."""
        return self._rect

    # context menu ----------------------------------------------------------
    def contextMenuEvent(self, event: QtWidgets.QGraphicsSceneContextMenuEvent):
        menu = QtWidgets.QMenu()
        move_action = menu.addAction("Move")
        delete_action = menu.addAction("Delete")
        action = menu.exec_(event.screenPos())
        view: PDFGraphicsView = self.scene().views()[0]  # type: ignore
        if action == move_action:
            view.startMovingItem(self)
        elif action == delete_action:
            view.highlightDeleted.emit(self)
            self.scene().removeItem(self)
        event.accept()


class LineItem(QtWidgets.QGraphicsLineItem):
    """Simple line item used for doodles."""

    def __init__(self, line: QtCore.QLineF, color: QtGui.QColor, parent=None):
        super().__init__(line, parent)
        pen = QtGui.QPen(color, 2)
        self.setPen(pen)
        self.setFlags(
            QtWidgets.QGraphicsItem.ItemIsSelectable
            | QtWidgets.QGraphicsItem.ItemIsMovable
        )
        self.page = -1


class PDFGraphicsView(QtWidgets.QGraphicsView):
    """Graphics view that displays PDF pages and supports drawing."""

    stampDropped = QtCore.pyqtSignal(object)
    highlightDeleted = QtCore.pyqtSignal(object)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setDragMode(QtWidgets.QGraphicsView.ScrollHandDrag)
        self.setMouseTracking(True)
        self.setAcceptDrops(True)
        self._scene = QtWidgets.QGraphicsScene(self)
        self.setScene(self._scene)

        # drawing helpers
        self.draw_mode = False
        self.draw_shape = "rect"  # "rect" or "line"
        self.drawing = False
        self.template_defined = False
        self.template_item: QtWidgets.QGraphicsRectItem | None = None
        self.start_point: QtCore.QPointF | None = None
        self.current_highlight_color = QtGui.QColor(255, 0, 0, 100)

        # move helpers
        self.moving_item: HighlightItem | None = None
        self.current_takeoff: dict | None = None

        # pdf
        self.doc: fitz.Document | None = None
        self.current_page: int = -1

    # ── drawing‑mode toggle ────────────────────────────────────────────
    def setDrawingMode(self, enabled: bool):
        self.draw_mode = enabled
        if enabled:
            self.setCursor(QtCore.Qt.CrossCursor)
            self.setDragMode(QtWidgets.QGraphicsView.NoDrag)
        else:
            self.setCursor(QtCore.Qt.ArrowCursor)
            self.setDragMode(QtWidgets.QGraphicsView.ScrollHandDrag)
            if self.template_item:
                self._scene.removeItem(self.template_item)
                self.template_item = None
            self.drawing = False
            self.template_defined = False

    def setDrawingShape(self, shape: str):
        """Choose the drawing shape (``rect`` or ``line``)."""
        self.draw_shape = shape

    # ------------------------------------------------------------------
    def cloneTemplate(self, item: QtWidgets.QGraphicsRectItem) -> HighlightItem:
        new_item = HighlightItem(item.rect(), self.current_highlight_color)
        new_item.deleted.connect(self.handleHighlightDeleted)
        new_item.page = self.current_page  # type: ignore[attr-defined]
        return new_item

    # ------------------------------------------------------------------
    def display_page(self, page_num: int):
        if not self.doc or page_num < 0 or page_num >= self.doc.page_count:
            return
        if hasattr(self, "_pixmap_item") and self._pixmap_item:
            self._scene.removeItem(self._pixmap_item)  # type: ignore[attr-defined]
        page = self.doc.load_page(page_num)
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
        img = QtGui.QImage(
            pix.samples, pix.width, pix.height, pix.stride, QtGui.QImage.Format_RGB888
        )
        pixmap = QtGui.QPixmap.fromImage(img)
        self._pixmap_item = self._scene.addPixmap(pixmap)  # type: ignore[attr-defined]
        self._pixmap_item.setZValue(-10)  # type: ignore[attr-defined]
        self.setSceneRect(QtCore.QRectF(pixmap.rect()))
        self.current_page = page_num

    # ------------------------------------------------------------------
    def load_pdf(self, pdf_path: str):
        try:
            self.doc = fitz.open(pdf_path)
        except Exception as e:
            QtWidgets.QMessageBox.warning(self, "Error", f"Failed to open PDF:\n{e}")
            return
        self._scene.clear()
        if self.doc.page_count:
            self.display_page(0)

    # ------------------------------------------------------------------
    # zoom with wheel ---------------------------------------------------
    def wheelEvent(self, event: QtGui.QWheelEvent):
        z_in, z_out = 1.25, 1 / 1.25
        self.scale(z_in if event.angleDelta().y() > 0 else z_out, z_in if event.angleDelta().y() > 0 else z_out)
        event.accept()

    # ------------------------------------------------------------------
    # fix ghost‑rectangle during keyboard pan --------------------------
    def scrollContentsBy(self, dx: int, dy: int):
        super().scrollContentsBy(dx, dy)
        if self.template_item:
            self.viewport().update()

    # ------------------------------------------------------------------
    def mousePressEvent(self, event: QtGui.QMouseEvent):
        if self.moving_item:
            if event.button() == QtCore.Qt.LeftButton:
                if self.current_takeoff is not None:
                    self.current_takeoff["highlights"].append(self.moving_item)
                    self.stampDropped.emit(self.moving_item)
                self.moving_item = None
                return
            elif event.button() == QtCore.Qt.RightButton:
                self._scene.removeItem(self.moving_item)
                self.highlightDeleted.emit(self.moving_item)
                self.moving_item = None
                return

        if self.draw_mode and event.button() == QtCore.Qt.RightButton:
            if isinstance(self.itemAt(event.pos()), HighlightItem):
                super().mousePressEvent(event)
            else:
                self.setDrawingMode(False)
            return

        if self.draw_mode and event.button() == QtCore.Qt.LeftButton:
            if not self.template_defined:
                self.drawing = True
                self.start_point = self.mapToScene(event.pos())
                if self.draw_shape == "rect":
                    self.template_item = QtWidgets.QGraphicsRectItem(
                        QtCore.QRectF(self.start_point, self.start_point)
                    )
                    self.template_item.setBrush(QtGui.QBrush(self.current_highlight_color))
                    self.template_item.setPen(QtGui.QPen(QtCore.Qt.NoPen))
                    self._scene.addItem(self.template_item)
                else:
                    pen = QtGui.QPen(self.current_highlight_color, 2)
                    self.template_item = QtWidgets.QGraphicsLineItem(
                        QtCore.QLineF(self.start_point, self.start_point)
                    )
                    self.template_item.setPen(pen)
                    self._scene.addItem(self.template_item)
                return
            else:
                if self.draw_shape == "rect":
                    stamped = self.cloneTemplate(self.template_item)  # type: ignore
                    stamped.setRect(self.template_item.rect())  # type: ignore
                else:
                    line = self.template_item.line()  # type: ignore
                    stamped = LineItem(line, self.current_highlight_color)
                    stamped.page = self.current_page
                self._scene.addItem(stamped)
                self.stampDropped.emit(stamped)
                return

        super().mousePressEvent(event)

    # ------------------------------------------------------------------
    def mouseMoveEvent(self, event: QtGui.QMouseEvent):
        if self.moving_item:
            return super().mouseMoveEvent(event)

        if self.draw_mode and self.drawing and self.template_item:
            p = self.mapToScene(event.pos())
            if self.draw_shape == "rect":
                self.template_item.setRect(QtCore.QRectF(self.start_point, p).normalized())
            else:
                line = QtCore.QLineF(self.start_point, p)
                self.template_item.setLine(line)  # type: ignore
            return

        if self.draw_mode and self.template_defined and not self.drawing and self.template_item:
            p = self.mapToScene(event.pos())
            if self.draw_shape == "rect":
                size = self.template_item.rect().size()
                self.template_item.setRect(QtCore.QRectF(p, size))
            else:
                line = self.template_item.line()  # type: ignore
                delta = line.p2() - line.p1()
                self.template_item.setLine(QtCore.QLineF(p, p + delta))  # type: ignore
            return

        super().mouseMoveEvent(event)

    # ------------------------------------------------------------------
    def mouseReleaseEvent(self, event: QtGui.QMouseEvent):
        if self.draw_mode and event.button() == QtCore.Qt.LeftButton and self.drawing:
            self.drawing = False
            self.template_defined = True
            if self.draw_shape == "rect":
                stamped = self.cloneTemplate(self.template_item)  # type: ignore
                stamped.setRect(self.template_item.rect())  # type: ignore
            else:
                line = self.template_item.line()  # type: ignore
                stamped = LineItem(line, self.current_highlight_color)
                stamped.page = self.current_page
            self._scene.addItem(stamped)
            self.stampDropped.emit(stamped)
            return
        super().mouseReleaseEvent(event)

    # ------------------------------------------------------------------
    def startMovingItem(self, item):
        self.setDrawingMode(True)
        self.template_defined = True
        self.drawing = False
        if self.template_item and self.template_item is not item and self.template_item.scene():
            self._scene.removeItem(self.template_item)
        self.template_item = item
        if isinstance(item, HighlightItem):
            self.current_highlight_color = item._color
            self.draw_shape = "rect"
        elif isinstance(item, LineItem):
            pen = item.pen()
            self.current_highlight_color = pen.color()
            self.draw_shape = "line"
        self.moving_item = None

    # ------------------------------------------------------------------
    def handleHighlightDeleted(self, item):
        if self.current_takeoff and item in self.current_takeoff["highlights"]:
            self.current_takeoff["highlights"].remove(item)

