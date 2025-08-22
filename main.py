import sys
import ezdxf
from ezdxf.recover import readfile as recover_readfile
from ezdxf.addons.drawing.pyqt import PyQtBackend
from ezdxf.math import Vec3

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QGraphicsScene, QGraphicsView, QMessageBox,
    QFileDialog, QWidget, QSplitter, QVBoxLayout, QGridLayout, QLabel,
    QLineEdit, QComboBox, QFrame
)
from PySide6.QtGui import QColor, QWheelEvent, QAction, QFont
from PySide6.QtCore import Qt, QPoint, QPointF, Signal

# Faktor für das Zoomen mit dem Mausrad
ZOOM_FACTOR = 1.2


class ClickableLineEdit(QLineEdit):
    """Ein QLineEdit, das ein 'clicked' Signal aussendet, wenn es angeklickt wird."""
    clicked = Signal()

    def mousePressEvent(self, event):
        self.clicked.emit()
        super().mousePressEvent(event)


class DXFWidget(QWidget):
    # (Dieser Teil des Codes ist unverändert)
    dimension_clicked = Signal(str)
    CLICK_RADIUS_PIXELS = 50

    def __init__(self, parent=None):
        super().__init__(parent)
        self.doc, self.msp = None, None
        self.scene = QGraphicsScene()
        self.view = QGraphicsView(self.scene)
        layout = QVBoxLayout(self);
        layout.setContentsMargins(0, 0, 0, 0);
        layout.addWidget(self.view)
        self.setLayout(layout)
        self.view.scale(1, -1)
        self.view.setTransformationAnchor(QGraphicsView.ViewportAnchor.AnchorUnderMouse)
        self.view.setResizeAnchor(QGraphicsView.ViewportAnchor.AnchorViewCenter)
        self.view.setDragMode(QGraphicsView.DragMode.NoDrag)
        self.view.mousePressEvent = self.handle_mouse_press
        self.view.wheelEvent = self.handle_wheel_event

    def handle_wheel_event(self, event: QWheelEvent):
        if event.angleDelta().y() > 0:
            self.view.scale(ZOOM_FACTOR, ZOOM_FACTOR)
        else:
            self.view.scale(1 / ZOOM_FACTOR, 1 / ZOOM_FACTOR)
        super(QGraphicsView, self.view).wheelEvent(event)

    def load_dxf(self, filepath):
        try:
            self.doc, auditor = recover_readfile(filepath)
            if auditor.has_errors: print(f"DXF-Fehler: {len(auditor.errors)} Fehler.")
            self.msp = self.doc.modelspace()
            self.draw_dxf()
        except Exception as e:
            QMessageBox.critical(self, "Fehler", f"DXF-Ladefehler: {e}")

    def draw_dxf(self):
        self.scene.clear();
        self.view.setBackgroundBrush(QColor(30, 30, 30))
        try:
            from ezdxf.addons.drawing.frontend import Frontend
            from ezdxf.addons.drawing.properties import RenderContext
            backend = PyQtBackend(self.scene);
            ctx = RenderContext(self.doc)
            frontend = Frontend(ctx, backend);
            frontend.draw_layout(self.msp, finalize=True)
        except Exception as e:
            print(f"Zeichnen fehlgeschlagen: {e}")
        self.view.setSceneRect(self.scene.itemsBoundingRect())
        self.view.fitInView(self.scene.itemsBoundingRect(), Qt.KeepAspectRatio)

    def handle_mouse_press(self, event):
        super(QGraphicsView, self.view).mousePressEvent(event)
        if event.button() == Qt.MouseButton.LeftButton and self.msp:
            view_pos_f = event.position();
            view_pos_int = view_pos_f.toPoint()
            scene_pos = self.view.mapToScene(view_pos_int)
            world_pos = Vec3(scene_pos.x(), scene_pos.y())
            p1 = self.view.mapToScene(view_pos_int)
            p2_pos = view_pos_f + QPointF(self.CLICK_RADIUS_PIXELS, 0)
            p2 = self.view.mapToScene(p2_pos.toPoint())
            search_radius_world = abs(p2.x() - p1.x())
            found_dimension, min_dist = None, float('inf')
            for dimension in self.msp.query('DIMENSION'):
                points_to_check = [p for p in [getattr(dimension.dxf, 'defpoint', None),
                                               getattr(dimension.dxf, 'text_midpoint', None)] if p]
                for point in points_to_check:
                    dist = world_pos.distance(point)
                    if dist < search_radius_world and dist < min_dist:
                        min_dist, found_dimension = dist, dimension
            if found_dimension:
                try:
                    measurement = found_dimension.get_measurement()
                    self.dimension_clicked.emit(f"{measurement:.4f}")
                except Exception as e:
                    QMessageBox.warning(self, "Fehler", f"Konnte Messwert nicht auslesen: {e}")


class MessprotokollWidget(QWidget):
    """Ein Widget, das das Messprotokoll-Formular als feste GUI darstellt."""
    field_selected = Signal(object)
    TOLERANCE_VALUES = ["", "0", "+0.01", "-0.01", "+0.02", "-0.02", "+0.05", "-0.05", "+0.1", "-0.1", "+0.2", "-0.2"]
    MESSMITTEL_OPTIONS = ["", "optisch", "Messschieber", "Bügelmessschraube", "Höhenmessgerät", "3D-Messmaschine"]

    def __init__(self, parent=None):
        super().__init__(parent)
        grid = QGridLayout(self)
        grid.setSpacing(1)

        self.nominal_fields = []
        self.upper_tol_combos = []
        self.lower_tol_combos = []
        self.soll_labels = []
        self.measurement_fields = []

        header_font = QFont();
        header_font.setPointSize(24);
        header_font.setBold(True)
        grid.addWidget(QLabel("Messprotokoll"), 0, 0, 1, 4)
        grid.addWidget(QLabel("Kunde:"), 1, 0);
        grid.addWidget(QLineEdit(), 1, 1, 1, 2)
        grid.addWidget(QLabel("Auftrag:"), 1, 3);
        grid.addWidget(QLineEdit(), 1, 4, 1, 2)
        line = QFrame();
        line.setFrameShape(QFrame.HLine);
        line.setFrameShadow(QFrame.Sunken)
        grid.addWidget(line, 2, 0, 1, 15)

        grid.addWidget(QLabel("Maß lt.\nZeichnung"), 5, 0, 5, 1)
        grid.addWidget(QLabel("SOLL ➡", font=QFont("Arial", 10, QFont.Bold)), 9, 0, 1, 2)

        for i in range(8):
            col_start = 2 + (i * 3)

            lbl_messmittel = QLabel("Messmittel");
            lbl_messmittel.setAlignment(Qt.AlignCenter)
            grid.addWidget(lbl_messmittel, 3, col_start, 1, 3)
            messmittel_combo = QComboBox();
            messmittel_combo.addItems(self.MESSMITTEL_OPTIONS)
            grid.addWidget(messmittel_combo, 4, col_start, 1, 3)

            lbl_mass = QLabel(f"Maß {i + 1}");
            lbl_mass.setAlignment(Qt.AlignCenter)
            grid.addWidget(lbl_mass, 5, col_start, 1, 3)

            nominal_field = ClickableLineEdit();
            nominal_field.setAlignment(Qt.AlignCenter)
            nominal_field.setReadOnly(True)
            nominal_field.setStyleSheet("QLineEdit:read-only { background-color: #e9f5e9; }")
            nominal_field.clicked.connect(lambda f=nominal_field: self.field_selected.emit(f))
            grid.addWidget(nominal_field, 6, col_start, 1, 3)
            self.nominal_fields.append(nominal_field)

            upper_tol_combo = QComboBox();
            upper_tol_combo.addItems(self.TOLERANCE_VALUES)
            lower_tol_combo = QComboBox();
            lower_tol_combo.addItems(self.TOLERANCE_VALUES)
            grid.addWidget(upper_tol_combo, 7, col_start, 1, 2);
            grid.addWidget(QLabel(" Größtmaß"), 7, col_start + 2)
            grid.addWidget(lower_tol_combo, 8, col_start, 1, 2);
            grid.addWidget(QLabel(" Kleinstmaß"), 8, col_start + 2)
            self.upper_tol_combos.append(upper_tol_combo);
            self.lower_tol_combos.append(lower_tol_combo)

            soll_label = QLabel("---");
            soll_label.setAlignment(Qt.AlignCenter)
            soll_label.setStyleSheet("font-weight: bold; border: 1px solid #ccc;")
            grid.addWidget(soll_label, 9, col_start, 1, 3)
            self.soll_labels.append(soll_label)

            nominal_field.textChanged.connect(lambda _, idx=i: self._update_soll_wert(idx))
            upper_tol_combo.currentTextChanged.connect(lambda _, idx=i: self._update_soll_wert(idx))
            lower_tol_combo.currentTextChanged.connect(lambda _, idx=i: self._update_soll_wert(idx))

        num_measurement_rows = 15
        for row in range(num_measurement_rows):
            row_label_text = f"Teil {row + 1}"
            if row == 0: row_label_text = "Anfahrteil"
            grid.addWidget(QLabel(row_label_text), 10 + row, 0, 1, 2)
            for i in range(8):
                col_start = 2 + (i * 3)
                field = ClickableLineEdit()
                field.setAlignment(Qt.AlignCenter);
                field.setReadOnly(True)
                field.setStyleSheet("QLineEdit:read-only { background-color: #f0f0f0; }")
                field.clicked.connect(lambda f=field: self.field_selected.emit(f))
                grid.addWidget(field, 10 + row, col_start, 1, 3)
                self.measurement_fields.append(field)

    def _update_soll_wert(self, index):
        try:
            nominal_text = self.nominal_fields[index].text().replace(',', '.')
            upper_tol_text = self.upper_tol_combos[index].currentText().replace(',', '.')
            lower_tol_text = self.lower_tol_combos[index].currentText().replace(',', '.')
            nominal = float(nominal_text or 0);
            upper_tol = float(upper_tol_text or 0);
            lower_tol = float(lower_tol_text or 0)
            soll_wert = nominal + (upper_tol + lower_tol) / 2.0
            self.soll_labels[index].setText(f"{soll_wert:.4f}".replace('.', ','))
        except (ValueError, TypeError):
            self.soll_labels[index].setText("---")


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.target_widget = None
        self.dxf_widget = DXFWidget()
        self.protokoll_widget = MessprotokollWidget()
        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(self.dxf_widget)
        splitter.addWidget(self.protokoll_widget)

        # === ÄNDERUNG HIER ===
        # Ersetzt setStretchFactor für eine exakte 50/50 Start-Aufteilung.
        splitter.setSizes([1000, 1000])

        self.setCentralWidget(splitter)

        menu_bar = self.menuBar()
        file_menu = menu_bar.addMenu("Datei")
        open_action = QAction("DXF Öffnen...", self)
        open_action.triggered.connect(self.open_file)
        file_menu.addAction(open_action)

        self.setWindowTitle("Messprotokoll-Assistent v5.1")
        self.setGeometry(50, 50, 1800, 1000)

        self.protokoll_widget.field_selected.connect(self.on_protokoll_field_selected)
        self.dxf_widget.dimension_clicked.connect(self.on_dimension_value_received)

    def on_protokoll_field_selected(self, widget):
        if self.target_widget:
            if self.target_widget in self.protokoll_widget.nominal_fields:
                self.target_widget.setStyleSheet("QLineEdit:read-only { background-color: #e9f5e9; }")
            else:
                self.target_widget.setStyleSheet("QLineEdit:read-only { background-color: #f0f0f0; }")

        self.target_widget = widget

        if self.target_widget in self.protokoll_widget.nominal_fields:
            self.target_widget.setStyleSheet("background-color: #90EE90;")
        elif self.target_widget in self.protokoll_widget.measurement_fields:
            self.target_widget.setStyleSheet("background-color: #aadeff;")

    def on_dimension_value_received(self, value):
        if self.target_widget:
            self.target_widget.setText(value.replace('.', ','))

            if self.target_widget in self.protokoll_widget.nominal_fields:
                self.target_widget.setStyleSheet("QLineEdit:read-only { background-color: #e9f5e9; }")
            else:
                self.target_widget.setStyleSheet("QLineEdit:read-only { background-color: #f0f0f0; }")
            self.target_widget = None
        else:
            QMessageBox.information(self, "Hinweis",
                                    "Bitte klicken Sie zuerst in ein Feld im Protokoll (Nennmaß oder Messwert).")

    def open_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "DXF-Datei öffnen", "", "DXF-Dateien (*.dxf)")
        if path:
            self.dxf_widget.load_dxf(path)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())