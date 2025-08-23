# ==============================================================================
#      1. IMPORTS
# ==============================================================================
import sys
import os  # Wichtig für die Pfad-Prüfung des Logos
import ezdxf
from ezdxf.recover import readfile as recover_readfile
from ezdxf.addons.drawing.pyqt import PyQtBackend
from ezdxf.math import Vec3

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QGraphicsScene, QGraphicsView, QMessageBox,
    QFileDialog, QWidget, QSplitter, QVBoxLayout, QGridLayout, QLabel,
    QLineEdit, QComboBox, QFrame, QHBoxLayout, QDateEdit, QStyleFactory
)
from PySide6.QtGui import QColor, QWheelEvent, QAction, QFont, QDragEnterEvent, QDropEvent, QPixmap
from PySide6.QtCore import Qt, QPoint, QPointF, Signal, QDate

# ==============================================================================
#      2. KONSTANTEN UND HILFSKLASSEN
# ==============================================================================

ZOOM_FACTOR = 1.2


class ClickableLineEdit(QLineEdit):
    clicked = Signal()

    def mousePressEvent(self, event):
        self.clicked.emit()
        super().mousePressEvent(event)


# ==============================================================================
#      3. HAUPT-WIDGETS
# ==============================================================================

# ------------------------------------------------------------------------------
#      3.1 DXFWidget: Linke Seite (DXF-Anzeige)
# ------------------------------------------------------------------------------
class DXFWidget(QWidget):
    # (Unverändert)
    dimension_clicked = Signal(str)
    CLICK_RADIUS_PIXELS = 50

    def __init__(self, parent=None):
        super().__init__(parent)
        self.doc, self.msp = None, None
        self.scene = QGraphicsScene()
        self.view = QGraphicsView(self.scene)
        self.view.setAcceptDrops(False)
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


# ------------------------------------------------------------------------------
#      3.2 MessprotokollWidget: Rechte Seite (Eingabeformular)
# ------------------------------------------------------------------------------
class MessprotokollWidget(QWidget):
    field_selected = Signal(object)
    field_manually_edited = Signal(object)

    _pos_vals = [f"+{i / 1000.0:.3f}" for i in range(5, 201, 5)]
    _neg_vals = [f"-{i / 1000.0:.3f}" for i in range(5, 201, 5)]
    UPPER_TOLERANCE_VALUES = ["", "0"] + _pos_vals
    LOWER_TOLERANCE_VALUES = ["", "0"] + _neg_vals
    MESSMITTEL_OPTIONS = ["", "optisch", "Messschieber", "Bügelmessschraube", "Höhenmessgerät", "3D-Messmaschine"]
    KUNDEN_LISTE = ["", "Tool Service GmbH", "Musterfirma AG", "Projekt X Kunde"]

    def __init__(self, parent=None):
        super().__init__(parent)
        main_layout = QVBoxLayout(self);
        main_layout.setSpacing(15)
        self.nominal_fields, self.upper_tol_combos, self.lower_tol_combos, self.soll_labels = [], [], [], []

        header_grid = QGridLayout()
        header_grid.setColumnStretch(1, 2);
        header_grid.setColumnStretch(10, 1)  # Stretch vor dem Logo
        header_grid.addWidget(QLabel("Messprotokoll-Assistent", font=QFont("Arial", 30, QFont.Bold)), 0, 0, 1, 5)
        kunde_combo = QComboBox();
        kunde_combo.setEditable(True);
        kunde_combo.addItems(self.KUNDEN_LISTE)
        header_grid.addWidget(QLabel("Kunde:"), 1, 0);
        header_grid.addWidget(kunde_combo, 1, 1)
        auftrag_layout = QHBoxLayout();
        auftrag_layout.addWidget(QLabel("Auftrag: AT-25 /"));
        auftrag_layout.addWidget(QLineEdit())
        header_grid.addLayout(auftrag_layout, 1, 3)
        pos_edit = QLineEdit();
        pos_edit.setFixedWidth(80)
        header_grid.addWidget(QLabel("Pos.:"), 1, 5);
        header_grid.addWidget(pos_edit, 1, 6)
        date_edit = QDateEdit(calendarPopup=True, date=QDate.currentDate())
        header_grid.addWidget(QLabel("Datum:"), 1, 8);
        header_grid.addWidget(date_edit, 1, 9)

        # === NEU: Logo hinzufügen ===
        # Stellen Sie sicher, dass 'app-logo.png' im selben Verzeichnis wie das Skript liegt.
        logo_label = QLabel()
        logo_label.setFixedSize(200, 200)
        logo_label.setScaledContents(True)
        logo_path = "app-logo.png"
        if os.path.exists(logo_path):
            logo_pixmap = QPixmap(logo_path)
            logo_label.setPixmap(logo_pixmap)
        else:
            print(f"WARNUNG: Logo-Datei '{logo_path}' nicht gefunden.")
        # Logo oben rechts im Grid platzieren
        header_grid.addWidget(logo_label, 0, 11, 2, 1, alignment=Qt.AlignTop | Qt.AlignRight)

        main_layout.addLayout(header_grid)

        for block_idx in range(2):
            block_frame = QFrame();
            block_frame.setFrameShape(QFrame.StyledPanel)
            grid = QGridLayout(block_frame);
            grid.setSpacing(10)
            grid.addWidget(QLabel("Maß lt.\nZeichnung"), 0, 0, 5, 1)
            grid.addWidget(QLabel("SOLL ➡", font=QFont("Arial", 10, QFont.Bold)), 4, 0, 1, 1)
            for col_idx in range(4):
                measure_index = block_idx * 4 + col_idx
                col_start = 1 + col_idx
                grid.addWidget(QLabel(f"Maß {measure_index + 1}", alignment=Qt.AlignCenter), 0, col_start)
                nominal_field = ClickableLineEdit(alignment=Qt.AlignCenter,
                                                  styleSheet="background-color: #e9f5e9; color: #333333;")
                nominal_field.clicked.connect(lambda f=nominal_field: self.field_selected.emit(f))
                nominal_field.textEdited.connect(lambda text, f=nominal_field: self.field_manually_edited.emit(f))
                grid.addWidget(nominal_field, 1, col_start)
                self.nominal_fields.append(nominal_field)
                grid.addWidget(QLabel("Messmittel", alignment=Qt.AlignCenter), 2, col_start)
                messmittel_combo = QComboBox();
                messmittel_combo.addItems(self.MESSMITTEL_OPTIONS)
                grid.addWidget(messmittel_combo, 3, col_start)
                tol_layout = QGridLayout()
                upper_tol_combo = QComboBox();
                upper_tol_combo.addItems(self.UPPER_TOLERANCE_VALUES);
                upper_tol_combo.setEditable(True)
                lower_tol_combo = QComboBox();
                lower_tol_combo.addItems(self.LOWER_TOLERANCE_VALUES);
                lower_tol_combo.setEditable(True)
                tol_layout.addWidget(upper_tol_combo, 0, 0);
                tol_layout.addWidget(QLabel("Größtmaß"), 0, 1)
                tol_layout.addWidget(lower_tol_combo, 1, 0);
                tol_layout.addWidget(QLabel("Kleinstmaß"), 1, 1)
                self.upper_tol_combos.append(upper_tol_combo);
                self.lower_tol_combos.append(lower_tol_combo)
                soll_label = QLabel("---", alignment=Qt.AlignCenter,
                                    styleSheet="font-weight: bold; border: 1px solid #ccc; padding: 6px;")
                grid.addLayout(tol_layout, 4, col_start)
                grid.addWidget(soll_label, 5, col_start)
                self.soll_labels.append(soll_label)
                nominal_field.textChanged.connect(lambda _, idx=measure_index: self._update_soll_wert(idx))
                upper_tol_combo.currentTextChanged.connect(lambda _, idx=measure_index: self._update_soll_wert(idx))
                lower_tol_combo.currentTextChanged.connect(lambda _, idx=measure_index: self._update_soll_wert(idx))
            main_layout.addWidget(block_frame)
        main_layout.addStretch()

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


# ==============================================================================
#      4. HAUPTFENSTER (MAIN WINDOW)
# ==============================================================================
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self._set_application_style()
        self.setAcceptDrops(True)
        self.target_widget = None
        self.dxf_widget = DXFWidget()
        self.protokoll_widget = MessprotokollWidget()
        splitter = QSplitter(Qt.Horizontal);
        splitter.addWidget(self.dxf_widget);
        splitter.addWidget(self.protokoll_widget)
        splitter.setSizes([1000, 1000])
        self.setCentralWidget(splitter)
        menu_bar = self.menuBar();
        file_menu = menu_bar.addMenu("Datei")
        open_action = QAction("DXF Öffnen...", self);
        open_action.triggered.connect(self.open_file);
        file_menu.addAction(open_action)
        self.setWindowTitle("Messprotokoll-Assistent (Fusion Style)");
        self.setGeometry(50, 50, 1800, 1000)
        self.protokoll_widget.field_selected.connect(self.on_protokoll_field_selected)
        self.dxf_widget.dimension_clicked.connect(self.on_dimension_value_received)
        self.protokoll_widget.field_manually_edited.connect(self.on_field_manually_edited)

    def _set_application_style(self):
        # (Unverändert)
        app = QApplication.instance()
        macos_style_found = False
        if sys.platform == "darwin":
            for style_key in QStyleFactory.keys():
                if style_key in ["Macintosh", "macOS"]:
                    app.setStyle(QStyleFactory.create(style_key))
                    macos_style_found = True
                    print(f"INFO: Nativer '{style_key}'-Stil angewendet.")
                    break
        if not macos_style_found:
            try:
                app.setStyle(QStyleFactory.create("Fusion"))
                print("INFO: Nativer macOS-Stil nicht gefunden oder nicht macOS. Fallback auf Fusion-Stil angewendet.")
            except Exception:
                print("WARNUNG: Fusion-Stil nicht verfügbar. Verwende Systemstandard-Stil.")
        app.setPalette(app.style().standardPalette())
        app.setStyleSheet("")

    def on_protokoll_field_selected(self, widget):
        if self.target_widget: self.target_widget.setStyleSheet("background-color: #e9f5e9; color: #333333;")
        self.target_widget = widget
        self.target_widget.setStyleSheet("background-color: #0078d7; color: white;")
        print("Ziel für Nennmaß gesetzt.")

    def on_dimension_value_received(self, value):
        if self.target_widget:
            self.target_widget.setText(value.replace('.', ','))
            self.target_widget.setStyleSheet("background-color: #e9f5e9; color: #333333;")
            self.target_widget = None
        else:
            QMessageBox.information(self, "Hinweis",
                                    "Bitte klicken Sie zuerst in ein 'Maß lt. Zeichnung'-Feld im Protokoll.")

    def on_field_manually_edited(self, widget):
        if self.target_widget == widget:
            print("Manuelle Eingabe erkannt. DXF-Ziel wird deaktiviert.")
            self.target_widget.setStyleSheet("background-color: #e9f5e9; color: #333333;")
            self.target_widget = None

    def open_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "DXF-Datei öffnen", "", "DXF-Dateien (*.dxf)")
        if path: self.dxf_widget.load_dxf(path)

    # --------------------------------------------------------------------------
    #      4.1 Drag-and-Drop-Funktionalität
    # --------------------------------------------------------------------------
    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                if url.isLocalFile() and url.toLocalFile().lower().endswith('.dxf'):
                    event.acceptProposedAction()
                    return
        event.ignore()

    def dropEvent(self, event: QDropEvent):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                file_path = url.toLocalFile()
                if file_path.lower().endswith('.dxf'):
                    print(f"INFO: Lade DXF-Datei per Drag & Drop: {file_path}")
                    self.dxf_widget.load_dxf(file_path)
                    event.acceptProposedAction()
                    return
        event.ignore()


# ==============================================================================
#      5. START DER APPLIKATION
# ==============================================================================
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())