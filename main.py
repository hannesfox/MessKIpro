# ==============================================================================
#      1. IMPORTS
# ==============================================================================
import sys
import os
import json  # Für das Laden der Toleranzdaten
import math  # Für die Berechnung der Seitenanzahl
import ezdxf
from ezdxf.recover import readfile as recover_readfile
from ezdxf.addons.drawing.pyqt import PyQtBackend
from ezdxf.math import Vec3

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QGraphicsScene, QGraphicsView, QMessageBox,
    QFileDialog, QWidget, QSplitter, QVBoxLayout, QGridLayout, QLabel,
    QLineEdit, QComboBox, QFrame, QHBoxLayout, QDateEdit, QStyleFactory,
    QPushButton
)
from PySide6.QtGui import (
    QColor, QWheelEvent, QAction, QFont, QDragEnterEvent, QDropEvent, QPixmap,
    QFontDatabase
)
from PySide6.QtCore import Qt, QPoint, QPointF, Signal, QDate

# ==============================================================================
#      2. KONSTANTEN, THEME UND HILFSKLASSEN
# ==============================================================================

ZOOM_FACTOR = 1.2

LIGHT_THEME_QSS = """
/* Allgemeines Widget-Styling */
QWidget {{
    /* Der Platzhalter {font_family} wird dynamisch ersetzt */
    font-family: '{font_family}', Arial, sans-serif;
    font-size: 10pt;
    color: #333333;
    background-color: #f5f5f5; /* Leicht grauer Hintergrund */
}}

/* Hauptfenster und Splitter */
QMainWindow {{
    background-color: #e9e9e9;
}}

QSplitter::handle {{
    background-color: #cccccc;
}}

QSplitter::handle:horizontal {{
    width: 2px;
}}

QSplitter::handle:vertical {{
    height: 2px;
}}

/* Eingabefelder: QLineEdit, QComboBox, QDateEdit */
QLineEdit, QComboBox, QDateEdit {{
    background-color: #ffffff;
    border: 1px solid #cccccc;
    border-radius: 4px;
    padding: 6px;
    selection-background-color: #0078d7;
    selection-color: #ffffff;
}}

QLineEdit:focus, QComboBox:focus, QDateEdit:focus {{
    border: 1px solid #0078d7; /* Blauer Akzent bei Fokus */
}}

QComboBox::drop-down {{
    subcontrol-origin: padding;
    subcontrol-position: top right;
    width: 20px;
    border-left-width: 1px;
    border-left-color: #cccccc;
    border-left-style: solid;
    border-top-right-radius: 3px;
    border-bottom-right-radius: 3px;
}}

QComboBox::down-arrow {{
    image: url(down_arrow.png); /* Fallback, falls kein Icon gefunden wird */
}}

/* Buttons */
QPushButton {{
    background-color: #e1e1e1;
    border: 1px solid #cccccc;
    border-radius: 4px;
    padding: 8px 16px;
    font-weight: bold;
}}

QPushButton:hover {{
    background-color: #d1d1d1;
    border-color: #bbbbbb;
}}

QPushButton:pressed {{
    background-color: #c1c1c1;
}}

QPushButton:disabled {{
    background-color: #eeeeee;
    color: #aaaaaa;
    border-color: #dddddd;
}}

/* Labels */
QLabel {{
    background-color: transparent; /* Labels sollen den Hintergrund des Parents haben */
}}

/* Rahmen (Frames) für die Maß-Blöcke */
QFrame {{
    border: 1px solid #dddddd;
    border-radius: 5px;
    background-color: #ffffff;
}}

/* Menü-Bar */
QMenuBar {{
    background-color: #f0f0f0;
}}
QMenuBar::item {{
    padding: 4px 8px;
    background: transparent;
}}
QMenuBar::item:selected {{
    background-color: #d6d6d6;
}}
QMenu {{
    background-color: #fdfdfd;
    border: 1px solid #cccccc;
}}
QMenu::item:selected {{
    background-color: #0078d7;
    color: #ffffff;
}}
"""


class ClickableLineEdit(QLineEdit):
    clicked = Signal()

    def mousePressEvent(self, event):
        self.clicked.emit()
        super().mousePressEvent(event)


class IsoFitsCalculator:
    """
    Kapselt die Logik zur Berechnung von ISO 286-1 Passungen.
    """

    def __init__(self, data_folder_path):
        self.tolerances_data = []
        self.available_fits = [""]
        self._load_data(data_folder_path)

    def _load_data(self, data_folder_path):
        tolerances_path = os.path.join(data_folder_path, "tolerances.json")
        try:
            with open(tolerances_path, 'r', encoding='utf-8') as f:
                self.tolerances_data = json.load(f)
            print(f"INFO: {len(self.tolerances_data)} Toleranzdatensätze geladen.")

            all_fits = set(entry["toleranzklasse"] for entry in self.tolerances_data)
            self.available_fits.extend(sorted(list(all_fits)))
            print(f"INFO: {len(all_fits)} einzigartige Toleranzklassen geladen.")
        except (FileNotFoundError, json.JSONDecodeError) as e:
            QMessageBox.critical(None, "Fataler Fehler", f"Datei 'tolerances.json' nicht gefunden/lesbar: {e}")
            sys.exit(1)

    def calculate(self, nominal_size, fit_string):
        if not self.tolerances_data: return None
        try:
            for entry in self.tolerances_data:
                if (entry["toleranzklasse"].lower() == fit_string.lower() and
                        entry["lowerlimit"] < nominal_size <= entry["upperlimit"]):
                    return entry["es"] / 1000.0, entry["ei"] / 1000.0
            return None
        except Exception as e:
            print(f"Fehler bei Toleranzberechnung: {e}")
            return None


# ==============================================================================
#      3. HAUPT-WIDGETS
# ==============================================================================

# ------------------------------------------------------------------------------
#      3.1 DXFWidget: Linke Seite (DXF-Anzeige)
# ------------------------------------------------------------------------------
class DXFWidget(QWidget):
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
        factor = ZOOM_FACTOR if event.angleDelta().y() > 0 else 1 / ZOOM_FACTOR
        self.view.scale(factor, factor)
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
            backend = PyQtBackend(self.scene)
            from ezdxf.addons.drawing.frontend import Frontend
            from ezdxf.addons.drawing.properties import RenderContext
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
            scene_pos = self.view.mapToScene(event.position().toPoint())
            world_pos = Vec3(scene_pos.x(), scene_pos.y())
            p1 = self.view.mapToScene(event.position().toPoint())
            p2_pos = event.position() + QPointF(self.CLICK_RADIUS_PIXELS, 0)
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
                    QMessageBox.warning(self, "Fehler", f"Messwert konnte nicht ausgelesen werden: {e}")


# ------------------------------------------------------------------------------
#      3.2 MessprotokollWidget: Rechte Seite (Eingabeformular)
# ------------------------------------------------------------------------------
class MessprotokollWidget(QWidget):
    field_selected = Signal(object)
    field_manually_edited = Signal(object)
    TOTAL_MEASURES = 18
    MEASURES_PER_PAGE = 6
    BLOCKS_PER_PAGE = 2
    MEASURES_PER_BLOCK = 3
    TOTAL_BLOCKS = TOTAL_MEASURES // MEASURES_PER_BLOCK
    _pos_vals = [f"+{i / 1000.0:.3f}" for i in range(5, 201, 5)]
    _neg_vals = [f"-{i / 1000.0:.3f}" for i in range(5, 201, 5)]
    TOLERANCE_VALUES = ["", "0"] + _neg_vals[::-1] + _pos_vals
    MESSMITTEL_OPTIONS = ["", "Aussen Mikrometer", "Digimar", "Endmaß", "Gewinde-lehrdorn", "Gewinde-lehrring",
                          "Haarlineal", "Innen Mikrometer", "Innenschnell-taster", "Lehrdorn", "Lehrring",
                          "MahrSurf M 310", "Maschinen- taster", "Mess-schieber", "Messuhr", "optisch", "Prüfstifte",
                          "Radius Lehre", "Rugotest", "Steigungs-lehre", "Subito", "Tiefenmaß", "Winkel-messer",
                          "Zeiss", "Zoller"]
    KUNDEN_LISTE = ["", "AGILOX Services GmbH", "Alpina Tec", "Alpine Metal Tech", "AMB", "Cloeren", "Collin", "Dtech",
                    "Econ", "Eicon", "Eiermacher", "Fill", "Gewa", "Gföllner", "Global Hydro Energy", "Gottfried",
                    "GreinerBio-One", "Gtech", "Haidlmair GmbH", "Hainzl", "HFP", "IFW", "IKIPM", "Kässbohrer",
                    "KI Automation", "Kiefel", "Knorr Bremse", "Kwapil & Co", "Laska", "Mark", "MBK Rinnerberger",
                    "MIBA Sinter", "Myonic", "Peak Technoligy", "Plastic Omnium", "Puhl", "RO-RA", "Rotax", "Schell",
                    "Schröckenfux", "Seisenbacher", "Sema", "SK Blechtechnik", "SMW", "STIWA", "Wuppermann"]

    def __init__(self, parent=None):
        super().__init__(parent)
        main_layout = QVBoxLayout(self);
        main_layout.setSpacing(15)
        script_dir = os.path.dirname(os.path.realpath(__file__))
        data_dir = os.path.join(script_dir, "Data")
        self.iso_calculator = IsoFitsCalculator(data_dir)
        self.nominal_fields, self.upper_tol_combos, self.lower_tol_combos, self.soll_labels, self.iso_fit_combos = [], [], [], [], []
        self.measure_blocks = []
        self.current_page = 0
        self.total_pages = math.ceil(self.TOTAL_BLOCKS / self.BLOCKS_PER_PAGE)
        header_grid = QGridLayout()
        header_grid.setColumnStretch(1, 2);
        header_grid.setColumnStretch(10, 1)

        # === ÄNDERUNG: Titel mit spezifischem Stylesheet vergrößern ===
        title_label = QLabel("Messprotokoll-Assistent")
        # Dieses spezifische Stylesheet überschreibt die globale QSS-Regel.
        title_label.setStyleSheet("""
            font-size: 32pt;
            font-weight: bold;
            color: #2c3e50;
            background-color: transparent;
        """)
        header_grid.addWidget(title_label, 0, 0, 1, 5)

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
        logo_label = QLabel();
        logo_label.setFixedSize(150, 150);
        logo_label.setScaledContents(True)
        logo_path = "app-logo.png"
        if os.path.exists(logo_path):
            logo_label.setPixmap(QPixmap(logo_path))
        header_grid.addWidget(logo_label, 0, 11, 2, 1, alignment=Qt.AlignTop | Qt.AlignRight)
        main_layout.addLayout(header_grid)

        for block_idx in range(self.TOTAL_BLOCKS):
            block_frame = QFrame();
            grid = QGridLayout(block_frame);
            grid.setSpacing(10)
            grid.addWidget(QLabel("Maß lt.\nZeichnung"), 0, 0, 7, 1)
            soll_qlabel = QLabel("SOLL ➡")
            soll_qlabel.setStyleSheet("font-weight: bold;")
            grid.addWidget(soll_qlabel, 6, 0, 1, 1)
            for col_idx in range(self.MEASURES_PER_BLOCK):
                measure_index = block_idx * self.MEASURES_PER_BLOCK + col_idx
                col_start = 1 + col_idx
                grid.addWidget(QLabel(f"Maß {measure_index + 1}", alignment=Qt.AlignCenter), 0, col_start)
                nominal_field = ClickableLineEdit(alignment=Qt.AlignCenter)
                nominal_field.clicked.connect(lambda f=nominal_field: self.field_selected.emit(f))
                nominal_field.textEdited.connect(lambda text, f=nominal_field: self.field_manually_edited.emit(f))
                grid.addWidget(nominal_field, 1, col_start)
                self.nominal_fields.append(nominal_field)
                grid.addWidget(QLabel("ISO-Toleranz", alignment=Qt.AlignCenter), 2, col_start)
                iso_fit_combo = QComboBox();
                iso_fit_combo.setEditable(True);
                iso_fit_combo.addItems(self.iso_calculator.available_fits)
                grid.addWidget(iso_fit_combo, 3, col_start)
                self.iso_fit_combos.append(iso_fit_combo)
                grid.addWidget(QLabel("Messmittel", alignment=Qt.AlignCenter), 4, col_start)
                messmittel_combo = QComboBox();
                messmittel_combo.addItems(self.MESSMITTEL_OPTIONS)
                grid.addWidget(messmittel_combo, 5, col_start)
                tol_layout = QGridLayout()
                upper_tol_combo = QComboBox();
                upper_tol_combo.addItems(self.TOLERANCE_VALUES);
                upper_tol_combo.setEditable(True)
                lower_tol_combo = QComboBox();
                lower_tol_combo.addItems(self.TOLERANCE_VALUES);
                lower_tol_combo.setEditable(True)
                tol_layout.addWidget(upper_tol_combo, 0, 0);
                tol_layout.addWidget(QLabel("Größtmaß"), 0, 1)
                tol_layout.addWidget(lower_tol_combo, 1, 0);
                tol_layout.addWidget(QLabel("Kleinstmaß"), 1, 1)
                self.upper_tol_combos.append(upper_tol_combo);
                self.lower_tol_combos.append(lower_tol_combo)
                soll_label = QLabel("---", alignment=Qt.AlignCenter,
                                    styleSheet="font-weight: bold; border: 1px solid #ccc; padding: 6px; background-color: #f0f0f0;")
                grid.addLayout(tol_layout, 6, col_start)
                grid.addWidget(soll_label, 7, col_start)
                self.soll_labels.append(soll_label)
                nominal_field.textChanged.connect(lambda _, idx=measure_index: self._trigger_iso_fit_calculation(idx))
                iso_fit_combo.currentTextChanged.connect(
                    lambda _, idx=measure_index: self._trigger_iso_fit_calculation(idx))
                upper_tol_combo.currentTextChanged.connect(lambda _, idx=measure_index: self._update_soll_wert(idx))
                lower_tol_combo.currentTextChanged.connect(lambda _, idx=measure_index: self._update_soll_wert(idx))
            main_layout.addWidget(block_frame)
            self.measure_blocks.append(block_frame)

        main_layout.addStretch()
        pagination_layout = QHBoxLayout()
        self.prev_button = QPushButton("<< Zurück")
        self.prev_button.clicked.connect(self._previous_page)
        self.page_label = QLabel("", alignment=Qt.AlignCenter)
        self.next_button = QPushButton("Vor >>")
        self.next_button.clicked.connect(self._next_page)
        pagination_layout.addStretch()
        pagination_layout.addWidget(self.prev_button)
        pagination_layout.addWidget(self.page_label)
        pagination_layout.addWidget(self.next_button)
        pagination_layout.addStretch()
        main_layout.addLayout(pagination_layout)
        self._update_page_view()

    def _update_page_view(self):
        start_block_idx = self.current_page * self.BLOCKS_PER_PAGE
        end_block_idx = start_block_idx + self.BLOCKS_PER_PAGE
        for i, block in enumerate(self.measure_blocks):
            block.setVisible(start_block_idx <= i < end_block_idx)
        self.page_label.setText(f"Seite {self.current_page + 1} / {self.total_pages}")
        self.prev_button.setEnabled(self.current_page > 0)
        self.next_button.setEnabled(self.current_page < self.total_pages - 1)

    def _previous_page(self):
        if self.current_page > 0:
            self.current_page -= 1
            self._update_page_view()

    def _next_page(self):
        if self.current_page < self.total_pages - 1:
            self.current_page += 1
            self._update_page_view()

    def _trigger_iso_fit_calculation(self, index):
        nominal_text = self.nominal_fields[index].text().replace(',', '.')
        fit_string = self.iso_fit_combos[index].currentText().strip()
        self._update_soll_wert(index)
        if not nominal_text or not fit_string: return
        try:
            nominal_value = float(nominal_text)
            result = self.iso_calculator.calculate(nominal_value, fit_string)
            if result is None: return
            upper_dev, lower_dev = result
            self.upper_tol_combos[index].blockSignals(True)
            self.lower_tol_combos[index].blockSignals(True)
            self.upper_tol_combos[index].setCurrentText(f"{upper_dev:+.3f}")
            self.lower_tol_combos[index].setCurrentText(f"{lower_dev:+.3f}")
            self.upper_tol_combos[index].blockSignals(False)
            self.lower_tol_combos[index].blockSignals(False)
            self._update_soll_wert(index)
        except (ValueError, TypeError) as e:
            print(f"Fehler bei ISO-Toleranz-Verarbeitung für Index {index}: {e}")

    def _update_soll_wert(self, index):
        try:
            nominal = float(self.nominal_fields[index].text().replace(',', '.') or 0)
            upper_tol = float(self.upper_tol_combos[index].currentText().replace(',', '.') or 0)
            lower_tol = float(self.lower_tol_combos[index].currentText().replace(',', '.') or 0)
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
        self.target_widget = None
        self._set_application_style()
        self.setAcceptDrops(True)
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
        self.setWindowTitle("Messprotokoll-Assistent");
        self.setGeometry(50, 50, 1800, 1000)
        self.protokoll_widget.field_selected.connect(self.on_protokoll_field_selected)
        self.dxf_widget.dimension_clicked.connect(self.on_dimension_value_received)
        self.protokoll_widget.field_manually_edited.connect(self.on_field_manually_edited)

    def _set_application_style(self):
        """Wendet den globalen Stil und das Theme an."""
        app = QApplication.instance()
        app.setStyle(QStyleFactory.create("Fusion"))
        default_font = QFontDatabase.systemFont(QFontDatabase.GeneralFont)
        font_family = default_font.family()
        print(f"INFO: Verwende System-Schriftart: '{font_family}'")
        formatted_qss = LIGHT_THEME_QSS.format(font_family=font_family)
        app.setStyleSheet(formatted_qss)

    def on_protokoll_field_selected(self, widget):
        if self.target_widget:
            self.target_widget.setStyleSheet("")
        self.target_widget = widget
        self.target_widget.setStyleSheet("background-color: #0078d7; color: white; border: 1px solid #005a9e;")
        print("Ziel für Nennmaß gesetzt.")

    def on_dimension_value_received(self, value):
        if self.target_widget:
            self.target_widget.setText(value.replace('.', ','))
            self.target_widget.setStyleSheet("")
            self.target_widget = None
        else:
            QMessageBox.information(self, "Hinweis", "Bitte klicken Sie zuerst in ein 'Maß lt. Zeichnung'-Feld.")

    def on_field_manually_edited(self, widget):
        if self.target_widget == widget:
            print("Manuelle Eingabe erkannt. DXF-Ziel wird deaktiviert.")
            self.target_widget.setStyleSheet("")
            self.target_widget = None

    def open_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "DXF-Datei öffnen", "", "DXF-Dateien (*.dxf)")
        if path: self.dxf_widget.load_dxf(path)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            url = event.mimeData().urls()[0]
            if url.isLocalFile() and url.toLocalFile().lower().endswith('.dxf'):
                event.acceptProposedAction()
            else:
                event.ignore()
        else:
            event.ignore()

    def dropEvent(self, event: QDropEvent):
        if event.mimeData().hasUrls():
            file_path = event.mimeData().urls()[0].toLocalFile()
            if file_path.lower().endswith('.dxf'):
                print(f"INFO: Lade DXF-Datei per Drag & Drop: {file_path}")
                self.dxf_widget.load_dxf(file_path)
                event.acceptProposedAction()
            else:
                event.ignore()
        else:
            event.ignore()


# ==============================================================================
#      5. START DER APPLIKATION
# ==============================================================================
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())