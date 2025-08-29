# ==============================================================================
#      1. IMPORTS
# ==============================================================================
import sys
import os
import json
import math
import ezdxf
import openpyxl
from openpyxl.cell.cell import MergedCell

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
    QColor, QWheelEvent, QAction, QFont, QDragEnterEvent, QDropEvent, QPixmap
)
from PySide6.QtCore import Qt, QPoint, QPointF, Signal, QDate

# qt-material wird NACH PySide6 importiert
from qt_material import apply_stylesheet

# ==============================================================================
#      2. KONSTANTEN UND HILFSKLASSEN
# ==============================================================================

ZOOM_FACTOR = 1.2


class ClickableLineEdit(QLineEdit):
    clicked = Signal()

    def mousePressEvent(self, event):
        self.clicked.emit()
        super().mousePressEvent(event)


class IsoFitsCalculator:
    def __init__(self, data_folder_path):
        self.tolerances_data = []
        self.available_fits = [""]
        self._load_data(data_folder_path)

    def _load_data(self, data_folder_path):
        tolerances_path = os.path.join(data_folder_path, "tolerances.json")
        try:
            with open(tolerances_path, 'r', encoding='utf-8') as f:
                self.tolerances_data = json.load(f)
            all_fits = set(entry["toleranzklasse"] for entry in self.tolerances_data)
            self.available_fits.extend(sorted(list(all_fits)))
        except (FileNotFoundError, json.JSONDecodeError) as e:
            QMessageBox.critical(None, "Fataler Fehler", f"Datei 'tolerances.json' nicht gefunden/lesbar: {e}")
            sys.exit(1)

    def calculate(self, nominal_size, fit_string):
        if not self.tolerances_data:
            return None
        try:
            for entry in self.tolerances_data:
                if (entry["toleranzklasse"].lower() == fit_string.lower() and
                        entry["lowerlimit"] < nominal_size <= entry["upperlimit"]):
                    return entry["es"] / 1000.0, entry["ei"] / 1000.0
            return None
        except Exception:
            return None


# ==============================================================================
#      3. HAUPT-WIDGETS
# ==============================================================================

class DXFWidget(QWidget):
    dimension_clicked = Signal(str)
    text_clicked = Signal(str)
    CLICK_RADIUS_PIXELS = 50

    def __init__(self, parent=None):
        super().__init__(parent)
        self.doc, self.msp = None, None
        self.selection_mode = 'dimension'
        self.scene = QGraphicsScene()
        self.view = QGraphicsView(self.scene)
        self.view.setAcceptDrops(False)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.view)
        self.setLayout(layout)
        self.view.scale(1, -1)
        self.view.setTransformationAnchor(QGraphicsView.ViewportAnchor.AnchorUnderMouse)
        self.view.setResizeAnchor(QGraphicsView.ViewportAnchor.AnchorViewCenter)
        self.view.setDragMode(QGraphicsView.DragMode.NoDrag)

        self._is_panning = False
        self._pan_start_pos = QPointF()

        self.view.mousePressEvent = self.handle_mouse_press
        self.view.mouseMoveEvent = self.handle_mouse_move
        self.view.mouseReleaseEvent = self.handle_mouse_release
        self.view.wheelEvent = self.handle_wheel_event

    def set_selection_mode(self, mode: str):
        if mode in ['dimension', 'text']:
            self.selection_mode = mode
            print(f"INFO: DXF-Auswahlmodus auf '{mode}' gesetzt.")
        else:
            print(f"WARNUNG: Unbekannter Auswahlmodus '{mode}'. Ignoriert.")

    def handle_wheel_event(self, event: QWheelEvent):
        factor = ZOOM_FACTOR if event.angleDelta().y() > 0 else 1 / ZOOM_FACTOR
        self.view.scale(factor, factor)
        super(QGraphicsView, self.view).wheelEvent(event)

    def load_dxf(self, filepath):
        try:
            self.doc, auditor = recover_readfile(filepath)
            self.msp = self.doc.modelspace()
            self.draw_dxf()
        except Exception as e:
            QMessageBox.critical(self, "Fehler", f"DXF-Ladefehler: {e}")

    def draw_dxf(self):
        self.scene.clear()
        self.view.setBackgroundBrush(QColor(0, 0, 0))
        try:
            from ezdxf.addons.drawing.frontend import Frontend
            from ezdxf.addons.drawing.properties import RenderContext
            backend = PyQtBackend(self.scene)
            ctx = RenderContext(self.doc)
            frontend = Frontend(ctx, backend)
            frontend.draw_layout(self.msp, finalize=True)
            self.view.fitInView(self.scene.itemsBoundingRect(), Qt.KeepAspectRatio)
        except Exception as e:
            print(f"Zeichnen fehlgeschlagen: {e}")

    def handle_mouse_press(self, event):
        if event.button() == Qt.MouseButton.RightButton:
            self._is_panning = True
            self._pan_start_pos = event.position()
            self.view.setCursor(Qt.CursorShape.ClosedHandCursor)
            event.accept()
        elif event.button() == Qt.MouseButton.LeftButton:
            if not self.msp:
                super(QGraphicsView, self.view).mousePressEvent(event)
                return

            scene_pos = self.view.mapToScene(event.position().toPoint())
            world_pos = Vec3(scene_pos.x(), scene_pos.y())
            p1 = scene_pos
            p2_pos = event.position() + QPointF(self.CLICK_RADIUS_PIXELS, 0)
            p2 = self.view.mapToScene(p2_pos.toPoint())
            search_radius = abs(p2.x() - p1.x())

            if self.selection_mode == 'dimension':
                self.find_closest_dimension(world_pos, search_radius)
            elif self.selection_mode == 'text':
                self.find_closest_text(world_pos, search_radius)
            event.accept()
        else:
            super(QGraphicsView, self.view).mousePressEvent(event)

    def handle_mouse_move(self, event):
        if self._is_panning:
            delta = event.position() - self._pan_start_pos
            self._pan_start_pos = event.position()
            h_bar = self.view.horizontalScrollBar()
            v_bar = self.view.verticalScrollBar()
            h_bar.setValue(h_bar.value() - int(delta.x()))
            v_bar.setValue(v_bar.value() - int(delta.y()))
            event.accept()
        else:
            super(QGraphicsView, self.view).mouseMoveEvent(event)

    def handle_mouse_release(self, event):
        if event.button() == Qt.MouseButton.RightButton and self._is_panning:
            self._is_panning = False
            self.view.setCursor(Qt.CursorShape.ArrowCursor)
            event.accept()
        else:
            super(QGraphicsView, self.view).mouseReleaseEvent(event)

    def find_closest_dimension(self, world_pos, radius):
        def point_to_line_segment_dist(p: Vec3, a: Vec3, b: Vec3) -> float:
            ab = b - a
            ap = p - a
            ab_mag_sq = ab.magnitude_square
            if ab_mag_sq == 0:
                return ap.magnitude
            t = ap.dot(ab) / ab_mag_sq
            if t < 0.0:
                closest_point = a
            elif t > 1.0:
                closest_point = b
            else:
                closest_point = a + ab * t
            return p.distance(closest_point)

        found_dimensions_in_radius = []
        for dimension in self.msp.query('DIMENSION'):
            try:
                virtual_primitives = dimension.virtual_entities()
                min_dist_for_this_dim = float('inf')
                found_primitive = False
                for primitive in virtual_primitives:
                    dist = -1
                    if primitive.dxftype() == 'TEXT':
                        text_pos = primitive.dxf.get('insert', None)
                        if text_pos: dist = world_pos.distance(text_pos)
                    elif primitive.dxftype() == 'LINE':
                        start = Vec3(primitive.dxf.start)
                        end = Vec3(primitive.dxf.end)
                        dist = point_to_line_segment_dist(world_pos, start, end)
                    if 0 <= dist < radius:
                        found_primitive = True
                        if dist < min_dist_for_this_dim:
                            min_dist_for_this_dim = dist
                if found_primitive:
                    measurement = dimension.get_measurement()
                    if isinstance(measurement, (int, float)):
                        found_dimensions_in_radius.append((min_dist_for_this_dim, measurement, dimension))
            except Exception:
                continue

        if found_dimensions_in_radius:
            found_dimensions_in_radius.sort(key=lambda x: x[0])
            closest_measurement = float(found_dimensions_in_radius[0][1])
            self.dimension_clicked.emit(f"{closest_measurement:.4f}")

    def find_closest_text(self, world_pos, radius):
        found_entity, min_dist = None, float('inf')
        for entity in self.msp.query('TEXT MTEXT'):
            if not entity.dxf.hasattr('insert'): continue
            dist = world_pos.distance(entity.dxf.insert)
            if dist < radius and dist < min_dist:
                min_dist, found_entity = dist, entity
        if found_entity:
            text_content = ""
            if found_entity.dxftype() == 'TEXT':
                text_content = found_entity.dxf.text
            elif found_entity.dxftype() == 'MTEXT':
                text_content = found_entity.plain_text()
            if text_content: self.text_clicked.emit(text_content.strip())


class MessprotokollWidget(QWidget):
    selection_mode_changed = Signal(str)
    field_selected = Signal(object)
    field_manually_edited = Signal(object)
    TOTAL_MEASURES = 18;
    MEASURES_PER_PAGE = 6;
    BLOCKS_PER_PAGE = 2;
    MEASURES_PER_BLOCK = 3
    TOTAL_BLOCKS = TOTAL_MEASURES // MEASURES_PER_BLOCK
    _pos_vals = [f"+{i / 1000.0:.3f}" for i in range(5, 201, 5)]
    _neg_vals = [f"-{i / 1000.0:.3f}" for i in range(5, 201, 5)]
    TOLERANCE_VALUES = ["", "0"] + _neg_vals[::-1] + _pos_vals
    MESSMITTEL_OPTIONS = ["", "Mess-schieber", "Aussen Mikrometer", "Digimar", "Endmaß", "Gewinde-lehrdorn", "Gewinde-lehrring",
                          "Haarlineal", "Innen Mikrometer", "Innenschnell-taster", "Lehrdorn", "Lehrring",
                          "MahrSurf M 310", "Maschinen- taster", "Messuhr", "optisch", "Prüfstifte",
                          "Radius Lehre", "Rugotest", "Steigungs-lehre", "Subito", "Tiefenmaß", "Winkel-messer",
                          "Zeiss", "Zoller"]

    def __init__(self, parent=None):
        super().__init__(parent)
        self.cell_mapping = {}
        self._load_mapping()
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(15)
        script_dir = os.path.dirname(os.path.realpath(__file__))
        data_dir = os.path.join(script_dir, "Data")
        self.iso_calculator = IsoFitsCalculator(data_dir)
        self.nominal_fields, self.upper_tol_combos, self.lower_tol_combos = [], [], []
        self.soll_labels, self.iso_fit_combos, self.messmittel_combos = [], [], []
        self.measure_blocks = []
        self.current_page = 0
        self.total_pages = math.ceil(self.TOTAL_BLOCKS / self.BLOCKS_PER_PAGE)
        self._create_header(main_layout)
        main_layout.addSpacing(25)
        self._create_measure_blocks(main_layout)
        main_layout.addStretch()
        self._create_footer_controls(main_layout)
        self._update_page_view()

    def _create_header(self, main_layout):
        header_grid = QGridLayout()
        header_grid.setColumnStretch(1, 2)
        header_grid.setColumnStretch(5, 3)
        title_label = QLabel("Messprotokoll-Assistent")
        title_label.setStyleSheet("font-size: 28pt; font-weight: bold;")
        header_grid.addWidget(title_label, 0, 0, 1, 5)

        self.zeichnungsnummer_field = ClickableLineEdit()
        self.zeichnungsnummer_field.clicked.connect(self._on_zeichnungsnummer_field_selected)
        header_grid.addWidget(QLabel("Zeichnungsnummer:"), 1, 0)
        header_grid.addWidget(self.zeichnungsnummer_field, 1, 1, 1, 2)

        auftrag_layout = QHBoxLayout()
        auftrag_layout.addWidget(QLabel("Auftrag: AT-25 /"))
        self.auftrag_edit = QLineEdit()
        auftrag_layout.addWidget(self.auftrag_edit)
        header_grid.addLayout(auftrag_layout, 1, 3)
        self.pos_edit = QLineEdit()
        self.pos_edit.setFixedWidth(80)
        header_grid.addWidget(QLabel("Pos.:"), 1, 5)
        header_grid.addWidget(self.pos_edit, 1, 6)
        self.date_edit = QDateEdit(calendarPopup=True, date=QDate.currentDate())
        header_grid.addWidget(QLabel("Datum:"), 1, 8)
        header_grid.addWidget(self.date_edit, 1, 9)
        self.oberflaeche_edit = QLineEdit()
        header_grid.addWidget(QLabel("Oberflächenbehandlung:"), 2, 0)
        header_grid.addWidget(self.oberflaeche_edit, 2, 1, 1, 2)
        self.bemerkungen_edit = QLineEdit()
        header_grid.addWidget(QLabel("Bemerkungen:"), 2, 3)
        header_grid.addWidget(self.bemerkungen_edit, 2, 4, 1, 6)
        logo_label = QLabel()
        logo_label.setFixedSize(150, 150)
        logo_label.setScaledContents(True)
        if os.path.exists("app-logo.png"):
            logo_label.setPixmap(QPixmap("app-logo.png"))
        header_grid.addWidget(logo_label, 0, 11, 3, 1, alignment=Qt.AlignTop | Qt.AlignRight)
        main_layout.addLayout(header_grid)

    def _create_measure_blocks(self, main_layout):
        for block_idx in range(self.TOTAL_BLOCKS):
            block_frame = QFrame()
            grid = QGridLayout(block_frame)
            grid.setSpacing(10)
            grid.addWidget(QLabel("Maß lt.\nZeichnung"), 0, 0, 7, 1)
            soll_qlabel = QLabel("SOLL ➡")
            soll_qlabel.setStyleSheet("font-weight: bold;")
            grid.addWidget(soll_qlabel, 6, 0, 1, 1)
            for col_idx in range(self.MEASURES_PER_BLOCK):
                idx = block_idx * self.MEASURES_PER_BLOCK + col_idx
                col = 1 + col_idx
                grid.addWidget(QLabel(f"Maß {idx + 1}", alignment=Qt.AlignCenter), 0, col)
                nf = ClickableLineEdit(alignment=Qt.AlignCenter)
                nf.clicked.connect(lambda widget=nf: self._on_measure_field_selected(widget))
                nf.textEdited.connect(lambda t, f=nf: self.field_manually_edited.emit(f))
                grid.addWidget(nf, 1, col);
                self.nominal_fields.append(nf)
                grid.addWidget(QLabel("ISO-Toleranz"), 2, col, alignment=Qt.AlignCenter)
                ifc = QComboBox();
                ifc.setEditable(True);
                ifc.addItems(self.iso_calculator.available_fits)
                grid.addWidget(ifc, 3, col);
                self.iso_fit_combos.append(ifc)
                grid.addWidget(QLabel("Messmittel"), 4, col, alignment=Qt.AlignCenter)
                mc = QComboBox();
                mc.addItems(self.MESSMITTEL_OPTIONS)
                grid.addWidget(mc, 5, col);
                self.messmittel_combos.append(mc)
                tol_layout = QGridLayout()
                utc = QComboBox();
                utc.addItems(self.TOLERANCE_VALUES);
                utc.setEditable(True)
                ltc = QComboBox();
                ltc.addItems(self.TOLERANCE_VALUES);
                ltc.setEditable(True)
                tol_layout.addWidget(utc, 0, 0);
                tol_layout.addWidget(QLabel("Größtmaß"), 0, 1)
                tol_layout.addWidget(ltc, 1, 0);
                tol_layout.addWidget(QLabel("Kleinstmaß"), 1, 1)
                self.upper_tol_combos.append(utc);
                self.lower_tol_combos.append(ltc)
                sl = QLabel("---", alignment=Qt.AlignCenter,
                            styleSheet="font-weight: bold; border: 1px solid grey; padding: 6px;")
                grid.addLayout(tol_layout, 6, col);
                grid.addWidget(sl, 7, col);
                self.soll_labels.append(sl)
                nf.textChanged.connect(lambda _, i=idx: self._trigger_iso_fit_calculation(i))
                ifc.currentTextChanged.connect(lambda _, i=idx: self._trigger_iso_fit_calculation(i))
                utc.currentTextChanged.connect(lambda _, i=idx: self._update_soll_wert(i))
                ltc.currentTextChanged.connect(lambda _, i=idx: self._update_soll_wert(i))
            main_layout.addWidget(block_frame)
            self.measure_blocks.append(block_frame)

    def _on_zeichnungsnummer_field_selected(self):
        self.field_selected.emit(self.zeichnungsnummer_field)
        self.selection_mode_changed.emit('text')

    def _on_measure_field_selected(self, widget):
        self.field_selected.emit(widget)
        self.selection_mode_changed.emit('dimension')

    def _create_footer_controls(self, main_layout):
        # === GEÄNDERT: Layout für Buttons angepasst ===
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

        # === NEU: Editieren-Button erstellen und verbinden ===
        self.edit_button = QPushButton("Vorhandenes Protokoll editieren")
        self.edit_button.clicked.connect(self._edit_protokoll)

        self.save_button = QPushButton("Protokoll Speichern")
        self.save_button.setProperty('class', 'success-color')
        self.save_button.clicked.connect(self._save_protokoll)

        bottom_layout = QHBoxLayout()
        bottom_layout.addLayout(pagination_layout, 1)
        bottom_layout.addStretch(1)
        bottom_layout.addWidget(self.edit_button)  # Button hinzugefügt
        bottom_layout.addWidget(self.save_button)
        main_layout.addLayout(bottom_layout)

    def _load_mapping(self):
        try:
            with open("mapping.json", 'r', encoding='utf-8') as f:
                self.cell_mapping = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError) as e:
            QMessageBox.critical(self, "Mapping Fehler",
                                 f"Die Datei 'mapping.json' konnte nicht geladen werden.\n\n{e}\n\nDie Speicherfunktion ist deaktiviert.")
            self.cell_mapping = {}

    def _get_writable_cell(self, sheet, cell_coord):
        cell = sheet[cell_coord]
        if isinstance(cell, MergedCell):
            for merged_range in sheet.merged_cells.ranges:
                if cell.coordinate in merged_range:
                    return sheet.cell(row=merged_range.min_row, column=merged_range.min_col)
        return cell

    # ==========================================================================
    # === NEUE METHODE ZUM EDITIEREN EINES BESTEHENDEN PROTOKOLLS ===
    # ==========================================================================
    def _edit_protokoll(self):
        """Lädt ein bestehendes Excel-Protokoll und aktualisiert es mit den
        in der GUI ausgefüllten Werten."""
        if not self.cell_mapping:
            QMessageBox.warning(self, "Aktion nicht möglich", "Die Mapping-Datei ist fehlerhaft oder fehlt.")
            return

        open_path, _ = QFileDialog.getOpenFileName(
            self,
            "Bestehendes Protokoll zum Editieren auswählen...",
            "",
            "Excel-Dateien (*.xlsx)"
        )

        if not open_path:
            return  # Benutzer hat den Dialog abgebrochen

        try:
            workbook = openpyxl.load_workbook(open_path)
            sheet = workbook["Tabelle1"]

            # --- Header-Daten schreiben (nur wenn ausgefüllt) ---
            header_map = self.cell_mapping.get("header", {})

            znr = self.zeichnungsnummer_field.text().strip()
            if znr and 'zeichnungsnummer' in header_map:
                self._get_writable_cell(sheet, header_map['zeichnungsnummer']).value = znr

            auftrag = self.auftrag_edit.text().strip()
            if auftrag and 'auftrag' in header_map:
                self._get_writable_cell(sheet, header_map['auftrag']).value = auftrag

            pos = self.pos_edit.text().strip()
            if pos and 'position' in header_map:
                self._get_writable_cell(sheet, header_map['position']).value = pos

            # Datum wird immer aktualisiert, da es einen Standardwert hat
            if 'datum' in header_map:
                self._get_writable_cell(sheet, header_map['datum']).value = self.date_edit.date().toString("dd.MM.yyyy")

            oberflaeche = self.oberflaeche_edit.text().strip()
            if oberflaeche and 'oberflaeche' in header_map:
                self._get_writable_cell(sheet,
                                        header_map['oberflaeche']).value = f"Oberflächenbehandlung: {oberflaeche}"

            bemerkungen = self.bemerkungen_edit.text().strip()
            if bemerkungen and 'bemerkungen' in header_map:
                self._get_writable_cell(sheet, header_map['bemerkungen']).value = f"Bemerkungen: {bemerkungen}"

            # --- Messdaten schreiben (nur wenn ausgefüllt) ---
            measure_map = self.cell_mapping.get("measures", [])
            for i in range(self.TOTAL_MEASURES):
                if i < len(measure_map):
                    cell_info = measure_map[i]

                    val = self.nominal_fields[i].text().strip()
                    if val and 'nominal' in cell_info: self._get_writable_cell(sheet, cell_info['nominal']).value = val

                    val = self.iso_fit_combos[i].currentText().strip()
                    if val and 'iso_fit' in cell_info: self._get_writable_cell(sheet, cell_info['iso_fit']).value = val

                    val = self.messmittel_combos[i].currentText().strip()
                    if val and 'messmittel' in cell_info: self._get_writable_cell(sheet,
                                                                                  cell_info['messmittel']).value = val

                    val = self.upper_tol_combos[i].currentText().strip()
                    if val and 'upper_tol' in cell_info: self._get_writable_cell(sheet,
                                                                                 cell_info['upper_tol']).value = val

                    val = self.lower_tol_combos[i].currentText().strip()
                    if val and 'lower_tol' in cell_info: self._get_writable_cell(sheet,
                                                                                 cell_info['lower_tol']).value = val

                    val = self.soll_labels[i].text().strip()
                    if val != "---" and 'soll' in cell_info: self._get_writable_cell(sheet,
                                                                                     cell_info['soll']).value = val

            workbook.save(filename=open_path)
            QMessageBox.information(self, "Erfolg",
                                    f"Das Protokoll '{os.path.basename(open_path)}' wurde erfolgreich aktualisiert.")

        except FileNotFoundError:
            QMessageBox.critical(self, "Fehler", f"Die Datei '{open_path}' wurde nicht gefunden.")
        except KeyError as e:
            QMessageBox.critical(self, "Excel Fehler",
                                 f"Ein Schlüssel im Mapping ('{e}') oder das Arbeitsblatt 'Tabelle1' wurde nicht gefunden.")
        except Exception as e:
            QMessageBox.critical(self, "Aktualisierungsfehler", f"Ein unerwarteter Fehler ist aufgetreten:\n\n{e}")

    def _save_protokoll(self):
        if not self.cell_mapping:
            QMessageBox.warning(self, "Speichern nicht möglich", "Die Mapping-Datei ist fehlerhaft oder fehlt.")
            return

        zeichnungsnummer = self.zeichnungsnummer_field.text().strip()
        auftrag = self.auftrag_edit.text().strip()
        pos = self.pos_edit.text().strip()

        if not zeichnungsnummer or not auftrag or not pos:
            QMessageBox.warning(self, "Fehlende Eingabe",
                                "Bitte füllen Sie die Felder 'Zeichnungsnummer', 'Auftrag' und 'Pos' aus.")
            return

        template_path = "LEERFORMULAR.xlsx"
        if not os.path.exists(template_path):
            QMessageBox.critical(self, "Vorlage fehlt", f"Die Vorlagendatei '{template_path}' wurde nicht gefunden.")
            return

        suggested_filename = f"{zeichnungsnummer}+{auftrag}+{pos}.xlsx"
        save_path, _ = QFileDialog.getSaveFileName(self, "Protokoll speichern unter...", suggested_filename,
                                                   "Excel-Dateien (*.xlsx)")

        if not save_path:
            return

        try:
            workbook = openpyxl.load_workbook(template_path)
            sheet = workbook["Tabelle1"]
            header_map = self.cell_mapping.get("header", {})

            if 'zeichnungsnummer' in header_map: self._get_writable_cell(sheet, header_map[
                'zeichnungsnummer']).value = zeichnungsnummer
            if 'auftrag' in header_map: self._get_writable_cell(sheet, header_map['auftrag']).value = auftrag
            if 'position' in header_map: self._get_writable_cell(sheet, header_map['position']).value = pos
            if 'datum' in header_map: self._get_writable_cell(sheet, header_map[
                'datum']).value = self.date_edit.date().toString("dd.MM.yyyy")

            text_oberflaeche = self.oberflaeche_edit.text()
            if text_oberflaeche and 'oberflaeche' in header_map:
                self._get_writable_cell(sheet,
                                        header_map['oberflaeche']).value = f"Oberflächenbehandlung: {text_oberflaeche}"

            text_bemerkungen = self.bemerkungen_edit.text()
            if text_bemerkungen and 'bemerkungen' in header_map:
                self._get_writable_cell(sheet, header_map['bemerkungen']).value = f"Bemerkungen: {text_bemerkungen}"

            measure_map = self.cell_mapping.get("measures", [])
            for i in range(self.TOTAL_MEASURES):
                if i < len(measure_map):
                    cell_info = measure_map[i]
                    if 'nominal' in cell_info: self._get_writable_cell(sheet, cell_info['nominal']).value = \
                    self.nominal_fields[i].text()
                    if 'iso_fit' in cell_info: self._get_writable_cell(sheet, cell_info['iso_fit']).value = \
                    self.iso_fit_combos[i].currentText()
                    if 'messmittel' in cell_info: self._get_writable_cell(sheet, cell_info['messmittel']).value = \
                    self.messmittel_combos[i].currentText()
                    if 'upper_tol' in cell_info: self._get_writable_cell(sheet, cell_info['upper_tol']).value = \
                    self.upper_tol_combos[i].currentText()
                    if 'lower_tol' in cell_info: self._get_writable_cell(sheet, cell_info['lower_tol']).value = \
                    self.lower_tol_combos[i].currentText()
                    if 'soll' in cell_info: self._get_writable_cell(sheet, cell_info['soll']).value = self.soll_labels[
                        i].text()

            workbook.save(filename=save_path)
            QMessageBox.information(self, "Erfolg", f"Das Protokoll wurde erfolgreich unter '{save_path}' gespeichert.")

        except KeyError as e:
            QMessageBox.critical(self, "Excel Fehler",
                                 f"Ein Schlüssel im Mapping ('{e}') oder das Arbeitsblatt 'Tabelle1' wurde nicht gefunden.")
        except Exception as e:
            QMessageBox.critical(self, "Speicherfehler",
                                 f"Ein unerwarteter Fehler ist beim Speichern aufgetreten:\n\n{e}")

    def _update_page_view(self):
        start = self.current_page * self.BLOCKS_PER_PAGE
        end = start + self.BLOCKS_PER_PAGE
        for i, block in enumerate(self.measure_blocks):
            block.setVisible(start <= i < end)
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
        self._update_soll_wert(index)
        nominal_text = self.nominal_fields[index].text().replace(',', '.')
        fit_string = self.iso_fit_combos[index].currentText().strip()
        if not nominal_text or not fit_string: return
        try:
            result = self.iso_calculator.calculate(float(nominal_text), fit_string)
            if result:
                up_dev, low_dev = result
                self.upper_tol_combos[index].blockSignals(True);
                self.lower_tol_combos[index].blockSignals(True)
                self.upper_tol_combos[index].setCurrentText(f"{up_dev:+.3f}")
                self.lower_tol_combos[index].setCurrentText(f"{low_dev:+.3f}")
                self.upper_tol_combos[index].blockSignals(False);
                self.lower_tol_combos[index].blockSignals(False)
                self._update_soll_wert(index)
        except (ValueError, TypeError):
            pass

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
        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(self.dxf_widget)
        splitter.addWidget(self.protokoll_widget)
        splitter.setSizes([1000, 1200])
        self.setCentralWidget(splitter)
        menu_bar = self.menuBar()
        file_menu = menu_bar.addMenu("Datei")
        open_action = QAction("DXF Öffnen...", self)
        open_action.triggered.connect(self.open_file)
        file_menu.addAction(open_action)
        self.setWindowTitle("Messprotokoll-Assistent")
        self.setGeometry(50, 50, 2200, 1200)
        self.protokoll_widget.selection_mode_changed.connect(self.dxf_widget.set_selection_mode)
        self.protokoll_widget.field_selected.connect(self.on_field_selected)
        self.dxf_widget.dimension_clicked.connect(self.on_dimension_value_received)
        self.dxf_widget.text_clicked.connect(self.on_text_value_received)
        self.protokoll_widget.field_manually_edited.connect(self.on_field_manually_edited)

    def _set_application_style(self):
        app = QApplication.instance()
        extra = {'accent_color': '#448AFF', 'secondaryLightColor': "#31363B"}
        apply_stylesheet(app, theme='dark_blue.xml', extra=extra)

    def on_field_selected(self, widget):
        if self.target_widget: self.target_widget.setStyleSheet("")
        self.target_widget = widget
        self.target_widget.setStyleSheet("border: 2px solid #448AFF;")

    def on_dimension_value_received(self, value):
        if self.target_widget:
            self.target_widget.setText(value.replace('.', ','))
            self.target_widget.setStyleSheet("")
            self.target_widget = None
            self.dxf_widget.set_selection_mode('dimension')
        else:
            QMessageBox.information(self, "Hinweis", "Bitte klicken Sie zuerst in ein 'Maß lt. Zeichnung'-Feld.")

    def on_text_value_received(self, value):
        if self.target_widget:
            self.target_widget.setText(value)
            self.target_widget.setStyleSheet("")
            self.target_widget = None
            self.dxf_widget.set_selection_mode('dimension')
        else:
            QMessageBox.information(self, "Hinweis", "Bitte klicken Sie zuerst in das 'Zeichnungsnummer'-Feld.")

    def on_field_manually_edited(self, widget):
        if self.target_widget == widget:
            self.target_widget.setStyleSheet("")
            self.target_widget = None
            self.dxf_widget.set_selection_mode('dimension')

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