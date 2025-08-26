# ==============================================================================
#      1. IMPORTS
# ==============================================================================
import sys
import os
import json
import ezdxf

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
    from openpyxl.drawing.image import Image as OpenpyxlImage
except ImportError:
    print("FEHLER: 'openpyxl' wurde nicht gefunden. Bitte installieren mit: pip install openpyxl Pillow")
    sys.exit(1)

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
TOTAL_MEASUREMENTS = 18


class ClickableLineEdit(QLineEdit):
    clicked = Signal()

    def mousePressEvent(self, event):
        self.clicked.emit()
        super().mousePressEvent(event)


class IsoFitsCalculator:
    # (Unverändert)
    def __init__(self, data_folder_path):
        self.tolerances_data = []
        self._load_data(data_folder_path)

    def _load_data(self, data_folder_path):
        tolerances_path = os.path.join(data_folder_path, "tolerances.json")
        try:
            with open(tolerances_path, 'r', encoding='utf-8') as f:
                self.tolerances_data = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError) as e:
            QMessageBox.critical(None, "Fataler Fehler",
                                 f"Die Datei 'tolerances.json' konnte nicht im 'Data'-Ordner gefunden werden.\n\n{e}")
            sys.exit(1)

    def calculate(self, nominal_size, fit_string):
        if not self.tolerances_data: return None
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
    # (Fast unverändert, nur auf 18 Maße erweitert)
    field_selected = Signal(object)
    field_manually_edited = Signal(object)

    ISO_FITS_COMMON = ["", "H6", "H7", "H8", "H9", "H11", "G7", "F7", "E9", "D10", "P7", "h6", "h7", "h9", "h11", "g6",
                       "f7", "e8", "d9", "k6", "js6", "s7"]
    MESSMITTEL_OPTIONS = ["", "optisch", "Messschieber", "Bügelmessschraube", "Höhenmessgerät", "3D-Messmaschine"]
    KUNDEN_LISTE = ["", "Tool Service GmbH", "Musterfirma AG", "Projekt X Kunde"]

    def __init__(self, parent=None):
        super().__init__(parent)
        main_layout = QVBoxLayout(self);
        main_layout.setSpacing(15)

        script_dir = os.path.dirname(os.path.realpath(__file__))
        data_dir = os.path.join(script_dir, "Data")
        self.iso_calculator = IsoFitsCalculator(data_dir)

        self.nominal_fields, self.upper_tol_combos, self.lower_tol_combos, self.soll_labels, self.iso_fit_combos, self.messmittel_combos = [], [], [], [], [], []

        header_grid = QGridLayout()
        header_grid.setColumnStretch(1, 2);
        header_grid.setColumnStretch(10, 1)
        header_grid.addWidget(QLabel("Messprotokoll-Assistent", font=QFont("Arial", 30, QFont.Bold)), 0, 0, 1, 5)
        self.kunde_combo = QComboBox();
        self.kunde_combo.setEditable(True);
        self.kunde_combo.addItems(self.KUNDEN_LISTE)
        header_grid.addWidget(QLabel("Kunde:"), 1, 0);
        header_grid.addWidget(self.kunde_combo, 1, 1)
        auftrag_layout = QHBoxLayout();
        auftrag_layout.addWidget(QLabel("Auftrag: AT-25 /"));
        self.auftrag_edit = QLineEdit();
        auftrag_layout.addWidget(self.auftrag_edit)
        header_grid.addLayout(auftrag_layout, 1, 3)
        self.pos_edit = QLineEdit();
        self.pos_edit.setFixedWidth(80)
        header_grid.addWidget(QLabel("Pos.:"), 1, 5);
        header_grid.addWidget(self.pos_edit, 1, 6)
        self.date_edit = QDateEdit(calendarPopup=True, date=QDate.currentDate())
        header_grid.addWidget(QLabel("Datum:"), 1, 8);
        header_grid.addWidget(self.date_edit, 1, 9)
        logo_label = QLabel();
        logo_label.setFixedSize(200, 200);
        logo_label.setScaledContents(True)
        logo_path = "app-logo.png"
        if os.path.exists(logo_path):
            logo_label.setPixmap(QPixmap(logo_path))
        else:
            print(f"WARNUNG: Logo-Datei '{logo_path}' nicht gefunden.")
        header_grid.addWidget(logo_label, 0, 11, 2, 1, alignment=Qt.AlignTop | Qt.AlignRight)
        main_layout.addLayout(header_grid)

        measures_per_row = 6
        num_rows = TOTAL_MEASUREMENTS // measures_per_row

        for row_idx in range(num_rows):
            block_frame = QFrame();
            block_frame.setFrameShape(QFrame.StyledPanel)
            grid = QGridLayout(block_frame);
            grid.setSpacing(10)
            grid.addWidget(QLabel("Maß lt.\nZeichnung"), 0, 0, 7, 1)
            grid.addWidget(QLabel("SOLL ➡", font=QFont("Arial", 10, QFont.Bold)), 6, 0, 1, 1)

            for col_idx in range(measures_per_row):
                measure_index = row_idx * measures_per_row + col_idx
                col_start = 1 + col_idx
                grid.addWidget(QLabel(f"Maß {measure_index + 1}", alignment=Qt.AlignCenter), 0, col_start)
                nominal_field = ClickableLineEdit(alignment=Qt.AlignCenter,
                                                  styleSheet="background-color: #e9f5e9; color: #333333;")
                nominal_field.clicked.connect(lambda f=nominal_field: self.field_selected.emit(f))
                nominal_field.textEdited.connect(lambda text, f=nominal_field: self.field_manually_edited.emit(f))
                grid.addWidget(nominal_field, 1, col_start)
                self.nominal_fields.append(nominal_field)
                grid.addWidget(QLabel("ISO-Fit", alignment=Qt.AlignCenter), 2, col_start)
                iso_fit_combo = QComboBox();
                iso_fit_combo.setEditable(True);
                iso_fit_combo.addItems(self.ISO_FITS_COMMON)
                grid.addWidget(iso_fit_combo, 3, col_start)
                self.iso_fit_combos.append(iso_fit_combo)
                grid.addWidget(QLabel("Messmittel", alignment=Qt.AlignCenter), 4, col_start)
                messmittel_combo = QComboBox();
                messmittel_combo.addItems(self.MESSMITTEL_OPTIONS)
                grid.addWidget(messmittel_combo, 5, col_start)
                self.messmittel_combos.append(messmittel_combo)
                tol_layout = QGridLayout()
                upper_tol_combo = QComboBox();
                upper_tol_combo.setEditable(True)
                lower_tol_combo = QComboBox();
                lower_tol_combo.setEditable(True)
                tol_layout.addWidget(upper_tol_combo, 0, 0);
                tol_layout.addWidget(QLabel("Größtmaß"), 0, 1)
                tol_layout.addWidget(lower_tol_combo, 1, 0);
                tol_layout.addWidget(QLabel("Kleinstmaß"), 1, 1)
                self.upper_tol_combos.append(upper_tol_combo);
                self.lower_tol_combos.append(lower_tol_combo)
                soll_label = QLabel("---", alignment=Qt.AlignCenter,
                                    styleSheet="font-weight: bold; border: 1px solid #ccc; padding: 6px;")
                grid.addLayout(tol_layout, 6, col_start)
                grid.addWidget(soll_label, 7, col_start)
                self.soll_labels.append(soll_label)
                nominal_field.textChanged.connect(lambda _, idx=measure_index: self._trigger_iso_fit_calculation(idx))
                iso_fit_combo.currentTextChanged.connect(
                    lambda _, idx=measure_index: self._trigger_iso_fit_calculation(idx))
                upper_tol_combo.currentTextChanged.connect(lambda _, idx=measure_index: self._update_soll_wert(idx))
                lower_tol_combo.currentTextChanged.connect(lambda _, idx=measure_index: self._update_soll_wert(idx))
            main_layout.addWidget(block_frame)
        main_layout.addStretch()

    def _trigger_iso_fit_calculation(self, index):
        # (Unverändert)
        nominal_text = self.nominal_fields[index].text().replace(',', '.')
        fit_string = self.iso_fit_combos[index].currentText().strip()
        self._update_soll_wert(index)
        if not nominal_text or not fit_string: return
        try:
            nominal_value = float(nominal_text)
            result = self.iso_calculator.calculate(nominal_value, fit_string)
            if result is None: return
            upper_dev, lower_dev = result
            upper_str = f"{upper_dev:+.3f}"
            lower_str = f"{lower_dev:+.3f}"
            self.upper_tol_combos[index].blockSignals(True)
            self.lower_tol_combos[index].blockSignals(True)
            self.upper_tol_combos[index].setCurrentText(upper_str)
            self.lower_tol_combos[index].setCurrentText(lower_str)
            self.upper_tol_combos[index].blockSignals(False)
            self.lower_tol_combos[index].blockSignals(False)
            self._update_soll_wert(index)
        except Exception as e:
            print(f"Fehler bei ISO-Fit-Verarbeitung für Index {index}: {e}")

    def _update_soll_wert(self, index):
        # (Unverändert)
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
        save_action = QAction("Speichern unter...", self);
        save_action.triggered.connect(self.save_to_styled_excel);
        file_menu.addAction(save_action)

        self.setWindowTitle("Messprotokoll-Assistent (Fusion Style)");
        self.setGeometry(50, 50, 1800, 1000)
        self.protokoll_widget.field_selected.connect(self.on_protokoll_field_selected)
        self.dxf_widget.dimension_clicked.connect(self.on_dimension_value_received)
        self.protokoll_widget.field_manually_edited.connect(self.on_field_manually_edited)

    def _set_application_style(self):
        # (Unverändert)
        app = QApplication.instance()
        if sys.platform == "darwin":
            app.setStyle(QStyleFactory.create("Macintosh"))
        else:
            app.setStyle(QStyleFactory.create("Fusion"))
        app.setPalette(app.style().standardPalette())
        app.setStyleSheet("")

    def on_protokoll_field_selected(self, widget):
        if self.target_widget: self.target_widget.setStyleSheet("background-color: #e9f5e9; color: #333333;")
        self.target_widget = widget
        self.target_widget.setStyleSheet("background-color: #0078d7; color: white;")

    def on_dimension_value_received(self, value):
        if self.target_widget:
            self.target_widget.setText(value.replace(',', '.'))
            self.target_widget.setStyleSheet("background-color: #e9f5e9; color: #333333;")
            self.target_widget = None
        else:
            QMessageBox.information(self, "Hinweis", "Bitte klicken Sie zuerst in ein 'Maß lt. Zeichnung'-Feld.")

    def on_field_manually_edited(self, widget):
        if self.target_widget == widget:
            self.target_widget.setStyleSheet("background-color: #e9f5e9; color: #333333;")
            self.target_widget = None

    def open_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "DXF-Datei öffnen", "", "DXF-Dateien (*.dxf)")
        if path: self.dxf_widget.load_dxf(path)

    def dragEnterEvent(self, event: QDragEnterEvent):
        # (Unverändert)
        if event.mimeData().hasUrls():
            for url in event.mimeData.urls():
                if url.isLocalFile() and url.toLocalFile().lower().endswith('.dxf'):
                    event.acceptProposedAction();
                    return
        event.ignore()

    def dropEvent(self, event: QDropEvent):
        # (Unverändert)
        if event.mimeData().hasUrls():
            for url in event.mimeData.urls():
                file_path = url.toLocalFile()
                if file_path.lower().endswith('.dxf'):
                    self.dxf_widget.load_dxf(file_path);
                    event.acceptProposedAction();
                    return
        event.ignore()

    # --------------------------------------------------------------------------
    #      4.2 Excel-Export-Funktionalität
    # --------------------------------------------------------------------------

    def save_to_styled_excel(self):
        save_path, _ = QFileDialog.getSaveFileName(self, "Messprotokoll speichern unter...", "",
                                                   "Excel-Dateien (*.xlsx)")
        if not save_path:
            return

        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Messprotokoll"

            # --- STYLES DEFINIEREN ---
            font_bold = Font(name='Arial', size=10, bold=True)
            font_header = Font(name='Arial', size=20, bold=True)
            font_small = Font(name='Arial', size=8)
            font_normal = Font(name='Arial', size=9)
            font_blue_logo = Font(name='Arial', size=14, bold=True, color='4472C4')

            fill_light_blue = PatternFill(start_color="C5D9F1", end_color="C5D9F1", fill_type="solid")
            fill_yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
            fill_gray_header = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")

            thin_border = Side(border_style="thin", color="000000")
            thick_border = Side(border_style="thick", color="000000")

            all_borders = Border(top=thin_border, left=thin_border,
                                 right=thin_border, bottom=thin_border)
            thick_box = Border(top=thick_border, left=thick_border,
                               right=thick_border, bottom=thick_border)

            center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
            left_align = Alignment(horizontal='left', vertical='center')
            right_align = Alignment(horizontal='right', vertical='center')

            # --- HEADER INFORMATIONEN (oben rechts) ---
            ws['U1'] = "Ersteller: Huemer Christine"
            ws['U2'] = "Revision: 06"
            ws['U3'] = "Letzte Änderung: 12.11.2024"
            ws['U4'] = "Geändert durch: Huemer Christine"

            for cell_ref in ['U1', 'U2', 'U3', 'U4']:
                ws[cell_ref].font = font_small
                ws[cell_ref].alignment = right_align

            # --- ZEILE 6: Kunde, Auftrag, Position, Datum ---
            # ERST die Werte setzen, DANN zusammenführen
            ws['A6'] = "Kunde:"
            ws['A6'].font = font_bold
            ws['A6'].border = thick_box

            ws['B6'] = self.protokoll_widget.kunde_combo.currentText()
            ws['B6'].border = thick_box

            ws['F6'] = "Auftrag: AT- 25 /"
            ws['F6'].font = font_bold
            ws['F6'].border = thick_box

            ws['G6'] = self.protokoll_widget.auftrag_edit.text()
            ws['G6'].border = thick_box

            ws['J6'] = "Pos.:"
            ws['J6'].font = font_bold
            ws['J6'].border = thick_box

            ws['K6'] = self.protokoll_widget.pos_edit.text()
            ws['K6'].border = thick_box

            ws['L6'] = "Dat.:"
            ws['L6'].font = font_bold
            ws['L6'].border = thick_box

            ws['M6'] = self.protokoll_widget.date_edit.date().toString("dd.MM.yyyy")
            ws['M6'].border = thick_box

            # Jetzt zusammenführen
            ws.merge_cells('B6:E6')  # Kunde Feld vergrößern
            ws.merge_cells('G6:I6')  # Auftrag Feld vergrößern

            # --- ZEILE 8-9: Messprotokoll Überschrift ---
            ws['A8'] = "Messprotokoll"
            ws['A8'].font = font_header
            ws['A8'].alignment = center_align
            ws.merge_cells('A8:T9')

            # Tool Service Logo
            ws['U8'] = "TOOL"
            ws['U9'] = "SERVICE"
            ws['U8'].font = font_blue_logo
            ws['U9'].font = font_blue_logo
            ws['U8'].alignment = center_align
            ws['U9'].alignment = center_align

            # --- ZEILE 11-12: Oberflächenbehandlung und Bemerkungen ---
            ws['A11'] = "Oberflächenbehandlung:"
            ws['A11'].font = font_bold
            ws['A11'].border = all_borders
            ws.merge_cells('A11:U11')

            ws['A12'] = "Bemerkungen:"
            ws['A12'].font = font_bold
            ws['A12'].border = all_borders
            ws.merge_cells('A12:U12')

            # --- HAUPTTABELLE AB ZEILE 14 ---
            start_row = 14

            # ZEILE 1: Messmittel Headers
            # Erst alle Werte setzen
            ws.cell(row=start_row, column=1, value='Messmittel')
            ws.cell(row=start_row, column=2, value='Messmittel')
            ws.cell(row=start_row, column=3, value='Messmittel')

            for i in range(18):
                ws.cell(row=start_row, column=4 + i, value='Messmittel')

            ws.cell(row=start_row, column=22, value='Messmittel')
            ws.cell(row=start_row, column=23, value='Messmittel')
            ws.cell(row=start_row, column=24, value='NAME')

            # Dann formatieren
            for col in range(1, 25):
                cell = ws.cell(row=start_row, column=col)
                cell.font = font_bold
                cell.fill = fill_gray_header
                cell.border = all_borders
                cell.alignment = center_align

            # NAME Spalte über zwei Zellen zusammenführen
            ws.merge_cells(f'X{start_row}:Y{start_row}')  # Spalten 24-25

            # ZEILE 2: Messmittel Werte
            row2 = start_row + 1

            # Erste drei Spalten
            ws.cell(row=row2, column=1, value='optisch')
            ws.cell(row=row2, column=2, value='optisch')
            ws.cell(row=row2, column=3, value='')

            # Messmittel für Maß 1-18
            for i in range(18):
                messmittel_text = self.protokoll_widget.messmittel_combos[i].currentText()
                ws.cell(row=row2, column=4 + i, value=messmittel_text)

            # Letzte Spalten
            ws.cell(row=row2, column=22, value='')
            ws.cell(row=row2, column=23, value='')
            ws.cell(row=row2, column=24, value='A')
            ws.cell(row=row2, column=25, value='B')

            # Formatierung
            for col in range(1, 26):
                cell = ws.cell(row=row2, column=col)
                cell.font = font_normal
                cell.border = all_borders
                cell.alignment = center_align

                # Gelbe Spalten für erste drei Spalten
                if col <= 3:
                    cell.fill = fill_yellow

            # ZEILE 3: Beschriftungen
            row3 = start_row + 2

            ws.cell(row=row3, column=1, value='Maß lt.\nZeichnung')
            ws.cell(row=row3, column=2, value='keine Be-\nschädig-\nungen')
            ws.cell(row=row3, column=3, value='keine\nGrate')

            for i in range(18):
                ws.cell(row=row3, column=4 + i, value=f'Maß {i + 1}')

            ws.cell(row=row3, column=22, value=f'Maß 17')
            ws.cell(row=row3, column=23, value=f'Maß 18')
            ws.cell(row=row3, column=24, value='Bewer-\ntung')

            # Formatierung
            for col in range(1, 26):
                cell = ws.cell(row=row3, column=col)
                cell.font = font_bold
                cell.border = all_borders
                cell.alignment = center_align

                # Färbung
                if col <= 3:
                    cell.fill = fill_yellow
                elif col == 24:  # Bewertung über beide Spalten
                    ws.merge_cells(f'X{row3}:Y{row3}')
                    break
                else:
                    cell.fill = fill_gray_header

            # ZEILE 4: Nominal-Werte (IST-Werte)
            row4 = start_row + 3

            # Erste 3 Spalten leer, aber gelb
            for col in range(1, 4):
                cell = ws.cell(row=row4, column=col, value="")
                cell.fill = fill_yellow
                cell.border = all_borders

            # Nominal-Werte für Maß 1-18
            for i in range(18):
                col = 4 + i
                try:
                    nominal_text = self.protokoll_widget.nominal_fields[i].text().replace(',', '.')
                    if nominal_text:
                        value = float(nominal_text)
                    else:
                        value = ""
                except (ValueError, TypeError):
                    value = self.protokoll_widget.nominal_fields[i].text()

                cell = ws.cell(row=row4, column=col, value=value)
                cell.border = all_borders
                cell.alignment = center_align

            # Letzte Spalten
            for col in range(22, 26):
                cell = ws.cell(row=row4, column=col, value="")
                cell.border = all_borders

            # ZEILE 5: SOLL-Werte
            row5 = start_row + 4

            # Erste Spalte: "SOLL ➡"
            cell = ws.cell(row=row5, column=1, value="SOLL ➡")
            cell.font = font_bold
            cell.fill = fill_light_blue
            cell.border = all_borders
            cell.alignment = center_align

            # Spalten 2-3: gelb und leer
            for col in range(2, 4):
                cell = ws.cell(row=row5, column=col, value="")
                cell.fill = fill_yellow
                cell.border = all_borders

            # SOLL-Werte für Maß 1-18
            for i in range(18):
                col = 4 + i
                try:
                    soll_text = self.protokoll_widget.soll_labels[i].text().replace(',', '.')
                    if soll_text and soll_text != "---":
                        value = float(soll_text)
                    else:
                        value = ""
                except (ValueError, TypeError):
                    value = self.protokoll_widget.soll_labels[i].text() if self.protokoll_widget.soll_labels[
                                                                               i].text() != "---" else ""

                cell = ws.cell(row=row5, column=col, value=value)
                cell.border = all_borders
                cell.alignment = center_align

            # Letzte Spalten
            for col in range(22, 26):
                cell = ws.cell(row=row5, column=col, value="")
                cell.border = all_borders

            # ZEILE 6: Anfahrteil
            row6 = start_row + 5

            cell = ws.cell(row=row6, column=1, value="Anfahrteil")
            cell.font = font_bold
            cell.fill = fill_light_blue
            cell.border = all_borders
            cell.alignment = center_align

            # Rest der Zeile
            for col in range(2, 26):
                cell = ws.cell(row=row6, column=col, value="")
                cell.border = all_borders
                if col <= 3:
                    cell.fill = fill_yellow

            # WEITERE MESSREIHEN (15 zusätzliche Zeilen)
            for row_offset in range(1, 16):
                current_row = row6 + row_offset

                # Erste Spalte blau
                cell = ws.cell(row=current_row, column=1, value="")
                cell.fill = fill_light_blue
                cell.border = all_borders

                # Rest der Zeile
                for col in range(2, 26):
                    cell = ws.cell(row=current_row, column=col, value="")
                    cell.border = all_borders
                    if col <= 3:
                        cell.fill = fill_yellow

            # --- SPALTENBREITEN OPTIMIEREN ---
            column_widths = {
                'A': 12,  # Maß lt. Zeichnung
                'B': 10,  # keine Beschädigungen
                'C': 8,  # keine Grate
            }

            # Maß-Spalten D-U (Maß 1-18)
            for i in range(18):
                col_letter = chr(ord('D') + i)
                column_widths[col_letter] = 7

            # Zusätzliche Spalten V, W für Maß 17, 18
            column_widths['V'] = 7  # Maß 17
            column_widths['W'] = 7  # Maß 18
            column_widths['X'] = 6  # NAME A
            column_widths['Y'] = 6  # NAME B
            column_widths['U'] = 25  # Header-Info

            for col, width in column_widths.items():
                ws.column_dimensions[col].width = width

            # --- ZEILENHÖHEN OPTIMIEREN ---
            row_heights = {
                start_row: 25,  # Messmittel Header
                start_row + 1: 20,  # Messmittel Werte
                start_row + 2: 35,  # Beschriftungen
                start_row + 3: 20,  # Nominal-Werte
                start_row + 4: 20,  # SOLL
                start_row + 5: 20,  # Anfahrteil
            }

            for row, height in row_heights.items():
                ws.row_dimensions[row].height = height

            # Weitere Messreihen
            for i in range(1, 16):
                ws.row_dimensions[start_row + 5 + i].height = 18

            # --- DRUCKBEREICH UND SEITENEINRICHTUNG ---
            ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
            ws.page_setup.paperSize = ws.PAPERSIZE_A4
            ws.page_margins.left = 0.5
            ws.page_margins.right = 0.5
            ws.page_margins.top = 0.75
            ws.page_margins.bottom = 0.75

            # Speichern
            wb.save(save_path)
            QMessageBox.information(self, "Erfolg", f"Messprotokoll erfolgreich gespeichert unter:\n{save_path}")

        except Exception as e:
            QMessageBox.critical(self, "Export-Fehler",
                                 f"Ein Fehler ist beim Speichern der Excel-Datei aufgetreten:\n{e}")
            import traceback
            print(f"Detaillierter Fehler: {traceback.format_exc()}")


# ==============================================================================
#      5. START DER APPLIKATION
# ==============================================================================
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())