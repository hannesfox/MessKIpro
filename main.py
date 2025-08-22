import sys
import ezdxf
from ezdxf.recover import readfile as recover_readfile
from ezdxf.addons.drawing.pyqt import PyQtBackend
from ezdxf.math import Vec3

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QGraphicsScene, QGraphicsView, QMessageBox,
    QFileDialog,
)
from PySide6.QtGui import QColor, QWheelEvent, QAction
from PySide6.QtCore import Qt, QPoint, QPointF

# Faktor für das Zoomen mit dem Mausrad
ZOOM_FACTOR = 1.2


class DXFViewer(QMainWindow):
    """
    Ein einfacher DXF-Viewer mit PySide6 und ezdxf, der das Auslesen von
    Bemaßungsinformationen per Mausklick ermöglicht.
    """
    # Radius in Pixeln auf dem Bildschirm, der für die Klick-Suche verwendet wird.
    CLICK_RADIUS_PIXELS = 50

    def __init__(self):
        """Initialisiert das Hauptfenster und die GUI-Komponenten."""
        super().__init__()
        self.doc = None
        self.msp = None

        self.scene = QGraphicsScene()
        self.view = QGraphicsView(self.scene)

        self.view.scale(1, -1)
        self.view.setTransformationAnchor(QGraphicsView.ViewportAnchor.AnchorUnderMouse)
        self.view.setResizeAnchor(QGraphicsView.ViewportAnchor.AnchorViewCenter)
        self.view.setDragMode(QGraphicsView.DragMode.NoDrag)

        self.setCentralWidget(self.view)

        menu_bar = self.menuBar()
        file_menu = menu_bar.addMenu("Datei")
        open_action = QAction("Öffnen...", self)
        open_action.triggered.connect(self.open_file)
        file_menu.addAction(open_action)

        self.setWindowTitle("DXF Bemaßungs-Viewer (PySide6)")
        self.setGeometry(100, 100, 1200, 800)

        self.view.mousePressEvent = self.handle_mouse_press
        self.view.wheelEvent = self.handle_wheel_event

    def handle_wheel_event(self, event: QWheelEvent):
        """Verarbeitet das Mausrad-Event zum Zoomen."""
        if event.angleDelta().y() > 0:
            self.view.scale(ZOOM_FACTOR, ZOOM_FACTOR)
        else:
            self.view.scale(1 / ZOOM_FACTOR, 1 / ZOOM_FACTOR)
        super(QGraphicsView, self.view).wheelEvent(event)

    def open_file(self):
        """Öffnet einen Datei-Dialog und lädt die ausgewählte DXF-Datei."""
        path, _ = QFileDialog.getOpenFileName(
            self, "DXF-Datei öffnen", "", "DXF-Dateien (*.dxf)"
        )
        if path:
            self.load_dxf(path)

    def load_dxf(self, filepath):
        """Lädt eine DXF-Datei mit ezdxf und startet den Zeichenvorgang."""
        try:
            self.doc, auditor = recover_readfile(filepath)
            if auditor.has_errors:
                print(f"DXF-Fehler gefunden: {len(auditor.errors)} Fehler.")
            self.msp = self.doc.modelspace()
            self.draw_dxf()
        except Exception as e:
            QMessageBox.critical(
                self, "Fehler", f"Ein unerwarteter Fehler ist aufgetreten: {e}"
            )
            print(f"Fehlerdetails: {e}")

    def draw_dxf(self):
        """Rendert den Modellbereich der DXF-Datei in die QGraphicsScene."""
        self.scene.clear()
        self.view.setBackgroundBrush(QColor(30, 30, 30))

        try:
            from ezdxf.addons.drawing.frontend import Frontend
            from ezdxf.addons.drawing.properties import RenderContext
            backend = PyQtBackend(self.scene)
            ctx = RenderContext(self.doc)
            frontend = Frontend(ctx, backend)
            frontend.draw_layout(self.msp, finalize=True)
        except Exception as e:
            print(f"Zeichnen fehlgeschlagen: {e}")
            QMessageBox.warning(self, "Warnung", "DXF konnte nicht gezeichnet werden.")

        self.view.setSceneRect(self.scene.itemsBoundingRect())
        self.view.fitInView(self.scene.itemsBoundingRect(), Qt.KeepAspectRatio)

    def handle_mouse_press(self, event):
        """Verarbeitet einen Mausklick, um nach Bemaßungen in der Nähe zu suchen."""
        super(QGraphicsView, self.view).mousePressEvent(event)

        if event.button() == Qt.MouseButton.LeftButton and self.msp:
            # event.position() verwenden, um DeprecationWarning zu vermeiden.
            view_pos_f = event.position()

            # === KORREKTUR HIER ===
            # Der Traceback zeigt, dass mapToScene ein QPoint (int) erwartet, kein QPointF (float).
            # Wir müssen das Ergebnis von event.position() explizit umwandeln.
            view_pos_int = view_pos_f.toPoint()

            scene_pos = self.view.mapToScene(view_pos_int)
            world_pos = Vec3(scene_pos.x(), scene_pos.y())

            # === ROBUSTE SUCHLOGIK ===
            p1 = self.view.mapToScene(view_pos_int)
            # Hier addieren wir zum präzisen QPointF und konvertieren erst dann zum QPoint für die mapToScene-Funktion.
            p2_pos = view_pos_f + QPointF(self.CLICK_RADIUS_PIXELS, 0)
            p2 = self.view.mapToScene(p2_pos.toPoint())

            search_radius_world = abs(p2.x() - p1.x())

            found_dimension = None
            min_dist = float('inf')

            for dimension in self.msp.query('DIMENSION'):
                points_to_check = []
                if hasattr(dimension.dxf, 'defpoint'):
                    points_to_check.append(dimension.dxf.defpoint)
                if hasattr(dimension.dxf, 'text_midpoint'):
                    points_to_check.append(dimension.dxf.text_midpoint)

                for point in points_to_check:
                    if point is None:
                        continue
                    dist = world_pos.distance(point)
                    if dist < search_radius_world and dist < min_dist:
                        min_dist = dist
                        found_dimension = dimension

            if found_dimension:
                try:
                    measurement = found_dimension.get_measurement()
                    text = found_dimension.dxf.text

                    if isinstance(measurement, Vec3):
                        measurement_text = f"X={measurement.x:.4f}, Y={measurement.y:.4f}"
                    else:
                        measurement_text = f"{measurement:.4f}"

                    if text == "<>" or text is None or text.strip() == "":
                        message = f"Messwert: {measurement_text}"
                    else:
                        message = f"Messwert: {measurement_text}\nBenutzerdefinierter Text: {text}"

                    QMessageBox.information(self, "Bemaßungsinformation", message)
                except Exception as e:
                    QMessageBox.warning(
                        self, "Fehler", f"Konnte Messwert nicht auslesen: {e}"
                    )


if __name__ == "__main__":
    app = QApplication(sys.argv)
    viewer = DXFViewer()
    viewer.show()
    sys.exit(app.exec())