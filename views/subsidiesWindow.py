from typing import Dict, Any
from decimal import Decimal, InvalidOperation
import csv
import os

from PyQt5 import QtWidgets
from PyQt5.QtCore import pyqtSignal, Qt
from PyQt5.QtGui import QFont, QCursor, QColor

from config.roles import Permisos, Roles
from config.settings import Settings
from utils.logger import app_logger
from services.subsidyService import SubsidioService


class SubsidiosWindow(QtWidgets.QWidget):
    back_requested = pyqtSignal()

    def __init__(self, user_data: dict, parent=None):
        super().__init__(parent)
        self.user_data = user_data or {}
        self.user_id = self.user_data.get("_id") or self.user_data.get("user_id")
        self.corredor_id = self.user_data.get("corredor_id") or self.user_id
        self.role = self.user_data.get("rol", "cliente")

        try:
            self.service = SubsidioService(corredor_id=self.corredor_id, user_id=self.user_id)
        except Exception as e:
            app_logger.error(f"Error iniciando SubsidioService: {e}")
            self.service = None

        self._can_edit = Permisos.tiene_permiso(self.role, "subsidios")
        self._is_admin = (self.role == Roles.ADMINISTRADOR)

        self._build_ui()
        self._load_data()

    def _build_ui(self):
        self.setWindowTitle("Gestión de Subsidios")
        self.setMinimumSize(Settings.WINDOW_MIN_WIDTH, Settings.WINDOW_MIN_HEIGHT)

        main_layout = QtWidgets.QVBoxLayout(self)
        main_layout.setContentsMargins(18, 12, 18, 12)
        main_layout.setSpacing(12)

        # Header con botón volver y título
        header_layout = QtWidgets.QHBoxLayout()
        back_button = QtWidgets.QPushButton("←")
        back_button.setFont(QFont("Segoe UI", 12))
        back_button.setCursor(QCursor(Qt.PointingHandCursor))
        back_button.clicked.connect(lambda: self.back_requested.emit())
        back_button.setStyleSheet("background: transparent; border: none; color: #3498db; padding: 6px 8px; font-size: 16px;")
        header_layout.addWidget(back_button)

        title = QtWidgets.QLabel("Gestión de Subsidios y Beneficios")
        title.setFont(QFont("Segoe UI", 16, QFont.Bold))
        title.setStyleSheet("color: #2c3e50; margin-left: 6px;")
        header_layout.addWidget(title)
        header_layout.addStretch()

        main_layout.addLayout(header_layout)

        # Toolbar
        self._add_toolbar(main_layout)

        # Content area (table + form)
        content_layout = QtWidgets.QHBoxLayout()
        content_layout.setSpacing(16)

        # Left: table (list of subsidios)
        self.table = QtWidgets.QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["ID", "Nombre", "Valor (%)", "ID Normativa"])
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.table.hideColumn(0)
        self.table.itemSelectionChanged.connect(self._on_table_select)
        self.table.setAlternatingRowColors(True)
        self.table.setStyleSheet("""
            QTableWidget { background: white; border: 1px solid #e6e9ee; }
            QHeaderView::section { background: qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 #E94E1B, stop:1 #ff7a43); color: white; padding: 8px; font-weight: bold; }
        """)
        content_layout.addWidget(self.table, 2)

        # Right: form
        form_frame = QtWidgets.QFrame()
        form_frame.setStyleSheet("QFrame { background: white; border: 1px solid #e6e9ee; border-radius: 8px; }")
        form_layout = QtWidgets.QFormLayout()
        form_layout.setContentsMargins(12, 12, 12, 12)
        form_layout.setSpacing(10)

        self.input_nombre = QtWidgets.QLineEdit()
        self.input_nombre.setPlaceholderText("Nombre del subsidio")
        self.input_valor = QtWidgets.QLineEdit()
        self.input_valor.setPlaceholderText("Ej: 12.5 o 12,5")
        self.input_idnorm = QtWidgets.QLineEdit()
        self.input_idnorm.setPlaceholderText("ID normativa (opcional)")

        form_layout.addRow("Nombre subsidio:", self.input_nombre)
        form_layout.addRow("Valor (% o decimal):", self.input_valor)
        form_layout.addRow("ID Normativa:", self.input_idnorm)

        form_frame.setLayout(form_layout)
        content_layout.addWidget(form_frame, 1)

        main_layout.addLayout(content_layout)

        # Bottom buttons
        btn_layout = QtWidgets.QHBoxLayout()
        btn_layout.setSpacing(10)

        self.btn_back = QtWidgets.QPushButton("Volver")
        self.btn_back.clicked.connect(lambda: self.back_requested.emit())
        self.btn_back.setCursor(QCursor(Qt.PointingHandCursor))
        self.btn_new = QtWidgets.QPushButton("Nuevo")
        self.btn_save = QtWidgets.QPushButton("Guardar")
        self.btn_delete = QtWidgets.QPushButton("Eliminar")
        self.btn_import = QtWidgets.QPushButton("Importar CSV")
        self.btn_export = QtWidgets.QPushButton("Exportar CSV")

        # Connect actions
        self.btn_new.clicked.connect(self._on_new)
        self.btn_save.clicked.connect(self._on_save)
        self.btn_delete.clicked.connect(self._on_delete)
        self.btn_import.clicked.connect(self._on_import)
        self.btn_export.clicked.connect(self._on_export)

        # Roles/permissions
        if not self._can_edit:
            self.input_nombre.setReadOnly(True)
            self.input_valor.setReadOnly(True)
            self.input_idnorm.setReadOnly(True)
            self.btn_new.setVisible(False)
            self.btn_save.setVisible(False)
            self.btn_delete.setVisible(False)
            self.btn_import.setVisible(False)
            self.btn_export.setToolTip("Exportar datos (solo lectura)")

        # Set properties for styled roles (used by stylesheet)
        self.btn_new.setProperty("role", "primary")
        self.btn_save.setProperty("role", "secondary")
        self.btn_delete.setProperty("role", "danger")
        self.btn_import.setProperty("role", "muted")
        self.btn_export.setProperty("role", "muted")
        self.btn_back.setProperty("role", "muted")

        btn_layout.addWidget(self.btn_back)
        btn_layout.addWidget(self.btn_new)
        btn_layout.addWidget(self.btn_save)
        btn_layout.addWidget(self.btn_delete)
        btn_layout.addStretch()
        btn_layout.addWidget(self.btn_import)
        btn_layout.addWidget(self.btn_export)

        main_layout.addLayout(btn_layout)

        # Tooltips
        self.btn_import.setToolTip("Importar subsidios desde CSV (actualiza o inserta filas)")
        self.btn_export.setToolTip("Exportar subsidios a CSV")

        # Apply styles
        self.apply_styles()

    def _add_toolbar(self, layout):
        toolbar_layout = QtWidgets.QHBoxLayout()
        toolbar_layout.setSpacing(10)

        btn_nueva = QtWidgets.QPushButton("➕ Nuevo Subsidio")
        btn_nueva.setCursor(QCursor(Qt.PointingHandCursor))
        btn_nueva.clicked.connect(self._on_new)
        btn_nueva.setProperty("role", "primary")
        btn_nueva.setStyleSheet("padding:8px 14px; border-radius:6px;")

        btn_refrescar = QtWidgets.QPushButton("🔄 Refrescar")
        btn_refrescar.setCursor(QCursor(Qt.PointingHandCursor))
        btn_refrescar.clicked.connect(self._load_data)
        btn_refrescar.setProperty("role", "secondary")
        btn_refrescar.setStyleSheet("padding:8px 12px; border-radius:6px;")

        btn_limpiar = QtWidgets.QPushButton("🗑️ Limpiar Todo")
        btn_limpiar.setCursor(QCursor(Qt.PointingHandCursor))
        btn_limpiar.clicked.connect(self._limpiar_todos_subsidios)
        btn_limpiar.setProperty("role", "danger")
        btn_limpiar.setVisible(self._is_admin)
        btn_limpiar.setStyleSheet("padding:8px 12px; border-radius:6px;")

        toolbar_layout.addWidget(btn_nueva)
        toolbar_layout.addWidget(btn_refrescar)
        toolbar_layout.addWidget(btn_limpiar)
        toolbar_layout.addStretch()

        # single info label (no duplication)
        self.label_info = QtWidgets.QLabel("")
        self.label_info.setFont(QFont("Segoe UI", 10))
        self.label_info.setStyleSheet("color: #2c3e50;")
        toolbar_layout.addWidget(self.label_info)

        layout.addLayout(toolbar_layout)

    def _load_data(self):
        try:
            self.table.setRowCount(0)
            if not self.service:
                raise RuntimeError("Servicio de subsidios no disponible")
            subsidios = self.service.list_all()
            for s in subsidios:
                self._add_row(s)
            total = len(subsidios)
            # update single label only
            self.label_info.setText(f"Total: {total} subsidios")
        except Exception as e:
            app_logger.error(f"Error cargando subsidios: {e}")
            QtWidgets.QMessageBox.warning(self, "Error", "No fue posible cargar los subsidios. Consulte logs.")

    def _add_row(self, s: Dict[str, Any]):
        r = self.table.rowCount()
        self.table.insertRow(r)
        self.table.setItem(r, 0, QtWidgets.QTableWidgetItem(s['id']))
        self.table.setItem(r, 1, QtWidgets.QTableWidgetItem(s['nombre_subsidio']))
        try:
            valor_pct = str((s['valor_porcentual'] * Decimal(100)).quantize(Decimal('0.01')))
        except Exception:
            valor_pct = str(s['valor_porcentual'])
        self.table.setItem(r, 2, QtWidgets.QTableWidgetItem(valor_pct))
        self.table.setItem(r, 3, QtWidgets.QTableWidgetItem(str(s.get('id_normativa') or "")))

    def _clear_form(self):
        self.input_nombre.clear()
        self.input_valor.clear()
        self.input_idnorm.clear()
        self.table.clearSelection()

    def _on_new(self):
        if not self._can_edit:
            QtWidgets.QMessageBox.warning(self, "Sin permiso", "No tiene permiso para crear subsidios.")
            return
        self._clear_form()

    def _on_table_select(self):
        items = self.table.selectedItems()
        if not items:
            return
        row = items[0].row()
        _id = self.table.item(row, 0).text()
        try:
            if not self.service:
                raise RuntimeError("Servicio de subsidios no disponible")
            s = self.service.get_by_id(_id)
            if s:
                self.input_nombre.setText(s['nombre_subsidio'])
                self.input_valor.setText(str((s['valor_porcentual'] * Decimal(100)).quantize(Decimal('0.01'))))
                self.input_idnorm.setText(str(s.get('id_normativa') or ""))
        except Exception as e:
            app_logger.error(f"Error obteniendo subsidio seleccionado: {e}")
            QtWidgets.QMessageBox.warning(self, "Error", "No fue posible obtener el subsidio seleccionado.")

    def _validate_and_build(self) -> Dict[str, Any]:
        nombre = self.input_nombre.text().strip()
        if not nombre:
            raise ValueError("El nombre del subsidio es obligatorio.")
        raw_val = self.input_valor.text().strip()
        if not raw_val:
            raise ValueError("El valor porcentual es obligatorio.")
        try:
            valor = Decimal(raw_val.replace(',', '.'))
        except InvalidOperation:
            raise ValueError("Valor porcentual no es un número válido.")
        if valor > 1:
            valor = (valor / Decimal(100)).quantize(Decimal('0.0001'))
        if valor < 0 or valor > 1:
            raise ValueError("El valor porcentual debe estar entre 0 y 1 (o 0-100).")
        id_norm_raw = self.input_idnorm.text().strip()
        id_norm_int = id_norm_raw if id_norm_raw else None

        items = self.table.selectedItems()
        if items:
            row = items[0].row()
            _id = self.table.item(row, 0).text()
            return {"id": _id, "nombre_subsidio": nombre, "valor_porcentual": valor, "id_normativa": id_norm_int}
        else:
            return {"nombre_subsidio": nombre, "valor_porcentual": valor, "id_normativa": id_norm_int}

    def _on_save(self):
        if not self._can_edit:
            QtWidgets.QMessageBox.warning(self, "Sin permiso", "No tiene permiso para guardar subsidios.")
            return
        try:
            data = self._validate_and_build()
            if not self.service:
                raise RuntimeError("Servicio de subsidios no disponible")
            self.service.add_or_update(data)
            QtWidgets.QMessageBox.information(self, "Éxito", "Subsidio guardado correctamente.")
            self._load_data()
            self._clear_form()
        except Exception as e:
            app_logger.error(f"Error guardando subsidio: {e}")
            QtWidgets.QMessageBox.warning(self, "Error", str(e))

    def _on_delete(self):
        if not self._can_edit:
            QtWidgets.QMessageBox.warning(self, "Sin permiso", "No tiene permiso para eliminar subsidios.")
            return
        items = self.table.selectedItems()
        if not items:
            QtWidgets.QMessageBox.warning(self, "Error", "Seleccione un subsidio para eliminar.")
            return
        row = items[0].row()
        _id = self.table.item(row, 0).text()
        reply = QtWidgets.QMessageBox.question(self, 'Confirmar', '¿Eliminar subsidio seleccionado?', QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
        if reply == QtWidgets.QMessageBox.Yes:
            try:
                if not self.service:
                    raise RuntimeError("Servicio de subsidios no disponible")
                self.service.delete(_id)
                self._load_data()
                self._clear_form()
            except Exception as e:
                app_logger.error(f"Error eliminando subsidio: {e}")
                QtWidgets.QMessageBox.warning(self, "Error", "No fue posible eliminar el subsidio. Consulte logs.")

    def _on_import(self):
        if not self._can_edit:
            QtWidgets.QMessageBox.warning(self, "Sin permiso", "No tiene permiso para importar subsidios.")
            return
        path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Importar CSV", os.path.expanduser("~"), "CSV Files (*.csv);;Todos (*)")
        if not path:
            return
        try:
            if not self.service:
                raise RuntimeError("Servicio de subsidios no disponible")
            stats = self.service.import_from_csv(path)
            QtWidgets.QMessageBox.information(self, "Importación completada",
                                              f"Filas: {stats['rows']}\nAñadidos: {stats['added']}\nActualizados: {stats['updated']}\nErrores: {stats['errors']}")
            self._load_data()
        except Exception as e:
            app_logger.error(f"Error importando CSV: {e}")
            QtWidgets.QMessageBox.warning(self, "Error", "No fue posible importar el CSV. Consulte logs.")

    def _on_export(self):
        path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Exportar CSV", os.path.expanduser("~"), "CSV Files (*.csv)")
        if not path:
            return
        try:
            if not self.service:
                raise RuntimeError("Servicio de subsidios no disponible")
            subsidios = self.service.list_all()
            with open(path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(['id', 'nombre_subsidio', 'valor_porcentual', 'id_normativa'])
                for s in subsidios:
                    writer.writerow([s['id'], s['nombre_subsidio'], str(s['valor_porcentual']), s.get('id_normativa') or ""])
            QtWidgets.QMessageBox.information(self, "Exportado", f"Exportado {len(subsidios)} subsidios a {path}")
        except Exception as e:
            app_logger.error(f"Error exportando CSV: {e}")
            QtWidgets.QMessageBox.warning(self, "Error", "No fue posible exportar el archivo.")

    def _limpiar_todos_subsidios(self):
        if not self._is_admin:
            QtWidgets.QMessageBox.warning(self, "Sin permiso", "Solo Administrador puede eliminar todos los subsidios.")
            return

        reply = QtWidgets.QMessageBox.warning(
            self,
            "⚠️ ADVERTENCIA CRÍTICA",
            "Está a punto de ELIMINAR TODOS los subsidios locales del sistema.\n\nEsta acción NO se puede deshacer.\n\n¿Desea continuar?",
            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
            QtWidgets.QMessageBox.No
        )
        if reply != QtWidgets.QMessageBox.Yes:
            return

        text, ok = QtWidgets.QInputDialog.getText(self, "CONFIRMACIÓN FINAL", "Escriba 'ELIMINAR TODO' para confirmar:")
        if not ok or text.strip() != "ELIMINAR TODO":
            QtWidgets.QMessageBox.information(self, "Cancelado", "Operación cancelada.")
            return

        try:
            stats = self.service.delete_all_subsidios()
            QtWidgets.QMessageBox.information(self, "Limpieza Completada", f"✅ Eliminados: {stats.get('deleted', 0)}\n❌ Errores: {stats.get('errors', 0)}")
            self._load_data()
        except Exception as e:
            app_logger.error(f"Error al eliminar todos los subsidios: {e}")
            QtWidgets.QMessageBox.critical(self, "Error", "Ocurrió un error al eliminar todos los subsidios. Consulte logs.")

    def apply_styles(self):
        """
        Estilos para la ventana de subsidios (toolbar, botones y tabla).
        """
        self.setStyleSheet("""
            QWidget { background-color: #f5f6fa; color: #2c3e50; }
            /* Toolbar roles */
            QPushButton[role="primary"] { background-color: #E94E1B; color: white; border-radius: 6px; padding: 8px 14px; }
            QPushButton[role="secondary"] { background-color: #3498db; color: white; border-radius: 6px; padding: 8px 12px; }
            QPushButton[role="success"] { background-color: #16a085; color: white; border-radius: 6px; padding: 8px 12px; }
            QPushButton[role="danger"] { background-color: #e74c3c; color: white; border-radius: 6px; padding: 8px 12px; }
            QPushButton[role="muted"] { background-color: #95a5a6; color: white; border-radius: 6px; padding: 6px 10px; }

            /* Table */
            QTableWidget { background-color: white; border: 1px solid #e6e9ee; }
            QHeaderView::section { background: qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 #E94E1B, stop:1 #ff7a43); color: white; padding: 8px; font-weight: bold; }
            QTableWidget::item { padding: 6px 8px; }

            /* Inputs and form */
            QLineEdit { padding: 8px; border: 1px solid #dfe6ee; border-radius: 6px; background: white; }
            QFrame { background: transparent; }

            /* Small labels */
            QLabel { color: #2c3e50; }
            QLabel[role="muted"] { color: #7f8c8d; font-size: 12px; }
        """)