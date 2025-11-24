from typing import List, Dict, Any
from datetime import datetime
from decimal import Decimal

from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QTableWidget, QTableWidgetItem, QHeaderView, QDateEdit,
    QComboBox, QFrame, QMessageBox, QAbstractItemView,
    QScrollArea, QDialog, QFormLayout, QLineEdit, QDoubleSpinBox,
    QGroupBox, QGridLayout, QListWidget, QListWidgetItem, QInputDialog
)
from PyQt5.QtCore import Qt, QDate, pyqtSignal
from PyQt5.QtGui import QFont, QColor, QCursor

from services.taxService import CalificacionTributariaService
from services.subsidyService import SubsidioService
from utils.logger import app_logger
from config.firebaseConfig import firebase_config


class CalificacionFormDialog(QDialog):
    """Diálogo para crear/editar calificaciones con selector de subsidios"""

    def __init__(self, parent, user_data: dict, modo="crear", calificacion: Dict[str, Any] = None):
        super().__init__(parent)
        self.user_data = user_data or {}
        self.user_id = self.user_data.get("_id")
        self.corredor_id = self.user_data.get("corredor_id") or self.user_id
        self.modo = modo
        self.calificacion = calificacion
        self.service = CalificacionTributariaService()

        try:
            self.subsidio_service = SubsidioService(corredor_id=self.corredor_id, user_id=self.user_id)
        except Exception as e:
            app_logger.error(f"No se pudo iniciar SubsidioService en formulario: {e}")
            self.subsidio_service = None

        self._subsidios_map: Dict[str, Dict[str, Any]] = {}
        self.init_ui()

        if modo == "editar" and calificacion:
            self.cargar_datos()

    def init_ui(self):
        titulo = "Nueva Calificación Tributaria" if self.modo == "crear" else "Editar Calificación Tributaria"
        self.setWindowTitle(titulo)
        self.setMinimumSize(900, 740)

        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(12)

        title_label = QLabel(titulo)
        title_label.setFont(QFont("Arial", 16, QFont.Bold))
        title_label.setStyleSheet("color: #2c3e50;")
        main_layout.addWidget(title_label)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)

        scroll_content = QWidget()
        scroll_layout = QVBoxLayout()
        scroll_layout.setSpacing(12)

        self.add_seccion_general(scroll_layout)
        self.add_seccion_factores(scroll_layout)
        self.add_validacion_suma(scroll_layout)

        scroll_content.setLayout(scroll_layout)
        scroll.setWidget(scroll_content)
        main_layout.addWidget(scroll)

        self.add_buttons(main_layout)

        self.setLayout(main_layout)
        self.apply_styles()

        if self.modo == "editar" and self.calificacion and not self.calificacion.get("esLocal", False):
            self.establecer_solo_lectura()

    def add_seccion_general(self, layout):
        group = QGroupBox("📝 Información General")
        group.setFont(QFont("Arial", 11, QFont.Bold))

        form_layout = QFormLayout()
        form_layout.setSpacing(8)

        self.input_cliente = QLineEdit()
        self.input_cliente.setPlaceholderText("Ej: 12345678-9")
        form_layout.addRow("Cliente RUT: *", self.input_cliente)

        self.input_fecha = QDateEdit()
        self.input_fecha.setCalendarPopup(True)
        self.input_fecha.setDate(QDate.currentDate())
        self.input_fecha.setDisplayFormat("dd/MM/yyyy")
        form_layout.addRow("Fecha Declaración: *", self.input_fecha)

        self.combo_tipo_impuesto = QComboBox()
        self.combo_tipo_impuesto.addItems(["IVA", "Renta", "Importación", "Exportación", "Otro"])
        form_layout.addRow("Tipo de Impuesto: *", self.combo_tipo_impuesto)

        self.combo_pais = QComboBox()
        self.combo_pais.addItems(["Chile", "Perú", "Colombia"])
        form_layout.addRow("País: *", self.combo_pais)

        self.input_monto = QDoubleSpinBox()
        self.input_monto.setRange(0.01, 999999999999.99)
        self.input_monto.setDecimals(2)
        self.input_monto.setPrefix("$ ")
        self.input_monto.setGroupSeparatorShown(True)
        self.input_monto.valueChanged.connect(self._on_monto_changed)
        form_layout.addRow("Monto Declarado: *", self.input_monto)

        # Subsidios selector
        self.subsidios_list = QListWidget()
        self.subsidios_list.setSelectionMode(QAbstractItemView.NoSelection)
        self.subsidios_list.itemChanged.connect(self._on_subsidio_item_changed)
        self.label_monto_ajustado = QLabel("Monto con Subsidios: $ 0.00")
        self.label_monto_ajustado.setFont(QFont("Arial", 10, QFont.Bold))
        self.label_monto_ajustado.setStyleSheet("color: #2c3e50;")

        form_layout.addRow("Subsidios aplicables (selecciona los que correspondan):", self.subsidios_list)
        form_layout.addRow("", self.label_monto_ajustado)

        group.setLayout(form_layout)
        layout.addWidget(group)

        self._load_subsidios_into_list()

    def _load_subsidios_into_list(self):
        self.subsidios_list.clear()
        self._subsidios_map.clear()
        if not self.subsidio_service:
            return
        try:
            subs = self.subsidio_service.list_all()
            for s in subs:
                sid = s.get("id")
                name = s.get("nombre_subsidio", str(sid))
                # valor_porcentual stored as Decimal in subsidyService
                try:
                    pct = (s.get("valor_porcentual") * Decimal(100)).quantize(Decimal("0.01"))
                except Exception:
                    pct = s.get("valor_porcentual")
                item = QListWidgetItem(f"{name} ({pct}% )")
                item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
                item.setCheckState(Qt.Unchecked)
                item.setData(Qt.UserRole, sid)
                self.subsidios_list.addItem(item)
                if sid is not None:
                    self._subsidios_map[str(sid)] = s
        except Exception as e:
            app_logger.error(f"Error cargando subsidios en formulario: {e}")

    def _on_subsidio_item_changed(self, item: QListWidgetItem):
        self._update_preview_monto()

    def _on_monto_changed(self, _):
        self._update_preview_monto()

    def add_seccion_factores(self, layout):
        group = QGroupBox("📊 Factores (19 valores entre 0 y 1)")
        group.setFont(QFont("Arial", 11, QFont.Bold))

        grid_layout = QGridLayout()
        grid_layout.setSpacing(8)

        self.factor_inputs: List[QDoubleSpinBox] = []
        for i in range(19):
            factor_num = i + 1
            label = QLabel(f"Factor {factor_num}:")
            label.setFont(QFont("Arial", 9))

            spinbox = QDoubleSpinBox()
            spinbox.setRange(0.0, 1.0)
            spinbox.setDecimals(4)
            spinbox.setSingleStep(0.01)
            spinbox.setValue(0.0)
            spinbox.setMinimumWidth(100)

            if 8 <= factor_num <= 19:
                spinbox.setProperty("highlighted", True)

            spinbox.valueChanged.connect(self.actualizar_suma_factores)

            self.factor_inputs.append(spinbox)

            row = i // 4
            col = (i % 4) * 2
            grid_layout.addWidget(label, row, col)
            grid_layout.addWidget(spinbox, row, col + 1)

        group.setLayout(grid_layout)
        layout.addWidget(group)

    def add_validacion_suma(self, layout):
        validacion_frame = QFrame()
        validacion_frame.setStyleSheet("""
            QFrame {
                background-color: #f8f9fa;
                border: 2px solid #dee2e6;
                border-radius: 5px;
                padding: 10px;
            }
        """)
        validacion_layout = QHBoxLayout()
        icon_label = QLabel("⚠️")
        icon_label.setFont(QFont("Arial", 16))
        validacion_layout.addWidget(icon_label)

        self.label_suma = QLabel("Suma Factores 8-19: 0.0000")
        self.label_suma.setFont(QFont("Arial", 12, QFont.Bold))
        validacion_layout.addWidget(self.label_suma)

        validacion_layout.addStretch()

        self.label_estado_suma = QLabel("✅ Válido")
        self.label_estado_suma.setFont(QFont("Arial", 11, QFont.Bold))
        validacion_layout.addWidget(self.label_estado_suma)

        validacion_frame.setLayout(validacion_layout)
        layout.addWidget(validacion_frame)

    def add_buttons(self, layout):
        button_layout = QHBoxLayout()
        button_layout.addStretch()

        btn_cancelar = QPushButton("❌ Cancelar")
        btn_cancelar.setMinimumSize(120, 40)
        btn_cancelar.clicked.connect(self.reject)
        button_layout.addWidget(btn_cancelar)

        self.btn_guardar = QPushButton("💾 Guardar")
        self.btn_guardar.setMinimumSize(120, 40)
        self.btn_guardar.clicked.connect(self.guardar)
        button_layout.addWidget(self.btn_guardar)

        layout.addLayout(button_layout)

    def actualizar_suma_factores(self):
        suma = sum(self.factor_inputs[i].value() for i in range(7, 19))
        self.label_suma.setText(f"Suma Factores 8-19: {suma:.4f}")
        if suma > 1.0:
            self.label_suma.setStyleSheet("color: #e74c3c;")
            self.label_estado_suma.setText("❌ Inválido (> 1.0)")
            self.label_estado_suma.setStyleSheet("color: #e74c3c;")
            self.btn_guardar.setEnabled(False)
        else:
            self.label_suma.setStyleSheet("color: #27ae60;")
            self.label_estado_suma.setText("✅ Válido")
            self.label_estado_suma.setStyleSheet("color: #27ae60;")
            self.btn_guardar.setEnabled(True)

    def cargar_datos(self):
        if not self.calificacion:
            return

        self.input_cliente.setText(self.calificacion.get("clienteId", ""))

        fecha_str = self.calificacion.get("fechaDeclaracion", "")
        if fecha_str:
            try:
                fecha_obj = datetime.strptime(fecha_str, "%Y-%m-%d")
                qdate = QDate(fecha_obj.year, fecha_obj.month, fecha_obj.day)
                self.input_fecha.setDate(qdate)
            except Exception:
                pass

        tipo = self.calificacion.get("tipoImpuesto", "")
        index = self.combo_tipo_impuesto.findText(tipo)
        if index >= 0:
            self.combo_tipo_impuesto.setCurrentIndex(index)

        pais = self.calificacion.get("pais", "")
        index = self.combo_pais.findText(pais)
        if index >= 0:
            self.combo_pais.setCurrentIndex(index)

        self.input_monto.setValue(self.calificacion.get("montoDeclarado", 0))

        factores = self.calificacion.get("factores", {})
        for i in range(19):
            valor = factores.get(f"factor_{i+1}", 0)
            self.factor_inputs[i].setValue(valor)

        # preseleccionar subsidios si existen
        subs_aplicados = self.calificacion.get("subsidiosAplicados", []) or []
        seleccion_ids = set()
        for s in subs_aplicados:
            if isinstance(s, dict):
                sid = s.get("id") or s.get("subsidio_id")
            else:
                sid = s
            if sid:
                seleccion_ids.add(str(sid))

        for i in range(self.subsidios_list.count()):
            item = self.subsidios_list.item(i)
            sid = str(item.data(Qt.UserRole))
            if sid in seleccion_ids:
                item.setCheckState(Qt.Checked)
            else:
                item.setCheckState(Qt.Unchecked)

        self.actualizar_suma_factores()
        self._update_preview_monto()

    def establecer_solo_lectura(self):
        self.input_cliente.setEnabled(False)
        self.input_fecha.setEnabled(False)
        self.combo_tipo_impuesto.setEnabled(False)
        self.combo_pais.setEnabled(False)
        self.input_monto.setEnabled(False)
        self.subsidios_list.setEnabled(False)
        for spinbox in self.factor_inputs:
            spinbox.setEnabled(False)
        self.btn_guardar.setEnabled(False)
        mensaje = QLabel("⚠️ SOLO LECTURA: Esta calificación es de la bolsa y no puede modificarse")
        mensaje.setFont(QFont("Arial", 10, QFont.Bold))
        mensaje.setStyleSheet("""
            QLabel {
                background-color: #fff3cd;
                color: #856404;
                border: 2px solid #ffc107;
                border-radius: 5px;
                padding: 10px;
            }
        """)
        mensaje.setAlignment(Qt.AlignCenter)
        self.layout().insertWidget(1, mensaje)

    def _gather_selected_subsidios_ids(self) -> List[str]:
        ids: List[str] = []
        for i in range(self.subsidios_list.count()):
            item = self.subsidios_list.item(i)
            if item.checkState() == Qt.Checked:
                sid = item.data(Qt.UserRole)
                ids.append(str(sid))
        return ids

    def _update_preview_monto(self):
        monto = Decimal(str(self.input_monto.value()))
        selected_ids = self._gather_selected_subsidios_ids()
        for sid in selected_ids:
            try:
                s = self._subsidios_map.get(str(sid))
                if not s and self.subsidio_service:
                    s = self.subsidio_service.get_by_id(sid)
                if not s:
                    continue
                vp = s.get("valor_porcentual", Decimal("0"))
                if not isinstance(vp, Decimal):
                    vp = Decimal(str(vp))
                monto = (monto * (Decimal("1") - vp)).quantize(Decimal("0.01"))
            except Exception as e:
                app_logger.error(f"Error calculando preview por subsidio {sid}: {e}")
                continue
        try:
            display = f"$ {float(monto):,.2f}"
        except Exception:
            display = f"{monto}"
        self.label_monto_ajustado.setText(f"Monto con Subsidios: {display}")

    def guardar(self):
        try:
            fecha_date = self.input_fecha.date().toPyDate()
            fecha_declaracion = datetime.combine(fecha_date, datetime.min.time())

            datos: Dict[str, Any] = {
                "cliente_id": self.input_cliente.text().strip(),
                "fecha_declaracion": fecha_declaracion,
                "tipo_impuesto": self.combo_tipo_impuesto.currentText(),
                "pais": self.combo_pais.currentText(),
                "monto_declarado": self.input_monto.value(),
            }

            factores = [spinbox.value() for spinbox in self.factor_inputs]
            datos["factores"] = factores

            # pasar subsidios seleccionados
            datos["subsidios_aplicados"] = self._gather_selected_subsidios_ids()

            if not datos["cliente_id"]:
                QMessageBox.warning(self, "Validación", "El RUT del cliente es obligatorio")
                return

            usuario_id = self.user_data.get("_id")
            user_rol = self.user_data.get("rol", "cliente")

            if self.modo == "crear":
                result = self.service.crear_calificacion(datos, usuario_id)
                if not result["success"] and result.get("conflicto", False):
                    dato_oficial = result.get("dato_oficial", {})
                    QMessageBox.warning(
                        self,
                        "⚠️ Conflicto con Dato Oficial",
                        f"Ya existe una calificación OFICIAL de bolsa:\n\n"
                        f"• Monto: ${dato_oficial.get('monto', 0):,.2f}\n"
                        f"• Fecha: {dato_oficial.get('fecha', 'N/A')}\n\n"
                        "No se puede crear una calificación local duplicada.\nCambie la fecha o el tipo de impuesto.",
                        QMessageBox.Ok
                    )
                    return
            else:
                result = self.service.actualizar_calificacion(
                    self.calificacion["_id"],
                    datos,
                    usuario_id,
                    user_rol
                )

            if result["success"]:
                self.accept()
            else:
                QMessageBox.warning(self, "Error", result.get("message", "Error desconocido"))

        except Exception as e:
            app_logger.error(f"Error al guardar: {str(e)}")
            QMessageBox.critical(self, "Error", f"Error al guardar: {str(e)}")

    def apply_styles(self):
        self.setStyleSheet("""
            /* Fondo general del diálogo */
            QDialog {
                background-color: #f2f4f7;
            }

            /* Card principal */
            QFrame#card_main {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                            stop:0 #ffffff, stop:0.6 #fbfbfb, stop:1 #f7f9fb);
                border: 1px solid #e3e7ee;
                border-radius: 6px;
                padding: 6px;
            }

            /* GroupBox estilo 'panel' */
            QGroupBox {
                font-weight: bold;
                color: #2c3e50;
                border: 1px solid #e6e9ee;
                border-radius: 6px;
                margin-top: 8px;
                background-color: #ffffff;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 4px 8px;
                color: #2c3e50;
                background-color: rgba(0,0,0,0);
            }

            /* Inputs (alineados y con borde suave) */
            QLineEdit, QComboBox, QDateEdit, QDoubleSpinBox {
                padding: 8px 10px;
                border: 1px solid #d7dfe8;
                border-radius: 4px;
                background-color: #ffffff;
                color: #2c3e50;
            }

            /* Botones primarios / secundarios */
            QPushButton#btn_guardar {
                background-color: #e76f23; /* naranja */
                color: white;
                padding: 10px 14px;
                border-radius: 6px;
                border: none;
                font-weight: 600;
            }
            QPushButton#btn_guardar:hover {
                background-color: #d15f12;
            }
            QPushButton#btn_cancelar {
                background-color: #95a5a6;
                color: white;
                padding: 10px 14px;
                border-radius: 6px;
                border: none;
            }
            QPushButton#btn_cancelar:hover {
                background-color: #7f8c8d;
            }

            /* Lista de subsidios */
            QListWidget {
                background-color: #fff;
                border: 1px solid #e6e9ee;
                border-radius: 6px;
                padding: 6px;
                min-height: 80px;
            }
            QListWidget::item {
                padding: 8px;
                margin: 2px 0;
            }

            /* Banda de validación */
            QFrame#validation_frame {
                background-color: #fff9e8;
                border: 1px solid #ffe6cc;
                border-radius: 6px;
                padding: 8px;
            }

            /* Labels destacados */
            QLabel {
                color: #2c3e50;
            }
            QLabel[role="muted"] {
                color: #7f8c8d;
                font-size: 12px;
            }
        """)


class GestionCalificacionesContent(QWidget):
    """Contenido de gestión de calificaciones (widget embebible en MainWindow)"""

    back_requested = pyqtSignal()

    def __init__(self, user_data: dict, parent=None):
        super().__init__(parent)
        self.user_data = user_data or {}
        self.user_rol = self.user_data.get("rol", "cliente")
        self.service = CalificacionTributariaService()
        self.calificaciones: List[Dict[str, Any]] = []
        
        # Cache para RUTs de clientes
        self.clientes_cache: Dict[str, str] = {}  # {cliente_id: rut}
        self.db = firebase_config.get_firestore_client()

        self.init_ui()
        self.refrescar_tabla()

    def init_ui(self):
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll_area.setStyleSheet("QScrollArea { border: none; background-color: #f5f6fa; }")

        content_widget = QWidget()
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(40, 30, 40, 30)
        main_layout.setSpacing(20)

        # Header con botón volver
        self.add_header(main_layout)

        # Toolbar
        self.add_toolbar(main_layout)

        # Filtros
        self.add_filters(main_layout)

        # Tabla
        self.add_table(main_layout)

        # Footer
        self.add_footer(main_layout)

        content_widget.setLayout(main_layout)
        scroll_area.setWidget(content_widget)

        widget_layout = QVBoxLayout()
        widget_layout.setContentsMargins(0, 0, 0, 0)
        widget_layout.addWidget(scroll_area)
        self.setLayout(widget_layout)

        # aplicar estilos
        self.apply_styles()

    def add_header(self, layout):
        header_layout = QHBoxLayout()

        back_button = QPushButton("←")
        back_button.setFont(QFont("Arial", 10))
        back_button.setCursor(QCursor(Qt.PointingHandCursor))
        back_button.clicked.connect(self.back_requested.emit)
        back_button.setStyleSheet("""
            QPushButton {
                background-color: transparent;
                border: none;
                color: #3498db;
                padding: 5px;
            }
            QPushButton:hover { color: #2980b9; }
        """)
        header_layout.addWidget(back_button)

        title = QLabel("Gestión de Calificaciones Tributarias")
        title.setFont(QFont("Arial", 18, QFont.Bold))
        title.setStyleSheet("color: #2c3e50;")
        header_layout.addWidget(title)

        header_layout.addStretch()
        layout.addLayout(header_layout)

    def add_toolbar(self, layout):
        toolbar_layout = QHBoxLayout()
        toolbar_layout.setSpacing(10)

        btn_nueva = QPushButton("➕ Nueva Calificación")
        btn_nueva.setFont(QFont("Arial", 10, QFont.Bold))
        btn_nueva.setMinimumHeight(40)
        btn_nueva.setCursor(QCursor(Qt.PointingHandCursor))
        btn_nueva.clicked.connect(self.abrir_formulario_crear)
        btn_nueva.setStyleSheet("""
            QPushButton {
                background-color: #E94E1B;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 10px 20px;
            }
            QPushButton:hover { background-color: #d64419; }
        """)
        toolbar_layout.addWidget(btn_nueva)

        btn_refrescar = QPushButton("🔄 Refrescar")
        btn_refrescar.setFont(QFont("Arial", 10))
        btn_refrescar.setMinimumHeight(40)
        btn_refrescar.setCursor(QCursor(Qt.PointingHandCursor))
        btn_refrescar.clicked.connect(self.refrescar_tabla)
        btn_refrescar.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 10px 20px;
            }
            QPushButton:hover { background-color: #2980b9; }
        """)
        toolbar_layout.addWidget(btn_refrescar)

        # Botón "Eliminar Todo" (solo visible para administradores)
        if self.user_rol == "administrador":
            btn_eliminar_todo = QPushButton("🗑️ Eliminar Todo")
            btn_eliminar_todo.setFont(QFont("Arial", 10, QFont.Bold))
            btn_eliminar_todo.setMinimumHeight(40)
            btn_eliminar_todo.setCursor(QCursor(Qt.PointingHandCursor))
            btn_eliminar_todo.clicked.connect(self.eliminar_todas_calificaciones)
            btn_eliminar_todo.setStyleSheet("""
                QPushButton {
                    background-color: #e74c3c;
                    color: white;
                    border: none;
                    border-radius: 5px;
                    padding: 10px 20px;
                }
                QPushButton:hover { background-color: #c0392b; }
            """)
            toolbar_layout.addWidget(btn_eliminar_todo)

        toolbar_layout.addStretch()

        self.label_contador = QLabel("Total: 0 calificaciones")
        self.label_contador.setFont(QFont("Arial", 10, QFont.Bold))
        self.label_contador.setStyleSheet("color: #2c3e50;")
        toolbar_layout.addWidget(self.label_contador)

        layout.addLayout(toolbar_layout)

    def add_filters(self, layout):
        filter_frame = QFrame()
        filter_frame.setStyleSheet("""
            QFrame {
                background-color: #f8f9fa;
                border: 1px solid #dee2e6;
                border-radius: 5px;
                padding: 10px;
            }
        """)
        filter_layout = QHBoxLayout()
        filter_layout.setSpacing(15)

        filter_layout.addWidget(QLabel("Fecha desde:"))
        self.date_desde = QDateEdit()
        self.date_desde.setCalendarPopup(True)
        self.date_desde.setDate(QDate.currentDate().addMonths(-6))
        self.date_desde.setDisplayFormat("dd/MM/yyyy")
        filter_layout.addWidget(self.date_desde)

        filter_layout.addWidget(QLabel("Fecha hasta:"))
        self.date_hasta = QDateEdit()
        self.date_hasta.setCalendarPopup(True)
        self.date_hasta.setDate(QDate.currentDate())
        self.date_hasta.setDisplayFormat("dd/MM/yyyy")
        filter_layout.addWidget(self.date_hasta)

        filter_layout.addWidget(QLabel("Tipo:"))
        self.combo_tipo = QComboBox()
        self.combo_tipo.addItems(["Todos", "IVA", "Renta", "Importación", "Exportación", "Otro"])
        filter_layout.addWidget(self.combo_tipo)

        filter_layout.addWidget(QLabel("País:"))
        self.combo_pais = QComboBox()
        self.combo_pais.addItems(["Todos", "Chile", "Perú", "Colombia"])
        filter_layout.addWidget(self.combo_pais)

        filter_layout.addStretch()

        btn_buscar = QPushButton("🔍 Buscar")
        btn_buscar.setCursor(QCursor(Qt.PointingHandCursor))
        btn_buscar.clicked.connect(self.aplicar_filtros)
        btn_buscar.setStyleSheet("""
            QPushButton {
                background-color: #27ae60;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 8px 16px;
            }
            QPushButton:hover { background-color: #229954; }
        """)
        filter_layout.addWidget(btn_buscar)

        btn_limpiar = QPushButton("🗑️ Limpiar")
        btn_limpiar.setCursor(QCursor(Qt.PointingHandCursor))
        btn_limpiar.clicked.connect(self.limpiar_filtros)
        btn_limpiar.setStyleSheet("""
            QPushButton {
                background-color: #95a5a6;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 8px 16px;
            }
            QPushButton:hover { background-color: #7f8c8d; }
        """)
        filter_layout.addWidget(btn_limpiar)

        filter_frame.setLayout(filter_layout)
        layout.addWidget(filter_frame)

    def add_table(self, layout):
        self.table = QTableWidget()
        self.table.setColumnCount(10)
        self.table.setHorizontalHeaderLabels([
            "ID", "Cliente RUT", "Fecha", "Tipo", "País",
            "Monto", "Suma 8-19", "Estado", "Editar", "Eliminar"
        ])

        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)

        header = self.table.horizontalHeader()
        for i in range(10):
            header.setSectionResizeMode(i, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.Stretch)

        layout.addWidget(self.table)

    def add_footer(self, layout):
        if self.user_rol == "administrador":
            footer_text = "💡 Admin: Puedes ver, editar y eliminar TODAS las calificaciones (locales y de bolsa)"
        else:
            footer_text = "💡 Tip: Solo puedes editar/eliminar tus calificaciones locales. Los datos de bolsa son de solo lectura."

        footer = QLabel(footer_text)
        footer.setFont(QFont("Arial", 8))
        footer.setStyleSheet("color: #7f8c8d;")
        layout.addWidget(footer)

    def actualizar_tabla(self, calificaciones: List[Dict[str, Any]]):
        self.calificaciones = calificaciones
        self.table.setRowCount(len(calificaciones))

        for row, cal in enumerate(calificaciones):
            # Columna 0: ID (primeros 8 caracteres - codificado)
            item_id = QTableWidgetItem(cal["_id"][:8])
            self.table.setItem(row, 0, item_id)

            # Columna 1: Cliente RUT (NO codificado)
            cliente_id = cal.get("clienteId", "")
            cliente_rut = self._obtener_rut_cliente(cliente_id)
            item_cliente = QTableWidgetItem(cliente_rut)
            self.table.setItem(row, 1, item_cliente)

            # Columna 2: Fecha
            fecha_str = cal.get("fechaDeclaracion", "")
            self.table.setItem(row, 2, QTableWidgetItem(fecha_str))

            # Columna 3: Tipo de impuesto
            self.table.setItem(row, 3, QTableWidgetItem(cal.get("tipoImpuesto", "")))
            
            # Columna 4: País
            self.table.setItem(row, 4, QTableWidgetItem(cal.get("pais", "")))

            # Columna 5: Monto declarado
            monto = cal.get("montoDeclarado", 0)
            item_monto = QTableWidgetItem(f"${monto:,.2f}")
            item_monto.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.table.setItem(row, 5, item_monto)

            # Columna 6: Suma factores 8-19
            factores = cal.get("factores", {})
            suma = sum(factores.get(f"factor_{i}", 0) for i in range(8, 20))
            item_suma = QTableWidgetItem(f"{suma:.4f}")
            item_suma.setTextAlignment(Qt.AlignCenter)
            if suma > 1.0:
                item_suma.setBackground(QColor(255, 200, 200))
                item_suma.setForeground(QColor(200, 0, 0))
            else:
                item_suma.setBackground(QColor(200, 255, 200))
                item_suma.setForeground(QColor(0, 150, 0))
            self.table.setItem(row, 6, item_suma)

            # Columna 7: Estado (Local/Bolsa)
            es_local = cal.get("esLocal", False)
            es_admin = self.user_rol == "administrador"
            estado = "Local" if es_local else "Bolsa"
            item_estado = QTableWidgetItem(estado)
            item_estado.setTextAlignment(Qt.AlignCenter)
            if es_local:
                item_estado.setBackground(QColor(200, 230, 255))
                item_estado.setForeground(QColor(0, 100, 200))
            else:
                item_estado.setBackground(QColor(230, 230, 230))
                item_estado.setForeground(QColor(100, 100, 100))
            self.table.setItem(row, 7, item_estado)

            # Columna 8: Botón Editar
            btn_editar = QPushButton("✏️")
            btn_editar.setEnabled(es_admin or es_local)
            btn_editar.clicked.connect(lambda checked, c=cal: self.abrir_formulario_editar(c))
            self.table.setCellWidget(row, 8, btn_editar)

            # Columna 9: Botón Eliminar
            btn_eliminar = QPushButton("🗑️")
            btn_eliminar.setEnabled(es_admin or es_local)
            btn_eliminar.clicked.connect(lambda checked, c=cal: self.eliminar_calificacion(c))
            self.table.setCellWidget(row, 9, btn_eliminar)

        self.label_contador.setText(f"Total: {len(calificaciones)} calificaciones")

    def _obtener_rut_cliente(self, cliente_id: str) -> str:
        """
        Obtiene el RUT del cliente desde Firestore.
        Usa caché para evitar consultas repetidas.
        
        Args:
            cliente_id (str): ID del documento del cliente en Firestore
            
        Returns:
            str: RUT del cliente o el cliente_id si no se encuentra
        """
        # ✅ NUEVO: Manejar caso donde cliente_id es una lista o None
        if not cliente_id:
            return "N/A"
        
        # Si cliente_id es una lista, tomar el primer elemento
        if isinstance(cliente_id, list):
            if len(cliente_id) > 0:
                cliente_id = str(cliente_id[0])
            else:
                return "N/A"
        
        # Convertir a string por si acaso
        cliente_id = str(cliente_id)
        
        # Verificar si ya está en caché
        if cliente_id in self.clientes_cache:
            return self.clientes_cache[cliente_id]
        
        # Si no está en caché, buscar en Firestore
        try:
            cliente_doc = self.db.collection("usuarios").document(cliente_id).get()
            
            if cliente_doc.exists:
                cliente_data = cliente_doc.to_dict()
                rut = cliente_data.get("rut", cliente_id)
                
                # Guardar en caché
                self.clientes_cache[cliente_id] = rut
                
                app_logger.debug(f"RUT obtenido para cliente {cliente_id}: {rut}")
                return rut
            else:
                app_logger.warning(f"Cliente {cliente_id} no encontrado en Firestore")
                # Guardar en caché para evitar consultas repetidas
                self.clientes_cache[cliente_id] = cliente_id
                return cliente_id
                
        except Exception as e:
            app_logger.error(f"Error al obtener RUT del cliente {cliente_id}: {e}")
            # Guardar en caché para evitar consultas repetidas
            self.clientes_cache[cliente_id] = cliente_id
            return cliente_id
        
    def abrir_formulario_crear(self):
        dialog = CalificacionFormDialog(self, self.user_data, modo="crear")
        if dialog.exec_():
            self.refrescar_tabla()
            QMessageBox.information(self, "Éxito", "Calificación creada exitosamente")

    def abrir_formulario_editar(self, calificacion: Dict[str, Any]):
        es_local = calificacion.get("esLocal", False)
        es_admin = self.user_rol == "administrador"
        if not es_admin and not es_local:
            QMessageBox.warning(self, "Acción no permitida", "No se puede editar una calificación de bolsa.")
            return
        dialog = CalificacionFormDialog(self, self.user_data, modo="editar", calificacion=calificacion)
        if dialog.exec_():
            self.refrescar_tabla()
            QMessageBox.information(self, "Éxito", "Calificación actualizada exitosamente")

    def eliminar_calificacion(self, calificacion: Dict[str, Any]):
        es_local = calificacion.get("esLocal", False)
        es_admin = self.user_rol == "administrador"
        if not es_admin and not es_local:
            QMessageBox.warning(self, "Acción no permitida", "No se puede eliminar una calificación de bolsa.")
            return

        if es_admin and not es_local:
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Warning)
            msg.setWindowTitle("⚠️ ADVERTENCIA: Eliminar Dato de Bolsa")
            msg.setText("¿Está seguro de eliminar este dato de la bolsa?")
            msg.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
            msg.setDefaultButton(QMessageBox.No)
            msg.setStyleSheet("""
                QMessageBox {
                    background-color: white;
                }
                QMessageBox QLabel {
                    color: #2c3e50;
                    min-width: 300px;
                }
                QMessageBox QPushButton {
                    background-color: #3498db;
                    color: white;
                    padding: 8px 16px;
                    border-radius: 4px;
                    min-width: 80px;
                }
            """)
            reply = msg.exec_()
        else:
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Question)
            msg.setWindowTitle("Confirmar Eliminación")
            msg.setText("¿Está seguro de eliminar esta calificación?")
            msg.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
            msg.setDefaultButton(QMessageBox.No)
            msg.setStyleSheet("""
                QMessageBox {
                    background-color: white;
                }
                QMessageBox QLabel {
                    color: #2c3e50;
                    min-width: 300px;
                }
                QMessageBox QPushButton {
                    background-color: #3498db;
                    color: white;
                    padding: 8px 16px;
                    border-radius: 4px;
                    min-width: 80px;
                }
            """)
            reply = msg.exec_()

        if reply == QMessageBox.Yes:
            result = self.service.eliminar_calificacion(
                calificacion["_id"],
                self.user_data.get("_id"),
                self.user_rol
            )
            if result.get("success"):
                self.refrescar_tabla()
                QMessageBox.information(self, "Éxito", "Calificación eliminada exitosamente")
            else:
                QMessageBox.warning(self, "Error", result.get("message", "Error desconocido"))

    def eliminar_todas_calificaciones(self):
        """
        Elimina TODAS las calificaciones del sistema.
        Solo accesible para administradores con doble confirmación.
        """
        if self.user_rol != "administrador":
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Warning)
            msg.setWindowTitle("Sin permiso")
            msg.setText("Solo los administradores pueden eliminar todas las calificaciones.")
            msg.setStyleSheet("""
                QMessageBox {
                    background-color: white;
                }
                QMessageBox QLabel {
                    color: #2c3e50;
                    min-width: 300px;
                }
            """)
            msg.exec_()
            return

        # Primera confirmación
        msg = QMessageBox(self)
        msg.setIcon(QMessageBox.Warning)
        msg.setWindowTitle("⚠️ ADVERTENCIA CRÍTICA")
        msg.setText(
            "Está a punto de ELIMINAR TODAS las calificaciones tributarias del sistema.\n\n"
            "Esto incluye:\n"
            "• Calificaciones de BOLSA (datos oficiales)\n"
            "• Calificaciones LOCALES (de todos los corredores)\n\n"
            "Esta acción NO se puede deshacer.\n\n"
            "¿Desea continuar?"
        )
        msg.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        msg.setDefaultButton(QMessageBox.No)
        msg.setStyleSheet("""
            QMessageBox {
                background-color: white;
            }
            QMessageBox QLabel {
                color: #2c3e50;
                min-width: 400px;
            }
            QMessageBox QPushButton {
                background-color: #3498db;
                color: white;
                padding: 8px 16px;
                border-radius: 4px;
                min-width: 80px;
            }
        """)
        reply = msg.exec_()

        if reply != QMessageBox.Yes:
            return

        # Segunda confirmación: Diálogo personalizado
        dialog = QDialog(self)
        dialog.setWindowTitle("CONFIRMACIÓN FINAL")
        dialog.setModal(True)
        dialog.setFixedSize(500, 200)
        dialog.setStyleSheet("""
            QDialog {
                background-color: white;
            }
            QLabel {
                color: #2c3e50;
                font-size: 13px;
            }
            QLineEdit {
                padding: 8px;
                border: 2px solid #3498db;
                border-radius: 4px;
                background-color: white;
                color: #2c3e50;
                font-size: 13px;
            }
            QPushButton {
                padding: 8px 16px;
                border-radius: 4px;
                font-weight: 600;
                min-width: 100px;
            }
            QPushButton#btn_confirmar {
                background-color: #e74c3c;
                color: white;
                border: none;
            }
            QPushButton#btn_confirmar:hover {
                background-color: #c0392b;
            }
            QPushButton#btn_cancelar {
                background-color: #95a5a6;
                color: white;
                border: none;
            }
            QPushButton#btn_cancelar:hover {
                background-color: #7f8c8d;
            }
        """)
        
        # Layout principal
        layout = QVBoxLayout()
        layout.setSpacing(15)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # Mensaje
        label = QLabel(
            "Para confirmar esta acción destructiva,\n"
            "escriba exactamente:\n\n"
            "ELIMINAR TODO\n\n"
            "(distingue mayúsculas y minúsculas)"
        )
        label.setAlignment(Qt.AlignCenter)
        label.setWordWrap(True)
        layout.addWidget(label)
        
        # Input
        input_text = QLineEdit()
        input_text.setPlaceholderText("Escriba aquí...")
        layout.addWidget(input_text)
        
        # Botones
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        
        btn_cancelar = QPushButton("Cancelar")
        btn_cancelar.setObjectName("btn_cancelar")
        btn_cancelar.clicked.connect(dialog.reject)
        button_layout.addWidget(btn_cancelar)
        
        btn_confirmar = QPushButton("Confirmar")
        btn_confirmar.setObjectName("btn_confirmar")
        btn_confirmar.clicked.connect(dialog.accept)
        button_layout.addWidget(btn_confirmar)
        
        layout.addLayout(button_layout)
        dialog.setLayout(layout)
        
        # Mostrar diálogo
        result = dialog.exec_()
        text = input_text.text().strip()
        
        if result != QDialog.Accepted or text != "ELIMINAR TODO":
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Information)
            msg.setWindowTitle("Cancelado")
            msg.setText("Operación cancelada. No se eliminó nada.")
            msg.setStyleSheet("""
                QMessageBox {
                    background-color: white;
                }
                QMessageBox QLabel {
                    color: #2c3e50;
                }
            """)
            msg.exec_()
            return

        # Proceder con la eliminación
        try:
            app_logger.warning(f"ADMIN {self.user_data.get('_id')} está eliminando TODAS las calificaciones")
            
            # ✅ NUEVO: Usar el service en lugar de acceso directo
            # Primero, obtener todas las calificaciones usando el servicio
            all_calificaciones = self.service.listar_calificaciones(
                self.user_data.get("_id"),
                self.user_rol
            )
            
            app_logger.info(f"Se encontraron {len(all_calificaciones)} calificaciones para eliminar")
            
            if len(all_calificaciones) == 0:
                msg = QMessageBox(self)
                msg.setIcon(QMessageBox.Information)
                msg.setWindowTitle("Sin datos")
                msg.setText("No hay calificaciones para eliminar en el sistema.")
                msg.setStyleSheet("""
                    QMessageBox {
                        background-color: white;
                    }
                    QMessageBox QLabel {
                        color: #2c3e50;
                    }
                """)
                msg.exec_()
                return
            
            deleted_count = 0
            errors_count = 0
            
            # Eliminar una por una usando el servicio
            for cal in all_calificaciones:
                try:
                    result = self.service.eliminar_calificacion(
                        cal["_id"],
                        self.user_data.get("_id"),
                        self.user_rol
                    )
                    if result.get("success"):
                        deleted_count += 1
                    else:
                        errors_count += 1
                        app_logger.error(f"Error al eliminar calificación {cal['_id']}: {result.get('message')}")
                except Exception as e:
                    errors_count += 1
                    app_logger.error(f"Error al eliminar calificación {cal['_id']}: {e}")

            app_logger.warning(
                f"Eliminación masiva completada. "
                f"Eliminadas: {deleted_count}, Errores: {errors_count}"
            )

            # Mostrar resultado
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Information)
            msg.setWindowTitle("Eliminación Completada")
            msg.setText(
                f"✅ Eliminadas: {deleted_count} calificaciones\n"
                f"❌ Errores: {errors_count}\n\n"
                f"La tabla se actualizará automáticamente."
            )
            msg.setStyleSheet("""
                QMessageBox {
                    background-color: white;
                }
                QMessageBox QLabel {
                    color: #2c3e50;
                    min-width: 300px;
                }
            """)
            msg.exec_()

            # Refrescar tabla
            self.refrescar_tabla()

        except Exception as e:
            app_logger.error(f"Error crítico al eliminar todas las calificaciones: {e}")
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Critical)
            msg.setWindowTitle("Error")
            msg.setText(
                f"Ocurrió un error al eliminar las calificaciones:\n\n{str(e)}\n\n"
                f"Consulte los logs para más detalles."
            )
            msg.setStyleSheet("""
                QMessageBox {
                    background-color: white;
                }
                QMessageBox QLabel {
                    color: #2c3e50;
                    min-width: 300px;
                }
            """)
            msg.exec_()

    def aplicar_filtros(self):
        filtros: Dict[str, Any] = {}
        filtros["fecha_desde"] = self.date_desde.date().toPyDate()
        filtros["fecha_hasta"] = self.date_hasta.date().toPyDate()
        tipo = self.combo_tipo.currentText()
        if tipo != "Todos":
            filtros["tipo_impuesto"] = tipo
        pais = self.combo_pais.currentText()
        if pais != "Todos":
            filtros["pais"] = pais

        calificaciones = self.service.listar_calificaciones(
            self.user_data.get("_id"),
            self.user_rol,
            filtros
        )
        self.actualizar_tabla(calificaciones)

    def limpiar_filtros(self):
        self.date_desde.setDate(QDate.currentDate().addMonths(-6))
        self.date_hasta.setDate(QDate.currentDate())
        self.combo_tipo.setCurrentIndex(0)
        self.combo_pais.setCurrentIndex(0)
        self.refrescar_tabla()

    def refrescar_tabla(self):
        """Refresca la tabla y limpia el caché de clientes"""
        self.clientes_cache.clear()  # Limpiar caché al refrescar
        
        calificaciones = self.service.listar_calificaciones(
            self.user_data.get("_id"),
            self.user_rol
        )
        self.actualizar_tabla(calificaciones)

    def apply_styles(self):
        """
        Estilos para la vista de gestión de calificaciones (tabla, toolbar, filtros).
        Busca un look similar al diseño anterior con header naranja y botones llamativos.
        """
        self.setStyleSheet("""
            /* Fondo del contenedor (dentro de la ventana principal) */
            QWidget {
                background-color: transparent;
            }

            /* Marco de filtros */
            QFrame {
                background-color: #ffffff;
                border: 1px solid #e6e9ee;
                border-radius: 8px;
            }

            /* Toolbar: botones naranja y azul */
            QPushButton {
                padding: 8px 12px;
                border-radius: 6px;
            }
            QPushButton[role="primary"] {
                background-color: #e76f23;
                color: white;
                font-weight: 600;
            }
            QPushButton[role="secondary"] {
                background-color: #3498db;
                color: white;
            }
            QPushButton[role="success"] {
                background-color: #16a085;
                color: white;
            }
            QPushButton[role="muted"] {
                background-color: #95a5a6;
                color: white;
            }

            /* Tabla */
            QTableWidget {
                background-color: #ffffff;
                border: 1px solid #e6e9ee;
                gridline-color: #f0f2f5;
            }
            QHeaderView::section {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #E94E1B, stop:1 #ff7a43);
                color: white;
                padding: 8px;
                font-weight: bold;
                border: none;
            }

            /* Celdas */
            QTableWidget::item {
                padding: 6px 8px;
            }

            /* Estado y resaltado para suma 8-19 */
            QTableWidget QTableCornerButton::section {
                background-color: #f9fafb;
            }

            /* Labels en footer */
            QLabel {
                color: #2c3e50;
            }
            QLabel[role="muted"] {
                color: #7f8c8d;
            }
        """)