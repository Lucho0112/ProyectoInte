from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QTableWidget, QTableWidgetItem, QHeaderView, QDateEdit,
    QComboBox, QFrame, QMessageBox, QAbstractItemView,
    QScrollArea, QDialog, QFormLayout, QLineEdit, QDoubleSpinBox,
    QGroupBox, QGridLayout
)
from PyQt5.QtCore import Qt, QDate, pyqtSignal
from PyQt5.QtGui import QFont, QColor, QCursor
from datetime import datetime
from services.taxService import CalificacionTributariaService
from utils.logger import app_logger


class CalificacionFormDialog(QDialog):
    """Diálogo para crear/editar calificaciones"""
    
    def __init__(self, parent, user_data: dict, modo="crear", calificacion=None):
        super().__init__(parent)
        self.user_data = user_data
        self.modo = modo
        self.calificacion = calificacion
        self.service = CalificacionTributariaService()
        
        self.init_ui()
        
        if modo == "editar" and calificacion:
            self.cargar_datos()
    
    def init_ui(self):
        """Inicializa la interfaz"""
        titulo = "Nueva Calificación Tributaria" if self.modo == "crear" else "Editar Calificación Tributaria"
        self.setWindowTitle(titulo)
        self.setMinimumSize(800, 700)
        
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)
        
        # Título
        title_label = QLabel(titulo)
        title_label.setFont(QFont("Arial", 16, QFont.Bold))
        title_label.setStyleSheet("color: #2c3e50;")
        main_layout.addWidget(title_label)
        
        # Scroll
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        
        scroll_content = QWidget()
        scroll_layout = QVBoxLayout()
        scroll_layout.setSpacing(15)
        
        # Información General
        self.add_seccion_general(scroll_layout)
        
        # Factores
        self.add_seccion_factores(scroll_layout)
        
        # Validación
        self.add_validacion_suma(scroll_layout)
        
        scroll_content.setLayout(scroll_layout)
        scroll.setWidget(scroll_content)
        main_layout.addWidget(scroll)
        
        # Botones
        self.add_buttons(main_layout)
        
        self.setLayout(main_layout)
        self.apply_styles()
        
        # Si no es local, solo lectura
        if self.modo == "editar" and self.calificacion and not self.calificacion.get("esLocal", False):
            self.establecer_solo_lectura()
    
    def add_seccion_general(self, layout):
        """Información general"""
        group = QGroupBox("📝 Información General")
        group.setFont(QFont("Arial", 11, QFont.Bold))
        
        form_layout = QFormLayout()
        form_layout.setSpacing(10)
        
        # Cliente ID
        self.input_cliente = QLineEdit()
        self.input_cliente.setPlaceholderText("Ej: 12345678-9")
        form_layout.addRow("Cliente RUT: *", self.input_cliente)
        
        # Fecha
        self.input_fecha = QDateEdit()
        self.input_fecha.setCalendarPopup(True)
        self.input_fecha.setDate(QDate.currentDate())
        self.input_fecha.setDisplayFormat("dd/MM/yyyy")
        form_layout.addRow("Fecha Declaración: *", self.input_fecha)
        
        # Tipo Impuesto
        self.combo_tipo_impuesto = QComboBox()
        self.combo_tipo_impuesto.addItems(["IVA", "Renta", "Importación", "Exportación", "Otro"])
        form_layout.addRow("Tipo de Impuesto: *", self.combo_tipo_impuesto)
        
        # País
        self.combo_pais = QComboBox()
        self.combo_pais.addItems(["Chile", "Perú", "Colombia"])
        form_layout.addRow("País: *", self.combo_pais)
        
        # Monto
        self.input_monto = QDoubleSpinBox()
        self.input_monto.setRange(0.01, 999999999999.99)
        self.input_monto.setDecimals(2)
        self.input_monto.setPrefix("$ ")
        self.input_monto.setGroupSeparatorShown(True)
        form_layout.addRow("Monto Declarado: *", self.input_monto)
        
        group.setLayout(form_layout)
        layout.addWidget(group)
    
    def add_seccion_factores(self, layout):
        """Factores 1-19"""
        group = QGroupBox("📊 Factores (19 valores entre 0 y 1)")
        group.setFont(QFont("Arial", 11, QFont.Bold))
        
        grid_layout = QGridLayout()
        grid_layout.setSpacing(10)
        
        self.factor_inputs = []
        
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
            
            # Destacar factores 8-19
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
        """Indicador de validación"""
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
        """Botones"""
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        
        btn_cancelar = QPushButton("❌ Cancelar")
        btn_cancelar.setFont(QFont("Arial", 10))
        btn_cancelar.setMinimumSize(120, 40)
        btn_cancelar.setCursor(QCursor(Qt.PointingHandCursor))
        btn_cancelar.clicked.connect(self.reject)
        btn_cancelar.setStyleSheet("""
            QPushButton {
                background-color: #95a5a6;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 10px;
            }
            QPushButton:hover { background-color: #7f8c8d; }
        """)
        button_layout.addWidget(btn_cancelar)
        
        self.btn_guardar = QPushButton("💾 Guardar")
        self.btn_guardar.setFont(QFont("Arial", 10, QFont.Bold))
        self.btn_guardar.setMinimumSize(120, 40)
        self.btn_guardar.setCursor(QCursor(Qt.PointingHandCursor))
        self.btn_guardar.clicked.connect(self.guardar)
        self.btn_guardar.setStyleSheet("""
            QPushButton {
                background-color: #27ae60;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 10px;
            }
            QPushButton:hover { background-color: #229954; }
            QPushButton:disabled {
                background-color: #bdc3c7;
                color: #7f8c8d;
            }
        """)
        button_layout.addWidget(self.btn_guardar)
        
        layout.addLayout(button_layout)
    
    def actualizar_suma_factores(self):
        """Valida suma factores 8-19"""
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
        """Carga datos para edición"""
        if not self.calificacion:
            return
        
        self.input_cliente.setText(self.calificacion.get("clienteId", ""))
        
        fecha_str = self.calificacion.get("fechaDeclaracion", "")
        if fecha_str:
            try:
                fecha_obj = datetime.strptime(fecha_str, "%Y-%m-%d")
                qdate = QDate(fecha_obj.year, fecha_obj.month, fecha_obj.day)
                self.input_fecha.setDate(qdate)
            except:
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
        
        self.actualizar_suma_factores()
    
    def establecer_solo_lectura(self):
        """Modo solo lectura"""
        self.input_cliente.setEnabled(False)
        self.input_fecha.setEnabled(False)
        self.combo_tipo_impuesto.setEnabled(False)
        self.combo_pais.setEnabled(False)
        self.input_monto.setEnabled(False)
        
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
    
    def guardar(self):
        """Guarda la calificación"""
        try:
            fecha_date = self.input_fecha.date().toPyDate()
            fecha_declaracion = datetime.combine(fecha_date, datetime.min.time())
            
            datos = {
                "cliente_id": self.input_cliente.text().strip(),
                "fecha_declaracion": fecha_declaracion,
                "tipo_impuesto": self.combo_tipo_impuesto.currentText(),
                "pais": self.combo_pais.currentText(),
                "monto_declarado": self.input_monto.value(),
                "subsidios_aplicados": []
            }
            
            factores = [spinbox.value() for spinbox in self.factor_inputs]
            datos["factores"] = factores
            
            if not datos["cliente_id"]:
                QMessageBox.warning(self, "Validación", "El RUT del cliente es obligatorio")
                return
            
            usuario_id = self.user_data.get("_id")
            user_rol = self.user_data.get("rol", "cliente")
            
            if self.modo == "crear":
                result = self.service.crear_calificacion(datos, usuario_id)
                
                # ← NUEVO: Manejar conflicto
                if not result["success"] and result.get("conflicto", False):
                    dato_oficial = result.get("dato_oficial", {})
                    
                    reply = QMessageBox.warning(
                        self,
                        "⚠️ Conflicto con Dato Oficial",
                        f"Ya existe una calificación OFICIAL de bolsa:\n\n"
                        f"• Monto: ${dato_oficial.get('monto', 0):,.2f}\n"
                        f"• Fecha: {dato_oficial.get('fecha', 'N/A')}\n\n"
                        f"No se puede crear una calificación local duplicada.\n"
                        f"Cambie la fecha o el tipo de impuesto.",
                        QMessageBox.Ok
                    )
                    return
            else:
                result = self.service.actualizar_calificacion(
                    self.calificacion["_id"],
                    datos,
                    usuario_id,
                    user_rol  # ← NUEVO: Pasar rol
                )
            
            if result["success"]:
                self.accept()
            else:
                QMessageBox.warning(self, "Error", result["message"])
        
        except Exception as e:
            app_logger.error(f"Error al guardar: {str(e)}")
            QMessageBox.critical(self, "Error", f"Error al guardar: {str(e)}")
    
    def apply_styles(self):
        """Estilos"""
        self.setStyleSheet("""
            QDialog { background-color: #f8f9fa; }
            QLabel { color: #2c3e50; }
            
            QLineEdit {
                padding: 8px 12px;
                border: 2px solid #dee2e6;
                border-radius: 6px;
                background-color: white;
                min-height: 25px;
            }
            QLineEdit:hover { border: 2px solid #E94E1B; background-color: #fef5f1; }
            QLineEdit:focus { border: 2px solid #E94E1B; }
            
            QComboBox {
                padding: 8px 12px;
                border: 2px solid #dee2e6;
                border-radius: 6px;
                background-color: white;
                min-height: 25px;
            }
            QComboBox:hover { border: 2px solid #E94E1B; background-color: #fef5f1; }
            QComboBox::drop-down { border: none; width: 30px; }
            QComboBox::down-arrow {
                image: none;
                border-left: 5px solid transparent;
                border-right: 5px solid transparent;
                border-top: 6px solid #E94E1B;
            }
            
            QDateEdit {
                padding: 8px 12px;
                border: 2px solid #dee2e6;
                border-radius: 6px;
                background-color: white;
                min-height: 25px;
            }
            QDateEdit:hover { border: 2px solid #E94E1B; background-color: #fef5f1; }
            
            QDoubleSpinBox {
                padding: 6px 10px;
                border: 2px solid #dee2e6;
                border-radius: 6px;
                background-color: white;
                min-height: 30px;
            }
            QDoubleSpinBox:hover { border: 2px solid #E94E1B; background-color: #fef5f1; }
            
            QDoubleSpinBox[highlighted="true"] {
                background-color: #fff9c4;
                border: 2px solid #fdd835;
            }
            QDoubleSpinBox[highlighted="true"]:hover {
                border: 2px solid #E94E1B;
                background-color: #fff5ba;
            }
            
            QGroupBox {
                border: 2px solid #dee2e6;
                border-radius: 8px;
                margin-top: 15px;
                padding-top: 20px;
                background-color: white;
            }
        """)


class GestionCalificacionesContent(QWidget):
    """Contenido de gestión de calificaciones (sin ser ventana independiente)"""
    
    back_requested = pyqtSignal()
    
    def __init__(self, user_data: dict, parent=None):
        super().__init__(parent)
        self.user_data = user_data
        self.user_rol = user_data.get("rol", "cliente")
        self.service = CalificacionTributariaService()
        self.calificaciones = []
        
        self.init_ui()
        self.refrescar_tabla()

    def obtener_rut_cliente(self, cliente_id: str) -> str:
        """
        Obtiene el RUT de un cliente dado su ID de Firebase
        
        Args:
            cliente_id (str): ID del cliente en Firebase
            
        Returns:
            str: RUT del cliente o "Cliente no encontrado"
        """
        try:
            from config.firebaseConfig import firebase_config
            db = firebase_config.get_firestore_client()
            
            # Buscar cliente por ID
            doc = db.collection("usuarios").document(cliente_id).get()
            
            if doc.exists:
                data = doc.to_dict()
                return data.get("rut", "RUT no disponible")
            else:
                return "Cliente no encontrado"
        
        except Exception as e:
            app_logger.error(f"Error al obtener RUT: {str(e)}")
            return "Error al cargar"
    
    def init_ui(self):
        """Inicializa la interfaz"""
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
    
    def add_header(self, layout):
        """Header con botón volver"""
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
        """Barra de herramientas"""
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

        if self.user_rol == "administrador":
            btn_limpiar = QPushButton("🗑️ Limpiar Todo")
            btn_limpiar.setFont(QFont("Arial", 10))
            btn_limpiar.setMinimumHeight(40)
            btn_limpiar.setCursor(QCursor(Qt.PointingHandCursor))
            btn_limpiar.clicked.connect(self.limpiar_todas_calificaciones)
            btn_limpiar.setStyleSheet("""
                QPushButton {
                    background-color: #e74c3c;
                    color: white;
                    border: none;
                    border-radius: 5px;
                    padding: 10px 20px;
                }
                QPushButton:hover { background-color: #c0392b; }
            """)
            toolbar_layout.addWidget(btn_limpiar)
        
        toolbar_layout.addStretch()
        
        self.label_contador = QLabel("Total: 0 calificaciones")
        self.label_contador.setFont(QFont("Arial", 10, QFont.Bold))
        self.label_contador.setStyleSheet("color: #2c3e50;")
        toolbar_layout.addWidget(self.label_contador)
        
        layout.addLayout(toolbar_layout)
    
    def add_filters(self, layout):
        """Panel de filtros"""
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
        """Tabla de calificaciones"""
        self.table = QTableWidget()
        self.table.setColumnCount(10)
        self.table.setHorizontalHeaderLabels([
            "ID", "Cliente", "Fecha", "Tipo", "País",
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
        
        self.table.setStyleSheet("""
            QTableWidget {
                background-color: white;
                alternate-background-color: #f8f9fa;
                border: 1px solid #dee2e6;
                border-radius: 5px;
            }
            QHeaderView::section {
                background-color: #E94E1B;
                color: white;
                padding: 8px;
                font-weight: bold;
            }
        """)
        
        layout.addWidget(self.table)
    
    def add_footer(self, layout):
        """Footer"""
        if self.user_rol == "administrador":
            footer_text = "💡 Admin: Puedes ver, editar y eliminar TODAS las calificaciones (locales y de bolsa)"
        else:
            footer_text = "💡 Tip: Solo puedes editar/eliminar tus calificaciones locales. Los datos de bolsa son de solo lectura."
        
        footer = QLabel(footer_text)
        footer.setFont(QFont("Arial", 8))
        footer.setStyleSheet("color: #7f8c8d;")
        layout.addWidget(footer)
    
    def actualizar_tabla(self, calificaciones: list):
        """Actualiza la tabla"""
        self.calificaciones = calificaciones
        self.table.setRowCount(len(calificaciones))
        
        for row, cal in enumerate(calificaciones):
            # ID
            item_id = QTableWidgetItem(cal["_id"][:8])
            self.table.setItem(row, 0, item_id)
            
            # Cliente
            cliente_id = cal.get("clienteId", "")
            rut_cliente = self.obtener_rut_cliente(cliente_id)
            self.table.setItem(row, 1, QTableWidgetItem(rut_cliente))
            
            # Fecha
            fecha_str = cal.get("fechaDeclaracion", "")
            self.table.setItem(row, 2, QTableWidgetItem(fecha_str))
            
            # Tipo
            self.table.setItem(row, 3, QTableWidgetItem(cal.get("tipoImpuesto", "")))
            
            # País
            self.table.setItem(row, 4, QTableWidgetItem(cal.get("pais", "")))
            
            # Monto
            monto = cal.get("montoDeclarado", 0)
            item_monto = QTableWidgetItem(f"${monto:,.2f}")
            item_monto.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.table.setItem(row, 5, item_monto)
            
            # Suma 8-19
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
            
            # Estado
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
            
            # Botón Editar
            btn_editar = QPushButton("✏️")
            btn_editar.setEnabled(es_admin or es_local)
            btn_editar.setCursor(QCursor(Qt.PointingHandCursor) if (es_admin or es_local) else QCursor(Qt.ForbiddenCursor))
            btn_editar.clicked.connect(lambda checked, c=cal: self.abrir_formulario_editar(c))
            btn_editar.setStyleSheet("""
                QPushButton {
                    background-color: #3498db;
                    color: white;
                    border: none;
                    border-radius: 3px;
                    padding: 5px 10px;
                }
                QPushButton:hover:enabled { background-color: #2980b9; }
                QPushButton:disabled {
                    background-color: #bdc3c7;
                    color: #7f8c8d;
                }
            """)
            self.table.setCellWidget(row, 8, btn_editar)
            
            # Botón Eliminar
            btn_eliminar = QPushButton("🗑️")
            btn_eliminar.setEnabled(es_admin or es_local)  # ← CAMBIO AQUÍ
            btn_eliminar.setCursor(QCursor(Qt.PointingHandCursor) if (es_admin or es_local) else QCursor(Qt.ForbiddenCursor))
            btn_eliminar.clicked.connect(lambda checked, c=cal: self.eliminar_calificacion(c))
            btn_eliminar.setStyleSheet("""
                QPushButton {
                    background-color: #e74c3c;
                    color: white;
                    border: none;
                    border-radius: 3px;
                    padding: 5px 10px;
                }
                QPushButton:hover:enabled { background-color: #c0392b; }
                QPushButton:disabled {
                    background-color: #bdc3c7;
                    color: #7f8c8d;
                }
            """)
            self.table.setCellWidget(row, 9, btn_eliminar)
        
        self.label_contador.setText(f"Total: {len(calificaciones)} calificaciones")
    
    def abrir_formulario_crear(self):
        """Abre formulario de creación"""
        dialog = CalificacionFormDialog(self, self.user_data, modo="crear")
        if dialog.exec_():
            self.refrescar_tabla()
            QMessageBox.information(self, "Éxito", "Calificación creada exitosamente")
    
    def abrir_formulario_editar(self, calificacion: dict):
        """Abre formulario de edición"""
        es_local = calificacion.get("esLocal", False)
        es_admin = self.user_rol == "administrador"
        
        # ← NUEVO: Solo advertir si NO es admin y NO es local
        if not es_admin and not es_local:
            QMessageBox.warning(
                self,
                "Acción no permitida",
                "No se puede editar una calificación de bolsa.\n\n"
                "Solo las calificaciones locales pueden ser modificadas."
            )
            return
        
        dialog = CalificacionFormDialog(self, self.user_data, modo="editar", calificacion=calificacion)
        if dialog.exec_():
            self.refrescar_tabla()
            QMessageBox.information(self, "Éxito", "Calificación actualizada exitosamente")
    
    def eliminar_calificacion(self, calificacion: dict):
        """Elimina una calificación"""
        es_local = calificacion.get("esLocal", False)
        es_admin = self.user_rol == "administrador"
        
        # ← NUEVO: Solo advertir si NO es admin y NO es local
        if not es_admin and not es_local:
            QMessageBox.warning(
                self,
                "Acción no permitida",
                "No se puede eliminar una calificación de bolsa."
            )
            return
        
        # ← NUEVO: Advertencia especial si admin elimina dato de bolsa
        if es_admin and not es_local:
            reply = QMessageBox.warning(
                self,
                "⚠️ ADVERTENCIA: Eliminar Dato de Bolsa",
                f"Está a punto de eliminar un DATO OFICIAL DE BOLSA.\n\n"
                f"Cliente: {calificacion.get('clienteId', 'N/A')}\n"
                f"Fecha: {calificacion.get('fechaDeclaracion', 'N/A')}\n\n"
                f"Esta acción debe ser justificada y registrada.\n"
                f"¿Está seguro?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No  # No por defecto
            )
        else:
            # Usuario normal eliminando su dato local
            reply = QMessageBox.question(
                self,
                "Confirmar Eliminación",
                f"¿Está seguro de eliminar esta calificación?\n\n"
                f"Cliente: {calificacion.get('clienteId', 'N/A')}\n"
                f"Fecha: {calificacion.get('fechaDeclaracion', 'N/A')}",
                QMessageBox.Yes | QMessageBox.No
            )
        
        if reply == QMessageBox.Yes:
            result = self.service.eliminar_calificacion(
                calificacion["_id"],
                self.user_data.get("_id"),
                self.user_rol  # ← NUEVO: Pasar rol
            )
            
            if result["success"]:
                self.refrescar_tabla()
                QMessageBox.information(self, "Éxito", "Calificación eliminada exitosamente")
            else:
                QMessageBox.warning(self, "Error", result["message"])
    
    def aplicar_filtros(self):
        """Aplica filtros"""
        filtros = {}
        
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
        """Limpia filtros"""
        self.date_desde.setDate(QDate.currentDate().addMonths(-6))
        self.date_hasta.setDate(QDate.currentDate())
        self.combo_tipo.setCurrentIndex(0)
        self.combo_pais.setCurrentIndex(0)
        self.refrescar_tabla()
    
    def refrescar_tabla(self):
        """Refresca la tabla"""
        calificaciones = self.service.listar_calificaciones(
            self.user_data.get("_id"),
            self.user_rol
        )
        self.actualizar_tabla(calificaciones)

    # AL FINAL de GestionCalificacionesContent:

    def limpiar_todas_calificaciones(self):
        """Elimina TODAS las calificaciones (solo admin)"""
        reply = QMessageBox.warning(
            self,
            "⚠️ ADVERTENCIA CRÍTICA",
            f"Está a punto de ELIMINAR TODAS las calificaciones del sistema.\n\n"
            f"Total actual: {len(self.calificaciones)} calificaciones\n\n"
            f"Esta acción NO se puede deshacer.\n"
            f"¿Está COMPLETAMENTE seguro?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            # Confirmar nuevamente
            reply2 = QMessageBox.critical(
                self,
                "ÚLTIMA CONFIRMACIÓN",
                "Escriba 'ELIMINAR TODO' para confirmar:",
                QMessageBox.Ok | QMessageBox.Cancel
            )
            
            if reply2 == QMessageBox.Ok:
                # Eliminar todas
                errores = 0
                eliminadas = 0
                
                for cal in self.calificaciones:
                    result = self.service.eliminar_calificacion(
                        cal["_id"],
                        self.user_data.get("_id"),
                        self.user_rol
                    )
                    if result["success"]:
                        eliminadas += 1
                    else:
                        errores += 1
                
                QMessageBox.information(
                    self,
                    "Limpieza Completada",
                    f"✅ Eliminadas: {eliminadas}\n❌ Errores: {errores}"
                )
                
                self.refrescar_tabla()