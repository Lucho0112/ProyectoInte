from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QTableWidget, QTableWidgetItem, QHeaderView, QDateEdit,
    QComboBox, QFrame, QMessageBox, QAbstractItemView,
    QScrollArea, QRadioButton, QButtonGroup, QFileDialog
)
from PyQt5.QtCore import Qt, QDate, pyqtSignal
from PyQt5.QtGui import QFont, QColor, QCursor
from datetime import datetime
from services.reportService import ReportService
from utils.logger import app_logger


class GenerarReportesContent(QWidget):
    """Contenido de generación de reportes"""
    
    back_requested = pyqtSignal()
    
    def __init__(self, user_data: dict, parent=None):
        super().__init__(parent)
        self.user_data = user_data
        self.user_rol = user_data.get("rol", "cliente")
        self.service = ReportService()
        self.datos_actuales = []
        
        self.init_ui()
    
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
        
        # Header
        self.add_header(main_layout)
        
        # Filtros
        self.add_filters(main_layout)
        
        # Vista previa
        self.add_preview(main_layout)
        
        # Botones de exportación
        self.add_export_buttons(main_layout)
        
        # Historial (solo admin y auditores)
        if self.user_rol in ["administrador", "auditor_tributario"]:
            self.add_history(main_layout)
        
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
        
        title = QLabel("📊 Generar Reportes y Exportaciones")
        title.setFont(QFont("Arial", 18, QFont.Bold))
        title.setStyleSheet("color: #2c3e50;")
        header_layout.addWidget(title)
        
        header_layout.addStretch()
        layout.addLayout(header_layout)
    
    def add_filters(self, layout):
        """Panel de filtros"""
        filter_frame = QFrame()
        filter_frame.setStyleSheet("""
            QFrame {
                background-color: white;
                border: 2px solid #dee2e6;
                border-radius: 10px;
                padding: 20px;
            }
        """)
        
        filter_layout = QVBoxLayout()
        filter_layout.setSpacing(15)
        
        # Título
        filter_title = QLabel("🔍 Filtros de Exportación")
        filter_title.setFont(QFont("Arial", 13, QFont.Bold))
        filter_title.setStyleSheet("color: #2c3e50;")
        filter_layout.addWidget(filter_title)
        
        # Grid de filtros
        grid_layout = QHBoxLayout()
        grid_layout.setSpacing(15)
        
        # Columna 1: Fechas
        col1 = QVBoxLayout()
        col1.setSpacing(8)
        
        col1.addWidget(QLabel("Fecha desde:"))
        self.date_desde = QDateEdit()
        self.date_desde.setCalendarPopup(True)
        self.date_desde.setDate(QDate.currentDate().addMonths(-6))
        self.date_desde.setDisplayFormat("dd/MM/yyyy")
        self.date_desde.setMinimumHeight(35)
        col1.addWidget(self.date_desde)
        
        col1.addWidget(QLabel("Fecha hasta:"))
        self.date_hasta = QDateEdit()
        self.date_hasta.setCalendarPopup(True)
        self.date_hasta.setDate(QDate.currentDate())
        self.date_hasta.setDisplayFormat("dd/MM/yyyy")
        self.date_hasta.setMinimumHeight(35)
        col1.addWidget(self.date_hasta)
        
        grid_layout.addLayout(col1)
        
        # Columna 2: Tipo y País
        col2 = QVBoxLayout()
        col2.setSpacing(8)
        
        col2.addWidget(QLabel("Tipo de Impuesto:"))
        self.combo_tipo = QComboBox()
        self.combo_tipo.addItems(["Todos", "IVA", "Renta", "Importación", "Exportación", "Otro"])
        self.combo_tipo.setMinimumHeight(35)
        col2.addWidget(self.combo_tipo)
        
        col2.addWidget(QLabel("País:"))
        self.combo_pais = QComboBox()
        self.combo_pais.addItems(["Todos", "Chile", "Perú", "Colombia"])
        self.combo_pais.setMinimumHeight(35)
        col2.addWidget(self.combo_pais)
        
        grid_layout.addLayout(col2)
        
        # Columna 3: Estado
        col3 = QVBoxLayout()
        col3.setSpacing(8)
        
        col3.addWidget(QLabel("Estado de datos:"))
        
        self.button_group = QButtonGroup()
        
        self.radio_ambos = QRadioButton("📊 Ambos (Local + Bolsa)")
        self.radio_ambos.setChecked(True)
        self.button_group.addButton(self.radio_ambos)
        col3.addWidget(self.radio_ambos)
        
        self.radio_local = QRadioButton("💼 Solo Local")
        self.button_group.addButton(self.radio_local)
        col3.addWidget(self.radio_local)
        
        self.radio_bolsa = QRadioButton("🏛️ Solo Bolsa")
        self.button_group.addButton(self.radio_bolsa)
        col3.addWidget(self.radio_bolsa)
        
        col3.addStretch()
        grid_layout.addLayout(col3)
        
        filter_layout.addLayout(grid_layout)
        
        # Botones
        buttons_layout = QHBoxLayout()
        buttons_layout.addStretch()
        
        btn_limpiar = QPushButton("🗑️ Limpiar Filtros")
        btn_limpiar.setFont(QFont("Arial", 10))
        btn_limpiar.setMinimumHeight(40)
        btn_limpiar.setCursor(QCursor(Qt.PointingHandCursor))
        btn_limpiar.clicked.connect(self.limpiar_filtros)
        btn_limpiar.setStyleSheet("""
            QPushButton {
                background-color: #95a5a6;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 10px 20px;
            }
            QPushButton:hover { background-color: #7f8c8d; }
        """)
        buttons_layout.addWidget(btn_limpiar)
        
        btn_aplicar = QPushButton("🔍 Aplicar Filtros")
        btn_aplicar.setFont(QFont("Arial", 10, QFont.Bold))
        btn_aplicar.setMinimumHeight(40)
        btn_aplicar.setCursor(QCursor(Qt.PointingHandCursor))
        btn_aplicar.clicked.connect(self.aplicar_filtros)
        btn_aplicar.setStyleSheet("""
            QPushButton {
                background-color: #E94E1B;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 10px 20px;
            }
            QPushButton:hover { background-color: #d64419; }
        """)
        buttons_layout.addWidget(btn_aplicar)
        
        filter_layout.addLayout(buttons_layout)
        filter_frame.setLayout(filter_layout)
        layout.addWidget(filter_frame)
    
    def add_preview(self, layout):
        """Vista previa de datos"""
        preview_frame = QFrame()
        preview_frame.setStyleSheet("""
            QFrame {
                background-color: white;
                border: 2px solid #dee2e6;
                border-radius: 10px;
                padding: 20px;
            }
        """)
        
        preview_layout = QVBoxLayout()
        preview_layout.setSpacing(10)
        
        # Header de vista previa
        preview_header = QHBoxLayout()
        
        preview_title = QLabel("👁️ Vista Previa (primeras 50 filas)")
        preview_title.setFont(QFont("Arial", 13, QFont.Bold))
        preview_title.setStyleSheet("color: #2c3e50;")
        preview_header.addWidget(preview_title)
        
        preview_header.addStretch()
        
        self.label_contador = QLabel("Total: 0 registros")
        self.label_contador.setFont(QFont("Arial", 11, QFont.Bold))
        self.label_contador.setStyleSheet("color: #E94E1B;")
        preview_header.addWidget(self.label_contador)
        
        preview_layout.addLayout(preview_header)
        
        # Tabla
        self.preview_table = QTableWidget()
        self.preview_table.setColumnCount(8)
        self.preview_table.setHorizontalHeaderLabels([
            "RUT Cliente", "Fecha", "Tipo", "País",
            "Monto", "Suma 8-19", "Estado", "Válido"
        ])
        
        self.preview_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.preview_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.preview_table.setAlternatingRowColors(True)
        self.preview_table.verticalHeader().setVisible(False)
        
        header = self.preview_table.horizontalHeader()
        for i in range(8):
            header.setSectionResizeMode(i, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        
        self.preview_table.setStyleSheet("""
            QTableWidget {
                background-color: white;
                alternate-background-color: #f8f9fa;
                border: 1px solid #dee2e6;
                border-radius: 5px;
            }
            QHeaderView::section {
                background-color: #2c3e50;
                color: white;
                padding: 8px;
                font-weight: bold;
                border: none;
            }
        """)
        
        self.preview_table.setMaximumHeight(400)
        preview_layout.addWidget(self.preview_table)
        
        preview_frame.setLayout(preview_layout)
        layout.addWidget(preview_frame)
    
    def add_export_buttons(self, layout):
        """Botones de exportación"""
        export_frame = QFrame()
        export_frame.setStyleSheet("""
            QFrame {
                background-color: #f8f9fa;
                border: 2px solid #dee2e6;
                border-radius: 10px;
                padding: 20px;
            }
        """)
        
        export_layout = QHBoxLayout()
        export_layout.setSpacing(15)
        
        # Icono
        icon_label = QLabel("📥")
        icon_label.setFont(QFont("Arial", 36))
        export_layout.addWidget(icon_label)
        
        # Texto
        text_layout = QVBoxLayout()
        text_layout.setSpacing(5)
        
        title_label = QLabel("Exportar Datos")
        title_label.setFont(QFont("Arial", 13, QFont.Bold))
        title_label.setStyleSheet("color: #2c3e50;")
        text_layout.addWidget(title_label)
        
        desc_label = QLabel("Selecciona el formato de exportación deseado")
        desc_label.setFont(QFont("Arial", 9))
        desc_label.setStyleSheet("color: #7f8c8d;")
        text_layout.addWidget(desc_label)
        
        export_layout.addLayout(text_layout)
        export_layout.addStretch()
        
        # Botones
        self.btn_csv = QPushButton("📄 Exportar CSV")
        self.btn_csv.setFont(QFont("Arial", 11, QFont.Bold))
        self.btn_csv.setMinimumSize(180, 50)
        self.btn_csv.setCursor(QCursor(Qt.PointingHandCursor))
        self.btn_csv.clicked.connect(self.exportar_csv)
        self.btn_csv.setEnabled(False)
        self.btn_csv.setStyleSheet("""
            QPushButton {
                background-color: #27ae60;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 10px 20px;
            }
            QPushButton:hover:enabled { background-color: #229954; }
            QPushButton:disabled {
                background-color: #bdc3c7;
                color: #7f8c8d;
            }
        """)
        export_layout.addWidget(self.btn_csv)
        
        self.btn_excel = QPushButton("📊 Exportar Excel")
        self.btn_excel.setFont(QFont("Arial", 11, QFont.Bold))
        self.btn_excel.setMinimumSize(180, 50)
        self.btn_excel.setCursor(QCursor(Qt.PointingHandCursor))
        self.btn_excel.clicked.connect(self.exportar_excel)
        self.btn_excel.setEnabled(False)
        self.btn_excel.setStyleSheet("""
            QPushButton {
                background-color: #2980b9;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 10px 20px;
            }
            QPushButton:hover:enabled { background-color: #21618c; }
            QPushButton:disabled {
                background-color: #bdc3c7;
                color: #7f8c8d;
            }
        """)
        export_layout.addWidget(self.btn_excel)
        
        export_frame.setLayout(export_layout)
        layout.addWidget(export_frame)
    
    def add_history(self, layout):
        """Historial de reportes (solo admin/auditor)"""
        history_frame = QFrame()
        history_frame.setStyleSheet("""
            QFrame {
                background-color: white;
                border: 2px solid #dee2e6;
                border-radius: 10px;
                padding: 20px;
            }
        """)
        
        history_layout = QVBoxLayout()
        
        title = QLabel("📚 Historial de Reportes Generados")
        title.setFont(QFont("Arial", 13, QFont.Bold))
        title.setStyleSheet("color: #2c3e50;")
        history_layout.addWidget(title)
        
        # Tabla de historial
        self.history_table = QTableWidget()
        self.history_table.setColumnCount(5)
        self.history_table.setHorizontalHeaderLabels([
            "Fecha", "Archivo", "Formato", "Registros", "Usuario"
        ])
        
        self.history_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.history_table.setAlternatingRowColors(True)
        self.history_table.verticalHeader().setVisible(False)
        self.history_table.setMaximumHeight(200)
        
        history_layout.addWidget(self.history_table)
        
        # Botón refrescar
        btn_refrescar = QPushButton("🔄 Refrescar Historial")
        btn_refrescar.setMinimumHeight(35)
        btn_refrescar.setCursor(QCursor(Qt.PointingHandCursor))
        btn_refrescar.clicked.connect(self.cargar_historial)
        btn_refrescar.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 8px 16px;
            }
            QPushButton:hover { background-color: #2980b9; }
        """)
        history_layout.addWidget(btn_refrescar)
        
        history_frame.setLayout(history_layout)
        layout.addWidget(history_frame)
        
        # Cargar historial inicial
        self.cargar_historial()
    
    def aplicar_filtros(self):
        """Aplica los filtros y carga datos"""
        filtros = self.obtener_filtros()
        
        # Obtener datos
        self.datos_actuales = self.service.obtener_datos_filtrados(
            filtros,
            self.user_data.get("_id"),
            self.user_rol
        )
        
        if not self.datos_actuales:
            QMessageBox.information(
                self,
                "Sin resultados",
                "No se encontraron datos con los filtros aplicados."
            )
            self.btn_csv.setEnabled(False)
            self.btn_excel.setEnabled(False)
            self.preview_table.setRowCount(0)
            self.label_contador.setText("Total: 0 registros")
            return
        
        # Actualizar vista previa
        self.actualizar_vista_previa(self.datos_actuales[:50])
        
        # Habilitar exportación
        self.btn_csv.setEnabled(True)
        self.btn_excel.setEnabled(True)
        
        self.label_contador.setText(f"Total: {len(self.datos_actuales)} registros")
    
    def obtener_filtros(self) -> dict:
        """Obtiene los filtros actuales"""
        filtros = {
            "fecha_desde": self.date_desde.date().toPyDate(),
            "fecha_hasta": self.date_hasta.date().toPyDate()
        }
        
        if self.combo_tipo.currentText() != "Todos":
            filtros["tipo_impuesto"] = self.combo_tipo.currentText()
        
        if self.combo_pais.currentText() != "Todos":
            filtros["pais"] = self.combo_pais.currentText()
        
        # Estado
        if self.radio_local.isChecked():
            filtros["estado"] = "local"
        elif self.radio_bolsa.isChecked():
            filtros["estado"] = "bolsa"
        else:
            filtros["estado"] = "ambos"
        
        return filtros
    
    def actualizar_vista_previa(self, datos: list):
        """Actualiza la tabla de vista previa"""
        self.preview_table.setRowCount(len(datos))
        
        for row, cal in enumerate(datos):
            # RUT Cliente
            rut = self.service.obtener_rut_cliente(cal.get("clienteId", ""))
            self.preview_table.setItem(row, 0, QTableWidgetItem(rut))
            
            # Fecha
            self.preview_table.setItem(row, 1, QTableWidgetItem(cal.get("fechaDeclaracion", "")))
            
            # Tipo
            self.preview_table.setItem(row, 2, QTableWidgetItem(cal.get("tipoImpuesto", "")))
            
            # País
            self.preview_table.setItem(row, 3, QTableWidgetItem(cal.get("pais", "")))
            
            # Monto
            monto = cal.get("montoDeclarado", 0)
            item_monto = QTableWidgetItem(f"${monto:,.2f}")
            item_monto.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.preview_table.setItem(row, 4, item_monto)
            
            # Suma 8-19
            factores = cal.get("factores", {})
            suma = sum(factores.get(f"factor_{i}", 0) for i in range(8, 20))
            item_suma = QTableWidgetItem(f"{suma:.4f}")
            item_suma.setTextAlignment(Qt.AlignCenter)
            
            if suma > 1.0:
                item_suma.setBackground(QColor(255, 200, 200))
            else:
                item_suma.setBackground(QColor(200, 255, 200))
            
            self.preview_table.setItem(row, 5, item_suma)
            
            # Estado
            estado = "Local" if cal.get("esLocal", False) else "Bolsa"
            item_estado = QTableWidgetItem(estado)
            item_estado.setTextAlignment(Qt.AlignCenter)
            self.preview_table.setItem(row, 6, item_estado)
            
            # Válido
            valido = "✅ Sí" if suma <= 1.0 else "❌ No"
            self.preview_table.setItem(row, 7, QTableWidgetItem(valido))
    
    def exportar_csv(self):
        """Exporta datos a CSV"""
        if not self.datos_actuales:
            return
        
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Guardar reporte CSV",
            f"reporte_tributario_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            "CSV (*.csv)"
        )
        
        if not file_path:
            return
        
        filtros = self.obtener_filtros()
        result = self.service.exportar_csv(
            file_path,
            self.datos_actuales,
            filtros,
            self.user_data.get("_id")
        )
        
        if result["success"]:
            QMessageBox.information(self, "Éxito", result["message"])
            if hasattr(self, 'history_table'):
                self.cargar_historial()
        else:
            QMessageBox.warning(self, "Error", result["message"])
    
    def exportar_excel(self):
        """Exporta datos a Excel"""
        if not self.datos_actuales:
            return
        
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Guardar reporte Excel",
            f"reporte_tributario_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            "Excel (*.xlsx)"
        )
        
        if not file_path:
            return
        
        filtros = self.obtener_filtros()
        result = self.service.exportar_excel(
            file_path,
            self.datos_actuales,
            filtros,
            self.user_data.get("_id")
        )
        
        if result["success"]:
            QMessageBox.information(self, "Éxito", result["message"])
            if hasattr(self, 'history_table'):
                self.cargar_historial()
        else:
            QMessageBox.warning(self, "Error", result["message"])
    
    def limpiar_filtros(self):
        """Limpia todos los filtros"""
        self.date_desde.setDate(QDate.currentDate().addMonths(-6))
        self.date_hasta.setDate(QDate.currentDate())
        self.combo_tipo.setCurrentIndex(0)
        self.combo_pais.setCurrentIndex(0)
        self.radio_ambos.setChecked(True)
        
        self.datos_actuales = []
        self.preview_table.setRowCount(0)
        self.label_contador.setText("Total: 0 registros")
        self.btn_csv.setEnabled(False)
        self.btn_excel.setEnabled(False)
    
    def cargar_historial(self):
        """Carga el historial de reportes"""
        if not hasattr(self, 'history_table'):
            return
        
        reportes = self.service.obtener_historial_reportes(
            self.user_data.get("_id"),
            self.user_rol
        )
        
        self.history_table.setRowCount(len(reportes))
        
        for row, reporte in enumerate(reportes):
            # Fecha
            fecha = reporte.get("fechaGeneracion")
            if fecha:
                fecha_str = fecha.strftime("%Y-%m-%d %H:%M") if hasattr(fecha, 'strftime') else str(fecha)
            else:
                fecha_str = "N/A"
            self.history_table.setItem(row, 0, QTableWidgetItem(fecha_str))
            
            # Archivo
            self.history_table.setItem(row, 1, QTableWidgetItem(reporte.get("nombreArchivo", "")))
            
            # Formato
            self.history_table.setItem(row, 2, QTableWidgetItem(reporte.get("formato", "")))
            
            # Registros
            registros = str(reporte.get("totalRegistros", 0))
            self.history_table.setItem(row, 3, QTableWidgetItem(registros))
            
            # Usuario (solo admin ve esto)
            usuario_id = reporte.get("usuarioGeneradorId", "")
            if self.user_rol == "administrador":
                usuario_text = usuario_id[:8] + "..."
            else:
                usuario_text = "Yo"
            self.history_table.setItem(row, 4, QTableWidgetItem(usuario_text))