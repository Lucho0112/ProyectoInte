from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QTableWidget, QTableWidgetItem, QHeaderView, QDateEdit,
    QComboBox, QFrame, QMessageBox, QAbstractItemView,
    QScrollArea, QLineEdit
)
from PyQt5.QtCore import Qt, QDate, pyqtSignal
from PyQt5.QtGui import QFont, QColor, QCursor
from datetime import datetime
from services.reportService import ReportService
from utils.logger import app_logger


class ConsultarDatosContent(QWidget):
    """Contenido de consulta y filtrado de datos tributarios"""
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
        
        # Toolbar
        self.add_toolbar(main_layout)
        
        # Filtros
        self.add_filters(main_layout)
        
        # Tabla de resultados
        self.add_results_table(main_layout)
        
        # Footer con tips
        self.add_footer(main_layout)
        
        content_widget.setLayout(main_layout)
        scroll_area.setWidget(content_widget)
        
        widget_layout = QVBoxLayout()
        widget_layout.setContentsMargins(0, 0, 0, 0)
        widget_layout.addWidget(scroll_area)
        self.setLayout(widget_layout)
        
        # Aplicar estilos
        self.apply_styles()
    
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
        
        title = QLabel("🔍 Consultar y Filtrar Datos Tributarios")
        title.setFont(QFont("Arial", 18, QFont.Bold))
        title.setStyleSheet("color: #2c3e50;")
        header_layout.addWidget(title)
        
        header_layout.addStretch()
        layout.addLayout(header_layout)
    
    def add_toolbar(self, layout):
        """Barra de herramientas"""
        toolbar_layout = QHBoxLayout()
        toolbar_layout.setSpacing(10)
        
        btn_buscar = QPushButton("🔍 Buscar")
        btn_buscar.setFont(QFont("Arial", 10, QFont.Bold))
        btn_buscar.setMinimumHeight(40)
        btn_buscar.setCursor(QCursor(Qt.PointingHandCursor))
        btn_buscar.clicked.connect(self.buscar_datos)
        btn_buscar.setProperty("role", "primary")
        toolbar_layout.addWidget(btn_buscar)
        
        btn_limpiar = QPushButton("🗑️ Limpiar Filtros")
        btn_limpiar.setFont(QFont("Arial", 10))
        btn_limpiar.setMinimumHeight(40)
        btn_limpiar.setCursor(QCursor(Qt.PointingHandCursor))
        btn_limpiar.clicked.connect(self.limpiar_filtros)
        btn_limpiar.setProperty("role", "muted")
        toolbar_layout.addWidget(btn_limpiar)
        
        btn_refrescar = QPushButton("🔄 Refrescar")
        btn_refrescar.setFont(QFont("Arial", 10))
        btn_refrescar.setMinimumHeight(40)
        btn_refrescar.setCursor(QCursor(Qt.PointingHandCursor))
        btn_refrescar.clicked.connect(self.refrescar_datos)
        btn_refrescar.setProperty("role", "secondary")
        toolbar_layout.addWidget(btn_refrescar)
        
        toolbar_layout.addStretch()
        
        self.label_contador = QLabel("Total: 0 registros")
        self.label_contador.setFont(QFont("Arial", 11, QFont.Bold))
        self.label_contador.setStyleSheet("color: #2c3e50;")
        toolbar_layout.addWidget(self.label_contador)
        
        layout.addLayout(toolbar_layout)
    
    def add_filters(self, layout):
        """Panel de filtros de consulta"""
        filter_frame = QFrame()
        filter_frame.setStyleSheet("""
            QFrame {
                background-color: white;
                border: 1px solid #e6e9ee;
                border-radius: 8px;
                padding: 15px;
            }
        """)
        
        filter_layout = QVBoxLayout()
        filter_layout.setSpacing(15)
        
        # Título
        filter_title = QLabel("🔍 Criterios de Búsqueda")
        filter_title.setFont(QFont("Arial", 13, QFont.Bold))
        filter_title.setStyleSheet("color: #2c3e50;")
        filter_layout.addWidget(filter_title)
        
        # Grid de filtros - Fila 1
        grid_layout1 = QHBoxLayout()
        grid_layout1.setSpacing(15)
        
        # Fechas
        fecha_layout = QVBoxLayout()
        fecha_layout.setSpacing(8)
        fecha_layout.addWidget(QLabel("Fecha desde:"))
        self.date_desde = QDateEdit()
        self.date_desde.setCalendarPopup(True)
        self.date_desde.setDate(QDate.currentDate().addMonths(-3))
        self.date_desde.setDisplayFormat("dd/MM/yyyy")
        self.date_desde.setMinimumHeight(35)
        fecha_layout.addWidget(self.date_desde)
        grid_layout1.addLayout(fecha_layout)
        
        fecha_layout2 = QVBoxLayout()
        fecha_layout2.setSpacing(8)
        fecha_layout2.addWidget(QLabel("Fecha hasta:"))
        self.date_hasta = QDateEdit()
        self.date_hasta.setCalendarPopup(True)
        self.date_hasta.setDate(QDate.currentDate())
        self.date_hasta.setDisplayFormat("dd/MM/yyyy")
        self.date_hasta.setMinimumHeight(35)
        fecha_layout2.addWidget(self.date_hasta)
        grid_layout1.addLayout(fecha_layout2)
        
        # Tipo de Impuesto
        tipo_layout = QVBoxLayout()
        tipo_layout.setSpacing(8)
        tipo_layout.addWidget(QLabel("Tipo de Impuesto:"))
        self.combo_tipo = QComboBox()
        self.combo_tipo.addItems(["Todos", "IVA", "Renta", "Importación", "Exportación", "Otro"])
        self.combo_tipo.setMinimumHeight(35)
        tipo_layout.addWidget(self.combo_tipo)
        grid_layout1.addLayout(tipo_layout)
        
        # País
        pais_layout = QVBoxLayout()
        pais_layout.setSpacing(8)
        pais_layout.addWidget(QLabel("País:"))
        self.combo_pais = QComboBox()
        self.combo_pais.addItems(["Todos", "Chile", "Perú", "Colombia"])
        self.combo_pais.setMinimumHeight(35)
        pais_layout.addWidget(self.combo_pais)
        grid_layout1.addLayout(pais_layout)
        
        filter_layout.addLayout(grid_layout1)
        
        # Grid de filtros - Fila 2
        grid_layout2 = QHBoxLayout()
        grid_layout2.setSpacing(15)
        
        # Monto mínimo
        monto_min_layout = QVBoxLayout()
        monto_min_layout.setSpacing(8)
        monto_min_layout.addWidget(QLabel("Monto mínimo ($):"))
        self.input_monto_min = QLineEdit()
        self.input_monto_min.setPlaceholderText("Ej: 1000")
        self.input_monto_min.setMinimumHeight(35)
        monto_min_layout.addWidget(self.input_monto_min)
        grid_layout2.addLayout(monto_min_layout)
        
        # Monto máximo
        monto_max_layout = QVBoxLayout()
        monto_max_layout.setSpacing(8)
        monto_max_layout.addWidget(QLabel("Monto máximo ($):"))
        self.input_monto_max = QLineEdit()
        self.input_monto_max.setPlaceholderText("Ej: 100000")
        self.input_monto_max.setMinimumHeight(35)
        monto_max_layout.addWidget(self.input_monto_max)
        grid_layout2.addLayout(monto_max_layout)
        
        # RUT Cliente (búsqueda)
        rut_layout = QVBoxLayout()
        rut_layout.setSpacing(8)
        rut_layout.addWidget(QLabel("RUT Cliente:"))
        self.input_rut = QLineEdit()
        self.input_rut.setPlaceholderText("Ej: 12345678-9")
        self.input_rut.setMinimumHeight(35)
        rut_layout.addWidget(self.input_rut)
        grid_layout2.addLayout(rut_layout)
        
        # Estado (Local/Bolsa)
        estado_layout = QVBoxLayout()
        estado_layout.setSpacing(8)
        estado_layout.addWidget(QLabel("Estado:"))
        self.combo_estado = QComboBox()
        self.combo_estado.addItems(["Todos", "Local", "Bolsa"])
        self.combo_estado.setMinimumHeight(35)
        estado_layout.addWidget(self.combo_estado)
        grid_layout2.addLayout(estado_layout)
        
        filter_layout.addLayout(grid_layout2)
        
        filter_frame.setLayout(filter_layout)
        layout.addWidget(filter_frame)
    
    def add_results_table(self, layout):
        """Tabla de resultados"""
        results_frame = QFrame()
        results_frame.setStyleSheet("""
            QFrame {
                background-color: white;
                border: 1px solid #e6e9ee;
                border-radius: 8px;
                padding: 15px;
            }
        """)
        
        results_layout = QVBoxLayout()
        results_layout.setSpacing(10)
        
        # Header
        results_header = QHBoxLayout()
        
        results_title = QLabel("📋 Resultados de la Búsqueda")
        results_title.setFont(QFont("Arial", 13, QFont.Bold))
        results_title.setStyleSheet("color: #2c3e50;")
        results_header.addWidget(results_title)
        
        results_header.addStretch()
        results_layout.addLayout(results_header)
        
        # Tabla
        self.results_table = QTableWidget()
        self.results_table.setColumnCount(10)
        self.results_table.setHorizontalHeaderLabels([
            "ID", "RUT Cliente", "Fecha", "Tipo", "País",
            "Monto", "Suma 8-19", "Estado", "Válido", "Ver Detalles"
        ])
        
        self.results_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.results_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.results_table.setAlternatingRowColors(True)
        self.results_table.verticalHeader().setVisible(False)
        
        # Ocultar columna ID
        self.results_table.setColumnHidden(0, True)
        
        header = self.results_table.horizontalHeader()
        for i in range(10):
            header.setSectionResizeMode(i, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        
        results_layout.addWidget(self.results_table)
        
        results_frame.setLayout(results_layout)
        layout.addWidget(results_frame)
    
    def add_footer(self, layout):
        """Footer con información útil"""
        if self.user_rol == "administrador":
            footer_text = "💡 Admin: Puedes consultar TODOS los datos del sistema (locales y de bolsa)"
        elif self.user_rol in ["analista_mercado", "auditor_tributario"]:
            footer_text = "💡 Puedes consultar todos los datos de la bolsa y tus datos locales"
        else:
            footer_text = "💡 Puedes consultar los datos de la bolsa y tus propios datos locales"
        
        footer = QLabel(footer_text)
        footer.setFont(QFont("Arial", 9))
        footer.setStyleSheet("color: #7f8c8d;")
        layout.addWidget(footer)
    
    def buscar_datos(self):
        """Realiza la búsqueda con los filtros aplicados"""
        filtros = self.obtener_filtros()
        
        try:
            # Obtener datos filtrados
            self.datos_actuales = self.service.obtener_datos_filtrados(
                filtros,
                self.user_data.get("_id"),
                self.user_rol
            )
            
            if not self.datos_actuales:
                QMessageBox.information(
                    self,
                    "Sin resultados",
                    "No se encontraron datos con los criterios especificados."
                )
                self.results_table.setRowCount(0)
                self.label_contador.setText("Total: 0 registros")
                return
            
            # Actualizar tabla
            self.actualizar_tabla(self.datos_actuales)
            self.label_contador.setText(f"Total: {len(self.datos_actuales)} registros")
            
        except Exception as e:
            app_logger.error(f"Error en búsqueda: {str(e)}")
            QMessageBox.warning(
                self,
                "Error",
                f"Ocurrió un error al buscar los datos: {str(e)}"
            )
    
    def obtener_filtros(self) -> dict:
        """Obtiene los filtros actuales"""
        filtros = {
            "fecha_desde": self.date_desde.date().toPyDate(),
            "fecha_hasta": self.date_hasta.date().toPyDate()
        }
        
        # Tipo de impuesto
        if self.combo_tipo.currentText() != "Todos":
            filtros["tipo_impuesto"] = self.combo_tipo.currentText()
        
        # País
        if self.combo_pais.currentText() != "Todos":
            filtros["pais"] = self.combo_pais.currentText()
        
        # Estado (Local/Bolsa)
        estado_texto = self.combo_estado.currentText()
        if estado_texto == "Local":
            filtros["estado"] = "local"
        elif estado_texto == "Bolsa":
            filtros["estado"] = "bolsa"
        else:
            filtros["estado"] = "ambos"
        
        # Monto mínimo
        monto_min = self.input_monto_min.text().strip()
        if monto_min:
            try:
                filtros["monto_minimo"] = float(monto_min)
            except ValueError:
                pass
        
        # Monto máximo
        monto_max = self.input_monto_max.text().strip()
        if monto_max:
            try:
                filtros["monto_maximo"] = float(monto_max)
            except ValueError:
                pass
        
        # RUT Cliente
        rut = self.input_rut.text().strip()
        if rut:
            filtros["rut_cliente"] = rut
        
        return filtros
    
    def actualizar_tabla(self, datos: list):
        """Actualiza la tabla de resultados"""
        self.results_table.setRowCount(len(datos))
        
        for row, cal in enumerate(datos):
            # ID (oculto)
            self.results_table.setItem(row, 0, QTableWidgetItem(cal.get("_id", "")))
            
            # RUT Cliente
            rut = self.service.obtener_rut_cliente(cal.get("clienteId", ""))
            self.results_table.setItem(row, 1, QTableWidgetItem(rut))
            
            # Fecha
            self.results_table.setItem(row, 2, QTableWidgetItem(cal.get("fechaDeclaracion", "")))
            
            # Tipo
            self.results_table.setItem(row, 3, QTableWidgetItem(cal.get("tipoImpuesto", "")))
            
            # País
            self.results_table.setItem(row, 4, QTableWidgetItem(cal.get("pais", "")))
            
            # Monto
            monto = cal.get("montoDeclarado", 0)
            item_monto = QTableWidgetItem(f"${monto:,.2f}")
            item_monto.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.results_table.setItem(row, 5, item_monto)
            
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
            
            self.results_table.setItem(row, 6, item_suma)
            
            # Estado
            es_local = cal.get("esLocal", False)
            estado = "Local" if es_local else "Bolsa"
            item_estado = QTableWidgetItem(estado)
            item_estado.setTextAlignment(Qt.AlignCenter)
            
            if es_local:
                item_estado.setBackground(QColor(200, 230, 255))
                item_estado.setForeground(QColor(0, 100, 200))
            else:
                item_estado.setBackground(QColor(230, 230, 230))
                item_estado.setForeground(QColor(100, 100, 100))
            
            self.results_table.setItem(row, 7, item_estado)
            
            # Válido
            valido = "✅ Sí" if suma <= 1.0 else "❌ No"
            self.results_table.setItem(row, 8, QTableWidgetItem(valido))
            
            # Botón Ver Detalles
            btn_ver = QPushButton("👁️ Ver")
            btn_ver.setProperty("role", "secondary")
            btn_ver.setCursor(QCursor(Qt.PointingHandCursor))
            btn_ver.clicked.connect(lambda checked, c=cal: self.ver_detalles(c))
            self.results_table.setCellWidget(row, 9, btn_ver)
    
    def ver_detalles(self, calificacion: dict):
        """Muestra los detalles completos de una calificación"""
        detalles = f"""
        📋 DETALLES DE LA CALIFICACIÓN
        
        🆔 ID: {calificacion.get('_id', 'N/A')[:12]}...
        👤 Cliente: {self.service.obtener_rut_cliente(calificacion.get('clienteId', ''))}
        📅 Fecha: {calificacion.get('fechaDeclaracion', 'N/A')}
        🏷️ Tipo: {calificacion.get('tipoImpuesto', 'N/A')}
        🌎 País: {calificacion.get('pais', 'N/A')}
        💰 Monto: ${calificacion.get('montoDeclarado', 0):,.2f}
        📊 Estado: {'Local' if calificacion.get('esLocal', False) else 'Bolsa'}
        
        📐 FACTORES:
        """
        
        factores = calificacion.get("factores", {})
        for i in range(1, 20):
            valor = factores.get(f"factor_{i}", 0)
            detalles += f"\n  Factor {i}: {valor:.4f}"
        
        suma_8_19 = sum(factores.get(f"factor_{i}", 0) for i in range(8, 20))
        detalles += f"\n\n✅ Suma Factores 8-19: {suma_8_19:.4f}"
        detalles += f"\n{'✅ Válido' if suma_8_19 <= 1.0 else '❌ Inválido (> 1.0)'}"
        
        # Subsidios aplicados
        subsidios = calificacion.get("subsidiosAplicados", [])
        if subsidios:
            detalles += "\n\n🎁 SUBSIDIOS APLICADOS:"
            for sub in subsidios:
                if isinstance(sub, dict):
                    detalles += f"\n  • {sub.get('nombre', 'N/A')}"
        
        QMessageBox.information(self, "Detalles de Calificación", detalles)
    
    def limpiar_filtros(self):
        """Limpia todos los filtros"""
        self.date_desde.setDate(QDate.currentDate().addMonths(-3))
        self.date_hasta.setDate(QDate.currentDate())
        self.combo_tipo.setCurrentIndex(0)
        self.combo_pais.setCurrentIndex(0)
        self.combo_estado.setCurrentIndex(0)
        self.input_monto_min.clear()
        self.input_monto_max.clear()
        self.input_rut.clear()
        
        self.datos_actuales = []
        self.results_table.setRowCount(0)
        self.label_contador.setText("Total: 0 registros")
    
    def refrescar_datos(self):
        """Refresca los datos con los filtros actuales"""
        if self.datos_actuales:
            self.buscar_datos()
        else:
            QMessageBox.information(
                self,
                "Información",
                "Aplica filtros y realiza una búsqueda primero."
            )
    
    def apply_styles(self):
        """Aplica estilos consistentes con los otros módulos"""
        self.setStyleSheet("""
            /* Fondo general */
            QWidget {
                background-color: transparent;
            }
            
            /* Frames */
            QFrame {
                background-color: #ffffff;
                border: 1px solid #e6e9ee;
                border-radius: 8px;
            }
            
            /* Inputs */
            QLineEdit, QComboBox, QDateEdit {
                padding: 8px 10px;
                border: 1px solid #d7dfe8;
                border-radius: 4px;
                background-color: #ffffff;
                color: #2c3e50;
            }
            
            /* Botones según rol */
            QPushButton[role="primary"] {
                background-color: #E94E1B;
                color: white;
                padding: 10px 14px;
                border-radius: 6px;
                border: none;
                font-weight: 600;
            }
            QPushButton[role="primary"]:hover {
                background-color: #d64419;
            }
            
            QPushButton[role="secondary"] {
                background-color: #3498db;
                color: white;
                padding: 8px 12px;
                border-radius: 6px;
                border: none;
            }
            QPushButton[role="secondary"]:hover {
                background-color: #2980b9;
            }
            
            QPushButton[role="muted"] {
                background-color: #95a5a6;
                color: white;
                padding: 10px 14px;
                border-radius: 6px;
                border: none;
            }
            QPushButton[role="muted"]:hover {
                background-color: #7f8c8d;
            }
            
            /* Tablas */
            QTableWidget {
                background-color: #ffffff;
                border: 1px solid #e6e9ee;
                gridline-color: #f0f2f5;
            }
            
            QHeaderView::section {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, 
                    stop:0 #E94E1B, stop:1 #ff7a43);
                color: white;
                padding: 8px;
                font-weight: bold;
                border: none;
            }
            
            QTableWidget::item {
                padding: 6px 8px;
            }
            
            /* Labels */
            QLabel {
                color: #2c3e50;
            }
        """)