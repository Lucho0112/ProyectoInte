from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QTableWidget, QTableWidgetItem, QHeaderView, QDateEdit,
    QComboBox, QFrame, QMessageBox, QAbstractItemView,
    QScrollArea, QLineEdit, QDialog
)
from PyQt5.QtCore import Qt, QDate, pyqtSignal
from PyQt5.QtGui import QFont, QColor, QCursor
from datetime import datetime
from services.reportService import ReportService
from utils.logger import app_logger


class DetallesCalificacionDialog(QDialog):
    """Diálogo mejorado para mostrar detalles de calificación"""
    
    def __init__(self, calificacion: dict, service: ReportService, parent=None):
        super().__init__(parent)
        self.calificacion = calificacion
        self.service = service
        self.init_ui()
    
    def init_ui(self):
        """Inicializa la interfaz del diálogo"""
        self.setWindowTitle("Detalles de Calificación")
        self.setMinimumSize(700, 650)
        self.setStyleSheet("""
            QDialog {
                background-color: white;
            }
        """)
        
        layout = QVBoxLayout()
        layout.setContentsMargins(25, 25, 25, 25)
        layout.setSpacing(15)
        
        # Título
        title = QLabel("📋 DETALLES DE LA CALIFICACIÓN")
        title.setFont(QFont("Arial", 16, QFont.Bold))
        title.setStyleSheet("color: #E94E1B; padding: 10px 0;")
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)
        
        # Área de scroll para contenido
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll.setStyleSheet("""
            QScrollArea {
                border: none;
                background-color: white;
            }
        """)
        
        content_widget = QWidget()
        content_layout = QVBoxLayout()
        content_layout.setSpacing(12)
        
        # Información General
        self.add_info_section(content_layout)
        
        # Factores
        self.add_factores_section(content_layout)
        
        # Validación
        self.add_validacion_section(content_layout)
        
        # Subsidios (si existen)
        subsidios = self.calificacion.get("subsidiosAplicados", [])
        if subsidios:
            self.add_subsidios_section(content_layout, subsidios)
        
        content_widget.setLayout(content_layout)
        scroll.setWidget(content_widget)
        layout.addWidget(scroll)
        
        # Botón cerrar
        btn_cerrar = QPushButton("✖ Cerrar")
        btn_cerrar.setFont(QFont("Arial", 11, QFont.Bold))
        btn_cerrar.setMinimumHeight(45)
        btn_cerrar.setCursor(QCursor(Qt.PointingHandCursor))
        btn_cerrar.clicked.connect(self.accept)
        btn_cerrar.setStyleSheet("""
            QPushButton {
                background-color: #95a5a6;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 10px 20px;
                font-size: 12px;
            }
            QPushButton:hover {
                background-color: #7f8c8d;
            }
        """)
        layout.addWidget(btn_cerrar)
        
        self.setLayout(layout)
    
    def add_info_section(self, layout):
        """Agrega sección de información general"""
        frame = QFrame()
        frame.setStyleSheet("""
            QFrame {
                background-color: #f8f9fa;
                border: 2px solid #e6e9ee;
                border-radius: 8px;
                padding: 15px;
            }
        """)
        
        frame_layout = QVBoxLayout()
        frame_layout.setSpacing(8)
        
        # Título de sección
        section_title = QLabel("📊 Información General")
        section_title.setFont(QFont("Arial", 12, QFont.Bold))
        section_title.setStyleSheet("color: #2c3e50;")
        frame_layout.addWidget(section_title)
        
        # Datos
        rut_cliente = self.service.obtener_rut_cliente(self.calificacion.get('clienteId', ''))
        cal_id = self.calificacion.get('_id', 'N/A')[:16] + "..."
        
        datos = [
            ("🆔 ID", cal_id),
            ("👤 Cliente", rut_cliente),
            ("📅 Fecha", self.calificacion.get('fechaDeclaracion', 'N/A')),
            ("🏷️ Tipo", self.calificacion.get('tipoImpuesto', 'N/A')),
            ("🌎 País", self.calificacion.get('pais', 'N/A')),
            ("💰 Monto", f"${self.calificacion.get('montoDeclarado', 0):,.2f}"),
            ("📊 Estado", 'Local' if self.calificacion.get('esLocal', False) else 'Bolsa')
        ]
        
        for label, valor in datos:
            row_layout = QHBoxLayout()
            
            lbl = QLabel(label)
            lbl.setFont(QFont("Arial", 10, QFont.Bold))
            lbl.setStyleSheet("color: #34495e;")
            lbl.setMinimumWidth(120)
            row_layout.addWidget(lbl)
            
            val = QLabel(str(valor))
            val.setFont(QFont("Arial", 10))
            val.setStyleSheet("color: #2c3e50;")
            row_layout.addWidget(val)
            
            row_layout.addStretch()
            frame_layout.addLayout(row_layout)
        
        frame.setLayout(frame_layout)
        layout.addWidget(frame)
    
    def add_factores_section(self, layout):
        """Agrega sección de factores"""
        frame = QFrame()
        frame.setStyleSheet("""
            QFrame {
                background-color: #f8f9fa;
                border: 2px solid #e6e9ee;
                border-radius: 8px;
                padding: 15px;
            }
        """)
        
        frame_layout = QVBoxLayout()
        frame_layout.setSpacing(8)
        
        # Título de sección
        section_title = QLabel("📐 Factores Tributarios")
        section_title.setFont(QFont("Arial", 12, QFont.Bold))
        section_title.setStyleSheet("color: #2c3e50;")
        frame_layout.addWidget(section_title)
        
        # Grid de factores (2 columnas)
        factores = self.calificacion.get("factores", {})
        
        for row_idx in range(0, 19, 2):
            row_layout = QHBoxLayout()
            row_layout.setSpacing(20)
            
            for col_offset in [0, 1]:
                i = row_idx + col_offset + 1
                if i > 19:
                    break
                
                valor = factores.get(f"factor_{i}", 0)
                
                factor_label = QLabel(f"Factor {i}:")
                factor_label.setFont(QFont("Arial", 9, QFont.Bold))
                factor_label.setStyleSheet("color: #34495e;")
                factor_label.setMinimumWidth(70)
                
                valor_label = QLabel(f"{valor:.4f}")
                valor_label.setFont(QFont("Arial", 9))
                valor_label.setStyleSheet("color: #2c3e50;")
                
                # Resaltar factores 8-19
                if 8 <= i <= 19:
                    factor_label.setStyleSheet("color: #E94E1B; font-weight: bold;")
                    valor_label.setStyleSheet("color: #E94E1B;")
                
                row_layout.addWidget(factor_label)
                row_layout.addWidget(valor_label)
            
            row_layout.addStretch()
            frame_layout.addLayout(row_layout)
        
        frame.setLayout(frame_layout)
        layout.addWidget(frame)
    
    def add_validacion_section(self, layout):
        """Agrega sección de validación"""
        frame = QFrame()
        
        factores = self.calificacion.get("factores", {})
        suma_8_19 = sum(factores.get(f"factor_{i}", 0) for i in range(8, 20))
        es_valido = suma_8_19 <= 1.0
        
        if es_valido:
            frame.setStyleSheet("""
                QFrame {
                    background-color: #d4edda;
                    border: 2px solid #28a745;
                    border-radius: 8px;
                    padding: 15px;
                }
            """)
        else:
            frame.setStyleSheet("""
                QFrame {
                    background-color: #f8d7da;
                    border: 2px solid #dc3545;
                    border-radius: 8px;
                    padding: 15px;
                }
            """)
        
        frame_layout = QVBoxLayout()
        
        # Título
        title_text = "✅ VALIDACIÓN EXITOSA" if es_valido else "❌ VALIDACIÓN FALLIDA"
        title = QLabel(title_text)
        title.setFont(QFont("Arial", 12, QFont.Bold))
        title.setStyleSheet("color: #28a745;" if es_valido else "color: #dc3545;")
        title.setAlignment(Qt.AlignCenter)
        frame_layout.addWidget(title)
        
        # Suma
        suma_label = QLabel(f"Suma Factores 8-19: {suma_8_19:.4f}")
        suma_label.setFont(QFont("Arial", 11, QFont.Bold))
        suma_label.setStyleSheet("color: #155724;" if es_valido else "color: #721c24;")
        suma_label.setAlignment(Qt.AlignCenter)
        frame_layout.addWidget(suma_label)
        
        # Mensaje
        if es_valido:
            msg = QLabel("La suma de factores es válida (≤ 1.0)")
        else:
            msg = QLabel(f"La suma de factores excede el límite permitido (> 1.0)")
        msg.setFont(QFont("Arial", 10))
        msg.setStyleSheet("color: #155724;" if es_valido else "color: #721c24;")
        msg.setAlignment(Qt.AlignCenter)
        msg.setWordWrap(True)
        frame_layout.addWidget(msg)
        
        frame.setLayout(frame_layout)
        layout.addWidget(frame)
    
    def add_subsidios_section(self, layout, subsidios):
        """Agrega sección de subsidios"""
        frame = QFrame()
        frame.setStyleSheet("""
            QFrame {
                background-color: #fff3cd;
                border: 2px solid #ffc107;
                border-radius: 8px;
                padding: 15px;
            }
        """)
        
        frame_layout = QVBoxLayout()
        frame_layout.setSpacing(8)
        
        # Título de sección
        section_title = QLabel("🎁 Subsidios Aplicados")
        section_title.setFont(QFont("Arial", 12, QFont.Bold))
        section_title.setStyleSheet("color: #856404;")
        frame_layout.addWidget(section_title)
        
        for sub in subsidios:
            if isinstance(sub, dict):
                sub_label = QLabel(f"• {sub.get('nombre', 'N/A')}")
                sub_label.setFont(QFont("Arial", 10))
                sub_label.setStyleSheet("color: #856404;")
                frame_layout.addWidget(sub_label)
        
        frame.setLayout(frame_layout)
        layout.addWidget(frame)


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
        back_button.setFont(QFont("Arial", 12))
        back_button.setCursor(QCursor(Qt.PointingHandCursor))
        back_button.clicked.connect(self.back_requested.emit)
        back_button.setStyleSheet("""
            QPushButton {
                background-color: transparent;
                border: none;
                color: #3498db;
                padding: 5px;
                font-size: 18px;
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
                border: 2px solid #dee2e6;
                border-radius: 10px;
                padding: 20px;
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
        fecha_label = QLabel("📅 Fecha desde:")
        fecha_label.setFont(QFont("Arial", 10, QFont.Bold))
        fecha_layout.addWidget(fecha_label)
        self.date_desde = QDateEdit()
        self.date_desde.setCalendarPopup(True)
        self.date_desde.setDate(QDate.currentDate().addMonths(-12))
        self.date_desde.setDisplayFormat("dd/MM/yyyy")
        self.date_desde.setMinimumHeight(38)
        fecha_layout.addWidget(self.date_desde)
        grid_layout1.addLayout(fecha_layout)
        
        fecha_layout2 = QVBoxLayout()
        fecha_layout2.setSpacing(8)
        fecha_label2 = QLabel("📅 Fecha hasta:")
        fecha_label2.setFont(QFont("Arial", 10, QFont.Bold))
        fecha_layout2.addWidget(fecha_label2)
        self.date_hasta = QDateEdit()
        self.date_hasta.setCalendarPopup(True)
        self.date_hasta.setDate(QDate.currentDate())
        self.date_hasta.setDisplayFormat("dd/MM/yyyy")
        self.date_hasta.setMinimumHeight(38)
        fecha_layout2.addWidget(self.date_hasta)
        grid_layout1.addLayout(fecha_layout2)
        
        # Tipo de Impuesto
        tipo_layout = QVBoxLayout()
        tipo_layout.setSpacing(8)
        tipo_label = QLabel("📋 Tipo de Impuesto:")
        tipo_label.setFont(QFont("Arial", 10, QFont.Bold))
        tipo_layout.addWidget(tipo_label)
        self.combo_tipo = QComboBox()
        self.combo_tipo.addItems(["Todos", "IVA", "Renta", "Importación", "Exportación", "Otro"])
        self.combo_tipo.setMinimumHeight(38)
        tipo_layout.addWidget(self.combo_tipo)
        grid_layout1.addLayout(tipo_layout)
        
        # País
        pais_layout = QVBoxLayout()
        pais_layout.setSpacing(8)
        pais_label = QLabel("🌎 País:")
        pais_label.setFont(QFont("Arial", 10, QFont.Bold))
        pais_layout.addWidget(pais_label)
        self.combo_pais = QComboBox()
        self.combo_pais.addItems(["Todos", "Chile", "Perú", "Colombia"])
        self.combo_pais.setMinimumHeight(38)
        pais_layout.addWidget(self.combo_pais)
        grid_layout1.addLayout(pais_layout)
        
        filter_layout.addLayout(grid_layout1)
        
        # Grid de filtros - Fila 2
        grid_layout2 = QHBoxLayout()
        grid_layout2.setSpacing(15)
        
        # Monto mínimo
        monto_min_layout = QVBoxLayout()
        monto_min_layout.setSpacing(8)
        monto_min_label = QLabel("💵 Monto mínimo ($):")
        monto_min_label.setFont(QFont("Arial", 10, QFont.Bold))
        monto_min_layout.addWidget(monto_min_label)
        self.input_monto_min = QLineEdit()
        self.input_monto_min.setPlaceholderText("Ej: 1000")
        self.input_monto_min.setMinimumHeight(38)
        monto_min_layout.addWidget(self.input_monto_min)
        grid_layout2.addLayout(monto_min_layout)
        
        # Monto máximo
        monto_max_layout = QVBoxLayout()
        monto_max_layout.setSpacing(8)
        monto_max_label = QLabel("💵 Monto máximo ($):")
        monto_max_label.setFont(QFont("Arial", 10, QFont.Bold))
        monto_max_layout.addWidget(monto_max_label)
        self.input_monto_max = QLineEdit()
        self.input_monto_max.setPlaceholderText("Ej: 100000")
        self.input_monto_max.setMinimumHeight(38)
        monto_max_layout.addWidget(self.input_monto_max)
        grid_layout2.addLayout(monto_max_layout)
        
        # RUT Cliente
        rut_layout = QVBoxLayout()
        rut_layout.setSpacing(8)
        rut_label = QLabel("🔖 RUT Cliente:")
        rut_label.setFont(QFont("Arial", 10, QFont.Bold))
        rut_layout.addWidget(rut_label)
        self.input_rut = QLineEdit()
        self.input_rut.setPlaceholderText("Ej: 12345678-9")
        self.input_rut.setMinimumHeight(38)
        rut_layout.addWidget(self.input_rut)
        grid_layout2.addLayout(rut_layout)
        
        # Estado
        estado_layout = QVBoxLayout()
        estado_layout.setSpacing(8)
        estado_label = QLabel("📊 Estado:")
        estado_label.setFont(QFont("Arial", 10, QFont.Bold))
        estado_layout.addWidget(estado_label)
        self.combo_estado = QComboBox()
        self.combo_estado.addItems(["Todos", "Local", "Bolsa"])
        self.combo_estado.setMinimumHeight(38)
        estado_layout.addWidget(self.combo_estado)
        grid_layout2.addLayout(estado_layout)
        
        filter_layout.addLayout(grid_layout2)
        
        filter_frame.setLayout(filter_layout)
        layout.addWidget(filter_frame)
    
    def add_results_table(self, layout):
        """Tabla de resultados CON BARRA NARANJA"""
        results_frame = QFrame()
        results_frame.setStyleSheet("""
            QFrame {
                background-color: white;
                border: 2px solid #dee2e6;
                border-radius: 10px;
                padding: 20px;
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
        
        # Vista previa
        preview_label = QLabel("Vista Previa (primeros 50 registros)")
        preview_label.setFont(QFont("Arial", 9))
        preview_label.setStyleSheet("color: #7f8c8d; padding: 5px 0;")
        results_layout.addWidget(preview_label)
        
        # ← BARRA NARANJA PERSONALIZADA
        header_bar = QFrame()
        header_bar.setFixedHeight(40)
        header_bar.setStyleSheet("""
            QFrame {
                background-color: #E94E1B;
                border-radius: 5px 5px 0 0;
            }
        """)
        
        header_layout = QHBoxLayout(header_bar)
        header_layout.setContentsMargins(5, 0, 5, 0)
        header_layout.setSpacing(0)
        
        # Títulos con anchos proporcionales
        columnas = [
            ("RUT Cliente", 140),
            ("Fecha", 100),
            ("Tipo", 100),
            ("País", 100),
            ("Monto", 130),
            ("Suma 8-19", 90),
            ("Estado", 80),
            ("Válido", 80),
            ("Ver", 80)
        ]
        
        for titulo, ancho in columnas:
            lbl = QLabel(titulo)
            lbl.setFont(QFont("Arial", 10, QFont.Bold))
            lbl.setStyleSheet("color: white; padding: 10px 5px;")
            lbl.setAlignment(Qt.AlignCenter)
            lbl.setFixedWidth(ancho)
            header_layout.addWidget(lbl)
        
        header_layout.addStretch()
        results_layout.addWidget(header_bar)
        
        # Tabla SIN header nativo
        self.results_table = QTableWidget()
        self.results_table.setColumnCount(10)
        self.results_table.setHorizontalHeaderLabels([
            "ID", "RUT", "Fecha", "Tipo", "País", "Monto", "Suma", "Estado", "Válido", "Ver"
        ])
        
        # OCULTAR header nativo
        self.results_table.horizontalHeader().setVisible(False)
        self.results_table.verticalHeader().setVisible(False)
        
        self.results_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.results_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.results_table.setAlternatingRowColors(True)
        self.results_table.setMinimumHeight(450)
        
        # Ocultar columna ID
        self.results_table.setColumnHidden(0, True)
        
        # ← ANCHOS EXACTOS para alinear con header naranja
        self.results_table.setColumnWidth(1, 140)  # RUT
        self.results_table.setColumnWidth(2, 100)  # Fecha
        self.results_table.setColumnWidth(3, 100)  # Tipo
        self.results_table.setColumnWidth(4, 100)  # País
        self.results_table.setColumnWidth(5, 130)  # Monto
        self.results_table.setColumnWidth(6, 90)   # Suma
        self.results_table.setColumnWidth(7, 80)   # Estado
        self.results_table.setColumnWidth(8, 80)   # Válido
        self.results_table.setColumnWidth(9, 80)   # Ver
        
        # Estilo tabla
        self.results_table.setStyleSheet("""
            QTableWidget {
                background-color: white;
                border: none;
                border-top: none;
                gridline-color: #e6e9ee;
            }
            QTableWidget::item {
                padding: 8px 5px;
                border-bottom: 1px solid #f0f0f0;
                font-size: 9pt;
            }
            QTableWidget::item:selected {
                background-color: #E8F4F8;
                color: #2c3e50;
            }
            QTableWidget::item:hover {
                background-color: #f8f9fa;
            }
        """)
        
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
        footer.setFont(QFont("Arial", 10))
        footer.setStyleSheet("color: #7f8c8d; padding: 12px; background-color: #ecf0f1; border-radius: 8px;")
        layout.addWidget(footer)
    
    def buscar_datos(self):
        """Realiza la búsqueda con los filtros aplicados"""
        filtros = self.obtener_filtros()
        
        try:
            datos_brutos = self.service.obtener_datos_filtrados(
                filtros,
                self.user_data.get("_id"),
                self.user_rol
            )
            
            self.datos_actuales = self.aplicar_filtros_locales(datos_brutos, filtros)
            
            if not self.datos_actuales:
                QMessageBox.information(
                    self,
                    "Sin resultados",
                    "No se encontraron datos con los criterios especificados.\n\n"
                    "💡 Sugerencia: Intenta ampliar el rango de fechas o reducir los filtros."
                )
                self.results_table.setRowCount(0)
                self.label_contador.setText("Total: 0 registros")
                return
            
            self.actualizar_tabla(self.datos_actuales)
            self.label_contador.setText(f"Total: {len(self.datos_actuales)} registros")
            
            QMessageBox.information(
                self,
                "Búsqueda completada",
                f"✅ Se encontraron {len(self.datos_actuales)} registros que coinciden con los criterios."
            )
            
        except Exception as e:
            app_logger.error(f"Error en búsqueda: {str(e)}", exc_info=True)
            QMessageBox.warning(
                self,
                "Error",
                f"❌ Ocurrió un error al buscar los datos:\n{str(e)}"
            )
    
    def aplicar_filtros_locales(self, datos: list, filtros: dict) -> list:
        """Aplica filtros adicionales en el frontend"""
        datos_filtrados = []
        
        for dato in datos:
            if "monto_minimo" in filtros:
                if dato.get("montoDeclarado", 0) < filtros["monto_minimo"]:
                    continue
            
            if "monto_maximo" in filtros:
                if dato.get("montoDeclarado", 0) > filtros["monto_maximo"]:
                    continue
            
            if "rut_cliente" in filtros and filtros["rut_cliente"]:
                rut_actual = self.service.obtener_rut_cliente(dato.get("clienteId", ""))
                rut_filtro = filtros["rut_cliente"].replace(".", "").replace("-", "").strip().upper()
                rut_actual_norm = rut_actual.replace(".", "").replace("-", "").strip().upper()
                
                if rut_actual_norm != rut_filtro and rut_actual != "N/A":
                    continue
            
            datos_filtrados.append(dato)
        
        return datos_filtrados
    
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
        
        estado_texto = self.combo_estado.currentText()
        if estado_texto == "Local":
            filtros["estado"] = "local"
        elif estado_texto == "Bolsa":
            filtros["estado"] = "bolsa"
        else:
            filtros["estado"] = "ambos"
        
        monto_min = self.input_monto_min.text().strip()
        if monto_min:
            try:
                filtros["monto_minimo"] = max(0, float(monto_min))
            except (ValueError, TypeError):
                app_logger.warning(f"Monto mínimo inválido: {monto_min}")
        
        monto_max = self.input_monto_max.text().strip()
        if monto_max:
            try:
                filtros["monto_maximo"] = max(0, float(monto_max))
            except (ValueError, TypeError):
                app_logger.warning(f"Monto máximo inválido: {monto_max}")
        
        rut = self.input_rut.text().strip()
        if rut:
            filtros["rut_cliente"] = rut
        
        return filtros
    
    def actualizar_tabla(self, datos: list):
        """Actualiza la tabla de resultados"""
        self.results_table.setRowCount(len(datos))
        
        for row, cal in enumerate(datos):
            # Col 0: ID (oculto)
            self.results_table.setItem(row, 0, QTableWidgetItem(cal.get("_id", "")))
            
            # Col 1: RUT Cliente
            rut = self.service.obtener_rut_cliente(cal.get("clienteId", ""))
            rut_item = QTableWidgetItem(rut)
            rut_item.setTextAlignment(Qt.AlignCenter)
            rut_item.setFont(QFont("Arial", 9))
            self.results_table.setItem(row, 1, rut_item)
            
            # Col 2: Fecha
            fecha_item = QTableWidgetItem(cal.get("fechaDeclaracion", ""))
            fecha_item.setTextAlignment(Qt.AlignCenter)
            fecha_item.setFont(QFont("Arial", 9))
            self.results_table.setItem(row, 2, fecha_item)
            
            # Col 3: Tipo
            tipo_item = QTableWidgetItem(cal.get("tipoImpuesto", ""))
            tipo_item.setTextAlignment(Qt.AlignCenter)
            tipo_item.setFont(QFont("Arial", 9))
            self.results_table.setItem(row, 3, tipo_item)
            
            # Col 4: País
            pais_item = QTableWidgetItem(cal.get("pais", ""))
            pais_item.setTextAlignment(Qt.AlignCenter)
            pais_item.setFont(QFont("Arial", 9))
            self.results_table.setItem(row, 4, pais_item)
            
            # Col 5: Monto
            monto = cal.get("montoDeclarado", 0)
            monto_item = QTableWidgetItem(f"${monto:,.2f}")
            monto_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            monto_item.setFont(QFont("Arial", 9))
            self.results_table.setItem(row, 5, monto_item)
            
            # Col 6: Suma 8-19
            factores = cal.get("factores", {})
            suma_8_19 = sum(factores.get(f"factor_{i}", 0) for i in range(8, 20))
            suma_item = QTableWidgetItem(f"{suma_8_19:.4f}")
            suma_item.setTextAlignment(Qt.AlignCenter)
            suma_item.setFont(QFont("Arial", 9))
            if suma_8_19 > 1.0:
                suma_item.setForeground(QColor("#e74c3c"))
                suma_item.setFont(QFont("Arial", 9, QFont.Bold))
            self.results_table.setItem(row, 6, suma_item)
            
            # Col 7: Estado
            estado = "Local" if cal.get("esLocal", False) else "Bolsa"
            estado_item = QTableWidgetItem(estado)
            estado_item.setTextAlignment(Qt.AlignCenter)
            estado_item.setFont(QFont("Arial", 9))
            if estado == "Local":
                estado_item.setForeground(QColor("#3498db"))
            else:
                estado_item.setForeground(QColor("#27ae60"))
            self.results_table.setItem(row, 7, estado_item)
            
            # Col 8: Válido
            valido = "✅ Sí" if suma_8_19 <= 1.0 else "❌ No"
            valido_item = QTableWidgetItem(valido)
            valido_item.setTextAlignment(Qt.AlignCenter)
            valido_item.setFont(QFont("Arial", 9))
            if suma_8_19 <= 1.0:
                valido_item.setForeground(QColor("#27ae60"))
            else:
                valido_item.setForeground(QColor("#e74c3c"))
            self.results_table.setItem(row, 8, valido_item)
            
            # Col 9: Botón Ver (CENTRADO)
            btn_widget = QWidget()
            btn_layout = QHBoxLayout(btn_widget)
            btn_layout.setContentsMargins(5, 3, 5, 3)
            btn_layout.setAlignment(Qt.AlignCenter)
            
            btn = QPushButton("👁 Ver")
            btn.setFont(QFont("Arial", 8, QFont.Bold))
            btn.setFixedSize(60, 24)
            btn.setCursor(QCursor(Qt.PointingHandCursor))
            btn.setStyleSheet("""
                QPushButton {
                    background-color: #3498db;
                    color: white;
                    border: none;
                    border-radius: 3px;
                }
                QPushButton:hover {
                    background-color: #2980b9;
                }
            """)
            btn.clicked.connect(lambda checked, c=cal: self.ver_detalles(c))
            
            btn_layout.addWidget(btn)
            self.results_table.setCellWidget(row, 9, btn_widget)
    
    def ver_detalles(self, calificacion: dict):
        """Muestra diálogo de detalles"""
        dialog = DetallesCalificacionDialog(calificacion, self.service, self)
        dialog.exec_()
    
    def limpiar_filtros(self):
        """Limpia todos los filtros"""
        self.date_desde.setDate(QDate.currentDate().addMonths(-12))
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
        
        QMessageBox.information(
            self,
            "Filtros limpiados",
            "✅ Todos los filtros han sido restablecidos."
        )
    
    def refrescar_datos(self):
        """Refresca los datos con los filtros actuales"""
        if self.datos_actuales:
            self.buscar_datos()
        else:
            QMessageBox.information(
                self,
                "Información",
                "💡 Aplica filtros y realiza una búsqueda primero."
            )
    
    def apply_styles(self):
        """Aplica estilos consistentes"""
        self.setStyleSheet("""
            QWidget {
                background-color: #f5f6fa;
                font-family: 'Segoe UI', Arial, sans-serif;
            }
            
            QLineEdit, QDateEdit, QComboBox {
                border: 2px solid #e6e9ee;
                border-radius: 6px;
                padding: 8px;
                background-color: white;
                font-size: 11px;
            }
            
            QLineEdit:focus, QDateEdit:focus, QComboBox:focus {
                border-color: #E94E1B;
            }
            
            QPushButton[role="primary"] {
                background-color: #E94E1B;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 10px 20px;
                font-weight: bold;
                font-size: 11px;
            }
            
            QPushButton[role="primary"]:hover {
                background-color: #d13d0f;
            }
            
            QPushButton[role="secondary"] {
                background-color: #3498db;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 10px 20px;
                font-weight: bold;
                font-size: 11px;
            }
            
            QPushButton[role="secondary"]:hover {
                background-color: #2980b9;
            }
            
            QPushButton[role="muted"] {
                background-color: #95a5a6;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 10px 20px;
                font-weight: bold;
                font-size: 11px;
            }
            
            QPushButton[role="muted"]:hover {
                background-color: #7f8c8d;
            }
            
            QLabel {
                color: #2c3e50;
            }
        """)