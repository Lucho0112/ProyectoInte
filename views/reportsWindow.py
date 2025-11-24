from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QTableWidget, QTableWidgetItem, QHeaderView, QDateEdit,
    QComboBox, QFrame, QMessageBox, QAbstractItemView,
    QScrollArea, QRadioButton, QButtonGroup, QFileDialog,
    QProgressDialog, QLineEdit
)
from PyQt5.QtCore import Qt, QDate, pyqtSignal, QThread
from PyQt5.QtGui import QFont, QColor, QCursor
from datetime import datetime
from services.reportService import ReportService
from utils.logger import app_logger
import traceback


class ExportWorker(QThread):
    """Worker thread para exportación en segundo plano"""
    progress = pyqtSignal(int)
    finished = pyqtSignal(dict)
    
    def __init__(self, service, method, *args, **kwargs):
        super().__init__()
        self.service = service
        self.method = method
        self.args = args
        self.kwargs = kwargs
    
    def run(self):
        """Ejecuta la exportación"""
        try:
            result = self.method(*self.args, **self.kwargs, progress_callback=self.emit_progress)
            self.finished.emit(result)
        except Exception as e:
            app_logger.error(f"Error en worker de exportación: {str(e)}\n{traceback.format_exc()}")
            self.finished.emit({"success": False, "message": str(e)})
    
    def emit_progress(self, value):
        """Emite el progreso"""
        self.progress.emit(value)


class GenerarReportesContent(QWidget):
    """Contenido de generación de reportes tributarios"""
    back_requested = pyqtSignal()
    
    def __init__(self, user_data: dict, parent=None):
        super().__init__(parent)
        self.user_data = user_data
        self.user_rol = user_data.get("rol", "cliente")
        self.service = ReportService()
        self.datos_actuales = []
        self.export_worker = None
        
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
        
        # Vista previa
        self.add_preview(main_layout)
        
        # Botones de exportación
        self.add_export_buttons(main_layout)
        
        # Historial (solo admin y auditores)
        if self.user_rol in ["administrador", "auditor_tributario"]:
            self.add_history(main_layout)
        
        # Footer
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
        
        title = QLabel("📊 Generar Reportes y Exportaciones")
        title.setFont(QFont("Arial", 18, QFont.Bold))
        title.setStyleSheet("color: #2c3e50;")
        header_layout.addWidget(title)
        
        header_layout.addStretch()
        layout.addLayout(header_layout)
    
    def add_toolbar(self, layout):
        """Barra de herramientas"""
        toolbar_layout = QHBoxLayout()
        toolbar_layout.setSpacing(10)
        
        btn_aplicar = QPushButton("🔍 Aplicar Filtros")
        btn_aplicar.setFont(QFont("Arial", 10, QFont.Bold))
        btn_aplicar.setMinimumHeight(40)
        btn_aplicar.setCursor(QCursor(Qt.PointingHandCursor))
        btn_aplicar.clicked.connect(self.aplicar_filtros)
        btn_aplicar.setProperty("role", "primary")
        toolbar_layout.addWidget(btn_aplicar)
        
        btn_limpiar = QPushButton("🗑️ Limpiar Filtros")
        btn_limpiar.setFont(QFont("Arial", 10))
        btn_limpiar.setMinimumHeight(40)
        btn_limpiar.setCursor(QCursor(Qt.PointingHandCursor))
        btn_limpiar.clicked.connect(self.limpiar_filtros)
        btn_limpiar.setProperty("role", "muted")
        toolbar_layout.addWidget(btn_limpiar)
        
        toolbar_layout.addStretch()
        
        self.label_contador = QLabel("Total: 0 registros")
        self.label_contador.setFont(QFont("Arial", 11, QFont.Bold))
        self.label_contador.setStyleSheet("color: #2c3e50;")
        toolbar_layout.addWidget(self.label_contador)
        
        layout.addLayout(toolbar_layout)
    
    def add_filters(self, layout):
        """Panel de filtros"""
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
        
        # Columna 3: Estado y RUT
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
        
        # RUT Cliente (opcional)
        col3.addWidget(QLabel("RUT Cliente (opcional):"))
        self.input_rut_filtro = QLineEdit()
        self.input_rut_filtro.setPlaceholderText("Ej: 12345678-9")
        self.input_rut_filtro.setMinimumHeight(35)
        col3.addWidget(self.input_rut_filtro)
        
        col3.addStretch()
        grid_layout.addLayout(col3)
        
        filter_layout.addLayout(grid_layout)
        filter_frame.setLayout(filter_layout)
        layout.addWidget(filter_frame)
    
    def add_preview(self, layout):
        """Vista previa de datos"""
        preview_frame = QFrame()
        preview_frame.setStyleSheet("""
            QFrame {
                background-color: white;
                border: 1px solid #e6e9ee;
                border-radius: 8px;
                padding: 15px;
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
        
        self.preview_table.setMaximumHeight(400)
        preview_layout.addWidget(self.preview_table)
        
        preview_frame.setLayout(preview_layout)
        layout.addWidget(preview_frame)
    
    def add_export_buttons(self, layout):
        """Botones de exportación"""
        export_frame = QFrame()
        export_frame.setStyleSheet("""
            QFrame {
                background-color: white;
                border: 1px solid #e6e9ee;
                border-radius: 8px;
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
        self.btn_csv.setProperty("role", "success")
        export_layout.addWidget(self.btn_csv)
        
        self.btn_excel = QPushButton("📊 Exportar Excel")
        self.btn_excel.setFont(QFont("Arial", 11, QFont.Bold))
        self.btn_excel.setMinimumSize(180, 50)
        self.btn_excel.setCursor(QCursor(Qt.PointingHandCursor))
        self.btn_excel.clicked.connect(self.exportar_excel)
        self.btn_excel.setEnabled(False)
        self.btn_excel.setProperty("role", "secondary")
        export_layout.addWidget(self.btn_excel)
        
        export_frame.setLayout(export_layout)
        layout.addWidget(export_frame)
    
    def add_history(self, layout):
        """Historial de reportes (solo admin/auditor)"""
        history_frame = QFrame()
        history_frame.setStyleSheet("""
            QFrame {
                background-color: white;
                border: 1px solid #e6e9ee;
                border-radius: 8px;
                padding: 15px;
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
        
        header = self.history_table.horizontalHeader()
        for i in range(5):
            header.setSectionResizeMode(i, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        
        history_layout.addWidget(self.history_table)
        
        # Botón refrescar
        btn_refrescar = QPushButton("🔄 Refrescar Historial")
        btn_refrescar.setMinimumHeight(35)
        btn_refrescar.setCursor(QCursor(Qt.PointingHandCursor))
        btn_refrescar.clicked.connect(self.cargar_historial)
        btn_refrescar.setProperty("role", "secondary")
        history_layout.addWidget(btn_refrescar)
        
        history_frame.setLayout(history_layout)
        layout.addWidget(history_frame)
        
        # Cargar historial inicial
        self.cargar_historial()
    
    def add_footer(self, layout):
        """Footer informativo"""
        if self.user_rol == "administrador":
            footer_text = "💡 Admin: Puedes exportar TODOS los datos del sistema y ver el historial completo"
        elif self.user_rol == "auditor_tributario":
            footer_text = "💡 Auditor: Puedes exportar datos de bolsa y tus datos locales, y ver el historial completo"
        else:
            footer_text = "💡 Tip: Puedes exportar datos de bolsa y tus propios datos locales"
        
        footer = QLabel(footer_text)
        footer.setFont(QFont("Arial", 9))
        footer.setStyleSheet("color: #7f8c8d;")
        layout.addWidget(footer)
    
    def aplicar_filtros(self):
        """Aplica los filtros y carga datos"""
        try:
            # ✅ CORRECCIÓN: Obtener filtros con validación
            filtros = self.obtener_filtros()
            
            # Log para debug
            app_logger.info(f"Filtros obtenidos: {filtros}")
            
            # Limpiar caché del servicio
            self.service.limpiar_cache()
            
            # Obtener datos
            self.datos_actuales = self.service.obtener_datos_filtrados(
                filtros,
                self.user_data.get("_id"),
                self.user_rol
            )
            
            if not self.datos_actuales:
                msg = QMessageBox(self)
                msg.setIcon(QMessageBox.Information)
                msg.setWindowTitle("Sin resultados")
                msg.setText("No se encontraron datos con los filtros aplicados.")
                msg.setStyleSheet("""
                    QMessageBox {
                        background-color: white;
                    }
                    QMessageBox QLabel {
                        color: #2c3e50;
                    }
                """)
                msg.exec_()
                
                self.btn_csv.setEnabled(False)
                self.btn_excel.setEnabled(False)
                self.preview_table.setRowCount(0)
                self.label_contador.setText("Total: 0 registros")
                return
            
            # Advertencia para grandes volúmenes
            if len(self.datos_actuales) > 5000:
                msg = QMessageBox(self)
                msg.setIcon(QMessageBox.Warning)
                msg.setWindowTitle("Gran volumen de datos")
                msg.setText(
                    f"Se encontraron {len(self.datos_actuales)} registros.\n\n"
                    f"La exportación puede tardar varios minutos.\n\n"
                    f"¿Desea continuar?"
                )
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
                """)
                
                if msg.exec_() != QMessageBox.Yes:
                    return
            
            # Actualizar vista previa (primeras 50)
            preview_data = self.datos_actuales[:self.service.MAX_PREVIEW]
            self.actualizar_vista_previa(preview_data)
            
            # Habilitar exportación
            self.btn_csv.setEnabled(True)
            self.btn_excel.setEnabled(True)
            
            total = len(self.datos_actuales)
            if total >= self.service.MAX_RECORDS:
                self.label_contador.setText(
                    f"Total: {total} registros (máximo alcanzado)"
                )
            else:
                self.label_contador.setText(f"Total: {total} registros")
        
        except Exception as e:
            app_logger.error(f"Error al aplicar filtros: {str(e)}\n{traceback.format_exc()}")
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Critical)
            msg.setWindowTitle("Error")
            msg.setText(f"Error al aplicar filtros:\n\n{str(e)}")
            msg.setStyleSheet("""
                QMessageBox {
                    background-color: white;
                }
                QMessageBox QLabel {
                    color: #2c3e50;
                }
            """)
            msg.exec_()
    
    def obtener_filtros(self) -> dict:
        """Obtiene los filtros actuales - SOLO valores simples (str, int, float, date)"""
        try:
            filtros = {}
            
            # ✅ Fechas - convertir a date de Python
            filtros["fecha_desde"] = self.date_desde.date().toPyDate()
            filtros["fecha_hasta"] = self.date_hasta.date().toPyDate()
            
            # ✅ Tipo de impuesto - solo string
            tipo_texto = self.combo_tipo.currentText()
            if tipo_texto != "Todos":
                filtros["tipo_impuesto"] = str(tipo_texto)
            
            # ✅ País - solo string
            pais_texto = self.combo_pais.currentText()
            if pais_texto != "Todos":
                filtros["pais"] = str(pais_texto)
            
            # ✅ RUT Cliente - solo string
            rut_filtro = self.input_rut_filtro.text().strip()
            if rut_filtro:
                filtros["rut_cliente"] = str(rut_filtro)
            
            # ✅ Estado - solo string
            if self.radio_local.isChecked():
                filtros["estado"] = "local"
            elif self.radio_bolsa.isChecked():
                filtros["estado"] = "bolsa"
            else:
                filtros["estado"] = "ambos"
            
            return filtros
            
        except Exception as e:
            app_logger.error(f"Error al obtener filtros: {str(e)}\n{traceback.format_exc()}")
            # Retornar filtros mínimos en caso de error
            return {
                "fecha_desde": datetime.now().date(),
                "fecha_hasta": datetime.now().date(),
                "estado": "ambos"
            }
    
    def actualizar_vista_previa(self, datos: list):
        
        """Actualiza la tabla de vista previa"""
        try:
            self.preview_table.setRowCount(len(datos))
            
            for row, cal in enumerate(datos):
                try:
                    # RUT Cliente
                    rut = self.service.obtener_rut_cliente(cal.get("clienteId", ""))
                    self.preview_table.setItem(row, 0, QTableWidgetItem(str(rut)))
                    
                    # Fecha
                    self.preview_table.setItem(row, 1, QTableWidgetItem(str(cal.get("fechaDeclaracion", ""))))
                    
                    # Tipo
                    self.preview_table.setItem(row, 2, QTableWidgetItem(str(cal.get("tipoImpuesto", ""))))
                    
                    # País
                    self.preview_table.setItem(row, 3, QTableWidgetItem(str(cal.get("pais", ""))))
                    
                    # Monto
                    monto = cal.get("montoDeclarado", 0)
                    try:
                        monto_float = float(monto)
                        item_monto = QTableWidgetItem(f"${monto_float:,.2f}")
                    except (ValueError, TypeError):
                        item_monto = QTableWidgetItem("$0.00")
                    item_monto.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                    self.preview_table.setItem(row, 4, item_monto)
                    
                    # ✅ CORRECCIÓN CRÍTICA: Suma 8-19 con manejo robusto de factores
                    factores = cal.get("factores", {})
                    suma = 0.0
                    
                    # Verificar si factores es un diccionario
                    if isinstance(factores, dict):
                        try:
                            suma = sum(float(factores.get(f"factor_{i}", 0)) for i in range(8, 20))
                        except (ValueError, TypeError) as e:
                            app_logger.warning(f"Error al sumar factores (dict) en fila {row}: {e}")
                            suma = 0.0
                    # Si factores es una lista (problema común)
                    elif isinstance(factores, (list, tuple)):
                        try:
                            # Si es lista de 19 elementos, sumar índices 7-18 (factores 8-19)
                            if len(factores) >= 19:
                                suma = sum(float(factores[i]) for i in range(7, 19))
                            else:
                                app_logger.warning(f"Lista de factores incompleta en fila {row}: {len(factores)} elementos")
                                suma = 0.0
                        except (ValueError, TypeError, IndexError) as e:
                            app_logger.warning(f"Error al sumar factores (list) en fila {row}: {e}")
                            suma = 0.0
                    else:
                        app_logger.warning(f"Tipo de factores no soportado en fila {row}: {type(factores)}")
                        suma = 0.0
                    
                    item_suma = QTableWidgetItem(f"{suma:.4f}")
                    item_suma.setTextAlignment(Qt.AlignCenter)
                    
                    if suma > 1.0:
                        item_suma.setBackground(QColor(255, 200, 200))
                        item_suma.setForeground(QColor(200, 0, 0))
                    else:
                        item_suma.setBackground(QColor(200, 255, 200))
                        item_suma.setForeground(QColor(0, 150, 0))
                    
                    self.preview_table.setItem(row, 5, item_suma)
                    
                    # Estado
                    estado = "Local" if cal.get("esLocal", False) else "Bolsa"
                    item_estado = QTableWidgetItem(estado)
                    item_estado.setTextAlignment(Qt.AlignCenter)
                    self.preview_table.setItem(row, 6, item_estado)
                    
                    # Válido
                    valido = "✅ Sí" if suma <= 1.0 else "❌ No"
                    self.preview_table.setItem(row, 7, QTableWidgetItem(valido))
                    
                except Exception as e:
                    app_logger.error(f"Error al procesar fila {row}: {str(e)}")
                    # Llenar con valores por defecto para esta fila
                    for col in range(8):
                        if self.preview_table.item(row, col) is None:
                            self.preview_table.setItem(row, col, QTableWidgetItem("N/A"))
                    continue
                    
        except Exception as e:
            app_logger.error(f"Error al actualizar vista previa: {str(e)}\n{traceback.format_exc()}")
    
    def exportar_csv(self):
        """Exporta datos a CSV con barra de progreso"""
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
        
        # Crear barra de progreso
        progress = QProgressDialog(
            "Generando reporte CSV...",
            "Cancelar",
            0, 100, self
        )
        progress.setWindowTitle("Exportando")
        progress.setWindowModality(Qt.WindowModal)
        progress.setMinimumDuration(0)
        progress.setValue(0)
        
        # Deshabilitar botones
        self.btn_csv.setEnabled(False)
        self.btn_excel.setEnabled(False)
        
        # Crear worker
        filtros = self.obtener_filtros()
        self.export_worker = ExportWorker(
            self.service,
            self.service.exportar_csv,
            file_path,
            self.datos_actuales,
            filtros,
            self.user_data.get("_id"),
            self.user_rol
        )
        
        # Conectar señales
        self.export_worker.progress.connect(progress.setValue)
        self.export_worker.finished.connect(
            lambda result: self.on_export_finished(result, progress, "CSV")
        )
        
        # Iniciar exportación
        self.export_worker.start()
    
    def exportar_excel(self):
        """Exporta datos a Excel con barra de progreso"""
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
        
        # Crear barra de progreso
        progress = QProgressDialog(
            "Generando reporte Excel...",
            "Cancelar",
            0, 100, self
        )
        progress.setWindowTitle("Exportando")
        progress.setWindowModality(Qt.WindowModal)
        progress.setMinimumDuration(0)
        progress.setValue(0)
        
        # Deshabilitar botones
        self.btn_csv.setEnabled(False)
        self.btn_excel.setEnabled(False)
        
        # Crear worker
        filtros = self.obtener_filtros()
        self.export_worker = ExportWorker(
            self.service,
            self.service.exportar_excel,
            file_path,
            self.datos_actuales,
            filtros,
            self.user_data.get("_id"),
            self.user_rol
        )
        
        # Conectar señales
        self.export_worker.progress.connect(progress.setValue)
        self.export_worker.finished.connect(
            lambda result: self.on_export_finished(result, progress, "Excel")
        )
        
        # Iniciar exportación
        self.export_worker.start()
    
    def on_export_finished(self, result: dict, progress: QProgressDialog, formato: str):
        """Maneja el resultado de la exportación"""
        progress.close()
        
        # Rehabilitar botones
        self.btn_csv.setEnabled(True)
        self.btn_excel.setEnabled(True)
        
        # Mostrar resultado
        msg = QMessageBox(self)
        
        if result.get("success", False):
            msg.setIcon(QMessageBox.Information)
            msg.setWindowTitle("Exportación Exitosa")
            msg.setText(result.get("message", "Exportación completada"))
            
            # Refrescar historial si existe
            if hasattr(self, 'history_table'):
                self.cargar_historial()
        else:
            msg.setIcon(QMessageBox.Warning)
            msg.setWindowTitle("Error en Exportación")
            msg.setText(result.get("message", "Error desconocido"))
        
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
    
    def limpiar_filtros(self):
        """Limpia todos los filtros"""
        self.date_desde.setDate(QDate.currentDate().addMonths(-6))
        self.date_hasta.setDate(QDate.currentDate())
        self.combo_tipo.setCurrentIndex(0)
        self.combo_pais.setCurrentIndex(0)
        self.input_rut_filtro.clear()
        self.radio_ambos.setChecked(True)
        
        self.datos_actuales = []
        self.preview_table.setRowCount(0)
        self.label_contador.setText("Total: 0 registros")
        self.btn_csv.setEnabled(False)
        self.btn_excel.setEnabled(False)
        
        # Limpiar caché
        self.service.limpiar_cache()
    
    def cargar_historial(self):
        """Carga el historial de reportes"""
        if not hasattr(self, 'history_table'):
            return
        
        try:
            reportes = self.service.obtener_historial_reportes(
                self.user_data.get("_id"),
                self.user_rol
            )
            
            self.history_table.setRowCount(len(reportes))
            
            for row, reporte in enumerate(reportes):
                try:
                    # Fecha
                    fecha = reporte.get("fechaGeneracion")
                    if fecha:
                        if hasattr(fecha, 'strftime'):
                            fecha_str = fecha.strftime("%Y-%m-%d %H:%M")
                        else:
                            fecha_str = str(fecha)
                    else:
                        fecha_str = "N/A"
                    self.history_table.setItem(row, 0, QTableWidgetItem(fecha_str))
                    
                    # Archivo
                    self.history_table.setItem(row, 1, QTableWidgetItem(str(reporte.get("nombreArchivo", ""))))
                    
                    # Formato
                    formato = str(reporte.get("formato", ""))
                    item_formato = QTableWidgetItem(formato)
                    if formato == "CSV":
                        item_formato.setForeground(QColor(39, 174, 96))  # Verde
                    else:
                        item_formato.setForeground(QColor(52, 152, 219))  # Azul
                    self.history_table.setItem(row, 2, item_formato)
                    
                    # Registros
                    registros = str(reporte.get("totalRegistros", 0))
                    item_registros = QTableWidgetItem(registros)
                    item_registros.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                    self.history_table.setItem(row, 3, item_registros)
                    
                    # Usuario (solo admin/auditor ve IDs completos)
                    usuario_id = str(reporte.get("usuarioGeneradorId", ""))
                    if self.user_rol in ["administrador", "auditor_tributario"]:
                        usuario_text = usuario_id[:12] + "..." if len(usuario_id) > 12 else usuario_id
                    else:
                        usuario_text = "Yo"
                    self.history_table.setItem(row, 4, QTableWidgetItem(usuario_text))
                    
                except Exception as e:
                    app_logger.error(f"Error al procesar reporte {row}: {str(e)}")
                    continue
        
        except Exception as e:
            app_logger.error(f"Error al cargar historial: {str(e)}\n{traceback.format_exc()}")
    
    def apply_styles(self):
        """Aplica estilos consistentes con el resto del sistema"""
        self.setStyleSheet("""
            /* Fondo general */
            QWidget {
                background-color: transparent;
            }
            
            /* Frames y GroupBox */
            QFrame {
                background-color: #ffffff;
                border: 1px solid #e6e9ee;
                border-radius: 8px;
            }
            
            /* Inputs */
            QLineEdit, QComboBox, QDateEdit, QDoubleSpinBox {
                padding: 8px 10px;
                border: 1px solid #d7dfe8;
                border-radius: 4px;
                background-color: #ffffff;
                color: #2c3e50;
            }
            
            QLineEdit:focus, QComboBox:focus, QDateEdit:focus {
                border: 2px solid #3498db;
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
                padding: 10px 14px;
                border-radius: 6px;
                border: none;
            }
            QPushButton[role="secondary"]:hover {
                background-color: #2980b9;
            }
            
            QPushButton[role="success"] {
                background-color: #27ae60;
                color: white;
                padding: 10px 14px;
                border-radius: 6px;
                border: none;
            }
            QPushButton[role="success"]:hover {
                background-color: #229954;
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
            
            QPushButton:disabled {
                background-color: #bdc3c7;
                color: #7f8c8d;
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
            
            /* Radio buttons */
            QRadioButton {
                color: #2c3e50;
                spacing: 8px;
            }
            
            QRadioButton::indicator {
                width: 18px;
                height: 18px;
            }
            
            QRadioButton::indicator:unchecked {
                border: 2px solid #d7dfe8;
                border-radius: 9px;
                background-color: white;
            }
            
            QRadioButton::indicator:checked {
                border: 2px solid #E94E1B;
                border-radius: 9px;
                background-color: #E94E1B;
            }
        """)