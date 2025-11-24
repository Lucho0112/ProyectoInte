from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QTableWidget, QTableWidgetItem, QHeaderView, QDateEdit,
    QComboBox, QFrame, QMessageBox, QAbstractItemView,
    QScrollArea, QRadioButton, QButtonGroup, QFileDialog,
    QProgressDialog, QApplication, QGroupBox
)
from PyQt5.QtCore import Qt, QDate, pyqtSignal, QThread
from PyQt5.QtGui import QFont, QColor, QCursor
from datetime import datetime
from services.reportService import ReportService
from utils.logger import app_logger


class ExportWorker(QThread):
    """Worker thread para exportación de datos sin bloquear la UI"""
    
    finished = pyqtSignal(dict)
    progress = pyqtSignal(int, str)
    
    def __init__(self, service, file_path, datos, filtros, usuario_id, formato):
        super().__init__()
        self.service = service
        self.file_path = file_path
        self.datos = datos
        self.filtros = filtros
        self.usuario_id = usuario_id
        self.formato = formato
    
    def run(self):
        """Ejecuta la exportación en segundo plano"""
        try:
            self.progress.emit(25, "Preparando datos...")
            
            if self.formato == "CSV":
                self.progress.emit(50, "Generando archivo CSV...")
                result = self.service.exportar_csv(
                    self.file_path,
                    self.datos,
                    self.filtros,
                    self.usuario_id
                )
            else:  # Excel
                self.progress.emit(50, "Generando archivo Excel...")
                result = self.service.exportar_excel(
                    self.file_path,
                    self.datos,
                    self.filtros,
                    self.usuario_id
                )
            
            self.progress.emit(90, "Registrando en historial...")
            self.progress.emit(100, "¡Completado!")
            
            self.finished.emit(result)
            
        except Exception as e:
            app_logger.error(f"Error en worker de exportación: {str(e)}")
            self.finished.emit({
                "success": False,
                "message": f"Error inesperado: {str(e)}"
            })


class GenerarReportesContent(QWidget):
    """Contenido de generación de reportes"""
    
    back_requested = pyqtSignal()
    
    def __init__(self, user_data: dict, parent=None):
        super().__init__(parent)
        self.user_data = user_data
        self.user_rol = user_data.get("rol", "cliente")
        self.service = ReportService()
        self.datos_actuales = []
        self.export_worker = None
        self.progress_dialog = None
        
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
        toolbar_layout.addWidget(btn_limpiar)
        
        toolbar_layout.addStretch()
        
        self.label_contador = QLabel("Total: 0 registros")
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
        
        # Fecha desde
        filter_layout.addWidget(QLabel("Fecha desde:"))
        self.date_desde = QDateEdit()
        self.date_desde.setCalendarPopup(True)
        self.date_desde.setDate(QDate.currentDate().addMonths(-6))
        self.date_desde.setDisplayFormat("dd/MM/yyyy")
        self.date_desde.setMinimumHeight(35)
        filter_layout.addWidget(self.date_desde)
        
        # Fecha hasta
        filter_layout.addWidget(QLabel("Fecha hasta:"))
        self.date_hasta = QDateEdit()
        self.date_hasta.setCalendarPopup(True)
        self.date_hasta.setDate(QDate.currentDate())
        self.date_hasta.setDisplayFormat("dd/MM/yyyy")
        self.date_hasta.setMinimumHeight(35)
        filter_layout.addWidget(self.date_hasta)
        
        # Tipo de impuesto
        filter_layout.addWidget(QLabel("Tipo:"))
        self.combo_tipo = QComboBox()
        self.combo_tipo.addItems(["Todos", "IVA", "Renta", "Importación", "Exportación", "Otro"])
        self.combo_tipo.setMinimumHeight(35)
        filter_layout.addWidget(self.combo_tipo)
        
        # País
        filter_layout.addWidget(QLabel("País:"))
        self.combo_pais = QComboBox()
        self.combo_pais.addItems(["Todos", "Chile", "Perú", "Colombia"])
        self.combo_pais.setMinimumHeight(35)
        filter_layout.addWidget(self.combo_pais)
        
        filter_layout.addStretch()
        
        # Botón buscar
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
                min-height: 35px;
            }
            QPushButton:hover { background-color: #229954; }
        """)
        filter_layout.addWidget(btn_buscar)
        
        filter_frame.setLayout(filter_layout)
        layout.addWidget(filter_frame)
        
        # Radio buttons para estado (debajo de filtros principales)
        estado_frame = QFrame()
        estado_frame.setStyleSheet("""
            QFrame {
                background-color: white;
                border: 1px solid #e6e9ee;
                border-radius: 5px;
                padding: 10px;
            }
        """)
        
        estado_layout = QHBoxLayout()
        estado_layout.addWidget(QLabel("Estado de datos:"))
        
        self.button_group = QButtonGroup()
        
        self.radio_ambos = QRadioButton("📊 Ambos (Local + Bolsa)")
        self.radio_ambos.setChecked(True)
        self.button_group.addButton(self.radio_ambos)
        estado_layout.addWidget(self.radio_ambos)
        
        self.radio_local = QRadioButton("💼 Solo Local")
        self.button_group.addButton(self.radio_local)
        estado_layout.addWidget(self.radio_local)
        
        self.radio_bolsa = QRadioButton("🏛️ Solo Bolsa")
        self.button_group.addButton(self.radio_bolsa)
        estado_layout.addWidget(self.radio_bolsa)
        
        estado_layout.addStretch()
        estado_frame.setLayout(estado_layout)
        layout.addWidget(estado_frame)
    
    def add_preview(self, layout):
        """Vista previa de datos"""
        # Header de vista previa
        preview_header = QHBoxLayout()
        
        preview_title = QLabel("Vista Previa (primeras 50 filas)")
        preview_title.setFont(QFont("Arial", 11, QFont.Bold))
        preview_title.setStyleSheet("color: #2c3e50;")
        preview_header.addWidget(preview_title)
        
        preview_header.addStretch()
        layout.addLayout(preview_header)
        
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
        altura_minima = 35 + (30 * 10) + 10
        self.preview_table.setMinimumHeight(altura_minima)
        altura_maxima = 35 + (30 * 15) + 10
        self.preview_table.setMaximumHeight(altura_maxima)
        
        layout.addWidget(self.preview_table)
    
    def add_export_buttons(self, layout):
        """Botones de exportación"""
        export_layout = QHBoxLayout()
        export_layout.setSpacing(15)
        
        export_layout.addWidget(QLabel("Exportar datos:"))
        
        export_layout.addStretch()
        
        # Botón CSV
        self.btn_csv = QPushButton("📄 Exportar CSV")
        self.btn_csv.setFont(QFont("Arial", 10, QFont.Bold))
        self.btn_csv.setMinimumSize(150, 45)
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
            QPushButton:hover:enabled {
                background-color: #229954;
            }
            QPushButton:disabled {
                background-color: #bdc3c7;
                color: #7f8c8d;
            }
        """)
        export_layout.addWidget(self.btn_csv)
        
        # Botón Excel
        self.btn_excel = QPushButton("📊 Exportar Excel")
        self.btn_excel.setFont(QFont("Arial", 10, QFont.Bold))
        self.btn_excel.setMinimumSize(150, 45)
        self.btn_excel.setCursor(QCursor(Qt.PointingHandCursor))
        self.btn_excel.clicked.connect(self.exportar_excel)
        self.btn_excel.setEnabled(False)
        self.btn_excel.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 10px 20px;
            }
            QPushButton:hover:enabled {
                background-color: #2980b9;
            }
            QPushButton:disabled {
                background-color: #bdc3c7;
                color: #7f8c8d;
            }
        """)
        export_layout.addWidget(self.btn_excel)
        
        layout.addLayout(export_layout)
    
    def add_history(self, layout):
        """Historial de reportes (solo admin/auditor)"""
        # Header del historial
        history_header = QHBoxLayout()
        
        history_title = QLabel("Historial de Reportes Generados")
        history_title.setFont(QFont("Arial", 11, QFont.Bold))
        history_title.setStyleSheet("color: #2c3e50;")
        history_header.addWidget(history_title)
        
        history_header.addStretch()
        
        btn_refrescar = QPushButton("🔄 Refrescar")
        btn_refrescar.setMinimumHeight(35)
        btn_refrescar.setCursor(QCursor(Qt.PointingHandCursor))
        btn_refrescar.clicked.connect(self.cargar_historial)
        btn_refrescar.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 8px 16px;
            }
            QPushButton:hover { background-color: #2980b9; }
        """)
        history_header.addWidget(btn_refrescar)
        
        layout.addLayout(history_header)
        
        # Tabla de historial
        self.history_table = QTableWidget()
        self.history_table.setColumnCount(5)
        self.history_table.setHorizontalHeaderLabels([
            "Fecha", "Archivo", "Formato", "Registros", "Usuario"
        ])
        
        self.history_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.history_table.setAlternatingRowColors(True)
        self.history_table.verticalHeader().setVisible(False)
        self.history_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        altura_minima = 35 + (30 * 10) + 10
        self.history_table.setMinimumHeight(altura_minima)
        altura_maxima = 35 + (30 * 12) + 10
        self.history_table.setMaximumHeight(altura_maxima)
        
        header = self.history_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.ResizeToContents)
        
        layout.addWidget(self.history_table)
        
        # Cargar historial inicial
        self.cargar_historial()
    
    def add_footer(self, layout):
        """Footer informativo"""
        if self.user_rol in ["administrador", "auditor_tributario"]:
            footer_text = "💡 Puede generar reportes de todas las calificaciones del sistema"
        else:
            footer_text = "💡 Los reportes incluirán solo los datos a los que tiene acceso según su rol"
        
        footer = QLabel(footer_text)
        footer.setFont(QFont("Arial", 8))
        footer.setStyleSheet("color: #7f8c8d;")
        layout.addWidget(footer)
    
    def validar_filtros(self) -> tuple:
        """
        Valida que los filtros sean correctos antes de aplicar
        
        Returns:
            tuple[bool, str]: (es_valido, mensaje_error)
        """
        # Validar rango de fechas
        fecha_desde = self.date_desde.date().toPyDate()
        fecha_hasta = self.date_hasta.date().toPyDate()
        
        if fecha_desde > fecha_hasta:
            return False, "La fecha 'desde' no puede ser mayor que la fecha 'hasta'"
        
        # Validar que el rango no sea muy antiguo (más de 5 años)
        from datetime import timedelta
        if (datetime.now().date() - fecha_desde).days > 1825:  # 5 años
            return False, "El rango de fechas no puede ser superior a 5 años"
        
        return True, ""
    
    def aplicar_filtros(self):
        """Aplica los filtros y carga datos"""
        try:
            # Validar filtros primero
            es_valido, mensaje_error = self.validar_filtros()
            if not es_valido:
                QMessageBox.warning(
                    self,
                    "⚠️ Filtros Inválidos",
                    mensaje_error
                )
                return
            
            filtros = self.obtener_filtros()
            
            # Mostrar mensaje de carga
            QApplication.setOverrideCursor(Qt.WaitCursor)
            app_logger.info(f"Aplicando filtros: {filtros}")
            
            # Obtener datos
            self.datos_actuales = self.service.obtener_datos_filtrados(
                filtros,
                self.user_data.get("_id"),
                self.user_rol
            )
            
            QApplication.restoreOverrideCursor()
            
            if not self.datos_actuales:
                QMessageBox.information(
                    self,
                    "📊 Sin resultados",
                    "No se encontraron datos con los filtros aplicados.\n\n"
                    "Intente modificar los criterios de búsqueda."
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
            
            # Actualizar contador con formato
            total = len(self.datos_actuales)
            self.label_contador.setText(f"Total: {total:,} registros")
            
            # Mostrar alerta si hay muchos registros
            if total > 10000:
                QMessageBox.information(
                    self,
                    "📊 Gran volumen de datos",
                    f"Se encontraron {total:,} registros.\n\n"
                    "La exportación puede tomar varios segundos."
                )
            
            app_logger.info(f"Filtros aplicados exitosamente: {total} registros")
            
        except Exception as e:
            QApplication.restoreOverrideCursor()
            app_logger.error(f"Error al aplicar filtros: {str(e)}", exc_info=True)
            QMessageBox.critical(
                self,
                "❌ Error",
                f"Error al aplicar filtros:\n{str(e)}"
            )
    
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
        try:
            self.preview_table.setRowCount(len(datos))
            
            for row, cal in enumerate(datos):
                try:
                    # RUT Cliente - manejar diferentes tipos
                    cliente_id = cal.get("clienteId", "")
                    
                    # Si clienteId es una lista, tomar el primer elemento
                    if isinstance(cliente_id, (list, tuple)):
                        cliente_id = cliente_id[0] if len(cliente_id) > 0 else ""
                    
                    # Convertir a string y validar
                    cliente_id = str(cliente_id) if cliente_id else ""
                    
                    if cliente_id and cliente_id != "None":
                        rut = self.service.obtener_rut_cliente(cliente_id)
                    else:
                        rut = "N/A"
                    
                    self.preview_table.setItem(row, 0, QTableWidgetItem(str(rut)))
                except Exception as e:
                    app_logger.error(f"Error obteniendo RUT fila {row}: {str(e)}", exc_info=True)
                    self.preview_table.setItem(row, 0, QTableWidgetItem("ERROR"))
                
                # Fecha
                try:
                    fecha = cal.get("fechaDeclaracion", "")
                    self.preview_table.setItem(row, 1, QTableWidgetItem(str(fecha)))
                except:
                    self.preview_table.setItem(row, 1, QTableWidgetItem(""))
                
                # Tipo
                try:
                    tipo = cal.get("tipoImpuesto", "")
                    self.preview_table.setItem(row, 2, QTableWidgetItem(str(tipo)))
                except:
                    self.preview_table.setItem(row, 2, QTableWidgetItem(""))
                
                # País
                try:
                    pais = cal.get("pais", "")
                    self.preview_table.setItem(row, 3, QTableWidgetItem(str(pais)))
                except:
                    self.preview_table.setItem(row, 3, QTableWidgetItem(""))
                
                # Monto
                try:
                    monto = cal.get("montoDeclarado", 0)
                    if isinstance(monto, (int, float)):
                        item_monto = QTableWidgetItem(f"${monto:,.2f}")
                    else:
                        item_monto = QTableWidgetItem("$0.00")
                    item_monto.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                    self.preview_table.setItem(row, 4, item_monto)
                except Exception as e:
                    app_logger.error(f"Error en monto fila {row}: {str(e)}")
                    self.preview_table.setItem(row, 4, QTableWidgetItem("$0.00"))
                
                # Suma 8-19 y Válido
                try:
                    factores = cal.get("factores", {})
                    
                    # Validar que factores sea un diccionario
                    if not isinstance(factores, dict):
                        app_logger.warning(f"Factores no es diccionario en fila {row}: {type(factores)}")
                        factores = {}
                    
                    suma = 0
                    for i in range(8, 20):
                        factor_key = f"factor_{i}"
                        factor_val = factores.get(factor_key, 0)
                        
                        # Convertir listas, tuplas u otros tipos a número
                        if isinstance(factor_val, (list, tuple)):
                            # Si es una lista, tomar el primer elemento si existe
                            factor_val = factor_val[0] if len(factor_val) > 0 else 0
                        
                        # Asegurar que sea numérico
                        if isinstance(factor_val, (int, float)):
                            suma += factor_val
                        else:
                            # Intentar convertir a float
                            try:
                                suma += float(factor_val)
                            except (ValueError, TypeError):
                                app_logger.warning(f"Factor {factor_key} no numérico en fila {row}: {factor_val}")
                                continue
                    
                    item_suma = QTableWidgetItem(f"{suma:.4f}")
                    item_suma.setTextAlignment(Qt.AlignCenter)
                    
                    if suma > 1.0:
                        item_suma.setBackground(QColor(255, 200, 200))
                        item_suma.setForeground(QColor(200, 0, 0))
                    else:
                        item_suma.setBackground(QColor(200, 255, 200))
                        item_suma.setForeground(QColor(0, 150, 0))
                    
                    self.preview_table.setItem(row, 5, item_suma)
                    
                    # Válido
                    valido = "✅ Sí" if suma <= 1.0 else "❌ No"
                    item_valido = QTableWidgetItem(valido)
                    item_valido.setTextAlignment(Qt.AlignCenter)
                    self.preview_table.setItem(row, 7, item_valido)
                    
                except Exception as e:
                    app_logger.error(f"Error calculando suma factores fila {row}: {str(e)}", exc_info=True)
                    self.preview_table.setItem(row, 5, QTableWidgetItem("ERROR"))
                    self.preview_table.setItem(row, 7, QTableWidgetItem("❌"))
                
                # Estado
                try:
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
                    
                    self.preview_table.setItem(row, 6, item_estado)
                except Exception as e:
                    app_logger.error(f"Error en estado fila {row}: {str(e)}")
                    self.preview_table.setItem(row, 6, QTableWidgetItem(""))
            
            app_logger.info(f"Vista previa actualizada con {len(datos)} registros")
            
        except Exception as e:
            app_logger.error(f"Error general al actualizar vista previa: {str(e)}", exc_info=True)
            QMessageBox.warning(
                self,
                "⚠️ Advertencia",
                f"Error al actualizar la vista previa:\n{str(e)}\n\nLos datos están disponibles para exportar."
            )
    
    def exportar_csv(self):
        """Exporta datos a CSV"""
        if not self.datos_actuales:
            QMessageBox.warning(
                self,
                "⚠️ Sin datos",
                "No hay datos para exportar. Aplique filtros primero."
            )
            return
        
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Guardar reporte CSV",
            f"reporte_tributario_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            "CSV (*.csv)"
        )
        
        if not file_path:
            return
        
        try:
            # Crear diálogo de progreso
            self.progress_dialog = QProgressDialog(
                "Exportando datos a CSV...",
                None,  # Sin botón cancelar
                0,
                100,
                self
            )
            self.progress_dialog.setWindowTitle("Exportando")
            self.progress_dialog.setWindowModality(Qt.WindowModal)
            self.progress_dialog.setMinimumDuration(0)
            self.progress_dialog.setValue(0)
            self.progress_dialog.show()
            QApplication.processEvents()
            
            # Crear worker
            filtros = self.obtener_filtros()
            self.export_worker = ExportWorker(
                self.service,
                file_path,
                self.datos_actuales,
                filtros,
                self.user_data.get("_id"),
                "CSV"
            )
            
            # Conectar señales
            self.export_worker.progress.connect(self.actualizar_progreso)
            self.export_worker.finished.connect(self.exportacion_completada)
            
            # Iniciar exportación
            self.export_worker.start()
            
        except Exception as e:
            if self.progress_dialog:
                self.progress_dialog.close()
            app_logger.error(f"Error al iniciar exportación CSV: {str(e)}", exc_info=True)
            QMessageBox.critical(
                self,
                "❌ Error",
                f"Error al exportar CSV:\n{str(e)}"
            )
    
    def exportar_excel(self):
        """Exporta datos a Excel"""
        if not self.datos_actuales:
            QMessageBox.warning(
                self,
                "⚠️ Sin datos",
                "No hay datos para exportar. Aplique filtros primero."
            )
            return
        
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Guardar reporte Excel",
            f"reporte_tributario_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            "Excel (*.xlsx)"
        )
        
        if not file_path:
            return
        
        try:
            # Crear diálogo de progreso
            self.progress_dialog = QProgressDialog(
                "Exportando datos a Excel...",
                None,  # Sin botón cancelar
                0,
                100,
                self
            )
            self.progress_dialog.setWindowTitle("Exportando")
            self.progress_dialog.setWindowModality(Qt.WindowModal)
            self.progress_dialog.setMinimumDuration(0)
            self.progress_dialog.setValue(0)
            self.progress_dialog.show()
            QApplication.processEvents()
            
            # Crear worker
            filtros = self.obtener_filtros()
            self.export_worker = ExportWorker(
                self.service,
                file_path,
                self.datos_actuales,
                filtros,
                self.user_data.get("_id"),
                "Excel"
            )
            
            # Conectar señales
            self.export_worker.progress.connect(self.actualizar_progreso)
            self.export_worker.finished.connect(self.exportacion_completada)
            
            # Iniciar exportación
            self.export_worker.start()
            
        except Exception as e:
            if self.progress_dialog:
                self.progress_dialog.close()
            app_logger.error(f"Error al iniciar exportación Excel: {str(e)}", exc_info=True)
            QMessageBox.critical(
                self,
                "❌ Error",
                f"Error al exportar Excel:\n{str(e)}"
            )
    
    def actualizar_progreso(self, valor: int, mensaje: str):
        """Actualiza el progreso de la exportación"""
        if self.progress_dialog and hasattr(self.progress_dialog, 'setLabelText'):
            try:
                self.progress_dialog.setValue(valor)
                self.progress_dialog.setLabelText(mensaje)
                QApplication.processEvents()
            except:
                pass
    
    def exportacion_completada(self, result: dict):
        """Maneja la finalización de la exportación"""
        if self.progress_dialog:
            self.progress_dialog.close()
            self.progress_dialog = None
        
        if result.get("success", False):
            QMessageBox.information(
                self,
                "✅ Éxito",
                result.get("message", "Exportación completada")
            )
            # Refrescar historial si existe
            if hasattr(self, 'history_table'):
                self.cargar_historial()
        else:
            QMessageBox.warning(
                self,
                "⚠️ Error",
                result.get("message", "Error desconocido en la exportación")
            )
    
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
        
        # Limpiar caché de RUTs
        self.service.limpiar_cache_rut()
        
        app_logger.info("Filtros limpiados")
        
        QMessageBox.information(
            self,
            "🗑️ Filtros limpiados",
            "Se han restablecido todos los filtros a sus valores predeterminados."
        )
    
    def cargar_historial(self):
        """Carga el historial de reportes"""
        if not hasattr(self, 'history_table'):
            return
        
        try:
            app_logger.info("Cargando historial de reportes")
            
            reportes = self.service.obtener_historial_reportes(
                self.user_data.get("_id"),
                self.user_rol
            )
            
            self.history_table.setRowCount(len(reportes))
            
            for row, reporte in enumerate(reportes):
                # Fecha
                fecha = reporte.get("fechaGeneracion")
                if fecha:
                    fecha_str = fecha.strftime("%d/%m/%Y %H:%M") if hasattr(fecha, 'strftime') else str(fecha)
                else:
                    fecha_str = "N/A"
                self.history_table.setItem(row, 0, QTableWidgetItem(fecha_str))
                
                # Archivo
                nombre_archivo = reporte.get("nombreArchivo", "")
                self.history_table.setItem(row, 1, QTableWidgetItem(nombre_archivo))
                
                # Formato con icono
                formato = reporte.get("formato", "")
                icono_formato = "📄" if formato == "CSV" else "📊"
                item_formato = QTableWidgetItem(f"{icono_formato} {formato}")
                item_formato.setTextAlignment(Qt.AlignCenter)
                self.history_table.setItem(row, 2, item_formato)
                
                # Registros
                registros = reporte.get("totalRegistros", 0)
                item_registros = QTableWidgetItem(f"{registros:,}")
                item_registros.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.history_table.setItem(row, 3, item_registros)
                
                # Usuario (solo admin ve esto)
                usuario_id = reporte.get("usuarioGeneradorId", "")
                if self.user_rol == "administrador":
                    usuario_text = usuario_id[:12] + "..." if len(usuario_id) > 12 else usuario_id
                else:
                    usuario_text = "Yo"
                item_usuario = QTableWidgetItem(usuario_text)
                item_usuario.setTextAlignment(Qt.AlignCenter)
                self.history_table.setItem(row, 4, item_usuario)
            
            app_logger.info(f"Historial cargado: {len(reportes)} reportes")
            
        except Exception as e:
            app_logger.error(f"Error al cargar historial: {str(e)}", exc_info=True)
            QMessageBox.warning(
                self,
                "⚠️ Advertencia",
                f"Error al cargar el historial de reportes:\n{str(e)}"
            )
    
    def apply_styles(self):
        """Aplica estilos consistentes con las demás gestiones"""
        self.setStyleSheet("""
            /* Fondo del contenedor */
            QWidget {
                background-color: transparent;
            }
            
            /* Marco de filtros */
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
            
            /* Tabla */
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
        """)