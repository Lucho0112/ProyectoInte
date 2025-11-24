import os
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, 
    QFrame, QFileDialog, QTableWidget, QTableWidgetItem, 
    QProgressBar, QTextEdit, QMessageBox, QHeaderView, QScrollArea, QAbstractItemView
)
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QFont, QCursor, QDragEnterEvent, QDropEvent


class FileDropZone(QFrame):
    """Zona de drag & drop para archivos"""
    
    file_dropped = pyqtSignal(str)
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setup_ui()
        self.setAcceptDrops(True)
    
    def setup_ui(self):
        self.setFixedHeight(200)
        self.setStyleSheet("""
            QFrame {
                background-color: #f8f9fa;
                border: 2px dashed #dee2e6;
                border-radius: 10px;
            }
            QFrame:hover {
                border-color: #E94E1B;
                background-color: #fef5f1;
            }
        """)
        
        layout = QVBoxLayout()
        layout.setAlignment(Qt.AlignCenter)
        
        icon_label = QLabel("📤")
        icon_label.setFont(QFont("Arial", 48))
        icon_label.setAlignment(Qt.AlignCenter)
        icon_label.setAttribute(Qt.WA_TransparentForMouseEvents)
        layout.addWidget(icon_label)
        
        text_label = QLabel("Arrastrar archivo CSV/Excel aquí\no hacer clic para seleccionar")
        text_label.setFont(QFont("Arial", 11))
        text_label.setAlignment(Qt.AlignCenter)
        text_label.setStyleSheet("color: #7f8c8d;")
        text_label.setAttribute(Qt.WA_TransparentForMouseEvents)
        layout.addWidget(text_label)
        
        self.setLayout(layout)
    
    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
    
    def dropEvent(self, event: QDropEvent):
        files = [u.toLocalFile() for u in event.mimeData().urls()]
        if files:
            file_path = files[0]
            if file_path.endswith(('.csv', '.xlsx', '.xls')):
                self.file_dropped.emit(file_path)
            else:
                QMessageBox.warning(self, "Archivo inválido", 
                                  "Por favor seleccione un archivo CSV o Excel")
    
    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            file_path, _ = QFileDialog.getOpenFileName(
                self, "Seleccionar archivo", "",
                "Archivos CSV/Excel (*.csv *.xlsx *.xls)"
            )
            if file_path:
                self.file_dropped.emit(file_path)


class CargaMasivaContent(QWidget):
    """Contenido de carga masiva (sin ser ventana independiente)"""
    
    back_requested = pyqtSignal()
    
    def __init__(self, user_data: dict, parent=None):
        super().__init__(parent)
        self.user_data = user_data
        self.file_path = None
        self.dataframe = None
        self.validation_passed = False
        self.init_ui()
    
    def init_ui(self):
        """Inicializa la interfaz"""
        # Scroll area principal
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll_area.setStyleSheet("QScrollArea { border: none; background-color: #f5f6fa; }")
        
        # Contenido
        content_widget = QWidget()
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(40, 30, 40, 30)
        main_layout.setSpacing(20)
        
        # Header con botón volver
        self.add_header(main_layout)
        
        # Zona de carga
        self.add_file_zone(main_layout)
        
        # Info del archivo
        self.add_file_info(main_layout)
        
        # Vista previa
        self.add_preview_table(main_layout)
        
        # Validaciones
        self.add_validation_area(main_layout)
        
        # Progreso
        self.add_progress_bar(main_layout)
        
        # Botones
        self.add_action_buttons(main_layout)
        
        content_widget.setLayout(main_layout)
        scroll_area.setWidget(content_widget)
        
        # Layout principal del widget
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
            QPushButton:hover {
                color: #2980b9;
            }
        """)
        header_layout.addWidget(back_button)
        
        title = QLabel("Carga Masiva de Datos Tributarios")
        title.setFont(QFont("Arial", 18, QFont.Bold))
        title.setStyleSheet("color: #2c3e50;")
        header_layout.addWidget(title)
        
        header_layout.addStretch()
        layout.addLayout(header_layout)
    
    def add_file_zone(self, layout):
        self.drop_zone = FileDropZone()
        self.drop_zone.file_dropped.connect(self.on_file_selected)
        layout.addWidget(self.drop_zone)
        
        # Panel de ayuda (se muestra cuando no hay archivo)
        self.add_help_panel(layout)
    
    def add_file_info(self, layout):
        info_frame = QFrame()
        info_frame.setStyleSheet("""
            QFrame {
                background-color: #e8f4f8;
                border-radius: 5px;
                padding: 10px;
            }
        """)
        info_layout = QVBoxLayout()
        
        self.file_name_label = QLabel("📄 Archivo: Ninguno seleccionado")
        self.file_name_label.setFont(QFont("Arial", 10))
        info_layout.addWidget(self.file_name_label)
        
        self.file_records_label = QLabel("📊 Registros detectados: 0")
        self.file_records_label.setFont(QFont("Arial", 10))
        info_layout.addWidget(self.file_records_label)
        
        info_frame.setLayout(info_layout)
        layout.addWidget(info_frame)
        info_frame.hide()
        self.info_frame = info_frame
    
    def add_help_panel(self, layout):
        """Panel de ayuda con instrucciones"""
        help_frame = QFrame()
        help_frame.setStyleSheet("""
            QFrame {
                background-color: white;
                border: 1px solid #dee2e6;
                border-radius: 10px;
                padding: 20px;
            }
        """)
        
        help_layout = QVBoxLayout()
        
        # Título
        title = QLabel("📋 Instrucciones de Carga Masiva")
        title.setFont(QFont("Arial", 14, QFont.Bold))
        title.setStyleSheet("color: #2c3e50; margin-bottom: 10px;")
        help_layout.addWidget(title)
        
        # Pasos
        steps = [
            "1. Descarga la plantilla CSV/Excel haciendo clic en 'Descargar Plantilla'",
            "2. Completa la plantilla con tus datos tributarios",
            "3. Arrastra el archivo o haz clic en la zona de carga",
            "4. Revisa la vista previa y validaciones",
            "5. Haz clic en 'Importar Datos' para confirmar"
        ]
        
        for step in steps:
            step_label = QLabel(step)
            step_label.setFont(QFont("Arial", 10))
            step_label.setStyleSheet("color: #555; margin: 5px 0px;")
            step_label.setWordWrap(True)
            help_layout.addWidget(step_label)
        
        help_layout.addSpacing(15)
        
        # Información adicional en grid
        info_grid = QHBoxLayout()
        
        # Tarjeta: Formatos aceptados
        format_card = self.create_info_card(
            "📂 Formatos",
            "CSV, XLSX, XLS"
        )
        info_grid.addWidget(format_card)
        
        # Tarjeta: Límite
        limit_card = self.create_info_card(
            "📊 Límite",
            "Hasta 1000 registros por archivo"
        )
        info_grid.addWidget(limit_card)
        
        # Tarjeta: Validaciones
        validation_card = self.create_info_card(
            "✅ Validaciones",
            "RUT, montos, factores 8-19 ≤ 1.0"
        )
        info_grid.addWidget(validation_card)
        
        help_layout.addLayout(info_grid)
        help_frame.setLayout(help_layout)
        layout.addWidget(help_frame)
        
        self.help_frame = help_frame

        help_layout.addSpacing(10)

        warning_label = QLabel(
            "⚠️ IMPORTANTE: Solo se pueden importar datos de clientes YA REGISTRADOS.\n"
            "Si un RUT no está registrado, la importación fallará."
        )
        warning_label.setFont(QFont("Arial", 9))
        warning_label.setStyleSheet("""
            QLabel {
                color: #856404;
                background-color: #fff3cd;
                border: 1px solid #ffc107;
                border-radius: 5px;
                padding: 10px;
            }
        """)
        warning_label.setWordWrap(True)
        help_layout.addWidget(warning_label)
    
    def create_info_card(self, title: str, content: str):
        """Crea una tarjeta de información"""
        card = QFrame()
        card.setStyleSheet("""
            QFrame {
                background-color: #f8f9fa;
                border: 1px solid #dee2e6;
                border-radius: 5px;
                padding: 15px;
            }
        """)
        
        card_layout = QVBoxLayout()
        
        title_label = QLabel(title)
        title_label.setFont(QFont("Arial", 10, QFont.Bold))
        title_label.setStyleSheet("color: #2c3e50;")
        card_layout.addWidget(title_label)
        
        content_label = QLabel(content)
        content_label.setFont(QFont("Arial", 9))
        content_label.setStyleSheet("color: #7f8c8d;")
        content_label.setWordWrap(True)
        card_layout.addWidget(content_label)
        
        card.setLayout(card_layout)
        return card
    
    def add_preview_table(self, layout):
        preview_label = QLabel("Vista Previa (primeras 10 filas)")
        preview_label.setFont(QFont("Arial", 12, QFont.Bold))
        preview_label.setStyleSheet("color: #2c3e50;")
        layout.addWidget(preview_label)
        
        self.preview_table = QTableWidget()
        self.preview_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.preview_table.setAlternatingRowColors(True)
        self.preview_table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.preview_table.setStyleSheet("""
            QTableWidget {
                border: 1px solid #dee2e6;
                border-radius: 5px;
                background-color: white;
            }
            QHeaderView::section {
                background-color: #E94E1B;
                color: white;
                padding: 8px;
                border: none;
                font-weight: bold;
                font-size: 10px;
            }
            QTableWidget::item {
                padding: 5px;
            }
        """)
        self.preview_table.setMaximumHeight(450)
        self.preview_table.setMinimumHeight(200)
        layout.addWidget(self.preview_table)
        
        # ← NUEVO: Leyenda de colores
        legend_frame = QFrame()
        legend_frame.setStyleSheet("""
            QFrame {
                background-color: #f8f9fa;
                border: 1px solid #dee2e6;
                border-radius: 5px;
                padding: 10px;
            }
        """)
        
        legend_layout = QHBoxLayout()
        legend_layout.setSpacing(15)
        
        legend_label = QLabel("🎨 Leyenda:")
        legend_label.setFont(QFont("Arial", 9, QFont.Bold))
        legend_layout.addWidget(legend_label)
        
        legends = [
            ("🟢 Montos", "#e8f5e9"),
            ("🟡 Factores 8-19", "#fff9c4"),
            ("🔵 Factores 1-7", "#e3f2fd"),
            ("🟣 Fechas", "#f3e5f5"),
            ("🔴 Suma > 1.0", "#ffcdd2")
        ]
        
        for text, color in legends:
            legend_item = QLabel(text)
            legend_item.setFont(QFont("Arial", 8))
            legend_item.setStyleSheet(f"""
                QLabel {{
                    background-color: {color};
                    padding: 3px 8px;
                    border-radius: 3px;
                    border: 1px solid {color};
                }}
            """)
            legend_layout.addWidget(legend_item)
        
        legend_layout.addStretch()
        legend_frame.setLayout(legend_layout)
        layout.addWidget(legend_frame)
        
        preview_label.hide()
        self.preview_table.hide()
        legend_frame.hide()  # ← También ocultar leyenda inicialmente
        
        self.preview_label = preview_label
        self.legend_frame = legend_frame
    
    def add_validation_area(self, layout):
        validation_label = QLabel("Validaciones:")
        validation_label.setFont(QFont("Arial", 12, QFont.Bold))
        layout.addWidget(validation_label)
        
        self.validation_text = QTextEdit()
        self.validation_text.setReadOnly(True)
        self.validation_text.setMaximumHeight(150)
        self.validation_text.setStyleSheet("""
            QTextEdit {
                border: 1px solid #dee2e6;
                border-radius: 5px;
                padding: 10px;
                background-color: #f8f9fa;
                font-family: 'Courier New';
                font-size: 10px;
            }
        """)
        layout.addWidget(self.validation_text)
        
        validation_label.hide()
        self.validation_text.hide()
        self.validation_label = validation_label
    
    def add_progress_bar(self, layout):
        self.progress_bar = QProgressBar()
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 1px solid #dee2e6;
                border-radius: 5px;
                text-align: center;
                height: 25px;
            }
            QProgressBar::chunk {
                background-color: #27ae60;
            }
        """)
        layout.addWidget(self.progress_bar)
        self.progress_bar.hide()
    
    def add_action_buttons(self, layout):
        buttons_layout = QHBoxLayout()
        buttons_layout.addStretch()
        
        self.template_button = QPushButton("📥 Descargar Plantilla")
        self.template_button.setFont(QFont("Arial", 10))
        self.template_button.setCursor(QCursor(Qt.PointingHandCursor))
        self.template_button.setFixedHeight(40)
        self.template_button.clicked.connect(self.download_template)
        self.template_button.setStyleSheet("""
            QPushButton {
                background-color: #95a5a6;
                color: white;
                border-radius: 5px;
                padding: 10px 20px;
            }
            QPushButton:hover { background-color: #7f8c8d; }
        """)
        buttons_layout.addWidget(self.template_button)
        
        self.cancel_button = QPushButton("Cancelar")
        self.cancel_button.setFont(QFont("Arial", 10))
        self.cancel_button.setCursor(QCursor(Qt.PointingHandCursor))
        self.cancel_button.setFixedHeight(40)
        self.cancel_button.clicked.connect(self.reset_form)
        self.cancel_button.setStyleSheet("""
            QPushButton {
                background-color: #e74c3c;
                color: white;
                border-radius: 5px;
                padding: 10px 20px;
            }
            QPushButton:hover { background-color: #c0392b; }
        """)
        buttons_layout.addWidget(self.cancel_button)
        self.cancel_button.hide()
        
        self.import_button = QPushButton("Importar Datos")
        self.import_button.setFont(QFont("Arial", 10, QFont.Bold))
        self.import_button.setCursor(QCursor(Qt.PointingHandCursor))
        self.import_button.setFixedHeight(40)
        self.import_button.clicked.connect(self.import_data)
        self.import_button.setEnabled(False)
        self.import_button.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                border-radius: 5px;
                padding: 10px 20px;
            }
            QPushButton:hover { background-color: #2980b9; }
            QPushButton:disabled { background-color: #bdc3c7; }
        """)
        buttons_layout.addWidget(self.import_button)
        self.import_button.hide()
        
        layout.addLayout(buttons_layout)
    
    def on_file_selected(self, file_path: str):
        self.file_path = file_path
        self.validate_file()
    
    def validate_file(self):
        from utils.csvValidator import CSVValidator
        from services.massiveLoadService import CargaMasivaService  # ← AGREGAR
        
        # Ocultar panel de ayuda cuando se carga archivo
        self.help_frame.hide()
        
        self.validation_text.setText("⏳ Validando archivo...")
        self.validation_label.show()
        self.validation_text.show()
        
        # Paso 1: Validar formato CSV
        is_valid, message, df = CSVValidator.validate_file(self.file_path)
        
        self.dataframe = df
        
        self.file_name_label.setText(f"📄 Archivo: {os.path.basename(self.file_path)}")
        self.file_records_label.setText(f"📊 Registros detectados: {len(df) if df is not None else 0}")
        self.info_frame.show()
        
        if not is_valid:
            # Error de formato CSV
            self.validation_text.setStyleSheet("""
                QTextEdit {
                    border: 1px solid #e74c3c;
                    border-radius: 5px;
                    padding: 10px;
                    background-color: #fadbd8;
                    color: #e74c3c;
                    font-family: 'Courier New';
                }
            """)
            self.validation_text.setText(f"❌ {message}")
            self.import_button.setEnabled(False)
            self.validation_passed = False
            
            if df is not None:
                self.show_preview(df)
            
            self.cancel_button.show()
            self.import_button.show()
            return
        
        # ← NUEVO: Paso 2: Validar que los clientes existan
        self.validation_text.setText("⏳ Validando clientes en el sistema...")
        
        service = CargaMasivaService()
        validacion_clientes = service.validate_all_clientes(df)
        
        if not validacion_clientes["valid"]:
            # Clientes faltantes
            missing_ruts = validacion_clientes["missing_ruts"]
            
            error_msg = f"❌ ARCHIVO VÁLIDO PERO:\n\n"
            error_msg += f"{validacion_clientes['message']}\n\n"
            error_msg += f"RUTs no registrados ({len(missing_ruts)}):\n"
            error_msg += "\n".join([f"• {rut}" for rut in missing_ruts[:10]])
            
            if len(missing_ruts) > 10:
                error_msg += f"\n... y {len(missing_ruts) - 10} más."
            
            error_msg += "\n\n⚠️ Registre estos clientes antes de importar."
            
            self.validation_text.setStyleSheet("""
                QTextEdit {
                    border: 1px solid #e74c3c;
                    border-radius: 5px;
                    padding: 10px;
                    background-color: #fadbd8;
                    color: #e74c3c;
                    font-family: 'Courier New';
                    font-size: 12px;
                }
            """)
            self.validation_text.setText(error_msg)
            self.import_button.setEnabled(False)
            self.validation_passed = False
            self.show_preview(df)
            self.legend_frame.show()
            
            self.cancel_button.show()
            self.import_button.show()
            return
        
        # ✅ TODO VÁLIDO
        self.validation_text.setStyleSheet("""
            QTextEdit {
                border: 1px solid #27ae60;
                border-radius: 5px;
                padding: 10px;
                background-color: #d5f4e6;
                color: #27ae60;
                font-family: 'Courier New';
                font-size: 12px;
            }
        """)
        self.validation_text.setText(
            f"✅ Archivo válido\n"
            f"✅ {validacion_clientes['message']}\n\n"
            f"Listo para importar {len(df)} registros."
        )
        self.import_button.setEnabled(True)
        self.validation_passed = True
        self.show_preview(df)
        self.legend_frame.show()
        
        self.cancel_button.show()
        self.import_button.show()
    
    # MODIFICAR el método show_preview() completo:

    def show_preview(self, df):
        if df is None or df.empty:
            return
        
        from PyQt5.QtGui import QColor
        
        preview_df = df.head(10)
        self.preview_table.setRowCount(len(preview_df))
        self.preview_table.setColumnCount(len(preview_df.columns))
        self.preview_table.setHorizontalHeaderLabels(preview_df.columns.tolist())
        
        for i, row in enumerate(preview_df.itertuples(index=False)):
            for j, value in enumerate(row):
                item = QTableWidgetItem(str(value))
                
                # ← NUEVO: Aplicar colores según tipo de columna
                col_name = preview_df.columns[j]
                
                # Colores para montos
                if 'monto' in col_name.lower():
                    item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                    item.setBackground(QColor(232, 245, 233))  # Verde claro
                    item.setForeground(QColor(46, 125, 50))     # Verde oscuro
                
                # Colores para factores
                elif 'factor' in col_name.lower():
                    factor_num = int(col_name.split('_')[1]) if '_' in col_name else 0
                    
                    # Factores 8-19 (críticos) - Amarillo
                    if 8 <= factor_num <= 19:
                        item.setBackground(QColor(255, 249, 196))  # Amarillo claro
                        item.setForeground(QColor(245, 124, 0))    # Naranja
                        item.setTextAlignment(Qt.AlignCenter)
                    else:
                        # Factores normales - Azul claro
                        item.setBackground(QColor(227, 242, 253))  # Azul claro
                        item.setForeground(QColor(13, 71, 161))    # Azul oscuro
                        item.setTextAlignment(Qt.AlignCenter)
                
                # Colores para fechas
                elif 'fecha' in col_name.lower():
                    item.setBackground(QColor(243, 229, 245))  # Púrpura claro
                    item.setForeground(QColor(106, 27, 154))   # Púrpura oscuro
                    item.setTextAlignment(Qt.AlignCenter)
                
                # Colores para RUT
                elif 'rut' in col_name.lower():
                    item.setBackground(QColor(224, 247, 250))  # Cyan claro
                    item.setForeground(QColor(0, 96, 100))     # Cyan oscuro
                
                # Colores para tipo/país
                elif col_name.lower() in ['tipo_impuesto', 'pais']:
                    item.setBackground(QColor(255, 243, 224))  # Naranja claro
                    item.setForeground(QColor(230, 81, 0))     # Naranja oscuro
                    item.setTextAlignment(Qt.AlignCenter)
                
                self.preview_table.setItem(i, j, item)
        
        # ← NUEVO: Ajustar ancho de columnas según contenido
        self.preview_table.resizeColumnsToContents()
        
        # ← NUEVO: Validar suma de factores 8-19 y colorear filas
        for i in range(len(preview_df)):
            suma_factores = sum(
                float(preview_df.iloc[i][f'factor_{j}']) 
                for j in range(8, 20)
            )
            
            # Si la suma > 1.0, colorear toda la fila de rojo
            if suma_factores > 1.0:
                for j in range(len(preview_df.columns)):
                    item = self.preview_table.item(i, j)
                    if item:
                        item.setBackground(QColor(255, 205, 210))  # Rojo claro
                        item.setForeground(QColor(198, 40, 40))    # Rojo oscuro
        
        self.preview_label.show()
        self.preview_table.show()
    
    def download_template(self):
        from utils.csvValidator import CSVValidator
        
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Guardar Plantilla",
            "plantilla_carga_masiva.csv",
            "CSV (*.csv);;Excel (*.xlsx)"
        )
        
        if file_path:
            success = CSVValidator.export_template(file_path)
            if success:
                QMessageBox.information(self, "Plantilla Descargada",
                    f"Plantilla guardada en:\n{file_path}")
    
    # MODIFICAR el método import_data() completo:

    def import_data(self):
        """Importa los datos validados a Firebase"""
        if not self.validation_passed or self.dataframe is None:
            return
        
        from services.massiveLoadService import CargaMasivaService
        
        # ← YA NO NECESITA validar clientes aquí (ya se validó antes)
        
        # Confirmar importación
        reply = QMessageBox.question(
            self, "Confirmar Importación",
            f"¿Desea importar {len(self.dataframe)} registros?",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.No:
            return
        
        # Mostrar barra de progreso
        self.progress_bar.show()
        self.progress_bar.setValue(0)
        self.import_button.setEnabled(False)
        
        service = CargaMasivaService()
        
        def update_progress(value):
            self.progress_bar.setValue(value)
        
        # Importar
        result = service.import_data(
            self.dataframe,
            self.user_data.get("_id"),
            progress_callback=update_progress
        )
        
        # Mostrar resultados
        if result["success"]:
            # ← NUEVO: Actualizar área de validación con resumen
            self.validation_text.setStyleSheet("""
                QTextEdit {
                    border: 1px solid #27ae60;
                    border-radius: 5px;
                    padding: 10px;
                    background-color: #d5f4e6;
                    color: #27ae60;
                    font-family: 'Courier New';
                    font-size: 12px;
                    font-weight: bold;
                }
            """)
            
            resumen = f"✅ IMPORTACIÓN COMPLETADA\n\n"
            resumen += f"📊 Creados: {result['created']}\n"
            resumen += f"🔄 Actualizados: {result['updated']}\n"
            resumen += f"❌ Errores: {len(result['errors'])}\n\n"
            resumen += f"Total procesado: {result['total_processed']} registros"
            
            self.validation_text.setText(resumen)
            
            # Mostrar notificación
            msg = f"Importación completada!\n\n"
            msg += f"✅ Creados: {result['created']}\n"
            msg += f"🔄 Actualizados: {result['updated']}\n"
            msg += f"❌ Errores: {len(result['errors'])}"
            QMessageBox.information(self, "Importación Exitosa", msg)
            
            # ← NUEVO: Auto-reset con delay de 2 segundos
            from PyQt5.QtCore import QTimer
            QTimer.singleShot(2000, self.reset_form)
            
        else:
            # Mostrar errores
            self.validation_text.setStyleSheet("""
                QTextEdit {
                    border: 1px solid #e74c3c;
                    border-radius: 5px;
                    padding: 10px;
                    background-color: #fadbd8;
                    color: #e74c3c;
                    font-family: 'Courier New';
                    font-size: 12px;
                }
            """)
            
            error_detail = "\n".join(result['errors'][:10]) if result['errors'] else "Error desconocido"
            if len(result['errors']) > 10:
                error_detail += f"\n... y {len(result['errors']) - 10} errores más."
            
            self.validation_text.setText(f"❌ IMPORTACIÓN CON ERRORES\n\n{error_detail}")
            
            QMessageBox.warning(
                self, 
                "Importación con errores", 
                f"Se procesaron {result['total_processed']} registros.\n\n"
                f"Errores:\n{error_detail}"
            )
            
            # ← NUEVO: En caso de error, esperar 5 segundos antes de resetear
            from PyQt5.QtCore import QTimer
            QTimer.singleShot(5000, self.reset_form)
        
        self.import_button.setEnabled(True)

    def reset_form(self):
        self.file_path = None
        self.dataframe = None
        self.validation_passed = False
        
        self.help_frame.show()  # Mostrar panel de ayuda nuevamente
        self.info_frame.hide()
        self.preview_label.hide()
        self.preview_table.hide()
        self.validation_label.hide()
        self.validation_text.hide()
        self.progress_bar.hide()
        self.cancel_button.hide()
        self.import_button.hide()
        self.legend_frame.hide()