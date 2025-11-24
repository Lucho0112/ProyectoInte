from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QTableWidget, QTableWidgetItem, QHeaderView, QComboBox,
    QFrame, QMessageBox, QAbstractItemView, QScrollArea,
    QDialog, QFormLayout, QLineEdit, QGroupBox
)
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QFont, QColor, QCursor
from datetime import datetime
from services.authService import AuthService
from utils.logger import app_logger, log_audit


class UserFormDialog(QDialog):
    """Diálogo para crear/editar usuarios"""
    
    def __init__(self, parent, user_data: dict, modo="crear", usuario: dict = None):
        super().__init__(parent)
        self.user_data = user_data
        self.modo = modo
        self.usuario = usuario
        self.service = AuthService()
        
        self.init_ui()
        
        if modo == "editar" and usuario:
            self.cargar_datos()
    
    def init_ui(self):
        titulo = "Nuevo Usuario" if self.modo == "crear" else "Editar Usuario"
        self.setWindowTitle(titulo)
        self.setMinimumSize(600, 500)
        
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)
        
        # Título
        title_label = QLabel(titulo)
        title_label.setFont(QFont("Arial", 16, QFont.Bold))
        title_label.setStyleSheet("color: #2c3e50;")
        main_layout.addWidget(title_label)
        
        # Scroll area para el formulario
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        
        scroll_content = QWidget()
        scroll_layout = QVBoxLayout()
        scroll_layout.setSpacing(15)
        
        # Sección: Información Personal
        self.add_seccion_personal(scroll_layout)
        
        # Sección: Información de Cuenta
        self.add_seccion_cuenta(scroll_layout)
        
        scroll_content.setLayout(scroll_layout)
        scroll.setWidget(scroll_content)
        main_layout.addWidget(scroll)
        
        # Botones
        self.add_buttons(main_layout)
        
        self.setLayout(main_layout)
        self.apply_styles()
    
    def add_seccion_personal(self, layout):
        group = QGroupBox("📝 Información Personal")
        group.setFont(QFont("Arial", 11, QFont.Bold))
        
        form_layout = QFormLayout()
        form_layout.setSpacing(10)
        
        # RUT
        self.input_rut = QLineEdit()
        self.input_rut.setPlaceholderText("Ej: 12345678-9")
        form_layout.addRow("RUT: *", self.input_rut)
        
        # Nombre
        self.input_nombre = QLineEdit()
        self.input_nombre.setPlaceholderText("Nombre")
        form_layout.addRow("Nombre: *", self.input_nombre)
        
        # Apellido Paterno
        self.input_apellido_p = QLineEdit()
        self.input_apellido_p.setPlaceholderText("Apellido Paterno")
        form_layout.addRow("Apellido Paterno: *", self.input_apellido_p)
        
        # Apellido Materno
        self.input_apellido_m = QLineEdit()
        self.input_apellido_m.setPlaceholderText("Apellido Materno")
        form_layout.addRow("Apellido Materno:", self.input_apellido_m)
        
        group.setLayout(form_layout)
        layout.addWidget(group)
    
    def add_seccion_cuenta(self, layout):
        group = QGroupBox("🔐 Información de Cuenta")
        group.setFont(QFont("Arial", 11, QFont.Bold))
        
        form_layout = QFormLayout()
        form_layout.setSpacing(10)
        
        # Correo
        self.input_correo = QLineEdit()
        self.input_correo.setPlaceholderText("correo@ejemplo.com")
        form_layout.addRow("Correo: *", self.input_correo)
        
        # Contraseña (solo en crear)
        if self.modo == "crear":
            self.input_password = QLineEdit()
            self.input_password.setEchoMode(QLineEdit.Password)
            self.input_password.setPlaceholderText("Mínimo 6 caracteres")
            form_layout.addRow("Contraseña: *", self.input_password)
            
            self.input_password_confirm = QLineEdit()
            self.input_password_confirm.setEchoMode(QLineEdit.Password)
            self.input_password_confirm.setPlaceholderText("Repita la contraseña")
            form_layout.addRow("Confirmar: *", self.input_password_confirm)
        
        # Rol
        self.combo_rol = QComboBox()
        self.combo_rol.addItems([
            "cliente",
            "analista_mercado",
            "auditor_tributario",
            "especialista_beneficios",
            "administrador"
        ])
        form_layout.addRow("Rol: *", self.combo_rol)
        
        group.setLayout(form_layout)
        layout.addWidget(group)
    
    def add_buttons(self, layout):
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        
        btn_cancelar = QPushButton("❌ Cancelar")
        btn_cancelar.setMinimumSize(120, 40)
        btn_cancelar.setCursor(QCursor(Qt.PointingHandCursor))
        btn_cancelar.clicked.connect(self.reject)
        btn_cancelar.setProperty("role", "muted")
        button_layout.addWidget(btn_cancelar)
        
        self.btn_guardar = QPushButton("💾 Guardar")
        self.btn_guardar.setMinimumSize(120, 40)
        self.btn_guardar.setCursor(QCursor(Qt.PointingHandCursor))
        self.btn_guardar.clicked.connect(self.guardar)
        self.btn_guardar.setProperty("role", "primary")
        button_layout.addWidget(self.btn_guardar)
        
        layout.addLayout(button_layout)
    
    def cargar_datos(self):
        """Carga los datos del usuario en el formulario"""
        if not self.usuario:
            return
        
        self.input_rut.setText(self.usuario.get("rut", ""))
        self.input_rut.setEnabled(False)  # RUT no se puede cambiar
        
        self.input_nombre.setText(self.usuario.get("nombre", ""))
        self.input_apellido_p.setText(self.usuario.get("apellido_P", ""))
        self.input_apellido_m.setText(self.usuario.get("apellido_M", ""))
        self.input_correo.setText(self.usuario.get("correo", ""))
        
        # Seleccionar rol
        rol = self.usuario.get("rol", "cliente")
        index = self.combo_rol.findText(rol)
        if index >= 0:
            self.combo_rol.setCurrentIndex(index)
    
    def guardar(self):
        """Valida y guarda el usuario"""
        try:
            # Validaciones básicas
            rut = self.input_rut.text().strip()
            nombre = self.input_nombre.text().strip()
            apellido_p = self.input_apellido_p.text().strip()
            correo = self.input_correo.text().strip()
            
            if not all([rut, nombre, apellido_p, correo]):
                QMessageBox.warning(
                    self,
                    "Validación",
                    "Los campos marcados con * son obligatorios"
                )
                return
            
            # Validar contraseña en modo crear
            if self.modo == "crear":
                password = self.input_password.text()
                password_confirm = self.input_password_confirm.text()
                
                if not password:
                    QMessageBox.warning(self, "Validación", "La contraseña es obligatoria")
                    return
                
                if len(password) < 6:
                    QMessageBox.warning(
                        self,
                        "Validación",
                        "La contraseña debe tener al menos 6 caracteres"
                    )
                    return
                
                if password != password_confirm:
                    QMessageBox.warning(
                        self,
                        "Validación",
                        "Las contraseñas no coinciden"
                    )
                    return
            
            # Preparar datos
            user_data = {
                "rut": rut,
                "nombre": nombre,
                "apellido_P": apellido_p,
                "apellido_M": self.input_apellido_m.text().strip(),
                "correo": correo,
                "rol": self.combo_rol.currentText()
            }
            
            if self.modo == "crear":
                user_data["password"] = self.input_password.text()
                
                # Llamar al método existente en AuthService (se llama 'register', no 'register_user')
                result = self.service.register(user_data)
                
                if result.get("success"):
                    # Registrar auditoría
                    log_audit(
                        action="USUARIO_CREADO",
                        user_id=self.user_data.get("_id"),
                        details={"rut": rut, "rol": user_data["rol"]}
                    )
                    
                    QMessageBox.information(
                        self,
                        "Éxito",
                        f"Usuario creado exitosamente\nRUT: {rut}"
                    )
                    self.accept()
                else:
                    QMessageBox.warning(self, "Error", result.get("message", "Error desconocido"))
            
            else:
                # EDITAR: usar AuthService.update_user
                try:
                    user_id = self.usuario.get("_id")
                    updates = {
                        "nombre": user_data["nombre"],
                        "apellido_P": user_data["apellido_P"],
                        "apellido_M": user_data["apellido_M"],
                        "correo": user_data["correo"],
                        "rol": user_data["rol"]
                    }
                    # Si desea permitirse cambiar contraseña en edición, añadir aquí:
                    # if hasattr(self, 'input_password') and self.input_password.text():
                    #     updates['contraseña'] = self.input_password.text()
                    
                    result = self.service.update_user(user_id, updates)
                    if result.get("success"):
                        log_audit(
                            action="USUARIO_EDITADO",
                            user_id=self.user_data.get("_id"),
                            details={"usuario_id": user_id, "fields": list(updates.keys())}
                        )
                        QMessageBox.information(self, "Éxito", "Usuario actualizado correctamente")
                        self.accept()
                    else:
                        QMessageBox.warning(self, "Error", result.get("message", "Error al actualizar usuario"))
                except Exception as e:
                    app_logger.error(f"Error actualizando usuario en UI: {e}")
                    QMessageBox.critical(self, "Error", f"Error al actualizar usuario: {str(e)}")
        
        except Exception as e:
            app_logger.error(f"Error al guardar usuario: {str(e)}")
            QMessageBox.critical(self, "Error", f"Error al guardar: {str(e)}")
    
    def apply_styles(self):
        self.setStyleSheet("""
            QDialog {
                background-color: #f2f4f7;
            }
            
            QGroupBox {
                font-weight: bold;
                color: #2c3e50;
                border: 1px solid #e6e9ee;
                border-radius: 6px;
                margin-top: 8px;
                background-color: #ffffff;
                padding: 15px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 4px 8px;
                color: #2c3e50;
            }
            
            QLineEdit, QComboBox {
                padding: 8px 10px;
                border: 1px solid #d7dfe8;
                border-radius: 4px;
                background-color: #ffffff;
                color: #2c3e50;
                min-height: 30px;
            }
            
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
        """)
        

class GestionUsuariosContent(QWidget):
    """Contenido de gestión de usuarios"""
    back_requested = pyqtSignal()
    
    def __init__(self, user_data: dict, parent=None):
        super().__init__(parent)
        self.user_data = user_data
        self.user_rol = user_data.get("rol", "cliente")
        self.service = AuthService()
        self.usuarios = []
        
        self.init_ui()
        self.refrescar_tabla()
    
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
        
        title = QLabel("👥 Gestión de Usuarios")
        title.setFont(QFont("Arial", 18, QFont.Bold))
        title.setStyleSheet("color: #2c3e50;")
        header_layout.addWidget(title)
        
        header_layout.addStretch()
        layout.addLayout(header_layout)
    
    def add_toolbar(self, layout):
        """Barra de herramientas"""
        toolbar_layout = QHBoxLayout()
        toolbar_layout.setSpacing(10)
        
        btn_nuevo = QPushButton("➕ Nuevo Usuario")
        btn_nuevo.setFont(QFont("Arial", 10, QFont.Bold))
        btn_nuevo.setMinimumHeight(40)
        btn_nuevo.setCursor(QCursor(Qt.PointingHandCursor))
        btn_nuevo.clicked.connect(self.abrir_formulario_crear)
        btn_nuevo.setProperty("role", "primary")
        toolbar_layout.addWidget(btn_nuevo)
        
        btn_refrescar = QPushButton("🔄 Refrescar")
        btn_refrescar.setFont(QFont("Arial", 10))
        btn_refrescar.setMinimumHeight(40)
        btn_refrescar.setCursor(QCursor(Qt.PointingHandCursor))
        btn_refrescar.clicked.connect(self.refrescar_tabla)
        btn_refrescar.setProperty("role", "secondary")
        toolbar_layout.addWidget(btn_refrescar)
        
        toolbar_layout.addStretch()
        
        self.label_contador = QLabel("Total: 0 usuarios")
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
        
        filter_layout = QHBoxLayout()
        filter_layout.setSpacing(15)
        
        filter_layout.addWidget(QLabel("Filtrar por rol:"))
        self.combo_filtro_rol = QComboBox()
        self.combo_filtro_rol.addItems([
            "Todos",
            "cliente",
            "analista_mercado",
            "auditor_tributario",
            "especialista_beneficios",
            "administrador"
        ])
        self.combo_filtro_rol.setMinimumHeight(35)
        filter_layout.addWidget(self.combo_filtro_rol)
        
        filter_layout.addStretch()
        
        btn_filtrar = QPushButton("🔍 Filtrar")
        btn_filtrar.setCursor(QCursor(Qt.PointingHandCursor))
        btn_filtrar.clicked.connect(self.aplicar_filtros)
        btn_filtrar.setProperty("role", "success")
        filter_layout.addWidget(btn_filtrar)
        
        btn_limpiar = QPushButton("🗑️ Limpiar")
        btn_limpiar.setCursor(QCursor(Qt.PointingHandCursor))
        btn_limpiar.clicked.connect(self.limpiar_filtros)
        btn_limpiar.setProperty("role", "muted")
        filter_layout.addWidget(btn_limpiar)
        
        filter_frame.setLayout(filter_layout)
        layout.addWidget(filter_frame)
    
    def add_table(self, layout):
        """Tabla de usuarios"""
        self.table = QTableWidget()
        self.table.setColumnCount(8)
        self.table.setHorizontalHeaderLabels([
            "RUT", "Nombre Completo", "Correo", "Rol",
            "Fecha Registro", "Último Acceso", "Editar", "Activo"
        ])
        
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        
        header = self.table.horizontalHeader()
        for i in range(8):
            header.setSectionResizeMode(i, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        
        layout.addWidget(self.table)
    
    def add_footer(self, layout):
        """Footer informativo"""
        footer_text = "💡 Solo los administradores pueden crear, editar y desactivar usuarios del sistema"
        
        footer = QLabel(footer_text)
        footer.setFont(QFont("Arial", 9))
        footer.setStyleSheet("color: #7f8c8d;")
        layout.addWidget(footer)
    
    def actualizar_tabla(self, usuarios: list):
        """Actualiza la tabla con los usuarios"""
        self.usuarios = usuarios
        self.table.setRowCount(len(usuarios))
        
        for row, usuario in enumerate(usuarios):
            # RUT
            self.table.setItem(row, 0, QTableWidgetItem(usuario.get("rut", "")))
            
            # Nombre completo
            nombre_completo = f"{usuario.get('nombre', '')} {usuario.get('apellido_P', '')} {usuario.get('apellido_M', '')}".strip()
            self.table.setItem(row, 1, QTableWidgetItem(nombre_completo))
            
            # Correo
            self.table.setItem(row, 2, QTableWidgetItem(usuario.get("correo", "")))
            
            # Rol
            rol = usuario.get("rol", "")
            item_rol = QTableWidgetItem(rol)
            item_rol.setTextAlignment(Qt.AlignCenter)
            
            # Colorear según rol
            if rol == "administrador":
                item_rol.setBackground(QColor(255, 200, 200))
                item_rol.setForeground(QColor(150, 0, 0))
            elif rol in ["analista_mercado", "auditor_tributario", "especialista_beneficios"]:
                item_rol.setBackground(QColor(200, 230, 255))
                item_rol.setForeground(QColor(0, 100, 200))
            else:
                item_rol.setBackground(QColor(230, 230, 230))
            
            self.table.setItem(row, 3, item_rol)
            
            # Fecha de registro
            fecha_registro = usuario.get("fechaRegistro")
            if fecha_registro:
                if hasattr(fecha_registro, 'strftime'):
                    fecha_str = fecha_registro.strftime("%Y-%m-%d")
                else:
                    fecha_str = str(fecha_registro)[:10]
            else:
                fecha_str = "N/A"
            self.table.setItem(row, 4, QTableWidgetItem(fecha_str))
            
            # Último acceso
            ultimo_acceso = usuario.get("ultimoAcceso")
            if ultimo_acceso:
                if hasattr(ultimo_acceso, 'strftime'):
                    acceso_str = ultimo_acceso.strftime("%Y-%m-%d %H:%M")
                else:
                    acceso_str = str(ultimo_acceso)
            else:
                acceso_str = "Nunca"
            self.table.setItem(row, 5, QTableWidgetItem(acceso_str))
            
            # Botón editar
            btn_editar = QPushButton("✏️")
            btn_editar.setCursor(QCursor(Qt.PointingHandCursor))
            btn_editar.clicked.connect(lambda checked, u=usuario: self.abrir_formulario_editar(u))
            self.table.setCellWidget(row, 6, btn_editar)
            
            # Activo / Inactivo - mostrar control de toggle
            is_active = usuario.get("activo", True)
            if is_active:
                btn_toggle = QPushButton("🚫")  # desactivar
                btn_toggle.setToolTip("Desactivar usuario")
            else:
                btn_toggle = QPushButton("♻️")  # reactivar
                btn_toggle.setToolTip("Reactivar usuario")
            
            btn_toggle.setCursor(QCursor(Qt.PointingHandCursor))
            btn_toggle.clicked.connect(lambda checked, u=usuario: self.toggle_usuario_activo(u))
            self.table.setCellWidget(row, 7, btn_toggle)
            
            # Visualmente marcar filas inactivas
            if not is_active:
                for c in range(self.table.columnCount()):
                    item = self.table.item(row, c)
                    if item:
                        item.setForeground(QColor(140, 140, 140))
                        font = item.font()
                        font.setItalic(True)
                        item.setFont(font)
        
        self.label_contador.setText(f"Total: {len(usuarios)} usuarios")
    
    def abrir_formulario_crear(self):
        """Abre el formulario para crear usuario"""
        dialog = UserFormDialog(self, self.user_data, modo="crear")
        if dialog.exec_():
            self.refrescar_tabla()
    
    def abrir_formulario_editar(self, usuario: dict):
        """Abre el formulario para editar usuario"""
        dialog = UserFormDialog(self, self.user_data, modo="editar", usuario=usuario)
        if dialog.exec_():
            self.refrescar_tabla()
    
    def toggle_usuario_activo(self, usuario: dict):
        """Alterna el estado activo/inactivo de un usuario (desactivar/reactivar)."""
        if not usuario:
            return
        user_id = usuario.get("_id")
        if not user_id:
            QMessageBox.warning(self, "Error", "No se pudo identificar el usuario seleccionado.")
            return
        
        is_active = usuario.get("activo", True)
        # Preparar dialog de confirmación con estilo claro (evitar cuadros negros)
        msg = QMessageBox(self)
        msg.setIcon(QMessageBox.Question)
        msg.setWindowTitle("Confirmar")
        if is_active:
            msg.setText(f"¿Está seguro de desactivar al usuario {usuario.get('nombre', '')}?")
        else:
            msg.setText(f"¿Desea reactivar al usuario {usuario.get('nombre', '')}?")
        msg.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        msg.setStyleSheet("QMessageBox { background-color: white; } QLabel { color: #222; }")
        resp = msg.exec_()
        if resp != QMessageBox.Yes:
            return
        
        try:
            if is_active:
                result = self.service.deactivate_user(user_id)
                action = "USUARIO_DESACTIVADO"
                success_msg = "Usuario desactivado correctamente"
            else:
                result = self.service.reactivate_user(user_id)
                action = "USUARIO_REACTIVADO"
                success_msg = "Usuario reactivado correctamente"
            
            if result.get("success"):
                log_audit(action, self.user_data.get("_id"), {"usuario_id": user_id})
                QMessageBox.information(self, "Éxito", success_msg)
                self.refrescar_tabla()
            else:
                QMessageBox.warning(self, "Error", result.get("message", "Error en la operación"))
        except Exception as e:
            app_logger.error(f"Error cambiando estado usuario: {e}")
            QMessageBox.critical(self, "Error", "Ocurrió un error al cambiar el estado del usuario. Consulte logs.")
    
    def aplicar_filtros(self):
        """Aplica filtros a la lista de usuarios"""
        rol_filtro = self.combo_filtro_rol.currentText()
        
        if rol_filtro == "Todos":
            self.refrescar_tabla()
        else:
            usuarios_filtrados = [
                u for u in self.usuarios 
                if u.get("rol") == rol_filtro
            ]
            self.actualizar_tabla(usuarios_filtrados)
    
    def limpiar_filtros(self):
        """Limpia los filtros"""
        self.combo_filtro_rol.setCurrentIndex(0)
        self.refrescar_tabla()
    
    def refrescar_tabla(self):
        """Refresca la tabla de usuarios"""
        try:
            # Obtener todos los usuarios desde Firebase
            usuarios_ref = self.service.usuarios_ref
            docs = usuarios_ref.stream()
            
            usuarios = []
            for doc in docs:
                data = doc.to_dict()
                data["_id"] = doc.id
                usuarios.append(data)
            
            # Mostrar todos (activos e inactivos) para permitir reactivación
            self.actualizar_tabla(usuarios)
        
        except Exception as e:
            app_logger.error(f"Error al cargar usuarios: {str(e)}")
            QMessageBox.warning(
                self,
                "Error",
                "No se pudieron cargar los usuarios. Verifique su conexión."
            )
    
    def apply_styles(self):
        """Aplica estilos consistentes"""
        self.setStyleSheet("""
            QWidget {
                background-color: transparent;
            }
            
            QFrame {
                background-color: #ffffff;
                border: 1px solid #e6e9ee;
                border-radius: 8px;
            }
            
            QLineEdit, QComboBox {
                padding: 8px 10px;
                border: 1px solid #d7dfe8;
                border-radius: 4px;
                background-color: #ffffff;
                color: #2c3e50;
            }
            
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
                padding: 8px 12px;
                border-radius: 4px;
                border: none;
            }
            QPushButton[role="success"]:hover {
                background-color: #229954;
            }
            
            QPushButton[role="muted"] {
                background-color: #95a5a6;
                color: white;
                padding: 8px 12px;
                border-radius: 4px;
                border: none;
            }
            QPushButton[role="muted"]:hover {
                background-color: #7f8c8d;
            }
            
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
            
            QLabel {
                color: #2c3e50;
            }
        """)