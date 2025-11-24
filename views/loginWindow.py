import os
from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
    QLabel, QLineEdit, QPushButton, QMessageBox, QFrame
)
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QFont, QPixmap, QIcon
from config.settings import Settings


class LoginWindow(QMainWindow):
    """Ventana de inicio de sesión"""
    
    # Señal que se emite cuando el login es exitoso
    login_successful = pyqtSignal(dict)  # Emite datos del usuario
    
    def __init__(self):
        super().__init__()
        self.init_ui()
    
    def init_ui(self):
        """Inicializa la interfaz de usuario"""
        self.setWindowTitle(f"{Settings.APP_NAME} - Inicio de Sesión")
        self.setFixedSize(450, 550)
        self.setStyleSheet(self.get_stylesheet())
        
        # Widget central
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Layout principal
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(40, 40, 40, 40)
        main_layout.setSpacing(20)
        
        # Agregar componentes
        self.add_header(main_layout)
        self.add_form(main_layout)
        self.add_buttons(main_layout)
        self.add_footer(main_layout)
        
        central_widget.setLayout(main_layout)
        
        # Centrar ventana en pantalla
        self.center_on_screen()
    
    def add_header(self, layout):
        """Agrega el encabezado con logo y título"""
        header_layout = QVBoxLayout()
        header_layout.setSpacing(10)
        
        # Logo
        logo_label = QLabel()
        logo_label.setAlignment(Qt.AlignCenter)
        
        # Intentar cargar logo.png
        logo_path = os.path.join(Settings.IMAGES_DIR, "logo.png")
        if os.path.exists(logo_path):
            pixmap = QPixmap(logo_path)
            # Escalar a tamaño apropiado manteniendo aspecto
            scaled_pixmap = pixmap.scaled(120, 120, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            logo_label.setPixmap(scaled_pixmap)
        else:
            # Fallback a emoji si no existe el logo
            logo_label.setStyleSheet("font-size: 48px; color: #2c3e50;")
            logo_label.setText("🏛️")
        
        header_layout.addWidget(logo_label)
        
        # Título
        title_label = QLabel("TribunSys")
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setFont(QFont("Arial", 24, QFont.Bold))
        title_label.setStyleSheet("color: #2c3e50;")
        header_layout.addWidget(title_label)
        
        # Subtítulo
        subtitle_label = QLabel("Sistema de Gestión Tributaria")
        subtitle_label.setAlignment(Qt.AlignCenter)
        subtitle_label.setFont(QFont("Arial", 10))
        subtitle_label.setStyleSheet("color: #7f8c8d;")
        header_layout.addWidget(subtitle_label)
        
        layout.addLayout(header_layout)
        
        # Separador
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setStyleSheet("background-color: #ecf0f1;")
        layout.addWidget(separator)
    
    def add_form(self, layout):
        """Agrega el formulario de login"""
        form_layout = QVBoxLayout()
        form_layout.setSpacing(15)
        
        # Campo RUT
        rut_label = QLabel("RUT:")
        rut_label.setFont(QFont("Arial", 10, QFont.Bold))
        form_layout.addWidget(rut_label)
        
        self.rut_input = QLineEdit()
        self.rut_input.setPlaceholderText("Ej: 12345678-9")
        self.rut_input.setMaxLength(10)
        self.rut_input.setMinimumHeight(25)  # Altura mínima para evitar corte
        self.rut_input.textChanged.connect(self.format_rut_input)
        form_layout.addWidget(self.rut_input)
        
        # Campo Contraseña
        password_label = QLabel("Contraseña:")
        password_label.setFont(QFont("Arial", 10, QFont.Bold))
        form_layout.addWidget(password_label)
        
        self.password_input = QLineEdit()
        self.password_input.setPlaceholderText("Ingrese su contraseña")
        self.password_input.setEchoMode(QLineEdit.Password)
        self.password_input.setMinimumHeight(25)  # Altura mínima para evitar corte
        form_layout.addWidget(self.password_input)
        
        # Conectar Enter para login
        self.password_input.returnPressed.connect(self.on_login_clicked)
        
        layout.addLayout(form_layout)
    
    def add_buttons(self, layout):
        """Agrega los botones de acción"""
        buttons_layout = QVBoxLayout()
        buttons_layout.setSpacing(10)
        
        # Botón Iniciar Sesión
        self.login_button = QPushButton("Iniciar Sesión")
        self.login_button.setMinimumHeight(40)
        self.login_button.setCursor(Qt.PointingHandCursor)
        self.login_button.clicked.connect(self.on_login_clicked)
        buttons_layout.addWidget(self.login_button)
        
        # Layout para botones secundarios
        secondary_buttons = QHBoxLayout()
        
        # Botón Olvidé mi contraseña
        self.forgot_password_button = QPushButton("¿Olvidó su contraseña?")
        self.forgot_password_button.setFlat(True)
        self.forgot_password_button.setCursor(Qt.PointingHandCursor)
        self.forgot_password_button.clicked.connect(self.on_forgot_password_clicked)
        secondary_buttons.addWidget(self.forgot_password_button)
        
        buttons_layout.addLayout(secondary_buttons)
        layout.addLayout(buttons_layout)
    
    def add_footer(self, layout):
        """Agrega el pie de página"""
        layout.addStretch()
        
        footer_label = QLabel(f"© 2025 Nuam | Versión {Settings.APP_VERSION}")
        footer_label.setAlignment(Qt.AlignCenter)
        footer_label.setFont(QFont("Arial", 8))
        footer_label.setStyleSheet("color: #95a5a6;")
        layout.addWidget(footer_label)
    
    def on_login_clicked(self):
        """Maneja el clic en el botón de login"""
        rut = self.rut_input.text().strip().upper()
        password = self.password_input.text()
        
        # Validaciones básicas
        if not rut:
            self.show_error("Por favor ingrese su RUT")
            self.rut_input.setFocus()
            return
        
        if not password:
            self.show_error("Por favor ingrese su contraseña")
            self.password_input.setFocus()
            return
        
        # Deshabilitar botón mientras se procesa
        self.login_button.setEnabled(False)
        self.login_button.setText("Iniciando sesión...")
        
        # Aquí el controlador tomará el control
        # Por ahora solo emitimos la señal con los datos
        self.attempt_login(rut, password)
    
    def attempt_login(self, rut, password):
        """
        Intenta realizar el login
        Esta función será sobrescrita por el controlador
        """
        # Placeholder - el controlador implementará la lógica real
        pass
    
    def on_forgot_password_clicked(self):
        """Maneja el clic en olvidé mi contraseña"""
        QMessageBox.information(
            self,
            "Recuperar Contraseña",
            "Por favor contacte al administrador del sistema para recuperar su contraseña.\n\n"
            "Email: soporte@nuam.cl"
        )
    
    def show_error(self, message):
        """Muestra un mensaje de error"""
        QMessageBox.critical(self, "Error de Autenticación", message)
        self.reset_login_button()
    
    def show_success(self, message):
        """Muestra un mensaje de éxito"""
        QMessageBox.information(self, "Éxito", message)
    
    def reset_login_button(self):
        """Restablece el estado del botón de login"""
        self.login_button.setEnabled(True)
        self.login_button.setText("Iniciar Sesión")
    
    def clear_form(self):
        """Limpia los campos del formulario"""
        self.rut_input.clear()
        self.password_input.clear()
        self.rut_input.setFocus()
    
    def center_on_screen(self):
        """Centra la ventana en la pantalla"""
        screen = self.screen().geometry()
        window = self.geometry()
        x = (screen.width() - window.width()) // 2
        y = (screen.height() - window.height()) // 2
        self.move(x, y)
    
    def get_stylesheet(self):
        """Retorna el stylesheet de la ventana"""
        return """
            QMainWindow {
                background-color: #f5f6fa;
            }
            
            QLabel {
                color: #2c3e50;
            }
            
            QLineEdit {
                padding: 10px;
                border: 2px solid #dfe6e9;
                border-radius: 5px;
                background-color: white;
                font-size: 12px;
                min-height: 25px;
            }
            
            QLineEdit:focus {
                border: 2px solid #3498db;
            }
            
            QPushButton {
                background-color: #3498db;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 10px;
                font-size: 12px;
                font-weight: bold;
            }
            
            QPushButton:hover {
                background-color: #2980b9;
            }
            
            QPushButton:pressed {
                background-color: #21618c;
            }
            
            QPushButton:disabled {
                background-color: #bdc3c7;
            }
            
            QPushButton[flat="true"] {
                background-color: transparent;
                color: #3498db;
                text-decoration: underline;
            }
            
            QPushButton[flat="true"]:hover {
                color: #2980b9;
            }
        """
    
    def format_rut_input(self, text):
        """Formatea el RUT mientras se escribe"""
        # Guardar posición del cursor
        cursor_pos = self.rut_input.cursorPosition()
        
        # Limpiar: solo números y K/k
        cleaned = ''.join(c for c in text.upper() if c.isdigit() or c == 'K')
        
        # Limitar longitud (máximo 9: 8 dígitos + 1 DV)
        if len(cleaned) > 9:
            cleaned = cleaned[:9]
        
        # Formatear SOLO si tiene todos los dígitos (8 o 9 caracteres)
        if len(cleaned) >= 9:
            # Separar cuerpo y dígito verificador
            cuerpo = cleaned[:-1]
            dv = cleaned[-1]
            formatted = f"{cuerpo}-{dv}"
        else:
            # Mientras escribe, no formatear
            formatted = cleaned
        
        # Actualizar solo si cambió
        if formatted != text:
            self.rut_input.blockSignals(True)
            self.rut_input.setText(formatted)
            
            # Mantener cursor en posición lógica
            # Si hay guión y el cursor estaba antes, mantener posición
            if '-' in formatted:
                if cursor_pos > len(cuerpo):
                    cursor_pos = len(formatted)
                else:
                    cursor_pos = min(cursor_pos, len(cuerpo))
            else:
                cursor_pos = min(cursor_pos, len(formatted))
            
            self.rut_input.setCursorPosition(cursor_pos)
            self.rut_input.blockSignals(False)