from PyQt5.QtWidgets import QWidget, QHBoxLayout, QLabel, QPushButton
from PyQt5.QtCore import Qt, pyqtSlot
from PyQt5.QtGui import QFont, QCursor
from utils.connectionManager import connection_manager


class ConnectionIndicator(QWidget):
    """
    Widget que muestra el estado de conexión a internet
    Se actualiza automáticamente cuando cambia la conexión
    """
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setup_ui()
        self.connect_signals()
        
        # Actualizar estado inicial
        self.update_status(connection_manager.is_online())
    
    def setup_ui(self):
        """Configura la interfaz del indicador"""
        layout = QHBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)
        
        # Icono de estado
        self.status_icon = QLabel("🟢")
        self.status_icon.setFont(QFont("Arial", 12))
        layout.addWidget(self.status_icon)
        
        # Texto de estado
        self.status_text = QLabel("Online")
        self.status_text.setFont(QFont("Arial", 9))
        self.status_text.setStyleSheet("color: #27ae60;")
        layout.addWidget(self.status_text)
        
        # Botón de reintentar (oculto por defecto)
        self.retry_button = QPushButton("🔄 Reintentar")
        self.retry_button.setFont(QFont("Arial", 8))
        self.retry_button.setCursor(QCursor(Qt.PointingHandCursor))
        self.retry_button.setFixedHeight(25)
        self.retry_button.clicked.connect(self.retry_connection)
        self.retry_button.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                border: none;
                border-radius: 3px;
                padding: 3px 8px;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
        """)
        self.retry_button.hide()
        layout.addWidget(self.retry_button)
        
        self.setLayout(layout)
    
    def connect_signals(self):
        """Conecta las señales del connection manager"""
        connection_manager.connection_lost.connect(self.on_connection_lost)
        connection_manager.connection_restored.connect(self.on_connection_restored)
    
    @pyqtSlot()
    def on_connection_lost(self):
        """Se ejecuta cuando se pierde la conexión"""
        self.update_status(False)
    
    @pyqtSlot()
    def on_connection_restored(self):
        """Se ejecuta cuando se restaura la conexión"""
        self.update_status(True)
    
    def update_status(self, is_online: bool):
        """
        Actualiza la visualización del estado
        
        Args:
            is_online (bool): True si hay conexión
        """
        if is_online:
            self.status_icon.setText("🟢")
            self.status_text.setText("Online")
            self.status_text.setStyleSheet("color: #27ae60;")
            self.retry_button.hide()
        else:
            self.status_icon.setText("🔴")
            self.status_text.setText("Sin conexión")
            self.status_text.setStyleSheet("color: #e74c3c;")
            self.retry_button.show()
    
    def retry_connection(self):
        """Reintenta verificar la conexión"""
        self.status_text.setText("Verificando...")
        self.status_text.setStyleSheet("color: #f39c12;")
        
        # Forzar verificación inmediata
        is_online = connection_manager.is_online()
        self.update_status(is_online)


class ConnectionStatusBar(QWidget):
    """
    Barra de estado expandida que muestra información detallada de conexión
    Se muestra cuando no hay conexión
    """
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setup_ui()
        self.connect_signals()
        self.hide()  # Oculto por defecto
    
    def setup_ui(self):
        """Configura la interfaz de la barra de estado"""
        layout = QHBoxLayout()
        layout.setContentsMargins(15, 10, 15, 10)
        
        # Icono
        icon_label = QLabel("⚠️")
        icon_label.setFont(QFont("Arial", 16))
        layout.addWidget(icon_label)
        
        # Mensaje
        self.message_label = QLabel(
            "Sin conexión a internet. Trabajando en modo offline. "
            "Algunas funciones no están disponibles."
        )
        self.message_label.setFont(QFont("Arial", 10))
        self.message_label.setWordWrap(True)
        layout.addWidget(self.message_label, 1)
        
        # Botón cerrar
        close_button = QPushButton("✕")
        close_button.setFixedSize(25, 25)
        close_button.setCursor(QCursor(Qt.PointingHandCursor))
        close_button.clicked.connect(self.hide)
        close_button.setStyleSheet("""
            QPushButton {
                background-color: transparent;
                border: none;
                color: #856404;
                font-size: 16px;
            }
            QPushButton:hover {
                color: #533f03;
            }
        """)
        layout.addWidget(close_button)
        
        self.setLayout(layout)
        self.setStyleSheet("""
            QWidget {
                background-color: #fff3cd;
                border: 2px solid #ffc107;
                border-radius: 5px;
            }
        """)
    
    def connect_signals(self):
        """Conecta las señales del connection manager"""
        connection_manager.connection_lost.connect(self.show)
        connection_manager.connection_restored.connect(self.on_connection_restored)
    
    @pyqtSlot()
    def on_connection_restored(self):
        """Se ejecuta cuando se restaura la conexión"""
        self.message_label.setText(
            "✅ Conexión restaurada. Todas las funciones están disponibles nuevamente."
        )
        self.setStyleSheet("""
            QWidget {
                background-color: #d5f4e6;
                border: 2px solid #27ae60;
                border-radius: 5px;
            }
        """)
        
        # Auto-ocultar después de 3 segundos
        from PyQt5.QtCore import QTimer
        QTimer.singleShot(3000, self.hide)