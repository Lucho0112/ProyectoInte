from PyQt5.QtWidgets import QPushButton, QVBoxLayout, QLabel, QFrame, QSizePolicy
from PyQt5.QtCore import Qt, pyqtSignal, QPropertyAnimation, QEasingCurve
from PyQt5.QtGui import QFont, QCursor


class CardButton(QFrame):
    """Tarjeta clickeable para módulos del sistema"""
    
    clicked = pyqtSignal(str)  # Emite el ID del módulo
    
    def __init__(self, module_id: str, title: str, icon: str, description: str = "", parent=None):
        super().__init__(parent)
        self.module_id = module_id
        self.title = title
        self.icon = icon
        self.description = description
        
        self.setup_ui()
        self.setup_animations()
    
    def setup_ui(self):
        """Configura la interfaz de la tarjeta"""
        self.setMinimumSize(200, 200)
        self.setMaximumSize(700, 200)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.setCursor(QCursor(Qt.PointingHandCursor))
        self.setObjectName("cardButton")
        
        # Layout principal
        layout = QVBoxLayout()
        layout.setAlignment(Qt.AlignCenter)
        layout.setSpacing(10)
        
        # Icono (por ahora texto, luego será SVG/FontAwesome)
        self.icon_label = QLabel(self.icon)
        self.icon_label.setAlignment(Qt.AlignCenter)
        self.icon_label.setFont(QFont("Arial", 36))
        self.icon_label.setStyleSheet("color: #E94E1B;")
        layout.addWidget(self.icon_label)
        
        # Título
        title_label = QLabel(self.title)
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setFont(QFont("Arial", 11, QFont.Bold))
        title_label.setWordWrap(True)
        title_label.setStyleSheet("color: #2c3e50;")
        layout.addWidget(title_label)
        
        # Descripción (opcional)
        if self.description:
            desc_label = QLabel(self.description)
            desc_label.setAlignment(Qt.AlignCenter)
            desc_label.setFont(QFont("Arial", 8))
            desc_label.setWordWrap(True)
            desc_label.setStyleSheet("color: #7f8c8d;")
            layout.addWidget(desc_label)
        
        self.setLayout(layout)
        self.apply_style()
    
    def setup_animations(self):
        """Configura animaciones de hover"""
        self.default_shadow = "0px 2px 5px rgba(0,0,0,0.1)"
        self.hover_shadow = "0px 8px 20px rgba(233, 78, 27, 0.3)"
    
    def apply_style(self):
        """Aplica estilos CSS a la tarjeta"""
        self.setStyleSheet(f"""
            #cardButton {{
                background-color: white;
                border: 2px solid #ecf0f1;
                border-radius: 10px;
                padding: 15px;
            }}
            #cardButton:hover {{
                border: 2px solid #E94E1B;
                background-color: #fef5f1;
            }}
        """)
    
    def mousePressEvent(self, event):
        """Maneja el clic en la tarjeta"""
        if event.button() == Qt.LeftButton:
            self.clicked.emit(self.module_id)
        super().mousePressEvent(event)
    
    def enterEvent(self, event):
        """Efecto hover - entrada"""
        self.setStyleSheet(f"""
            #cardButton {{
                background-color: #fef5f1;
                border: 2px solid #E94E1B;
                border-radius: 10px;
                padding: 15px;
            }}
        """)
        super().enterEvent(event)
    
    def leaveEvent(self, event):
        """Efecto hover - salida"""
        self.apply_style()
        super().leaveEvent(event)