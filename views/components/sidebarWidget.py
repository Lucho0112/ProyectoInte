from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QPushButton, QLabel, 
    QFrame, QScrollArea, QSizePolicy
)
from PyQt5.QtCore import Qt, pyqtSignal, QPropertyAnimation, QEasingCurve, QSize
from PyQt5.QtGui import QFont, QCursor


class SidebarButton(QPushButton):
    """Botón personalizado para el sidebar"""
    
    def __init__(self, text: str, icon: str, module_id: str, parent=None):
        super().__init__(parent)
        self.module_id = module_id
        self.icon = icon
        self.full_text = text
        self.setText(f"{icon}  {text}")
        self.setFont(QFont("Arial", 10))
        self.setCursor(QCursor(Qt.PointingHandCursor))
        self.setMinimumHeight(45)
        self.setup_style()
    
    def setup_style(self):
        """Aplica estilos al botón"""
        self.setStyleSheet("""
            QPushButton {
                background-color: transparent;
                border: none;
                text-align: left;
                padding: 12px 15px;
                color: #2c3e50;
            }
            QPushButton:hover {
                background-color: #fef5f1;
                border-left: 3px solid #E94E1B;
            }
            QPushButton:pressed {
                background-color: #fde8df;
            }
        """)


class SidebarWidget(QWidget):
    """Widget de sidebar colapsable"""
    
    # Señales
    module_selected = pyqtSignal(str)  # Emite ID del módulo
    
    # Dimensiones
    EXPANDED_WIDTH = 250  # Aumentado para texto completo
    COLLAPSED_WIDTH = 66
    
    def __init__(self, user_role: str, parent=None):
        super().__init__(parent)
        self.user_role = user_role
        self.is_expanded = False
        
        self.setup_ui()
        self.load_menu_items()

        self.setFixedWidth(self.COLLAPSED_WIDTH)
        self._collapse_buttons_text()

    def _collapse_buttons_text(self):
        for i in range(self.buttons_layout.count()):
            widget = self.buttons_layout.itemAt(i).widget()
            if isinstance(widget, SidebarButton):
                widget.setText(widget.icon)
    
    def setup_ui(self):
        """Configura la interfaz del sidebar"""
        self.setStyleSheet("""
            QWidget {
                background-color: #f8f9fa;
                border-right: 1px solid #dee2e6;
            }
        """)
        
        # Layout principal
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        
        # Botón de colapsar/expandir - ALINEADO A LA IZQUIERDA
        self.toggle_button = QPushButton("☰")
        self.toggle_button.setFont(QFont("Arial", 16))
        self.toggle_button.setFixedHeight(50)
        self.toggle_button.setCursor(QCursor(Qt.PointingHandCursor))
        self.toggle_button.clicked.connect(self.toggle_sidebar)
        self.toggle_button.setStyleSheet("""
            QPushButton {
                background-color: transparent;
                border: none;
                border-bottom: 1px solid #dee2e6;
                text-align: left;
                padding-left: 20px;
            }
            QPushButton:hover {
                background-color: #e9ecef;
            }
        """)
        main_layout.addWidget(self.toggle_button)
        
        # Área de scroll para los botones
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        scroll_area.setStyleSheet("""
            QScrollArea {
                border: none;
                background-color: transparent;
            }
        """)
        
        # Widget contenedor de botones
        self.buttons_widget = QWidget()
        self.buttons_layout = QVBoxLayout()
        self.buttons_layout.setContentsMargins(0, 10, 0, 10)
        self.buttons_layout.setSpacing(5)
        self.buttons_widget.setLayout(self.buttons_layout)
        
        scroll_area.setWidget(self.buttons_widget)
        main_layout.addWidget(scroll_area, 1)
        
        self.setLayout(main_layout)
    
    def load_menu_items(self):
        """Carga los items del menú según el rol"""
        # Definir módulos por rol
        modules = self.get_modules_by_role()
        
        for module in modules:
            btn = SidebarButton(
                text=module["text"],
                icon=module["icon"],
                module_id=module["id"]
            )
            btn.clicked.connect(lambda checked, m=module["id"]: self.on_module_clicked(m))
            self.buttons_layout.addWidget(btn)

        self.buttons_layout.addStretch()
    
    def get_modules_by_role(self):
        """Retorna los módulos disponibles según el rol"""
        from config.roles import ModulosConfig
        
        # Obtener módulos permitidos según el rol
        modules = ModulosConfig.get_modulos_por_rol(self.user_role)
        
        # Convertir al formato que espera el sidebar
        return [
            {
                "id": mod["id"],
                "text": mod["title"],
                "icon": mod["icon"]
            }
            for mod in modules
        ]
    
    def toggle_sidebar(self):
        """Colapsa o expande el sidebar"""
        if self.is_expanded:
            self.collapse()
        else:
            self.expand()
    
    def collapse(self):
        """Colapsa el sidebar (solo iconos)"""
        self.animate_width(self.COLLAPSED_WIDTH)
        self.is_expanded = False
        
        # Ocultar textos de los botones
        for i in range(self.buttons_layout.count()):
            widget = self.buttons_layout.itemAt(i).widget()
            if isinstance(widget, SidebarButton):
                # Mostrar solo el icono
                widget.setText(widget.icon)
    
    def expand(self):
        """Expande el sidebar (iconos + texto)"""
        self.animate_width(self.EXPANDED_WIDTH)
        self.is_expanded = True
        
        # Mostrar textos de los botones
        modules = self.get_modules_by_role()
        for i in range(self.buttons_layout.count()):
            widget = self.buttons_layout.itemAt(i).widget()
            if isinstance(widget, SidebarButton) and i < len(modules):
                widget.setText(f"{modules[i]['icon']}  {modules[i]['text']}")
    
    def animate_width(self, target_width):
        """Anima el cambio de ancho del sidebar"""
        self.animation = QPropertyAnimation(self, b"minimumWidth")
        self.animation.setDuration(300)
        self.animation.setStartValue(self.width())
        self.animation.setEndValue(target_width)
        self.animation.setEasingCurve(QEasingCurve.InOutQuad)
        self.animation.start()
        
        self.animation2 = QPropertyAnimation(self, b"maximumWidth")
        self.animation2.setDuration(300)
        self.animation2.setStartValue(self.width())
        self.animation2.setEndValue(target_width)
        self.animation2.setEasingCurve(QEasingCurve.InOutQuad)
        self.animation2.start()
    
    def on_module_clicked(self, module_id: str):
        """Maneja el clic en un módulo"""
        self.module_selected.emit(module_id)