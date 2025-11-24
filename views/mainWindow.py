import os
from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QFrame, QGridLayout, QScrollArea,
    QMenu, QAction, QSizePolicy, QStackedWidget
)
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QFont, QPixmap, QCursor
from config.settings import Settings
from views.components.sidebarWidget import SidebarWidget
from views.components.cardButton import CardButton
from utils.connectionManager import connection_manager
from views.components.connectionIndicator import ConnectionIndicator, ConnectionStatusBar


class MainWindow(QMainWindow):
    """Ventana principal del sistema"""
    
    # Señales
    logout_requested = pyqtSignal()
    theme_changed = pyqtSignal(str)  # 'light' o 'dark'
    
    def __init__(self, user_data: dict):
        super().__init__()
        self.user_data = user_data
        self.current_theme = "light"

        connection_manager.start_monitoring()

        self.init_ui()
    
    def init_ui(self):
        """Inicializa la interfaz de usuario"""
        self.setWindowTitle(f"{Settings.APP_NAME} - Sistema Principal")
        
        # Tamaño inicial más grande pero NO maximizado
        self.resize(1280, 900)
        
        # Widget central
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Layout principal (horizontal: sidebar + content)
        main_layout = QHBoxLayout()
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        
        # Sidebar
        self.sidebar = SidebarWidget(user_role=self.user_data.get("rol", "cliente"))
        self.sidebar.module_selected.connect(self.on_module_selected)
        main_layout.addWidget(self.sidebar)
        
        # Área de contenido (con stack de widgets)
        content_widget = QWidget()
        self.content_layout = QVBoxLayout()
        self.content_layout.setContentsMargins(0, 0, 0, 0)
        self.content_layout.setSpacing(0)
        
        # Header (SIN LÍNEAS DEBAJO)
        self.add_header(self.content_layout)
        
        # Stack de contenidos (área intercambiable)
        self.content_stack = QStackedWidget()
        self.content_layout.addWidget(self.content_stack)
        
        # Crear páginas
        self.create_home_page()                      # Index 0
        self.create_carga_masiva_page()              # Index 1
        self.create_gestionar_calificaciones_page()  # Index 2
        self.create_gestionar_subsidios_page()       # Index 3
        self.create_reportes_page()                  # Index 4 
        self.create_consultar_page()                 # Index 5 
        self.create_usuarios_page()
        
        # Mostrar página de inicio
        self.content_stack.setCurrentIndex(0)
        
        # Footer (CENTRADO)
        self.add_footer(self.content_layout)
        
        content_widget.setLayout(self.content_layout)
        main_layout.addWidget(content_widget)
        
        central_widget.setLayout(main_layout)
        
        # Aplicar tema
        self.apply_theme()
    
    def create_home_page(self):
        """Crea la página de inicio (banner + cards + resumen)"""
        home_widget = QWidget()
        
        # Área de scroll
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll_area.setStyleSheet("QScrollArea { border: none; background-color: #f5f6fa; }")
        
        # Contenido
        scroll_content = QWidget()
        scroll_layout = QVBoxLayout()
        scroll_layout.setContentsMargins(40, 30, 40, 30)
        scroll_layout.setSpacing(25)
        
        # Banner
        self.add_banner(scroll_layout)
        
        # Cards
        self.add_module_cards(scroll_layout)
        
        # Resumen (solo para usuarios internos)
        if self.user_data.get("rol") != "cliente":
            self.add_summary_widget(scroll_layout)
        
        scroll_content.setLayout(scroll_layout)
        scroll_area.setWidget(scroll_content)
        
        home_layout = QVBoxLayout()
        home_layout.setContentsMargins(0, 0, 0, 0)
        home_layout.addWidget(scroll_area)
        home_widget.setLayout(home_layout)
        
        self.content_stack.addWidget(home_widget)
    
    def create_carga_masiva_page(self):
        """Crea la página de carga masiva"""
        from views.massiveLoadWindow import CargaMasivaContent
        
        carga_masiva_widget = CargaMasivaContent(self.user_data)
        carga_masiva_widget.back_requested.connect(self.show_home)
        self.content_stack.addWidget(carga_masiva_widget)

    def create_gestionar_calificaciones_page(self):
        """Crea la página de gestión de calificaciones"""
        from views.taxManagementWindow import GestionCalificacionesContent
        
        gestion_cal_widget = GestionCalificacionesContent(self.user_data)
        gestion_cal_widget.back_requested.connect(self.show_home)
        self.content_stack.addWidget(gestion_cal_widget)

    def create_gestionar_subsidios_page(self):
        """Crea la página de gestión de subsidios y la añade al stack"""
        try:
            from views.subsidiesWindow import SubsidiosWindow
        except Exception as e:
            from utils.logger import app_logger
            app_logger.error(f"No se pudo importar SubsidiosWindow: {e}")
            return

        try:
            subsidios_widget = SubsidiosWindow(self.user_data)
            subsidios_widget.back_requested.connect(self.show_home)
            self.content_stack.addWidget(subsidios_widget)
        except Exception as e:
            from utils.logger import app_logger
            app_logger.error(f"Error creando instancia de SubsidiosWindow: {e}")

    def create_reportes_page(self):
        """Crea la página de generación de reportes - NUEVO"""
        try:
            from views.reportsWindow import GenerarReportesContent
            
            reportes_widget = GenerarReportesContent(self.user_data)
            reportes_widget.back_requested.connect(self.show_home)
            self.content_stack.addWidget(reportes_widget)
        except Exception as e:
            from utils.logger import app_logger
            app_logger.error(f"Error creando página de reportes: {e}")

    def create_consultar_page(self):

        try:
            from views.queryWindow import ConsultarDatosContent
            
            consultar_widget = ConsultarDatosContent(self.user_data)
            consultar_widget.back_requested.connect(self.show_home)
            self.content_stack.addWidget(consultar_widget)
        except Exception as e:
            from utils.logger import app_logger
            app_logger.error(f"Error creando página de consulta: {e}")

    def create_usuarios_page(self):

        try:
            from views.userManagementWindow import GestionUsuariosContent
            
            usuarios_widget = GestionUsuariosContent(self.user_data)
            usuarios_widget.back_requested.connect(self.show_home)
            self.content_stack.addWidget(usuarios_widget)
        except Exception as e:
            from utils.logger import app_logger
            app_logger.error(f"Error creando página de usuarios: {e}")

    def show_usuarios(self):
        """Muestra la página de gestión de usuarios"""
        if not self.check_connection_before_operation("Gestión de Usuarios"):
            return

        self.content_stack.setCurrentIndex(6)

    def show_home(self):
        """Muestra la página de inicio"""
        self.content_stack.setCurrentIndex(0)
    
    def show_carga_masiva(self):
        """Muestra la página de carga masiva"""

        if not self.check_connection_before_operation("Carga Masiva"):
            return

        self.content_stack.setCurrentIndex(1)

    def show_gestionar_calificaciones(self):
        """Muestra la página de gestión de calificaciones"""

        if not self.check_connection_before_operation("Gestionar Calificaciones"):
            return

        self.content_stack.setCurrentIndex(2)

    def show_gestionar_subsidios(self):
        """Muestra la página de gestión de subsidios"""
        if not self.check_connection_before_operation("Gestionar Subsidios"):
            return

        # Intentar seleccionar la página correspondiente en el stack.
        # Asumimos que fue añadida después de las otras páginas; si no, la buscamos por tipo.
        # Preferimos localizar por nombre de clase agregado en create_gestionar_subsidios_page.
        for idx in range(self.content_stack.count()):
            w = self.content_stack.widget(idx)
            # buscar por nombre de clase para mayor robustez
            if w.__class__.__name__ in ("GestionSubsidiosContent", "SubsidiosWindow", "SubsidiosWidget"):
                self.content_stack.setCurrentIndex(idx)
                return

        # Fallback: si no existe, intentar recargar la página (recrear) y seleccionar la última añadida
        try:
            self.create_gestionar_subsidios_page()
            # seleccionar última
            self.content_stack.setCurrentIndex(self.content_stack.count() - 1)
        except Exception:
            self.show_home()

    def show_reportes(self):
        """Muestra la página de generación de reportes - NUEVO"""
        if not self.check_connection_before_operation("Generar Reportes"):
            return

        self.content_stack.setCurrentIndex(4)

    def show_consultar(self):
        """Muestra la página de consulta de datos - NUEVO"""
        if not self.check_connection_before_operation("Consultar Datos"):
            return

        self.content_stack.setCurrentIndex(5)
    
    def add_header(self, layout):
        """Agrega el header con logo y menú de usuario - SIN LÍNEAS"""
        header = QFrame()
        header.setFixedHeight(70)
        header.setStyleSheet("QFrame { background-color: white; }")
        
        header_layout = QHBoxLayout()
        header_layout.setContentsMargins(30, 10, 30, 10)
        
        # Logo + Título
        logo_layout = QHBoxLayout()
        logo_layout.setSpacing(10)
        
        # Logo
        logo_label = QLabel()
        logo_path = os.path.join(Settings.IMAGES_DIR, "logo.png")
        if os.path.exists(logo_path):
            pixmap = QPixmap(logo_path)
            scaled_pixmap = pixmap.scaled(40, 40, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            logo_label.setPixmap(scaled_pixmap)
        else:
            logo_label.setText("⚡")
            logo_label.setFont(QFont("Arial", 24))
            logo_label.setStyleSheet("color: #E94E1B;")
        
        logo_layout.addWidget(logo_label)
        
        # Título
        title_label = QLabel("TribunSys")
        title_label.setFont(QFont("Arial", 16, QFont.Bold))
        title_label.setStyleSheet("color: #2c3e50;")
        logo_layout.addWidget(title_label)
        
        header_layout.addLayout(logo_layout)
        header_layout.addStretch()

        # Indicador de conexión
        self.connection_indicator = ConnectionIndicator()
        header_layout.addWidget(self.connection_indicator)
        
        # Separador visual
        separator = QLabel("|")
        separator.setStyleSheet("color: #dee2e6; padding: 0 10px;")
        header_layout.addWidget(separator)
        
        # Info del usuario + menú desplegable (SIN ROL)
        user_layout = QHBoxLayout()
        user_layout.setSpacing(15)
        
        # Nombre del usuario solamente
        user_info = QLabel(f"Bienvenido, {self.user_data.get('nombre', '')} {self.user_data.get('apellido_P', '')}")
        user_info.setFont(QFont("Arial", 10))
        user_info.setStyleSheet("color: #2c3e50;")
        user_layout.addWidget(user_info)
        
        # Botón de menú de usuario
        self.user_menu_button = QPushButton("👤 ▼")
        self.user_menu_button.setFont(QFont("Arial", 10))
        self.user_menu_button.setCursor(QCursor(Qt.PointingHandCursor))
        self.user_menu_button.setFixedSize(70, 35)
        self.user_menu_button.setStyleSheet("""
            QPushButton {
                background-color: #f8f9fa;
                border: 1px solid #dee2e6;
                border-radius: 5px;
                padding: 5px;
            }
            QPushButton:hover {
                background-color: #e9ecef;
            }
        """)
        self.user_menu_button.clicked.connect(self.show_user_menu)
        user_layout.addWidget(self.user_menu_button)
        
        header_layout.addLayout(user_layout)
        header.setLayout(header_layout)
        layout.addWidget(header)

        self.connection_status_bar = ConnectionStatusBar()
        layout.addWidget(self.connection_status_bar)
    
    def show_user_menu(self):
        """Muestra el menú desplegable del usuario"""
        menu = QMenu(self)
        menu.setStyleSheet("""
            QMenu {
                background-color: white;
                border: 1px solid #dee2e6;
                border-radius: 5px;
                padding: 5px;
            }
            QMenu::item {
                padding: 8px 30px 8px 10px;
                border-radius: 3px;
            }
            QMenu::item:selected {
                background-color: #fef5f1;
                color: #E94E1B;
            }
        """)
        
        # Mi perfil
        profile_action = QAction("👤 Mi Perfil", self)
        profile_action.triggered.connect(self.open_profile)
        menu.addAction(profile_action)
        
        # Cambiar contraseña
        password_action = QAction("🔑 Cambiar Contraseña", self)
        password_action.triggered.connect(self.change_password)
        menu.addAction(password_action)
        
        menu.addSeparator()
        
        # Tema (Dark Mode)
        theme_text = "🌙 Modo Oscuro" if self.current_theme == "light" else "☀️ Modo Claro"
        theme_action = QAction(theme_text, self)
        theme_action.triggered.connect(self.toggle_theme)
        menu.addAction(theme_action)
        
        menu.addSeparator()
        
        # Cerrar sesión
        logout_action = QAction("🚪 Cerrar Sesión", self)
        logout_action.triggered.connect(self.logout)
        menu.addAction(logout_action)
        
        # Mostrar el menú debajo del botón
        menu.exec_(self.user_menu_button.mapToGlobal(self.user_menu_button.rect().bottomLeft()))
    
    def add_banner(self, layout):
        """Agrega el banner de bienvenida - CARGA banner.png"""
        banner = QFrame()
        banner.setFixedHeight(180)
        banner.setStyleSheet("""
            QFrame {
                border-radius: 10px;
                background-color: white;
            }
        """)
        
        banner_layout = QVBoxLayout()
        banner_layout.setContentsMargins(0, 0, 0, 0)
        
        # Intentar cargar imagen del banner
        banner_path = os.path.join(Settings.IMAGES_DIR, "banner.png")
        
        if os.path.exists(banner_path):
            # Cargar imagen personalizada
            banner_label = QLabel()
            banner_label.setAlignment(Qt.AlignCenter)
            pixmap = QPixmap(banner_path)
            scaled_pixmap = pixmap.scaledToHeight(180, Qt.SmoothTransformation)
            banner_label.setPixmap(scaled_pixmap)
            banner_label.setScaledContents(False)
            banner_layout.addWidget(banner_label)
        else:
            # Fallback: gradiente con texto
            banner.setStyleSheet("""
                QFrame {
                    background: qlineargradient(
                        x1:0, y1:0, x2:1, y2:0,
                        stop:0 #E94E1B,
                        stop:1 #ff7043
                    );
                    border-radius: 10px;
                }
            """)
            banner_layout.setContentsMargins(30, 20, 30, 20)
            
            title = QLabel("INTEGRACIÓN DE LAS BOLSAS DE")
            title.setFont(QFont("Arial", 16, QFont.Bold))
            title.setStyleSheet("color: white;")
            banner_layout.addWidget(title)
            
            subtitle = QLabel("CHILE, COLOMBIA Y PERÚ")
            subtitle.setFont(QFont("Arial", 20, QFont.Bold))
            subtitle.setStyleSheet("color: white;")
            banner_layout.addWidget(subtitle)
            
            description = QLabel("NUAM se destaca por un crecimiento sostenido a través de un mercado más amplio, líquido y eficiente")
            description.setFont(QFont("Arial", 9))
            description.setStyleSheet("color: rgba(255,255,255,0.9);")
            description.setWordWrap(True)
            banner_layout.addWidget(description)
            
            banner_layout.addStretch()
        
        banner.setLayout(banner_layout)
        layout.addWidget(banner)
    
    def add_module_cards(self, layout):
        """Agrega las tarjetas de módulos - SE ADAPTAN AL TAMAÑO"""
        cards_label = QLabel("Módulos del Sistema")
        cards_label.setFont(QFont("Arial", 14, QFont.Bold))
        cards_label.setStyleSheet("color: #2c3e50; margin-top: 5px; margin-bottom: 15px;")
        layout.addWidget(cards_label)
        
        # Grid de cards adaptable
        grid = QGridLayout()
        grid.setSpacing(25)
        
        modules = self.get_modules_by_role()
        
        row, col = 0, 0
        max_cols = 3
        
        for module in modules:
            card = CardButton(
                module_id=module["id"],
                title=module["title"],
                icon=module["icon"],
                description=module.get("description", "")
            )
            card.clicked.connect(self.on_module_selected)
            grid.addWidget(card, row, col)
            
            col += 1
            if col >= max_cols:
                col = 0
                row += 1
        
        layout.addLayout(grid)
    
    def get_modules_by_role(self):
        """
        Retorna los módulos disponibles según el rol
        
        Usa ModulosConfig para asegurar consistencia entre:
        - Cards del home
        - Sidebar
        - Permisos de navegación
        """
        from config.roles import ModulosConfig
        
        rol = self.user_data.get("rol", "cliente")
        
        # Obtener módulos permitidos para este rol
        modulos_permitidos = ModulosConfig.get_modulos_por_rol(rol)
        
        # Ya vienen con el formato correcto: id, title, icon, description
        return modulos_permitidos
    
    def add_summary_widget(self, layout):
        """Agrega el widget de resumen"""
        summary = QFrame()
        summary.setStyleSheet("""
            QFrame {
                background-color: white;
                border: 1px solid #dee2e6;
                border-radius: 10px;
                padding: 20px;
            }
        """)
        
        summary_layout = QVBoxLayout()
        
        title = QLabel("📈 Resumen Rápido")
        title.setFont(QFont("Arial", 13, QFont.Bold))
        title.setStyleSheet("color: #2c3e50;")
        summary_layout.addWidget(title)
        
        # Datos de ejemplo (luego serán dinámicos desde Firebase)
        stats = [
            "• Última actualización: Hoy",
            "• Calificaciones activas: Cargando...",
            "• Subsidios aplicados este mes: Cargando..."
        ]
        
        for stat in stats:
            stat_label = QLabel(stat)
            stat_label.setFont(QFont("Arial", 10))
            stat_label.setStyleSheet("color: #7f8c8d;")
            summary_layout.addWidget(stat_label)
        
        summary.setLayout(summary_layout)
        layout.addWidget(summary, 1)
    
    def add_footer(self, layout):
        """Agrega el footer - TEXTO CENTRADO"""
        footer = QFrame()
        footer.setFixedHeight(45)
        footer.setStyleSheet("""
            QFrame {
                background-color: white;
                border-top: 1px solid #dee2e6;
            }
        """)
        
        footer_layout = QHBoxLayout()
        footer_layout.setContentsMargins(0, 0, 0, 0)
        
        # Texto completamente centrado
        footer_text = QLabel(f"© 2025 Nuam  |  v{Settings.APP_VERSION}")
        footer_text.setFont(QFont("Arial", 9))
        footer_text.setStyleSheet("color: #95a5a6;")
        footer_text.setAlignment(Qt.AlignCenter)
        
        footer_layout.addWidget(footer_text)
        footer.setLayout(footer_layout)
        layout.addWidget(footer)
    
    def on_module_selected(self, module_id: str):
        """Maneja la selección de un módulo - CORREGIDO"""
        print(f"Módulo seleccionado: {module_id}")
        
        if module_id == "carga_masiva":
            self.show_carga_masiva()
        elif module_id == "calificaciones":
            self.show_gestionar_calificaciones()
        elif module_id == "subsidios":
            self.show_gestionar_subsidios()
        elif module_id == "reportes":
            self.show_reportes()
        elif module_id == "consultar":
            self.show_consultar()
        elif module_id == "usuarios":
            self.show_usuarios()
        else:
            # Si el módulo no está implementado, volver al home
            print(f"Módulo '{module_id}' no implementado aún")
            self.show_home()
    
    def open_profile(self):
        """Abre el perfil del usuario"""
        print("Abrir perfil")
        # TODO: Implementar
    
    def change_password(self):
        """Abre el diálogo de cambio de contraseña"""
        print("Cambiar contraseña")
        # TODO: Implementar
    
    def toggle_theme(self):
        """Cambia entre tema claro y oscuro"""
        self.current_theme = "dark" if self.current_theme == "light" else "light"
        self.apply_theme()
        self.theme_changed.emit(self.current_theme)
        print(f"Tema cambiado a: {self.current_theme}")
    
    def apply_theme(self):
        """Aplica el tema actual"""
        # TODO: Implementar en FASE 2
        pass
    
    def logout(self):
        """Cierra la sesión"""
        self.logout_requested.emit()
        self.close()

    def closeEvent(self, event):
        """
        Maneja el cierre de la ventana
        Detiene el monitoreo de conexión
        """
        # Detener monitoreo de conexión
        connection_manager.stop_monitoring()
        
        # Cerrar normalmente
        super().closeEvent(event)

    def check_connection_before_operation(self, operation_name: str = "Esta operación") -> bool:
        """
        Verifica si hay conexión antes de realizar una operación
        Muestra mensaje al usuario si no hay conexión
        
        Args:
            operation_name (str): Nombre de la operación para el mensaje
            
        Returns:
            bool: True si hay conexión, False si no
        """
        if not connection_manager.is_online():
            from PyQt5.QtWidgets import QMessageBox
            QMessageBox.warning(
                self,
                "Sin conexión",
                f"{operation_name} requiere conexión a internet.\n\n"
                "Por favor, verifica tu conexión y vuelve a intentarlo."
            )
            return False
        
        return True