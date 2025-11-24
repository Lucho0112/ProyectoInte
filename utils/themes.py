class Theme:
    """Clase base para temas"""
    
    # Colores principales
    PRIMARY = "#E94E1B"
    PRIMARY_HOVER = "#d63d0f"
    PRIMARY_LIGHT = "#fef5f1"
    
    # Colores de texto
    TEXT_PRIMARY = "#2c3e50"
    TEXT_SECONDARY = "#7f8c8d"
    TEXT_DISABLED = "#95a5a6"
    
    # Fondos
    BACKGROUND = "#f5f6fa"
    BACKGROUND_CARD = "#ffffff"
    BACKGROUND_SIDEBAR = "#f8f9fa"
    
    # Bordes
    BORDER = "#dee2e6"
    BORDER_LIGHT = "#ecf0f1"
    
    # Estados
    SUCCESS = "#27ae60"
    WARNING = "#f39c12"
    DANGER = "#e74c3c"
    INFO = "#3498db"


class LightTheme(Theme):
    """Tema claro (por defecto)"""
    
    NAME = "light"
    
    # Colores de texto
    TEXT_PRIMARY = "#2c3e50"
    TEXT_SECONDARY = "#7f8c8d"
    TEXT_DISABLED = "#95a5a6"
    
    # Fondos
    BACKGROUND = "#f5f6fa"
    BACKGROUND_CARD = "#ffffff"
    BACKGROUND_SIDEBAR = "#f8f9fa"
    BACKGROUND_HEADER = "#ffffff"
    BACKGROUND_FOOTER = "#ffffff"
    
    # Bordes
    BORDER = "#dee2e6"
    BORDER_LIGHT = "#ecf0f1"
    
    # Sidebar
    SIDEBAR_BUTTON_HOVER = "#fef5f1"
    SIDEBAR_BUTTON_PRESSED = "#fde8df"


class DarkTheme(Theme):
    """Tema oscuro"""
    
    NAME = "dark"
    
    # Colores de texto (invertidos)
    TEXT_PRIMARY = "#ecf0f1"
    TEXT_SECONDARY = "#bdc3c7"
    TEXT_DISABLED = "#7f8c8d"
    
    # Fondos oscuros
    BACKGROUND = "#1a1d23"
    BACKGROUND_CARD = "#2c3e50"
    BACKGROUND_SIDEBAR = "#252932"
    BACKGROUND_HEADER = "#2c3e50"
    BACKGROUND_FOOTER = "#2c3e50"
    
    # Bordes oscuros
    BORDER = "#34495e"
    BORDER_LIGHT = "#3d4b5e"
    
    # Sidebar oscuro
    SIDEBAR_BUTTON_HOVER = "#34495e"
    SIDEBAR_BUTTON_PRESSED = "#2c3e50"


class ThemeManager:
    """Gestor de temas de la aplicación"""
    
    _current_theme = LightTheme
    
    @classmethod
    def get_current_theme(cls):
        """Retorna el tema actual"""
        return cls._current_theme
    
    @classmethod
    def set_theme(cls, theme_name: str):
        """
        Establece el tema actual
        
        Args:
            theme_name (str): 'light' o 'dark'
        """
        if theme_name == "dark":
            cls._current_theme = DarkTheme
        else:
            cls._current_theme = LightTheme
    
    @classmethod
    def is_dark_mode(cls):
        """Retorna True si el tema actual es oscuro"""
        return cls._current_theme.NAME == "dark"
    
    @classmethod
    def get_main_window_style(cls):
        """Retorna el stylesheet para la ventana principal"""
        theme = cls._current_theme
        
        return f"""
            QMainWindow {{
                background-color: {theme.BACKGROUND};
            }}
            
            QWidget {{
                color: {theme.TEXT_PRIMARY};
            }}
            
            QScrollArea {{
                border: none;
                background-color: {theme.BACKGROUND};
            }}
        """
    
    @classmethod
    def get_header_style(cls):
        """Retorna el stylesheet para el header"""
        theme = cls._current_theme
        
        return f"""
            QFrame {{
                background-color: {theme.BACKGROUND_HEADER};
                border-bottom: 1px solid {theme.BORDER};
            }}
            
            QLabel {{
                color: {theme.TEXT_PRIMARY};
            }}
            
            QPushButton {{
                background-color: {theme.BACKGROUND_SIDEBAR};
                border: 1px solid {theme.BORDER};
                border-radius: 5px;
                padding: 5px;
                color: {theme.TEXT_PRIMARY};
            }}
            
            QPushButton:hover {{
                background-color: {theme.BORDER_LIGHT};
            }}
        """
    
    @classmethod
    def get_sidebar_style(cls):
        """Retorna el stylesheet para el sidebar"""
        theme = cls._current_theme
        
        return f"""
            QWidget {{
                background-color: {theme.BACKGROUND_SIDEBAR};
                border-right: 1px solid {theme.BORDER};
            }}
            
            QPushButton {{
                background-color: transparent;
                border: none;
                text-align: left;
                padding: 12px 15px;
                color: {theme.TEXT_PRIMARY};
            }}
            
            QPushButton:hover {{
                background-color: {theme.SIDEBAR_BUTTON_HOVER};
                border-left: 3px solid {Theme.PRIMARY};
            }}
            
            QPushButton:pressed {{
                background-color: {theme.SIDEBAR_BUTTON_PRESSED};
            }}
        """
    
    @classmethod
    def get_card_style(cls):
        """Retorna el stylesheet para las cards"""
        theme = cls._current_theme
        
        return f"""
            #cardButton {{
                background-color: {theme.BACKGROUND_CARD};
                border: 2px solid {theme.BORDER_LIGHT};
                border-radius: 10px;
                padding: 15px;
            }}
            
            #cardButton:hover {{
                border: 2px solid {Theme.PRIMARY};
                background-color: {Theme.PRIMARY_LIGHT if theme.NAME == 'light' else theme.SIDEBAR_BUTTON_HOVER};
            }}
            
            QLabel {{
                color: {theme.TEXT_PRIMARY};
            }}
        """
    
    @classmethod
    def get_banner_style(cls):
        """Retorna el stylesheet para el banner"""
        theme = cls._current_theme
        
        return f"""
            QFrame {{
                background: qlineargradient(
                    x1:0, y1:0, x2:1, y2:0,
                    stop:0 {Theme.PRIMARY},
                    stop:1 #ff7043
                );
                border-radius: 10px;
            }}
            
            QLabel {{
                color: white;
            }}
        """
    
    @classmethod
    def get_summary_style(cls):
        """Retorna el stylesheet para el widget de resumen"""
        theme = cls._current_theme
        
        return f"""
            QFrame {{
                background-color: {theme.BACKGROUND_CARD};
                border: 1px solid {theme.BORDER};
                border-radius: 10px;
                padding: 20px;
            }}
            
            QLabel {{
                color: {theme.TEXT_PRIMARY};
            }}
        """
    
    @classmethod
    def get_footer_style(cls):
        """Retorna el stylesheet para el footer"""
        theme = cls._current_theme
        
        return f"""
            QFrame {{
                background-color: {theme.BACKGROUND_FOOTER};
                border-top: 1px solid {theme.BORDER};
            }}
            
            QLabel {{
                color: {theme.TEXT_DISABLED};
            }}
        """
    
    @classmethod
    def get_menu_style(cls):
        """Retorna el stylesheet para menús desplegables"""
        theme = cls._current_theme
        
        return f"""
            QMenu {{
                background-color: {theme.BACKGROUND_CARD};
                border: 1px solid {theme.BORDER};
                border-radius: 5px;
                padding: 5px;
            }}
            
            QMenu::item {{
                padding: 8px 30px 8px 10px;
                border-radius: 3px;
                color: {theme.TEXT_PRIMARY};
            }}
            
            QMenu::item:selected {{
                background-color: {Theme.PRIMARY_LIGHT if theme.NAME == 'light' else theme.SIDEBAR_BUTTON_HOVER};
                color: {Theme.PRIMARY};
            }}
        """


# Instancia global del gestor de temas
theme_manager = ThemeManager()