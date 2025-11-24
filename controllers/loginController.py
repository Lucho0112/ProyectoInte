from PyQt5.QtCore import QObject, pyqtSlot
from views.loginWindow import LoginWindow
from services.authService import AuthService
from utils.logger import app_logger, log_audit
from utils.validators import validate_rut


class LoginController(QObject):
    """Controlador para gestionar el proceso de login"""
    
    def __init__(self):
        super().__init__()
        self.auth_service = AuthService()
        self.login_window = None
        self.current_user = None
    
    def show_login(self):
        """Muestra la ventana de login"""
        self.login_window = LoginWindow()
        
        # Sobrescribir el método attempt_login de la ventana
        self.login_window.attempt_login = self.handle_login
        
        # Conectar señal de login exitoso
        self.login_window.login_successful.connect(self.on_login_success)
        
        self.login_window.show()
        app_logger.info("Ventana de login mostrada")
    
    def handle_login(self, rut, password):
        """
        Maneja el intento de login
        
        Args:
            rut (str): RUT del usuario
            password (str): Contraseña del usuario
        """
        try:
            # Validar RUT
            is_valid, error_message = validate_rut(rut)
            if not is_valid:
                self.login_window.show_error(error_message)
                return
            
            # Intentar autenticar
            app_logger.info(f"Intentando autenticar usuario con RUT: {rut}")
            
            result = self.auth_service.login(rut, password)
            
            if result["success"]:
                self.current_user = result["user"]
                
                # Registrar auditoría
                log_audit(
                    action="LOGIN_EXITOSO",
                    user_id=self.current_user.get("_id", "unknown"),
                    details={"rut": rut}
                )
                
                # Emitir señal de éxito
                self.login_window.login_successful.emit(self.current_user)
                
            else:
                # Mostrar error
                self.login_window.show_error(result["message"])
                
                # Registrar intento fallido
                log_audit(
                    action="LOGIN_FALLIDO",
                    user_id="N/A",
                    details={"rut": rut, "error": result["message"]}
                )
        
        except Exception as e:
            error_msg = f"Error inesperado durante el login: {str(e)}"
            app_logger.error(error_msg)
            self.login_window.show_error("Error al iniciar sesión. Por favor intente nuevamente.")
        
        finally:
            # Restablecer botón de login
            self.login_window.reset_login_button()
    
    @pyqtSlot(dict)
    def on_login_success(self, user_data):
        """
        Maneja el login exitoso
        
        Args:
            user_data (dict): Datos del usuario autenticado
        """
        app_logger.info(f"Login exitoso para usuario: {user_data.get('nombre', 'N/A')}")
        
        # Cerrar ventana de login
        self.login_window.close()
        
        # Abrir ventana principal
        self.open_main_window(user_data)
    
    def open_main_window(self, user_data):
        """
        Abre la ventana principal del sistema
        
        Args:
            user_data (dict): Datos del usuario autenticado
        """
        from views.mainWindow import MainWindow
        
        self.main_window = MainWindow(user_data)
        self.main_window.logout_requested.connect(self.on_logout_requested)
        self.main_window.show()
        
        app_logger.info("Ventana principal abierta")
    
    def on_logout_requested(self):
        """Maneja la solicitud de cerrar sesión"""
        if self.current_user:
            user_id = self.current_user.get("_id")
            self.auth_service.logout(user_id)
        
        # Cerrar ventana principal
        if hasattr(self, 'main_window'):
            self.main_window.close()
        
        # Volver a mostrar login
        self.show_login()
        
        app_logger.info("Sesión cerrada, volviendo a login")
    
    def get_current_user(self):
        """Retorna el usuario actual autenticado"""
        return self.current_user