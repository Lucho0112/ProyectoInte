import sys
from PyQt5.QtWidgets import QApplication, QMessageBox
from PyQt5.QtCore import Qt

# Importar configuraciones
from config.settings import Settings
from utils.logger import LoggerSetup, app_logger

# Importar controlador de login
from controllers.loginController import LoginController


class TribunSysApp:
    """Clase principal de la aplicación"""
    
    def __init__(self):
        """Inicializa la aplicación"""
        self.app = None
        self.login_controller = None
        
    def setup(self):
        """Configura la aplicación antes de iniciar"""
        try:
            # Configurar logging
            LoggerSetup.setup()
            app_logger.info("="*60)
            app_logger.info(f"Iniciando {Settings.APP_NAME} v{Settings.APP_VERSION}")
            app_logger.info(f"Entorno: {Settings.ENVIRONMENT}")
            app_logger.info("="*60)
            
            # Verificar que los directorios necesarios existan
            Settings.ensure_directories()
            app_logger.info("Directorios de la aplicación verificados")
            
            # Inicializar PyQt Application
            self.app = QApplication(sys.argv)
            self.app.setApplicationName(Settings.APP_NAME)
            self.app.setApplicationVersion(Settings.APP_VERSION)
            
            # Configurar atributos de alta DPI
            QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
            QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
            
            app_logger.info("Aplicación PyQt5 inicializada")
            
            return True
            
        except Exception as e:
            error_msg = f"Error crítico al inicializar la aplicación: {str(e)}"
            app_logger.critical(error_msg)
            self.show_critical_error(error_msg)
            return False
    
    def run(self):
        """Ejecuta la aplicación"""
        try:
            # Configurar aplicación
            if not self.setup():
                return 1
            
            # Verificar conexión con Firebase
            if not self.check_firebase_connection():
                return 1
            
            # Inicializar y mostrar ventana de login
            self.login_controller = LoginController()
            self.login_controller.show_login()
            
            app_logger.info("Aplicación lista. Esperando interacción del usuario...")
            
            # Iniciar loop de eventos
            exit_code = self.app.exec_()
            
            app_logger.info(f"Aplicación cerrada con código: {exit_code}")
            return exit_code
            
        except Exception as e:
            error_msg = f"Error durante la ejecución: {str(e)}"
            app_logger.critical(error_msg)
            self.show_critical_error(error_msg)
            return 1
    
    def check_firebase_connection(self):
        """
        Verifica la conexión con Firebase
        
        Returns:
            bool: True si la conexión es exitosa
        """
        try:
            from config.firebaseConfig import firebase_config
            
            # Intentar obtener cliente de Firestore
            db = firebase_config.get_firestore_client()
            
            if db:
                app_logger.info("Conexión con Firebase establecida correctamente")
                return True
            else:
                raise Exception("No se pudo obtener el cliente de Firestore")
                
        except Exception as e:
            error_msg = f"Error al conectar con Firebase: {str(e)}"
            app_logger.error(error_msg)
            
            self.show_critical_error(
                "No se pudo establecer conexión con Firebase.\n\n"
                "Verifique:\n"
                "1. Que el archivo .env existe y tiene las credenciales correctas\n"
                "2. Que el archivo serviceAccountKey.json está en la carpeta config/\n"
                "3. Que tiene conexión a Internet\n\n"
                f"Error técnico: {str(e)}"
            )
            return False
    
    def show_critical_error(self, message):
        """
        Muestra un error crítico al usuario
        
        Args:
            message (str): Mensaje de error
        """
        if self.app:
            QMessageBox.critical(
                None,
                f"{Settings.APP_NAME} - Error Crítico",
                message
            )
        else:
            print(f"ERROR CRÍTICO: {message}")


def main():
    """Función principal"""
    app = TribunSysApp()
    sys.exit(app.run())


if __name__ == "__main__":
    main()