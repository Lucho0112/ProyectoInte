import sys
import os
from loguru import logger
from config.settings import Settings


class LoggerSetup:
    """Configuración centralizada del sistema de logging"""
    
    _configured = False
    
    @classmethod
    def setup(cls):
        """Configura el sistema de logging"""
        if cls._configured:
            return
        
        # Remover configuración por defecto
        logger.remove()
        
        # Limpiar log anterior solo en desarrollo
        if Settings.is_development() and os.path.exists(Settings.LOG_FILE_PATH):
            try:
                os.remove(Settings.LOG_FILE_PATH)
                print("✓ Log anterior limpiado (modo desarrollo)")
            except Exception as e:
                print(f"⚠ No se pudo limpiar log anterior: {e}")
        
        # Configurar salida a consola con colores
        logger.add(
            sys.stdout,
            colorize=True,
            format="<green>{time:YYYY-MM-DD HH:mm:ss}</green> | "
                   "<level>{level: <8}</level> | "
                   "<cyan>{name}</cyan>:<cyan>{function}</cyan>:<cyan>{line}</cyan> | "
                   "<level>{message}</level>",
            level=Settings.LOG_LEVEL
        )
        
        # Configurar salida a archivo con rotación
        logger.add(
            Settings.LOG_FILE_PATH,
            rotation=Settings.LOG_MAX_SIZE,
            retention=Settings.LOG_BACKUP_COUNT,
            compression="zip",
            format="{time:YYYY-MM-DD HH:mm:ss} | {level: <8} | "
                   "{name}:{function}:{line} | {message}",
            level="DEBUG",  # Guardar todos los niveles en archivo
            encoding="utf-8"
        )
        
        # Archivo separado para errores críticos
        logger.add(
            Settings.LOGS_DIR + "/errors.log",
            rotation="1 week",
            retention="1 month",
            level="ERROR",
            format="{time:YYYY-MM-DD HH:mm:ss} | {level} | {name}:{function}:{line} | {message}",
            encoding="utf-8"
        )
        
        cls._configured = True
        logger.info("Sistema de logging configurado correctamente")
    
    @classmethod
    def get_logger(cls, name: str = None):
        """
        Obtiene una instancia del logger
        
        Args:
            name (str): Nombre del módulo que usa el logger
            
        Returns:
            Logger: Instancia configurada de loguru
        """
        if not cls._configured:
            cls.setup()
        
        if name:
            return logger.bind(module=name)
        return logger


# Instancia global del logger
app_logger = LoggerSetup.get_logger("TribunSys")


def log_info(message: str, **kwargs):
    """Registra mensaje informativo"""
    app_logger.info(message, **kwargs)


def log_error(message: str, **kwargs):
    """Registra mensaje de error"""
    app_logger.error(message, **kwargs)


def log_warning(message: str, **kwargs):
    """Registra mensaje de advertencia"""
    app_logger.warning(message, **kwargs)


def log_debug(message: str, **kwargs):
    """Registra mensaje de debug"""
    app_logger.debug(message, **kwargs)


def log_critical(message: str, **kwargs):
    """Registra mensaje crítico"""
    app_logger.critical(message, **kwargs)


def log_audit(action: str, user_id: str, details: dict):
    """
    Registra una acción de auditoría
    
    Args:
        action (str): Acción realizada
        user_id (str): ID del usuario que realizó la acción
        details (dict): Detalles adicionales de la acción
    """
    app_logger.info(
        f"AUDITORÍA | Acción: {action} | Usuario: {user_id} | Detalles: {details}"
    )