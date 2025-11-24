from functools import wraps
from typing import Any, Callable
from utils.connectionManager import connection_manager
from utils.logger import app_logger


class FirebaseOperationError(Exception):
    """Excepción personalizada para operaciones de Firebase"""
    pass


class OfflineError(FirebaseOperationError):
    """Error cuando no hay conexión a internet"""
    pass


def requires_connection(func: Callable) -> Callable:
    """
    Decorador para funciones que requieren conexión a internet
    
    Usage:
        @requires_connection
        def mi_funcion_firebase(self, ...):
            # código que usa Firebase
    """
    @wraps(func)
    def wrapper(*args, **kwargs):
        if not connection_manager.is_online():
            error_msg = "No hay conexión a internet. Esta operación requiere conexión."
            app_logger.error(f"{func.__name__}: {error_msg}")
            
            # Retornar estructura estándar de error
            return {
                "success": False,
                "message": error_msg,
                "offline": True
            }
        
        try:
            return func(*args, **kwargs)
        
        except Exception as e:
            # Verificar si el error es por pérdida de conexión durante la operación
            if not connection_manager.is_online():
                error_msg = "Se perdió la conexión durante la operación"
                app_logger.error(f"{func.__name__}: {error_msg}")
                return {
                    "success": False,
                    "message": error_msg,
                    "offline": True
                }
            
            # Si no es problema de conexión, propagar el error original
            app_logger.error(f"{func.__name__}: {str(e)}")
            return {
                "success": False,
                "message": f"Error: {str(e)}",
                "offline": False
            }
    
    return wrapper


def safe_firebase_operation(func: Callable) -> Callable:
    """
    Decorador más suave que intenta la operación pero no falla si no hay conexión
    Útil para operaciones opcionales como logs o auditoría
    
    Usage:
        @safe_firebase_operation
        def log_to_firebase(self, ...):
            # Esta función no lanzará error si no hay conexión
    """
    @wraps(func)
    def wrapper(*args, **kwargs):
        if not connection_manager.is_online():
            app_logger.warning(f"{func.__name__}: Omitida (sin conexión)")
            return None
        
        try:
            return func(*args, **kwargs)
        except Exception as e:
            app_logger.warning(f"{func.__name__}: Error pero continuando: {str(e)}")
            return None
    
    return wrapper


class FirebaseServiceBase:
    """
    Clase base para servicios de Firebase con manejo offline
    """
    
    def __init__(self):
        self._offline_queue = []
    
    def is_online(self) -> bool:
        """Verifica si hay conexión"""
        return connection_manager.is_online()
    
    def queue_operation(self, operation_name: str, operation_data: dict):
        """
        Encola una operación para ejecutar cuando se restaure la conexión
        
        Args:
            operation_name (str): Nombre de la operación
            operation_data (dict): Datos necesarios para la operación
        """
        self._offline_queue.append({
            "operation": operation_name,
            "data": operation_data,
            "timestamp": app_logger.info(f"Operación encolada: {operation_name}")
        })
    
    def process_queue(self):
        """
        Procesa las operaciones encoladas cuando se restaure la conexión
        """
        if not self.is_online():
            return
        
        processed = 0
        failed = 0
        
        while self._offline_queue:
            operation = self._offline_queue.pop(0)
            
            try:
                # Aquí cada servicio debe implementar su lógica
                # de procesamiento según el tipo de operación
                self._process_queued_operation(operation)
                processed += 1
            except Exception as e:
                app_logger.error(f"Error procesando operación encolada: {str(e)}")
                failed += 1
        
        app_logger.info(f"Cola procesada: {processed} éxitos, {failed} fallos")
    
    def _process_queued_operation(self, operation: dict):
        """
        Método a sobrescribir en cada servicio para procesar operaciones encoladas
        
        Args:
            operation (dict): Operación a procesar
        """
        raise NotImplementedError("Cada servicio debe implementar este método")


def handle_firebase_error(error: Exception, operation: str = "Operación") -> dict:
    """
    Maneja errores de Firebase de forma estandarizada
    
    Args:
        error (Exception): Error capturado
        operation (str): Nombre de la operación
        
    Returns:
        dict: Respuesta estandarizada de error
    """
    error_str = str(error)
    
    # Detectar tipos comunes de errores
    if "permission" in error_str.lower() or "denied" in error_str.lower():
        message = f"{operation} falló: Permisos insuficientes"
    elif "not found" in error_str.lower():
        message = f"{operation} falló: Recurso no encontrado"
    elif "network" in error_str.lower() or "timeout" in error_str.lower():
        message = f"{operation} falló: Error de red o timeout"
    else:
        message = f"{operation} falló: {error_str}"
    
    app_logger.error(message)
    
    return {
        "success": False,
        "message": message,
        "offline": not connection_manager.is_online()
    }