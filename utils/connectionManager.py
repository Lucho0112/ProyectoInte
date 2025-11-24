import socket
import threading
import time
from typing import Callable, Optional
from PyQt5.QtCore import QObject, pyqtSignal
from utils.logger import app_logger


class ConnectionManager(QObject):
    """
    Gestor de conexión a internet con monitoreo continuo
    Emite señales cuando cambia el estado de conexión
    """
    
    # Señales
    connection_lost = pyqtSignal()
    connection_restored = pyqtSignal()
    connection_status_changed = pyqtSignal(bool)  # True = online, False = offline
    
    _instance = None
    
    def __new__(cls):
        """Singleton para un único gestor de conexión"""
        if cls._instance is None:
            cls._instance = super(ConnectionManager, cls).__new__(cls)
            cls._instance._initialized = False
        return cls._instance
    
    def __init__(self):
        if self._initialized:
            return
        
        super().__init__()
        self._initialized = True
        self._is_online = True
        self._monitoring = False
        self._monitor_thread = None
        self._check_interval = 5  # Segundos entre verificaciones
        self._consecutive_failures = 0
        self._max_failures = 2  # Fallos consecutivos antes de marcar offline
    
    def is_online(self) -> bool:
        """
        Verifica si hay conexión a internet
        
        Returns:
            bool: True si hay conexión, False si no
        """
        try:
            # Intenta conectar a Google DNS (rápido y confiable)
            socket.create_connection(("8.8.8.8", 53), timeout=3)
            return True
        except OSError:
            pass
        
        try:
            # Backup: intenta conectar a Cloudflare DNS
            socket.create_connection(("1.1.1.1", 53), timeout=3)
            return True
        except OSError:
            pass
        
        return False
    
    def check_firebase_connection(self) -> bool:
        """
        Verifica específicamente la conexión con Firebase
        
        Returns:
            bool: True si Firebase está accesible
        """
        try:
            from config.firebaseConfig import firebase_config
            
            # Intenta una operación mínima en Firestore
            db = firebase_config.get_firestore_client()
            
            # Operación ligera: verificar si se puede acceder a colecciones
            collections = db.collections()
            list(collections)  # Forzar la ejecución
            
            return True
        except Exception as e:
            app_logger.warning(f"Firebase no accesible: {str(e)}")
            return False
    
    def start_monitoring(self):
        """Inicia el monitoreo continuo de conexión"""
        if self._monitoring:
            return
        
        self._monitoring = True
        self._monitor_thread = threading.Thread(target=self._monitor_loop, daemon=True)
        self._monitor_thread.start()
        
        app_logger.info("Monitoreo de conexión iniciado")
    
    def stop_monitoring(self):
        """Detiene el monitoreo de conexión"""
        self._monitoring = False
        if self._monitor_thread:
            self._monitor_thread.join(timeout=2)
        
        app_logger.info("Monitoreo de conexión detenido")
    
    def _monitor_loop(self):
        """Loop de monitoreo que se ejecuta en un thread separado"""
        while self._monitoring:
            current_status = self.is_online()
            
            if current_status:
                # Conexión OK
                self._consecutive_failures = 0
                
                if not self._is_online:
                    # Conexión restaurada
                    self._is_online = True
                    app_logger.info("✅ Conexión a internet restaurada")
                    self.connection_restored.emit()
                    self.connection_status_changed.emit(True)
            else:
                # Sin conexión
                self._consecutive_failures += 1
                
                if self._is_online and self._consecutive_failures >= self._max_failures:
                    # Conexión perdida
                    self._is_online = False
                    app_logger.warning("❌ Conexión a internet perdida")
                    self.connection_lost.emit()
                    self.connection_status_changed.emit(False)
            
            # Esperar antes de siguiente verificación
            time.sleep(self._check_interval)
    
    def get_status(self) -> dict:
        """
        Retorna el estado actual de conexión
        
        Returns:
            dict: {
                "online": bool,
                "firebase_accessible": bool,
                "consecutive_failures": int
            }
        """
        return {
            "online": self._is_online,
            "firebase_accessible": self.check_firebase_connection() if self._is_online else False,
            "consecutive_failures": self._consecutive_failures
        }


# Instancia global
connection_manager = ConnectionManager()