import os
from dotenv import load_dotenv


load_dotenv()


class Settings:
    """Configuraciones generales de la aplicación"""
    
    # Información de la aplicación
    APP_NAME = os.getenv("APP_NAME", "TribunSys")
    APP_VERSION = os.getenv("APP_VERSION", "1.0.0")
    ENVIRONMENT = os.getenv("ENVIRONMENT", "development")
    
    # Seguridad
    SECRET_KEY = os.getenv("SECRET_KEY", "default_secret_key_change_this")
    BCRYPT_ROUNDS = int(os.getenv("BCRYPT_ROUNDS", "12"))
    
    # Logging
    LOG_LEVEL = os.getenv("LOG_LEVEL", "INFO")
    LOG_FILE_PATH = os.getenv("LOG_FILE_PATH", "./logs/tribunsys.log")
    LOG_MAX_SIZE = 10 * 1024 * 1024  # 10 MB
    LOG_BACKUP_COUNT = 5
    
    # Colecciones de Firestore
    COLLECTION_USUARIOS = "usuarios"
    COLLECTION_DATOS_TRIBUTARIOS = "datosTributarios"
    COLLECTION_SUBSIDIOS = "subsidios"
    COLLECTION_REPORTES = "reportes"
    COLLECTION_AUDITORIA = "auditoria"
    
    # Validaciones
    MAX_LOGIN_ATTEMPTS = 3
    SESSION_TIMEOUT = 3600  # 1 hora en segundos
    
    # Carga masiva
    MAX_BATCH_SIZE = 1000
    MAX_UPLOAD_TIME = 10  # segundos
    ALLOWED_FILE_EXTENSIONS = ['.csv', '.xlsx', '.xls']
    
    # Validación de factores (M-01, A-01)
    MIN_FACTOR_VALUE = 0
    MAX_FACTOR_VALUE = 1
    FACTOR_SUM_RANGE = (8, 19)  # Factores del 8 al 19
    MAX_FACTOR_SUM = 1.0
    
    # Rendimiento
    MAX_RESPONSE_TIME = 5  # segundos
    CACHE_ENABLED = True
    CACHE_TTL = 300  # 5 minutos
    
    # UI Settings
    WINDOW_MIN_WIDTH = 1024
    WINDOW_MIN_HEIGHT = 768
    THEME = "light"
    
    # Paths
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    RESOURCES_DIR = os.path.join(BASE_DIR, "resources")
    TEMPLATES_DIR = os.path.join(RESOURCES_DIR, "templates")
    IMAGES_DIR = os.path.join(RESOURCES_DIR, "images")
    LOGS_DIR = os.path.join(BASE_DIR, "logs")
    
    @classmethod
    def ensure_directories(cls):
        """Asegura que los directorios necesarios existan"""
        directories = [
            cls.LOGS_DIR,
            cls.RESOURCES_DIR,
            cls.TEMPLATES_DIR,
            cls.IMAGES_DIR
        ]
        for directory in directories:
            os.makedirs(directory, exist_ok=True)
    
    @classmethod
    def is_production(cls):
        """Verifica si está en entorno de producción"""
        return cls.ENVIRONMENT.lower() == "production"
    
    @classmethod
    def is_development(cls):
        """Verifica si está en entorno de desarrollo"""
        return cls.ENVIRONMENT.lower() == "development"


# Crear directorios al importar
Settings.ensure_directories()