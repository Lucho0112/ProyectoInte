import os
import firebase_admin
from firebase_admin import credentials, firestore, auth
import pyrebase
from dotenv import load_dotenv
from utils.helpers import resource_path



# Cargar variables de entorno
env_path = resource_path(".env")
load_dotenv(env_path)


class FirebaseConfig:
    """Clase para gestionar la configuración de Firebase"""
    
    _instance = None
    _initialized = False
    
    def __new__(cls):
        """Patrón Singleton para asegurar una única instancia"""
        if cls._instance is None:
            cls._instance = super(FirebaseConfig, cls).__new__(cls)
        return cls._instance
    
    def __init__(self):
        """Inicializa la configuración de Firebase"""
        if not self._initialized:
            self.firebase_config = {
                "apiKey": os.getenv("FIREBASE_API_KEY"),
                "authDomain": os.getenv("FIREBASE_AUTH_DOMAIN"),
                "projectId": os.getenv("FIREBASE_PROJECT_ID"),
                "storageBucket": os.getenv("FIREBASE_STORAGE_BUCKET"),
                "messagingSenderId": os.getenv("FIREBASE_MESSAGING_SENDER_ID"),
                "appId": os.getenv("FIREBASE_APP_ID"),
                "databaseURL": os.getenv("FIREBASE_DATABASE_URL")
            }

            relative_path = os.getenv("FIREBASE_CREDENTIALS_PATH")
            self.credentials_path = resource_path(relative_path)
            self._initialize_firebase()
            FirebaseConfig._initialized = True
    
    def _initialize_firebase(self):
        """Inicializa Firebase Admin SDK y Pyrebase"""
        try:
            # Inicializar Firebase Admin SDK (para Firestore)
            if not firebase_admin._apps:
                cred = credentials.Certificate(self.credentials_path)
                firebase_admin.initialize_app(cred)
                print("✅ Firebase Admin SDK inicializado correctamente")
            
            # Inicializar Pyrebase (para Authentication)
            self.firebase = pyrebase.initialize_app(self.firebase_config)
            self.auth_client = self.firebase.auth()
            
            # Obtener cliente de Firestore
            self.db = firestore.client()
            
            print("✅ Firebase Pyrebase inicializado correctamente")
            
        except Exception as e:
            print(f"❌ Error al inicializar Firebase: {str(e)}")
            raise
    
    def get_firestore_client(self):
        """Retorna el cliente de Firestore"""
        return self.db
    
    def get_auth_client(self):
        """Retorna el cliente de autenticación"""
        return self.auth_client
    
    def get_admin_auth(self):
        """Retorna el módulo de autenticación de Admin SDK"""
        return auth


# Instancia global de configuración
firebase_config = FirebaseConfig()