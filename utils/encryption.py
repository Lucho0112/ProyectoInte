import bcrypt
from config.settings import Settings


class EncryptionUtils:
    """Clase para manejar encriptación de contraseñas"""
    
    @staticmethod
    def hash_password(password: str) -> str:
        """
        Hashea una contraseña usando bcrypt con salt
        
        Args:
            password (str): Contraseña en texto plano
            
        Returns:
            str: Hash de la contraseña
        """
        if not password:
            raise ValueError("La contraseña no puede estar vacía")
        
        # Generar salt y hashear
        salt = bcrypt.gensalt(rounds=Settings.BCRYPT_ROUNDS)
        hashed = bcrypt.hashpw(password.encode('utf-8'), salt)
        
        return hashed.decode('utf-8')
    
    @staticmethod
    def verify_password(password: str, hashed_password: str) -> bool:
        """
        Verifica si una contraseña coincide con su hash
        
        Args:
            password (str): Contraseña en texto plano
            hashed_password (str): Hash almacenado
            
        Returns:
            bool: True si la contraseña es correcta, False en caso contrario
        """
        if not password or not hashed_password:
            return False
        
        try:
            return bcrypt.checkpw(
                password.encode('utf-8'),
                hashed_password.encode('utf-8')
            )
        except Exception as e:
            print(f"Error al verificar contraseña: {str(e)}")
            return False
    
    @staticmethod
    def is_strong_password(password: str) -> tuple[bool, str]:
        """
        Valida si una contraseña cumple con requisitos de seguridad
        
        Args:
            password (str): Contraseña a validar
            
        Returns:
            tuple[bool, str]: (es_válida, mensaje_error)
        """
        if len(password) < 8:
            return False, "La contraseña debe tener al menos 8 caracteres"
        
        if not any(char.isupper() for char in password):
            return False, "La contraseña debe contener al menos una mayúscula"
        
        if not any(char.islower() for char in password):
            return False, "La contraseña debe contener al menos una minúscula"
        
        if not any(char.isdigit() for char in password):
            return False, "La contraseña debe contener al menos un número"
        
        special_chars = "!@#$%^&*()_+-=[]{}|;:,.<>?"
        if not any(char in special_chars for char in password):
            return False, "La contraseña debe contener al menos un carácter especial"
        
        return True, "Contraseña válida"


# Funciones auxiliares para uso directo
def hash_password(password: str) -> str:
    """Wrapper para hashear contraseña"""
    return EncryptionUtils.hash_password(password)


def verify_password(password: str, hashed_password: str) -> bool:
    """Wrapper para verificar contraseña"""
    return EncryptionUtils.verify_password(password, hashed_password)


def validate_password_strength(password: str) -> tuple[bool, str]:
    """Wrapper para validar fortaleza de contraseña"""
    return EncryptionUtils.is_strong_password(password)