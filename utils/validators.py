import re
from typing import Tuple


def clean_rut(rut: str) -> str:
    """
    Limpia el RUT removiendo puntos y guiones
    
    Args:
        rut (str): RUT a limpiar
        
    Returns:
        str: RUT limpio
    """
    return rut.replace(".", "").replace("-", "").upper()


def format_rut(rut: str) -> str:
    """
    Formatea un RUT en el formato XX.XXX.XXX-X
    
    Args:
        rut (str): RUT sin formato
        
    Returns:
        str: RUT formateado
    """
    rut = clean_rut(rut)
    if len(rut) < 2:
        return rut
    
    dv = rut[-1]
    numero = rut[:-1]
    
    # Agregar puntos
    numero_formateado = ""
    for i, digit in enumerate(reversed(numero)):
        if i > 0 and i % 3 == 0:
            numero_formateado = "." + numero_formateado
        numero_formateado = digit + numero_formateado
    
    return f"{numero_formateado}-{dv}"


def calculate_rut_dv(rut_number: str) -> str:
    """
    Calcula el dígito verificador de un RUT
    
    Args:
        rut_number (str): Número del RUT sin DV
        
    Returns:
        str: Dígito verificador
    """
    reversed_digits = map(int, reversed(str(rut_number)))
    factors = [2, 3, 4, 5, 6, 7]
    
    s = sum(d * factors[i % 6] for i, d in enumerate(reversed_digits))
    remainder = s % 11
    dv = 11 - remainder
    
    if dv == 11:
        return "0"
    elif dv == 10:
        return "K"
    else:
        return str(dv)


def validate_rut(rut: str) -> Tuple[bool, str]:
    """
    Valida un RUT chileno
    
    Args:
        rut (str): RUT a validar
        
    Returns:
        Tuple[bool, str]: (es_válido, mensaje_error)
    """
    if not rut:
        return False, "El RUT no puede estar vacío"
    
    # Limpiar RUT
    rut_clean = clean_rut(rut)
    
    # Validar formato básico
    if len(rut_clean) < 2:
        return False, "El RUT debe tener al menos 2 caracteres"
    
    # Separar número y dígito verificador
    rut_number = rut_clean[:-1]
    dv_provided = rut_clean[-1]
    
    # Validar que el número sea numérico
    if not rut_number.isdigit():
        return False, "El RUT debe contener solo números antes del dígito verificador"
    
    # Validar longitud
    if len(rut_number) < 7 or len(rut_number) > 8:
        return False, "El RUT debe tener entre 7 y 8 dígitos"
    
    # Calcular dígito verificador
    dv_calculated = calculate_rut_dv(rut_number)
    
    # Comparar
    if dv_provided != dv_calculated:
        return False, "El dígito verificador del RUT es inválido"
    
    return True, "RUT válido"


def validate_email(email: str) -> Tuple[bool, str]:
    """
    Valida un email
    
    Args:
        email (str): Email a validar
        
    Returns:
        Tuple[bool, str]: (es_válido, mensaje_error)
    """
    if not email:
        return False, "El email no puede estar vacío"
    
    # Patrón regex para validar email
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    
    if not re.match(pattern, email):
        return False, "El formato del email es inválido"
    
    return True, "Email válido"


def validate_phone(phone: str) -> Tuple[bool, str]:
    """
    Valida un teléfono chileno
    
    Args:
        phone (str): Teléfono a validar
        
    Returns:
        Tuple[bool, str]: (es_válido, mensaje_error)
    """
    if not phone:
        return False, "El teléfono no puede estar vacío"
    
    # Limpiar caracteres especiales
    phone_clean = re.sub(r'[^\d+]', '', phone)
    
    # Validar longitud (9 dígitos o con +56)
    if len(phone_clean) == 9 and phone_clean.isdigit():
        return True, "Teléfono válido"
    elif len(phone_clean) == 12 and phone_clean.startswith("+56"):
        return True, "Teléfono válido"
    else:
        return False, "El teléfono debe tener 9 dígitos"


def validate_factor_sum(factors: dict, start: int = 8, end: int = 19) -> Tuple[bool, str]:
    """
    Valida que la suma de factores esté dentro del rango permitido
    Requerimiento A-01
    
    Args:
        factors (dict): Diccionario con factores {factor_id: valor}
        start (int): Factor inicial (por defecto 8)
        end (int): Factor final (por defecto 19)
        
    Returns:
        Tuple[bool, str]: (es_válido, mensaje_error)
    """
    total = 0
    
    for i in range(start, end + 1):
        factor_key = f"factor_{i}"
        if factor_key in factors:
            value = factors[factor_key]
            
            # Validar que sea numérico
            try:
                value = float(value)
            except (ValueError, TypeError):
                return False, f"El {factor_key} debe ser un valor numérico"
            
            # Validar rango
            if value < 0 or value > 1:
                return False, f"El {factor_key} debe estar entre 0 y 1"
            
            total += value
    
    # Validar suma total
    if total > 1.0:
        return False, f"La suma de los factores {start} al {end} no puede superar 1.0 (actual: {total:.2f})"
    
    return True, f"Suma de factores válida: {total:.2f}"


def validate_required_fields(data: dict, required_fields: list) -> Tuple[bool, str]:
    """
    Valida que los campos requeridos estén presentes y no vacíos
    
    Args:
        data (dict): Datos a validar
        required_fields (list): Lista de campos requeridos
        
    Returns:
        Tuple[bool, str]: (es_válido, mensaje_error)
    """
    for field in required_fields:
        if field not in data or not data[field]:
            return False, f"El campo '{field}' es requerido"
    
    return True, "Todos los campos requeridos están presentes"