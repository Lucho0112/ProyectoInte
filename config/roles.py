class Roles:
    """Definición de roles del sistema"""
    
    ADMINISTRADOR = "administrador"
    ANALISTA_MERCADO = "analista_mercado"
    AUDITOR_TRIBUTARIO = "auditor_tributario"
    ESPECIALISTA_BENEFICIOS = "especialista_beneficios"
    CLIENTE = "cliente"
    
    ALL_ROLES = [
        ADMINISTRADOR,
        ANALISTA_MERCADO,
        AUDITOR_TRIBUTARIO,
        ESPECIALISTA_BENEFICIOS,
        CLIENTE
    ]


class Permisos:
    """Define los permisos por rol"""
    
    PERMISOS_POR_ROL = {
        Roles.ADMINISTRADOR: {
            "calificaciones": True,
            "carga_masiva": True,
            "subsidios": True,
            "consultar": True,
            "reportes": True,
            "usuarios": True
        },
        Roles.ANALISTA_MERCADO: {
            "calificaciones": True,
            "carga_masiva": True,
            "subsidios": False,
            "consultar": True,
            "reportes": True,
            "usuarios": False
        },
        Roles.AUDITOR_TRIBUTARIO: {
            "calificaciones": False,
            "carga_masiva": False,
            "subsidios": False,
            "consultar": True,
            "reportes": True,
            "usuarios": False
        },
        Roles.ESPECIALISTA_BENEFICIOS: {
            "calificaciones": False,
            "carga_masiva": False,
            "subsidios": True,
            "consultar": True,
            "reportes": True,
            "usuarios": False
        },
        Roles.CLIENTE: {
            "calificaciones": False,
            "carga_masiva": False,
            "subsidios": False,
            "consultar": True,  # Solo sus propios datos
            "reportes": True,   # Solo sus propios reportes
            "usuarios": False
        }
    }
    
    @classmethod
    def tiene_permiso(cls, rol: str, modulo: str) -> bool:
        """
        Verifica si un rol tiene permiso para acceder a un módulo
        
        Args:
            rol (str): Rol del usuario
            modulo (str): ID del módulo
            
        Returns:
            bool: True si tiene permiso, False en caso contrario
        """
        if rol not in cls.PERMISOS_POR_ROL:
            return False
        
        return cls.PERMISOS_POR_ROL[rol].get(modulo, False)
    
    @classmethod
    def get_modulos_disponibles(cls, rol: str) -> list:
        """
        Retorna la lista de módulos disponibles para un rol
        
        Args:
            rol (str): Rol del usuario
            
        Returns:
            list: Lista de IDs de módulos permitidos
        """
        if rol not in cls.PERMISOS_POR_ROL:
            return []
        
        return [
            modulo for modulo, permitido 
            in cls.PERMISOS_POR_ROL[rol].items() 
            if permitido
        ]


class ModulosConfig:
    """Configuración de módulos del sistema"""
    
    MODULOS = {
        "calificaciones": {
            "id": "calificaciones",
            "title": "Gestionar Calificaciones",
            "icon": "📊",
            "description": "CRUD de calificaciones tributarias"
        },
        "carga_masiva": {
            "id": "carga_masiva",
            "title": "Carga Masiva",
            "icon": "📤",
            "description": "Importar datos CSV/Excel"
        },
        "subsidios": {
            "id": "subsidios",
            "title": "Gestionar Subsidios",
            "icon": "🎁",
            "description": "Administrar beneficios"
        },
        "consultar": {
            "id": "consultar",
            "title": "Consultar y Filtrar Datos",
            "icon": "🔍",
            "description": "Buscar información tributaria"
        },
        "reportes": {
            "id": "reportes",
            "title": "Generar Reportes",
            "icon": "📄",
            "description": "Crear reportes tributarios"
        },
        "usuarios": {
            "id": "usuarios",
            "title": "Gestionar Usuarios",
            "icon": "👥",
            "description": "Administrar accesos (Solo Admin)"
        }
    }
    
    @classmethod
    def get_modulo(cls, modulo_id: str) -> dict:
        """Retorna la configuración de un módulo"""
        return cls.MODULOS.get(modulo_id, {})
    
    @classmethod
    def get_modulos_por_rol(cls, rol: str) -> list:
        """
        Retorna los módulos disponibles para un rol específico
        
        Args:
            rol (str): Rol del usuario
            
        Returns:
            list: Lista de configuraciones de módulos
        """
        modulos_permitidos = Permisos.get_modulos_disponibles(rol)
        return [
            cls.MODULOS[mod_id] 
            for mod_id in modulos_permitidos 
            if mod_id in cls.MODULOS
        ]