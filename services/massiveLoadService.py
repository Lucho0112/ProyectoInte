import pandas as pd
from datetime import datetime, timezone, timedelta
from typing import Dict, List, Tuple
from config.firebaseConfig import firebase_config
from config.settings import Settings
from utils.logger import app_logger, log_audit
from services.firebaseWrapper import requires_connection, FirebaseServiceBase


class CargaMasivaService:
    """Servicio para importar datos tributarios masivamente"""
    
    def __init__(self):
        super().__init__()  # ← Agregar esta línea
        self.db = firebase_config.get_firestore_client()
        self.datos_ref = self.db.collection(Settings.COLLECTION_DATOS_TRIBUTARIOS)
        self.clientes_ref = self.db.collection(Settings.COLLECTION_USUARIOS)
        self.chile_tz = timezone(timedelta(hours=-3))

    def get_chile_time(self):
        """Retorna la fecha/hora actual en zona horaria de Chile"""
        return datetime.now(self.chile_tz)
    
    @requires_connection
    def import_data(self, df: pd.DataFrame, usuario_carga_id: str, 
                   progress_callback=None) -> Dict:
        """
        Importa datos desde un DataFrame a Firebase
        
        Args:
            df (pd.DataFrame): DataFrame validado con los datos
            usuario_carga_id (str): ID del usuario que realiza la carga
            progress_callback (callable): Función para reportar progreso
            
        Returns:
            Dict: Resultado de la importación {
                "success": bool,
                "created": int,
                "updated": int,
                "errors": List[str],
                "total_processed": int
            }
        """
        created_count = 0
        updated_count = 0
        errors = []
        total = len(df)
        
        app_logger.info(f"Iniciando importación de {total} registros por usuario {usuario_carga_id}")
        
        for idx, row in df.iterrows():
            try:
                # Actualizar progreso
                if progress_callback:
                    progress = int((idx + 1) / total * 100)
                    progress_callback(progress)
                
                # Buscar o crear cliente
                cliente_id = self.validate_cliente(row['cliente_rut'])
                
                if not cliente_id:
                    errors.append(f"Fila {idx + 2}: No se pudo obtener cliente con RUT {row['cliente_rut']}")
                    continue
                
                # Preparar datos tributarios
                dato_tributario = self.prepare_dato_tributario(
                    row, 
                    cliente_id, 
                    usuario_carga_id
                )
                
                # Verificar si ya existe (upsert)
                existing_id = self.find_existing_dato_bolsa(
                    cliente_id,
                    row['fecha_declaracion'],
                    row['tipo_impuesto']
                )
                
                if existing_id:
                    # Actualizar existente
                    self.datos_ref.document(existing_id).update(dato_tributario)
                    updated_count += 1
                    app_logger.debug(f"Dato de bolsa actualizado: {existing_id}")
                else:
                    # Crear nuevo
                    self.datos_ref.add(dato_tributario)
                    created_count += 1
                    app_logger.debug(f"Nuevo dato de bolsa creado para cliente {cliente_id}")
                
            except Exception as e:
                error_msg = f"Fila {idx + 2}: {str(e)}"
                errors.append(error_msg)
                app_logger.error(f"Error en fila {idx + 2}: {str(e)}")
                continue
        
        # Registrar auditoría
        log_audit(
            action="CARGA_MASIVA",
            user_id=usuario_carga_id,
            details={
                "total": total,
                "created": created_count,
                "updated": updated_count,
                "errors": len(errors)
            }
        )
        
        success = len(errors) == 0 or (created_count + updated_count) > 0
        
        app_logger.info(f"Importación completada. Creados: {created_count}, "
                       f"Actualizados: {updated_count}, Errores: {len(errors)}")
        
        return {
            "success": success,
            "created": created_count,
            "updated": updated_count,
            "errors": errors,
            "total_processed": created_count + updated_count
        }
    
    def validate_cliente(self, rut: str) -> tuple:
        """
        Valida que un cliente exista en la base de datos
        NO CREA clientes nuevos
        
        Args:
            rut (str): RUT del cliente
            
        Returns:
            tuple: (cliente_id, error_message)
                Si existe: (cliente_id, None)
                Si no existe: (None, "mensaje de error")
        """
        try:
            # Buscar cliente existente
            query = self.clientes_ref\
                .where("rut", "==", rut)\
                .where("rol", "==", "cliente")\
                .limit(1)
            results = query.stream()
            
            for doc in results:
                app_logger.debug(f"Cliente encontrado: {rut}")
                return doc.id, None
            
            # Cliente NO existe
            error_msg = f"Cliente con RUT {rut} no está registrado en el sistema"
            app_logger.warning(error_msg)
            return None, error_msg
            
        except Exception as e:
            error_msg = f"Error al buscar cliente {rut}: {str(e)}"
            app_logger.error(error_msg)
            return None, error_msg
    
    def prepare_dato_tributario(self, row: pd.Series, cliente_id: str, 
                               usuario_carga_id: str) -> Dict:
        """
        Prepara el diccionario de datos tributarios para Firebase
        
        Args:
            row (pd.Series): Fila del DataFrame
            cliente_id (str): ID del cliente
            usuario_carga_id (str): ID del usuario que carga
            
        Returns:
            Dict: Datos formateados para Firestore
        """
        # Preparar factores
        factores = {}
        for i in range(1, 20):
            factores[f"factor_{i}"] = float(row[f"factor_{i}"])
        
        # Construir documento
        dato = {
            "clienteId": cliente_id,
            "usuarioCargaId": usuario_carga_id,
            "propietarioRegistroId": usuario_carga_id,
            "fechaDeclaracion": row['fecha_declaracion'].strftime("%Y-%m-%d"),
            "montoDeclarado": float(row['monto_declarado']),
            "tipoImpuesto": row['tipo_impuesto'],
            "pais": row['pais'],
            "factores": factores,
            "subsidiosAplicados": [],
            "esLocal": False,
            "fechaCreacion": self.get_chile_time(),
            "fechaModificacion": self.get_chile_time(),
            "activo": True
        }
        
        return dato
    
    def find_existing_dato(self, cliente_id: str, fecha: str, 
                          tipo_impuesto: str) -> str:
        """
        Busca si ya existe un dato tributario con los mismos criterios
        
        Args:
            cliente_id (str): ID del cliente
            fecha (str): Fecha de declaración
            tipo_impuesto (str): Tipo de impuesto
            
        Returns:
            str: ID del documento si existe, None si no existe
        """
        try:
            # Convertir fecha a string para comparación
            if isinstance(fecha, pd.Timestamp):
                fecha_str = fecha.strftime("%Y-%m-%d")
            else:
                fecha_str = str(fecha)
            
            # Buscar documento existente
            query = self.datos_ref.where("clienteId", "==", cliente_id)\
                                 .where("fechaDeclaracion", "==", fecha_str)\
                                 .where("tipoImpuesto", "==", tipo_impuesto)\
                                 .limit(1)
            
            results = query.stream()
            
            for doc in results:
                return doc.id
            
            return None
            
        except Exception as e:
            app_logger.error(f"Error al buscar dato existente: {str(e)}")
            return None
    
    def find_existing_dato_bolsa(self, cliente_id: str, fecha: str, tipo_impuesto: str) -> str:
        """
        Busca si ya existe un dato de BOLSA (no local) con los mismos criterios
        Esto asegura que la carga masiva solo actualice datos de bolsa, no locales
        
        Args:
            cliente_id (str): ID del cliente
            fecha (str): Fecha de declaración
            tipo_impuesto (str): Tipo de impuesto
            
        Returns:
            str: ID del documento si existe, None si no existe
        """
        try:
            # Convertir fecha a string
            if isinstance(fecha, pd.Timestamp):
                fecha_str = fecha.strftime("%Y-%m-%d")
            else:
                fecha_str = str(fecha)
            
            # ← CLAVE: Buscar SOLO datos con esLocal=False (bolsa)
            query = self.datos_ref\
                .where("clienteId", "==", cliente_id)\
                .where("fechaDeclaracion", "==", fecha_str)\
                .where("tipoImpuesto", "==", tipo_impuesto)\
                .where("esLocal", "==", False)\
                .limit(1)
            
            results = query.stream()
            
            for doc in results:
                return doc.id
            
            return None
            
        except Exception as e:
            app_logger.error(f"Error al buscar dato de bolsa: {str(e)}")
            return None
    
    def validate_before_import(self, df: pd.DataFrame) -> Tuple[bool, str]:
        """
        Validación final antes de importar (redundante pero seguro)
        
        Args:
            df (pd.DataFrame): DataFrame a validar
            
        Returns:
            Tuple[bool, str]: (es_válido, mensaje)
        """
        if df is None or df.empty:
            return False, "DataFrame vacío"
        
        # Validar límite de registros
        if len(df) > Settings.MAX_BATCH_SIZE:
            return False, f"El archivo excede el límite de {Settings.MAX_BATCH_SIZE} registros"
        
        return True, "Validación exitosa"
    
    def validate_all_clientes(self, df: pd.DataFrame) -> Dict:
        """
        Pre-valida que TODOS los clientes del archivo existan
        Útil para mostrar errores antes de importar
        
        Args:
            df (pd.DataFrame): DataFrame con los datos a importar
            
        Returns:
            Dict: {
                "valid": bool,
                "missing_ruts": List[str],
                "total_checked": int,
                "message": str
            }
        """
        try:
            missing_ruts = []
            unique_ruts = df['cliente_rut'].unique()
            
            app_logger.info(f"Validando {len(unique_ruts)} RUTs únicos...")
            
            for rut in unique_ruts:
                cliente_id, error = self.validate_cliente(rut)
                if not cliente_id:
                    missing_ruts.append(rut)
            
            if missing_ruts:
                return {
                    "valid": False,
                    "missing_ruts": missing_ruts,
                    "total_checked": len(unique_ruts),
                    "message": f"Se encontraron {len(missing_ruts)} RUTs no registrados"
                }
            
            return {
                "valid": True,
                "missing_ruts": [],
                "total_checked": len(unique_ruts),
                "message": f"Todos los {len(unique_ruts)} RUTs están registrados"
            }
            
        except Exception as e:
            app_logger.error(f"Error al validar clientes: {str(e)}")
            return {
                "valid": False,
                "missing_ruts": [],
                "total_checked": 0,
                "message": f"Error en validación: {str(e)}"
            }