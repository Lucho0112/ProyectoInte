import pandas as pd
from datetime import datetime, date, timezone, timedelta
from typing import Dict, List
from config.firebaseConfig import firebase_config
from config.settings import Settings
from utils.logger import app_logger, log_audit
from services.firebaseWrapper import requires_connection
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows


class ReportService:
    """Servicio para generar y exportar reportes"""
    
    def __init__(self):
        self.db = firebase_config.get_firestore_client()
        self.datos_ref = self.db.collection(Settings.COLLECTION_DATOS_TRIBUTARIOS)
        self.reportes_ref = self.db.collection(Settings.COLLECTION_REPORTES)
        self.usuarios_ref = self.db.collection(Settings.COLLECTION_USUARIOS)
        self.chile_tz = timezone(timedelta(hours=-3))
    
    def get_chile_time(self):
        """Retorna la fecha/hora actual en zona horaria de Chile"""
        return datetime.now(self.chile_tz)
    
    @requires_connection
    def obtener_datos_filtrados(self, filtros: Dict, usuario_id: str, rol: str) -> List[Dict]:
        """
        Obtiene calificaciones tributarias según filtros
        
        Args:
            filtros (Dict): Filtros a aplicar
            usuario_id (str): ID del usuario que genera el reporte
            rol (str): Rol del usuario
            
        Returns:
            List[Dict]: Lista de calificaciones
        """
        try:
            # ← CAMBIO: Query simple sin múltiples where() para evitar necesitar índice
            # Obtenemos TODOS los documentos activos y filtramos en memoria
            query = self.datos_ref.where("activo", "==", True)
            
            # Ejecutar query base
            docs = query.stream()
            
            calificaciones = []
            for doc in docs:
                data = doc.to_dict()
                data["_id"] = doc.id
                
                # ← NUEVO: Filtrar en memoria (Python) en vez de en Firebase
                
                # Filtro de fechas
                if "fecha_desde" in filtros and filtros["fecha_desde"]:
                    fecha_str = filtros["fecha_desde"].strftime("%Y-%m-%d")
                    if data.get("fechaDeclaracion", "") < fecha_str:
                        continue
                
                if "fecha_hasta" in filtros and filtros["fecha_hasta"]:
                    fecha_str = filtros["fecha_hasta"].strftime("%Y-%m-%d")
                    if data.get("fechaDeclaracion", "") > fecha_str:
                        continue
                
                # Filtro de tipo impuesto
                if "tipo_impuesto" in filtros and filtros["tipo_impuesto"]:
                    if data.get("tipoImpuesto", "") != filtros["tipo_impuesto"]:
                        continue
                
                # Filtro de país
                if "pais" in filtros and filtros["pais"]:
                    if data.get("pais", "") != filtros["pais"]:
                        continue
                
                # Filtrar por estado (local/bolsa)
                es_local = data.get("esLocal", False)
                estado_filtro = filtros.get("estado", "ambos")
                
                if estado_filtro == "local" and not es_local:
                    continue
                if estado_filtro == "bolsa" and es_local:
                    continue
                
                # Permisos según rol
                es_propietario = data.get("propietarioRegistroId") == usuario_id
                
                if rol == "administrador":
                    calificaciones.append(data)
                elif not es_local:
                    calificaciones.append(data)
                elif es_propietario:
                    calificaciones.append(data)
            
            app_logger.info(f"Datos filtrados: {len(calificaciones)} registros")
            return calificaciones
        
        except Exception as e:
            app_logger.error(f"Error al obtener datos filtrados: {str(e)}")
            return []
    
    def obtener_rut_cliente(self, cliente_id: str) -> str:
        """Obtiene el RUT de un cliente por su ID"""
        try:
            doc = self.usuarios_ref.document(cliente_id).get()
            if doc.exists:
                return doc.to_dict().get("rut", "N/A")
            return "N/A"
        except:
            return "N/A"
    
    def preparar_dataframe(self, calificaciones: List[Dict]) -> pd.DataFrame:
        """
        Convierte las calificaciones a DataFrame para exportación
        
        Args:
            calificaciones (List[Dict]): Lista de calificaciones
            
        Returns:
            pd.DataFrame: DataFrame con los datos formateados
        """
        if not calificaciones:
            return pd.DataFrame()
        
        data = []
        for cal in calificaciones:
            # Obtener RUT del cliente
            rut_cliente = self.obtener_rut_cliente(cal.get("clienteId", ""))
            
            # Calcular suma factores 8-19
            factores = cal.get("factores", {})
            suma_8_19 = sum(factores.get(f"factor_{i}", 0) for i in range(8, 20))
            
            # Fila base
            fila = {
                "RUT Cliente": rut_cliente,
                "Fecha Declaración": cal.get("fechaDeclaracion", ""),
                "Tipo Impuesto": cal.get("tipoImpuesto", ""),
                "País": cal.get("pais", ""),
                "Monto Declarado": cal.get("montoDeclarado", 0),
            }
            
            # Agregar factores 1-19
            for i in range(1, 20):
                fila[f"Factor {i}"] = factores.get(f"factor_{i}", 0)
            
            # Agregar suma
            fila["Suma Factores 8-19"] = suma_8_19
            fila["Estado"] = "Local" if cal.get("esLocal", False) else "Bolsa"
            fila["Válido"] = "Sí" if suma_8_19 <= 1.0 else "No (>1.0)"
            
            data.append(fila)
        
        df = pd.DataFrame(data)
        return df
    
    def exportar_csv(self, file_path: str, calificaciones: List[Dict], 
                    filtros: Dict, usuario_id: str) -> Dict:
        """
        Exporta calificaciones a CSV
        
        Args:
            file_path (str): Ruta donde guardar el archivo
            calificaciones (List[Dict]): Datos a exportar
            filtros (Dict): Filtros aplicados
            usuario_id (str): ID del usuario
            
        Returns:
            Dict: Resultado de la exportación
        """
        try:
            if not calificaciones:
                return {
                    "success": False,
                    "message": "No hay datos para exportar"
                }
            
            # Crear DataFrame
            df = self.preparar_dataframe(calificaciones)
            
            # Exportar a CSV
            df.to_csv(file_path, index=False, encoding='utf-8-sig')
            
            # Registrar en Firebase
            self.registrar_reporte(
                usuario_id=usuario_id,
                tipo_reporte="exportacion_calificaciones",
                filtros=filtros,
                total_registros=len(calificaciones),
                formato="CSV",
                nombre_archivo=file_path.split("/")[-1]
            )
            
            app_logger.info(f"CSV exportado: {file_path}")
            
            return {
                "success": True,
                "message": f"CSV generado exitosamente con {len(calificaciones)} registros",
                "file_path": file_path
            }
        
        except Exception as e:
            app_logger.error(f"Error al exportar CSV: {str(e)}")
            return {
                "success": False,
                "message": f"Error al exportar: {str(e)}"
            }
    
    def exportar_excel(self, file_path: str, calificaciones: List[Dict], 
                      filtros: Dict, usuario_id: str) -> Dict:
        """
        Exporta calificaciones a Excel con formato
        
        Args:
            file_path (str): Ruta donde guardar el archivo
            calificaciones (List[Dict]): Datos a exportar
            filtros (Dict): Filtros aplicados
            usuario_id (str): ID del usuario
            
        Returns:
            Dict: Resultado de la exportación
        """
        try:
            if not calificaciones:
                return {
                    "success": False,
                    "message": "No hay datos para exportar"
                }
            
            # Crear DataFrame
            df = self.preparar_dataframe(calificaciones)
            
            # Crear workbook
            wb = Workbook()
            
            # HOJA 1: Datos
            ws_datos = wb.active
            ws_datos.title = "Calificaciones Tributarias"
            
            # Escribir datos
            for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
                for c_idx, value in enumerate(row, 1):
                    cell = ws_datos.cell(row=r_idx, column=c_idx, value=value)
                    
                    # Formato de encabezado
                    if r_idx == 1:
                        cell.font = Font(bold=True, color="FFFFFF")
                        cell.fill = PatternFill(start_color="E94E1B", end_color="E94E1B", fill_type="solid")
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                    else:
                        cell.alignment = Alignment(horizontal="left" if c_idx <= 4 else "right")
            
            # Ajustar anchos de columna
            for column in ws_datos.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                ws_datos.column_dimensions[column_letter].width = adjusted_width
            
            # HOJA 2: Resumen
            ws_resumen = wb.create_sheet("Resumen")
            
            # Título
            ws_resumen['A1'] = "RESUMEN DE EXPORTACIÓN"
            ws_resumen['A1'].font = Font(size=14, bold=True, color="E94E1B")
            
            # Metadatos
            ws_resumen['A3'] = "Fecha de generación:"
            ws_resumen['B3'] = self.get_chile_time().strftime("%Y-%m-%d %H:%M:%S")
            
            ws_resumen['A4'] = "Total de registros:"
            ws_resumen['B4'] = len(calificaciones)
            
            ws_resumen['A5'] = "Filtros aplicados:"
            ws_resumen['B5'] = str(filtros)
            
            # Estadísticas
            ws_resumen['A7'] = "ESTADÍSTICAS"
            ws_resumen['A7'].font = Font(size=12, bold=True)
            
            # Calcular estadísticas
            locales = sum(1 for c in calificaciones if c.get("esLocal", False))
            bolsa = len(calificaciones) - locales
            
            total_monto = sum(c.get("montoDeclarado", 0) for c in calificaciones)
            
            ws_resumen['A8'] = "Registros Locales:"
            ws_resumen['B8'] = locales
            
            ws_resumen['A9'] = "Registros Bolsa:"
            ws_resumen['B9'] = bolsa
            
            ws_resumen['A10'] = "Monto Total Declarado:"
            ws_resumen['B10'] = f"${total_monto:,.2f}"
            
            # Ajustar anchos
            ws_resumen.column_dimensions['A'].width = 30
            ws_resumen.column_dimensions['B'].width = 40
            
            # Guardar
            wb.save(file_path)
            
            # Registrar en Firebase
            self.registrar_reporte(
                usuario_id=usuario_id,
                tipo_reporte="exportacion_calificaciones",
                filtros=filtros,
                total_registros=len(calificaciones),
                formato="Excel",
                nombre_archivo=file_path.split("/")[-1]
            )
            
            app_logger.info(f"Excel exportado: {file_path}")
            
            return {
                "success": True,
                "message": f"Excel generado exitosamente con {len(calificaciones)} registros",
                "file_path": file_path
            }
        
        except Exception as e:
            app_logger.error(f"Error al exportar Excel: {str(e)}")
            return {
                "success": False,
                "message": f"Error al exportar: {str(e)}"
            }
    
    @requires_connection
    def registrar_reporte(self, usuario_id: str, tipo_reporte: str, 
                         filtros: Dict, total_registros: int, 
                         formato: str, nombre_archivo: str) -> bool:
        """
        Registra un reporte generado en Firebase
        
        Args:
            usuario_id (str): ID del usuario que generó el reporte
            tipo_reporte (str): Tipo de reporte
            filtros (Dict): Filtros aplicados
            total_registros (int): Total de registros exportados
            formato (str): Formato del reporte (CSV/Excel)
            nombre_archivo (str): Nombre del archivo generado
            
        Returns:
            bool: True si se registró correctamente
        """
        try:
            # ← NUEVO: Convertir datetime.date a datetime para Firestore
            filtros_firestore = {}
            for key, value in filtros.items():
                if isinstance(value, date) and not isinstance(value, datetime):
                    # Convertir date a datetime
                    filtros_firestore[key] = datetime.combine(value, datetime.min.time())
                else:
                    filtros_firestore[key] = value
            
            # Crear diccionario del reporte directamente
            reporte_data = {
                "usuarioGeneradorId": usuario_id,
                "tipoReporte": tipo_reporte,
                "filtrosAplicados": filtros_firestore,  # ← Usar versión convertida
                "totalRegistros": total_registros,
                "formato": formato,
                "nombreArchivo": nombre_archivo,
                "fechaGeneracion": self.get_chile_time()
            }
            
            # Guardar en Firestore
            self.reportes_ref.add(reporte_data)
            
            # Registrar auditoría
            log_audit(
                action="REPORTE_GENERADO",
                user_id=usuario_id,
                details={
                    "tipo": tipo_reporte,
                    "formato": formato,
                    "registros": total_registros,
                    "archivo": nombre_archivo
                }
            )
            
            app_logger.info(f"Reporte registrado: {nombre_archivo}")
            return True
        
        except Exception as e:
            app_logger.error(f"Error al registrar reporte: {str(e)}")
            return False
    
    @requires_connection
    def obtener_historial_reportes(self, usuario_id: str, rol: str) -> List[Dict]:
        """
        Obtiene el historial de reportes generados
        
        Args:
            usuario_id (str): ID del usuario
            rol (str): Rol del usuario
            
        Returns:
            List[Dict]: Lista de reportes
        """
        try:
            # Admin ve todos, usuarios ven solo los suyos
            if rol == "administrador":
                query = self.reportes_ref.order_by("fechaGeneracion", direction="DESCENDING").limit(50)
            else:
                query = self.reportes_ref\
                    .where("usuarioGeneradorId", "==", usuario_id)\
                    .order_by("fechaGeneracion", direction="DESCENDING")\
                    .limit(50)
            
            docs = query.stream()
            
            reportes = []
            for doc in docs:
                data = doc.to_dict()
                data["_id"] = doc.id
                reportes.append(data)
            
            return reportes
        
        except Exception as e:
            app_logger.error(f"Error al obtener historial: {str(e)}")
            return []