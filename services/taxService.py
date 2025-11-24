from datetime import datetime
from typing import Dict, List, Optional, Tuple, Any
from decimal import Decimal
from config.firebaseConfig import firebase_config
from config.settings import Settings
from utils.logger import app_logger, log_audit
from utils.validators import validate_factor_sum
from services.firebaseWrapper import requires_connection, FirebaseServiceBase
from services.subsidyService import SubsidioService


class CalificacionTributariaService(FirebaseServiceBase):
    def __init__(self):
        self.db = firebase_config.get_firestore_client()
        self.datos_ref = self.db.collection(Settings.COLLECTION_DATOS_TRIBUTARIOS)
        self.clientes_ref = self.db.collection(Settings.COLLECTION_USUARIOS)

    def _resolve_and_apply_subsidios(
        self,
        corredor_id: str,
        subsidios_ids: List[str],
        monto_original: float
    ) -> Tuple[List[Dict[str, Any]], float]:
        applied: List[Dict[str, Any]] = []
        try:
            svc = SubsidioService(corredor_id=str(corredor_id), user_id=str(corredor_id))
        except Exception as e:
            app_logger.warning(f"No se pudo inicializar SubsidioService para corredor {corredor_id}: {e}")
            return applied, float(monto_original)

        monto = Decimal(str(monto_original))
        for sid in (subsidios_ids or []):
            try:
                s = svc.get_by_id(sid)
                if not s:
                    app_logger.warning(f"Subsidio {sid} no encontrado para corredor {corredor_id}")
                    continue
                vp = s.get("valor_porcentual", Decimal("0"))
                if not isinstance(vp, Decimal):
                    vp = Decimal(str(vp))
                monto = (monto * (Decimal("1") - vp)).quantize(Decimal("0.01"))
                applied.append({
                    "id": s["id"],
                    "nombre_subsidio": s["nombre_subsidio"],
                    "valor_porcentual": str(vp)
                })
            except Exception as e:
                app_logger.error(f"Error al resolver subsidio {sid}: {e}")
                continue

        return applied, float(monto)

    @requires_connection
    def crear_calificacion(self, datos: Dict, usuario_id: str, corredor_id: Optional[str] = None) -> Dict:
        corredor_id = corredor_id or usuario_id
        try:
            es_valido, mensaje = self._validar_datos(datos)
            if not es_valido:
                return {"success": False, "message": mensaje}

            cliente_id, error = self._validate_cliente(datos.get("cliente_id"))
            if not cliente_id:
                return {
                    "success": False,
                    "message": f"No se puede crear la calificación: {error}",
                    "cliente_no_registrado": True
                }

            fecha_str = datos["fecha_declaracion"].strftime("%Y-%m-%d")
            conflicto = self.buscar_conflicto_oficial(
                cliente_id,
                fecha_str,
                datos["tipo_impuesto"]
            )

            if conflicto:
                return {
                    "success": False,
                    "message": "Ya existe una calificación oficial de bolsa con estos datos",
                    "conflicto": True,
                    "dato_oficial": {
                        "id": conflicto["_id"],
                        "monto": conflicto.get("montoDeclarado", 0),
                        "fecha": conflicto.get("fechaDeclaracion", "")
                    }
                }

            subsidios_ids = datos.get("subsidios_aplicados") or []
            if not isinstance(subsidios_ids, list):
                subsidios_ids = []
            subsidios_ids = [str(x) for x in subsidios_ids if x]

            applied_subsidios, monto_ajustado = self._resolve_and_apply_subsidios(corredor_id, subsidios_ids, datos["monto_declarado"])

            calificacion = {
                "clienteId": cliente_id,
                "usuarioCargaId": usuario_id,
                "propietarioRegistroId": usuario_id,
                "fechaDeclaracion": datos["fecha_declaracion"].strftime("%Y-%m-%d"),
                "tipoImpuesto": datos["tipo_impuesto"],
                "pais": datos["pais"],
                "montoDeclarado": float(datos["monto_declarado"]),
                "montoConSubsidios": float(monto_ajustado),
                "factores": self._preparar_factores(datos["factores"]),
                "subsidiosAplicados": applied_subsidios,
                "esLocal": True,
                "fechaCreacion": datetime.utcnow(),
                "fechaModificacion": datetime.utcnow(),
                "activo": True
            }

            try:
                doc_ref = self.datos_ref.document()  # generar id explícitamente
                doc_ref.set(calificacion)
                calificacion_id = doc_ref.id
            except Exception as e:
                app_logger.error(f"Error guardando calificación en Firestore: {e}")
                return {"success": False, "message": "Error al guardar la calificación en la base de datos"}

            log_audit(
                action="CALIFICACION_CREADA",
                user_id=usuario_id,
                details={"calificacion_id": calificacion_id, "cliente_id": cliente_id, "subsidios": applied_subsidios}
            )

            app_logger.info(f"Calificación creada: {calificacion_id}")

            return {"success": True, "message": "Calificación creada exitosamente", "calificacion_id": calificacion_id}

        except Exception as e:
            app_logger.error(f"Error al crear calificación: {str(e)}")
            return {"success": False, "message": f"Error al crear: {str(e)}"}

    def actualizar_calificacion(self, calificacion_id: str, datos: Dict, usuario_id: str, rol: str = None, corredor_id: Optional[str] = None) -> Dict:
        corredor_id = corredor_id or usuario_id
        try:
            doc = self.datos_ref.document(calificacion_id).get()
            if not doc.exists:
                return {"success": False, "message": "Calificación no encontrada"}

            calificacion_actual = doc.to_dict()

            if rol == "administrador":
                pass
            else:
                if not calificacion_actual.get("esLocal", False):
                    return {"success": False, "message": "No se puede modificar una calificación de bolsa"}
                if calificacion_actual.get("propietarioRegistroId") != usuario_id:
                    return {"success": False, "message": "No tiene permisos para editar esta calificación"}

            es_valido, mensaje = self._validar_datos(datos)
            if not es_valido:
                return {"success": False, "message": mensaje}

            subsidios_ids = datos.get("subsidios_aplicados") or []
            if not isinstance(subsidios_ids, list):
                subsidios_ids = []
            subsidios_ids = [str(x) for x in subsidios_ids if x]

            applied_subsidios, monto_ajustado = self._resolve_and_apply_subsidios(corredor_id, subsidios_ids, datos["monto_declarado"])

            actualizacion = {
                "fechaDeclaracion": datos["fecha_declaracion"].strftime("%Y-%m-%d"),
                "tipoImpuesto": datos["tipo_impuesto"],
                "pais": datos["pais"],
                "montoDeclarado": float(datos["monto_declarado"]),
                "montoConSubsidios": float(monto_ajustado),
                "factores": self._preparar_factores(datos["factores"]),
                "subsidiosAplicados": applied_subsidios,
                "fechaModificacion": datetime.utcnow()
            }

            try:
                self.datos_ref.document(calificacion_id).update(actualizacion)
            except Exception as e:
                app_logger.error(f"Error actualizando calificación en Firestore: {e}")
                return {"success": False, "message": "Error al actualizar la calificación en la base de datos"}

            log_audit(
                action="CALIFICACION_ACTUALIZADA",
                user_id=usuario_id,
                details={"calificacion_id": calificacion_id, "subsidios": applied_subsidios}
            )

            app_logger.info(f"Calificación actualizada: {calificacion_id}")

            return {"success": True, "message": "Calificación actualizada exitosamente"}

        except Exception as e:
            app_logger.error(f"Error al actualizar calificación: {str(e)}")
            return {"success": False, "message": f"Error al actualizar: {str(e)}"}

    def eliminar_calificacion(self, calificacion_id: str, usuario_id: str, rol: str = None) -> Dict:
        try:
            doc = self.datos_ref.document(calificacion_id).get()
            if not doc.exists:
                return {"success": False, "message": "Calificación no encontrada"}

            calificacion = doc.to_dict()

            if rol == "administrador":
                pass
            else:
                if not calificacion.get("esLocal", False):
                    return {"success": False, "message": "No se puede eliminar una calificación de bolsa"}
                if calificacion.get("propietarioRegistroId") != usuario_id:
                    return {"success": False, "message": "No tiene permisos para eliminar esta calificación"}

            try:
                self.datos_ref.document(calificacion_id).update({
                    "activo": False,
                    "fechaEliminacion": datetime.utcnow()
                })
            except Exception as e:
                app_logger.error(f"Error desactivando calificación en Firestore: {e}")
                return {"success": False, "message": "Error al eliminar la calificación en la base de datos"}

            log_audit(
                action="CALIFICACION_ELIMINADA",
                user_id=usuario_id,
                details={"calificacion_id": calificacion_id}
            )

            app_logger.info(f"Calificación eliminada: {calificacion_id}")

            return {"success": True, "message": "Calificación eliminada exitosamente"}

        except Exception as e:
            app_logger.error(f"Error al eliminar calificación: {str(e)}")
            return {"success": False, "message": f"Error al eliminar: {str(e)}"}

    def obtener_calificacion(self, calificacion_id: str) -> Optional[Dict]:
        try:
            doc = self.datos_ref.document(calificacion_id).get()
            if not doc.exists:
                return None
            data = doc.to_dict()
            data["_id"] = doc.id
            return data
        except Exception as e:
            app_logger.error(f"Error al obtener calificación: {str(e)}")
            return None

    def buscar_conflicto_oficial(self, cliente_id: str, fecha: str, tipo_impuesto: str) -> Optional[Dict]:
        try:
            if hasattr(fecha, 'strftime'):
                fecha_str = fecha.strftime("%Y-%m-%d")
            else:
                fecha_str = str(fecha)

            query = self.datos_ref.where("clienteId", "==", cliente_id)\
                                .where("fechaDeclaracion", "==", fecha_str)\
                                .where("tipoImpuesto", "==", tipo_impuesto)\
                                .where("esLocal", "==", False)\
                                .where("activo", "==", True)\
                                .limit(1)

            results = query.stream()
            for doc in results:
                data = doc.to_dict()
                data["_id"] = doc.id
                return data
            return None
        except Exception as e:
            app_logger.error(f"Error al buscar conflicto oficial: {str(e)}")
            return None

    def listar_calificaciones(self, usuario_id: str, rol: str = None, filtros: Optional[Dict] = None) -> List[Dict]:
        try:
            query = self.datos_ref.where("activo", "==", True)
            if filtros:
                if "fecha_desde" in filtros:
                    fecha_str = filtros["fecha_desde"].strftime("%Y-%m-%d")
                    query = query.where("fechaDeclaracion", ">=", fecha_str)
                if "fecha_hasta" in filtros:
                    fecha_str = filtros["fecha_hasta"].strftime("%Y-%m-%d")
                    query = query.where("fechaDeclaracion", "<=", fecha_str)
                if "tipo_impuesto" in filtros:
                    query = query.where("tipoImpuesto", "==", filtros["tipo_impuesto"])
                if "pais" in filtros:
                    query = query.where("pais", "==", filtros["pais"])

            docs = query.stream()
            calificaciones = []
            for doc in docs:
                data = doc.to_dict()
                data["_id"] = doc.id
                es_local = data.get("esLocal", False)
                es_propietario = data.get("propietarioRegistroId") == usuario_id
                if rol == "administrador":
                    calificaciones.append(data)
                elif not es_local:
                    calificaciones.append(data)
                elif es_propietario:
                    calificaciones.append(data)
            app_logger.info(f"Listadas {len(calificaciones)} calificaciones para rol {rol}")
            return calificaciones
        except Exception as e:
            app_logger.error(f"Error al listar calificaciones: {str(e)}")
            return []

    def _validar_datos(self, datos: Dict) -> tuple:
        requeridos = ["fecha_declaracion", "tipo_impuesto", "pais", "monto_declarado", "factores"]
        for campo in requeridos:
            if campo not in datos or not datos[campo]:
                return False, f"El campo '{campo}' es requerido"
        if datos["monto_declarado"] <= 0:
            return False, "El monto declarado debe ser positivo"
        if len(datos["factores"]) != 19:
            return False, "Se requieren exactamente 19 factores"
        for i, valor in enumerate(datos["factores"], 1):
            if not (0 <= valor <= 1):
                return False, f"El factor {i} debe estar entre 0 y 1"
        factores_dict = {f"factor_{i}": datos["factores"][i-1] for i in range(1, 20)}
        es_valido, mensaje = validate_factor_sum(factores_dict, start=8, end=19)
        if not es_valido:
            return False, mensaje
        return True, "Validación exitosa"

    def _preparar_factores(self, factores: List[float]) -> Dict:
        return {f"factor_{i}": float(factores[i-1]) for i in range(1, 20)}

    def _validate_cliente(self, rut: str) -> tuple:
        try:
            query = self.clientes_ref\
                .where("rut", "==", rut)\
                .where("rol", "==", "cliente")\
                .limit(1)
            results = query.stream()
            for doc in results:
                app_logger.debug(f"Cliente encontrado: {rut}")
                return doc.id, None
            error_msg = f"Cliente con RUT {rut} no está registrado en el sistema"
            app_logger.warning(error_msg)
            return None, error_msg
        except Exception as e:
            error_msg = f"Error al buscar cliente {rut}: {str(e)}"
            app_logger.error(error_msg)
            return None, error_msg