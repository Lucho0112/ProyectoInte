from datetime import datetime, timezone, timedelta
from config.firebaseConfig import firebase_config
from config.settings import Settings
from utils.logger import app_logger
from utils.encryption import hash_password, verify_password
from services.firebaseWrapper import requires_connection, safe_firebase_operation


class AuthService:
    """Servicio para gestionar autenticación de usuarios"""

    def __init__(self):
        self.db = firebase_config.get_firestore_client()
        self.auth_client = firebase_config.get_auth_client()
        self.usuarios_ref = self.db.collection(Settings.COLLECTION_USUARIOS)
        # Zona horaria de Chile (UTC-3 o UTC-4 según horario de verano)
        self.chile_tz = timezone(timedelta(hours=-3))

    def get_chile_time(self):
        """Retorna la fecha/hora actual en zona horaria de Chile"""
        return datetime.now(self.chile_tz)

    @requires_connection
    def login(self, rut: str, password: str) -> dict:
        """
        Autentica un usuario

        Args:
            rut (str): RUT del usuario
            password (str): Contraseña en texto plano

        Returns:
            dict: {
                "success": bool,
                "message": str,
                "user": dict (si success=True)
            }
        """
        try:
            # Buscar usuario por RUT
            query = self.usuarios_ref.where("rut", "==", rut).limit(1)
            users = query.stream()

            user_doc = None
            for doc in users:
                user_doc = doc
                break

            if not user_doc:
                app_logger.warning(f"Intento de login con RUT no registrado: {rut}")
                return {
                    "success": False,
                    "message": "RUT o contraseña incorrectos"
                }

            # Obtener datos del usuario
            user_data = user_doc.to_dict()
            user_data["_id"] = user_doc.id
            email = user_data.get("correo") or user_data.get("email")

            if not email:
                app_logger.error(f"Usuario {rut} no tiene email/correo registrado")
                return {
                    "success": False,
                    "message": "Error de configuración de cuenta. Contacte al administrador."
                }

            try:
                auth_response = self.auth_client.sign_in_with_email_and_password(email, password)
                # Si llega aquí, la autenticación fue exitosa
            except Exception as auth_error:
                app_logger.warning(f"Autenticación fallida para RUT: {rut}")
                return {
                    "success": False,
                    "message": "RUT o contraseña incorrectos"
                }

            # Actualizar último acceso con hora de Chile
            self.usuarios_ref.document(user_doc.id).update({
                "ultimoAcceso": self.get_chile_time()
            })

            app_logger.info(f"Login exitoso para usuario: {rut}")

            # No retornar la contraseña
            user_data.pop("contraseña", None)

            return {
                "success": True,
                "message": "Login exitoso",
                "user": user_data
            }

        except Exception as e:
            app_logger.error(f"Error durante el login: {str(e)}")
            return {
                "success": False,
                "message": "Error al autenticar. Por favor intente nuevamente."
            }

    @requires_connection
    def register(self, user_data: dict) -> dict:
        """
        Registra un nuevo usuario

        Args:
            user_data (dict): Datos del usuario a registrar

        Returns:
            dict: {
                "success": bool,
                "message": str,
                "user_id": str (si success=True)
            }
        """
        try:
            rut = user_data.get("rut")
            password = user_data.get("password")
            email = user_data.get("correo")

            # Verificar si el RUT ya existe
            existing_user = self.usuarios_ref.where("rut", "==", rut).limit(1).get()

            if len(list(existing_user)) > 0:
                return {
                    "success": False,
                    "message": "El RUT ya está registrado en el sistema"
                }

            try:
                auth_user = self.auth_client.create_user_with_email_and_password(email, password)
                uid = auth_user['localId']
            except Exception as e:
                return {
                    "success": False,
                    "message": f"Error al crear cuenta de autenticación: {str(e)}"
                }

            user_data_firestore = {
                "rut": rut,
                "nombre": user_data.get("nombre"),
                "apellido_P": user_data.get("apellido_P"),
                "apellido_M": user_data.get("apellido_M"),
                "correo": email,
                "rol": user_data.get("rol", "cliente"),
                "contraseña": hash_password(password),  # Por seguridad también
                "fechaRegistro": self.get_chile_time(),
                "ultimoAcceso": None,
                "activo": True
            }

            # Guardar con UID como document ID
            self.usuarios_ref.document(uid).set(user_data_firestore)

            app_logger.info(f"Usuario registrado: {rut} con UID: {uid}")

            return {
                "success": True,
                "message": "Usuario registrado exitosamente",
                "user_id": uid
            }

        except Exception as e:
            app_logger.error(f"Error durante el registro: {str(e)}")
            return {
                "success": False,
                "message": "Error al registrar usuario. Por favor intente nuevamente."
            }

    @safe_firebase_operation
    def logout(self, user_id: str):
        """
        Cierra la sesión del usuario

        Args:
            user_id (str): ID del usuario
        """
        try:
            # Registrar hora de cierre de sesión con hora de Chile
            self.usuarios_ref.document(user_id).update({
                "ultimaSesionCerrada": self.get_chile_time()
            })

            app_logger.info(f"Sesión cerrada para usuario ID: {user_id}")

        except Exception as e:
            app_logger.error(f"Error al cerrar sesión: {str(e)}")

    @requires_connection
    def change_password(self, user_id: str, old_password: str, new_password: str) -> dict:
        """
        Cambia la contraseña de un usuario

        Args:
            user_id (str): ID del usuario
            old_password (str): Contraseña actual
            new_password (str): Nueva contraseña

        Returns:
            dict: {"success": bool, "message": str}
        """
        try:
            # Obtener usuario
            user_doc = self.usuarios_ref.document(user_id).get()

            if not user_doc.exists:
                return {
                    "success": False,
                    "message": "Usuario no encontrado"
                }

            user_data = user_doc.to_dict()

            # Verificar contraseña actual
            if not verify_password(old_password, user_data.get("contraseña", "")):
                return {
                    "success": False,
                    "message": "La contraseña actual es incorrecta"
                }

            # Actualizar contraseña
            from utils.encryption import hash_password
            new_hashed_password = hash_password(new_password)

            self.usuarios_ref.document(user_id).update({
                "contraseña": new_hashed_password,
                "fechaCambioContraseña": self.get_chile_time()
            })

            app_logger.info(f"Contraseña cambiada para usuario ID: {user_id}")

            return {
                "success": True,
                "message": "Contraseña actualizada exitosamente"
            }

        except Exception as e:
            app_logger.error(f"Error al cambiar contraseña: {str(e)}")
            return {
                "success": False,
                "message": "Error al cambiar contraseña"
            }

    @requires_connection
    def update_user(self, user_id: str, updates: dict) -> dict:
        """
        Actualiza los campos del documento de usuario en Firestore.
        No modifica credenciales en Firebase Auth (solo datos en Firestore).
        """
        try:
            # Normalizar fields: no permitir keys inesperadas o vacías
            allowed = {"nombre", "apellido_P", "apellido_M", "correo", "rol", "contraseña"}
            filtered = {k: v for k, v in updates.items() if k in allowed}
            if not filtered:
                return {"success": False, "message": "No hay campos válidos para actualizar"}

            # Si se envía contraseña, guardamos el hash
            if "contraseña" in filtered:
                filtered["contraseña"] = hash_password(filtered["contraseña"])

            filtered["fechaModificacion"] = self.get_chile_time()

            self.usuarios_ref.document(user_id).update(filtered)
            app_logger.info(f"Usuario actualizado: {user_id} -> {list(filtered.keys())}")
            # opcional: registrar auditoría (si existe log_audit importable)
            try:
                from utils.logger import log_audit
                log_audit("USUARIO_ACTUALIZADO", user_id, {"fields": list(filtered.keys())})
            except Exception:
                pass
            return {"success": True, "message": "Usuario actualizado correctamente"}
        except Exception as e:
            app_logger.error(f"Error actualizando usuario {user_id}: {e}")
            return {"success": False, "message": f"Error al actualizar usuario: {str(e)}"}

    @requires_connection
    def deactivate_user(self, user_id: str) -> dict:
        """
        Marca al usuario como inactivo (campo 'activo' = False) y registra fecha de eliminación.
        """
        try:
            self.usuarios_ref.document(user_id).update({
                "activo": False,
                "fechaEliminacion": self.get_chile_time()
            })
            app_logger.info(f"Usuario desactivado: {user_id}")
            try:
                from utils.logger import log_audit
                log_audit("USUARIO_DESACTIVADO", user_id, {})
            except Exception:
                pass
            return {"success": True, "message": "Usuario desactivado correctamente"}
        except Exception as e:
            app_logger.error(f"Error desactivando usuario {user_id}: {e}")
            return {"success": False, "message": f"Error al desactivar usuario: {str(e)}"}

    @requires_connection
    def reactivate_user(self, user_id: str) -> dict:
        """
        Reactiva un usuario (campo 'activo' = True) y elimina fechaEliminacion.
        """
        try:
            updates = {"activo": True}
            updates["fechaEliminacion"] = None
            updates["fechaModificacion"] = self.get_chile_time()
            self.usuarios_ref.document(user_id).update(updates)
            app_logger.info(f"Usuario reactivado: {user_id}")
            try:
                from utils.logger import log_audit
                log_audit("USUARIO_REACTIVADO", user_id, {})
            except Exception:
                pass
            return {"success": True, "message": "Usuario reactivado correctamente"}
        except Exception as e:
            app_logger.error(f"Error reactivando usuario {user_id}: {e}")
            return {"success": False, "message": f"Error al reactivar usuario: {str(e)}"}