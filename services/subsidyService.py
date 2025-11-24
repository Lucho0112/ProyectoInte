import os
import sqlite3
from typing import List, Optional, Dict, Any
from decimal import Decimal
import csv
import uuid
from datetime import datetime

from config.settings import Settings
from utils.logger import app_logger, log_audit

try:
    import firebase_admin
    from firebase_admin import credentials, firestore
    FIREBASE_AVAILABLE = True
except Exception:
    FIREBASE_AVAILABLE = False


class SubsidioService:
    def __init__(self, corredor_id: str, data_dir: Optional[str]=None, firebase_cfg: Optional[Dict[str, Any]]=None, user_id: Optional[str]=None):
        self.corredor_id = corredor_id
        default_dir = getattr(Settings, "DATA_DIR", None) or os.path.join(os.path.expanduser("~"), ".tribunsys")
        self.data_dir = data_dir or default_dir
        os.makedirs(self.data_dir, exist_ok=True)
        self.db_path = os.path.join(self.data_dir, f"subsidios_{self.corredor_id}.db")
        self.user_id = user_id
        self._init_local_db()
        self.firestore_client = None
        if firebase_cfg and FIREBASE_AVAILABLE:
            self._init_firestore(firebase_cfg)
        else:
            if firebase_cfg and not FIREBASE_AVAILABLE:
                app_logger.warning("Firebase configurado pero firebase_admin no está instalado. Saltando sincronización.")

    def _init_local_db(self):
        conn = sqlite3.connect(self.db_path)
        cur = conn.cursor()
        cur.execute("""
            CREATE TABLE IF NOT EXISTS subsidios (
                id TEXT PRIMARY KEY,
                nombre_subsidio TEXT NOT NULL,
                valor_porcentual TEXT NOT NULL,
                id_normativa TEXT
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS subsidios_aplicados (
                id TEXT PRIMARY KEY,
                subsidio_id TEXT NOT NULL,
                dato_tributario_id TEXT NOT NULL,
                fecha_aplicacion TEXT NOT NULL,
                detalles TEXT
            )
        """)
        conn.commit()
        conn.close()
        app_logger.debug(f"Inicializada DB local de subsidios: {self.db_path}")

    def _init_firestore(self, firebase_cfg: Dict[str, Any]):
        try:
            if isinstance(firebase_cfg, str) and os.path.exists(firebase_cfg):
                cred = credentials.Certificate(firebase_cfg)
                firebase_admin.initialize_app(cred)
            else:
                if not firebase_admin._apps:
                    app_logger.warning("firebase_cfg no es ruta válida y firebase no está inicializado.")
            self.firestore_client = firestore.client()
            app_logger.info("Firestore inicializado para subsidios (opcional).")
        except Exception as e:
            self.firestore_client = None
            app_logger.error(f"No se pudo inicializar Firestore: {e}")

    def _row_to_dict(self, row) -> Dict[str, Any]:
        return {
            "id": row[0],
            "nombre_subsidio": row[1],
            "valor_porcentual": Decimal(row[2]),
            "id_normativa": row[3]
        }

    def list_all(self) -> List[Dict[str, Any]]:
        conn = sqlite3.connect(self.db_path)
        cur = conn.cursor()
        cur.execute("SELECT id, nombre_subsidio, valor_porcentual, id_normativa FROM subsidios ORDER BY nombre_subsidio")
        rows = cur.fetchall()
        conn.close()
        return [self._row_to_dict(r) for r in rows]

    def get_by_id(self, subsidio_id: str) -> Optional[Dict[str, Any]]:
        conn = sqlite3.connect(self.db_path)
        cur = conn.cursor()
        cur.execute("SELECT id, nombre_subsidio, valor_porcentual, id_normativa FROM subsidios WHERE id = ?", (subsidio_id,))
        r = cur.fetchone()
        conn.close()
        if r:
            return self._row_to_dict(r)
        return None

    def add_or_update(self, subsidio: Dict[str, Any]) -> None:
        sid = subsidio.get("id") or str(uuid.uuid4())
        nombre = str(subsidio["nombre_subsidio"]).strip()
        vp = subsidio["valor_porcentual"]
        if not isinstance(vp, Decimal):
            vp = Decimal(str(vp))
        if vp > 1:
            vp = (vp / Decimal(100)).quantize(Decimal("0.0001"))
        id_norm = subsidio.get("id_normativa")

        conn = sqlite3.connect(self.db_path)
        cur = conn.cursor()
        cur.execute("BEGIN")
        try:
            cur.execute("""
                INSERT INTO subsidios (id, nombre_subsidio, valor_porcentual, id_normativa)
                VALUES (?, ?, ?, ?)
                ON CONFLICT(id) DO UPDATE SET
                    nombre_subsidio=excluded.nombre_subsidio,
                    valor_porcentual=excluded.valor_porcentual,
                    id_normativa=excluded.id_normativa
            """, (sid, nombre, str(vp), id_norm))
            conn.commit()
            app_logger.info(f"Subsidio guardado: {sid} - {nombre}")
            if self.user_id:
                log_audit("SUBSIDIO_GUARDADO", self.user_id, {"subsidio_id": sid, "nombre": nombre})
        except Exception as e:
            conn.rollback()
            app_logger.error(f"Error guardando subsidio {sid}: {e}")
            raise
        finally:
            conn.close()

        if self.firestore_client:
            try:
                doc_ref = self.firestore_client.collection("corredores").document(self.corredor_id).collection(Settings.COLLECTION_SUBSIDIOS).document(sid)
                doc_ref.set({
                    "id": sid,
                    "nombre_subsidio": nombre,
                    "valor_porcentual": str(vp),
                    "id_normativa": id_norm
                })
            except Exception as e:
                app_logger.warning(f"Error sincronizando subsidio a Firestore (no crítico): {e}")

    def delete(self, subsidio_id: str) -> None:
        conn = sqlite3.connect(self.db_path)
        cur = conn.cursor()
        try:
            cur.execute("DELETE FROM subsidios WHERE id = ?", (subsidio_id,))
            conn.commit()
            app_logger.info(f"Subsidio eliminado: {subsidio_id}")
            if self.user_id:
                log_audit("SUBSIDIO_ELIMINADO", self.user_id, {"subsidio_id": subsidio_id})
        except Exception as e:
            conn.rollback()
            app_logger.error(f"Error eliminando subsidio {subsidio_id}: {e}")
            raise
        finally:
            conn.close()

        if self.firestore_client:
            try:
                doc_ref = self.firestore_client.collection("corredores").document(self.corredor_id).collection(Settings.COLLECTION_SUBSIDIOS).document(subsidio_id)
                doc_ref.delete()
            except Exception as e:
                app_logger.warning(f"Error eliminando subsidio en Firestore (no crítico): {e}")

    def import_from_csv(self, csv_path: str) -> Dict[str, int]:
        stats = {"added": 0, "updated": 0, "errors": 0, "rows": 0}
        if not os.path.exists(csv_path):
            app_logger.error(f"CSV no encontrado: {csv_path}")
            return stats

        conn = sqlite3.connect(self.db_path)
        cur = conn.cursor()
        try:
            cur.execute("BEGIN")
            with open(csv_path, newline='', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    stats['rows'] += 1
                    try:
                        nombre = (row.get('nombre_subsidio') or row.get('nombre') or row.get('name') or "").strip()
                        valor_raw = (row.get('valor_porcentual') or row.get('valor') or row.get('porcentaje') or "").strip()
                        id_norm = (row.get('id_normativa') or row.get('id_norm') or "").strip()
                        if not nombre or not valor_raw:
                            stats['errors'] += 1
                            continue
                        valor = Decimal(valor_raw.replace(',', '.'))
                        if valor > 1:
                            valor = (valor / Decimal(100)).quantize(Decimal('0.0001'))
                        id_norm_val = id_norm if id_norm else None

                        existing_id = None
                        if id_norm_val is not None:
                            cur.execute("SELECT id FROM subsidios WHERE id_normativa = ? LIMIT 1", (id_norm_val,))
                            rowid = cur.fetchone()
                            if rowid:
                                existing_id = rowid[0]
                        if not existing_id:
                            cur.execute("SELECT id FROM subsidios WHERE nombre_subsidio = ? LIMIT 1", (nombre,))
                            rowid = cur.fetchone()
                            if rowid:
                                existing_id = rowid[0]

                        if existing_id:
                            cur.execute("""
                                UPDATE subsidios SET nombre_subsidio = ?, valor_porcentual = ?, id_normativa = ? WHERE id = ?
                            """, (nombre, str(valor), id_norm_val, existing_id))
                            stats['updated'] += 1
                        else:
                            new_id = str(uuid.uuid4())
                            cur.execute("""
                                INSERT INTO subsidios (id, nombre_subsidio, valor_porcentual, id_normativa)
                                VALUES (?, ?, ?, ?)
                            """, (new_id, nombre, str(valor), id_norm_val))
                            stats['added'] += 1
                    except Exception as e:
                        app_logger.error(f"Error procesando fila CSV #{stats['rows']}: {e}")
                        stats['errors'] += 1
            conn.commit()
            app_logger.info(f"Importación CSV completada: {csv_path} -> {stats}")
            if self.user_id:
                log_audit("SUBSIDIO_IMPORT_CSV", self.user_id, {"file": csv_path, "stats": stats})
        except Exception as e:
            conn.rollback()
            app_logger.error(f"Error durante import_from_csv: {e}")
            stats['errors'] += 1
        finally:
            conn.close()
        return stats

    def export_to_csv(self, path: str) -> int:
        subs = self.list_all()
        with open(path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(['id', 'nombre_subsidio', 'valor_porcentual', 'id_normativa'])
            for s in subs:
                writer.writerow([s['id'], str(s['nombre_subsidio']), str(s['valor_porcentual']), s.get('id_normativa') or ""])
        return len(subs)

    def aplicar_subsidio_a_dato(self, subsidio_id: str, dato_tributario_id: str, detalles: str = "") -> None:
        s = self.get_by_id(subsidio_id)
        if not s:
            raise ValueError("Subsidio no encontrado")
        conn = sqlite3.connect(self.db_path)
        cur = conn.cursor()
        try:
            cur.execute("BEGIN")
            aplic_id = str(uuid.uuid4())
            fecha = datetime.utcnow().isoformat() + "Z"
            cur.execute("""
                INSERT INTO subsidios_aplicados (id, subsidio_id, dato_tributario_id, fecha_aplicacion, detalles)
                VALUES (?, ?, ?, ?, ?)
            """, (aplic_id, subsidio_id, dato_tributario_id, fecha, detalles))
            conn.commit()
            app_logger.info(f"Subsidio {subsidio_id} aplicado a dato {dato_tributario_id} (local) por user {self.user_id}")
            if self.user_id:
                log_audit("SUBSIDIO_APLICADO", self.user_id, {"subsidio_id": subsidio_id, "dato_id": dato_tributario_id})
        except Exception as e:
            conn.rollback()
            app_logger.error(f"Error aplicando subsidio localmente: {e}")
            raise
        finally:
            conn.close()

    def delete_all_subsidios(self) -> Dict[str, int]:
        stats = {"deleted": 0, "errors": 0}
        conn = sqlite3.connect(self.db_path)
        cur = conn.cursor()
        try:
            cur.execute("BEGIN")
            cur.execute("SELECT COUNT(*) FROM subsidios")
            count = cur.fetchone()[0]
            cur.execute("DELETE FROM subsidios")
            cur.execute("DELETE FROM subsidios_aplicados")
            conn.commit()
            stats["deleted"] = count
            app_logger.info(f"Se eliminaron {count} subsidios locales del corredor {self.corredor_id}")
            if self.user_id:
                log_audit("SUBSIDIO_ELIMINAR_TODOS", self.user_id, {"deleted": count})
        except Exception as e:
            conn.rollback()
            app_logger.error(f"Error eliminando todos los subsidios: {e}")
            stats["errors"] += 1
            raise
        finally:
            conn.close()

        if self.firestore_client:
            try:
                col_ref = self.firestore_client.collection("corredores").document(self.corredor_id).collection(Settings.COLLECTION_SUBSIDIOS)
                docs = col_ref.stream()
                batch = self.firestore_client.batch()
                deleted = 0
                for d in docs:
                    batch.delete(d.reference)
                    deleted += 1
                if deleted:
                    batch.commit()
                    app_logger.info(f"Se eliminaron {deleted} subsidios en Firestore (corredor {self.corredor_id})")
            except Exception as e:
                app_logger.warning(f"No se pudo eliminar subsidiios en Firestore (no crítico): {e}")

        return stats