import pandas as pd
from typing import Tuple, List, Dict
from utils.validators import validate_rut, validate_factor_sum
from config.settings import Settings


class CSVValidator:
    """Validador de archivos CSV para carga masiva"""
    
    # Columnas requeridas en el CSV
    REQUIRED_COLUMNS = [
        'cliente_rut',
        'fecha_declaracion',
        'monto_declarado',
        'tipo_impuesto',
        'pais'
    ] + [f'factor_{i}' for i in range(1, 20)]  # factor_1 hasta factor_19
    
    # Tipos de impuesto válidos
    TIPOS_IMPUESTO_VALIDOS = ['Renta', 'IVA', 'Retenciones', 'Patente', 'Timbre']
    
    # Países válidos
    PAISES_VALIDOS = ['Chile', 'Peru', 'Colombia']
    
    @classmethod
    def validate_file(cls, file_path: str) -> Tuple[bool, str, pd.DataFrame]:
        """
        Valida un archivo CSV/Excel completo
        
        Args:
            file_path (str): Ruta del archivo a validar
            
        Returns:
            Tuple[bool, str, pd.DataFrame]: (es_válido, mensaje, dataframe)
        """
        try:
            # Leer archivo según extensión
            if file_path.endswith('.csv'):
                df = pd.read_csv(file_path)
            elif file_path.endswith(('.xlsx', '.xls')):
                df = pd.read_excel(file_path)
            else:
                return False, "Formato de archivo no soportado. Use CSV o Excel.", None
            
            # Validar que no esté vacío
            if df.empty:
                return False, "El archivo está vacío.", None
            
            # Validar columnas requeridas
            is_valid, message = cls._validate_columns(df)
            if not is_valid:
                return False, message, None
            
            # Validar tipos de datos y formato
            is_valid, message = cls._validate_data_types(df)
            if not is_valid:
                return False, message, None
            
            # Validar datos específicos (RUTs, factores, etc.)
            is_valid, message, errors = cls._validate_data_content(df)
            if not is_valid:
                # Si hay errores, retornar el dataframe con la columna de errores
                df['_errores'] = errors
                return False, message, df
            
            return True, f"Archivo válido. {len(df)} registros listos para importar.", df
            
        except Exception as e:
            return False, f"Error al leer el archivo: {str(e)}", None
    
    @classmethod
    def _validate_columns(cls, df: pd.DataFrame) -> Tuple[bool, str]:
        """Valida que todas las columnas requeridas estén presentes"""
        missing_columns = [col for col in cls.REQUIRED_COLUMNS if col not in df.columns]
        
        if missing_columns:
            return False, f"Columnas faltantes: {', '.join(missing_columns)}"
        
        return True, "Columnas correctas"
    
    @classmethod
    def _validate_data_types(cls, df: pd.DataFrame) -> Tuple[bool, str]:
        """Valida los tipos de datos de las columnas"""
        try:
            # Convertir fecha_declaracion a datetime
            df['fecha_declaracion'] = pd.to_datetime(df['fecha_declaracion'], errors='coerce')
            
            # Verificar que las fechas sean válidas
            if df['fecha_declaracion'].isna().any():
                invalid_rows = df[df['fecha_declaracion'].isna()].index.tolist()
                return False, f"Fechas inválidas en filas: {invalid_rows[:5]}"
            
            # Convertir monto_declarado a numérico
            df['monto_declarado'] = pd.to_numeric(df['monto_declarado'], errors='coerce')
            
            if df['monto_declarado'].isna().any():
                invalid_rows = df[df['monto_declarado'].isna()].index.tolist()
                return False, f"Montos inválidos en filas: {invalid_rows[:5]}"
            
            # Convertir factores a numérico
            for i in range(1, 20):
                col = f'factor_{i}'
                df[col] = pd.to_numeric(df[col], errors='coerce')
                
                if df[col].isna().any():
                    invalid_rows = df[df[col].isna()].index.tolist()
                    return False, f"Factor {i} inválido en filas: {invalid_rows[:5]}"
            
            return True, "Tipos de datos correctos"
            
        except Exception as e:
            return False, f"Error al validar tipos de datos: {str(e)}"
    
    @classmethod
    def _validate_data_content(cls, df: pd.DataFrame) -> Tuple[bool, str, List[str]]:
        """Valida el contenido de los datos (RUTs, rangos, etc.)"""
        errors = []
        
        for idx, row in df.iterrows():
            row_errors = []
            
            # Validar RUT
            rut = str(row['cliente_rut'])
            is_valid_rut, rut_msg = validate_rut(rut)
            if not is_valid_rut:
                row_errors.append(f"RUT inválido: {rut_msg}")
            
            # Validar monto positivo
            if row['monto_declarado'] <= 0:
                row_errors.append("Monto debe ser positivo")
            
            # Validar tipo de impuesto
            if row['tipo_impuesto'] not in cls.TIPOS_IMPUESTO_VALIDOS:
                row_errors.append(f"Tipo de impuesto inválido. Debe ser: {', '.join(cls.TIPOS_IMPUESTO_VALIDOS)}")
            
            # Validar país
            if row['pais'] not in cls.PAISES_VALIDOS:
                row_errors.append(f"País inválido. Debe ser: {', '.join(cls.PAISES_VALIDOS)}")
            
            # Validar rangos de factores (todos entre 0 y 1)
            for i in range(1, 20):
                factor_val = row[f'factor_{i}']
                if not (0 <= factor_val <= 1):
                    row_errors.append(f"Factor_{i} fuera de rango [0-1]: {factor_val}")
            
            # Validar suma de factores 8-19 (REQUERIMIENTO A-01)
            factores_dict = {f'factor_{i}': row[f'factor_{i}'] for i in range(8, 20)}
            is_valid_sum, sum_msg = validate_factor_sum(factores_dict, start=8, end=19)
            
            if not is_valid_sum:
                row_errors.append(sum_msg)
            
            # Agregar errores de esta fila
            if row_errors:
                errors.append(f"Fila {idx + 2}: " + "; ".join(row_errors))
            else:
                errors.append("")  # Sin errores
        
        # Verificar si hay errores
        total_errors = [e for e in errors if e]
        
        if total_errors:
            # Mostrar solo los primeros 10 errores
            error_preview = "\n".join(total_errors[:10])
            if len(total_errors) > 10:
                error_preview += f"\n... y {len(total_errors) - 10} errores más."
            
            return False, f"Se encontraron {len(total_errors)} errores:\n{error_preview}", errors
        
        return True, "Todos los datos son válidos", errors
    
    @classmethod
    def get_template_dataframe(cls) -> pd.DataFrame:
        """
        Genera un DataFrame con la estructura esperada para usar como plantilla
        
        Returns:
            pd.DataFrame: DataFrame vacío con las columnas correctas
        """
        data = {
            'cliente_rut': ['12345678-5', '21232535-K'],
            'fecha_declaracion': ['2024-11-20', '2024-11-21'],
            'monto_declarado': [1500000, 2300000],
            'tipo_impuesto': ['Renta', 'IVA'],
            'pais': ['Peru', 'Chile']
        }
        
        # Agregar factores del 1 al 19 con valores de ejemplo
        for i in range(1, 20):
            if i >= 8 and i <= 19:
                # Factores críticos con valores pequeños
                data[f'factor_{i}'] = [0.05, 0.03]
            else:
                # Factores base
                data[f'factor_{i}'] = [0.00, 0.00]
        
        return pd.DataFrame(data)
    
    @classmethod
    def export_template(cls, file_path: str) -> bool:
        """
        Exporta una plantilla CSV de ejemplo
        
        Args:
            file_path (str): Ruta donde guardar la plantilla
            
        Returns:
            bool: True si se exportó correctamente
        """
        try:
            df = cls.get_template_dataframe()
            
            if file_path.endswith('.csv'):
                df.to_csv(file_path, index=False)
            elif file_path.endswith(('.xlsx', '.xls')):
                df.to_excel(file_path, index=False)
            else:
                return False
            
            return True
            
        except Exception as e:
            print(f"Error al exportar plantilla: {str(e)}")
            return False