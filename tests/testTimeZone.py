from datetime import datetime, timezone, timedelta

# Zona horaria de Chile (UTC-3)
chile_tz = timezone(timedelta(hours=-3))

# Hora actual en Chile
hora_chile = datetime.now(chile_tz)
hora_utc = datetime.now(timezone.utc)

print("="*60)
print("VERIFICACIÓN DE ZONA HORARIA - CHILE")
print("="*60)
print(f"Hora UTC:   {hora_utc.strftime('%Y-%m-%d %H:%M:%S %Z')}")
print(f"Hora Chile: {hora_chile.strftime('%Y-%m-%d %H:%M:%S %Z')}")
print(f"Diferencia: {(hora_chile.utcoffset().total_seconds() / 3600):.0f} horas")
print("="*60)

# Verificar formato para Firebase
print("\nFormato para Firebase:")
print(f"ISO Format: {hora_chile.isoformat()}")
print(f"Timestamp:  {hora_chile.timestamp()}")