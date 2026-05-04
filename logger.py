"""Logging de eventos de uso en Supabase. Nunca propaga excepciones hacia la app."""
from db import get_client


def log_event(action, ip, files, mode, rows=None):
    """
    Registra un evento en la tabla usage_events.

    action : "upload" | "confirm"
    ip     : IP del cliente (string)
    files  : lista de dicts {name, size_kb}
    mode   : "unificar" | "formatear"
    rows   : filas procesadas (int, opcional)
    """
    try:
        client = get_client()
        if client is None:
            return
        data = {
            'mode':           mode,
            'action':         action,
            'ip':             ip,
            'file_names':     [f['name'] for f in files],
            'total_size_kb':  round(sum(f.get('size_kb', 0) for f in files), 2),
            'rows_processed': rows,
        }
        client.table('usage_events').insert(data).execute()
    except Exception as e:
        print(f'[logger] Error registrando evento: {e}')
