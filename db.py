"""Conexión singleton a Supabase. Nunca propaga excepciones hacia la app."""
import os
from dotenv import load_dotenv

load_dotenv()

_client = None


def get_client():
    """Retorna el cliente Supabase inicializado una sola vez.
    Si faltan credenciales o hay error de conexión, retorna None silenciosamente."""
    global _client
    if _client is not None:
        return _client
    try:
        from supabase import create_client
        url = os.getenv('SUPABASE_URL', '')
        key = os.getenv('SUPABASE_KEY', '')
        if not url or not key:
            print('[db] SUPABASE_URL o SUPABASE_KEY no definidos — logging desactivado.')
            return None
        _client = create_client(url, key)
    except Exception as e:
        print(f'[db] Error inicializando Supabase: {e}')
    return _client
