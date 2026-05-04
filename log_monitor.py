"""
Script de monitoreo de uso — Unificador Excel.
Uso: python log_monitor.py [--days N]
"""
import argparse
from collections import Counter
from datetime import datetime, timezone, timedelta

from db import get_client


def fmt_size(kb):
    """Formatea KB a unidad legible (KB / MB / GB)."""
    if kb >= 1024 * 1024:
        return f'{kb / (1024 * 1024):.2f} GB'
    if kb >= 1024:
        return f'{kb / 1024:.2f} MB'
    return f'{kb:.1f} KB'


def bar(value, max_val, width=10):
    """Genera barra de progreso ASCII."""
    filled = round(value / max_val * width) if max_val else 0
    return '█' * filled + '░' * (width - filled)


def parse_ts(event):
    """Convierte el campo timestamp del evento a datetime con timezone."""
    ts = event.get('timestamp', '')
    try:
        return datetime.fromisoformat(ts.replace('Z', '+00:00'))
    except Exception:
        return datetime.min.replace(tzinfo=timezone.utc)


def main():
    parser = argparse.ArgumentParser(description='Monitor de uso — Unificador Excel')
    parser.add_argument(
        '--days', type=int, default=None,
        help='Número de días a analizar (default: todo el historial)',
    )
    args = parser.parse_args()

    client = get_client()
    if client is None:
        print('Error: no se pudo conectar a Supabase. Verifica SUPABASE_URL y SUPABASE_KEY en .env')
        return

    # Construir query y traer datos
    query = client.table('usage_events').select('*')
    if args.days:
        since = (datetime.now(timezone.utc) - timedelta(days=args.days)).isoformat()
        query = query.gte('timestamp', since)

    res    = query.execute()
    events = res.data or []

    if not events:
        print('Sin eventos en el período seleccionado.')
        return

    now         = datetime.now(timezone.utc)
    today_start = now.replace(hour=0, minute=0, second=0, microsecond=0)
    week_start  = now - timedelta(days=7)
    month_start = now - timedelta(days=30)

    # --- Cálculos generales ---
    total     = len(events)
    ips       = {e.get('ip') for e in events if e.get('ip')}
    unificar  = sum(1 for e in events if e.get('mode') == 'unificar')
    formatear = sum(1 for e in events if e.get('mode') == 'formatear')

    hoy    = sum(1 for e in events if parse_ts(e) >= today_start)
    semana = sum(1 for e in events if parse_ts(e) >= week_start)
    mes    = sum(1 for e in events if parse_ts(e) >= month_start)

    total_kb = sum(e.get('total_size_kb') or 0 for e in events)
    avg_kb   = total_kb / total if total else 0

    # Archivo más grande (por evento, no por fila individual)
    biggest_name = biggest_kb = None
    for e in events:
        names = e.get('file_names') or []
        kb    = e.get('total_size_kb') or 0
        if kb > (biggest_kb or 0) and names:
            biggest_kb   = kb
            biggest_name = names[0]

    # Top 5 IPs
    ip_events = Counter(e.get('ip') for e in events if e.get('ip'))
    ip_kb     = {}
    for e in events:
        ip = e.get('ip')
        if ip:
            ip_kb[ip] = ip_kb.get(ip, 0) + (e.get('total_size_kb') or 0)
    top_ips = ip_events.most_common(5)

    # Archivos más frecuentes por nombre
    file_counter = Counter()
    for e in events:
        for name in (e.get('file_names') or []):
            file_counter[name] += 1
    top_files = file_counter.most_common(5)

    # Actividad por hora — últimas 24h
    last_24h      = [e for e in events if parse_ts(e) >= now - timedelta(hours=24)]
    hour_counter  = Counter(parse_ts(e).hour for e in last_24h)
    max_hour      = max(hour_counter.values(), default=1)
    active_hours  = sorted(hour_counter.keys())

    # --- Mostrar ---
    periodo = f'últimos {args.days} días' if args.days else 'todo el historial'
    sep     = '═' * 44

    print(f'\n{sep}')
    print(' RESUMEN DE USO — Unificador Excel')
    print(sep)
    print(f' Período analizado:   {periodo}')
    print(f' Total eventos:       {total}')
    print(f' IPs únicas:          {len(ips)}')

    print(f'\n USO POR MODO')
    print(f' ├─ Unificar:         {unificar} eventos')
    print(f' └─ Formatear:        {formatear} eventos')

    print(f'\n FRECUENCIA')
    print(f' ├─ Hoy:              {hoy} eventos')
    print(f' ├─ Últimos 7 días:   {semana} eventos')
    print(f' └─ Último mes:       {mes} eventos')

    print(f'\n VOLUMEN PROCESADO')
    print(f' ├─ Total:            {fmt_size(total_kb)}')
    print(f' ├─ Promedio por uso: {fmt_size(avg_kb)}')
    if biggest_name:
        print(f' └─ Archivo más grande: {biggest_name} ({fmt_size(biggest_kb)})')

    print(f'\n TOP 5 IPs MÁS ACTIVAS')
    for i, (ip, count) in enumerate(top_ips, 1):
        kb = ip_kb.get(ip, 0)
        print(f' {i}. {ip:<18} — {count} eventos — {fmt_size(kb)}')

    print(f'\n ARCHIVOS MÁS FRECUENTES (por nombre)')
    for i, (name, count) in enumerate(top_files, 1):
        print(f' {i}. {name:<30} — {count} veces')

    print(f'\n ACTIVIDAD POR HORA (últimas 24h)')
    if active_hours:
        for h in active_hours:
            count = hour_counter[h]
            print(f' {h:02d}h {bar(count, max_hour)}  {count}')
    else:
        print(' (sin actividad en las últimas 24h)')

    print(f'{sep}\n')


if __name__ == '__main__':
    main()
