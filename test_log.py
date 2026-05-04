# test_log.py
from db import get_client

client = get_client()
result = client.table("usage_events").insert({
    "mode": "test",
    "action": "upload",
    "ip": "127.0.0.1",
    "file_names": ["prueba.xlsx"],
    "total_size_kb": 10.5,
    "rows_processed": 100
}).execute()
print(result)