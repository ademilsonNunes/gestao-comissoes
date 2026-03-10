import sqlite3
from pathlib import Path
from comissoes.config import DB_PATH

conn = sqlite3.connect(str(DB_PATH))
cur = conn.cursor()
cur.execute("SELECT COUNT(*) FROM representantes")
print("rep_count", cur.fetchone()[0])
cur.execute("SELECT id,codvend,nome,email FROM representantes ORDER BY id")
print(cur.fetchall())
conn.close()
