import sqlite3
from comissoes.config import DB_PATH
conn = sqlite3.connect(str(DB_PATH))
cur = conn.cursor()
cur.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")
print([r[0] for r in cur.fetchall()])
conn.close()
