from pathlib import Path
import sqlite3
from ..config import DB_PATH


def get_conn() -> sqlite3.Connection:
    Path(DB_PATH).parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(DB_PATH))
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_schema() -> None:
    def _has_column(cur: sqlite3.Cursor, table: str, column: str) -> bool:
        cur.execute(f"PRAGMA table_info({table})")
        return any(str(r[1]) == column for r in cur.fetchall())

    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        "CREATE TABLE IF NOT EXISTS apuracoes (id INTEGER PRIMARY KEY, mes INTEGER NOT NULL, ano INTEGER NOT NULL, arquivo_nome TEXT, arquivo_hash TEXT, data_importacao TEXT, status TEXT DEFAULT 'ativa')"
    )
    cur.execute(
        "CREATE UNIQUE INDEX IF NOT EXISTS ux_apuracoes_periodo ON apuracoes(mes, ano)"
    )
    cur.execute(
        "CREATE TABLE IF NOT EXISTS representantes (id INTEGER PRIMARY KEY, codvend TEXT UNIQUE, nome TEXT, email TEXT, corpo_email TEXT, ativo INTEGER DEFAULT 1)"
    )
    cur.execute(
        "CREATE TABLE IF NOT EXISTS supervisores (id INTEGER PRIMARY KEY, nome TEXT, email TEXT)"
    )
    cur.execute(
        "CREATE TABLE IF NOT EXISTS regras_comissao (id INTEGER PRIMARY KEY, codvend TEXT NOT NULL, codcliente TEXT, rede TEXT, uf TEXT, codprod TEXT, percentual REAL NOT NULL, prioridade INTEGER DEFAULT 100, ativo INTEGER DEFAULT 1, descricao TEXT)"
    )
    cur.execute(
        "CREATE TABLE IF NOT EXISTS lancamentos (id INTEGER PRIMARY KEY, apuracao_id INTEGER, emp TEXT, tp TEXT, codvend TEXT, super TEXT, vend TEXT, nf TEXT, pedido TEXT, item TEXT, codprod TEXT, produto TEXT, dtemissao TEXT, vencto TEXT, dtbaixa TEXT, codcliente TEXT, rede TEXT, uf TEXT, cliente TEXT, class_cli TEXT, vlrbruto REAL, vlrliq REAL, comis_cli REAL, comis_vend REAL, comis_prod REAL, tcomisprod REAL, mes INTEGER, ano INTEGER, tipo TEXT, FOREIGN KEY(apuracao_id) REFERENCES apuracoes(id) ON DELETE CASCADE)"
    )
    cur.execute(
        "CREATE UNIQUE INDEX IF NOT EXISTS ux_lancamentos_dedupe ON lancamentos(nf, pedido, item, codprod, codvend, mes, ano, tipo)"
    )
    cur.execute(
        "CREATE TABLE IF NOT EXISTS comissoes (id INTEGER PRIMARY KEY, apuracao_id INTEGER, codvend TEXT, mes INTEGER, ano INTEGER, total_vlrliq REAL, total_comissao REAL, ajuste_desconto REAL DEFAULT 0, ajuste_premio REAL DEFAULT 0, ajuste_obs TEXT, status TEXT, FOREIGN KEY(apuracao_id) REFERENCES apuracoes(id) ON DELETE CASCADE)"
    )
    cur.execute(
        "CREATE TABLE IF NOT EXISTS emails_envio (id INTEGER PRIMARY KEY, representante_id INTEGER, data TEXT, status TEXT, destinatario TEXT, tipo TEXT, FOREIGN KEY(representante_id) REFERENCES representantes(id))"
    )
    cur.execute(
        "CREATE TABLE IF NOT EXISTS ajustes (id INTEGER PRIMARY KEY, lancamento_id INTEGER, campo TEXT, valor TEXT, motivo TEXT, data TEXT, FOREIGN KEY(lancamento_id) REFERENCES lancamentos(id))"
    )
    cur.execute(
        "CREATE TABLE IF NOT EXISTS configuracoes (id INTEGER PRIMARY KEY, smtp_host TEXT, smtp_port INTEGER, smtp_user TEXT, smtp_pass TEXT, smtp_from TEXT)"
    )
    cur.execute(
        "CREATE TABLE IF NOT EXISTS comissao_aglutinacao (id INTEGER PRIMARY KEY, codvend_origem TEXT UNIQUE NOT NULL, codvend_destino TEXT NOT NULL, ativo INTEGER DEFAULT 1, descricao TEXT)"
    )
    cur.execute(
        "CREATE TABLE IF NOT EXISTS importacoes (id INTEGER PRIMARY KEY, apuracao_id INTEGER, arquivo_nome TEXT, arquivo_hash TEXT UNIQUE, mes INTEGER, ano INTEGER, dbc INTEGER, devolucoes INTEGER, data_importacao TEXT, FOREIGN KEY(apuracao_id) REFERENCES apuracoes(id) ON DELETE SET NULL)"
    )
    if not _has_column(cur, "lancamentos", "apuracao_id"):
        cur.execute("ALTER TABLE lancamentos ADD COLUMN apuracao_id INTEGER")
    if not _has_column(cur, "comissoes", "apuracao_id"):
        cur.execute("ALTER TABLE comissoes ADD COLUMN apuracao_id INTEGER")
    if not _has_column(cur, "comissoes", "ajuste_desconto"):
        cur.execute("ALTER TABLE comissoes ADD COLUMN ajuste_desconto REAL DEFAULT 0")
    if not _has_column(cur, "comissoes", "ajuste_premio"):
        cur.execute("ALTER TABLE comissoes ADD COLUMN ajuste_premio REAL DEFAULT 0")
    if not _has_column(cur, "comissoes", "ajuste_obs"):
        cur.execute("ALTER TABLE comissoes ADD COLUMN ajuste_obs TEXT")
    if not _has_column(cur, "importacoes", "apuracao_id"):
        cur.execute("ALTER TABLE importacoes ADD COLUMN apuracao_id INTEGER")
    if not _has_column(cur, "configuracoes", "reabrir_senha_hash"):
        cur.execute("ALTER TABLE configuracoes ADD COLUMN reabrir_senha_hash TEXT")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_lancamentos_apuracao ON lancamentos(apuracao_id)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_comissoes_apuracao ON comissoes(apuracao_id)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_importacoes_apuracao ON importacoes(apuracao_id)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_aglutinacao_destino ON comissao_aglutinacao(codvend_destino)")
    conn.commit()
    conn.close()
