import sqlite3
from pathlib import Path


DB_PATH = Path(__file__).resolve().parents[1] / "comissoes.db"


def main() -> None:
    conn = sqlite3.connect(str(DB_PATH))
    cur = conn.cursor()

    cur.execute("SELECT COUNT(*) FROM lancamentos")
    antes_lanc = int(cur.fetchone()[0] or 0)

    cur.execute(
        """
        DELETE FROM lancamentos
        WHERE id IN (
            SELECT l1.id
            FROM lancamentos l1
            WHERE EXISTS (
                SELECT 1
                FROM lancamentos l2
                WHERE l2.nf = l1.nf
                  AND l2.pedido = l1.pedido
                  AND l2.item = l1.item
                  AND l2.codprod = l1.codprod
                  AND l2.codvend = l1.codvend
                  AND l2.tipo = l1.tipo
                  AND (l2.ano > l1.ano OR (l2.ano = l1.ano AND l2.mes > l1.mes))
            )
        )
        """
    )
    removidos = int(cur.rowcount or 0)

    # Remove consolidado sem base de lancamentos.
    cur.execute(
        """
        DELETE FROM comissoes
        WHERE NOT EXISTS (
            SELECT 1 FROM lancamentos l
            WHERE l.codvend = comissoes.codvend
              AND l.mes = comissoes.mes
              AND l.ano = comissoes.ano
        )
        """
    )
    removidos_com = int(cur.rowcount or 0)

    cur.execute("SELECT COUNT(*) FROM lancamentos")
    depois_lanc = int(cur.fetchone()[0] or 0)

    conn.commit()
    conn.close()

    print(f"DB: {DB_PATH}")
    print(f"Lancamentos antes: {antes_lanc}")
    print(f"Lancamentos removidos: {removidos}")
    print(f"Lancamentos depois: {depois_lanc}")
    print(f"Comissoes removidas sem base: {removidos_com}")


if __name__ == "__main__":
    main()
