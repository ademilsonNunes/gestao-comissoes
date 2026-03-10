from pathlib import Path

import pytest
from openpyxl import Workbook

from comissoes.app import app
from comissoes.database.database import init_schema, get_conn
import comissoes.database.database as db_module
from comissoes.services import importacao


def _make_workbook(path: Path, mes: int, ano: int) -> None:
    wb = Workbook()
    ws_db = wb.active
    ws_db.title = "DBCOMISS"
    headers = [
        "CODVEND",
        "VEND",
        "NF",
        "PEDIDO",
        "ITEM",
        "CODPROD",
        "PRODUTO",
        "VLRBRUTO",
        "VLRLIQ",
        "COMIS_VEND",
        "COMIS_PROD",
        "TCOMISPROD",
        "MES",
        "ANO",
        "DTBAIXA",
    ]
    ws_db.append(headers)
    ws_db.append(["000571", "REP TESTE", "NF1", "P1", "1", "SKU1", "PROD1", 100, 90, 2.5, None, None, mes, ano, "01/01/2026"])

    ws_ds = wb.create_sheet("DEVOLUCOES")
    ws_ds.append(headers)
    ws_ds.append(["000571", "REP TESTE", "NF2", "P2", "1", "SKU1", "PROD1", 50, 45, None, None, 1.1, mes, ano, "01/01/2026"])
    wb.save(path)


@pytest.fixture()
def temp_db(tmp_path, monkeypatch):
    db_path = tmp_path / "test.db"
    monkeypatch.setattr(db_module, "DB_PATH", db_path)
    init_schema()
    return db_path


def test_importacao_infer_mes_ano_prioriza_colunas_mes_ano(temp_db, tmp_path):
    arq = tmp_path / "base.xlsx"
    _make_workbook(arq, mes=3, ano=2026)
    res = importacao.importar_vendas_arquivo(arq)
    assert res["mes"] == 3
    assert res["ano"] == 2026


def test_importacao_bloqueia_mes_ano_diferente_para_mesmo_arquivo(temp_db, tmp_path):
    arq = tmp_path / "base.xlsx"
    _make_workbook(arq, mes=3, ano=2026)
    importacao.importar_vendas_arquivo(arq)
    with pytest.raises(ValueError, match="arquivo_ja_importado_em_outro_periodo"):
        importacao.importar_vendas_arquivo(arq, mes_override=1, ano_override=2026)


def test_reimportacao_mesmo_periodo_substitui_periodo_sem_duplicar(temp_db, tmp_path):
    arq = tmp_path / "base.xlsx"
    _make_workbook(arq, mes=3, ano=2026)
    importacao.importar_vendas_arquivo(arq)
    importacao.importar_vendas_arquivo(arq, mes_override=3, ano_override=2026)
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT COUNT(*) FROM lancamentos WHERE mes=3 AND ano=2026")
    total = cur.fetchone()[0]
    conn.close()
    assert total == 2


def test_email_lote_exige_ids():
    with app.test_client() as c:
        r = c.post("/api/email/lote", json={"ids": [], "assunto": "Teste", "anexos": []})
        assert r.status_code == 400
        assert (r.get_json(silent=True) or {}).get("error") == "ids_requeridos"


def test_ajuste_percentual_recalcula_e_persiste(temp_db, tmp_path):
    arq = tmp_path / "base.xlsx"
    _make_workbook(arq, mes=5, ano=2026)
    importacao.importar_vendas_arquivo(arq)

    with app.test_client() as c:
        c.post("/api/comissoes/calcular", json={"mes": 5, "ano": 2026})
        coms = c.get("/api/comissoes/5/2026").get_json(silent=True) or []
        assert coms
        cid = int(coms[0]["id"])
        lancs = c.get(f"/api/comissoes/{cid}/lancamentos").get_json(silent=True) or []
        assert lancs
        lid = int(lancs[0]["id"])
        r = c.put(
            f"/api/lancamentos/{lid}/percentuais",
            json={"comis_vend": 1.0, "comis_prod": 2.5, "motivo": "ajuste teste"},
        )
        assert r.status_code == 200

    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT vlrliq, comis_vend, comis_prod, tcomisprod FROM lancamentos WHERE id=?", (lid,))
    row = cur.fetchone()
    conn.close()
    assert row is not None
    assert float(row[1] or 0) == pytest.approx(1.0, rel=1e-9)
    assert float(row[2] or 0) == pytest.approx(2.5, rel=1e-9)
    assert float(row[3] or 0) == pytest.approx(float(row[0] or 0) * 0.025, rel=1e-9)


def test_excluir_apuracao_remove_dados_periodo(temp_db, tmp_path):
    arq = tmp_path / "base.xlsx"
    _make_workbook(arq, mes=6, ano=2026)
    importacao.importar_vendas_arquivo(arq)

    with app.test_client() as c:
        r = c.delete("/api/apuracoes/6/2026")
        assert r.status_code == 200
        body = r.get_json(silent=True) or {}
        assert body.get("deleted") is True

    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT COUNT(*) FROM apuracoes WHERE mes=6 AND ano=2026")
    assert int(cur.fetchone()[0] or 0) == 0
    cur.execute("SELECT COUNT(*) FROM lancamentos WHERE mes=6 AND ano=2026")
    assert int(cur.fetchone()[0] or 0) == 0
    cur.execute("SELECT COUNT(*) FROM comissoes WHERE mes=6 AND ano=2026")
    assert int(cur.fetchone()[0] or 0) == 0
    conn.close()
