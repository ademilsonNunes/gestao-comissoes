from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
import hashlib
import unicodedata

from openpyxl import load_workbook

from ..config import ARTFATOS_DIR
from ..database.database import get_conn, init_schema
from ..database.models import (
    criar_ou_substituir_apuracao_periodo,
    inserir_lancamentos,
    listar_codvend_distintos,
    upsert_representante,
)


def _normalize_name(value: str) -> str:
    text = unicodedata.normalize("NFKD", str(value or "")).encode("ascii", "ignore").decode("ascii")
    return text.lower().strip()


def _arquivo_hash_sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        while True:
            chunk = f.read(1024 * 1024)
            if not chunk:
                break
            h.update(chunk)
    return h.hexdigest()


def _check_hash_conflict(file_hash: str, mes: int, ano: int) -> None:
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT mes, ano FROM importacoes WHERE arquivo_hash=?", (file_hash,))
    row = cur.fetchone()
    conn.close()
    if not row:
        return
    mes_old, ano_old = int(row[0] or 0), int(row[1] or 0)
    if mes_old != mes or ano_old != ano:
        raise ValueError(
            f"arquivo_ja_importado_em_outro_periodo: arquivo ja importado em {mes_old:02d}/{ano_old}, "
            f"novo periodo solicitado {mes:02d}/{ano}"
        )


def _registrar_importacao(path: Path, file_hash: str, mes: int, ano: int, dbc: int, devolucoes: int, apuracao_id: int) -> None:
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        """
        INSERT INTO importacoes(apuracao_id, arquivo_nome, arquivo_hash, mes, ano, dbc, devolucoes, data_importacao)
        VALUES(?,?,?,?,?,?,?,?)
        ON CONFLICT(arquivo_hash) DO UPDATE SET
            apuracao_id=excluded.apuracao_id,
            arquivo_nome=excluded.arquivo_nome,
            mes=excluded.mes,
            ano=excluded.ano,
            dbc=excluded.dbc,
            devolucoes=excluded.devolucoes,
            data_importacao=excluded.data_importacao
        """,
        (int(apuracao_id or 0) or None, path.name, file_hash, mes, ano, dbc, devolucoes, datetime.utcnow().isoformat()),
    )
    conn.commit()
    conn.close()


def _replace_period_data(mes: int, ano: int) -> None:
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT id FROM apuracoes WHERE mes=? AND ano=?", (mes, ano))
    row = cur.fetchone()
    if row:
        apuracao_id = int(row[0] or 0)
        cur.execute("DELETE FROM ajustes WHERE lancamento_id IN (SELECT id FROM lancamentos WHERE apuracao_id=?)", (apuracao_id,))
        cur.execute("DELETE FROM importacoes WHERE apuracao_id=?", (apuracao_id,))
        cur.execute("DELETE FROM comissoes WHERE apuracao_id=?", (apuracao_id,))
        cur.execute("DELETE FROM lancamentos WHERE apuracao_id=?", (apuracao_id,))
        cur.execute("DELETE FROM apuracoes WHERE id=?", (apuracao_id,))
    cur.execute("DELETE FROM lancamentos WHERE mes=? AND ano=?", (mes, ano))
    cur.execute("DELETE FROM comissoes WHERE mes=? AND ano=?", (mes, ano))
    cur.execute("DELETE FROM importacoes WHERE mes=? AND ano=?", (mes, ano))
    conn.commit()
    conn.close()


def importar_representantes_base() -> int:
    init_schema()
    arquivo = ARTFATOS_DIR / "BASE ENVIO COMISSAO JANEIRO.xlsm"
    wb = load_workbook(filename=str(arquivo), data_only=True)
    ws = wb["Planilha1"] if "Planilha1" in wb.sheetnames else wb.active
    count = 0
    mapa_codvend_nome: Dict[str, str] = {}
    try:
        conn = get_conn()
        cur = conn.cursor()
        cur.execute("SELECT DISTINCT codvend, vend FROM lancamentos WHERE COALESCE(codvend,'')<>''")
        for cod, nome_vend in cur.fetchall():
            key = _normalize_name(nome_vend or "")
            if key and key not in mapa_codvend_nome:
                mapa_codvend_nome[key] = str(cod)
        conn.close()
    except Exception:
        pass

    hdr_row_idx = 1
    for r in range(1, 10):
        vals = [str(c or "").strip().lower() for c in next(ws.iter_rows(min_row=r, max_row=r, values_only=True))]
        if any(v for v in vals):
            hdr_row_idx = r
            break
    header = [str(c or "").strip().lower() for c in next(ws.iter_rows(min_row=hdr_row_idx, max_row=hdr_row_idx, values_only=True))]
    idx = {h: i for i, h in enumerate(header)}

    def col(*names):
        for n in names:
            if n in idx:
                return idx[n]
        return None

    i_cod = col("codvend", "codigo", "cod", "código")
    i_email = col("email", "e-mail")
    i_nome = col("nome", "name")
    i_corpo = col("corpo_email", "corpo do email", "corpo do e-mail", "mensagem", "corpo")
    start_row = hdr_row_idx + 1
    for row in ws.iter_rows(min_row=start_row, values_only=True):
        codvend = str((row[i_cod] if i_cod is not None else row[0]) or "").strip()
        email = str((row[i_email] if i_email is not None else row[1]) or "").strip()
        nome = str((row[i_nome] if i_nome is not None else row[2]) or "").strip()
        corpo = str((row[i_corpo] if i_corpo is not None else row[3]) or "").strip()
        if not codvend and nome:
            nome_key = _normalize_name(nome)
            codvend = mapa_codvend_nome.get(nome_key, "")
            if not codvend:
                for k, v in mapa_codvend_nome.items():
                    if nome_key and (nome_key in k or k in nome_key):
                        codvend = v
                        break
        if codvend:
            upsert_representante(codvend, nome, email, corpo)
            count += 1
    try:
        conn = get_conn()
        cur = conn.cursor()
        cur.execute("UPDATE representantes SET ativo=0 WHERE COALESCE(codvend,'')=''")
        conn.commit()
        conn.close()
    except Exception:
        pass
    if count == 0:
        for cod in listar_codvend_distintos():
            upsert_representante(cod, "", "", "")
            count += 1
    return count


def _sheet_to_dicts(ws) -> List[Dict[str, Any]]:
    headers = [str(c or "").strip() for c in next(ws.iter_rows(min_row=1, max_row=1, values_only=True))]
    registros = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        d = {}
        for i, h in enumerate(headers):
            d[h] = row[i]
        registros.append(d)
    return registros


def _parse_date(val: Any) -> Optional[datetime]:
    if isinstance(val, datetime):
        return val
    if isinstance(val, str):
        s = val.strip()
        for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%m/%d/%Y"):
            try:
                return datetime.strptime(s, fmt)
            except Exception:
                pass
    return None


def _to_int(v: Any) -> Optional[int]:
    if v is None:
        return None
    try:
        if isinstance(v, str) and not v.strip():
            return None
        return int(float(v))
    except Exception:
        return None


def _infer_mes_ano(registros: List[Dict[str, Any]]) -> Tuple[int, int]:
    meses_col = []
    anos_col = []
    for r in registros[:2000]:
        m = _to_int(r.get("MES"))
        a = _to_int(r.get("ANO"))
        if m and 1 <= m <= 12:
            meses_col.append(m)
        if a and 2000 <= a <= 2100:
            anos_col.append(a)

    if meses_col and anos_col:
        mes = max(set(meses_col), key=meses_col.count)
        ano = max(set(anos_col), key=anos_col.count)
        return mes, ano

    meses_dt = []
    anos_dt = []
    for r in registros[:2000]:
        for k in ("DTBAIXA", "DTEMISSAO", "VENCTO"):
            dt = _parse_date(r.get(k))
            if dt:
                meses_dt.append(dt.month)
                anos_dt.append(dt.year)
                break

    if meses_dt and anos_dt:
        mes = max(set(meses_dt), key=meses_dt.count)
        ano = max(set(anos_dt), key=anos_dt.count)
        confianca = meses_dt.count(mes) / max(1, len(meses_dt))
        if confianca < 0.55:
            raise ValueError("periodo_ambiguo: informe mes/ano manualmente no upload")
        return mes, ano

    raise ValueError("periodo_nao_identificado: informe mes/ano manualmente no upload")


def _escolher_abas_vendas(path: Path) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
    wb = load_workbook(filename=str(path), data_only=True)
    db_name = None
    dv_name = None
    for n in wb.sheetnames:
        nl = _normalize_name(n)
        if "dbcomiss" in nl:
            db_name = n
        if "devolu" in nl:
            dv_name = n
    db = wb[db_name] if db_name else wb[wb.sheetnames[0]]
    dv = wb[dv_name] if dv_name else wb[wb.sheetnames[-1]]
    return _sheet_to_dicts(db), _sheet_to_dicts(dv)


def importar_vendas() -> Dict[str, int]:
    init_schema()
    arquivo = ARTFATOS_DIR / "Comissão Março 2026.xlsx"
    if not arquivo.exists():
        candidatos = sorted(ARTFATOS_DIR.glob("Comissão*.xlsx"), reverse=True)
        if not candidatos:
            raise FileNotFoundError("arquivo_padrao_nao_encontrado")
        arquivo = candidatos[0]
    return importar_vendas_arquivo(arquivo, mes_override=3, ano_override=2026)


def importar_vendas_arquivo(path: Path, mes_override: Optional[int] = None, ano_override: Optional[int] = None) -> Dict[str, int]:
    init_schema()
    registros_db, registros_dv = _escolher_abas_vendas(path)
    mes_inf, ano_inf = _infer_mes_ano(registros_db or registros_dv)
    mes = int(mes_override) if mes_override else mes_inf
    ano = int(ano_override) if ano_override else ano_inf
    if not (1 <= mes <= 12):
        raise ValueError("mes_invalido")
    if not (2000 <= ano <= 2100):
        raise ValueError("ano_invalido")

    file_hash = _arquivo_hash_sha256(path)
    _check_hash_conflict(file_hash, mes, ano)
    _replace_period_data(mes, ano)
    apuracao_id = criar_ou_substituir_apuracao_periodo(
        mes=mes,
        ano=ano,
        arquivo_nome=path.name,
        arquivo_hash=file_hash,
        data_importacao=datetime.utcnow().isoformat(),
    )

    for r in registros_db:
        r["MES"] = mes
        r["ANO"] = ano
    for r in registros_dv:
        r["MES"] = mes
        r["ANO"] = ano

    inserir_lancamentos(registros_db, "DB", apuracao_id=apuracao_id)
    inserir_lancamentos(registros_dv, "DS", apuracao_id=apuracao_id)
    _registrar_importacao(path, file_hash, mes, ano, len(registros_db), len(registros_dv), apuracao_id=apuracao_id)
    return {"dbc": len(registros_db), "devolucoes": len(registros_dv), "mes": mes, "ano": ano, "apuracao_id": apuracao_id}
