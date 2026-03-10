from typing import List, Dict, Any, Optional, Tuple
from .database import get_conn


def upsert_representante(codvend: str, nome: str, email: str, corpo_email: str) -> None:
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        "INSERT INTO representantes(codvend, nome, email, corpo_email, ativo) VALUES(?,?,?,?,1) ON CONFLICT(codvend) DO UPDATE SET nome=excluded.nome, email=excluded.email, corpo_email=excluded.corpo_email, ativo=1",
        (codvend, nome, email, corpo_email),
    )
    conn.commit()
    conn.close()


def listar_representantes() -> List[Dict[str, Any]]:
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT id, codvend, nome, email, corpo_email, ativo FROM representantes")
    rows = cur.fetchall()
    conn.close()
    return [
        {
            "id": r[0],
            "codvend": r[1],
            "nome": r[2],
            "email": r[3],
            "corpo_email": r[4],
            "ativo": r[5],
        }
        for r in rows
    ]


def desativar_representante(rep_id: int) -> None:
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("UPDATE representantes SET ativo=0 WHERE id=?", (rep_id,))
    conn.commit()
    conn.close()


def atualizar_representante(rep_id: int, dados: Dict[str, Any]) -> None:
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        "UPDATE representantes SET nome=?, email=?, corpo_email=? WHERE id=?",
        (dados.get("nome", ""), dados.get("email", ""), dados.get("corpo_email", ""), rep_id),
    )
    conn.commit()
    conn.close()


def inserir_lancamentos(registros: List[Dict[str, Any]], tipo: str, apuracao_id: Optional[int] = None) -> None:
    def _to_float_or_none(v: Any) -> Optional[float]:
        if v is None:
            return None
        if isinstance(v, str) and not v.strip():
            return None
        try:
            return float(v)
        except Exception:
            return None

    conn = get_conn()
    cur = conn.cursor()
    tipo_norm = (tipo or "").upper()
    for r in registros:
        vlrbruto = _to_float_or_none(r.get("VLRBRUTO"))
        vlrliq = _to_float_or_none(r.get("VLRLIQ"))
        comis_cli = _to_float_or_none(r.get("COMIS_CLI"))
        comis_vend = _to_float_or_none(r.get("COMIS_VEND"))
        comis_prod = _to_float_or_none(r.get("COMIS_PROD"))
        tcomisprod = _to_float_or_none(r.get("TCOMISPROD"))

        # Garantir devolucoes como negativas, mesmo quando vierem positivas do Excel.
        if tipo_norm == "DS":
            if vlrbruto is not None:
                vlrbruto = -abs(vlrbruto)
            if vlrliq is not None:
                vlrliq = -abs(vlrliq)
            if tcomisprod is not None:
                tcomisprod = -abs(tcomisprod)

        cur.execute(
            "INSERT OR IGNORE INTO lancamentos(apuracao_id, emp, tp, codvend, super, vend, nf, pedido, item, codprod, produto, dtemissao, vencto, dtbaixa, codcliente, rede, uf, cliente, class_cli, vlrbruto, vlrliq, comis_cli, comis_vend, comis_prod, tcomisprod, mes, ano, tipo) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
            (
                int(apuracao_id or 0) or None,
                r.get("EMP", ""),
                r.get("TP", ""),
                r.get("CODVEND", ""),
                r.get("SUPER", ""),
                r.get("VEND", ""),
                r.get("NF", ""),
                r.get("PEDIDO", ""),
                r.get("ITEM", ""),
                r.get("CODPROD", ""),
                r.get("PRODUTO", ""),
                r.get("DTEMISSAO", ""),
                r.get("VENCTO", ""),
                r.get("DTBAIXA", ""),
                r.get("CODCLIENTE", ""),
                r.get("REDE", ""),
                r.get("UF", ""),
                r.get("CLIENTE", ""),
                r.get("CLASS_CLI", ""),
                vlrbruto,
                vlrliq,
                comis_cli,
                comis_vend,
                comis_prod,
                tcomisprod,
                int(r.get("MES", 0) or 0),
                int(r.get("ANO", 0) or 0),
                tipo,
            ),
        )
    conn.commit()
    conn.close()


def representantes_faltantes() -> List[str]:
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        "SELECT DISTINCT l.codvend FROM lancamentos l LEFT JOIN representantes r ON r.codvend=l.codvend WHERE r.id IS NULL AND l.codvend<>''"
    )
    rows = cur.fetchall()
    conn.close()
    return [r[0] for r in rows]


def listar_codvend_distintos() -> List[str]:
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT DISTINCT codvend FROM lancamentos WHERE codvend<>''")
    rows = cur.fetchall()
    conn.close()
    return [r[0] for r in rows]


def listar_regras(codvend: str = "") -> List[Dict[str, Any]]:
    conn = get_conn()
    cur = conn.cursor()
    if codvend:
        cur.execute(
            """
            SELECT id, codvend, codcliente, rede, uf, codprod, percentual, prioridade, ativo, descricao
            FROM regras_comissao
            WHERE codvend=?
            ORDER BY prioridade ASC, id ASC
            """,
            (codvend,),
        )
    else:
        cur.execute(
            """
            SELECT id, codvend, codcliente, rede, uf, codprod, percentual, prioridade, ativo, descricao
            FROM regras_comissao
            ORDER BY codvend ASC, prioridade ASC, id ASC
            """
        )
    rows = cur.fetchall()
    conn.close()
    return [
        {
            "id": r[0],
            "codvend": r[1],
            "codcliente": r[2] or "",
            "rede": r[3] or "",
            "uf": r[4] or "",
            "codprod": r[5] or "",
            "percentual": float(r[6] or 0),
            "prioridade": int(r[7] or 100),
            "ativo": int(r[8] or 0),
            "descricao": r[9] or "",
        }
        for r in rows
    ]


def salvar_regra(dados: Dict[str, Any]) -> int:
    conn = get_conn()
    cur = conn.cursor()
    rid = int(dados.get("id", 0) or 0)
    payload = (
        str(dados.get("codvend", "")).strip(),
        str(dados.get("codcliente", "")).strip(),
        str(dados.get("rede", "")).strip(),
        str(dados.get("uf", "")).strip().upper(),
        str(dados.get("codprod", "")).strip(),
        float(dados.get("percentual", 0) or 0),
        int(dados.get("prioridade", 100) or 100),
        int(dados.get("ativo", 1) or 1),
        str(dados.get("descricao", "")).strip(),
    )
    if rid:
        cur.execute(
            """
            UPDATE regras_comissao
            SET codvend=?, codcliente=?, rede=?, uf=?, codprod=?, percentual=?, prioridade=?, ativo=?, descricao=?
            WHERE id=?
            """,
            (*payload, rid),
        )
    else:
        cur.execute(
            """
            INSERT INTO regras_comissao(codvend, codcliente, rede, uf, codprod, percentual, prioridade, ativo, descricao)
            VALUES(?,?,?,?,?,?,?,?,?)
            """,
            payload,
        )
        rid = int(cur.lastrowid or 0)
    conn.commit()
    conn.close()
    return rid


def remover_regra(rid: int) -> None:
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("DELETE FROM regras_comissao WHERE id=?", (rid,))
    conn.commit()
    conn.close()


def _calc_fallback_comissao(l: Dict[str, Any]) -> float:
    vlrliq = float(l.get("vlrliq") or 0)
    comis_vend = l.get("comis_vend")
    comis_prod = l.get("comis_prod")
    tcomisprod = l.get("tcomisprod")
    comis_cli = l.get("comis_cli")

    # Algumas bases trazem COMIS_VEND=0 em todos os itens; nesses casos,
    # usar COMIS_PROD/TCOMISPROD para não zerar comissão indevidamente.
    if tcomisprod is not None and abs(float(tcomisprod or 0)) > 1e-12:
        return float(tcomisprod or 0)
    if comis_prod is not None and abs(float(comis_prod or 0)) > 1e-12:
        return vlrliq * (float(comis_prod) / 100.0)
    if comis_vend is not None and abs(float(comis_vend or 0)) > 1e-12:
        return vlrliq * (float(comis_vend) / 100.0)
    if comis_cli is not None and abs(float(comis_cli or 0)) > 1e-12:
        return vlrliq * (float(comis_cli) / 100.0)
    return 0.0


def _match_regra(l: Dict[str, Any], regra: Dict[str, Any]) -> bool:
    if not int(regra.get("ativo", 0)):
        return False
    if regra.get("codvend") and str(regra.get("codvend")) != str(l.get("codvend", "")):
        return False
    if regra.get("codcliente") and str(regra.get("codcliente")) != str(l.get("codcliente", "")):
        return False
    if regra.get("rede") and str(regra.get("rede")).lower() != str(l.get("rede", "")).lower():
        return False
    if regra.get("uf") and str(regra.get("uf")).upper() != str(l.get("uf", "")).upper():
        return False
    if regra.get("codprod") and str(regra.get("codprod")) != str(l.get("codprod", "")):
        return False
    return True


def _peso_regra(regra: Dict[str, Any]) -> int:
    especificidade = 0
    for key in ("codcliente", "rede", "uf", "codprod"):
        if str(regra.get(key, "")).strip():
            especificidade += 1
    # Menor prioridade vence; em empate, regra mais específica vence.
    return int(regra.get("prioridade", 100) or 100) * 100 - especificidade


def _normalizar_codvend(codvend: Any) -> str:
    return str(codvend or "").strip()


def _carregar_mapa_aglutinacao(cur) -> Dict[str, str]:
    cur.execute(
        """
        SELECT codvend_origem, codvend_destino
        FROM comissao_aglutinacao
        WHERE COALESCE(ativo, 1)=1
        """
    )
    rows = cur.fetchall()
    return {_normalizar_codvend(r[0]): _normalizar_codvend(r[1]) for r in rows if _normalizar_codvend(r[0]) and _normalizar_codvend(r[1])}


def _resolver_codvend_consolidado(codvend: str, mapa: Dict[str, str]) -> str:
    atual = _normalizar_codvend(codvend)
    if not atual:
        return ""
    visitados = set()
    while atual in mapa and mapa[atual] and mapa[atual] not in visitados:
        visitados.add(atual)
        proximo = _normalizar_codvend(mapa.get(atual))
        if not proximo or proximo == atual:
            break
        atual = proximo
    return atual


def _codvends_do_grupo(master_codvend: str, mapa: Dict[str, str]) -> List[str]:
    master = _normalizar_codvend(master_codvend)
    if not master:
        return []
    grupo = set()
    for cod in set(list(mapa.keys()) + list(mapa.values()) + [master]):
        if _resolver_codvend_consolidado(cod, mapa) == master:
            grupo.add(cod)
    grupo.add(master)
    return sorted(grupo)


def calcular_periodo(mes: int, ano: int) -> List[Dict[str, Any]]:
    conn = get_conn()
    cur = conn.cursor()
    mapa_aglutinacao = _carregar_mapa_aglutinacao(cur)
    origens_aglutinadas = sorted({k for k, v in mapa_aglutinacao.items() if k and v})
    cur.execute(
        "SELECT apuracao_id, codvend, status, COALESCE(ajuste_desconto,0), COALESCE(ajuste_premio,0), COALESCE(ajuste_obs,'') FROM comissoes WHERE mes=? AND ano=?",
        (mes, ano),
    )
    status_existente: Dict[Tuple[Optional[int], str], str] = {}
    ajustes_existentes: Dict[Tuple[Optional[int], str], Dict[str, Any]] = {}
    prioridade_status = {"rascunho": 0, "aprovado": 1, "enviado": 2}
    for ap_id, codvend, status, desconto, premio, obs in cur.fetchall():
        codvend_consolidado = _resolver_codvend_consolidado(str(codvend or ""), mapa_aglutinacao)
        key = ((int(ap_id) if ap_id is not None else None), codvend_consolidado)
        novo = str(status or "rascunho").lower()
        atual = status_existente.get(key, "rascunho")
        if prioridade_status.get(novo, 0) >= prioridade_status.get(atual, 0):
            status_existente[key] = novo
        ajustes_existentes[key] = {
            "ajuste_desconto": float(desconto or 0),
            "ajuste_premio": float(premio or 0),
            "ajuste_obs": str(obs or ""),
        }
    cur.execute("DELETE FROM comissoes WHERE mes=? AND ano=?", (mes, ano))
    if origens_aglutinadas:
        ph = ",".join(["?"] * len(origens_aglutinadas))
        cur.execute(
            f"""
            SELECT l.apuracao_id, l.codvend, l.codcliente, l.rede, l.uf, l.codprod, l.vlrliq, l.comis_vend, l.comis_prod, l.tcomisprod, l.comis_cli
            FROM lancamentos l
            LEFT JOIN representantes r ON r.codvend = l.codvend
            WHERE l.mes=? AND l.ano=? AND (COALESCE(r.ativo, 1)=1 OR l.codvend IN ({ph}))
            """,
            (mes, ano, *origens_aglutinadas),
        )
    else:
        cur.execute(
            """
            SELECT l.apuracao_id, l.codvend, l.codcliente, l.rede, l.uf, l.codprod, l.vlrliq, l.comis_vend, l.comis_prod, l.tcomisprod, l.comis_cli
            FROM lancamentos l
            LEFT JOIN representantes r ON r.codvend = l.codvend
            WHERE l.mes=? AND l.ano=? AND COALESCE(r.ativo, 1)=1
            """,
            (mes, ano),
        )
    lancamentos = [
        {
            "apuracao_id": (int(r[0]) if r[0] is not None else None),
            "codvend": r[1] or "",
            "codvend_original": r[1] or "",
            "codcliente": r[2] or "",
            "rede": r[3] or "",
            "uf": r[4] or "",
            "codprod": r[5] or "",
            "vlrliq": float(r[6] or 0),
            "comis_vend": (float(r[7]) if r[7] is not None else None),
            "comis_prod": (float(r[8]) if r[8] is not None else None),
            "tcomisprod": (float(r[9]) if r[9] is not None else None),
            "comis_cli": (float(r[10]) if r[10] is not None else None),
        }
        for r in cur.fetchall()
    ]

    regras = listar_regras()
    regras_por_codvend: Dict[str, List[Dict[str, Any]]] = {}
    for rg in regras:
        regras_por_codvend.setdefault(str(rg.get("codvend", "")), []).append(rg)

    acumulado: Dict[Tuple[Optional[int], str], Dict[str, float]] = {}
    for l in lancamentos:
        apuracao_id = l.get("apuracao_id")
        codvend_original = str(l.get("codvend_original", ""))
        codvend = _resolver_codvend_consolidado(codvend_original, mapa_aglutinacao)
        reg_match = None
        candidatas = regras_por_codvend.get(codvend_original, [])
        if candidatas:
            validas = [rg for rg in candidatas if _match_regra(l, rg)]
            if validas:
                reg_match = sorted(validas, key=_peso_regra)[0]

        vlrliq = float(l.get("vlrliq", 0) or 0)
        if reg_match:
            comissao = vlrliq * (float(reg_match.get("percentual", 0) or 0) / 100.0)
        else:
            comissao = _calc_fallback_comissao(l)

        key = (int(apuracao_id) if apuracao_id is not None else None, codvend)
        if key not in acumulado:
            acumulado[key] = {"total_vlrliq": 0.0, "total_comissao": 0.0}
        acumulado[key]["total_vlrliq"] += vlrliq
        acumulado[key]["total_comissao"] += comissao

    resultados = []
    for (apuracao_id, codvend), tot in acumulado.items():
        total_liq = float(tot.get("total_vlrliq", 0) or 0)
        total_comissao = float(tot.get("total_comissao", 0) or 0)
        percent = (total_comissao / total_liq * 100.0) if total_liq else 0.0
        resultados.append({"apuracao_id": apuracao_id, "codvend": codvend, "total_vlrliq": total_liq, "percent": percent, "total_comissao": total_comissao})

    resultados.sort(key=lambda x: (int(x.get("apuracao_id") or 0), str(x.get("codvend", ""))))
    cur.executemany(
        "INSERT INTO comissoes(apuracao_id, codvend, mes, ano, total_vlrliq, total_comissao, ajuste_desconto, ajuste_premio, ajuste_obs, status) VALUES(?,?,?,?,?,?,?,?,?,?)",
        [
            (
                r.get("apuracao_id"),
                r["codvend"],
                mes,
                ano,
                r["total_vlrliq"],
                r["total_comissao"],
                ajustes_existentes.get(
                    ((int(r.get("apuracao_id")) if r.get("apuracao_id") is not None else None), str(r.get("codvend", ""))),
                    {},
                ).get("ajuste_desconto", 0.0),
                ajustes_existentes.get(
                    ((int(r.get("apuracao_id")) if r.get("apuracao_id") is not None else None), str(r.get("codvend", ""))),
                    {},
                ).get("ajuste_premio", 0.0),
                ajustes_existentes.get(
                    ((int(r.get("apuracao_id")) if r.get("apuracao_id") is not None else None), str(r.get("codvend", ""))),
                    {},
                ).get("ajuste_obs", ""),
                status_existente.get(
                    ((int(r.get("apuracao_id")) if r.get("apuracao_id") is not None else None), str(r.get("codvend", ""))),
                    "rascunho",
                ),
            )
            for r in resultados
        ],
    )
    conn.commit()
    conn.close()
    return resultados


def obter_comissao_periodo(mes: int, ano: int) -> List[Dict[str, Any]]:
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        "SELECT id, apuracao_id, codvend, total_vlrliq, total_comissao, COALESCE(ajuste_desconto,0), COALESCE(ajuste_premio,0), COALESCE(ajuste_obs,''), status FROM comissoes WHERE mes=? AND ano=? ORDER BY codvend ASC",
        (mes, ano),
    )
    rows = cur.fetchall()
    conn.close()
    return [
        {
            "id": r[0],
            "apuracao_id": r[1],
            "codvend": r[2],
            "total_vlrliq": r[3],
            "total_comissao": r[4],
            "ajuste_desconto": r[5],
            "ajuste_premio": r[6],
            "ajuste_obs": r[7],
            "total_comissao_final": float(r[4] or 0) - float(r[5] or 0) + float(r[6] or 0),
            "status": r[8],
        }
        for r in rows
    ]


def listar_periodos_comissoes() -> List[Dict[str, Any]]:
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        """
        SELECT mes, ano, COUNT(*)
        FROM comissoes
        GROUP BY ano, mes
        ORDER BY ano DESC, mes DESC
        """
    )
    rows = cur.fetchall()
    conn.close()
    return [{"mes": int(r[0]), "ano": int(r[1]), "total": int(r[2])} for r in rows]


def listar_periodos_lancamentos() -> List[Dict[str, Any]]:
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        """
        SELECT mes, ano, COUNT(*)
        FROM lancamentos
        WHERE COALESCE(mes,0) > 0 AND COALESCE(ano,0) > 0
        GROUP BY ano, mes
        ORDER BY ano DESC, mes DESC
        """
    )
    rows = cur.fetchall()
    conn.close()
    return [{"mes": int(r[0]), "ano": int(r[1]), "total": int(r[2])} for r in rows]


def aprovar_comissao(cid: int) -> None:
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("UPDATE comissoes SET status='aprovado' WHERE id=?", (cid,))
    conn.commit()
    conn.close()


def cancelar_aprovacao_comissao(cid: int) -> bool:
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT status FROM comissoes WHERE id=?", (cid,))
    r = cur.fetchone()
    if not r:
        conn.close()
        return False
    status = str(r[0] or "").lower()
    if status == "enviado":
        conn.close()
        return False
    cur.execute("UPDATE comissoes SET status='rascunho' WHERE id=?", (cid,))
    conn.commit()
    conn.close()
    return True


def marcar_comissao_enviada(cid: int) -> bool:
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT status FROM comissoes WHERE id=?", (cid,))
    r = cur.fetchone()
    if not r:
        conn.close()
        return False
    status = str(r[0] or "").lower()
    if status != "aprovado":
        conn.close()
        return False
    cur.execute("UPDATE comissoes SET status='enviado' WHERE id=?", (cid,))
    conn.commit()
    conn.close()
    return True


def atualizar_ajustes_financeiros_comissao(cid: int, desconto: float, premio: float, observacao: str = "") -> bool:
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT id FROM comissoes WHERE id=?", (cid,))
    r = cur.fetchone()
    if not r:
        conn.close()
        return False
    cur.execute(
        "UPDATE comissoes SET ajuste_desconto=?, ajuste_premio=?, ajuste_obs=? WHERE id=?",
        (float(desconto or 0), float(premio or 0), str(observacao or ""), cid),
    )
    conn.commit()
    conn.close()
    return True


def obter_representante(rep_id: int) -> Dict[str, Any]:
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT id, codvend, nome, email, corpo_email FROM representantes WHERE id=?", (rep_id,))
    r = cur.fetchone()
    conn.close()
    if not r:
        return {}
    return {"id": r[0], "codvend": r[1], "nome": r[2], "email": r[3], "corpo_email": r[4]}


def obter_representante_por_codvend(codvend: str) -> Dict[str, Any]:
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT id, codvend, nome, email, corpo_email FROM representantes WHERE codvend=?", (codvend,))
    r = cur.fetchone()
    conn.close()
    if not r:
        return {}
    return {"id": r[0], "codvend": r[1], "nome": r[2], "email": r[3], "corpo_email": r[4]}


def obter_ultima_comissao_por_codvend(codvend: str) -> Dict[str, Any]:
    conn = get_conn()
    cur = conn.cursor()
    mapa_aglutinacao = _carregar_mapa_aglutinacao(cur)
    codvend_busca = _resolver_codvend_consolidado(codvend, mapa_aglutinacao)
    cur.execute(
        "SELECT id, codvend, mes, ano, total_vlrliq, total_comissao, COALESCE(ajuste_desconto,0), COALESCE(ajuste_premio,0), COALESCE(ajuste_obs,''), status FROM comissoes WHERE codvend=? ORDER BY ano DESC, mes DESC LIMIT 1",
        (codvend_busca,),
    )
    r = cur.fetchone()
    conn.close()
    if not r:
        return {}
    total_final = float(r[5] or 0) - float(r[6] or 0) + float(r[7] or 0)
    return {
        "id": r[0],
        "codvend": r[1],
        "mes": r[2],
        "ano": r[3],
        "total_vlrliq": r[4],
        "total_comissao": r[5],
        "ajuste_desconto": r[6],
        "ajuste_premio": r[7],
        "ajuste_obs": r[8],
        "total_comissao_final": total_final,
        "status": r[9],
    }


def registrar_email_envio(rep_id: int, destinatario: str, status: str, tipo: str) -> None:
    import datetime
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        "INSERT INTO emails_envio(representante_id, data, status, destinatario, tipo) VALUES(?,?,?,?,?)",
        (rep_id, datetime.datetime.utcnow().isoformat(), status, destinatario, tipo),
    )
    conn.commit()
    conn.close()


def listar_historico_email() -> List[Dict[str, Any]]:
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT id, representante_id, data, status, destinatario, tipo FROM emails_envio ORDER BY id DESC")
    rows = cur.fetchall()
    conn.close()
    return [
        {"id": r[0], "representante_id": r[1], "data": r[2], "status": r[3], "destinatario": r[4], "tipo": r[5]}
        for r in rows
    ]


def obter_comissao_por_id(cid: int) -> Dict[str, Any]:
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT id, apuracao_id, codvend, mes, ano, total_vlrliq, total_comissao, COALESCE(ajuste_desconto,0), COALESCE(ajuste_premio,0), COALESCE(ajuste_obs,''), status FROM comissoes WHERE id=?", (cid,))
    r = cur.fetchone()
    conn.close()
    if not r:
        return {}
    total_final = float(r[6] or 0) - float(r[7] or 0) + float(r[8] or 0)
    return {
        "id": r[0],
        "apuracao_id": r[1],
        "codvend": r[2],
        "mes": r[3],
        "ano": r[4],
        "total_vlrliq": r[5],
        "total_comissao": r[6],
        "ajuste_desconto": r[7],
        "ajuste_premio": r[8],
        "ajuste_obs": r[9],
        "total_comissao_final": total_final,
        "status": r[10],
    }


def status_comissao_por_lancamento(lancamento_id: int) -> str:
    conn = get_conn()
    cur = conn.cursor()
    mapa_aglutinacao = _carregar_mapa_aglutinacao(cur)
    cur.execute("SELECT apuracao_id, codvend, mes, ano FROM lancamentos WHERE id=?", (lancamento_id,))
    base = cur.fetchone()
    if not base:
        conn.close()
        return ""
    apuracao_id = int(base[0]) if base[0] is not None else None
    codvend = _resolver_codvend_consolidado(str(base[1] or ""), mapa_aglutinacao)
    mes = int(base[2] or 0)
    ano = int(base[3] or 0)
    if apuracao_id:
        cur.execute(
            "SELECT status FROM comissoes WHERE codvend=? AND mes=? AND ano=? AND apuracao_id=? ORDER BY id DESC LIMIT 1",
            (codvend, mes, ano, apuracao_id),
        )
        r = cur.fetchone()
        if not r:
            cur.execute(
                "SELECT status FROM comissoes WHERE codvend=? AND mes=? AND ano=? ORDER BY id DESC LIMIT 1",
                (codvend, mes, ano),
            )
            r = cur.fetchone()
    else:
        cur.execute(
            "SELECT status FROM comissoes WHERE codvend=? AND mes=? AND ano=? ORDER BY id DESC LIMIT 1",
            (codvend, mes, ano),
        )
        r = cur.fetchone()
    conn.close()
    return str(r[0] or "").lower() if r else ""


def obter_lancamentos_por_comissao(cid: int) -> List[Dict[str, Any]]:
    com = obter_comissao_por_id(cid)
    if not com:
        return []
    conn = get_conn()
    cur = conn.cursor()
    mapa_aglutinacao = _carregar_mapa_aglutinacao(cur)
    codvends_grupo = _codvends_do_grupo(str(com.get("codvend", "")), mapa_aglutinacao)
    apuracao_id = com.get("apuracao_id")
    if not codvends_grupo:
        conn.close()
        return []
    placeholders = ",".join(["?"] * len(codvends_grupo))
    if apuracao_id:
        cur.execute(
            f"SELECT id, codvend, emp, vend, nf, pedido, item, codprod, produto, dtemissao, vencto, dtbaixa, rede, uf, cliente, vlrliq, comis_vend, comis_prod, tcomisprod FROM lancamentos WHERE codvend IN ({placeholders}) AND mes=? AND ano=? AND apuracao_id=?",
            (*codvends_grupo, com["mes"], com["ano"], apuracao_id),
        )
    else:
        cur.execute(
            f"SELECT id, codvend, emp, vend, nf, pedido, item, codprod, produto, dtemissao, vencto, dtbaixa, rede, uf, cliente, vlrliq, comis_vend, comis_prod, tcomisprod FROM lancamentos WHERE codvend IN ({placeholders}) AND mes=? AND ano=?",
            (*codvends_grupo, com["mes"], com["ano"]),
        )
    rows = cur.fetchall()
    conn.close()
    return [
        {
            "id": r[0],
            "codvend": r[1] or "",
            "emp": r[2] or "",
            "vend": r[3] or "",
            "nf": r[4],
            "pedido": r[5],
            "item": r[6],
            "codprod": r[7],
            "produto": r[8],
            "dtemissao": r[9] or "",
            "vencto": r[10] or "",
            "dtbaixa": r[11] or "",
            "rede": r[12] or "",
            "uf": r[13] or "",
            "cliente": r[14] or "",
            "vlrliq": float(r[15] or 0),
            "comis_vend": (float(r[16]) if r[16] is not None else None),
            "comis_prod": (float(r[17]) if r[17] is not None else None),
            "tcomisprod": (float(r[18]) if r[18] is not None else None),
        }
        for r in rows
    ]


def criar_ou_substituir_apuracao_periodo(
    mes: int,
    ano: int,
    arquivo_nome: str = "",
    arquivo_hash: str = "",
    data_importacao: str = "",
) -> int:
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT id FROM apuracoes WHERE mes=? AND ano=?", (mes, ano))
    existente = cur.fetchone()
    if existente:
        apuracao_id = int(existente[0] or 0)
        cur.execute("DELETE FROM ajustes WHERE lancamento_id IN (SELECT id FROM lancamentos WHERE apuracao_id=?)", (apuracao_id,))
        cur.execute("DELETE FROM comissoes WHERE apuracao_id=?", (apuracao_id,))
        cur.execute("DELETE FROM lancamentos WHERE apuracao_id=?", (apuracao_id,))
        cur.execute("DELETE FROM importacoes WHERE apuracao_id=?", (apuracao_id,))
        cur.execute("DELETE FROM apuracoes WHERE id=?", (apuracao_id,))
    cur.execute(
        "INSERT INTO apuracoes(mes, ano, arquivo_nome, arquivo_hash, data_importacao, status) VALUES(?,?,?,?,?,?)",
        (mes, ano, arquivo_nome, arquivo_hash, data_importacao, "ativa"),
    )
    apuracao_id = int(cur.lastrowid or 0)
    conn.commit()
    conn.close()
    return apuracao_id


def obter_apuracao_por_periodo(mes: int, ano: int) -> Dict[str, Any]:
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        "SELECT id, mes, ano, arquivo_nome, arquivo_hash, data_importacao, status FROM apuracoes WHERE mes=? AND ano=?",
        (mes, ano),
    )
    r = cur.fetchone()
    conn.close()
    if not r:
        return {}
    return {
        "id": int(r[0]),
        "mes": int(r[1]),
        "ano": int(r[2]),
        "arquivo_nome": r[3] or "",
        "arquivo_hash": r[4] or "",
        "data_importacao": r[5] or "",
        "status": r[6] or "ativa",
    }


def excluir_apuracao_periodo(mes: int, ano: int) -> Dict[str, Any]:
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT id FROM apuracoes WHERE mes=? AND ano=?", (mes, ano))
    r = cur.fetchone()
    if not r:
        conn.close()
        return {"deleted": False, "mes": mes, "ano": ano}
    apuracao_id = int(r[0] or 0)
    cur.execute("DELETE FROM ajustes WHERE lancamento_id IN (SELECT id FROM lancamentos WHERE apuracao_id=?)", (apuracao_id,))
    cur.execute("DELETE FROM comissoes WHERE apuracao_id=?", (apuracao_id,))
    cur.execute("DELETE FROM lancamentos WHERE apuracao_id=?", (apuracao_id,))
    cur.execute("DELETE FROM importacoes WHERE apuracao_id=?", (apuracao_id,))
    cur.execute("DELETE FROM apuracoes WHERE id=?", (apuracao_id,))
    conn.commit()
    conn.close()
    return {"deleted": True, "apuracao_id": apuracao_id, "mes": mes, "ano": ano}


def _atualizar_total_comissao_codvend(conn, codvend: str, mes: int, ano: int, apuracao_id: Optional[int]) -> None:
    cur = conn.cursor()
    mapa_aglutinacao = _carregar_mapa_aglutinacao(cur)
    codvend_master = _resolver_codvend_consolidado(codvend, mapa_aglutinacao)
    grupo = _codvends_do_grupo(codvend_master, mapa_aglutinacao)
    if not grupo:
        grupo = [codvend_master]
    placeholders = ",".join(["?"] * len(grupo))
    if apuracao_id:
        cur.execute(
            f"""
            SELECT COALESCE(SUM(vlrliq),0), COALESCE(SUM(tcomisprod),0)
            FROM lancamentos
            WHERE codvend IN ({placeholders}) AND mes=? AND ano=? AND apuracao_id=?
            """,
            (*grupo, mes, ano, int(apuracao_id)),
        )
    else:
        cur.execute(
            f"""
            SELECT COALESCE(SUM(vlrliq),0), COALESCE(SUM(tcomisprod),0)
            FROM lancamentos
            WHERE codvend IN ({placeholders}) AND mes=? AND ano=?
            """,
            (*grupo, mes, ano),
        )
    total_vlrliq, total_comissao = cur.fetchone() or (0.0, 0.0)

    if apuracao_id:
        cur.execute(
            "SELECT id FROM comissoes WHERE codvend=? AND mes=? AND ano=? AND apuracao_id=? LIMIT 1",
            (codvend_master, mes, ano, int(apuracao_id)),
        )
        row = cur.fetchone()
        if not row:
            cur.execute(
                "SELECT id FROM comissoes WHERE codvend=? AND mes=? AND ano=? ORDER BY id DESC LIMIT 1",
                (codvend_master, mes, ano),
            )
            row = cur.fetchone()
    else:
        cur.execute(
            "SELECT id FROM comissoes WHERE codvend=? AND mes=? AND ano=? ORDER BY id DESC LIMIT 1",
            (codvend_master, mes, ano),
        )
        row = cur.fetchone()
    if row:
        cur.execute(
            "UPDATE comissoes SET total_vlrliq=?, total_comissao=? WHERE id=?",
            (float(total_vlrliq or 0), float(total_comissao or 0), int(row[0])),
        )
    else:
        cur.execute(
            "INSERT INTO comissoes(apuracao_id, codvend, mes, ano, total_vlrliq, total_comissao, status) VALUES(?,?,?,?,?,?,?)",
            (int(apuracao_id or 0) or None, codvend_master, mes, ano, float(total_vlrliq or 0), float(total_comissao or 0), "rascunho"),
        )


def atualizar_percentuais_lancamento(
    lancamento_id: int,
    comis_vend: Optional[float],
    comis_prod: Optional[float],
    motivo: str,
) -> Dict[str, Any]:
    import datetime

    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        "SELECT apuracao_id, codvend, mes, ano, vlrliq, comis_vend, comis_prod FROM lancamentos WHERE id=?",
        (lancamento_id,),
    )
    row = cur.fetchone()
    if not row:
        conn.close()
        return {}

    apuracao_id = int(row[0]) if row[0] is not None else None
    codvend = str(row[1] or "")
    mes = int(row[2] or 0)
    ano = int(row[3] or 0)
    vlrliq = float(row[4] or 0)
    atual_vend = float(row[5] or 0)
    atual_prod = float(row[6] or 0)

    novo_vend = atual_vend if comis_vend is None else float(comis_vend)
    novo_prod = atual_prod if comis_prod is None else float(comis_prod)
    percentual_base = novo_prod if abs(novo_prod) > 1e-12 else novo_vend
    nova_comissao = vlrliq * (percentual_base / 100.0)

    cur.execute(
        "UPDATE lancamentos SET comis_vend=?, comis_prod=?, tcomisprod=? WHERE id=?",
        (novo_vend, novo_prod, nova_comissao, lancamento_id),
    )
    cur.execute(
        "INSERT INTO ajustes(lancamento_id, campo, valor, motivo, data) VALUES(?,?,?,?,?)",
        (
            lancamento_id,
            "percentuais",
            f"comis_vend={novo_vend:.6f};comis_prod={novo_prod:.6f};tcomisprod={nova_comissao:.6f}",
            motivo or "Ajuste de percentual",
            datetime.datetime.utcnow().isoformat(),
        ),
    )

    _atualizar_total_comissao_codvend(conn, codvend, mes, ano, apuracao_id)
    conn.commit()
    conn.close()
    return {
        "id": int(lancamento_id),
        "apuracao_id": apuracao_id,
        "codvend": codvend,
        "mes": mes,
        "ano": ano,
        "vlrliq": vlrliq,
        "comis_vend": float(novo_vend),
        "comis_prod": float(novo_prod),
        "tcomisprod": float(nova_comissao),
    }


def registrar_ajuste(lancamento_id: int, campo: str, valor: str, motivo: str) -> None:
    import datetime
    conn = get_conn()
    cur = conn.cursor()
    campos_editaveis = {"comis_vend", "comis_prod", "tcomisprod", "vlrliq"}
    campo_norm = (campo or "").strip().lower()
    codvend = ""
    mes = 0
    ano = 0
    apuracao_id = None
    cur.execute("SELECT apuracao_id, codvend, mes, ano FROM lancamentos WHERE id=?", (lancamento_id,))
    base = cur.fetchone()
    if base:
        apuracao_id = int(base[0]) if base[0] is not None else None
        codvend = str(base[1] or "")
        mes = int(base[2] or 0)
        ano = int(base[3] or 0)
    if campo_norm in campos_editaveis:
        try:
            novo_valor = float(valor)
        except Exception:
            novo_valor = 0.0
        cur.execute(f"UPDATE lancamentos SET {campo_norm}=? WHERE id=?", (novo_valor, lancamento_id))
        if campo_norm in {"comis_vend", "comis_prod", "vlrliq"}:
            cur.execute("SELECT vlrliq, comis_vend, comis_prod FROM lancamentos WHERE id=?", (lancamento_id,))
            r = cur.fetchone() or (0.0, 0.0, 0.0)
            vlrliq = float(r[0] or 0)
            perc_vend = float(r[1] or 0)
            perc_prod = float(r[2] or 0)
            perc_base = perc_prod if abs(perc_prod) > 1e-12 else perc_vend
            cur.execute("UPDATE lancamentos SET tcomisprod=? WHERE id=?", (vlrliq * (perc_base / 100.0), lancamento_id))
    cur.execute(
        "INSERT INTO ajustes(lancamento_id, campo, valor, motivo, data) VALUES(?,?,?,?,?)",
        (lancamento_id, campo_norm or campo, valor, motivo, datetime.datetime.utcnow().isoformat()),
    )
    if codvend and mes and ano:
        _atualizar_total_comissao_codvend(conn, codvend, mes, ano, apuracao_id)
    conn.commit()
    conn.close()


def obter_configuracoes() -> Dict[str, Any]:
    conn = get_conn()
    cur = conn.cursor()
    try:
        cur.execute("CREATE TABLE IF NOT EXISTS configuracoes (id INTEGER PRIMARY KEY, smtp_host TEXT, smtp_port INTEGER, smtp_user TEXT, smtp_pass TEXT, smtp_from TEXT)")
        conn.commit()
    except Exception:
        pass
    cur.execute("PRAGMA table_info(configuracoes)")
    colunas = {str(r[1]) for r in cur.fetchall()}
    if "reabrir_senha_hash" not in colunas:
        cur.execute("ALTER TABLE configuracoes ADD COLUMN reabrir_senha_hash TEXT")
        conn.commit()
    cur.execute("SELECT id, smtp_host, smtp_port, smtp_user, smtp_pass, smtp_from, COALESCE(reabrir_senha_hash,'') FROM configuracoes LIMIT 1")
    r = cur.fetchone()
    if not r:
        conn.close()
        return {"smtp_host": "", "smtp_port": 587, "smtp_user": "", "smtp_pass": "", "smtp_from": "", "has_reabrir_senha": False}
    conn.close()
    return {"id": r[0], "smtp_host": r[1], "smtp_port": r[2], "smtp_user": r[3], "smtp_pass": r[4], "smtp_from": r[5], "has_reabrir_senha": bool(str(r[6] or "").strip())}


def resumo_auditoria() -> Dict[str, Any]:
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT COUNT(*) FROM representantes")
    reps_total = int(cur.fetchone()[0] or 0)
    cur.execute("SELECT COUNT(*) FROM representantes WHERE ativo=1")
    reps_ativos = int(cur.fetchone()[0] or 0)
    cur.execute("SELECT COUNT(*) FROM representantes WHERE ativo=1 AND COALESCE(codvend,'')<>'' AND COALESCE(email,'')=''")
    reps_sem_email = int(cur.fetchone()[0] or 0)
    cur.execute("SELECT COUNT(*) FROM regras_comissao WHERE ativo=1")
    regras_ativas = int(cur.fetchone()[0] or 0)
    cur.execute(
        """
        SELECT COUNT(DISTINCT l.codvend)
        FROM lancamentos l
        LEFT JOIN representantes r ON r.codvend=l.codvend
        WHERE COALESCE(l.codvend,'')<>'' AND r.id IS NULL
        """
    )
    codvend_sem_cadastro = int(cur.fetchone()[0] or 0)
    cur.execute(
        """
        SELECT COUNT(DISTINCT r.codvend)
        FROM representantes r
        LEFT JOIN regras_comissao rc ON rc.codvend=r.codvend AND rc.ativo=1
        WHERE r.ativo=1 AND rc.id IS NULL
        """
    )
    reps_sem_regra = int(cur.fetchone()[0] or 0)
    conn.close()
    return {
        "representantes_total": reps_total,
        "representantes_ativos": reps_ativos,
        "representantes_sem_email": reps_sem_email,
        "regras_ativas": regras_ativas,
        "codvend_sem_cadastro": codvend_sem_cadastro,
        "representantes_sem_regra": reps_sem_regra,
    }


def salvar_configuracoes(cfg: Dict[str, Any]) -> None:
    import hashlib
    conn = get_conn()
    cur = conn.cursor()
    try:
        cur.execute("CREATE TABLE IF NOT EXISTS configuracoes (id INTEGER PRIMARY KEY, smtp_host TEXT, smtp_port INTEGER, smtp_user TEXT, smtp_pass TEXT, smtp_from TEXT)")
        conn.commit()
    except Exception:
        pass
    cur.execute("PRAGMA table_info(configuracoes)")
    colunas = {str(r[1]) for r in cur.fetchall()}
    if "reabrir_senha_hash" not in colunas:
        cur.execute("ALTER TABLE configuracoes ADD COLUMN reabrir_senha_hash TEXT")
        conn.commit()
    senha = str(cfg.get("reabrir_senha", "") or "").strip()
    limpar_senha = int(cfg.get("limpar_reabrir_senha", 0) or 0) == 1
    senha_hash = hashlib.sha256(senha.encode("utf-8")).hexdigest() if senha else ""
    cur.execute("SELECT id, COALESCE(reabrir_senha_hash,'') FROM configuracoes LIMIT 1")
    r = cur.fetchone()
    if r:
        hash_final = senha_hash if senha else ("" if limpar_senha else str(r[1] or ""))
        cur.execute(
            "UPDATE configuracoes SET smtp_host=?, smtp_port=?, smtp_user=?, smtp_pass=?, smtp_from=?, reabrir_senha_hash=? WHERE id=?",
            (cfg.get("smtp_host", ""), int(cfg.get("smtp_port", 587)), cfg.get("smtp_user", ""), cfg.get("smtp_pass", ""), cfg.get("smtp_from", ""), hash_final, r[0]),
        )
    else:
        cur.execute(
            "INSERT INTO configuracoes(smtp_host, smtp_port, smtp_user, smtp_pass, smtp_from, reabrir_senha_hash) VALUES(?,?,?,?,?,?)",
            (cfg.get("smtp_host", ""), int(cfg.get("smtp_port", 587)), cfg.get("smtp_user", ""), cfg.get("smtp_pass", ""), cfg.get("smtp_from", ""), senha_hash),
        )
    conn.commit()
    conn.close()


def reabrir_comissao_enviada(cid: int, senha: str) -> bool:
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT status FROM comissoes WHERE id=?", (cid,))
    r = cur.fetchone()
    if not r:
        conn.close()
        return False
    status = str(r[0] or "").lower()
    if status != "enviado":
        conn.close()
        return False
    if not validar_senha_reabertura(senha, conn=conn):
        conn.close()
        return False
    cur.execute("UPDATE comissoes SET status='rascunho' WHERE id=?", (cid,))
    conn.commit()
    conn.close()
    return True


def obter_hash_senha_reabertura(conn=None) -> str:
    close_conn = False
    if conn is None:
        conn = get_conn()
        close_conn = True
    cur = conn.cursor()
    cur.execute("SELECT COALESCE(reabrir_senha_hash,'') FROM configuracoes LIMIT 1")
    row = cur.fetchone()
    if close_conn:
        conn.close()
    return str(row[0] or "") if row else ""


def senha_reabertura_configurada() -> bool:
    return bool(obter_hash_senha_reabertura().strip())


def validar_senha_reabertura(senha: str, conn=None) -> bool:
    import hashlib

    hash_salvo = obter_hash_senha_reabertura(conn=conn)
    if not hash_salvo:
        return False
    hash_recebido = hashlib.sha256(str(senha or "").encode("utf-8")).hexdigest()
    return hash_recebido == hash_salvo


def listar_aglutinacoes() -> List[Dict[str, Any]]:
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        """
        SELECT id, codvend_origem, codvend_destino, ativo, COALESCE(descricao,'')
        FROM comissao_aglutinacao
        ORDER BY codvend_destino ASC, codvend_origem ASC
        """
    )
    rows = cur.fetchall()
    conn.close()
    return [
        {
            "id": int(r[0]),
            "codvend_origem": str(r[1] or ""),
            "codvend_destino": str(r[2] or ""),
            "ativo": int(r[3] or 0),
            "descricao": str(r[4] or ""),
        }
        for r in rows
    ]


def salvar_aglutinacao(dados: Dict[str, Any]) -> int:
    import sqlite3
    conn = get_conn()
    cur = conn.cursor()
    aid = int(dados.get("id", 0) or 0)
    origem = _normalizar_codvend(dados.get("codvend_origem"))
    destino = _normalizar_codvend(dados.get("codvend_destino"))
    if not origem or not destino:
        conn.close()
        raise ValueError("codvend_origem_e_destino_requeridos")
    if origem == destino:
        conn.close()
        raise ValueError("origem_igual_destino")
    ativo_raw = dados.get("ativo", 1)
    if ativo_raw is None or str(ativo_raw).strip() == "":
        ativo = 1
    else:
        ativo = int(ativo_raw)
    descricao = str(dados.get("descricao", "") or "").strip()
    try:
        if aid:
            cur.execute(
                "UPDATE comissao_aglutinacao SET codvend_origem=?, codvend_destino=?, ativo=?, descricao=? WHERE id=?",
                (origem, destino, ativo, descricao, aid),
            )
        else:
            cur.execute(
                "INSERT INTO comissao_aglutinacao(codvend_origem, codvend_destino, ativo, descricao) VALUES(?,?,?,?)",
                (origem, destino, ativo, descricao),
            )
            aid = int(cur.lastrowid or 0)
    except sqlite3.IntegrityError:
        conn.rollback()
        conn.close()
        raise ValueError("codvend_origem_ja_cadastrado")
    conn.commit()
    conn.close()
    return aid


def remover_aglutinacao(aid: int) -> None:
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("DELETE FROM comissao_aglutinacao WHERE id=?", (aid,))
    conn.commit()
    conn.close()
