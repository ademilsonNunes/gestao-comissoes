import os
import time
from urllib.parse import urlparse
from datetime import timedelta

from flask import Flask, jsonify, request, send_file, render_template, session, redirect, url_for
from .database.database import init_schema
from .database import models
from .services.importacao import (
    importar_representantes_base,
    importar_vendas,
    importar_vendas_arquivo,
    importar_vendas_query_banco,
    testar_conexao_sql,
)
from .services.calculo import calcular, consolidado
from .services.relatorio import gerar_pdf_representante, gerar_pdf_consolidado
from .services.email_service import enviar_email, enviar_email_cfg


app = Flask(__name__)
app.secret_key = os.getenv("APP_SECRET_KEY", "gestao-comissoes-local-secret-change-me")
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
app.config["PERMANENT_SESSION_LIFETIME"] = timedelta(days=1)
init_schema()

SESSION_IDLE_MINUTES = int(os.getenv("SESSION_IDLE_MINUTES", "30"))
LOGIN_MAX_ATTEMPTS = int(os.getenv("LOGIN_MAX_ATTEMPTS", "5"))
LOGIN_WINDOW_MINUTES = int(os.getenv("LOGIN_WINDOW_MINUTES", "15"))
LOGIN_LOCKOUT_MINUTES = int(os.getenv("LOGIN_LOCKOUT_MINUTES", "15"))
_LOGIN_ATTEMPTS: dict[str, dict[str, list[float] | float]] = {}


def _senha_mascarada(v: str) -> bool:
    s = str(v or "").strip()
    return bool(s) and (set(s) <= set("*•●")) and len(s) >= 4


@app.after_request
def _charset_utf8(response):
    ctype = str(response.headers.get("Content-Type", "")).lower()
    if ctype.startswith("text/html") and "charset=" not in ctype:
        response.headers["Content-Type"] = "text/html; charset=utf-8"
    elif (
        ctype.startswith("application/javascript")
        or ctype.startswith("text/javascript")
    ) and "charset=" not in ctype:
        response.headers["Content-Type"] = "text/javascript; charset=utf-8"
    elif ctype.startswith("text/css") and "charset=" not in ctype:
        response.headers["Content-Type"] = "text/css; charset=utf-8"
    return response


def _is_safe_next_path(next_path: str) -> bool:
    if not next_path:
        return False
    parsed = urlparse(next_path)
    return parsed.scheme == "" and parsed.netloc == "" and next_path.startswith("/")


def _client_key() -> str:
    fwd = str(request.headers.get("X-Forwarded-For", "") or "").strip()
    if fwd:
        return fwd.split(",")[0].strip()
    return str(request.remote_addr or "unknown")


def _prune_attempts(item: dict[str, list[float] | float], now: float) -> None:
    window_seconds = LOGIN_WINDOW_MINUTES * 60
    fails = [ts for ts in list(item.get("fails", [])) if now - float(ts) <= window_seconds]
    item["fails"] = fails
    locked_until = float(item.get("locked_until", 0) or 0)
    if locked_until and now >= locked_until:
        item["locked_until"] = 0


def _is_login_locked(key: str) -> tuple[bool, int]:
    now = time.time()
    item = _LOGIN_ATTEMPTS.get(key, {"fails": [], "locked_until": 0})
    _prune_attempts(item, now)
    _LOGIN_ATTEMPTS[key] = item
    locked_until = float(item.get("locked_until", 0) or 0)
    if locked_until > now:
        remaining = max(1, int((locked_until - now) // 60) + 1)
        return True, remaining
    return False, 0


def _register_login_failure(key: str) -> None:
    now = time.time()
    item = _LOGIN_ATTEMPTS.get(key, {"fails": [], "locked_until": 0})
    _prune_attempts(item, now)
    fails = list(item.get("fails", []))
    fails.append(now)
    item["fails"] = fails
    if len(fails) >= LOGIN_MAX_ATTEMPTS:
        item["locked_until"] = now + (LOGIN_LOCKOUT_MINUTES * 60)
        item["fails"] = []
    _LOGIN_ATTEMPTS[key] = item


def _clear_login_failures(key: str) -> None:
    if key in _LOGIN_ATTEMPTS:
        del _LOGIN_ATTEMPTS[key]


@app.before_request
def _enforce_authentication():
    endpoint = request.endpoint or ""
    if endpoint in {"static", "login_page", "login_submit", "logout"}:
        return None
    if request.path.startswith("/static/"):
        return None
    if session.get("auth_ok"):
        now = time.time()
        last_seen = float(session.get("last_seen_ts", 0) or 0)
        if last_seen and (now - last_seen) > (SESSION_IDLE_MINUTES * 60):
            session.clear()
            if request.path.startswith("/api/"):
                return jsonify({"error": "sessao_expirada"}), 401
            return redirect(url_for("login_page", next=request.path))
        session["last_seen_ts"] = now
        return None
    if request.path.startswith("/api/"):
        return jsonify({"error": "nao_autenticado"}), 401
    return redirect(url_for("login_page", next=request.path))


def _fmt_brl(v: float) -> str:
    s = f"{float(v or 0):,.2f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def _mes_nome(mes: int) -> str:
    nomes = [
        "",
        "janeiro",
        "fevereiro",
        "marco",
        "abril",
        "maio",
        "junho",
        "julho",
        "agosto",
        "setembro",
        "outubro",
        "novembro",
        "dezembro",
    ]
    return nomes[mes] if 1 <= int(mes or 0) <= 12 else ""


def _proximo_mes_ano(mes: int, ano: int) -> tuple[int, int]:
    mes = int(mes or 0)
    ano = int(ano or 0)
    if mes <= 0 or ano <= 0:
        return mes, ano
    if mes == 12:
        return 1, ano + 1
    return mes + 1, ano


def _montar_corpo_email(rep: dict, dados: dict) -> str:
    nome = str(rep.get("nome", "") or rep.get("codvend", "") or "").strip() or "Representante"
    mes = int(dados.get("mes", 0) or 0)
    ano = int(dados.get("ano", 0) or 0)
    mes_pag, ano_pag = _proximo_mes_ano(mes, ano)
    ref_txt = f"{mes:02d}/{ano}"
    pag_txt = f"{mes_pag:02d}/{ano_pag}" if mes and ano else ""
    ref_nome = _mes_nome(mes)
    pag_nome = _mes_nome(mes_pag)
    total_liq = float(dados.get("total_vlrliq", 0) or 0)
    total_comissao = float(dados.get("total_comissao", 0) or 0)
    desconto = float(dados.get("ajuste_desconto", 0) or 0)
    premio = float(dados.get("ajuste_premio", 0) or 0)
    total_final = float(dados.get("total_comissao_final", total_comissao - desconto + premio) or 0)
    ajuste_obs = str(dados.get("ajuste_obs", "") or "").strip()
    extra = str(rep.get("corpo_email", "") or "").strip()

    linhas = [
        f"Prezado(a) {nome},",
        "",
        f"Segue o fechamento da comissao de {pag_txt} referente ao periodo {ref_txt}.",
        f"Competencia: {pag_nome} de {ano_pag} | Referencia: {ref_nome} de {ano}",
        "",
        "Resumo financeiro:",
        f"- Faturamento liquido: R$ {_fmt_brl(total_liq)}",
        f"- Comissao base: R$ {_fmt_brl(total_comissao)}",
        f"- Descontos: R$ {_fmt_brl(desconto)}",
        f"- Premiacao: R$ {_fmt_brl(premio)}",
        f"- Total de comissao a pagar: R$ {_fmt_brl(total_final)}",
    ]
    if ajuste_obs:
        linhas.extend(["", f"Observacao dos ajustes: {ajuste_obs}"])
    linhas.extend(
        [
            "",
            "O relatorio detalhado da comissao segue em anexo (PDF).",
            "",
            "Favor enviar nota fiscal para que possamos dar andamento no pagamento.",
            "",
            "Atenciosamente,",
            "Equipe Financeira",
        ]
    )
    if extra:
        linhas.extend(["", "Informacoes adicionais:", extra])
    return "\n".join(linhas)


def _montar_dados_comissao(rep: dict, com: dict, lancamentos: list[dict]) -> dict:
    if lancamentos:
        total_liq = float(sum(float(l.get("vlrliq", 0) or 0) for l in lancamentos))
        total_com = float(sum(float(l.get("tcomisprod", 0) or 0) for l in lancamentos))
    else:
        total_liq = float(com.get("total_vlrliq", 0) or 0)
        total_com = float(com.get("total_comissao", 0) or 0)
    desconto = float(com.get("ajuste_desconto", 0) or 0)
    premio = float(com.get("ajuste_premio", 0) or 0)
    total_final = float(total_com - desconto + premio)
    percent = (total_final / total_liq * 100.0) if total_liq else 0.0
    codvends_aglutinados = sorted({str(l.get("codvend", "") or "").strip() for l in (lancamentos or []) if str(l.get("codvend", "") or "").strip()})
    if not codvends_aglutinados and com.get("codvend"):
        codvends_aglutinados = [str(com.get("codvend"))]
    return {
        "nome": rep.get("nome", ""),
        "mes": com.get("mes"),
        "ano": com.get("ano"),
        "total_vlrliq": total_liq,
        "percent": percent,
        "total_comissao": total_com,
        "ajuste_desconto": desconto,
        "ajuste_premio": premio,
        "ajuste_obs": com.get("ajuste_obs", ""),
        "total_comissao_final": total_final,
        "lancamentos": lancamentos,
        "codvends_aglutinados": codvends_aglutinados,
    }


@app.get("/login")
def login_page():
    init_schema()
    if session.get("auth_ok"):
        return redirect(url_for("index_page"))
    next_path = request.args.get("next", "/")
    if not _is_safe_next_path(next_path):
        next_path = "/"
    return render_template(
        "login.html",
        error="",
        setup_mode=not models.senha_reabertura_configurada(),
        next_path=next_path,
    )


@app.post("/login")
def login_submit():
    init_schema()
    senha = str(request.form.get("senha", "") or "")
    senha_confirm = str(request.form.get("senha_confirm", "") or "")
    next_path = request.form.get("next", "/")
    if not _is_safe_next_path(next_path):
        next_path = "/"
    key = _client_key()
    locked, remaining = _is_login_locked(key)
    if locked:
        return render_template("login.html", error=f"Acesso temporariamente bloqueado. Tente novamente em {remaining} minuto(s).", setup_mode=False, next_path=next_path), 429

    setup_mode = not models.senha_reabertura_configurada()
    if setup_mode:
        if len(senha) < 6:
            return render_template("login.html", error="Defina uma senha com pelo menos 6 caracteres.", setup_mode=True, next_path=next_path), 400
        if senha != senha_confirm:
            return render_template("login.html", error="Confirmação de senha não confere.", setup_mode=True, next_path=next_path), 400
        cfg = models.obter_configuracoes()
        cfg["reabrir_senha"] = senha
        cfg["limpar_reabrir_senha"] = 0
        models.salvar_configuracoes(cfg)
        _clear_login_failures(key)
        session.permanent = True
        session["auth_ok"] = True
        session["last_seen_ts"] = time.time()
        return redirect(next_path)

    if not models.validar_senha_reabertura(senha):
        _register_login_failure(key)
        return render_template("login.html", error="Senha inválida.", setup_mode=False, next_path=next_path), 401
    _clear_login_failures(key)
    session.permanent = True
    session["auth_ok"] = True
    session["last_seen_ts"] = time.time()
    return redirect(next_path)


@app.post("/logout")
def logout():
    session.clear()
    return redirect(url_for("login_page"))


@app.get("/")
def index_page():
    init_schema()
    return render_template("index.html")

@app.get("/importacao")
def importacao_page():
    init_schema()
    return render_template("importacao.html")

@app.get("/representantes")
def representantes_page():
    init_schema()
    return render_template("representantes.html")

@app.get("/regras")
def regras_page():
    init_schema()
    return render_template("regras.html")

@app.get("/apuracao")
def apuracao_page():
    init_schema()
    return render_template("apuracao.html")

@app.get("/envio")
def envio_page():
    init_schema()
    return render_template("envio.html")

@app.get("/configuracoes")
def configuracoes_page():
    init_schema()
    return render_template("configuracoes.html")

@app.get("/api/representantes")
def listar_representantes():
    init_schema()
    return jsonify(models.listar_representantes())


@app.post("/api/representantes")
def criar_representante():
    payload = request.get_json(silent=True) or {}
    models.upsert_representante(payload.get("codvend",""), payload.get("nome",""), payload.get("email",""), payload.get("corpo_email",""))
    return jsonify({"status": "ok"}), 201


@app.put("/api/representantes/<int:rep_id>")
def atualizar_representante(rep_id: int):
    payload = request.get_json(silent=True) or {}
    models.atualizar_representante(rep_id, payload)
    return jsonify({"status": "ok"})


@app.delete("/api/representantes/<int:rep_id>")
def desativar_representante(rep_id: int):
    models.desativar_representante(rep_id)
    return jsonify({"status": "ok"}), 204


@app.post("/api/representantes/importar")
def importar_representantes():
    n = importar_representantes_base()
    return jsonify({"importados": n})


@app.post("/api/importacao/upload")
def upload_importacao():
    f = request.files.get("arquivo")
    mes = request.form.get("mes")
    ano = request.form.get("ano")
    if not f:
        return jsonify({"error":"arquivo_requerido"}), 400
    import time
    from pathlib import Path
    from .config import ARTFATOS_DIR
    updir = ARTFATOS_DIR / "uploads"
    updir.mkdir(parents=True, exist_ok=True)
    suffix = Path(f.filename).suffix.lower()
    if suffix not in (".xlsx",".xlsm"):
        return jsonify({"error":"formato_invalido"}), 400
    p = updir / f"{int(time.time())}_{f.filename}"
    f.save(str(p))
    mes_i = None
    ano_i = None
    try:
        mes_i = int(mes) if mes else None
    except Exception:
        mes_i = None
    try:
        ano_i = int(ano) if ano else None
    except Exception:
        ano_i = None
    try:
        res = importar_vendas_arquivo(p, mes_override=mes_i, ano_override=ano_i)
    except ValueError as e:
        return jsonify({"error": str(e)}), 400
    return jsonify(res)


@app.post("/api/importacao/padrao")
def importar_arquivo_padrao():
    return jsonify(importar_vendas())


@app.post("/api/importacao/query")
def importar_via_query():
    payload = request.get_json(silent=True) or {}
    try:
        mes = int(payload.get("mes", 0) or 0)
        ano = int(payload.get("ano", 0) or 0)
    except Exception:
        return jsonify({"error": "mes_ano_invalidos"}), 400
    conn_str = str(payload.get("conn_str", "") or "").strip() or None
    incluir_devolucoes_raw = payload.get("incluir_devolucoes", None)
    incluir_devolucoes = None if incluir_devolucoes_raw is None else bool(int(incluir_devolucoes_raw))
    try:
        res = importar_vendas_query_banco(
            mes=mes,
            ano=ano,
            conn_str=conn_str,
            incluir_devolucoes=incluir_devolucoes,
        )
    except ValueError as e:
        return jsonify({"error": str(e)}), 400
    except RuntimeError as e:
        return jsonify({"error": str(e)}), 500
    except Exception as e:
        return jsonify({"error": f"falha_importacao_query: {e}"}), 500
    return jsonify(res)


@app.get("/api/importacao/<int:imp_id>/status")
def status_importacao(imp_id: int):
    return jsonify({"status": "ok", "id": imp_id})


@app.get("/api/importacao/<int:imp_id>/pendencias")
def pendencias_importacao(imp_id: int):
    return jsonify({"pendentes": models.representantes_faltantes(), "id": imp_id})


@app.get("/api/auditoria/cadastro")
def auditoria_cadastro():
    init_schema()
    return jsonify(models.resumo_auditoria())


@app.get("/api/regras")
def listar_regras():
    init_schema()
    codvend = request.args.get("codvend", "")
    return jsonify(models.listar_regras(codvend))


@app.post("/api/regras")
def criar_regra():
    init_schema()
    payload = request.get_json(silent=True) or {}
    rid = models.salvar_regra(payload)
    return jsonify({"status": "ok", "id": rid}), 201


@app.put("/api/regras/<int:rid>")
def atualizar_regra(rid: int):
    init_schema()
    payload = request.get_json(silent=True) or {}
    payload["id"] = rid
    models.salvar_regra(payload)
    return jsonify({"status": "ok", "id": rid})


@app.delete("/api/regras/<int:rid>")
def excluir_regra(rid: int):
    init_schema()
    models.remover_regra(rid)
    return jsonify({"status": "ok"}), 204


@app.get("/api/aglutinacoes")
def listar_aglutinacoes():
    init_schema()
    return jsonify(models.listar_aglutinacoes())


@app.post("/api/aglutinacoes")
def criar_aglutinacao():
    init_schema()
    payload = request.get_json(silent=True) or {}
    try:
        aid = models.salvar_aglutinacao(payload)
    except ValueError as e:
        return jsonify({"error": str(e)}), 400
    return jsonify({"status": "ok", "id": aid}), 201


@app.put("/api/aglutinacoes/<int:aid>")
def atualizar_aglutinacao(aid: int):
    init_schema()
    payload = request.get_json(silent=True) or {}
    payload["id"] = aid
    try:
        models.salvar_aglutinacao(payload)
    except ValueError as e:
        return jsonify({"error": str(e)}), 400
    return jsonify({"status": "ok", "id": aid})


@app.delete("/api/aglutinacoes/<int:aid>")
def excluir_aglutinacao(aid: int):
    init_schema()
    models.remover_aglutinacao(aid)
    return jsonify({"status": "ok"}), 204


@app.post("/api/comissoes/calcular")
def calcular_comissoes():
    payload = request.get_json(silent=True) or {}
    mes = int(payload.get("mes",0) or 0)
    ano = int(payload.get("ano",0) or 0)
    return jsonify(calcular(mes, ano))


@app.get("/api/comissoes/<int:mes>/<int:ano>")
def consolidado_periodo(mes: int, ano: int):
    return jsonify(consolidado(mes, ano))


@app.get("/api/comissoes/periodos")
def periodos_comissoes():
    init_schema()
    return jsonify(models.listar_periodos_comissoes())


@app.get("/api/lancamentos/periodos")
def periodos_lancamentos():
    init_schema()
    return jsonify(models.listar_periodos_lancamentos())


@app.get("/api/apuracoes/<int:mes>/<int:ano>")
def obter_apuracao_periodo(mes: int, ano: int):
    init_schema()
    return jsonify(models.obter_apuracao_por_periodo(mes, ano))


@app.delete("/api/apuracoes/<int:mes>/<int:ano>")
def excluir_apuracao_periodo(mes: int, ano: int):
    init_schema()
    res = models.excluir_apuracao_periodo(mes, ano)
    if not res.get("deleted"):
        return jsonify({"error": "apuracao_nao_encontrada", "mes": mes, "ano": ano}), 404
    return jsonify(res)


@app.get("/api/comissoes/<int:cid>/lancamentos")
def lancamentos_comissao(cid: int):
    return jsonify(models.obter_lancamentos_por_comissao(cid))


@app.put("/api/lancamentos/<int:lid>")
def ajustar_lancamento(lid: int):
    status = models.status_comissao_por_lancamento(lid)
    if status in {"aprovado", "enviado"}:
        return jsonify({"error": "comissao_bloqueada", "status": status}), 409
    payload = request.get_json(silent=True) or {}
    campo = payload.get("campo", "")
    valor = str(payload.get("valor", ""))
    motivo = payload.get("motivo", "")
    models.registrar_ajuste(lid, campo, valor, motivo)
    return jsonify({"status": "ok", "id": lid})


@app.put("/api/lancamentos/<int:lid>/percentuais")
def ajustar_percentuais_lancamento(lid: int):
    status = models.status_comissao_por_lancamento(lid)
    if status in {"aprovado", "enviado"}:
        return jsonify({"error": "comissao_bloqueada", "status": status}), 409
    payload = request.get_json(silent=True) or {}
    motivo = str(payload.get("motivo", "") or "").strip()
    vend_raw = payload.get("comis_vend", None)
    prod_raw = payload.get("comis_prod", None)
    comis_vend = None
    comis_prod = None
    try:
        if vend_raw is not None and str(vend_raw).strip() != "":
            comis_vend = float(vend_raw)
    except Exception:
        return jsonify({"error": "comis_vend_invalido"}), 400
    try:
        if prod_raw is not None and str(prod_raw).strip() != "":
            comis_prod = float(prod_raw)
    except Exception:
        return jsonify({"error": "comis_prod_invalido"}), 400

    res = models.atualizar_percentuais_lancamento(lid, comis_vend, comis_prod, motivo)
    if not res:
        return jsonify({"error": "lancamento_nao_encontrado"}), 404
    return jsonify({"status": "ok", "lancamento": res})


@app.post("/api/comissoes/<int:cid>/aprovar")
def aprovar_comissao(cid: int):
    models.aprovar_comissao(cid)
    return jsonify({"status": "ok", "id": cid})


@app.post("/api/comissoes/<int:cid>/cancelar-aprovacao")
def cancelar_aprovacao_comissao(cid: int):
    ok = models.cancelar_aprovacao_comissao(cid)
    if not ok:
        return jsonify({"error": "nao_permitido"}), 409
    return jsonify({"status": "ok", "id": cid})


@app.post("/api/comissoes/<int:cid>/reabrir")
def reabrir_comissao(cid: int):
    payload = request.get_json(silent=True) or {}
    senha = str(payload.get("senha", "") or "")
    if not senha:
        return jsonify({"error": "senha_requerida"}), 400
    ok = models.reabrir_comissao_enviada(cid, senha)
    if not ok:
        return jsonify({"error": "nao_permitido_ou_senha_invalida"}), 409
    return jsonify({"status": "ok", "id": cid})


@app.put("/api/comissoes/<int:cid>/ajustes-financeiros")
def atualizar_ajustes_financeiros_comissao(cid: int):
    payload = request.get_json(silent=True) or {}
    com = models.obter_comissao_por_id(cid)
    if not com:
        return jsonify({"error": "not_found"}), 404
    if str(com.get("status", "")).lower() == "enviado":
        return jsonify({"error": "comissao_enviada_bloqueada"}), 409
    try:
        desconto = float(payload.get("desconto", 0) or 0)
        premio = float(payload.get("premio", 0) or 0)
    except Exception:
        return jsonify({"error": "valores_invalidos"}), 400
    obs = str(payload.get("observacao", "") or "")
    ok = models.atualizar_ajustes_financeiros_comissao(cid, desconto, premio, obs)
    if not ok:
        return jsonify({"error": "not_found"}), 404
    com_upd = models.obter_comissao_por_id(cid)
    return jsonify({"status": "ok", "comissao": com_upd})


@app.post("/api/comissoes/<int:cid>/enviar-email")
def enviar_email_por_comissao(cid: int):
    payload = request.get_json(silent=True) or {}
    com = models.obter_comissao_por_id(cid)
    if not com:
        return jsonify({"error": "not_found"}), 404
    status = str(com.get("status", "")).lower()
    if status != "aprovado":
        return jsonify({"error": "comissao_nao_aprovada", "status": status}), 409
    rep = models.obter_representante_por_codvend(str(com.get("codvend", "")))
    if not rep:
        return jsonify({"error": "representante_nao_encontrado"}), 404
    if not str(rep.get("email", "")).strip():
        return jsonify({"error": "email_representante_ausente"}), 400
    cfg = models.obter_configuracoes()
    lancamentos = models.obter_lancamentos_por_comissao(cid)
    dados = _montar_dados_comissao(rep, com, lancamentos)
    anexos = [gerar_pdf_representante(rep["codvend"], dados)]
    corpo = _montar_corpo_email(rep, dados)
    ok = enviar_email_cfg(cfg, rep["email"], payload.get("assunto", "Relatorio de Comissao"), corpo, anexos)
    models.registrar_email_envio(int(rep["id"]), rep["email"], "ok" if ok else "erro", "unitario")
    if ok:
        models.marcar_comissao_enviada(cid)
    return jsonify({"status": "ok" if ok else "erro", "id": cid})


@app.get("/api/comissoes/<int:cid>/pdf")
def gerar_pdf_por_comissao(cid: int):
    com = models.obter_comissao_por_id(cid)
    if not com:
        return jsonify({"error": "not_found"}), 404
    rep = models.obter_representante_por_codvend(str(com.get("codvend", "")))
    rep_nome = rep.get("nome", "") if rep else ""
    lancamentos = models.obter_lancamentos_por_comissao(cid)
    dados = _montar_dados_comissao({"nome": rep_nome}, com, lancamentos)
    path = gerar_pdf_representante(str(com.get("codvend", "")), dados)
    try:
        return send_file(path, mimetype="application/pdf")
    except Exception:
        return jsonify({"pdf": path})


@app.get("/api/relatorios/<int:rid>/pdf")
def gerar_pdf(rid: int):
    rep = models.obter_representante(rid)
    if not rep:
        return jsonify({"error":"not_found"}), 404
    com = models.obter_ultima_comissao_por_codvend(rep["codvend"])
    lancamentos = models.obter_lancamentos_por_comissao(int(com.get("id", 0) or 0)) if com.get("id") else []
    dados = _montar_dados_comissao(rep, com, lancamentos)
    path = gerar_pdf_representante(rep["codvend"], dados)
    try:
        return send_file(path, mimetype="application/pdf")
    except Exception:
        return jsonify({"pdf": path})

@app.get("/api/relatorios/consolidado/<int:mes>/<int:ano>/csv")
def consolidado_csv(mes: int, ano: int):
    import csv
    from pathlib import Path
    from .config import ARTFATOS_DIR
    coms = models.obter_comissao_periodo(mes, ano)
    out_dir = ARTFATOS_DIR / "relatorios"
    out_dir.mkdir(parents=True, exist_ok=True)
    path = out_dir / f"consolidado_{mes:02d}_{ano}.csv"
    with path.open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["codvend","total_vlrliq","total_comissao_base","ajuste_desconto","ajuste_premio","total_comissao_final","status"])
        for r in coms:
            w.writerow([
                r.get("codvend",""),
                r.get("total_vlrliq",0),
                r.get("total_comissao",0),
                r.get("ajuste_desconto",0),
                r.get("ajuste_premio",0),
                r.get("total_comissao_final",0),
                r.get("status",""),
            ])
    try:
        return send_file(str(path), mimetype="text/csv")
    except Exception:
        return jsonify({"csv": str(path)})

@app.get("/api/relatorios/consolidado/<int:mes>/<int:ano>/pdf")
def consolidado_pdf(mes: int, ano: int):
    coms = models.obter_comissao_periodo(mes, ano)
    path = gerar_pdf_consolidado(mes, ano, coms)
    try:
        return send_file(path, mimetype="application/pdf")
    except Exception:
        return jsonify({"pdf": path})

@app.post("/api/email/enviar/<int:rid>")
def enviar_email_unitario(rid: int):
    payload = request.get_json(silent=True) or {}
    rep = models.obter_representante(rid)
    if not rep:
        return jsonify({"error":"not_found"}), 404
    if not str(rep.get("email", "")).strip():
        return jsonify({"error": "email_representante_ausente"}), 400
    cfg = models.obter_configuracoes()
    anexos = payload.get("anexos", []) or []
    com = None
    if not anexos:
        com = models.obter_ultima_comissao_por_codvend(rep["codvend"])
        if not com:
            return jsonify({"error": "comissao_nao_encontrada"}), 404
        status = str(com.get("status", "")).lower()
        if status != "aprovado":
            return jsonify({"error": "comissao_nao_aprovada", "status": status}), 409
        lancamentos = models.obter_lancamentos_por_comissao(int(com.get("id", 0) or 0)) if com.get("id") else []
        dados = _montar_dados_comissao(rep, com, lancamentos)
        anexos = [gerar_pdf_representante(rep["codvend"], dados)]
    corpo = _montar_corpo_email(rep, dados if com else {})
    ok = enviar_email_cfg(cfg, rep["email"], payload.get("assunto",""), corpo, anexos)
    models.registrar_email_envio(rid, rep["email"], "ok" if ok else "erro", "unitario")
    if ok and com and com.get("id"):
        models.marcar_comissao_enviada(int(com.get("id")))
    return jsonify({"status": "ok" if ok else "erro"})


@app.post("/api/email/lote")
def enviar_email_lote():
    payload = request.get_json(silent=True) or {}
    ids = payload.get("ids", [])
    if not isinstance(ids, list) or not ids:
        return jsonify({"error": "ids_requeridos"}), 400
    resultados = []
    for rid in ids:
        rep = models.obter_representante(int(rid))
        if not rep:
            resultados.append({"id": rid, "status": "erro", "motivo": "not_found"})
            continue
        if not str(rep.get("email", "")).strip():
            resultados.append({"id": rid, "status": "erro", "motivo": "email_representante_ausente"})
            continue
        cfg = models.obter_configuracoes()
        anexos = payload.get("anexos", []) or []
        com = None
        if not anexos:
            com = models.obter_ultima_comissao_por_codvend(rep["codvend"])
            if not com:
                resultados.append({"id": rid, "status": "erro", "motivo": "comissao_nao_encontrada"})
                continue
            status = str(com.get("status", "")).lower()
            if status != "aprovado":
                resultados.append({"id": rid, "status": "erro", "motivo": "comissao_nao_aprovada"})
                continue
            lancamentos = models.obter_lancamentos_por_comissao(int(com.get("id", 0) or 0)) if com.get("id") else []
            dados = _montar_dados_comissao(rep, com, lancamentos)
            anexos = [gerar_pdf_representante(rep["codvend"], dados)]
        corpo = _montar_corpo_email(rep, dados if com else {})
        ok = enviar_email_cfg(cfg, rep["email"], payload.get("assunto",""), corpo, anexos)
        models.registrar_email_envio(int(rid), rep["email"], "ok" if ok else "erro", "lote")
        if ok and com and com.get("id"):
            models.marcar_comissao_enviada(int(com.get("id")))
        resultados.append({"id": rid, "status": "ok" if ok else "erro"})
    return jsonify({"resultados": resultados})


@app.get("/api/email/historico")
def historico_email():
    return jsonify(models.listar_historico_email())


@app.get("/api/configuracoes")
def obter_configuracoes():
    init_schema()
    cfg = models.obter_configuracoes()
    cfg_pub = dict(cfg)
    cfg_pub["has_smtp_pass"] = bool(str(cfg.get("smtp_pass", "") or "").strip())
    cfg_pub["has_sql_pass"] = bool(str(cfg.get("sql_pass", "") or "").strip())
    cfg_pub["smtp_pass"] = ""
    cfg_pub["sql_pass"] = ""
    return jsonify(cfg_pub)


@app.put("/api/configuracoes")
def salvar_configuracoes():
    payload = request.get_json(silent=True) or {}
    init_schema()
    models.salvar_configuracoes(payload)
    return jsonify({"status": "ok"})


@app.post("/api/configuracoes/testar-smtp")
def testar_smtp():
    payload = request.get_json(silent=True) or {}
    cfg = models.obter_configuracoes()
    host = payload.get("smtp_host", cfg.get("smtp_host", ""))
    port = int(payload.get("smtp_port", cfg.get("smtp_port", 587)))
    user = payload.get("smtp_user", cfg.get("smtp_user", ""))
    passwd = str(payload.get("smtp_pass", "") or "").strip()
    if (not passwd) or _senha_mascarada(passwd):
        passwd = cfg.get("smtp_pass", "")
    from_addr = payload.get("smtp_from", cfg.get("smtp_from", ""))
    from comissoes.services.email_service import smtplib
    def _smtp_try(auth_pass: str) -> None:
        with smtplib.SMTP(host, port, timeout=20) as s:
            s.ehlo()
            if s.has_extn("starttls"):
                s.starttls()
                s.ehlo()
            if user:
                s.login(user, auth_pass)

    try:
        _smtp_try(passwd)
        return jsonify({"status": "ok"})
    except Exception as e:
        # Fallback: se senha enviada falhar, tenta a senha salva (evita quebra por autofill incorreto no front).
        saved_pass = str(cfg.get("smtp_pass", "") or "")
        if saved_pass and saved_pass != passwd:
            try:
                _smtp_try(saved_pass)
                return jsonify({"status": "ok", "fallback": "smtp_pass_salva"})
            except Exception:
                pass
        return jsonify({"status": "erro", "motivo": str(e)})


@app.post("/api/configuracoes/testar-sql")
def testar_sql():
    payload = request.get_json(silent=True) or {}
    conn_str = str(payload.get("conn_str", "") or "").strip() or None
    try:
        cfg = models.obter_configuracoes()
        if payload:
            for k, v in payload.items():
                if k in {"sql_pass", "smtp_pass"} and (
                    (not str(v or "").strip()) or _senha_mascarada(str(v or ""))
                ):
                    continue
                cfg[k] = v
        res = testar_conexao_sql(conn_str=conn_str, cfg=cfg)
        return jsonify(res)
    except ValueError as e:
        return jsonify({"status": "erro", "motivo": str(e)}), 400
    except RuntimeError as e:
        return jsonify({"status": "erro", "motivo": str(e)}), 500
    except Exception as e:
        return jsonify({"status": "erro", "motivo": str(e)}), 500


@app.post("/api/configuracoes/testar_sql")
def testar_sql_alias():
    # Compatibilidade com versões antigas do front-end.
    return testar_sql()


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
