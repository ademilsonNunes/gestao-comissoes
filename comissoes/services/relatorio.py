from datetime import datetime
from pathlib import Path
from typing import List, Dict, Any

from reportlab.lib.pagesizes import A4, A3, landscape
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader

from ..config import ARTFATOS_DIR, BASE_DIR


def _fmt_brl(v: float) -> str:
    s = f"{float(v or 0):,.2f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def _fmt_date(v: Any) -> str:
    if not v:
        return ""
    if isinstance(v, datetime):
        return v.strftime("%d/%m/%Y")
    s = str(v).strip()
    if not s:
        return ""
    for fmt in ("%Y-%m-%d", "%Y-%m-%d %H:%M:%S", "%d/%m/%Y"):
        try:
            return datetime.strptime(s, fmt).strftime("%d/%m/%Y")
        except Exception:
            pass
    return s[:10]


def _find_logo_path() -> Path | None:
    candidates = [
        BASE_DIR / "comissoes" / "static" / "Logo-Suprema1.png",
        BASE_DIR / "comissoes" / "static" / "logo-suprema.webp",
        BASE_DIR / "comissoes" / "assets" / "logo-suprema.webp",
        BASE_DIR / "comissoes" / "assets" / "logo.png",
        BASE_DIR / "comissoes" / "assets" / "logo.jpg",
        ARTFATOS_DIR / "logo.png",
        ARTFATOS_DIR / "logo.jpg",
    ]
    for p in candidates:
        if p.exists():
            return p
    return None


def _draw_header(c: canvas.Canvas, title: str) -> None:
    w, h = c._pagesize
    logo_path = _find_logo_path()
    if logo_path:
        logo = ImageReader(str(logo_path))
        c.drawImage(logo, 36, h - 72, width=130, height=36, preserveAspectRatio=True, mask="auto")
    c.setFont("Helvetica-Bold", 15)
    c.drawString(180, h - 50, title)
    c.setFont("Helvetica", 9)
    c.drawRightString(w - 36, h - 48, datetime.now().strftime("Gerado em %d/%m/%Y %H:%M"))
    c.line(36, h - 80, w - 36, h - 80)


def _mes_nome(mes: int) -> str:
    nomes = [
        "",
        "Janeiro",
        "Fevereiro",
        "Marco",
        "Abril",
        "Maio",
        "Junho",
        "Julho",
        "Agosto",
        "Setembro",
        "Outubro",
        "Novembro",
        "Dezembro",
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


def _agrupar_lancamentos_por_pedido(lancamentos: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    grupos: Dict[str, Dict[str, Any]] = {}
    ordem: List[str] = []
    for l in lancamentos:
        pedido_raw = str(l.get("pedido", "") or "").strip()
        key = pedido_raw if pedido_raw else f"__sem_pedido__{l.get('id', '')}"
        if key not in grupos:
            grupos[key] = {
                "emp": l.get("emp", "") or "",
                "pedido": pedido_raw,
                "vend": l.get("vend", "") or "",
                "dtemissao": l.get("dtemissao", "") or "",
                "vencto": l.get("vencto", "") or "",
                "dtbaixa": l.get("dtbaixa", "") or "",
                "rede": l.get("rede", "") or "",
                "uf": l.get("uf", "") or "",
                "cliente": l.get("cliente", "") or "",
                "vlrliq": 0.0,
                "tcomisprod": 0.0,
                "perc_min": None,
                "perc_max": None,
            }
            ordem.append(key)

        g = grupos[key]
        vlrliq_item = float(l.get("vlrliq", 0) or 0)
        tcom_item = float(l.get("tcomisprod", 0) or 0)
        g["vlrliq"] += vlrliq_item
        g["tcomisprod"] += tcom_item

        vend = float(l.get("comis_vend", 0) or 0)
        prod = float(l.get("comis_prod", 0) or 0)
        perc_item = prod if abs(prod) > 1e-12 else vend
        if g["perc_min"] is None or perc_item < float(g["perc_min"]):
            g["perc_min"] = perc_item
        if g["perc_max"] is None or perc_item > float(g["perc_max"]):
            g["perc_max"] = perc_item

        # Completa campos faltantes com o primeiro valor não vazio encontrado.
        for campo in ("emp", "vend", "dtemissao", "vencto", "dtbaixa", "rede", "uf", "cliente"):
            if not g.get(campo):
                g[campo] = l.get(campo, "") or ""

    return [grupos[k] for k in ordem]


def gerar_pdf_representante(codvend: str, dados: Dict[str, Any]) -> str:
    out_dir = ARTFATOS_DIR / "relatorios"
    out_dir.mkdir(parents=True, exist_ok=True)
    path = out_dir / f"{codvend}.pdf"

    c = canvas.Canvas(str(path), pagesize=landscape(A3))
    w, h = landscape(A3)

    mes = int(dados.get("mes", 0) or 0)
    ano = int(dados.get("ano", 0) or 0)
    mes_pag, ano_pag = _proximo_mes_ano(mes, ano)
    periodo_ref = f"{_mes_nome(mes)} de {ano}" if mes and ano else ""
    periodo_pag = f"{_mes_nome(mes_pag)} de {ano_pag}" if mes and ano else ""
    title = f"Relatorio de Comissao {periodo_pag} (Ref. {periodo_ref})".strip()
    _draw_header(c, title)

    rep_nome = str(dados.get("nome") or codvend or "").strip()
    cods_aglutinados = [str(x).strip() for x in (dados.get("codvends_aglutinados") or []) if str(x).strip()]
    cods_aglutinados = sorted(set(cods_aglutinados))
    aglutinado = len(cods_aglutinados) > 1
    c.setFont("Helvetica-Bold", 11)
    c.drawString(36, h - 102, f"Representante: {rep_nome}")
    c.setFont("Helvetica", 9)
    c.drawString(36, h - 118, f"Codigo: {codvend}")
    if aglutinado:
        c.drawString(200, h - 118, f"Codigos consolidados: {', '.join(cods_aglutinados)}")

    # Cabeçalho da tabela no padrão do PDF de referência.
    cols = [
        ("EMP", 36),
        ("PEDIDO", 86),
        ("VEND", 146),
        ("DTEMISSAO", 350),
        ("VENCTO", 430),
        ("DTBAIXA", 510),
        ("REDE", 590),
        ("UF", 665),
        ("CLIENTE", 695),
        ("VLRLIQ", 930),
        ("TCOMISPROD", 1020),
        ("%", 1120),
    ]

    y = h - 142
    c.setFont("Helvetica-Bold", 9)
    for label, x in cols:
        c.drawString(x, y, label)
    c.line(36, y - 4, w - 36, y - 4)
    y -= 18

    lancamentos = _agrupar_lancamentos_por_pedido(list(dados.get("lancamentos") or []))
    c.setFont("Helvetica", 8)
    total_liq = 0.0
    total_com = 0.0

    for l in lancamentos:
        if y < 68:
            c.showPage()
            _draw_header(c, title)
            y = h - 102
            c.setFont("Helvetica-Bold", 9)
            for label, x in cols:
                c.drawString(x, y, label)
            c.line(36, y - 4, w - 36, y - 4)
            y -= 18
            c.setFont("Helvetica", 8)

        vlrliq = float(l.get("vlrliq", 0) or 0)
        tcom = float(l.get("tcomisprod", 0) or 0)
        perc_calc = (tcom / vlrliq * 100.0) if abs(vlrliq) > 1e-12 else 0.0
        pmin = l.get("perc_min")
        pmax = l.get("perc_max")
        if pmin is None or pmax is None:
            perc_txt = f"{perc_calc:.2f}%"
        elif abs(float(pmin) - float(pmax)) < 1e-9:
            perc_txt = f"{float(pmin):.2f}%"
        else:
            perc_txt = f"{float(pmin):.2f}% a {float(pmax):.2f}%"

        total_liq += vlrliq
        total_com += tcom

        values = [
            str(l.get("emp", "") or ""),
            str(l.get("pedido", "") or ""),
            str(l.get("vend", "") or "")[:36],
            _fmt_date(l.get("dtemissao")),
            _fmt_date(l.get("vencto")),
            _fmt_date(l.get("dtbaixa")),
            str(l.get("rede", "") or "")[:14],
            str(l.get("uf", "") or "")[:2],
            str(l.get("cliente", "") or "")[:40],
            f"R$ {_fmt_brl(vlrliq)}",
            f"R$ {_fmt_brl(tcom)}",
            perc_txt,
        ]
        for (_, x), v in zip(cols, values):
            c.drawString(x, y, v)
        y -= 15

    if not lancamentos:
        total_liq = float(dados.get("total_vlrliq", 0) or 0)
        total_com = float(dados.get("total_comissao", 0) or 0)
        c.setFont("Helvetica", 10)
        c.drawString(36, y, "Sem lancamentos detalhados para este representante.")
        y -= 16

    desconto = float(dados.get("ajuste_desconto", 0) or 0)
    premio = float(dados.get("ajuste_premio", 0) or 0)
    total_final = float(dados.get("total_comissao_final", total_com - desconto + premio) or 0)
    ajuste_obs = str(dados.get("ajuste_obs", "") or "").strip()

    y -= 10
    c.setFont("Helvetica-Bold", 11)
    c.drawString(36, y, "Resumo financeiro do periodo")
    y -= 18
    c.setFont("Helvetica", 10)
    c.drawString(36, y, f"Faturamento liquido: R$ {_fmt_brl(total_liq)}")
    y -= 14
    c.drawString(36, y, f"Comissao base: R$ {_fmt_brl(total_com)}")
    if aglutinado:
        y -= 14
        c.drawString(36, y, f"Comissao aglutinada de {len(cods_aglutinados)} codigos: {', '.join(cods_aglutinados)}")
    y -= 14
    c.drawString(36, y, f"Descontos: R$ {_fmt_brl(desconto)}")
    y -= 14
    c.drawString(36, y, f"Premiacao: R$ {_fmt_brl(premio)}")
    y -= 16
    c.setFont("Helvetica-Bold", 11)
    c.drawString(36, y, f"Comissao final a pagar: R$ {_fmt_brl(total_final)}")
    if ajuste_obs:
        y -= 16
        c.setFont("Helvetica", 10)
        c.drawString(36, y, f"Observacao dos ajustes: {ajuste_obs[:140]}")
    y -= 18
    c.setFont("Helvetica-Bold", 10)
    c.drawString(36, y, "Favor enviar nota fiscal para darmos andamento ao pagamento.")

    c.showPage()
    c.save()
    return str(path)


def gerar_pdf_consolidado(mes: int, ano: int, comissoes: List[Dict[str, Any]]) -> str:
    out_dir = ARTFATOS_DIR / "relatorios"
    out_dir.mkdir(parents=True, exist_ok=True)
    path = out_dir / f"consolidado_{mes:02d}_{ano}.pdf"
    c = canvas.Canvas(str(path), pagesize=A4)
    mes_pag, ano_pag = _proximo_mes_ano(mes, ano)
    _draw_header(c, f"Consolidado {mes_pag:02d}/{ano_pag} (Ref. {mes:02d}/{ano})")
    c.setFont("Helvetica-Bold", 11)
    y = 700
    c.drawString(40, y, "CODVEND")
    c.drawString(140, y, "VlrLiq")
    c.drawString(240, y, "Comissao Base")
    c.drawString(350, y, "Desc")
    c.drawString(430, y, "Premio")
    c.drawString(520, y, "Comissao Final")
    c.drawString(660, y, "Status")
    y -= 20
    c.setFont("Helvetica", 10)
    total_liq = 0.0
    total_com = 0.0
    for r in comissoes:
        if y < 120:
            c.showPage()
            _draw_header(c, f"Consolidado {mes_pag:02d}/{ano_pag} (Ref. {mes:02d}/{ano})")
            c.setFont("Helvetica-Bold", 11)
            y = 700
            c.drawString(40, y, "CODVEND")
            c.drawString(140, y, "VlrLiq")
            c.drawString(240, y, "Comissao Base")
            c.drawString(350, y, "Desc")
            c.drawString(430, y, "Premio")
            c.drawString(520, y, "Comissao Final")
            c.drawString(660, y, "Status")
            y -= 20
            c.setFont("Helvetica", 10)
        total_liq += float(r.get("total_vlrliq", 0) or 0)
        total_base = float(r.get("total_comissao", 0) or 0)
        desconto = float(r.get("ajuste_desconto", 0) or 0)
        premio = float(r.get("ajuste_premio", 0) or 0)
        total_final = float(r.get("total_comissao_final", total_base - desconto + premio) or 0)
        total_com += total_final
        c.drawString(40, y, str(r.get("codvend", "")))
        c.drawString(140, y, f"R$ {_fmt_brl(r.get('total_vlrliq', 0))}")
        c.drawString(240, y, f"R$ {_fmt_brl(total_base)}")
        c.drawString(350, y, f"R$ {_fmt_brl(desconto)}")
        c.drawString(430, y, f"R$ {_fmt_brl(premio)}")
        c.drawString(520, y, f"R$ {_fmt_brl(total_final)}")
        c.drawString(660, y, str(r.get("status", "")))
        y -= 18
    y -= 10
    c.setFont("Helvetica-Bold", 11)
    c.drawString(40, y, "Totais")
    c.drawString(140, y, f"R$ {_fmt_brl(total_liq)}")
    c.drawString(520, y, f"R$ {_fmt_brl(total_com)}")
    c.showPage()
    c.save()
    return str(path)
