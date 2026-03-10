import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from pathlib import Path
from typing import List
from ..config import SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASS, SMTP_FROM

SMTP_CC_FIXO = "ti@sobelsuprema.com.br"


def _build_recipients(destinatario: str) -> List[str]:
    to_addr = str(destinatario or "").strip()
    if not to_addr:
        return []
    cc_addr = str(SMTP_CC_FIXO or "").strip()
    if cc_addr and cc_addr.lower() != to_addr.lower():
        return [to_addr, cc_addr]
    return [to_addr]


def enviar_email(destinatario: str, assunto: str, corpo: str, anexos: List[str]) -> bool:
    if not SMTP_HOST or not SMTP_USER or not SMTP_FROM:
        return False
    recipients = _build_recipients(destinatario)
    if not recipients:
        return False
    msg = MIMEMultipart()
    msg["From"] = SMTP_FROM
    msg["To"] = recipients[0]
    if len(recipients) > 1:
        msg["Cc"] = ", ".join(recipients[1:])
    msg["Subject"] = assunto
    msg.attach(MIMEText(corpo, "plain"))
    for a in anexos:
        p = Path(a)
        if not p.exists():
            continue
        part = MIMEBase("application", "octet-stream")
        part.set_payload(p.read_bytes())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename={p.name}")
        msg.attach(part)
    try:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as s:
            s.starttls()
            s.login(SMTP_USER, SMTP_PASS)
            s.sendmail(SMTP_FROM, recipients, msg.as_string())
        return True
    except Exception:
        return False


def enviar_email_cfg(cfg: dict, destinatario: str, assunto: str, corpo: str, anexos: List[str]) -> bool:
    host = cfg.get("smtp_host", "")
    port = int(cfg.get("smtp_port", 587))
    user = cfg.get("smtp_user", "")
    passwd = cfg.get("smtp_pass", "")
    from_addr = cfg.get("smtp_from", "")
    if not host or not user or not from_addr:
        return False
    recipients = _build_recipients(destinatario)
    if not recipients:
        return False
    msg = MIMEMultipart()
    msg["From"] = from_addr
    msg["To"] = recipients[0]
    if len(recipients) > 1:
        msg["Cc"] = ", ".join(recipients[1:])
    msg["Subject"] = assunto
    msg.attach(MIMEText(corpo, "plain"))
    for a in anexos:
        p = Path(a)
        if not p.exists():
            continue
        part = MIMEBase("application", "octet-stream")
        part.set_payload(p.read_bytes())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename={p.name}")
        msg.attach(part)
    try:
        with smtplib.SMTP(host, port) as s:
            s.starttls()
            if user:
                s.login(user, passwd)
            s.sendmail(from_addr, recipients, msg.as_string())
        return True
    except Exception:
        return False
