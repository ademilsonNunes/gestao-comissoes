from typing import List, Dict, Any
from ..database.models import calcular_periodo, obter_comissao_periodo
from ..database.database import init_schema


def calcular(mes: int, ano: int) -> List[Dict[str, Any]]:
    init_schema()
    return calcular_periodo(mes, ano)


def consolidado(mes: int, ano: int) -> List[Dict[str, Any]]:
    init_schema()
    return obter_comissao_periodo(mes, ano)
