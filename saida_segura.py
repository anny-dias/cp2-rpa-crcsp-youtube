"""Publica o Excel sem perder a coleta quando o Windows bloqueia um arquivo."""
import logging
import os
import shutil
import tempfile
import time
from datetime import datetime
from pathlib import Path
from uuid import uuid4

LOG = logging.getLogger(__name__)


def criar_temporario(destino):
    destino = Path(destino)
    destino.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(dir=destino.parent, prefix=destino.stem + "_",
                                     suffix=".tmp.xlsx", delete=False) as handle:
        return Path(handle.name)


def publicar_excel(temporario, destino):
    """Retorna o caminho efetivo; em falha total mantém o temporário recuperável."""
    temporario, destino = Path(temporario), Path(destino)
    for tentativa in range(3):
        try:
            os.replace(temporario, destino)
            return destino
        except PermissionError:
            if tentativa < 2:
                time.sleep(0.25 * (tentativa + 1))

    alternativo = destino.with_name(
        f"{destino.stem}_{datetime.now():%Y%m%d_%H%M%S}_{uuid4().hex[:8]}{destino.suffix}")
    try:
        # Copiar também funciona quando a origem permite leitura, mas não renomeação.
        # 'xb' impede sobrescrever qualquer arquivo existente.
        with temporario.open('rb') as origem, alternativo.open('xb') as saida:
            shutil.copyfileobj(origem, saida)
    except OSError as exc:
        raise OSError(
            f"Não foi possível publicar o Excel. A coleta permanece em: {temporario}. "
            "Feche o Excel e confira as permissões da pasta antes de tentar novamente."
        ) from exc
    try:
        temporario.unlink()
    except OSError:
        LOG.warning("Arquivo temporário mantido para recuperação: %s", temporario)
    LOG.warning("Arquivo bloqueado para substituição; Excel salvo em: %s", alternativo)
    return alternativo
