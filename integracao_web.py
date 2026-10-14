"""Integração HTTP simples com o YouTube para validar o link informado.

Esta etapa é complementar ao yt-dlp e existe para tornar explícita a
integração externa do projeto acadêmico. Não substitui o pipeline principal.
"""
from __future__ import annotations

import requests


def validar_video_youtube(url: str, timeout: int = 10) -> dict:
    """Consulta o endpoint oEmbed do YouTube e retorna metadados básicos.

    Raises:
        requests.RequestException: em falha de rede/HTTP.
        ValueError: se a resposta não contiver os campos esperados.
    """
    endpoint = "https://www.youtube.com/oembed"
    resposta = requests.get(
        endpoint,
        params={"url": url, "format": "json"},
        timeout=timeout,
        headers={"User-Agent": "CP2-RPA-CRCSP/1.0"},
    )
    resposta.raise_for_status()
    dados = resposta.json()
    if not dados.get("title"):
        raise ValueError("O YouTube não retornou um título para o vídeo informado.")
    return {
        "titulo_oembed": dados.get("title", ""),
        "autor_oembed": dados.get("author_name", ""),
        "url_autor_oembed": dados.get("author_url", ""),
    }
