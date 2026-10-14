"""Validação local e normalização de links, sem chamadas à rede."""
import re
from urllib.parse import parse_qs, urlsplit


def normalizar_link(link: str) -> tuple[str, str]:
    value = link.strip().strip('\"\'')
    if '://' not in value:
        value = 'https://' + value
    parsed = urlsplit(value)
    if parsed.scheme not in {'https', 'http'} or parsed.username or parsed.password:
        raise ValueError('Informe um link HTTP/HTTPS do YouTube.')
    host = (parsed.hostname or '').lower()
    parts = parsed.path.strip('/').split('/')
    vid = None
    if host in {'youtu.be', 'www.youtu.be'} and len(parts) == 1:
        vid = parts[0]
    elif host in {'youtube.com', 'www.youtube.com', 'm.youtube.com', 'music.youtube.com',
                  'youtube-nocookie.com', 'www.youtube-nocookie.com'}:
        if parsed.path.rstrip('/') == '/watch':
            ids = parse_qs(parsed.query).get('v', [])
            if len(ids) == 1:
                vid = ids[0]
        elif len(parts) == 2 and parts[0] in {'live', 'shorts', 'embed', 'v'}:
            vid = parts[1]
    if not vid or not re.fullmatch(r'[A-Za-z0-9_-]{11}', vid):
        raise ValueError('Link inválido. Cole o endereço de um vídeo ou transmissão, não de um canal ou playlist.')
    return f'https://www.youtube.com/watch?v={vid}', vid


def solicitar_link() -> str:
    while True:
        try:
            value = input('Cole o link do vídeo do YouTube que deseja analisar: ').strip()
        except (EOFError, KeyboardInterrupt):
            raise ValueError('Entrada cancelada. Você também pode informar o link como argumento.') from None
        try:
            return normalizar_link(value)[0]
        except ValueError as exc:
            print(exc)
