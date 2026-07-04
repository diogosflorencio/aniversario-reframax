import json
import sys
import urllib.error
import urllib.request

from version import VERSION

SERVER_URL = "https://raw.githubusercontent.com/diogosflorencio/aniversario-reframax/main/server.json"
SERVER_KEY = "status"
SERVER_VALUE = "ativo"
TIMEOUT_SEGUNDOS = 10


def app_compilado():
    return getattr(sys, "frozen", False)


def verificar_servidor_ativo():
    if not app_compilado():
        return True, ""

    try:
        req = urllib.request.Request(
            SERVER_URL,
            headers={"User-Agent": f"GeradorAniversariantes/{VERSION}"},
        )
        with urllib.request.urlopen(req, timeout=TIMEOUT_SEGUNDOS) as response:
            dados = json.loads(response.read().decode("utf-8"))

        if dados.get(SERVER_KEY) == SERVER_VALUE:
            return True, ""

        return False, "Servidor fora do ar. Falar com Diogo =/"

    except (urllib.error.URLError, urllib.error.HTTPError, TimeoutError, json.JSONDecodeError, OSError):
        return False, "Servidor fora do ar. Falar com Diogo =/"
