import json
import os
from datetime import date


DEFAULTS = {
    "excel": "",
    "mes_index": 0,
    "layout_id": "padrao",
    "estilo": "Estilo padrão",
    "carregar_ao_iniciar": True,
    "lembrete_projeto": {
        "ativo": False,
        "dia_mes": 25,
    },
    "geracao_automatica": {
        "ativo": False,
        "dia_mes": 1,
        "imprimir": True,
    },
    "ultimo_lembrete": "",
    "ultima_geracao_auto": "",
}


def caminho_arquivo(base_dir):
    return os.path.join(base_dir, "preferencias.json")


def carregar(base_dir):
    path = caminho_arquivo(base_dir)
    if not os.path.isfile(path):
        return dict(DEFAULTS)

    try:
        with open(path, encoding="utf-8") as f:
            dados = json.load(f)
    except (json.JSONDecodeError, OSError):
        return dict(DEFAULTS)

    prefs = dict(DEFAULTS)
    for chave, valor in dados.items():
        if isinstance(valor, dict) and chave in prefs and isinstance(prefs[chave], dict):
            prefs[chave].update(valor)
        else:
            prefs[chave] = valor
    return prefs


def salvar(base_dir, prefs):
    path = caminho_arquivo(base_dir)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(prefs, f, ensure_ascii=False, indent=2)


def hoje_iso():
    return date.today().isoformat()


def dia_mes_hoje():
    return date.today().day
