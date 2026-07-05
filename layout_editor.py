import copy
import json
import os


def carregar_arquivo(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def salvar_arquivo(path, dados):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=2)


def clonar_config(config):
    return copy.deepcopy({k: v for k, v in config.items() if not str(k).startswith("_")})


def caminho_relativo_modelo(nome_arquivo):
    return f"../../modelos/{nome_arquivo}"


def caminho_relativo_fonte(nome_arquivo):
    return f"../../fontes/{nome_arquivo}"


def y_para_editor(bloco, altura_imagem, padrao=100):
    if not bloco:
        return padrao
    if bloco.get("y_abs") is not None:
        return int(bloco["y_abs"])
    if bloco.get("y_rel") is not None:
        return int(altura_imagem * bloco["y_rel"])
    return padrao


def x_rel_para_editor(bloco, padrao=0.5):
    if not bloco:
        return padrao
    if bloco.get("x_rel") is not None:
        return float(bloco["x_rel"])
    return padrao


def bloco_texto_de_editor(y_abs, x_rel, fonte_tamanho, cor=None, visivel=True):
    bloco = {
        "y_abs": int(y_abs),
        "x_rel": round(float(x_rel), 4),
        "centralizado_h": True,
        "fonte_tamanho": int(fonte_tamanho),
        "visivel": bool(visivel),
    }
    if cor is not None:
        bloco["cor"] = list(cor)
    return bloco


def lista_de_editor(inicio_y, x_rel, espacamento, fonte_tamanho, largura_linha, cor_texto=None):
    bloco = {
        "inicio_y": int(inicio_y),
        "x_rel": round(float(x_rel), 4),
        "espacamento": int(espacamento),
        "fonte_tamanho": int(fonte_tamanho),
        "largura_linha": int(largura_linha),
    }
    if cor_texto is not None:
        bloco["cor_texto"] = list(cor_texto)
    return bloco
