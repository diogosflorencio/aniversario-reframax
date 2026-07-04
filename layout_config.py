import json
import glob
import os
from datetime import date


PASTA_MODELOS = "modelos"
PASTA_FONTES = "fontes"
EXT_MODELOS = (".png", ".jpg", ".jpeg", ".webp", ".bmp")
EXT_FONTES = (".ttf", ".otf")


def listar_arquivos_pasta(base_dir, pasta, extensoes):
    caminho = os.path.join(base_dir, pasta)
    if not os.path.isdir(caminho):
        return []
    return sorted(
        nome for nome in os.listdir(caminho)
        if os.path.splitext(nome)[1].lower() in extensoes
    )


def listar_modelos(base_dir):
    return listar_arquivos_pasta(base_dir, PASTA_MODELOS, EXT_MODELOS)


def listar_fontes(base_dir):
    return listar_arquivos_pasta(base_dir, PASTA_FONTES, EXT_FONTES)


def caminho_modelo(base_dir, nome_arquivo):
    return os.path.join(base_dir, PASTA_MODELOS, nome_arquivo)


def caminho_fonte(base_dir, nome_arquivo):
    return os.path.join(base_dir, PASTA_FONTES, nome_arquivo)


def defaults_layout(config):
    modelo = None
    if config.get("imagem_fundo"):
        modelo = os.path.basename(resolver_caminho(config["_dir"], config["imagem_fundo"]))

    mes = config.get("mes", {})
    lista = config.get("lista", {})
    fonte_mes = os.path.basename(
        resolver_caminho(config["_dir"], mes.get("fonte_arquivo", "../../fontes/OpenSans-Regular.ttf"))
    )
    fonte_lista = os.path.basename(
        resolver_caminho(config["_dir"], lista.get("fonte_arquivo", "../../fontes/OpenSans-Regular.ttf"))
    )
    return {"modelo": modelo, "fonte_mes": fonte_mes, "fonte_lista": fonte_lista}


def _layouts_dir(base_dir):
    return os.path.join(base_dir, "layouts")


def listar_layouts(base_dir):
    """Retorna configs de todos os layouts encontrados em layouts/*/layout.json."""
    layouts = []
    for path in sorted(glob.glob(os.path.join(_layouts_dir(base_dir), "*/layout.json"))):
        with open(path, encoding="utf-8") as f:
            config = json.load(f)
        config["_dir"] = os.path.dirname(path)
        config["_path"] = path
        layouts.append(config)
    return layouts


def carregar_layout(base_dir, layout_id):
    for config in listar_layouts(base_dir):
        if config.get("id") == layout_id:
            return config
    return None


def resolver_caminho(layout_dir, caminho_relativo):
    return os.path.normpath(os.path.join(layout_dir, caminho_relativo))


def caminho_imagem_fundo(config):
    return resolver_caminho(config["_dir"], config["imagem_fundo"])


def posicao_mes(config, altura_imagem):
    mes = config.get("mes", {})
    y = mes.get("y_abs")
    if y is None:
        y = int(altura_imagem * mes.get("y_rel", 0.1))
    return {
        "y": y,
        "centralizado_h": mes.get("centralizado_h", True),
        "fonte_tamanho": mes.get("fonte_tamanho", 200),
        "fonte_arquivo": resolver_caminho(config["_dir"], mes.get("fonte_arquivo", "../../fontes/OpenSans-Regular.ttf")),
        "cor": tuple(mes.get("cor", [43, 57, 100])),
    }


def tipo_layout(config):
    return config.get("tipo", "mensal")


def nome_mes(numero):
    return date(2020, numero, 1).strftime("%B").capitalize()


def config_titulo_ano(config, altura_imagem):
    titulo = config.get("titulo_ano", {})
    y = titulo.get("y_abs")
    if y is None:
        y = int(altura_imagem * titulo.get("y_rel", 0.04))
    return {
        "y": y,
        "centralizado_h": titulo.get("centralizado_h", True),
        "fonte_tamanho": titulo.get("fonte_tamanho", 120),
        "cor": tuple(titulo.get("cor", [43, 57, 100])),
    }


def config_titulo_mes(config, numero_mes, altura_imagem):
    mes_cfg = config.get("meses", {}).get(str(numero_mes), {})
    titulo = mes_cfg.get("titulo", {})
    y = titulo.get("y_abs")
    if y is None and titulo.get("y_rel") is not None:
        y = int(altura_imagem * titulo["y_rel"])
    return {
        "y": y,
        "x_rel": titulo.get("x_rel"),
        "centralizado_h": titulo.get("centralizado_h", True),
        "fonte_tamanho": titulo.get("fonte_tamanho", 36),
        "cor": tuple(titulo.get("cor", [43, 57, 100])),
        "visivel": titulo.get("visivel", True),
    }


def config_lista(config, mes_numero=None):
    lista = dict(config.get("lista", {}))
    if tipo_layout(config) == "anual" and mes_numero:
        mes_cfg = config.get("meses", {}).get(str(mes_numero), {})
        lista.update(mes_cfg.get("lista", {}))
        if lista.get("x_rel") is None:
            titulo = mes_cfg.get("titulo", {})
            if titulo.get("x_rel") is not None:
                lista["x_rel"] = titulo["x_rel"]
    return {
        "inicio_y": lista.get("inicio_y", 450),
        "espacamento": lista.get("espacamento", 70),
        "fonte_tamanho": lista.get("fonte_tamanho", 30),
        "fonte_arquivo": resolver_caminho(config["_dir"], lista.get("fonte_arquivo", "../../fontes/OpenSans-Regular.ttf")),
        "cor_texto": tuple(lista.get("cor_texto", [43, 57, 100])),
        "largura_linha": lista.get("largura_linha", 600),
        "x_rel": lista.get("x_rel"),
        "slots": lista.get("slots", []),
    }


def slot_para_indice(config, indice, y_calculada, mes_numero=None):
    """Retorna posições e forma de um slot; usa slots fixos se definidos."""
    slots = config_lista(config, mes_numero).get("slots") or []
    if indice < len(slots):
        slot = slots[indice]
        return {
            "y": slot.get("y_abs", y_calculada),
            "nome_x_rel": slot.get("nome", {}).get("x_rel"),
            "dia_x_rel": slot.get("dia", {}).get("x_rel"),
            "forma": slot.get("forma"),
        }
    return {"y": y_calculada, "nome_x_rel": None, "dia_x_rel": None, "forma": None}


def forma_padrao(config):
    return config.get("forma_padrao")


def config_estilo_layout(config, estilo_chave):
    return config.get("estilos", {}).get(estilo_chave, {})
