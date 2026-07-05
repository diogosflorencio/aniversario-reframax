import json
import glob
import os
from datetime import date


PASTA_MODELOS = "modelos"
PASTA_FONTES = "fontes"
EXT_MODELOS = (".png", ".jpg", ".jpeg", ".webp", ".bmp")
EXT_FONTES = (".ttf", ".otf")
FONTE_PADRAO = "../../fontes/OpenSans-Regular.ttf"


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


def _stem_arquivo(nome):
    return os.path.splitext(nome)[0].lower()


def modelo_referencia_layout(config):
    if not config or not config.get("imagem_fundo"):
        return None
    return os.path.basename(resolver_caminho(config["_dir"], config["imagem_fundo"]))


def modelos_correspondem(nome_a, nome_b):
    if not nome_a or not nome_b:
        return False
    return _stem_arquivo(nome_a) == _stem_arquivo(nome_b)


def modelo_corresponde_layout(config, nome_modelo):
    return modelos_correspondem(modelo_referencia_layout(config), nome_modelo)


def encontrar_layout_por_modelo(base_dir, nome_modelo):
    if not nome_modelo or str(nome_modelo).startswith("Nenhum"):
        return None
    for config in listar_layouts(base_dir):
        if modelo_corresponde_layout(config, nome_modelo):
            return config.get("id")
    return None


def defaults_layout(config):
    modelo = None
    if config.get("imagem_fundo"):
        modelo = os.path.basename(resolver_caminho(config["_dir"], config["imagem_fundo"]))

    mes = config.get("mes", {})
    lista = config.get("lista", {})
    titulo_ano = config.get("titulo_ano", {})
    fonte_titulo = mes.get("fonte_arquivo") or titulo_ano.get("fonte_arquivo") or FONTE_PADRAO
    fonte_mes = os.path.basename(resolver_caminho(config["_dir"], fonte_titulo))
    fonte_lista = os.path.basename(
        resolver_caminho(config["_dir"], lista.get("fonte_arquivo", FONTE_PADRAO))
    )
    cores = defaults_cores(config)
    return {
        "modelo": modelo,
        "fonte_mes": fonte_mes,
        "fonte_lista": fonte_lista,
        "cor_ano": list(cores["cor_ano"]),
        "cor_mes": list(cores["cor_mes"]),
        "cor_nomes": list(cores["cor_nomes"]),
    }


def defaults_cores(config):
    mes = config.get("mes", {})
    titulo_ano = config.get("titulo_ano", {})
    lista = config.get("lista", {})
    titulo_mes = config.get("meses", {}).get("1", {}).get("titulo", {})
    padrao = [43, 57, 100]
    cor_mes = mes.get("cor") or titulo_mes.get("cor") or titulo_ano.get("cor") or padrao
    cor_ano = titulo_ano.get("cor") or mes.get("cor") or padrao
    cor_nomes = lista.get("cor_texto") or padrao
    return {
        "cor_ano": tuple(cor_ano),
        "cor_mes": tuple(cor_mes),
        "cor_nomes": tuple(cor_nomes),
    }


def _layouts_dir(base_dir):
    return os.path.join(base_dir, "layouts")


def layout_chave(config):
    """Identificador único do layout = nome da pasta em layouts/."""
    return os.path.basename(config["_dir"])


def listar_layouts(base_dir):
    """Retorna configs de todos os layouts encontrados em layouts/*/layout.json."""
    layouts = []
    for path in sorted(glob.glob(os.path.join(_layouts_dir(base_dir), "*/layout.json"))):
        if os.path.basename(os.path.dirname(path)).startswith("_"):
            continue
        with open(path, encoding="utf-8") as f:
            config = json.load(f)
        config["_dir"] = os.path.dirname(path)
        config["_path"] = path
        layouts.append(config)
    return layouts


def carregar_layout(base_dir, chave):
    if not chave:
        return None
    layouts = listar_layouts(base_dir)
    for config in layouts:
        if layout_chave(config) == chave:
            return config
    for config in layouts:
        if config.get("id") == chave:
            return config
    return None


def resolver_caminho(layout_dir, caminho_relativo):
    return os.path.normpath(os.path.join(layout_dir, caminho_relativo))


def caminho_imagem_fundo(config):
    return resolver_caminho(config["_dir"], config["imagem_fundo"])


def tipo_layout(config):
    return config.get("tipo", "mensal")


def nome_mes(numero):
    return date(2020, numero, 1).strftime("%B").capitalize()


def _resolver_y(bloco, altura_imagem, y_rel_padrao=0.1):
    if not bloco:
        return int(altura_imagem * y_rel_padrao)
    if bloco.get("y_abs") is not None:
        return bloco["y_abs"]
    if bloco.get("y_rel") is not None:
        return int(altura_imagem * bloco["y_rel"])
    return int(altura_imagem * y_rel_padrao)


def _resolver_fonte_arquivo(config, bloco, *blocos_fallback):
    if bloco and bloco.get("fonte_arquivo"):
        return resolver_caminho(config["_dir"], bloco["fonte_arquivo"])
    for fb in blocos_fallback:
        if fb and fb.get("fonte_arquivo"):
            return resolver_caminho(config["_dir"], fb["fonte_arquivo"])
    return resolver_caminho(config["_dir"], FONTE_PADRAO)


def _config_texto(bloco, config, altura_imagem, y_rel_padrao=0.1, tamanho_padrao=30, blocos_fonte=()):
    bloco = bloco or {}
    return {
        "y": _resolver_y(bloco, altura_imagem, y_rel_padrao),
        "x_abs": bloco.get("x_abs"),
        "x_rel": bloco.get("x_rel"),
        "centralizado_h": bloco.get("centralizado_h", True),
        "fonte_tamanho": bloco.get("fonte_tamanho", tamanho_padrao),
        "fonte_arquivo": _resolver_fonte_arquivo(config, bloco, *blocos_fonte),
        "cor": tuple(bloco.get("cor", [43, 57, 100])),
        "visivel": bloco.get("visivel", True),
    }


def posicao_mes(config, altura_imagem):
    """Modo mensal: título grande com o nome do mês selecionado."""
    mes = config.get("mes", {})
    return _config_texto(mes, config, altura_imagem, y_rel_padrao=0.1, tamanho_padrao=200)


def config_titulo_ano(config, altura_imagem):
    """Modo anual: número do ano no topo do cartaz."""
    titulo = config.get("titulo_ano", {})
    mes = config.get("mes", {})
    return _config_texto(
        titulo, config, altura_imagem,
        y_rel_padrao=0.04, tamanho_padrao=120,
        blocos_fonte=(mes,),
    )


def config_titulo_mes(config, numero_mes, altura_imagem):
    """Modo anual: nome de cada mês (Janeiro, Fevereiro...) dentro de meses/N/titulo."""
    mes_cfg = config.get("meses", {}).get(str(numero_mes), {})
    titulo = mes_cfg.get("titulo", {})
    titulo_ano = config.get("titulo_ano", {})
    mes = config.get("mes", {})
    cfg = _config_texto(
        titulo, config, altura_imagem,
        y_rel_padrao=0.0, tamanho_padrao=36,
        blocos_fonte=(titulo_ano, mes),
    )
    if titulo.get("y_abs") is None and titulo.get("y_rel") is None:
        cfg["y"] = None
    return cfg


def config_lista(config, mes_numero=None):
    """Lista de aniversariantes. No anual, meses/N/lista sobrescreve os valores globais."""
    lista = dict(config.get("lista", {}))
    mes_cfg = {}
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
        "fonte_arquivo": _resolver_fonte_arquivo(config, lista, config.get("mes"), config.get("titulo_ano")),
        "cor_texto": tuple(lista.get("cor_texto", [43, 57, 100])),
        "largura_linha": lista.get("largura_linha", 600),
        "x_rel": lista.get("x_rel"),
        "x_abs": lista.get("x_abs"),
        "deslocamento_y": lista.get("deslocamento_y", 0),
        "slots": lista.get("slots", []),
    }


def slot_para_indice(config, indice, y_calculada, mes_numero=None):
    """Posição de cada linha da lista. slots[] permite coordenadas fixas por índice."""
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


def anchor_x_bloco(cfg, largura_bloco, largura_imagem):
    """Calcula X inicial de um bloco de texto usando x_abs, x_rel ou centro da imagem."""
    if cfg.get("x_abs") is not None:
        return cfg["x_abs"]
    x_rel = cfg.get("x_rel")
    if x_rel is not None:
        return largura_imagem * x_rel - largura_bloco / 2
    return (largura_imagem - largura_bloco) / 2


def calcular_x_texto(cfg, largura_texto, largura_imagem):
    """Calcula X de um rótulo (título) usando x_abs, x_rel ou centralizado_h."""
    if cfg.get("x_abs") is not None:
        return cfg["x_abs"]
    x_rel = cfg.get("x_rel")
    if x_rel is not None:
        if cfg.get("centralizado_h", True):
            return largura_imagem * x_rel - largura_texto / 2
        return largura_imagem * x_rel
    if cfg.get("centralizado_h", True):
        return (largura_imagem - largura_texto) / 2
    return cfg.get("x_abs", 0)
