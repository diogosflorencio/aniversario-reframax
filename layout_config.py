import json
import glob
import os


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


def config_lista(config):
    lista = config.get("lista", {})
    return {
        "inicio_y": lista.get("inicio_y", 450),
        "espacamento": lista.get("espacamento", 70),
        "fonte_tamanho": lista.get("fonte_tamanho", 30),
        "fonte_arquivo": resolver_caminho(config["_dir"], lista.get("fonte_arquivo", "../../fontes/OpenSans-Regular.ttf")),
        "cor_texto": tuple(lista.get("cor_texto", [43, 57, 100])),
        "largura_linha": lista.get("largura_linha", 600),
        "slots": lista.get("slots", []),
    }


def slot_para_indice(config, indice, y_calculada):
    """Retorna posições e forma de um slot; usa slots fixos se definidos."""
    slots = config_lista(config).get("slots") or []
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
