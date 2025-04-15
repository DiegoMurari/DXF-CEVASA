import os
import json

CONFIG_FILE = "config.json"  # ou use: os.path.join(os.path.expanduser("~"), ".dxf_config.json")

def salvar_pasta_saida(pasta):
    with open(CONFIG_FILE, "w") as f:
        json.dump({"pasta_saida": pasta}, f)

def carregar_pasta_saida():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r") as f:
            config = json.load(f)
            return config.get("pasta_saida", "")
    return ""
