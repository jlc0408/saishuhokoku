import configparser
import os

CONFIG_PATH = os.path.join(os.path.dirname(__file__), "config.ini")

DEFAULTS = {
    "BEAMS": {
        "url": "https://ap.salesforce.com/secur/login_portal.jsp?orgId=00D10000000IDQS&portalId=06010000000Lc5O",
        "username": "",
        "password": "",
    },
    "EDGE": {
        "edge_path": r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    },
}


def load_config() -> configparser.ConfigParser:
    cfg = configparser.ConfigParser()
    if os.path.exists(CONFIG_PATH):
        cfg.read(CONFIG_PATH, encoding="utf-8")
    # 不足セクション・キーを補完
    for section, values in DEFAULTS.items():
        if not cfg.has_section(section):
            cfg.add_section(section)
        for key, val in values.items():
            if not cfg.has_option(section, key):
                cfg.set(section, key, val)
    return cfg


def save_config(cfg: configparser.ConfigParser):
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        cfg.write(f)


def get(section: str, key: str) -> str:
    cfg = load_config()
    return cfg.get(section, key, fallback="")


def set_value(section: str, key: str, value: str):
    cfg = load_config()
    if not cfg.has_section(section):
        cfg.add_section(section)
    cfg.set(section, key, value)
    save_config(cfg)
