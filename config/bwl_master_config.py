import os

BASE_DIR = os.path.dirname(os.path.dirname(__file__))
RES_DIR = os.path.join(BASE_DIR, "ressources")
DOWNLOAD_DIR = os.path.join(RES_DIR, "downloads")
EXTRACT_DIR = os.path.join(DOWNLOAD_DIR, "extracted")
OUTPUT_DIR = RES_DIR
MODULE_MAP_CSV = os.path.join(RES_DIR, "modul_mengen_stat_vwl_bwl.csv")
OUTPUT_CSV = os.path.join(RES_DIR, "bewerber_evaluierung.csv")


REQ_NOTE_MAX = 2.4
REQUIREMENTS = {
    "BWL": 60.0,
    "VWL": 20.0,
    "Statistik": 15.0
}

DOM_ECTS_MAP = {
    "volkswirtschaftslehre": "VWL",
    "statistik": "Statistik",
    "betriebswirtschaftslehre": "BWL"
}
WHITELIST_UNIS = os.path.join(RES_DIR, "whitelist_unis_bwl.csv")
