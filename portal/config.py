from __future__ import annotations
import os
from pathlib import Path

PROJECT_DIR = Path(__file__).resolve().parents[1]
DATA_DIR = PROJECT_DIR / 'data'
STATIC_DIR = PROJECT_DIR / 'static'
DB_PATH = DATA_DIR / 'cm_data_hub.db'
HOST = os.environ.get('TOUCHLINE_HOST', '127.0.0.1')
PORT = int(os.environ.get('TOUCHLINE_PORT', '5100'))
APP_NAME = 'Touchline Web Portal'
VERSION='1.7.1'
