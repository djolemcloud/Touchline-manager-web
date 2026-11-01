from __future__ import annotations
import json
from functools import lru_cache
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
MAP_PATH = ROOT / 'HISTORY' / 'TROPHY_MAP.json'
LOCAL_FALLBACK = '/static/history/trophies/fallback_animated.webp'

@lru_cache(maxsize=1)
def _map():
    try:
        return json.loads(MAP_PATH.read_text(encoding='utf-8'))
    except Exception:
        return {'fallbacks':{}, 'competitions':{}}

def trophy_for(history_key: str|None, comp_type: str|None):
    data=_map(); key=(history_key or '').strip(); kind=(comp_type or '').strip().lower()
    rec=dict((data.get('competitions') or {}).get(key) or {})
    if not rec and key:
        # Historical competition keys often split into regional/group variants.
        # Use the longest configured prefix so every history card stays inside
        # the same local CGI trophy family without relying on remote images.
        families=data.get('families') or {}
        matches=[(prefix,value) for prefix,value in families.items() if key.startswith(prefix)]
        if matches:
            prefix,value=max(matches,key=lambda item:len(item[0]))
            rec=dict(value or {})
    if not rec:
        image=(data.get('fallbacks') or {}).get(kind) or LOCAL_FALLBACK
        rec={'image':image,'source':'Touchline local CGI fallback','license':'Original Touchline project render'}
    rec.setdefault('image',LOCAL_FALLBACK)
    rec.setdefault('source','')
    rec.setdefault('license','')
    return rec
