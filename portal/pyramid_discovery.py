from __future__ import annotations

import re
import unicodedata
from urllib.parse import urljoin
from bs4 import BeautifulSoup

SECTIONS = {
    'srpske lige': 3,
    'zone': 4,
    'okruzne lige': 5,
    'medjuopstinske lige': 6,
    'opstinske lige': 6,
    'druge opstinske lige': 6,
}


def _clean(value):
    return re.sub(r'\s+', ' ', str(value or '')).strip()


def _norm(value):
    s=_clean(value).lower().replace('đ','dj')
    s=unicodedata.normalize('NFKD', s)
    s=''.join(ch for ch in s if not unicodedata.combining(ch))
    return re.sub(r'[^a-z0-9]+',' ',s).strip()




def _valid_competition_name(name, level):
    """Guard the live index against global-menu / cross-section leakage."""
    n=_norm(name)
    if level == 3:
        return n in {
            'srpska liga beograd','srpska liga vojvodina','srpska liga istok','srpska liga zapad'
        }
    if level == 4:
        return ('zona' in n) or n.startswith('vojvodjanska liga ')
    if level == 5:
        return (('okruzna liga' in n) or n.startswith('pfl ') or
                n.startswith('prva beogradska liga ') or n=='prva niska liga' or
                n=='prva liga kragujevca' or n.startswith('sremska liga'))
    if level == 6:
        # Sixth tier is intentionally broad, but must never accept national L1/L2,
        # Serbian League (L3), zones (L4), district leagues (L5), cups or youth/women.
        forbidden=(
            n in {'mozzart bet super liga srbije','superliga srbije','super liga srbije',
                  'mozzart bet prva liga srbije','prva liga srbije'} or
            n.startswith('srpska liga ') or ' zona' in (' '+n) or n.startswith('zona ') or
            'okruzna liga' in n or n.startswith('pfl ') or n.startswith('prva beogradska liga ') or
            'kup ' in n or n.startswith('kup') or 'zenska' in n or 'omladin' in n or 'kadet' in n
        )
        return not forbidden
    return False


def discover_index(html: str):
    """Return current senior football competitions grouped into levels 3-6.

    Links are collected only between known SrbijaSport section headings. This
    prevents cup, foreign, youth, news and other-sport links from leaking into
    the Serbian senior pyramid.
    """
    soup=BeautifulSoup(html or '', 'lxml')
    headings=set(SECTIONS)
    stops=headings | {
        'kup takmicenja','lige petice','evropa','prijateljske utakmice',
        'sva takmicenja','top takmicenja',
    }
    blocks={3:[],4:[],5:[],6:[]}
    level6_blocks=[]

    for node in soup.find_all(string=True):
        head=_norm(node)
        if head not in headings:
            continue
        level=SECTIONS[head]
        found=[]; seen=set()
        for el in node.parent.next_elements:
            if isinstance(el, str):
                t=_norm(el)
                if found and t in stops and t != head:
                    break
                continue
            if getattr(el, 'name', None) != 'a':
                continue
            href=_clean(el.get('href') or '')
            m=re.search(r'(?:^|/)league/(\d+)(?:-|/|$)', href)
            if not m:
                continue
            source_id=m.group(1)
            if source_id in seen:
                continue
            name=_clean(el.get_text(' ', strip=True))
            if not name:
                img=el.find('img')
                if img:
                    name=_clean(img.get('alt') or img.get('title') or '')
            if not name or not _valid_competition_name(name, level):
                continue
            seen.add(source_id)
            found.append({
                'source_id':source_id,
                'name':name,
                'url':urljoin('https://srbijasport.net/', href),
                'level':level,
                'section_heading':_clean(node),
            })
        if not found:
            continue
        if level == 6:
            level6_blocks.append(found)
        elif len(found) > len(blocks[level]):
            blocks[level]=found

    # Level 6 is intentionally the union of Međuopštinske + Opštinske +
    # Druge opštinske, matching the Touchline product decision that the public
    # "Opštinske lige" card lists every sixth-level league individually.
    all6=[]; seen6=set()
    for block in level6_blocks:
        for item in block:
            if item['source_id'] in seen6:
                continue
            seen6.add(item['source_id']); all6.append(item)
    blocks[6]=all6

    # General de-duplication by SrbijaSport competition ID.
    for level,items in list(blocks.items()):
        out=[]; seen=set()
        for item in items:
            if item['source_id'] in seen:
                continue
            seen.add(item['source_id']); out.append(item)
        blocks[level]=out
    return blocks
