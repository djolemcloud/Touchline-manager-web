from __future__ import annotations

"""Central L1/L2 historical table importer.

Touchline 1.5.3 deliberately treats the league tables as the master historical
source and derives club history from those tables.  This module does *not* call
Wikipedia per club.  It downloads RSSSF season pages, caches them locally and
upserts every final table row against the existing Touchline master club IDs.

Coverage target:
- level 1: 1946/47 -> 2025/26 (1945/46 Serbian republic seed remains in
  history_sync.seed_history())
- level 2: 1947/48 -> 2025/26, including historical groups when the second tier
  was split into East/West/North/South/etc.
"""

import re
import time
import hashlib
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

import requests
from bs4 import BeautifulSoup

from .config import DATA_DIR
from .data import connect
from . import history_sync as hs

RSSSF_JOEG_BASE = 'https://www.rsssf.org/tablesj/'
RSSSF_SERBIA_BASE = 'https://www.rsssf.org/tabless/'
CACHE_DIR = DATA_DIR / 'history_cache' / 'rsssf_l12'
CACHE_DIR.mkdir(parents=True, exist_ok=True)

UA = 'TouchlineHistory/1.5.3 (local Serbian football archive; central L1-L2 tables)'
CENTRAL_NOTE = 'CENTRAL_L12_RSSSF_1.5.3'


@dataclass
class L12Row:
    position: int
    club_name: str
    played: int
    wins: int
    draws: int
    losses: int
    gf: int
    ga: int
    points: float
    tail: str = ''


# Historical RSSSF tables use several row layouts.  The parser deliberately
# reads the numeric tail from the right so it accepts all of these forms:
#   1.Club              30 18  8  4 50-24 44
#   1 Club              34 20  9  5 54:15 49
#   1.OFK BEOGRAD Beograd 30 14 9 7 ( 50: 26) 37
#   1. Zemun            38 23  6 (5) 9 59-35 51
# The optional parenthesised number is the late-Yugoslav shoot-out statistic,
# not a W/D/L field, and is intentionally ignored.
_ROW_TAIL_RE = re.compile(
    r'(?P<played>\d{1,3})\s+(?P<wins>\d{1,3})\s+(?P<draws>\d{1,3})'
    r'(?:\s+\(\s*\d+\s*\))?\s+(?P<losses>\d{1,3})\s+'
    r'\(?\s*(?P<gf>\d{1,3})(?:\s*[-:]\s*|\s+)(?P<ga>\d{1,3})\s*\)?\s+'
    r'(?P<points>-?\d+(?:\.\d+)?)\b(?P<tail>.*)$'
)
_POS_RE = re.compile(r'^\s*(?P<pos>\d{1,2})\s*[.)]?\s*(?P<body>.*)$')


def _clean_club_label(v: str) -> str:
    name = re.sub(r'\s+', ' ', (v or '').strip(' -\t'))
    # Some fixed-width RSSSF tables repeat the locality after a club whose name
    # already ends with it: "OFK BEOGRAD Beograd", "ZAGREB Zagreb".  Removing
    # only an exact duplicated final token is conservative and improves master
    # club matching without guessing locality for ordinary names.
    parts = name.split()
    if len(parts) >= 2 and _fold(parts[-1]) == _fold(parts[-2]):
        name = ' '.join(parts[:-1])
    return name


def parse_row(line: str):
    raw = (line or '').replace('\xa0', ' ')
    mpos = _POS_RE.match(raw)
    if not mpos:
        return None
    mtail = _ROW_TAIL_RE.search(mpos.group('body'))
    if not mtail:
        return None
    club = _clean_club_label(mpos.group('body')[:mtail.start()])
    if not club or len(club) > 90:
        return None
    try:
        return L12Row(
            int(mpos.group('pos')), club,
            int(mtail.group('played')), int(mtail.group('wins')),
            int(mtail.group('draws')), int(mtail.group('losses')),
            int(mtail.group('gf')), int(mtail.group('ga')),
            float(mtail.group('points')), (mtail.group('tail') or '').strip()
        )
    except Exception:
        return None

def _clean_line(v: str) -> str:
    return re.sub(r'\s+', ' ', (v or '').strip())


def _fold(v: str) -> str:
    import unicodedata
    s = (v or '').lower().replace('đ', 'dj')
    s = unicodedata.normalize('NFKD', s)
    s = ''.join(ch for ch in s if not unicodedata.combining(ch))
    return re.sub(r'\s+', ' ', s).strip()


def _season_for_page_year(page_year: int) -> str:
    """Normalize historical page year to Touchline season label.

    The product convention is one YYYY/YY public timeline.  RSSSF's short
    calendar championships 1950-1952 are therefore normalized to 1949/50,
    1950/51 and 1951/52, matching the existing Touchline convention.
    """
    y = int(page_year)
    if y in (1950, 1951, 1952):
        start = y - 1
        return f'{start}/{str(y)[-2:]}'
    return f'{y-1}/{str(y)[-2:]}'


def _modern_season(start_year: int) -> str:
    y = int(start_year)
    return f'{y}/{str(y+1)[-2:]}'


def _joeg_url(page_year: int) -> str:
    return f'{RSSSF_JOEG_BASE}joeg{int(page_year)%100:02d}.html'


def _serv_url(start_year: int) -> str:
    end = int(start_year) + 1
    suffix = f'{end%100:02d}' if int(start_year) <= 2008 else str(end)
    return f'{RSSSF_SERBIA_BASE}serv{suffix}.html'


def _cache_path(url: str) -> Path:
    stem = re.sub(r'[^a-zA-Z0-9._-]+', '_', url.split('/')[-1] or 'page.html')
    digest = hashlib.sha1(url.encode('utf-8')).hexdigest()[:10]
    return CACHE_DIR / f'{digest}_{stem}'


def _fetch_html(session: requests.Session, url: str, refresh: bool = False, timeout: int = 30) -> str:
    path = _cache_path(url)
    if path.is_file() and not refresh:
        return path.read_text(encoding='utf-8', errors='ignore')

    last = None
    for attempt in range(4):
        try:
            r = session.get(url, headers={'User-Agent': UA}, timeout=timeout)
            if r.status_code in (429, 503):
                raise RuntimeError(f'HTTP {r.status_code}')
            r.raise_for_status()
            html = r.text
            path.write_text(html, encoding='utf-8')
            return html
        except Exception as e:
            last = e
            time.sleep(1.0 + attempt * 1.5)
    raise RuntimeError(f'{url}: {last}')


def _html_to_lines(html: str) -> list[str]:
    soup = BeautifulSoup(html or '', 'lxml')
    # RSSSF is mostly PRE text; get_text preserves the logical row lines.
    return [x.rstrip() for x in soup.get_text('\n').splitlines()]


def _is_table_marker(line: str) -> bool:
    s = _fold(_clean_line(line)).rstrip(':')
    return s in ('final table', 'table') or s.startswith('final table ')


def _collect_table(lines: list[str], marker_index: int) -> list[L12Row]:
    rows: list[L12Row] = []
    started = False
    quiet = 0
    for raw in lines[marker_index+1:]:
        s = _clean_line(raw)
        low = _fold(s)
        if _is_table_marker(s) and started:
            break
        row = parse_row(raw)
        if row:
            rows.append(row)
            started = True
            quiet = 0
            continue
        if not started:
            if low.startswith(('round ', 'cross-table', 'cross table')):
                break
            continue
        # separator lines and short notes inside a table are harmless; real new
        # sections end collection immediately.
        if not s or set(s) <= set('-=._ '):
            continue
        if low.startswith(('nb:', 'nb ', 'promoted:', 'relegated:', 'round ', 'cross-table', 'cross table', 'topscorer', 'top scorer')):
            break
        if re.match(r'^(kup|cup|druga |second level|second league|third level|treca |srpska liga|regional leagues)', low):
            break
        quiet += 1
        if quiet >= 2:
            break
    # A credible league table should contain more than a handful of teams.
    return rows if len(rows) >= 6 else []


def _first_index(lines: list[str], patterns: Iterable[str], start: int = 0):
    regs = [re.compile(p, re.I) for p in patterns]
    for i in range(start, len(lines)):
        s = _clean_line(lines[i])
        if len(s) > 120:
            continue
        if any(r.search(s) for r in regs):
            return i
    return None


def _next_table_marker(lines: list[str], start: int, limit: int = 140):
    stop = min(len(lines), start + max(1, limit))
    for i in range(start, stop):
        if _is_table_marker(lines[i]):
            return i
    return None


def _section_heading_for_table(lines: list[str], patterns: Iterable[str], table_index: int, lookback: int = 120):
    """Return the nearest real section heading before a Final Table.

    RSSSF pages often repeat league names in a table-of-contents near the top.
    Choosing the *first* match therefore starts the parser too early and can let
    Regional/Third-Level tables overwrite the real L1 table.  The nearest
    heading before the actual table marker is the robust signal.
    """
    regs = [re.compile(p, re.I) for p in patterns]
    lo = max(0, table_index - lookback)
    for i in range(table_index - 1, lo - 1, -1):
        s = _fold(_clean_line(lines[i]))
        if not s or len(s) > 120:
            continue
        if any(r.search(s) for r in regs):
            return i
    return None


def _section_end(lines: list[str], start: int, level: int):
    # Hard boundaries must include *all* lower-level section families used by
    # RSSSF.  In particular, old Yugoslav pages can jump from L1 directly to
    # "Regional Leagues" without a national second division.
    if level == 1:
        pats = [
            r'^kup\b', r'^cup\b', r'^druga(?: savezna)? liga\b', r'^ii(?:\s+savezna)?\s+liga\b',
            r'^second level\b', r'^second league\b', r'^prva liga\b', r'^zonal leagues?\b',
            r'^regional leagues?\b', r'^republic leagues?\b',
            r'^medjurepublic', r'^inter[- ]?republic', r'^srpska liga\b',
            r'^third level\b', r'^tre[cć]a(?: savezna)? liga\b'
        ]
    else:
        pats = [
            r'^kup\b', r'^cup\b', r'^third level\b', r'^tre[cć]a(?: savezna)? liga\b',
            r'^srpska liga\b', r'^serbian league\b', r'^regional leagues?\b',
            r'^regional league\b', r'^republic leagues?\b'
        ]
    regs = [re.compile(p, re.I) for p in pats]
    for i in range(start + 1, len(lines)):
        s = _fold(_clean_line(lines[i]))
        if not s or len(s) > 120:
            continue
        if any(r.search(s) for r in regs):
            return i
    return len(lines)


def _find_range_near_table(lines: list[str], heading_patterns: Iterable[str], level: int, fallback_first_table: bool = False):
    # Look at actual table markers and bind each one to the nearest matching
    # heading.  This deliberately ignores table-of-contents entries.
    for ti, raw in enumerate(lines):
        if not _is_table_marker(raw):
            continue
        h = _section_heading_for_table(lines, heading_patterns, ti)
        if h is not None:
            end = _section_end(lines, h, level)
            if ti < end:
                return h, end
    if fallback_first_table:
        # FR Yugoslavia pages often place the L1 Final Table directly below the
        # page title, without a "First League" heading.
        ti = _next_table_marker(lines, 0, min(len(lines), 180))
        if ti is not None:
            return 0, _section_end(lines, 0, level)
    return None


def _find_historical_l1_range(lines: list[str]):
    r = _find_range_near_table(
        lines, [r'^prva savezna liga\b', r'^first(?: federal)? league\b', r'^meridian superliga\b'],
        1, fallback_first_table=True
    )
    return r if r else (0, 0)


def _find_historical_l2_range(lines: list[str], search_start: int = 0):
    # The second tier changed name and shape repeatedly: Druga savezna liga,
    # II Savezna Liga, Zonal leagues, Second Level and Yugoslavia 1/B are all
    # genuine level-2 headings in different eras.  Start after L1 so TOC links
    # near the top cannot be mistaken for the real section.
    pats = [
        r'^druga(?: savezna)? liga\b', r'^ii(?:\s+savezna)?\s+liga\b',
        r'^second level\b', r'^second(?: federal)? league\b',
        r'^zonal leagues?\b', r'^yugoslavia\s+1/b\b', r'^1/b\b',
        r'^promotion playoffs? 1953\b'
    ]
    start = _first_index(lines, pats, start=search_start)
    if start is None:
        return None
    return start, _section_end(lines, start, 2)


def _find_modern_l1_range(lines: list[str]):
    return _find_range_near_table(
        lines, [r'^.*super\s*liga\b', r'^.*superliga\b'], 1, fallback_first_table=True
    )


def _find_modern_l2_range(lines: list[str], search_start: int = 0):
    start = _first_index(lines, [r'^prva liga\b', r'^first league\b', r'^second level\b'], start=search_start)
    if start is None:
        return None
    return start, _section_end(lines, start, 2)


def _parse_single_classification(lines: list[str], start: int, end: int) -> list[L12Row]:
    """Parse a level-1/unified level-2 section.

    Later tables overwrite earlier positions.  This intentionally handles the
    Serbian playoff format: regular-stage 1..16 is replaced by championship
    1..8 and relegation 9..16 final tables when present.
    """
    final: dict[int, L12Row] = {}
    for i in range(start, end):
        if not _is_table_marker(lines[i]):
            continue
        for row in _collect_table(lines[:end], i):
            final[row.position] = row
    rows = [final[k] for k in sorted(final)]
    return rows if len(rows) >= 6 else []


_GROUP_TRANSLATIONS = {
    'east': ('Istok', 'East'), 'eastern': ('Istok', 'East'), 'istok': ('Istok', 'East'),
    'west': ('Zapad', 'West'), 'western': ('Zapad', 'West'), 'zapad': ('Zapad', 'West'),
    'north': ('Sever', 'North'), 'northern': ('Sever', 'North'), 'sever': ('Sever', 'North'),
    'south': ('Jug', 'South'), 'southern': ('Jug', 'South'), 'jug': ('Jug', 'South'),
    'a': ('A', 'A'), 'b': ('B', 'B'), 'c': ('C', 'C'),
}
_ZONE_ORD = {
    'first': 'I', '1': 'I', 'i': 'I',
    'second': 'II', '2': 'II', 'ii': 'II',
    'third': 'III', '3': 'III', 'iii': 'III',
    'fourth': 'IV', '4': 'IV', 'iv': 'IV',
    'fifth': 'V', '5': 'V', 'v': 'V',
}


def _direct_region_group(low: str):
    patterns = [
        r'\bgroup\s+(sever|istok|zapad|jug|north|south|east|west|eastern|western|northern|southern)\b',
        r'\b(?:second league|druga(?: savezna)? liga|ii(?:\s+savezna)?\s+liga)\s+(north|south|east|west|sever|jug|istok|zapad)\b',
        r'^group\s+(north|south|east|west|sever|jug|istok|zapad)$',
    ]
    for pat in patterns:
        m = re.search(pat, low)
        if m:
            return _GROUP_TRANSLATIONS[m.group(1)]
    if low in _GROUP_TRANSLATIONS and low not in ('a','b','c'):
        return _GROUP_TRANSLATIONS[low]
    if low in ('eastern group','western group','northern group','southern group'):
        return _GROUP_TRANSLATIONS[low.split()[0]]
    return None


def _zone_heading(low: str):
    # First Zone / Second Zone / I Zona / Zone IV
    m = re.search(r'^(first|second|third|fourth|fifth|[1-5]|i{1,3}|iv|v)\.?\s+(?:zone|zona)\b', low)
    if not m:
        m = re.search(r'^(?:zone|zona)\s+(first|second|third|fourth|fifth|[1-5]|i{1,3}|iv|v)\b', low)
    if not m:
        return None
    return _ZONE_ORD.get(m.group(1))


def _group_from_context(lines: list[str], marker: int, section_start: int):
    # RSSSF can place a full list of rounds/results between a group heading and
    # its final table.  Therefore look back much farther than a few dozen lines,
    # but stop at the previous table marker so a heading from the preceding
    # division cannot leak into the current one.
    zone = None
    letter = None
    lo = max(section_start, marker - 320)
    for j in range(marker-1, lo-1, -1):
        raw = _clean_line(lines[j])
        if not raw or len(raw) > 120:
            continue
        if _is_table_marker(raw):
            break
        low = _fold(raw).rstrip(':')
        if low.startswith(('nb:', 'promoted', 'relegated', 'round ', 'cross-table', 'cross table')):
            continue
        direct = _direct_region_group(low)
        if direct:
            return direct
        m = re.match(r'^group\s+([abc])$', low)
        if m and letter is None:
            letter = m.group(1).upper()
            continue
        z = _zone_heading(low)
        if z and zone is None:
            zone = z
            if letter:
                return (f'Zona {zone}-{letter}', f'Zone {zone}-{letter}')
            return (f'Zona {zone}', f'Zone {zone}')
    if letter:
        return (f'Grupa {letter}', f'Group {letter}')
    return None


def _row_blocks(lines: list[str], start: int, end: int):
    """Find credible standings blocks even when RSSSF has no Table marker.

    A number of 1970s/1980s pages print the group heading followed immediately
    by rows (and use colon-separated goals), so marker-only parsing necessarily
    creates holes.  Match runs of standings rows and tolerate blank/separator
    lines inside a block.
    """
    blocks = []
    i = max(0, start)
    while i < min(end, len(lines)):
        first = parse_row(lines[i])
        if not first:
            i += 1
            continue
        rows = [first]
        j = i + 1
        misses = 0
        while j < min(end, len(lines)):
            raw = _clean_line(lines[j])
            row = parse_row(lines[j])
            if row:
                rows.append(row); misses = 0; j += 1; continue
            low = _fold(raw)
            if not raw or set(raw) <= set('-=._ '):
                j += 1; continue
            if low.startswith(('nb:', 'promoted:', 'relegated:')):
                break
            # A new division/group heading terminates the current markerless
            # standings block.  Without this, East and West tables get merged
            # into one block because both restart at position 1.
            if re.search(r'^(?:ii(?:\s+savezna)?\s+liga|druga(?:\s+savezna)?\s+liga|group\s+|(?:eastern|western|northern|southern)\s+group|(?:first|second|third|fourth|fifth)\s+zone|(?:i{1,3}|iv|v)\.?\s+zona|zonal leagues)', low):
                break
            misses += 1
            if misses >= 2:
                break
            j += 1
        # positions should be mostly unique and a real division normally has 6+
        # teams.  Tiny playoff groups are handled separately, not as the league.
        # If this row-run is immediately below an explicit Table/Final Table
        # marker, marker-driven parsing already owns it; do not add a duplicate
        # markerless `unified` block.
        prev = i - 1
        while prev >= start and (not _clean_line(lines[prev]) or set(_clean_line(lines[prev])) <= set('-=._ ')):
            prev -= 1
        owned_by_marker = prev >= start and _is_table_marker(lines[prev])
        uniq = {r.position for r in rows}
        if len(rows) >= 6 and len(uniq) >= 6 and not owned_by_marker:
            blocks.append((i, rows))
            i = max(j, i+1)
        else:
            i += 1
    return blocks


def _parse_group_classifications(lines: list[str], start: int, end: int):
    found: dict[str, tuple[tuple[str,str] | None, list[L12Row]]] = {}

    # Marker-driven tables (the common case).
    for i in range(start, end):
        if not _is_table_marker(lines[i]):
            continue
        rows = _collect_table(lines[:end], i)
        if not rows:
            continue
        group = _group_from_context(lines, i, start)
        key = (group[1].lower() if group else 'unified')
        if key not in found or len(rows) > len(found[key][1]):
            found[key] = (group, rows)

    # Markerless historical tables.  This also catches `II SAVEZNA LIGA ISTOK`
    # pages and RSSSF's fixed-width colon-goal layouts.
    for marker, rows in _row_blocks(lines, start, end):
        group = _group_from_context(lines, marker, start)
        key = (group[1].lower() if group else 'unified')
        if key not in found or len(rows) > len(found[key][1]):
            found[key] = (group, rows)

    return list(found.values())

def parse_historical_page(html: str, page_year: int):
    lines = _html_to_lines(html)
    l1_start, l1_end = _find_historical_l1_range(lines)
    l1 = _parse_single_classification(lines, l1_start, l1_end)
    r2 = _find_historical_l2_range(lines, l1_end)
    l2 = _parse_group_classifications(lines, *r2) if r2 else []
    return {'season': _season_for_page_year(page_year), 'l1': l1, 'l2': l2}


def parse_modern_page(html: str, start_year: int):
    lines = _html_to_lines(html)
    r1 = _find_modern_l1_range(lines)
    r2 = _find_modern_l2_range(lines, r1[1] if r1 else 0)
    l1 = _parse_single_classification(lines, *r1) if r1 else []
    l2 = _parse_group_classifications(lines, *r2) if r2 else []
    return {'season': _modern_season(start_year), 'l1': l1, 'l2': l2}


def _year(season: str):
    m = re.match(r'^(\d{4})', str(season or ''))
    return int(m.group(1)) if m else 9999


def _l1_comp(con, season: str):
    y = _year(season)
    if y <= 1991:
        key = 'yugoslavia-first-1946-1992'
    elif y <= 2002:
        key = 'fr-yugoslavia-first-1992-2003'
    elif y <= 2005:
        key = 'serbia-montenegro-first-2003-2006'
    else:
        key = 'serbia-top-flight'
    row = con.execute('SELECT id FROM history_competitions WHERE history_key=?', (key,)).fetchone()
    return row[0] if row else None


def _slug(v: str):
    s = _fold(v)
    s = re.sub(r'[^a-z0-9]+', '-', s).strip('-')
    return s or 'group'


def _l2_base(season: str):
    y = _year(season)
    if y <= 1991:
        return ('yugoslavia-second-league', 'Druga savezna liga Jugoslavije', 'Yugoslav Second League', 36)
    if y <= 2002:
        return ('fr-yugoslavia-second-1992-2003', 'Druga liga SR Jugoslavije', 'FR Yugoslavia Second League', 33)
    if y <= 2005:
        return ('serbia-montenegro-second-2003-2006', 'Druga liga Srbije i Crne Gore', 'Serbia and Montenegro Second League', 34)
    return ('serbia-second-tier', 'Prva liga Srbije', 'Serbian First League', 20)


def _l2_comp(con, season: str, group):
    base_key, base_sr, base_en, sort = _l2_base(season)
    if not group:
        row = con.execute('SELECT id FROM history_competitions WHERE history_key=?', (base_key,)).fetchone()
        if row:
            return row[0]
        return hs._competition(con, base_key, base_sr, base_en, 'league', 2, None, sort, None, None, None)
    sr, en = group
    key = f'{base_key}-{_slug(en)}'
    row = con.execute('SELECT id FROM history_competitions WHERE history_key=?', (key,)).fetchone()
    if row:
        return row[0]
    return hs._competition(con, key, f'{base_sr} – {sr}', f'{base_en} – {en}', 'league', 2, None, sort+1, None, None, None)


def _plausible_table(rows: list[L12Row], min_rows=6):
    if len(rows) < min_rows:
        return False
    positions = [r.position for r in rows]
    if len(set(positions)) != len(positions):
        return False
    if min(positions) != 1:
        return False
    # Final tables should have positive games and sensible W+D+L totals.
    good = sum(1 for r in rows if r.played > 0 and (r.wins + r.draws + r.losses) >= max(1, r.played - 2))
    return good >= max(4, len(rows)//2)



def _expected_l2_group_count(season: str):
    """Historical number of parallel second-tier divisions.

    This is used as a completeness guard: a page that yields only one group in
    a four-group era is partial and must not replace a previously good season.
    1951/52 is the normalized label for the cancelled 1952 competition.
    1952/53 is a special republic/inter-republic transition and is handled as a
    documented exception rather than pretending the promotion playoff is the
    whole second tier.
    """
    y = _year(season)
    if season == '1951/52':
        return 0
    if season == '1952/53':
        return None
    if 1947 <= y <= 1950:
        return 1
    if 1953 <= y <= 1954:
        return 1
    if 1955 <= y <= 1957:
        return 5
    if 1958 <= y <= 1967:
        return 2
    if 1968 <= y <= 1972:
        return 4
    if 1973 <= y <= 1987:
        return 2
    if 1988 <= y <= 1995:
        return 1
    if 1996 <= y <= 1998:
        return 2
    if 1999 <= y <= 2003:
        return 3
    if y >= 2004:
        return 1
    return None


def _l2_special_note(season: str):
    if season == '1951/52':
        return ('Druga liga nije igrana — takmičenje 1952. je otkazano.',
                'Second League not played — the 1952 competition was cancelled.')
    if season == '1952/53':
        return ('Poseban prelazni sistem republičkih/međurepubličkih liga; federalni promotion playoff nije potpuna tabela drugog ranga.',
                'Special republic/inter-republic transition; the federal promotion playoff is not a complete second-tier table.')
    return None

def _expected_l1_champion(season: str):
    # Existing Touchline verified champion constants are intentionally partial in
    # Yugoslav years (they mostly list Serbian champions).  Use them when they do
    # contain the season; otherwise trust the locked RSSSF L1 section.
    m = dict(hs.YUGOSLAV_CHAMPIONS)
    m.update(dict(hs.FRY_CHAMPIONS))
    m.update(hs.SUPER_CHAMPIONS)
    return m.get(season)


def _same_identity(con, a: str, b: str):
    ai = hs._club_id(con, a)
    bi = hs._club_id(con, b)
    if ai is not None and bi is not None:
        return int(ai) == int(bi)
    return hs._norm_match(a) == hs._norm_match(b)


def _validate_l1(con, season: str, rows: list[L12Row]):
    if not _plausible_table(rows):
        return False, 'neuverljiva L1 tabela'
    expected = _expected_l1_champion(season)
    if expected and not _same_identity(con, rows[0].club_name, expected):
        return False, f'prvak {rows[0].club_name}, očekivano {expected}'
    return True, ''


def _replace_level_season(con, level: int, season: str):
    # Central tables are the source-of-truth for league rows.  Cups/Europe are
    # separate comp_types and are therefore untouched.
    con.execute('''DELETE FROM history_club_seasons
        WHERE season_label=? AND competition_id IN (
            SELECT id FROM history_competitions WHERE comp_type='league' AND level=?
        )''', (season, int(level)))


def _write_table(con, comp_id: int, season: str, rows: list[L12Row], url: str):
    matched = unmatched = 0
    for r in rows:
        tail = _fold(r.tail)
        _, cid = hs._upsert_history(
            con, comp_id, season, r.club_name,
            position=r.position, played=r.played, wins=r.wins, draws=r.draws,
            losses=r.losses, gf=r.gf, ga=r.ga, points=r.points,
            champion=1 if r.position == 1 else 0,
            runner_up=1 if r.position == 2 else 0,
            promoted=1 if 'promot' in tail else 0,
            relegated=1 if 'relegat' in tail else 0,
            highlight=1 if r.position <= 2 else 0,
            source_name='RSSSF', source_url=url, source_note=CENTRAL_NOTE,
            fill_only=False
        )
        if cid:
            matched += 1
        else:
            unmatched += 1
    return matched, unmatched


def sync_all(refresh: bool = False, historical_from: int = 1947, historical_to: int = 2006,
             modern_from: int = 2006, modern_to: int = 2025):
    """Disabled in 1.6.4: domestic league history is Excel-only.

    The parser is intentionally kept in the source tree for forensic/reference
    purposes, but it must never rebuild public league history again.
    """
    return {
        'disabled': 1,
        'reason': 'Domestic league history is locked to HISTORY/TOUCHLINE_HISTORY_MASTER.xlsx',
        'source': 'Touchline History Master Excel',
    }


def _sync_all_rsssf_legacy(refresh: bool = False, historical_from: int = 1947, historical_to: int = 2006,
             modern_from: int = 2006, modern_to: int = 2025):
    """Legacy RSSSF rebuild retained only as unreachable reference code."""
    hs.seed_history()
    con = connect()
    hs.ensure_history_schema(con)
    # Genuine second-tier historical exceptions: show them explicitly instead
    # of leaving apparent holes in the season list.
    for _season in ('1951/52','1952/53'):
        _notes = _l2_special_note(_season)
        if _notes:
            con.execute("""INSERT INTO history_season_notes(level,season_label,note_sr,note_en,source_url,updated_at)
                VALUES(2,?,?,?,?,CURRENT_TIMESTAMP)
                ON CONFLICT(level,season_label) DO UPDATE SET note_sr=excluded.note_sr,note_en=excluded.note_en,source_url=excluded.source_url,updated_at=CURRENT_TIMESTAMP""",
                (_season,_notes[0],_notes[1],_joeg_url(int(_season[:4])+1)))
    con.commit()
    out = {
        'l1_seasons': 0, 'l1_rows': 0, 'l2_seasons': 0, 'l2_groups': 0,
        'l2_rows': 0, 'matched': 0, 'unmatched': 0, 'pages': 0,
        'errors': [], 'missing_l1': [], 'missing_l2': []
    }
    session = requests.Session()

    def process(parsed, url):
        season = parsed['season']
        l1 = parsed.get('l1') or []
        l2 = parsed.get('l2') or []

        if l1:
            ok, why = _validate_l1(con, season, l1)
            if ok:
                comp = _l1_comp(con, season)
                if comp:
                    _replace_level_season(con, 1, season)
                    ma, un = _write_table(con, comp, season, l1, url)
                    out['l1_seasons'] += 1; out['l1_rows'] += len(l1)
                    out['matched'] += ma; out['unmatched'] += un
                else:
                    out['errors'].append(f'{season} L1: competition missing')
            else:
                out['errors'].append(f'{season} L1 ODBAČEN: {why}')
                out['missing_l1'].append(season)
        else:
            out['missing_l1'].append(season)

        credible_groups = [(g, rows) for g, rows in l2 if _plausible_table(rows)]
        expected_groups = _expected_l2_group_count(season)
        complete = bool(credible_groups)
        if expected_groups is not None:
            complete = (len(credible_groups) >= expected_groups) if expected_groups > 0 else False

        if complete:
            # Only delete previous L2 rows after we know the season has the
            # expected number of central second-tier divisions.  This prevents a
            # single parsed West group from erasing East/North/South history.
            _replace_level_season(con, 2, season)
            for group, rows in credible_groups:
                comp = _l2_comp(con, season, group)
                ma, un = _write_table(con, comp, season, rows, url)
                out['l2_groups'] += 1; out['l2_rows'] += len(rows)
                out['matched'] += ma; out['unmatched'] += un
            out['l2_seasons'] += 1
        else:
            # 1946/47 had no regular federal L2.  1951/52 (calendar 1952) was
            # cancelled, and 1952/53 used a special republic/inter-republic
            # structure whose promotion playoff is not a full second-tier table.
            if season not in ('1946/47','1951/52'):
                out['missing_l2'].append(season)
            if expected_groups and credible_groups:
                out['errors'].append(f'{season} L2 DELIMIČNO: {len(credible_groups)}/{expected_groups} grupa')
        hs._reconcile_history_club_links(con)
        hs._remove_shadow_history_rows(con)
        con.commit()

    for page_year in range(int(historical_from), int(historical_to)+1):
        url = _joeg_url(page_year)
        print(f'[HISTORY L1/L2] {_season_for_page_year(page_year)} RSSSF...', flush=True)
        try:
            html = _fetch_html(session, url, refresh=refresh)
            process(parse_historical_page(html, page_year), url)
            out['pages'] += 1
            time.sleep(0.10)
        except Exception as e:
            out['errors'].append(f'{_season_for_page_year(page_year)}: {e}')

    for start_year in range(int(modern_from), int(modern_to)+1):
        url = _serv_url(start_year)
        print(f'[HISTORY L1/L2] {_modern_season(start_year)} RSSSF...', flush=True)
        try:
            html = _fetch_html(session, url, refresh=refresh)
            process(parse_modern_page(html, start_year), url)
            out['pages'] += 1
            time.sleep(0.10)
        except Exception as e:
            out['errors'].append(f'{_modern_season(start_year)}: {e}')

    con.execute("INSERT OR REPLACE INTO history_meta(key,value,updated_at) VALUES('l12_central_version','1.5.3',CURRENT_TIMESTAMP)")
    con.execute("INSERT OR REPLACE INTO history_meta(key,value,updated_at) VALUES('l12_last_summary',?,CURRENT_TIMESTAMP)",
                (f"L1 {out['l1_seasons']} seasons/{out['l1_rows']} rows; L2 {out['l2_seasons']} seasons/{out['l2_groups']} groups/{out['l2_rows']} rows",))
    con.commit(); con.close()
    return out


def validate_current_clubs():
    """Return coverage report for current men's L1/L2 clubs."""
    con = connect()
    hs.ensure_history_schema(con)
    clubs = con.execute('''SELECT c.id,c.name,l.name league_name,l.level
        FROM clubs c JOIN leagues l ON l.id=c.league_id
        WHERE l.section='senior' AND l.level IN (1,2) AND COALESCE(c.source_active,1)=1
        ORDER BY l.level,c.name''').fetchall()
    rows = []
    for c in clubs:
        seasons = con.execute('''SELECT COUNT(DISTINCT h.season_label) n,
            MIN(h.season_label) first_season,MAX(h.season_label) last_season
            FROM history_club_seasons h JOIN history_competitions hc ON hc.id=h.competition_id
            WHERE h.club_id=? AND hc.comp_type='league' AND hc.level IN (1,2)''', (c['id'],)).fetchone()
        rows.append({
            'club_id': c['id'], 'club': c['name'], 'current_level': c['level'],
            'seasons': int(seasons['n'] or 0), 'first': seasons['first_season'],
            'last': seasons['last_season']
        })
    meta = con.execute("SELECT value FROM history_meta WHERE key='l12_last_summary'").fetchone()
    con.close()
    return {'summary': meta['value'] if meta else 'L1-L2 central sync još nije pokrenut', 'clubs': rows}


# ---------------------------------------------------------------------------
# Small offline parser fixtures used by SELF_TEST.  They intentionally cover a
# unified modern table, a historical split second tier and a playoff top flight.
# ---------------------------------------------------------------------------
FIXTURE_1960 = '''
Prva savezna liga 1960/61
Final Table:
 1.Partizan (Beograd) 22 15 2 5 53-23 32 Champions
 2.Crvena zvezda (Beograd) 22 13 5 4 38-21 31
 3.Hajduk (Split) 22 13 4 5 34-22 30
 4.Dinamo (Zagreb) 22 10 7 5 36-27 27
 5.Vojvodina (Novi Sad) 22 10 3 9 32-29 23
 6.OFK Beograd 22 8 7 7 30-30 23
 7.Rijeka 22 10 2 10 32-36 22
 8.Sarajevo 22 6 6 10 33-39 18
Kup 1960/61
Druga savezna liga 1960/61
Eastern Group
Final Table:
 1.Novi Sad 22 15 2 5 48-21 32 Promoted
 2.Buducnost 22 12 5 5 39-21 29
 3.Radnicki Sombor 22 12 3 7 43-30 27
 4.Srem 22 9 5 8 37-35 23
 5.Radnicki Nis 22 8 7 7 36-36 23
 6.Spartak 22 5 11 6 21-22 21
Western Group
Final Table:
 1.Radnicki Split 22 16 2 4 60-24 34 Promoted
 2.Tresnjevka 22 13 3 6 43-29 29
 3.Borac Banja Luka 22 13 2 7 55-34 28
 4.Zeljeznicar Sarajevo 22 11 3 8 33-29 25
 5.Sibenik 22 11 2 9 44-36 24
 6.Zagreb 22 8 3 11 46-54 19
'''

FIXTURE_2016 = '''
Jelen Superliga
Regular Stage
Final Table:
 1.Crvena zvezda (Beograd) 30 25 4 1 75-25 79 Championship Playoff
 2.Partizan (Beograd) 30 23 4 3 59-17 73 Championship Playoff
 3.Vojvodina (Novi Sad) 30 18 5 7 51-26 59 Championship Playoff
 4.Napredak (Krusevac) 30 15 7 8 35-23 52 Championship Playoff
 5.Mladost (Lucani) 30 14 6 10 35-29 48 Championship Playoff
 6.Radnicki (Nis) 30 12 8 10 37-36 44 Championship Playoff
 7.Vozdovac 30 13 4 13 35-38 43 Championship Playoff
 8.Javor 30 11 9 10 32-37 42 Championship Playoff
 9.Cukaricki 30 11 7 12 40-42 40 Relegation Playoff
Championship Playoff
Final Table:
 1.Partizan (Beograd) 37 30 4 3 78-22 58 Champions
 2.Crvena zvezda (Beograd) 37 30 4 3 93-33 55
 3.Vojvodina (Novi Sad) 37 22 6 9 58-31 43
 4.Mladost (Lucani) 37 18 6 13 46-44 36
 5.Radnicki (Nis) 37 15 9 13 47-46 32
 6.Napredak (Krusevac) 37 16 8 13 44-36 30
 7.Vozdovac 37 14 6 17 41-51 27
 8.Javor 37 11 10 16 34-50 22
Relegation Playoff
Final Table:
 9.Cukaricki 37 15 7 15 54-49 32
 10.Spartak 37 14 9 14 47-55 32
 11.Rad 37 11 9 17 29-45 25
 12.Backa 37 9 7 21 32-53 20
 13.Radnik 37 8 9 20 30-55 19
 14.Metalac 37 7 10 20 25-50 18
 15.Novi Pazar 37 6 8 23 28-66 14
 16.Borac Cacak 37 5 8 24 24-61 12
Kup Srbije
Prva Liga
Table:
 1.Macva Sabac 30 19 6 5 48-25 63 Promoted
 2.Zemun 30 17 8 5 47-24 59 Promoted
 3.Sloboda Uzice 30 15 7 8 43-30 52
 4.Indjija 30 14 8 8 40-30 50
 5.Dinamo Vranje 30 13 8 9 38-31 47
 6.CSK Celarevo 30 12 9 9 35-30 45
'''
