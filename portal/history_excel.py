from __future__ import annotations

import hashlib
import re
import unicodedata
from pathlib import Path
from typing import Iterable

from openpyxl import load_workbook

from .config import PROJECT_DIR
from .data import connect
from . import history_sync as hs

MASTER_PATH = PROJECT_DIR / 'HISTORY' / 'TOUCHLINE_HISTORY_MASTER.xlsx'
SOURCE_NAME = 'Touchline History Master Excel'
MASTER_VERSION = 'v6-complete-top-flight-excel-only'

CYR_TO_LAT = str.maketrans({
    'А':'A','Б':'B','В':'V','Г':'G','Д':'D','Ђ':'Đ','Е':'E','Ж':'Ž','З':'Z','И':'I','Ј':'J','К':'K','Л':'L','Љ':'Lj','М':'M','Н':'N','Њ':'Nj','О':'O','П':'P','Р':'R','С':'S','Т':'T','Ћ':'Ć','У':'U','Ф':'F','Х':'H','Ц':'C','Ч':'Č','Џ':'Dž','Ш':'Š',
    'а':'a','б':'b','в':'v','г':'g','д':'d','ђ':'đ','е':'e','ж':'ž','з':'z','и':'i','ј':'j','к':'k','л':'l','љ':'lj','м':'m','н':'n','њ':'nj','о':'o','п':'p','р':'r','с':'s','т':'t','ћ':'ć','у':'u','ф':'f','х':'h','ц':'c','ч':'č','џ':'dž','ш':'š',
})

CANONICAL_L3 = ('Srpska liga Beograd','Srpska liga Vojvodina','Srpska liga Istok','Srpska liga Zapad')


LEAGUE_SOURCE_POLICY = 'Touchline History Master Excel'


def _enforce_excel_only_league_history(con):
    """Remove every historical league row not originating from the master Excel.

    This affects only history_club_seasons / history_all_time_cache. Current/live
    standings, memberships, matches, players, staff and transfers are separate
    data layers and are intentionally untouched.
    """
    cur=con.execute("""DELETE FROM history_club_seasons
        WHERE competition_id IN (SELECT id FROM history_competitions WHERE comp_type='league')
          AND COALESCE(source_name,'')<>?""",(LEAGUE_SOURCE_POLICY,))
    removed=max(cur.rowcount or 0,0)
    # The old all-time cache was RSSSF-derived. It is hidden in the public UI and
    # must not survive an Excel-only league-history policy.
    con.execute("DELETE FROM history_all_time_cache WHERE competition_id IN (SELECT id FROM history_competitions WHERE comp_type='league')")
    con.execute("UPDATE history_competitions SET source_name=? WHERE comp_type='league'",(LEAGUE_SOURCE_POLICY,))
    return removed


def _reset_all_historical_league_rows(con):
    """Rebuild the historical domestic league layer from Excel on every import.

    Excel is the sole authority. Deleting all league rows before import prevents
    stale rows from older Excel versions or legacy syncs from surviving forever.
    Cups and Europe/world rows are not touched.
    """
    cur=con.execute("""DELETE FROM history_club_seasons
        WHERE competition_id IN (SELECT id FROM history_competitions WHERE comp_type='league')""")
    removed=max(cur.rowcount or 0,0)
    con.execute("DELETE FROM history_all_time_cache WHERE competition_id IN (SELECT id FROM history_competitions WHERE comp_type='league')")
    return removed


def _assert_single_excel_league_per_club_season(con):
    """Fail closed if Excel mapping would put one club in two leagues in one season."""
    rows=con.execute("""SELECT h.club_id,h.season_label,COUNT(DISTINCT h.competition_id) n,
               GROUP_CONCAT(hc.name_sr,' | ') comps
        FROM history_club_seasons h
        JOIN history_competitions hc ON hc.id=h.competition_id
        WHERE hc.comp_type='league' AND h.source_name=? AND h.club_id IS NOT NULL
        GROUP BY h.club_id,h.season_label
        HAVING COUNT(DISTINCT h.competition_id)>1
        ORDER BY h.season_label,h.club_id""",(LEAGUE_SOURCE_POLICY,)).fetchall()
    if rows:
        sample='; '.join(f"club_id={r['club_id']} {r['season_label']}: {r['comps']}" for r in rows[:10])
        raise RuntimeError('Excel league conflict: isti klub je mapiran u više liga iste sezone. '+sample)

    # Also fail closed for still-unlinked historical names. This prevents the
    # same Excel spelling from silently appearing in two domestic leagues in
    # the same season before it is linked to a master club_id.
    rows=con.execute("""SELECT lower(trim(h.club_name_override)) club_name,h.season_label,
               COUNT(DISTINCT h.competition_id) n,GROUP_CONCAT(hc.name_sr,' | ') comps
        FROM history_club_seasons h
        JOIN history_competitions hc ON hc.id=h.competition_id
        WHERE hc.comp_type='league' AND h.source_name=? AND h.club_id IS NULL
          AND trim(COALESCE(h.club_name_override,''))<>''
        GROUP BY lower(trim(h.club_name_override)),h.season_label
        HAVING COUNT(DISTINCT h.competition_id)>1
        ORDER BY h.season_label,club_name""",(LEAGUE_SOURCE_POLICY,)).fetchall()
    if rows:
        sample='; '.join(f"{r['club_name']} {r['season_label']}: {r['comps']}" for r in rows[:10])
        raise RuntimeError('Excel league conflict: isti istorijski naziv je u više liga iste sezone. '+sample)
    return 0


def _install_excel_only_league_guards(con):
    """Install database-level guards so the policy cannot be bypassed later."""
    # These triggers intentionally protect only the historical league layer.
    # Cups/Europe/world and all live/current tables remain unaffected.
    con.executescript(f"""
    DROP TRIGGER IF EXISTS trg_history_league_excel_only_insert;
    CREATE TRIGGER trg_history_league_excel_only_insert
    BEFORE INSERT ON history_club_seasons
    WHEN EXISTS (SELECT 1 FROM history_competitions hc
                 WHERE hc.id=NEW.competition_id AND hc.comp_type='league')
         AND COALESCE(NEW.source_name,'')<>'{LEAGUE_SOURCE_POLICY}'
    BEGIN
      SELECT RAISE(ABORT,'Domestic league history is Excel-only');
    END;

    DROP TRIGGER IF EXISTS trg_history_league_excel_only_update;
    CREATE TRIGGER trg_history_league_excel_only_update
    BEFORE UPDATE OF competition_id,source_name ON history_club_seasons
    WHEN EXISTS (SELECT 1 FROM history_competitions hc
                 WHERE hc.id=NEW.competition_id AND hc.comp_type='league')
         AND COALESCE(NEW.source_name,'')<>'{LEAGUE_SOURCE_POLICY}'
    BEGIN
      SELECT RAISE(ABORT,'Domestic league history is Excel-only');
    END;

    DROP TRIGGER IF EXISTS trg_history_one_league_per_club_season_insert;
    CREATE TRIGGER trg_history_one_league_per_club_season_insert
    BEFORE INSERT ON history_club_seasons
    WHEN EXISTS (SELECT 1 FROM history_competitions hc
                 WHERE hc.id=NEW.competition_id AND hc.comp_type='league')
         AND NEW.club_id IS NOT NULL
         AND EXISTS (SELECT 1
                     FROM history_club_seasons h
                     JOIN history_competitions hc2 ON hc2.id=h.competition_id
                     WHERE hc2.comp_type='league'
                       AND h.club_id=NEW.club_id
                       AND h.season_label=NEW.season_label
                       AND h.competition_id<>NEW.competition_id)
    BEGIN
      SELECT RAISE(ABORT,'One club can have only one domestic league per season');
    END;

    DROP TRIGGER IF EXISTS trg_history_one_league_per_club_season_update;
    CREATE TRIGGER trg_history_one_league_per_club_season_update
    BEFORE UPDATE OF club_id,competition_id,season_label ON history_club_seasons
    WHEN EXISTS (SELECT 1 FROM history_competitions hc
                 WHERE hc.id=NEW.competition_id AND hc.comp_type='league')
         AND NEW.club_id IS NOT NULL
         AND EXISTS (SELECT 1
                     FROM history_club_seasons h
                     JOIN history_competitions hc2 ON hc2.id=h.competition_id
                     WHERE hc2.comp_type='league'
                       AND h.club_id=NEW.club_id
                       AND h.season_label=NEW.season_label
                       AND h.competition_id<>NEW.competition_id
                       AND h.id<>OLD.id)
    BEGIN
      SELECT RAISE(ABORT,'One club can have only one domestic league per season');
    END;
    """)
    return 4


# Public/history DB convention: one row per football season. Some Yugoslav Cup
# editions were published by calendar year; only transitional cup labels whose
# matching league season is split-year are folded. 1950/1951/1952 stay as
# calendar seasons because the league itself used calendar seasons then.
LEGACY_CUP_SEASON_MAP = {
    '1947':'1946/47',
    '1948':'1947/48',
    '1949':'1948/49',
    '1953':'1952/53',
    '1954':'1953/54',
    '1955':'1954/55',
    '1973':'1972/73',
}


def _canonical_cup_season(value):
    label=str(value or '').strip()
    return LEGACY_CUP_SEASON_MAP.get(label,label)


def latin(value):
    return str(value or '').translate(CYR_TO_LAT).strip()


def _norm(value):
    s=latin(value).lower().replace('đ','dj')
    s=unicodedata.normalize('NFKD',s)
    s=''.join(ch for ch in s if not unicodedata.combining(ch))
    s=re.sub(r'\b(fk|rfk|ofk|nk|sk|fkz)\b',' ',s)
    s=re.sub(r'[^a-z0-9]+',' ',s)
    return re.sub(r'\s+',' ',s).strip()


def _slug(value):
    return re.sub(r'[^a-z0-9]+','-',_norm(value)).strip('-') or 'all'


def _club_norm(value):
    """Conservative club-name normalization for identity matching.

    Unlike generic _norm(), keep identity-bearing prefixes such as OFK/BSK/NK.
    Removing OFK caused historical 'Beograd' to collide with OFK Beograd.
    """
    s=latin(value).lower().replace('đ','dj')
    s=unicodedata.normalize('NFKD',s)
    s=''.join(ch for ch in s if not unicodedata.combining(ch))
    s=re.sub(r'\bfk\b',' ',s)
    s=re.sub(r'[^a-z0-9]+',' ',s)
    return re.sub(r'\s+',' ',s).strip()


def _sha256(path:Path):
    h=hashlib.sha256()
    with path.open('rb') as f:
        for chunk in iter(lambda:f.read(1024*1024),b''):
            h.update(chunk)
    return h.hexdigest()


def _sheet_rows(wb,name):
    if name not in wb.sheetnames:
        raise RuntimeError(f'Nedostaje obavezni sheet: {name}')
    ws=wb[name]
    rows=list(ws.iter_rows(values_only=True))
    if not rows:
        return [],[]
    headers=[str(x or '').strip() for x in rows[0]]
    out=[]
    for row in rows[1:]:
        if not any(v is not None and str(v).strip() for v in row):
            continue
        out.append({headers[i]:row[i] if i<len(row) else None for i in range(len(headers))})
    return headers,out


def _current_league_id(con,name):
    r=con.execute("SELECT id FROM leagues WHERE country_id=1 AND section='senior' AND name=? ORDER BY id LIMIT 1",(name,)).fetchone()
    return int(r[0]) if r else None


def repair_competition_scope(con):
    """Keep national L1/L2 and the four Serbian Leagues in their canonical scopes.

    Older live-discovery builds could leave duplicate 'Mozzart Bet ...' rows at
    level 3. They remain in the database for forensic safety but are inactive.
    """
    changed=0
    for name,level in (('SuperLiga Srbije',1),('Prva liga Srbije',2)):
        cur=con.execute("UPDATE leagues SET level=?,section='senior',source_active=1 WHERE country_id=1 AND name=?",(level,name))
        changed+=max(cur.rowcount or 0,0)
    q=','.join('?' for _ in CANONICAL_L3)
    cur=con.execute(f"UPDATE leagues SET source_active=0 WHERE country_id=1 AND COALESCE(section,'senior')='senior' AND level=3 AND name NOT IN ({q})",CANONICAL_L3)
    changed+=max(cur.rowcount or 0,0)
    for name in CANONICAL_L3:
        cur=con.execute("UPDATE leagues SET level=3,section='senior',source_active=1 WHERE country_id=1 AND name=?",(name,))
        changed+=max(cur.rowcount or 0,0)
    # Explicitly retire leaked national duplicates even if an old build gave them another level.
    cur=con.execute("""UPDATE leagues SET source_active=0
        WHERE country_id=1 AND source_type='srbijasport'
          AND lower(name) IN ('mozzart bet super liga srbije','mozzart bet prva liga srbije')""")
    changed+=max(cur.rowcount or 0,0)
    return changed


def _club_index(con):
    idx={}
    rows=con.execute("""SELECT c.id,c.name,c.cm_name,c.tm_name,COALESCE(l.level,99) level,
        COALESCE(l.section,'senior') section,COALESCE(c.source_active,1) active
        FROM clubs c LEFT JOIN leagues l ON l.id=c.league_id
        ORDER BY CASE WHEN COALESCE(l.section,'senior')='senior' THEN 0 ELSE 1 END,
                 COALESCE(c.source_active,1) DESC,COALESCE(l.level,99),c.id""").fetchall()
    for r in rows:
        priority=(0 if (r['section'] or 'senior')=='senior' else 1, 0 if r['active'] else 1, int(r['level'] or 99), int(r['id']))
        for raw in (r['name'],r['cm_name'],r['tm_name']):
            n=_club_norm(raw)
            if n:
                idx.setdefault(n,[]).append((priority,int(r['id'])))
    return idx


FIXED_ALIASES={
    'crvena zvezda':'red star belgrade', 'crvena zvezda beograd':'red star belgrade',
    'partizan':'partizan belgrade', 'partizan beograd':'partizan belgrade',
    'vojvodina':'vojvodina novi sad',
    'javor ivanjica':'javor matis', 'javor':'javor matis',
    'borac cacak':'borac 1926',
    'radnicki kragujevac':'radnicki 1923',
    'tsc':'tsc backa topola',
}

def _club_id_from_candidates(idx,candidates:Iterable[str]):
    for raw in candidates:
        n=_club_norm(raw)
        if not n:
            continue
        n=FIXED_ALIASES.get(n,n)
        items=idx.get(n,[])
        if not items:
            continue
        # Prefer the highest active senior level; only accept ties when they point
        # to the same master club. This avoids merging two different 'Mladost' etc.
        items=sorted(items,key=lambda x:x[0])
        best_pri=items[0][0][:3]
        best=sorted(set(cid for pri,cid in items if pri[:3]==best_pri))
        if len(best)==1:
            return best[0]
    return None


def _display_name(name,place=None):
    name=latin(name)
    place=latin(place)
    if not name:
        return ''
    if place and _norm(place) not in _norm(name):
        return f'{name} ({place})'
    return name


def _upsert_excel(con,comp_id,season,name,position,champion,runner_up,source_url,note='',place=None,wiki_title=None,stage_sr=None,
                  played=None,wins=None,draws=None,losses=None,gf=None,ga=None,points=None,promoted=0,relegated=0):
    display=_display_name(name,place)
    candidates=[wiki_title,display,name]
    cid=_club_id_from_candidates(_club_index_cache,candidates)
    status,_=hs._upsert_history(
        con,comp_id,str(season),display,position=position,
        played=played,wins=wins,draws=draws,losses=losses,gf=gf,ga=ga,points=points,
        champion=1 if champion else 0,runner_up=1 if runner_up else 0,
        promoted=1 if promoted else 0,relegated=1 if relegated else 0,
        highlight=1 if champion else 0,stage_sr=stage_sr,
        source_name=SOURCE_NAME,source_url=source_url,source_note=note,
        fill_only=False,forced_club_id=cid)
    # _upsert_history chooses its own override when forced_club_id is None.
    # Make sure historical display is the Excel master spelling for unmatched rows.
    if not cid:
        con.execute("""UPDATE history_club_seasons SET club_name_override=?
            WHERE competition_id=? AND season_label=? AND club_id IS NULL
              AND lower(COALESCE(club_name_override,''))=lower(?)""",
            (display,comp_id,str(season),hs._canonical_source_name(display)))
    return status,cid


def _ensure_comp(con,key,name_sr,name_en,ctype,level,current_lid,sort_order,active_from=None,active_to=None,source_url=None):
    return hs._competition(con,key,name_sr,name_en,ctype,level,current_lid,sort_order,active_from,active_to,source_url)


def _l1_spec(comp):
    n=_norm(comp)
    if 'prvenstvo jugoslavije' in n:
        return ('kingdom-yugoslavia-first-1923-1940','Prvenstvo Kraljevine SHS / Jugoslavije','Kingdom of Yugoslavia Championship',1)
    if 'prva savezna liga jugoslavije' in n:
        return ('yugoslavia-first-1946-1992','Prva savezna liga Jugoslavije','Yugoslav First League',1)
    if 'prva liga sr jugoslavije' in n:
        return ('fr-yugoslavia-first-1992-2003','Prva liga SR Jugoslavije','FR Yugoslavia First League',1)
    if 'prva liga srbije i crne gore' in n:
        return ('serbia-montenegro-first-2003-2006','Prva liga Srbije i Crne Gore','Serbia and Montenegro First League',1)
    if 'superliga srbije' in n:
        return ('serbia-top-flight','SuperLiga Srbije','Serbian SuperLiga',1)
    return None


def _cup_spec(comp):
    n=_norm(comp)
    if n=='kup jugoslavije': return ('yugoslavia-cup','Kup Jugoslavije','Yugoslav Cup')
    if 'kup sr jugoslavije' in n: return ('fr-yugoslavia-cup','Kup SR Jugoslavije','FR Yugoslavia Cup')
    if 'kup srbije i crne gore' in n: return ('serbia-montenegro-cup','Kup Srbije i Crne Gore','Serbia and Montenegro Cup')
    if n=='kup srbije': return ('serbia-cup','Kup Srbije','Serbian Cup')
    return None


def _l2_base(comp):
    n=_norm(comp)
    if 'druga savezna liga jugoslavije' in n:
        return ('yugoslavia-second-league','Druga savezna liga Jugoslavije','Yugoslav Second League')
    if 'druga liga sr jugoslavije' in n:
        return ('fr-yugoslavia-second-1992-2003','Druga liga SR Jugoslavije','FR Yugoslavia Second League')
    if 'drugi rang srbije i crne gore' in n:
        return ('serbia-montenegro-second-2003-2006','Drugi rang Srbije i Crne Gore','Serbia and Montenegro Second Tier')
    if 'prva liga srbije' in n:
        return ('serbia-second-tier','Prva liga Srbije','Serbian First League')
    return None


def _l3_spec(comp):
    n=_norm(comp)
    mapping={
        'srpska liga beograd':('serbia-third-beograd','Srpska liga Beograd','Serbian League Belgrade'),
        'srpska liga vojvodina':('serbia-third-vojvodina','Srpska liga Vojvodina','Serbian League Vojvodina'),
        'srpska liga istok':('serbia-third-istok','Srpska liga Istok','Serbian League East'),
        'srpska liga zapad':('serbia-third-zapad','Srpska liga Zapad','Serbian League West'),
    }
    return mapping.get(n)


def _reset_authority(con,comp_ids,seasons=None):
    if not comp_ids: return
    q=','.join('?' for _ in comp_ids)
    params=list(comp_ids)
    sql=f"UPDATE history_club_seasons SET is_champion=0,is_runner_up=0 WHERE competition_id IN ({q})"
    if seasons:
        sq=','.join('?' for _ in seasons)
        sql+=f" AND season_label IN ({sq})"; params+=list(seasons)
    con.execute(sql,params)


def import_master(path:Path|str=MASTER_PATH,force=True):
    global _club_index_cache
    path=Path(path)
    if not path.exists():
        raise RuntimeError(f'History Excel nije pronađen: {path}')
    digest=_sha256(path)
    con=connect(); hs.ensure_history_schema(con); repair_competition_scope(con)
    removed_non_excel_league=_enforce_excel_only_league_history(con)
    guards_installed=_install_excel_only_league_guards(con)
    con.commit()
    if not force:
        row=con.execute("SELECT value FROM history_meta WHERE key='excel_master_hash'").fetchone()
        if row and row['value']==digest:
            con.close(); return {'skipped':1,'hash':digest,'removed_non_excel_league':removed_non_excel_league,'league_source':'Excel only','db_guards':guards_installed}

    wb=load_workbook(path,read_only=True,data_only=True)
    required=('PRVENSTVO_1945+','PRVENSTVO_KOMPLETNE_TABELE','PRVENSTVO_PRE_1945','KUP','DRUGI_RANG_1946+','SRPSKE_LIGE_3_RANG')
    for s in required:
        if s not in wb.sheetnames: raise RuntimeError(f'Nedostaje sheet {s}')
    super_lid=_current_league_id(con,'SuperLiga Srbije')
    prva_lid=_current_league_id(con,'Prva liga Srbije')
    l3_lids={name:_current_league_id(con,name) for name in CANONICAL_L3}
    counts={'inserted':0,'updated':0,'linked':0,'unlinked':0,'l1':0,'l1_rows':0,'pre':0,'cup':0,'l2':0,'l3':0,'notes':0,'scope_repairs':0,'removed_non_excel_league':removed_non_excel_league,'removed_prior_league_rows':0,'db_guards':guards_installed}
    counts['scope_repairs']=repair_competition_scope(con)
    _club_index_cache=_club_index(con)
    # FINAL FREEZE: every import is a clean rebuild of domestic league history.
    counts['removed_prior_league_rows']=_reset_all_historical_league_rows(con)

    # 1.6.3 integrity repair: old builds incorrectly represented Vojvodina's
    # 1976 first place in an Intertoto/Rapan group as an overall Intertoto Cup
    # title. Remove that synthetic honour before rebuilding authoritative data.
    intertoto=con.execute("SELECT id FROM history_competitions WHERE history_key='intertoto-cup'").fetchone()
    voj=hs._club_id(con,'Vojvodina (Novi Sad)')
    if intertoto and voj:
        con.execute("DELETE FROM history_club_seasons WHERE competition_id=? AND club_id=? AND season_label='1975/76' AND (COALESCE(stage_sr,'') LIKE '%Intertoto%' OR COALESCE(is_champion,0)=1)",(intertoto[0],voj))

    # Remove obsolete normalized calendar-year rows created by older web-seed
    # builds. The Excel master intentionally keeps the historically correct
    # 1950/1951/1952 calendar seasons and calendar-year cups.
    old_first=con.execute("SELECT id FROM history_competitions WHERE history_key='yugoslavia-first-1946-1992'").fetchone()
    if old_first:
        con.execute("DELETE FROM history_club_seasons WHERE competition_id=? AND season_label IN ('1949/50','1950/51','1951/52') AND source_name<>?",(old_first[0],SOURCE_NAME))
    old_cup=con.execute("SELECT id FROM history_competitions WHERE history_key='yugoslavia-cup'").fetchone()
    if old_cup:
        # Delete both legacy calendar labels and canonical targets, then let the
        # Excel master repopulate authoritative winner/finalist rows once.
        cup_labels=sorted(set(LEGACY_CUP_SEASON_MAP)|set(LEGACY_CUP_SEASON_MAP.values()))
        marks=','.join('?' for _ in cup_labels)
        con.execute(f"DELETE FROM history_club_seasons WHERE competition_id=? AND season_label IN ({marks})",[old_cup[0],*cup_labels])

    # League rows outside the Excel master were removed above. The workbook is
    # the sole authority for domestic historical league rows and medal/title flags.
    comp_ids=[]; season_map={}

    def rows(name):
        ws=wb[name]; it=ws.iter_rows(values_only=True); headers=[str(x or '').strip() for x in next(it)]
        for rr in it:
            if not any(v is not None and str(v).strip() for v in rr): continue
            yield {headers[i]:rr[i] if i<len(rr) else None for i in range(len(headers))}

    # Pre-war first tier.
    pre_comp=None
    for r in rows('PRVENSTVO_PRE_1945'):
        spec=_l1_spec(r.get('Takmičenje'))
        if not spec: continue
        if pre_comp is None:
            pre_comp=_ensure_comp(con,spec[0],spec[1],spec[2],'league',1,super_lid,1,'1923','1939/40',r.get('Izvor URL'))
            comp_ids.append(pre_comp)
        season=str(r.get('Sezona') or '').strip(); status=str(r.get('Status') or '')
        if status!='FINAL':
            con.execute("INSERT OR REPLACE INTO history_season_notes(level,season_label,note_sr,note_en,source_url) VALUES(1,?,?,?,?,?)".replace('VALUES(1,?,?,?,?,?)','VALUES(1,?,?,?,?)'),(season,latin(r.get('Napomena') or 'Takmičenje nije održano'),'Competition not held',r.get('Izvor URL')))
            counts['notes']+=1; continue
        for field,pos,champ,run in [('Prvak',1,1,0),('Drugoplasirani',2,0,1)]:
            if r.get(field):
                st,cid=_upsert_excel(con,pre_comp,season,r[field],pos,champ,run,r.get('Izvor URL'),r.get('ID') or '')
                counts[st]+=1; counts['linked' if cid else 'unlinked']+=1
        counts['pre']+=1

    # Post-war/top flight: COMPLETE FINAL TABLES are the authority.
    # PRVENSTVO_1945+ remains a human-readable summary/QA sheet, but club-history
    # rows come from PRVENSTVO_KOMPLETNE_TABELE so 4th, 10th, 15th etc. are kept.
    l1_comps={}
    for r in rows('PRVENSTVO_KOMPLETNE_TABELE'):
        if str(r.get('Status') or '').strip().upper() not in ('','FINAL','FINAL_GROUP'):
            continue
        spec=_l1_spec(r.get('Takmičenje'))
        if not spec: continue
        key=spec[0]
        if key not in l1_comps:
            l1_comps[key]=_ensure_comp(con,key,spec[1],spec[2],'league',1,super_lid,2,r.get('Sezona'),None,r.get('Izvor URL'))
            comp_ids.append(l1_comps[key])
        comp_id=l1_comps[key]
        season=str(r.get('Sezona') or '').strip()
        try: pos=int(r.get('Pozicija')) if r.get('Pozicija') not in (None,'') else None
        except Exception: pos=None
        note_text=latin(r.get('Napomena') or '')
        low_note=note_text.lower()
        stage=latin(r.get('Grupa / faza') or '').strip()
        if _norm(stage) in ('','jedinstvena'): stage=None
        championship_group = not (stage and _norm(stage).startswith('1 b'))
        st,cid=_upsert_excel(
            con,comp_id,season,r.get('Klub'),pos,(pos==1 and championship_group),(pos==2 and championship_group),r.get('Izvor URL'),r.get('ID') or '',
            wiki_title=r.get('Klub iz izvora'),stage_sr=stage,
            played=r.get('Odigrano'),wins=r.get('Pobede'),draws=r.get('Nerešeno'),losses=r.get('Porazi'),
            gf=r.get('Golovi za'),ga=r.get('Golovi protiv'),points=r.get('Bodovi'),
            promoted=('promot' in low_note),relegated=('relegat' in low_note))
        counts[st]+=1; counts['linked' if cid else 'unlinked']+=1
        counts['l1_rows']+=1
    counts['l1']=len(set(str(r.get('Sezona') or '').strip() for r in rows('PRVENSTVO_KOMPLETNE_TABELE') if str(r.get('Status') or '').strip().upper() in ('','FINAL','FINAL_GROUP')))

    # Cups.
    cup_comps={}
    for r in rows('KUP'):
        spec=_cup_spec(r.get('Takmičenje'))
        if not spec: continue
        if spec[0] not in cup_comps:
            cup_comps[spec[0]]=_ensure_comp(con,spec[0],spec[1],spec[2],'cup',None,None,40,None,None,r.get('Izvor URL'))
            comp_ids.append(cup_comps[spec[0]])
        comp_id=cup_comps[spec[0]]; season=_canonical_cup_season(r.get('Sezona'))
        if str(r.get('Status') or '')!='FINAL': continue
        if r.get('Osvajač'):
            st,cid=_upsert_excel(con,comp_id,season,r['Osvajač'],None,1,0,r.get('Izvor URL'),r.get('ID') or '',stage_sr='Osvajač')
            counts[st]+=1; counts['linked' if cid else 'unlinked']+=1
        if r.get('Finalista'):
            st,cid=_upsert_excel(con,comp_id,season,r['Finalista'],None,0,1,r.get('Izvor URL'),r.get('ID') or '',stage_sr='Finalista')
            counts[st]+=1; counts['linked' if cid else 'unlinked']+=1
        counts['cup']+=1

    # Second tier: one competition identity per historical group/zone when needed.
    l2_comps={}
    for r in rows('DRUGI_RANG_1946+'):
        base=_l2_base(r.get('Takmičenje'))
        if not base: continue
        group=latin(r.get('Grupa / zona') or '').strip()
        unified=_norm(group) in ('','jedinstvena') or group=='—'
        key=base[0] if unified else f"{base[0]}-{_slug(group)}"
        name_sr=base[1] if unified else f'{base[1]} – {group}'
        name_en=base[2] if unified else f'{base[2]} – {group}'
        if key not in l2_comps:
            current=prva_lid if base[0]=='serbia-second-tier' else None
            l2_comps[key]=_ensure_comp(con,key,name_sr,name_en,'league',2,current,20,None,None,r.get('Izvor URL'))
            comp_ids.append(l2_comps[key])
        comp_id=l2_comps[key]; season=str(r.get('Sezona') or '').strip()
        for field,pos,champ,run in [('Prvak',1,1,0),('Drugoplasirani',2,0,1),('Treće mesto',3,0,0)]:
            if r.get(field):
                st,cid=_upsert_excel(con,comp_id,season,r[field],pos,champ,run,r.get('Izvor URL'),r.get('ID') or '')
                counts[st]+=1; counts['linked' if cid else 'unlinked']+=1
        counts['l2']+=1

    # Serbian League (third tier), keeping region identity exact.
    l3_comps={}
    for r in rows('SRPSKE_LIGE_3_RANG'):
        spec=_l3_spec(r.get('Takmičenje'))
        if not spec: continue
        canonical=spec[1]
        if spec[0] not in l3_comps:
            l3_comps[spec[0]]=_ensure_comp(con,spec[0],spec[1],spec[2],'league',3,l3_lids.get(canonical),30,None,None,r.get('Izvor URL'))
            comp_ids.append(l3_comps[spec[0]])
        comp_id=l3_comps[spec[0]]; season=str(r.get('Sezona') or '').strip()
        entries=[
            ('Prvak','Prvak - mesto','Prvak - Wiki naslov',1,1,0),
            ('Viceprvak','Viceprvak - mesto','Viceprvak - Wiki naslov',2,0,1),
        ]
        for field,place_field,wiki_field,pos,champ,run in entries:
            if r.get(field):
                st,cid=_upsert_excel(con,comp_id,season,r[field],pos,champ,run,r.get('Izvor URL'),r.get('ID') or '',place=r.get(place_field),wiki_title=r.get(wiki_field))
                counts[st]+=1; counts['linked' if cid else 'unlinked']+=1
        counts['l3']+=1

    # Excel is authoritative for champion/runner-up flags in all imported lineages.
    # Reset first, then reapply from the workbook in a second lightweight pass.
    # We do this at the end because some rows may have been created above.
    # Only flags are reset; detailed standings remain untouched.
    imported_ids=sorted(set(comp_ids))
    _reset_authority(con,imported_ids)

    # Reapply authoritative flags from rows just written (identified by source name).
    con.execute("""UPDATE history_club_seasons SET is_champion=CASE WHEN position=1 THEN 1 ELSE 0 END,
        is_runner_up=CASE WHEN position=2 THEN 1 ELSE 0 END
        WHERE source_name=? AND competition_id IN (%s)""" % ','.join('?' for _ in imported_ids), [SOURCE_NAME,*imported_ids])
    # Cup positions are intentionally NULL, so use verified stage labels.
    con.execute("""UPDATE history_club_seasons SET is_champion=1,is_runner_up=0 WHERE source_name=? AND stage_sr='Osvajač'""",(SOURCE_NAME,))
    con.execute("""UPDATE history_club_seasons SET is_champion=0,is_runner_up=1 WHERE source_name=? AND stage_sr='Finalista'""",(SOURCE_NAME,))

    # Final guard: one club cannot be in two domestic leagues in one season.
    _assert_single_excel_league_per_club_season(con)

    # Extend men's L1 lineage to the Kingdom-era competition.
    con.execute("INSERT OR REPLACE INTO history_meta(key,value,updated_at) VALUES('excel_master_hash',?,CURRENT_TIMESTAMP)",(digest,))
    con.execute("INSERT OR REPLACE INTO history_meta(key,value,updated_at) VALUES('excel_master_version',?,CURRENT_TIMESTAMP)",(MASTER_VERSION,))
    con.execute("INSERT OR REPLACE INTO history_meta(key,value,updated_at) VALUES('excel_master_path',?,CURRENT_TIMESTAMP)",(str(path),))
    con.commit(); wb.close(); con.close()
    counts['hash']=digest; counts['path']=str(path); counts['competitions']=len(imported_ids)
    return counts


def validate_master(path:Path|str=MASTER_PATH):
    path=Path(path)
    if not path.exists(): raise RuntimeError(f'History Excel nije pronađen: {path}')
    wb=load_workbook(path,read_only=True,data_only=True)
    expected={'PRVENSTVO_1945+':81,'PRVENSTVO_KOMPLETNE_TABELE':1264,'PRVENSTVO_PRE_1945':18,'KUP':80,'DRUGI_RANG_1946+':150,'SRPSKE_LIGE_3_RANG':91}
    out={}
    for sheet,want in expected.items():
        if sheet not in wb.sheetnames: out[sheet]={'ok':False,'rows':0,'expected':want}; continue
        rows=sum(1 for r in wb[sheet].iter_rows(min_row=2,values_only=True) if any(v is not None and str(v).strip() for v in r))
        out[sheet]={'ok':rows==want,'rows':rows,'expected':want}
    wb.close()
    return {'ok':all(x['ok'] for x in out.values()),'sheets':out,'hash':_sha256(path),'path':str(path)}


def status():
    con=connect()
    try:
        meta={r['key']:r['value'] for r in con.execute("SELECT key,value FROM history_meta WHERE key LIKE 'excel_master_%'").fetchall()}
        return meta
    finally:
        con.close()
