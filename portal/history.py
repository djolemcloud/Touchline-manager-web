from __future__ import annotations
from .data import connect
from . import sync_engine as engine




FIRST_TIER_LINEAGE_KEYS = (
    'kingdom-yugoslavia-first-1923-1940',
    'yugoslavia-first-1946-1992',
    'fr-yugoslavia-first-1992-2003',
    'serbia-montenegro-first-2003-2006',
    'serbia-top-flight',
)
SECOND_TIER_LINEAGE_PREFIXES = (
    'yugoslavia-second-league',
    'fr-yugoslavia-second-1992-2003',
    'serbia-montenegro-second-2003-2006',
    'serbia-second-tier',
)

def _league_lineage_ids(con, league_id:int):
    """Resolve a live league to the correct historical competition lineage.

    Only the two *men's senior* national divisions are allowed to expand into
    the 1945/46->today L1/L2 lineage. Women's/youth/regional competitions may
    share the same numeric level, but must never inherit men's history merely
    because level==1 or level==2.
    """
    lr=con.execute("SELECT level,COALESCE(section,'senior') section,name FROM leagues WHERE id=?",(int(league_id),)).fetchone()
    if not lr:
        return []
    level=int(lr['level']) if lr['level'] is not None else None
    section=(lr['section'] or 'senior').lower()
    name=(lr['name'] or '').strip()
    is_mens_top = section=='senior' and name=='SuperLiga Srbije' and level==1
    is_mens_second = section=='senior' and name=='Prva liga Srbije' and level==2
    if is_mens_top:
        qs=','.join('?' for _ in FIRST_TIER_LINEAGE_KEYS)
        rows=con.execute(f"SELECT id FROM history_competitions WHERE comp_type='league' AND history_key IN ({qs})",FIRST_TIER_LINEAGE_KEYS).fetchall()
        return [int(r['id']) for r in rows]
    if is_mens_second:
        cond=' OR '.join('history_key=? OR history_key LIKE ?' for _ in SECOND_TIER_LINEAGE_PREFIXES)
        params=[]
        for pref in SECOND_TIER_LINEAGE_PREFIXES:
            params.extend([pref,pref+'-%'])
        rows=con.execute(f"SELECT id FROM history_competitions WHERE comp_type='league' AND level=2 AND ({cond})",params).fetchall()
        return [int(r['id']) for r in rows]
    # Every other competition is isolated to its own explicit historical master.
    rows=con.execute("SELECT id FROM history_competitions WHERE current_league_id=? AND comp_type='league'",(int(league_id),)).fetchall()
    return [int(r['id']) for r in rows]

def _lineage_clause(con, league_id:int, column='h.competition_id'):
    ids=_league_lineage_ids(con,league_id)
    if not ids:
        return '1=0',[]
    return f"{column} IN ({','.join('?' for _ in ids)})",ids

def comp_name(row):
    if engine.ui_lang() == 'en':
        return row['name_en'] or row['name_sr']
    return row['name_sr']




# Historical Yugoslav Cup editions were sometimes named only by the calendar
# year even when the surrounding league history is shown as a football season.
# Keep the historically calendar-year league era (1950-1952) intact, but fold
# the other cup-only calendar labels into the matching football-season row so
# club history can never show two visual rows for the same sporting season.
LEGACY_CUP_PRESENTATION_SEASONS = {
    '1947':'1946/47',
    '1948':'1947/48',
    '1949':'1948/49',
    '1953':'1952/53',
    '1954':'1953/54',
    '1955':'1954/55',
    '1973':'1972/73',
}


def _presentation_season(row):
    label=str(row['season_label'] or '').strip()
    if (row['comp_type'] or '').lower()=='cup':
        return LEGACY_CUP_PRESENTATION_SEASONS.get(label,label)
    return label


def club_seasons(club_id: int):
    con=connect()
    rows=con.execute('''SELECT h.*,hc.name_sr,hc.name_en,hc.comp_type,hc.level,hc.current_league_id,hc.sort_order
        FROM history_club_seasons h
        JOIN history_competitions hc ON hc.id=h.competition_id
        WHERE h.club_id=? AND (hc.comp_type<>'league' OR h.source_name='Touchline History Master Excel')
        ORDER BY CAST(substr(h.season_label,1,4) AS INTEGER) DESC,h.season_label DESC,hc.sort_order,hc.name_sr''',(int(club_id),)).fetchall()
    con.close(); return rows


def club_summary(club_id: int):
    con=connect()
    row=con.execute('''SELECT
        COUNT(DISTINCT h.season_label) seasons,
        SUM(CASE WHEN h.is_champion=1 THEN 1 ELSE 0 END) titles,
        SUM(CASE WHEN h.is_champion=1 AND hc.comp_type='league' AND hc.level=1 THEN 1 ELSE 0 END) league_titles,
        SUM(CASE WHEN h.is_runner_up=1 AND hc.comp_type='league' AND hc.level=1 THEN 1 ELSE 0 END) league_runner_up,
        SUM(CASE WHEN h.is_runner_up=1 THEN 1 ELSE 0 END) runner_up,
        SUM(CASE WHEN h.is_champion=1 AND hc.comp_type='cup' THEN 1 ELSE 0 END) cups,
        SUM(CASE WHEN h.is_champion=1 AND hc.comp_type IN ('europe','world') THEN 1 ELSE 0 END) international_titles,
        SUM(CASE WHEN (h.is_champion=1 OR h.is_runner_up=1 OR COALESCE(h.highlight,0)=1) AND hc.comp_type IN ('europe','world') THEN 1 ELSE 0 END) international
        FROM history_club_seasons h JOIN history_competitions hc ON hc.id=h.competition_id
        WHERE h.club_id=?''',(int(club_id),)).fetchone()
    con.close()
    return dict(row) if row else {'seasons':0,'titles':0,'league_titles':0,'league_runner_up':0,'runner_up':0,'cups':0,'international_titles':0,'international':0}


def club_season_timeline(club_id: int):
    """Return one presentation row per football season.

    League, cup and international records remain normalized in the database,
    while this function groups them into the one-row-per-season view used by the
    public club history page.
    """
    rows=club_seasons(club_id)
    seasons={}
    for r in rows:
        season=_presentation_season(r)
        item=seasons.setdefault(season,{'season_label':season,'league':[],'cup':[],'europe':[],'other':[]})
        kind=(r['comp_type'] or 'other').lower(); rec=dict(r)
        if kind=='league': item['league'].append(rec)
        elif kind=='cup': item['cup'].append(rec)
        elif kind in ('europe','world'): item['europe'].append(rec)
        else: item['other'].append(rec)
    def year_key(label):
        try: return int(str(label)[:4])
        except Exception: return -1
    return sorted(seasons.values(),key=lambda x:(year_key(x['season_label']),x['season_label']),reverse=True)


def result_label(row):
    if not row: return '—'
    lang=engine.ui_lang(); kind=(row.get('comp_type') or '').lower()
    stage=(row.get('stage_en') if lang=='en' else row.get('stage_sr'))
    # International milestones often have a more meaningful verified label
    # than the generic Champion/Runner-up (e.g. "Šampion Evrope").
    if kind in ('europe','world') and stage:
        return stage
    if row.get('is_champion'):
        if kind=='cup': return 'Winner' if lang=='en' else 'Osvajač'
        if kind=='league': return 'Champion' if lang=='en' else 'Prvak'
        return 'Champion' if lang=='en' else 'Šampion'
    if row.get('is_runner_up'):
        if kind=='cup': return 'Runner-up' if lang=='en' else 'Finalista'
        if kind=='league': return 'Runner-up' if lang=='en' else 'Vicešampion'
        return 'Runner-up' if lang=='en' else 'Finalista'
    return stage or '—'


def club_international_highlights(club_id:int):
    """Individual international achievements, including stage-only highlights."""
    con=connect()
    rows=con.execute('''SELECT h.*,hc.history_key,hc.name_sr,hc.name_en,hc.comp_type,hc.level,hc.sort_order
        FROM history_club_seasons h JOIN history_competitions hc ON hc.id=h.competition_id
        WHERE h.club_id=? AND hc.comp_type IN ('europe','world') AND COALESCE(h.highlight,0)=1
        ORDER BY h.is_champion DESC,h.is_runner_up DESC,
          CASE hc.comp_type WHEN 'world' THEN 1 WHEN 'europe' THEN 2 ELSE 9 END,
          hc.sort_order,CAST(substr(h.season_label,1,4) AS INTEGER) DESC''',(int(club_id),)).fetchall()
    con.close(); return rows


def league_champions(league_id: int):
    con=connect(); clause,params=_lineage_clause(con,league_id)
    rows=con.execute(f'''SELECT h.*,hc.name_sr,hc.name_en,hc.comp_type,
        COALESCE(c.cm_name,c.name) club_name_en,c.name club_name_sr
        FROM history_club_seasons h
        JOIN history_competitions hc ON hc.id=h.competition_id
        LEFT JOIN clubs c ON c.id=h.club_id
        WHERE {clause} AND h.is_champion=1 AND h.source_name='Touchline History Master Excel'
        ORDER BY CAST(substr(h.season_label,1,4) AS INTEGER) DESC,h.season_label DESC,hc.sort_order''',params).fetchall()
    con.close(); return rows


def league_season_notes(league_id: int):
    con=connect()
    try:
        row=con.execute("SELECT level,COALESCE(section,'senior') section,name FROM leagues WHERE id=?",(int(league_id),)).fetchone()
        if (not row or int(row['level'] or 0) != 2 or (row['section'] or 'senior').lower()!='senior'
                or (row['name'] or '')!='Prva liga Srbije'):
            return []
        present={r[0] for r in con.execute("SELECT name FROM sqlite_master WHERE type='table'")}
        if 'history_season_notes' not in present:
            return []
        return con.execute("""SELECT season_label,note_sr,note_en,source_url
            FROM history_season_notes WHERE level=2
            ORDER BY CAST(substr(season_label,1,4) AS INTEGER) DESC,season_label DESC""").fetchall()
    finally:
        con.close()


def league_title_ranking(league_id: int):
    con=connect(); clause,params=_lineage_clause(con,league_id)
    rows=con.execute(f'''SELECT h.club_id,
        COALESCE(MAX(CASE WHEN ?='en' THEN COALESCE(c.cm_name,c.name) ELSE c.name END),MAX(c.name),MAX(h.club_name_override)) club_name,
        COUNT(*) titles,MIN(h.season_label) first_title,MAX(h.season_label) last_title
        FROM history_club_seasons h
        JOIN history_competitions hc ON hc.id=h.competition_id
        LEFT JOIN clubs c ON c.id=h.club_id
        WHERE {clause} AND h.is_champion=1 AND h.source_name='Touchline History Master Excel'
        GROUP BY h.club_id,COALESCE(h.club_name_override,'')
        ORDER BY titles DESC,last_title DESC,club_name''',[engine.ui_lang(),*params]).fetchall()
    con.close(); return rows


def league_all_time(league_id: int):
    con=connect(); clause,params=_lineage_clause(con,league_id)
    rows=con.execute(f'''SELECT h.club_id,
        COALESCE(MAX(CASE WHEN ?='en' THEN COALESCE(c.cm_name,c.name) ELSE c.name END),MAX(c.name),MAX(h.club_name_override)) club_name,
        COUNT(DISTINCT h.season_label) seasons,SUM(COALESCE(h.played,0)) played,SUM(COALESCE(h.wins,0)) wins,SUM(COALESCE(h.draws,0)) draws,
        SUM(COALESCE(h.losses,0)) losses,SUM(COALESCE(h.goals_for,0)) goals_for,SUM(COALESCE(h.goals_against,0)) goals_against,
        SUM(COALESCE(h.points,0)) points,SUM(CASE WHEN h.is_champion=1 THEN 1 ELSE 0 END) titles
        FROM history_club_seasons h JOIN history_competitions hc ON hc.id=h.competition_id
        LEFT JOIN clubs c ON c.id=h.club_id
        WHERE {clause} AND hc.comp_type='league' AND h.source_name='Touchline History Master Excel' AND h.played IS NOT NULL
        GROUP BY h.club_id,COALESCE(h.club_name_override,'')
        ORDER BY points DESC,wins DESC,goals_for DESC''',[engine.ui_lang(),*params]).fetchall()
    con.close(); return rows


def club_year_span(club_id:int):
    con=connect(); row=con.execute('''SELECT MIN(CAST(substr(season_label,1,4) AS INTEGER)) first_year,
                              MAX(CAST(substr(season_label,1,4) AS INTEGER)) last_year
                       FROM history_club_seasons WHERE club_id=?''',(int(club_id),)).fetchone(); con.close()
    return ((row['first_year'] if row else None),(row['last_year'] if row else None))


def league_history_meta(league_id:int):
    con=connect(); ids=_league_lineage_ids(con,league_id)
    if not ids:
        con.close(); return None
    q=','.join('?' for _ in ids)
    row=con.execute(f'''SELECT MIN(active_from) active_from,MAX(COALESCE(active_to,'9999/99')) active_to
        FROM history_competitions WHERE id IN ({q})''',ids).fetchone()
    con.close(); return row


def club_honour_groups(club_id:int):
    """Aggregate domestic league/cup honours by historical competition era."""
    con=connect()
    rows=con.execute('''SELECT hc.id,hc.history_key,hc.name_sr,hc.name_en,hc.comp_type,hc.level,hc.sort_order,
        SUM(CASE WHEN h.is_champion=1 THEN 1 ELSE 0 END) title_count,
        SUM(CASE WHEN h.is_runner_up=1 THEN 1 ELSE 0 END) runner_up_count,
        MIN(CASE WHEN h.is_champion=1 OR h.is_runner_up=1 THEN h.season_label END) first_season,
        MAX(CASE WHEN h.is_champion=1 OR h.is_runner_up=1 THEN h.season_label END) last_season
        FROM history_club_seasons h JOIN history_competitions hc ON hc.id=h.competition_id
        WHERE h.club_id=? AND hc.comp_type IN ('league','cup') AND (h.is_champion=1 OR h.is_runner_up=1)
          AND (hc.comp_type<>'league' OR h.source_name='Touchline History Master Excel')
        GROUP BY hc.id
        HAVING SUM(CASE WHEN h.is_champion=1 THEN 1 ELSE 0 END)>0 OR SUM(CASE WHEN h.is_runner_up=1 THEN 1 ELSE 0 END)>0
        ORDER BY CASE hc.comp_type WHEN 'league' THEN 1 WHEN 'cup' THEN 2 ELSE 9 END,hc.sort_order,hc.name_sr''',(int(club_id),)).fetchall()
    con.close(); return rows


def club_major_totals(club_id:int):
    con=connect(); row=con.execute('''SELECT
        SUM(CASE WHEN h.is_champion=1 AND hc.comp_type='league' AND hc.level=1 THEN 1 ELSE 0 END) national_titles,
        SUM(CASE WHEN h.is_champion=1 AND hc.comp_type='cup' THEN 1 ELSE 0 END) national_cups
        FROM history_club_seasons h JOIN history_competitions hc ON hc.id=h.competition_id
        WHERE h.club_id=? AND (hc.comp_type<>'league' OR h.source_name='Touchline History Master Excel')''',(int(club_id),)).fetchone(); con.close()
    return {'national_titles':int((row['national_titles'] if row else 0) or 0),'national_cups':int((row['national_cups'] if row else 0) or 0)}
