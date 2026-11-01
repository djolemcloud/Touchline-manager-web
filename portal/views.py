from __future__ import annotations
from .data import connect
from .stats import league_stats, aggregate
from .ui import layout, notice, crumb, visual_card, rich_card, group_card, metric_strip, esc
from . import sync_engine as engine
from . import history as history_data
from . import trophies as trophy_data
from .catalog import HOME_LEVELS, FSS_HUB, DEVELOPMENT_HUB, REGIONS, YOUTH_IMAGES, WOMEN_SECOND_ORDER, LEVEL6_REGION_ORDER


def _league_rows(cid=1, section='senior', level=None):
    con=connect(); sql="SELECT * FROM leagues WHERE country_id=? AND COALESCE(section,'senior')=? AND COALESCE(source_active,1)=1"; p=[int(cid),section]
    if level is not None:
        sql+=' AND level=?'; p.append(int(level))
    if section=='senior' and int(level or 0)==3:
        # Public Srpske lige page is intentionally fixed to the four canonical
        # third-tier Serbian leagues. Old discovery builds could leave leaked
        # Super/Prva rows in the database; never render them here.
        names=('Srpska liga Beograd','Srpska liga Vojvodina','Srpska liga Istok','Srpska liga Zapad')
        sql+=' AND name IN (?,?,?,?)'; p.extend(names)
    sql+=' ORDER BY level,name'; rows=con.execute(sql,p).fetchall(); con.close(); return rows


def _country(cid=1):
    con=connect(); row=con.execute('SELECT * FROM countries WHERE id=?',(int(cid),)).fetchone(); con.close(); return row




def _display_league_name(name):
    return engine.display_league_name(name)


def _trophy_media(history_key, comp_type, alt):
    rec=trophy_data.trophy_for(history_key, comp_type)
    src=esc(rec.get('image') or '/static/history/trophies/fallback.svg')
    alt=esc(alt or 'Trofej')
    fallback='/static/history/trophies/fallback.svg'
    return f'<div class="honour-trophy"><img src="{src}" alt="{alt}" loading="lazy" referrerpolicy="no-referrer" onerror="this.onerror=null;this.src=\'{fallback}\';"></div>'


def _standings_block(owner_type,owner_id):
    comp,rows=engine.standings_for_owner(owner_type,int(owner_id))
    if not comp:
        return '<section class="standings"><h2>Tabela</h2><div class="panel"><p>Tabela još nije povučena. Osvežava se iz Superadmin centra za sinhronizaciju.</p></div></section>'
    has_goals=any(r['goals_for'] is not None or r['goals_against'] is not None for r in rows)
    head='<tr><th>#</th><th>Tim</th><th>OU</th><th>P</th><th>N</th><th>I</th>'
    if has_goals: head+='<th>DG</th><th>PG</th><th>GR</th>'
    head+='<th>B</th></tr>'
    body=[]
    link_con=connect() if owner_type=='league' else None
    for r in rows:
        club_id=r['club_id']
        if not club_id and link_con is not None:
            club_id=engine._standing_match_club(link_con,int(owner_id),r['team_name'])
        team=f'<a href="/club?id={club_id}">{esc(r["team_name"])}</a>' if club_id else esc(r['team_name'])
        pos=str(r['position'] or '')+(f' <small>{esc(r["position_note"])}</small>' if r['position_note'] else '')
        line=f'<tr><td>{pos}</td><td>{team}</td><td>{r["played"] if r["played"] is not None else ""}</td><td>{r["wins"] if r["wins"] is not None else ""}</td><td>{r["draws"] if r["draws"] is not None else ""}</td><td>{r["losses"] if r["losses"] is not None else ""}</td>'
        if has_goals:
            line+=f'<td>{r["goals_for"] if r["goals_for"] is not None else ""}</td><td>{r["goals_against"] if r["goals_against"] is not None else ""}</td><td>{r["goal_diff"] if r["goal_diff"] is not None else ""}</td>'
        line+=f'<td><b>{r["points"] if r["points"] is not None else ""}</b></td></tr>'; body.append(line)
    if link_con is not None: link_con.close()
    meta=f'{esc(comp["season"] or "")} · {esc((comp["source_type"] or "").upper())} · osveženo {esc(comp["updated_at"] or "")}'
    return f'<section class="standings"><h2>Tabela</h2><p class="standingsmeta">{meta}</p><div class="tablewrap"><table><thead>{head}</thead><tbody>{"".join(body)}</tbody></table></div></section>'


def home(msg=''):
    cid=1
    levels={lv:_league_rows(cid,'senior',lv) for lv in range(1,6)}
    st=league_stats(country_id=cid,section='senior')
    agg={lv:aggregate([st.get(r['id']) for r in levels[lv]]) for lv in levels}
    cards=[]
    for lv in range(1,6):
        cfg=HOME_LEVELS[lv]; rows=levels[lv]
        href=f'/league?id={rows[0]["id"]}' if lv in (1,2) and rows else f'/level?country_id={cid}&level={lv}'
        extra=(f'{agg[lv]["clubs"]} klubova · {agg[lv]["mapped"]}/{agg[lv]["clubs"]} TM mapirano' if lv in (1,2)
               else f'{len(rows)} liga · {agg[lv]["clubs"]} klubova')
        cards.append(visual_card(cfg['image'],cfg['title'],cfg['label'],href,extra))
    cfg=HOME_LEVELS[6]
    cards.append(visual_card(cfg['image'],cfg['title'],cfg['label'],f'/level6?country_id={cid}',''))
    cards.append(visual_card('/static/branding/nacionalna_takmicenja.png','Nacionalna takmičenja','FSS',f'/fss?country_id={cid}','Reprezentacije · omladinska · ženska · regionalni savezi'))
    sponsors='''<section class="sponsors"><div class="sponsorheading"><small>GENERALNI SPONZORI SAJTA</small></div><div class="sponsorlogos">
    <div class="sponsorlogo"><img src="/static/sponsors/msmobile_logo.png" alt="MS Mobile"></div>
    <div class="sponsorlogo"><img src="/static/sponsors/superior_logo.png" alt="Superior"></div>
    <div class="sponsorlogo"><img src="/static/sponsors/moxom_logo_v2.png" alt="Moxom"></div></div></section>'''
    return layout('Touchline Web Portal',notice(msg)+'<h1>Srbija</h1><div id="competitions" class="visualgrid homegrid">'+''.join(cards)+'</div>'+sponsors)


def fss_hub(cid=1,msg=''):
    cards=''.join(visual_card(x['image'],x['title'],x['label'],f"{x['route']}?country_id={cid}",'') for x in FSS_HUB)
    return layout('Nacionalna takmičenja',notice(msg)+crumb([('Srbija','/'),('Nacionalna takmičenja',None)])+'<h1>Nacionalna takmičenja</h1><div class="visualstack">'+cards+'</div>')


def national_hub(cid=1,msg=''):
    con=connect(); teams=con.execute("SELECT * FROM national_teams WHERE country_id=? AND team_type IN ('senior','u21') ORDER BY CASE team_type WHEN 'senior' THEN 1 ELSE 2 END",(int(cid),)).fetchall(); con.close()
    cards=''.join(visual_card('/'+(t['image_path'] or 'static/branding/fss.jpg'),t['short_name'] or t['name'],'NACIONALNA SELEKCIJA',f'/national_team?id={t["id"]}','') for t in teams)
    return layout('Seniorski tim i Srbija U21',notice(msg)+crumb([('Nacionalna takmičenja',f'/fss?country_id={cid}'),('Seniorski tim i Srbija U21',None)])+'<h1>Seniorski tim i Srbija U21</h1><div class="visualgrid">'+cards+'</div>')


def development_hub(cid=1,msg=''):
    cards=''.join(visual_card(x['image'],x['title'],x['label'],f"{x['route']}?country_id={cid}",'') for x in DEVELOPMENT_HUB)
    return layout('Omladinska i ženska takmičenja',notice(msg)+crumb([('Nacionalna takmičenja',f'/fss?country_id={cid}'),('Omladinska i ženska takmičenja',None)])+'<h1>Omladinska i ženska takmičenja</h1><div class="visualgrid developmentgrid">'+cards+'</div>')


def youth(cid=1,msg=''):
    rows=_league_rows(cid,'youth')
    # Canonical display order: U19 (Omladinska) first, then U17 (Kadetska).
    rows=sorted(rows,key=lambda r: ({'U19':1,'U17':2}.get((r['age_group'] or '').upper(),9), r['name']))
    st=league_stats(league_ids=[r['id'] for r in rows]); total=aggregate([st.get(r['id']) for r in rows])
    cards=''.join(rich_card(title=_display_league_name(r['name']),subtitle='FSS · OMLADINSKI FUDBAL',href=f'/league?id={r["id"]}',stats=st.get(r['id'],{}),image=YOUTH_IMAGES.get((r['age_group'] or '').upper())) for r in rows)
    body=notice(msg)+crumb([('Nacionalna takmičenja',f'/fss?country_id={cid}'),('Omladinska i ženska takmičenja',f'/fss/development?country_id={cid}'),('Omladinska takmičenja',None)])+'<h1>Omladinska takmičenja</h1>'+metric_strip([('LIGE',len(rows)),('KLUBOVI',total['clubs']),('TM MAPIRANO',f'{total["mapped"]}/{total["clubs"]}'),('IGRAČI',total['players']),('STRUČNI ŠTAB',total['staff'])])+'<div class="richgrid">'+cards+'</div>'
    return layout('Omladinska takmičenja',body)


def women(cid=1,msg=''):
    con=connect(); teams=con.execute("SELECT * FROM national_teams WHERE country_id=? AND team_type IN ('women_senior','women_u19') ORDER BY CASE team_type WHEN 'women_senior' THEN 1 ELSE 2 END",(int(cid),)).fetchall(); leagues=con.execute("SELECT * FROM leagues WHERE country_id=? AND section='women' ORDER BY level,name",(int(cid),)).fetchall(); con.close()
    team_cards=''.join(visual_card('/'+(t['image_path'] or 'static/branding/fss.jpg'),t['name'],'NACIONALNE SELEKCIJE ŽENA',f'/national_team?id={t["id"]}','') for t in teams)
    st=league_stats(league_ids=[r['id'] for r in leagues]); club=[]
    for r in leagues:
        if r['level'] in (1,2): club.append(rich_card(title=_display_league_name(r['name']),subtitle='FSS · ŽENSKI FUDBAL',href=f'/league?id={r["id"]}',stats=st.get(r['id'],{}),image='/'+r['image_path'] if r['image_path'] else None))
    second=[r for r in leagues if r['level']==3]; second_stats=aggregate([st.get(r['id']) for r in second])
    club.append(f'<a class="womenhubvisualcard" href="/women/second?country_id={cid}"><img src="/static/branding/druga_zenska_liga_hub.png" alt="Druga ženska liga"><div class="hub-overlay-stats"><strong>{second_stats["clubs"]} klubova</strong><span>3 regionalne lige</span></div></a>')
    body=notice(msg)+crumb([('Nacionalna takmičenja',f'/fss?country_id={cid}'),('Omladinska i ženska takmičenja',f'/fss/development?country_id={cid}'),('Ženska takmičenja',None)])+'<h1>Ženska takmičenja</h1><h2>Nacionalne selekcije žena</h2><div class="visualgrid">'+team_cards+'</div><h2>Klupska takmičenja žena</h2><div class="richgrid">'+''.join(club)+'</div>'
    return layout('Ženska takmičenja',body)


def women_second(cid=1,msg=''):
    rows=sorted(_league_rows(cid,'women',3),key=lambda r:WOMEN_SECOND_ORDER.get(r['name'],9)); st=league_stats(league_ids=[r['id'] for r in rows])
    cards=''.join(rich_card(title=_display_league_name(r['name']),subtitle='DRUGA ŽENSKA LIGA',href=f'/league?id={r["id"]}',stats=st.get(r['id'],{}),image='/'+r['image_path'] if r['image_path'] else None) for r in rows)
    return layout('Druga ženska liga',notice(msg)+crumb([('Ženska takmičenja',f'/women?country_id={cid}'),('Druga ženska liga',None)])+'<h1>Druga ženska liga</h1><div class="richgrid">'+cards+'</div>')


def regions(cid=1,msg=''):
    cards=''.join(visual_card(cfg['image'],cfg['title'],'REGIONALNI SAVEZ',f'/fss/region?country_id={cid}&code={code}','') for code,cfg in REGIONS.items())
    return layout('Regionalni savezi',notice(msg)+crumb([('Nacionalna takmičenja',f'/fss?country_id={cid}'),('Regionalni savezi',None)])+'<h1>Regionalni savezi</h1><div class="visualgrid regionsgrid">'+cards+'</div>')


def _region_rows(cid,code):
    rows=_league_rows(cid,'senior')
    code=(code or '').upper()
    return [r for r in rows if int(r['level'] or 0) in (3,4,5,6) and (r['region_code'] or '').upper()==code]


def region(cid=1,code='FSB',msg=''):
    code=(code or 'FSB').upper(); cfg=REGIONS.get(code,{'title':code}); rows=_region_rows(cid,code)
    st=league_stats(league_ids=[r['id'] for r in rows]); sections=[]
    labels={3:'NIVO 3 · SRPSKA LIGA',4:'NIVO 4 · ZONSKE LIGE',5:'NIVO 5 · OKRUŽNE LIGE',6:'NIVO 6 · OPŠTINSKE / MEĐUOPŠTINSKE LIGE'}
    for lv in (3,4,5,6):
        rr=[r for r in rows if int(r['level'] or 0)==lv]
        cards=''.join(rich_card(title=_display_league_name(r['name']),subtitle=f'NIVO {lv} · {str(r["source_type"] or "").upper()}',href=f'/league?id={r["id"]}',stats=st.get(r['id'],{})) for r in rr)
        if not cards:
            cards='<div class="panel"><p>Za ovaj nivo trenutno nema evidentiranih liga u bazi.</p></div>'
        sections.append(f'<section class="regionlevel"><h2>{esc(labels[lv])}</h2><div class="richgrid">{cards}</div></section>')
    content=''.join(sections) if sections else '<div class="panel"><p>Za ovaj regionalni savez trenutno nema evidentiranih liga.</p></div>'
    title=cfg['title']
    return layout(title,notice(msg)+crumb([('Regionalni savezi',f'/fss/regions?country_id={cid}'),(title,None)])+f'<h1>{esc(title)}</h1>'+content)

def fsb_first(cid=1,msg=''):
    rows=[r for r in _league_rows(cid,'senior',5) if r['name'].startswith('Prva Beogradska liga')]; st=league_stats(league_ids=[r['id'] for r in rows])
    cards=''.join(rich_card(title=_display_league_name(r['name']),subtitle='PRVA BEOGRADSKA LIGA',href=f'/league?id={r["id"]}',stats=st.get(r['id'],{})) for r in rows)
    return layout('Prva beogradska liga',notice(msg)+crumb([('Fudbalski savez Beograda',f'/fss/region?country_id={cid}&code=FSB'),('Prva beogradska liga',None)])+'<h1>Prva beogradska liga</h1><div class="richgrid">'+cards+'</div>')


def level(cid=1,lv=3,msg=''):
    rows=_league_rows(cid,'senior',lv); st=league_stats(league_ids=[r['id'] for r in rows]); title={3:'Srpske lige',4:'Zonske lige',5:'Okružne lige'}.get(int(lv),f'Nivo {lv}')
    cards=''.join(rich_card(title=_display_league_name(r['name']),subtitle=f'NIVO {lv} · {str(r["source_type"] or "").upper()}',href=f'/league?id={r["id"]}',stats=st.get(r['id'],{})) for r in rows)
    return layout(title,notice(msg)+crumb([('Srbija','/'),(title,None)])+f'<h1>{esc(title)}</h1><div class="richgrid">{cards}</div>')


def level6(cid=1,msg=''):
    rows=_league_rows(cid,'senior',6); st=league_stats(league_ids=[r['id'] for r in rows])
    cards=''.join(rich_card(title=_display_league_name(r['name']),subtitle='NIVO 6 · SRBIJASPORT',href=f'/league?id={r["id"]}',stats=st.get(r['id'],{})) for r in rows)
    return layout('Opštinske lige',notice(msg)+crumb([('Srbija','/'),('Opštinske lige',None)])+'<h1>Opštinske lige</h1><p>Sve lige šestog nivoa na jednom mestu — međuopštinske, opštinske i druge opštinske lige.</p><div class="richgrid">'+cards+'</div>')

def level6_region(cid=1,code='FSB',msg=''):
    code=(code or 'FSB').upper(); cfg=REGIONS.get(code,{'title':code}); rows=[r for r in _league_rows(cid,'senior',6) if (r['region_code'] or '').upper()==code]
    st=league_stats(league_ids=[r['id'] for r in rows]); cards=''.join(rich_card(title=_display_league_name(r['name']),subtitle='NIVO 6 · SRBIJASPORT',href=f'/league?id={r["id"]}',stats=st.get(r['id'],{})) for r in rows)
    content=cards or '<div class="panel"><p>Za ovaj region trenutno nema evidentiranih liga šestog nivoa.</p></div>'
    return layout(cfg['title']+' · Nivo 6',notice(msg)+crumb([('Opštinske lige',f'/level6?country_id={cid}'),(cfg['title'],None)])+f'<h1>{esc(cfg["title"])}</h1><div class="richgrid">{content}</div>')


def _crest_url(info):
    if not info or not info['crest_local']: return ''
    p=str(info['crest_local']).replace('\\','/').lstrip('/')
    return '/'+p


CLUB_SR_NAMES={
    'Red Star Belgrade':'Crvena zvezda Beograd','Partizan Belgrade':'Partizan Beograd',
    'Vojvodina Novi Sad':'Vojvodina Novi Sad','Radnicki 1923':'Radnički 1923',
    'IMT Novi Beograd':'IMT Novi Beograd','Radnik Surdulica':'Radnik Surdulica',
    'Mladost Lucani':'Mladost Lučani','Novi Pazar':'Novi Pazar','Zeleznicar Pancevo':'Železničar Pančevo',
    'Cukaricki':'Čukarički','OFK Beograd':'OFK Beograd','Radnicki Nis':'Radnički Niš','Macva Sabac':'Mačva Šabac',
    'Zemun':'Zemun','Borac 1926':'Borac 1926 Čačak','TSC Backa Topola':'TSC Bačka Topola','OFK Vrsac':'OFK Vršac',
    'Javor Matis':'Javor Matis','Spartak Subotica':'Spartak Subotica','Smederevo 1924':'Smederevo 1924',
    'GFK Dubocica 1923':'GFK Dubočica 1923','Napredak Krusevac':'Napredak Kruševac',
    'Metalac Gornji Milanovac':'Metalac Gornji Milanovac','Vozdovac':'Voždovac','Loznica':'Loznica',
    'Bor 1919':'Bor 1919','Proleter 023':'Proleter 023','Jedinstvo Ub':'Jedinstvo Ub','RFK Graficar':'RFK Grafičar',
    'Teleoptik':'Teleoptik'
}


def _display_club_name(row_or_name):
    if isinstance(row_or_name,str):
        raw=row_or_name
    else:
        raw=(row_or_name['cm_name'] or row_or_name['name']) if engine.ui_lang()=='en' else (row_or_name['name'] or row_or_name['cm_name'])
    if engine.ui_lang()=='sr': return CLUB_SR_NAMES.get(raw,raw)
    return raw


def _asset_url(path):
    if not path: return ''
    p=str(path).replace('\\','/').lstrip('/')
    return '/'+p


def _display_stadium_name(name):
    """Normalize common source-language stadium names for the active UI language."""
    raw=str(name or 'Stadion').strip()
    sr={
        'Partizan Stadium':'Stadion Partizana',
        'Stadion Cukaricki':'Stadion Čukarički',
        'Stadion Karadjordje':'Stadion Karađorđe',
        'Stadion Dragan Dzajic':'Stadion Dragan Džajić',
        'Čukarički Stadium':'Stadion Čukarički',
        'Cukaricki Stadium':'Stadion Čukarički',
        'Karađorđe Stadium':'Stadion Karađorđe',
        'Karadjordje Stadium':'Stadion Karađorđe',
        'Dragan Džajić Stadium':'Stadion Dragan Džajić',
        'Dragan Dzajic Stadium':'Stadion Dragan Džajić',
        'Rajko Mitić Stadium':'Stadion Rajko Mitić',
        'Rajko Mitic Stadium':'Stadion Rajko Mitić',
        'Rajko Mitić':'Stadion Rajko Mitić',
        'Rajko Mitic':'Stadion Rajko Mitić',
    }
    en={
        'Stadion Partizana':'Partizan Stadium',
        'Stadion Čukarički':'Čukarički Stadium',
        'Stadion Cukaricki':'Čukarički Stadium',
        'Stadion Karađorđe':'Karađorđe Stadium',
        'Stadion Karadjordje':'Karađorđe Stadium',
        'Stadion Dragan Džajić':'Dragan Džajić Stadium',
        'Stadion Dragan Dzajic':'Dragan Džajić Stadium',
        'Stadion Rajko Mitić':'Rajko Mitić Stadium',
        'Rajko Mitić':'Rajko Mitić Stadium',
        'Rajko Mitic':'Rajko Mitić Stadium',
        'Čika Dača Stadion':'Čika Dača Stadium',
        'Gradski stadion Novi Pazar':'Novi Pazar City Stadium',
        'Gradski stadion Surdulica':'Surdulica City Stadium',
        'Sportski Centar Stara Pazova':'Stara Pazova Sports Centre',
    }
    if engine.ui_lang()=='en':
        if raw in en: return en[raw]
        if raw.startswith('Gradski stadion '): return raw[len('Gradski stadion '):]+' City Stadium'
        if raw.startswith('Stadion '): return raw[len('Stadion '):]+' Stadium'
        if raw.endswith(' Stadion'): return raw[:-len(' Stadion')]+' Stadium'
        return raw
    if raw in sr: return sr[raw]
    if raw.endswith(' Stadium'): return 'Stadion '+raw[:-len(' Stadium')]
    return raw


def _stadium_asset(club_id, image_local=''):
    if image_local:
        return _asset_url(image_local)
    # Simple local override: dropping static/stadiums/<club_id>.(jpg|png|webp)
    # is enough; no DB edit is required.
    from .config import STATIC_DIR
    for ext in ('.jpg','.jpeg','.png','.webp'):
        p=STATIC_DIR/'stadiums'/f'{int(club_id)}{ext}'
        if p.is_file():
            return '/static/stadiums/'+p.name
    return ''


def _stadium_visual(stadium, club_id=None):
    image_local=(stadium['image_local'] if stadium and 'image_local' in stadium.keys() else '')
    src=_stadium_asset(club_id,image_local) if club_id is not None else _asset_url(image_local)
    if src:
        return f'<img class="stadium-photo" src="{esc(src)}" alt="{esc(_display_stadium_name(stadium["name"] or "Stadion"))}">'
    return '<div class="stadium-photo stadium-placeholder"><span class="stadium-icon">⌂</span><small>Fotografija stadiona još nije učitana.</small></div>'




def _jersey_assets(club_id):
    from .config import STATIC_DIR
    out=[]
    for idx in (1,2,3):
        for ext in ('.png','.webp','.jpg','.jpeg'):
            p=STATIC_DIR/'jerseys'/f'{int(club_id)}_{idx}{ext}'
            if p.is_file():
                out.append(f'/static/jerseys/{p.name}')
                break
    return out


def _jersey_card_visual(club_id, club_name='Klub'):
    kits=_jersey_assets(club_id)
    if kits:
        return f'<div class="clubcardjersey"><img class="jersey-card-img" src="{esc(kits[0])}" alt="{esc(club_name)} dres"></div>'
    return ''


def _jersey_panel(club_id, club_name='Klub'):
    kits=_jersey_assets(club_id)
    if not kits:
        return ''
    labels=['1. garnitura','2. garnitura','3. garnitura']
    items=''.join(
        f'<figure class="jersey-item"><img src="{esc(src)}" alt="{esc(club_name)} {labels[i]}"><figcaption>{labels[i]}</figcaption></figure>'
        for i,src in enumerate(kits)
    )
    return f'''<div class="clubhero-stadium jersey-panel"><div class="jersey-gallery">{items}</div><div class="stadium-caption jersey-caption"><strong>Dresovi 2026/27</strong><span>{len(kits)} dostupne garniture</span></div></div>'''


# Position groups for squad tables (source data uses Transfermarkt English labels).
_POS_GROUPS=(
    ('gk','Golmani','GK',('goalkeeper',)),
    ('df','Odbrana','DF',('back','defender','sweeper')),
    ('mf','Vezni red','MF',('midfield',)),
    ('fw','Napad','FW',('forward','striker','winger')),
)


def _pos_group(position):
    p=str(position or '').lower()
    for key,label,short,keys in _POS_GROUPS:
        if any(k in p for k in keys): return key
    return 'other'


def _age_from_dob(dob):
    """'21/05/2004' -> 22 (age today); returns '' when the date is unknown."""
    import datetime as _dt
    try:
        d,m,y=[int(x) for x in str(dob or '').strip()[:10].split('/')]
        t=_dt.date.today(); return t.year-y-((t.month,t.day)<(m,d))
    except Exception:
        return ''


def _squad_table(players, columns):
    """Players grouped GK -> DF -> MF -> FW with a heading row per group.
    columns: list of (header, cell_fn(row)->html, css_class)."""
    groups={k:[] for k,*_ in _POS_GROUPS}; groups['other']=[]
    for p in players: groups[_pos_group(p['position'])].append(p)
    head=''.join(f'<th class="{cls}">{esc(h)}</th>' for h,_,cls in columns)
    body=[]
    meta={k:(label,short) for k,label,short,_ in _POS_GROUPS}; meta['other']=('Ostali','—')
    for key in [k for k,*_ in _POS_GROUPS]+['other']:
        rows=groups[key]
        if not rows: continue
        label,_short=meta[key]
        body.append(f'<tr class="grouprow g-{key}"><th colspan="{len(columns)}">{esc(label)}<span>{len(rows)}</span></th></tr>')
        for r in rows:
            body.append('<tr>'+''.join(f'<td class="{cls}">{fn(r)}</td>' for _,fn,cls in columns)+'</tr>')
    return f'<div class="tablewrap"><table class="squadtable"><thead><tr>{head}</tr></thead><tbody>{"".join(body)}</tbody></table></div>'


def _pos_cell(r):
    key=_pos_group(r['position']); short=dict((k,s_) for k,_,s_,_ in _POS_GROUPS).get(key,'—')
    return f'<span class="pos pos-{key}">{short}</span>{esc(r["position"])}' if r['position'] else '—'


def _dob_cell(r):
    age=_age_from_dob(r['date_of_birth'])
    return esc(r['date_of_birth'])+(f'<span class="age">({age})</span>' if age!='' else '')


def _club_tabs(cid,active='overview'):
    items=[('overview','Pregled',f'/club?id={cid}'),('players','Igrači',f'/club?id={cid}#players'),('staff','Stručni štab',f'/club?id={cid}#staff'),('transfers','Transferi',f'/transfers?id={cid}'),('history','Istorija',f'/club/history?id={cid}')]
    return '<nav class="sectiontabs">'+''.join(f'<a class="{("active" if key==active else "")}" href="{href}">{esc(label)}</a>' for key,label,href in items)+'</nav>'


def _league_tabs(lid,active='overview'):
    items=[('overview','Pregled',f'/league?id={lid}'),('clubs','Klubovi',f'/league?id={lid}#clubs'),('table','Tabela',f'/league?id={lid}#table'),('history','Istorija',f'/league/history?id={lid}')]
    return '<nav class="sectiontabs">'+''.join(f'<a class="{("active" if key==active else "")}" href="{href}">{esc(label)}</a>' for key,label,href in items)+'</nav>'


def league(lid,msg=''):
    con=connect(); l=con.execute('SELECT * FROM leagues WHERE id=?',(int(lid),)).fetchone()
    if not l: con.close(); return layout('Not found','<h1>Liga nije pronađena</h1>')
    clubs=con.execute('''SELECT c.*,ci.crest_local,s.name stadium_name,s.capacity stadium_capacity,s.image_local stadium_image,
        (SELECT COUNT(*) FROM players p WHERE p.club_id=c.id) pc,(SELECT COUNT(*) FROM staff st WHERE st.club_id=c.id) sc
        FROM clubs c LEFT JOIN club_info ci ON ci.club_id=c.id LEFT JOIN stadiums s ON s.club_id=c.id
        WHERE c.league_id=? AND COALESCE(c.source_active,1)=1 ORDER BY COALESCE(c.cm_name,c.name)''',(int(lid),)).fetchall(); con.close()
    st=league_stats(league_ids=[int(lid)]).get(int(lid),{}); club_cards=[]
    premium=int(l['level'] or 0) in (1,2)
    for c in clubs:
        crest=f'<img src="/{esc(str(c["crest_local"]).replace(chr(92),"/").lstrip("/"))}" alt="">' if c['crest_local'] else '<span class="crestfallback">GRB</span>'
        name=_display_club_name(c)
        if premium:
            jersey_media=_jersey_card_visual(c['id'], name)
            stadium_media=''
            if jersey_media:
                stadium_media=jersey_media
            else:
                stadium_src=_stadium_asset(c['id'],c['stadium_image'])
                if stadium_src:
                    stadium_media=f'<img src="{esc(stadium_src)}" alt="{esc(_display_stadium_name(c["stadium_name"] or "Stadion"))}">'
            cap=((f'{int(c["stadium_capacity"]):,}' if engine.ui_lang()=='en' else f'{int(c["stadium_capacity"]):,}'.replace(',','.')) if c['stadium_capacity'] else '—')
            stadium_meta=f'<strong>{esc(_display_stadium_name(c["stadium_name"] or "Stadion"))}</strong><span>Kapacitet: {esc(cap)}</span>'
            side=f'<div class="clubcardstadium">{stadium_media}</div>' if stadium_media else ''
            club_cards.append(f'''<a class="clubcard clubcard-premium" href="/club?id={c["id"]}">
                <div class="clubcard-media{'' if side else ' solo'}"><div class="clubcardcrest">{crest}</div>{side}</div>
                <div class="clubcard-copy"><h3>{esc(name)}</h3><div class="clubcard-stadium-meta">{stadium_meta}</div><p>{int(c["pc"] or 0)} igrača · {int(c["sc"] or 0)} staff</p></div></a>''')
        else:
            club_cards.append(f'<a class="clubcard" href="/club?id={c["id"]}"><div class="clubcardcrest">{crest}</div><h3>{esc(name)}</h3><p>{int(c["pc"] or 0)} igrača · {int(c["sc"] or 0)} staff</p><small>{esc(c["tm_match_status"] or "")}</small></a>')
    hero=rich_card(title=_display_league_name(l['name']),subtitle=f'NIVO {l["level"]} · {str(l["source_type"] or "").upper()}',href='#clubs',stats=st,image='/'+l['image_path'] if l['image_path'] else None,cta='Klubovi ↓')
    exports=f'<div class="pageactions"><a class="button" href="/export_league_master?id={lid}">MASTER XLSX</a><a class="button secondary" href="/export_league_players?id={lid}">Igrači XLSX</a><a class="button secondary" href="/export_league_staff?id={lid}">Stručni štab XLSX</a><a class="button secondary" href="/export_league_crests?id={lid}">Grbovi ZIP</a></div>'
    history_cta=''
    if int(l['level'] or 0)<=3:
        history_cta=f'''<a class="league-history-teaser" href="/league/history?id={lid}"><div><small>ISTORIJA TAKMIČENJA</small><h2>Istorija lige</h2><p>Šampioni po sezonama · najtrofejniji klubovi</p></div><span>→</span></a>'''
    standings=_standings_block('league',lid).replace('<section class="standings">','<section class="standings" id="table">',1)
    league_name=_display_league_name(l['name'])
    return layout(league_name,notice(msg)+crumb([('Srbija','/'),(league_name,None)])+hero+_league_tabs(lid)+exports+history_cta+f'<h2 id="clubs">Klubovi</h2><div class="clubgrid">{"".join(club_cards)}</div>'+standings)


def club(cid,msg=''):
    con=connect(); c=con.execute('''SELECT c.*,l.name league_name,l.id league_id,l.level league_level FROM clubs c JOIN leagues l ON l.id=c.league_id WHERE c.id=?''',(int(cid),)).fetchone()
    if not c: con.close(); return layout('Not found','<h1>Klub nije pronađen</h1>')
    info=con.execute('SELECT * FROM club_info WHERE club_id=?',(int(cid),)).fetchone(); stadium=con.execute('SELECT * FROM stadiums WHERE club_id=?',(int(cid),)).fetchone()
    players=con.execute('SELECT * FROM players WHERE club_id=? ORDER BY position,name',(int(cid),)).fetchall(); staff=con.execute('SELECT * FROM staff WHERE club_id=? ORDER BY section,position,name',(int(cid),)).fetchall()
    season_id=engine.current_season_id(); transfer_summary=con.execute('SELECT * FROM club_transfer_summary WHERE club_id=? AND season_id=?',(int(cid),season_id)).fetchone(); transfer_count=con.execute('SELECT COUNT(*) FROM club_transfers WHERE club_id=? AND season_id=?',(int(cid),season_id)).fetchone()[0]
    con.close()
    club_name=_display_club_name(c)
    crest=f'<img class="club-profile-crest" src="{esc(_crest_url(info))}" alt="{esc(club_name)}">' if _crest_url(info) else '<div class="club-profile-crest crestfallback">GRB NIJE DOSTUPAN</div>'
    city=(info['city'] or info['postal_city']) if info else ''
    meta=[]
    if info and info['founded']: meta.append(f'<span><b>Osnovan</b>{esc(info["founded"])}</span>')
    if city: meta.append(f'<span><b>Grad</b>{esc(city)}</span>')
    if info and info['website']:
        site=str(info['website']); href=site if site.startswith('http') else 'https://'+site
        meta.append(f'<span><b>Zvanični sajt</b><a href="{esc(href)}" target="_blank" rel="noopener">{esc(site.replace("https://","").replace("http://","").rstrip("/"))}</a></span>')
    club_data=''.join(meta) or '<p>Podaci o klubu još nisu učitani.</p>'
    stadium_name=_display_stadium_name(stadium['name'] if stadium and stadium['name'] else 'Stadion')
    stadium_capacity=((f'{int(stadium["capacity"]):,}' if engine.ui_lang()=='en' else f'{int(stadium["capacity"]):,}'.replace(',','.')) if stadium and stadium['capacity'] else '—')
    stadium_panel=_jersey_panel(cid,club_name)
    if not stadium_panel and _stadium_asset(cid,(stadium['image_local'] if stadium and 'image_local' in stadium.keys() else '')):
        stadium_panel=f'''<div class="clubhero-stadium">{_stadium_visual(stadium,cid)}<div class="stadium-caption"><strong>{esc(stadium_name)}</strong><span>Kapacitet: {esc(stadium_capacity)}</span></div></div>'''
    if not stadium_panel and stadium and stadium['name']:
        # no kit and no photo: show the stadium as a plain fact instead of an empty placeholder
        meta.append(f'<span><b>Stadion</b>{esc(stadium_name)} · {esc(stadium_capacity)}</span>')
        club_data=''.join(meta)
    players_table=_squad_table(players,[
        ('#',lambda r: esc(r['squad_number']) or '—',''),
        ('Ime',lambda r: esc(r['name']),'name'),
        ('Pozicija',_pos_cell,''),
        ('Rođen',_dob_cell,'hide-sm'),
        ('Nacionalnost',lambda r: esc(r['nationality']),'hide-sm'),
        ('Vrednost',lambda r: esc(r['market_value']) or '—','num'),
    ])
    s_rows=''.join(f'<tr><td>{esc(x["section"])}</td><td>{esc(x["name"])}</td><td>{esc(x["position"])}</td><td>{esc(x["date_of_birth"])}</td><td>{esc(x["nationality"])}</td></tr>' for x in staff)
    exports=f'<div class="pageactions compact-actions"><a class="button" href="/export_club_players?id={cid}">Igrači XLSX</a><a class="button secondary" href="/export_club_staff?id={cid}">Stručni štab XLSX</a></div>'
    if transfer_summary:
        transfer_meta=(f'<p>{int(transfer_summary["arrivals_count"] or 0)} dolazaka · {int(transfer_summary["departures_count"] or 0)} odlazaka</p><strong>Bilans {esc(transfer_summary["balance"] or "—")}</strong>')
    elif transfer_count:
        transfer_meta=f'<p>{int(transfer_count)} transfer zapisa u sezoni {esc(engine.season_label_from_id(season_id))}</p><strong>OTVORI TRANSFERE →</strong>'
    else:
        transfer_meta='<p>Transferi još nisu učitani za aktuelnu sezonu.</p><strong>OTVORI TRANSFERE →</strong>'
    hsum=history_data.club_summary(cid)
    first_hist,_=history_data.club_year_span(cid)
    history_start=str(first_hist) if first_hist else '—'
    history_meta=(f'<p>{int(hsum.get("seasons") or 0)} sezona u bazi · {int(hsum.get("league_titles") or 0)} ligaških titula · {int(hsum.get("cups") or 0)} kupova</p><strong>OTVORI ISTORIJU →</strong>')
    modules=f'''<div class="club-module-grid">
        <a class="club-module transfer-module" href="/transfers?id={cid}"><div class="module-bg transfer-bg"></div><div class="module-copy"><small>AKTUELNA SEZONA {esc(engine.season_label_from_id(season_id))}</small><h2>Transferi</h2>{transfer_meta}</div></a>
        <a class="club-module history-module" href="/club/history?id={cid}"><div class="history-mark"><span>{esc(history_start)}</span><i></i><span>DANAS</span></div><div class="module-copy"><small>ISTORIJA KLUBA</small><h2>Istorija</h2>{history_meta}</div></a>
    </div>'''
    league_name=_display_league_name(c['league_name'])
    hero=f'''<section class="clubhero clubhero-v2{'' if stadium_panel else ' no-side'}"><div class="clubhero-crest">{crest}</div><div class="clubhero-copy"><small>{esc(league_name)}</small><h1>{esc(club_name)}</h1><div class="clubfacts">{club_data}</div></div>{stadium_panel}</section>'''
    body=notice(msg)+crumb([(league_name,f'/league?id={c["league_id"]}'),(club_name,None)])+hero+_club_tabs(cid)+exports+modules+f'<h2 id="players"><span>Igrači</span> <small>({len(players)})</small></h2>{players_table}<h2 id="staff"><span>Stručni štab</span> <small>({len(staff)})</small></h2><div class="tablewrap"><table><thead><tr><th>Sekcija</th><th>Ime</th><th>Uloga</th><th>DOB</th><th>Nacionalnost</th></tr></thead><tbody>{s_rows}</tbody></table></div>'
    return layout(club_name,body)


def club_history(club_id,msg=''):
    con=connect(); c=con.execute('''SELECT c.*,l.name league_name,l.id league_id FROM clubs c JOIN leagues l ON l.id=c.league_id WHERE c.id=?''',(int(club_id),)).fetchone(); info=con.execute('SELECT * FROM club_info WHERE club_id=?',(int(club_id),)).fetchone() if c else None; con.close()
    if not c: return layout('Not found','<h1>Klub nije pronađen</h1>')
    name=_display_club_name(c); rows=history_data.club_seasons(club_id); sm=history_data.club_summary(club_id)
    honour_groups=history_data.club_honour_groups(club_id); international=history_data.club_international_highlights(club_id); major_totals=history_data.club_major_totals(club_id)
    crest=f'<img src="{esc(_crest_url(info))}" alt="{esc(name)}">' if _crest_url(info) else '<span class="crestfallback">GRB</span>'
    metrics=metric_strip([('SEZONE',int(sm.get('seasons') or 0)),('DRŽAVNA PRVENSTVA',major_totals['national_titles']),('NACIONALNI KUPOVI',major_totals['national_cups']),('VICEŠAMPION',int(sm.get('league_runner_up') or 0)),('MEĐUNARODNO',int(sm.get('international') or 0))])
    high=[]
    if major_totals['national_titles']:
        high.append(f'<article class="honour-card total league">{_trophy_media("","league","Državna prvenstva")}<small>UKUPNO</small><h3>Državna prvenstva</h3><strong>{major_totals["national_titles"]} ×</strong></article>')
    if major_totals['national_cups']:
        high.append(f'<article class="honour-card total cup">{_trophy_media("","cup","Nacionalni kupovi")}<small>UKUPNO</small><h3>Nacionalni kupovi</h3><strong>{major_totals["national_cups"]} ×</strong></article>')
    # International achievements stay individual: European/world titles and
    # notable finals/stages are milestones, not domestic-style repeated counts.
    for r in international:
        comp=history_data.comp_name(r); label=history_data.result_label(dict(r)); kind=(r['comp_type'] or '')
        high.append(f'<article class="honour-card international {esc(kind)}">{_trophy_media(r["history_key"] if "history_key" in r.keys() else "",kind,comp)}<small>{esc(r["season_label"] or "")}</small><h3>{esc(comp)}</h3><strong>{esc(label)}</strong></article>')
    for r in honour_groups:
        comp=history_data.comp_name(r); wins=int(r['title_count'] or 0); runners=int(r['runner_up_count'] or 0); kind=(r['comp_type'] or '')
        labels=[]
        if wins:
            win_label='Osvajač' if kind=='cup' else 'Prvak'
            labels.append(f'{win_label} {wins} ×')
        if runners:
            run_label='Vicešampion' if kind=='league' else 'Finalista'
            labels.append(f'{run_label} {runners} ×')
        span=(r['first_season'] or '') if r['first_season']==r['last_season'] else f'{r["first_season"] or ""}–{r["last_season"] or ""}'
        high.append(f'<article class="honour-card {esc(kind)}">{_trophy_media(r["history_key"] if "history_key" in r.keys() else "",kind,comp)}<small>{esc(span)}</small><h3>{esc(comp)}</h3><strong>{esc(" · ".join(labels))}</strong></article>')
    highlight_block='<div class="honours-grid">'+''.join(high)+'</div>' if high else '<div class="history-empty"><p>Nema još unetih istorijskih sezona za ovaj klub.</p><small>Istorijski model je spreman; podaci će se dopunjavati iz proverenih izvora.</small></div>'
    # Public timeline is one row per season. The normalized DB may contain
    # separate league/cup/Europe rows for the same year.
    timeline=history_data.club_season_timeline(club_id)
    body_rows=[]
    for season in timeline:
        leagues=season['league']
        primary=leagues[0] if leagues else None
        league_parts=[]
        for lr in leagues:
            nm=history_data.comp_name(lr)
            league_parts.append(esc(nm))
        league_cell='<br>'.join(league_parts) if league_parts else '—'
        tier=(primary.get('level') if primary else None) or '—'
        if primary:
            pos='🏆 1' if primary.get('is_champion') else ('🥈 2' if primary.get('is_runner_up') else (str(primary.get('position')) if primary.get('position') else '—'))
        else:
            pos='—'
        def comp_results(items):
            bits=[]
            for rr in items:
                result=history_data.result_label(rr)
                bits.append(f'<span class="history-result"><b>{esc(history_data.comp_name(rr))}</b><small>{esc(result)}</small></span>')
            return ''.join(bits) if bits else '—'
        cup_cell=comp_results(season['cup'])
        eur_cell=comp_results(season['europe'])
        body_rows.append(f'<tr><td><b>{esc(season["season_label"])}</b></td><td>{league_cell}</td><td>{esc(tier)}</td><td>{esc(pos)}</td><td>{cup_cell}</td><td>{eur_cell}</td></tr>')
    table=('<div class="tablewrap history-season-table"><table><thead><tr><th>Sezona</th><th>Liga</th><th>Rang</th><th>Plasman</th><th>Kup</th><th>Evropa / svet</th></tr></thead><tbody>'+''.join(body_rows)+'</tbody></table></div>') if body_rows else ''
    hero=f'<section class="history-hero"><div class="history-crest">{crest}</div><div><small>ISTORIJA KLUBA</small><h1>{esc(name)}</h1><p>Sezonski plasmani kroz sva takmičenja, nezavisno od trenutnog ranga kluba.</p></div></section>'
    league_name=_display_league_name(c['league_name'])
    page=notice(msg)+crumb([(league_name,f'/league?id={c["league_id"]}'),(name,f'/club?id={club_id}'),('Istorija',None)])+hero+_club_tabs(club_id,'history')+metrics+'<h2>Najveći uspesi</h2>'+highlight_block+'<h2>Plasmani</h2>'+table
    return layout(name+' · Istorija',page)


def league_history(lid,msg=''):
    con=connect(); l=con.execute('SELECT * FROM leagues WHERE id=?',(int(lid),)).fetchone(); con.close()
    if not l: return layout('Not found','<h1>Liga nije pronađena</h1>')
    ranking=history_data.league_title_ranking(lid); champs=history_data.league_champions(lid); season_notes=history_data.league_season_notes(lid)
    def club_history_link(cid,label):
        shown=esc(_display_club_name(label or '—'))
        return f'<a href="/club?id={int(cid)}">{shown}</a>' if cid is not None else shown
    rank_html=''.join(f'<div class="title-rank-row"><span>{i}</span><strong>{club_history_link(r["club_id"],r["club_name"])}</strong><b>{int(r["titles"])} ×</b></div>' for i,r in enumerate(ranking,1))
    rank_block=f'<div class="title-ranking">{rank_html}</div>' if rank_html else '<div class="history-empty"><p>Nema još unetih istorijskih sezona za ovo takmičenje.</p><small>Istorijski model je spreman; podaci će se dopunjavati iz proverenih izvora.</small></div>'
    champ_items=[]
    for r in champs:
        raw=r['club_name_en'] if engine.ui_lang()=='en' else r['club_name_sr']; raw=raw or r['club_name_override'] or '—'
        champ_items.append((r['season_label'],0,f'<tr><td><b>{esc(r["season_label"])}</b></td><td>{club_history_link(r["club_id"],raw)}</td><td>{esc(history_data.comp_name(r))}</td></tr>'))
    for n in season_notes:
        note=n['note_en'] if engine.ui_lang()=='en' else n['note_sr']
        champ_items.append((n['season_label'],1,f'<tr class="history-note-row"><td><b>{esc(n["season_label"])}</b></td><td>—</td><td><em>{esc(note or "—")}</em></td></tr>'))
    champ_items.sort(key=lambda x:(int(str(x[0])[:4]) if str(x[0])[:4].isdigit() else -1,-x[1]),reverse=True)
    champ_rows=[x[2] for x in champ_items]
    champ_table=('<div class="tablewrap"><table><thead><tr><th>Sezona</th><th>Šampion</th><th>Takmičenje / status</th></tr></thead><tbody>'+''.join(champ_rows)+'</tbody></table></div>') if champ_rows else '<div class="history-empty compact"><p>Istorijski podaci još nisu učitani.</p></div>'
    league_name=_display_league_name(l['name'])
    hmeta=history_data.league_history_meta(lid)
    era=(hmeta['active_from'] if hmeta and hmeta['active_from'] else '—')
    hero=f'<section class="history-hero league-history-hero"><div class="history-year">{esc(era)}<br><span>→</span><br>DANAS</div><div><small>ISTORIJA TAKMIČENJA</small><h1>{esc(league_name)}</h1><p>Šampioni po sezonama i najtrofejniji klubovi iz zajedničkog History Excel mastera.</p></div></section>'
    page=notice(msg)+crumb([('Srbija','/'),(league_name,f'/league?id={lid}'),('Istorija',None)])+hero+_league_tabs(lid,'history')+'<h2>Najtrofejniji klubovi</h2>'+rank_block+'<h2>Šampioni po sezonama</h2>'+champ_table
    return layout(league_name+' · '+engine.tr_text('Istorija',engine.ui_lang()),page)


def national_team(team_id,msg=''):
    con=connect(); t=con.execute('SELECT * FROM national_teams WHERE id=?',(int(team_id),)).fetchone()
    if not t: con.close(); return layout('Not found','<h1>Reprezentacija nije pronađena</h1>')
    players=con.execute('SELECT * FROM national_team_players WHERE national_team_id=? ORDER BY position,name',(int(team_id),)).fetchall(); staff=con.execute('SELECT * FROM national_team_staff WHERE national_team_id=? ORDER BY section,position,name',(int(team_id),)).fetchall(); con.close()
    img='/'+t['image_path'] if t['image_path'] else '/static/branding/fss.jpg'
    players_table=_squad_table(players,[
        ('Ime',lambda r: esc(r['name']),'name'),
        ('Pozicija',_pos_cell,''),
        ('Rođen',_dob_cell,'hide-sm'),
        ('Klub',lambda r: esc(r['current_club']) or '—',''),
    ])
    s=''.join(f'<tr><td>{esc(x["section"])}</td><td>{esc(x["name"])}</td><td>{esc(x["position"])}</td><td>{esc(x["date_of_birth"])}</td></tr>' for x in staff)
    exports=f'<div class="pageactions"><a class="button" href="/export_national_players?id={team_id}">Igrači XLSX</a><a class="button secondary" href="/export_national_staff?id={team_id}">Stručni štab XLSX</a></div>'
    # Self-heal configured national-team standings. If the table is missing or
    # empty, refresh it from this running portal instance before rendering.
    _comp,_rows=engine.standings_for_owner('national_team',team_id)
    if t['standings_source_url'] and (not _comp or not _rows):
        try:
            engine.refresh_national_team_standings(int(team_id))
        except Exception:
            pass
    standings=_standings_block('national_team',team_id)
    if not t['standings_source_url'] and (t['standings_competition_name'] or ''):
        standings=f'<section class="standings"><h2>Tabela</h2><div class="panel"><p>{esc(t["standings_competition_name"])} trenutno nema ligašku tabelu za prikaz.</p></div></section>'
    comp_label=''
    if t['standings_competition_name']:
        comp_label=f'<p><b>{esc(t["standings_competition_name"])}</b>{(" · Grupa "+esc(t["standings_group_code"])) if t["standings_group_code"] else ""}</p>'
    body=notice(msg)+f'<section class="teamhero"><img src="{esc(img)}"><div><small>NACIONALNA SELEKCIJA</small><h1>{esc(t["name"])}</h1><p>{len(players)} igrača · {len(staff)} članova stručnog štaba</p>{comp_label}</div></section>{exports}{standings}<h2>Igrači <small>({len(players)})</small></h2>{players_table}<h2>Stručni štab</h2><div class="tablewrap"><table><thead><tr><th>Sekcija</th><th>Ime</th><th>Uloga</th><th>DOB</th></tr></thead><tbody>{s}</tbody></table></div>'
    return layout(t['name'],body)


def transfers(club_id,msg=''):
    con=connect(); c=con.execute('SELECT c.*,l.name league_name FROM clubs c JOIN leagues l ON l.id=c.league_id WHERE c.id=?',(int(club_id),)).fetchone()
    if not c: con.close(); return layout('Not found','<h1>Klub nije pronađen</h1>')
    sid=engine.current_season_id(); rows=con.execute('SELECT * FROM club_transfers WHERE club_id=? AND season_id=? ORDER BY direction,transfer_date,player_name',(int(club_id),sid)).fetchall(); summary=con.execute('SELECT * FROM club_transfer_summary WHERE club_id=? AND season_id=?',(int(club_id),sid)).fetchone(); con.close()
    tr=''.join(f'<tr><td>{esc(r["direction"])}</td><td>{esc(r["player_name"])}</td><td>{esc(r["position"])}</td><td>{esc(r["other_club"])}</td><td>{esc(r["fee"])}</td><td>{esc(r["transfer_date"])}</td></tr>' for r in rows)
    exp=f'<div class="pageactions"><a class="button" href="/export_club_transfers?id={club_id}">Transferi XLSX</a></div>'
    summ=f'<div class="panel"><b>{esc(summary["season_label"] if summary else engine.season_label_from_id(sid))}</b> · {len(rows)} transfer zapisa</div>'
    return layout('Transferi',notice(msg)+crumb([(c['cm_name'] or c['name'],f'/club?id={c["id"]}'),('Transferi',None)])+f'<h1>Transferi · {esc(c["cm_name"] or c["name"])}</h1>{exp}{summ}<div class="tablewrap"><table><thead><tr><th>Smer</th><th>Igrač</th><th>Pozicija</th><th>Drugi klub</th><th>Naknada</th><th>Datum</th></tr></thead><tbody>{tr}</tbody></table></div>')


def mapping_workbench(cid=1,level=4,msg=''):
    country,rows=engine.mapping_workbench_rows(int(cid),int(level))
    if not country: return layout('Not found','<h1>Country not found</h1>')
    con=connect(); total=con.execute('''SELECT COUNT(*) FROM clubs c JOIN leagues l ON l.id=c.league_id WHERE l.country_id=? AND l.level=? AND COALESCE(c.source_active,1)=1''',(cid,level)).fetchone()[0]; mapped=con.execute("SELECT COUNT(*) FROM clubs c JOIN leagues l ON l.id=c.league_id WHERE l.country_id=? AND l.level=? AND COALESCE(c.source_active,1)=1 AND COALESCE(c.tm_url,'')<>''",(cid,level)).fetchone()[0]; con.close()
    review=sum((r['tm_match_status'] or '').lower()=='review' for r in rows); no_tm=sum((r['tm_match_status'] or '').lower()=='no_tm_yet' for r in rows); unresolved=len(rows)-review-no_tm
    trs=[]
    for r in rows:
        status=(r['tm_match_status'] or 'unresolved').lower(); label='NO TM YET' if status=='no_tm_yet' else status.upper(); candidate='—'
        if r['tm_candidate_url']:
            candidate=f'<a href="{esc(r["tm_candidate_url"])}" target="_blank">{esc(r["tm_candidate_name"] or r["tm_candidate_url"])}</a>'
        actions=f'''<div class="inlineactions"><form method="post" action="/mapping/retry"><input type="hidden" name="club_id" value="{r['id']}"><input type="hidden" name="country_id" value="{cid}"><input type="hidden" name="level" value="{level}"><button class="secondary">Retry</button></form>'''
        if r['tm_candidate_url']: actions+=f'''<form method="post" action="/mapping/accept"><input type="hidden" name="club_id" value="{r['id']}"><input type="hidden" name="country_id" value="{cid}"><input type="hidden" name="level" value="{level}"><button>Accept</button></form>'''
        actions+=f'''<form method="post" action="/mapping/no_tm"><input type="hidden" name="club_id" value="{r['id']}"><input type="hidden" name="country_id" value="{cid}"><input type="hidden" name="level" value="{level}"><button class="danger">NO TM YET</button></form></div>'''
        manual=f'''<form class="manualmap" method="post" action="/mapping/manual"><input type="hidden" name="club_id" value="{r['id']}"><input type="hidden" name="country_id" value="{cid}"><input type="hidden" name="level" value="{level}"><input name="tm_url" placeholder="Transfermarkt URL"><button>Save</button></form>'''
        trs.append(f'<tr><td>{esc(r["league_name"])}</td><td><b>{esc(r["source_name"] or r["name"])}</b></td><td><span class="matchbadge {esc(status)}">{esc(label)}</span></td><td>{candidate}</td><td>{actions}</td><td>{manual}</td></tr>')
    tools=f'''<div class="pageactions"><form data-busy="1" method="post" action="/mapping/retry_level"><input type="hidden" name="country_id" value="{cid}"><input type="hidden" name="level" value="{level}"><button>Retry ALL unresolved/review</button></form><form data-busy="1" method="post" action="/mapping/recheck_no_tm"><input type="hidden" name="country_id" value="{cid}"><input type="hidden" name="level" value="{level}"><button class="secondary">Recheck NO TM YET</button></form><a class="button secondary" href="/export_level_mapping_issues?country_id={cid}&level={level}">Open Mapping XLSX</a><a class="button secondary" href="/export_level_mapping_backup?country_id={cid}&level={level}">Backup ZIP</a><form method="post" action="/mapping/import" enctype="multipart/form-data"><input type="hidden" name="country_id" value="{cid}"><input type="hidden" name="level" value="{level}"><input type="file" name="mapping_file" accept=".xlsx" required><button class="secondary">Import XLSX</button></form></div>'''
    body=notice(msg)+crumb([('SUPERADMIN',f'/admin/sync?country_id={cid}'),(f'Mapping Level {level}',None)])+f'<h1>Mapping Workbench · Level {level}</h1>'+metric_strip([('AKTIVNI',total),('MAPPED',mapped),('REVIEW',review),('UNRESOLVED',unresolved),('NO TM',no_tm)])+tools+f'<div class="tablewrap"><table><thead><tr><th>Liga</th><th>Klub</th><th>Status</th><th>Kandidat</th><th>Akcije</th><th>Ručni URL</th></tr></thead><tbody>{"".join(trs) if trs else "<tr><td colspan=\"6\">Svi aktivni klubovi su verifikovani.</td></tr>"}</tbody></table></div>'
    return layout('Mapping Workbench',body,admin=True)
