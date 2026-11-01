from __future__ import annotations
from urllib.parse import quote
from .data import connect
from .ui import layout, notice, crumb, esc
from . import sync_engine as engine
from . import history_sync
from . import history_l12
from . import history_excel
from .config import VERSION, DB_PATH, PORT


def page(cid=1,msg='',details=False):
    con=connect()
    # Fast dashboard query: aggregate club status before joining to leagues.
    # Never join players x staff here; that creates a multiplicative temporary result.
    leagues=[dict(r) for r in con.execute('''WITH cc AS (
      SELECT c.league_id,
        SUM(CASE WHEN COALESCE(c.source_active,1)=1 THEN 1 ELSE 0 END) clubs,
        SUM(CASE WHEN COALESCE(c.source_active,1)=1 AND COALESCE(c.tm_url,'')<>'' THEN 1 ELSE 0 END) mapped,
        SUM(CASE WHEN COALESCE(c.source_active,1)=1 AND c.tm_match_status='review' THEN 1 ELSE 0 END) review,
        SUM(CASE WHEN COALESCE(c.source_active,1)=1 AND c.tm_match_status='no_tm_yet' THEN 1 ELSE 0 END) no_tm
      FROM clubs c GROUP BY c.league_id
    )
    SELECT l.*,COALESCE(cc.clubs,0) clubs,COALESCE(cc.mapped,0) mapped,
      COALESCE(cc.review,0) review,COALESCE(cc.no_tm,0) no_tm
    FROM leagues l LEFT JOIN cc ON cc.league_id=l.id
    WHERE l.country_id=? AND COALESCE(l.source_active,1)=1
    ORDER BY COALESCE(l.section,'senior'),l.level,l.name''',(int(cid),)).fetchall()]

    totals=dict(con.execute('''SELECT
      (SELECT COUNT(*) FROM clubs c JOIN leagues l ON l.id=c.league_id
         WHERE l.country_id=? AND COALESCE(l.source_active,1)=1 AND COALESCE(c.source_active,1)=1) clubs,
      (SELECT COUNT(*) FROM clubs c JOIN leagues l ON l.id=c.league_id
         WHERE l.country_id=? AND COALESCE(l.source_active,1)=1 AND COALESCE(c.source_active,1)=1 AND COALESCE(c.tm_url,'')<>'') mapped,
      (SELECT COUNT(*) FROM players p JOIN clubs c ON c.id=p.club_id JOIN leagues l ON l.id=c.league_id
         WHERE l.country_id=? AND COALESCE(l.source_active,1)=1 AND COALESCE(c.source_active,1)=1) players,
      (SELECT COUNT(*) FROM staff s JOIN clubs c ON c.id=s.club_id JOIN leagues l ON l.id=c.league_id
         WHERE l.country_id=? AND COALESCE(l.source_active,1)=1 AND COALESCE(c.source_active,1)=1) staff''',(int(cid),int(cid),int(cid),int(cid))).fetchone())

    teams=con.execute('''SELECT nt.*,
      (SELECT COUNT(*) FROM national_team_players p WHERE p.national_team_id=nt.id) players,
      (SELECT COUNT(*) FROM national_team_staff s WHERE s.national_team_id=nt.id) staff
      FROM national_teams nt WHERE nt.country_id=? ORDER BY nt.id''',(int(cid),)).fetchall()

    # Detailed per-league player/staff counts are loaded only on demand.
    player_counts={}
    staff_counts={}
    if details:
        player_counts={int(r[0]):int(r[1]) for r in con.execute('''SELECT c.league_id,COUNT(p.id)
          FROM clubs c JOIN leagues l ON l.id=c.league_id LEFT JOIN players p ON p.club_id=c.id
          WHERE l.country_id=? AND COALESCE(l.source_active,1)=1 AND COALESCE(c.source_active,1)=1 GROUP BY c.league_id''',(int(cid),)).fetchall()}
        staff_counts={int(r[0]):int(r[1]) for r in con.execute('''SELECT c.league_id,COUNT(s.id)
          FROM clubs c JOIN leagues l ON l.id=c.league_id LEFT JOIN staff s ON s.club_id=c.id
          WHERE l.country_id=? AND COALESCE(l.source_active,1)=1 AND COALESCE(c.source_active,1)=1 GROUP BY c.league_id''',(int(cid),)).fetchall()}
    con.close()

    total=lambda key:int(totals.get(key,0) or 0)
    top=[r for r in leagues if (r.get('section') or 'senior')=='senior' and int(r.get('level') or 0) in (1,2)]
    lower=[r for r in leagues if (r.get('section') or 'senior')=='senior' and int(r.get('level') or 0) in (3,4,5,6)]
    youth=[r for r in leagues if (r.get('section') or '')=='youth']
    women=[r for r in leagues if (r.get('section') or '')=='women']
    def section_summary(items):
        return f"{len(items)} liga · {sum(int(x.get('clubs') or 0) for x in items)} klubova"

    league_opts=''.join(f'<option value="{r["id"]}">{esc(engine.tr_text(r.get("section") or "senior",engine.ui_lang()))} · L{r.get("level")} · {esc(engine.display_league_name(r["name"]))}</option>' for r in leagues)
    levels=sorted({int(r['level']) for r in leagues if r.get('level') is not None and (r.get('section') or '')=='senior'})
    level_opts=''.join(f'<option value="{x}">Nivo {x}</option>' for x in levels)
    rows=''
    if details:
        rows=''.join(f'<tr><td>{esc(engine.tr_text(r.get("section") or "senior",engine.ui_lang()))}</td><td>{r.get("level")}</td><td>{esc(engine.display_league_name(r["name"]))}</td><td>{r.get("clubs",0)}</td><td>{r.get("mapped",0)}/{r.get("clubs",0)}</td><td>{r.get("review",0)}</td><td>{r.get("no_tm",0)}</td><td>{player_counts.get(int(r["id"]),0)}</td><td>{staff_counts.get(int(r["id"]),0)}</td></tr>' for r in leagues)
    def _team_status(t):
        comp=esc(t['standings_competition_name'] or '')
        grp=esc(t['standings_group_code'] or '')
        if comp and grp: return f'{comp} · Grupa {grp}'
        if comp: return comp
        return 'Tabela nije konfigurisana'
    nt_all=''.join(f'<div class="admin-team-row"><b>{esc(t["name"])}</b><span>{t["players"]} igrača · {t["staff"]} staff · {_team_status(t)}</span></div>' for t in teams)
    team_opts=''.join(f'<option value="{t["id"]}">{esc(t["name"])}</option>' for t in teams)
    def form(action,label,extra=''):
        return f'<form data-busy="1" method="post" action="{action}"><input type="hidden" name="country_id" value="{cid}">{extra}<button>{esc(label)}</button></form>'

    if details:
        status_table=f'''<section class="adminpanel"><small>DIJAGNOSTIKA</small><h2>Status svih liga</h2><p>Detaljna tabela je učitana na zahtev.</p><div class="tablewrap"><table><thead><tr><th>Sekcija</th><th>Nivo</th><th>Liga</th><th>Klubovi</th><th>TM</th><th>Review</th><th>NO TM</th><th>Igrači</th><th>Staff</th></tr></thead><tbody>{rows}</tbody></table></div><div class="pageactions"><a class="button secondary" href="/admin/sync?country_id={cid}">Sakrij detaljni status</a></div></section>'''
    else:
        status_table=f'''<section class="adminpanel"><small>DIJAGNOSTIKA</small><h2>Status svih liga</h2><p>Detaljni brojači po svih {len(leagues)} liga više se ne računaju pri svakom ulasku u Superadmin. Učitaj ih samo kada su potrebni.</p><div class="pageactions"><a class="button secondary" href="/admin/sync?country_id={cid}&details=1">Učitaj detaljni status liga</a></div></section>'''

    body=notice(msg)+crumb([('Početna','/'),('SUPERADMIN',None)])+f'''
<section class="adminhero"><small>SUPERADMIN · DATA OPERATIONS</small><h1>Touchline Centar sinhronizacije</h1><p>GLOBAL obuhvata sve lige. Svi ligaški blokovi koriste isti prioritet i isti numerisani redosled 01–07. Generalna noćna sinhronizacija pokreće kompletan ciklus jednim klikom.<br><span class="muted">Runtime: build {VERSION} · port {PORT} · DB {esc(str(DB_PATH))}</span></p></section>
<div class="adminstats"><div><span>LIGE</span><strong>{len(leagues)}</strong></div><div><span>KLUBOVI</span><strong>{total('clubs')}</strong></div><div><span>TM MAPPED</span><strong>{total('mapped')}/{total('clubs')}</strong></div><div><span>IGRAČI</span><strong>{total('players')}</strong></div><div><span>STAFF</span><strong>{total('staff')}</strong></div></div>

<section class="adminpanel dangerpanel"><small>GENERALNA SINHRONIZACIJA</small><h2>Generalno čišćenje i update</h2><p>Jedno dugme za kompletno noćno osvežavanje. Redosled: 01 sastavi → 02 TM mapping → 03 igrači/staff → 04 transferi → 05 grbovi → 06 informacije o klubovima → 07 tabele → reprezentacije.</p><div class="adminactions singleaction">
{form('/admin/action/night_sync','GENERALNO ČIŠĆENJE I UPDATE · NOĆNA SINHRONIZACIJA · 01→07')}
</div></section>

<section class="adminpanel"><small>GLOBAL · SVE LIGE</small><h2>Centralne akcije · isti redosled 01–07</h2><div class="adminactions">
{form('/admin/action/all_memberships','01 · SASTAVI SVIH LIGA')}
{form('/admin/action/all_tm','02 · TM MAPPING SVIH LIGA')}
{form('/admin/action/all_data','03 · IGRAČI + STAFF','<label class="check"><input type="checkbox" name="fetch_dob" value="1"> DOB stručnog štaba</label>')}
{form('/admin/action/all_transfers','04 · TRANSFERI')}
{form('/admin/action/all_crests','05 · GRBOVI')}
{form('/admin/action/all_info','06 · INFORMACIJE O KLUBOVIMA')}
{form('/admin/action/standings_all','07 · TABELE')}
</div></section>

<section class="adminpanel nationaladmin"><small>NACIONALNE SELEKCIJE</small><h2>Reprezentacije · rosteri, staff i tabele</h2><p>Sve operacije reprezentacija su na jednom mestu. Za reprezentacije su primenljive 01 (roster/staff) i 07 (tabela); klupske operacije 02–06 se ne primenjuju. Srbija A je automatski vezana za UEFA Nations League A · Grupa A2; U21 trenutno nema ligašku tabelu.</p><div class="adminactions">
{form('/admin/action/national','01 · REPREZENTACIJE · ROSTERI + STAFF','<label class="check"><input type="checkbox" name="fetch_dob" value="1"> DOB stručnog štaba</label>')}
{form('/admin/action/national_standings','07 · REPREZENTACIJE · TABELE')}
</div>{nt_all}<details style="margin-top:14px"><summary>Napredni override izvora tabele</summary><form method="post" action="/admin/action/national_standings_source"><input type="hidden" name="country_id" value="{cid}"><select name="team_id">{team_opts}</select><input name="standings_url" placeholder="FSS / UEFA / Transfermarkt URL za override" style="min-width:420px"><button>Sačuvaj / promeni izvor</button></form></details></section>

<section class="adminpanel"><small>TOP TIER · NIVO 1–2</small><h2>SuperLiga + Prva liga Srbije</h2><p>{section_summary(top)}</p><div class="adminactions">
{form('/admin/action/top_memberships','01 · SASTAVI LIGA')}
{form('/admin/action/top_tm','02 · TM MAPPING')}
{form('/admin/action/top_data','03 · IGRAČI + STAFF')}
{form('/admin/action/top_transfers','04 · TRANSFERI')}
{form('/admin/action/top_crests','05 · GRBOVI')}
{form('/admin/action/top_info','06 · INFORMACIJE O KLUBOVIMA')}
{form('/admin/action/top_standings','07 · TABELE')}
</div></section>

<section class="adminpanel historyadmin"><small>ISTORIJSKA BAZA · EXCEL MASTER</small><h2>Prvenstvo · Kup · Drugi rang · Srpske lige</h2><p>Od verzije 1.6.4 kompletna DOMAĆA LIGAŠKA istorija u History sloju ima samo jedan dozvoljeni izvor: zajednički Excel <b>HISTORY/TOUCHLINE_HISTORY_MASTER.xlsx</b>, isti format koji će koristiti i Touchline Manager.</p>
{form('/admin/action/history_excel_import','ISTORIJA · UVEZI / OSVEŽI MASTER EXCEL')}
{form('/admin/action/history_excel_validate','ISTORIJA · PROVERI MASTER EXCEL')}
<p class="muted">LIGA = isključivo Excel. RSSSF/Wikipedia/COMET ne mogu više da upisuju ligaške redove u klupsku istoriju. Aktuelne tabele, sastavi, transferi i evropski uspesi ostaju odvojeni i ne diraju se.</p></section>

<section class="adminpanel"><small>NIŽE SENIORSKE · NIVO 3–6</small><h2>Srpske · Zonske · Okružne · Opštinske</h2><p>{section_summary(lower)}</p><div class="adminactions">
{form('/admin/action/lower_memberships','01 · SASTAVI LIGA')}
{form('/admin/action/lower_tm','02 · TM MAPPING')}
{form('/admin/action/lower_data','03 · IGRAČI + STAFF')}
{form('/admin/action/lower_transfers','04 · TRANSFERI')}
{form('/admin/action/lower_crests','05 · GRBOVI')}
{form('/admin/action/lower_info','06 · INFORMACIJE O KLUBOVIMA')}
{form('/admin/action/lower_standings','07 · TABELE')}
{form('/admin/action/lower_discover','POSEBNO · OTKRIJ / USKLADI SVE LIGE NIVO 3–6')}
</div></section>

<section class="adminpanel"><small>OMLADINSKA TAKMIČENJA</small><h2>Omladinska liga Srbije + Kadetska liga Srbije</h2><p>{section_summary(youth)}</p><div class="adminactions">
{form('/admin/action/youth_memberships','01 · SASTAVI LIGA')}
{form('/admin/action/youth_tm','02 · TM MAPPING')}
{form('/admin/action/youth_data','03 · IGRAČI + STAFF')}
{form('/admin/action/youth_transfers','04 · TRANSFERI')}
{form('/admin/action/youth_crests','05 · GRBOVI')}
{form('/admin/action/youth_info','06 · INFORMACIJE O KLUBOVIMA')}
{form('/admin/action/youth_standings','07 · TABELE')}
</div></section>

<section class="adminpanel womenadmin"><small>ŽENSKI FUDBAL</small><h2>Ženske lige i selekcije</h2><p>{section_summary(women)}</p><div class="adminactions">
{form('/admin/action/women_memberships','01 · SASTAVI LIGA')}
{form('/admin/action/women_tm','02 · TM MAPPING')}
{form('/admin/action/women_data','03 · IGRAČICE + STAFF')}
{form('/admin/action/women_transfers','04 · TRANSFERI')}
{form('/admin/action/women_crests','05 · GRBOVI')}
{form('/admin/action/women_info','06 · INFORMACIJE O KLUBOVIMA')}
{form('/admin/action/women_standings','07 · TABELE')}
{form('/admin/action/women_fss','POSEBNO · FSS ROSTERI SUPER + PRVA')}
</div></section>

<section class="adminpanel"><small>NAPREDNE KONTROLE</small><h2>Po nivou / po ligi / mapping</h2>
<form data-busy="1" method="post" action="/admin/action/level"><input type="hidden" name="country_id" value="{cid}"><select name="level">{level_opts}</select><select name="operation"><option value="membership">Sastavi liga</option><option value="tm">TM mapping</option><option value="data">Igrači + staff</option><option value="transfers">Transferi</option><option value="standings">Tabela</option></select><button>Pokreni nivo</button></form>
<form data-busy="1" method="post" action="/admin/action/league"><input type="hidden" name="country_id" value="{cid}"><select name="league_id">{league_opts}</select><select name="operation"><option value="membership">Sastav lige</option><option value="tm">TM mapping</option><option value="data">Igrači + staff</option><option value="transfers">Transferi</option><option value="standings">Tabela</option></select><button>Pokreni ligu</button></form>
<div class="pageactions"><a class="button secondary" href="/mapping_workbench?country_id={cid}&level=3">Mapping Nivo 3</a><a class="button secondary" href="/mapping_workbench?country_id={cid}&level=4">Mapping Nivo 4</a><a class="button secondary" href="/mapping_workbench?country_id={cid}&level=5">Mapping Nivo 5</a><a class="button secondary" href="/mapping_workbench?country_id={cid}&level=6">Mapping Nivo 6</a></div></section>

{status_table}'''
    return layout('SUPERADMIN · Sync Center',body,admin=True)

def _select_leagues(cid, levels=None, section=None):
    con=connect(); sql='SELECT * FROM leagues WHERE country_id=? AND COALESCE(source_active,1)=1'; p=[int(cid)]
    if section is not None:
        sql+=' AND COALESCE(section,\'senior\')=?'; p.append(section)
    if levels:
        marks=','.join('?' for _ in levels); sql+=f' AND level IN ({marks})'; p.extend([int(x) for x in levels])
    sql+=' ORDER BY COALESCE(level,999),name'; rows=con.execute(sql,p).fetchall(); con.close(); return rows


def _run_league_group(cid, operation, levels=None, section=None, fetch_staff_dob=False):
    leagues=_select_leagues(cid,levels,section)
    if any(int(l['level'] or 0) in (3,4,5,6) for l in leagues):
        try:
            wanted=tuple(sorted({int(l['level'] or 0) for l in leagues if int(l['level'] or 0) in (3,4,5,6)}))
            if wanted: engine.discover_srbijasport_pyramid(cid,False,levels=wanted)
        except Exception: pass
        leagues=_select_leagues(cid,levels,section)
    out={'leagues':len(leagues),'ok':0,'rows':0,'found':0,'matched':0,'review':0,'unresolved':0,
         'clubs':0,'players':0,'staff':0,'mapped':0,'with_data':0,'errors':[]}
    for l in leagues:
        try:
            if operation=='membership':
                r=engine.refresh_league_membership_live(l); out['found']+=int(r.get('found',0) or 0); out['ok']+=1
            elif operation=='tm':
                r=engine.resolve_tm_for_league(l['id']); out['matched']+=int(r.get('matched',0) or 0); out['review']+=int(r.get('review',0) or 0); out['unresolved']+=int(r.get('unresolved',0) or 0); out['ok']+=1
            elif operation=='data':
                r=engine.batch_update_league(l['id'],fetch_staff_dob); out['clubs']+=int(r.get('clubs',0) or 0); out['ok']+=int(r.get('ok',0) or 0); out['players']+=int(r.get('players',0) or 0); out['staff']+=int(r.get('staff',0) or 0)
            elif operation=='transfers':
                r=engine.batch_update_league_transfers(l['id']); out['mapped']+=int(r.get('mapped',0) or 0); out['ok']+=int(r.get('ok',0) or 0); out['rows']+=int(r.get('rows',0) or 0); out['with_data']+=1 if int(r.get('rows',0) or 0)>0 else 0
            elif operation=='standings':
                r=engine.refresh_league_standings(l['id']); out['rows']+=int(r.get('rows',0) or 0); out['ok']+=1
        except Exception as e:
            out['errors'].append(f"{l['name']}: {e}")
    return out

def run_action(name, form):
    cid=int(form.get('country_id','1') or 1)
    try:
        if name=='night_sync':
            a=engine.refresh_country_memberships_all(cid,False)
            b=engine.resolve_tm_for_country_all(cid)
            c=engine.batch_update_country_all_leagues(cid,fetch_staff_dob=form.get('fetch_dob')=='1',refresh_membership=False)
            d=engine.batch_update_country_transfers_all(cid)
            e=engine.enrich_all_clubs(cid,'crests')
            f=engine.enrich_all_clubs(cid,'info')
            g=engine.refresh_country_standings(cid)
            nr=engine.update_all_national_teams(cid,fetch_staff_dob=form.get('fetch_dob')=='1')
            ns=engine.refresh_country_national_standings(cid)
            errs=len(a.get('errors',[]))+len(c.get('errors',[]))+len(e.get('errors',[]))+len(f.get('errors',[]))+len(g.get('errors',[]))+len(nr.get('errors',[]))+len(ns.get('errors',[]))
            return f"NOĆNA SINHRONIZACIJA: 01 sastavi {a.get('ok',0)}/{a.get('leagues',0)}; 02 TM +{b.get('matched',0)}; 03 podaci {c.get('clubs_ok',0)}/{c.get('clubs',0)}; 04 transferi {d.get('ok',0)}/{d.get('mapped',0)}; 05 grbovi {e.get('crests',0)}; 06 info {f.get('info',0)}; 07 tabele {g.get('ok',0)}/{g.get('leagues',0)}; reprezentacije {nr.get('ok',0)}/{nr.get('teams',0)}, tabele {ns.get('ok',0)}/{ns.get('teams',0)}; greške {errs}."
        if name=='history_excel_import':
            r=history_excel.import_master(force=True)
            return (f"Excel history master: L1 {r['l1']} + pre-1945 {r['pre']}; kup {r['cup']}; "
                    f"L2 {r['l2']}; Srpske lige {r['l3']}; povezano sa master klubovima {r['linked']}, "
                    f"istorijski nazivi bez master veze {r['unlinked']}; takmičenja {r['competitions']}.")
        if name=='history_excel_validate':
            r=history_excel.validate_master()
            detail=', '.join(f"{k} {v['rows']}/{v['expected']}" for k,v in r['sheets'].items())
            return ('Excel history master OK: ' if r['ok'] else 'Excel history master PROVERA NIJE PROŠLA: ')+detail
        if name=='all_crests':
            r=engine.enrich_all_clubs(cid,'crests'); return f"GLOBAL grbovi: {r['ok']}/{r['eligible']} obrađeno; grbovi {r['crests']}; greške {len(r['errors'])}."
        if name=='all_info':
            r=engine.enrich_all_clubs(cid,'info')
            return f"GLOBAL info klubova: {r['ok']}/{r['eligible']} obrađeno; info {r['info']}; greške {len(r['errors'])}."
        if name.startswith('top_'):
            op=name.split('_',1)[1]
            if op in ('crests','info'):
                if op=='crests':
                    r=engine.enrich_clubs_group(cid,'senior',(1,2),'crests'); return f"Top tier {op}: {r['ok']}/{r['eligible']} obrađeno; grbovi {r['crests']}; info {r['info']}; greške {len(r['errors'])}."
                r=engine.enrich_clubs_group(cid,'senior',(1,2),'info')
                return f"Top tier info: {r['ok']}/{r['eligible']} obrađeno; info {r['info']}; greške {len(r['errors'])}."
            r=_run_league_group(cid,op,levels=(1,2),section='senior')
            if op=='membership': return f"Top tier sastavi: {r['ok']}/{r['leagues']} liga; {r['found']} članstava; greške {len(r['errors'])}."
            if op=='tm': return f"Top tier TM: {r['ok']}/{r['leagues']} liga; +{r['matched']} mapped; review {r['review']}; unresolved {r['unresolved']}."
            if op=='data': return f"Top tier podaci: {r['ok']} klubova OK; igrači {r['players']}; staff {r['staff']}; greške {len(r['errors'])}."
            if op=='transfers': return f"Top tier transferi: {r['ok']}/{r['mapped']} mapiranih klubova; redova {r['rows']}."
            if op=='standings': return f"Top tier tabele: {r['ok']}/{r['leagues']} liga; {r['rows']} redova; greške {len(r['errors'])}."
        if name.startswith('lower_'):
            op=name.split('_',1)[1]
            if op in ('crests','info'):
                if op=='crests':
                    r=engine.enrich_clubs_group(cid,'senior',(3,4,5,6),'crests'); return f"Niže lige {op}: {r['ok']}/{r['eligible']} obrađeno; grbovi {r['crests']}; info {r['info']}; greške {len(r['errors'])}."
                r=engine.enrich_clubs_group(cid,'senior',(3,4,5,6),'info')
                return f"Niže lige info: {r['ok']}/{r['eligible']} obrađeno; info {r['info']}; greške {len(r['errors'])}."
            r=_run_league_group(cid,op,levels=(3,4,5,6),section='senior')
            if op=='membership': return f"Niže lige sastavi: {r['ok']}/{r['leagues']} liga; {r['found']} članstava; greške {len(r['errors'])}."
            if op=='tm': return f"Niže lige TM: {r['ok']}/{r['leagues']} liga; +{r['matched']} mapped; review {r['review']}; unresolved {r['unresolved']}."
            if op=='data': return f"Niže lige podaci: {r['ok']} klubova OK; igrači {r['players']}; staff {r['staff']}; greške {len(r['errors'])}."
            if op=='transfers': return f"Niže lige transferi: {r['ok']}/{r['mapped']} mapiranih klubova; redova {r['rows']}."
            if op=='standings': return f"Niže lige tabele: {r['ok']}/{r['leagues']} liga; {r['rows']} redova; greške {len(r['errors'])}."
        if name=='lower_enrich':
            r=engine.enrich_lower_clubs_srbijasport(cid); return f"Niži klubovi SrbijaSport: {r['ok']}/{r['clubs']} obrađeno; grbovi {r['crests']}; info {r.get('info',0)}; greške {len(r['errors'])}."+((' Detalj: '+' | '.join(r['errors'][:3])) if r.get('errors') else '')
        if name.startswith('youth_'):
            op=name.split('_',1)[1]
            if op in ('crests','info'):
                r=engine.enrich_clubs_group(cid,'youth',None,'crests' if op=='crests' else 'info'); return f"Omladinske {op}: {r['ok']}/{r['eligible']} obrađeno; grbovi {r['crests']}; info {r['info']}; greške {len(r['errors'])}."
            r=_run_league_group(cid,op,section='youth')
            if op=='membership': return f"Omladinske lige sastavi: {r['ok']}/{r['leagues']} liga; {r['found']} članstava."
            if op=='tm': return f"Omladinske lige TM: {r['ok']}/{r['leagues']} liga; +{r['matched']} mapped; review {r['review']}."
            if op=='data': return f"Omladinske lige podaci: {r['ok']} klubova OK; igrači {r['players']}; staff {r['staff']}."
            if op=='transfers': return f"Omladinske lige transferi: {r['ok']}/{r['mapped']} mapiranih klubova; redova {r['rows']}."
            if op=='standings': return f"Omladinske tabele: {r['ok']}/{r['leagues']} liga; {r['rows']} redova; greške {len(r['errors'])}."
        if name=='women_crests':
            r=engine.enrich_clubs_group(cid,'women',None,'crests'); return f"Žene grbovi: {r['ok']}/{r['eligible']} obrađeno; grbovi {r['crests']}; greške {len(r['errors'])}."
        if name=='women_info':
            r=engine.enrich_clubs_group(cid,'women',None,'info'); return f"Žene info: {r['ok']}/{r['eligible']} obrađeno; info {r['info']}; greške {len(r['errors'])}."
        if name=='women_standings':
            r=_run_league_group(cid,'standings',section='women'); return f"Ženske tabele: {r['ok']}/{r['leagues']} liga; {r['rows']} redova; greške {len(r['errors'])}."
        if name=='all_memberships':
            r=engine.refresh_country_memberships_all(cid,False); return f"Sastavi liga: {r['ok']}/{r['leagues']} liga osveženo; {r['found']} članstava."
        if name=='all_tm':
            r=engine.resolve_tm_for_country_all(cid); return f"TM mapping: {r['leagues']} liga; novo mapirano {r['matched']}, review {r['review']}, unresolved {r['unresolved']}, NO TM {r['no_tm_yet']}."
        if name=='all_data':
            r=engine.batch_update_country_all_leagues(cid,fetch_staff_dob=form.get('fetch_dob')=='1',refresh_membership=False); return f"Podaci: {r['clubs_ok']}/{r['clubs']} klubova OK; igrači {r['players']}; staff {r['staff']}."
        if name=='all_transfers':
            r=engine.batch_update_country_transfers_all(cid); return f"Transferi: {r['ok']}/{r['mapped']} mapiranih klubova provereno; {r['with_data']} sa podacima."
        if name=='national':
            r=engine.update_all_national_teams(cid,fetch_staff_dob=form.get('fetch_dob')=='1'); msg=f"Reprezentacije: {r['ok']}/{r['teams']} timova; igrači {r['players']}; staff {r['staff']}."; return msg+((' Greške: '+' | '.join(r['errors'][:3])) if r.get('errors') else '')
        if name=='standings_all':
            r=engine.refresh_country_standings(cid); return f"Tabele: {r['ok']}/{r['leagues']} liga; {r['rows']} redova. Greške {len(r['errors'])}."
        if name=='lower_discover':
            r=engine.discover_srbijasport_pyramid(cid,False,levels=(3,4,5,6)); lv=r.get('levels',{}); return f"Otkrivanje liga: N3 {lv.get(3,0)}, N4 {lv.get(4,0)}, N5 {lv.get(5,0)}, N6 {lv.get(6,0)}; novo {r.get('inserted',0)}, osveženo {r.get('updated',0)}, neaktivno {r.get('deactivated',0)}; greške {len(r.get('errors',[]))}."
        if name=='level6_discover':
            r=engine.discover_level6_srbijasport(cid,True); return f"Nivo 6: otkriveno {r['discovered']}/{r['expected']} liga; sync {r['synced']}; klubova {r['clubs']}; nedostaje {len(r['missing'])}; greške {len(r['errors'])}."
        if name=='level6_standings':
            r=engine.refresh_level_standings(cid,6); return f"Nivo 6 tabele: {r['ok']}/{r['leagues']} liga; {r['rows']} redova; greške {len(r['errors'])}."
        if name=='full':
            a=engine.refresh_country_memberships_all(cid,False); b=engine.resolve_tm_for_country_all(cid); c=engine.batch_update_country_all_leagues(cid,False,False); d=engine.batch_update_country_transfers_all(cid); e=engine.refresh_country_standings(cid); nr=engine.update_all_national_teams(cid,False); n=engine.refresh_country_national_standings(cid); return f"FULL SYNC: membership {a['ok']}/{a['leagues']}; TM +{b['matched']}; data {c['clubs_ok']}/{c['clubs']}; transfers {d['ok']}/{d['mapped']}; tabele {e['ok']}/{e['leagues']}; reprezentacije rosteri {nr['ok']}/{nr['teams']}; tabele reprezentacija {n['ok']}/{n['teams']}."
        if name=='women_memberships':
            r=engine.refresh_section_memberships(cid,'women'); return f"Ženske lige membership: {r['ok']}/{r['leagues']} liga; {r['found']} klubova."
        if name=='women_fss':
            r=engine.update_women_fss_rosters_country(cid); return f"FSS rosteri žena: {r['ok']}/{r['leagues']} liga; igrači {r['players']}."
        if name=='women_tm':
            r=engine.resolve_tm_for_section(cid,'women'); return f"Žene TM: checked {r['checked']}; mapped {r['matched']}; review {r['review']}; unresolved {r['unresolved']}."
        if name=='women_data':
            r=engine.batch_update_section(cid,'women',False); return f"Žene podaci: {r['ok']}/{r['mapped']} mapiranih klubova; igrači {r['players']}; staff {r['staff']}."
        if name=='women_transfers':
            r=engine.batch_update_section_transfers(cid,'women'); return f"Žene transferi: {r['ok']}/{r['mapped']} klubova; {r['with_data']} sa podacima."
        if name=='women_national':
            r=engine.update_women_national_teams(cid); return f"Ženske reprezentacije: {r['ok']}/{r['teams']} timova; igrači {r['players']}; staff {r['staff']}."
        if name=='women_enrich':
            r=engine.enrich_women_clubs_srbijasport(cid); return f"Ženski klubovi SrbijaSport: {r['ok']}/{r['clubs']} obrađeno; grbovi {r['crests']}; info {r.get('info',0)}; stadioni {r.get('stadiums',0)}; greške {len(r['errors'])}."+((' Detalj: '+' | '.join(r['errors'][:3])) if r.get('errors') else '')
        if name=='national_standings':
            r=engine.refresh_country_national_standings(cid); msg=f"Reprezentacije tabele: {r['ok']}/{r['teams']} selekcija; izračunato {r['rows']} redova; vidljivo {r.get('visible_rows',0)}; greške {len(r['errors'])}."; return msg+((' Detalj: '+' | '.join(r['errors'][:3])) if r.get('errors') else '')
        if name=='national_standings_source':
            tid=int(form.get('team_id','0')); url=engine.set_national_standings_source(tid,form.get('standings_url','')); return 'Izvor tabele sačuvan.' if url else 'Izvor tabele obrisan.'
        if name=='level':
            lv=int(form.get('level','0')); op=form.get('operation','membership')
            if op=='membership': r=engine.refresh_level_memberships(cid,lv); return f"Nivo {lv}: {r['active']}/{r['expected']} aktivnih klubova."
            if op=='tm': r=engine.resolve_tm_for_level(cid,lv); return f"Nivo {lv} TM: +{r['matched']} mapped; review {r['review']}; unresolved {r['unresolved']}."
            if op=='data': r=engine.batch_update_level(cid,lv,False); return f"Nivo {lv} podaci: {r['ok']}/{r['mapped']} OK; igrači {r['players']}; staff {r['staff']}."
            if op=='standings': r=engine.refresh_level_standings(cid,lv); return f"Nivo {lv} tabele: {r['ok']}/{r['leagues']} liga; {r['rows']} redova."
            r=engine.batch_update_level_transfers(cid,lv); return f"Nivo {lv} transferi: {r['ok']}/{r['mapped']} provereno."
        if name=='league':
            lid=int(form.get('league_id','0')); op=form.get('operation','membership'); con=connect(); league=con.execute('SELECT * FROM leagues WHERE id=?',(lid,)).fetchone(); con.close()
            if not league: return 'Liga nije pronađena.'
            if op=='membership': r=engine.refresh_league_membership_live(league); return f"{league['name']}: pronađeno {r['found']} klubova."
            if op=='tm': r=engine.resolve_tm_for_league(lid); return f"{league['name']} TM: mapped {r['matched']}; review {r['review']}; unresolved {r['unresolved']}."
            if op=='data': r=engine.batch_update_league(lid,False); return f"{league['name']}: {r['ok']}/{r['clubs']} klubova OK; igrači {r['players']}; staff {r['staff']}."
            if op=='standings': r=engine.refresh_league_standings(lid); return f"{league['name']} tabela: {r['rows']} redova sa izvora {r['source']}."
            r=engine.batch_update_league_transfers(lid); return f"{league['name']} transferi: {r['ok']}/{r['mapped']} provereno."
        return 'Nepoznata akcija.'
    except Exception as e:
        return 'GREŠKA: '+str(e)
