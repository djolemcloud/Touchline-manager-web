from __future__ import annotations
import mimetypes, re, threading, webbrowser, time
from email.parser import BytesParser
from email.policy import default as email_policy_default
from http.cookies import SimpleCookie
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse, parse_qs, quote
from .config import HOST, PORT, STATIC_DIR, VERSION
from . import views, admin
from .data import connect
from . import sync_engine as engine


_LIFECYCLE_LOCK=threading.Lock()
_LIFECYCLE_TABS=set()
_LIFECYCLE_SERVER=None
_LIFECYCLE_TIMER=None

def _cancel_shutdown_timer():
    global _LIFECYCLE_TIMER
    with _LIFECYCLE_LOCK:
        if _LIFECYCLE_TIMER is not None:
            try: _LIFECYCLE_TIMER.cancel()
            except Exception: pass
            _LIFECYCLE_TIMER=None

def _register_tab(tab_id):
    if not tab_id: return
    global _LIFECYCLE_TIMER
    with _LIFECYCLE_LOCK:
        _LIFECYCLE_TABS.add(tab_id)
        if _LIFECYCLE_TIMER is not None:
            try: _LIFECYCLE_TIMER.cancel()
            except Exception: pass
            _LIFECYCLE_TIMER=None

def _schedule_shutdown_if_last(tab_id):
    global _LIFECYCLE_TIMER
    if not tab_id: return
    with _LIFECYCLE_LOCK:
        _LIFECYCLE_TABS.discard(tab_id)
        if _LIFECYCLE_TABS or _LIFECYCLE_SERVER is None:
            return
        if _LIFECYCLE_TIMER is not None:
            try: _LIFECYCLE_TIMER.cancel()
            except Exception: pass
        def stop_if_still_empty():
            global _LIFECYCLE_TIMER
            with _LIFECYCLE_LOCK:
                if _LIFECYCLE_TABS or _LIFECYCLE_SERVER is None:
                    _LIFECYCLE_TIMER=None; return
                srv=_LIFECYCLE_SERVER; _LIFECYCLE_TIMER=None
            print('[LIFECYCLE] Poslednji Touchline prozor je zatvoren. Gasim lokalni server.', flush=True)
            threading.Thread(target=srv.shutdown,daemon=True).start()
        _LIFECYCLE_TIMER=threading.Timer(5.0,stop_if_still_empty)
        _LIFECYCLE_TIMER.daemon=True; _LIFECYCLE_TIMER.start()

_LIFECYCLE_SCRIPT=r'''<script>(function(){
try{
  if(!window.name || !window.name.startsWith('touchline-tab-')){window.name='touchline-tab-'+Date.now()+'-'+Math.random().toString(36).slice(2);}
  const id=window.name; let navigating=false;
  const ping=()=>fetch('/__touchline_tab_open?id='+encodeURIComponent(id),{cache:'no-store',keepalive:true}).catch(()=>{});
  ping();
  document.addEventListener('click',e=>{const a=e.target.closest&&e.target.closest('a'); if(a&&a.href&&new URL(a.href,location.href).origin===location.origin){navigating=true;}},true);
  document.addEventListener('submit',()=>{navigating=true;},true);
  window.addEventListener('pageshow',()=>{navigating=false;ping();});
  window.addEventListener('pagehide',()=>{if(!navigating){try{navigator.sendBeacon('/__touchline_tab_close?id='+encodeURIComponent(id),'');}catch(e){}}});
}catch(e){}
})();</script>'''


def q1(q,name,default=''):
    return (q.get(name) or [default])[0]


def safe_name(name, fallback='export'):
    x=re.sub(r'[^A-Za-z0-9_.-]+','_',str(name or fallback)).strip('_')
    return x or fallback


class Handler(BaseHTTPRequestHandler):
    server_version=f'TouchlinePortal/{VERSION}'

    def is_local(self):
        return self.client_address[0] in ('127.0.0.1','::1')

    def lang(self):
        c=SimpleCookie(); c.load(self.headers.get('Cookie',''))
        return c.get('tl_lang').value if c.get('tl_lang') and c.get('tl_lang').value in ('sr','en') else 'sr'

    def send_html(self,html,status=200,headers=None):
        if '</body>' in html:
            html=html.replace('</body>',_LIFECYCLE_SCRIPT+'</body>',1)
        else:
            html=html+_LIFECYCLE_SCRIPT
        data=html.encode('utf-8'); self.send_response(status); self.send_header('Content-Type','text/html; charset=utf-8'); self.send_header('Content-Length',str(len(data)))
        for k,v in (headers or {}).items(): self.send_header(k,v)
        self.end_headers(); self.wfile.write(data)

    def redirect(self,url,headers=None):
        self.send_response(303); self.send_header('Location',url)
        for k,v in (headers or {}).items(): self.send_header(k,v)
        self.end_headers()

    def send_file(self,path):
        p=Path(path)
        if not p.is_file(): return self.send_html('<h1>404</h1>',404)
        data=p.read_bytes(); self.send_response(200); self.send_header('Content-Type',mimetypes.guess_type(p.name)[0] or 'application/octet-stream'); self.send_header('Content-Length',str(len(data))); self.end_headers(); self.wfile.write(data)

    def send_download(self,data,filename,ctype='application/octet-stream'):
        self.send_response(200); self.send_header('Content-Type',ctype); self.send_header('Content-Disposition',f'attachment; filename="{filename}"'); self.send_header('Content-Length',str(len(data))); self.end_headers(); self.wfile.write(data)

    def post_data(self):
        n=int(self.headers.get('Content-Length','0') or 0); raw=self.rfile.read(n); ctype=self.headers.get('Content-Type',''); fields={}; files={}
        if ctype.lower().startswith('multipart/form-data'):
            msg=BytesParser(policy=email_policy_default).parsebytes((f'Content-Type: {ctype}\r\nMIME-Version: 1.0\r\n\r\n').encode()+raw)
            if msg.is_multipart():
                for part in msg.iter_parts():
                    name=part.get_param('name',header='content-disposition')
                    if not name: continue
                    payload=part.get_payload(decode=True) or b''; filename=part.get_filename()
                    if filename: files[name]={'filename':filename,'data':payload}
                    else: fields[name]=payload.decode(part.get_content_charset() or 'utf-8',errors='replace')
        else:
            fields={k:v[-1] for k,v in parse_qs(raw.decode('utf-8',errors='replace')).items()}
        return fields,files

    def do_GET(self):
        engine.set_ui_lang(self.lang()); u=urlparse(self.path); path=u.path; q=parse_qs(u.query); msg=q1(q,'msg','')
        try:
            if path=='/__touchline_tab_open':
                _register_tab(q1(q,'id','')); self.send_response(204); self.end_headers(); return
            if path.startswith('/static/'):
                rel=path[len('/static/'):]; target=(STATIC_DIR/rel).resolve()
                if STATIC_DIR.resolve() not in target.parents and target!=STATIC_DIR.resolve(): return self.send_html('<h1>403</h1>',403)
                return self.send_file(target)
            if path=='/set_lang':
                lang='en' if q1(q,'lang')=='en' else 'sr'; ref=self.headers.get('Referer','/')
                try:
                    pu=urlparse(ref); ref=(pu.path or '/')+(('?'+pu.query) if pu.query else '')
                except Exception: ref='/'
                return self.redirect(ref,{'Set-Cookie':f'tl_lang={lang}; Path=/; SameSite=Lax; Max-Age=31536000'})
            if path=='/': return self.send_html(views.home(msg))
            if path=='/country': return self.redirect('/'+(('?msg='+quote(msg)) if msg else ''))
            if path=='/fss': return self.send_html(views.fss_hub(int(q1(q,'country_id','1')),msg))
            if path in ('/fss/national','/national'): return self.send_html(views.national_hub(int(q1(q,'country_id','1')),msg))
            if path=='/fss/development': return self.send_html(views.development_hub(int(q1(q,'country_id','1')),msg))
            if path=='/youth': return self.send_html(views.youth(int(q1(q,'country_id','1')),msg))
            if path=='/women': return self.send_html(views.women(int(q1(q,'country_id','1')),msg))
            if path=='/women/second': return self.send_html(views.women_second(int(q1(q,'country_id','1')),msg))
            if path=='/fss/regions': return self.send_html(views.regions(int(q1(q,'country_id','1')),msg))
            if path=='/fss/region': return self.send_html(views.region(int(q1(q,'country_id','1')),q1(q,'code','FSB'),msg))
            if path=='/fss/region/fsb/first': return self.send_html(views.fsb_first(int(q1(q,'country_id','1')),msg))
            if path=='/level': return self.send_html(views.level(int(q1(q,'country_id','1')),int(q1(q,'level','3')),msg))
            if path=='/level6': return self.send_html(views.level6(int(q1(q,'country_id','1')),msg))
            if path=='/level6/region': return self.send_html(views.level6_region(int(q1(q,'country_id','1')),q1(q,'code','FSB'),msg))
            if path=='/league': return self.send_html(views.league(int(q1(q,'id','0')),msg))
            if path=='/club': return self.send_html(views.club(int(q1(q,'id','0')),msg))
            if path=='/club/history': return self.send_html(views.club_history(int(q1(q,'id','0')),msg))
            if path=='/league/history': return self.send_html(views.league_history(int(q1(q,'id','0')),msg))
            if path=='/transfers': return self.send_html(views.transfers(int(q1(q,'id','0')),msg))
            if path=='/national_team': return self.send_html(views.national_team(int(q1(q,'id','0')),msg))
            if path=='/admin/sync':
                if not self.is_local(): return self.send_html('<h1>403</h1><p>SUPERADMIN local access only.</p>',403)
                t0=time.perf_counter()
                details=q1(q,'details','0')=='1'
                html=admin.page(int(q1(q,'country_id','1')),msg,details=details)
                ms=(time.perf_counter()-t0)*1000.0
                print(f'[PERF] /admin/sync render {ms:.1f} ms' + (' · details' if details else ''), flush=True)
                return self.send_html(html,headers={'Server-Timing':f'app;dur={ms:.1f}','X-Touchline-Render-Ms':f'{ms:.1f}'})
            if path=='/mapping_workbench':
                if not self.is_local(): return self.send_html('<h1>403</h1>',403)
                return self.send_html(views.mapping_workbench(int(q1(q,'country_id','1')),int(q1(q,'level','4')),msg))

            # Exports
            if path=='/export_country_master':
                cid=int(q1(q,'id','1')); return self.send_download(engine.country_master_workbook_bytes(cid),'SERBIA_MASTER.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            if path=='/export_country_crests':
                cid=int(q1(q,'id','1')); return self.send_download(engine.country_crest_zip_bytes(cid),'SERBIA_CRESTS.zip','application/zip')
            if path=='/export_league_master':
                lid=int(q1(q,'id','0')); return self.send_download(engine.master_workbook_bytes(lid),'LEAGUE_MASTER.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            if path=='/export_league_crests':
                lid=int(q1(q,'id','0')); return self.send_download(engine.crest_zip_bytes(lid),'LEAGUE_CRESTS.zip','application/zip')
            if path=='/export_league_players':
                lid=int(q1(q,'id','0')); con=connect(); league=con.execute('SELECT * FROM leagues WHERE id=?',(lid,)).fetchone(); rows=con.execute('''SELECT p.*,COALESCE(c.cm_name,c.name) club_name FROM players p JOIN clubs c ON c.id=p.club_id WHERE c.league_id=? ORDER BY club_name,p.position,p.name''',(lid,)).fetchall(); con.close(); return self.send_download(engine.players_workbook_bytes(rows),safe_name(league['name'] if league else 'league')+'_players.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            if path=='/export_league_staff':
                lid=int(q1(q,'id','0')); con=connect(); league=con.execute('SELECT * FROM leagues WHERE id=?',(lid,)).fetchone(); rows=con.execute('''SELECT s.*,COALESCE(c.cm_name,c.name) club_name FROM staff s JOIN clubs c ON c.id=s.club_id WHERE c.league_id=? ORDER BY club_name,s.section,s.position,s.name''',(lid,)).fetchall(); con.close(); return self.send_download(engine.staff_workbook_bytes(rows),safe_name(league['name'] if league else 'league')+'_staff.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            if path=='/export_club_players':
                cid=int(q1(q,'id','0')); con=connect(); club=con.execute('SELECT * FROM clubs WHERE id=?',(cid,)).fetchone(); rows=con.execute('''SELECT p.*,COALESCE(c.cm_name,c.name) club_name FROM players p JOIN clubs c ON c.id=p.club_id WHERE p.club_id=? ORDER BY p.position,p.name''',(cid,)).fetchall(); con.close(); return self.send_download(engine.players_workbook_bytes(rows),safe_name((club['cm_name'] or club['name']) if club else 'club')+'_players.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            if path=='/export_club_staff':
                cid=int(q1(q,'id','0')); con=connect(); club=con.execute('SELECT * FROM clubs WHERE id=?',(cid,)).fetchone(); rows=con.execute('''SELECT s.*,COALESCE(c.cm_name,c.name) club_name FROM staff s JOIN clubs c ON c.id=s.club_id WHERE s.club_id=? ORDER BY s.section,s.position,s.name''',(cid,)).fetchall(); con.close(); return self.send_download(engine.staff_workbook_bytes(rows),safe_name((club['cm_name'] or club['name']) if club else 'club')+'_staff.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            if path=='/export_club_transfers':
                cid=int(q1(q,'id','0')); sid=engine.current_season_id(); con=connect(); club=con.execute('SELECT * FROM clubs WHERE id=?',(cid,)).fetchone(); summary=con.execute('SELECT * FROM club_transfer_summary WHERE club_id=? AND season_id=?',(cid,sid)).fetchone(); rows=con.execute('SELECT * FROM club_transfers WHERE club_id=? AND season_id=? ORDER BY direction,player_name',(cid,sid)).fetchall(); con.close(); return self.send_download(engine.transfers_workbook_bytes(club,summary,rows),safe_name((club['cm_name'] or club['name']) if club else 'club')+'_transfers.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            if path=='/export_national_players':
                tid=int(q1(q,'id','0')); con=connect(); team=con.execute('SELECT * FROM national_teams WHERE id=?',(tid,)).fetchone(); rows=con.execute('''SELECT p.*,nt.name team_name FROM national_team_players p JOIN national_teams nt ON nt.id=p.national_team_id WHERE p.national_team_id=? ORDER BY p.position,p.name''',(tid,)).fetchall(); con.close(); return self.send_download(engine.national_players_workbook_bytes(rows),safe_name(team['name'] if team else 'team')+'_players.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            if path=='/export_national_staff':
                tid=int(q1(q,'id','0')); con=connect(); team=con.execute('SELECT * FROM national_teams WHERE id=?',(tid,)).fetchone(); rows=con.execute('''SELECT s.*,nt.name team_name FROM national_team_staff s JOIN national_teams nt ON nt.id=s.national_team_id WHERE s.national_team_id=? ORDER BY s.section,s.position,s.name''',(tid,)).fetchall(); con.close(); return self.send_download(engine.national_staff_workbook_bytes(rows),safe_name(team['name'] if team else 'team')+'_staff.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            if path=='/export_level_mapping_issues':
                cid=int(q1(q,'country_id','1')); lv=int(q1(q,'level','4')); return self.send_download(engine.mapping_workbook_bytes(cid,lv,True),f'LEVEL_{lv}_OPEN_MAPPING.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            if path=='/export_level_mapping_backup':
                cid=int(q1(q,'country_id','1')); lv=int(q1(q,'level','4')); return self.send_download(engine.level_mapping_backup_zip_bytes(cid,lv),f'LEVEL_{lv}_MAPPING_BACKUP.zip','application/zip')
            return self.send_html('<h1>404</h1><p>Stranica nije pronađena.</p>',404)
        except Exception as e:
            return self.send_html(f'<h1>Greška</h1><pre>{esc_html(str(e))}</pre>',500)

    def do_POST(self):
        engine.set_ui_lang(self.lang()); u=urlparse(self.path); path=u.path
        if path=='/__touchline_tab_close':
            _schedule_shutdown_if_last(q1(parse_qs(u.query),'id','')); self.send_response(204); self.end_headers(); return
        form,files=self.post_data(); cid=int(form.get('country_id','1') or 1)
        try:
            if path.startswith('/admin/action/'):
                if not self.is_local(): return self.send_html('<h1>403</h1>',403)
                name=path.rsplit('/',1)[-1]; msg=admin.run_action(name,form); return self.redirect('/admin/sync?country_id='+str(cid)+'&msg='+quote(msg))
            if path.startswith('/mapping/'):
                if not self.is_local(): return self.send_html('<h1>403</h1>',403)
                lv=int(form.get('level','4') or 4); club_id=int(form.get('club_id','0') or 0)
                if path=='/mapping/retry_level':
                    r=engine.retry_clean_level(cid,lv); msg=f"Retry: checked {r['checked']}; mapped {r['matched']}; review {r['review']}; unresolved {r['unresolved']}."
                elif path=='/mapping/recheck_no_tm':
                    r=engine.recheck_no_tm_level(cid,lv); msg=f"NO TM recheck: checked {r['checked']}; mapped {r['matched']}; review {r['review']}; still NO TM {r['no_tm_yet']}."
                elif path=='/mapping/retry':
                    r=engine.retry_tm_clean_club(club_id); msg=r.get('message','Done')
                elif path=='/mapping/accept':
                    msg='Candidate accepted: '+engine.accept_tm_candidate(club_id)
                elif path=='/mapping/no_tm':
                    msg=engine.mark_no_tm_yet(club_id)
                elif path=='/mapping/manual':
                    msg='Mapping saved: '+engine.set_manual_tm_url(club_id,form.get('tm_url',''))
                elif path=='/mapping/import':
                    up=files.get('mapping_file')
                    if not up: raise RuntimeError('Choose XLSX file first')
                    r=engine.import_mapping_xlsx_bytes(up['data'],cid,lv); msg=f"Imported: {r['mapped']} mappings; skipped {r['skipped']}; not found {r['not_found']}."
                else: raise RuntimeError('Unknown mapping action')
                return self.redirect(f'/mapping_workbench?country_id={cid}&level={lv}&msg='+quote(msg))
            return self.redirect('/')
        except Exception as e:
            target='/admin/sync?country_id='+str(cid) if path.startswith('/admin/') else f'/mapping_workbench?country_id={cid}&level={form.get("level","4")}'
            return self.redirect(target+'&msg='+quote('GREŠKA: '+str(e)))

    def log_message(self,fmt,*args):
        return


def esc_html(v):
    import html
    return html.escape(str(v or ''))


def run(open_browser=True):
    global _LIFECYCLE_SERVER
    engine.init_db(); url=f'http://{HOST}:{PORT}'
    print(f'Touchline Web Portal {VERSION} running at',url)
    print('SUPERADMIN:',url+'/admin/sync?country_id=1')
    print('Zatvaranje poslednjeg Touchline browser prozora automatski gasi server.')
    print('Press Ctrl+C to stop.')
    srv=ThreadingHTTPServer((HOST,PORT),Handler); _LIFECYCLE_SERVER=srv
    if open_browser: threading.Timer(1.0,lambda:webbrowser.open(url)).start()
    try: srv.serve_forever()
    finally:
        _cancel_shutdown_timer(); srv.server_close(); _LIFECYCLE_SERVER=None
