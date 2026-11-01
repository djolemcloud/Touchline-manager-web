from __future__ import annotations
from html import escape
from urllib.parse import quote
from .config import VERSION
from . import sync_engine as engine


def esc(v):
    # 0 is a real value (e.g. "0 titula") and must be shown, only None/False/'' are empty
    if v is None or v is False: return ''
    return escape(str(v))


def tr(v):
    return engine.tr_text(v, engine.ui_lang())


# Header search: live suggestions from /api/search, Enter opens the full /search page.
SEARCH_JS = r"""<script>(function(){
var f=document.getElementById('headSearch'); if(!f) return;
var i=f.querySelector('input'), pop=f.querySelector('.searchpop'), t=null, last='', sel=-1;
function esc(s){return String(s==null?'':s).replace(/[&<>"]/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c];});}
var SHORT={gk:'GK',df:'DF',mf:'MF',fw:'FW',other:'—'};
function close(){pop.classList.remove('open');pop.innerHTML='';sel=-1;}
function render(d,q){
  var h='';
  if(d.clubs.length){h+='<div class="sp-head">'+esc(f.dataset.clubs)+'</div>';d.clubs.forEach(function(c){
    h+='<a class="sp-item" href="'+esc(c.href)+'"><span class="sp-crest">'+(c.crest?'<img src="'+esc(c.crest)+'" alt="">':'')+'</span><span class="sp-text"><b>'+esc(c.name)+'</b><small>'+esc(c.league)+'</small></span></a>';});}
  if(d.players.length){h+='<div class="sp-head">'+esc(f.dataset.players)+'</div>';d.players.forEach(function(p){
    h+='<a class="sp-item" href="'+esc(p.href)+'"><span class="pos pos-'+esc(p.pos)+'">'+(SHORT[p.pos]||'—')+'</span><span class="sp-text"><b>'+esc(p.name)+'</b><small>'+esc(p.team)+'</small></span></a>';});}
  if(!h) h='<div class="sp-empty">'+esc(f.dataset.empty)+'</div>';
  else h+='<a class="sp-all" href="/search?q='+encodeURIComponent(q)+'">'+esc(f.dataset.all)+' →</a>';
  pop.innerHTML=h; pop.classList.add('open'); sel=-1;
}
i.addEventListener('input',function(){
  var q=i.value.trim(); clearTimeout(t);
  if(q.length<2){close();return;}
  t=setTimeout(function(){ last=q;
    fetch('/api/search?q='+encodeURIComponent(q),{cache:'no-store'}).then(function(r){return r.json();}).then(function(d){ if(q===last) render(d,q); }).catch(function(){});
  },160);
});
i.addEventListener('keydown',function(e){
  var items=pop.querySelectorAll('a'); if(!items.length) return;
  if(e.key==='ArrowDown'||e.key==='ArrowUp'){ e.preventDefault(); sel=(sel+(e.key==='ArrowDown'?1:-1)+items.length)%items.length;
    items.forEach(function(a,k){a.classList.toggle('active',k===sel);}); }
  else if(e.key==='Enter'&&sel>=0){ e.preventDefault(); items[sel].click(); }
  else if(e.key==='Escape'){ close(); i.blur(); }
});
document.addEventListener('click',function(e){ if(!f.contains(e.target)) close(); });
document.addEventListener('keydown',function(e){ if(e.key==='/'&&document.activeElement.tagName!=='INPUT'&&document.activeElement.tagName!=='TEXTAREA'){ e.preventDefault(); i.focus(); } });
})();</script>"""


def layout(title: str, body: str, *, admin: bool=False) -> str:
    lang = engine.ui_lang()
    body = engine.localize_html(body, lang)
    title = engine.tr_text(title, lang)
    sr = ' active' if lang == 'sr' else ''
    en = ' active' if lang == 'en' else ''
    home_label=engine.tr_text('Početna',lang)
    comp_label=engine.tr_text('Takmičenja',lang)
    country_label=engine.tr_text('Srbija',lang)
    back_label=engine.tr_text('Nazad',lang)
    tagline=engine.tr_text('Fudbalski podaci i analitika',lang)
    busy=engine.tr_text('Operacija je pokrenuta...',lang)
    search_ph=engine.tr_text('Pretraži klub ili igrača',lang)
    search_lbl=engine.tr_text('Pretraga',lang)
    s_clubs=engine.tr_text('Klubovi',lang); s_players=engine.tr_text('Igrači',lang)
    s_empty=engine.tr_text('Nema rezultata.',lang); s_all=engine.tr_text('Svi rezultati',lang)
    dont_close=engine.tr_text('Ne zatvaraj ovaj prozor dok se operacija ne završi.',lang)
    admin_link = '<a class="top-admin" href="/admin/sync?country_id=1">SUPERADMIN</a>' if not admin else '<a class="top-admin active" href="/admin/sync?country_id=1">SUPERADMIN</a>'
    return f'''<!doctype html>
<html lang="{lang}"><head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<meta name="theme-color" content="#080d14">
<link rel="preconnect" href="https://fonts.googleapis.com"><link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap">
<title>{esc(title)}</title><link rel="stylesheet" href="/static/style.css?v=200">
</head><body>
<header class="app-header" id="appHeader"><div class="wrap head"><a class="brand" href="/"><span class="brandmark">TL</span><span class="brandtext"><b>Touchline</b><small>{esc(tagline)}</small></span></a><nav class="topnav"><a href="/">{esc(home_label)}</a><a href="/#competitions">{esc(comp_label)}</a><a href="/fss?country_id=1">FSS</a><a class="navsearch" href="/search">{esc(search_lbl)}</a><a class="navadmin" href="/admin/sync?country_id=1">SUPERADMIN</a></nav><div class="head-actions"><form class="headsearch" id="headSearch" action="/search" method="get" role="search" data-clubs="{esc(s_clubs)}" data-players="{esc(s_players)}" data-empty="{esc(s_empty)}" data-all="{esc(s_all)}"><input type="search" name="q" placeholder="{esc(search_ph)}" aria-label="{esc(search_lbl)}" autocomplete="off"><kbd>/</kbd><div class="searchpop"></div></form>{admin_link}<nav class="langswitch" aria-label="Language"><a class="{sr}" href="/set_lang?lang=sr">SR</a><span>/</span><a class="{en}" href="/set_lang?lang=en">EN</a></nav><button class="navtoggle" id="navToggle" type="button" aria-label="Meni" aria-expanded="false"><span></span></button></div></div></header>
<main class="wrap app-main"><button id="globalBack" class="global-back" type="button" aria-label="{esc(back_label)}">← {esc(back_label)}</button>{body}</main>
<footer class="wrap app-footer"><span>Touchline Web Portal {esc(VERSION)}</span><span>{esc(country_label)}</span></footer>
<div id="adminBusy" class="adminbusy"><div class="adminbusybox"><div class="adminspinner"></div><h2>SUPERADMIN</h2><p id="adminBusyAction">{esc(busy)}</p><small>{esc(dont_close)}</small></div></div>
<script>document.addEventListener('submit',function(ev){{var f=ev.target;if(f&&f.dataset&&f.dataset.busy==='1'){{var o=document.getElementById('adminBusy');if(o)o.classList.add('show');var b=f.querySelector('button');if(b){{document.getElementById('adminBusyAction').textContent=b.textContent;b.disabled=true;}}}}}});(function(){{var b=document.getElementById('globalBack');if(!b)return;if(location.pathname==='/'||location.pathname==='/index'){{b.style.display='none';return;}}b.addEventListener('click',function(){{if(history.length>1)history.back();else location.href='/';}});var c=document.querySelector('.crumb');if(c)c.insertBefore(b,c.firstChild);}})();(function(){{var t=document.getElementById('navToggle'),h=document.getElementById('appHeader');if(!t||!h)return;t.addEventListener('click',function(){{var o=h.classList.toggle('nav-open');t.setAttribute('aria-expanded',o?'true':'false');}});}})();</script>{SEARCH_JS}
</body></html>'''


def notice(msg=''):
    return f'<div class="notice">{esc(msg)}</div>' if msg else ''


def crumb(parts):
    out=[]
    for label, href in parts:
        out.append(f'<a href="{esc(href)}">{esc(label)}</a>' if href else esc(label))
    return '<div class="crumb">'+' / '.join(out)+'</div>'


def visual_card(img, title, subtitle, href, stats=''):
    stat = f'<div class="visualstats">{esc(stats)}</div>' if stats else ''
    return f'''<a class="visualcard" href="{esc(href)}">
<img src="{esc(img)}" alt="{esc(title)}"><div class="visualmeta"><small>{esc(subtitle)}</small><h2>{esc(title)}</h2>{stat}</div></a>'''


def stat_cell(value, label):
    return f'<div class="statcell"><strong>{esc(value)}</strong><span>{esc(label)}</span></div>'


def rich_card(*, title, subtitle, href, stats, image=None, cta='Otvori ligu →'):
    image_html = f'<img class="richcard-image" src="{esc(image)}" alt="{esc(title)}">' if image else ''
    cells = ''.join([
        stat_cell(stats.get('clubs',0), 'klubovi'),
        stat_cell(f"{stats.get('mapped',0)}/{stats.get('clubs',0)}", 'TM mapirano'),
        stat_cell(stats.get('players',0), 'igrači'),
        stat_cell(stats.get('staff',0), 'stručni štab'),
        stat_cell(stats.get('transfers',0), 'transferi 26/27'),
        stat_cell(f"{stats.get('updated_clubs',0)}/{stats.get('clubs',0)}", 'klubova ažurirano'),
    ])
    return f'''<article class="richcard">{image_html}<div class="richcard-body"><small>{esc(subtitle)}</small><h2>{esc(title)}</h2><div class="richstats">{cells}</div><a class="richcta" href="{esc(href)}">{esc(cta)}</a></div></article>'''


def group_card(title, subtitle, href, stats, cta='Otvori sekciju →'):
    return rich_card(title=title, subtitle=subtitle, href=href, stats=stats, image=None, cta=cta)


def metric_strip(items):
    return '<div class="metricstrip">'+''.join(f'<div><span>{esc(k)}</span><strong>{esc(v)}</strong></div>' for k,v in items)+'</div>'
