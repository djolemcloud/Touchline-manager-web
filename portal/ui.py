from __future__ import annotations
from html import escape
from urllib.parse import quote
from .config import VERSION
from . import sync_engine as engine


def esc(v):
    return escape(str(v or ''))


def tr(v):
    return engine.tr_text(v, engine.ui_lang())


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
    dont_close=engine.tr_text('Ne zatvaraj ovaj prozor dok se operacija ne završi.',lang)
    admin_link = '<a class="top-admin" href="/admin/sync?country_id=1">SUPERADMIN</a>' if not admin else '<a class="top-admin active" href="/admin/sync?country_id=1">SUPERADMIN</a>'
    return f'''<!doctype html>
<html lang="{lang}"><head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<meta name="theme-color" content="#071019">
<title>{esc(title)}</title><link rel="stylesheet" href="/static/style.css?v=130">
</head><body>
<header class="app-header"><div class="wrap head"><a class="brand" href="/"><span class="brandmark">TL</span><span class="brandtext"><b>Touchline</b><small>{esc(tagline)}</small></span></a><nav class="topnav"><a href="/">{esc(home_label)}</a><a href="/#competitions">{esc(comp_label)}</a><a href="/fss?country_id=1">FSS</a></nav><div class="head-actions">{admin_link}<nav class="langswitch" aria-label="Language"><a class="{sr}" href="/set_lang?lang=sr">SR</a><span>/</span><a class="{en}" href="/set_lang?lang=en">EN</a></nav></div></div></header>
<main class="wrap app-main"><button id="globalBack" class="global-back" type="button" aria-label="{esc(back_label)}">← {esc(back_label)}</button>{body}</main>
<footer class="wrap app-footer"><span>Touchline Web Portal {esc(VERSION)}</span><span>{esc(country_label)}</span></footer>
<div id="adminBusy" class="adminbusy"><div class="adminbusybox"><div class="adminspinner"></div><h2>SUPERADMIN</h2><p id="adminBusyAction">{esc(busy)}</p><small>{esc(dont_close)}</small></div></div>
<script>document.addEventListener('submit',function(ev){{var f=ev.target;if(f&&f.dataset&&f.dataset.busy==='1'){{var o=document.getElementById('adminBusy');if(o)o.classList.add('show');var b=f.querySelector('button');if(b){{document.getElementById('adminBusyAction').textContent=b.textContent;b.disabled=true;}}}}}});(function(){{var b=document.getElementById('globalBack');if(!b)return;if(location.pathname==='/'||location.pathname==='/index'){{b.style.display='none';return;}}b.addEventListener('click',function(){{if(history.length>1)history.back();else location.href='/';}});}})();</script>
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
