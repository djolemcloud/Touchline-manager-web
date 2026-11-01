from __future__ import annotations
import py_compile
import sqlite3
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))


def check(cond, msg):
    if not cond:
        raise AssertionError(msg)


def main():
    print('TOUCHLINE WEB PORTAL 1.7.1 - CLEAN FREEZE VERIFY', flush=True)

    # Syntax
    py_files = [ROOT/'run.py', ROOT/'self_test.py'] + sorted((ROOT/'portal').glob('*.py'))
    for p in py_files:
        if p.exists():
            py_compile.compile(str(p), doraise=True)
    print('PYTHON SYNTAX OK:', len([p for p in py_files if p.exists()]), 'files', flush=True)

    # Imports
    from portal.config import DB_PATH, STATIC_DIR, VERSION
    from portal import views, admin, sync_engine, history_excel, trophies
    check(VERSION == '1.7.1', f'Expected VERSION 1.7.1, got {VERSION}')
    print('MODULE IMPORTS OK · build', VERSION, flush=True)

    # DB integrity + core tables
    check(DB_PATH.exists(), f'Database missing: {DB_PATH}')
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    integrity = con.execute('PRAGMA integrity_check').fetchone()[0]
    check(integrity == 'ok', f'Database integrity failed: {integrity}')
    present = {r[0] for r in con.execute("SELECT name FROM sqlite_master WHERE type='table'")}
    needed = {'countries','leagues','clubs','players','staff','club_info','stadiums','history_competitions','history_club_seasons'}
    check(needed.issubset(present), 'Missing DB tables: '+', '.join(sorted(needed-present)))
    counts = {t: con.execute(f'SELECT COUNT(*) FROM {t}').fetchone()[0] for t in ['leagues','clubs','players','staff']}
    check(counts['clubs'] >= 100 and counts['leagues'] >= 50, f'Unexpected DB counts: {counts}')

    # Abandoned stadium-photo layer must be clean, metadata may remain.
    scol={r[1] for r in con.execute('PRAGMA table_info(stadiums)')}
    if {'image_local','image_url'}.issubset(scol):
        media_refs=con.execute("SELECT COUNT(*) FROM stadiums WHERE COALESCE(image_local,'')<>'' OR COALESCE(image_url,'')<>''").fetchone()[0]
        check(media_refs == 0, f'Stadium photo DB references remain: {media_refs}')

    # Every current L1/L2 senior club must have at least one jersey asset.
    clubs=con.execute("""SELECT c.id,COALESCE(c.cm_name,c.name) club,l.level
        FROM clubs c JOIN leagues l ON l.id=c.league_id
        WHERE l.level IN (1,2) AND COALESCE(l.section,'senior')='senior'
          AND COALESCE(c.source_active,1)=1 AND COALESCE(l.source_active,1)=1
        ORDER BY l.level,club""").fetchall()
    missing=[]
    for c in clubs:
        ok=False
        for idx in (1,2,3):
            for ext in ('.png','.webp','.jpg','.jpeg'):
                if (STATIC_DIR/'jerseys'/f"{c['id']}_{idx}{ext}").exists():
                    ok=True; break
            if ok: break
        if not ok: missing.append(f"L{c['level']} {c['club']} ({c['id']})")
    check(not missing, 'Missing L1/L2 jersey assets: '+', '.join(missing))
    print(f'JERSEYS OK: {len(clubs)}/{len(clubs)} active L1-L2 clubs', flush=True)
    con.close()

    # History master/trophy mapping still present.
    check((ROOT/'HISTORY'/'TOUCHLINE_HISTORY_MASTER.xlsx').exists(), 'History master missing')
    check((ROOT/'HISTORY'/'TROPHY_MAP.json').exists(), 'Trophy map missing')
    check(trophies.trophy_for('european-cup','europe').get('image',''), 'Trophy mapping not loaded')

    # Active UI must keep jersey layer and no 06B stadium action.
    vsrc=(ROOT/'portal'/'views.py').read_text(encoding='utf-8')
    asrc=(ROOT/'portal'/'admin.py').read_text(encoding='utf-8')
    check('_jersey_assets' in vsrc and 'Dresovi 2026/27' in vsrc, 'Jersey UI layer missing')
    check('final_stadium' not in asrc and '06B' not in asrc, 'Abandoned stadium sync still active in Superadmin')

    # Junk should not exist in clean tree.
    # Importing runtime modules may recreate empty cache/runtime directories.
    # They must contain no shipped artifacts.
    for d,label in [(ROOT/'REPORTS','REPORTS'),(ROOT/'static'/'stadiums','static/stadiums'),(ROOT/'data'/'history_cache','history_cache')]:
        if d.exists():
            files=[x for x in d.rglob('*') if x.is_file()]
            check(not files, f'{label} contains stale shipped files: {[str(x) for x in files[:5]]}')
    check(not (ROOT/'06B').exists(), 'legacy 06B marker remains')

    print('DB OK:', counts, flush=True)
    print('CLEAN TREE OK', flush=True)
    print('ALL OFFLINE TESTS PASSED', flush=True)


if __name__ == '__main__':
    main()
