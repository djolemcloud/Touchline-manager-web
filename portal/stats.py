from __future__ import annotations
from .data import connect
from . import sync_engine as engine


def league_stats(league_ids=None, *, country_id=None, section=None, level=None):
    con = connect()
    where=['1=1']; params=[]
    if league_ids is not None:
        ids=[int(x) for x in league_ids]
        if not ids:
            con.close(); return {}
        where.append('l.id IN (%s)' % ','.join('?'*len(ids))); params.extend(ids)
    if country_id is not None:
        where.append('l.country_id=?'); params.append(int(country_id))
    if section is not None:
        where.append("COALESCE(l.section,'senior')=?"); params.append(section)
    if level is not None:
        where.append('l.level=?'); params.append(int(level))
    sql=f'''SELECT l.id,l.name,l.level,l.section,l.age_group,l.image_path,
      COUNT(DISTINCT CASE WHEN COALESCE(c.source_active,1)=1 THEN c.id END) clubs,
      COUNT(DISTINCT CASE WHEN COALESCE(c.source_active,1)=1 AND COALESCE(c.tm_url,'')<>'' THEN c.id END) mapped,
      COUNT(DISTINCT CASE WHEN COALESCE(c.source_active,1)=1 THEN p.id END) players,
      COUNT(DISTINCT CASE WHEN COALESCE(c.source_active,1)=1 THEN s.id END) staff
      FROM leagues l
      LEFT JOIN clubs c ON c.league_id=l.id
      LEFT JOIN players p ON p.club_id=c.id
      LEFT JOIN staff s ON s.club_id=c.id
      WHERE {' AND '.join(where)} GROUP BY l.id'''
    rows=con.execute(sql, params).fetchall()
    sid=engine.current_season_id()
    result={}
    for r in rows:
        tr=con.execute('''SELECT COUNT(ct.id),COUNT(DISTINCT cts.club_id) FROM clubs c
          LEFT JOIN club_transfers ct ON ct.club_id=c.id AND ct.season_id=?
          LEFT JOIN club_transfer_summary cts ON cts.club_id=c.id AND cts.season_id=?
          WHERE c.league_id=? AND COALESCE(c.source_active,1)=1''',(sid,sid,r['id'])).fetchone()
        result[r['id']]={
            'league':r,'clubs':int(r['clubs'] or 0),'mapped':int(r['mapped'] or 0),
            'players':int(r['players'] or 0),'staff':int(r['staff'] or 0),
            'transfers':int(tr[0] or 0),'updated_clubs':int(tr[1] or 0),
        }
    con.close(); return result


def aggregate(items):
    keys=('clubs','mapped','players','staff','transfers','updated_clubs')
    return {k:sum(int((x or {}).get(k,0)) for x in items) for k in keys}


def club_stats(club_id):
    con=connect(); sid=engine.current_season_id()
    row=con.execute('''SELECT c.*,
      (SELECT COUNT(*) FROM players p WHERE p.club_id=c.id) players,
      (SELECT COUNT(*) FROM staff s WHERE s.club_id=c.id) staff,
      (SELECT COUNT(*) FROM club_transfers ct WHERE ct.club_id=c.id AND ct.season_id=?) transfers,
      (SELECT COUNT(*) FROM club_transfer_summary cs WHERE cs.club_id=c.id AND cs.season_id=?) updated
      FROM clubs c WHERE c.id=?''',(sid,sid,int(club_id))).fetchone(); con.close(); return row
