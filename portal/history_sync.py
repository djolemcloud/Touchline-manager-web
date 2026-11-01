from __future__ import annotations

import re
import unicodedata
import json
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from typing import Iterable
from urllib.parse import quote

import requests
from bs4 import BeautifulSoup

from .data import connect

RSSSF_BASE = 'https://www.rsssf.org/tabless/'
RSSSF_CHAMPIONS = 'https://www.rsssf.org/tablesj/joegchamp.html'
RSSSF_CUPS = 'https://www.rsssf.org/tablesj/joegcuphist.html'
RSSSF_TOP_ALLTIME = 'https://www.rsssf.org/tabless/servalltime.html'
RED_STAR_SEASONS_URL = 'https://en.wikipedia.org/wiki/List_of_Red_Star_Belgrade_seasons'
PARTIZAN_SEASONS_URL = 'https://en.wikipedia.org/wiki/List_of_FK_Partizan_seasons'
VOJVODINA_HISTORY_URL = 'https://en.wikipedia.org/wiki/FK_Vojvodina'
RED_STAR_OFFICIAL_URL = 'https://www.crvenazvezdafk.com/en/klub/muzej'
PARTIZAN_OFFICIAL_URL = 'https://www.partizan.rs/vesti/istorijat'
VOJVODINA_OFFICIAL_HISTORY_URL = 'https://www.fkvojvodina.rs/wp-content/uploads/2012/01/o_organizatoru_FK_VOJVODINA_KUP_2012.pdf'
HISTORY_SEED_VERSION = '1.4.0'
EXCEL_LEAGUE_SOURCE = 'Touchline History Master Excel'

# Modern Serbian top-flight champions. These are seeded so league history is
# useful immediately, before the optional full RSSSF table sync is run.
SUPER_CHAMPIONS = {
    '2006/07':'Crvena zvezda (Beograd)', '2007/08':'Partizan (Beograd)',
    '2008/09':'Partizan (Beograd)', '2009/10':'Partizan (Beograd)',
    '2010/11':'Partizan (Beograd)', '2011/12':'Partizan (Beograd)',
    '2012/13':'Partizan (Beograd)', '2013/14':'Crvena zvezda (Beograd)',
    '2014/15':'Partizan (Beograd)', '2015/16':'Crvena zvezda (Beograd)',
    '2016/17':'Partizan (Beograd)', '2017/18':'Crvena zvezda (Beograd)',
    '2018/19':'Crvena zvezda (Beograd)', '2019/20':'Crvena zvezda (Beograd)',
    '2020/21':'Crvena zvezda (Beograd)', '2021/22':'Crvena zvezda (Beograd)',
    '2022/23':'Crvena zvezda (Beograd)', '2023/24':'Crvena zvezda (Beograd)',
    '2024/25':'Crvena zvezda (Beograd)', '2025/26':'Crvena zvezda (Beograd)',
}

PRVA_CHAMPIONS = {
    '2006/07':'Mladost (Lučani)', '2007/08':'Javor (Ivanjica)',
    '2008/09':'BSK (Borča)', '2009/10':'Inđija', '2010/11':'BASK (Beograd)',
    '2011/12':'Radnički (Niš)', '2012/13':'Napredak (Kruševac)',
    '2013/14':'Mladost (Lučani)', '2014/15':'Radnik (Surdulica)',
    '2015/16':'Napredak (Kruševac)', '2016/17':'Mačva (Šabac)',
    '2017/18':'Proleter (Novi Sad)', '2018/19':'TSC (Bačka Topola)',
    '2019/20':'Zlatibor (Čajetina)', '2020/21':'Radnički 1923 (Kragujevac)',
    '2021/22':'Mladost GAT (Novi Sad)', '2022/23':'IMT (Novi Beograd)',
    '2023/24':'OFK Beograd', '2024/25':'Radnik (Surdulica)',
    '2025/26':'Zemun',
}


# Bundled complete final classifications for the two most recent completed
# seasons. This gives the installed portal meaningful club-history rows even
# before the optional online History Sync is run. Tuple layout:
# (position, club, P, W, D, L, GF, GA, points, promoted, relegated)
BUNDLED_FINAL_TABLES = {
    ('2024/25','super'): [
        (1,'Crvena zvezda (Beograd)',37,32,4,1,123,35,100,0,0),
        (2,'Partizan (Beograd)',37,21,10,6,73,40,73,0,0),
        (3,'Novi Pazar (Novi Pazar)',37,15,9,13,60,65,54,0,0),
        (4,'OFK Beograd',37,15,8,14,53,54,53,0,0),
        (5,'Radnički 1923 (Kragujevac)',37,15,8,14,60,53,53,0,0),
        (6,'Vojvodina (Novi Sad)',37,14,11,12,57,49,53,0,0),
        (7,'TSC (Bačka Topola)',37,15,5,17,59,58,50,0,0),
        (8,'Mladost (Lučani)',37,12,11,14,38,48,47,0,0),
        (9,'Čukarički (Beograd)',37,12,13,12,47,49,49,0,0),
        (10,'Železničar (Pančevo)',37,13,10,14,49,43,49,0,0),
        (11,'IMT (Novi Beograd)',37,13,9,15,49,55,48,0,0),
        (12,'Spartak (Subotica)',37,11,11,15,35,51,44,0,0),
        (13,'Radnički (Niš)',37,11,10,16,50,67,43,0,0),
        (14,'Napredak (Kruševac)',37,11,9,17,35,48,42,0,0),
        (15,'Tekstilac (Odžaci)',37,11,4,22,33,65,31,0,1),
        (16,'Jedinstvo (Ub)',37,7,4,26,32,73,25,0,1),
    ],
    ('2024/25','prva'): [
        (1,'Radnik (Surdulica)',37,23,9,5,64,19,78,1,0),
        (2,'Javor (Ivanjica)',37,19,12,6,44,25,69,1,0),
        (3,'Mladost (Novi Sad)',37,17,14,6,43,31,65,0,0),
        (4,'Mačva (Šabac)',37,16,9,12,40,33,57,0,0),
        (5,'Voždovac (Beograd)',37,14,13,10,38,29,55,0,0),
        (6,'Grafičar (Beograd)',37,13,14,10,49,49,53,0,0),
        (7,'OFK Vršac',37,13,10,14,34,39,49,0,0),
        (8,'Borac 1926 (Čačak)',37,13,9,15,46,49,48,0,0),
        (9,'Zemun (Zemun)',37,11,15,11,46,44,48,0,0),
        (10,'Trajal (Kruševac)',37,13,9,15,45,44,45,0,0),
        (11,'GFK Dubočica (Leskovac)',37,10,11,16,43,57,41,0,0),
        (12,'Smederevo 1924',37,10,11,16,31,47,41,0,0),
        (13,'Radnički (Sr. Mitrovica)',37,8,16,13,29,35,40,0,1),
        (14,'GFK Sloven (Ruma)',37,8,15,14,31,44,39,0,1),
        (15,'Inđija (Inđija)',37,7,11,19,33,52,32,0,1),
        (16,'Sloboda (Užice)',37,7,10,20,23,42,31,0,1),
    ],
    ('2025/26','super'): [
        (1,'Crvena zvezda (Beograd)',37,27,5,5,100,31,86,0,0),
        (2,'Vojvodina (Novi Sad)',37,23,7,7,66,35,76,0,0),
        (3,'Partizan (Beograd)',37,22,7,8,72,45,73,0,0),
        (4,'Železničar (Pančevo)',37,16,11,10,50,37,59,0,0),
        (5,'Novi Pazar (Novi Pazar)',37,14,10,13,46,54,52,0,0),
        (6,'OFK Beograd',37,12,13,12,49,47,49,0,0),
        (7,'Radnik (Surdulica)',37,11,13,13,44,48,46,0,0),
        (8,'Čukarički (Beograd)',37,10,15,12,47,50,45,0,0),
        (9,'IMT (Novi Beograd)',37,13,11,13,41,51,50,0,0),
        (10,'Radnički (Niš)',37,14,7,16,46,49,49,0,0),
        (11,'Mladost (Lučani)',37,11,13,13,33,52,46,0,0),
        (12,'Radnički 1923 (Kragujevac)',37,9,18,10,41,42,45,0,0),
        (13,'TSC (Bačka Topola)',37,11,12,14,45,45,45,0,1),
        (14,'Javor (Ivanjica)',37,10,12,15,41,54,42,0,1),
        (15,'Spartak (Subotica)',37,4,9,24,37,72,21,0,1),
        (16,'Napredak (Kruševac)',37,3,9,25,30,76,18,0,1),
    ],
    ('2025/26','prva'): [
        (1,'Zemun (Zemun)',35,20,10,5,55,33,70,1,0),
        (2,'Mačva (Šabac)',35,19,9,7,47,29,66,1,0),
        (3,'Voždovac (Beograd)',35,19,8,8,60,29,65,0,0),
        (4,'Loznica',35,17,10,8,37,26,61,0,0),
        (5,'OFK Vršac',35,10,16,9,35,36,46,0,0),
        (6,'Jedinstvo (Ub)',35,11,11,13,46,52,44,0,0),
        (7,'Smederevo 1924',39,17,8,14,53,39,59,0,0),
        (8,'SU Dinamo Jug (Vranje)',39,16,14,9,44,33,58,0,0),
        (9,'Borac 1926 (Čačak)',39,14,12,13,43,40,54,0,0),
        (10,'Grafičar (Beograd)',39,14,10,15,41,38,52,0,0),
        (11,'Tekstilac (Odžaci)',39,12,13,14,47,52,49,0,1),
        (12,'Kabel (Novi Sad)',39,10,13,16,38,56,43,0,1),
        (13,'FAP (Priboj)',39,10,10,19,41,58,40,0,1),
        (14,'GFK Dubočica (Leskovac)',39,8,13,18,39,60,37,0,0),
        (15,'Trajal (Kruševac)',39,9,8,22,40,56,35,0,1),
        (16,'Ušće (Novi Beograd)',39,6,11,22,31,60,29,0,1),
    ],
}

# Cup finalists from the national cup continuity. Each tuple is season,winner,runner-up.
SERBIA_CUP_FINALS = [
    ('1991/92','Partizan (Beograd)','Crvena zvezda (Beograd)'),
    ('1992/93','Crvena zvezda (Beograd)','Partizan (Beograd)'),
    ('1993/94','Partizan (Beograd)','Spartak (Subotica)'),
    ('1994/95','Crvena zvezda (Beograd)','Obilić (Beograd)'),
    ('1995/96','Crvena zvezda (Beograd)','Partizan (Beograd)'),
    ('1996/97','Crvena zvezda (Beograd)','Vojvodina (Novi Sad)'),
    ('1997/98','Partizan (Beograd)','Obilić (Beograd)'),
    ('1998/99','Crvena zvezda (Beograd)','Partizan (Beograd)'),
    ('1999/00','Crvena zvezda (Beograd)','Napredak (Kruševac)'),
    ('2000/01','Partizan (Beograd)','Crvena zvezda (Beograd)'),
    ('2001/02','Crvena zvezda (Beograd)','Sartid Smederevo'),
    ('2002/03','Sartid Smederevo','Crvena zvezda (Beograd)'),
    ('2003/04','Crvena zvezda (Beograd)','Budućnost Banatski Dvor'),
    ('2004/05','Železnik (Beograd)','Crvena zvezda (Beograd)'),
    ('2005/06','Crvena zvezda (Beograd)','OFK Beograd'),
    ('2006/07','Crvena zvezda (Beograd)','Vojvodina (Novi Sad)'),
    ('2007/08','Partizan (Beograd)','Zemun'),
    ('2008/09','Partizan (Beograd)','Sevojno'),
    ('2009/10','Crvena zvezda (Beograd)','Vojvodina (Novi Sad)'),
    ('2010/11','Partizan (Beograd)','Vojvodina (Novi Sad)'),
    ('2011/12','Crvena zvezda (Beograd)','Borac (Čačak)'),
    ('2012/13','Jagodina','Vojvodina (Novi Sad)'),
    ('2013/14','Vojvodina (Novi Sad)','Jagodina'),
    ('2014/15','Čukarički (Beograd)','Partizan (Beograd)'),
    ('2015/16','Partizan (Beograd)','Javor (Ivanjica)'),
    ('2016/17','Partizan (Beograd)','Crvena zvezda (Beograd)'),
    ('2017/18','Partizan (Beograd)','Mladost (Lučani)'),
    ('2018/19','Partizan (Beograd)','Crvena zvezda (Beograd)'),
    ('2019/20','Vojvodina (Novi Sad)','Partizan (Beograd)'),
    ('2020/21','Crvena zvezda (Beograd)','Partizan (Beograd)'),
    ('2021/22','Crvena zvezda (Beograd)','Partizan (Beograd)'),
    ('2022/23','Crvena zvezda (Beograd)','Čukarički (Beograd)'),
    ('2023/24','Crvena zvezda (Beograd)','Vojvodina (Novi Sad)'),
    ('2024/25','Crvena zvezda (Beograd)','Vojvodina (Novi Sad)'),
    ('2025/26','Crvena zvezda (Beograd)','Vojvodina (Novi Sad)'),
]

YUGOSLAV_CHAMPIONS = [
    ('1946/47','Partizan (Beograd)'),
    ('1948/49','Partizan (Beograd)'), ('1950/51','Crvena zvezda (Beograd)'),
    ('1952/53','Crvena zvezda (Beograd)'), ('1955/56','Crvena zvezda (Beograd)'),
    ('1956/57','Crvena zvezda (Beograd)'), ('1958/59','Crvena zvezda (Beograd)'),
    ('1959/60','Crvena zvezda (Beograd)'), ('1960/61','Partizan (Beograd)'),
    ('1961/62','Partizan (Beograd)'), ('1962/63','Partizan (Beograd)'),
    ('1963/64','Crvena zvezda (Beograd)'), ('1964/65','Partizan (Beograd)'),
    ('1965/66','Vojvodina (Novi Sad)'), ('1967/68','Crvena zvezda (Beograd)'),
    ('1968/69','Crvena zvezda (Beograd)'), ('1969/70','Crvena zvezda (Beograd)'),
    ('1972/73','Crvena zvezda (Beograd)'), ('1975/76','Partizan (Beograd)'),
    ('1976/77','Crvena zvezda (Beograd)'), ('1977/78','Partizan (Beograd)'),
    ('1979/80','Crvena zvezda (Beograd)'), ('1980/81','Crvena zvezda (Beograd)'),
    ('1982/83','Partizan (Beograd)'), ('1983/84','Crvena zvezda (Beograd)'),
    ('1985/86','Partizan (Beograd)'), ('1986/87','Partizan (Beograd)'),
    ('1987/88','Crvena zvezda (Beograd)'), ('1988/89','Vojvodina (Novi Sad)'),
    ('1989/90','Crvena zvezda (Beograd)'), ('1990/91','Crvena zvezda (Beograd)'),
    ('1991/92','Crvena zvezda (Beograd)'),
]

FRY_CHAMPIONS = [
    ('1992/93','Partizan (Beograd)'),
    ('1993/94','Partizan (Beograd)'), ('1994/95','Crvena zvezda (Beograd)'),
    ('1995/96','Partizan (Beograd)'), ('1996/97','Partizan (Beograd)'),
    ('1997/98','Obilić (Beograd)'), ('1998/99','Partizan (Beograd)'),
    ('1999/00','Crvena zvezda (Beograd)'), ('2000/01','Crvena zvezda (Beograd)'),
    ('2001/02','Partizan (Beograd)'), ('2002/03','Partizan (Beograd)'),
    ('2003/04','Crvena zvezda (Beograd)'), ('2004/05','Partizan (Beograd)'),
    ('2005/06','Crvena zvezda (Beograd)'),
]

# Verified post-war league placings for the three major Serbian clubs.
# Public display uses one football-season convention (YYYY/YY) even for the
# short calendar championships of 1950-1952; source notes retain provenance.
# Sources: RSSSF, club season archives, Wikipedia season lists and foot.dk.
RED_STAR_LEAGUE_PLACINGS = [
    ('1945/46','serbia-republic-1946',1),
    ('1946/47','yugoslavia-first-1946-1992',3),('1947/48','yugoslavia-first-1946-1992',5),('1948/49','yugoslavia-first-1946-1992',2),
    ('1949/50','yugoslavia-first-1946-1992',2),('1950/51','yugoslavia-first-1946-1992',1),('1951/52','yugoslavia-first-1946-1992',2),
    ('1952/53','yugoslavia-first-1946-1992',1),('1953/54','yugoslavia-first-1946-1992',3),('1954/55','yugoslavia-first-1946-1992',4),
    ('1955/56','yugoslavia-first-1946-1992',1),('1956/57','yugoslavia-first-1946-1992',1),('1957/58','yugoslavia-first-1946-1992',4),
    ('1958/59','yugoslavia-first-1946-1992',1),('1959/60','yugoslavia-first-1946-1992',1),('1960/61','yugoslavia-first-1946-1992',2),
    ('1961/62','yugoslavia-first-1946-1992',4),('1962/63','yugoslavia-first-1946-1992',7),('1963/64','yugoslavia-first-1946-1992',1),
    ('1964/65','yugoslavia-first-1946-1992',3),('1965/66','yugoslavia-first-1946-1992',5),('1966/67','yugoslavia-first-1946-1992',5),
    ('1967/68','yugoslavia-first-1946-1992',1),('1968/69','yugoslavia-first-1946-1992',1),('1969/70','yugoslavia-first-1946-1992',1),
    ('1970/71','yugoslavia-first-1946-1992',6),('1971/72','yugoslavia-first-1946-1992',2),('1972/73','yugoslavia-first-1946-1992',1),
    ('1973/74','yugoslavia-first-1946-1992',3),('1974/75','yugoslavia-first-1946-1992',3),('1975/76','yugoslavia-first-1946-1992',4),
    ('1976/77','yugoslavia-first-1946-1992',1),('1977/78','yugoslavia-first-1946-1992',2),('1978/79','yugoslavia-first-1946-1992',3),
    ('1979/80','yugoslavia-first-1946-1992',1),('1980/81','yugoslavia-first-1946-1992',1),('1981/82','yugoslavia-first-1946-1992',2),
    ('1982/83','yugoslavia-first-1946-1992',5),('1983/84','yugoslavia-first-1946-1992',1),('1984/85','yugoslavia-first-1946-1992',4),
    ('1985/86','yugoslavia-first-1946-1992',2),('1986/87','yugoslavia-first-1946-1992',3),('1987/88','yugoslavia-first-1946-1992',1),
    ('1988/89','yugoslavia-first-1946-1992',2),('1989/90','yugoslavia-first-1946-1992',1),('1990/91','yugoslavia-first-1946-1992',1),('1991/92','yugoslavia-first-1946-1992',1),
    ('1992/93','fr-yugoslavia-first-1992-2003',2),('1993/94','fr-yugoslavia-first-1992-2003',2),('1994/95','fr-yugoslavia-first-1992-2003',1),
    ('1995/96','fr-yugoslavia-first-1992-2003',2),('1996/97','fr-yugoslavia-first-1992-2003',2),('1997/98','fr-yugoslavia-first-1992-2003',2),
    ('1998/99','fr-yugoslavia-first-1992-2003',3),('1999/00','fr-yugoslavia-first-1992-2003',1),('2000/01','fr-yugoslavia-first-1992-2003',1),
    ('2001/02','fr-yugoslavia-first-1992-2003',2),('2002/03','fr-yugoslavia-first-1992-2003',2),
    ('2003/04','serbia-montenegro-first-2003-2006',1),('2004/05','serbia-montenegro-first-2003-2006',2),('2005/06','serbia-montenegro-first-2003-2006',1),
    ('2006/07','serbia-top-flight',1),('2007/08','serbia-top-flight',2),('2008/09','serbia-top-flight',3),('2009/10','serbia-top-flight',2),
    ('2010/11','serbia-top-flight',2),('2011/12','serbia-top-flight',2),('2012/13','serbia-top-flight',2),('2013/14','serbia-top-flight',1),
    ('2014/15','serbia-top-flight',2),('2015/16','serbia-top-flight',1),('2016/17','serbia-top-flight',2),('2017/18','serbia-top-flight',1),
    ('2018/19','serbia-top-flight',1),('2019/20','serbia-top-flight',1),('2020/21','serbia-top-flight',1),('2021/22','serbia-top-flight',1),
    ('2022/23','serbia-top-flight',1),('2023/24','serbia-top-flight',1),('2024/25','serbia-top-flight',1),('2025/26','serbia-top-flight',1),
]

PARTIZAN_LEAGUE_PLACINGS = [
    ('1946/47','yugoslavia-first-1946-1992',1),('1947/48','yugoslavia-first-1946-1992',3),('1948/49','yugoslavia-first-1946-1992',1),
    ('1949/50','yugoslavia-first-1946-1992',3),('1950/51','yugoslavia-first-1946-1992',4),('1951/52','yugoslavia-first-1946-1992',6),
    ('1952/53','yugoslavia-first-1946-1992',3),('1953/54','yugoslavia-first-1946-1992',2),('1954/55','yugoslavia-first-1946-1992',5),
    ('1955/56','yugoslavia-first-1946-1992',2),('1956/57','yugoslavia-first-1946-1992',4),('1957/58','yugoslavia-first-1946-1992',2),
    ('1958/59','yugoslavia-first-1946-1992',2),('1959/60','yugoslavia-first-1946-1992',3),('1960/61','yugoslavia-first-1946-1992',1),
    ('1961/62','yugoslavia-first-1946-1992',1),('1962/63','yugoslavia-first-1946-1992',1),('1963/64','yugoslavia-first-1946-1992',5),
    ('1964/65','yugoslavia-first-1946-1992',1),('1965/66','yugoslavia-first-1946-1992',11),('1966/67','yugoslavia-first-1946-1992',3),
    ('1967/68','yugoslavia-first-1946-1992',2),('1968/69','yugoslavia-first-1946-1992',3),('1969/70','yugoslavia-first-1946-1992',2),
    ('1970/71','yugoslavia-first-1946-1992',5),('1971/72','yugoslavia-first-1946-1992',5),('1972/73','yugoslavia-first-1946-1992',4),
    ('1973/74','yugoslavia-first-1946-1992',4),('1974/75','yugoslavia-first-1946-1992',6),('1975/76','yugoslavia-first-1946-1992',1),
    ('1976/77','yugoslavia-first-1946-1992',4),('1977/78','yugoslavia-first-1946-1992',1),('1978/79','yugoslavia-first-1946-1992',15),
    ('1979/80','yugoslavia-first-1946-1992',13),('1980/81','yugoslavia-first-1946-1992',8),('1981/82','yugoslavia-first-1946-1992',6),
    ('1982/83','yugoslavia-first-1946-1992',1),('1983/84','yugoslavia-first-1946-1992',2),('1984/85','yugoslavia-first-1946-1992',3),
    ('1985/86','yugoslavia-first-1946-1992',1),('1986/87','yugoslavia-first-1946-1992',1),('1987/88','yugoslavia-first-1946-1992',2),
    ('1988/89','yugoslavia-first-1946-1992',6),('1989/90','yugoslavia-first-1946-1992',4),('1990/91','yugoslavia-first-1946-1992',3),('1991/92','yugoslavia-first-1946-1992',2),
    ('1992/93','fr-yugoslavia-first-1992-2003',1),('1993/94','fr-yugoslavia-first-1992-2003',1),('1994/95','fr-yugoslavia-first-1992-2003',2),
    ('1995/96','fr-yugoslavia-first-1992-2003',1),('1996/97','fr-yugoslavia-first-1992-2003',1),('1997/98','fr-yugoslavia-first-1992-2003',3),
    ('1998/99','fr-yugoslavia-first-1992-2003',1),('1999/00','fr-yugoslavia-first-1992-2003',2),('2000/01','fr-yugoslavia-first-1992-2003',2),
    ('2001/02','fr-yugoslavia-first-1992-2003',1),('2002/03','fr-yugoslavia-first-1992-2003',1),
    ('2003/04','serbia-montenegro-first-2003-2006',2),('2004/05','serbia-montenegro-first-2003-2006',1),('2005/06','serbia-montenegro-first-2003-2006',2),
    ('2006/07','serbia-top-flight',2),('2007/08','serbia-top-flight',1),('2008/09','serbia-top-flight',1),('2009/10','serbia-top-flight',1),
    ('2010/11','serbia-top-flight',1),('2011/12','serbia-top-flight',1),('2012/13','serbia-top-flight',1),('2013/14','serbia-top-flight',2),
    ('2014/15','serbia-top-flight',1),('2015/16','serbia-top-flight',2),('2016/17','serbia-top-flight',1),('2017/18','serbia-top-flight',2),
    ('2018/19','serbia-top-flight',3),('2019/20','serbia-top-flight',2),('2020/21','serbia-top-flight',2),('2021/22','serbia-top-flight',2),
    ('2022/23','serbia-top-flight',4),('2023/24','serbia-top-flight',2),('2024/25','serbia-top-flight',2),('2025/26','serbia-top-flight',3),
]

VOJVODINA_LEAGUE_PLACINGS = [
    ('1946/47','serbia-regional-1946-47',1),('1947/48','yugoslavia-second-league',2),('1948/49','yugoslavia-first-1946-1992',9),
    ('1949/50','yugoslavia-second-league',2),('1950/51','yugoslavia-first-1946-1992',8),('1951/52','yugoslavia-first-1946-1992',8),
    ('1952/53','yugoslavia-first-1946-1992',4),('1953/54','yugoslavia-first-1946-1992',5),('1954/55','yugoslavia-first-1946-1992',6),
    ('1955/56','yugoslavia-first-1946-1992',5),('1956/57','yugoslavia-first-1946-1992',2),('1957/58','yugoslavia-first-1946-1992',5),
    ('1958/59','yugoslavia-first-1946-1992',3),('1959/60','yugoslavia-first-1946-1992',4),('1960/61','yugoslavia-first-1946-1992',5),
    ('1961/62','yugoslavia-first-1946-1992',2),('1962/63','yugoslavia-first-1946-1992',12),('1963/64','yugoslavia-first-1946-1992',7),
    ('1964/65','yugoslavia-first-1946-1992',9),('1965/66','yugoslavia-first-1946-1992',1),('1966/67','yugoslavia-first-1946-1992',4),
    ('1967/68','yugoslavia-first-1946-1992',13),('1968/69','yugoslavia-first-1946-1992',4),('1969/70','yugoslavia-first-1946-1992',11),
    ('1970/71','yugoslavia-first-1946-1992',15),('1971/72','yugoslavia-first-1946-1992',4),('1972/73','yugoslavia-first-1946-1992',13),
    ('1973/74','yugoslavia-first-1946-1992',12),('1974/75','yugoslavia-first-1946-1992',2),('1975/76','yugoslavia-first-1946-1992',5),
    ('1976/77','yugoslavia-first-1946-1992',14),('1977/78','yugoslavia-first-1946-1992',8),('1978/79','yugoslavia-first-1946-1992',12),
    ('1979/80','yugoslavia-first-1946-1992',16),('1980/81','yugoslavia-first-1946-1992',10),('1981/82','yugoslavia-first-1946-1992',10),
    ('1982/83','yugoslavia-first-1946-1992',9),('1983/84','yugoslavia-first-1946-1992',10),('1984/85','yugoslavia-first-1946-1992',16),
    ('1985/86','yugoslavia-first-1946-1992',18),('1986/87','yugoslavia-second-league-west',1),('1987/88','yugoslavia-first-1946-1992',10),
    ('1988/89','yugoslavia-first-1946-1992',1),('1989/90','yugoslavia-first-1946-1992',11),('1990/91','yugoslavia-first-1946-1992',7),('1991/92','yugoslavia-first-1946-1992',3),
    ('1992/93','fr-yugoslavia-first-1992-2003',3),('1993/94','fr-yugoslavia-first-1992-2003',3),('1994/95','fr-yugoslavia-first-1992-2003',3),
    ('1995/96','fr-yugoslavia-first-1992-2003',3),('1996/97','fr-yugoslavia-first-1992-2003',3),('1997/98','fr-yugoslavia-first-1992-2003',4),
    ('1998/99','fr-yugoslavia-first-1992-2003',4),('1999/00','fr-yugoslavia-first-1992-2003',10),('2000/01','fr-yugoslavia-first-1992-2003',6),
    ('2001/02','fr-yugoslavia-first-1992-2003',9),('2002/03','fr-yugoslavia-first-1992-2003',6),
    ('2003/04','serbia-montenegro-first-2003-2006',9),('2004/05','serbia-montenegro-first-2003-2006',8),('2005/06','serbia-montenegro-first-2003-2006',9),
    ('2006/07','serbia-top-flight',3),('2007/08','serbia-top-flight',3),('2008/09','serbia-top-flight',2),('2009/10','serbia-top-flight',5),
    ('2010/11','serbia-top-flight',3),('2011/12','serbia-top-flight',3),('2012/13','serbia-top-flight',3),('2013/14','serbia-top-flight',4),
    ('2014/15','serbia-top-flight',4),('2015/16','serbia-top-flight',4),('2016/17','serbia-top-flight',3),('2017/18','serbia-top-flight',8),
    ('2018/19','serbia-top-flight',7),('2019/20','serbia-top-flight',3),('2020/21','serbia-top-flight',4),('2021/22','serbia-top-flight',7),
    ('2022/23','serbia-top-flight',5),('2023/24','serbia-top-flight',4),('2024/25','serbia-top-flight',6),('2025/26','serbia-top-flight',2),
]

YUGOSLAV_CUP_FINALS = [
    ('1946/47','Partizan (Beograd)','Naša krila (Zemun)'),
    ('1947/48','Crvena zvezda (Beograd)','Partizan (Beograd)'),
    ('1948/49','Crvena zvezda (Beograd)','Naša krila (Zemun)'),
    ('1949/50','Crvena zvezda (Beograd)','Dinamo (Zagreb)'),
    ('1950/51','Dinamo (Zagreb)','Vojvodina (Novi Sad)'),
    ('1951/52','Partizan (Beograd)','Crvena zvezda (Beograd)'),
    ('1952/53','OFK Beograd','Hajduk (Split)'), ('1953/54','Partizan (Beograd)','Crvena zvezda (Beograd)'),
    ('1954/55','OFK Beograd','Hajduk (Split)'), ('1956/57','Partizan (Beograd)','Radnički (Beograd)'),
    ('1957/58','Crvena zvezda (Beograd)','Velež (Mostar)'),
    ('1958/59','Crvena zvezda (Beograd)','Partizan (Beograd)'),
    ('1959/60','Dinamo (Zagreb)','Partizan (Beograd)'),
    ('1961/62','OFK Beograd','Spartak (Subotica)'),
    ('1963/64','Crvena zvezda (Beograd)','Dinamo (Zagreb)'),
    ('1965/66','OFK Beograd','Dinamo (Zagreb)'),
    ('1967/68','Crvena zvezda (Beograd)','Bor'),
    ('1969/70','Crvena zvezda (Beograd)','Olimpija (Ljubljana)'),
    ('1970/71','Crvena zvezda (Beograd)','Sloboda (Tuzla)'),
    ('1972/73','Hajduk (Split)','Crvena zvezda (Beograd)'),
    ('1978/79','Rijeka','Partizan (Beograd)'),
    ('1979/80','Dinamo (Zagreb)','Crvena zvezda (Beograd)'),
    ('1981/82','Crvena zvezda (Beograd)','Dinamo (Zagreb)'),
    ('1983/84','Hajduk (Split)','Crvena zvezda (Beograd)'),
    ('1984/85','Crvena zvezda (Beograd)','Dinamo (Zagreb)'),
    ('1987/88','Borac (Banja Luka)','Crvena zvezda (Beograd)'),
    ('1988/89','Partizan (Beograd)','Velež (Mostar)'),
    ('1989/90','Crvena zvezda (Beograd)','Hajduk (Split)'),
    ('1990/91','Hajduk (Split)','Crvena zvezda (Beograd)'),
]

# Official UEFA club facts / history highlights used by the Hall of Fame block.
EUROPE_HIGHLIGHTS = [
    # Season labels follow the football season rather than the calendar year of
    # the final: 1991 events belong to 1990/91, 1978 trophy to 1977/78, etc.
    ('Crvena zvezda (Beograd)','1990/91','european-cup',1,0,'Šampion Evrope','European champions','UEFA'),
    ('Crvena zvezda (Beograd)','1990/91','intercontinental-cup',1,0,'Svetski / interkontinentalni šampion','World / Intercontinental champions','FK Crvena zvezda / UEFA'),
    ('Crvena zvezda (Beograd)','1990/91','uefa-super-cup',0,1,'Finalista UEFA Superkupa','UEFA Super Cup runners-up','UEFA'),
    ('Crvena zvezda (Beograd)','1978/79','uefa-cup',0,1,'Finalista Kupa UEFA','UEFA Cup runners-up','UEFA'),
    ('Crvena zvezda (Beograd)','1957/58','mitropa-cup',1,0,'Osvajač Srednjoevropskog kupa','Mitropa Cup winners','FK Crvena zvezda'),
    ('Crvena zvezda (Beograd)','1967/68','mitropa-cup',1,0,'Osvajač Srednjoevropskog kupa','Mitropa Cup winners','FK Crvena zvezda'),
    ('Partizan (Beograd)','1965/66','european-cup',0,1,'Finalista Kupa evropskih šampiona','European Cup runners-up','UEFA'),
    ('Partizan (Beograd)','1977/78','mitropa-cup',1,0,'Osvajač Srednjoevropskog kupa','Mitropa Cup winners','FK Partizan'),
    ('Vojvodina (Novi Sad)','1966/67','european-cup',0,0,'Četvrtfinale','Quarter-finals','UEFA / Wikipedia'),
    ('Vojvodina (Novi Sad)','1976/77','mitropa-cup',1,0,'Osvajač Srednjoevropskog kupa','Mitropa Cup winners','Wikipedia'),
    ('Vojvodina (Novi Sad)','1956/57','mitropa-cup',0,1,'Finalista Srednjoevropskog kupa','Mitropa Cup runners-up','Wikipedia'),
    ('Vojvodina (Novi Sad)','1997/98','intertoto-cup',0,1,'Finalista Intertoto kupa','Intertoto Cup runners-up','Wikipedia'),
    ('Vojvodina (Novi Sad)','1961/62','fairs-cup',0,0,'Četvrtfinale','Quarter-finals','Wikipedia'),
    ('Vojvodina (Novi Sad)','1967/68','fairs-cup',0,0,'Četvrtfinale','Quarter-finals','Wikipedia'),
]

# RSSSF all-time SuperLiga table 2006/07-2025/26. Keeping this compact
# snapshot in the package makes the league history page useful immediately;
# a full season sync later supersedes it without changing the UI.
TOP_ALLTIME = [
 ('CRVENA ZVEZDA BEOGRAD',20,679,526,93,60,1581,485,1671),
 ('PARTIZAN BEOGRAD',20,679,470,115,94,1401,526,1525),
 ('VOJVODINA NOVI SAD',20,679,328,168,183,980,690,1152),
 ('ČUKARIČKI BEOGRAD',17,587,237,153,197,774,689,864),
 ('SPARTAK ŽK SUBOTICA',17,581,195,149,237,660,769,734),
 ('RADNIČKI NIŠ',14,491,188,128,175,605,599,692),
 ('MLADOST LUČANI',13,464,160,125,179,515,626,605),
 ('NAPREDAK KRUŠEVAC',15,520,156,123,241,531,703,591),
 ('JAVOR MATIS IVANJICA',14,473,136,141,196,467,596,549),
 ('VOŽDOVAC BEOGRAD',12,419,145,96,178,467,549,531),
 ('NOVI PAZAR',12,417,136,98,183,458,590,506),
 ('OFK BEOGRAD',12,389,139,85,165,447,500,502),
 ('RAD BEOGRAD',13,429,129,105,195,418,551,492),
 ('TSC BAČKA TOPOLA',7,253,117,59,77,423,314,410),
 ('RADNIK SURDULICA',10,364,104,98,162,378,511,410),
 ('RADNIČKI 1923 KRAGUJEVAC',9,305,92,89,124,346,412,365),
 ('BORAC ČAČAK',10,329,89,89,151,277,408,356),
 ('JAGODINA',8,250,87,60,103,268,296,321),
 ('HAJDUK 1912 KULA',7,218,64,59,95,194,248,251),
 ('METALAC GORNJI MILANOVAC',7,239,59,66,114,217,334,243),
 ('SMEDEREVO 1924',6,185,50,42,93,153,240,192),
 ('SLOBODA UŽICE',4,120,45,32,43,136,145,167),
 ('PROLETER NOVI SAD',4,142,40,36,66,129,187,156),
 ('ŽELEZNIČAR PANČEVO',3,111,39,30,42,146,145,147),
 ('IMT NOVI BEOGRAD',3,111,37,29,45,133,159,140),
 ('BSK BORČA',4,120,32,30,58,94,170,126),
 ('OFK BAČKA BAČKA PALANKA',4,149,32,26,91,123,246,122),
 ('MAČVA ŠABAC',4,142,30,30,82,106,224,120),
 ('BANAT ZRENJANIN',3,98,25,26,47,91,141,101),
 ('DONJI SREM PEĆINCI',3,90,22,26,42,80,116,92),
 ('INĐIJA TOYO TIRES',3,98,24,14,60,84,161,86),
 ('KOLUBARA LAZAREVAC',2,74,25,12,37,70,122,81),
 ('ZEMUN',3,106,18,26,62,97,163,80),
 ('BEŽANIJA NOVI BEOGRAD',2,65,17,16,32,67,89,67),
 ('MLADOST APA APATIN',1,32,11,8,13,25,33,41),
 ('DINAMO VRANJE',1,37,9,6,22,24,67,33),
 ('TEKSTILAC ODŽACI',1,37,11,4,22,33,65,31),
 ('MLADOST GAT NOVI SAD',1,37,6,12,19,25,49,30),
 ('ZLATIBOR ČAJETINA',1,38,7,8,23,28,64,29),
 ('MLADI RADNIK 1926 POŽAREVAC',1,30,5,10,15,19,47,25),
 ('JEDINSTVO UB',1,37,7,4,26,32,73,25),
]

ALIASES = {
 'crvena zvezda beograd':'Red Star Belgrade','c zvezda beograd':'Red Star Belgrade','crvena zvezda':'Red Star Belgrade','red star belgrade':'Red Star Belgrade',
 'partizan beograd':'Partizan Belgrade','partizan':'Partizan Belgrade',
 'vojvodina novi sad':'Vojvodina Novi Sad','vojvodina':'Vojvodina Novi Sad',
 'radnicki 1923 kragujevac':'Radnicki 1923','radnicki kragujevac':'Radnicki 1923',
 'imt novi beograd':'IMT Novi Beograd','imt':'IMT Novi Beograd','radnik surdulica':'Radnik Surdulica',
 'mladost lucani':'Mladost Lucani','novi pazar':'Novi Pazar','zeleznicar pancevo':'Zeleznicar Pancevo',
 'cukaricki beograd':'Cukaricki','fk cukaricki':'Cukaricki','cukaricki':'Cukaricki','ofk beograd':'OFK Beograd','radnicki nis':'Radnicki Nis',
 'macva sabac':'Macva Sabac','macva':'Macva Sabac','zemun':'Zemun','borac cacak':'Borac 1926','borac 1926 cacak':'Borac 1926',
 'tsc backa topola':'TSC Backa Topola','tsc':'TSC Backa Topola','ofk vrsac':'OFK Vrsac','javor ivanjica':'Javor Matis','javor matis ivanjica':'Javor Matis',
 'spartak subotica':'Spartak Subotica','spartak zk subotica':'Spartak Subotica','spartak zk subotica':'Spartak Subotica','smederevo':'Smederevo 1924','smederevo 1924':'Smederevo 1924',
 'sartid smederevo':'Smederevo 1924','gfk dubocica leskovac':'GFK Dubocica 1923','napredak krusevac':'Napredak Krusevac',
 'metalac gornji milanovac':'Metalac Gornji Milanovac','metalac g milanovac':'Metalac Gornji Milanovac',
 'vozdovac beograd':'Vozdovac','vozdovac':'Vozdovac','loznica':'Loznica','jedinstvo ub':'Jedinstvo Ub',
 'graficar beograd':'RFK Graficar','rfk graficar beograd':'RFK Graficar','teleoptik zemun':'Teleoptik','teleoptik':'Teleoptik',
 'habitpharm javor ivanjica':'Javor Matis','habitpharm javor':'Javor Matis','cukaricki stankom beograd':'Cukaricki','cukaricki stankom':'Cukaricki',
 'spartak zlatibor v subotica':'Spartak Subotica','spartak zlatibor voda subotica':'Spartak Subotica','fk crvena zvezda beograd':'Red Star Belgrade',
 'fk partizan beograd':'Partizan Belgrade','fk vojvodina novi sad':'Vojvodina Novi Sad','fc sartid smederevo':'Smederevo 1924','fk zemun':'Zemun',
 'aik backa topola':'TSC Backa Topola','aik tsc backa topola':'TSC Backa Topola',
 'dubocica leskovac':'GFK Dubocica 1923','gfk dubocica':'GFK Dubocica 1923','fk bor':'Bor 1919','bor bor':'Bor 1919',
}

@dataclass
class TableRow:
    position:int
    club_name:str
    played:int
    wins:int
    draws:int
    losses:int
    gf:int
    ga:int
    points:float
    tail:str=''


def norm(value:str)->str:
    s=(value or '').lower().replace('đ','dj')
    s=unicodedata.normalize('NFKD',s)
    s=''.join(ch for ch in s if not unicodedata.combining(ch))
    s=re.sub(r'\[[^\]]*\]',' ',s)
    s=re.sub(r'\([^)]*\)',' ',s)
    s=s.replace('fk ',' ').replace('ofk ','ofk ')
    s=re.sub(r'[^a-z0-9]+',' ',s)
    return re.sub(r'\s+',' ',s).strip()


def _canonical_source_name(name:str)->str:
    # Preserve locality in unmatched historical identities; it is essential for
    # distinguishing clubs such as Mladost (Lučani) and Mladost (Novi Sad).
    return re.sub(r'\s+',' ',name or '').strip()


def _norm_match(value:str)->str:
    """Normalize identity while PRESERVING locality text in parentheses.

    RSSSF commonly writes e.g. ``Mladost (Lučani)`` or ``Radnički (Niš)``.
    Dropping the parenthetical text makes those identities ambiguous, so matching
    must keep it even though display normalization may omit it.
    """
    s=(value or '').lower().replace('đ','dj')
    s=unicodedata.normalize('NFKD',s)
    s=''.join(ch for ch in s if not unicodedata.combining(ch))
    s=re.sub(r'\[[^\]]*\]',' ',s)
    s=re.sub(r'[^a-z0-9]+',' ',s)
    return re.sub(r'\s+',' ',s).strip()


def _club_id(con, source_name:str):
    key=_norm_match(source_name)
    short=norm(source_name)
    target=ALIASES.get(key) or ALIASES.get(short)
    clubs=con.execute('SELECT id,name,cm_name FROM clubs').fetchall()
    if target:
        t=_norm_match(target)
        for r in clubs:
            if _norm_match(r['name'])==t or _norm_match(r['cm_name'] or '')==t:
                return r['id']
    # Conservative exact normalized match only; never fuzzy-merge identities.
    for r in clubs:
        names={_norm_match(r['name']),_norm_match(r['cm_name'] or '')}
        if key and key in names:
            return r['id']
        if short and short in names and sum(1 for x in clubs if short in {_norm_match(x['name']),_norm_match(x['cm_name'] or '')})==1:
            return r['id']
    return None


def _reconcile_history_club_links(con):
    """Attach previously-unmatched historical rows after identity rules improve."""
    moved=0; merged=0
    rows=con.execute("SELECT * FROM history_club_seasons WHERE club_id IS NULL AND COALESCE(club_name_override,'')<>''").fetchall()
    for r in rows:
        cid=_club_id(con,r['club_name_override'])
        if not cid: continue
        existing=con.execute('SELECT * FROM history_club_seasons WHERE club_id=? AND competition_id=? AND season_label=?',(cid,r['competition_id'],r['season_label'])).fetchone()
        if existing:
            con.execute('''UPDATE history_club_seasons SET
              is_champion=MAX(COALESCE(is_champion,0),?),is_runner_up=MAX(COALESCE(is_runner_up,0),?),
              promoted=MAX(COALESCE(promoted,0),?),relegated=MAX(COALESCE(relegated,0),?),highlight=MAX(COALESCE(highlight,0),?),
              position=COALESCE(position,?),played=COALESCE(played,?),wins=COALESCE(wins,?),draws=COALESCE(draws,?),losses=COALESCE(losses,?),
              goals_for=COALESCE(goals_for,?),goals_against=COALESCE(goals_against,?),points=COALESCE(points,?),
              stage_sr=COALESCE(stage_sr,?),stage_en=COALESCE(stage_en,?) WHERE id=?''',
              (r['is_champion'],r['is_runner_up'],r['promoted'],r['relegated'],r['highlight'],r['position'],r['played'],r['wins'],r['draws'],r['losses'],r['goals_for'],r['goals_against'],r['points'],r['stage_sr'],r['stage_en'],existing['id']))
            con.execute('DELETE FROM history_club_seasons WHERE id=?',(r['id'],)); merged+=1
        else:
            con.execute('UPDATE history_club_seasons SET club_id=?,club_name_override=NULL WHERE id=?',(cid,r['id'])); moved+=1
    return moved,merged



def _remove_shadow_history_rows(con):
    """Remove stale duplicate rows created by older, weaker identity rules."""
    removed=0
    # League tables: one club per final position. Prefer a linked row, otherwise
    # the newest row (which preserves locality in its override name).
    groups=con.execute('''SELECT competition_id,season_label,position,COUNT(*) n
        FROM history_club_seasons WHERE position IS NOT NULL
        GROUP BY competition_id,season_label,position HAVING COUNT(*)>1''').fetchall()
    for g in groups:
        rows=con.execute('''SELECT id,club_id FROM history_club_seasons
            WHERE competition_id=? AND season_label=? AND position=?
            ORDER BY (club_id IS NOT NULL) DESC,id DESC''',(g['competition_id'],g['season_label'],g['position'])).fetchall()
        for r in rows[1:]:
            con.execute('DELETE FROM history_club_seasons WHERE id=?',(r['id'],)); removed+=1
    # Knockout honours do not have league positions; champion and runner-up are
    # unique inside one competition/season.
    for flag in ('is_champion','is_runner_up'):
        groups=con.execute(f'''SELECT competition_id,season_label,COUNT(*) n FROM history_club_seasons
            WHERE position IS NULL AND {flag}=1 GROUP BY competition_id,season_label HAVING COUNT(*)>1''').fetchall()
        for g in groups:
            rows=con.execute(f'''SELECT id,club_id FROM history_club_seasons
                WHERE competition_id=? AND season_label=? AND position IS NULL AND {flag}=1
                ORDER BY (club_id IS NOT NULL) DESC,id DESC''',(g['competition_id'],g['season_label'])).fetchall()
            for r in rows[1:]:
                con.execute('DELETE FROM history_club_seasons WHERE id=?',(r['id'],)); removed+=1
    return removed

def ensure_history_schema(con):
    con.executescript('''
    CREATE TABLE IF NOT EXISTS history_all_time_cache(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      competition_id INTEGER NOT NULL,club_id INTEGER,club_name_override TEXT,
      seasons INTEGER,played INTEGER,wins INTEGER,draws INTEGER,losses INTEGER,
      goals_for INTEGER,goals_against INTEGER,points REAL,titles INTEGER DEFAULT 0,
      source_name TEXT,source_url TEXT,updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
      UNIQUE(competition_id,club_name_override),
      FOREIGN KEY(competition_id) REFERENCES history_competitions(id) ON DELETE CASCADE,
      FOREIGN KEY(club_id) REFERENCES clubs(id) ON DELETE SET NULL);
    CREATE INDEX IF NOT EXISTS idx_history_alltime_comp ON history_all_time_cache(competition_id,points DESC);
    CREATE TABLE IF NOT EXISTS history_meta(
      key TEXT PRIMARY KEY,value TEXT,updated_at TEXT DEFAULT CURRENT_TIMESTAMP
    );
    CREATE TABLE IF NOT EXISTS history_season_notes(
      level INTEGER NOT NULL,season_label TEXT NOT NULL,note_sr TEXT,note_en TEXT,source_url TEXT,
      updated_at TEXT DEFAULT CURRENT_TIMESTAMP,PRIMARY KEY(level,season_label)
    );
    ''')


def _competition(con,key,name_sr,name_en,ctype='league',level=None,current_league_id=None,sort_order=0,active_from=None,active_to=None,source_url=None):
    con.execute('''INSERT OR IGNORE INTO history_competitions(history_key,current_league_id,name_sr,name_en,comp_type,country_code,level,active_from,active_to,sort_order,source_name,source_url)
                   VALUES(?,?,?,?,?,?,?,?,?,?,?,?)''',(key,current_league_id,name_sr,name_en,ctype,'SRB',level,active_from,active_to,sort_order,'RSSSF / UEFA',source_url))
    con.execute('''UPDATE history_competitions SET current_league_id=COALESCE(?,current_league_id),name_sr=?,name_en=?,comp_type=?,level=?,active_from=?,active_to=?,sort_order=?,source_url=COALESCE(?,source_url)
                   WHERE history_key=?''',(current_league_id,name_sr,name_en,ctype,level,active_from,active_to,sort_order,source_url,key))
    return con.execute('SELECT id FROM history_competitions WHERE history_key=?',(key,)).fetchone()[0]


def _upsert_history(con,competition_id,season,club_name,position=None,played=None,wins=None,draws=None,losses=None,gf=None,ga=None,points=None,champion=0,runner_up=0,promoted=0,relegated=0,highlight=0,stage_sr=None,stage_en=None,source_name='RSSSF',source_url=None,source_note=None,fill_only=False,forced_club_id=None):
    # 1.6.4 FINAL FREEZE rule: historical DOMESTIC LEAGUE rows are Excel-only.
    # Europe/world/cups may still use their verified dedicated sources, but no
    # RSSSF/Wikipedia/COMET parser is allowed to write a league row into the
    # club-history table. Current/live standings live in separate tables.
    _ctype=con.execute('SELECT comp_type FROM history_competitions WHERE id=?',(competition_id,)).fetchone()
    if _ctype and (_ctype[0] or '').lower()=='league' and source_name != EXCEL_LEAGUE_SOURCE:
        cid=int(forced_club_id) if forced_club_id is not None else _club_id(con,club_name)
        return 'updated',cid
    cid=int(forced_club_id) if forced_club_id is not None else _club_id(con,club_name)
    override=None if cid else _canonical_source_name(club_name)
    if cid:
        existing=con.execute('SELECT id FROM history_club_seasons WHERE club_id=? AND competition_id=? AND season_label=?',(cid,competition_id,season)).fetchone()
    else:
        existing=con.execute('SELECT id FROM history_club_seasons WHERE club_id IS NULL AND competition_id=? AND season_label=? AND lower(COALESCE(club_name_override,\'\'))=lower(?)',(competition_id,season,override)).fetchone()
    if existing:
        if fill_only:
            con.execute('''UPDATE history_club_seasons SET
              position=COALESCE(position,?),played=COALESCE(played,?),wins=COALESCE(wins,?),draws=COALESCE(draws,?),losses=COALESCE(losses,?),
              goals_for=COALESCE(goals_for,?),goals_against=COALESCE(goals_against,?),points=COALESCE(points,?),
              is_champion=MAX(COALESCE(is_champion,0),?),is_runner_up=MAX(COALESCE(is_runner_up,0),?),
              promoted=MAX(COALESCE(promoted,0),?),relegated=MAX(COALESCE(relegated,0),?),highlight=MAX(COALESCE(highlight,0),?),
              stage_sr=COALESCE(stage_sr,?),stage_en=COALESCE(stage_en,?),source_name=COALESCE(source_name,?),source_url=COALESCE(source_url,?),source_note=COALESCE(source_note,?)
              WHERE id=?''',(position,played,wins,draws,losses,gf,ga,points,champion,runner_up,promoted,relegated,highlight,stage_sr,stage_en,source_name,source_url,source_note,existing[0]))
        else:
            con.execute('''UPDATE history_club_seasons SET club_name_override=?,position=?,played=?,wins=?,draws=?,losses=?,goals_for=?,goals_against=?,points=?,
              is_champion=?,is_runner_up=?,promoted=?,relegated=?,highlight=MAX(COALESCE(highlight,0),?),stage_sr=COALESCE(?,stage_sr),stage_en=COALESCE(?,stage_en),
              source_name=?,source_url=?,source_note=?,updated_at=CURRENT_TIMESTAMP WHERE id=?''',
              (override,position,played,wins,draws,losses,gf,ga,points,champion,runner_up,promoted,relegated,highlight,stage_sr,stage_en,source_name,source_url,source_note,existing[0]))
        return 'updated',cid
    con.execute('''INSERT INTO history_club_seasons(club_id,competition_id,season_label,club_name_override,position,stage_sr,stage_en,played,wins,draws,losses,goals_for,goals_against,points,is_champion,is_runner_up,promoted,relegated,highlight,source_name,source_url,source_note)
                   VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
                (cid,competition_id,season,override,position,stage_sr,stage_en,played,wins,draws,losses,gf,ga,points,champion,runner_up,promoted,relegated,highlight,source_name,source_url,source_note))
    return 'inserted',cid



def _season_start_year(label:str):
    m=re.match(r'^\s*(\d{4})',str(label or ''))
    return int(m.group(1)) if m else None


def _migrate_history_keys(con):
    """Normalize old history keys without losing user data.

    1.2.x used one combined SRJ/SCG league and cup. 1.3.0 keeps the public
    timeline continuous, but stores the historically correct competition era
    for every season.
    """
    renames={
        'yugoslavia-first-1945-1991':'yugoslavia-first-1946-1992',
        'yugoslavia-serbia-montenegro-first-1991-2006':'legacy-first-1992-2006',
        'yugoslavia-serbia-montenegro-first-1992-2006':'legacy-first-1992-2006',
    }
    for old,new in renames.items():
        oldrow=con.execute('SELECT id FROM history_competitions WHERE history_key=?',(old,)).fetchone()
        newrow=con.execute('SELECT id FROM history_competitions WHERE history_key=?',(new,)).fetchone()
        if oldrow and not newrow:
            con.execute('UPDATE history_competitions SET history_key=? WHERE id=?',(new,oldrow[0]))
        elif oldrow and newrow and oldrow[0]!=newrow[0]:
            con.execute('UPDATE history_club_seasons SET competition_id=? WHERE competition_id=?',(newrow[0],oldrow[0]))
            try:
                con.execute('UPDATE history_all_time_cache SET competition_id=? WHERE competition_id=?',(newrow[0],oldrow[0]))
            except Exception:
                pass
            con.execute('DELETE FROM history_competitions WHERE id=?',(oldrow[0],))

    # In 1.2.x this same key covered 1992/93-2005/06. In 1.3.0 it means only
    # the actual Serbia-and-Montenegro era (2003/04-2005/06). Rename only the
    # legacy wide-range row; never churn the canonical competition on restart.
    cup=con.execute("SELECT id,active_from,active_to FROM history_competitions WHERE history_key='serbia-montenegro-cup'").fetchone()
    if cup and (cup['active_from'] or '')!='2003/04':
        old_legacy=con.execute("SELECT id FROM history_competitions WHERE history_key='legacy-cup-1992-2006'").fetchone()
        if old_legacy and old_legacy['id']!=cup['id']:
            con.execute('UPDATE history_club_seasons SET competition_id=? WHERE competition_id=?',(old_legacy['id'],cup['id']))
            con.execute('DELETE FROM history_competitions WHERE id=?',(cup['id'],))
        else:
            con.execute("UPDATE history_competitions SET history_key='legacy-cup-1992-2006' WHERE id=?",(cup['id'],))


def _migrate_legacy_era_rows(con,fry_id,scg_id,fry_cup_id,scg_cup_id):
    """Split legacy combined SRJ/SCG competitions at the 2003/04 boundary."""
    moved=0
    legacy=con.execute("SELECT id FROM history_competitions WHERE history_key='legacy-first-1992-2006'").fetchone()
    if legacy:
        for r in con.execute('SELECT id,season_label FROM history_club_seasons WHERE competition_id=?',(legacy[0],)).fetchall():
            y=_season_start_year(r['season_label'])
            target=fry_id if y is not None and y<=2002 else scg_id
            cur=con.execute('UPDATE history_club_seasons SET competition_id=? WHERE id=?',(target,r['id']))
            moved+=max(cur.rowcount or 0,0)
        try:
            con.execute('DELETE FROM history_all_time_cache WHERE competition_id=?',(legacy[0],))
        except Exception:
            pass
        con.execute('DELETE FROM history_competitions WHERE id=?',(legacy[0],))

    legacy=con.execute("SELECT id FROM history_competitions WHERE history_key='legacy-cup-1992-2006'").fetchone()
    if legacy:
        for r in con.execute('SELECT id,season_label FROM history_club_seasons WHERE competition_id=?',(legacy[0],)).fetchall():
            y=_season_start_year(r['season_label'])
            target=fry_cup_id if y is not None and y<=2002 else scg_cup_id
            cur=con.execute('UPDATE history_club_seasons SET competition_id=? WHERE id=?',(target,r['id']))
            moved+=max(cur.rowcount or 0,0)
        con.execute('DELETE FROM history_competitions WHERE id=?',(legacy[0],))
    return moved


def _normalize_legacy_history_rows(con,big3_ids,yug_cup_id):
    """Remove obsolete calendar-year seed rows from 1.2.x.

    The user-facing project convention is one football-season label (YYYY/YY).
    Verified normalized rows are re-seeded immediately afterwards.
    """
    changed=0
    ids=[int(x) for x in big3_ids if x is not None]
    if ids:
        q=','.join('?'*len(ids))
        cur=con.execute(f"""DELETE FROM history_club_seasons
            WHERE club_id IN ({q})
              AND season_label IN ('1950','1951','1952')
              AND competition_id IN (SELECT id FROM history_competitions WHERE comp_type='league')""",ids)
        changed+=max(cur.rowcount or 0,0)
        cur=con.execute(f"""DELETE FROM history_club_seasons
            WHERE club_id IN ({q})
              AND length(trim(season_label))=4
              AND competition_id IN (SELECT id FROM history_competitions WHERE comp_type IN ('europe','world'))""",ids)
        changed+=max(cur.rowcount or 0,0)

    # Old cup seed used calendar years. It contained only final/winner records;
    # replace those with the normalized season-labelled verified final list.
    cur=con.execute("""DELETE FROM history_club_seasons
        WHERE competition_id=? AND length(trim(season_label))=4
          AND (is_champion=1 OR is_runner_up=1 OR stage_sr IN ('Osvajač','Finalista','Finalistaa'))""",(yug_cup_id,))
    changed+=max(cur.rowcount or 0,0)
    return changed


def _same_club(con,row,expected_name):
    expected_id=_club_id(con,expected_name)
    if expected_id is not None and row['club_id'] is not None:
        return int(row['club_id'])==int(expected_id)
    actual=row['club_name_override']
    if not actual and row['club_id'] is not None:
        cr=con.execute('SELECT name,cm_name FROM clubs WHERE id=?',(row['club_id'],)).fetchone()
        actual=(cr['cm_name'] or cr['name']) if cr else ''
    return _norm_match(actual or '')==_norm_match(expected_name)


def _repair_modern_competition_contamination(con,competition_id,expected_champions):
    # Remove an online-synced season when its extracted champion is impossible.
    removed=0
    for season,expected in expected_champions.items():
        champs=con.execute('''SELECT * FROM history_club_seasons
            WHERE competition_id=? AND season_label=? AND is_champion=1''',(competition_id,season)).fetchall()
        if any(not _same_club(con,r,expected) for r in champs) or len(champs)>1:
            cur=con.execute('''DELETE FROM history_club_seasons
                WHERE competition_id=? AND season_label=?
                  AND COALESCE(source_note,'') LIKE 'Final table%' ''',(competition_id,season))
            removed+=cur.rowcount if cur.rowcount and cur.rowcount>0 else 0
            leftovers=con.execute('SELECT * FROM history_club_seasons WHERE competition_id=? AND season_label=? AND is_champion=1',(competition_id,season)).fetchall()
            for r in leftovers:
                if not _same_club(con,r,expected):
                    con.execute('DELETE FROM history_club_seasons WHERE id=?',(r['id'],)); removed+=1
    return removed


def _repair_history_integrity(con,super_id,prva_id,yug_id,fry_id,scg_id,yug_cup_id,fry_cup_id,scg_cup_id,big3_ids):
    repaired=0
    cur=con.execute("UPDATE history_club_seasons SET stage_sr=REPLACE(stage_sr,'Finalistaa','Finalista') WHERE stage_sr LIKE '%Finalistaa%'")
    repaired+=max(cur.rowcount or 0,0)

    # Era boundaries: these rows belong to the adjacent historical competition.
    cur=con.execute("DELETE FROM history_club_seasons WHERE competition_id=? AND season_label='1945/46'",(yug_id,))
    repaired+=max(cur.rowcount or 0,0)
    for bad_id in (fry_id,scg_id):
        cur=con.execute("DELETE FROM history_club_seasons WHERE competition_id=? AND season_label='1991/92'",(bad_id,))
        repaired+=max(cur.rowcount or 0,0)
    for bad_id in (fry_cup_id,scg_cup_id):
        cur=con.execute("DELETE FROM history_club_seasons WHERE competition_id=? AND season_label='1991/92'",(bad_id,))
        repaired+=max(cur.rowcount or 0,0)

    repaired+=_normalize_legacy_history_rows(con,big3_ids,yug_cup_id)
    repaired+=_repair_modern_competition_contamination(con,super_id,SUPER_CHAMPIONS)
    repaired+=_repair_modern_competition_contamination(con,prva_id,PRVA_CHAMPIONS)

    # Vojvodina's 1976 Rapan/Intertoto result was first place in a group, not a
    # single overall cup title. Older builds incorrectly promoted it to an
    # 'Intertoto Cup winner' honour. Remove that row permanently.
    intertoto=con.execute("SELECT id FROM history_competitions WHERE history_key='intertoto-cup'").fetchone()
    voj=_club_id(con,'Vojvodina (Novi Sad)')
    if intertoto and voj:
        cur=con.execute("DELETE FROM history_club_seasons WHERE competition_id=? AND club_id=? AND season_label='1975/76' AND (COALESCE(stage_sr,'') LIKE '%Intertoto%' OR COALESCE(is_champion,0)=1)",(intertoto[0],voj))
        repaired+=max(cur.rowcount or 0,0)
    return repaired


def seed_history(force=False):
    con=connect()
    ensure_history_schema(con)
    if not force:
        marker=con.execute("SELECT value FROM history_meta WHERE key='seed_version'").fetchone()
        if marker and marker['value']==HISTORY_SEED_VERSION:
            con.close()
            return {'inserted':0,'updated':0,'matched':0,'unmatched':0,'alltime':0,
                    'reconciled':0,'merged':0,'repaired':0,'migrated':0,'deduplicated':0,'skipped':1}
    _migrate_history_keys(con)

    super_lid=con.execute("SELECT id FROM leagues WHERE name='SuperLiga Srbije' ORDER BY id LIMIT 1").fetchone()
    prva_lid=con.execute("SELECT id FROM leagues WHERE name='Prva liga Srbije' ORDER BY id LIMIT 1").fetchone()

    # Domestic league lineage. These are separate historical competition
    # identities in storage, but one continuous club-history timeline in UI.
    super_id=_competition(con,'serbia-top-flight','SuperLiga Srbije','Serbian SuperLiga','league',1,super_lid[0] if super_lid else None,10,'2006/07',None,RSSSF_CHAMPIONS)
    prva_id=_competition(con,'serbia-second-tier','Prva liga Srbije','Serbian First League','league',2,prva_lid[0] if prva_lid else None,20,'2006/07',None,None)
    republic_id=_competition(con,'serbia-republic-1946','Prvenstvo NR Srbije',"People's Republic of Serbia League",'league',1,None,29,'1945/46','1945/46',RED_STAR_OFFICIAL_URL)
    yug_id=_competition(con,'yugoslavia-first-1946-1992','Prva savezna liga Jugoslavije','Yugoslav First League','league',1,None,30,'1946/47','1991/92',RSSSF_CHAMPIONS)
    fry_id=_competition(con,'fr-yugoslavia-first-1992-2003','Prva liga SR Jugoslavije','FR Yugoslavia First League','league',1,None,31,'1992/93','2002/03',RSSSF_CHAMPIONS)
    scg_id=_competition(con,'serbia-montenegro-first-2003-2006','Prva liga Srbije i Crne Gore','Serbia and Montenegro First League','league',1,None,32,'2003/04','2005/06',RSSSF_CHAMPIONS)

    # Historical lower tiers needed for a truthful club path (not counted as
    # national championships).
    _competition(con,'serbia-regional-1946-47','Prvenstvo Srbije – I grupa','Serbian Championship – Group I','league',3,None,35,'1946/47','1946/47',VOJVODINA_HISTORY_URL)
    _competition(con,'yugoslavia-second-league','Druga savezna liga Jugoslavije','Yugoslav Second League','league',2,None,36,'1947/48','1950/51',RSSSF_CHAMPIONS)
    _competition(con,'yugoslavia-second-league-west','Druga savezna liga Jugoslavije – Zapad','Yugoslav Second League – West','league',2,None,37,'1986/87','1986/87',RSSSF_CHAMPIONS)
    # Generic importer 1.4.0 uses these canonical historical second-tier IDs.
    _competition(con,'fr-yugoslavia-second-1992-2003','Druga liga SR Jugoslavije','FR Yugoslavia Second League','league',2,None,33,'1992/93','2002/03',None)
    _competition(con,'serbia-montenegro-second-2003-2006','Druga liga Srbije i Crne Gore','Serbia and Montenegro Second League','league',2,None,34,'2003/04','2005/06',None)
    con.execute("UPDATE history_competitions SET active_from='1947/48',active_to='1991/92',name_sr='Druga savezna liga Jugoslavije',name_en='Yugoslav Second League',level=2 WHERE history_key='yugoslavia-second-league'")

    # National-cup lineage.
    yug_cup_id=_competition(con,'yugoslavia-cup','Kup maršala Tita / Kup Jugoslavije','Marshal Tito Cup / Yugoslav Cup','cup',None,None,40,'1946/47','1991/92',RSSSF_CUPS)
    fry_cup_id=_competition(con,'fr-yugoslavia-cup','Kup SR Jugoslavije','FR Yugoslavia Cup','cup',None,None,41,'1992/93','2002/03',RSSSF_CUPS)
    scg_cup_id=_competition(con,'serbia-montenegro-cup','Kup Srbije i Crne Gore','Serbia and Montenegro Cup','cup',None,None,42,'2003/04','2005/06',RSSSF_CUPS)
    srb_cup_id=_competition(con,'serbia-cup','Kup Srbije','Serbian Cup','cup',None,None,43,'2006/07',None,RSSSF_CUPS)

    # International milestones.
    _competition(con,'european-cup','Kup evropskih šampiona','European Cup','europe',None,None,50,'1955/56','1991/92','https://www.uefa.com/uefachampionsleague/')
    _competition(con,'uefa-cup','Kup UEFA','UEFA Cup','europe',None,None,51,'1971/72','2008/09','https://www.uefa.com/uefaeuropaleague/')
    _competition(con,'uefa-super-cup','UEFA Superkup','UEFA Super Cup','europe',None,None,52,'1973/74',None,'https://www.uefa.com/uefasupercup/')
    _competition(con,'intercontinental-cup','Interkontinentalni kup','Intercontinental Cup','world',None,None,53,'1960/61','2004/05',RED_STAR_OFFICIAL_URL)
    _competition(con,'mitropa-cup','Srednjoevropski / Mitropa kup','Mitropa / Central European Cup','europe',None,None,54,None,None,'https://www.rsssf.org/tablesm/mit.html')
    _competition(con,'intertoto-cup','Intertoto kup','Intertoto Cup','europe',None,None,55,None,None,VOJVODINA_HISTORY_URL)
    _competition(con,'fairs-cup','Kup sajamskih gradova','Inter-Cities Fairs Cup','europe',None,None,56,'1955/56','1970/71','https://www.uefa.com/')

    # Split 1.2.x combined competition rows after the destination identities
    # exist, so existing installations are migrated in-place.
    migrated=_migrate_legacy_era_rows(con,fry_id,scg_id,fry_cup_id,scg_cup_id)
    comp_by_key={r['history_key']:r['id'] for r in con.execute('SELECT id,history_key FROM history_competitions')}

    big3_ids=[
        _club_id(con,'Crvena zvezda (Beograd)'),
        _club_id(con,'Partizan (Beograd)'),
        _club_id(con,'Vojvodina (Novi Sad)'),
    ]
    reconciled,merged=_reconcile_history_club_links(con)
    repaired=_repair_history_integrity(con,super_id,prva_id,yug_id,fry_id,scg_id,yug_cup_id,fry_cup_id,scg_cup_id,big3_ids)

    counts={'inserted':0,'updated':0,'matched':0,'unmatched':0,'alltime':0,
            'reconciled':reconciled,'merged':merged,'repaired':repaired,'migrated':migrated}

    def add(comp,season,club,**kw):
        status,cid=_upsert_history(con,comp,season,club,fill_only=True,**kw)
        counts[status]+=1
        counts['matched' if cid else 'unmatched']+=1

    comp_levels={r['id']:r['level'] for r in con.execute('SELECT id,level FROM history_competitions')}
    def seed_placements(club_name,rows,source_name,source_url):
        for season,key,pos in rows:
            comp=comp_by_key.get(key)
            if not comp:
                counts['unmatched']+=1
                continue
            is_top=(comp_levels.get(comp) or 0)==1
            status,cid=_upsert_history(
                con,comp,season,club_name,position=pos,
                champion=1 if pos==1 else 0,
                runner_up=1 if pos==2 else 0,
                highlight=1 if (pos==1 and is_top) else 0,
                source_name=source_name,source_url=source_url,
                source_note='Verified season placement; public season labels normalized to YYYY/YY.',
                fill_only=True)
            counts[status]+=1
            counts['matched' if cid else 'unmatched']+=1

    # Complete league paths for the three major Serbian clubs.
    seed_placements('Crvena zvezda (Beograd)',RED_STAR_LEAGUE_PLACINGS,'Wikipedia / FK Crvena zvezda / RSSSF',RED_STAR_SEASONS_URL)
    seed_placements('Partizan (Beograd)',PARTIZAN_LEAGUE_PLACINGS,'Wikipedia / FK Partizan',PARTIZAN_SEASONS_URL)
    seed_placements('Vojvodina (Novi Sad)',VOJVODINA_LEAGUE_PLACINGS,'Wikipedia / FK Vojvodina / RSSSF',VOJVODINA_HISTORY_URL)

    # Champions seed makes competition history useful even before online sync.
    for season,club in SUPER_CHAMPIONS.items():
        add(super_id,season,club,position=1,champion=1,highlight=1,source_url=RSSSF_CHAMPIONS)
    for season,club in PRVA_CHAMPIONS.items():
        add(prva_id,season,club,position=1,champion=1,highlight=1,source_url='https://www.rsssf.org/tabless/serv2champ.html')

    # Complete bundled classifications for the two latest completed seasons.
    for (season,kind),rows in BUNDLED_FINAL_TABLES.items():
        comp=super_id if kind=='super' else prva_id
        start_year=int(season[:4]); src=rsssf_url_for_season(start_year)
        for pos,club,p,w,d,l,gf,ga,pts,prom,rel in rows:
            status,cid=_upsert_history(con,comp,season,club,position=pos,played=p,wins=w,draws=d,losses=l,gf=gf,ga=ga,points=pts,
                champion=1 if pos==1 else 0,runner_up=1 if pos==2 else 0,promoted=prom,relegated=rel,
                highlight=1 if pos<=2 else 0,source_name='RSSSF',source_url=src,source_note='Bundled final classification')
            counts[status]+=1
            counts['matched' if cid else 'unmatched']+=1

    for season,club in YUGOSLAV_CHAMPIONS:
        add(yug_id,season,club,position=1,champion=1,highlight=1,source_url=RSSSF_CHAMPIONS)
    for season,club in FRY_CHAMPIONS:
        y=_season_start_year(season)
        comp=fry_id if y is not None and y<=2002 else scg_id
        add(comp,season,club,position=1,champion=1,highlight=1,source_url=RSSSF_CHAMPIONS)

    # Verified national cup finals, normalized to football-season labels.
    for season,winner,runner in YUGOSLAV_CUP_FINALS:
        add(yug_cup_id,season,winner,champion=1,highlight=1,stage_sr='Osvajač',stage_en='Winner',source_url=RSSSF_CUPS)
        add(yug_cup_id,season,runner,runner_up=1,highlight=1,stage_sr='Finalista',stage_en='Runner-up',source_url=RSSSF_CUPS)
    for season,winner,runner in SERBIA_CUP_FINALS:
        y=_season_start_year(season)
        if y is not None and y<=1991:
            comp=yug_cup_id
        elif y is not None and y<=2002:
            comp=fry_cup_id
        elif y is not None and y<=2005:
            comp=scg_cup_id
        else:
            comp=srb_cup_id
        add(comp,season,winner,champion=1,highlight=1,stage_sr='Osvajač',stage_en='Winner',source_url=RSSSF_CUPS)
        add(comp,season,runner,runner_up=1,highlight=1,stage_sr='Finalista',stage_en='Runner-up',source_url=RSSSF_CUPS)

    for club,season,key,ch,ru,sr,en,src in EUROPE_HIGHLIGHTS:
        comp=comp_by_key.get(key)
        if not comp:
            continue
        source_url='https://www.uefa.com/'
        if 'Crvena' in club:
            source_url=RED_STAR_OFFICIAL_URL
        elif 'Partizan' in club:
            source_url=PARTIZAN_OFFICIAL_URL
        elif 'Vojvodina' in club:
            source_url=VOJVODINA_HISTORY_URL
        add(comp,season,club,champion=ch,runner_up=ru,highlight=1,stage_sr=sr,stage_en=en,source_name=src,source_url=source_url)

    counts['deduplicated']=_remove_shadow_history_rows(con)

    # Seed current SuperLiga all-time snapshot.
    con.execute('DELETE FROM history_all_time_cache WHERE competition_id=?',(super_id,))
    title_counts_by_id={}
    title_counts_by_name={}
    for r in con.execute("SELECT club_id,club_name_override,COUNT(*) n FROM history_club_seasons WHERE competition_id=? AND is_champion=1 GROUP BY club_id,club_name_override",(super_id,)).fetchall():
        if r['club_id'] is not None:
            title_counts_by_id[int(r['club_id'])]=int(r['n'] or 0)
        elif r['club_name_override']:
            title_counts_by_name[norm(r['club_name_override'])]=int(r['n'] or 0)
    for name,seasons,played,w,d,l,gf,ga,points in TOP_ALLTIME:
        cid=_club_id(con,name); override=None if cid else _canonical_source_name(name)
        t=title_counts_by_id.get(int(cid),0) if cid is not None else title_counts_by_name.get(norm(name),0)
        con.execute("INSERT INTO history_all_time_cache(competition_id,club_id,club_name_override,seasons,played,wins,draws,losses,goals_for,goals_against,points,titles,source_name,source_url) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)",(super_id,cid,override,seasons,played,w,d,l,gf,ga,points,t,'RSSSF',RSSSF_TOP_ALLTIME))
        counts['alltime']+=1

    # Known source-contaminated women club identity. Competition phase belongs
    # to the league, never to club.name.
    con.execute("UPDATE clubs SET name='Partizan Sloga Zemun', cm_name=CASE WHEN COALESCE(cm_name,'')='' OR cm_name LIKE '%Play Out%' THEN 'Partizan Sloga Zemun' ELSE cm_name END, source_name='Partizan Sloga Zemun' WHERE source_club_url LIKE '%/club/237-partizan%' AND (name LIKE '%Play Out%' OR name LIKE 'Partizan Zemun%')")
    con.execute("INSERT OR REPLACE INTO history_meta(key,value,updated_at) VALUES('seed_version',?,CURRENT_TIMESTAMP)",(HISTORY_SEED_VERSION,))

    con.commit()
    con.close()
    return counts

def rsssf_url_for_season(start_year:int)->str:
    end=(start_year+1)%100
    if start_year<=2008:
        suffix=f'{end:02d}'
    else:
        suffix=str(start_year+1)
    return f'{RSSSF_BASE}serv{suffix}.html'

_ROW_RE=re.compile(r'^\s*(\d+)\.\s*(.*?)\s+(\d{1,3})\s+(\d{1,3})\s+(\d{1,3})\s+(\d{1,3})\s+(\d+)\s*-\s*(\d+)\s+(-?\d+(?:\.\d+)?)\b(.*)$')


def parse_table_row(line:str):
    m=_ROW_RE.match(line or '')
    if not m: return None
    return TableRow(int(m.group(1)),m.group(2).strip(),int(m.group(3)),int(m.group(4)),int(m.group(5)),int(m.group(6)),int(m.group(7)),int(m.group(8)),float(m.group(9)),m.group(10).strip())


def _section_lines(text:str,kind:str):
    lines=[x.rstrip() for x in (text or '').splitlines()]
    starts=[]
    for i,line in enumerate(lines):
        clean=re.sub(r'\s+',' ',line.strip().lower())
        if len(clean)>90: continue
        if kind=='super' and re.search(r'\bsuper\s*liga\b',clean):
            starts.append(i)
        if kind=='prva' and (re.search(r'\bprva\s+liga\b',clean) or clean.startswith('second level')):
            starts.append(i)
    for start in starts:
        sample='\n'.join(lines[start:start+140]).lower()
        if 'final table' not in sample: continue
        end=len(lines)
        for j in range(start+1,len(lines)):
            low=re.sub(r'\s+',' ',lines[j].strip().lower())
            if not low: continue
            if kind=='super':
                if (re.match(r'^(second level|prva liga\b|kup srbije\b|serbian cup\b|third level\b)',low)
                    or re.match(r'^####?\s*(kup|prva liga|second level)',low)):
                    end=j; break
            else:
                # Older RSSSF pages often jump from the second level straight
                # into regional headings such as "NIŠKA ZONA" or "BEOGRADSKA
                # ZONA" without a generic "Third Level" marker.  Treat any
                # short zone/third-tier heading as the hard end of Prva liga.
                if (re.match(r'^(third level|srpska liga\b|serbian league\b)',low)
                    or re.search(r'\b(zona|zone)\b',low)
                    or re.match(r'^####?\s*(third level|srpska liga|serbian league)',low)):
                    end=j; break
        return lines[start:end]
    return []


def parse_final_table(text:str,kind:str):
    lines=_section_lines(text,kind)
    if not lines: return []
    final={}
    active=False
    for raw in lines:
        line=raw.strip('\n')
        low=line.strip().lower()
        if 'final table' in low:
            active=True; continue
        if not active: continue
        row=parse_table_row(line)
        if row:
            # In playoff formats later final tables overwrite regular-stage
            # positions; championship/relegation groups together form the final
            # season classification.
            final[row.position]=row
            continue
    return [final[k] for k in sorted(final)]


def _parsed_champion_matches(con,kind,season,rows):
    expected=(SUPER_CHAMPIONS if kind=='super' else PRVA_CHAMPIONS).get(season)
    if not expected: return True
    first=next((r for r in rows if r.position==1),None)
    if not first: return False
    expected_id=_club_id(con,expected); actual_id=_club_id(con,first.club_name)
    if expected_id is not None and actual_id is not None:
        return int(expected_id)==int(actual_id)
    return _norm_match(first.club_name)==_norm_match(expected)


def fetch_season(start_year:int,timeout=25):
    url=rsssf_url_for_season(start_year)
    headers={'User-Agent':'Mozilla/5.0 TouchlineHistory/1.3.0 (+local football archive tool)'}
    r=requests.get(url,headers=headers,timeout=timeout); r.raise_for_status()
    text=BeautifulSoup(r.text,'lxml').get_text('\n')
    return url,text


def sync_modern_history(start_year=2006,end_year=2025):
    """Disabled: completed domestic league history is Excel-only from 1.6.4."""
    return {'disabled':1,'source':EXCEL_LEAGUE_SOURCE,'reason':'Domestic league history is Excel-only'}


def _sync_modern_history_rsssf_legacy(start_year=2006,end_year=2025):
    """Legacy online rebuild retained only as unreachable reference code."""
    seed_history()
    con=connect(); ensure_history_schema(con)
    super_id=con.execute("SELECT id FROM history_competitions WHERE history_key='serbia-top-flight'").fetchone()[0]
    prva_id=con.execute("SELECT id FROM history_competitions WHERE history_key='serbia-second-tier'").fetchone()[0]
    out={'seasons':0,'super_rows':0,'prva_rows':0,'matched':0,'unmatched':0,'errors':[]}
    for year in range(int(start_year),int(end_year)+1):
        season=f'{year}/{str(year+1)[-2:]}'
        try:
            url,text=fetch_season(year)
            parsed_any=False
            for kind,comp_id,key in [('super',super_id,'super_rows'),('prva',prva_id,'prva_rows')]:
                rows=parse_final_table(text,kind)
                if not rows:
                    out['errors'].append(f'{season} {kind}: final table nije pronađena'); continue
                if not _parsed_champion_matches(con,kind,season,rows):
                    actual=next((r.club_name for r in rows if r.position==1),'?')
                    expected=(SUPER_CHAMPIONS if kind=='super' else PRVA_CHAMPIONS).get(season,'?')
                    out['errors'].append(f'{season} {kind}: BLOK ODBAČEN — prvak {actual}, očekivano {expected}')
                    continue
                parsed_any=True
                # Replace only rows created by an earlier online table sync.
                # Verified seeds/highlights are preserved and merged by upsert.
                con.execute("DELETE FROM history_club_seasons WHERE competition_id=? AND season_label=? AND COALESCE(source_note,'') LIKE 'Final table%'",(comp_id,season))
                for row in rows:
                    tail=(row.tail or '').lower()
                    status,cid=_upsert_history(con,comp_id,season,row.club_name,position=row.position,played=row.played,wins=row.wins,draws=row.draws,losses=row.losses,gf=row.gf,ga=row.ga,points=row.points,
                      champion=1 if row.position==1 else 0,runner_up=1 if row.position==2 else 0,promoted=1 if 'promot' in tail else 0,relegated=1 if 'relegat' in tail else 0,
                      source_name='RSSSF',source_url=url,source_note='Final table / final playoff classification')
                    out[key]+=1; out['matched' if cid else 'unmatched']+=1
            if parsed_any: out['seasons']+=1
            con.commit()
        except Exception as e:
            out['errors'].append(f'{season}: {e}')
    con.commit(); con.close(); return out


# ---------------------------------------------------------------------------
# 1.4.0 generic club-history importer
# ---------------------------------------------------------------------------
WIKIPEDIA_API = 'https://en.wikipedia.org/w/api.php'
WIKI_UA = 'TouchlineHistory/1.4.0 (local Serbian football archive; educational data sync)'

def _fold_ascii(v:str)->str:
    s=(str(v or '')).lower().replace('đ','dj')
    s=unicodedata.normalize('NFKD',s)
    return ''.join(ch for ch in s if not unicodedata.combining(ch))


# Reliable title hints are only hints. The importer still searches Wikipedia and
# validates that a page actually contains a season-by-season table before use.
WIKI_PAGE_HINTS = {
    'redstarbelgrade': ['List of Red Star Belgrade seasons'],
    'partizanbelgrade': ['List of FK Partizan seasons'],
    'vojvodinanovisad': ['FK Vojvodina'],
    'radnickinis': ['List of FK Radnički Niš seasons'],
    'radnicki1923': ['FK Radnički 1923'],
    'napredakkrusevac': ['FK Napredak Kruševac'],
    'ofkbeograd': ['OFK Beograd'],
    'cukaricki': ['FK Čukarički'],
    'macvasabac': ['FK Mačva Šabac'],
    'mladostlucani': ['FK Mladost Lučani'],
    'novipazar': ['FK Novi Pazar'],
    'radniksurdulica': ['FK Radnik Surdulica'],
    'zeleznicarpancevo': ['FK Železničar Pančevo'],
    'imtnovibeograd': ['FK IMT'],
    'zemun': ['FK Zemun'],
    'javormatis': ['FK Javor Ivanjica'],
    'ofkvrsac': ['OFK Vršac'],
    'spartaksubotica': ['FK Spartak Subotica'],
    'smederevo1924': ['FK Smederevo 1924'],
    'vozdovac': ['FK Voždovac'],
    'gfkdubocica1923': ['GFK Dubočica'],
    'metalacgornjimilanovac': ['FK Metalac Gornji Milanovac'],
    'loznica': ['FK Loznica'],
    'jedinstvoub': ['FK Jedinstvo Ub'],
    'rfkgraficar': ['RFK Grafičar Beograd'],
    'teleoptik': ['FK Teleoptik'],
    'tscbackatopola': ['FK TSC'],
}


def _wiki_key(v:str)->str:
    return re.sub(r'[^a-z0-9]+','',_fold_ascii(str(v or '')).lower())


def _clean_wiki_text(v:str)->str:
    x=re.sub(r'\[[^\]]*\]','',str(v or ''))
    x=re.sub(r'\s+',' ',x).strip()
    return x


def _season_label(v:str):
    """Normalize Wikipedia season labels to Touchline YYYY/YY."""
    x=_clean_wiki_text(v).replace('–','-').replace('—','-').replace('−','-')
    m=re.search(r'(?<!\d)(19\d{2}|20\d{2})\s*[-/]\s*(\d{2,4})(?!\d)',x)
    if m:
        y=int(m.group(1)); e=m.group(2)
        ey=int(e) if len(e)==4 else (y//100)*100+int(e)
        if ey < y: ey += 100
        return f'{y:04d}/{str(ey)[-2:]}'
    # A few old source tables use a single calendar year for a football season.
    # We do NOT manufacture a second year unless it is clearly a season table.
    return None


def _wiki_int(v):
    x=_clean_wiki_text(v).replace(',','')
    m=re.search(r'-?\d+',x)
    return int(m.group()) if m else None


def _wiki_position(v):
    x=_clean_wiki_text(v)
    m=re.search(r'(?<!\d)(\d{1,2})(?:st|nd|rd|th)?\b',x,re.I)
    return int(m.group(1)) if m else None


def _wiki_rows_from_html(html_text:str):
    """Extract season rows from Wikipedia-style season tables.

    We deliberately key off a season in the first cell and the conventional
    League/P/W/D/L/GF/GA/Pts/Pos order instead of depending on one exact table
    class or translated heading. Pages with no real season table simply return
    no rows and are reported, never guessed.
    """
    soup=BeautifulSoup(html_text or '','lxml')
    best=[]
    for table in soup.find_all('table'):
        rows=[]
        for tr in table.find_all('tr'):
            cells=tr.find_all(['th','td'],recursive=False)
            if len(cells)<2:
                continue
            vals=[]
            for c in cells:
                for sup in c.find_all('sup'):
                    sup.decompose()
                vals.append(_clean_wiki_text(c.get_text(' ',strip=True)))
            season=_season_label(vals[0])
            if not season:
                continue
            league=_clean_wiki_text(vals[1])
            if not league or len(league)>100:
                continue
            rec={'season':season,'league':league,'played':None,'wins':None,'draws':None,'losses':None,
                 'gf':None,'ga':None,'points':None,'position':None,'cup':None,'europe_comp':None,'europe_stage':None}
            # The overwhelming majority of club-season tables use this order.
            if len(vals)>=10:
                stats=[_wiki_int(x) for x in vals[2:9]]
                if len(stats)>=7 and stats[0] is not None:
                    rec.update({'played':stats[0],'wins':stats[1],'draws':stats[2],'losses':stats[3],
                                'gf':stats[4],'ga':stats[5],'points':stats[6]})
                    rec['position']=_wiki_position(vals[9])
                    if len(vals)>10: rec['cup']=vals[10]
                    if len(vals)>11: rec['europe_comp']=vals[11]
                    if len(vals)>12: rec['europe_stage']=vals[12]
            if rec['position'] is None:
                # Fallback for compact tables: position is normally one of the
                # last short cells. Avoid treating match counts as position by
                # searching only after the league cell and preferring ordinals.
                for x in vals[2:]:
                    if re.search(r'\b\d{1,2}(?:st|nd|rd|th)\b',x,re.I):
                        rec['position']=_wiki_position(x); break
            if rec['position'] is not None or rec['played'] is not None:
                rows.append(rec)
        if len(rows)>len(best):
            best=rows
    # Unique season+league rows while preserving order.
    out=[]; seen=set()
    for r in best:
        k=(r['season'],_norm_match(r['league']))
        if k in seen: continue
        seen.add(k); out.append(r)
    return out


def _wiki_search_titles(club_name:str,limit=8):
    queries=[f'"List of {club_name} seasons"',f'{club_name} seasons',club_name]
    found=[]
    for q in queries:
        try:
            r=requests.get(WIKIPEDIA_API,params={'action':'query','list':'search','srsearch':q,'format':'json','srlimit':limit},
                           headers={'User-Agent':WIKI_UA},timeout=12)
            r.raise_for_status()
            for item in r.json().get('query',{}).get('search',[]):
                title=item.get('title')
                if title and title not in found: found.append(title)
        except Exception:
            continue
        if found: break
    return found


def _wiki_page_html(title:str):
    r=requests.get(WIKIPEDIA_API,params={'action':'parse','page':title,'prop':'text','format':'json','formatversion':2},
                   headers={'User-Agent':WIKI_UA},timeout=18)
    r.raise_for_status(); data=r.json()
    if 'error' in data: raise RuntimeError(data['error'].get('info','Wikipedia parse error'))
    return data.get('parse',{}).get('text','')


def _candidate_titles(club_name:str):
    key=_wiki_key(club_name); out=[]; hinted=set(WIKI_PAGE_HINTS.get(key,[]))
    for t in hinted:
        if t not in out: out.append(t)
    for t in _wiki_search_titles(club_name):
        if t not in out: out.append(t)
    # Rank list/season pages first, then ordinary club page candidates. A page
    # with a different numeric identity (e.g. Borac 1926 -> old FK Borac) is not
    # auto-accepted unless explicitly whitelisted above: continuity is a master
    # identity decision, not a fuzzy-history guess.
    folded=_fold_ascii(club_name).lower()
    tokens=[x for x in re.split(r'\W+',folded) if len(x)>2 and x not in ('belgrade','beograd','novi','gornji','milanovac')]
    number_tokens=re.findall(r'\b\d{3,4}\b',folded)
    def relevant(t):
        if t in hinted: return True
        n=_fold_ascii(t).lower()
        if number_tokens and any(num not in n for num in number_tokens): return False
        words=[x for x in tokens if not x.isdigit()]
        return bool(words) and words[0] in n
    out=[t for t in out if relevant(t)]
    def score(t):
        n=_fold_ascii(t).lower(); sc=0
        if n.startswith('list of '): sc+=80
        if ' seasons' in n: sc+=80
        if all(tok in n for tok in tokens[:2]): sc+=35
        if re.match(r'^\d{4}',n): sc-=100
        if ' season' in n and ' seasons' not in n: sc-=40
        return sc
    return sorted(out,key=score,reverse=True)


def _fetch_club_season_table(club_name:str):
    errors=[]
    for title in _candidate_titles(club_name)[:8]:
        try:
            rows=_wiki_rows_from_html(_wiki_page_html(title))
            if len(rows)>=3:
                url='https://en.wikipedia.org/wiki/'+quote(title.replace(' ','_'))
                return {'title':title,'url':url,'rows':rows,'errors':errors}
            errors.append(f'{title}: nema sezonsku tabelu')
        except Exception as e:
            errors.append(f'{title}: {e}')
    return {'title':None,'url':None,'rows':[],'errors':errors}


def _sr_league_name(en_name:str)->str:
    n=_clean_wiki_text(en_name)
    repl=[
        ('Serbian SuperLiga','SuperLiga Srbije'),('Serbian First League','Prva liga Srbije'),
        ('Yugoslav First League','Prva savezna liga Jugoslavije'),('Yugoslav Second League','Druga savezna liga Jugoslavije'),
        ('First League of FR Yugoslavia','Prva liga SR Jugoslavije'),('Second League of FR Yugoslavia','Druga liga SR Jugoslavije'),
        ('First League of Serbia and Montenegro','Prva liga Srbije i Crne Gore'),('Second League of Serbia and Montenegro','Druga liga Srbije i Crne Gore'),
        ('Serbian League Belgrade','Srpska liga Beograd'),('Serbian League Vojvodina','Srpska liga Vojvodina'),
        ('Serbian League East','Srpska liga Istok'),('Serbian League West','Srpska liga Zapad'),
        ('Serbian Republic League','Srpska republička liga'),
    ]
    for a,b in repl:
        if _norm_match(n)==_norm_match(a): return b
    return n


def _dynamic_league_level(league_name:str,season:str):
    n=_norm_match(league_name); y=_season_start_year(season) or 9999
    if 'superliga' in n: return 1
    if n in ('div 1','division 1','1 division','1st division','first division'): return 1
    if n in ('div 2','division 2','2 division','2nd division','second division'): return 2
    if 'first league' in n or 'prva liga' in n:
        return 1 if y<=2005 else 2
    if 'second league' in n or 'druga liga' in n: return 2
    if 'serbian league' in n or 'srpska liga' in n or 'republic league' in n or 'republican league' in n: return 3
    if 'zone' in n or 'zona' in n or 'fourth league' in n: return 4
    return None


def _competition_for_source_league(con,season:str,league_name:str):
    """Resolve a source league to a canonical historical competition.

    Named national first/second tiers are normalized across the state changes;
    lower or unusual competitions get a stable source-derived identity instead
    of being guessed into the wrong national tier.
    """
    y=_season_start_year(season) or 9999; n=_norm_match(league_name)
    canonical=None
    div1 = n in ('div 1','division 1','1 division','1st division','first division') or bool(re.fullmatch(r'div(?:ision)?\s*1',n))
    div2 = n in ('div 2','division 2','2 division','2nd division','second division') or bool(re.fullmatch(r'div(?:ision)?\s*2',n))
    if ('superliga' in n) and y>=2006:
        canonical='serbia-top-flight'
    elif div1 and y>=2006:
        canonical='serbia-top-flight'
    elif div2 and y>=2006:
        canonical='serbia-second-tier'
    elif ('first league' in n or 'prva liga' in n or div1) and y<=1991:
        canonical='yugoslavia-first-1946-1992'
    elif ('first league' in n or 'prva liga' in n or div1) and 1992<=y<=2002:
        canonical='fr-yugoslavia-first-1992-2003'
    elif ('first league' in n or 'prva liga' in n or div1) and 2003<=y<=2005:
        canonical='serbia-montenegro-first-2003-2006'
    elif ('first league' in n or 'prva liga' in n) and y>=2006:
        canonical='serbia-second-tier'
    elif ('second league' in n or 'druga liga' in n or div2) and y<=1991:
        canonical='yugoslavia-second-league'
    elif ('second league' in n or 'druga liga' in n or div2) and 1992<=y<=2002:
        canonical='fr-yugoslavia-second-1992-2003'
    elif ('second league' in n or 'druga liga' in n or div2) and 2003<=y<=2005:
        canonical='serbia-montenegro-second-2003-2006'
    if canonical:
        r=con.execute('SELECT id FROM history_competitions WHERE history_key=?',(canonical,)).fetchone()
        if r: return r[0]

    level=_dynamic_league_level(league_name,season)
    slug=re.sub(r'[^a-z0-9]+','-',_fold_ascii(league_name).lower()).strip('-')[:70] or 'unknown'
    era='modern' if y>=2006 else ('scg' if y>=2003 else ('fry' if y>=1992 else 'yug'))
    key=f'source-{era}-{slug}'
    name_en=_clean_wiki_text(league_name); name_sr=_sr_league_name(name_en)
    sort=60+(level or 8)
    return _competition(con,key,name_sr,name_en,'league',level,None,sort,None,None,None)


def _ensure_extended_domestic_competitions(con):
    _competition(con,'fr-yugoslavia-second-1992-2003','Druga liga SR Jugoslavije','FR Yugoslavia Second League','league',2,None,33,'1992/93','2002/03',None)
    _competition(con,'serbia-montenegro-second-2003-2006','Druga liga Srbije i Crne Gore','Serbia and Montenegro Second League','league',2,None,34,'2003/04','2005/06',None)
    # Widen the existing historical second tier; older versions had a narrow
    # active range because it was created only for Vojvodina's control rows.
    con.execute("UPDATE history_competitions SET active_from='1947/48',active_to='1991/92',name_sr='Druga savezna liga Jugoslavije',name_en='Yugoslav Second League',level=2 WHERE history_key='yugoslavia-second-league'")


def sync_current_club_histories(max_workers=6,completed_through=2025):
    """Disabled: club domestic league history is Excel-only from 1.6.4."""
    return {'disabled':1,'source':EXCEL_LEAGUE_SOURCE,'reason':'Domestic league history is Excel-only'}


def _sync_current_club_histories_wikipedia_legacy(max_workers=6,completed_through=2025):
    """Legacy Wikipedia club-history importer retained only as reference code."""
    seed_history()
    con=connect(); ensure_history_schema(con); _ensure_extended_domestic_competitions(con); con.commit()
    league_rows=con.execute("SELECT id FROM leagues WHERE name IN ('SuperLiga Srbije','Prva liga Srbije') AND COALESCE(section,'senior')='senior'").fetchall()
    lids=[int(r[0]) for r in league_rows]
    if not lids:
        con.close(); return {'clubs':0,'ok':0,'rows':0,'failed':0,'errors':['Nisu pronađene muške lige nivoa 1-2.'],'details':[]}
    marks=','.join('?'*len(lids))
    clubs=con.execute(f"SELECT id,name,cm_name FROM clubs WHERE league_id IN ({marks}) ORDER BY id",lids).fetchall()
    club_data=[{'id':int(r['id']),'name':(r['cm_name'] or r['name']),'display':r['name']} for r in clubs]
    con.close()

    fetched={}; details=[]
    with ThreadPoolExecutor(max_workers=max(1,min(int(max_workers),8))) as ex:
        futs={ex.submit(_fetch_club_season_table,c['name']):c for c in club_data}
        for fut in as_completed(futs):
            c=futs[fut]
            try: fetched[c['id']]=fut.result()
            except Exception as e: fetched[c['id']]={'title':None,'url':None,'rows':[],'errors':[str(e)]}

    con=connect(); ensure_history_schema(con); _ensure_extended_domestic_competitions(con)
    out={'clubs':len(club_data),'ok':0,'rows':0,'failed':0,'errors':[],'details':[]}
    for c in club_data:
        res=fetched.get(c['id']) or {'rows':[],'errors':['bez rezultata']}
        rows=res.get('rows') or []
        if not rows:
            out['failed']+=1
            err='; '.join((res.get('errors') or [])[:2]) or 'nije pronađena sezonska tabela'
            out['errors'].append(f"{c['display']}: {err}")
            out['details'].append({'club':c['display'],'status':'failed','rows':0,'source':res.get('title')})
            continue
        imported=0
        for r in rows:
            sy=_season_start_year(r['season'])
            if sy is None or sy>int(completed_through):
                continue
            comp=_competition_for_source_league(con,r['season'],r['league'])
            level=con.execute('SELECT level FROM history_competitions WHERE id=?',(comp,)).fetchone()['level']
            pos=r.get('position')
            _upsert_history(con,comp,r['season'],c['display'],position=pos,played=r.get('played'),wins=r.get('wins'),draws=r.get('draws'),losses=r.get('losses'),
                            gf=r.get('gf'),ga=r.get('ga'),points=r.get('points'),champion=1 if pos==1 else 0,runner_up=1 if pos==2 else 0,
                            highlight=1 if (pos==1 and level==1) else 0,source_name='Wikipedia',source_url=res.get('url'),
                            source_note='Wikipedia club-season import 1.4.0',fill_only=False,forced_club_id=c['id'])
            imported+=1
        con.commit(); out['ok']+=1; out['rows']+=imported
        out['details'].append({'club':c['display'],'status':'ok','rows':imported,'source':res.get('title')})
    con.execute("INSERT OR REPLACE INTO history_meta(key,value,updated_at) VALUES('all_clubs_last_sync',?,CURRENT_TIMESTAMP)",
                (json.dumps({'ok':out['ok'],'failed':out['failed'],'rows':out['rows']},ensure_ascii=False),))
    con.commit(); con.close(); return out


if __name__=='__main__':
    print(seed_history())
