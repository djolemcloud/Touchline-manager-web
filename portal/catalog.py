from __future__ import annotations

# Central UI/navigation catalog. Add or change cards here instead of patching page code.
HOME_LEVELS = {
    1: dict(title='SuperLiga Srbije', label='NIVO 1', image='/static/branding/super_liga.jpg'),
    2: dict(title='Prva liga Srbije', label='NIVO 2', image='/static/branding/prva_liga.jpg'),
    3: dict(title='Srpske lige', label='NIVO 3', image='/static/branding/srpske_lige.jpg'),
    4: dict(title='Zonske lige', label='NIVO 4', image='/static/branding/zonske_lige.jpg'),
    5: dict(title='Okružne lige', label='NIVO 5', image='/static/branding/okruzne_lige.png'),
    6: dict(title='Opštinske lige', label='NIVO 6', image='/static/branding/medjuopstinske_lige.png'),
}

FSS_HUB = (
    dict(title='Srbija nacionalni A tim i U21 tim', label='NACIONALNE SELEKCIJE', image='/static/branding/nacionalni_timovi_hub.png', route='/fss/national'),
    dict(title='Omladinska i ženska takmičenja', label='FSS · TAKMIČENJA', image='/static/branding/omladinska_zenska_hub.png', route='/fss/development'),
    dict(title='Regionalni savezi', label='FSS · REGIONALNI SAVEZI', image='/static/branding/regionalni_savezi_hub.png', route='/fss/regions'),
)

DEVELOPMENT_HUB = (
    dict(title='Omladinska takmičenja', label='FSS · RAZVOJNA TAKMIČENJA', image='/static/branding/omladinska_takmicenja_hub.png', route='/youth'),
    dict(title='Ženska takmičenja', label='FSS · ŽENSKI FUDBAL', image='/static/branding/zenska_takmicenja_hub.png', route='/women'),
)

REGIONS = {
    'FSB': dict(title='Fudbalski savez Beograda', image='/static/branding/regional_fsb.png', keywords=()),
    'FSV': dict(title='Fudbalski savez Vojvodine', image='/static/branding/regional_fsv.png', keywords=('Vojvodina','Vojvođanska','Vojvodjanska','PFL','novosadska','Sremska','Subotica','Sombor','Pančevo','Pancevo','Zrenjanin')),
    'WEST': dict(title='Region Zapadne Srbije', image='/static/branding/regional_west.png', keywords=('Zapad','Drina','Morava','Raška','Raska','Mačva','Macva','Kolubara','Zlatibor','Moravička','Moravicka','Šumadija','Sumadija','Kragujevac','Rasinska')),
    'EAST': dict(title='Region Istočne Srbije', image='/static/branding/regional_east.png', keywords=('Istok','Centar','Jug','Dunav','Niš','Nis','Pirot','Bor','Zaječar','Zajecar','Toplič','Toplic','Jablani','Pomorav','Braničev','Branicev','Podunav')),
    'KIM': dict(title='Kosovo i Metohija', image='/static/branding/regional_kim.png', keywords=()),
}

YOUTH_IMAGES = {
    'U19': '/static/branding/omladinska_u19_card.png',
    'U17': '/static/branding/kadetska_u17_card.png',
}

WOMEN_SECOND_ORDER = {
    'Druga ženska liga Beograd': 1,
    'Druga ženska liga Zapad': 2,
    'Druga ženska liga Istok': 3,
}

# Level 6 / Međuopštinske lige, SrbijaSport 2026/27.
# Region is explicit to avoid fragile name-based navigation after discovery.
LEVEL6_LEAGUES = {
    # Fudbalski savez Beograda
    'Međuopštinska liga Beograd - grupa "A"': 'FSB',
    'Međuopštinska liga Beograd - grupa "B"': 'FSB',
    'Međuopštinska liga Beograd - grupa "C"': 'FSB',
    'Opštinska liga Obrenovac': 'FSB',
    'Prva opštinska liga Lazarevac': 'FSB',

    # Fudbalski savez Vojvodine
    'Gradska liga Subotica': 'FSV',
    'Druga južnobanatska liga "Zapad"': 'FSV',
    'Druga južnobanatska liga "Istok"': 'FSV',
    'MFL Sombor 1. razred': 'FSV',
    'Opštinska liga Bačka Palanka': 'FSV',
    'Međuopštinska liga Srem - Zapad': 'FSV',
    'Druga novosadska liga': 'FSV',
    '"A" liga FSG Zrenjanin': 'FSV',
    'Međuopštinska liga Srem - Istok': 'FSV',
    'Međuopštinska liga Kikinda-Žitište': 'FSV',
    'Potiska međuopštinska liga': 'FSV',
    'MFL Bačka Topola - Mali Iđoš - Vrbas - Bečej - Srbobran': 'FSV',

    # Fudbalski savez regiona Istočne Srbije
    'Međuopštinska liga FSRO': 'EAST',
    'Međuopštinska liga Jug': 'EAST',
    'Međuopštinska liga “Dunav“': 'EAST',
    'Međuopštinska liga “Mlava”': 'EAST',
    'Međuopštinska liga Sever': 'EAST',
    'Međuopštinska jablanička liga': 'EAST',
    'Gradska liga Kruševac - grupa Rasina': 'EAST',
    'Gradska liga Kruševac - grupa Morava': 'EAST',
    'Jasenička opštinska liga': 'EAST',
    'Druga rasinska okružna liga "Zapad"': 'EAST',
    'Druga rasinska okružna liga "Istok"': 'EAST',
    'Gradska liga Pirot': 'EAST',
    'Opštinska liga Aleksinac': 'EAST',
    'Druga Niška liga': 'EAST',
    'Opštinska liga Velika Plana': 'EAST',
    'Prva gradska liga Smederevo': 'EAST',

    # Fudbalski savez regiona Zapadne Srbije
    'Međuopštinska liga Kolubara - grupa Zapad': 'WEST',
    'Međuopštinska liga Kolubara - grupa Istok': 'WEST',
    'Međuopštinska liga "Pocerina"': 'WEST',
    'Međuopštinska liga “Mačva”': 'WEST',
    'Gradska liga Čačak': 'WEST',
    'Međuopštinska liga “Jadar”': 'WEST',
    'Druga liga Kragujevca': 'WEST',
    'Šumadijska liga "Sever"': 'WEST',
    'Šumadijska liga "Jug"': 'WEST',
    'Opštinska liga Gornji Milanovac': 'WEST',
    'Opštinska liga Lučani': 'WEST',
    'Opštinska liga Požega': 'WEST',
    'Opštinska liga Užice': 'WEST',
}


# Explicit senior-league regional ownership used by regional federation pages.
# Keep this deterministic instead of relying on fuzzy title keywords.
REGION_LEAGUES = {
    # Level 3
    'Srpska liga Beograd': 'FSB',
    'Srpska liga Vojvodina': 'FSV',
    'Srpska liga Zapad': 'WEST',
    'Srpska liga Istok': 'EAST',

    # Level 4
    'Beogradska zona': 'FSB',
    'Vojvođanska liga "Jug"': 'FSV',
    'Vojvođanska liga "Sever"': 'FSV',
    'Vojvođanska liga "Istok"': 'FSV',
    'Zona "Drina"': 'WEST',
    'Zona "Morava"': 'WEST',
    'Zona "Zapad"': 'WEST',
    'Zona "Dunav"': 'EAST',
    'Zona "Centar"': 'EAST',
    'Zona "Istok"': 'EAST',
    'Zona "Jug"': 'EAST',

    # Level 5
    'Prva Beogradska liga - grupa "A"': 'FSB',
    'Prva Beogradska liga - grupa "B"': 'FSB',
    'Prva Beogradska liga - grupa "C"': 'FSB',
    'PFL Pančevo': 'FSV',
    'PFL Sombor': 'FSV',
    'PFL Subotica': 'FSV',
    'PFL Zrenjanin': 'FSV',
    'Prva novosadska liga': 'FSV',
    'Sremska liga': 'FSV',
    'Kolubarska okružna liga': 'WEST',
    'Mačvanska okružna liga': 'WEST',
    'Moravička okružna liga': 'WEST',
    'Prva liga Kragujevca': 'WEST',
    'Raška okružna liga': 'WEST',
    'Zlatiborska okružna liga': 'WEST',
    'Šumadijska okružna liga': 'WEST',
    'Borska okružna liga': 'EAST',
    'Braničevska okružna liga': 'EAST',
    'Jablanička okružna liga - grupa Istok': 'EAST',
    'Jablanička okružna liga - grupa Zapad': 'EAST',
    'Nišavska okružna liga': 'EAST',
    'Pirotska okružna liga': 'EAST',
    'Podunavska okružna liga': 'EAST',
    'Pomoravska okružna liga': 'EAST',
    'Prva Niška liga': 'EAST',
    'Prva rasinska okružna liga': 'EAST',
    'Toplička okružna liga': 'EAST',
    'Zaječarska okružna liga': 'EAST',
}

LEVEL6_REGION_ORDER = ('FSB','FSV','WEST','EAST','KIM')


# ---------------------------------------------------------------------------
# Standings zones (coloured stripe + legend under a league table).
# Key = league id from the database (1 = SuperLiga Srbije, 2 = Prva liga Srbije).
# Each zone: (from_position, to_position, kind, label shown in the legend).
# kind is one of: champion, ucl, uel, uecl, promotion, playoff, barrage, playout, relegation
# Leagues not listed here fall back to the zone notes that come with the
# source table (e.g. SrbijaSport "Play Off", "Baraž", ▲ / ▼).
# ---------------------------------------------------------------------------
STANDINGS_ZONES = {
    # SuperLiga Srbije: after the first phase the table splits into two groups
    1: [(1, 6, 'playoff', 'Plej-of · borba za titulu'), (7, 14, 'playout', 'Plej-aut · borba za opstanak')],
    # Prva liga Srbije: same split, top 6 to the championship group
    2: [(1, 6, 'playoff', 'Plej-of · borba za titulu'), (7, 16, 'playout', 'Plej-aut · borba za opstanak')],
}
