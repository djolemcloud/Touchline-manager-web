# -*- coding: utf-8 -*-
from __future__ import annotations
from functools import lru_cache

import argparse
import csv
import html as html_lib
import io
import json
import re
import sqlite3
import threading
import time
import unicodedata
import webbrowser
import zipfile
from email.parser import BytesParser
from email.policy import default as email_policy_default
from difflib import SequenceMatcher
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urljoin, urlparse, quote_plus

import requests
from bs4 import BeautifulSoup, Tag
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

from .catalog import LEVEL6_LEAGUES, REGION_LEAGUES
from .pyramid_discovery import discover_index as discover_srbijasport_index

BASE_DIR = Path(__file__).resolve().parents[1]
DB_PATH = BASE_DIR / 'data' / 'cm_data_hub.db'
CSS_PATH = BASE_DIR / 'static' / 'style.css'
CREST_DIR = BASE_DIR / 'static' / 'crests'
CREST_DIR.mkdir(parents=True, exist_ok=True)
BRAND_DIR = BASE_DIR / 'static' / 'branding'
BRAND_DIR.mkdir(parents=True, exist_ok=True)
SPONSOR_DIR = BASE_DIR / 'static' / 'sponsors'
SPONSOR_DIR.mkdir(parents=True, exist_ok=True)
STADIUM_DIR = BASE_DIR / 'static' / 'stadiums'
STADIUM_DIR.mkdir(parents=True, exist_ok=True)
REPORT_DIR = BASE_DIR / 'REPORTS'
REPORT_DIR.mkdir(parents=True, exist_ok=True)
HOST = '127.0.0.1'
PORT = 5100


# ---------------------------------------------------------------------------
# UI language / i18n
# Serbian is the default. Language is stored in a browser cookie and kept
# per request/thread so the local threaded server remains safe.
# ---------------------------------------------------------------------------
LANG_CTX = threading.local()
LANG_CTX.code = 'sr'

def set_ui_lang(code):
    LANG_CTX.code = 'en' if str(code).lower() == 'en' else 'sr'

def ui_lang():
    return getattr(LANG_CTX, 'code', 'sr')

# Pairs are written as Serbian -> English. Exact full-text matches are applied
# first; regex rules below handle counters and other dynamic messages.
UI_TEXT = {
    'Srbija': 'Serbia',
    'Srbija / FSS': 'Serbia / FSS',
    'Podaci o fudbalskim ligama Srbije': 'Serbian football leagues data',
    'CENTRALNA BAZA SRBIJE': 'SERBIA CENTRAL DATABASE',
    'Fudbalski savez Srbije · centralni pregled svih nivoa.': 'Football Association of Serbia · central overview of all levels.',
    'GENERALNI SPONZORI SAJTA': 'OFFICIAL SITE SPONSORS',
    'Takmičenja Srbije': 'Serbian competitions',
    'Seniorske lige': 'Senior leagues',
    'Omladinski fudbal': 'Youth football',
    'Omladinski fudbal Srbije': 'Youth football of Serbia',
    'Omladinska liga Srbije': 'Serbian U19 League',
    'Kadetska liga Srbije': 'Serbian U17 League',
    'Omladinska liga Srbije U19 · Kadetska liga Srbije U17': 'Serbian U19 League · Serbian U17 League',
    'Omladinska liga Srbije U19 i Kadetska liga Srbije U17.': 'Serbian U19 League and Serbian U17 League.',
    'SuperLiga Srbije': 'Serbian SuperLiga',
    'Prva liga Srbije': 'Serbian First League',
    'Srpske lige': 'Serbian Leagues',
    'Zonske lige': 'Zone Leagues',
    'Treći rang Srbije': 'Serbian third tier',
    'Četvrti rang Srbije': 'Serbian fourth tier',
    'Izaberi jednu od četiri Srpske lige.': 'Choose one of the four Serbian Leagues.',
    'Izaberi jednu od jedanaest Zonskih liga.': 'Choose one of the eleven Zone Leagues.',
    'Fudbalski savez Srbije': 'Football Association of Serbia',
    'Seniorski tim i Srbija U21': 'Senior team and Serbia U21',
    'Omladinska i ženska takmičenja': 'Youth and women competitions',
    'Regionalni savezi': 'Regional associations',
    'Omladinska takmičenja': 'Youth competitions',
    'Ženska takmičenja': 'Women competitions',
    'Nacionalne selekcije žena': 'Women national teams',
    'Klupska takmičenja žena': 'Women club competitions',
    'Srbija A – žene': 'Serbia Women A',
    'Srbija U19 – žene': 'Serbia Women U19',
    'Super liga Srbije (ž)': 'Serbian Women Super League',
    'Prva liga Srbije (ž)': 'Serbian Women First League',
    'Druga ženska liga': 'Women Second League',
    'Druga ženska liga Beograd': 'Women Second League Belgrade',
    'Druga ženska liga Istok': 'Women Second League East',
    'Druga ženska liga Zapad': 'Women Second League West',
    'Otvori sekciju →': 'Open section →',
    'Otvori takmičenje →': 'Open competition →',
    'Otvori reprezentaciju →': 'Open national team →',
    'FSS · CENTRALNI HUB': 'FAS · CENTRAL HUB',
    'FSS · REPREZENTACIJE': 'FAS · NATIONAL TEAMS',
    'FSS · RAZVOJNA TAKMIČENJA': 'FAS · DEVELOPMENT COMPETITIONS',
    'FSS · ŽENSKI FUDBAL': 'FAS · WOMEN FOOTBALL',
    'FSS · REGIONALNI SAVEZI': 'FAS · REGIONAL ASSOCIATIONS',
    'Nacionalni tim Srbije': 'Serbia national team',
    'NACIONALNE SELEKCIJE': 'NATIONAL TEAMS',
    'REPREZENTACIJA SRBIJE': 'SERBIA NATIONAL TEAM',
    'SENIORSKI TIM': 'SENIOR TEAM',
    'MLADA REPREZENTACIJA': 'U21 NATIONAL TEAM',
    'Otvori reprezentacije →': 'Open national teams →',
    'Otvori tim →': 'Open team →',
    'Ažuriraj obe reprezentacije': 'Update both national teams',
    'Ažuriraj reprezentaciju LIVE': 'Update national team LIVE',
    'CENTRALNI SYNC · SVE LIGE': 'CENTRAL SYNC · ALL LEAGUES',
    'Centralno ažuriranje svih liga Srbije': 'Central update of all Serbian leagues',
    'Ovaj blok automatski radi nad svim ligama koje postoje u bazi. Nove ženske, omladinske i buduće Level 6 lige ulaze automatski čim ih dodamo.': 'This block automatically works across all leagues stored in the database. New women’s, youth and future Level 6 leagues are included automatically as soon as they are added.',
    'Prvo osveži sastave svih liga iz njihovog izvora': 'Refresh all league memberships from their configured source first',
    'Uključi transfere svih TM-mapiranih klubova': 'Include transfers for all TM-mapped clubs',
    'Ažuriraj SVE LIGE LIVE': 'Update ALL LEAGUES LIVE',
    'SrbijaSport: osveži SVE lige': 'SrbijaSport: refresh ALL leagues',
    'Transfermarkt: mapiraj SVE lige': 'Transfermarkt: map ALL leagues',
    'Transfermarkt: ažuriraj SVE transfere': 'Transfermarkt: update ALL transfers',
    'Akcije nisu vezane za fiksne nivoe: čitaju aktuelne lige iz baze u trenutku pokretanja.': 'Actions are not tied to fixed levels: they read the current leagues from the database at runtime.',
    'Podaci o reprezentaciji': 'National team information',
    'PODACI': 'DATA',
    'PODACI O KLUBU': 'CLUB DATA',
    'Podaci o klubu': 'Club information',
    'Podaci o klubu još nisu učitani.': 'Club information has not been loaded yet.',
    'Podaci o stadionu još nisu učitani.': 'Stadium information has not been loaded yet.',
    'STADION': 'STADIUM',
    'Stadion': 'Stadium',
    'Igrači': 'Players',
    'IGRAČI': 'PLAYERS',
    'IGRAČI U BAZI': 'PLAYERS IN DATABASE',
    'Igrači još nisu učitani.': 'Players have not been loaded yet.',
    'Stručni štab': 'Staff',
    'STRUČNI ŠTAB': 'STAFF',
    'Stručni štab još nije učitan.': 'Staff has not been loaded yet.',
    'Klub': 'Club',
    'KLUBOVI': 'CLUBS',
    'LIGE': 'LEAGUES',
    'Broj': 'No.',
    'Ime': 'Name',
    'Igrač': 'Player',
    'Pozicija': 'Position',
    'Datum rođenja': 'Date of birth',
    'God.': 'Age',
    'Godine': 'Age',
    'Državljanstvo': 'Nationality',
    'Tržišna vrednost': 'Market value',
    'Ukupna tržišna vrednost': 'Total market value',
    'Prosečna starost': 'Average age',
    'Stranci': 'Foreign players',
    'Reprezentativci': 'National team players',
    'Ime za igru': 'Game-safe name',
    'Sekcija': 'Section',
    'Imenovan': 'Appointed',
    'Prethodni klub': 'Previous club',
    'Grad': 'City',
    'Grad / država': 'City / country',
    'Adresa': 'Address',
    'Telefon': 'Phone',
    'Web sajt': 'Website',
    'Zvanični naziv': 'Official name',
    'Osnovano': 'Founded',
    'Kapacitet': 'Capacity',
    'Savez': 'Association',
    'Ažurirano': 'Updated',
    'Nedostupno': 'Unavailable',
    'Grb nije dostupan': 'Crest unavailable',
    'Transferi': 'Transfers',
    'TRANSFERI': 'TRANSFERS',
    'AKTUELNA SEZONA': 'CURRENT SEASON',
    'DOLASCI': 'ARRIVALS',
    'Dolasci': 'Arrivals',
    'ODLASCI': 'DEPARTURES',
    'Odlasci': 'Departures',
    'Bilans': 'Balance',
    'Prihod': 'Income',
    'Potrošeno': 'Spent',
    'Naknada': 'Fee',
    'Novi klub': 'New club',
    'Nema evidentiranih dolazaka.': 'No recorded arrivals.',
    'Nema evidentiranih odlazaka.': 'No recorded departures.',
    'Nema klubova.': 'No clubs.',
    'Nema Transfermarkt profila': 'No Transfermarkt profile',
    'Klub još nema Transfermarkt profil. Transferi će biti dostupni čim klub bude mapiran.': 'The club does not have a Transfermarkt profile yet. Transfers will be available once the club is mapped.',
    'Klikni za transfere aktuelne sezone': 'Open current-season transfers',
    'Ažuriraj transfere LIVE': 'Update transfers LIVE',
    'Ažuriraj SVE transfere lige': 'Update ALL league transfers',
    'Ažuriraj SVE transfere nivoa': 'Update ALL level transfers',
    'Ažuriraj SVE seniorske transfere': 'Update ALL senior transfers',
    'Transferi ažurirani': 'Transfers updated',
    'Transfer bilans': 'Transfer balance',
    'Transferi XLSX': 'Transfers XLSX',
    'Ažuriraj podatke + grb LIVE': 'Update data + crest LIVE',
    'Ažuriraj podatke →': 'Update data →',
    'Ažuriraj igrače LIVE': 'Update players LIVE',
    'Ažuriraj stručni štab LIVE': 'Update staff LIVE',
    'Ažuriraj celu ligu': 'Update entire league',
    'Ažuriraj sve mapirane klubove': 'Update all mapped clubs',
    'Ažuriraj SVE omladinske ekipe': 'Update ALL youth teams',
    'Ažuriraj CELU SRBIJU LIVE': 'Update ALL SERBIA LIVE',
    'Ažuriranje cele Srbije': 'Update all Serbia',
    'Ažuriranje radi samo nad mapiranim klubovima.': 'Update runs only on mapped clubs.',
    'CENTRALNI REFRESH · FSS': 'CENTRAL REFRESH · FSS',
    'Centralni refresh obuhvata seniore + U19 + U17. Posebno dugme za transfere ispod odnosi se na seniorske lige.': 'Central refresh includes seniors + U19 + U17. The separate transfers button below applies to senior leagues.',
    'Prvo osveži sastave svih liga': 'Refresh all league memberships first',
    'Prvo osveži sastave obe lige': 'Refresh both league memberships first',
    'Osveži samo sastave liga': 'Refresh league memberships only',
    'Uključi transfere aktuelne sezone': 'Include current-season transfers',
    'Dohvati datum rođenja': 'Fetch date of birth',
    'Dohvati datume rođenja stručnog štaba': 'Fetch staff dates of birth',
    'Dohvati datume rođenja stručnog štaba (sporije)': 'Fetch staff dates of birth (slower)',
    'Otvori': 'Open',
    'Otvori ligu →': 'Open league →',
    'Otvori lige →': 'Open leagues →',
    'Otvori takmičenje →': 'Open competition →',
    'Otvori omladinski deo →': 'Open youth section →',
    'Pronađi klub LIVE': 'Find club LIVE',
    'Sačuvaj povezivanje': 'Save mapping',
    'Povezivanje kluba': 'Club mapping',
    'Ručni URL kluba': 'Manual club URL',
    'Sačuvani HTML kluba': 'Saved club HTML',
    'Sačuvani HTML ekipe': 'Saved team HTML',
    'Sačuvani HTML stručnog štaba': 'Saved staff HTML',
    'Sačuvani Transfermarkt Transfers HTML': 'Saved Transfermarkt Transfers HTML',
    'Uvezi HTML kluba': 'Import club HTML',
    'Uvezi HTML igrača': 'Import players HTML',
    'Uvezi HTML stručnog štaba': 'Import staff HTML',
    'Uvezi Transfers HTML': 'Import Transfers HTML',
    'Igrači XLSX': 'Players XLSX',
    'Stručni štab XLSX': 'Staff XLSX',
    'Export MASTER XLSX': 'Export MASTER XLSX',
    'Export crests ZIP': 'Export crests ZIP',
    'Export SERBIA MASTER XLSX': 'Export SERBIA MASTER XLSX',
    'Export SERBIA crests ZIP': 'Export SERBIA crests ZIP',
    'Export YOUTH MASTER XLSX': 'Export YOUTH MASTER XLSX',
    'MAPPING BACKUP & CONTROL': 'MAPPING BACKUP & CONTROL',
    'Protect verified mappings': 'Protect verified mappings',
    'Sačuvaj kompletan SrbijaSport → Transfermarkt mapping nezavisno od SQLite baze, ili izvezi samo klubove koje još treba proveriti.': 'Save the complete SrbijaSport → Transfermarkt mapping independently of the SQLite database, or export only clubs that still need review.',
    'Backup ZIP sadrži XLSX + CSV + issues CSV + summary JSON. Ne menja bazu.': 'Backup ZIP contains XLSX + CSV + issues CSV + summary JSON. It does not modify the database.',
    'Open Mapping Workbench': 'Open Mapping Workbench',
    'Download Mapping Backup ZIP': 'Download Mapping Backup ZIP',
    'Download Open Mapping XLSX': 'Download Open Mapping XLSX',
    'Refresh memberships': 'Refresh memberships',
    'Resolve Transfermarkt mappings': 'Resolve Transfermarkt mappings',
    'Resolve unresolved/review mappings': 'Resolve unresolved/review mappings',
    'Resolver actions': 'Resolver actions',
    'Retry ALL unresolved/review': 'Retry ALL unresolved/review',
    'Retry cleaned search': 'Retry cleaned search',
    'Clean search name': 'Clean search name',
    'Recheck NO TM YET': 'Recheck NO TM YET',
    'Recheck TM now': 'Recheck TM now',
    'Refresh club mapping LIVE': 'Refresh club mapping LIVE',
    'Accept candidate': 'Accept candidate',
    'Save mapping': 'Save mapping',
    'Mark NO TM YET': 'Mark NO TM YET',
    'Import XLSX mappings': 'Import XLSX mappings',
    'Import filled Mapping/Issues XLSX': 'Import filled Mapping/Issues XLSX',
    'Update clubs from SrbijaSport LIVE': 'Update clubs from SrbijaSport LIVE',
    'Update ZONSKA LEAGUES membership': 'Update ZONE LEAGUES membership',
    '11 liga · očekivano 140 klubova. Ovaj korak povlači samo sastave liga sa SrbijaSporta, bez Transfermarkt mapiranja.': '11 leagues · expected 140 clubs. This step fetches league memberships from SrbijaSport only, without Transfermarkt mapping.',
    'Ove akcije rade samo nad ovim nivoom takmičenja. Centralni Serbia update ostaje odvojen.': 'These actions apply only to this competition level. The central Serbia update remains separate.',
    'Normalni resolver preskače NO TM YET klubove. Njih kasnije proveravaj kroz Mapping Workbench → Recheck NO TM YET. Update preskače sve klubove bez TM URL-a.': 'The normal resolver skips NO TM YET clubs. Recheck them later through Mapping Workbench → Recheck NO TM YET. Update skips all clubs without a TM URL.',
    'Normalni retry preskače NO TM YET klubove. Njih periodično proverava posebno dugme.': 'Normal retry skips NO TM YET clubs. A separate button checks them periodically.',
    'Klubovi bez TM URL-a. NO TM YET = ručno provereno, profil trenutno nije pronađen, ali ostaje otvoren za kasnije.': 'Clubs without a TM URL. NO TM YET = manually checked, profile not currently found, but left open for later.',
    'možeš ga bilo kada recheck-ovati ili nalepiti novi TM URL; tada automatski postaje MATCHED.': 'you can recheck it at any time or paste a new TM URL; it will then automatically become MATCHED.',
    'SrbijaSport defines league membership. Transfermarkt enrichment is separate.': 'SrbijaSport defines league membership. Transfermarkt enrichment is separate.',
    'League source: SrbijaSport': 'League source: SrbijaSport',
    'League source: Transfermarkt': 'League source: Transfermarkt',
    'All active clubs are verified/mapped.': 'All active clubs are verified/mapped.',
    'Club not found': 'Club not found',
    'Country not found': 'Country not found',
    'League not found': 'League not found',
    'National team not found': 'National team not found',
    'Competition level not supported.': 'Competition level not supported.',
    'Error': 'Error',
    'Countries': 'Countries',
    'Home': 'Home',
    '← All Serbia leagues': '← All Serbia leagues',
    'Local prototype v0.18.4 · Srbija': 'Local prototype v0.18.5 · Serbia',
    'Početna': 'Home',
    'Države': 'Countries',
    'Greška': 'Error',
    'Liga nije pronađena': 'League not found',
    'Klub nije pronađen': 'Club not found',
    'Država nije pronađena': 'Country not found',
    'Reprezentacija nije pronađena': 'National team not found',
    'Nivo takmičenja nije podržan.': 'Competition level not supported.',
    'ALATI NIVOA': 'LEVEL TOOLS',
    'grupne akcije': 'group actions',
    'BACKUP I KONTROLA MAPIRANJA': 'MAPPING BACKUP & CONTROL',
    'Zaštiti potvrđena mapiranja': 'Protect verified mappings',
    'Radna tabla mapiranja': 'Mapping Workbench',
    'Akcije resolvera': 'Resolver actions',
    'Očišćen naziv za pretragu': 'Clean search name',
    'Prihvati kandidata': 'Accept candidate',
    'Kandidat': 'Candidate',
    'Ručni TM URL': 'Manual TM URL',
    'Poslednje ažuriranje:': 'Last update:',
    'verifikovano': 'verified',
    'provera': 'review',
    'nerešeno': 'unresolved',
    'bez TM URL-a': 'without TM URL',
    'TM mapirano': 'TM mapped',
    'mapirano': 'mapped',
    'klubova ažurirano': 'clubs updated',
    'dolazaka': 'arrivals',
    'odlazaka': 'departures',
    'Ažurira svih': 'Updates all',
    'seniorskih liga i obe nacionalne omladinske lige': 'senior leagues and both national youth leagues',
    'Ovaj proces može trajati duže.': 'This process may take longer.',
    'Ažuriranje lige završeno:': 'League update completed:',
    'Nivo': 'Level',
    'ažuriran:': 'updated:',
    'mapiranih klubova OK.': 'mapped clubs OK.',
    'Preskočeno bez mappinga': 'Skipped without mapping',
    'provereno': 'checked',
    'sa transferima': 'with transfers',
    'bez transfera': 'without transfers',
    'preskočeno bez TM profila': 'skipped without TM profile',
    'Seniorski transferi:': 'Senior transfers:',
    'Transferi Srbije:': 'Serbia transfers:',
    'Transferi lige:': 'League transfers:',
    'Transferi nivoa': 'Level transfers',
    'Uvezeni transferi': 'Imported transfers',
    'Sastavi omladinskih liga osveženi.': 'Youth league memberships refreshed.',
    'Greške:': 'Errors:',
    'Greška': 'Error',
    'Izvezi mapiranje XLSX': 'Export Mapping XLSX',
    'Izvezi mapiranje CSV': 'Export Mapping CSV',
    'Preuzmi backup mapiranja ZIP': 'Download Mapping Backup ZIP',
    'Preuzmi otvorena mapiranja XLSX': 'Download Open Mapping XLSX',
    'Uvezi XLSX mapiranja': 'Import XLSX mappings',
    'Uvezi popunjen Mapping/Issues XLSX': 'Import filled Mapping/Issues XLSX',
    'Osveži sastave liga': 'Refresh memberships',
    'Reši nerešena/pregled mapiranja': 'Resolve unresolved/review mappings',
    'Ponovi SVE nerešene/provera': 'Retry ALL unresolved/review',
    'Ponovi pretragu očišćenog naziva': 'Retry cleaned search',
    'Ponovo proveri TM sada': 'Recheck TM now',
    'Osveži mapiranje kluba LIVE': 'Refresh club mapping LIVE',
    'Označi NO TM YET': 'Mark NO TM YET',
    'Ažuriraj klubove sa SrbijaSport LIVE': 'Update clubs from SrbijaSport LIVE',
    'Ažuriraj sastave ZONSKIH LIGA': 'Update ZONE LEAGUES membership',
    'FSS · OMLADINSKI FUDBAL': 'FSS · YOUTH FOOTBALL',
    'CENTRALNI Serbia update ostaje odvojen.': 'The central Serbia update remains separate.',
    'Centralni Serbia update ostaje odvojen.': 'The central Serbia update remains separate.',
    'Izvor lige:': 'League source:',
    'IZVOR': 'SOURCE',
    'KLUBOVI AŽURIRANI': 'CLUBS UPDATED',
    'aktivno': 'active',
    'Bilans': 'Balance',
    'Nikad': 'Never',
    'NIKAD': 'NEVER',
    'Srbija U21': 'Serbia U21',
    'Otvori Radnu tablu mapiranja': 'Open Mapping Workbench',
    'Sve aktivne ekipe su verifikovane/mapirane.': 'All active clubs are verified/mapped.',
    'Grupne alatke': 'Batch helpers',
    'Kontrola mapiranja': 'Mapping control',
    'mapping kontrola': 'mapping control',
    'IZVOR TRANSFERMARKT': 'SOURCE TRANSFERMARKT',
    'IZVOR SRBIJASPORT': 'SOURCE SRBIJASPORT',
    'FSS Omladinska sekcija': 'FSS Youth Section',
    'Odaberite XLSX fajl prvo': 'Choose an XLSX file first',
    'Odaberite HTML fajl prvo': 'Choose an HTML file first',
    'Odaberite HTML fajl kluba prvo': 'Choose a club HTML file first',
    'Odaberite HTML fajl kluba/ekipe prvo': 'Choose a club/squad HTML file first',
    'Transfermarkt URL kluba nije podešen': 'Transfermarkt club URL is not configured',
    'Podaci o klubu, stadion i grb ažurirani': 'Club Info, Stadium and Crest updated',
    'Podaci o klubu, stadion i grb uvezeni iz sačuvanog HTML-a': 'Club Info, Stadium and Crest imported from saved HTML',
    'LIVE ažuriranje igrača uvezlo je': 'Live players update imported',
    'LIVE ažuriranje stručnog štaba uvezlo je': 'Live staff update imported',
    'redova': 'rows',
    'Uvezeno': 'Imported',
    'redova igrača iz sačuvanog HTML-a': 'player rows from saved HTML',
    'redova stručnog štaba iz sačuvanog HTML-a': 'staff rows from saved HTML',
    'Broj igrača': 'Number of players',
    'URL grba': 'Crest URL',
    'Uvezi HTML ekipe': 'Import Team HTML',
    'Uvezi HTML stručnog štaba': 'Import Staff HTML',
    'LIGA': 'LEAGUE',
    'Serbia Football Leagues Data Hub': 'Serbia Football Leagues Data Hub',
}

# English strings that already exist in the legacy UI and need a Serbian
# rendering when Serbian is selected. Only presentation strings are included;
# source/provider names and IDs remain untouched.
LEGACY_EN_SR = {
    'Countries':'Države','Home':'Početna','Error':'Greška','Club not found':'Klub nije pronađen',
    'Country not found':'Država nije pronađena','League not found':'Liga nije pronađena',
    'National team not found':'Reprezentacija nije pronađena','Competition level not supported.':'Nivo takmičenja nije podržan.',
    'Open league →':'Otvori ligu →','Open leagues →':'Otvori lige →','Open team →':'Otvori tim →',
    'players':'igrača','Players':'Igrači','PLAYERS':'IGRAČI','staff':'stručni štab','Staff':'Stručni štab','STAFF':'STRUČNI ŠTAB',
    'clubs':'klubova','Club':'Klub','LEAGUES':'LIGE','leagues':'liga','mapped':'mapirano','TM mapped':'TM mapirano',
    'No':'Broj','Name':'Ime','Player':'Igrač','Position':'Pozicija','Date of birth':'Datum rođenja','Age':'Godine',
    'Nationality':'Državljanstvo','Market value':'Tržišna vrednost','Game-safe name':'Ime za igru','Section':'Sekcija',
    'Appointed':'Imenovan','Last update:':'Poslednje ažuriranje:','active':'aktivno','verified':'verifikovano','Never':'Nikad','Balance':'Bilans','League source:':'Izvor lige:','SOURCE':'IZVOR','CLUBS UPDATED':'KLUBOVI AŽURIRANI',
    'review':'provera','unresolved':'nerešeno','without TM URL':'bez TM URL-a','Candidate':'Kandidat',
    'Manual TM URL':'Ručni TM URL','Save mapping':'Sačuvaj povezivanje','Accept candidate':'Prihvati kandidata',
    'Refresh memberships':'Osveži sastave liga','Resolve unresolved/review mappings':'Reši nerešena/pregled mapiranja',
    'Retry ALL unresolved/review':'Ponovi SVE nerešene/provera','Retry cleaned search':'Ponovi pretragu očišćenog naziva',
    'Clean search name':'Očišćen naziv za pretragu','Recheck TM now':'Ponovo proveri TM sada',
    'Refresh club mapping LIVE':'Osveži mapiranje kluba LIVE','Mark NO TM YET':'Označi NO TM YET',
    'Import league HTML':'Uvezi HTML lige','Import SrbijaSport HTML':'Uvezi SrbijaSport HTML',
    'Import Staff HTML':'Uvezi HTML stručnog štaba','Import Team HTML':'Uvezi HTML ekipe','Crest URL':'URL grba','Number of players':'Broj igrača',
    'Saved Transfermarkt league HTML':'Sačuvani Transfermarkt HTML lige',
    'Saved Transfermarkt staff HTML':'Sačuvani Transfermarkt HTML stručnog štaba',
    'Saved Transfermarkt team HTML':'Sačuvani Transfermarkt HTML ekipe',
    'Saved SrbijaSport league HTML':'Sačuvani SrbijaSport HTML lige',
    'Fetch Staff DOB':'Dohvati datum rođenja stručnog štaba','No players imported yet.':'Igrači još nisu uvezeni.',
    'No staff imported yet.':'Stručni štab još nije uvezen.','No clubs.':'Nema klubova.',
    'League':'Liga','Status':'Status','SrbijaSport club':'SrbijaSport klub','TM Match':'TM podudaranje',
    'Batch helpers':'Grupne alatke','Resolver actions':'Akcije resolvera','MAPPING BACKUP & CONTROL':'BACKUP I KONTROLA MAPIRANJA',
    'Protect verified mappings':'Zaštiti potvrđena mapiranja','Download Mapping Backup ZIP':'Preuzmi backup mapiranja ZIP',
    'Download Open Mapping XLSX':'Preuzmi otvorena mapiranja XLSX','Export Mapping XLSX':'Izvezi mapiranje XLSX',
    'Export Mapping CSV':'Izvezi mapiranje CSV','Export OPEN + NO TM YET XLSX':'Izvezi OPEN + NO TM YET XLSX',
    'Import XLSX mappings':'Uvezi XLSX mapiranja','Import filled Mapping/Issues XLSX':'Uvezi popunjen Mapping/Issues XLSX',
    'Update clubs from SrbijaSport LIVE':'Ažuriraj klubove sa SrbijaSport LIVE','All active clubs are verified/mapped.':'Sve aktivne ekipe su verifikovane/mapirane.','Open Mapping Workbench':'Otvori Radnu tablu mapiranja',
    'League source: Transfermarkt':'Izvor lige: Transfermarkt','League source: SrbijaSport':'Izvor lige: SrbijaSport',
    '← All Serbia leagues':'← Sve lige Srbije','MAPPING WORKBENCH':'RADNA TABLA MAPIRANJA',
    'mapping control':'kontrola mapiranja','Recheck NO TM YET':'Ponovo proveri NO TM YET',
    'Export MASTER XLSX':'Izvezi MASTER XLSX','Export crests ZIP':'Izvezi grbove ZIP',
    'Export SERBIA MASTER XLSX':'Izvezi SERBIA MASTER XLSX','Export SERBIA crests ZIP':'Izvezi grbove Srbije ZIP',
    'Export YOUTH MASTER XLSX':'Izvezi YOUTH MASTER XLSX','Export OPEN + NO TM YET XLSX':'Izvezi OPEN + NO TM YET XLSX',
    'FSS Youth Section':'FSS Omladinska sekcija','Choose an XLSX file first':'Odaberite XLSX fajl prvo',
    'Choose an HTML file first':'Odaberite HTML fajl prvo','Choose a club HTML file first':'Odaberite HTML fajl kluba prvo',
    'Choose a club/squad HTML file first':'Odaberite HTML fajl kluba/ekipe prvo',
    'Transfermarkt club URL is not configured':'Transfermarkt URL kluba nije podešen',
    'Club Info, Stadium and Crest updated':'Podaci o klubu, stadion i grb ažurirani',
    'Club Info, Stadium and Crest imported from saved HTML':'Podaci o klubu, stadion i grb uvezeni iz sačuvanog HTML-a',
    'Live players update imported':'LIVE ažuriranje igrača uvezlo je','Live staff update imported':'LIVE ažuriranje stručnog štaba uvezlo je',
    'Imported':'Uvezeno','player rows from saved HTML':'redova igrača iz sačuvanog HTML-a','staff rows from saved HTML':'redova stručnog štaba iz sačuvanog HTML-a',
    'rows':'redova','Done':'Završeno','Errors:':'Greške:','Search warnings:':'Upozorenja pretrage:',
    'Candidate accepted:':'Kandidat prihvaćen:','Mapping saved:':'Mapiranje sačuvano:',
    'Mapping XLSX imported:':'Mapping XLSX uvezen:','mappings saved':'mapiranja sačuvano','blank rows skipped':'praznih redova preskočeno','source rows not found':'izvornih redova nije pronađeno',
    'Update ZONE LEAGUES membership':'Ažuriraj sastave ZONSKIH LIGA','Update ZONSKA LEAGUES membership':'Ažuriraj sastave ZONSKIH LIGA',
    'Serbian football leagues data':'Podaci o fudbalskim ligama Srbije','Mapping Workbench':'Radna tabla mapiranja','Open Mapping Workbench':'Otvori Radnu tablu mapiranja','Serbia':'Srbija','Serbia U21':'Srbija U21',
    'LEVEL 1':'NIVO 1','LEVEL 2':'NIVO 2','LEVEL 3':'NIVO 3','LEVEL 4':'NIVO 4','LEVEL 3 TOOLS':'ALATI NIVOA 3','LEVEL 4 TOOLS':'ALATI NIVOA 4',
    'COUNTRY DATA HUB':'CENTRALNA BAZA DRŽAVE','FSS · YOUTH SECTION':'FSS · OMLADINSKA SEKCIJA',
    'FSS · YOUTH COMPETITION':'FSS · OMLADINSKO TAKMIČENJE','FSS · OMLADINSKI FUDBAL':'FSS · OMLADINSKI FUDBAL',
    'MAPPING BACKUP & CONTROL':'BACKUP I KONTROLA MAPIRANJA','CURRENT SEASON':'AKTUELNA SEZONA',
}

# Generate reverse exact dictionary from SR->EN entries. Entries that are the
# same in both languages are harmless.
UI_TEXT.update({
    'Početna': 'Home', 'Takmičenja': 'Competitions', 'Istorija': 'History', 'Pregled': 'Overview',
    'Klubovi': 'Clubs', 'Tabela': 'Table', 'Transferi': 'Transfers', 'Istorija kluba': 'Club history',
    'Istorija lige': 'League history', 'Plasmani': 'Placings', 'Trofeji': 'Honours', 'Evropa': 'Europe',
    'Najveći uspesi': 'Major achievements', 'Sezona': 'Season', 'Takmičenje': 'Competition',
    'Rang': 'Tier', 'Plasman': 'Position', 'Učinak': 'Record', 'Kup / faza': 'Cup / stage',
    'Najtrofejniji klubovi': 'Most successful clubs', 'Šampioni po sezonama': 'Champions by season',
    'Večna tabela': 'All-time table', 'Titule': 'Titles', 'Vicešampion': 'Runner-up',
    'Vicešampioni': 'Runner-up finishes', 'Međunarodni uspesi': 'International achievements',
    'Istorijski podaci još nisu učitani.': 'Historical data has not been loaded yet.',
    'Istorijski model je spreman; podaci će se dopunjavati iz proverenih izvora.': 'The history model is ready; data will be filled from verified sources.',
    'Fotografija stadiona još nije učitana.': 'Stadium photo has not been loaded yet.',
    'Kapacitet': 'Capacity', 'Osnovan': 'Founded', 'Grad': 'City', 'Zvanični sajt': 'Official website',
    'AKTUELNA SEZONA': 'CURRENT SEASON', 'ISTORIJA KLUBA': 'CLUB HISTORY',
    'ISTORIJA TAKMIČENJA': 'COMPETITION HISTORY', 'OTVORI ISTORIJU': 'OPEN HISTORY',
    'OTVORI TRANSFERE': 'OPEN TRANSFERS', 'sezona u bazi': 'seasons in database',
    'osvojene titule': 'titles won', 'finala / vicešampion': 'finals / runner-up',
    'međunarodni uspesi': 'international achievements', 'Šampion': 'Champion', 'Finalista': 'Finalist',
    'Nema još unetih istorijskih sezona za ovaj klub.': 'No historical seasons have been entered for this club yet.',
    'Nema još unetih istorijskih sezona za ovo takmičenje.': 'No historical seasons have been entered for this competition yet.',
    'Operacija je pokrenuta...': 'Operation started...',
    'Ne zatvaraj ovaj prozor dok se operacija ne završi.': 'Do not close this window until the operation finishes.',
    'Šampion': 'Champion', 'Sezone': 'Seasons', 'Golovi': 'Goals', 'Bodovi': 'Points',
    'SEZONE': 'SEASONS', 'TITULE': 'TITLES', 'VICEŠAMPION': 'RUNNER-UP', 'KUPOVI': 'CUPS', 'MEĐUNARODNO': 'INTERNATIONAL',
    'DANAS': 'TODAY',
    # Core portal labels / full SR-EN separation.
    'Fudbalski podaci i analitika': 'Football data & intelligence',
    'Nacionalnost': 'Nationality', 'Vrednost': 'Value', 'Uloga': 'Role', 'Klub': 'Club',
    'Klubovi ↓': 'Clubs ↓', 'Grbovi ZIP': 'Crests ZIP', 'Stadion': 'Stadium',
    'Transferi još nisu učitani za aktuelnu sezonu.': 'Transfers have not been loaded for the current season yet.',
    'transfer zapisa u sezoni': 'transfer records in season',
    'Sezonski plasmani kroz sva takmičenja, nezavisno od trenutnog ranga kluba.': 'Season-by-season results across all competitions, regardless of the club’s current tier.',
    'Šampioni po sezonama i najtrofejniji klubovi iz zajedničkog History Excel mastera.': 'Champions by season and most successful clubs from the shared History Excel master.',
    'Šampioni po sezonama · najtrofejniji klubovi': 'Champions by season · most successful clubs',
    'Podaci o klubu': 'Club information', 'Podaci o stadionu': 'Stadium information',
    'Aktuelna sezona': 'Current season', 'Istorijski pregled': 'Historical overview',
    # Player positions (Transfermarkt source values -> Serbian in SR mode).
    'Golman': 'Goalkeeper', 'Štoper': 'Centre-Back', 'Levi bek': 'Left-Back', 'Desni bek': 'Right-Back',
    'Defanzivni vezni': 'Defensive Midfield', 'Centralni vezni': 'Central Midfield',
    'Ofanzivni vezni': 'Attacking Midfield', 'Levi vezni': 'Left Midfield', 'Desni vezni': 'Right Midfield',
    'Levo krilo': 'Left Winger', 'Desno krilo': 'Right Winger', 'Centarfor': 'Centre-Forward',
    'Povučeni napadač': 'Second Striker',
    # Staff sections and roles.
    'Uprava': 'Management', 'Medicinski tim': 'Medical department', 'Skauting': 'Scouting',
    'Trener': 'Manager', 'Pomoćni trener': 'Assistant Manager', 'Trener golmana': 'Goalkeeping Coach',
    'Kondicioni trener': 'Fitness coach', 'Glavni analitičar': 'Chief Analyst',
    'Analitičar utakmica': 'Match Analyst', 'Video analitičar': 'Video Analyst',
    'Glavni skaut': 'Chief Scout', 'Skaut': 'Scout', 'Šef skautinga': 'Head of Scouting',
    'Šef omladinskog skautinga': 'Head of Youth Scouting', 'Trener omladinaca': 'Youth Coach',
    'Direktor omladinske škole': 'Director of youth department', 'Šef trenerskog rada akademije': 'Head of Academy Coaching',
    'Menadžer akademije': 'Academy manager', 'Klupski lekar': 'Club Doctor', 'Šef medicinske službe': 'Head of Medical',
    'Fizioterapeut': 'Physiotherapist', 'Ekonom': 'Kit Manager', 'Sportski direktor': 'Sporting Director',
    'Tehnički direktor': 'Technical Director', 'Generalni direktor': 'General Manager',
    'Izvršni direktor': 'Chief Executive Officer', 'Direktor profesionalnog fudbala': 'Managing Director Professional Football',
    'Direktor sporta': 'Managing Director Sport', 'Sportski koordinator': 'Sports coordinator',
    'Predsednik': 'President', 'Počasni predsednik': 'Honorary President', 'Potpredsednik': 'Vice-President',
    'Predsednik uprave': 'Chairman', 'Član uprave': 'Board member', 'Klupski predstavnik': 'Club representative',
    'Generalni sekretar': 'Secretary general', 'Savetnik': 'Advisor',
    # Cities and country/nationality names used by current Serbian squads.
    'Beograd': 'Belgrade',
    'Angola': 'Angola', 'Argentina': 'Argentina', 'Italija': 'Italy', 'Jermenija': 'Armenia',
    'Srbija': 'Serbia', 'Austrija': 'Austria', 'Slovenija': 'Slovenia', 'Azerbejdžan': 'Azerbaijan',
    'Belorusija': 'Belarus', 'Belgija': 'Belgium', 'Alžir': 'Algeria', 'Bosna i Hercegovina': 'Bosnia-Herzegovina',
    'Hrvatska': 'Croatia', 'Nemačka': 'Germany', 'Mađarska': 'Hungary', 'Brazil': 'Brazil',
    'Portugal': 'Portugal', 'Bugarska': 'Bulgaria', 'Burkina Faso': 'Burkina Faso', 'Obala Slonovače': "Cote d'Ivoire",
    'Kamerun': 'Cameroon', 'Nigerija': 'Nigeria', 'Kanada': 'Canada', 'Japan': 'Japan',
    'Centralnoafrička Republika': 'Central African Republic', 'Francuska': 'France', 'Kina': 'China',
    'Kolumbija': 'Colombia', 'Kongo': 'Congo', 'Kurasao': 'Curacao', 'Holandija': 'Netherlands',
    'Kipar': 'Cyprus', 'Češka': 'Czech Republic', 'Engleska': 'England', 'Ekvatorijalna Gvineja': 'Equatorial Guinea',
    'Španija': 'Spain', 'Estonija': 'Estonia', 'DR Kongo': 'DR Congo', 'Mali': 'Mali', 'Senegal': 'Senegal',
    'Gana': 'Ghana', 'Gibraltar': 'Gibraltar', 'Grčka': 'Greece', 'Albanija': 'Albania',
    'Gvineja': 'Guinea', 'Gvineja-Bisao': 'Guinea-Bissau', 'Izrael': 'Israel', 'Jamajka': 'Jamaica',
    'Južna Koreja': 'Korea, South', 'Liberija': 'Liberia', 'Luksemburg': 'Luxembourg', 'Martinik': 'Martinique',
    'Crna Gora': 'Montenegro', 'Maroko': 'Morocco', 'Surinam': 'Suriname', 'Niger': 'Niger',
    'Severna Makedonija': 'North Macedonia', 'Peru': 'Peru', 'Rumunija': 'Romania', 'Rusija': 'Russia',
    'Južna Afrika': 'South Africa', 'Švajcarska': 'Switzerland', 'Sjedinjene Američke Države': 'United States',
    'Norveška': 'Norway', 'Švedska': 'Sweden', 'Gambija': 'The Gambia', 'Togo': 'Togo',
    'Ukrajina': 'Ukraine', 'Uzbekistan': 'Uzbekistan', 'Australija': 'Australia',
    # Main competition names used throughout the public portal.
    'SuperLiga Srbije': 'Serbian SuperLiga', 'Prva liga Srbije': 'Serbian First League',
    'Srpska liga Beograd': 'Serbian League Belgrade', 'Srpska liga Vojvodina': 'Serbian League Vojvodina',
    'Srpska liga Zapad': 'Serbian League West', 'Srpska liga Istok': 'Serbian League East',
    'Omladinska liga Srbije': 'Serbian U19 League', 'Kadetska liga Srbije': 'Serbian U17 League',
    'Super liga Srbije (ž)': "Serbian Women's Super League", 'Prva liga Srbije (ž)': "Serbian Women's First League",
    'Druga ženska liga Beograd': "Serbian Women's Second League Belgrade",
    'Druga ženska liga Zapad': "Serbian Women's Second League West", 'Druga ženska liga Istok': "Serbian Women's Second League East",
    # Portal navigation / category labels.
    'Nacionalna takmičenja': 'National competitions', 'Srbija nacionalni A tim i U21 tim': 'Serbia senior national team and U21',
    'NACIONALNE SELEKCIJE': 'NATIONAL TEAMS', 'NACIONALNA SELEKCIJA': 'NATIONAL TEAM',
    'FSS · TAKMIČENJA': 'FSS · COMPETITIONS', 'Omladinska i ženska takmičenja': 'Youth and women competitions',
    'FSS · RAZVOJNA TAKMIČENJA': 'FAS · DEVELOPMENT COMPETITIONS', 'FSS · ŽENSKI FUDBAL': "FAS · WOMEN'S FOOTBALL",
    'FSS · REGIONALNI SAVEZI': 'FAS · REGIONAL ASSOCIATIONS', 'Omladinska takmičenja': 'Youth competitions',
    'Ženska takmičenja': 'Women competitions', 'Nacionalne selekcije žena': "Women's national teams",
    'Klupska takmičenja žena': "Women's club competitions", 'NACIONALNE SELEKCIJE ŽENA': "WOMEN'S NATIONAL TEAMS",
    'DRUGA ŽENSKA LIGA': "WOMEN'S SECOND LEAGUE", 'Druga ženska liga': "Women's Second League",
    'regionalne lige': 'regional leagues', 'Regionalni savezi': 'Regional associations',
    'REGIONALNI SAVEZ': 'REGIONAL ASSOCIATION', 'Fudbalski savez Beograda': 'Belgrade Football Association',
    'Fudbalski savez Vojvodine': 'Vojvodina Football Association', 'Region Zapadne Srbije': 'Western Serbia Region',
    'Region Istočne Srbije': 'Eastern Serbia Region', 'Kosovo i Metohija': 'Kosovo and Metohija',
    'Srpske lige': 'Serbian Leagues', 'Zonske lige': 'Zone Leagues', 'Okružne lige': 'District Leagues',
    'Opštinske lige': 'Municipal Leagues', 'Sve lige šestog nivoa na jednom mestu — međuopštinske, opštinske i druge opštinske lige.': 'All sixth-tier leagues in one place — intermunicipal, municipal and secondary municipal leagues.',
    'NIVO 3 · SRPSKA LIGA': 'LEVEL 3 · SERBIAN LEAGUE', 'NIVO 4 · ZONSKE LIGE': 'LEVEL 4 · ZONE LEAGUES',
    'NIVO 5 · OKRUŽNE LIGE': 'LEVEL 5 · DISTRICT LEAGUES',
    'NIVO 6 · OPŠTINSKE / MEĐUOPŠTINSKE LIGE': 'LEVEL 6 · MUNICIPAL / INTERMUNICIPAL LEAGUES',
    'Za ovaj nivo trenutno nema evidentiranih liga u bazi.': 'There are currently no recorded leagues at this level.',
    'Za ovaj regionalni savez trenutno nema evidentiranih liga.': 'There are currently no recorded leagues for this regional association.',
    'Za ovaj region trenutno nema evidentiranih liga šestog nivoa.': 'There are currently no recorded sixth-tier leagues for this region.',
    'Tabela još nije povučena. Osvežava se iz Superadmin Sync Centra.': 'The table has not been loaded yet. Refresh it from the Superadmin Sync Center.',
    'Tabela još nije povučena. Osvežava se iz Superadmin centra za sinhronizaciju.': 'The table has not been loaded yet. Refresh it from the Superadmin Sync Center.',
    'SUPERADMIN · CENTAR SINHRONIZACIJE': 'SUPERADMIN · SYNC CENTER',
    'Touchline Centar sinhronizacije': 'Touchline Sync Center',
    'Stručni štab XLSX': 'Staff XLSX', 'DOB stručnog štaba': 'Staff DOB',
    'Tabela': 'Table', 'Tim': 'Team', 'osveženo': 'updated',
    'LIGAŠKE TITULE': 'LEAGUE TITLES', 'ligaških titula': 'league titles', 'kupova': 'cups',
    'Liga': 'League', 'Kup': 'Cup', 'Evropa / svet': 'Europe / world',
    'Reprezentacije · omladinska · ženska · regionalni savezi': 'National teams · youth · women · regional associations',
    # Superadmin surface.
    'seniorski': 'senior', 'žene': 'women', 'omladinski': 'youth',
    'SUPERADMIN · CENTAR SINHRONIZACIJE': 'SUPERADMIN · SYNC CENTER', 'Touchline Centar sinhronizacije': 'Touchline Sync Center',
    'DOB stručnog štaba': 'Staff DOB', 'Stručni štab XLSX': 'Staff XLSX',
    'GLOBAL obuhvata sve lige. Svi ligaški blokovi koriste isti prioritet i isti numerisani redosled 01–07. Generalna noćna sinhronizacija pokreće kompletan ciklus jednim klikom.': 'GLOBAL covers all leagues. Every league block uses the same priority and the same numbered 01–07 sequence. The general nightly sync runs the complete cycle with one click.',
    'GENERALNA SINHRONIZACIJA': 'GENERAL SYNC', 'Generalno čišćenje i update': 'General cleanup and update',
    'Jedno dugme za kompletno noćno osvežavanje. Redosled: 01 sastavi → 02 TM mapping → 03 igrači/staff → 04 transferi → 05 grbovi → 06 informacije o klubovima → 07 tabele → reprezentacije.': 'One button for the complete nightly refresh. Sequence: 01 memberships → 02 TM mapping → 03 players/staff → 04 transfers → 05 crests → 06 club information → 07 tables → national teams.',
    'GENERALNO ČIŠĆENJE I UPDATE · NOĆNA SINHRONIZACIJA · 01→07': 'GENERAL CLEANUP AND UPDATE · NIGHTLY SYNC · 01→07',
    'GLOBAL · SVE LIGE': 'GLOBAL · ALL LEAGUES', 'Centralne akcije · isti redosled 01–07': 'Central actions · same 01–07 sequence',
    '01 · SASTAVI SVIH LIGA': '01 · ALL LEAGUE MEMBERSHIPS', '02 · TM MAPPING SVIH LIGA': '02 · TM MAPPING ALL LEAGUES',
    '03 · IGRAČI + STAFF': '03 · PLAYERS + STAFF', '04 · TRANSFERI': '04 · TRANSFERS', '05 · GRBOVI': '05 · CRESTS',
    '06 · INFORMACIJE O KLUBOVIMA': '06 · CLUB INFORMATION', '07 · TABELE': '07 · TABLES',
    'Reprezentacije · rosteri, staff i tabele': 'National teams · rosters, staff and tables',
    'Sve operacije reprezentacija su na jednom mestu. Za reprezentacije su primenljive 01 (roster/staff) i 07 (tabela); klupske operacije 02–06 se ne primenjuju. Srbija A je automatski vezana za UEFA Nations League A · Grupa A2; U21 trenutno nema ligašku tabelu.': 'All national-team operations are in one place. National teams use 01 (roster/staff) and 07 (table); club operations 02–06 do not apply. Serbia A is automatically linked to UEFA Nations League A · Group A2; U21 currently has no league table.',
    '01 · REPREZENTACIJE · ROSTERI + STAFF': '01 · NATIONAL TEAMS · ROSTERS + STAFF',
    '07 · REPREZENTACIJE · TABELE': '07 · NATIONAL TEAMS · TABLES', 'Grupa': 'Group',
    'Tabela nije konfigurisana': 'Table is not configured', 'Napredni override izvora tabele': 'Advanced table-source override',
    'Međunarodne prijateljske utakmice': 'International Friendlies', 'UEFA Liga nacija A': 'UEFA Nations League A',
    'Sačuvaj / promeni izvor': 'Save / change source', 'TOP TIER · NIVO 1–2': 'TOP TIER · LEVEL 1–2',
    'SuperLiga + Prva liga Srbije': 'Serbian SuperLiga + Serbian First League',
    '01 · SASTAVI LIGA': '01 · LEAGUE MEMBERSHIPS', '02 · TM MAPPING': '02 · TM MAPPING',
    'NIŽE SENIORSKE · NIVO 3–6': 'LOWER SENIOR · LEVELS 3–6',
    'Srpske · Zonske · Okružne · Međuopštinske': 'Serbian · Zone · District · Intermunicipal leagues',
    'POSEBNO · NIVO 6 · OTKRIJ + POVUCI LIGE': 'SPECIAL · LEVEL 6 · DISCOVER + FETCH LEAGUES',
    'OMLADINSKA TAKMIČENJA': 'YOUTH COMPETITIONS', 'ŽENSKI FUDBAL': "WOMEN'S FOOTBALL",
    'Ženske lige i selekcije': "Women's leagues and national teams", '03 · IGRAČICE + STAFF': '03 · PLAYERS + STAFF',
    'POSEBNO · FSS ROSTERI SUPER + PRVA': 'SPECIAL · FSS ROSTERS SUPER + FIRST',
    'NAPREDNE KONTROLE': 'ADVANCED CONTROLS', 'Po nivou / po ligi / mapping': 'By level / by league / mapping',
    'Sastavi liga': 'League memberships', 'Igrači + staff': 'Players + staff', 'Pokreni nivo': 'Run level',
    'Sastav lige': 'League membership', 'Pokreni ligu': 'Run league', 'Status svih liga': 'All-league status',
    'Sekcija': 'Section', 'Lige': 'Leagues', 'LIGE': 'LEAGUES', 'KLUBOVI': 'CLUBS', 'IGRAČI': 'PLAYERS',
    'STRUČNI ŠTAB': 'STAFF',
})

SR_EN = dict(UI_TEXT)
EN_SR = {en: sr for sr, en in SR_EN.items()}
EN_SR.update(LEGACY_EN_SR)
EN_SR.update({
    'Ove akcije rade samo nad ovim nivoom takmičenja. Centralni Serbia update ostaje odvojen.': 'Ove akcije rade samo nad ovim nivoom takmičenja. Centralno ažuriranje Srbije ostaje odvojeno.',
    'Normalni resolver preskače NO TM YET klubove. Njih kasnije proveravaj kroz Mapping Workbench → Recheck NO TM YET. Update preskače sve klubove bez TM URL-a.': 'Normalni resolver preskače NO TM YET klubove. Njih kasnije proveravaj kroz Radnu tablu mapiranja → Ponovo proveri NO TM YET. Ažuriranje preskače sve klubove bez TM URL-a.',
    'Normalni retry preskače NO TM YET klubove. Njih periodično proverava posebno dugme.': 'Normalni ponovni pokušaj preskače NO TM YET klubove. Njih periodično proverava posebno dugme.',
    'NO TM YET: možeš ga bilo kada recheck-ovati ili nalepiti novi TM URL; tada automatski postaje MATCHED.': 'NO TM YET: možeš ga bilo kada ponovo proveriti ili nalepiti novi TM URL; tada automatski postaje MATCHED.',
    'Sačuvaj kompletan SrbijaSport → Transfermarkt mapping nezavisno od SQLite baze, ili izvezi samo klubove koje još treba proveriti.': 'Sačuvaj kompletno SrbijaSport → Transfermarkt mapiranje nezavisno od SQLite baze, ili izvezi samo klubove koje još treba proveriti.',
    # Source aliases / spelling variants.
    'Belgrad': 'Beograd', 'Goal keeper coach': 'Trener golmana', 'Conditioning Coach': 'Kondicioni trener',
    'Coaching Staff': 'Stručni štab',
    'Striker': 'Napadač',
})
# EN source aliases that should normalize in EN mode as well.
SR_EN.update({'Belgrad':'Belgrade','Goal keeper coach':'Goalkeeping Coach','Conditioning Coach':'Fitness coach','Napadač':'Striker'})
_EXTRA_HISTORY_UI={'Nazad':'Back','Državna prvenstva':'National championships','DRŽAVNA PRVENSTVA':'NATIONAL CHAMPIONSHIPS','Nacionalni kupovi':'National cups','NACIONALNI KUPOVI':'NATIONAL CUPS','UKUPNO':'TOTAL','Prvak':'Champion','Osvajač':'Winner'}
SR_EN.update(_EXTRA_HISTORY_UI)
EN_SR.update({v:k for k,v in _EXTRA_HISTORY_UI.items()})
# 2.0 UI refresh labels (sr -> en only, so English source data is never rewritten in SR mode)
SR_EN.update({'Golmani':'Goalkeepers','Odbrana':'Defence','Vezni red':'Midfield','Napad':'Attack','Ostali':'Other','Rođen':'Born','Meni':'Menu','Stadion':'Stadium','Kapacitet':'Capacity'})


def _translate_dynamic(text, lang):
    s=text
    if lang=='sr':
        # Dynamic English counters/messages from legacy pages.
        reps=[
            (r'\b(\d+) leagues\b', r'\1 liga'), (r'\b(\d+) league\b', r'\1 liga'),
            (r'\b(\d+) clubs\b', r'\1 klubova'), (r'\b(\d+) players\b', r'\1 igrača'),
            (r'\b(\d+) staff\b', r'\1 članova stručnog štaba'), (r'\b(\d+) transfers\b', r'\1 transfera'),
            (r'\b(\d+)/(\d+) TM mapped\b', r'\1/\2 TM mapirano'),
            (r'\bLast update:\s*', 'Poslednje ažuriranje: '),
            (r'\bLEVEL (\d+) TOOLS\b', r'ALATI NIVOA \1'),
            (r'\bgroup actions\b', 'grupne akcije'),
            (r'\b(\d+) arrivals\b', r'\1 dolazaka'), (r'\b(\d+) departures\b', r'\1 odlazaka'),
            (r'\b(\d+) cups\b', r'\1 kupova'),
            (r'\bGroup ([A-Z0-9-]+)\b', r'Grupa \1'),
            (r'GENERALNO ČIŠĆENJE I UPDATE', 'GENERALNO ČIŠĆENJE I AŽURIRANJE'),
            (r'02 · TM MAPPING', '02 · TM MAPIRANJE'),
            (r'03 · IGRAČI \+ STAFF', '03 · IGRAČI + STRUČNI ŠTAB'),
            (r'03 · IGRAČICE \+ STAFF', '03 · IGRAČICE + STRUČNI ŠTAB'),
            (r'ROSTERI \+ STAFF', 'SASTAVI + STRUČNI ŠTAB'),
            (r'\bupdate\b', 'ažuriranje'), (r'\bstaff\b', 'stručni štab'),
            (r'\brosteri\b', 'sastavi'), (r'\broster\b', 'sastav'),
            (r'\bmapping\b', 'mapiranje'), (r'\bsync\b', 'sinhronizacija'),
            (r'\bDATA OPERATIONS\b', 'OPERACIJE NAD PODACIMA'),
            (r'\bRuntime:\s*', 'Okruženje: '), (r'\bbuild\b', 'verzija'),
            (r'\b(\d+) clubs updated\b', r'\1 klubova ažurirano'),
            (r'\bLEVEL (\d+)\b', r'NIVO \1'),
            (r'\bSOURCE\b', 'IZVOR'),
            (r'\bLeague source:\s*', 'Izvor lige: '),
            (r'\bOpen Mapping Workbench\b', 'Otvori Radnu tablu mapiranja'),
            (r'\bAll active clubs are verified/mapped\.', 'Sve aktivne ekipe su verifikovane/mapirane.'),
            (r'\bNever\b', 'Nikad'),
            (r'\bMAPPING WORKBENCH\b', 'RADNA TABLA MAPIRANJA'),
            (r'\bmapping control\b', 'kontrola mapiranja'),
            (r'\bRecheck NO TM YET\b', 'Ponovo proveri NO TM YET'),
            (r'← All Serbia leagues', '← Sve lige Srbije'),
            (r'\bchecked\b', 'provereno'), (r'\bNew matched\b', 'novo mapirano'),
            (r'\bstill NO TM YET\b', 'i dalje NO TM YET'), (r'\bSearch warnings:\b', 'Upozorenja pretrage:'),
            (r'\bmatched\b', 'mapirano'), (r'\bblank rows skipped\b', 'praznih redova preskočeno'),
            (r'\bsource rows not found\b', 'izvornih redova nije pronađeno'),
            (r'\brows\b', 'redova'),
        ]
    else:
        reps=[
            (r'\b(\d+) liga\b', r'\1 leagues'), (r'\b(\d+) lige\b', r'\1 leagues'), (r'\b(\d+) klubova\b', r'\1 clubs'),
            (r'\b(\d+) igrača\b', r'\1 players'), (r'\b(\d+) članova stručnog štaba\b', r'\1 staff'),
            (r'\b(\d+) transfera\b', r'\1 transfers'),
            (r'\b(\d+)/(\d+) TM mapirano\b', r'\1/\2 TM mapped'),
            (r'\bPoslednje ažuriranje:\s*', 'Last update: '),
            (r'\bklubova ažurirano\b', 'clubs updated'),
            (r'\bstručni štab\b', 'staff'),
            (r'\b(\d+) dolazaka\b', r'\1 arrivals'), (r'\b(\d+) odlazaka\b', r'\1 departures'),
            (r'\b(\d+) kupova\b', r'\1 cups'),
            (r'\bGrupa ([A-Z0-9-]+)\b', r'Group \1'),
            (r'\bažuriranje\b', 'update'), (r'\bmapiranje\b', 'mapping'),
            (r'\bLEVEL (\d+) TOOLS\b', r'LEVEL \1 TOOLS'),
            (r'\bALATI NIVOA (\d+)\b', r'LEVEL \1 TOOLS'),
            (r'\b(\d+) transferi\b', r'\1 transfers'),
            (r'\bLEVEL (\d+)\b', r'LEVEL \1'),
            (r'\bNIVO (\d+)\b', r'LEVEL \1'),
            (r'\bKLUBOVI AŽURIRANI\b', 'CLUBS UPDATED'),
            (r'\baktivno\b', 'active'),
            (r'\bBilans\b', 'Balance'),
            (r'\bNikad\b', 'Never'),
            (r'\bSrbija U21\b', 'Serbia U21'),
            (r'\bFSS · OMLADINSKI FUDBAL\b', 'FSS · YOUTH FOOTBALL'),
            (r'\bprovereno\b', 'checked'), (r'\bnovo mapirano\b', 'new matched'),
            (r'\bi dalje NO TM YET\b', 'still NO TM YET'),
            (r'\bGreške:\b', 'Errors:'), (r'\bUpozorenja pretrage:\b', 'Search warnings:'),
            (r'\bredova\b', 'rows'),
            (r'\btransferi\b', 'transfers'),
            (r'\bTransferi\b', 'Transfers'),
        ]
    for pat,repl in reps:
        s=re.sub(pat,repl,s,flags=re.I)
    return s


def _translation_bundle(lang):
    table=EN_SR if lang=='sr' else SR_EN
    lower={str(k).lower():v for k,v in table.items()}
    pairs=[(str(src),str(dst)) for src,dst in sorted(table.items(), key=lambda kv: len(str(kv[0])), reverse=True) if len(str(src))>=8]
    embedded_map={src.lower():dst for src,dst in pairs}
    # One compiled alternation per language is dramatically faster than
    # compiling one regex per dictionary entry for every rendered text node.
    if pairs:
        alt='|'.join(re.escape(src) for src,_ in pairs)
        embedded_re=re.compile(r'(?<!\w)(?:'+alt+r')(?!\w)',re.I)
    else:
        embedded_re=None
    return table,lower,embedded_map,embedded_re

_TRANSLATION_BUNDLES={lang:_translation_bundle(lang) for lang in ('sr','en')}

@lru_cache(maxsize=16384)
def _tr_core_cached(core,lang):
    table,lower_map,embedded_map,embedded_re=_TRANSLATION_BUNDLES[lang]
    if core in table:
        return str(table[core])
    hit=lower_map.get(core.lower())
    if hit is not None:
        return str(hit).upper() if core.isupper() else str(hit)
    if ' / ' in core:
        parts=[p.strip() for p in core.split(' / ')]
        translated=[]; changed=False
        for part in parts:
            phit=table.get(part)
            if phit is None: phit=lower_map.get(part.lower())
            if phit is not None:
                translated.append(str(phit)); changed=True
            else:
                translated.append(part)
        if changed:
            return ' / '.join(translated)
        return _translate_dynamic(core,lang)
    # Embedded phrase replacement used to rebuild the maps and sort the full
    # translation dictionary for every HTML text node.  History pages can have
    # hundreds of nodes, so cache the translated core and reuse precomputed
    # bundles instead.  This removes the multi-second/minute UI stall.
    out=core
    if embedded_re is not None:
        out=embedded_re.sub(lambda m: str(embedded_map.get(m.group(0).lower(),m.group(0))),out)
    return _translate_dynamic(out,lang)

def tr_text(text, lang=None):
    """Translate one visible UI text node, preserving surrounding whitespace."""
    lang=lang or ui_lang()
    if text is None: return text
    raw=str(text)
    if not raw.strip(): return raw
    lead=raw[:len(raw)-len(raw.lstrip())]; trail=raw[len(raw.rstrip()):]
    return lead+_tr_core_cached(raw.strip(),lang)+trail


def localize_html(fragment, lang=None):
    """Translate rendered HTML text nodes + selected accessibility attributes."""
    lang=lang or ui_lang()
    try:
        soup=BeautifulSoup(fragment,'html.parser')
        for node in list(soup.find_all(string=True)):
            if node.parent and node.parent.name in ('script','style'):
                continue
            new=tr_text(str(node),lang)
            if new!=str(node): node.replace_with(new)
        for tag in soup.find_all(True):
            for attr in ('title','placeholder','aria-label'):
                if tag.has_attr(attr): tag[attr]=tr_text(tag.get(attr),lang)
        return str(soup)
    except Exception:
        return fragment


def display_league_name(name, lang=None):
    """Translate league/competition display names without touching stored source names."""
    lang=lang or ui_lang()
    raw=str(name or '')
    if lang!='en': return raw
    exact=tr_text(raw,'en')
    if exact!=raw: return exact
    out=raw
    prefixes={
        'Borska okružna liga':'Bor District League','Braničevska okružna liga':'Braničevo District League',
        'Jablanička okružna liga':'Jablanica District League','Kolubarska okružna liga':'Kolubara District League',
        'Mačvanska okružna liga':'Mačva District League','Moravička okružna liga':'Moravica District League',
        'Nišavska okružna liga':'Nišava District League','Pirotska okružna liga':'Pirot District League',
        'Podunavska okružna liga':'Podunavlje District League','Pomoravska okružna liga':'Pomoravlje District League',
        'Raška okružna liga':'Raška District League','Toplička okružna liga':'Toplica District League',
        'Zaječarska okružna liga':'Zaječar District League','Zlatiborska okružna liga':'Zlatibor District League',
        'Šumadijska okružna liga':'Šumadija District League','Sremska liga':'Srem League',
        'Beogradska zona':'Belgrade Zone League','Prva Beogradska liga':'Belgrade First League',
        'Prva Niška liga':'Niš First League','Prva liga Kragujevca':'Kragujevac First League',
        'Prva novosadska liga':'Novi Sad First League','Prva rasinska okružna liga':'Rasina First District League',
        'Druga Niška liga':'Niš Second League','Druga južnobanatska liga':'South Banat Second League',
        'Druga liga Kragujevca':'Kragujevac Second League','Druga novosadska liga':'Novi Sad Second League',
        'Druga rasinska okružna liga':'Rasina Second District League','Potiska međuopštinska liga':'Tisa Intermunicipal League',
        'Međuopštinska jablanička liga':'Jablanica Intermunicipal League','Jasenička opštinska liga':'Jasenica Municipal League',
        'Prva gradska liga Smederevo':'Smederevo First City League','Prva opštinska liga Lazarevac':'Lazarevac First Municipal League',
    }
    for src,dst in sorted(prefixes.items(),key=lambda kv:len(kv[0]),reverse=True):
        if out.startswith(src): out=dst+out[len(src):]; break
    for src,dst in [
        ('Vojvođanska liga','Vojvodina League'),('Srpska liga','Serbian League'),
        ('Međuopštinska liga','Intermunicipal League'),('opštinska liga','Municipal League'),
        ('Gradska liga','City League'),('Šumadijska liga','Šumadija League'),
        (' liga',' League'),('Zona ','Zone League '),('grupa','Group'),('Istok','East'),('Zapad','West'),('Sever','North'),('Jug','South')]:
        out=out.replace(src,dst)
    return re.sub(r'\s+',' ',out).strip()

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36',
    'Accept-Language': 'en-US,en;q=0.9,sr;q=0.8',
    'Referer': 'https://www.transfermarkt.com/',
}

SECTION_ALIASES = {
    'coaching staff': 'Coaching Staff', 'coaching': 'Coaching Staff', 'trainerstab': 'Coaching Staff',
    'management': 'Management', 'vorstand': 'Management',
    'scouting': 'Scouting', 'scouts': 'Scouting',
    'medical department': 'Medical department', 'medical staff': 'Medical department',
    'medical': 'Medical department', 'medizinische abteilung': 'Medical department',
}

MONTHS = {'jan':1,'feb':2,'mar':3,'apr':4,'may':5,'jun':6,'jul':7,'aug':8,'sep':9,'oct':10,'nov':11,'dec':12}
MONTH_RE = r'(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)'
DATE_PATTERNS = [
    re.compile(rf'\b{MONTH_RE}\s+\d{{1,2}},\s+\d{{4}}\b', re.I),
    re.compile(r'\b\d{1,2}/\d{1,2}/\d{4}\b'),
    re.compile(r'\b\d{1,2}\.\d{1,2}\.\d{4}\b'),
    re.compile(r'\b\d{4}-\d{2}-\d{2}\b'),
]
AGE_RE = re.compile(r'\((\d{1,3})\)')

YOUTH_2026_27 = {
    'Omladinska liga Srbije': {
        'tm_name':'Omladinska liga Srbije','tm_comp_id':'SB19','age_group':'U19','expected':14,
        'tm_url':'https://www.transfermarkt.com/omladinska-liga-srbije/startseite/wettbewerb/SB19',
        'clubs':[
            ('Crvena zvezda U19','Red Star Belgrade U19','11137','roter-stern-belgrad-u19'),
            ('FK IMT U19','FK IMT Belgrad U19','83839','fk-imt-belgrad-u19'),
            ('FK Mladost Lučani U19','FK Mladost Lucani U19','66821','fk-mladost-lucani-u19'),
            ('FK Jedinstvo Ub U19','FK Jedinstvo Ub U19','75551','fk-jedinstvo-ub-u19'),
            ('FK TSC Bačka Topola U19','FK TSC Backa Topola U19','68302','fk-tsc-backa-topola-u19'),
            ('FK Čukarički U19','FK Čukarički U19','15058','fk-cukaricki-u19'),
            ('OFK Beograd U19','OFK Beograd U19','15072','ofk-belgrad-u19'),
            ('FK Spartak Subotica U19','FK Spartak Subotica U19','35851','fk-spartak-subotica-u19'),
            ('FK Vojvodina Novi Sad U19','FK Vojvodina Novi Sad U19','50306','fk-vojvodina-novi-sad-u19'),
            ('FK Ušće Novi Beograd U19','FK Usce Novi Beograd U19','113128','fk-usce-novi-beograd-u19'),
            ('FK Novi Pazar U19','FK Novi Pazar U19','27127','fk-novi-pazar-u19'),
            ('FK Partizan U19','FK Partizan Belgrade U19','12373','fk-partizan-belgrad-u19'),
            ('FK Brodarac U19','FK Brodarac Belgrad U19','67496','fk-brodarac-belgrad-u19'),
            ('FK Voždovac U19','FK Vozdovac U19','54517','fk-vozdovac-u19'),
        ],
    },
    'Kadetska liga Srbije': {
        'tm_name':'Kadetska liga Srbije','tm_comp_id':'SB17','age_group':'U17','expected':14,
        'tm_url':'https://www.transfermarkt.com/kadetska-liga-srbije/startseite/wettbewerb/SB17',
        'clubs':[
            ('FK Čukarički U17','FK Čukarički U17','56959','fk-cukaricki-u17'),
            ('FK IMT U17','FK IMT Belgrad U17','83840','fk-imt-belgrad-u17'),
            ('FK Železničar Pančevo U17','FK Zeleznicar Pancevo U17','98635','zeleznicar-pancevo-u17'),
            ('FK TSC Bačka Topola U17','FK TSC Backa Topola U17','68303','fk-tsc-backa-topola-u17'),
            ('FK Mladost Lučani U17','FK Mladost Lucani U17','73595','fk-mladost-lucani-u17'),
            ('FK Partizan U17','FK Partizan Belgrade U17','66822','fk-partizan-belgrad-u17'),
            ('Grafičar U17','Graficar Belgrad U17','76276','graficar-belgrad-u17'),
            ('FK Radnički 1923 Kragujevac U17','FK Radnicki 1923 Kragujevac U17','90882','fk-radnicki-1923-kragujevac-u17'),
            ('Crvena zvezda U17','Red Star Belgrade U17','54866','roter-stern-belgrad-u17'),
            ('FK Teleoptik Zemun U17','FK Teleoptik Zemun U17','68979','fk-teleoptik-zemun-u17'),
            ('FK Vojvodina Novi Sad U17','FK Vojvodina Novi Sad U17','71671','fk-vojvodina-novi-sad-u17'),
            ('FK Spartak Subotica U17','FK Spartak Subotica U17','71678','fk-spartak-subotica-u17'),
            ('FK Voždovac U17','FK Vozdovac U17','73591','fk-vozdovac-u17'),
            ('011 Beograd U17','011 Belgrad U17','98680','011-belgrad-u17'),
        ],
    },
}

SERBIA_2026_27 = {
    'SuperLiga Srbije': {
        'tm_url': 'https://www.transfermarkt.com/super-liga-srbije/startseite/wettbewerb/SER1',
        'clubs': [
            ('Red Star Belgrade','Red Star Belgrade'), ('Vojvodina Novi Sad','Vojvodina Novi Sad'),
            ('Radnicki 1923','Radnicki 1923'), ('IMT Novi Beograd','IMT Novi Beograd'),
            ('Radnik Surdulica','Radnik Surdulica'), ('Mladost Lucani','Mladost Lucani'),
            ('Novi Pazar','Novi Pazar'), ('Zeleznicar Pancevo','Zeleznicar Pancevo'),
            ('Cukaricki','Cukaricki'), ('OFK Beograd','OFK Beograd'), ('Radnicki Nis','Radnicki Nis'),
            ('Macva Sabac','Macva Sabac'), ('Partizan Belgrade','Partizan Belgrade'), ('Zemun','Zemun'),
        ],
    },
    'Prva liga Srbije': {
        'tm_url': 'https://www.transfermarkt.com/prva-liga-srbije/startseite/wettbewerb/SER2',
        'clubs': [
            ('Borac 1926','Borac 1926'), ('TSC Backa Topola','TSC Backa Topola'), ('OFK Vrsac','OFK Vrsac'),
            ('Javor Matis','Javor Matis'), ('Spartak Subotica','Spartak Subotica'), ('Smederevo 1924','Smederevo 1924'),
            ('GFK Dubocica 1923','GFK Dubocica 1923'), ('Napredak Krusevac','Napredak Krusevac'),
            ('Metalac Gornji Milanovac','Metalac Gornji Milanovac'), ('Vozdovac','Vozdovac'), ('Loznica','Loznica'),
            ('Bor 1919','Bor 1919'), ('Proleter 023','Proleter 023'), ('Jedinstvo Ub','Jedinstvo Ub'),
            ('RFK Graficar','RFK Graficar'), ('Teleoptik','Teleoptik'),
        ],
    },
}


SERBIASPORT_2026_27 = {
    'Srpska liga Beograd': {
        'url': 'https://srbijasport.net/league/8682-srpska-liga-beograd',
        'expected': 14,
    },
    'Srpska liga Vojvodina': {
        'url': 'https://srbijasport.net/league/8669-srpska-liga-vojvodina',
        'expected': 16,
    },
    'Srpska liga Istok': {
        'url': 'https://srbijasport.net/league/8659-srpska-liga-istok',
        'expected': 16,
    },
    'Srpska liga Zapad': {
        'url': 'https://srbijasport.net/league/8650-srpska-liga-zapad',
        'expected': 16,
    },
}


ZONES_2026_27 = {
    'Beogradska zona': {'url': 'https://srbijasport.net/league/8692-beogradska-zona', 'expected': 14},
    'Vojvođanska liga "Jug"': {'url': 'https://srbijasport.net/league/8670-vojvodjanska-liga-jug', 'expected': 14},
    'Vojvođanska liga "Sever"': {'url': 'https://srbijasport.net/league/8671-vojvodjanska-liga-sever', 'expected': 14},
    'Vojvođanska liga "Istok"': {'url': 'https://srbijasport.net/league/8672-vojvodjanska-liga-istok', 'expected': 14},
    'Zona "Dunav"': {'url': 'https://srbijasport.net/league/8653-zona-dunav', 'expected': 12},
    'Zona "Morava"': {'url': 'https://srbijasport.net/league/8654-zona-morava', 'expected': 12},
    'Zona "Drina"': {'url': 'https://srbijasport.net/league/8656-zona-drina', 'expected': 12},
    'Zona "Centar"': {'url': 'https://srbijasport.net/league/8689-zona-centar', 'expected': 12},
    'Zona "Zapad"': {'url': 'https://srbijasport.net/league/8693-zona-zapad', 'expected': 12},
    'Zona "Istok"': {'url': 'https://srbijasport.net/league/8732-zona-istok', 'expected': 12},
    'Zona "Jug"': {'url': 'https://srbijasport.net/league/8733-zona-jug', 'expected': 12},
}


# Level 5 / Okružne lige — SrbijaSport 2026/27.
# expected=0 intentionally: the parser reads the dedicated Timovi section and
# stops before fixtures, so membership size is taken from the live source.
LEVEL5_2026_27 = {
    'PFL Sombor': {'url': 'https://srbijasport.net/league/8644-pfl-sombor', 'expected': 0},
    'Raška okružna liga': {'url': 'https://srbijasport.net/league/8645-raska-okruzna-liga', 'expected': 0},
    'PFL Pančevo': {'url': 'https://srbijasport.net/league/8646-pfl-pancevo', 'expected': 0},
    'PFL Subotica': {'url': 'https://srbijasport.net/league/8657-pfl-subotica', 'expected': 0},
    'Podunavska okružna liga': {'url': 'https://srbijasport.net/league/8660-podunavska-okruzna-liga', 'expected': 0},
    'Kolubarska okružna liga': {'url': 'https://srbijasport.net/league/8661-kolubarska-okruzna-liga', 'expected': 0},
    'Nišavska okružna liga': {'url': 'https://srbijasport.net/league/8683-nisavska-okruzna-liga', 'expected': 0},
    'Prva rasinska okružna liga': {'url': 'https://srbijasport.net/league/8684-prva-rasinska-okruzna-liga', 'expected': 0},
    'Prva Beogradska liga - grupa "B"': {'url': 'https://srbijasport.net/league/8685-prva-beogradska-liga---grupa-b', 'expected': 0},
    'Braničevska okružna liga': {'url': 'https://srbijasport.net/league/8687-branicevska-okruzna-liga', 'expected': 0},
    'Zaječarska okružna liga': {'url': 'https://srbijasport.net/league/8694-zajecarska-okruzna-liga', 'expected': 0},
    'Prva novosadska liga': {'url': 'https://srbijasport.net/league/8695-prva-novosadska-liga', 'expected': 0},
    'Pomoravska okružna liga': {'url': 'https://srbijasport.net/league/8697-pomoravska-okruzna-liga', 'expected': 0},
    'Jablanička okružna liga - grupa Istok': {'url': 'https://srbijasport.net/league/8700-jablanicka-okruzna-liga---grupa-istok', 'expected': 0},
    'Jablanička okružna liga - grupa Zapad': {'url': 'https://srbijasport.net/league/8701-jablanicka-okruzna-liga---grupa-zapad', 'expected': 0},
    'Borska okružna liga': {'url': 'https://srbijasport.net/league/8703-borska-okruzna-liga', 'expected': 0},
    'Mačvanska okružna liga': {'url': 'https://srbijasport.net/league/8706-macvanska-okruzna-liga', 'expected': 0},
    'Sremska liga': {'url': 'https://srbijasport.net/league/8707-sremska-liga', 'expected': 0},
    'Pirotska okružna liga': {'url': 'https://srbijasport.net/league/8715-pirotska-okruzna-liga', 'expected': 0},
    'Prva Beogradska liga - grupa "A"': {'url': 'https://srbijasport.net/league/8719-prva-beogradska-liga---grupa-a', 'expected': 0},
    'Prva Beogradska liga - grupa "C"': {'url': 'https://srbijasport.net/league/8720-prva-beogradska-liga---grupa-c', 'expected': 0},
    'Moravička okružna liga': {'url': 'https://srbijasport.net/league/8746-moravicka-okruzna-liga', 'expected': 0},
    'PFL Zrenjanin': {'url': 'https://srbijasport.net/league/8748-pfl-zrenjanin', 'expected': 0},
    'Šumadijska okružna liga': {'url': 'https://srbijasport.net/league/8753-sumadijska-okruzna-liga', 'expected': 0},
    'Zlatiborska okružna liga': {'url': 'https://srbijasport.net/league/8755-zlatiborska-okruzna-liga', 'expected': 0},
    'Prva Niška liga': {'url': 'https://srbijasport.net/league/8779-prva-niska-liga', 'expected': 0},
    'Prva liga Kragujevca': {'url': 'https://srbijasport.net/league/8788-prva-liga-kragujevca', 'expected': 0},
    'Toplička okružna liga': {'url': 'https://srbijasport.net/league/8795-toplicka-okruzna-liga', 'expected': 0},
}


# Women's football — Serbia 2026/27.
# Membership is taken from SrbijaSport for all five club competitions. The two
# national divisions also retain their official FSS competition URL so that
# current player rosters can be assembled from official FSS match reports.
WOMEN_2026_27 = {
    'Super liga Srbije (ž)': {
        'level': 1,
        'url': 'https://srbijasport.net/league/8651-superliga-srbije',
        'expected': 8,
        'fss_url': 'https://fss.rs/takmicenje/super-liga-srbije-z-26-27/?script=lat',
        'image_path': 'static/branding/women_super_liga.png',
    },
    'Prva liga Srbije (ž)': {
        'level': 2,
        'url': 'https://srbijasport.net/league/8652-prva-liga-srbije',
        'expected': 8,
        'fss_url': 'https://fss.rs/takmicenje/prva-liga-srbije-z-26-27/?script=lat',
        'image_path': 'static/branding/women_prva_liga.png',
    },
    'Druga ženska liga Beograd': {
        'level': 3,
        'url': 'https://srbijasport.net/league/8959-druga-zenska-liga-beograd',
        'expected': 6,
        'fss_url': '',
        'image_path': 'static/branding/women_druga_beograd.png',
    },
    'Druga ženska liga Istok': {
        'level': 3,
        'url': 'https://srbijasport.net/league/8907-druga-zenska-liga-istok',
        'expected': 5,
        'fss_url': '',
        'image_path': 'static/branding/women_druga_istok.png',
    },
    'Druga ženska liga Zapad': {
        'level': 3,
        'url': 'https://srbijasport.net/league/8836-druga-zenska-liga-zapad',
        'expected': 7,
        'fss_url': '',
        'image_path': 'static/branding/women_druga_zapad.png',
    },
}

WOMEN_NATIONAL_2026 = [
    {
        'name': 'Srbija A – žene', 'short_name': 'Srbija žene', 'team_type': 'women_senior',
        'source_url': 'https://fss.rs/selekcije/a-selekcija-zene/?script=lat',
        'source_key': 'fss-women-senior', 'slug': 'a-selekcija-zene',
        'image_path': 'static/branding/women_nt_a.png',
    },
    {
        'name': 'Srbija U19 – žene', 'short_name': 'Srbija U19 žene', 'team_type': 'women_u19',
        'source_url': 'https://fss.rs/selekcije/u-19-zene/?script=lat',
        'source_key': 'fss-women-u19', 'slug': 'u-19-zene',
        'image_path': 'static/branding/women_nt_u19.png',
    },
]


# Verified Transfermarkt mappings for club names as exposed by SrbijaSport.
# Keys use the same normalized form as tm_match_normalize().
VERIFIED_TM_CLUBS = {
    "jedinstvo sp stara pazova": {
        "tm_name": "FK Jedinstvo Stara Pazova",
        "tm_club_id": "38827",
        "tm_slug": "fk-jedinstvo-stara-pazova",
    },
    "mladost bj backi jarak": {
        "tm_name": "FK Mladost Backi Jarak",
        "tm_club_id": "19746",
        "tm_slug": "fk-mladost-backi-jarak",
    },
    "backa palanka": {
        "tm_name": "OFK Backa",
        "tm_club_id": "4650",
        "tm_slug": "ofk-backa",
    },
    "gradnulica zrenjanin": {
        "tm_name": "OFK Gradnulica",
        "tm_club_id": "112147",
        "tm_slug": "ofk-gradnulica",
    },
    "mladost apa apatin": {
        "tm_name": "FK Mladost Apatin",
        "tm_club_id": "904",
        "tm_slug": "fk-mladost-apatin",
    },
    "vrbas": {
        "tm_name": "OFK Vrbas",
        "tm_club_id": "2152",
        "tm_slug": "ofk-vrbas",
    },
    "veternik": {
        "tm_name": "FK Veternik",
        "tm_club_id": "4653",
        "tm_slug": "fk-veternik",
    },
    "zeleznicar indjija": {
        "tm_name": "Zeleznicar Indjija",
        "tm_club_id": "94939",
        "tm_slug": "zeleznicar-indjija",
    },
    # Verified mappings: Srpska liga Beograd / Istok / Zapad
    "bask beograd": {
        "tm_name": "FK BASK Belgrad",
        "tm_club_id": "8975",
        "tm_slug": "fk-bask-belgrad",
    },
    "bezanija beograd": {
        "tm_name": "FK Bezanija",
        "tm_club_id": "10319",
        "tm_slug": "fk-bezanija",
    },
    "brodarac novi beograd": {
        "tm_name": "FK Brodarac Belgrad",
        "tm_club_id": "43656",
        "tm_slug": "fk-brodarac-belgrad",
    },
    "gsp polet dorcol beograd": {
        "tm_name": "FK GSP Polet Belgrad",
        "tm_club_id": "46949",
        "tm_slug": "fk-gsp-polet-belgrad",
    },
    "pkb 1950 dudovica": {
        "tm_name": "FK PKB Padinska Skela",
        "tm_club_id": "22420",
        "tm_slug": "fk-pkb-padinska-skela",
    },
    "t6 nika glogonjski rit": {
        "tm_name": "T6 Nika Lepusnica",
        "tm_club_id": "94942",
        "tm_slug": "lepusnica",
    },
    "torlak beograd": {
        "tm_name": "FK Torlak Kumodraz",
        "tm_club_id": "64164",
        "tm_slug": "fk-torlak-kumodraz",
    },
    "usce novi beograd novi beograd": {
        "tm_name": "FK Usce Novi Beograd",
        "tm_club_id": "57461",
        "tm_slug": "fk-usce-novi-beograd",
    },
    "usce novi beograd": {
        "tm_name": "FK Usce Novi Beograd",
        "tm_club_id": "57461",
        "tm_slug": "fk-usce-novi-beograd",
    },
    "zvezdara beograd": {
        "tm_name": "FK Zvezdara",
        "tm_club_id": "4654",
        "tm_slug": "fk-zvezdara",
    },
    "jagodina": {
        "tm_name": "GFK Jagodina",
        "tm_club_id": "12132",
        "tm_slug": "gfk-jagodina",
    },
    "malosiste": {
        "tm_name": "FK Malosiste",
        "tm_club_id": "96499",
        "tm_slug": "fk-malosiste",
    },
    "morava 1918 cuprija": {
        "tm_name": "FK Morava Cuprija",
        "tm_club_id": "37985",
        "tm_slug": "fk-morava-cuprija",
    },
    "brzi brod brzi brod": {
        "tm_name": "OFK Brzi Brod Nis",
        "tm_club_id": "94234",
        "tm_slug": "ofk-brzi-brod-nis",
    },
    "brzi brod": {
        "tm_name": "OFK Brzi Brod Nis",
        "tm_club_id": "94234",
        "tm_slug": "ofk-brzi-brod-nis",
    },
    "sindjelic nis": {
        "tm_name": "FK Sindjelic Nis",
        "tm_club_id": "25196",
        "tm_slug": "fk-sindjelic-nis",
    },
    "radnicki p pirot": {
        "tm_name": "FK Radnicki Pirot",
        "tm_club_id": "11045",
        "tm_slug": "fk-radnicki-pirot",
    },
    "radnicki pirot": {
        "tm_name": "FK Radnicki Pirot",
        "tm_club_id": "11045",
        "tm_slug": "fk-radnicki-pirot",
    },
    "radnicki s svilajnac": {
        "tm_name": "FK Radnicki Svilajnac",
        "tm_club_id": "26227",
        "tm_slug": "fk-radnicki-svilajnac",
    },
    "radnicki svilajnac": {
        "tm_name": "FK Radnicki Svilajnac",
        "tm_club_id": "26227",
        "tm_slug": "fk-radnicki-svilajnac",
    },
    "sloga l leskovac": {
        "tm_name": "FK Sloga Leskovac",
        "tm_club_id": "74428",
        "tm_slug": "fk-sloga-leskovac",
    },
    "sloga leskovac": {
        "tm_name": "FK Sloga Leskovac",
        "tm_club_id": "74428",
        "tm_slug": "fk-sloga-leskovac",
    },
    "timok 1919 zajecar": {
        "tm_name": "FK Timok Zajecar",
        "tm_club_id": "10460",
        "tm_slug": "fk-timok-zajecar",
    },
    "djerdap kladovo": {
        "tm_name": "FK Djerdap Kladovo",
        "tm_club_id": "42154",
        "tm_slug": "fk-djerdap-kladovo",
    },
    "buducnost krusik 2014 valjevo": {
        "tm_name": "FK Buducnost Krusik Valjevo",
        "tm_club_id": "15614",
        "tm_slug": "fk-buducnost-krusik-valjevo",
    },
    "omladinac z zablace": {
        "tm_name": "Omladinac Zablace",
        "tm_club_id": "96394",
        "tm_slug": "omladinac-zablace",
    },
    "omladinac zablace": {
        "tm_name": "Omladinac Zablace",
        "tm_club_id": "96394",
        "tm_slug": "omladinac-zablace",
    },
    "radnicki v valjevo": {
        "tm_name": "Radnicki Valjevo",
        "tm_club_id": "96450",
        "tm_slug": "radnicki-valjevo",
    },
    "radnicki valjevo": {
        "tm_name": "Radnicki Valjevo",
        "tm_club_id": "96450",
        "tm_slug": "radnicki-valjevo",
    },
    "radjevac krupanj": {
        "tm_name": "FK Radjevac Krupanj",
        "tm_club_id": "81595",
        "tm_slug": "fk-radjevac-krupanj",
    },
    "sloga p pozega": {
        "tm_name": "FK Sloga Pozega",
        "tm_club_id": "26427",
        "tm_slug": "fk-sloga-pozega",
    },
    "sloga pozega": {
        "tm_name": "FK Sloga Pozega",
        "tm_club_id": "26427",
        "tm_slug": "fk-sloga-pozega",
    },
    "djerdap golubac": {
        "tm_name": "Djerdap Golubac",
        "tm_club_id": "95590",
        "tm_slug": "djerdap-golubac",
    },
    "sumadija 1903 kragujevac": {
        "tm_name": "FK Sumadija 1903",
        "tm_club_id": "58463",
        "tm_slug": "fk-sumadija-1903",
    },

}

def verified_tm_candidate(source_name):
    key=tm_match_normalize(source_name)
    data=VERIFIED_TM_CLUBS.get(key)
    if not data:
        return None
    out=dict(data)
    out["tm_url"]=f'https://www.transfermarkt.com/{out["tm_slug"]}/startseite/verein/{out["tm_club_id"]}'
    out["staff_url"]=f'https://www.transfermarkt.com/{out["tm_slug"]}/mitarbeiter/verein/{out["tm_club_id"]}'
    out["score"]=1.0
    return out

TM_CLUB_ALIASES = {
    'roter stern belgrad':'Red Star Belgrade','crvena zvezda':'Red Star Belgrade','fk crvena zvezda':'Red Star Belgrade','red star belgrade':'Red Star Belgrade',
    'fk vojvodina novi sad':'Vojvodina Novi Sad','vojvodina':'Vojvodina Novi Sad',
    'fk radnicki 1923':'Radnicki 1923','radnicki 1923':'Radnicki 1923','fk radnicki 1923 kragujevac':'Radnicki 1923',
    'fk imt':'IMT Novi Beograd','imt':'IMT Novi Beograd','imt novi beograd':'IMT Novi Beograd','fk imt belgrad':'IMT Novi Beograd',
    'fk radnik surdulica':'Radnik Surdulica','radnik surdulica':'Radnik Surdulica',
    'fk mladost lucani':'Mladost Lucani','mladost lucani':'Mladost Lucani',
    'fk novi pazar':'Novi Pazar','novi pazar':'Novi Pazar',
    'fk zeleznicar pancevo':'Zeleznicar Pancevo','zeleznicar pancevo':'Zeleznicar Pancevo',
    'fk cukaricki':'Cukaricki','cukaricki':'Cukaricki',
    'ofk beograd':'OFK Beograd','ofk belgrad':'OFK Beograd',
    'fk radnicki nis':'Radnicki Nis','radnicki nis':'Radnicki Nis',
    'fk macva sabac':'Macva Sabac','macva sabac':'Macva Sabac',
    'fk partizan':'Partizan Belgrade','partizan belgrade':'Partizan Belgrade','fk partizan belgrad':'Partizan Belgrade','fk partizan belgrade':'Partizan Belgrade',
    'fk zemun':'Zemun','zemun':'Zemun',
    'fk borac 1926':'Borac 1926','borac 1926':'Borac 1926','fk borac cacak':'Borac 1926','fk borac 1926 cacak':'Borac 1926',
    'fk tsc backa topola':'TSC Backa Topola','tsc backa topola':'TSC Backa Topola','fk tsc':'TSC Backa Topola',
    'ofk vrsac':'OFK Vrsac',
    'fk javor matis':'Javor Matis','javor matis':'Javor Matis','fk javor matis ivanjica':'Javor Matis',
    'fk spartak subotica':'Spartak Subotica','spartak subotica':'Spartak Subotica','fk spartak zdrepceva krv':'Spartak Subotica',
    'fk smederevo 1924':'Smederevo 1924','smederevo 1924':'Smederevo 1924',
    'gfk dubocica':'GFK Dubocica 1923','gfk dubocica 1923':'GFK Dubocica 1923','fk dubocica leskovac':'GFK Dubocica 1923',
    'fk napredak krusevac':'Napredak Krusevac','napredak krusevac':'Napredak Krusevac',
    'fk metalac gornji milanovac':'Metalac Gornji Milanovac','metalac gornji milanovac':'Metalac Gornji Milanovac',
    'fk vozdovac':'Vozdovac','vozdovac':'Vozdovac','fk loznica':'Loznica','loznica':'Loznica',
    'ofk bor 1919':'Bor 1919','bor 1919':'Bor 1919','fk bor 1919':'Bor 1919',
    'fk proleter 023':'Proleter 023','proleter 023':'Proleter 023','fk proleter 023 zrenjanin':'Proleter 023',
    'fk jedinstvo ub':'Jedinstvo Ub','jedinstvo ub':'Jedinstvo Ub',
    'rfk graficar beograd':'RFK Graficar','rfk graficar':'RFK Graficar','graficar belgrad':'RFK Graficar',
    'fk teleoptik':'Teleoptik','teleoptik':'Teleoptik','fk teleoptik zemun':'Teleoptik',
}


def esc(v): return html_lib.escape(str(v or ''))
def clean_text(v): return re.sub(r'\s+',' ',str(v or '').replace('\xa0',' ')).strip()

def db():
    con=sqlite3.connect(DB_PATH); con.row_factory=sqlite3.Row; return con


def norm_club_name(value):
    s=clean_text(value).lower().replace('đ','dj')
    s=unicodedata.normalize('NFKD',s)
    s=''.join(ch for ch in s if not unicodedata.combining(ch))
    s=re.sub(r'[^a-z0-9]+',' ',s)
    return re.sub(r'\s+',' ',s).strip()


def seed_serbia_clubs(con):
    cur=con.cursor(); country=cur.execute("SELECT id FROM countries WHERE name='Serbia'").fetchone()
    if not country: return
    cid=country[0]
    for lname,info in SERBIA_2026_27.items():
        league=cur.execute('SELECT id FROM leagues WHERE country_id=? AND name=?',(cid,lname)).fetchone()
        if not league: continue
        lid=league[0]; cur.execute('UPDATE leagues SET tm_url=? WHERE id=?',(info['tm_url'],lid))
        for display,cm in info['clubs']:
            ex=cur.execute('SELECT id FROM clubs WHERE league_id=? AND (lower(name)=lower(?) OR lower(cm_name)=lower(?))',(lid,display,cm)).fetchone()
            if not ex:
                cur.execute('INSERT INTO clubs(league_id,name,tm_name,cm_name,tm_club_id,tm_slug,tm_url,staff_url) VALUES(?,?,?,?,?,?,?,?)',(lid,display,'',cm,'','','',''))
    con.commit()


def init_db():
    DB_PATH.parent.mkdir(exist_ok=True)
    con=sqlite3.connect(DB_PATH, timeout=1.5)
    # WAL keeps read-only UI pages responsive while a sync task is writing.
    con.execute('PRAGMA busy_timeout=1500')
    con.execute('PRAGMA journal_mode=WAL')
    con.execute('PRAGMA synchronous=NORMAL')
    con.executescript("""
    PRAGMA foreign_keys=ON;
    CREATE TABLE IF NOT EXISTS countries(id INTEGER PRIMARY KEY AUTOINCREMENT,name TEXT NOT NULL UNIQUE,tm_name TEXT,code TEXT);
    CREATE TABLE IF NOT EXISTS leagues(id INTEGER PRIMARY KEY AUTOINCREMENT,country_id INTEGER NOT NULL,name TEXT NOT NULL,tm_name TEXT,tm_url TEXT,level INTEGER DEFAULT 1,UNIQUE(country_id,name),FOREIGN KEY(country_id) REFERENCES countries(id) ON DELETE CASCADE);
    CREATE TABLE IF NOT EXISTS clubs(id INTEGER PRIMARY KEY AUTOINCREMENT,league_id INTEGER NOT NULL,name TEXT NOT NULL,tm_name TEXT,cm_name TEXT,tm_club_id TEXT,tm_slug TEXT,tm_url TEXT,staff_url TEXT,last_staff_update TEXT,UNIQUE(league_id,name),FOREIGN KEY(league_id) REFERENCES leagues(id) ON DELETE CASCADE);
    CREATE TABLE IF NOT EXISTS staff(id INTEGER PRIMARY KEY AUTOINCREMENT,club_id INTEGER NOT NULL,section TEXT,name TEXT,nationality TEXT,date_of_birth TEXT,age TEXT,appointed TEXT,contract_expires TEXT,last_club TEXT,profile_url TEXT,dob_source TEXT,position TEXT,updated_at TEXT DEFAULT CURRENT_TIMESTAMP,UNIQUE(club_id,section,profile_url),FOREIGN KEY(club_id) REFERENCES clubs(id) ON DELETE CASCADE);
    CREATE TABLE IF NOT EXISTS players(id INTEGER PRIMARY KEY AUTOINCREMENT,club_id INTEGER NOT NULL,squad_number TEXT,name TEXT NOT NULL,tm_player_id TEXT,game_safe_name TEXT,safe_name_length INTEGER,position TEXT,date_of_birth TEXT,age TEXT,nationality TEXT,market_value TEXT,profile_url TEXT,updated_at TEXT DEFAULT CURRENT_TIMESTAMP,UNIQUE(club_id,profile_url),FOREIGN KEY(club_id) REFERENCES clubs(id) ON DELETE CASCADE);
    CREATE TABLE IF NOT EXISTS club_info(
      club_id INTEGER PRIMARY KEY,
      official_name TEXT,founded TEXT,address_line1 TEXT,postal_city TEXT,city TEXT,country TEXT,
      phone TEXT,fax TEXT,website TEXT,squad_size TEXT,average_age TEXT,foreigners TEXT,foreigners_pct TEXT,
      national_team_players TEXT,transfer_record TEXT,crest_url TEXT,crest_local TEXT,
      updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
      FOREIGN KEY(club_id) REFERENCES clubs(id) ON DELETE CASCADE);
    CREATE TABLE IF NOT EXISTS stadiums(
      club_id INTEGER PRIMARY KEY,name TEXT,capacity INTEGER,city TEXT,
      image_url TEXT,image_local TEXT,image_source TEXT,opened TEXT,surface TEXT,
      updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
      FOREIGN KEY(club_id) REFERENCES clubs(id) ON DELETE CASCADE);

    CREATE TABLE IF NOT EXISTS history_competitions(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      history_key TEXT NOT NULL UNIQUE,
      current_league_id INTEGER,
      name_sr TEXT NOT NULL,name_en TEXT,
      comp_type TEXT NOT NULL DEFAULT 'league',
      country_code TEXT DEFAULT 'SRB',level INTEGER,
      active_from TEXT,active_to TEXT,sort_order INTEGER DEFAULT 0,
      notes TEXT,source_name TEXT,source_url TEXT,
      FOREIGN KEY(current_league_id) REFERENCES leagues(id) ON DELETE SET NULL);
    CREATE TABLE IF NOT EXISTS history_competition_aliases(
      id INTEGER PRIMARY KEY AUTOINCREMENT,competition_id INTEGER NOT NULL,
      name_sr TEXT NOT NULL,name_en TEXT,from_season TEXT,to_season TEXT,
      UNIQUE(competition_id,name_sr,from_season,to_season),
      FOREIGN KEY(competition_id) REFERENCES history_competitions(id) ON DELETE CASCADE);
    CREATE TABLE IF NOT EXISTS history_club_seasons(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      club_id INTEGER,competition_id INTEGER NOT NULL,season_label TEXT NOT NULL,
      club_name_override TEXT,position INTEGER,stage_sr TEXT,stage_en TEXT,
      played INTEGER,wins INTEGER,draws INTEGER,losses INTEGER,
      goals_for INTEGER,goals_against INTEGER,points REAL,
      is_champion INTEGER DEFAULT 0,is_runner_up INTEGER DEFAULT 0,
      promoted INTEGER DEFAULT 0,relegated INTEGER DEFAULT 0,highlight INTEGER DEFAULT 0,
      source_name TEXT,source_url TEXT,source_note TEXT,updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
      UNIQUE(club_id,competition_id,season_label),
      FOREIGN KEY(club_id) REFERENCES clubs(id) ON DELETE CASCADE,
      FOREIGN KEY(competition_id) REFERENCES history_competitions(id) ON DELETE CASCADE);
    CREATE INDEX IF NOT EXISTS idx_history_club ON history_club_seasons(club_id,season_label);
    CREATE INDEX IF NOT EXISTS idx_history_comp ON history_club_seasons(competition_id,season_label);

    CREATE TABLE IF NOT EXISTS club_transfer_summary(
      club_id INTEGER NOT NULL,
      season_id INTEGER NOT NULL,
      season_label TEXT,
      arrivals_count INTEGER DEFAULT 0,
      departures_count INTEGER DEFAULT 0,
      expenditure TEXT,
      income TEXT,
      balance TEXT,
      updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
      PRIMARY KEY(club_id,season_id),
      FOREIGN KEY(club_id) REFERENCES clubs(id) ON DELETE CASCADE);

    CREATE TABLE IF NOT EXISTS club_transfers(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      club_id INTEGER NOT NULL,
      season_id INTEGER NOT NULL,
      season_label TEXT,
      direction TEXT NOT NULL,
      player_name TEXT NOT NULL,
      player_url TEXT,
      player_tm_id TEXT,
      position TEXT,
      age TEXT,
      nationality TEXT,
      market_value TEXT,
      other_club TEXT,
      other_club_url TEXT,
      other_club_tm_id TEXT,
      other_league TEXT,
      other_country TEXT,
      fee TEXT,
      transfer_date TEXT,
      transfer_detail_url TEXT,
      updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
      FOREIGN KEY(club_id) REFERENCES clubs(id) ON DELETE CASCADE);

    CREATE INDEX IF NOT EXISTS idx_club_transfers_club_season
      ON club_transfers(club_id,season_id,direction);

    CREATE TABLE IF NOT EXISTS national_teams(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      country_id INTEGER NOT NULL,
      name TEXT NOT NULL,
      short_name TEXT,
      team_type TEXT NOT NULL,
      tm_name TEXT,
      tm_team_id TEXT NOT NULL,
      tm_slug TEXT NOT NULL,
      tm_url TEXT NOT NULL,
      staff_url TEXT NOT NULL,
      image_path TEXT,
      official_name TEXT,
      association TEXT,
      founded TEXT,
      squad_size TEXT,
      average_age TEXT,
      total_market_value TEXT,
      crest_url TEXT,
      crest_local TEXT,
      last_info_update TEXT,
      last_players_update TEXT,
      last_staff_update TEXT,
      UNIQUE(country_id,team_type),
      UNIQUE(tm_team_id),
      FOREIGN KEY(country_id) REFERENCES countries(id) ON DELETE CASCADE);

    CREATE TABLE IF NOT EXISTS national_team_players(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      national_team_id INTEGER NOT NULL,
      squad_number TEXT,
      name TEXT NOT NULL,
      tm_player_id TEXT,
      game_safe_name TEXT,
      safe_name_length INTEGER,
      position TEXT,
      date_of_birth TEXT,
      age TEXT,
      nationality TEXT,
      market_value TEXT,
      profile_url TEXT,
      updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
      UNIQUE(national_team_id,profile_url),
      FOREIGN KEY(national_team_id) REFERENCES national_teams(id) ON DELETE CASCADE);

    CREATE TABLE IF NOT EXISTS national_team_staff(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      national_team_id INTEGER NOT NULL,
      section TEXT,
      name TEXT,
      nationality TEXT,
      date_of_birth TEXT,
      age TEXT,
      appointed TEXT,
      contract_expires TEXT,
      last_club TEXT,
      profile_url TEXT,
      dob_source TEXT,
      position TEXT,
      updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
      UNIQUE(national_team_id,section,profile_url),
      FOREIGN KEY(national_team_id) REFERENCES national_teams(id) ON DELETE CASCADE);

    CREATE TABLE IF NOT EXISTS standings_competitions(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      country_id INTEGER NOT NULL,
      owner_type TEXT NOT NULL,
      owner_id INTEGER NOT NULL,
      name TEXT NOT NULL,
      season TEXT,
      group_name TEXT,
      source_type TEXT,
      source_url TEXT,
      updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
      UNIQUE(owner_type,owner_id,name,season,group_name));
    CREATE TABLE IF NOT EXISTS standings_rows(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      competition_id INTEGER NOT NULL,
      position INTEGER,
      position_note TEXT,
      team_name TEXT NOT NULL,
      club_id INTEGER,
      played INTEGER,wins INTEGER,draws INTEGER,losses INTEGER,
      goals_for INTEGER,goals_against INTEGER,goal_diff INTEGER,points INTEGER,
      form TEXT,note TEXT,
      UNIQUE(competition_id,team_name),
      FOREIGN KEY(competition_id) REFERENCES standings_competitions(id) ON DELETE CASCADE);
    CREATE INDEX IF NOT EXISTS idx_standings_owner ON standings_competitions(owner_type,owner_id);
    CREATE INDEX IF NOT EXISTS idx_standings_rows_comp ON standings_rows(competition_id,position);

    CREATE TABLE IF NOT EXISTS update_log(id INTEGER PRIMARY KEY AUTOINCREMENT,club_id INTEGER,module TEXT,source TEXT,rows_imported INTEGER,status TEXT,details TEXT,created_at TEXT DEFAULT CURRENT_TIMESTAMP);
    """)
    # Backward-compatible schema upgrades.
    league_cols={r[1] for r in con.execute('PRAGMA table_info(leagues)').fetchall()}
    if 'source_type' not in league_cols: con.execute("ALTER TABLE leagues ADD COLUMN source_type TEXT DEFAULT 'transfermarkt'")
    if 'source_url' not in league_cols: con.execute('ALTER TABLE leagues ADD COLUMN source_url TEXT')
    if 'source_expected_clubs' not in league_cols: con.execute('ALTER TABLE leagues ADD COLUMN source_expected_clubs INTEGER')
    if 'section' not in league_cols: con.execute("ALTER TABLE leagues ADD COLUMN section TEXT DEFAULT 'senior'")
    if 'age_group' not in league_cols: con.execute('ALTER TABLE leagues ADD COLUMN age_group TEXT')
    if 'tm_comp_id' not in league_cols: con.execute('ALTER TABLE leagues ADD COLUMN tm_comp_id TEXT')
    if 'fss_url' not in league_cols: con.execute('ALTER TABLE leagues ADD COLUMN fss_url TEXT')
    if 'image_path' not in league_cols: con.execute('ALTER TABLE leagues ADD COLUMN image_path TEXT')
    if 'region_code' not in league_cols: con.execute('ALTER TABLE leagues ADD COLUMN region_code TEXT')
    if 'standings_source_type' not in league_cols: con.execute('ALTER TABLE leagues ADD COLUMN standings_source_type TEXT')
    if 'standings_source_url' not in league_cols: con.execute('ALTER TABLE leagues ADD COLUMN standings_source_url TEXT')
    if 'source_active' not in league_cols: con.execute('ALTER TABLE leagues ADD COLUMN source_active INTEGER DEFAULT 1')
    con.execute("UPDATE leagues SET section='senior' WHERE COALESCE(section,'')=''")
    con.execute("UPDATE leagues SET source_active=1 WHERE source_active IS NULL")

    cols={r[1] for r in con.execute('PRAGMA table_info(clubs)').fetchall()}
    if 'last_players_update' not in cols: con.execute('ALTER TABLE clubs ADD COLUMN last_players_update TEXT')
    if 'source_name' not in cols: con.execute('ALTER TABLE clubs ADD COLUMN source_name TEXT')
    if 'source_club_url' not in cols: con.execute('ALTER TABLE clubs ADD COLUMN source_club_url TEXT')
    if 'source_active' not in cols: con.execute('ALTER TABLE clubs ADD COLUMN source_active INTEGER DEFAULT 1')
    if 'tm_match_status' not in cols: con.execute("ALTER TABLE clubs ADD COLUMN tm_match_status TEXT DEFAULT 'unresolved'")
    if 'tm_match_score' not in cols: con.execute('ALTER TABLE clubs ADD COLUMN tm_match_score REAL')
    if 'tm_candidate_name' not in cols: con.execute('ALTER TABLE clubs ADD COLUMN tm_candidate_name TEXT')
    if 'tm_candidate_url' not in cols: con.execute('ALTER TABLE clubs ADD COLUMN tm_candidate_url TEXT')
    if 'tm_match_note' not in cols: con.execute('ALTER TABLE clubs ADD COLUMN tm_match_note TEXT')
    if 'last_transfers_update' not in cols: con.execute('ALTER TABLE clubs ADD COLUMN last_transfers_update TEXT')

    stadium_cols={r[1] for r in con.execute('PRAGMA table_info(stadiums)').fetchall()}
    if 'image_url' not in stadium_cols: con.execute('ALTER TABLE stadiums ADD COLUMN image_url TEXT')
    if 'image_local' not in stadium_cols: con.execute('ALTER TABLE stadiums ADD COLUMN image_local TEXT')
    if 'image_source' not in stadium_cols: con.execute('ALTER TABLE stadiums ADD COLUMN image_source TEXT')
    if 'opened' not in stadium_cols: con.execute('ALTER TABLE stadiums ADD COLUMN opened TEXT')
    if 'surface' not in stadium_cols: con.execute('ALTER TABLE stadiums ADD COLUMN surface TEXT')

    player_cols={r[1] for r in con.execute('PRAGMA table_info(players)').fetchall()}
    if 'tm_player_id' not in player_cols: con.execute('ALTER TABLE players ADD COLUMN tm_player_id TEXT')
    nt_cols={r[1] for r in con.execute('PRAGMA table_info(national_teams)').fetchall()}
    if 'source_type' not in nt_cols: con.execute("ALTER TABLE national_teams ADD COLUMN source_type TEXT DEFAULT 'transfermarkt'")
    if 'source_url' not in nt_cols: con.execute('ALTER TABLE national_teams ADD COLUMN source_url TEXT')
    if 'standings_source_type' not in nt_cols: con.execute('ALTER TABLE national_teams ADD COLUMN standings_source_type TEXT')
    if 'standings_source_url' not in nt_cols: con.execute('ALTER TABLE national_teams ADD COLUMN standings_source_url TEXT')
    if 'standings_group_code' not in nt_cols: con.execute('ALTER TABLE national_teams ADD COLUMN standings_group_code TEXT')
    if 'standings_competition_code' not in nt_cols: con.execute('ALTER TABLE national_teams ADD COLUMN standings_competition_code TEXT')
    if 'standings_competition_name' not in nt_cols: con.execute('ALTER TABLE national_teams ADD COLUMN standings_competition_name TEXT')
    if 'standings_reference_url' not in nt_cols: con.execute('ALTER TABLE national_teams ADD COLUMN standings_reference_url TEXT')
    con.execute("UPDATE national_teams SET source_type='transfermarkt' WHERE COALESCE(source_type,'')=''")
    con.execute("UPDATE national_teams SET source_url=tm_url WHERE COALESCE(source_url,'')='' AND COALESCE(tm_url,'')<>''")

    nt_player_cols={r[1] for r in con.execute('PRAGMA table_info(national_team_players)').fetchall()}
    if 'tm_player_id' not in nt_player_cols: con.execute('ALTER TABLE national_team_players ADD COLUMN tm_player_id TEXT')
    if 'current_club' not in nt_player_cols: con.execute('ALTER TABLE national_team_players ADD COLUMN current_club TEXT')
    con.execute('CREATE INDEX IF NOT EXISTS idx_players_tm_player_id ON players(tm_player_id)')
    con.execute('CREATE INDEX IF NOT EXISTS idx_nt_players_tm_player_id ON national_team_players(tm_player_id)')
    con.execute('CREATE INDEX IF NOT EXISTS idx_clubs_league_active ON clubs(league_id,source_active)')
    con.execute('CREATE INDEX IF NOT EXISTS idx_players_club_id ON players(club_id)')
    con.execute('CREATE INDEX IF NOT EXISTS idx_staff_club_id ON staff(club_id)')
    con.execute('CREATE INDEX IF NOT EXISTS idx_nt_players_team_id ON national_team_players(national_team_id)')
    con.execute('CREATE INDEX IF NOT EXISTS idx_nt_staff_team_id ON national_team_staff(national_team_id)')

    # Backfill TM Player IDs from existing profile URLs.
    for pid,purl in con.execute("SELECT id,profile_url FROM players WHERE COALESCE(tm_player_id,'')='' AND COALESCE(profile_url,'')<>''").fetchall():
        m=re.search(r'/spieler/(\d+)',purl or '')
        if m: con.execute('UPDATE players SET tm_player_id=? WHERE id=?',(m.group(1),pid))
    for pid,purl in con.execute("SELECT id,profile_url FROM national_team_players WHERE COALESCE(tm_player_id,'')='' AND COALESCE(profile_url,'')<>''").fetchall():
        m=re.search(r'/spieler/(\d+)',purl or '')
        if m: con.execute('UPDATE national_team_players SET tm_player_id=? WHERE id=?',(m.group(1),pid))

    cur=con.cursor(); cur.execute('INSERT OR IGNORE INTO countries(name,tm_name,code) VALUES(?,?,?)',('Serbia','Serbia','SRB'))
    cid=cur.execute("SELECT id FROM countries WHERE name='Serbia'").fetchone()[0]
    cur.execute('INSERT OR IGNORE INTO leagues(country_id,name,tm_name,level) VALUES(?,?,?,?)',(cid,'SuperLiga Srbije','Super liga Srbije',1))
    cur.execute('INSERT OR IGNORE INTO leagues(country_id,name,tm_name,level) VALUES(?,?,?,?)',(cid,'Prva liga Srbije','Prva Liga Srbije',2))
    # Canonical men's national masters are protected identities.  A lower-league
    # discovery run must never be able to reclassify them as Level 3.
    cur.execute("""UPDATE leagues SET
        level=CASE name WHEN 'SuperLiga Srbije' THEN 1 WHEN 'Prva liga Srbije' THEN 2 ELSE level END,
        source_type='transfermarkt',source_url=COALESCE(tm_url,source_url),section='senior',source_active=1
        WHERE country_id=? AND name IN ('SuperLiga Srbije','Prva liga Srbije')""",(cid,))
    cur.execute("UPDATE leagues SET fss_url='https://fss.rs/takmicenje/mozzart-bet-super-liga-srbije-26-27/?script=lat',standings_source_type='fss',standings_source_url='https://fss.rs/takmicenje/mozzart-bet-super-liga-srbije-26-27/?script=lat' WHERE country_id=? AND name='SuperLiga Srbije'",(cid,))
    cur.execute("UPDATE leagues SET fss_url='https://fss.rs/takmicenje/mozzart-bet-prva-liga-srbije-26-27/?script=lat',standings_source_type='fss',standings_source_url='https://fss.rs/takmicenje/mozzart-bet-prva-liga-srbije-26-27/?script=lat' WHERE country_id=? AND name='Prva liga Srbije'",(cid,))

    # History masters: current league pages point to one long-lived competition identity.
    # Historical aliases/eras can be added without changing the live league rows.
    for _lname,_key,_en,_level in (
        ('SuperLiga Srbije','serbia-top-flight','Serbian top flight',1),
        ('Prva liga Srbije','serbia-second-tier','Serbian second tier',2),
    ):
        _lr=cur.execute('SELECT id FROM leagues WHERE country_id=? AND name=?',(cid,_lname)).fetchone()
        if _lr:
            cur.execute('''INSERT OR IGNORE INTO history_competitions(history_key,current_league_id,name_sr,name_en,comp_type,country_code,level,sort_order)
                VALUES(?,?,?,?,?,?,?,?)''',(_key,_lr[0],_lname,_en,'league','SRB',_level,_level))
            cur.execute('UPDATE history_competitions SET current_league_id=? WHERE history_key=?',(_lr[0],_key))

    for lname,spec in SERBIASPORT_2026_27.items():
        cur.execute('INSERT OR IGNORE INTO leagues(country_id,name,tm_name,level,source_type,source_url,source_expected_clubs) VALUES(?,?,?,?,?,?,?)',
                    (cid,lname,'',3,'srbijasport',spec['url'],spec['expected']))
        cur.execute("UPDATE leagues SET level=3,section='senior',source_type='srbijasport',source_url=?,source_expected_clubs=?,standings_source_type='srbijasport',standings_source_url=?,source_active=1 WHERE country_id=? AND name=?",
                    (spec['url'],spec['expected'],spec['url'],cid,lname))

    # Level 4 / Zonske lige: SrbijaSport membership first, TM mapping later.
    for lname,spec in ZONES_2026_27.items():
        cur.execute('INSERT OR IGNORE INTO leagues(country_id,name,tm_name,level,source_type,source_url,source_expected_clubs) VALUES(?,?,?,?,?,?,?)',
                    (cid,lname,'',4,'srbijasport',spec['url'],spec['expected']))
        cur.execute("UPDATE leagues SET level=4,section='senior',source_type='srbijasport',source_url=?,source_expected_clubs=?,standings_source_type='srbijasport',standings_source_url=? WHERE country_id=? AND name=?",
                    (spec['url'],spec['expected'],spec['url'],cid,lname))

    # Level 5 / Okružne lige: SrbijaSport owns competition membership.
    for lname,spec in LEVEL5_2026_27.items():
        cur.execute('INSERT OR IGNORE INTO leagues(country_id,name,tm_name,level,source_type,source_url,source_expected_clubs) VALUES(?,?,?,?,?,?,?)',
                    (cid,lname,'',5,'srbijasport',spec['url'],spec['expected']))
        cur.execute("UPDATE leagues SET level=5,section='senior',source_type='srbijasport',source_url=?,source_expected_clubs=?,standings_source_type='srbijasport',standings_source_url=? WHERE country_id=? AND name=?",
                    (spec['url'],spec['expected'],spec['url'],cid,lname))

    # Explicit regional ownership for the senior pyramid (Levels 3-5).
    # Level 6 is seeded below from LEVEL6_LEAGUES.
    for lname,region_code in REGION_LEAGUES.items():
        cur.execute("UPDATE leagues SET region_code=? WHERE country_id=? AND name=?",(region_code,cid,lname))

    # Level 6 / Međuopštinske lige. URLs are discovered live from the
    # SrbijaSport 2026/27 competition index so the project is not tied to
    # brittle numeric competition IDs.
    for lname,region_code in LEVEL6_LEAGUES.items():
        cur.execute('''INSERT OR IGNORE INTO leagues(
            country_id,name,tm_name,level,source_type,source_url,source_expected_clubs,section,region_code,standings_source_type
            ) VALUES(?,?,?,?,?,?,?,?,?,?)''',
            (cid,lname,'',6,'srbijasport','',0,'senior',region_code,'srbijasport'))
        cur.execute('''UPDATE leagues SET level=6,section='senior',source_type='srbijasport',
            region_code=?,standings_source_type='srbijasport' WHERE country_id=? AND name=?''',
            (region_code,cid,lname))

    # Women's club competitions. All five use SrbijaSport for current
    # membership; Super/Prva also keep official FSS URLs for roster extraction.
    for lname,spec in WOMEN_2026_27.items():
        cur.execute('''INSERT OR IGNORE INTO leagues(
            country_id,name,tm_name,level,source_type,source_url,source_expected_clubs,section,age_group,fss_url,image_path
            ) VALUES(?,?,?,?,?,?,?,?,?,?,?)''',
            (cid,lname,'',spec['level'],'srbijasport',spec['url'],spec['expected'],'women','Senior Women',spec['fss_url'],spec['image_path']))
        standings_type='fss' if spec['fss_url'] else 'srbijasport'
        standings_url=spec['fss_url'] or spec['url']
        cur.execute('''UPDATE leagues SET level=?,section='women',source_type='srbijasport',source_url=?,
            source_expected_clubs=?,age_group='Senior Women',fss_url=?,image_path=?,standings_source_type=?,standings_source_url=?,source_active=1 WHERE country_id=? AND name=?''',
            (spec['level'],spec['url'],spec['expected'],spec['fss_url'],spec['image_path'],standings_type,standings_url,cid,lname))

    # FSS Youth Section: national U19 and U17 leagues.
    for lname,spec in YOUTH_2026_27.items():
        cur.execute('''INSERT OR IGNORE INTO leagues(
            country_id,name,tm_name,tm_url,level,source_type,source_url,source_expected_clubs,section,age_group,tm_comp_id
            ) VALUES(?,?,?,?,?,?,?,?,?,?,?)''',
            (cid,lname,spec['tm_name'],spec['tm_url'],0,'transfermarkt',spec['tm_url'],spec['expected'],'youth',spec['age_group'],spec['tm_comp_id']))
        cur.execute('''UPDATE leagues SET tm_name=?,tm_url=?,level=0,source_type='transfermarkt',source_url=?,
            source_expected_clubs=?,section='youth',age_group=?,tm_comp_id=? WHERE country_id=? AND name=?''',
            (spec['tm_name'],spec['tm_url'],spec['tm_url'],spec['expected'],spec['age_group'],spec['tm_comp_id'],cid,lname))
        youth_fss_url={
            'Omladinska liga Srbije':'https://fss.rs/takmicenje/omladinska-liga-srbije-26-27/?script=lat',
            'Kadetska liga Srbije':'https://fss.rs/takmicenje/kadetska-liga-srbije-26-27/?script=lat',
        }.get(lname,'')
        if youth_fss_url:
            cur.execute("UPDATE leagues SET fss_url=?,standings_source_type='fss',standings_source_url=? WHERE country_id=? AND name=?",(youth_fss_url,youth_fss_url,cid,lname))
        ylid=cur.execute('SELECT id FROM leagues WHERE country_id=? AND name=?',(cid,lname)).fetchone()[0]
        current_ids=[]
        for display,tm_name,tmid,slug in spec['clubs']:
            current_ids.append(tmid)
            tm_url=f'https://www.transfermarkt.com/{slug}/startseite/verein/{tmid}'
            staff_url=f'https://www.transfermarkt.com/{slug}/mitarbeiter/verein/{tmid}'
            existing=cur.execute('SELECT id FROM clubs WHERE league_id=? AND tm_club_id=?',(ylid,tmid)).fetchone()
            if existing:
                cur.execute('''UPDATE clubs SET name=?,cm_name=?,tm_name=?,tm_slug=?,tm_url=?,staff_url=?,source_name=?,
                    source_active=1,tm_match_status='matched',tm_match_score=100 WHERE id=?''',
                    (display,display,tm_name,slug,tm_url,staff_url,tm_name,existing[0]))
            else:
                cur.execute('''INSERT INTO clubs(league_id,name,tm_name,cm_name,tm_club_id,tm_slug,tm_url,staff_url,
                    source_name,source_active,tm_match_status,tm_match_score) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)''',
                    (ylid,display,tm_name,display,tmid,slug,tm_url,staff_url,tm_name,1,'matched',100))
        # Keep old seasonal youth teams for history, but mark them inactive if they are no longer in the seeded current season.
        if current_ids:
            marks=','.join('?' for _ in current_ids)
            cur.execute(f'''UPDATE clubs SET source_active=0 WHERE league_id=? AND COALESCE(tm_club_id,'')<>''
                AND tm_club_id NOT IN ({marks})''',(ylid,*current_ids))

    # National selections are intentionally separate from league clubs.
    national_seed=[
        ('Serbia','Srbija','senior','Serbia','3438','serbia',
         'https://www.transfermarkt.com/serbia/startseite/verein/3438',
         'https://www.transfermarkt.com/serbia/mitarbeiter/verein/3438',
         'static/branding/serbia_nt.png','transfermarkt','https://www.transfermarkt.com/serbia/startseite/verein/3438'),
        ('Serbia U21','Srbija U21','u21','Serbia U21','9566','serbia-u21',
         'https://www.transfermarkt.com/serbia-u21/startseite/verein/9566',
         'https://www.transfermarkt.com/serbia-u21/mitarbeiter/verein/9566',
         'static/branding/serbia_u21.png','transfermarkt','https://www.transfermarkt.com/serbia-u21/startseite/verein/9566'),
    ]
    for w in WOMEN_NATIONAL_2026:
        national_seed.append((w['name'],w['short_name'],w['team_type'],w['name'],w['source_key'],w['slug'],
            w['source_url'],'',w['image_path'],'fss',w['source_url']))

    for nname,short_name,team_type,tm_name,tmid,slug,tm_url,staff_url,image_path,source_type,source_url in national_seed:
        cur.execute("""INSERT OR IGNORE INTO national_teams(
            country_id,name,short_name,team_type,tm_name,tm_team_id,tm_slug,tm_url,staff_url,image_path,source_type,source_url
            ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)""",
            (cid,nname,short_name,team_type,tm_name,tmid,slug,tm_url,staff_url,image_path,source_type,source_url))
        cur.execute("""UPDATE national_teams SET
            name=?,short_name=?,tm_name=?,tm_team_id=?,tm_slug=?,tm_url=?,staff_url=?,image_path=?,source_type=?,source_url=?
            WHERE country_id=? AND team_type=?""",
            (nname,short_name,tm_name,tmid,slug,tm_url,staff_url,image_path,source_type,source_url,cid,team_type))
        if team_type=='senior':
            # Transfermarkt identifies Serbia with UEFA Nations League A (UNLA), Group 2.
            # UEFA is the authoritative live source for the group table.
            cur.execute("""UPDATE national_teams SET
                standings_source_type='uefa',
                standings_source_url='https://www.uefa.com/uefanationsleague/standings/',
                standings_group_code='A2',
                standings_competition_code='UNLA',
                standings_competition_name='UEFA Nations League A',
                standings_reference_url='https://www.transfermarkt.com/uefa-nations-league-a/spieltag/pokalwettbewerb/UNLA/saison_id/2026/gruppe/2'
                WHERE country_id=? AND team_type='senior'""",(cid,))
        elif team_type=='u21':
            # Transfermarkt currently lists Serbia U21 under International Friendlies,
            # so there is no persistent league table to display.
            cur.execute("UPDATE national_teams SET standings_competition_name=COALESCE(NULLIF(standings_competition_name,''),'International Friendlies') WHERE country_id=? AND team_type='u21'",(cid,))
        else:
            cur.execute("UPDATE national_teams SET standings_source_type='fss' WHERE country_id=? AND team_type=? AND COALESCE(standings_source_type,'')=''",(cid,team_type))
    lid=cur.execute("SELECT id FROM leagues WHERE country_id=? AND name='SuperLiga Srbije'",(cid,)).fetchone()[0]
    if not cur.execute("SELECT id FROM clubs WHERE tm_club_id='159'").fetchone():
        cur.execute('INSERT INTO clubs(league_id,name,tm_name,cm_name,tm_club_id,tm_slug,tm_url,staff_url) VALUES(?,?,?,?,?,?,?,?)',(lid,'Red Star Belgrade','Roter Stern Belgrad','Red Star Belgrade','159','roter-stern-belgrad','https://www.transfermarkt.com/roter-stern-belgrad/startseite/verein/159','https://www.transfermarkt.com/roter-stern-belgrad/mitarbeiter/verein/159'))
    con.commit(); seed_serbia_clubs(con)
    con.execute("UPDATE clubs SET source_active=COALESCE(source_active,1)")
    con.execute("UPDATE clubs SET tm_match_status='matched' WHERE COALESCE(tm_club_id,'')<>''")
    con.execute("UPDATE leagues SET source_url=tm_url WHERE source_type='transfermarkt' AND COALESCE(source_url,'')=''")
    con.commit(); con.close()
    # History seed is idempotent and intentionally runs after the structural DB
    # migration. It provides verified champions/cup finals immediately while the
    # optional History Sync can later fill every final table row.
    try:
        from .history_sync import seed_history
        seed_history()
    except Exception as e:
        print('[HISTORY SEED WARNING]',e)
    # Excel History Master is authoritative for domestic league/cup honours.
    # Import only when the workbook hash changes; this is fast on normal startup.
    try:
        from .history_excel import import_master
        import_master(force=False)
    except Exception as e:
        print('[HISTORY EXCEL WARNING]',e)


def normalize_date(d):
    d=clean_text(d)
    m=re.fullmatch(rf'({MONTH_RE})\s+(\d{{1,2}}),\s+(\d{{4}})',d,re.I)
    if m: return f'{int(m.group(2)):02d}/{MONTHS[m.group(1).lower()[:3]]:02d}/{m.group(3)}'
    m=re.fullmatch(r'(\d{1,2})[./](\d{1,2})[./](\d{4})',d)
    if m: return f'{int(m.group(1)):02d}/{int(m.group(2)):02d}/{m.group(3)}'
    m=re.fullmatch(r'(\d{4})-(\d{2})-(\d{2})',d)
    if m: return f'{m.group(3)}/{m.group(2)}/{m.group(1)}'
    return d


def extract_date(text):
    text=clean_text(text)
    for pat in DATE_PATTERNS:
        m=pat.search(text)
        if m: return normalize_date(m.group(0))
    return ''

def extract_age(text):
    text=clean_text(text); m=AGE_RE.search(text)
    if m: return m.group(1)
    return text if re.fullmatch(r'\d{1,3}',text) else ''


def fetch_html(url,session=None):
    s=session or requests.Session(); r=s.get(url,headers=HEADERS,timeout=35)
    if r.status_code in (403,429): raise RuntimeError(f'Transfermarkt blocked request (HTTP {r.status_code}). Use saved HTML file.')
    r.raise_for_status()
    if 'captcha' in r.text.lower() or 'access denied' in r.text.lower(): raise RuntimeError('Transfermarkt returned captcha/access-denied page. Use saved HTML file.')
    return r.text


def parse_transfermarkt_league_clubs(html):
    soup=BeautifulSoup(html,'lxml'); found={}
    for a in soup.find_all('a',href=True):
        href=a.get('href',''); m=re.search(r'/([^/?#]+)/startseite/verein/(\d+)',href)
        if not m: continue
        slug,cid=m.group(1),m.group(2); name=clean_text(a.get_text(' ',strip=True))
        if not name:
            img=a.find('img'); name=clean_text((img.get('alt') or img.get('title')) if img else '')
        if name: found[cid]={'tm_name':name,'tm_club_id':cid,'tm_slug':slug,'tm_url':f'https://www.transfermarkt.com/{slug}/startseite/verein/{cid}','staff_url':f'https://www.transfermarkt.com/{slug}/mitarbeiter/verein/{cid}'}
    return list(found.values())


def match_seeded_club(existing,tm_name):
    n=norm_club_name(tm_name); alias=TM_CLUB_ALIASES.get(n)
    if alias:
        for c in existing:
            if (c['cm_name'] or c['name']).lower()==alias.lower(): return c
    for c in existing:
        for f in ('name','cm_name','tm_name'):
            if c[f] and norm_club_name(c[f])==n: return c
    nt=set(n.split()); best=None; best_score=0
    for c in existing:
        ct=set(norm_club_name(c['cm_name'] or c['name']).split())
        if not nt or not ct: continue
        score=len(nt&ct)/max(len(nt),len(ct))
        if score>best_score: best_score=score; best=c
    return best if best_score>=0.70 else None


def apply_transfermarkt_league_clubs(lid,parsed,insert_unmatched=False):
    con=db(); existing=con.execute('SELECT * FROM clubs WHERE league_id=? ORDER BY id',(lid,)).fetchall(); updated=0; unmatched=[]
    if insert_unmatched:
        con.execute('UPDATE clubs SET source_active=0 WHERE league_id=?',(lid,))
    for item in parsed:
        match=next((c for c in existing if clean_text(c['tm_club_id'])==clean_text(item['tm_club_id'])),None)
        if not match:
            match=match_seeded_club(existing,item['tm_name'])
        if match:
            con.execute('''UPDATE clubs SET tm_name=?,tm_club_id=?,tm_slug=?,tm_url=?,staff_url=?,source_name=?,source_active=1,
                tm_match_status='matched',tm_match_score=100 WHERE id=?''',
                (item['tm_name'],item['tm_club_id'],item['tm_slug'],item['tm_url'],item['staff_url'],item['tm_name'],match['id'])); updated+=1
        elif insert_unmatched:
            con.execute('''INSERT INTO clubs(league_id,name,tm_name,cm_name,tm_club_id,tm_slug,tm_url,staff_url,
                source_name,source_active,tm_match_status,tm_match_score) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)''',
                (lid,item['tm_name'],item['tm_name'],item['tm_name'],item['tm_club_id'],item['tm_slug'],item['tm_url'],item['staff_url'],
                 item['tm_name'],1,'matched',100)); updated+=1
        else:
            unmatched.append(item['tm_name'])
    con.commit(); con.close(); return updated,unmatched



def fetch_public_html(url, referer='https://srbijasport.net/'):
    headers=dict(HEADERS)
    headers['Referer']=referer
    r=requests.get(url,headers=headers,timeout=35)
    r.raise_for_status()
    return r.text



def fss_latin_url(url):
    # Force Latin-script FSS pages so names stay consistent with the local DB.
    u=clean_text(url)
    if not u:
        return ''
    if 'script=lat' in u:
        return u
    return u+('&' if '?' in u else '?')+'script=lat'


SERBIAN_CYR_MAP=str.maketrans({
    'А':'A','Б':'B','В':'V','Г':'G','Д':'D','Ђ':'Đ','Е':'E','Ж':'Ž','З':'Z','И':'I','Ј':'J','К':'K','Л':'L','Љ':'Lj','М':'M','Н':'N','Њ':'Nj','О':'O','П':'P','Р':'R','С':'S','Т':'T','Ћ':'Ć','У':'U','Ф':'F','Х':'H','Ц':'C','Ч':'Č','Џ':'Dž','Ш':'Š',
    'а':'a','б':'b','в':'v','г':'g','д':'d','ђ':'đ','е':'e','ж':'ž','з':'z','и':'i','ј':'j','к':'k','л':'l','љ':'lj','м':'m','н':'n','њ':'nj','о':'o','п':'p','р':'r','с':'s','т':'t','ћ':'ć','у':'u','ф':'f','х':'h','ц':'c','ч':'č','џ':'dž','ш':'š',
})


def serbian_to_latin(value):
    return clean_text(str(value or '').translate(SERBIAN_CYR_MAP))


def _fss_norm(value):
    s=serbian_to_latin(value).lower().replace('đ','dj')
    s=unicodedata.normalize('NFKD',s)
    s=''.join(ch for ch in s if not unicodedata.combining(ch))
    s=re.sub(r'[^a-z0-9]+',' ',s)
    return re.sub(r'\s+',' ',s).strip()


FSS_POSITION_HEADINGS={
    'golmani':'Goalkeepers','golman':'Goalkeepers',
    'odbrambeni igraci':'Defenders','odbrana':'Defenders',
    'vezni igraci':'Midfielders','vezni':'Midfielders',
    'napadaci':'Forwards','napad':'Forwards',
}
FSS_STAFF_HEADING={'strucni stab','staff'}
FSS_STAFF_ROLES={
    'selektor','trener','tim menadzer','trener golmana','kondicioni trener',
    'analiticar','fizioterapeut','doktor','ekonom','portparol','team manager',
    'assistant coach','coach','goalkeeping coach','fitness coach','physiotherapist','doctor','analyst',
}


def parse_fss_selection_html(html,base_url=''):
    # FSS selection pages expose positional groups followed by (current club),
    # then a STRUČNI ŠTAB block with name/role pairs. Parsing stripped strings
    # keeps this robust against minor theme markup changes.
    soup=BeautifulSoup(html,'lxml')
    lines=[serbian_to_latin(x) for x in soup.stripped_strings]
    lines=[x for x in lines if x]
    players=[]; staff=[]; pos=''; pending=''; in_staff=False; started=False
    staff_head={_fss_norm(x) for x in FSS_STAFF_HEADING}
    staff_roles={_fss_norm(x) for x in FSS_STAFF_ROLES}
    stop_staff_markers={_fss_norm(x) for x in {
        'nastupi za reprezentaciju','raspored i rezultati','rezultati','tabela','strelci',
        'arhiva','sponzori','takmicenja','takmičenja','savez','info','selekcije',
        'treneri','grassroots','sudije','media','medija','kontakt'
    }}
    for line in lines:
        nk=_fss_norm(line)
        if nk in FSS_POSITION_HEADINGS:
            started=True; in_staff=False; pending=''; pos=FSS_POSITION_HEADINGS[nk]; continue
        if nk in staff_head:
            started=True; in_staff=True; pending=''; pos=''; continue
        if not started:
            continue
        if in_staff:
            if nk in stop_staff_markers or nk.startswith('nastupi za reprezentaciju'):
                break
            if nk in staff_roles:
                if pending:
                    key=_fss_norm(pending)
                    staff.append({
                        'section':'National Team Staff','name':pending,'nationality':'Serbia',
                        'date_of_birth':'','age':'','appointed':'','contract_expires':'','last_club':'',
                        'profile_url':f'{base_url}#staff-{quote_plus(key)}','dob_source':'FSS','position':line,
                    })
                    pending=''
                continue
            if re.fullmatch(r'\d+[.:/-]?.*',line) or len(line)>90:
                continue
            pending=line
            continue

        if not pos:
            continue
        if line.startswith('(') and line.endswith(')') and pending:
            club=clean_text(line[1:-1])
            key=_fss_norm(pending); safe=game_safe_name(pending)
            players.append({
                'squad_number':'','name':pending,'tm_player_id':'','game_safe_name':safe,
                'safe_name_length':len(safe),'position':pos,'date_of_birth':'','age':'',
                'nationality':'Serbia','market_value':'','profile_url':f'{base_url}#player-{quote_plus(key)}',
                'current_club':club,
            })
            pending=''
        elif not (line.startswith('(') and line.endswith(')')):
            if nk not in ('selekcije','info','savez','treneri','grassroots') and len(line)<=80:
                pending=line

    pseen={}; sseen={}
    for r in players:
        pseen.setdefault(_fss_norm(r['name']),r)
    for r in staff:
        sseen.setdefault((_fss_norm(r['name']),_fss_norm(r['position'])),r)
    return list(pseen.values()),list(sseen.values())


def fss_match_report_links(html):
    soup=BeautifulSoup(html,'lxml'); out=[]; seen=set()
    for a in soup.find_all('a',href=True):
        href=clean_text(a.get('href'))
        if '/izvestaj-sa-utakmice/' not in href:
            continue
        u=fss_latin_url(urljoin('https://fss.rs/',href))
        if u not in seen:
            seen.add(u); out.append(u)
    return out


def parse_fss_match_report_rosters(html,report_url=''):
    # Deliberately extract rosters only. Match score/events are out of scope for
    # the current website phase.
    soup=BeautifulSoup(html,'lxml')
    title=soup.select_one('h1.fss-rez__title')
    teams=[serbian_to_latin(x.get_text(' ',strip=True)) for x in title.find_all('span')] if title else []
    if len(teams)<2:
        return []
    cols=soup.select('div.fss-rez__oneteam')
    if len(cols)<2:
        return []
    by_team=[[],[]]
    assignments=[(0,0),(1,1)]
    if len(cols)>=4:
        assignments.extend([(2,0),(3,1)])
    for ci,ti in assignments:
        for item in cols[ci].select('.fss-rez__start'):
            vals=[clean_text(x.get_text(' ',strip=True)) for x in item.select('.fss-rez__name span')]
            if not vals or not vals[0]:
                continue
            name=serbian_to_latin(vals[0]); number=vals[1] if len(vals)>1 else ''
            safe=game_safe_name(name); key=_fss_norm(name)
            by_team[ti].append({
                'squad_number':number,'name':name,'tm_player_id':'','game_safe_name':safe,
                'safe_name_length':len(safe),'position':'','date_of_birth':'','age':'',
                'nationality':'Serbia','market_value':'',
                'profile_url':f'{report_url}#player-{quote_plus(key)}',
            })
    result=[]
    for i in range(2):
        dedup={}
        for r in by_team[i]:
            k=_fss_norm(r['name'])
            if k not in dedup or (not dedup[k].get('squad_number') and r.get('squad_number')):
                dedup[k]=r
        result.append({'team_name':teams[i],'players':list(dedup.values())})
    return result


def _women_club_tokens(value):
    n=norm_club_name(value)
    return [t for t in n.split() if t not in {'zfk','fk','ofk','gfk','zgfk','zfu','zenski','zenska','fudbalski','klub'}]


def match_women_club(rows,source_name):
    st=_women_club_tokens(source_name)
    if not st:
        return None
    sset=set(st); best=None; best_score=0.0
    for c in rows:
        vals=[c['name'],c['cm_name'],c['source_name'],c['tm_name']]
        local_best=0.0
        for v in vals:
            if not v: continue
            tt=_women_club_tokens(v)
            if not tt: continue
            tset=set(tt)
            if ' '.join(st)==' '.join(tt):
                local_best=1.0; break
            local_best=max(local_best,len(sset&tset)/max(1,min(len(sset),len(tset))))
        if local_best>best_score:
            best_score=local_best; best=c
    return best if best_score>=0.72 else None


def save_fss_players_merge(club_id,records):
    # Merge FSS match-report players without deleting richer TM records.
    con=db()
    existing=con.execute('SELECT * FROM players WHERE club_id=?',(club_id,)).fetchall()
    by_name={_fss_norm(r['name']):r for r in existing}
    inserted=updated=0
    for r in records:
        k=_fss_norm(r.get('name'))
        if not k: continue
        old=by_name.get(k)
        if old:
            con.execute('''UPDATE players SET squad_number=CASE WHEN COALESCE(squad_number,'')='' THEN ? ELSE squad_number END,
                position=CASE WHEN COALESCE(position,'')='' THEN ? ELSE position END,
                profile_url=CASE WHEN COALESCE(profile_url,'')='' THEN ? ELSE profile_url END,
                updated_at=CURRENT_TIMESTAMP WHERE id=?''',
                (r.get('squad_number'),r.get('position'),r.get('profile_url'),old['id']))
            updated+=1
        else:
            con.execute('''INSERT INTO players(club_id,squad_number,name,tm_player_id,game_safe_name,safe_name_length,position,
                date_of_birth,age,nationality,market_value,profile_url,updated_at)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?,CURRENT_TIMESTAMP)''',
                (club_id,r.get('squad_number'),r.get('name'),'',r.get('game_safe_name'),r.get('safe_name_length'),
                 r.get('position'),r.get('date_of_birth'),r.get('age'),r.get('nationality'),r.get('market_value'),r.get('profile_url')))
            inserted+=1
    con.execute('UPDATE clubs SET last_players_update=CURRENT_TIMESTAMP WHERE id=?',(club_id,))
    con.execute('''INSERT INTO update_log(club_id,module,source,rows_imported,status,details)
        VALUES(?,?,?,?,?,?)''',(club_id,'players','FSS match reports',len(records),'OK',f'inserted {inserted}; merged {updated}'))
    con.commit(); con.close()
    return {'inserted':inserted,'updated':updated,'rows':len(records)}


def update_women_league_fss_rosters(lid):
    con=db()
    league=con.execute("SELECT * FROM leagues WHERE id=? AND section='women'",(lid,)).fetchone()
    clubs=con.execute('SELECT * FROM clubs WHERE league_id=? AND COALESCE(source_active,1)=1',(lid,)).fetchall()
    con.close()
    if not league:
        raise RuntimeError('Women league not found.')
    if not league['fss_url']:
        return {'reports':0,'clubs':0,'players':0,'club_ids':[],'errors':[]}
    league_html=fetch_public_html(fss_latin_url(league['fss_url']),'https://fss.rs/')
    links=fss_match_report_links(league_html)
    union={}; club_ids=set(); errors=[]
    for idx,u in enumerate(links,1):
        try:
            html=fetch_public_html(u,'https://fss.rs/')
            for team in parse_fss_match_report_rosters(html,u):
                club=match_women_club(clubs,team['team_name'])
                if not club:
                    errors.append(f"Unmatched FSS team: {team['team_name']}")
                    continue
                club_ids.add(club['id'])
                bucket=union.setdefault(club['id'],{})
                for r in team['players']:
                    k=_fss_norm(r['name'])
                    if k not in bucket or (not bucket[k].get('squad_number') and r.get('squad_number')):
                        bucket[k]=r
        except Exception as e:
            errors.append(f'{u}: {e}')
        if idx<len(links):
            time.sleep(0.15)
    total=0
    for club_id,vals in union.items():
        records=list(vals.values()); total+=len(records)
        save_fss_players_merge(club_id,records)
    return {'reports':len(links),'clubs':len(union),'players':total,'club_ids':list(club_ids),'errors':errors}


def update_women_fss_rosters_country(cid):
    con=db()
    leagues=con.execute('''SELECT id,name FROM leagues WHERE country_id=? AND section='women'
        AND COALESCE(fss_url,'')<>'' ORDER BY level,name''',(cid,)).fetchall()
    con.close()
    out={'leagues':len(leagues),'ok':0,'reports':0,'clubs':0,'players':0,'errors':[]}
    for l in leagues:
        try:
            r=update_women_league_fss_rosters(l['id'])
            out['ok']+=1; out['reports']+=r['reports']; out['clubs']+=r['clubs']; out['players']+=r['players']
            out['errors'].extend(f"{l['name']}: {e}" for e in r['errors'])
        except Exception as e:
            out['errors'].append(f"{l['name']}: {e}")
    return out

def clean_srbijasport_team_name(value):
    s=clean_text(value)
    # Team cards append current rank and, for promoted/relegated clubs, the name
    # of another competition. Strip that annotation but preserve years in names.
    league_words=(r'Mozzart|Superliga|Super liga|Prva|Druga|Treća|Srpska|Vojvođanska|Vojvodjanska|Zona|Beogradska|'
                  r'Okružna|okružna|Opštinska|opštinska|Međuopštinska|međuopštinska|Gradska|gradska|'
                  r'PFL|MFL|GFL|UOFL|Šumadijska|Sremska|Potiska')
    s=re.sub(rf'\s+\d{{1,2}}\s+(?:{league_words})\b.*$','',s,flags=re.I)
    s=re.sub(r'\s+\d{1,2}\s*$','',s)
    s=re.sub(r'\s+(?:Superliga|Super liga|Prva liga|Druga liga)\b.*(?:Play\s*[- ]?Out|Play\s*[- ]?Off).*$','',s,flags=re.I)
    s=clean_text(s)
    toks=s.split()
    for n in (3,2,1):
        if len(toks)>=2*n and [x.lower() for x in toks[-n:]]==[x.lower() for x in toks[-2*n:-n]]:
            toks=toks[:-n]; s=' '.join(toks); break
    return clean_text(s)


def parse_srbijasport_league_clubs(html, expected_count=0):
    """Parse the current SrbijaSport membership from the dedicated Timovi block.

    This avoids accidentally importing clubs mentioned later in fixtures,
    standings, higher/lower-rank links or news. expected_count remains supported
    for older configured levels but Level 5 intentionally uses the live block size.
    """
    soup=BeautifulSoup(html,'lxml')

    def club_item(a):
        href=clean_text(a.get('href') or '')
        if not re.search(r'(?:^|/)club/\d+(?:-|$)',href):
            return None
        full=urljoin('https://srbijasport.net/',href)
        m=re.search(r'/club/(\d+)(?:-([^/?#]+))?',full)
        if not m:
            return None
        name=clean_srbijasport_team_name(a.get_text(' ',strip=True))
        if not name:
            img=a.find('img')
            if img:
                name=clean_srbijasport_team_name(img.get('alt') or img.get('title') or '')
        if not name:
            name=clean_text((m.group(2) or '').replace('-',' '))
        if len(name)<2:
            return None
        return {'source_id':m.group(1),'source_name':name,'source_club_url':full}

    # SrbijaSport renders an explicit Timovi section followed by "Kola Dani...".
    # Try every exact Timovi text node and retain the candidate block containing
    # the most unique club links; this is robust to the Timovi label in tab nav.
    candidates=[]
    for txt in soup.find_all(string=lambda x: clean_text(x)=='Timovi' if x else False):
        found=[]; seen=set()
        for el in txt.parent.next_elements:
            if isinstance(el,str):
                t=clean_text(el)
                if found and (t=='Kola' or t.startswith('Kola Dani')):
                    break
                continue
            if getattr(el,'name',None)!='a':
                continue
            item=club_item(el)
            if not item or item['source_id'] in seen:
                continue
            seen.add(item['source_id']); found.append(item)
            if expected_count and len(found)>=expected_count:
                break
        if found:
            candidates.append(found)

    if candidates:
        found=max(candidates,key=len)
        if expected_count:
            return found[:expected_count]
        return found

    # Conservative fallback for an older/changed page structure.
    found=[]; seen=set()
    for a in soup.find_all('a',href=True):
        item=club_item(a)
        if not item or item['source_id'] in seen:
            continue
        seen.add(item['source_id']); found.append(item)
        if expected_count and len(found)>=expected_count:
            break
    return found


def _norm_comp_name(value):
    s=unicodedata.normalize('NFKD',clean_text(value)).encode('ascii','ignore').decode('ascii').lower()
    s=s.replace('“','"').replace('”','"').replace('„','"')
    return re.sub(r'[^a-z0-9]+',' ',s).strip()


def _region_for_discovered_league(name):
    if name in REGION_LEAGUES:
        return REGION_LEAGUES[name]
    if name in LEVEL6_LEAGUES:
        return LEVEL6_LEAGUES[name]
    n=_norm_comp_name(name)
    # Fallback for newly discovered leagues not yet present in the retained
    # catalog. This affects only navigation/grouping; league identity remains
    # the SrbijaSport competition ID.
    keyword_groups=(
        ('FSB',('beograd','obrenovac','lazarevac','mladenovac','grocka','sopot')),
        ('FSV',('vojvod','backa','sombor','subotica','srem','novosad','novi sad','zrenjanin',
                'kikinda','zitiste','potiska','juznobanat','vrbas','becej','srbobran','pancevo')),
        ('WEST',('kolubar','macva','pocerin','sabac','cacak','moravic','raska','zlatibor','sumadij',
                 'kragujevac','kraljevo','uzice','loznica','valjevo','drina')),
        ('EAST',('nis','pirot','bor','zajecar','toplic','jablan','pomorav','branicev','podunav',
                 'rasin','krusevac','aleksinac','smederevo','velika plana','mlava','dunav','istok','jug')),
    )
    for code,words in keyword_groups:
        if any(word in n for word in words):
            return code
    return ''


def discover_srbijasport_pyramid(country_id=1, sync_memberships=False, levels=(3,4,5,6)):
    """Discover/update the current lower senior pyramid from SrbijaSport.

    The live competition index is authoritative for Levels 3-6. Leagues which
    existed in a previous season are retained for history/reference but marked
    source_active=0 so they do not leak into the current portal or sync queue.
    """
    html=fetch_public_html('https://srbijasport.net/sport/1/takmicenja','https://srbijasport.net/')
    grouped=discover_srbijasport_index(html)
    wanted={int(x) for x in levels if int(x) in (3,4,5,6)}
    out={'discovered':0,'inserted':0,'updated':0,'deactivated':0,'synced':0,'clubs':0,'levels':{},'errors':[]}
    con=db()
    try:
        league_cols={r[1] for r in con.execute('PRAGMA table_info(leagues)').fetchall()}
        if 'source_active' not in league_cols:
            con.execute('ALTER TABLE leagues ADD COLUMN source_active INTEGER DEFAULT 1')
            con.commit()
        # First retire the previous live snapshot for only the requested lower
        # levels. Discovered rows below are reactivated. Nothing is deleted.
        for level in sorted(wanted):
            cur=con.execute("""UPDATE leagues SET source_active=0
                WHERE country_id=? AND level=? AND COALESCE(section,'senior')='senior'
                AND LOWER(COALESCE(source_type,''))='srbijasport' AND COALESCE(source_active,1)=1""",
                (country_id,level))
            out['deactivated']+=max(0,int(cur.rowcount or 0))

        for level in sorted(wanted):
            items=grouped.get(level,[])
            out['levels'][level]=len(items)
            for item in items:
                out['discovered']+=1
                row=con.execute("""SELECT * FROM leagues WHERE country_id=?
                    AND COALESCE(section,'senior')='senior' AND level BETWEEN 3 AND 6
                    AND source_url LIKE ?""",
                    (country_id,f'%/league/{item["source_id"]}%')).fetchone()
                if not row:
                    row=con.execute("""SELECT * FROM leagues WHERE country_id=? AND name=?
                        AND COALESCE(section,'senior')='senior' AND level BETWEEN 3 AND 6""",
                        (country_id,item['name'])).fetchone()
                region=_region_for_discovered_league(item['name'])
                if row:
                    con.execute("""UPDATE leagues SET name=?,level=?,section='senior',source_type='srbijasport',
                        source_url=?,source_active=1,
                        standings_source_type='srbijasport',standings_source_url=?,
                        region_code=CASE WHEN COALESCE(region_code,'')='' THEN ? ELSE region_code END
                        WHERE id=?""",
                        (item['name'],level,item['url'],item['url'],region,row['id']))
                    item['db_id']=int(row['id']); out['updated']+=1
                else:
                    cur=con.execute("""INSERT INTO leagues(country_id,name,tm_name,level,source_type,source_url,
                        source_expected_clubs,section,region_code,standings_source_type,standings_source_url,source_active)
                        VALUES(?,?,?,?,?,?,?,?,?,?,?,1)""",
                        (country_id,item['name'],'',level,'srbijasport',item['url'],0,'senior',region,'srbijasport',item['url']))
                    item['db_id']=int(cur.lastrowid); out['inserted']+=1
        con.commit()
        # Report how many rows remained retired after the fresh snapshot.
        if wanted:
            marks=','.join('?' for _ in wanted)
            q="""SELECT COUNT(*) FROM leagues WHERE country_id=? AND level IN ("""+marks+""")
                AND COALESCE(section,'senior')='senior' AND LOWER(COALESCE(source_type,''))='srbijasport'
                AND COALESCE(source_active,1)=0"""
            retired=int(con.execute(q,[country_id,*sorted(wanted)]).fetchone()[0] or 0)
            out['deactivated']=retired
    finally:
        con.close()
    if sync_memberships:
        for level in sorted(wanted):
            for item in grouped.get(level,[]):
                try:
                    page=fetch_public_html(item['url'],'https://srbijasport.net/')
                    records=parse_srbijasport_league_clubs(page,0)
                    result=sync_srbijasport_league(item['db_id'],records)
                    out['synced']+=1
                    out['clubs']+=int(result.get('active') or len(records))
                except Exception as e:
                    out['errors'].append(f"{item['name']}: {e}")
    return out


def discover_level6_srbijasport(country_id=1, sync_memberships=True):
    """Compatibility wrapper; Level 6 is now discovered from the live index."""
    r=discover_srbijasport_pyramid(country_id,sync_memberships,levels=(6,))
    n=r.get('levels',{}).get(6,0)
    return {'expected':n,'discovered':n,'synced':r.get('synced',0),'clubs':r.get('clubs',0),
            'missing':[],'errors':r.get('errors',[]),'inserted':r.get('inserted',0),'updated':r.get('updated',0)}


def _int_cell(value):
    m=re.search(r'-?\d+',clean_text(value))
    return int(m.group(0)) if m else None


def _clean_srbijasport_team_label(value):
    """Split SrbijaSport zone/status decorations from the actual club label."""
    raw=clean_text(value)
    note=[]
    # Symbols are often rendered directly in the club/status cell.
    if raw.startswith(('▲','▼')):
        note.append(raw[0]); raw=clean_text(raw[1:])
    patterns=[
        r'^Doigravanje\s*\(Play\s*Off\)\s*',
        r'^Play\s*Off\s*',
        r'^Play\s*Out\s*',
        r'^Prelazi\s+u\s+ligu\s+nižeg\s+ranga\s*',
        r'^Prelazi\s+u\s+ligu\s+viseg\s+ranga\s*',
        r'^Prelazi\s+u\s+ligu\s+višeg\s+ranga\s*',
        r'^Liga\s+višeg\s+ranga\s*',
        r'^Liga\s+viseg\s+ranga\s*',
        r'^Baraž\s*', r'^Baraz\s*',
    ]
    changed=True
    while changed:
        changed=False
        for pat in patterns:
            m=re.match(pat,raw,flags=re.I)
            if m:
                label=clean_text(m.group(0)).strip()
                if label: note.append(label)
                raw=clean_text(raw[m.end():]); changed=True; break
        if raw.startswith(('▲','▼')):
            note.append(raw[0]); raw=clean_text(raw[1:]); changed=True
    return raw, ' '.join(note).strip()


def _srbijasport_row_from_segment(position, position_note, segment):
    """Parse one standings row from tokens between two sequential rank markers.

    Team names may themselves contain numbers (PKB 1950, BSK 1926, Timok 1919),
    therefore stats are parsed from the RIGHT side of the row, never by stopping
    on the first numeric token.
    """
    vals=[clean_text(x) for x in segment if clean_text(x)]
    if not vals: return None
    # Find the trailing run of integer stats. SrbijaSport commonly exposes either
    # 5 values (O,P,N,I,B) or 7/8 values when goals are also present.
    trailing=[]
    j=len(vals)-1
    while j>=0 and re.fullmatch(r'-?\d+',vals[j]):
        trailing.append(int(vals[j])); j-=1
    trailing=list(reversed(trailing))
    if len(trailing)<5:
        return None
    team_raw=clean_text(' '.join(vals[:j+1]))
    team,note2=_clean_srbijasport_team_label(team_raw)
    if not team: return None
    note=' '.join(x for x in (position_note,note2) if x).strip()
    played,wins,draws,losses=trailing[:4]
    points=trailing[-1]
    gf=ga=gd=None
    extra=trailing[4:-1]
    if len(extra)>=2:
        gf,ga=extra[0],extra[1]
        gd=(gf-ga) if gf is not None and ga is not None else None
    if len(extra)>=3:
        gd=extra[2]
    return {
        'position':position,'position_note':note,'team_name':team,
        'played':played,'wins':wins,'draws':draws,'losses':losses,
        'goals_for':gf,'goals_against':ga,'goal_diff':gd,'points':points,
    }


def _parse_srbijasport_standings_anchors(soup):
    """Parse SrbijaSport standings by club anchors inside the Tabela section.

    This is intentionally independent of TM mappings. It uses /club/<id> links
    from the source table and reads stats from the smallest row-like parent.
    """
    candidates=[]
    table_labels=soup.find_all(string=lambda x: _norm_comp_name(clean_text(x))=='tabela' if x else False)
    for label in table_labels:
        rows=[]; seen_pos=set(); seen_clubs=set()
        for el in label.parent.next_elements:
            if isinstance(el,str):
                nt=_norm_comp_name(clean_text(el))
                if rows and nt in ('strelci','statistika'):
                    break
                continue
            if getattr(el,'name',None)!='a': continue
            href=clean_text(el.get('href') or '')
            if not re.search(r'(?:^|/)club/\d+(?:-|$)',href): continue
            team=clean_srbijasport_team_name(el.get_text(' ',strip=True))
            if not team: continue
            rowbox=None; strings=None; anc=el
            for _ in range(8):
                anc=getattr(anc,'parent',None)
                if anc is None: break
                vals=[clean_text(x) for x in anc.stripped_strings if clean_text(x)]
                if len(vals)>40: continue
                nums=[v for v in vals if re.fullmatch(r'-?\d+',v)]
                # rank + at least O/P/N/I/B
                if len(nums)>=6 and any(team in v or v in team for v in vals):
                    rowbox=anc; strings=vals; break
            if rowbox is None or not strings: continue
            # Locate team token; anchor text normally stays a single stripped string.
            ti=next((i for i,v in enumerate(strings) if clean_text(v)==team),None)
            if ti is None:
                ti=next((i for i,v in enumerate(strings) if team in clean_text(v)),None)
            if ti is None: continue
            before=strings[:ti]; after=strings[ti+1:]
            rank=None; note_parts=[]
            for v in reversed(before):
                m=re.match(r'^\s*(\d{1,2})\.?\s*(.*)$',v)
                if m:
                    rank=int(m.group(1)); rem=clean_text(m.group(2)).strip(' .:-')
                    if rem: note_parts.insert(0,rem)
                    break
                nv=clean_text(v)
                if nv and len(nv)<80: note_parts.insert(0,nv)
            if not rank or rank in seen_pos: continue
            nums=[]
            for v in after:
                if re.fullmatch(r'-?\d+',v): nums.append(int(v))
            if len(nums)<5: continue
            # The source may expose O/P/N/I/B only, or O/P/N/I/DG/PG[/GR]/B.
            played,wins,draws,losses=nums[:4]; points=nums[-1]
            extra=nums[4:-1]; gf=ga=gd=None
            if len(extra)>=2:
                gf,ga=extra[0],extra[1]; gd=gf-ga
            if len(extra)>=3: gd=extra[2]
            team2,note2=_clean_srbijasport_team_label(team)
            note=' '.join(x for x in (' '.join(note_parts),note2) if x).strip()
            rows.append({'position':rank,'position_note':note,'team_name':team2,
                         'played':played,'wins':wins,'draws':draws,'losses':losses,
                         'goals_for':gf,'goals_against':ga,'goal_diff':gd,'points':points})
            seen_pos.add(rank); seen_clubs.add(team2)
        if rows:
            rows.sort(key=lambda r:r['position'])
            candidates.append(rows)
    return max(candidates,key=len) if candidates else []


def _parse_srbijasport_standings_text(soup):
    """Fallback for SrbijaSport div/list standings.

    Robust against zone labels attached to rank cells and numeric club names.
    Rows are segmented by SEQUENTIAL rank markers (1,2,3,...) and statistics are
    parsed from the right edge, so club names such as PKB 1950 are preserved.
    """
    tokens=[clean_text(x) for x in soup.stripped_strings]
    tokens=[x for x in tokens if x]
    norm=[_norm_comp_name(x) for x in tokens]
    starts=[]
    for i,x in enumerate(norm):
        if x!='tabela': continue
        window=norm[i+1:i+20]
        need=('poz','tim','utak','pob','ner','por','bod')
        found=[]; cur=0
        for n in need:
            hit=None
            for j in range(cur,len(window)):
                if window[j]==n or window[j].startswith(n+' '): hit=j; break
            if hit is None:
                found=[]; break
            found.append(hit); cur=hit+1
        if found: starts.append(i+1+found[-1]+1)
    best=[]
    stop_words=('dodajte tabelu na vas sajt','strelci','statistika','kola','utakmice','detaljnije')
    for st in starts:
        # Locate sequential positions. Requiring 1->2->3 prevents years in club
        # names (1919/1926/1950) from being mistaken for rank markers.
        markers=[]; cursor=st; expected=1
        while cursor<len(tokens) and expected<=64:
            found=None; pnote=''
            for k in range(cursor,min(len(tokens),cursor+120)):
                nt=_norm_comp_name(tokens[k])
                if expected>1 and any(nt==w or nt.startswith(w+' ') for w in stop_words):
                    break
                m=re.match(r'^\s*(\d{1,2})\.?\s*(.*)$',tokens[k])
                if m and int(m.group(1))==expected:
                    found=k; pnote=clean_text(m.group(2)).strip(' .:-'); break
            if found is None: break
            markers.append((expected,found,pnote)); cursor=found+1; expected+=1
        parsed=[]
        for idx,(pos,mi,pnote) in enumerate(markers):
            end=markers[idx+1][1] if idx+1<len(markers) else min(len(tokens),mi+80)
            segment=tokens[mi+1:end]
            # Trim obvious post-table sections for the final row.
            cut=len(segment)
            for q,v in enumerate(segment):
                nv=_norm_comp_name(v)
                if any(nv==w or nv.startswith(w+' ') for w in stop_words): cut=q; break
            row=_srbijasport_row_from_segment(pos,pnote,segment[:cut])
            if row: parsed.append(row)
        if len(parsed)>len(best): best=parsed
    return best

def parse_standings_html(html):
    """Parse a league table from FSS or SrbijaSport HTML."""
    soup=BeautifulSoup(html,'lxml')
    best=[]
    for table in soup.find_all('table'):
        headers=[clean_text(x.get_text(' ',strip=True)) for x in table.find_all('th')]
        norm=[_norm_comp_name(x) for x in headers]
        if not headers: continue
        def idx(*names):
            for n in names:
                nn=_norm_comp_name(n)
                for i,h in enumerate(norm):
                    if h==nn or h.startswith(nn+' '): return i
            return None
        ipos=idx('Poz','P','Pozicija'); iteam=idx('Tim','Naziv kluba','Klub')
        ipl=idx('Utak','OU','Odigrane utakmice'); iw=idx('Pob','BP','Broj pobeda')
        idr=idx('Ner','BN','Broj nerešenih'); il=idx('Por','BI','Broj izgubljenih')
        igf=idx('DG','Dati golovi'); iga=idx('PG','Primljeni golovi'); igd=idx('GR','Gol razlika'); ipts=idx('Bod','B','Bodovi')
        if iteam is None or ipts is None: continue
        parsed=[]
        for tr in table.find_all('tr'):
            cells=[clean_text(x.get_text(' ',strip=True)) for x in tr.find_all(['td','th'])]
            if not cells or len(cells)<=max(iteam,ipts): continue
            team=cells[iteam]
            if not team or _norm_comp_name(team) in ('tim','naziv kluba','klub'): continue
            pos_raw=cells[ipos] if ipos is not None and ipos<len(cells) else str(len(parsed)+1)
            pm=re.match(r'\s*(\d+)\s*(.*)',pos_raw)
            pos=int(pm.group(1)) if pm else len(parsed)+1; pnote=clean_text(pm.group(2)).strip('.:- ') if pm else ''
            if ipos is not None and iteam is not None and iteam>ipos+1:
                mid=' '.join(clean_text(x) for x in cells[ipos+1:iteam] if clean_text(x)).strip()
                if mid: pnote=(' '.join(x for x in (pnote,mid) if x)).strip()
            team,note2=_clean_srbijasport_team_label(team)
            if note2: pnote=(' '.join(x for x in (pnote,note2) if x)).strip()
            gf=_int_cell(cells[igf]) if igf is not None and igf<len(cells) else None
            ga=_int_cell(cells[iga]) if iga is not None and iga<len(cells) else None
            gd=_int_cell(cells[igd]) if igd is not None and igd<len(cells) else None
            if gd is None and gf is not None and ga is not None: gd=gf-ga
            parsed.append({
                'position':pos,'position_note':pnote,'team_name':team,
                'played':_int_cell(cells[ipl]) if ipl is not None and ipl<len(cells) else None,
                'wins':_int_cell(cells[iw]) if iw is not None and iw<len(cells) else None,
                'draws':_int_cell(cells[idr]) if idr is not None and idr<len(cells) else None,
                'losses':_int_cell(cells[il]) if il is not None and il<len(cells) else None,
                'goals_for':gf,'goals_against':ga,'goal_diff':gd,
                'points':_int_cell(cells[ipts]),
            })
        if len(parsed)>len(best): best=parsed
    if best:
        return best
    anchored=_parse_srbijasport_standings_anchors(soup)
    if anchored:
        return anchored
    return _parse_srbijasport_standings_text(soup)

_STANDINGS_CANONICAL_ALIASES={
    'crvena zvezda':'red star belgrade', 'crvena zvezda beograd':'red star belgrade',
    'partizan':'partizan belgrade', 'partizan beograd':'partizan belgrade',
    'vojvodina':'vojvodina novi sad', 'vojvodina novi sad':'vojvodina novi sad',
    'zeleznicar':'zeleznicar pancevo', 'zeleznicar pancevo':'zeleznicar pancevo',
    'imt':'imt novi beograd', 'imt mrbit':'imt novi beograd', 'imt novi beograd':'imt novi beograd',
    'radnicki 1923':'radnicki 1923', 'radnicki 1923 kragujevac':'radnicki 1923',
    'cukaricki':'cukaricki', 'cukaricki doo':'cukaricki', 'cukaricki beograd':'cukaricki',
    'mladost':'mladost lucani', 'mladost lucani':'mladost lucani',
    'radnik':'radnik surdulica', 'radnik surdulica':'radnik surdulica',
    'novi pazar':'novi pazar', 'ofk beograd':'ofk beograd',
    'macva':'macva sabac', 'macva sabac':'macva sabac',
    'radnicki':'radnicki nis', 'radnicki nis':'radnicki nis',
    'zemun':'zemun',
    'tsc':'tsc backa topola', 'tsc backa topola':'tsc backa topola',
    'napredak':'napredak krusevac', 'napredak krusevac':'napredak krusevac',
    'spartak':'spartak subotica', 'spartak subotica':'spartak subotica',
    'metalac':'metalac gornji milanovac', 'metalac gornji milanovac':'metalac gornji milanovac',
    'jedinstvo':'jedinstvo ub', 'jedinstvo ub':'jedinstvo ub',
}

def _standing_match_club(con,league_id,team_name):
    wanted=norm_club_name(re.sub(r'\s+-\d+\s*$','',team_name or ''))
    if not wanted: return None
    wanted=_STANDINGS_CANONICAL_ALIASES.get(wanted,wanted)
    rows=con.execute("SELECT id,name,cm_name,source_name FROM clubs WHERE league_id=?",(league_id,)).fetchall()
    exact=[]
    for r in rows:
        vals={norm_club_name(r['name'] or ''),norm_club_name(r['cm_name'] or ''),norm_club_name(r['source_name'] or '')}
        vals={_STANDINGS_CANONICAL_ALIASES.get(v,v) for v in vals if v}
        if wanted in vals: exact.append(r['id'])
    return exact[0] if len(exact)==1 else None


def save_standings(owner_type,owner_id,name,rows,source_type,source_url,country_id=1,season='2026/2027',group_name=''):
    con=db()
    old=con.execute("SELECT id FROM standings_competitions WHERE owner_type=? AND owner_id=? AND name=? AND season=? AND COALESCE(group_name,'')=?",
                    (owner_type,owner_id,name,season,group_name or '')).fetchone()
    if old:
        comp_id=old['id']; con.execute('DELETE FROM standings_rows WHERE competition_id=?',(comp_id,))
        con.execute('UPDATE standings_competitions SET source_type=?,source_url=?,updated_at=CURRENT_TIMESTAMP WHERE id=?',(source_type,source_url,comp_id))
    else:
        cur=con.execute('INSERT INTO standings_competitions(country_id,owner_type,owner_id,name,season,group_name,source_type,source_url) VALUES(?,?,?,?,?,?,?,?)',
                        (country_id,owner_type,owner_id,name,season,group_name or '',source_type,source_url))
        comp_id=cur.lastrowid
    for r in rows:
        club_id=_standing_match_club(con,owner_id,r['team_name']) if owner_type=='league' else None
        con.execute('''INSERT INTO standings_rows(competition_id,position,position_note,team_name,club_id,played,wins,draws,losses,goals_for,goals_against,goal_diff,points,form,note)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
            (comp_id,r.get('position'),r.get('position_note',''),r.get('team_name',''),club_id,r.get('played'),r.get('wins'),r.get('draws'),r.get('losses'),r.get('goals_for'),r.get('goals_against'),r.get('goal_diff'),r.get('points'),r.get('form',''),r.get('note','')))
    con.commit(); con.close(); return {'competition_id':comp_id,'rows':len(rows)}


def _srbijasport_embed_standings_url(url):
    """Return SrbijaSport's dedicated standings embed URL for a league page."""
    m=re.search(r'/league/(\d+)', clean_text(url))
    return f'https://srbijasport.net/embed/standings/{m.group(1)}' if m else ''


def _validate_complete_standings(rows, expected=0):
    """Reject partial/corrupted tables before replacing the last good snapshot."""
    if not rows:
        raise RuntimeError('Tabela nije pronađena na izvornoj stranici.')
    positions=[int(r.get('position') or 0) for r in rows]
    if len(set(positions))!=len(positions) or min(positions or [0])!=1:
        raise RuntimeError(f'Tabela je nepotpuna ili ima duple pozicije: {positions}.')
    maxpos=max(positions)
    missing=[x for x in range(1,maxpos+1) if x not in set(positions)]
    if missing:
        raise RuntimeError(f'Tabela je nepotpuna; nedostaju pozicije: {missing}.')
    if expected and len(rows)<int(expected):
        raise RuntimeError(f'Tabela je nepotpuna: pronađeno {len(rows)} od očekivanih {expected} klubova.')
    # Basic sanity: no negative match counts and core W/D/L fields must exist.
    for r in rows:
        vals=(r.get('played'),r.get('wins'),r.get('draws'),r.get('losses'))
        if any(v is None for v in vals) or any(int(v)<0 for v in vals):
            raise RuntimeError(f'Neispravne statistike u redu {r.get("position")}: {vals}.')
    return rows

def refresh_league_standings(league_id):
    con=db(); l=con.execute('SELECT * FROM leagues WHERE id=?',(league_id,)).fetchone()
    active_count=con.execute('SELECT COUNT(*) FROM clubs WHERE league_id=? AND COALESCE(source_active,1)=1',(league_id,)).fetchone()[0] if l else 0
    con.close()
    if not l: raise RuntimeError('Liga nije pronađena.')
    source=(l['standings_source_type'] or '').lower() or ('srbijasport' if int(l['level'] or 0)>=3 else 'fss')
    url=l['standings_source_url'] or (l['fss_url'] if source=='fss' else l['source_url'])
    if not url: raise RuntimeError('Liga nema konfigurisan izvor tabele.')
    rows=[]; fetched_url=url; errors=[]
    if source=='srbijasport':
        embed=_srbijasport_embed_standings_url(url)
        if embed:
            try:
                html=fetch_public_html(embed,'https://srbijasport.net/')
                rows=parse_standings_html(html)
                fetched_url=embed
            except Exception as e:
                errors.append(f'embed: {e}')
        # Fallback to the normal league page only when embed did not yield a table.
        if not rows:
            try:
                html=fetch_public_html(url,'https://srbijasport.net/')
                rows=parse_standings_html(html)
                fetched_url=url
            except Exception as e:
                errors.append(f'league page: {e}')
    else:
        html=fetch_public_html(url,'https://fss.rs/')
        rows=parse_standings_html(html)
    expected=int(l['source_expected_clubs'] or 0) or int(active_count or 0)
    try:
        _validate_complete_standings(rows,expected)
    except Exception as e:
        suffix=(' | '.join(errors)) if errors else ''
        raise RuntimeError(str(e)+(f' ({suffix})' if suffix else ''))
    result=save_standings('league',league_id,l['name'],rows,source,url,l['country_id'])
    result.update({'league':l['name'],'source':source,'fetched_url':fetched_url}); return result

def refresh_country_standings(country_id=1):
    con=db(); leagues=con.execute("SELECT id,name FROM leagues WHERE country_id=? AND COALESCE(source_active,1)=1 AND COALESCE(standings_source_url,'')<>'' ORDER BY section,level,name",(country_id,)).fetchall(); con.close()
    out={'leagues':len(leagues),'ok':0,'rows':0,'errors':[]}
    for l in leagues:
        try:
            r=refresh_league_standings(l['id']); out['ok']+=1; out['rows']+=r['rows']
        except Exception as e: out['errors'].append(f"{l['name']}: {e}")
    return out


def standings_for_owner(owner_type,owner_id):
    con=db(); comp=con.execute('SELECT * FROM standings_competitions WHERE owner_type=? AND owner_id=? ORDER BY updated_at DESC,id DESC LIMIT 1',(owner_type,owner_id)).fetchone()
    rows=con.execute('SELECT * FROM standings_rows WHERE competition_id=? ORDER BY position,id',(comp['id'],)).fetchall() if comp else []
    con.close(); return comp,rows


def parse_srbijasport_club_info(html,source_url=''):
    soup=BeautifulSoup(html,'lxml')
    lines=[clean_text(x) for x in soup.stripped_strings if clean_text(x)]
    def after(label):
        nl=_norm_comp_name(label)
        for i,v in enumerate(lines):
            if _norm_comp_name(v)==nl and i+1<len(lines):
                nxt=lines[i+1]
                if _norm_comp_name(nxt) not in ('detaljnije','odigrano','planirano','tabela'): return nxt
        return ''
    title=''
    for tag in soup.find_all(['h1','h2']):
        v=clean_text(tag.get_text(' ',strip=True))
        if v and len(v)<100: title=v; break
    founded=after('Osnovan'); website=after('Web adresa'); address=after('Adresa')
    stadium_name=after('Stadion')
    if _norm_comp_name(stadium_name)=='image stadion': stadium_name=''
    crest=''
    # SrbijaSport club crests use a stable public endpoint /img/club/<club_id>/500.
    # Prefer this deterministic URL over social/og images, which can be generic.
    mclub=re.search(r'/club/(\d+)', source_url or '')
    if mclub:
        crest=f'https://srbijasport.net/img/club/{mclub.group(1)}/500'
    if not crest:
        og=soup.find('meta',attrs={'property':'og:image'})
        if og and og.get('content'): crest=urljoin(source_url or 'https://srbijasport.net/',og.get('content'))
    if not crest:
        for img in soup.find_all('img',src=True):
            alt=clean_text(img.get('alt') or img.get('title') or '')
            src=img.get('src') or ''
            if '/img/club/' in src and '/500' in src:
                crest=urljoin(source_url or 'https://srbijasport.net/',src); break
            if alt and title and (_norm_comp_name(alt) in _norm_comp_name(title) or _norm_comp_name(title) in _norm_comp_name(alt)) and 'stadion' not in _norm_comp_name(alt):
                crest=urljoin(source_url or 'https://srbijasport.net/',src); break
    info={'official_name':title,'founded':founded,'address_line1':address,'postal_city':'','city':'','country':'Serbia','phone':'','fax':'','website':website,
          'squad_size':'','average_age':'','foreigners':'','foreigners_pct':'','national_team_players':'','transfer_record':'','crest_url':crest}
    stadium={'name':stadium_name,'capacity':None,'city':''}
    return info,stadium


def enrich_women_clubs_srbijasport(country_id=1):
    con=db(); rows=con.execute('''SELECT c.id,c.source_club_url,COALESCE(c.cm_name,c.name) club_name FROM clubs c JOIN leagues l ON l.id=c.league_id
        WHERE l.country_id=? AND l.section='women' AND COALESCE(c.source_active,1)=1 AND COALESCE(c.source_club_url,'')<>'' ORDER BY l.level,l.name,c.name''',(country_id,)).fetchall(); con.close()
    out={'clubs':len(rows),'ok':0,'crests':0,'info':0,'errors':[]}
    for c in rows:
        try:
            html=fetch_public_html(c['source_club_url'],'https://srbijasport.net/')
            info,stadium=parse_srbijasport_club_info(html,c['source_club_url'])
            dbx=db(); old=dbx.execute('SELECT * FROM club_info WHERE club_id=?',(c['id'],)).fetchone(); dbx.close()
            crest_local=(old['crest_local'] if old and old['crest_local'] else '')
            if crest_local and not (BASE_DIR / crest_local).exists():
                crest_local=''
            crest_url=(info.get('crest_url','') or (old['crest_url'] if old and old['crest_url'] else ''))
            if not crest_local and crest_url:
                crest_local=cache_crest(c['id'],crest_url) or ''
            con2=db(); con2.execute('''INSERT INTO club_info(club_id,official_name,founded,address_line1,postal_city,city,country,phone,fax,website,squad_size,average_age,foreigners,foreigners_pct,national_team_players,transfer_record,crest_url,crest_local,updated_at)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,CURRENT_TIMESTAMP)
                ON CONFLICT(club_id) DO UPDATE SET
                official_name=CASE WHEN COALESCE(club_info.official_name,'')='' THEN excluded.official_name ELSE club_info.official_name END,
                founded=CASE WHEN COALESCE(club_info.founded,'')='' THEN excluded.founded ELSE club_info.founded END,
                address_line1=CASE WHEN COALESCE(club_info.address_line1,'')='' THEN excluded.address_line1 ELSE club_info.address_line1 END,
                city=CASE WHEN COALESCE(club_info.city,'')='' THEN excluded.city ELSE club_info.city END,
                website=CASE WHEN COALESCE(club_info.website,'')='' THEN excluded.website ELSE club_info.website END,
                crest_url=CASE WHEN COALESCE(club_info.crest_url,'')='' THEN excluded.crest_url ELSE club_info.crest_url END,
                crest_local=CASE WHEN COALESCE(club_info.crest_local,'')='' THEN excluded.crest_local ELSE club_info.crest_local END,
                updated_at=CURRENT_TIMESTAMP''',
                (c['id'],info.get('official_name',''),info.get('founded',''),info.get('address_line1',''),'',info.get('city',''),'Serbia','','',info.get('website',''),'','','','','','',crest_url,crest_local))
            if stadium.get('name'):
                con2.execute('''INSERT INTO stadiums(club_id,name,capacity,city,updated_at) VALUES(?,?,?,?,CURRENT_TIMESTAMP)
                    ON CONFLICT(club_id) DO UPDATE SET name=CASE WHEN COALESCE(stadiums.name,'')='' THEN excluded.name ELSE stadiums.name END,
                    city=CASE WHEN COALESCE(stadiums.city,'')='' THEN excluded.city ELSE stadiums.city END,updated_at=CURRENT_TIMESTAMP''',(c['id'],stadium.get('name',''),stadium.get('capacity'),stadium.get('city','')))
            con2.commit(); con2.close(); out['ok']+=1; out['info']+=1
            if crest_local: out['crests']+=1
        except Exception as e: out['errors'].append(f"{c['club_name']}: {e}")
    return out





def _stadium_image_abs(src, source_url=''):
    src=clean_text(src or '')
    if not src:
        return ''
    if src.startswith('//'):
        return 'https:'+src
    if src.startswith('/'):
        return urljoin(source_url or 'https://', src)
    return src


def _img_src(img, source_url=''):
    for raw in (img.get('data-src'), img.get('data-lazy-src'), img.get('src'), img.get('data-original')):
        if raw:
            src=_stadium_image_abs(raw, source_url)
            if src:
                return src
    return ''


def parse_srbijasport_stadium_image(html, source_url=''):
    # Strict SrbijaSport parser: only the club-info image explicitly labelled Stadion.
    if not html:
        return ''
    soup=BeautifulSoup(html,'lxml')
    for img in soup.find_all('img'):
        src=_img_src(img, source_url)
        if not src:
            continue
        meta=' '.join([clean_text(img.get('alt') or ''), clean_text(img.get('title') or '')]).lower()
        if 'stadion' in meta and '/img/clubinfo/' in urlparse(src).path.lower():
            return src
    return ''


def parse_transfermarkt_stadium_image(html, source_url='', stadium_name=''):
    # Strict TM parser: stadium-specific page only; reject crests, trophies and logos.
    if not html or '/stadion/verein/' not in (source_url or ''):
        return ''
    soup=BeautifulSoup(html,'lxml')
    wanted=_norm_media_key(stadium_name)
    best=''
    best_score=-999
    for img in soup.find_all('img'):
        src=_img_src(img, source_url)
        if not src:
            continue
        meta=' '.join([clean_text(img.get('alt') or ''), clean_text(img.get('title') or ''), ' '.join(img.get('class',[]) or []), src]).lower()
        path=urlparse(src).path.lower()
        if any(x in path for x in ('/images/wappen/','/logo/','/logos/','/flaggen/')) or any(x in meta for x in ('wappen','crest','logo','trophy','pokal','pehar')):
            continue
        score=0
        if 'stadion' in meta or 'stadium' in meta or 'arena' in meta:
            score+=8
        if wanted:
            m=_norm_media_key(meta)
            if wanted in m:
                score+=8
            else:
                toks=[t for t in wanted.split() if len(t)>=4]
                score+=sum(2 for t in toks if t in m)
        if score>best_score:
            best_score=score; best=src
    return best if best_score>=8 else ''


def parse_stadium_image_html(html, source_url='', stadium_name=''):
    # Compatibility wrapper using only trusted source-specific parsers.
    host=(urlparse(source_url).netloc or '').lower()
    if 'srbijasport.net' in host:
        return parse_srbijasport_stadium_image(html, source_url)
    if 'transfermarkt.' in host:
        return parse_transfermarkt_stadium_image(html, source_url, stadium_name)
    return ''



OFFICIAL_STADIUM_PATHS = {
    'crvenazvezdafk.com': ['/klub/stadion-rajko-mitic','/sr-latn/klub/stadion-rajko-mitic'],
    'www.crvenazvezdafk.com': ['/klub/stadion-rajko-mitic','/sr-latn/klub/stadion-rajko-mitic'],
    'partizan.rs': ['/vesti/stadion'],
    'www.partizan.rs': ['/vesti/stadion'],
    'fkradnickinis.rs': ['/stadion/'],
    'www.fkradnickinis.rs': ['/stadion/'],
    'fkmacva.com': ['/klub/stadion/'],
    'www.fkmacva.com': ['/klub/stadion/'],
    'fkvozdovac.rs': ['/stadion/'],
    'www.fkvozdovac.rs': ['/stadion/'],
    'fkbor1919.rs': ['/stadion/'],
    'www.fkbor1919.rs': ['/stadion/'],
    'fkborac1926.com': ['/stadion/'],
    'www.fkborac1926.com': ['/stadion/'],
    'fkmetalac.rs': ['/stadion/'],
    'www.fkmetalac.rs': ['/stadion/'],
    'fkloznica.rs': ['/stadion-lagator/'],
    'www.fkloznica.rs': ['/stadion-lagator/'],
    'fkjedinstvoub.com': ['/stadion'],
    'www.fkjedinstvoub.com': ['/stadion'],
}


def _website_url(value):
    value=clean_text(value or '')
    if not value:
        return ''
    if value.startswith('//'):
        return 'https:'+value
    if not re.match(r'^https?://',value,re.I):
        value='https://'+value.lstrip('/')
    return value


def _srcset_best(img, source_url=''):
    raw=clean_text(img.get('srcset') or img.get('data-srcset') or '')
    if not raw:
        return ''
    items=[]
    for part in raw.split(','):
        bit=part.strip().split()
        if not bit:
            continue
        url=_stadium_image_abs(bit[0],source_url)
        width=0
        if len(bit)>1:
            m=re.match(r'(\d+)[wx]?',bit[1])
            if m:
                width=int(m.group(1))
        if url:
            items.append((width,url))
    return max(items,key=lambda x:x[0])[1] if items else ''


def _img_src_official(img, source_url=''):
    return _srcset_best(img,source_url) or _img_src(img,source_url)


def _bad_media_url(src='', meta=''):
    blob=(clean_text(src)+' '+clean_text(meta)).lower()
    bad=('logo','crest','grb','wappen','favicon','avatar','icon','sponsor','partner','banner-ad','advert','trophy','pokal','pehar','medal','badge')
    return any(x in blob for x in bad)


def _official_page_context(soup, page_url='', stadium_name=''):
    parts=[urlparse(page_url).path]
    if soup.title:
        parts.append(soup.title.get_text(' ',strip=True))
    for tag in soup.find_all(['h1','h2'],limit=8):
        parts.append(tag.get_text(' ',strip=True))
    ctx=_norm_media_key(' '.join(parts))
    if any(k in ctx for k in ('stadion','stadium','arena','stadio')):
        return True
    wanted=_norm_media_key(stadium_name)
    if wanted and len(wanted)>=5:
        toks=[t for t in wanted.split() if len(t)>=4 and t not in ('stadion','stadium','gradski')]
        return bool(toks) and sum(1 for t in toks if t in ctx)>=max(1,min(2,len(toks)))
    return False


def _near_stadium_section_images(soup, page_url='', stadium_name=''):
    out=[]
    wanted=_norm_media_key(stadium_name)
    for node in soup.find_all(['h1','h2','h3','h4','strong','b','div','section']):
        txt=_norm_media_key(node.get_text(' ',strip=True)[:500])
        if not txt:
            continue
        match=('stadion' in txt or 'stadium' in txt or 'arena' in txt)
        if wanted and wanted in txt:
            match=True
        if not match:
            continue
        containers=[node]
        par=node.parent
        for _ in range(3):
            if par is None:
                break
            containers.append(par)
            par=par.parent
        for c in containers:
            for img in c.find_all('img',limit=12):
                src=_img_src_official(img,page_url)
                if src:
                    meta=' '.join([clean_text(img.get('alt') or ''),clean_text(img.get('title') or '')])
                    if not _bad_media_url(src,meta):
                        out.append(src)
            if out:
                break
        if out:
            break
    return list(dict.fromkeys(out))


def parse_official_stadium_image(html, page_url='', stadium_name=''):
    if not html:
        return ''
    soup=BeautifulSoup(html,'lxml')
    page_is_stadium=_official_page_context(soup,page_url,stadium_name)
    wanted=_norm_media_key(stadium_name)
    scored=[]
    for img in soup.find_all('img'):
        src=_img_src_official(img,page_url)
        if not src:
            continue
        meta=' '.join([clean_text(img.get('alt') or ''),clean_text(img.get('title') or ''),' '.join(img.get('class',[]) or [])])
        if _bad_media_url(src,meta):
            continue
        nm=_norm_media_key(meta+' '+src)
        score=0
        if any(x in nm for x in ('stadion','stadium','arena')):
            score+=12
        if wanted:
            toks=[t for t in wanted.split() if len(t)>=4 and t not in ('stadion','stadium','gradski')]
            score+=sum(3 for t in toks if t in nm)
        try:
            w=int(re.sub(r'\D','',str(img.get('width') or '0')) or 0)
            h=int(re.sub(r'\D','',str(img.get('height') or '0')) or 0)
            if w>=600 or h>=350:
                score+=2
        except Exception:
            pass
        if page_is_stadium:
            score+=4
        scored.append((score,src))
    if scored:
        scored.sort(reverse=True,key=lambda x:x[0])
        if scored[0][0]>=8:
            return scored[0][1]
    # On a confirmed stadium page, OG/Twitter hero is acceptable unless clearly a logo/crest.
    if page_is_stadium:
        for attr in [('property','og:image'),('name','twitter:image'),('property','twitter:image')]:
            tag=soup.find('meta',attrs={attr[0]:attr[1]})
            src=_stadium_image_abs((tag.get('content') if tag else ''),page_url)
            if src and not _bad_media_url(src,''):
                return src
        nearby=_near_stadium_section_images(soup,page_url,stadium_name)
        if nearby:
            return nearby[0]
    return ''


def _same_site(url, root_url):
    a=(urlparse(url).netloc or '').lower().removeprefix('www.')
    b=(urlparse(root_url).netloc or '').lower().removeprefix('www.')
    return bool(a and b and a==b)


def _official_stadium_candidates(website, stadium_name='', club_name=''):
    root=_website_url(website)
    if not root:
        return []
    u=urlparse(root)
    base=f'{u.scheme}://{u.netloc}/'
    host=(u.netloc or '').lower()
    candidates=[]
    for path in OFFICIAL_STADIUM_PATHS.get(host,[]):
        candidates.append(urljoin(base,path))
    # Common routes cover many club CMS implementations.
    for path in ('/stadion/','/stadion','/stadium/','/stadium','/klub/stadion/','/klub/stadion','/stadion.html','/stadion-lagator/'):
        candidates.append(urljoin(base,path))
    # Discover explicit internal stadium links from the official homepage.
    try:
        html=fetch_public_html(root,base)
        soup=BeautifulSoup(html,'lxml')
        for a in soup.find_all('a',href=True):
            href=urljoin(root,a.get('href'))
            if not _same_site(href,root):
                continue
            text=_norm_media_key((a.get_text(' ',strip=True) or '')+' '+href)
            if any(k in text for k in ('stadion','stadium','arena')):
                candidates.append(href)
        # The homepage itself can contain a dedicated stadium block/gallery.
        if _near_stadium_section_images(soup,root,stadium_name):
            candidates.append(root)
    except Exception:
        pass
    # Sitemap discovery is cheap and often finds a hidden Stadion page.
    for sm in (urljoin(base,'sitemap.xml'),urljoin(base,'wp-sitemap.xml')):
        try:
            xml=fetch_public_html(sm,base)
            for loc in re.findall(r'<loc>(.*?)</loc>',xml,re.I|re.S):
                loc=html_lib.unescape(clean_text(loc))
                if _same_site(loc,root) and any(k in _norm_media_key(loc) for k in ('stadion','stadium','arena')):
                    candidates.append(loc)
        except Exception:
            pass
    out=[]
    for x in candidates:
        if x and x not in out:
            out.append(x)
    return out[:18]


def fetch_official_stadium_image(website, stadium_name='', club_name=''):
    root=_website_url(website)
    if not root:
        return None
    host=(urlparse(root).netloc or '').lower().removeprefix('www.')
    for page in _official_stadium_candidates(root,stadium_name,club_name):
        try:
            html=fetch_public_html(page,root)
            img=parse_official_stadium_image(html,page,stadium_name)
            if img:
                return {'page_url':page,'image_url':img,'source':f'Official · {host}'}
        except Exception:
            continue
    return None


def _guess_ext(url='', ctype=''):
    ctype=(ctype or '').split(';',1)[0].strip().lower()
    ext_map={'image/png':'.png','image/jpeg':'.jpg','image/jpg':'.jpg','image/webp':'.webp','image/svg+xml':'.svg','image/gif':'.gif'}
    if ctype in ext_map:
        return ext_map[ctype]
    path=(urlparse(url).path or '').lower()
    for ext in ('.png','.jpg','.jpeg','.webp','.svg','.gif'):
        if path.endswith(ext):
            return '.jpg' if ext=='.jpeg' else ext
    return '.jpg'


def _download_image_bytes(url, referer=''):
    headers={**HEADERS,'Accept':'image/avif,image/webp,image/apng,image/svg+xml,image/*,*/*;q=0.8'}
    if referer:
        headers['Referer']=referer
    r=requests.get(url,headers=headers,timeout=35)
    r.raise_for_status()
    if not r.content:
        raise ValueError('empty image content')
    ctype=(r.headers.get('Content-Type') or '').split(';',1)[0].strip().lower()
    if ctype and not (ctype.startswith('image/') or ctype=='application/octet-stream'):
        raise ValueError(f'not an image: {ctype}')
    return r.content, _guess_ext(url, ctype)


def _write_stadium_asset(club_id, payload, ext):
    club_id=int(club_id)
    ext='.'+ext.lstrip('.') if ext else '.jpg'
    if ext=='.jpeg':
        ext='.jpg'
    for old in STADIUM_DIR.glob(f'{club_id}.*'):
        try:
            old.unlink()
        except Exception:
            pass
    dest=STADIUM_DIR / f'{club_id}{ext}'
    dest.write_bytes(payload)
    return f'static/stadiums/{dest.name}'


def _upsert_stadium_media(club_id, image_url='', image_local='', image_source=''):
    con=db()
    con.execute("""INSERT INTO stadiums(club_id,image_url,image_local,image_source,updated_at)
        VALUES(?,?,?,?,CURRENT_TIMESTAMP)
        ON CONFLICT(club_id) DO UPDATE SET
        image_url=CASE WHEN COALESCE(excluded.image_url,'')<>'' THEN excluded.image_url ELSE stadiums.image_url END,
        image_local=CASE WHEN COALESCE(excluded.image_local,'')<>'' THEN excluded.image_local ELSE stadiums.image_local END,
        image_source=CASE WHEN COALESCE(excluded.image_source,'')<>'' THEN excluded.image_source ELSE stadiums.image_source END,
        updated_at=CURRENT_TIMESTAMP""",(int(club_id),image_url,image_local,image_source))
    con.commit(); con.close()


def _norm_media_key(value):
    txt=unicodedata.normalize('NFKD', clean_text(value or ''))
    txt=''.join(ch for ch in txt if not unicodedata.combining(ch))
    return re.sub(r'\s+',' ',txt).strip().lower()


_GENERIC_STADIUM_NAMES={_norm_media_key(x) for x in [
    'Stadion','Gradski stadion','City Stadium','Municipal Stadium','Stadion kraj Morave',
    'Stadion kraj Nišave','Stadion kraj Timoka','Stadion kraj Skrobare','Stadion pored Tise'
]}


def _is_generic_stadium_name(name):
    return _norm_media_key(name) in _GENERIC_STADIUM_NAMES


def _tm_stadium_urls(tm_url='', tm_club_id=''):
    # Stadium-specific Transfermarkt URLs only. Never parse the normal club page for media.
    tmid=clean_text(tm_club_id or '')
    if not tmid:
        return []
    try:
        u=urlparse(tm_url or 'https://www.transfermarkt.com')
        root=f'{u.scheme or "https"}://{u.netloc or "www.transfermarkt.com"}'
    except Exception:
        root='https://www.transfermarkt.com'
    return list(dict.fromkeys([root+f'/stadion/verein/{tmid}', root+f'/-/stadion/verein/{tmid}']))


def _commons_candidates(query):
    if not query:
        return []
    api='https://commons.wikimedia.org/w/api.php'
    params={
        'action':'query','generator':'search','gsrnamespace':'6','gsrsearch':query,'gsrlimit':'6',
        'prop':'imageinfo','iiprop':'url|mime|size','iiurlwidth':'1600','format':'json'
    }
    r=requests.get(api,params=params,headers=HEADERS,timeout=25)
    r.raise_for_status()
    data=r.json()
    pages=list((data.get('query') or {}).get('pages',{}).values())
    out=[]
    for p in sorted(pages,key=lambda x:int(x.get('index',9999))):
        ii=(p.get('imageinfo') or [{}])[0]
        url=ii.get('thumburl') or ii.get('url') or ''
        mime=(ii.get('mime') or '').lower()
        if not url or (mime and not mime.startswith('image/')):
            continue
        out.append({'title':p.get('title',''),'url':url,'mime':mime})
    return out


def _wikipedia_page_image(query, lang='sr'):
    if not query:
        return None
    api=f'https://{lang}.wikipedia.org/w/api.php'
    s=requests.get(api,params={'action':'query','list':'search','srsearch':query,'srlimit':'3','format':'json'},headers=HEADERS,timeout=25)
    s.raise_for_status()
    hits=(s.json().get('query') or {}).get('search',[])
    if not hits:
        return None
    pageids='|'.join(str(h['pageid']) for h in hits[:3])
    p=requests.get(api,params={'action':'query','pageids':pageids,'prop':'pageimages','piprop':'original','format':'json'},headers=HEADERS,timeout=25)
    p.raise_for_status()
    pages=(p.json().get('query') or {}).get('pages',{})
    for pid in pageids.split('|'):
        node=pages.get(pid) or pages.get(int(pid)) or {}
        original=node.get('original') or {}
        src=original.get('source') or ''
        if src:
            return {'title':node.get('title',''),'url':src,'mime':''}
    return None


def _stadium_queries(stadium_name='', club_name='', city=''):
    queries=[]
    sname=clean_text(stadium_name or '')
    cname=clean_text(club_name or '')
    ccity=clean_text(city or '')
    if sname:
        queries.extend([
            f'"{sname}" {ccity} Serbia',
            f'"{sname}" stadium Serbia',
            f'{sname} {cname}',
            sname,
        ])
    if cname:
        queries.extend([
            f'{cname} stadium Serbia',
            f'{cname} stadion',
        ])
    out=[]
    for q in queries:
        q=clean_text(q)
        if q and q not in out:
            out.append(q)
    return out


def _trusted_existing_source(value):
    src=(value or '').strip().lower()
    return src.startswith('official') or src.startswith('srbijasport') or src.startswith('transfermarkt stadium') or src.startswith('manual')


def _clear_stadium_media_row(club_id):
    con=db()
    row=con.execute('SELECT image_local FROM stadiums WHERE club_id=?',(int(club_id),)).fetchone()
    local=(row['image_local'] if row and row['image_local'] else '')
    con.execute("UPDATE stadiums SET image_url='',image_local='',image_source='',updated_at=CURRENT_TIMESTAMP WHERE club_id=?",(int(club_id),))
    con.commit(); con.close()
    if local:
        try:
            f=BASE_DIR/local
            if f.exists() and STADIUM_DIR in f.resolve().parents:
                f.unlink()
        except Exception:
            pass


def cleanup_untrusted_stadium_media(country_id=1, levels=(1,2,3)):
    wanted=tuple(int(x) for x in levels)
    marks=','.join('?' for _ in wanted)
    con=db()
    sql=f"""SELECT s.club_id,s.image_local,s.image_source
        FROM stadiums s JOIN clubs c ON c.id=s.club_id JOIN leagues l ON l.id=c.league_id
        WHERE l.country_id=? AND COALESCE(l.section,'senior')='senior' AND l.level IN ({marks})
          AND COALESCE(s.image_local,'')<>''"""
    rows=con.execute(sql,(int(country_id),*wanted)).fetchall()
    con.close()
    removed=0
    for r in rows:
        if not _trusted_existing_source(r['image_source'] or ''):
            _clear_stadium_media_row(r['club_id']); removed+=1
    return {'checked':len(rows),'removed':removed}


def ensure_stadium_media(club_id, stadium_name='', club_name='', city='', source_club_url='', tm_url='', tm_club_id='', website=''):
    con=db(); row=con.execute('SELECT image_local,image_url,image_source FROM stadiums WHERE club_id=?',(int(club_id),)).fetchone(); con.close()
    if row and row['image_local'] and (BASE_DIR / row['image_local']).exists() and _trusted_existing_source(row['image_source'] or ''):
        return {'status':'EXISTING','image_local':row['image_local'],'image_url':row['image_url'] or '','image_source':row['image_source'] or 'trusted local cache'}
    if row and row['image_local']:
        _clear_stadium_media_row(club_id)

    # 1) Official club website: preferred source for L1/L2 when available.
    if website:
        try:
            found=fetch_official_stadium_image(website,stadium_name,club_name)
            if found and found.get('image_url'):
                payload,ext=_download_image_bytes(found['image_url'],found.get('page_url',''))
                image_local=_write_stadium_asset(club_id,payload,ext)
                source=found.get('source') or 'Official club website'
                _upsert_stadium_media(club_id,found['image_url'],image_local,source)
                return {'status':'AUTO FOUND','image_local':image_local,'image_url':found['image_url'],'image_source':source}
        except Exception:
            pass

    # 2) SrbijaSport: exact Stadion club-info image only.
    if source_club_url and 'srbijasport.net' in source_club_url:
        try:
            html=fetch_public_html(source_club_url,'https://srbijasport.net/')
            image_url=parse_srbijasport_stadium_image(html,source_club_url)
            if image_url:
                payload,ext=_download_image_bytes(image_url,source_club_url)
                image_local=_write_stadium_asset(club_id,payload,ext)
                _upsert_stadium_media(club_id,image_url,image_local,'SrbijaSport · Stadion')
                return {'status':'AUTO FOUND','image_local':image_local,'image_url':image_url,'image_source':'SrbijaSport · Stadion'}
        except Exception:
            pass

    # 3) Transfermarkt: stadium-specific page only.
    if tm_club_id:
        session=requests.Session()
        for url in _tm_stadium_urls(tm_url,tm_club_id):
            try:
                html=fetch_html(url,session)
                image_url=parse_transfermarkt_stadium_image(html,url,stadium_name)
                if not image_url:
                    continue
                payload,ext=_download_image_bytes(image_url,url)
                image_local=_write_stadium_asset(club_id,payload,ext)
                _upsert_stadium_media(club_id,image_url,image_local,'Transfermarkt Stadium')
                return {'status':'AUTO FOUND','image_local':image_local,'image_url':image_url,'image_source':'Transfermarkt Stadium'}
            except Exception:
                continue

    # No Wikipedia / Commons / generic web fallback by design.
    return {'status':'NOT FOUND','image_local':'','image_url':'','image_source':''}


def _write_stadium_report(rows, tag='stadium_update'):
    if not rows:
        return ''
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    path=REPORT_DIR / f'{tag}_{datetime.now().strftime("%Y%m%d_%H%M%S")}.csv'
    headers=['level','league','club','stadium_name','city','status','image_source','image_local','image_url']
    with path.open('w',newline='',encoding='utf-8-sig') as f:
        w=csv.DictWriter(f,fieldnames=headers)
        w.writeheader()
        for row in rows:
            w.writerow({h:row.get(h,'') for h in headers})
    return str(path.relative_to(BASE_DIR)).replace('\\','/')


def enrich_lower_clubs_srbijasport(country_id=1):
    """Fill missing crest/basic info for senior Level 3-6 clubs from SrbijaSport.

    This is independent of Transfermarkt mapping and uses the same proven
    SrbijaSport profile + crest endpoint flow as women's club enrichment.
    """
    con=db(); rows=con.execute("""SELECT c.id,c.source_club_url,COALESCE(c.cm_name,c.name) club_name,l.name league_name,l.level
        FROM clubs c JOIN leagues l ON l.id=c.league_id
        WHERE l.country_id=? AND COALESCE(l.section,'senior')='senior' AND l.level IN (3,4,5,6)
          AND COALESCE(c.source_active,1)=1 AND COALESCE(c.source_club_url,'')<>''
        ORDER BY l.level,l.name,c.name""",(country_id,)).fetchall(); con.close()
    out={'clubs':len(rows),'ok':0,'crests':0,'info':0,'errors':[]}
    for c in rows:
        try:
            html=fetch_public_html(c['source_club_url'],'https://srbijasport.net/')
            info,stadium=parse_srbijasport_club_info(html,c['source_club_url'])
            dbx=db(); old=dbx.execute('SELECT * FROM club_info WHERE club_id=?',(c['id'],)).fetchone(); dbx.close()
            crest_local=(old['crest_local'] if old and old['crest_local'] else '')
            if crest_local and not (BASE_DIR / crest_local).exists():
                crest_local=''
            crest_url=(info.get('crest_url','') or (old['crest_url'] if old and old['crest_url'] else ''))
            if not crest_local and crest_url:
                crest_local=cache_crest(c['id'],crest_url) or ''
            con2=db(); con2.execute("""INSERT INTO club_info(club_id,official_name,founded,address_line1,postal_city,city,country,phone,fax,website,squad_size,average_age,foreigners,foreigners_pct,national_team_players,transfer_record,crest_url,crest_local,updated_at)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,CURRENT_TIMESTAMP)
                ON CONFLICT(club_id) DO UPDATE SET
                official_name=CASE WHEN COALESCE(club_info.official_name,'')='' THEN excluded.official_name ELSE club_info.official_name END,
                founded=CASE WHEN COALESCE(club_info.founded,'')='' THEN excluded.founded ELSE club_info.founded END,
                address_line1=CASE WHEN COALESCE(club_info.address_line1,'')='' THEN excluded.address_line1 ELSE club_info.address_line1 END,
                city=CASE WHEN COALESCE(club_info.city,'')='' THEN excluded.city ELSE club_info.city END,
                website=CASE WHEN COALESCE(club_info.website,'')='' THEN excluded.website ELSE club_info.website END,
                crest_url=CASE WHEN COALESCE(club_info.crest_url,'')='' THEN excluded.crest_url ELSE club_info.crest_url END,
                crest_local=CASE WHEN COALESCE(club_info.crest_local,'')='' THEN excluded.crest_local ELSE club_info.crest_local END,
                updated_at=CURRENT_TIMESTAMP""",
                (c['id'],info.get('official_name',''),info.get('founded',''),info.get('address_line1',''),'',info.get('city',''),'Serbia','','',info.get('website',''),'','','','','','',crest_url,crest_local))
            if stadium.get('name'):
                con2.execute("""INSERT INTO stadiums(club_id,name,capacity,city,updated_at) VALUES(?,?,?,?,CURRENT_TIMESTAMP)
                    ON CONFLICT(club_id) DO UPDATE SET name=CASE WHEN COALESCE(stadiums.name,'')='' THEN excluded.name ELSE stadiums.name END,
                    city=CASE WHEN COALESCE(stadiums.city,'')='' THEN excluded.city ELSE stadiums.city END,updated_at=CURRENT_TIMESTAMP""",
                    (c['id'],stadium.get('name',''),stadium.get('capacity'),stadium.get('city','')))
            con2.commit(); con2.close(); out['ok']+=1; out['info']+=1
            if crest_local: out['crests']+=1
        except Exception as e:
            out['errors'].append(f"L{c['level']} {c['league_name']} / {c['club_name']}: {e}")
    return out

def _merge_club_info_missing(club_id, info, stadium=None, crest_local=""):
    """Merge enrichment data without overwriting already populated values."""
    stadium=stadium or {}
    con=db()
    old=con.execute("SELECT * FROM club_info WHERE club_id=?",(club_id,)).fetchone()
    old_crest=(old["crest_local"] if old and old["crest_local"] else "")
    final_crest=old_crest or crest_local or ""
    crest_url=(info.get("crest_url","") or (old["crest_url"] if old and old["crest_url"] else ""))
    con.execute("""INSERT INTO club_info(club_id,official_name,founded,address_line1,postal_city,city,country,phone,fax,website,squad_size,average_age,foreigners,foreigners_pct,national_team_players,transfer_record,crest_url,crest_local,updated_at)
        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,CURRENT_TIMESTAMP)
        ON CONFLICT(club_id) DO UPDATE SET
        official_name=CASE WHEN COALESCE(club_info.official_name,'')='' THEN excluded.official_name ELSE club_info.official_name END,
        founded=CASE WHEN COALESCE(club_info.founded,'')='' THEN excluded.founded ELSE club_info.founded END,
        address_line1=CASE WHEN COALESCE(club_info.address_line1,'')='' THEN excluded.address_line1 ELSE club_info.address_line1 END,
        postal_city=CASE WHEN COALESCE(club_info.postal_city,'')='' THEN excluded.postal_city ELSE club_info.postal_city END,
        city=CASE WHEN COALESCE(club_info.city,'')='' THEN excluded.city ELSE club_info.city END,
        country=CASE WHEN COALESCE(club_info.country,'')='' THEN excluded.country ELSE club_info.country END,
        phone=CASE WHEN COALESCE(club_info.phone,'')='' THEN excluded.phone ELSE club_info.phone END,
        fax=CASE WHEN COALESCE(club_info.fax,'')='' THEN excluded.fax ELSE club_info.fax END,
        website=CASE WHEN COALESCE(club_info.website,'')='' THEN excluded.website ELSE club_info.website END,
        squad_size=CASE WHEN COALESCE(club_info.squad_size,'')='' THEN excluded.squad_size ELSE club_info.squad_size END,
        average_age=CASE WHEN COALESCE(club_info.average_age,'')='' THEN excluded.average_age ELSE club_info.average_age END,
        foreigners=CASE WHEN COALESCE(club_info.foreigners,'')='' THEN excluded.foreigners ELSE club_info.foreigners END,
        foreigners_pct=CASE WHEN COALESCE(club_info.foreigners_pct,'')='' THEN excluded.foreigners_pct ELSE club_info.foreigners_pct END,
        national_team_players=CASE WHEN COALESCE(club_info.national_team_players,'')='' THEN excluded.national_team_players ELSE club_info.national_team_players END,
        transfer_record=CASE WHEN COALESCE(club_info.transfer_record,'')='' THEN excluded.transfer_record ELSE club_info.transfer_record END,
        crest_url=CASE WHEN COALESCE(club_info.crest_url,'')='' THEN excluded.crest_url ELSE club_info.crest_url END,
        crest_local=CASE WHEN COALESCE(club_info.crest_local,'')='' THEN excluded.crest_local ELSE club_info.crest_local END,
        updated_at=CURRENT_TIMESTAMP""",
        (club_id,info.get("official_name",""),info.get("founded",""),info.get("address_line1",""),info.get("postal_city",""),info.get("city",""),info.get("country","Serbia"),info.get("phone",""),info.get("fax",""),info.get("website",""),info.get("squad_size",""),info.get("average_age",""),info.get("foreigners",""),info.get("foreigners_pct",""),info.get("national_team_players",""),info.get("transfer_record",""),crest_url,final_crest))
    if stadium.get("name"):
        con.execute("""INSERT INTO stadiums(club_id,name,capacity,city,updated_at) VALUES(?,?,?,?,CURRENT_TIMESTAMP)
            ON CONFLICT(club_id) DO UPDATE SET
            name=CASE WHEN COALESCE(stadiums.name,'')='' THEN excluded.name ELSE stadiums.name END,
            capacity=CASE WHEN stadiums.capacity IS NULL THEN excluded.capacity ELSE stadiums.capacity END,
            city=CASE WHEN COALESCE(stadiums.city,'')='' THEN excluded.city ELSE stadiums.city END,
            updated_at=CURRENT_TIMESTAMP""",(club_id,stadium.get("name",""),stadium.get("capacity"),stadium.get("city","")))
    con.commit(); con.close()


def enrich_clubs_group(country_id=1, section=None, levels=None, mode="both"):
    """Generic crest/info enrichment used by all admin competition blocks."""
    con=db()
    sql="""SELECT c.id,c.source_club_url,c.tm_url,c.tm_club_id,COALESCE(c.cm_name,c.name) club_name,l.name league_name,l.level,COALESCE(l.section,'senior') section
        FROM clubs c JOIN leagues l ON l.id=c.league_id WHERE l.country_id=? AND COALESCE(c.source_active,1)=1"""
    params=[country_id]
    if section is not None:
        sql+=" AND COALESCE(l.section,'senior')=?"; params.append(section)
    if levels:
        marks=','.join('?' for _ in levels); sql+=f" AND l.level IN ({marks})"; params.extend([int(x) for x in levels])
    sql+=' ORDER BY l.level,l.name,c.name'
    clubs=con.execute(sql,params).fetchall(); con.close()
    out={'clubs':len(clubs),'eligible':0,'ok':0,'crests':0,'info':0,'stadium_auto':0,'stadium_ambiguous':0,'stadium_missing':0,'stadium_existing':0,'report_path':'','errors':[]}
    session=requests.Session()
    report_rows=[]
    for c in clubs:
        try:
            info={}; stadium={}
            if c['source_club_url'] and 'srbijasport.net' in c['source_club_url']:
                html=fetch_public_html(c['source_club_url'],'https://srbijasport.net/')
                info,stadium=parse_srbijasport_club_info(html,c['source_club_url'])
            elif c['tm_url']:
                html=fetch_html(c['tm_url'],session)
                info,stadium=parse_club_info_html(html,c['tm_club_id'])
            else:
                continue
            out['eligible']+=1
            con0=db(); old=con0.execute('SELECT * FROM club_info WHERE club_id=?',(c['id'],)).fetchone(); con0.close()
            crest_local=(old['crest_local'] if old and old['crest_local'] else '')
            if crest_local and not (BASE_DIR / crest_local).exists(): crest_local=''
            if mode in ('crests','both') and not crest_local and info.get('crest_url'):
                crest_local=cache_crest(c['id'],info.get('crest_url')) or ''
            if mode=='crests':
                _merge_club_info_missing(c['id'],{'crest_url':info.get('crest_url',''),'country':'Serbia'}, {}, crest_local)
            else:
                _merge_club_info_missing(c['id'],info,stadium,crest_local if mode=='both' else '')
            if mode in ('info','both') and (c['section'] or 'senior')=='senior' and int(c['level'] or 0) in (1,2,3):
                sres=ensure_stadium_media(c['id'], stadium.get('name',''), c['club_name'], stadium.get('city','') or info.get('city',''), (c['source_club_url'] or ''), (c['tm_url'] or ''), (c['tm_club_id'] or ''), info.get('website',''))
                status=(sres.get('status') or '').upper()
                if status=='AUTO FOUND': out['stadium_auto']+=1
                elif status=='AMBIGUOUS': out['stadium_ambiguous']+=1
                elif status=='EXISTING': out['stadium_existing']+=1
                else: out['stadium_missing']+=1
                report_rows.append({'level':int(c['level'] or 0),'league':c['league_name'],'club':c['club_name'],'stadium_name':stadium.get('name',''),'city':stadium.get('city','') or info.get('city',''),'status':sres.get('status',''),'image_source':sres.get('image_source',''),'image_local':sres.get('image_local',''),'image_url':sres.get('image_url','')})
            out['ok']+=1
            if crest_local: out['crests']+=1
            if mode in ('info','both'): out['info']+=1
        except Exception as e:
            out['errors'].append(f"{c['league_name']} / {c['club_name']}: {e}")
    if mode in ('info','both') and report_rows:
        out['report_path']=_write_stadium_report(report_rows,'stadium_update')
    return out



def final_stadium_sync(country_id=1, levels=(1,2,3)):
    """Independent final stadium pass for current senior clubs.

    It does not depend on a successful club-info refresh. Existing stadium/name/city
    metadata is used when available and club name becomes the search fallback.
    Every eligible club is written to the CSV report.
    """
    wanted=tuple(int(x) for x in (levels or (1,2,3)))
    marks=','.join('?' for _ in wanted)
    con=db()
    rows=con.execute(f"""SELECT c.id,c.source_club_url,c.tm_url,c.tm_club_id,
        COALESCE(c.cm_name,c.name) club_name,l.name league_name,l.level,
        COALESCE(s.name,'') stadium_name,COALESCE(s.city,'') stadium_city,
        COALESCE(ci.city,ci.postal_city,'') info_city,COALESCE(ci.website,'') website
        FROM clubs c JOIN leagues l ON l.id=c.league_id
        LEFT JOIN stadiums s ON s.club_id=c.id
        LEFT JOIN club_info ci ON ci.club_id=c.id
        WHERE l.country_id=? AND COALESCE(l.section,'senior')='senior'
          AND l.level IN ({marks}) AND COALESCE(c.source_active,1)=1
        ORDER BY l.level,l.name,COALESCE(c.cm_name,c.name)""", (int(country_id),*wanted)).fetchall()
    con.close()
    out={'clubs':len(rows),'auto':0,'ambiguous':0,'existing':0,'not_found':0,'errors':0,'report_path':''}
    report=[]
    for r in rows:
        try:
            res=ensure_stadium_media(
                r['id'], r['stadium_name'] or '', r['club_name'] or '',
                r['stadium_city'] or r['info_city'] or '',
                r['source_club_url'] or '', r['tm_url'] or '', r['tm_club_id'] or '', r['website'] or '')
            status=(res.get('status') or 'NOT FOUND').upper()
            if status=='AUTO FOUND': out['auto']+=1
            elif status=='AMBIGUOUS': out['ambiguous']+=1
            elif status=='EXISTING': out['existing']+=1
            else: out['not_found']+=1
            report.append({'level':int(r['level'] or 0),'league':r['league_name'],'club':r['club_name'],
                'stadium_name':r['stadium_name'],'city':r['stadium_city'] or r['info_city'],
                'status':res.get('status','NOT FOUND'),'image_source':res.get('image_source',''),
                'image_local':res.get('image_local',''),'image_url':res.get('image_url','')})
        except Exception as e:
            out['errors']+=1; out['not_found']+=1
            report.append({'level':int(r['level'] or 0),'league':r['league_name'],'club':r['club_name'],
                'stadium_name':r['stadium_name'],'city':r['stadium_city'] or r['info_city'],
                'status':'NOT FOUND','image_source':'ERROR: '+str(e)[:200],
                'image_local':'','image_url':''})
    out['report_path']=_write_stadium_report(report,'stadium_final') if report else ''
    return out


def enrich_all_clubs(country_id=1, mode='both'):
    return enrich_clubs_group(country_id,None,None,mode)

def refresh_level_standings(country_id,level):
    con=db(); rows=con.execute("SELECT id,name FROM leagues WHERE country_id=? AND level=? AND COALESCE(source_active,1)=1 AND COALESCE(standings_source_url,'')<>'' ORDER BY name",(country_id,level)).fetchall(); con.close()
    out={'leagues':len(rows),'ok':0,'rows':0,'errors':[]}
    for l in rows:
        try:
            r=refresh_league_standings(l['id']); out['ok']+=1; out['rows']+=r['rows']
        except Exception as e: out['errors'].append(f"{l['name']}: {e}")
    return out



UEFA_NATIONS_GROUP_TEAMS_2026 = {
    'A2': ['Netherlands','Greece','Germany','Serbia'],
}
UEFA_NATIONS_RESULTS_2026 = 'https://www.uefa.com/uefanationsleague/news/02a2-1fea18abbcbc-456e846509e7-1000--2026-27-uefa-nations-league-all-the-league-phase-fixtures-a/'
UEFA_NATIONS_RESULTS_BY_TEAM_2026 = 'https://www.uefa.com/uefanationsleague/news/02a2-1fea7079900a-801b2c5f6c9e-1000--2026-27-uefa-nations-league-league-phase-fixtures-and-results-by-team/'

# Last officially published completed A2 results known to this build (UEFA, through Matchday 4).
# This is a resilience fallback only: live UEFA standings/results are always attempted first.
UEFA_NATIONS_KNOWN_RESULTS_2026 = {
    'A2': [
        ('Netherlands',1,1,'Germany'),
        ('Serbia',1,2,'Greece'),
        ('Germany',0,1,'Greece'),
        ('Serbia',1,2,'Netherlands'),
        ('Germany',2,0,'Serbia'),
        ('Greece',2,2,'Netherlands'),
        ('Greece',0,0,'Germany'),
        ('Netherlands',2,1,'Serbia'),
    ],
}
UEFA_NATIONS_KNOWN_RESULTS_UPDATED = '2026-10-06'

_UEFA_TEAM_ALIASES = {
    'Netherlands': ('Netherlands','Niederlande','Holandija'),
    'Greece': ('Greece','Griechenland','Grčka','Grcka'),
    'Germany': ('Germany','Deutschland','Nemačka','Nemacka'),
    'Serbia': ('Serbia','Serbien','Srbija'),
}

def _uefa_canonical_team(value, teams):
    nv=_norm_comp_name(value)
    for team in teams:
        for alias in _UEFA_TEAM_ALIASES.get(team,(team,)):
            if nv==_norm_comp_name(alias):
                return team
    return ''

def parse_uefa_group_standings_html(html, group_code):
    """Parse one UEFA Nations League group table from the official standings page."""
    teams=UEFA_NATIONS_GROUP_TEAMS_2026.get(clean_text(group_code).upper(),[])
    if not teams: return []
    soup=BeautifulSoup(html,'lxml')
    tokens=[clean_text(x) for x in soup.stripped_strings if clean_text(x)]
    group=clean_text(group_code).upper()
    # Restrict parsing to the requested group when a heading is available.
    start=0; end=len(tokens)
    for i,tok in enumerate(tokens):
        nt=_norm_comp_name(tok)
        if group.lower() in nt and any(w in nt for w in ('group','gruppe','grupa')):
            start=i+1; break
    for j in range(start,len(tokens)):
        nt=_norm_comp_name(tokens[j])
        if j>start and any(w in nt for w in ('group','gruppe','grupa')) and re.search(r'\b[a-d]\d\b',nt):
            end=j; break
    segment=tokens[start:end]
    found=[]
    for idx,tok in enumerate(segment):
        team=_uefa_canonical_team(tok,teams)
        if not team: continue
        nums=[]
        for v in segment[idx+1:idx+24]:
            if _uefa_canonical_team(v,teams): break
            if re.fullmatch(r'-?\d+',v): nums.append(int(v))
            if len(nums)>=8: break
        if len(nums)>=8:
            played,wins,draws,losses,gf,ga,gd,pts=nums[:8]
            found.append({'team_name':team,'played':played,'wins':wins,'draws':draws,'losses':losses,
                          'goals_for':gf,'goals_against':ga,'goal_diff':gd,'points':pts,'position_note':'','form':'','note':''})
    # De-duplicate by team and require the full group.
    unique=[]; seen=set()
    for r in found:
        if r['team_name'] in seen: continue
        seen.add(r['team_name']); unique.append(r)
    if len(unique)!=len(teams): return []
    # Preserve UEFA's rendered order from the standings page.
    for i,r in enumerate(unique,1): r['position']=i
    return unique

def _uefa_group_matches_from_results_html(html, group_code):
    teams=UEFA_NATIONS_GROUP_TEAMS_2026.get(clean_text(group_code).upper(),[])
    if not teams: return []
    soup=BeautifulSoup(html,'lxml')
    text=' '.join(clean_text(x) for x in soup.stripped_strings if clean_text(x))
    aliases=[]; alias_to_team={}
    for team in teams:
        for alias in _UEFA_TEAM_ALIASES.get(team,(team,)):
            aliases.append(re.escape(alias)); alias_to_team[_norm_comp_name(alias)]=team
    alt='|'.join(sorted(aliases,key=len,reverse=True))
    group=re.escape(clean_text(group_code).upper())
    pat=re.compile(r'\b'+group+r'\s+('+alt+r')\s+(\d+)\s*[-–:]\s*(\d+)\s+('+alt+r')\b',re.I)
    out=[]; seen=set()
    for m in pat.finditer(text):
        home=alias_to_team.get(_norm_comp_name(m.group(1))); away=alias_to_team.get(_norm_comp_name(m.group(4)))
        if not home or not away or home==away: continue
        item=(home,int(m.group(2)),int(m.group(3)),away)
        key=(home,away,item[1],item[2])
        if key in seen: continue
        seen.add(key); out.append(item)
    return out

def _rank_uefa_group(teams, matches):
    data={t:{'team_name':t,'played':0,'wins':0,'draws':0,'losses':0,'goals_for':0,'goals_against':0,'goal_diff':0,'points':0,
             'position_note':'','form':'','note':''} for t in teams}
    for home,hg,ag,away in matches:
        if home not in data or away not in data: continue
        h=data[home]; a=data[away]
        h['played']+=1; a['played']+=1; h['goals_for']+=hg; h['goals_against']+=ag; a['goals_for']+=ag; a['goals_against']+=hg
        if hg>ag: h['wins']+=1; a['losses']+=1; h['points']+=3
        elif hg<ag: a['wins']+=1; h['losses']+=1; a['points']+=3
        else: h['draws']+=1; a['draws']+=1; h['points']+=1; a['points']+=1
    for r in data.values(): r['goal_diff']=r['goals_for']-r['goals_against']
    # UEFA uses head-to-head criteria first. Compute a mini-table inside each equal-points set,
    # then fall back to overall goal difference/goals scored/wins.
    by_points={}
    for t,r in data.items(): by_points.setdefault(r['points'],[]).append(t)
    ordered=[]
    for pts in sorted(by_points,reverse=True):
        tied=by_points[pts]
        if len(tied)==1: ordered.extend(tied); continue
        mini={t:{'p':0,'gd':0,'gf':0} for t in tied}
        for home,hg,ag,away in matches:
            if home not in mini or away not in mini: continue
            mini[home]['gf']+=hg; mini[home]['gd']+=hg-ag; mini[away]['gf']+=ag; mini[away]['gd']+=ag-hg
            if hg>ag: mini[home]['p']+=3
            elif hg<ag: mini[away]['p']+=3
            else: mini[home]['p']+=1; mini[away]['p']+=1
        tied.sort(key=lambda t:(mini[t]['p'],mini[t]['gd'],mini[t]['gf'],data[t]['goal_diff'],data[t]['goals_for'],data[t]['wins'],t), reverse=True)
        ordered.extend(tied)
    rows=[]
    for i,t in enumerate(ordered,1):
        r=dict(data[t]); r['position']=i; rows.append(r)
    return rows

def parse_uefa_group_results_html(html, group_code):
    teams=UEFA_NATIONS_GROUP_TEAMS_2026.get(clean_text(group_code).upper(),[])
    matches=_uefa_group_matches_from_results_html(html,group_code)
    return _rank_uefa_group(teams,matches) if teams and matches else []

def _uefa_group_matches_anywhere_html(html, group_code):
    """Parse completed matches for one configured UEFA group even when the page
    lists fixtures by team instead of prefixing each match with the group code.
    Duplicate match lines are de-duplicated.
    """
    teams=UEFA_NATIONS_GROUP_TEAMS_2026.get(clean_text(group_code).upper(),[])
    if not teams: return []
    soup=BeautifulSoup(html,'lxml')
    text=' '.join(clean_text(x) for x in soup.stripped_strings if clean_text(x))
    aliases=[]; alias_to_team={}
    for team in teams:
        for alias in _UEFA_TEAM_ALIASES.get(team,(team,)):
            aliases.append(re.escape(alias)); alias_to_team[_norm_comp_name(alias)]=team
    alt='|'.join(sorted(aliases,key=len,reverse=True))
    # Accept both '24/09: Serbia 1-2 Greece' and plain 'Serbia 1-2 Greece'.
    pat=re.compile(r'(?:\b\d{1,2}/\d{1,2}:?\s*)?('+alt+r')\s+(\d+)\s*[-–:]\s*(\d+)\s+('+alt+r')\b',re.I)
    out=[]; seen=set()
    for m in pat.finditer(text):
        home=alias_to_team.get(_norm_comp_name(m.group(1))); away=alias_to_team.get(_norm_comp_name(m.group(4)))
        if not home or not away or home==away: continue
        item=(home,int(m.group(2)),int(m.group(3)),away)
        if item in seen: continue
        seen.add(item); out.append(item)
    return out

def parse_uefa_group_results_anywhere_html(html, group_code):
    teams=UEFA_NATIONS_GROUP_TEAMS_2026.get(clean_text(group_code).upper(),[])
    matches=_uefa_group_matches_anywhere_html(html,group_code)
    return _rank_uefa_group(teams,matches) if teams and matches else []

def known_uefa_group_standings(group_code):
    """Return a table calculated from the last bundled official UEFA results snapshot.
    Used only when live UEFA endpoints cannot be fetched/parsed.
    """
    group=clean_text(group_code).upper()
    teams=UEFA_NATIONS_GROUP_TEAMS_2026.get(group,[])
    matches=UEFA_NATIONS_KNOWN_RESULTS_2026.get(group,[])
    return _rank_uefa_group(teams,matches) if teams and matches else []

def refresh_national_team_standings(team_id):
    con=db(); t=con.execute('SELECT * FROM national_teams WHERE id=?',(team_id,)).fetchone(); con.close()
    if not t: raise RuntimeError('Reprezentacija nije pronađena.')
    url=t['standings_source_url'] or ''
    if not url: raise RuntimeError('Za ovu selekciju još nije podešen izvor takmičenja/grupe.')
    source=(t['standings_source_type'] or 'fss').lower()
    group_code=clean_text(t['standings_group_code'] or '')
    comp_name=clean_text(t['standings_competition_name'] or t['name'])
    rows=[]; used_url=url
    if source=='uefa':
        attempt_errors=[]
        # UEFA standings is JS-rendered for some clients. A failed/empty direct
        # standings request must never prevent the official results fallback.
        try:
            html=fetch_public_html(url,'https://www.uefa.com/')
            rows=parse_uefa_group_standings_html(html,group_code)
        except Exception as exc:
            attempt_errors.append(f'standings: {exc}')
        if not rows:
            for fallback_url in (UEFA_NATIONS_RESULTS_2026, UEFA_NATIONS_RESULTS_BY_TEAM_2026):
                try:
                    html=fetch_public_html(fallback_url,'https://www.uefa.com/')
                    rows=parse_uefa_group_results_html(html,group_code)
                    if not rows:
                        rows=parse_uefa_group_results_anywhere_html(html,group_code)
                    if rows:
                        used_url=fallback_url
                        break
                    attempt_errors.append(f'{fallback_url}: nema parsiranih rezultata')
                except Exception as exc:
                    attempt_errors.append(f'{fallback_url}: {exc}')
        if not rows:
            # Final resilience fallback. This prevents an empty national-team table
            # when UEFA blocks plain HTTP clients or temporarily changes its HTML.
            rows=known_uefa_group_standings(group_code)
            if rows:
                used_url=f'{UEFA_NATIONS_RESULTS_2026}#bundled-{UEFA_NATIONS_KNOWN_RESULTS_UPDATED}'
            else:
                detail=' | '.join(attempt_errors[-3:])
                raise RuntimeError(f'UEFA tabela za grupu {group_code or "?"} nije pronađena.'+(f' {detail}' if detail else ''))
    elif source=='transfermarkt':
        html=fetch_public_html(url,'https://www.transfermarkt.com/')
        rows=parse_standings_html(html)
        if not rows: raise RuntimeError('Tabela nije pronađena na podešenoj Transfermarkt stranici.')
    else:
        html=fetch_public_html(url,'https://fss.rs/')
        rows=parse_standings_html(html)
        if not rows: raise RuntimeError('Tabela nije pronađena na podešenoj FSS stranici.')
    result=save_standings('national_team',team_id,comp_name,rows,source,used_url,t['country_id'],group_name=group_code)
    check_comp,check_rows=standings_for_owner('national_team',team_id)
    result['visible_rows']=len(check_rows)
    if len(check_rows)!=len(rows):
        raise RuntimeError(f'Tabela je izračunata ({len(rows)}), ali read-back vidi {len(check_rows)} redova za team_id={team_id}.')
    result.update({'team':t['name'],'source':source,'group':group_code,'competition':comp_name}); return result


def refresh_country_national_standings(country_id=1):
    con=db(); teams=con.execute("SELECT id,name FROM national_teams WHERE country_id=? AND COALESCE(standings_source_url,'')<>'' ORDER BY id",(country_id,)).fetchall(); con.close()
    out={'teams':len(teams),'ok':0,'rows':0,'visible_rows':0,'errors':[]}
    for t in teams:
        try:
            r=refresh_national_team_standings(t['id']); out['ok']+=1; out['rows']+=r['rows']; out['visible_rows']+=r.get('visible_rows',0)
        except Exception as e: out['errors'].append(f"{t['name']}: {e}")
    return out


def set_national_standings_source(team_id,url):
    url=clean_text(url)
    con=db(); t=con.execute('SELECT * FROM national_teams WHERE id=?',(team_id,)).fetchone()
    if not t: con.close(); raise RuntimeError('Reprezentacija nije pronađena.')
    if not url:
        con.execute("UPDATE national_teams SET standings_source_url=NULL,standings_source_type=NULL,standings_group_code=NULL WHERE id=?",(team_id,)); con.commit(); con.close(); return ''
    source='fss'; group=''; comp_code=''; comp_name=''
    if 'uefa.com/' in url:
        source='uefa'
    elif 'transfermarkt.' in url:
        source='transfermarkt'
        mg=re.search(r'/gruppe/(\d+)',url); mc=re.search(r'/pokalwettbewerb/([A-Za-z0-9_-]+)',url)
        if mg and mc and mc.group(1).upper()=='UNLA':
            group='A'+mg.group(1); comp_code='UNLA'; comp_name='UEFA Nations League A'
            # Use UEFA as live standings source but preserve the TM reference URL.
            con.execute("UPDATE national_teams SET standings_source_type='uefa',standings_source_url='https://www.uefa.com/uefanationsleague/standings/',standings_group_code=?,standings_competition_code=?,standings_competition_name=?,standings_reference_url=? WHERE id=?",
                        (group,comp_code,comp_name,url,team_id)); con.commit(); con.close(); return url
    elif 'fss.rs/' not in url:
        con.close(); raise RuntimeError('Dozvoljeni izvori za nacionalne selekcije su FSS, UEFA ili Transfermarkt.')
    con.execute("UPDATE national_teams SET standings_source_type=?,standings_source_url=?,standings_group_code=CASE WHEN ?<>'' THEN ? ELSE standings_group_code END,standings_competition_code=CASE WHEN ?<>'' THEN ? ELSE standings_competition_code END,standings_competition_name=CASE WHEN ?<>'' THEN ? ELSE standings_competition_name END WHERE id=?",
                (source,url,group,group,comp_code,comp_code,comp_name,comp_name,team_id)); con.commit(); con.close(); return url


def sync_srbijasport_league(lid, records):
    """
    Srbijasport owns league membership; existing TM data is preserved.
    If the same Srbijasport club already exists elsewhere, move it to this league
    instead of creating a duplicate.
    """
    con=db()
    league=con.execute('SELECT * FROM leagues WHERE id=?',(lid,)).fetchone()
    if not league:
        con.close()
        raise RuntimeError('League not found.')

    con.execute('UPDATE clubs SET source_active=0 WHERE league_id=?',(lid,))
    added=0
    updated=0

    for item in records:
        row=con.execute('SELECT * FROM clubs WHERE source_club_url=?',(item['source_club_url'],)).fetchone()

        if not row:
            # Conservative name match inside current league.
            candidates=con.execute('SELECT * FROM clubs WHERE league_id=?',(lid,)).fetchall()
            wanted=norm_club_name(item['source_name'])
            for c in candidates:
                if wanted and wanted in (
                    norm_club_name(c['source_name'] or ''),
                    norm_club_name(c['name'] or ''),
                    norm_club_name(c['cm_name'] or ''),
                ):
                    row=c
                    break

        if row:
            con.execute("""UPDATE clubs SET league_id=?,source_name=?,source_club_url=?,source_active=1,
                           name=CASE WHEN COALESCE(name,'')='' THEN ? ELSE name END,
                           cm_name=CASE WHEN COALESCE(cm_name,'')='' THEN ? ELSE cm_name END
                           WHERE id=?""",
                        (lid,item['source_name'],item['source_club_url'],
                         item['source_name'],game_safe_name(item['source_name']),row['id']))
            updated+=1
        else:
            con.execute("""INSERT INTO clubs(
                league_id,name,tm_name,cm_name,tm_club_id,tm_slug,tm_url,staff_url,
                source_name,source_club_url,source_active,tm_match_status)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?)""",
                (lid,item['source_name'],'',game_safe_name(item['source_name']),'','','','',
                 item['source_name'],item['source_club_url'],1,'unresolved'))
            added+=1

    if records:
        con.execute('UPDATE leagues SET source_expected_clubs=?,source_active=1 WHERE id=?',(len(records),lid))
    con.commit()
    active=con.execute('SELECT COUNT(*) FROM clubs WHERE league_id=? AND source_active=1',(lid,)).fetchone()[0]
    con.close()
    return {'found':len(records),'active':active,'added':added,'updated':updated}


def tm_match_normalize(value):
    s=norm_club_name(value)
    tokens=s.split()
    # Generic club prefixes are weak evidence, not identity.
    weak={'fk','ofk','gfk','rfk','su','fc','fudbalski','klub'}
    tokens=[t for t in tokens if t not in weak]

    # Srbijasport often shows name + city twice (OFK Vrbas Vrbas).
    dedup=[]
    for t in tokens:
        if not dedup or t!=dedup[-1]:
            dedup.append(t)
    return ' '.join(dedup)


def tm_name_score(source_name, tm_name):
    a=tm_match_normalize(source_name)
    b=tm_match_normalize(tm_name)
    if not a or not b:
        return 0.0

    seq=SequenceMatcher(None,a,b).ratio()
    at=set(a.split()); bt=set(b.split())
    jac=len(at & bt)/max(1,len(at | bt))
    contains=1.0 if (a in b or b in a) else 0.0
    score=(0.55*seq)+(0.35*jac)+(0.10*contains)
    return round(min(1.0,score),4)


def parse_tm_search_results(html):
    soup=BeautifulSoup(html,'lxml')
    results={}
    for a in soup.find_all('a',href=True):
        href=clean_text(a.get('href') or '')
        m=re.search(r'/([^/?#]+)/startseite/verein/(\d+)',href)
        if not m:
            continue
        slug,tmid=m.group(1),m.group(2)
        name=clean_text(a.get_text(' ',strip=True))
        if not name:
            img=a.find('img')
            if img:
                name=clean_text(img.get('alt') or img.get('title') or '')
        if not name:
            continue
        results[tmid]={
            'tm_name':name,
            'tm_club_id':tmid,
            'tm_slug':slug,
            'tm_url':f'https://www.transfermarkt.com/{slug}/startseite/verein/{tmid}',
            'staff_url':f'https://www.transfermarkt.com/{slug}/mitarbeiter/verein/{tmid}',
        }
    return list(results.values())


def transfermarkt_search_club(source_name):
    query=clean_text(source_name)
    url='https://www.transfermarkt.com/schnellsuche/ergebnis/schnellsuche?query='+quote_plus(query)
    html=fetch_html(url)
    candidates=parse_tm_search_results(html)
    for c in candidates:
        c['score']=tm_name_score(source_name,c['tm_name'])
    candidates.sort(key=lambda x:x['score'],reverse=True)
    return candidates


def apply_tm_candidate(club_id,candidate,status='matched',note=''):
    con=db()
    con.execute("""UPDATE clubs SET tm_name=?,tm_club_id=?,tm_slug=?,tm_url=?,staff_url=?,
                   tm_match_status=?,tm_match_score=?,tm_candidate_name=?,tm_candidate_url=?,tm_match_note=?
                   WHERE id=?""",
                (candidate['tm_name'],candidate['tm_club_id'],candidate['tm_slug'],
                 candidate['tm_url'],candidate['staff_url'],status,candidate.get('score'),
                 candidate['tm_name'],candidate['tm_url'],note,club_id))
    con.commit(); con.close()


def resolve_tm_for_club(club_id, force=False):
    con=db()
    club=con.execute('''SELECT cl.*,l.section league_section,l.name league_name FROM clubs cl
        JOIN leagues l ON l.id=cl.league_id WHERE cl.id=?''',(club_id,)).fetchone()
    con.close()
    if not club:
        return {'status':'error','message':'Club not found.'}
    if club['tm_url'] and not force:
        return {'status':'matched','message':'Already mapped.'}

    source_name=club['source_name'] or club['name']
    is_women=(club['league_section'] or '').lower()=='women'

    # Men's verified aliases must never be auto-applied to women's clubs with a
    # similar name. Women's mapping starts from a women-specific TM search and
    # only auto-accepts a candidate that is explicitly identifiable as women.
    verified=None if is_women else verified_tm_candidate(source_name)
    if verified:
        apply_tm_candidate(club_id,verified,'matched',f'Verified exact mapping from {source_name}')
        return {'status':'matched','candidate':verified,
                'message':f"Verified match: {verified['tm_name']} (TM ID {verified['tm_club_id']})"}

    search_name=(source_name+' women') if is_women else source_name
    candidates=transfermarkt_search_club(search_name)
    if is_women:
        for c in candidates:
            # Re-score against the real club name, not the added search hint.
            c['score']=tm_name_score(source_name,c['tm_name'])
            marker=norm_club_name(c['tm_name'])
            c['women_marker']=bool(re.search(r'\b(women|woman|frauen|damen|fem|zfk|zfk)\b',marker))
            if c['women_marker']:
                c['score']=min(1.0,c['score']+0.08)
        candidates.sort(key=lambda x:x['score'],reverse=True)
    if not candidates:
        con=db()
        con.execute("""UPDATE clubs SET tm_match_status='unresolved',tm_match_score=NULL,
                       tm_candidate_name=NULL,tm_candidate_url=NULL,
                       tm_match_note='No Transfermarkt candidate found' WHERE id=?""",(club_id,))
        con.commit(); con.close()
        return {'status':'unresolved','message':'No Transfermarkt candidate found.'}

    best=candidates[0]
    second=candidates[1]['score'] if len(candidates)>1 else 0.0
    gap=best['score']-second

    # Auto-accept only confident results. For women's clubs additionally
    # require an explicit women marker to avoid mapping to the men's team.
    auto_ok=(best['score']>=0.82 and (gap>=0.07 or best['score']>=0.94))
    if is_women:
        auto_ok=auto_ok and bool(best.get('women_marker')) and best['score']>=0.88
    if auto_ok:
        apply_tm_candidate(club_id,best,'matched',f'Auto matched from {source_name}')
        return {'status':'matched','candidate':best,'message':f"Matched: {best['tm_name']} ({best['score']:.2f})"}

    con=db()
    con.execute("""UPDATE clubs SET tm_match_status='review',tm_match_score=?,
                   tm_candidate_name=?,tm_candidate_url=?,
                   tm_match_note=? WHERE id=?""",
                (best['score'],best['tm_name'],best['tm_url'],
                 f'Best automatic candidate; second score {second:.2f}',club_id))
    con.commit(); con.close()
    return {'status':'review','candidate':best,'message':f"Review: {best['tm_name']} ({best['score']:.2f})"}


def resolve_tm_for_league(lid):
    con=db()
    no_tm_yet=con.execute("""SELECT COUNT(*) FROM clubs WHERE league_id=? AND source_active=1
                              AND COALESCE(tm_url,'')='' AND tm_match_status='no_tm_yet'""",(lid,)).fetchone()[0]
    rows=con.execute("""SELECT id FROM clubs WHERE league_id=? AND source_active=1
                        AND COALESCE(tm_url,'')='' AND COALESCE(tm_match_status,'unresolved')<>'no_tm_yet'
                        ORDER BY name""",(lid,)).fetchall()
    con.close()
    out={'total':len(rows),'matched':0,'review':0,'unresolved':0,'no_tm_yet':no_tm_yet,'errors':[]}
    for i,r in enumerate(rows,1):
        try:
            result=resolve_tm_for_club(r['id'])
            st=result['status']
            if st in out: out[st]+=1
        except Exception as e:
            out['errors'].append(str(e))
        if i<len(rows): time.sleep(0.9)
    return out

def resolve_tm_for_level(cid,level):
    """Resolve Transfermarkt mappings for a whole level by reusing the proven per-league resolver."""
    con=db()
    leagues=con.execute("""SELECT id,name FROM leagues
        WHERE country_id=? AND level=? AND COALESCE(section,'senior')='senior'
        ORDER BY name""",(cid,level)).fetchall()
    total_active=con.execute("""SELECT COUNT(*) FROM clubs cl JOIN leagues l ON l.id=cl.league_id
        WHERE l.country_id=? AND l.level=? AND COALESCE(cl.source_active,1)=1""",(cid,level)).fetchone()[0]
    already_mapped=con.execute("""SELECT COUNT(*) FROM clubs cl JOIN leagues l ON l.id=cl.league_id
        WHERE l.country_id=? AND l.level=? AND COALESCE(cl.source_active,1)=1 AND COALESCE(cl.tm_url,'')<>''""",(cid,level)).fetchone()[0]
    con.close()

    out={'active':total_active,'already_mapped':already_mapped,'checked':0,'no_tm_yet':0,
         'matched':0,'review':0,'unresolved':0,'errors':[],'leagues':len(leagues)}

    for idx,l in enumerate(leagues,1):
        print(f'[LEVEL {level} LEAGUE MAP {idx}/{len(leagues)}] {l["name"]}')
        try:
            result=resolve_tm_for_league(l['id'])
            out['checked'] += int(result.get('total',0) or 0)
            out['matched'] += int(result.get('matched',0) or 0)
            out['review'] += int(result.get('review',0) or 0)
            out['unresolved'] += int(result.get('unresolved',0) or 0)
            out['no_tm_yet'] += int(result.get('no_tm_yet',0) or 0)
            for err in result.get('errors',[]): out['errors'].append(f'{l["name"]}: {err}')
        except Exception as e:
            out['errors'].append(f'{l["name"]}: {e}')
        if idx<len(leagues): time.sleep(1.0)
    return out




def refresh_section_memberships(cid,section):
    con=db(); leagues=con.execute('''SELECT * FROM leagues WHERE country_id=? AND section=? AND COALESCE(source_active,1)=1
        ORDER BY COALESCE(level,999),name''',(cid,section)).fetchall(); con.close()
    out={'leagues':len(leagues),'ok':0,'found':0,'errors':[]}
    for l in leagues:
        try:
            r=refresh_league_membership_live(l)
            out['ok']+=1; out['found']+=int(r.get('found',0) or 0)
        except Exception as e:
            out['errors'].append(f"{l['name']}: {e}")
    return out


def resolve_tm_for_section(cid,section):
    con=db(); leagues=con.execute('''SELECT id,name FROM leagues WHERE country_id=? AND section=? AND COALESCE(source_active,1)=1
        ORDER BY COALESCE(level,999),name''',(cid,section)).fetchall(); con.close()
    out={'leagues':len(leagues),'checked':0,'matched':0,'review':0,'unresolved':0,'no_tm_yet':0,'errors':[]}
    for l in leagues:
        try:
            r=resolve_tm_for_league(l['id'])
            out['checked']+=int(r.get('total',0) or 0)
            for k in ('matched','review','unresolved','no_tm_yet'):
                out[k]+=int(r.get(k,0) or 0)
            out['errors'].extend(f"{l['name']}: {e}" for e in r.get('errors',[]))
        except Exception as e:
            out['errors'].append(f"{l['name']}: {e}")
    return out


def batch_update_section(cid,section,fetch_staff_dob=False):
    con=db(); leagues=con.execute('''SELECT id,name FROM leagues WHERE country_id=? AND section=? AND COALESCE(source_active,1)=1
        ORDER BY COALESCE(level,999),name''',(cid,section)).fetchall(); con.close()
    out={'leagues':len(leagues),'clubs':0,'ok':0,'players':0,'staff':0,'skipped':0,'errors':[]}
    for l in leagues:
        try:
            r=batch_update_league(l['id'],fetch_staff_dob)
            out['clubs']+=int(r.get('clubs',0) or 0); out['ok']+=int(r.get('ok',0) or 0)
            out['players']+=int(r.get('players',0) or 0); out['staff']+=int(r.get('staff',0) or 0)
            out['skipped']+=int(r.get('skipped',0) or 0)
            out['errors'].extend(f"{l['name']}: {e}" for e in r.get('errors',[]))
        except Exception as e:
            out['errors'].append(f"{l['name']}: {e}")
    return out


def batch_update_section_transfers(cid,section):
    con=db(); leagues=con.execute('''SELECT id,name FROM leagues WHERE country_id=? AND section=? AND COALESCE(source_active,1)=1
        ORDER BY COALESCE(level,999),name''',(cid,section)).fetchall(); con.close()
    out={'leagues':len(leagues),'mapped':0,'ok':0,'skipped':0,'rows':0,'errors':[]}
    for l in leagues:
        try:
            r=batch_update_league_transfers(l['id'])
            out['mapped']+=int(r.get('mapped',0) or 0); out['ok']+=int(r.get('ok',0) or 0)
            out['skipped']+=int(r.get('skipped',0) or 0); out['rows']+=int(r.get('rows',0) or 0)
            out['errors'].extend(f"{l['name']}: {e}" for e in r.get('errors',[]))
        except Exception as e:
            out['errors'].append(f"{l['name']}: {e}")
    return out


def update_women_national_teams(country_id):
    con=db(); teams=con.execute("""SELECT id,name FROM national_teams WHERE country_id=?
        AND team_type IN ('women_senior','women_u19') ORDER BY id""",(country_id,)).fetchall(); con.close()
    out={'teams':len(teams),'ok':0,'players':0,'staff':0,'errors':[]}
    for t in teams:
        try:
            r=update_national_team_live(t['id'],False)
            out['ok']+=1; out['players']+=int(r.get('players',0) or 0); out['staff']+=int(r.get('staff',0) or 0)
        except Exception as e:
            out['errors'].append(f"{t['name']}: {e}")
    return out

def resolve_tm_for_country_all(cid):
    """Resolve Transfermarkt mappings for every league currently stored for a country.
    Future leagues are picked up automatically; no hard-coded level/section list."""
    con=db()
    leagues=con.execute("""SELECT id,name,level,section FROM leagues
        WHERE country_id=? AND COALESCE(source_active,1)=1 ORDER BY COALESCE(level,999),COALESCE(section,'senior'),name""",(cid,)).fetchall()
    total_active=con.execute("""SELECT COUNT(*) FROM clubs cl JOIN leagues l ON l.id=cl.league_id
        WHERE l.country_id=? AND COALESCE(l.source_active,1)=1 AND COALESCE(cl.source_active,1)=1""",(cid,)).fetchone()[0]
    already_mapped=con.execute("""SELECT COUNT(*) FROM clubs cl JOIN leagues l ON l.id=cl.league_id
        WHERE l.country_id=? AND COALESCE(l.source_active,1)=1 AND COALESCE(cl.source_active,1)=1 AND COALESCE(cl.tm_url,'')<>''""",(cid,)).fetchone()[0]
    con.close()
    out={'active':total_active,'already_mapped':already_mapped,'checked':0,'no_tm_yet':0,
         'matched':0,'review':0,'unresolved':0,'errors':[],'leagues':len(leagues)}
    for idx,l in enumerate(leagues,1):
        print(f'[ALL LEAGUES TM MAP {idx}/{len(leagues)}] {l["name"]}')
        try:
            r=resolve_tm_for_league(l['id'])
            for key in ('total','matched','review','unresolved','no_tm_yet'):
                dest='checked' if key=='total' else key
                out[dest]+=int(r.get(key,0) or 0)
            out['errors'].extend(f'{l["name"]}: {e}' for e in r.get('errors',[]))
        except Exception as e:
            out['errors'].append(f'{l["name"]}: {e}')
        if idx<len(leagues): time.sleep(0.8)
    return out


def refresh_country_memberships_all(cid, srbijasport_only=False):
    """Refresh memberships from each league's configured source.
    If srbijasport_only=True, refresh only leagues backed by SrbijaSport.
    New leagues are included automatically when they have source_type/source_url or tm_url configured."""
    # Reconcile the full current lower-league pyramid first. Hard-coded
    # season IDs are no longer authoritative.
    try: discover_srbijasport_pyramid(cid,False,levels=(3,4,5,6))
    except Exception: pass
    con=db()
    if srbijasport_only:
        leagues=con.execute("""SELECT * FROM leagues WHERE country_id=?
            AND LOWER(COALESCE(source_type,''))='srbijasport' AND COALESCE(source_url,'')<>''
            AND COALESCE(source_active,1)=1
            ORDER BY COALESCE(level,999),name""",(cid,)).fetchall()
    else:
        leagues=con.execute("""SELECT * FROM leagues WHERE country_id=?
            AND COALESCE(source_active,1)=1
            ORDER BY COALESCE(level,999),COALESCE(section,'senior'),name""",(cid,)).fetchall()
    con.close()
    out={'leagues':len(leagues),'ok':0,'found':0,'errors':[],'details':[]}
    for idx,l in enumerate(leagues,1):
        print(f'[ALL MEMBERSHIP {idx}/{len(leagues)}] {l["name"]}')
        try:
            rr=refresh_league_membership_live(l)
            out['ok']+=1
            out['found']+=int(rr.get('found',0) or 0)
            out['details'].append(f"{l['name']}: {rr.get('source','')} {rr.get('found',0)}; {rr.get('membership','')}")
        except Exception as e:
            out['errors'].append(f'{l["name"]}: {e}')
    return out


def batch_update_country_all_leagues(cid, fetch_staff_dob=False, refresh_membership=False):
    """Update players/staff/info for every league currently stored for a country.
    No hard-coded levels or sections: future women's/youth/Level 6 leagues are included automatically."""
    con=db()
    country=con.execute('SELECT * FROM countries WHERE id=?',(cid,)).fetchone()
    leagues=con.execute("""SELECT * FROM leagues WHERE country_id=? AND COALESCE(source_active,1)=1
        ORDER BY COALESCE(level,999),COALESCE(section,'senior'),name""",(cid,)).fetchall()
    con.close()
    if not country: raise RuntimeError('Country not found.')
    out={'country':country['name'],'leagues':len(leagues),'clubs':0,'clubs_ok':0,
         'players':0,'staff':0,'membership':[],'errors':[]}
    for idx,l in enumerate(leagues,1):
        lname=l['name']
        print(f'[ALL LEAGUES UPDATE {idx}/{len(leagues)}] {lname}')
        if refresh_membership:
            try:
                rr=refresh_league_membership_live(l)
                out['membership'].append(f"{lname}: {rr.get('source','')} found {rr.get('found',0)}; {rr.get('membership','')}")
            except Exception as e:
                out['errors'].append(f'{lname} membership: {e}')
        try:
            b=batch_update_league(l['id'],fetch_staff_dob)
            out['clubs']+=int(b.get('clubs',0) or 0); out['clubs_ok']+=int(b.get('ok',0) or 0)
            out['players']+=int(b.get('players',0) or 0); out['staff']+=int(b.get('staff',0) or 0)
            out['errors'].extend(f'{lname}: {e}' for e in b.get('errors',[]))
        except Exception as e:
            out['errors'].append(f'{lname} update: {e}')
    return out


def batch_update_country_transfers_all(cid):
    """Update transfers for all active TM-mapped clubs in every league/section."""
    con=db()
    active_total=con.execute("""SELECT COUNT(*) FROM clubs cl JOIN leagues l ON l.id=cl.league_id
        WHERE l.country_id=? AND COALESCE(l.source_active,1)=1 AND COALESCE(cl.source_active,1)=1""",(cid,)).fetchone()[0]
    clubs=con.execute("""SELECT cl.*,l.name league_name,l.level league_level,l.section league_section FROM clubs cl
        JOIN leagues l ON l.id=cl.league_id WHERE l.country_id=? AND COALESCE(l.source_active,1)=1 AND COALESCE(cl.source_active,1)=1
        AND COALESCE(cl.tm_url,'')<>'' ORDER BY COALESCE(l.level,999),COALESCE(l.section,'senior'),l.name,cl.name""",(cid,)).fetchall()
    con.close()
    out={'active':active_total,'mapped':len(clubs),'skipped':max(0,active_total-len(clubs)),
         'ok':0,'with_data':0,'empty':0,'rows':0,'arrivals':0,'departures':0,'errors':[]}
    session=requests.Session()
    for i,c in enumerate(clubs,1):
        cname=c['cm_name'] or c['name']
        print(f'[ALL TRANSFERS {i}/{len(clubs)}] {c["league_name"]} / {cname}')
        try:
            sid=current_season_id(); url=transfer_detailed_url(c,sid)
            html=fetch_html(url,session); summary,records=parse_transfers_html(html,url,sid)
            low=html.lower()
            if not records and not (summary.get('arrivals_count') or summary.get('departures_count')):
                if 'arrivals' not in low and 'departures' not in low and 'transfer record' not in low:
                    raise RuntimeError('Stranica transfera nije mogla da se obradi.')
            save_club_transfers(c['id'],summary,records,'batch all leagues transfers')
            out['ok']+=1; out['rows']+=len(records)
            out['arrivals']+=int(summary.get('arrivals_count') or 0); out['departures']+=int(summary.get('departures_count') or 0)
            if records or summary.get('arrivals_count') or summary.get('departures_count'): out['with_data']+=1
            else: out['empty']+=1
        except Exception as e:
            out['errors'].append(f'{c["league_name"]} / {cname}: {e}')
        if i<len(clubs): time.sleep(0.8)
    return out

def mapping_status(row):
    if clean_text(row['tm_url'] or ''): return 'MATCHED'
    raw=clean_text(row['tm_match_status'] or '').lower()
    if raw=='review': return 'REVIEW'
    if raw=='no_tm_yet': return 'NO TM YET'
    return 'UNRESOLVED'

def mapping_method(row):
    note=clean_text(row['tm_match_note'] or ''); low=note.lower(); status=mapping_status(row)
    if 'manual' in low: return 'MANUAL'
    if 'verified' in low: return 'VERIFIED'
    if 'auto matched' in low: return 'AUTO'
    if status=='REVIEW': return 'REVIEW CANDIDATE'
    if status=='NO TM YET': return 'NO TM YET'
    if status=='MATCHED': return 'RESOLVER / EXISTING'
    return 'UNRESOLVED'

def level_mapping_rows(cid,level,issues_only=False):
    con=db()
    country=con.execute('SELECT * FROM countries WHERE id=?',(cid,)).fetchone()
    rows=con.execute("""SELECT cl.*,l.name league_name,l.level league_level,
        l.source_type league_source,l.source_url league_source_url
        FROM clubs cl JOIN leagues l ON l.id=cl.league_id
        WHERE l.country_id=? AND l.level=? AND COALESCE(cl.source_active,1)=1
        ORDER BY l.name,COALESCE(cl.source_name,cl.name),cl.name""",(cid,level)).fetchall()
    con.close()
    if not country:
        raise RuntimeError('Country not found.')

    out=[]
    for r in rows:
        status=mapping_status(r)
        if issues_only and status=='MATCHED':
            continue
        out.append({
            'Country':country['name'],
            'Level':level,
            'League':r['league_name'],
            'League Source':r['league_source'],
            'SrbijaSport Name':r['source_name'] or r['name'],
            'SrbijaSport URL':r['source_club_url'],
            'Data Hub Name':r['cm_name'] or r['name'],
            'TM Status':status,
            'Mapping Method':mapping_method(r),
            'TM Score':r['tm_match_score'],
            'TM Name':r['tm_name'],
            'TM ID':r['tm_club_id'],
            'TM Slug':r['tm_slug'],
            'TM URL':r['tm_url'],
            'TM Staff URL':r['staff_url'],
            'Candidate Name':r['tm_candidate_name'],
            'Candidate URL':r['tm_candidate_url'],
            'Mapping Note':r['tm_match_note'],
            'Players Updated':r['last_players_update'],
            'Staff Updated':r['last_staff_update'],
        })
    return country,out


MAPPING_COLUMNS = [
    'Country','Level','League','League Source','SrbijaSport Name','SrbijaSport URL',
    'Data Hub Name','TM Status','Mapping Method','TM Score','TM Name','TM ID','TM Slug',
    'TM URL','TM Staff URL','Candidate Name','Candidate URL','Mapping Note',
    'Players Updated','Staff Updated'
]


def level_mapping_csv_bytes(cid,level,issues_only=False):
    country,rows=level_mapping_rows(cid,level,issues_only)
    sio=io.StringIO()
    writer=csv.writer(sio)
    writer.writerow(MAPPING_COLUMNS)
    for r in rows:
        writer.writerow([r.get(c,'') for c in MAPPING_COLUMNS])
    return sio.getvalue().encode('utf-8-sig')


def _mapping_status_style(cell,status):
    status=clean_text(status).upper()
    if status=='MATCHED':
        cell.fill=PatternFill('solid',fgColor='1D6B50'); cell.font=Font(color='FFFFFF',bold=True)
    elif status=='REVIEW':
        cell.fill=PatternFill('solid',fgColor='99751A'); cell.font=Font(color='FFFFFF',bold=True)
    elif status=='NO TM YET':
        cell.fill=PatternFill('solid',fgColor='385A78'); cell.font=Font(color='FFFFFF',bold=True)
    else:
        cell.fill=PatternFill('solid',fgColor='8A3544'); cell.font=Font(color='FFFFFF',bold=True)

def mapping_workbook_bytes(cid,level,issues_only=False):
    country,rows=level_mapping_rows(cid,level,issues_only)
    _,all_rows=level_mapping_rows(cid,level,False)
    matched=sum(1 for r in all_rows if r['TM Status']=='MATCHED')
    review=sum(1 for r in all_rows if r['TM Status']=='REVIEW')
    unresolved=sum(1 for r in all_rows if r['TM Status']=='UNRESOLVED')
    no_tm_yet=sum(1 for r in all_rows if r['TM Status']=='NO TM YET')
    verified=matched+no_tm_yet
    wb=Workbook(); ws=wb.active; ws.title='Summary'
    title='Srpske lige' if level==3 else ('Zonske lige' if level==4 else f'Level {level}')
    for row in [
        ['Field','Value'],['Country',country['name']],['Competition group',title],['Level',level],
        ['Active clubs',len(all_rows)],['Matched',matched],['No TM Yet',no_tm_yet],
        ['Verified state',f'{verified}/{len(all_rows)}'],['Review',review],['Unresolved',unresolved],
        ['Export scope','OPEN + NO TM YET' if issues_only else 'ALL active clubs'],
        ['Generated',time.strftime('%Y-%m-%d %H:%M:%S')]]: ws.append(row)
    style_sheet(ws,[28,70])
    data=wb.create_sheet('Mapping' if not issues_only else 'Issues'); data.append(MAPPING_COLUMNS)
    for r in rows: data.append([r.get(c,'') for c in MAPPING_COLUMNS])
    widths=[14,8,30,18,30,55,28,16,20,12,28,12,28,58,58,28,58,42,20,20]; style_sheet(data,widths)
    if data.max_row>1:
        data.auto_filter.ref=f'A1:T{data.max_row}'
        for ri in range(2,data.max_row+1): _mapping_status_style(data.cell(ri,8),data.cell(ri,8).value)
    if not issues_only:
        issues=[r for r in all_rows if r['TM Status']!='MATCHED']; iss=wb.create_sheet('Issues'); iss.append(MAPPING_COLUMNS)
        for r in issues: iss.append([r.get(c,'') for c in MAPPING_COLUMNS])
        style_sheet(iss,widths)
        if iss.max_row>1:
            iss.auto_filter.ref=f'A1:T{iss.max_row}'
            for ri in range(2,iss.max_row+1): _mapping_status_style(iss.cell(ri,8),iss.cell(ri,8).value)
    bio=io.BytesIO(); wb.save(bio); return bio.getvalue()

def level_mapping_backup_zip_bytes(cid,level):
    country,rows=level_mapping_rows(cid,level,False)
    matched=sum(1 for r in rows if r['TM Status']=='MATCHED')
    review=sum(1 for r in rows if r['TM Status']=='REVIEW')
    unresolved=sum(1 for r in rows if r['TM Status']=='UNRESOLVED')
    no_tm_yet=sum(1 for r in rows if r['TM Status']=='NO TM YET'); verified=matched+no_tm_yet
    title='SRPSKE_LIGE' if level==3 else ('ZONSKE_LIGE' if level==4 else f'LEVEL_{level}')
    summary={'country':country['name'],'level':level,'group':title,'generated':time.strftime('%Y-%m-%d %H:%M:%S'),
             'active_clubs':len(rows),'matched':matched,'no_tm_yet':no_tm_yet,'verified':verified,
             'review':review,'unresolved':unresolved}
    bio=io.BytesIO()
    with zipfile.ZipFile(bio,'w',zipfile.ZIP_DEFLATED) as z:
        z.writestr(f'{title}_MAPPING.xlsx',mapping_workbook_bytes(cid,level,False))
        z.writestr(f'{title}_MAPPING.csv',level_mapping_csv_bytes(cid,level,False))
        z.writestr(f'{title}_OPEN_AND_NO_TM_YET.csv',level_mapping_csv_bytes(cid,level,True))
        z.writestr('mapping_summary.json',json.dumps(summary,ensure_ascii=False,indent=2).encode('utf-8'))
        readme=('CM DATA HUB - MAPPING BACKUP\\n'+f'Country: {country["name"]}\\nLevel: {level}\\nGroup: {title}\\n'
                +f'Active clubs: {len(rows)}\\nMatched: {matched}\\nNo TM Yet: {no_tm_yet}\\nVerified: {verified}/{len(rows)}\\nReview: {review}\\nUnresolved: {unresolved}\\n\\n'
                +'NO TM YET means manually checked: no profile found yet. It remains recheckable and manually mappable.\\n')
        z.writestr('README.txt',readme.encode('utf-8-sig'))
    return bio.getvalue()

def clean_mapping_search_name(value):
    # Create a cleaner club name for lower-league Transfermarkt searches.
    s=clean_text(value)
    s=re.sub(r'\s+', ' ', s).strip()
    noise_patterns=[
        r'\s+\d{1,2}\s+PFL\s+[^,|]+$',
        r'\s+\d{1,2}\s+(?:Sremska|Mačvanska|Macvanska|Područna|Podrucna|Međuopštinska|Medjuopstinska|Opštinska|Opstinska)\s+liga.*$',
        r'\s+\d{1,2}\s+(?:liga|zone|zona)\b.*$',
    ]
    for pat in noise_patterns:
        s=re.sub(pat,'',s,flags=re.I)
    s=re.sub(r'\s+\d{1,2}\s*$','',s)
    words=s.split()
    if len(words)>=2 and words[-1].casefold()==words[-2].casefold():
        words.pop()
    elif len(words)>=4 and ' '.join(words[-2:]).casefold()==' '.join(words[-4:-2]).casefold():
        words=words[:-2]
    s=' '.join(words)
    s=re.sub(r'\s*\([^)]{1,6}\)\s*',' ',s)
    return clean_text(s)


def mapping_search_variants(value):
    original=clean_text(value)
    clean=clean_mapping_search_name(original)
    variants=[]
    def add(v):
        v=clean_text(v)
        if v and v.casefold() not in [x.casefold() for x in variants]:
            variants.append(v)
    add(original); add(clean)
    for v in list(variants):
        no_prefix=re.sub(r'^(?:FK|OFK|GFK|RFK|FKK|SK|BSK|NK)\s+','',v,flags=re.I)
        add(no_prefix)
        if no_prefix and no_prefix.casefold()==v.casefold():
            add('FK '+v)
    add(re.sub(r'\s*\([^)]{1,6}\)\s*',' ',original))
    return variants[:6]


def transfermarkt_search_variants(source_name,max_candidates=5):
    variants=mapping_search_variants(source_name)
    merged={}; errors=[]
    for query in variants:
        try:
            candidates=transfermarkt_search_club(query)
        except Exception as e:
            errors.append(f'{query}: {e}')
            continue
        for c in candidates:
            key=c['tm_club_id']
            score=max(tm_name_score(v,c['tm_name']) for v in variants)
            item=dict(c); item['score']=score; item['query']=query
            if key not in merged or score>merged[key]['score']:
                merged[key]=item
        time.sleep(0.35)
    out=sorted(merged.values(),key=lambda x:x['score'],reverse=True)
    return out[:max_candidates],variants,errors


def retry_tm_clean_club(club_id):
    con=db(); club=con.execute('SELECT * FROM clubs WHERE id=?',(club_id,)).fetchone(); con.close()
    if not club:
        return {'status':'error','message':'Club not found.'}
    if clean_text(club['tm_url'] or ''):
        return {'status':'matched','message':'Club is already mapped.'}
    source_name=club['source_name'] or club['name']
    candidates,variants,errors=transfermarkt_search_variants(source_name,5)
    clean_name=clean_mapping_search_name(source_name)
    if not candidates:
        con=db(); con.execute("""UPDATE clubs SET tm_match_status='unresolved',tm_match_score=NULL,
                       tm_candidate_name=NULL,tm_candidate_url=NULL,tm_match_note=? WHERE id=?""",
                    (f'Cleaned search found no candidate. Clean name: {clean_name}',club_id)); con.commit(); con.close()
        return {'status':'unresolved','message':f'No candidate found for {clean_name}.','errors':errors}
    best=candidates[0]; second=candidates[1]['score'] if len(candidates)>1 else 0.0; gap=best['score']-second
    if best['score']>=0.92 and (gap>=0.08 or best['score']>=0.97):
        apply_tm_candidate(club_id,best,'matched',f'Cleaned-name auto match from {source_name}; query={best.get("query","")}')
        return {'status':'matched','candidate':best,'message':f'Cleaned search matched {best["tm_name"]} ({best["score"]:.2f}).'}
    con=db(); con.execute("""UPDATE clubs SET tm_match_status='review',tm_match_score=?,tm_candidate_name=?,tm_candidate_url=?,tm_match_note=? WHERE id=?""",
                (best['score'],best['tm_name'],best['tm_url'],f'Cleaned search candidate; query={best.get("query","")}; clean={clean_name}; second={second:.2f}',club_id)); con.commit(); con.close()
    return {'status':'review','candidate':best,'message':f'Review candidate: {best["tm_name"]} ({best["score"]:.2f}).','errors':errors}


def retry_clean_level(cid,level):
    con=db(); rows=con.execute("""SELECT cl.id,cl.name,l.name league_name FROM clubs cl JOIN leagues l ON l.id=cl.league_id
        WHERE l.country_id=? AND l.level=? AND COALESCE(cl.source_active,1)=1 AND COALESCE(cl.tm_url,'')=''
        AND COALESCE(cl.tm_match_status,'unresolved')<>'no_tm_yet' ORDER BY l.name,cl.name""",(cid,level)).fetchall(); con.close()
    out={'checked':len(rows),'matched':0,'review':0,'unresolved':0,'errors':[]}
    for i,r in enumerate(rows,1):
        print(f'[CLEAN MAP {i}/{len(rows)}] {r["league_name"]} / {r["name"]}')
        try:
            result=retry_tm_clean_club(r['id']); st=result.get('status','unresolved')
            if st in out: out[st]+=1
            if result.get('errors'): out['errors'].extend(result['errors'][:2])
        except Exception as e: out['errors'].append(f'{r["league_name"]} / {r["name"]}: {e}')
        if i<len(rows): time.sleep(0.55)
    return out

def mark_no_tm_yet(club_id):
    con=db(); club=con.execute('SELECT * FROM clubs WHERE id=?',(club_id,)).fetchone()
    if not club: con.close(); raise RuntimeError('Club not found.')
    if clean_text(club['tm_url'] or ''): con.close(); raise RuntimeError('Club is already mapped to Transfermarkt.')
    stamp=time.strftime('%Y-%m-%d')
    con.execute("""UPDATE clubs SET tm_match_status='no_tm_yet',tm_match_score=NULL,
                   tm_candidate_name=NULL,tm_candidate_url=NULL,tm_match_note=? WHERE id=?""",
                (f'Manually checked {stamp}: no Transfermarkt profile found yet',club_id))
    con.commit(); con.close(); return 'Marked NO TM YET. It can be rechecked or manually mapped later.'


def recheck_no_tm_club(club_id):
    con=db(); club=con.execute('SELECT * FROM clubs WHERE id=?',(club_id,)).fetchone(); con.close()
    if not club: return {'status':'error','message':'Club not found.'}
    if clean_text(club['tm_url'] or ''): return {'status':'matched','message':'Club is already mapped.'}
    source_name=club['source_name'] or club['name']; candidates,variants,errors=transfermarkt_search_variants(source_name,5)
    clean_name=clean_mapping_search_name(source_name); stamp=time.strftime('%Y-%m-%d')
    if not candidates:
        con=db(); con.execute("""UPDATE clubs SET tm_match_status='no_tm_yet',tm_match_score=NULL,
                       tm_candidate_name=NULL,tm_candidate_url=NULL,tm_match_note=? WHERE id=?""",
                    (f'Rechecked {stamp}: still no Transfermarkt profile candidate. Clean name: {clean_name}',club_id)); con.commit(); con.close()
        return {'status':'no_tm_yet','message':f'Still no Transfermarkt candidate for {clean_name}.','errors':errors}
    best=candidates[0]; second=candidates[1]['score'] if len(candidates)>1 else 0.0; gap=best['score']-second
    if best['score']>=0.92 and (gap>=0.08 or best['score']>=0.97):
        apply_tm_candidate(club_id,best,'matched',f'NO TM YET recheck auto match from {source_name}; query={best.get("query","")}')
        return {'status':'matched','candidate':best,'message':f'New Transfermarkt profile matched: {best["tm_name"]} ({best["score"]:.2f}).'}
    con=db(); con.execute("""UPDATE clubs SET tm_match_status='review',tm_match_score=?,tm_candidate_name=?,tm_candidate_url=?,tm_match_note=? WHERE id=?""",
                (best['score'],best['tm_name'],best['tm_url'],f'NO TM YET recheck candidate {stamp}; query={best.get("query","")}; second={second:.2f}',club_id)); con.commit(); con.close()
    return {'status':'review','candidate':best,'message':f'New candidate found for review: {best["tm_name"]} ({best["score"]:.2f}).','errors':errors}


def recheck_no_tm_level(cid,level):
    con=db(); rows=con.execute("""SELECT cl.id,cl.name,l.name league_name FROM clubs cl JOIN leagues l ON l.id=cl.league_id
        WHERE l.country_id=? AND l.level=? AND COALESCE(cl.source_active,1)=1
        AND COALESCE(cl.tm_url,'')='' AND cl.tm_match_status='no_tm_yet' ORDER BY l.name,cl.name""",(cid,level)).fetchall(); con.close()
    out={'checked':len(rows),'matched':0,'review':0,'no_tm_yet':0,'errors':[]}
    for i,r in enumerate(rows,1):
        print(f'[NO TM RECHECK {i}/{len(rows)}] {r["league_name"]} / {r["name"]}')
        try:
            result=recheck_no_tm_club(r['id']); st=result.get('status','no_tm_yet')
            if st in out: out[st]+=1
            if result.get('errors'): out['errors'].extend(result['errors'][:2])
        except Exception as e: out['errors'].append(f'{r["league_name"]} / {r["name"]}: {e}')
        if i<len(rows): time.sleep(0.55)
    return out


def accept_tm_candidate(club_id):
    con=db(); club=con.execute('SELECT * FROM clubs WHERE id=?',(club_id,)).fetchone(); con.close()
    if not club: raise RuntimeError('Club not found.')
    url=clean_text(club['tm_candidate_url'] or '')
    if not url: raise RuntimeError('This club has no stored Transfermarkt candidate.')
    canonical=set_manual_tm_url(club_id,url)
    con=db(); con.execute("UPDATE clubs SET tm_match_note='Accepted Mapping Workbench candidate',tm_match_score=1.0 WHERE id=?",(club_id,)); con.commit(); con.close()
    return canonical


def import_mapping_xlsx_bytes(data,cid,level):
    wb=load_workbook(io.BytesIO(data),data_only=True,read_only=True)
    ws=None
    for name in ('Issues','Mapping'):
        if name in wb.sheetnames: ws=wb[name]; break
    if ws is None: raise RuntimeError('Workbook must contain an Issues or Mapping sheet.')
    headers={clean_text(cell.value):idx for idx,cell in enumerate(next(ws.iter_rows(min_row=1,max_row=1)),start=1)}
    for h in ('SrbijaSport URL','TM URL'):
        if h not in headers: raise RuntimeError(f'Missing required column: {h}')
    result={'rows':0,'mapped':0,'skipped':0,'not_found':0,'errors':[]}
    for row in ws.iter_rows(min_row=2,values_only=True):
        result['rows']+=1
        source_url=clean_text(row[headers['SrbijaSport URL']-1] if len(row)>=headers['SrbijaSport URL'] else '')
        tm_url=clean_text(row[headers['TM URL']-1] if len(row)>=headers['TM URL'] else '')
        if not source_url or not tm_url:
            result['skipped']+=1; continue
        con=db(); club=con.execute("""SELECT cl.* FROM clubs cl JOIN leagues l ON l.id=cl.league_id WHERE cl.source_club_url=? AND l.country_id=? AND l.level=?""",(source_url,cid,level)).fetchone(); con.close()
        if not club:
            result['not_found']+=1; continue
        try:
            set_manual_tm_url(club['id'],tm_url)
            con=db(); con.execute("UPDATE clubs SET tm_match_note='Imported from Mapping XLSX',tm_match_score=1.0 WHERE id=?",(club['id'],)); con.commit(); con.close(); result['mapped']+=1
        except Exception as e:
            result['errors'].append(f'{source_url}: {e}')
    return result


def mapping_workbench_rows(cid,level):
    con=db(); country=con.execute('SELECT * FROM countries WHERE id=?',(cid,)).fetchone()
    rows=con.execute("""SELECT cl.*,l.name league_name FROM clubs cl JOIN leagues l ON l.id=cl.league_id
        WHERE l.country_id=? AND l.level=? AND COALESCE(cl.source_active,1)=1 AND COALESCE(cl.tm_url,'')=''
        ORDER BY CASE WHEN cl.tm_match_status='review' THEN 0 WHEN cl.tm_match_status='no_tm_yet' THEN 2 ELSE 1 END,
        l.name,COALESCE(cl.source_name,cl.name)""",(cid,level)).fetchall(); con.close()
    return country,rows

def page_mapping_workbench(cid,level,msg=''):
    country,rows=mapping_workbench_rows(cid,level)
    if not country: return layout('Not found','<h1>Country not found</h1>')
    title='Srpske lige' if level==3 else ('Zonske lige' if level==4 else f'Level {level}')
    review=sum(1 for r in rows if clean_text(r['tm_match_status'] or '').lower()=='review')
    no_tm_yet=sum(1 for r in rows if clean_text(r['tm_match_status'] or '').lower()=='no_tm_yet')
    unresolved=len(rows)-review-no_tm_yet
    con=db(); total_active=con.execute("""SELECT COUNT(*) FROM clubs cl JOIN leagues l ON l.id=cl.league_id
        WHERE l.country_id=? AND l.level=? AND COALESCE(cl.source_active,1)=1""",(cid,level)).fetchone()[0]
    mapped=con.execute("""SELECT COUNT(*) FROM clubs cl JOIN leagues l ON l.id=cl.league_id
        WHERE l.country_id=? AND l.level=? AND COALESCE(cl.source_active,1)=1 AND COALESCE(cl.tm_url,'')<>''""",(cid,level)).fetchone()[0]; con.close()
    verified=mapped+no_tm_yet; notice=f'<div class="notice">{esc(msg)}</div>' if msg else ''
    trs=[]
    for r in rows:
        source=r['source_name'] or r['name']; clean_name=clean_mapping_search_name(source); raw=clean_text(r['tm_match_status'] or '').lower()
        status='REVIEW' if raw=='review' else ('NO TM YET' if raw=='no_tm_yet' else 'UNRESOLVED'); status_class='no_tm_yet' if status=='NO TM YET' else status.lower()
        candidate='—'; accept=''
        if r['tm_candidate_url']:
            score=f'{float(r["tm_match_score"]):.3f}' if r['tm_match_score'] is not None else '—'
            candidate=f'<a href="{esc(r["tm_candidate_url"])}">{esc(r["tm_candidate_name"] or r["tm_candidate_url"])}</a><div class="muted smalltext">score {score}</div>'
            accept=f'<form method="post" action="/accept_tm_candidate"><input type="hidden" name="club_id" value="{r["id"]}"><input type="hidden" name="country_id" value="{cid}"><input type="hidden" name="level" value="{level}"><button class="acceptbtn">Accept candidate</button></form>'
        if status=='NO TM YET':
            resolver=f'<form method="post" action="/recheck_no_tm_club"><input type="hidden" name="club_id" value="{r["id"]}"><input type="hidden" name="country_id" value="{cid}"><input type="hidden" name="level" value="{level}"><button class="secondary">Recheck TM now</button></form>'
        else:
            resolver=(f'<form method="post" action="/retry_tm_clean_club"><input type="hidden" name="club_id" value="{r["id"]}"><input type="hidden" name="country_id" value="{cid}"><input type="hidden" name="level" value="{level}"><button class="secondary">Retry cleaned search</button></form>'+accept+
                      f'<form method="post" action="/mark_no_tm_yet"><input type="hidden" name="club_id" value="{r["id"]}"><input type="hidden" name="country_id" value="{cid}"><input type="hidden" name="level" value="{level}"><button class="notmyetbtn">Mark NO TM YET</button></form>')
        trs.append(f'<tr><td><b>{esc(r["league_name"])}</b></td><td><strong>{esc(source)}</strong><div class="muted smalltext">{esc(r["source_club_url"] or "")}</div></td><td>{esc(clean_name)}</td><td><span class="matchbadge {status_class}">{status}</span></td><td>{candidate}</td><td>{resolver}</td><td><form method="post" action="/set_tm_url_workbench"><input type="hidden" name="club_id" value="{r["id"]}"><input type="hidden" name="country_id" value="{cid}"><input type="hidden" name="level" value="{level}"><input class="tmurlinput" name="tm_url" placeholder="Paste Transfermarkt club URL" value="{esc(r["tm_candidate_url"] or "")}"><button>Save mapping</button></form></td></tr>')
    body=(notice+f'<div class="crumb"><a href="/level?country_id={cid}&level={level}">{esc(title)}</a> / Mapping Workbench</div>'
      +f'<div class="workbenchhero"><div><small>MAPPING WORKBENCH · LEVEL {level}</small><h1>{esc(title)} — mapping control</h1><p>Klubovi bez TM URL-a. NO TM YET = ručno provereno, profil trenutno nije pronađen, ali ostaje otvoren za kasnije.</p></div><div class="workbenchstats"><span><b>{len(rows)}</b> without TM URL</span><span><b>{review}</b> review</span><span><b>{unresolved}</b> unresolved</span><span><b>{no_tm_yet}</b> no TM yet</span><span><b>{verified}/{total_active}</b> verified</span></div></div>'
      +f'<section class="workbenchtools"><div><h2>Batch helpers</h2><p class="muted">Normalni retry preskače NO TM YET klubove. Njih periodično proverava posebno dugme.</p></div><div class="workbenchtoolgrid"><form method="post" action="/retry_tm_clean_level"><input type="hidden" name="country_id" value="{cid}"><input type="hidden" name="level" value="{level}"><button>Retry ALL unresolved/review</button></form><form method="post" action="/recheck_no_tm_level"><input type="hidden" name="country_id" value="{cid}"><input type="hidden" name="level" value="{level}"><button class="secondary">Recheck NO TM YET</button></form><a class="button secondarylink" href="/export_level_mapping_issues?country_id={cid}&level={level}">Download Open Mapping XLSX</a><form method="post" action="/import_mapping_xlsx" enctype="multipart/form-data"><input type="hidden" name="country_id" value="{cid}"><input type="hidden" name="level" value="{level}"><label>Import filled Mapping/Issues XLSX<input type="file" name="mapping_file" accept=".xlsx" required></label><button class="secondary">Import XLSX mappings</button></form></div><p class="muted"><b>NO TM YET:</b> možeš ga bilo kada recheck-ovati ili nalepiti novi TM URL; tada automatski postaje MATCHED.</p></section>'
      +f'<div class="workbenchtable"><table><thead><tr><th>League</th><th>SrbijaSport club</th><th>Clean search name</th><th>Status</th><th>Candidate</th><th>Resolver actions</th><th>Manual TM URL</th></tr></thead><tbody>{"".join(trs) if trs else "<tr><td colspan=\"7\">All active clubs are verified/mapped.</td></tr>"}</tbody></table></div>')
    return layout('Mapping Workbench',body)

def set_manual_tm_url(club_id,url):
    url=clean_text(url)
    m=re.search(r'https?://(?:www\.)?transfermarkt\.[^/]+/([^/?#]+)/(?:startseite|kader|mitarbeiter|datenfakten|transfers|leistungsdaten)/verein/(\d+)',url,re.I)
    if not m:
        # More permissive fallback for any TM URL containing /verein/<id>
        m2=re.search(r'https?://(?:www\.)?transfermarkt\.[^/]+/([^/?#]+)/[^?#]*?/verein/(\d+)',url,re.I)
        if not m2:
            raise RuntimeError('Could not read Transfermarkt slug / club ID from this URL.')
        slug,tmid=m2.group(1),m2.group(2)
    else:
        slug,tmid=m.group(1),m.group(2)

    canonical=f'https://www.transfermarkt.com/{slug}/startseite/verein/{tmid}'
    staff=f'https://www.transfermarkt.com/{slug}/mitarbeiter/verein/{tmid}'
    tm_name=''
    try:
        html=fetch_html(canonical)
        soup=BeautifulSoup(html,'lxml')
        h1=soup.find('h1')
        tm_name=clean_text(h1.get_text(' ',strip=True)) if h1 else ''
    except Exception:
        pass

    con=db()
    club=con.execute('SELECT * FROM clubs WHERE id=?',(club_id,)).fetchone()
    if not club:
        con.close()
        raise RuntimeError('Club not found.')
    con.execute("""UPDATE clubs SET tm_name=?,tm_club_id=?,tm_slug=?,tm_url=?,staff_url=?,
                   tm_match_status='matched',tm_match_score=1.0,tm_candidate_name=NULL,
                   tm_candidate_url=NULL,tm_match_note='Manual Transfermarkt mapping'
                   WHERE id=?""",
                (tm_name or club['tm_name'] or club['name'],tmid,slug,canonical,staff,club_id))
    con.commit(); con.close()
    return canonical


def normalize_section(text):
    t=clean_text(text).lower()
    for k,v in SECTION_ALIASES.items():
        if k in t: return v
    return ''

def table_section(table):
    node=table
    for _ in range(100):
        node=node.find_previous()
        if node is None: break
        if isinstance(node,Tag):
            cls=' '.join(node.get('class',[])); name=node.name.lower() if node.name else ''
            if name in ('h1','h2','h3','h4') or 'headline' in cls.lower():
                sec=normalize_section(node.get_text(' ',strip=True))
                if sec: return sec
    return ''

def headers_for(table):
    thead=table.find('thead'); return [clean_text(x.get_text(' ',strip=True)) for x in thead.find_all('th')] if thead else []

def flags_from_td(td):
    vals=[]
    for img in td.find_all('img'):
        x=clean_text(img.get('title') or img.get('alt'))
        if x and x not in vals: vals.append(x)
    return ' / '.join(vals)


def profile_from_row(row,base_url):
    for a in row.find_all('a',href=True):
        href=a.get('href',''); txt=clean_text(a.get_text(' ',strip=True))
        if txt and ('/profil/trainer/' in href or '/trainerprofil/' in href or '/profil/mitarbeiter/' in href): return txt,urljoin(base_url,href)
    for a in row.find_all('a',href=True):
        href=a.get('href',''); txt=clean_text(a.get_text(' ',strip=True))
        if txt and '/profil/' in href and '/verein/' not in href: return txt,urljoin(base_url,href)
    return '',''

def derive_position(name,values,hmap):
    for key in ('Function','Position','Role'):
        if key in hmap:
            v=clean_text(hmap[key])
            if name and v.lower().startswith(name.lower()): return clean_text(v[len(name):])
            if v and v!=name: return v
    for key,v in hmap.items():
        if 'name' in key.lower() and 'position' in key.lower():
            v=clean_text(v); return clean_text(v[len(name):]) if name and v.lower().startswith(name.lower()) else v
    for v in values:
        v=clean_text(v)
        if name and v.lower().startswith(name.lower()) and v.lower()!=name.lower():
            p=clean_text(v[len(name):])
            if p: return p
    return ''


def parse_staff_html(html,base_url,club_name):
    soup=BeautifulSoup(html,'lxml'); records=[]
    for table in soup.find_all('table'):
        section=table_section(table)
        if not section: continue
        headers=headers_for(table); rows=table.find('tbody').find_all('tr') if table.find('tbody') else table.find_all('tr')
        for row in rows:
            tds=row.find_all('td',recursive=False)
            if len(tds)<3: continue
            name,profile=profile_from_row(row,base_url)
            if not name: continue
            vals=[]
            for td in tds:
                txt=clean_text(td.get_text(' ',strip=True)); fl=flags_from_td(td); vals.append(fl if fl and not txt else txt)
            hmap={h:v for h,v in zip(headers,vals) if h} if len(headers)==len(vals) else {}
            nationality=dob=age=appointed=expires=last=''
            for k,v in hmap.items():
                kl=k.lower(); vv=clean_text(v)
                if 'date of birth' in kl or 'birthday' in kl or 'born' in kl: dob=extract_date(vv) or dob; age=extract_age(vv) or age
                elif kl.strip()=='age': age=extract_age(vv) or vv
                elif kl in ('nat.','nat','nationality'): nationality=vv
                elif 'appointed' in kl: appointed=vv
                elif 'contract expires' in kl or 'contract until' in kl: expires=vv
                elif 'last club' in kl: last=vv
            if not nationality:
                flags=[]
                for td in tds:
                    f=flags_from_td(td)
                    if f and f not in flags: flags.append(f)
                nationality=' / '.join(flags)
            records.append(dict(section=section,name=name,club=club_name,nationality=nationality,date_of_birth=dob,age=age,appointed=appointed,contract_expires=expires,last_club=last,profile_url=profile,dob_source='',position=derive_position(name,vals,hmap)))
    dedup={}
    for r in records:
        key=(r['section'],r['profile_url'] or r['name'].lower()); score=sum(bool(v) for v in r.values())
        if key not in dedup or score>dedup[key][0]: dedup[key]=(score,r)
    return [v[1] for v in dedup.values()]


def extract_dob_profile(html):
    soup=BeautifulSoup(html,'lxml')
    for script in soup.find_all('script',type='application/ld+json'):
        txt=script.string or script.get_text(' ',strip=True)
        if not txt: continue
        try:
            obj=json.loads(txt); items=obj if isinstance(obj,list) else [obj]
            for item in items:
                if isinstance(item,dict):
                    dob=item.get('birthDate') or item.get('birthdate')
                    if dob: return normalize_date(str(dob))
        except Exception: pass
    page=clean_text(soup.get_text(' ',strip=True))
    for label in ('Date of birth/Age','Date of birth','Born','Geb./Alter','Geburtsdatum'):
        for m in re.finditer(re.escape(label)+r'\s*:?\s*([^|]{0,120})',page,re.I):
            d=extract_date(m.group(1))
            if d: return d
    return ''

def enrich_dob(records):
    session=requests.Session(); cache={}
    for r in records:
        if r.get('date_of_birth') or not r.get('profile_url'): continue
        u=r['profile_url']
        if u in cache:
            r['date_of_birth']=cache[u]; r['dob_source']='cache' if cache[u] else 'cache:not-found'; continue
        try:
            dob=extract_dob_profile(fetch_html(u,session)); cache[u]=dob; r['date_of_birth']=dob; r['dob_source']='online' if dob else 'online:not-found'; time.sleep(1.2)
        except Exception: r['dob_source']='online:error'

def save_staff(club_id,records,source):
    con=db(); con.execute('DELETE FROM staff WHERE club_id=?',(club_id,))
    for r in records:
        con.execute("""INSERT OR REPLACE INTO staff(club_id,section,name,nationality,date_of_birth,age,appointed,contract_expires,last_club,profile_url,dob_source,position,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,CURRENT_TIMESTAMP)""",(club_id,r.get('section'),r.get('name'),r.get('nationality'),r.get('date_of_birth'),r.get('age'),r.get('appointed'),r.get('contract_expires'),r.get('last_club'),r.get('profile_url'),r.get('dob_source'),r.get('position')))
    con.execute('UPDATE clubs SET last_staff_update=CURRENT_TIMESTAMP WHERE id=?',(club_id,)); con.execute('INSERT INTO update_log(club_id,module,source,rows_imported,status,details) VALUES(?,?,?,?,?,?)',(club_id,'staff',source,len(records),'OK','')); con.commit(); con.close()


def game_safe_name(name):
    s=clean_text(name); repl={'Đ':'Dj','đ':'dj','Ł':'L','ł':'l','Ø':'O','ø':'o','Æ':'Ae','æ':'ae','Œ':'Oe','œ':'oe','ß':'ss'}
    for a,b in repl.items(): s=s.replace(a,b)
    s=unicodedata.normalize('NFKD',s); s=''.join(ch for ch in s if not unicodedata.combining(ch)); s=re.sub(r'[^\x20-\x7E]+','',s)
    return clean_text(s)


def player_from_row(row,base_url,headers):
    tds=row.find_all('td',recursive=False)
    if len(tds)<4: return None
    name=url=''; player_td=None
    for td in tds:
        for a in td.find_all('a',href=True):
            href=a.get('href','')
            if '/profil/spieler/' in href:
                nm=clean_text(a.get_text(' ',strip=True))
                if nm: name=nm; url=urljoin(base_url,href); player_td=td; break
        if name: break
    if not name: return None
    vals=[]
    for td in tds:
        txt=clean_text(td.get_text(' ',strip=True)); fl=flags_from_td(td); vals.append(fl if fl and not txt else txt)
    hmap={h:v for h,v in zip(headers,vals) if h} if len(headers)==len(vals) else {}
    squad=dob=age=nationality=market=''
    for k,v in hmap.items():
        kl=k.lower().strip(); vv=clean_text(v)
        if kl in ('#','no.','no','number','squad no','squad number'): squad=vv
        elif 'date of birth' in kl or 'age' in kl: dob=extract_date(vv) or dob; age=extract_age(vv) or age
        elif kl in ('nat.','nat','nationality'): nationality=vv
        elif 'market value' in kl: market=vv
    if not squad and vals and re.fullmatch(r'\d{1,3}|-',clean_text(vals[0])): squad='' if clean_text(vals[0])=='-' else clean_text(vals[0])
    if not dob and len(vals)>=3: dob=extract_date(vals[2]); age=extract_age(vals[2]) or age
    # Only use the legacy positional nationality fallback when the table has no headers.
    # National-team squad tables use that position for the current club.
    if not nationality and not headers and len(tds)>=4: nationality=flags_from_td(tds[3])
    if not market and len(vals)>=5: market=clean_text(vals[-1])
    position=''
    if player_td is not None:
        combined=clean_text(player_td.get_text(' ',strip=True))
        if combined.lower().startswith(name.lower()): position=clean_text(combined[len(name):])
        if not position:
            inline=player_td.find('table',class_=lambda x:x and 'inline-table' in x)
            if inline:
                cells=[clean_text(x.get_text(' ',strip=True)) for x in inline.find_all('td')]; cand=[c for c in cells if c and c.lower()!=name.lower()]
                if cand: position=cand[-1]
    # National-team squad tables also contain the player's current club.
    # Transfermarkt often renders that cell as a linked crest with no visible text,
    # so read both the header value and the linked club image/title.
    current_club=''
    for k,v in hmap.items():
        if 'club' in k.lower():
            current_club=clean_text(v)
            if current_club:
                break
    if not current_club:
        for td in tds:
            if td is player_td:
                continue
            for a in td.find_all('a',href=True):
                href=a.get('href','')
                if not re.search(r'/verein/\d+',href):
                    continue
                current_club=clean_text(a.get('title') or '')
                if not current_club:
                    img=a.find('img')
                    if img:
                        current_club=clean_text(img.get('title') or img.get('alt') or '')
                if not current_club:
                    current_club=clean_text(a.get_text(' ',strip=True))
                if current_club:
                    break
            if current_club:
                break
    safe=game_safe_name(name)
    tm_player_id=''
    m=re.search(r'/spieler/(\d+)',url or '')
    if m: tm_player_id=m.group(1)
    return {'squad_number':squad,'name':name,'tm_player_id':tm_player_id,'game_safe_name':safe,'safe_name_length':len(safe),'position':position,'date_of_birth':dob,'age':age,'nationality':nationality,'market_value':market,'profile_url':url,'current_club':current_club}


def parse_players_html(html,base_url):
    soup=BeautifulSoup(html,'lxml'); best=[]
    for table in soup.find_all('table'):
        if '/profil/spieler/' not in str(table): continue
        if 'inline-table' in (table.get('class') or []): continue
        headers=headers_for(table); ht=' | '.join(headers).lower()
        if headers and not ('player' in ht or 'date of birth' in ht or 'market value' in ht): continue
        tbody=table.find('tbody'); rows=tbody.find_all('tr',recursive=False) if tbody else table.find_all('tr'); parsed=[]
        for row in rows:
            r=player_from_row(row,base_url,headers)
            if r: parsed.append(r)
        if len(parsed)>len(best): best=parsed
    dedup={}
    for r in best:
        k=r['profile_url'] or r['name'].lower(); score=sum(bool(v) for v in r.values())
        if k not in dedup or score>dedup[k][0]: dedup[k]=(score,r)
    return [v[1] for v in dedup.values()]


def save_players(club_id,records,source):
    con=db(); con.execute('DELETE FROM players WHERE club_id=?',(club_id,))
    for r in records:
        con.execute("""INSERT OR REPLACE INTO players(club_id,squad_number,name,tm_player_id,game_safe_name,safe_name_length,position,date_of_birth,age,nationality,market_value,profile_url,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,CURRENT_TIMESTAMP)""",(club_id,r.get('squad_number'),r.get('name'),r.get('tm_player_id'),r.get('game_safe_name'),r.get('safe_name_length'),r.get('position'),r.get('date_of_birth'),r.get('age'),r.get('nationality'),r.get('market_value'),r.get('profile_url')))
    con.execute('UPDATE clubs SET last_players_update=CURRENT_TIMESTAMP WHERE id=?',(club_id,)); con.execute('INSERT INTO update_log(club_id,module,source,rows_imported,status,details) VALUES(?,?,?,?,?,?)',(club_id,'players',source,len(records),'OK','')); con.commit(); con.close()


def style_sheet(ws,widths):
    fill=PatternFill('solid',fgColor='17365D'); font=Font(color='FFFFFF',bold=True)
    for c in ws[1]: c.fill=fill; c.font=font; c.alignment=Alignment(horizontal='center')
    ws.freeze_panes='A2'
    for i,w in enumerate(widths,1): ws.column_dimensions[get_column_letter(i)].width=w


def staff_workbook_bytes(rows):
    wb=Workbook(); ws=wb.active; ws.title='Staff'; ws.append(['Section','Name','Club','Nationality','Date of birth','Age','Appointed','Contract expires','Last club','Profile_URL','DOB_Source','Position'])
    for r in rows: ws.append([r['section'],r['name'],r['club_name'],r['nationality'],r['date_of_birth'],r['age'],r['appointed'],r['contract_expires'],r['last_club'],r['profile_url'],r['dob_source'],r['position']])
    style_sheet(ws,[20,24,28,26,16,8,14,16,28,55,18,28]); bio=io.BytesIO(); wb.save(bio); return bio.getvalue()


def players_workbook_bytes(rows):
    wb=Workbook(); ws=wb.active; ws.title='Players'; ws.append(['Club','Squad No','Player Name','TM Player ID','Game Safe Name','Safe Name Length','Position','Date of Birth','Age','Nationality','Market Value','Player URL'])
    for r in rows: ws.append([r['club_name'],r['squad_number'],r['name'],r['tm_player_id'],r['game_safe_name'],r['safe_name_length'],r['position'],r['date_of_birth'],r['age'],r['nationality'],r['market_value'],r['profile_url']])
    style_sheet(ws,[28,10,24,13,24,16,22,16,8,28,14,58]); bio=io.BytesIO(); wb.save(bio); return bio.getvalue()



def current_season_id(now=None):
    now=now or datetime.now()
    return now.year if now.month>=7 else now.year-1


def season_label_from_id(season_id):
    y=int(season_id)
    return f'{str(y)[-2:]}/{str(y+1)[-2:]}'


def transfer_detailed_url(club,season_id=None):
    season_id=int(season_id or current_season_id())
    tm_id=clean_text(club['tm_club_id'] or '')
    slug=clean_text(club['tm_slug'] or '')
    if (not slug or not tm_id) and club['tm_url']:
        m=re.search(r'transfermarkt\.[^/]+/([^/]+)/[^?]*/verein/(\d+)',club['tm_url'])
        if m:
            slug=slug or m.group(1)
            tm_id=tm_id or m.group(2)
    if not tm_id:
        raise RuntimeError('Club has no Transfermarkt ID.')
    if not slug:
        raise RuntimeError('Club has no Transfermarkt slug.')
    return (f'https://www.transfermarkt.com/{slug}/transfers/verein/{tm_id}'
            f'/saison_id/{season_id}/pos//detailpos/0/w_s//skip_loans//skip_youth//plus/1')


def _tm_id_from_url(url,kind='spieler'):
    url=clean_text(url)
    if kind=='spieler':
        m=re.search(r'/spieler/(\d+)',url)
    elif kind=='verein':
        m=re.search(r'/verein/(\d+)',url)
    else:
        m=None
    return m.group(1) if m else ''


def _transfer_section_table(soup,title):
    wanted=title.lower()
    for h in soup.find_all('h2',class_='content-box-headline'):
        if clean_text(h.get_text(' ',strip=True)).lower()!=wanted:
            continue
        box=h.find_parent(class_='box')
        if not box:
            continue
        table=box.find('table',class_='items')
        if table:
            return table
    return None


def _transfer_position(player_td):
    inline=player_td.find('table',class_='inline-table')
    if inline:
        rows=inline.find_all('tr')
        if len(rows)>1:
            return clean_text(rows[1].get_text(' ',strip=True))
    textv=clean_text(player_td.get_text(' ',strip=True))
    a=player_td.find('a',href=re.compile(r'/profil/spieler/'))
    name=clean_text(a.get_text(' ',strip=True)) if a else ''
    if name and textv.startswith(name):
        return clean_text(textv[len(name):])
    return ''


def _transfer_nationality(td):
    vals=[]
    for img in td.find_all('img'):
        title=clean_text(img.get('title') or img.get('alt') or '')
        if title and title not in vals:
            vals.append(title)
    return ' / '.join(vals)


def _transfer_fee_and_date(td):
    date=''
    date_node=td.find('i',class_='normaler-text')
    if date_node:
        date=clean_text(date_node.get_text(' ',strip=True))
    a=td.find('a')
    if a:
        parts=[clean_text(s) for s in a.stripped_strings]
        parts=[x for x in parts if x and x!=date]
        fee=clean_text(' '.join(parts))
    else:
        fee=clean_text(td.get_text(' ',strip=True))
        if date and fee.endswith(date):
            fee=clean_text(fee[:-len(date)])
    detail=''
    if a and '/jumplist/transfers/' in (a.get('href') or ''):
        detail=urljoin('https://www.transfermarkt.com/',a.get('href'))
    return fee,date,detail


def parse_transfer_table(soup,title,direction,base_url):
    table=_transfer_section_table(soup,title)
    if not table:
        return []

    headers=[clean_text(th.get_text(' ',strip=True)) for th in table.find_all('th')]
    header_map={h.lower():i for i,h in enumerate(headers) if h}
    pidx=header_map.get('player',1)
    age_idx=header_map.get('age',2)
    mv_idx=header_map.get('market value')
    nat_idx=header_map.get('nat.')
    club_idx=header_map.get('left')
    if club_idx is None:
        club_idx=header_map.get('joined')
    fee_idx=header_map.get('fee',len(headers)-1)

    tbody=table.find('tbody')
    rows=tbody.find_all('tr',recursive=False) if tbody else []
    out=[]

    for row in rows:
        tds=row.find_all('td',recursive=False)
        if not tds or pidx>=len(tds):
            continue

        ptd=tds[pidx]
        pa=ptd.find('a',href=re.compile(r'/profil/spieler/'))
        if not pa:
            continue

        player_name=clean_text(pa.get_text(' ',strip=True))
        player_url=urljoin(base_url,pa.get('href') or '')
        position=_transfer_position(ptd)
        age=clean_text(tds[age_idx].get_text(' ',strip=True)) if age_idx is not None and age_idx<len(tds) else ''
        market_value=clean_text(tds[mv_idx].get_text(' ',strip=True)) if mv_idx is not None and mv_idx<len(tds) else ''
        nationality=_transfer_nationality(tds[nat_idx]) if nat_idx is not None and nat_idx<len(tds) else ''

        other_club=''
        other_club_url=''
        other_club_tm_id=''
        other_league=''
        other_country=''
        if club_idx is not None and club_idx<len(tds):
            ctd=tds[club_idx]
            club_a=ctd.find('a',href=re.compile(r'/startseite/verein/'))
            if club_a:
                other_club=clean_text(club_a.get('title') or club_a.get_text(' ',strip=True))
                other_club_url=urljoin(base_url,club_a.get('href') or '')
                other_club_tm_id=_tm_id_from_url(other_club_url,'verein')
            league_a=ctd.find('a',href=re.compile(r'/transfers/wettbewerb/'))
            if league_a:
                other_league=clean_text(league_a.get_text(' ',strip=True))
            flag=ctd.find('img',class_=re.compile(r'flaggenrahmen'))
            if flag:
                other_country=clean_text(flag.get('title') or flag.get('alt') or '')

        fee=''
        transfer_date=''
        detail_url=''
        if fee_idx is not None and fee_idx<len(tds):
            fee,transfer_date,detail_url=_transfer_fee_and_date(tds[fee_idx])

        out.append({
            'direction':direction,
            'player_name':player_name,
            'player_url':player_url,
            'player_tm_id':_tm_id_from_url(player_url,'spieler'),
            'position':position,
            'age':age,
            'nationality':nationality,
            'market_value':market_value,
            'other_club':other_club,
            'other_club_url':other_club_url,
            'other_club_tm_id':other_club_tm_id,
            'other_league':other_league,
            'other_country':other_country,
            'fee':fee,
            'transfer_date':transfer_date,
            'transfer_detail_url':detail_url,
        })
    return out


def parse_transfer_summary(soup):
    result={'arrivals_count':0,'departures_count':0,'expenditure':'','income':'','balance':''}
    for h in soup.find_all('h2',class_='content-box-headline'):
        if clean_text(h.get_text(' ',strip=True)).lower()!='transfer record':
            continue
        box=h.find_parent(class_='box')
        table=box.find('table') if box else None
        if not table:
            continue
        for tr in table.find_all('tr'):
            cells=[clean_text(x.get_text(' ',strip=True)) for x in tr.find_all(['th','td'])]
            if not cells:
                continue
            label=cells[0].lower()
            if label=='income' and len(cells)>=3:
                try: result['departures_count']=int(re.sub(r'\D','',cells[1]) or 0)
                except: pass
                result['income']=cells[2]
            elif label=='expenditure' and len(cells)>=3:
                try: result['arrivals_count']=int(re.sub(r'\D','',cells[1]) or 0)
                except: pass
                result['expenditure']=cells[2]
            elif label=='overall balance' and len(cells)>=2:
                result['balance']=cells[-1]
        break
    return result


def parse_transfers_html(html,base_url,season_id=None):
    soup=BeautifulSoup(html,'lxml')
    sid=int(season_id or current_season_id())
    label=season_label_from_id(sid)

    for h in soup.find_all('h2',class_='content-box-headline'):
        txt=clean_text(h.get_text(' ',strip=True))
        m=re.fullmatch(r'Transfers\s+(\d{2}/\d{2})',txt,re.I)
        if m:
            label=m.group(1)
            break

    arrivals=parse_transfer_table(soup,'Arrivals','arrival',base_url)
    departures=parse_transfer_table(soup,'Departures','departure',base_url)
    summary=parse_transfer_summary(soup)

    # If the Transfer record box is absent, derive counts from rows.
    if not summary['arrivals_count']:
        summary['arrivals_count']=len(arrivals)
    if not summary['departures_count']:
        summary['departures_count']=len(departures)

    summary['season_id']=sid
    summary['season_label']=label
    return summary,arrivals+departures


def save_club_transfers(club_id,summary,records,source='Transfermarkt'):
    season_id=int(summary['season_id'])
    season_label=summary.get('season_label') or season_label_from_id(season_id)
    con=db()

    # Current-season-only policy: remove old seasons for this club.
    con.execute('DELETE FROM club_transfers WHERE club_id=?',(club_id,))
    con.execute('DELETE FROM club_transfer_summary WHERE club_id=?',(club_id,))

    for r in records:
        con.execute("""INSERT INTO club_transfers(
            club_id,season_id,season_label,direction,player_name,player_url,player_tm_id,
            position,age,nationality,market_value,other_club,other_club_url,other_club_tm_id,
            other_league,other_country,fee,transfer_date,transfer_detail_url,updated_at)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,CURRENT_TIMESTAMP)""",(
            club_id,season_id,season_label,r.get('direction'),r.get('player_name'),
            r.get('player_url'),r.get('player_tm_id'),r.get('position'),r.get('age'),
            r.get('nationality'),r.get('market_value'),r.get('other_club'),
            r.get('other_club_url'),r.get('other_club_tm_id'),r.get('other_league'),
            r.get('other_country'),r.get('fee'),r.get('transfer_date'),
            r.get('transfer_detail_url')))

    con.execute("""INSERT OR REPLACE INTO club_transfer_summary(
        club_id,season_id,season_label,arrivals_count,departures_count,
        expenditure,income,balance,updated_at)
        VALUES(?,?,?,?,?,?,?,?,CURRENT_TIMESTAMP)""",(
        club_id,season_id,season_label,summary.get('arrivals_count',0),
        summary.get('departures_count',0),summary.get('expenditure',''),
        summary.get('income',''),summary.get('balance','')))

    con.execute('UPDATE clubs SET last_transfers_update=CURRENT_TIMESTAMP WHERE id=?',(club_id,))
    con.execute("""INSERT INTO update_log(club_id,module,source,rows_imported,status,details)
        VALUES(?,?,?,?,?,?)""",(club_id,'transfers',source,len(records),'OK',
        f"Season {season_label}; arrivals {summary.get('arrivals_count',0)}; departures {summary.get('departures_count',0)}"))
    con.commit()
    con.close()


def update_transfers_live(club_id):
    con=db()
    club=con.execute('SELECT * FROM clubs WHERE id=?',(club_id,)).fetchone()
    con.close()
    if not club:
        raise RuntimeError('Club not found.')
    if not club['tm_url']:
        raise RuntimeError('Club has no Transfermarkt mapping.')

    sid=current_season_id()
    url=transfer_detailed_url(club,sid)
    html=fetch_html(url)
    summary,records=parse_transfers_html(html,url,sid)

    if not records and not (summary.get('arrivals_count') or summary.get('departures_count')):
        # Valid empty season is allowed, but reject obviously blocked/non-transfer pages.
        low=html.lower()
        if 'arrivals' not in low and 'departures' not in low and 'transfer record' not in low:
            raise RuntimeError('Transfermarkt transfer page could not be parsed.')

    save_club_transfers(club_id,summary,records,'live Transfermarkt detailed view')
    return {'url':url,'summary':summary,'records':len(records)}


def batch_update_league_transfers(lid):
    con=db()
    active_total=con.execute("""SELECT COUNT(*) FROM clubs
        WHERE league_id=? AND COALESCE(source_active,1)=1""",(lid,)).fetchone()[0]
    clubs=con.execute("""SELECT * FROM clubs
        WHERE league_id=? AND COALESCE(source_active,1)=1
        AND COALESCE(tm_url,'')<>'' ORDER BY name""",(lid,)).fetchall()
    con.close()
    out={'active':active_total,'mapped':len(clubs),'skipped':max(0,active_total-len(clubs)),
         'ok':0,'with_data':0,'empty':0,'rows':0,'arrivals':0,'departures':0,'errors':[]}
    session=requests.Session()
    for i,c in enumerate(clubs,1):
        cname=c['cm_name'] or c['name']
        print(f'[LIGA TRANSFERI {i}/{len(clubs)}] {cname}')
        try:
            sid=current_season_id()
            url=transfer_detailed_url(c,sid)
            html=fetch_html(url,session)
            summary,records=parse_transfers_html(html,url,sid)
            low=html.lower()
            if not records and not (summary.get('arrivals_count') or summary.get('departures_count')):
                if 'arrivals' not in low and 'departures' not in low and 'transfer record' not in low:
                    raise RuntimeError('Stranica transfera nije mogla da se obradi.')
            save_club_transfers(c['id'],summary,records,'batch league transfers')
            out['ok']+=1
            out['rows']+=len(records)
            out['arrivals']+=int(summary.get('arrivals_count') or 0)
            out['departures']+=int(summary.get('departures_count') or 0)
            if records or summary.get('arrivals_count') or summary.get('departures_count'):
                out['with_data']+=1
            else:
                out['empty']+=1
        except Exception as e:
            out['errors'].append(f'{cname}: {e}')
        if i<len(clubs): time.sleep(0.8)
    return out


def batch_update_level_transfers(cid,level):
    con=db()
    active_total=con.execute("""SELECT COUNT(*) FROM clubs cl JOIN leagues l ON l.id=cl.league_id
        WHERE l.country_id=? AND l.level=? AND COALESCE(cl.source_active,1)=1""",(cid,level)).fetchone()[0]
    clubs=con.execute("""SELECT cl.*,l.name league_name FROM clubs cl JOIN leagues l ON l.id=cl.league_id
        WHERE l.country_id=? AND l.level=? AND COALESCE(cl.source_active,1)=1
        AND COALESCE(cl.tm_url,'')<>'' ORDER BY l.name,cl.name""",(cid,level)).fetchall()
    con.close()
    out={'active':active_total,'mapped':len(clubs),'skipped':max(0,active_total-len(clubs)),
         'ok':0,'with_data':0,'empty':0,'rows':0,'arrivals':0,'departures':0,'errors':[]}
    session=requests.Session()
    for i,c in enumerate(clubs,1):
        cname=c['cm_name'] or c['name']
        print(f'[LEVEL {level} TRANSFERI {i}/{len(clubs)}] {c["league_name"]} / {cname}')
        try:
            sid=current_season_id()
            url=transfer_detailed_url(c,sid)
            html=fetch_html(url,session)
            summary,records=parse_transfers_html(html,url,sid)
            low=html.lower()
            if not records and not (summary.get('arrivals_count') or summary.get('departures_count')):
                if 'arrivals' not in low and 'departures' not in low and 'transfer record' not in low:
                    raise RuntimeError('Stranica transfera nije mogla da se obradi.')
            save_club_transfers(c['id'],summary,records,f'batch level {level} transfers')
            out['ok']+=1
            out['rows']+=len(records)
            out['arrivals']+=int(summary.get('arrivals_count') or 0)
            out['departures']+=int(summary.get('departures_count') or 0)
            if records or summary.get('arrivals_count') or summary.get('departures_count'):
                out['with_data']+=1
            else:
                out['empty']+=1
        except Exception as e:
            out['errors'].append(f'{c["league_name"]} / {cname}: {e}')
        if i<len(clubs): time.sleep(0.8)
    return out


def batch_update_country_transfers(cid):
    con=db()
    active_total=con.execute("""SELECT COUNT(*) FROM clubs cl JOIN leagues l ON l.id=cl.league_id
        WHERE l.country_id=? AND COALESCE(l.section,'senior')='senior' AND COALESCE(cl.source_active,1)=1""",(cid,)).fetchone()[0]
    clubs=con.execute("""SELECT cl.*,l.name league_name,l.level league_level FROM clubs cl
        JOIN leagues l ON l.id=cl.league_id
        WHERE l.country_id=? AND COALESCE(l.section,'senior')='senior' AND COALESCE(cl.source_active,1)=1
        AND COALESCE(cl.tm_url,'')<>'' ORDER BY l.level,l.name,cl.name""",(cid,)).fetchall()
    con.close()
    out={'active':active_total,'mapped':len(clubs),'skipped':max(0,active_total-len(clubs)),
         'ok':0,'with_data':0,'empty':0,'rows':0,'arrivals':0,'departures':0,'errors':[]}
    session=requests.Session()
    for i,c in enumerate(clubs,1):
        cname=c['cm_name'] or c['name']
        print(f'[SRBIJA TRANSFERI {i}/{len(clubs)}] {c["league_name"]} / {cname}')
        try:
            sid=current_season_id()
            url=transfer_detailed_url(c,sid)
            html=fetch_html(url,session)
            summary,records=parse_transfers_html(html,url,sid)
            low=html.lower()
            if not records and not (summary.get('arrivals_count') or summary.get('departures_count')):
                if 'arrivals' not in low and 'departures' not in low and 'transfer record' not in low:
                    raise RuntimeError('Stranica transfera nije mogla da se obradi.')
            save_club_transfers(c['id'],summary,records,'batch Serbia transfers')
            out['ok']+=1
            out['rows']+=len(records)
            out['arrivals']+=int(summary.get('arrivals_count') or 0)
            out['departures']+=int(summary.get('departures_count') or 0)
            if records or summary.get('arrivals_count') or summary.get('departures_count'):
                out['with_data']+=1
            else:
                out['empty']+=1
        except Exception as e:
            out['errors'].append(f'{c["league_name"]} / {cname}: {e}')
        if i<len(clubs): time.sleep(0.8)
    return out


def transfers_workbook_bytes(club,summary,rows):
    locations=current_player_locations()
    wb=Workbook()
    ws=wb.active
    ws.title='Pregled'
    ws.append(['Klub','Sezona','Dolasci','Odlasci','Potrošeno','Prihod','Bilans','Poslednje ažuriranje'])
    ws.append([
        club['cm_name'] or club['name'],summary['season_label'] if summary else season_label_from_id(current_season_id()),
        summary['arrivals_count'] if summary else 0,summary['departures_count'] if summary else 0,
        summary['expenditure'] if summary else '',summary['income'] if summary else '',
        summary['balance'] if summary else '',summary['updated_at'] if summary else ''
    ])
    style_sheet(ws,[28,12,10,10,14,14,14,20])

    wt=wb.create_sheet('Transferi')
    headers=['Smer','Igrač','TM Player ID','Trenutni klub u Srbiji','Trenutna liga u Srbiji',
             'Pozicija','Godine','Državljanstvo','Tržišna vrednost','Drugi klub','TM Club ID',
             'Liga','Država','Naknada','Datum','Player URL','Club URL','Transfer URL']
    wt.append(headers)
    for r in rows:
        loc=locations.get(clean_text(r['player_tm_id'] or ''),{})
        wt.append([
            'DOLAZAK' if r['direction']=='arrival' else 'ODLAZAK',
            r['player_name'],r['player_tm_id'],loc.get('club',''),loc.get('league',''),
            r['position'],r['age'],r['nationality'],r['market_value'],r['other_club'],
            r['other_club_tm_id'],r['other_league'],r['other_country'],r['fee'],
            r['transfer_date'],r['player_url'],r['other_club_url'],r['transfer_detail_url']
        ])
    style_sheet(wt,[12,24,13,30,28,22,8,24,16,28,12,24,18,16,14,52,52,58])
    bio=io.BytesIO()
    wb.save(bio)
    return bio.getvalue()


def _info_table_values(soup):
    out={}
    table=soup.select_one('div.info-table')
    if not table:
        return out
    children=[x for x in table.children if isinstance(x,Tag)]
    i=0
    while i<len(children):
        node=children[i]
        cls=node.get('class',[])
        if 'info-table__content--regular' in cls:
            label=clean_text(node.get_text(' ',strip=True)).rstrip(':')
            vals=[]
            j=i+1
            while j<len(children):
                nxt=children[j]
                if 'info-table__content--regular' in nxt.get('class',[]):
                    break
                val=clean_text(nxt.get_text(' ',strip=True))
                if val:
                    vals.append(val)
                j+=1
            out[label]=vals
            i=j
        else:
            i+=1
    return out


def parse_club_info_html(html,tm_club_id=''):
    soup=BeautifulSoup(html,'lxml')
    header={}
    for li in soup.select('li.data-header__label'):
        txt=clean_text(li.get_text(' ',strip=True))
        if ':' in txt:
            k,v=txt.split(':',1)
            header[clean_text(k)]=clean_text(v)

    facts=_info_table_values(soup)
    official=(facts.get('Official club name') or [''])[0]
    addr=facts.get('Address') or []
    address_line1=addr[0] if len(addr)>0 else ''
    postal_city=addr[1] if len(addr)>1 else ''
    country=addr[2] if len(addr)>2 else ''
    city=re.sub(r'^\s*\d[\d\s-]*\s*','',postal_city).strip()
    phone=(facts.get('Tel') or [''])[0]
    fax=(facts.get('Fax') or [''])[0]
    website=(facts.get('Website') or [''])[0]
    founded=(facts.get('Osnovano') or [''])[0]

    foreigners_raw=header.get('Foreigners','')
    fm=re.match(r'\s*(\d+)\s*(.*)',foreigners_raw)
    foreigners=fm.group(1) if fm else foreigners_raw
    foreigners_pct=''
    pm=re.search(r'([\d.,]+)\s*%',foreigners_raw)
    if pm:
        foreigners_pct=pm.group(1)+' %'

    stadium_raw=header.get('Stadium','')
    cap=None
    cm=re.search(r'([\d.,]+)\s+Seats?',stadium_raw,re.I)
    stadium_name=stadium_raw
    if cm:
        cap_digits=re.sub(r'\D','',cm.group(1))
        cap=int(cap_digits) if cap_digits else None
        stadium_name=clean_text(stadium_raw[:cm.start()])

    crest_url=''
    preferred=[]
    fallback=[]

    # Transfermarkt may lazy-load club crests.
    for img in soup.find_all('img'):
        for raw_src in (
            img.get('src') or '',
            img.get('data-src') or '',
            img.get('data-lazy-src') or ''
        ):
            src=clean_text(raw_src)
            if not src:
                continue
            if src.startswith('//'):
                src='https:'+src
            elif src.startswith('/'):
                src='https://tmssl.akamaized.net'+src
            if '/images/wappen/' not in src:
                continue

            if tm_club_id and re.search(
                rf'/wappen/(?:head|medium|small)/{re.escape(str(tm_club_id))}\.',
                src
            ):
                preferred.append(src)
            else:
                fallback.append(src)

    # OpenGraph fallback.
    if not preferred:
        og=soup.find('meta',attrs={'property':'og:image'})
        if og:
            src=clean_text(og.get('content') or '')
            if src and '/images/wappen/' in src:
                preferred.append(src)

    if preferred:
        crest_url=preferred[0]
    elif fallback:
        crest_url=fallback[0]
    elif tm_club_id:
        crest_url=f'https://tmssl.akamaized.net/images/wappen/head/{tm_club_id}.png'

    info={
        'official_name':official,'founded':founded,'address_line1':address_line1,
        'postal_city':postal_city,'city':city,'country':country,'phone':phone,'fax':fax,
        'website':website,'squad_size':header.get('Broj igrača',''),
        'average_age':header.get('Prosečna starost',''),'foreigners':foreigners,
        'foreigners_pct':foreigners_pct,'national_team_players':header.get('National team players',''),
        'transfer_record':header.get('Current transfer record',''),'crest_url':crest_url,
    }
    stadium={'name':stadium_name,'capacity':cap,'city':city}
    return info,stadium


def cache_crest(club_id,crest_url):
    if not club_id or not crest_url:
        return ''
    candidates=[crest_url]
    if re.search(r'srbijasport\.net/img/club/\d+/500(?:$|\?)', crest_url):
        candidates += [re.sub(r'/500(?=$|\?)','/300',crest_url), re.sub(r'/500(?=$|\?)','/200',crest_url)]
    last_error=None
    for candidate_url in dict.fromkeys(candidates):
        try:
            r=requests.get(candidate_url,headers={**HEADERS,'Referer':'https://srbijasport.net/','Accept':'image/avif,image/webp,image/apng,image/svg+xml,image/*,*/*;q=0.8'},timeout=30)
            r.raise_for_status()
            if not r.content:
                continue
            ctype=(r.headers.get('Content-Type') or '').split(';',1)[0].strip().lower()
            if ctype and not (ctype.startswith('image/') or ctype=='application/octet-stream'):
                continue
            ext_map={'image/png':'.png','image/jpeg':'.jpg','image/jpg':'.jpg','image/webp':'.webp','image/svg+xml':'.svg','image/gif':'.gif'}
            ext=ext_map.get(ctype)
            if not ext:
                path=urlparse(candidate_url).path.lower()
                ext=next((x for x in ('.png','.jpg','.jpeg','.webp','.svg','.gif') if path.endswith(x)),'.jpg')
            dest=CREST_DIR / f'{club_id}{ext}'
            dest.write_bytes(r.content)
            return f'static/crests/{dest.name}'
        except Exception as e:
            last_error=e
    if last_error:
        print('Crest cache failed',club_id,last_error)
    return ''


def save_club_info(club_id,info,stadium,download_crest=True):
    con=db()
    crest_local=''
    old=con.execute('SELECT crest_local FROM club_info WHERE club_id=?',(club_id,)).fetchone()
    if old and old['crest_local']:
        crest_local=old['crest_local']
    if download_crest and info.get('crest_url'):
        crest_local=cache_crest(club_id,info.get('crest_url')) or crest_local
    con.execute("""INSERT OR REPLACE INTO club_info(
      club_id,official_name,founded,address_line1,postal_city,city,country,phone,fax,website,
      squad_size,average_age,foreigners,foreigners_pct,national_team_players,transfer_record,
      crest_url,crest_local,updated_at)
      VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,CURRENT_TIMESTAMP)""",(
      club_id,info.get('official_name'),info.get('founded'),info.get('address_line1'),
      info.get('postal_city'),info.get('city'),info.get('country'),info.get('phone'),
      info.get('fax'),info.get('website'),info.get('squad_size'),info.get('average_age'),
      info.get('foreigners'),info.get('foreigners_pct'),info.get('national_team_players'),
      info.get('transfer_record'),info.get('crest_url'),crest_local))
    # Do not use INSERT OR REPLACE here: REPLACE deletes/recreates the row and
    # would wipe manually/automatically collected stadium imagery. Keep media
    # fields intact while refreshing factual stadium metadata.
    con.execute("""INSERT INTO stadiums(club_id,name,capacity,city,updated_at)
      VALUES(?,?,?,?,CURRENT_TIMESTAMP)
      ON CONFLICT(club_id) DO UPDATE SET
        name=excluded.name,capacity=excluded.capacity,city=excluded.city,updated_at=CURRENT_TIMESTAMP""",
      (club_id,stadium.get('name'),stadium.get('capacity'),stadium.get('city')))
    con.execute("""INSERT INTO update_log(club_id,module,source,rows_imported,status,details)
      VALUES(?,?,?,?,?,?)""",(club_id,'club_info','Transfermarkt',1,'OK',''))
    con.commit()
    con.close()
    return crest_local


def batch_update_league(lid,fetch_staff_dob=False):
    con=db()
    league=con.execute('SELECT * FROM leagues WHERE id=?',(lid,)).fetchone()
    clubs=con.execute('SELECT * FROM clubs WHERE league_id=? AND COALESCE(source_active,1)=1 ORDER BY name',(lid,)).fetchall()
    con.close()
    summary={'clubs':len(clubs),'ok':0,'players':0,'staff':0,'skipped':0,'errors':[]}
    if not league:
        return summary

    ok_ids=set()
    is_women=(league['section'] or '').lower()=='women'
    # Official FSS reports are a roster source for the two national women's
    # divisions. They are merged first; richer TM records may enrich them later.
    if is_women and league['fss_url']:
        try:
            fr=update_women_league_fss_rosters(lid)
            summary['players']+=int(fr.get('players',0) or 0)
            ok_ids.update(fr.get('club_ids',[]))
            summary['errors'].extend(fr.get('errors',[]))
        except Exception as e:
            summary['errors'].append('FSS roster sync: '+str(e))

    session=requests.Session()
    for idx,c in enumerate(clubs,1):
        cname=c['cm_name'] or c['name']
        print(f'[BATCH {idx}/{len(clubs)}] {cname}')
        if not c['tm_url']:
            # Women's football can function from FSS/SrbijaSport without TM.
            # Other legacy sections keep the old explicit missing-TM warning.
            if is_women:
                summary['skipped']+=1
                continue
            summary['errors'].append(cname+': missing Transfermarkt club URL')
            continue
        club_ok=True
        try:
            start_html=fetch_html(c['tm_url'],session)
            info,stadium=parse_club_info_html(start_html,c['tm_club_id'])
            save_club_info(c['id'],info,stadium,True)
            players=parse_players_html(start_html,c['tm_url'])
            save_players(c['id'],players,'batch live Transfermarkt')
            summary['players']+=len(players)
        except Exception as e:
            club_ok=False
            summary['errors'].append(cname+' players/info: '+str(e))
        try:
            if c['staff_url']:
                staff_html=fetch_html(c['staff_url'],session)
                staff=parse_staff_html(staff_html,c['staff_url'],cname)
                if fetch_staff_dob:
                    enrich_dob(staff)
                save_staff(c['id'],staff,'batch live Transfermarkt')
                summary['staff']+=len(staff)
        except Exception as e:
            club_ok=False
            summary['errors'].append(cname+' staff: '+str(e))
        if club_ok:
            ok_ids.add(c['id'])
        time.sleep(0.8)
    summary['ok']=len(ok_ids)
    return summary


def _write_sheet(wb,title,headers,rows,widths=None):
    ws=wb.create_sheet(title)
    ws.append(headers)
    for row in rows:
        ws.append(row)
    style_sheet(ws,widths or [20]*len(headers))
    return ws




def current_player_locations():
    """TM Player ID -> current Serbian club/league based on latest roster tables."""
    con=db()
    rows=con.execute("""SELECT p.tm_player_id,COALESCE(cl.cm_name,cl.name) club_name,l.name league_name
        FROM players p JOIN clubs cl ON cl.id=p.club_id JOIN leagues l ON l.id=cl.league_id
        WHERE COALESCE(p.tm_player_id,'')<>'' AND COALESCE(cl.source_active,1)=1
        ORDER BY l.level,l.name,club_name""").fetchall()
    con.close()
    out={}
    for r in rows:
        key=clean_text(r['tm_player_id'])
        if not key: continue
        item=out.setdefault(key,{'clubs':[],'leagues':[]})
        if r['club_name'] not in item['clubs']: item['clubs'].append(r['club_name'])
        if r['league_name'] not in item['leagues']: item['leagues'].append(r['league_name'])
    return {k:{'club':' / '.join(v['clubs']),'league':' / '.join(v['leagues'])} for k,v in out.items()}


def league_transfer_stats_map(cid,level=None):
    sid=current_season_id()
    con=db()
    params=[sid,cid]
    level_sql=''
    if level is not None:
        level_sql=' AND l.level=?'
        params.append(int(level))
    rows=con.execute(f"""SELECT l.id league_id,
        COUNT(ct.id) transfer_rows,
        SUM(CASE WHEN ct.direction='arrival' THEN 1 ELSE 0 END) arrivals,
        SUM(CASE WHEN ct.direction='departure' THEN 1 ELSE 0 END) departures,
        COUNT(DISTINCT CASE WHEN COALESCE(ct.transfer_detail_url,'')<>'' THEN ct.transfer_detail_url END) unique_events
        FROM leagues l
        LEFT JOIN clubs cl ON cl.league_id=l.id AND COALESCE(cl.source_active,1)=1
        LEFT JOIN club_transfers ct ON ct.club_id=cl.id AND ct.season_id=?
        WHERE l.country_id=? AND COALESCE(l.section,'senior')='senior'{level_sql}
        GROUP BY l.id""",tuple(params)).fetchall()

    params2=[sid,cid]
    if level is not None: params2.append(int(level))
    updated=con.execute(f"""SELECT l.id league_id,COUNT(DISTINCT cts.club_id) updated_clubs
        FROM leagues l
        LEFT JOIN clubs cl ON cl.league_id=l.id AND COALESCE(cl.source_active,1)=1
        LEFT JOIN club_transfer_summary cts ON cts.club_id=cl.id AND cts.season_id=?
        WHERE l.country_id=? AND COALESCE(l.section,'senior')='senior'{level_sql}
        GROUP BY l.id""",tuple(params2)).fetchall()
    con.close()

    out={r['league_id']:{
        'rows':int(r['transfer_rows'] or 0),
        'arrivals':int(r['arrivals'] or 0),
        'departures':int(r['departures'] or 0),
        'unique':int(r['unique_events'] or 0),
        'clubs':0,
    } for r in rows}
    for r in updated:
        out.setdefault(r['league_id'],{'rows':0,'arrivals':0,'departures':0,'unique':0,'clubs':0})
        out[r['league_id']]['clubs']=int(r['updated_clubs'] or 0)
    return out


def club_transfer_stats_map(lid):
    sid=current_season_id()
    con=db()
    rows=con.execute("""SELECT cl.id club_id,
        COUNT(ct.id) transfer_rows,
        SUM(CASE WHEN ct.direction='arrival' THEN 1 ELSE 0 END) arrivals,
        SUM(CASE WHEN ct.direction='departure' THEN 1 ELSE 0 END) departures,
        MAX(cts.updated_at) updated_at
        FROM clubs cl
        LEFT JOIN club_transfers ct ON ct.club_id=cl.id AND ct.season_id=?
        LEFT JOIN club_transfer_summary cts ON cts.club_id=cl.id AND cts.season_id=?
        WHERE cl.league_id=? AND COALESCE(cl.source_active,1)=1
        GROUP BY cl.id""",(sid,sid,lid)).fetchall()
    con.close()
    return {r['club_id']:{
        'rows':int(r['transfer_rows'] or 0),
        'arrivals':int(r['arrivals'] or 0),
        'departures':int(r['departures'] or 0),
        'updated':r['updated_at'] or '',
    } for r in rows}


def country_dashboard_stats(cid):
    con=db()
    country=con.execute('SELECT * FROM countries WHERE id=?',(cid,)).fetchone()
    if not country:
        con.close()
        return None

    league_count=con.execute("SELECT COUNT(*) FROM leagues WHERE country_id=? AND COALESCE(section,'senior')='senior' AND COALESCE(source_active,1)=1",(cid,)).fetchone()[0]
    club_count=con.execute("""SELECT COUNT(*) FROM clubs cl JOIN leagues l ON l.id=cl.league_id
                              WHERE l.country_id=? AND COALESCE(l.section,'senior')='senior' AND COALESCE(cl.source_active,1)=1""",(cid,)).fetchone()[0]
    mapped_count=con.execute("""SELECT COUNT(*) FROM clubs cl JOIN leagues l ON l.id=cl.league_id
                                WHERE l.country_id=? AND COALESCE(l.section,'senior')='senior' AND COALESCE(cl.source_active,1)=1
                                AND COALESCE(cl.tm_url,'')<>''""",(cid,)).fetchone()[0]
    player_count=con.execute("""SELECT COUNT(*) FROM players p
        JOIN clubs cl ON cl.id=p.club_id JOIN leagues l ON l.id=cl.league_id
        WHERE l.country_id=? AND COALESCE(l.section,'senior')='senior' AND COALESCE(cl.source_active,1)=1""",(cid,)).fetchone()[0]
    staff_count=con.execute("""SELECT COUNT(*) FROM staff s
        JOIN clubs cl ON cl.id=s.club_id JOIN leagues l ON l.id=cl.league_id
        WHERE l.country_id=? AND COALESCE(l.section,'senior')='senior' AND COALESCE(cl.source_active,1)=1""",(cid,)).fetchone()[0]
    info_count=con.execute("""SELECT COUNT(*) FROM club_info ci
        JOIN clubs cl ON cl.id=ci.club_id JOIN leagues l ON l.id=cl.league_id
        WHERE l.country_id=? AND COALESCE(l.section,'senior')='senior' AND COALESCE(cl.source_active,1)=1""",(cid,)).fetchone()[0]
    stadium_count=con.execute("""SELECT COUNT(*) FROM stadiums s
        JOIN clubs cl ON cl.id=s.club_id JOIN leagues l ON l.id=cl.league_id
        WHERE l.country_id=? AND COALESCE(l.section,'senior')='senior' AND COALESCE(cl.source_active,1)=1
        AND COALESCE(s.name,'')<>''""",(cid,)).fetchone()[0]
    sid=current_season_id()
    tr=con.execute("""SELECT COUNT(*) rows,
        SUM(CASE WHEN ct.direction='arrival' THEN 1 ELSE 0 END) arrivals,
        SUM(CASE WHEN ct.direction='departure' THEN 1 ELSE 0 END) departures,
        COUNT(DISTINCT CASE WHEN COALESCE(ct.transfer_detail_url,'')<>'' THEN ct.transfer_detail_url END) unique_events
        FROM club_transfers ct JOIN clubs cl ON cl.id=ct.club_id JOIN leagues l ON l.id=cl.league_id
        WHERE l.country_id=? AND COALESCE(l.section,'senior')='senior' AND COALESCE(cl.source_active,1)=1 AND ct.season_id=?""",(cid,sid)).fetchone()
    transfer_clubs=con.execute("""SELECT COUNT(DISTINCT cts.club_id)
        FROM club_transfer_summary cts JOIN clubs cl ON cl.id=cts.club_id JOIN leagues l ON l.id=cl.league_id
        WHERE l.country_id=? AND COALESCE(l.section,'senior')='senior' AND COALESCE(cl.source_active,1)=1 AND cts.season_id=?""",(cid,sid)).fetchone()[0]
    con.close()
    return {
        'country':country, 'leagues':league_count, 'clubs':club_count,
        'mapped':mapped_count, 'unmapped':max(0,club_count-mapped_count),
        'players':player_count, 'staff':staff_count,
        'club_info':info_count, 'stadiums':stadium_count,
        'transfer_rows':int(tr['rows'] or 0),'transfer_arrivals':int(tr['arrivals'] or 0),
        'transfer_departures':int(tr['departures'] or 0),'transfer_unique':int(tr['unique_events'] or 0),
        'transfer_clubs':int(transfer_clubs or 0),'season_label':season_label_from_id(sid),
    }


def refresh_league_membership_live(league):
    source=(league['source_type'] or 'transfermarkt').lower()
    if source=='srbijasport':
        if not league['source_url']:
            raise RuntimeError('Missing SrbijaSport source URL.')
        html=fetch_public_html(league['source_url'],'https://srbijasport.net/')
        parsed=parse_srbijasport_league_clubs(html,league['source_expected_clubs'] or 0)
        result=sync_srbijasport_league(league['id'],parsed)
        if int(league['level'] or 0) >= 4:
            return {
                'source':'SrbijaSport', 'found':result['found'],
                'membership':f"active {result['active']}, added {result['added']}, updated {result['updated']}",
                'resolved':'TM mapping postponed for Level 4',
            }
        resolve=resolve_tm_for_league(league['id'])
        return {
            'source':'SrbijaSport', 'found':result['found'],
            'membership':f"active {result['active']}, added {result['added']}, updated {result['updated']}",
            'resolved':f"matched {resolve['matched']}, review {resolve['review']}, unresolved {resolve['unresolved']}",
        }

    if not league['tm_url']:
        raise RuntimeError('Missing Transfermarkt league URL.')
    html=fetch_html(league['tm_url'])
    parsed=parse_transfermarkt_league_clubs(html)
    updated,unmatched=apply_transfermarkt_league_clubs(league['id'],parsed,insert_unmatched=((league['section'] or 'senior')=='youth'))
    return {
        'source':'Transfermarkt', 'found':len(parsed),
        'membership':f'updated {updated}',
        'resolved':('unmatched '+', '.join(unmatched)) if unmatched else 'all mapped',
    }



def update_zone_memberships(cid):
    """Update all Level 4 memberships from SrbijaSport only; no TM work."""
    con=db()
    country=con.execute('SELECT * FROM countries WHERE id=?',(cid,)).fetchone()
    leagues=con.execute("""SELECT * FROM leagues
                           WHERE country_id=? AND level=4 AND source_type='srbijasport'
                           ORDER BY name""",(cid,)).fetchall()
    con.close()
    if not country:
        raise RuntimeError('Country not found.')

    out={'country':country['name'],'leagues':len(leagues),
         'expected':sum(int(l['source_expected_clubs'] or 0) for l in leagues),
         'found':0,'active':0,'added':0,'updated':0,'details':[],'errors':[]}

    for i,league in enumerate(leagues,1):
        print(f'[ZONES {i}/{len(leagues)}] {league["name"]}')
        try:
            html=fetch_public_html(league['source_url'],'https://srbijasport.net/')
            parsed=parse_srbijasport_league_clubs(html,league['source_expected_clubs'] or 0)
            result=sync_srbijasport_league(league['id'],parsed)
            out['found']+=result['found']
            out['active']+=result['active']
            out['added']+=result['added']
            out['updated']+=result['updated']
            expected=int(league['source_expected_clubs'] or 0)
            flag='OK' if result['active']==expected else 'CHECK'
            out['details'].append(f'{league["name"]}: {result["active"]}/{expected} {flag}')
        except Exception as e:
            out['errors'].append(f'{league["name"]}: {e}')
    return out



def refresh_level_memberships(cid,level):
    """Refresh all SrbijaSport memberships for one grouped level."""
    if int(level) in (3,4,5,6):
        try: discover_srbijasport_pyramid(cid,False,levels=(int(level),))
        except Exception: pass
    con=db()
    leagues=con.execute("""SELECT * FROM leagues
        WHERE country_id=? AND level=? AND source_type='srbijasport'
        AND COALESCE(source_active,1)=1
        ORDER BY name""",(cid,level)).fetchall()
    con.close()

    out={'leagues':len(leagues),'expected':0,'active':0,'added':0,'updated':0,'details':[],'errors':[]}
    for i,league in enumerate(leagues,1):
        expected=int(league['source_expected_clubs'] or 0)
        out['expected']+=expected
        print(f'[LEVEL {level} MEMBERSHIP {i}/{len(leagues)}] {league["name"]}')
        try:
            html=fetch_public_html(league['source_url'],'https://srbijasport.net/')
            parsed=parse_srbijasport_league_clubs(html,expected)
            result=sync_srbijasport_league(league['id'],parsed)
            out['active']+=result['active']
            out['added']+=result['added']
            out['updated']+=result['updated']
            state='OK' if (expected==0 or result['active']==expected) else 'CHECK'
            ratio=str(result['active']) if expected==0 else f'{result["active"]}/{expected}'
            out['details'].append(f'{league["name"]}: {ratio} {state}')
        except Exception as e:
            out['errors'].append(f'{league["name"]}: {e}')
    return out


def batch_update_level(cid,level,fetch_staff_dob=False):
    """Update every active mapped club across all leagues in one level."""
    con=db()
    active_total=con.execute("""SELECT COUNT(*) FROM clubs cl
        JOIN leagues l ON l.id=cl.league_id
        WHERE l.country_id=? AND l.level=? AND COALESCE(cl.source_active,1)=1""",
        (cid,level)).fetchone()[0]
    clubs=con.execute("""SELECT cl.*,l.name league_name FROM clubs cl
        JOIN leagues l ON l.id=cl.league_id
        WHERE l.country_id=? AND l.level=? AND COALESCE(cl.source_active,1)=1
        AND COALESCE(cl.tm_url,'')<>''
        ORDER BY l.name,cl.name""",(cid,level)).fetchall()
    con.close()

    out={
        'active':active_total,'mapped':len(clubs),'skipped':max(0,active_total-len(clubs)),
        'ok':0,'players':0,'staff':0,'errors':[]
    }
    session=requests.Session()

    for i,c in enumerate(clubs,1):
        cname=c['cm_name'] or c['name']
        print(f'[LEVEL {level} UPDATE {i}/{len(clubs)}] {c["league_name"]} / {cname}')
        club_ok=True
        try:
            start_html=fetch_html(c['tm_url'],session)
            info,stadium=parse_club_info_html(start_html,c['tm_club_id'])
            save_club_info(c['id'],info,stadium,True)
            players=parse_players_html(start_html,c['tm_url'])
            save_players(c['id'],players,f'level {level} batch live Transfermarkt')
            out['players']+=len(players)
        except Exception as e:
            club_ok=False
            out['errors'].append(f'{c["league_name"]} / {cname} players/info: {e}')

        try:
            if c['staff_url']:
                staff_html=fetch_html(c['staff_url'],session)
                staff=parse_staff_html(staff_html,c['staff_url'],cname)
                if fetch_staff_dob:
                    enrich_dob(staff)
                save_staff(c['id'],staff,f'level {level} batch live Transfermarkt')
                out['staff']+=len(staff)
        except Exception as e:
            club_ok=False
            out['errors'].append(f'{c["league_name"]} / {cname} staff: {e}')

        if club_ok:
            out['ok']+=1
        if i < len(clubs):
            time.sleep(0.8)

    return out


def batch_update_country(cid, fetch_staff_dob=False, refresh_membership=False):
    con=db()
    country=con.execute('SELECT * FROM countries WHERE id=?',(cid,)).fetchone()
    leagues=con.execute("SELECT * FROM leagues WHERE country_id=? AND COALESCE(section,'senior')='senior' AND COALESCE(source_active,1)=1 ORDER BY level,name",(cid,)).fetchall()
    con.close()
    if not country:
        raise RuntimeError('Country not found.')

    result={'country':country['name'],'leagues':len(leagues),'clubs':0,'clubs_ok':0,
            'players':0,'staff':0,'membership':[],'errors':[]}

    for idx,league in enumerate(leagues,1):
        lname=league['name']
        print(f'[COUNTRY {idx}/{len(leagues)}] {lname}')
        if refresh_membership:
            try:
                rr=refresh_league_membership_live(league)
                result['membership'].append(
                    f"{lname}: {rr['source']} found {rr['found']}; {rr['membership']}; {rr['resolved']}"
                )
            except Exception as e:
                result['errors'].append(f'{lname} membership: {e}')
        try:
            batch=batch_update_league(league['id'],fetch_staff_dob)
            result['clubs']+=batch['clubs']; result['clubs_ok']+=batch['ok']
            result['players']+=batch['players']; result['staff']+=batch['staff']
            result['errors'].extend(f'{lname}: {err}' for err in batch['errors'])
        except Exception as e:
            result['errors'].append(f'{lname} update: {e}')
    return result


def country_master_workbook_bytes(cid):
    con=db()
    country=con.execute('SELECT * FROM countries WHERE id=?',(cid,)).fetchone()
    if not country:
        con.close(); raise RuntimeError('Country not found.')
    leagues=con.execute("SELECT * FROM leagues WHERE country_id=? AND COALESCE(section,'senior')='senior' AND COALESCE(source_active,1)=1 ORDER BY level,name",(cid,)).fetchall()
    clubs=con.execute("""SELECT cl.*,l.name league_name,l.level league_level,l.source_type league_source,
        ci.official_name,ci.founded,ci.address_line1,ci.postal_city,ci.city,ci.country,
        ci.phone,ci.fax,ci.website,ci.squad_size,ci.average_age,ci.foreigners,ci.foreigners_pct,
        ci.national_team_players,ci.transfer_record,ci.crest_url,ci.crest_local,
        ci.updated_at club_info_updated,s.name stadium_name,s.capacity stadium_capacity,s.city stadium_city
        FROM clubs cl JOIN leagues l ON l.id=cl.league_id
        LEFT JOIN club_info ci ON ci.club_id=cl.id LEFT JOIN stadiums s ON s.club_id=cl.id
        WHERE l.country_id=? AND COALESCE(l.section,'senior')='senior' AND COALESCE(cl.source_active,1)=1
        ORDER BY l.level,l.name,cl.name""",(cid,)).fetchall()
    players=con.execute("""SELECT p.*,COALESCE(cl.cm_name,cl.name) club_name,l.name league_name,l.level league_level
        FROM players p JOIN clubs cl ON cl.id=p.club_id JOIN leagues l ON l.id=cl.league_id
        WHERE l.country_id=? AND COALESCE(l.section,'senior')='senior' AND COALESCE(cl.source_active,1)=1
        ORDER BY l.level,l.name,club_name,p.position,p.name""",(cid,)).fetchall()
    staff=con.execute("""SELECT s.*,COALESCE(cl.cm_name,cl.name) club_name,l.name league_name,l.level league_level
        FROM staff s JOIN clubs cl ON cl.id=s.club_id JOIN leagues l ON l.id=cl.league_id
        WHERE l.country_id=? AND COALESCE(l.section,'senior')='senior' AND COALESCE(cl.source_active,1)=1
        ORDER BY l.level,l.name,club_name,s.section,s.position,s.name""",(cid,)).fetchall()
    sid=current_season_id()
    transfers=con.execute("""SELECT ct.*,COALESCE(cl.cm_name,cl.name) source_club,l.name source_league,l.level source_level
        FROM club_transfers ct JOIN clubs cl ON cl.id=ct.club_id JOIN leagues l ON l.id=cl.league_id
        WHERE l.country_id=? AND COALESCE(l.section,'senior')='senior' AND COALESCE(cl.source_active,1)=1 AND ct.season_id=?
        ORDER BY l.level,l.name,source_club,ct.direction,ct.player_name""",(cid,sid)).fetchall()
    summaries=con.execute("""SELECT cts.* FROM club_transfer_summary cts
        JOIN clubs cl ON cl.id=cts.club_id JOIN leagues l ON l.id=cl.league_id
        WHERE l.country_id=? AND COALESCE(l.section,'senior')='senior' AND COALESCE(cl.source_active,1)=1 AND cts.season_id=?""",(cid,sid)).fetchall()
    con.close()

    wb=Workbook(); wb.remove(wb.active)
    matched=sum(1 for r in clubs if r['tm_url'])
    _write_sheet(wb,'Summary',['Field','Value'],[
        ['Country',country['name']],['Leagues',len(leagues)],['Clubs',len(clubs)],
        ['Transfermarkt mapped',matched],['Transfermarkt unmapped',len(clubs)-matched],
        ['Players',len(players)],['Staff',len(staff)],
        [f'Transfer records {season_label_from_id(sid)}',len(transfers)],
        [f'Clubs with transfer update {season_label_from_id(sid)}',len(summaries)],
        ['Note','Raw Serbia-wide CM DATA HUB export. Final CM0102 import mapping will be added later.']
    ],[28,90])

    _write_sheet(wb,'Leagues',
        ['Level','League','Membership Source','Source URL','Expected Clubs','Active Clubs','TM Mapped'],
        [[l['level'],l['name'],l['source_type'],l['source_url'] or l['tm_url'],l['source_expected_clubs'],
          sum(1 for c in clubs if c['league_name']==l['name']),
          sum(1 for c in clubs if c['league_name']==l['name'] and c['tm_url'])] for l in leagues],
        [10,28,20,70,16,14,14])

    club_headers=['Country','Level','League','League Source','Club','League Source Name','Source Club URL',
        'TM Match Status','TM Match Score','Transfermarkt Name','TM Club ID','Official Club Name','Osnovano',
        'Address','Postal/City','City','TM Country','Phone','Fax','Website','Squad Size','Average Age',
        'Foreigners','Foreigners %','National Team Players','Transfer Record','Stadium','Stadium City',
        'Stadium Capacity','Crest URL','Crest Local File','TM Club URL','TM Staff URL','Club Info Updated',
        'Players Updated','Staff Updated']
    club_rows=[]
    for r in clubs:
        club_rows.append([country['name'],r['league_level'],r['league_name'],r['league_source'],
            r['cm_name'] or r['name'],r['source_name'],r['source_club_url'],r['tm_match_status'],
            r['tm_match_score'],r['tm_name'],r['tm_club_id'],r['official_name'],r['founded'],
            r['address_line1'],r['postal_city'],r['city'],r['country'],r['phone'],r['fax'],r['website'],
            r['squad_size'],r['average_age'],r['foreigners'],r['foreigners_pct'],r['national_team_players'],
            r['transfer_record'],r['stadium_name'],r['stadium_city'],r['stadium_capacity'],r['crest_url'],
            r['crest_local'],r['tm_url'],r['staff_url'],r['club_info_updated'],r['last_players_update'],r['last_staff_update']])
    _write_sheet(wb,'Clubs',club_headers,club_rows,
        [14,8,28,18,28,28,55,16,12,28,10,32,14,28,22,20,18,18,18,28,12,12,12,12,18,16,32,22,14,55,28,55,55,20,20,20])

    city_seen=set(); city_rows=[]
    for r in clubs:
        city=clean_text(r['city'] or r['stadium_city']); ctry=clean_text(r['country'] or country['name'])
        key=(city.lower(),ctry.lower())
        if city and key not in city_seen:
            city_seen.add(key); city_rows.append([city,ctry,r['postal_city']])
    _write_sheet(wb,'Cities',['City','Country','Source postal/city'],city_rows,[26,20,24])

    _write_sheet(wb,'Stadiums',['League','Stadium','Club','City','Capacity','TM Club ID'],
        [[r['league_name'],r['stadium_name'],r['cm_name'] or r['name'],r['stadium_city'],r['stadium_capacity'],r['tm_club_id']]
         for r in clubs if r['stadium_name']], [28,34,28,22,12,12])

    _write_sheet(wb,'Players',
        ['Level','League','Club','Squad No','Player Name','TM Player ID','Game Safe Name','Safe Name Length','Position','Date of Birth','Age','Nationality','Market Value','Player URL'],
        [[r['league_level'],r['league_name'],r['club_name'],r['squad_number'],r['name'],r['tm_player_id'],r['game_safe_name'],
          r['safe_name_length'],r['position'],r['date_of_birth'],r['age'],r['nationality'],r['market_value'],r['profile_url']]
         for r in players], [8,28,28,10,24,13,24,16,22,16,8,28,14,58])

    _write_sheet(wb,'Staff',
        ['Level','League','Section','Name','Club','Nationality','Date of birth','Age','Appointed','Contract expires','Last club','Profile_URL','DOB_Source','Position'],
        [[r['league_level'],r['league_name'],r['section'],r['name'],r['club_name'],r['nationality'],r['date_of_birth'],
          r['age'],r['appointed'],r['contract_expires'],r['last_club'],r['profile_url'],r['dob_source'],r['position']]
         for r in staff], [8,28,22,24,28,28,16,8,14,16,28,58,18,28])

    locations={}
    for r in players:
        key=clean_text(r['tm_player_id'] or '')
        if key:
            locations[key]=(r['club_name'],r['league_name'])
    _write_sheet(wb,'Transfers',
        ['Level','League','Club','Season','Direction','Player Name','TM Player ID','Current Serbia Club','Current Serbia League',
         'Position','Age','Nationality','Market Value','Other Club','TM Other Club ID','Other League','Other Country',
         'Fee','Date','Player URL','Other Club URL','Transfer URL'],
        [[r['source_level'],r['source_league'],r['source_club'],r['season_label'],
          'DOLAZAK' if r['direction']=='arrival' else 'ODLAZAK',r['player_name'],r['player_tm_id'],
          locations.get(clean_text(r['player_tm_id'] or ''),('',''))[0],
          locations.get(clean_text(r['player_tm_id'] or ''),('',''))[1],
          r['position'],r['age'],r['nationality'],r['market_value'],r['other_club'],r['other_club_tm_id'],
          r['other_league'],r['other_country'],r['fee'],r['transfer_date'],r['player_url'],r['other_club_url'],r['transfer_detail_url']]
         for r in transfers],
        [8,28,28,10,12,24,13,28,28,22,8,24,16,28,13,24,18,16,14,52,52,58])

    bio=io.BytesIO(); wb.save(bio); return bio.getvalue()


def country_crest_zip_bytes(cid):
    con=db()
    country=con.execute('SELECT * FROM countries WHERE id=?',(cid,)).fetchone()
    rows=con.execute("""SELECT cl.id,cl.tm_club_id,COALESCE(cl.cm_name,cl.name) club_name,
        l.name league_name,l.level league_level,ci.crest_url,ci.crest_local
        FROM clubs cl JOIN leagues l ON l.id=cl.league_id LEFT JOIN club_info ci ON ci.club_id=cl.id
        WHERE l.country_id=? AND COALESCE(l.section,'senior')='senior' AND COALESCE(cl.source_active,1)=1
        ORDER BY l.level,l.name,cl.name""",(cid,)).fetchall()
    con.close()
    if not country: raise RuntimeError('Country not found.')

    bio=io.BytesIO(); manifest=[]
    with zipfile.ZipFile(bio,'w',zipfile.ZIP_DEFLATED) as z:
        for r in rows:
            local=r['crest_local']; p=BASE_DIR/local if local else None
            if (not p or not p.exists()) and r['crest_url']:
                local=cache_crest(r['id'],r['crest_url']); p=BASE_DIR/local if local else None
            safe_league=re.sub(r'[^A-Za-z0-9_.-]+','_',r['league_name'])
            safe_club=re.sub(r'[^A-Za-z0-9_.-]+','_',r['club_name'])
            zip_name=''
            if p and p.exists():
                ext=p.suffix.lower(); zip_name=f'Level_{r["league_level"]}/{safe_league}/{safe_club}{ext}'; z.write(p,zip_name)
            manifest.append([r['league_level'],r['league_name'],r['club_name'],r['tm_club_id'],r['crest_url'],zip_name])
        sio=io.StringIO(); w=csv.writer(sio); w.writerow(['Level','League','Club','TM Club ID','Crest URL','ZIP File']); w.writerows(manifest)
        z.writestr('SERBIA_crest_manifest.csv',sio.getvalue().encode('utf-8-sig'))
    return bio.getvalue()

def master_workbook_bytes(lid):
    con=db()
    league=con.execute("""SELECT l.*,c.name country_name FROM leagues l
        JOIN countries c ON c.id=l.country_id WHERE l.id=?""",(lid,)).fetchone()
    clubs=con.execute("""SELECT cl.*,ci.*,s.name stadium_name,s.capacity stadium_capacity,s.city stadium_city,
        ci.updated_at club_info_updated
        FROM clubs cl LEFT JOIN club_info ci ON ci.club_id=cl.id
        LEFT JOIN stadiums s ON s.club_id=cl.id
        WHERE cl.league_id=? AND COALESCE(cl.source_active,1)=1 ORDER BY cl.name""",(lid,)).fetchall()
    players=con.execute("""SELECT p.*,COALESCE(cl.cm_name,cl.name) club_name
        FROM players p JOIN clubs cl ON cl.id=p.club_id
        WHERE cl.league_id=? AND COALESCE(cl.source_active,1)=1 ORDER BY club_name,p.position,p.name""",(lid,)).fetchall()
    staff=con.execute("""SELECT s.*,COALESCE(cl.cm_name,cl.name) club_name
        FROM staff s JOIN clubs cl ON cl.id=s.club_id
        WHERE cl.league_id=? AND COALESCE(cl.source_active,1)=1 ORDER BY club_name,s.section,s.position,s.name""",(lid,)).fetchall()
    sid=current_season_id()
    transfers=con.execute("""SELECT ct.*,COALESCE(cl.cm_name,cl.name) source_club
        FROM club_transfers ct JOIN clubs cl ON cl.id=ct.club_id
        WHERE cl.league_id=? AND COALESCE(cl.source_active,1)=1 AND ct.season_id=?
        ORDER BY source_club,ct.direction,ct.player_name""",(lid,sid)).fetchall()
    con.close()

    wb=Workbook()
    wb.remove(wb.active)
    _write_sheet(wb,'Summary',['Field','Value'],[
        ['Country',league['country_name']],['League',league['name']],
        ['Clubs',len(clubs)],['Players',len(players)],['Staff',len(staff)],
        [f'Transfer records {season_label_from_id(sid)}',len(transfers)],
        ['Note','Raw CM DATA HUB export. Final CM0102 import columns will be mapped later.']
    ],[26,80])

    club_headers=['Country','League','Club','League Source Name','Source Club URL','TM Match Status','TM Match Score',
        'Transfermarkt Name','TM Club ID','Official Club Name','Osnovano',
        'Address','Postal/City','City','TM Country','Phone','Fax','Website','Squad Size','Average Age',
        'Foreigners','Foreigners %','National Team Players','Transfer Record','Crest URL','Crest Local File',
        'TM Club URL','TM Staff URL','Club Info Updated','Players Updated','Staff Updated']
    club_rows=[]
    for r in clubs:
        club_rows.append([
            league['country_name'],league['name'],r['cm_name'] or r['name'],r['source_name'],r['source_club_url'],
            r['tm_match_status'],r['tm_match_score'],r['tm_name'],r['tm_club_id'],
            r['official_name'],r['founded'],r['address_line1'],r['postal_city'],r['city'],r['country'],
            r['phone'],r['fax'],r['website'],r['squad_size'],r['average_age'],r['foreigners'],
            r['foreigners_pct'],r['national_team_players'],r['transfer_record'],r['crest_url'],r['crest_local'],
            r['tm_url'],r['staff_url'],r['club_info_updated'],r['last_players_update'],r['last_staff_update']
        ])
    _write_sheet(wb,'Clubs',club_headers,club_rows,
        [14,22,28,28,55,16,12,28,10,32,14,28,22,20,18,18,18,28,12,12,12,12,18,16,55,28,55,55,20,20,20])

    city_seen=set()
    city_rows=[]
    for r in clubs:
        city=clean_text(r['city'])
        country=clean_text(r['country'] or league['country_name'])
        key=(city.lower(),country.lower())
        if city and key not in city_seen:
            city_seen.add(key)
            city_rows.append([city,country,r['postal_city']])
    _write_sheet(wb,'Cities',['City','Country','Source postal/city'],city_rows,[24,20,24])

    stadium_rows=[[r['stadium_name'],r['cm_name'] or r['name'],r['stadium_city'],r['stadium_capacity'],r['tm_club_id']]
                  for r in clubs if r['stadium_name']]
    _write_sheet(wb,'Stadiums',['Stadium','Club','City','Capacity','TM Club ID'],stadium_rows,[32,28,22,12,12])

    _write_sheet(wb,'Players',
        ['Club','Squad No','Player Name','TM Player ID','Game Safe Name','Safe Name Length','Position','Date of Birth','Age','Nationality','Market Value','Player URL'],
        [[r['club_name'],r['squad_number'],r['name'],r['tm_player_id'],r['game_safe_name'],r['safe_name_length'],r['position'],r['date_of_birth'],r['age'],r['nationality'],r['market_value'],r['profile_url']] for r in players],
        [28,10,24,13,24,16,22,16,8,28,14,58])

    _write_sheet(wb,'Staff',
        ['Section','Name','Club','Nationality','Date of birth','Age','Appointed','Contract expires','Last club','Profile_URL','DOB_Source','Position'],
        [[r['section'],r['name'],r['club_name'],r['nationality'],r['date_of_birth'],r['age'],r['appointed'],r['contract_expires'],r['last_club'],r['profile_url'],r['dob_source'],r['position']] for r in staff],
        [22,24,28,28,16,8,14,16,28,58,18,28])

    all_locations=current_player_locations()
    _write_sheet(wb,'Transfers',
        ['Club','Season','Direction','Player Name','TM Player ID','Current Serbia Club','Current Serbia League',
         'Position','Age','Nationality','Market Value','Other Club','TM Other Club ID','Other League','Other Country',
         'Fee','Date','Player URL','Other Club URL','Transfer URL'],
        [[r['source_club'],r['season_label'],'DOLAZAK' if r['direction']=='arrival' else 'ODLAZAK',
          r['player_name'],r['player_tm_id'],
          all_locations.get(clean_text(r['player_tm_id'] or ''),{}).get('club',''),
          all_locations.get(clean_text(r['player_tm_id'] or ''),{}).get('league',''),
          r['position'],r['age'],r['nationality'],r['market_value'],r['other_club'],r['other_club_tm_id'],
          r['other_league'],r['other_country'],r['fee'],r['transfer_date'],r['player_url'],r['other_club_url'],r['transfer_detail_url']]
         for r in transfers],
        [28,10,12,24,13,28,28,22,8,24,16,28,13,24,18,16,14,52,52,58])

    bio=io.BytesIO()
    wb.save(bio)
    return bio.getvalue()


def crest_zip_bytes(lid):
    con=db()
    rows=con.execute("""SELECT cl.id,cl.tm_club_id,COALESCE(cl.cm_name,cl.name) club_name,
        ci.crest_url,ci.crest_local FROM clubs cl LEFT JOIN club_info ci ON ci.club_id=cl.id
        WHERE cl.league_id=? AND COALESCE(cl.source_active,1)=1 ORDER BY cl.name""",(lid,)).fetchall()
    con.close()
    bio=io.BytesIO()
    manifest=[]
    with zipfile.ZipFile(bio,'w',zipfile.ZIP_DEFLATED) as z:
        for r in rows:
            local=r['crest_local']
            p=BASE_DIR/local if local else None
            if (not p or not p.exists()) and r['crest_url']:
                local=cache_crest(r['id'],r['crest_url'])
                p=BASE_DIR/local if local else None
            if p and p.exists():
                safe=re.sub(r'[^A-Za-z0-9_.-]+','_',r['club_name'])
                ext=p.suffix.lower()
                z.write(p,f'{safe}{ext}')
                manifest.append([r['club_name'],r['tm_club_id'],r['crest_url'],f'{safe}{ext}'])
            else:
                manifest.append([r['club_name'],r['tm_club_id'],r['crest_url'],''])
        sio=io.StringIO()
        w=csv.writer(sio)
        w.writerow(['Club','TM Club ID','Crest URL','ZIP File'])
        w.writerows(manifest)
        z.writestr('crest_manifest.csv',sio.getvalue().encode('utf-8-sig'))
    return bio.getvalue()


def level_brand_image(level):
    level=int(level or 0)
    return {
        1:'/static/branding/super_liga.jpg',
        2:'/static/branding/prva_liga.jpg',
        3:'/static/branding/srpske_lige.jpg',
        4:'/static/branding/zonske_lige.jpg',
        5:'/static/branding/okruzne_lige.jpg',
    }.get(level,'/static/branding/fss.jpg')


def level_title(level):
    return {
        1:'SuperLiga Srbije',
        2:'Prva liga Srbije',
        3:'Srpske lige',
        4:'Zonske lige',
        5:'Okružne lige',
    }.get(int(level or 0),f'Level {level}')


def level_group_stats(leagues, level):
    rows=[r for r in leagues if int(r['level'] or 0)==int(level)]
    return {
        'rows':rows,
        'leagues':len(rows),
        'clubs':sum(int(r['cc'] or 0) for r in rows),
        'mapped':sum(int(r['mapped'] or 0) for r in rows),
        'players':sum(int(r['pc'] or 0) for r in rows),
        'staff':sum(int(r['sc'] or 0) for r in rows),
    }


def parse_national_team_info_html(html,tm_team_id=''):
    soup=BeautifulSoup(html,'lxml')
    header={}
    for li in soup.select('li.data-header__label'):
        txt=clean_text(li.get_text(' ',strip=True))
        if ':' in txt:
            k,v=txt.split(':',1)
            header[clean_text(k)]=clean_text(v)

    facts=_info_table_values(soup)
    official_name=''
    h1=soup.find('h1')
    if h1:
        official_name=clean_text(h1.get_text(' ',strip=True))

    association=''
    for key in ('Savez','Football Savez','Confederation'):
        vals=facts.get(key) or []
        if vals:
            association=vals[0]
            break

    founded=(facts.get('Osnovano') or [''])[0]

    crest_url=''
    preferred=[]
    fallback=[]
    for img in soup.find_all('img'):
        for raw in (img.get('src') or '',img.get('data-src') or '',img.get('data-lazy-src') or ''):
            src=clean_text(raw)
            if not src:
                continue
            if src.startswith('//'):
                src='https:'+src
            elif src.startswith('/'):
                src='https://tmssl.akamaized.net'+src
            if '/wappen/' not in src and '/images/wappen/' not in src:
                continue
            if tm_team_id and str(tm_team_id) in src:
                preferred.append(src)
            else:
                fallback.append(src)

    if preferred:
        crest_url=preferred[0]
    elif fallback:
        crest_url=fallback[0]
    elif tm_team_id:
        crest_url=f'https://tmssl.akamaized.net/images/wappen/head/{tm_team_id}.png'

    return {
        'official_name':official_name,
        'association':association,
        'founded':founded,
        'squad_size':header.get('Broj igrača',''),
        'average_age':header.get('Prosečna starost',''),
        'total_market_value':header.get('Ukupna tržišna vrednost','') or header.get('Market value',''),
        'crest_url':crest_url,
    }


def cache_national_crest(team_id,tm_team_id,crest_url):
    if not crest_url:
        return ''
    folder=BASE_DIR/'static'/'national_crests'
    folder.mkdir(parents=True,exist_ok=True)
    try:
        ext='.png'
        p=urlparse(crest_url).path.lower()
        for candidate in ('.png','.jpg','.jpeg','.webp','.svg'):
            if p.endswith(candidate):
                ext=candidate
                break
        dest=folder/f'{team_id}_{tm_team_id}{ext}'
        r=requests.get(crest_url,headers={**HEADERS,'Accept':'image/avif,image/webp,image/apng,image/svg+xml,image/*,*/*;q=0.8'},timeout=30)
        r.raise_for_status()
        if not r.content:
            return ''
        dest.write_bytes(r.content)
        return f'static/national_crests/{dest.name}'
    except Exception as e:
        print('National crest cache failed:',e)
        return ''


def save_national_team_info(team_id,info,download_crest=True):
    con=db()
    team=con.execute('SELECT * FROM national_teams WHERE id=?',(team_id,)).fetchone()
    if not team:
        con.close()
        raise RuntimeError('National team not found.')

    crest_local=team['crest_local'] or ''
    if download_crest and info.get('crest_url'):
        crest_local=cache_national_crest(team_id,team['tm_team_id'],info['crest_url']) or crest_local

    con.execute("""UPDATE national_teams SET
        official_name=?,association=?,founded=?,squad_size=?,average_age=?,total_market_value=?,
        crest_url=?,crest_local=?,last_info_update=CURRENT_TIMESTAMP WHERE id=?""",(
        info.get('official_name'),info.get('association'),info.get('founded'),
        info.get('squad_size'),info.get('average_age'),info.get('total_market_value'),
        info.get('crest_url'),crest_local,team_id))
    con.commit()
    con.close()


def save_national_players(team_id,records):
    con=db()
    con.execute('DELETE FROM national_team_players WHERE national_team_id=?',(team_id,))
    for r in records:
        con.execute('''INSERT OR REPLACE INTO national_team_players(
            national_team_id,squad_number,name,tm_player_id,game_safe_name,safe_name_length,position,
            date_of_birth,age,nationality,market_value,profile_url,current_club,updated_at)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,CURRENT_TIMESTAMP)''',(
            team_id,r.get('squad_number'),r.get('name'),r.get('tm_player_id'),r.get('game_safe_name'),
            r.get('safe_name_length'),r.get('position'),r.get('date_of_birth'),r.get('age'),
            r.get('nationality'),r.get('market_value'),r.get('profile_url'),r.get('current_club')))
    con.execute('UPDATE national_teams SET last_players_update=CURRENT_TIMESTAMP WHERE id=?',(team_id,))
    con.commit()
    con.close()


def save_national_staff(team_id,records):
    con=db()
    con.execute('DELETE FROM national_team_staff WHERE national_team_id=?',(team_id,))
    for r in records:
        con.execute("""INSERT OR REPLACE INTO national_team_staff(
            national_team_id,section,name,nationality,date_of_birth,age,appointed,contract_expires,
            last_club,profile_url,dob_source,position,updated_at)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,CURRENT_TIMESTAMP)""",(
            team_id,r.get('section'),r.get('name'),r.get('nationality'),r.get('date_of_birth'),
            r.get('age'),r.get('appointed'),r.get('contract_expires'),r.get('last_club'),
            r.get('profile_url'),r.get('dob_source'),r.get('position')))
    con.execute('UPDATE national_teams SET last_staff_update=CURRENT_TIMESTAMP WHERE id=?',(team_id,))
    con.commit()
    con.close()


def update_national_team_live(team_id,fetch_staff_dob=False):
    con=db()
    team=con.execute('SELECT * FROM national_teams WHERE id=?',(team_id,)).fetchone()
    con.close()
    if not team:
        raise RuntimeError('National team not found.')

    source=(team['source_type'] or 'transfermarkt').lower() if 'source_type' in team.keys() else 'transfermarkt'
    if source=='fss':
        url=team['source_url'] or team['tm_url']
        if not url:
            raise RuntimeError('Missing FSS selection URL.')
        html=fetch_public_html(fss_latin_url(url),'https://fss.rs/')
        players,staff=parse_fss_selection_html(html,fss_latin_url(url))
        save_national_players(team_id,players)
        save_national_staff(team_id,staff)
        con=db()
        con.execute('''UPDATE national_teams SET official_name=?,association='Fudbalski savez Srbije',
            squad_size=?,last_info_update=CURRENT_TIMESTAMP WHERE id=?''',
            (team['name'],str(len(players)),team_id))
        con.commit(); con.close()
        return {'players':len(players),'staff':len(staff),'source':'FSS'}

    session=requests.Session()
    team_html=fetch_html(team['tm_url'],session)

    info=parse_national_team_info_html(team_html,team['tm_team_id'])
    save_national_team_info(team_id,info,True)

    players=parse_players_html(team_html,team['tm_url'])
    save_national_players(team_id,players)

    staff=[]
    if team['staff_url']:
        staff_html=fetch_html(team['staff_url'],session)
        staff=parse_staff_html(staff_html,team['staff_url'],team['name'])
        if fetch_staff_dob:
            enrich_dob(staff)
        save_national_staff(team_id,staff)

    return {'players':len(players),'staff':len(staff),'source':'Transfermarkt'}


def update_all_national_teams(country_id,fetch_staff_dob=False):
    con=db()
    teams=con.execute('SELECT id,name FROM national_teams WHERE country_id=? ORDER BY id',(country_id,)).fetchall()
    con.close()
    result={'teams':len(teams),'ok':0,'players':0,'staff':0,'errors':[]}
    for i,t in enumerate(teams,1):
        try:
            print(f'[NATIONAL {i}/{len(teams)}] {t["name"]}')
            r=update_national_team_live(t['id'],fetch_staff_dob)
            result['ok']+=1
            result['players']+=r['players']
            result['staff']+=r['staff']
        except Exception as e:
            result['errors'].append(f'{t["name"]}: {e}')
        if i<len(teams):
            time.sleep(0.8)
    return result


def national_players_workbook_bytes(rows):
    wb=Workbook()
    ws=wb.active
    ws.title='Players'
    ws.append(['Team','Squad No','Player Name','TM Player ID','Game Safe Name','Safe Name Length','Position',
               'Current Club','Date of Birth','Age','Nationality','Market Value','Player URL'])
    for r in rows:
        ws.append([r['team_name'],r['squad_number'],r['name'],r['tm_player_id'],r['game_safe_name'],
                   r['safe_name_length'],r['position'],r['current_club'],r['date_of_birth'],r['age'],
                   r['nationality'],r['market_value'],r['profile_url']])
    style_sheet(ws,[26,10,24,13,24,16,22,26,16,8,28,14,58])
    bio=io.BytesIO(); wb.save(bio); return bio.getvalue()


def national_staff_workbook_bytes(rows):
    wb=Workbook()
    ws=wb.active
    ws.title='Staff'
    ws.append(['Team','Section','Name','Position','Nationality','Date of birth','Age',
               'Appointed','Contract expires','Last club','Profile_URL','DOB_Source'])
    for r in rows:
        ws.append([r['team_name'],r['section'],r['name'],r['position'],r['nationality'],
                   r['date_of_birth'],r['age'],r['appointed'],r['contract_expires'],
                   r['last_club'],r['profile_url'],r['dob_source']])
    style_sheet(ws,[26,22,24,28,28,16,8,14,16,28,58,18])
    bio=io.BytesIO(); wb.save(bio); return bio.getvalue()


