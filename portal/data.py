from __future__ import annotations
import sqlite3
from .config import DB_PATH


def connect():
    # Keep normal connections responsive even while a sync operation is writing.
    # WAL is enabled during init_db(); busy_timeout here prevents long browser stalls.
    con = sqlite3.connect(DB_PATH, timeout=1.5)
    con.row_factory = sqlite3.Row
    con.execute('PRAGMA foreign_keys=ON')
    con.execute('PRAGMA busy_timeout=1500')
    return con


def one(sql, params=()):
    con = connect()
    try:
        return con.execute(sql, params).fetchone()
    finally:
        con.close()


def all(sql, params=()):
    con = connect()
    try:
        return con.execute(sql, params).fetchall()
    finally:
        con.close()
