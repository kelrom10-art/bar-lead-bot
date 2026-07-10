#!/usr/bin/env python3
"""
Bar Lead Manager — Unified Render Deployment
FastAPI (REST + PWA) + Telegram Bot + APScheduler + Multi-User Auth
PostgreSQL via psycopg2

Environment variables required:
  DATABASE_URL   — PostgreSQL connection string
  BOT_TOKEN      — Telegram bot token
  OWNER_CHAT_ID  — Telegram chat ID of the owner
  PORT           — Port to listen on (Render sets this automatically)
"""

import asyncio, hashlib, hmac as _hmac, io, json, logging, os, secrets, threading, uuid
from contextlib import asynccontextmanager, contextmanager
from datetime import datetime, timedelta
from typing import Optional
from zoneinfo import ZoneInfo
import urllib.request as _ureq

import psycopg2
import psycopg2.extras
import psycopg2.pool

from fastapi import FastAPI, HTTPException, Depends
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse, StreamingResponse, FileResponse
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials
from pydantic import BaseModel, Field, field_validator
import uvicorn
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

from telegram import (
    Update, InlineKeyboardButton, InlineKeyboardMarkup,
    ReplyKeyboardMarkup, KeyboardButton,
)
from telegram.ext import (
    Application, CommandHandler, MessageHandler, ConversationHandler,
    CallbackQueryHandler, ContextTypes, filters,
)
from apscheduler.schedulers.asyncio import AsyncIOScheduler

# ── Config ───────────────────────────────────────────────────────────────────
BOT_TOKEN     = os.getenv("BOT_TOKEN", "")
OWNER_CHAT_ID = int(os.getenv("OWNER_CHAT_ID", "0"))
DATABASE_URL  = os.getenv("DATABASE_URL", "")
PORT          = int(os.getenv("PORT", "8080"))
TZ            = ZoneInfo("Asia/Jerusalem")
STATIC_DIR    = os.path.dirname(os.path.abspath(__file__))

logging.basicConfig(format="%(asctime)s | %(levelname)s | %(message)s", level=logging.INFO)
logger = logging.getLogger(__name__)

# ── Status constants ──────────────────────────────────────────────────────────
ST_NEW          = "new"
ST_PRE_REMINDED = "pre_reminded"
ST_CALLED       = "called"
ST_WON          = "won"
ST_LOST         = "lost"
ST_FOLLOW_UP    = "follow_up"
ST_OLD          = "old"

STATUS_HEB = {
    ST_NEW: "חדש", ST_PRE_REMINDED: "תזכורת נשלחה", ST_CALLED: "עובד",
    ST_WON: "נסגר ✅", ST_LOST: "אבד ❌", ST_FOLLOW_UP: "המשך טיפול", ST_OLD: "ישן",
}
STATUS_EMOJI = {
    ST_NEW: "🆕", ST_PRE_REMINDED: "⏰", ST_CALLED: "📞",
    ST_WON: "✅", ST_LOST: "❌", ST_FOLLOW_UP: "🔄", ST_OLD: "📁",
}

# ─────────────────────────────────────────────────────────────────────────────
#  DATABASE (PostgreSQL via psycopg2)
# ─────────────────────────────────────────────────────────────────────────────

_pool: Optional[psycopg2.pool.ThreadedConnectionPool] = None

def get_pool() -> psycopg2.pool.ThreadedConnectionPool:
    global _pool
    if _pool is None:
        _pool = psycopg2.pool.ThreadedConnectionPool(1, 10, DATABASE_URL)
    return _pool

@contextmanager
def _db():
    pool = get_pool()
    conn = pool.getconn()
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        pool.putconn(conn)

def _now() -> str:
    return datetime.now(TZ).strftime("%Y-%m-%dT%H:%M:%S")

def ensure_db():
    with _db() as conn:
        cur = conn.cursor()
        # Leads table
        cur.execute("""
            CREATE TABLE IF NOT EXISTS leads (
                id             TEXT PRIMARY KEY,
                name           TEXT NOT NULL,
                phone          TEXT NOT NULL,
                call_time      TEXT,
                status         TEXT DEFAULT 'new',
                sale_amount    REAL DEFAULT 0,
                notes          TEXT DEFAULT '',
                created_at     TEXT,
                updated_at     TEXT,
                follow_up_time TEXT,
                tags           TEXT DEFAULT '',
                created_by     TEXT DEFAULT ''
            )
        """)
        cur.execute("ALTER TABLE leads ADD COLUMN IF NOT EXISTS tags TEXT DEFAULT ''")
        cur.execute("ALTER TABLE leads ADD COLUMN IF NOT EXISTS created_by TEXT DEFAULT ''")
        # Users table
        cur.execute("""
            CREATE TABLE IF NOT EXISTS users (
                id            TEXT PRIMARY KEY,
                username      TEXT UNIQUE NOT NULL,
                password_hash TEXT NOT NULL,
                display_name  TEXT DEFAULT '',
                created_at    TEXT
            )
        """)
        # Sessions table
        cur.execute("""
            CREATE TABLE IF NOT EXISTS sessions (
                token      TEXT PRIMARY KEY,
                user_id    TEXT NOT NULL,
                created_at TEXT,
                expires_at TEXT
            )
        """)
    logger.info("DB ✓")

# ── Auth helpers ──────────────────────────────────────────────────────────────

def hash_password(password: str) -> str:
    """PBKDF2-HMAC-SHA256 with random salt."""
    salt = secrets.token_hex(16)
    key = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt.encode("utf-8"), 100_000)
    return f"{salt}:{key.hex()}"

def verify_password(password: str, stored: str) -> bool:
    try:
        salt, key_hex = stored.split(":", 1)
        key = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt.encode("utf-8"), 100_000)
        return _hmac.compare_digest(key.hex(), key_hex)
    except Exception:
        return False

# ── User & Session DB ─────────────────────────────────────────────────────────

def db_count_users() -> int:
    with _db() as conn:
        cur = conn.cursor()
        cur.execute("SELECT COUNT(*) FROM users")
        return cur.fetchone()[0]

def db_get_user_by_username(username: str) -> Optional[dict]:
    with _db() as conn:
        cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        cur.execute("SELECT * FROM users WHERE username=%s", (username.lower().strip(),))
        row = cur.fetchone()
    return dict(row) if row else None

def db_create_user(username: str, password: str, display_name: str = "") -> str:
    uid = str(uuid.uuid4())[:8].upper()
    with _db() as conn:
        cur = conn.cursor()
        cur.execute(
            "INSERT INTO users (id, username, password_hash, display_name, created_at) VALUES (%s,%s,%s,%s,%s)",
            (uid, username.lower().strip(), hash_password(password), display_name or username, _now())
        )
    return uid

def db_create_session(user_id: str) -> str:
    token = secrets.token_urlsafe(32)
    expires = (datetime.now(TZ) + timedelta(days=30)).strftime("%Y-%m-%dT%H:%M:%S")
    with _db() as conn:
        cur = conn.cursor()
        cur.execute(
            "INSERT INTO sessions (token, user_id, created_at, expires_at) VALUES (%s,%s,%s,%s)",
            (token, user_id, _now(), expires)
        )
    return token

def db_get_user_by_token(token: str) -> Optional[dict]:
    now = _now()
    with _db() as conn:
        cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        cur.execute("""
            SELECT u.id, u.username, u.display_name, u.created_at
            FROM users u
            JOIN sessions s ON s.user_id = u.id
            WHERE s.token = %s AND s.expires_at > %s
        """, (token, now))
        row = cur.fetchone()
    return dict(row) if row else None

def db_delete_session(token: str):
    with _db() as conn:
        cur = conn.cursor()
        cur.execute("DELETE FROM sessions WHERE token=%s", (token,))

def db_list_users() -> list:
    with _db() as conn:
        cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        cur.execute("SELECT id, username, display_name, created_at FROM users ORDER BY created_at")
        return [dict(r) for r in cur.fetchall()]

# ── FastAPI Auth Dependency ───────────────────────────────────────────────────

_bearer = HTTPBearer(auto_error=False)

def get_current_user(credentials: Optional[HTTPAuthorizationCredentials] = Depends(_bearer)) -> dict:
    if not credentials:
        raise HTTPException(401, "נדרשת התחברות")
    user = db_get_user_by_token(credentials.credentials)
    if not user:
        raise HTTPException(401, "הסשן פג תוקף — התחבר מחדש")
    return user

# ── Shared DB operations ──────────────────────────────────────────────────────

def db_get_lead(lid: str) -> Optional[dict]:
    with _db() as conn:
        cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        cur.execute("SELECT * FROM leads WHERE id=%s", (lid,))
        row = cur.fetchone()
    return dict(row) if row else None

def db_get_all() -> list:
    with _db() as conn:
        cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        cur.execute("SELECT * FROM leads ORDER BY call_time")
        return [dict(r) for r in cur.fetchall()]

def db_get_active() -> list:
    with _db() as conn:
        cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        cur.execute(
            "SELECT * FROM leads WHERE status IN (%s,%s,%s,%s) ORDER BY call_time",
            (ST_NEW, ST_PRE_REMINDED, ST_CALLED, ST_FOLLOW_UP)
        )
        return [dict(r) for r in cur.fetchall()]

def db_get_today() -> list:
    today = datetime.now(TZ).strftime("%Y-%m-%d")
    with _db() as conn:
        cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        cur.execute(
            "SELECT * FROM leads WHERE call_time LIKE %s ORDER BY call_time",
            (f"{today}%",)
        )
        return [dict(r) for r in cur.fetchall()]

def db_get_old() -> list:
    with _db() as conn:
        cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        cur.execute("SELECT * FROM leads WHERE status=%s ORDER BY updated_at DESC", (ST_OLD,))
        return [dict(r) for r in cur.fetchall()]

def db_search(query: str) -> list:
    q = f"%{query.lower()}%"
    with _db() as conn:
        cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        cur.execute(
            "SELECT * FROM leads WHERE LOWER(name) LIKE %s OR phone LIKE %s",
            (q, q)
        )
        return [dict(r) for r in cur.fetchall()]

def db_add_lead(name: str, phone: str, call_dt) -> str:
    lid = str(uuid.uuid4())[:8].upper()
    now = _now()
    call_time = call_dt.strftime("%Y-%m-%dT%H:%M:%S") if call_dt else None
    with _db() as conn:
        cur = conn.cursor()
        cur.execute(
            "INSERT INTO leads (id, name, phone, call_time, status, created_at, updated_at) "
            "VALUES (%s,%s,%s,%s,%s,%s,%s)",
            (lid, name, phone, call_time, ST_NEW, now, now)
        )
    return lid

def db_add_lead_full(name: str, phone: str, call_time: Optional[str],
                     notes: str, tags: str, created_by: str = "") -> str:
    lid = str(uuid.uuid4())[:8].upper()
    now = _now()
    with _db() as conn:
        cur = conn.cursor()
        cur.execute(
            "INSERT INTO leads (id, name, phone, call_time, status, notes, tags, created_by, created_at, updated_at) "
            "VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",
            (lid, name, phone, call_time, ST_NEW, notes or "", tags or "", created_by, now, now)
        )
    return lid

def db_update_lead(lid: str, **kw):
    kw["updated_at"] = _now()
    if kw.get("status") in (ST_WON, ST_LOST, ST_OLD) and "follow_up_time" not in kw:
        kw["follow_up_time"] = None
    cols = ", ".join(f"{k}=%s" for k in kw)
    vals = list(kw.values()) + [lid]
    with _db() as conn:
        cur = conn.cursor()
        cur.execute(f"UPDATE leads SET {cols} WHERE id=%s", vals)

def db_delete_lead(lid: str):
    with _db() as conn:
        cur = conn.cursor()
        cur.execute("DELETE FROM leads WHERE id=%s", (lid,))

def db_get_leads_pre_reminder() -> list:
    now = datetime.now(TZ)
    out = []
    with _db() as conn:
        cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        cur.execute(
            "SELECT * FROM leads WHERE status=%s AND call_time IS NOT NULL", (ST_NEW,)
        )
        rows = cur.fetchall()
    for r in rows:
        try:
            dt   = datetime.strptime(r["call_time"][:19], "%Y-%m-%dT%H:%M:%S").replace(tzinfo=TZ)
            diff = (dt - now).total_seconds() / 60
            if 0 <= diff <= 20:
                out.append(dict(r))
        except Exception:
            pass
    return out

def db_get_leads_post_call() -> list:
    now = datetime.now(TZ)
    out = []
    with _db() as conn:
        cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        cur.execute(
            "SELECT * FROM leads WHERE status IN (%s,%s) AND call_time IS NOT NULL",
            (ST_NEW, ST_PRE_REMINDED)
        )
        rows = cur.fetchall()
    for r in rows:
        try:
            dt   = datetime.strptime(r["call_time"][:19], "%Y-%m-%dT%H:%M:%S").replace(tzinfo=TZ)
            diff = (now - dt).total_seconds() / 60
            if diff >= 30:
                out.append(dict(r))
        except Exception:
            pass
    return out

def db_get_leads_followup_due() -> list:
    now_str = _now()
    with _db() as conn:
        cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        cur.execute(
            "SELECT * FROM leads WHERE status=%s AND follow_up_time IS NOT NULL AND follow_up_time <= %s",
            (ST_FOLLOW_UP, now_str)
        )
        return [dict(r) for r in cur.fetchall()]

def db_get_stale_leads() -> list:
    threshold = (datetime.now(TZ) - timedelta(days=3)).strftime("%Y-%m-%dT%H:%M:%S")
    with _db() as conn:
        cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        cur.execute(
            "SELECT * FROM leads WHERE status=%s AND updated_at <= %s",
            (ST_CALLED, threshold)
        )
        return [dict(r) for r in cur.fetchall()]

def db_stats_since(since: str) -> dict:
    with _db() as conn:
        cur = conn.cursor()
        cur.execute("SELECT COUNT(*) FROM leads WHERE created_at >= %s", (since,))
        total = cur.fetchone()[0]
        cur.execute(
            "SELECT COUNT(*), COALESCE(SUM(sale_amount),0) FROM leads "
            "WHERE status=%s AND created_at >= %s",
            (ST_WON, since)
        )
        won = cur.fetchone()
        cur.execute(
            "SELECT COUNT(*) FROM leads WHERE status=%s AND created_at >= %s",
            (ST_LOST, since)
        )
        lost = cur.fetchone()[0]
        cur.execute(
            "SELECT COUNT(*) FROM leads WHERE status IN (%s,%s,%s,%s) AND created_at >= %s",
            (ST_NEW, ST_PRE_REMINDED, ST_CALLED, ST_FOLLOW_UP, since)
        )
        active = cur.fetchone()[0]
    return {
        "total": total,
        "won_count": won[0],
        "won_amount": float(won[1]),
        "lost": lost,
        "active": active
    }

# ─────────────────────────────────────────────────────────────────────────────
#  TELEGRAM BOT
# ─────────────────────────────────────────────────────────────────────────────

ASK_NAME, ASK_PHONE, ASK_TIME = range(3)

MENU_BUTTONS = {
    "➕ הוספת ליד", "📋 לידים ישנים", "📊 סיכום", "⏰ תזכורות פעילות",
    "📅 היום", "🔍 חיפוש ליד", "📤 יצוא Excel", "✏️ עריכת ליד",
}

MAIN_MENU = ReplyKeyboardMarkup(
    [
        [KeyboardButton("➕ הוספת ליד"),      KeyboardButton("📋 לידים ישנים")],
        [KeyboardButton("📊 סיכום"),           KeyboardButton("⏰ תזכורות פעילות")],
        [KeyboardButton("📅 היום"),            KeyboardButton("🔍 חיפוש ליד")],
        [KeyboardButton("📤 יצוא Excel"),      KeyboardButton("✏️ עריכת ליד")],
    ],
    resize_keyboard=True,
    is_persistent=True,
)

def parse_time(text: str):
    now = datetime.now(TZ)
    t   = text.strip()
    if "מחר" in t:
        base = (now + timedelta(days=1)).replace(second=0, microsecond=0)
        for w in t.split():
            if ":" in w:
                try:
                    h, m = map(int, w.split(":"))
                    return base.replace(hour=h, minute=m)
                except Exception:
                    pass
        return base.replace(hour=9, minute=0)
    if "שעות" in t or "שעה" in t:
        for w in t.split():
            if w.isdigit():
                return now + timedelta(hours=int(w))
        return now + timedelta(hours=1)
    if "דקות" in t or "דקה" in t:
        for w in t.split():
            if w.isdigit():
                return now + timedelta(minutes=int(w))
        return now + timedelta(minutes=30)
    parts = t.split()
    if len(parts) == 2 and "/" in parts[0] and ":" in parts[1]:
        try:
            day, month = map(int, parts[0].split("/"))
            h, m       = map(int, parts[1].split(":"))
            dt = now.replace(month=month, day=day, hour=h, minute=m, second=0, microsecond=0)
            if dt < now:
                dt = dt.replace(year=now.year + 1)
            return dt
        except Exception:
            pass
    if ":" in t:
        try:
            h, m = map(int, t.split(":"))
            dt   = now.replace(hour=h, minute=m, second=0, microsecond=0)
            return dt if dt > now else dt + timedelta(days=1)
        except Exception:
            pass
    return None

def validate_phone(phone: str) -> bool:
    digits = "".join(c for c in phone if c.isdigit())
    if digits.startswith("972"):
        digits = "0" + digits[3:]
    return len(digits) == 10 and digits.startswith("05")

def outcome_kb(lid: str) -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("✅ נסגר! 🎉",        callback_data=f"out_{lid}_won")],
        [InlineKeyboardButton("⏰ נדחה — קבע זמן",  callback_data=f"out_{lid}_snooze"),
         InlineKeyboardButton("🔄 ליד ישן",          callback_data=f"out_{lid}_old")],
        [InlineKeyboardButton("❌ לא רלוונטי",       callback_data=f"out_{lid}_lost")],
    ])

def snooze_kb(lid: str) -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("בעוד שעה ⏰",  callback_data=f"snz_{lid}_1h"),
         InlineKeyboardButton("מחר 9:00 🌅",  callback_data=f"snz_{lid}_tomorrow")],
        [InlineKeyboardButton("קבע ידנית ✏️", callback_data=f"snz_{lid}_manual")],
    ])

def edit_field_kb(lid: str) -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("👤 שינוי שם",        callback_data=f"edf_{lid}_name")],
        [InlineKeyboardButton("📱 שינוי טלפון",      callback_data=f"edf_{lid}_phone")],
        [InlineKeyboardButton("📅 שינוי שעת שיחה",  callback_data=f"edf_{lid}_time")],
        [InlineKeyboardButton("🗑 מחיקת ליד",        callback_data=f"edf_{lid}_delete")],
    ])

def build_excel() -> io.BytesIO:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "לידים"
    hf = Font(bold=True, color="FFFFFF")
    hb = PatternFill("solid", fgColor="2B579A")
    ca = Alignment(horizontal="center")
    col_labels = ["מזהה", "שם", "טלפון", "שעת שיחה", "סטטוס", "סכום מכירה", "הערות", "נוצר ע״י", "נוצר"]
    col_keys   = ["id",   "name", "phone", "call_time", "status", "sale_amount", "notes", "created_by", "created_at"]
    for ci, label in enumerate(col_labels, 1):
        cell = ws.cell(row=1, column=ci, value=label)
        cell.font = hf; cell.fill = hb; cell.alignment = ca
    for ri, row in enumerate(db_get_all(), 2):
        for ci, key in enumerate(col_keys, 1):
            val = row.get(key, "")
            if key == "status":
                val = STATUS_HEB.get(str(val), str(val))
            elif key in ("call_time", "created_at") and val:
                try:
                    val = datetime.strptime(str(val)[:19], "%Y-%m-%dT%H:%M:%S").strftime("%d/%m/%Y %H:%M")
                except Exception:
                    pass
            elif key == "sale_amount" and val:
                try:
                    val = f"₪{float(val):,.0f}"
                except Exception:
                    pass
            ws.cell(row=ri, column=ci, value=str(val) if val else "")
    for col in ws.columns:
        max_len = max((len(str(c.value or "")) for c in col), default=0)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 40)
    buf = io.BytesIO()
    wb.save(buf); buf.seek(0)
    return buf

async def add_start(u: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data.clear()
    await u.message.reply_text(
        "👤 *מה שם הלקוח?*\n\n_שלח /cancel לביטול_", parse_mode="Markdown")
    return ASK_NAME

async def got_name(u: Update, ctx: ContextTypes.DEFAULT_TYPE):
    name = u.message.text.strip()
    if name in MENU_BUTTONS:
        return await cancel_add(u, ctx)
    ctx.user_data["name"] = name
    await u.message.reply_text(
        f"✅ שם: *{name}*\n\n📱 *מה מספר הטלפון?*\n_לדוגמא: 0501234567_",
        parse_mode="Markdown")
    return ASK_PHONE

async def got_phone(u: Update, ctx: ContextTypes.DEFAULT_TYPE):
    phone = u.message.text.strip()
    if phone in MENU_BUTTONS:
        return await cancel_add(u, ctx)
    if not validate_phone(phone):
        await u.message.reply_text(
            "❌ מספר טלפון לא תקין!\nנא להכניס מספר ישראלי: *0501234567*",
            parse_mode="Markdown")
        return ASK_PHONE
    ctx.user_data["phone"] = phone
    await u.message.reply_text(
        f"✅ טלפון: *{phone}*\n\n"
        "📅 *מתי השיחה?*\n"
        "_לדוגמא:_\n"
        "`16:30` — היום בשעה 16:30\n"
        "`מחר 10:00` — מחר בעשר\n"
        "`25/06 15:30` — תאריך ספציפי\n"
        "`בעוד שעה` — עוד שעה מעכשיו",
        parse_mode="Markdown")
    return ASK_TIME

async def got_time(u: Update, ctx: ContextTypes.DEFAULT_TYPE):
    text = u.message.text.strip()
    if text in MENU_BUTTONS:
        return await cancel_add(u, ctx)
    call_dt = parse_time(text)
    if not call_dt:
        await u.message.reply_text(
            "❓ לא הבנתי. נסה:\n`16:30` | `מחר 10:00` | `25/06 15:30`",
            parse_mode="Markdown")
        return ASK_TIME
    name  = ctx.user_data["name"]
    phone = ctx.user_data["phone"]
    lid   = db_add_lead(name, phone, call_dt)
    ctx.user_data.clear()
    await u.message.reply_text(
        f"🎯 *ליד נוסף בהצלחה!*\n\n"
        f"👤 *{name}*\n"
        f"📱 {phone}\n"
        f"📅 שיחה: *{call_dt.strftime('%d/%m/%Y בשעה %H:%M')}*\n\n"
        f"⏰ תקבל תזכורת 20 דקות לפני",
        parse_mode="Markdown", reply_markup=MAIN_MENU)
    return ConversationHandler.END

async def cancel_add(u: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data.clear()
    if u.message:
        await u.message.reply_text("❌ הוספת ליד בוטלה.", reply_markup=MAIN_MENU)
    return ConversationHandler.END

async def cmd_start(u: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await u.message.reply_text(
        "👋 *שלום! בוט ניהול הלידים שלך* 🍸\n\nבחר פעולה:",
        parse_mode="Markdown", reply_markup=MAIN_MENU)

async def show_today(u: Update, ctx: ContextTypes.DEFAULT_TYPE):
    leads = db_get_today()
    active = [l for l in leads if l.get("status") in (ST_NEW, ST_PRE_REMINDED, ST_CALLED, ST_FOLLOW_UP)]
    if not active:
        await u.message.reply_text("📅 אין שיחות מתוכננות להיום! 🎉", reply_markup=MAIN_MENU)
        return
    lines = [f"📅 *שיחות להיום — {datetime.now(TZ).strftime('%d/%m/%Y')} ({len(active)}):*\n"]
    for ld in active:
        ct = str(ld.get("call_time", ""))
        t  = ""
        try:
            t = datetime.strptime(ct[:19], "%Y-%m-%dT%H:%M:%S").strftime("%H:%M")
        except Exception:
            pass
        lines.append(
            f"🕐 *{t}* — {STATUS_EMOJI.get(ld.get('status'), '❓')} "
            f"*{ld['name']}* | 📱{ld.get('phone', '—')}"
        )
    await u.message.reply_text("\n".join(lines), parse_mode="Markdown", reply_markup=MAIN_MENU)

async def show_old_leads(u: Update, ctx: ContextTypes.DEFAULT_TYPE):
    old = db_get_old()
    if not old:
        await u.message.reply_text("📁 אין לידים ישנים.", reply_markup=MAIN_MENU)
        return
    await u.message.reply_text(f"📁 *לידים ישנים ({len(old)}):*", parse_mode="Markdown")
    for ld in old:
        ct       = str(ld.get("call_time", "") or "")
        date_str = ct[:10].replace("-", "/") if ct else "—"
        kb = InlineKeyboardMarkup([
            [InlineKeyboardButton("📞 לחזור אליו", callback_data=f"recall_{ld['id']}"),
             InlineKeyboardButton("✅ נסגר",        callback_data=f"out_{ld['id']}_won")],
            [InlineKeyboardButton("❌ לא רלוונטי", callback_data=f"out_{ld['id']}_lost")],
        ])
        await u.message.reply_text(
            f"📁 *{ld['name']}*\n📱 {ld.get('phone', '—')}\n📅 שיחה אחרונה: {date_str}",
            reply_markup=kb, parse_mode="Markdown")

async def show_reminders(u: Update, ctx: ContextTypes.DEFAULT_TYPE):
    active = db_get_active()
    if not active:
        await u.message.reply_text("✅ אין תזכורות פעילות!", reply_markup=MAIN_MENU)
        return
    lines = [f"⏰ *תזכורות פעילות ({len(active)}):*\n"]
    for ld in active:
        ct = str(ld.get("call_time") or ld.get("follow_up_time") or "")
        t  = ""
        if ct:
            try:
                t = f" | ⏰ {datetime.strptime(ct[:19], '%Y-%m-%dT%H:%M:%S').strftime('%d/%m %H:%M')}"
            except Exception:
                pass
        lines.append(
            f"{STATUS_EMOJI.get(ld.get('status'), '❓')} *{ld['name']}* "
            f"— 📱{ld.get('phone', '—')}{t}"
        )
    await u.message.reply_text("\n".join(lines), parse_mode="Markdown", reply_markup=MAIN_MENU)

async def show_summary(u: Update, ctx: ContextTypes.DEFAULT_TYPE):
    now     = datetime.now(TZ)
    weekly  = db_stats_since((now - timedelta(days=7)).strftime("%Y-%m-%dT%H:%M:%S"))
    monthly = db_stats_since(now.replace(day=1, hour=0, minute=0, second=0).strftime("%Y-%m-%dT%H:%M:%S"))
    rate    = lambda s: f"{s['won_count']/s['total']*100:.0f}%" if s["total"] else "—"
    await u.message.reply_text(
        "📊 *סיכום ביצועים*\n\n"
        f"📅 *שבוע אחרון:*\n"
        f"  לידים: {weekly['total']} | ✅ נסגרו: {weekly['won_count']} | ❌ אבדו: {weekly['lost']}\n"
        f"  💰 הכנסות: ₪{weekly['won_amount']:,.0f} | המרה: {rate(weekly)}\n\n"
        f"📆 *החודש הנוכחי:*\n"
        f"  לידים: {monthly['total']} | ✅ נסגרו: {monthly['won_count']} | ❌ אבדו: {monthly['lost']}\n"
        f"  💰 הכנסות: ₪{monthly['won_amount']:,.0f} | המרה: {rate(monthly)}\n\n"
        f"🟢 פעילים: {len(db_get_active())} | 📁 ישנים: {len(db_get_old())}",
        parse_mode="Markdown", reply_markup=MAIN_MENU)

async def start_search(u: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data["state"] = "searching"
    await u.message.reply_text("🔍 *חיפוש ליד*\n\nהכנס שם או מספר טלפון:", parse_mode="Markdown")

async def export_excel_bot(u: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await u.message.reply_text("⏳ מכין קובץ Excel...", reply_markup=MAIN_MENU)
    try:
        buf   = build_excel()
        fname = f"leads_{datetime.now(TZ).strftime('%d%m%Y_%H%M')}.xlsx"
        await u.message.reply_document(
            document=buf, filename=fname,
            caption=f"📊 כל הלידים — {datetime.now(TZ).strftime('%d/%m/%Y %H:%M')}")
    except Exception as e:
        logger.error(f"export_excel_bot: {e}")
        await u.message.reply_text(f"❌ שגיאה ביצירת קובץ: {e}")

async def start_edit(u: Update, ctx: ContextTypes.DEFAULT_TYPE):
    active = db_get_active()
    if not active:
        await u.message.reply_text("❌ אין לידים פעילים לעריכה.", reply_markup=MAIN_MENU)
        return
    buttons = []
    for ld in active[:10]:
        ct = str(ld.get("call_time", ""))
        t  = ""
        try:
            t = f" {datetime.strptime(ct[:19], '%Y-%m-%dT%H:%M:%S').strftime('%d/%m %H:%M')}"
        except Exception:
            pass
        buttons.append([InlineKeyboardButton(
            f"{STATUS_EMOJI.get(ld.get('status'), '❓')} {ld['name']}{t}",
            callback_data=f"edit_{ld['id']}"
        )])
    await u.message.reply_text(
        "✏️ *בחר ליד לעריכה:*",
        reply_markup=InlineKeyboardMarkup(buttons), parse_mode="Markdown")

async def handle_text(u: Update, ctx: ContextTypes.DEFAULT_TYPE):
    text  = u.message.text.strip()
    state = ctx.user_data.get("state")

    if state == "ask_amount":
        lid = ctx.user_data.get("lid")
        try:
            amount = float(text.replace("₪", "").replace(",", ""))
            ld     = db_get_lead(lid)
            db_update_lead(lid, status=ST_WON, sale_amount=amount)
            ctx.user_data.clear()
            await u.message.reply_text(
                f"🎉 *כל הכבוד!*\n*{ld['name']}* — עסקה נסגרה\n💰 *₪{amount:,.0f}*",
                parse_mode="Markdown", reply_markup=MAIN_MENU)
        except ValueError:
            await u.message.reply_text(
                "❓ הכנס סכום בשקלים (לדוגמא: `3500`)", parse_mode="Markdown")
        return

    if state in ("ask_snooze", "ask_recall"):
        lid = ctx.user_data.get("lid")
        dt  = parse_time(text)
        if dt:
            if state == "ask_recall":
                db_update_lead(lid, status=ST_NEW,
                               call_time=dt.strftime("%Y-%m-%dT%H:%M:%S"),
                               follow_up_time=None)
            else:
                db_update_lead(lid, status=ST_FOLLOW_UP,
                               follow_up_time=dt.strftime("%Y-%m-%dT%H:%M:%S"))
            ctx.user_data.clear()
            await u.message.reply_text(
                f"⏰ תזכורת נקבעה ל-*{dt.strftime('%d/%m/%Y %H:%M')}* ✓",
                parse_mode="Markdown", reply_markup=MAIN_MENU)
        else:
            await u.message.reply_text(
                "❓ נסה: `16:30` | `מחר 10:00` | `25/06 15:00`", parse_mode="Markdown")
        return

    if state == "searching":
        results = db_search(text)
        ctx.user_data.clear()
        if not results:
            await u.message.reply_text(
                f"🔍 לא נמצאו לידים עבור: *{text}*",
                parse_mode="Markdown", reply_markup=MAIN_MENU)
            return
        await u.message.reply_text(
            f"🔍 *{len(results)} תוצאות עבור \"{text}\":*", parse_mode="Markdown")
        for ld in results[:8]:
            ct = str(ld.get("call_time", ""))
            t  = ""
            try:
                t = f"\n📅 {datetime.strptime(ct[:19], '%Y-%m-%dT%H:%M:%S').strftime('%d/%m/%Y %H:%M')}"
            except Exception:
                pass
            kb = InlineKeyboardMarkup([[
                InlineKeyboardButton("✏️ עריכה",       callback_data=f"edit_{ld['id']}"),
                InlineKeyboardButton("📋 עדכן סטטוס",  callback_data=f"out_{ld['id']}_snooze"),
            ]])
            await u.message.reply_text(
                f"{STATUS_EMOJI.get(ld.get('status'), '❓')} *{ld['name']}*\n"
                f"📱 {ld.get('phone', '—')}{t}",
                parse_mode="Markdown", reply_markup=kb)
        return

    if state == "edit_value":
        lid   = ctx.user_data.get("lid")
        field = ctx.user_data.get("field")
        if field == "name":
            db_update_lead(lid, name=text)
            ctx.user_data.clear()
            await u.message.reply_text(
                f"✅ שם עודכן ל-*{text}*", parse_mode="Markdown", reply_markup=MAIN_MENU)
        elif field == "phone":
            if not validate_phone(text):
                await u.message.reply_text(
                    "❌ מספר לא תקין! נסה שוב (לדוגמא: 0501234567)", parse_mode="Markdown")
                return
            db_update_lead(lid, phone=text)
            ctx.user_data.clear()
            await u.message.reply_text(
                f"✅ טלפון עודכן ל-*{text}*", parse_mode="Markdown", reply_markup=MAIN_MENU)
        elif field == "time":
            dt = parse_time(text)
            if not dt:
                await u.message.reply_text(
                    "❓ נסה: `16:30` | `מחר 10:00` | `25/06 15:00`", parse_mode="Markdown")
                return
            db_update_lead(lid, call_time=dt.strftime("%Y-%m-%dT%H:%M:%S"))
            ctx.user_data.clear()
            await u.message.reply_text(
                f"✅ שעת שיחה עודכנה ל-*{dt.strftime('%d/%m/%Y %H:%M')}*",
                parse_mode="Markdown", reply_markup=MAIN_MENU)
        return

    ctx.user_data.clear()
    if "לידים ישנים" in text:
        await show_old_leads(u, ctx)
    elif "סיכום" in text:
        await show_summary(u, ctx)
    elif "תזכורות" in text:
        await show_reminders(u, ctx)
    elif "היום" in text:
        await show_today(u, ctx)
    elif "חיפוש" in text:
        await start_search(u, ctx)
    elif "Excel" in text or "יצוא" in text:
        await export_excel_bot(u, ctx)
    elif "עריכת ליד" in text:
        await start_edit(u, ctx)
    else:
        await u.message.reply_text("💬 לחץ על אחת האפשרויות בתפריט 👇", reply_markup=MAIN_MENU)

async def callback_handler(u: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = u.callback_query
    await q.answer()
    data = q.data

    if data.startswith("edit_"):
        lid = data.replace("edit_", "")
        ld  = db_get_lead(lid)
        if not ld:
            await q.edit_message_text("❌ ליד לא נמצא"); return
        ct = str(ld.get("call_time", ""))
        t  = ""
        try:
            t = datetime.strptime(ct[:19], "%Y-%m-%dT%H:%M:%S").strftime("%d/%m/%Y %H:%M")
        except Exception:
            pass
        await q.edit_message_text(
            f"✏️ *עריכת ליד*\n\n👤 {ld['name']}\n📱 {ld.get('phone', '—')}\n📅 {t}\n\n*מה לשנות?*",
            reply_markup=edit_field_kb(lid), parse_mode="Markdown")
        return

    if data.startswith("edf_"):
        _, lid, field = data.split("_", 2)
        ld = db_get_lead(lid)
        if not ld:
            await q.edit_message_text("❌ ליד לא נמצא"); return
        if field == "delete":
            db_delete_lead(lid)
            await q.edit_message_text(f"🗑 *{ld['name']}* נמחק.", parse_mode="Markdown"); return
        prompts = {
            "name":  f"👤 הכנס שם חדש עבור *{ld['name']}*:",
            "phone": f"📱 הכנס טלפון חדש עבור *{ld['name']}*:",
            "time":  f"📅 הכנס שעת שיחה חדשה עבור *{ld['name']}*:\n`16:30` | `מחר 10:00` | `25/06 15:00`",
        }
        ctx.user_data.update({"state": "edit_value", "lid": lid, "field": field})
        await q.edit_message_text(prompts.get(field, "?"), parse_mode="Markdown"); return

    if data.startswith("out_"):
        _, lid, action = data.split("_", 2)
        ld = db_get_lead(lid)
        if not ld:
            await q.edit_message_text("❌ ליד לא נמצא"); return
        if action == "won":
            ctx.user_data.update({"state": "ask_amount", "lid": lid})
            await q.edit_message_text(
                f"🎉 *כמה שילם {ld['name']}?*\n_(הכנס סכום בשקלים)_", parse_mode="Markdown")
        elif action == "lost":
            db_update_lead(lid, status=ST_LOST)
            await q.edit_message_text(f"❌ *{ld['name']}* סומן כלא רלוונטי.", parse_mode="Markdown")
        elif action == "snooze":
            await q.edit_message_text(
                f"⏰ *{ld['name']}* — מתי לחזור אליו?",
                reply_markup=snooze_kb(lid), parse_mode="Markdown")
        elif action == "old":
            db_update_lead(lid, status=ST_OLD)
            await q.edit_message_text(f"📁 *{ld['name']}* הועבר ללידים ישנים.", parse_mode="Markdown")
        return

    if data.startswith("snz_"):
        parts  = data.split("_")
        lid    = parts[1]
        option = "_".join(parts[2:])
        now    = datetime.now(TZ)
        if option == "1h":
            new_dt = now + timedelta(hours=1)
        elif option == "tomorrow":
            new_dt = (now + timedelta(days=1)).replace(hour=9, minute=0, second=0, microsecond=0)
        elif option == "manual":
            ctx.user_data.update({"state": "ask_snooze", "lid": lid})
            ld = db_get_lead(lid)
            await q.edit_message_text(
                f"⏰ מתי לחזור אל *{ld['name']}*?\n`16:30` | `מחר 10:00` | `בעוד שעה`",
                parse_mode="Markdown"); return
        else:
            return
        db_update_lead(lid, status=ST_FOLLOW_UP,
                       follow_up_time=new_dt.strftime("%Y-%m-%dT%H:%M:%S"))
        ld = db_get_lead(lid)
        await q.edit_message_text(
            f"⏰ *{ld['name']}* — אחזור ב-*{new_dt.strftime('%d/%m %H:%M')}* 👍",
            parse_mode="Markdown"); return

    if data.startswith("recall_"):
        lid = data.replace("recall_", "")
        ld  = db_get_lead(lid)
        ctx.user_data.update({"state": "ask_recall", "lid": lid})
        await q.edit_message_text(
            f"📞 מתי השיחה עם *{ld['name']}*?\n`16:30` | `מחר 10:00` | `25/06 15:00`",
            parse_mode="Markdown")

# ── Scheduled jobs ────────────────────────────────────────────────────────────

async def job_check_reminders(app):
    for ld in db_get_leads_pre_reminder():
        try:
            dt = datetime.strptime(ld["call_time"][:19], "%Y-%m-%dT%H:%M:%S").replace(tzinfo=TZ)
            await app.bot.send_message(
                chat_id=OWNER_CHAT_ID,
                text=f"⏰ *שיחה בעוד ~20 דקות!*\n\n"
                     f"👤 *{ld['name']}*\n"
                     f"📱 {ld.get('phone', '—')}\n"
                     f"🕐 שיחה בשעה *{dt.strftime('%H:%M')}*",
                parse_mode="Markdown")
            db_update_lead(ld["id"], status=ST_PRE_REMINDED)
        except Exception as e:
            logger.error(f"pre_reminder [{ld.get('id')}]: {e}")
    for ld in db_get_leads_post_call():
        try:
            await app.bot.send_message(
                chat_id=OWNER_CHAT_ID,
                text=f"📞 *מה קרה עם {ld['name']}?*\n📱 {ld.get('phone', '—')}",
                reply_markup=outcome_kb(ld["id"]), parse_mode="Markdown")
            db_update_lead(ld["id"], status=ST_CALLED)
        except Exception as e:
            logger.error(f"post_call [{ld.get('id')}]: {e}")
    for ld in db_get_leads_followup_due():
        try:
            await app.bot.send_message(
                chat_id=OWNER_CHAT_ID,
                text=f"🔄 *זמן לחזור אל {ld['name']}!*\n📱 {ld.get('phone', '—')}",
                reply_markup=outcome_kb(ld["id"]), parse_mode="Markdown")
            db_update_lead(ld["id"], status=ST_CALLED)
        except Exception as e:
            logger.error(f"followup_due [{ld.get('id')}]: {e}")

async def job_morning_briefing(app):
    active = db_get_active()
    old    = db_get_old()
    if not active and not old:
        return
    now   = datetime.now(TZ)
    lines = [f"☀️ *בוקר טוב! {now.strftime('%d/%m/%Y')}*\n"]
    if active:
        lines.append(f"📋 *שיחות ותזכורות ({len(active)}):*")
        for ld in active[:10]:
            ct = str(ld.get("call_time") or ld.get("follow_up_time") or "")
            t  = ""
            if ct:
                try:
                    t = f" ⏰{datetime.strptime(ct[:19], '%Y-%m-%dT%H:%M:%S').strftime('%H:%M')}"
                except Exception:
                    pass
            lines.append(f"{STATUS_EMOJI.get(ld.get('status'), '❓')} *{ld['name']}*{t}")
    if old:
        lines.append(f"\n📁 *{len(old)} לידים ישנים*")
    await app.bot.send_message(chat_id=OWNER_CHAT_ID, text="\n".join(lines), parse_mode="Markdown")

async def job_weekly_report(app):
    s    = db_stats_since((datetime.now(TZ) - timedelta(days=7)).strftime("%Y-%m-%dT%H:%M:%S"))
    rate = f"{s['won_count']/s['total']*100:.0f}%" if s["total"] else "—"
    await app.bot.send_message(
        chat_id=OWNER_CHAT_ID,
        text=f"📊 *דוח שבועי*\n\n"
             f"לידים: {s['total']} | ✅ נסגרו: {s['won_count']} | ❌ אבדו: {s['lost']}\n"
             f"💰 הכנסות: ₪{s['won_amount']:,.0f} | המרה: {rate}\n\nשבוע מוצלח! 🍸",
        parse_mode="Markdown")

async def job_monthly_report(app):
    now  = datetime.now(TZ)
    s    = db_stats_since(now.replace(day=1, hour=0, minute=0, second=0).strftime("%Y-%m-%dT%H:%M:%S"))
    rate = f"{s['won_count']/s['total']*100:.0f}%" if s["total"] else "—"
    await app.bot.send_message(
        chat_id=OWNER_CHAT_ID,
        text=f"📆 *דוח חודשי — {now.strftime('%m/%Y')}*\n\n"
             f"לידים: {s['total']} | ✅ נסגרו: {s['won_count']} | ❌ אבדו: {s['lost']}\n"
             f"💰 הכנסות: ₪{s['won_amount']:,.0f} | המרה: {rate}\n\nחודש מוצלח! 💪",
        parse_mode="Markdown")

async def job_stale_leads(app):
    stale = db_get_stale_leads()
    if not stale:
        return
    lines = [f"🔔 *{len(stale)} לידים מחכים לעדכון (3+ ימים):*\n"]
    for ld in stale[:10]:
        try:
            upd      = datetime.strptime(ld["updated_at"][:19], "%Y-%m-%dT%H:%M:%S")
            days_ago = (datetime.now(TZ).replace(tzinfo=None) - upd).days
        except Exception:
            days_ago = "?"
        lines.append(f"📞 *{ld['name']}* — {ld.get('phone', '—')} _(לפני {days_ago} ימים)_")
    await app.bot.send_message(
        chat_id=OWNER_CHAT_ID,
        text="\n".join(lines) + "\n\nפתח את האפליקציה לעדכון סטטוס 👆",
        parse_mode="Markdown")

def run_bot_thread():
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)

    if not BOT_TOKEN or not OWNER_CHAT_ID:
        logger.warning("BOT_TOKEN or OWNER_CHAT_ID not set — bot not started")
        return

    menu_regex = (
        "^(📋 לידים ישנים"
        "|📊 סיכום"
        "|⏰ תזכורות פעילות"
        "|📅 היום"
        "|🔍 חיפוש ליד"
        "|📤 יצוא Excel"
        "|✏️ עריכת ליד)$"
    )

    async def post_init(application: Application) -> None:
        s = AsyncIOScheduler(timezone="Asia/Jerusalem")
        s.add_job(job_check_reminders,  "interval", minutes=1,                  args=[application])
        s.add_job(job_morning_briefing, "cron",     hour=9,  minute=0,          args=[application])
        s.add_job(job_weekly_report,    "cron",     day_of_week="sun", hour=10, args=[application])
        s.add_job(job_monthly_report,   "cron",     day=1,   hour=10,           args=[application])
        s.add_job(job_stale_leads,      "cron",     hour=10, minute=30,         args=[application])
        s.start()
        logger.info("Scheduler ✓")

    conv = ConversationHandler(
        entry_points=[MessageHandler(filters.Regex("^➕ הוספת ליד$"), add_start)],
        states={
            ASK_NAME:  [MessageHandler(filters.TEXT & ~filters.COMMAND, got_name)],
            ASK_PHONE: [MessageHandler(filters.TEXT & ~filters.COMMAND, got_phone)],
            ASK_TIME:  [MessageHandler(filters.TEXT & ~filters.COMMAND, got_time)],
        },
        fallbacks=[
            CommandHandler("cancel", cancel_add),
            MessageHandler(filters.Regex(menu_regex), cancel_add),
        ],
        allow_reentry=True,
    )

    bot_app = Application.builder().token(BOT_TOKEN).post_init(post_init).build()
    bot_app.add_handler(CommandHandler("start", cmd_start))
    bot_app.add_handler(conv)
    bot_app.add_handler(CallbackQueryHandler(callback_handler))
    bot_app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))

    logger.info("Bot polling started 🚀")
    bot_app.run_polling(drop_pending_updates=True, stop_signals=None)

# ─────────────────────────────────────────────────────────────────────────────
#  FASTAPI APPLICATION
# ─────────────────────────────────────────────────────────────────────────────

@asynccontextmanager
async def lifespan(app: FastAPI):
    ensure_db()
    t = threading.Thread(target=run_bot_thread, daemon=True, name="telegram-bot")
    t.start()
    logger.info("🍸 Bar Lead Manager started")
    yield

fastapi_app = FastAPI(title="Bar Lead Manager API", lifespan=lifespan)
fastapi_app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"], allow_methods=["*"], allow_headers=["*"],
)

# ── Pydantic models ───────────────────────────────────────────────────────────

class LeadCreate(BaseModel):
    name:      str = Field(..., min_length=1, max_length=200)
    phone:     str = Field(..., min_length=1, max_length=50)
    call_time: Optional[str] = None
    notes:     Optional[str] = ""
    tags:      Optional[str] = ""

    @field_validator("name", "phone", mode="before")
    @classmethod
    def strip_and_require(cls, v):
        if isinstance(v, str):
            v = v.strip()
        return v

class LeadUpdate(BaseModel):
    name:           Optional[str]   = Field(None, min_length=1, max_length=200)
    phone:          Optional[str]   = Field(None, min_length=1, max_length=50)
    call_time:      Optional[str]   = None
    status:         Optional[str]   = None
    sale_amount:    Optional[float] = None
    notes:          Optional[str]   = None
    follow_up_time: Optional[str]   = None
    tags:           Optional[str]   = None

class RegisterRequest(BaseModel):
    username:     str = Field(..., min_length=3, max_length=50)
    password:     str = Field(..., min_length=6)
    display_name: Optional[str] = ""

class LoginRequest(BaseModel):
    username: str
    password: str

def enrich(d: dict) -> dict:
    d["status_heb"]   = STATUS_HEB.get(d.get("status", ""), d.get("status", ""))
    d["status_emoji"] = STATUS_EMOJI.get(d.get("status", ""), "")
    return d

# ── Auth Routes ───────────────────────────────────────────────────────────────

@fastapi_app.get("/api/app-config")
def app_config():
    has_users = db_count_users() > 0
    return {"auth_type": "login", "has_users": has_users, "version": "3.0"}

@fastapi_app.post("/api/auth/register", status_code=201)
def register(body: RegisterRequest):
    if db_get_user_by_username(body.username):
        raise HTTPException(400, "שם המשתמש כבר קיים")
    uid   = db_create_user(body.username.strip(), body.password, body.display_name or body.username)
    token = db_create_session(uid)
    user  = db_get_user_by_token(token)
    return {"token": token, "user": user}

@fastapi_app.post("/api/auth/login")
def login(body: LoginRequest):
    user = db_get_user_by_username(body.username.strip())
    if not user or not verify_password(body.password, user["password_hash"]):
        raise HTTPException(401, "שם משתמש או סיסמה שגויים")
    token = db_create_session(user["id"])
    return {
        "token": token,
        "user": {"id": user["id"], "username": user["username"], "display_name": user["display_name"]}
    }

@fastapi_app.post("/api/auth/logout")
def logout(credentials: Optional[HTTPAuthorizationCredentials] = Depends(_bearer)):
    if credentials:
        db_delete_session(credentials.credentials)
    return {"ok": True}

@fastapi_app.get("/api/auth/me")
def get_me(user: dict = Depends(get_current_user)):
    return user

@fastapi_app.get("/api/auth/users")
def list_users_api(user: dict = Depends(get_current_user)):
    return db_list_users()

# ── Protected API Routes ──────────────────────────────────────────────────────

@fastapi_app.get("/health")
def health():
    return {"status": "ok"}

@fastapi_app.get("/api/leads")
def get_leads(status: Optional[str] = None, q: Optional[str] = None,
              user: dict = Depends(get_current_user)):
    if q:
        rows = db_search(q)
    elif status:
        statuses = status.split(",")
        with _db() as conn:
            cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
            placeholders = ",".join(["%s"] * len(statuses))
            cur.execute(
                f"SELECT * FROM leads WHERE status IN ({placeholders}) ORDER BY call_time",
                statuses
            )
            rows = [dict(r) for r in cur.fetchall()]
    else:
        with _db() as conn:
            cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
            cur.execute("SELECT * FROM leads ORDER BY call_time DESC")
            rows = [dict(r) for r in cur.fetchall()]
    return [enrich(r) for r in rows]

@fastapi_app.get("/api/leads/today")
def get_today_api(user: dict = Depends(get_current_user)):
    return [enrich(r) for r in db_get_today()]

@fastapi_app.get("/api/leads/{lid}")
def get_lead_api(lid: str, user: dict = Depends(get_current_user)):
    row = db_get_lead(lid)
    if not row:
        raise HTTPException(404, "Lead not found")
    return enrich(row)

@fastapi_app.post("/api/leads", status_code=201)
def create_lead(body: LeadCreate, user: dict = Depends(get_current_user)):
    created_by = user.get("display_name") or user.get("username", "")
    lid = db_add_lead_full(body.name, body.phone, body.call_time, body.notes or "", body.tags or "", created_by)
    msg = f"🆕 <b>ליד חדש!</b>\n👤 {body.name}\n📱 {body.phone}\n➕ נוסף ע״י {created_by}"
    if body.call_time:
        msg += f"\n📅 {body.call_time}"
    if body.notes:
        msg += f"\n📝 {body.notes}"
    _send_tg(msg)
    return {"id": lid, "message": "Lead created"}

@fastapi_app.put("/api/leads/{lid}")
def update_lead_api(lid: str, body: LeadUpdate, user: dict = Depends(get_current_user)):
    with _db() as conn:
        cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        cur.execute("SELECT * FROM leads WHERE id=%s", (lid,))
        existing = cur.fetchone()
    if not existing:
        raise HTTPException(404, "Lead not found")
    raw = body.model_dump(exclude_unset=True)
    updates = {k: v for k, v in raw.items() if not (k in ("name", "phone") and not v)}
    if not updates:
        return {"message": "No changes"}
    if "call_time" in updates and updates["call_time"]:
        cur_status = dict(existing).get("status", "")
        if cur_status in (ST_PRE_REMINDED, ST_CALLED) and "status" not in updates:
            updates["status"] = ST_NEW
    new_status = updates.get("status")
    if new_status in (ST_WON, ST_LOST, ST_OLD) and "follow_up_time" not in updates:
        updates["follow_up_time"] = None
    db_update_lead(lid, **updates)
    return {"message": "Updated"}

@fastapi_app.delete("/api/leads/{lid}")
def delete_lead_api(lid: str, user: dict = Depends(get_current_user)):
    db_delete_lead(lid)
    return {"message": "Deleted"}

@fastapi_app.get("/api/stats")
def get_stats(user: dict = Depends(get_current_user)):
    now       = datetime.now(TZ)
    week_ago  = (now - timedelta(days=7)).strftime("%Y-%m-%dT%H:%M:%S")
    month_ago = now.replace(day=1, hour=0, minute=0, second=0).strftime("%Y-%m-%dT%H:%M:%S")

    with _db() as conn:
        cur = conn.cursor()
        cur.execute("SELECT COUNT(*) FROM leads"); all_total = cur.fetchone()[0]
        cur.execute(
            "SELECT COUNT(*) FROM leads WHERE status IN (%s,%s,%s,%s)",
            (ST_NEW, ST_PRE_REMINDED, ST_CALLED, ST_FOLLOW_UP)
        ); all_active = cur.fetchone()[0]
        cur.execute("SELECT COUNT(*) FROM leads WHERE status=%s", (ST_OLD,))
        all_old = cur.fetchone()[0]

        status_counts = {}
        for s in [ST_NEW, ST_PRE_REMINDED, ST_CALLED, ST_WON, ST_LOST, ST_FOLLOW_UP, ST_OLD]:
            cur.execute("SELECT COUNT(*) FROM leads WHERE status=%s", (s,))
            status_counts[s] = cur.fetchone()[0]

        cur.execute(
            "SELECT created_at, updated_at FROM leads WHERE status=%s", (ST_WON,)
        )
        won_rows = cur.fetchall()

        cur.execute(
            "SELECT AVG(sale_amount) FROM leads WHERE status=%s AND sale_amount > 0", (ST_WON,)
        )
        avg_sale_row = cur.fetchone()
        avg_sale = round(avg_sale_row[0] or 0, 0)

        daily = []
        for i in range(6, -1, -1):
            day    = now - timedelta(days=i)
            prefix = day.strftime("%Y-%m-%d")
            cur.execute(
                "SELECT COUNT(*) FROM leads WHERE created_at LIKE %s", (f"{prefix}%",)
            )
            cnt = cur.fetchone()[0]
            daily.append({"date": prefix, "label": day.strftime("%d/%m"), "count": cnt})

    if won_rows:
        days_list = []
        for r in won_rows:
            try:
                c = datetime.strptime(str(r[0])[:19], "%Y-%m-%dT%H:%M:%S")
                u = datetime.strptime(str(r[1])[:19], "%Y-%m-%dT%H:%M:%S")
                days_list.append((u - c).total_seconds() / 86400)
            except Exception:
                pass
        avg_days = round(sum(days_list) / len(days_list), 1) if days_list else 0
    else:
        avg_days = 0

    return {
        "weekly":            db_stats_since(week_ago),
        "monthly":           db_stats_since(month_ago),
        "total":             all_total,
        "active":            all_active,
        "old":               all_old,
        "status_counts":     status_counts,
        "avg_days_to_close": avg_days,
        "avg_sale":          avg_sale,
        "daily":             daily,
    }

@fastapi_app.get("/api/export")
def export_excel(user: dict = Depends(get_current_user)):
    buf   = build_excel()
    fname = f"leads_{datetime.now(TZ).strftime('%d%m%Y_%H%M')}.xlsx"
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={fname}"}
    )

@fastapi_app.get("/api/test-telegram")
def test_telegram(user: dict = Depends(get_current_user)):
    if not BOT_TOKEN:
        return {"ok": False, "error": "BOT_TOKEN לא הוגדר"}
    if not OWNER_CHAT_ID:
        return {"ok": False, "error": "OWNER_CHAT_ID לא הוגדר"}
    try:
        req = _ureq.Request(f"https://api.telegram.org/bot{BOT_TOKEN}/getMe")
        res = json.loads(_ureq.urlopen(req, timeout=5).read())
        if res.get("ok"):
            bot = res["result"]
            return {"ok": True, "bot_name": bot["first_name"],
                    "bot_username": bot["username"],
                    "message": f"✅ הבוט @{bot['username']} מחובר!"}
        return {"ok": False, "error": str(res)}
    except Exception as e:
        return {"ok": False, "error": str(e)}

# ── Static files ──────────────────────────────────────────────────────────────

@fastapi_app.get("/sw.js")
def serve_sw():
    p = os.path.join(STATIC_DIR, "sw.js")
    if os.path.exists(p):
        return FileResponse(p, media_type="application/javascript")
    return HTMLResponse("// no sw", 200)

@fastapi_app.get("/manifest.json")
def serve_manifest():
    p = os.path.join(STATIC_DIR, "manifest.json")
    if os.path.exists(p):
        return FileResponse(p, media_type="application/json")
    raise HTTPException(404)

@fastapi_app.get("/", response_class=HTMLResponse)
def serve_app():
    html_path = os.path.join(STATIC_DIR, "index.html")
    if os.path.exists(html_path):
        with open(html_path, encoding="utf-8") as f:
            return HTMLResponse(f.read())
    return HTMLResponse("<h1>index.html not found</h1>", 404)

# ── Telegram helper ───────────────────────────────────────────────────────────

def _send_tg(text: str):
    if not BOT_TOKEN or not OWNER_CHAT_ID:
        return
    def _do():
        try:
            payload = json.dumps({
                "chat_id": OWNER_CHAT_ID, "text": text, "parse_mode": "HTML"
            }).encode()
            req = _ureq.Request(
                f"https://api.telegram.org/bot{BOT_TOKEN}/sendMessage",
                data=payload, headers={"Content-Type": "application/json"}
            )
            _ureq.urlopen(req, timeout=6)
        except Exception:
            pass
    threading.Thread(target=_do, daemon=True).start()

# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    print(f"🍸 Bar Lead Manager — http://0.0.0.0:{PORT}")
    uvicorn.run(fastapi_app, host="0.0.0.0", port=PORT)
