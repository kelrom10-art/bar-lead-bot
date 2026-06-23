#!/usr/bin/env python3
"""
בוט ניהול לידים לבר 🍸
Google Sheets (אחסון) + Google Calendar (תזכורות)
"""

import os, json, base64, logging, threading, uuid
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from http.server import HTTPServer, BaseHTTPRequestHandler

import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

from telegram import (
    Update, InlineKeyboardButton, InlineKeyboardMarkup,
    ReplyKeyboardMarkup, KeyboardButton,
)
from telegram.ext import (
    Application, CommandHandler, MessageHandler, ConversationHandler,
    CallbackQueryHandler, ContextTypes, filters,
)
from apscheduler.schedulers.asyncio import AsyncIOScheduler

BOT_TOKEN         = os.getenv("BOT_TOKEN", "")
OWNER_CHAT_ID     = int(os.getenv("OWNER_CHAT_ID", "0"))
GOOGLE_CREDS_JSON = os.getenv("GOOGLE_CREDENTIALS_JSON", "")
SPREADSHEET_ID    = os.getenv("GOOGLE_SPREADSHEET_ID", "")
CALENDAR_ID       = os.getenv("GOOGLE_CALENDAR_ID", "primary")
TZ                = ZoneInfo("Asia/Jerusalem")

logging.basicConfig(format="%(asctime)s | %(levelname)s | %(message)s", level=logging.INFO)
logger = logging.getLogger(__name__)

ASK_NAME, ASK_PHONE, ASK_TIME = range(3)

ST_NEW="new"; ST_PRE_REMINDED="pre_reminded"; ST_CALLED="called"
ST_WON="won"; ST_LOST="lost"; ST_FOLLOW_UP="follow_up"; ST_OLD="old"

STATUS_EMOJI = {
    ST_NEW: "🆕", ST_PRE_REMINDED: "⏰", ST_CALLED: "📞",
    ST_WON: "✅", ST_LOST: "❌", ST_FOLLOW_UP: "🔄", ST_OLD: "📁",
}

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/calendar",
]

HEADERS = [
    "id", "name", "phone", "call_time", "status",
    "sale_amount", "notes", "created_at", "updated_at",
    "cal_event_id", "follow_up_time",
]

MAIN_MENU = ReplyKeyboardMarkup(
    [
        [KeyboardButton("➕ הוספת ליד"),      KeyboardButton("📋 לידים ישנים")],
        [KeyboardButton("📊 סיכום"),           KeyboardButton("⏰ תזכורות פעילות")],
    ],
    resize_keyboard=True,
    is_persistent=True,
)

_sheet = None
_cal   = None

def _get_creds():
    raw = GOOGLE_CREDS_JSON
    if not raw:
        raise ValueError("GOOGLE_CREDENTIALS_JSON לא הוגדר!")
    try:
        info = json.loads(base64.b64decode(raw).decode("utf-8"))
    except Exception:
        info = json.loads(raw)
    return Credentials.from_service_account_info(info, scopes=SCOPES)

def init_google():
    global _sheet, _cal
    creds  = _get_creds()
    gc     = gspread.authorize(creds)
    _cal   = build("calendar", "v3", credentials=creds)
    sp     = gc.open_by_key(SPREADSHEET_ID)
    try:
        _sheet = sp.worksheet("Leads")
    except gspread.WorksheetNotFound:
        _sheet = sp.add_worksheet("Leads", rows=2000, cols=len(HEADERS))
        _sheet.append_row(HEADERS)
    if not _sheet.row_values(1) or _sheet.cell(1, 1).value != "id":
        _sheet.insert_row(HEADERS, 1)
    logger.info("Google APIs ✓")

def _now():
    return datetime.now(TZ).strftime("%Y-%m-%dT%H:%M:%S")

def _rows():
    return _sheet.get_all_records()

def _find_row(lead_id):
    cell = _sheet.find(str(lead_id), in_column=1)
    return cell.row if cell else None

def _col(key):
    return HEADERS.index(key) + 1

def add_lead(name, phone, call_dt):
    lid = str(uuid.uuid4())[:8].upper()
    _sheet.append_row([
        lid, name, phone, call_dt.strftime("%Y-%m-%dT%H:%M:%S"), ST_NEW,
        "", "", _now(), _now(), "", "",
    ])
    return lid

def update_lead(lid, **kw):
    row = _find_row(lid)
    if not row:
        logger.error(f"ליד {lid} לא נמצא"); return
    kw["updated_at"] = _now()
    for k, v in kw.items():
        if k in HEADERS:
            _sheet.update_cell(row, _col(k), "" if v is None else str(v))

def get_lead(lid):
    return next((r for r in _rows() if str(r.get("id")) == str(lid)), None)

def get_active_leads():
    return [r for r in _rows()
            if r.get("status") in (ST_NEW, ST_PRE_REMINDED, ST_CALLED, ST_FOLLOW_UP)
            and r.get("name")]

def get_old_leads():
    return [r for r in _rows() if r.get("status") == ST_OLD and r.get("name")]

def get_leads_pre_reminder():
    now = datetime.now(TZ)
    out = []
    for r in _rows():
        if r.get("status") != ST_NEW or not r.get("call_time"):
            continue
        try:
            dt   = datetime.strptime(str(r["call_time"])[:19], "%Y-%m-%dT%H:%M:%S").replace(tzinfo=TZ)
            diff = (dt - now).total_seconds() / 60
            if 0 <= diff <= 20:
                out.append(r)
        except Exception:
            pass
    return out

def get_leads_post_call():
    now = datetime.now(TZ)
    out = []
    for r in _rows():
        if r.get("status") not in (ST_NEW, ST_PRE_REMINDED) or not r.get("call_time"):
            continue
        try:
            dt   = datetime.strptime(str(r["call_time"])[:19], "%Y-%m-%dT%H:%M:%S").replace(tzinfo=TZ)
            diff = (now - dt).total_seconds() / 60
            if diff >= 30:
                out.append(r)
        except Exception:
            pass
    return out

def get_leads_followup_due():
    now = _now()
    return [r for r in _rows()
            if r.get("status") == ST_FOLLOW_UP
            and r.get("follow_up_time")
            and str(r["follow_up_time"]) <= now]

def get_stats(since_dt):
    since = since_dt.strftime("%Y-%m-%dT%H:%M:%S")
    rows  = [r for r in _rows() if r.get("created_at", "") >= since]
    won   = [r for r in rows if r.get("status") == ST_WON]
    lost  = [r for r in rows if r.get("status") == ST_LOST]
    return {
        "total":      len(rows),
        "won_count":  len(won),
        "won_amount": sum(float(r.get("sale_amount") or 0) for r in won),
        "lost":       len(lost),
        "active":     len([r for r in rows if r.get("status") in (ST_NEW, ST_PRE_REMINDED, ST_CALLED, ST_FOLLOW_UP)]),
        "old":        len([r for r in rows if r.get("status") == ST_OLD]),
    }

def cal_create(lid, name, phone, call_dt):
    try:
        ev = {
            "summary":     f"📞 {name}",
            "description": f"טלפון: {phone}\nמזהה ליד: {lid}",
            "start": {"dateTime": call_dt.isoformat(), "timeZone": "Asia/Jerusalem"},
            "end":   {"dateTime": (call_dt + timedelta(minutes=30)).isoformat(), "timeZone": "Asia/Jerusalem"},
            "reminders": {"useDefault": False, "overrides": [{"method": "popup", "minutes": 20}]},
        }
        res = _cal.events().insert(calendarId=CALENDAR_ID, body=ev).execute()
        logger.info(f"אירוע יומן נוצר: {res['id']}")
        return res["id"]
    except Exception as e:
        logger.error(f"cal_create: {e}"); return None

def cal_update(event_id, call_dt):
    if not event_id: return
    try:
        ev = _cal.events().get(calendarId=CALENDAR_ID, eventId=event_id).execute()
        ev["start"]["dateTime"] = call_dt.isoformat()
        ev["end"]["dateTime"]   = (call_dt + timedelta(minutes=30)).isoformat()
        _cal.events().update(calendarId=CALENDAR_ID, eventId=event_id, body=ev).execute()
    except Exception as e:
        logger.error(f"cal_update: {e}")

def cal_delete(event_id):
    if not event_id: return
    try:
        _cal.events().delete(calendarId=CALENDAR_ID, eventId=event_id).execute()
    except Exception as e:
        logger.error(f"cal_delete: {e}")

def parse_time(text):
    now = datetime.now(TZ)
    t   = text.strip()
    if "מחר" in t:
        base = (now + timedelta(days=1)).replace(second=0, microsecond=0)
        for w in t.split():
            if ":" in w:
                try: h, m = map(int, w.split(":")); return base.replace(hour=h, minute=m)
                except: pass
        return base.replace(hour=9, minute=0)
    if "שעות" in t or "שעה" in t:
        for w in t.split():
            if w.isdigit(): return now + timedelta(hours=int(w))
        return now + timedelta(hours=1)
    if "דקות" in t or "דקה" in t:
        for w in t.split():
            if w.isdigit(): return now + timedelta(minutes=int(w))
        return now + timedelta(minutes=30)
    parts = t.split()
    if len(parts) == 2 and "/" in parts[0] and ":" in parts[1]:
        try:
            day, month = map(int, parts[0].split("/"))
            h, m       = map(int, parts[1].split(":"))
            dt = now.replace(month=month, day=day, hour=h, minute=m, second=0, microsecond=0)
            if dt < now: dt = dt.replace(year=now.year + 1)
            return dt
        except: pass
    if ":" in t:
        try:
            h, m = map(int, t.split(":"))
            dt   = now.replace(hour=h, minute=m, second=0, microsecond=0)
            return dt if dt > now else dt + timedelta(days=1)
        except: pass
    return None

def outcome_kb(lid):
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("✅ נסגר! 🎉",       callback_data=f"out_{lid}_won")],
        [InlineKeyboardButton("⏰ נדחה — קבע זמן", callback_data=f"out_{lid}_snooze"),
         InlineKeyboardButton("🔄 ליד ישן",        callback_data=f"out_{lid}_old")],
        [InlineKeyboardButton("❌ לא רלוונטי",     callback_data=f"out_{lid}_lost")],
    ])

def snooze_kb(lid):
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("בעוד שעה ⏰",   callback_data=f"snz_{lid}_1h"),
         InlineKeyboardButton("מחר 9:00 🌅",   callback_data=f"snz_{lid}_tomorrow")],
        [InlineKeyboardButton("קבע ידנית ✏️",  callback_data=f"snz_{lid}_manual")],
    ])

async def add_start(u: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data.clear()
    await u.message.reply_text("👤 *מה שם הלקוח?*", parse_mode="Markdown")
    return ASK_NAME

async def got_name(u: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data["name"] = u.message.text.strip()
    await u.message.reply_text(
        f"✅ שם: *{ctx.user_data['name']}*\n\n📱 *מה מספר הטלפון?*",
        parse_mode="Markdown")
    return ASK_PHONE

async def got_phone(u: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data["phone"] = u.message.text.strip()
    await u.message.reply_text(
        f"✅ טלפון: *{ctx.user_data['phone']}*\n\n"
        "📅 *מתי השיחה?*\n"
        "_לדוגמא:_\n"
        "`16:30` — היום בשעה 16:30\n"
        "`מחר 10:00` — מחר בעשר\n"
        "`25/06 15:30` — תאריך ספציפי\n"
        "`בעוד שעה` — עוד שעה מעכשיו",
        parse_mode="Markdown")
    return ASK_TIME

async def got_time(u: Update, ctx: ContextTypes.DEFAULT_TYPE):
    call_dt = parse_time(u.message.text.strip())
    if not call_dt:
        await u.message.reply_text(
            "❓ לא הבנתי. נסה:\n`16:30` | `מחר 10:00` | `25/06 15:30`",
            parse_mode="Markdown")
        return ASK_TIME
    name  = ctx.user_data["name"]
    phone = ctx.user_data["phone"]
    lid   = add_lead(name, phone, call_dt)
    eid   = cal_create(lid, name, phone, call_dt)
    if eid:
        update_lead(lid, cal_event_id=eid)
    ctx.user_data.clear()
    await u.message.reply_text(
        f"🎯 *ליד נוסף בהצלחה!*\n\n"
        f"👤 *{name}*\n"
        f"📱 {phone}\n"
        f"📅 שיחה: *{call_dt.strftime('%d/%m/%Y בשעה %H:%M')}*\n\n"
        f"{'📆 אירוע נוצר ב-Google Calendar ✓' if eid else '⚠️ יומן לא זמין'}\n"
        f"⏰ תקבל תזכורת 20 דקות לפני",
        parse_mode="Markdown",
        reply_markup=MAIN_MENU,
    )
    return ConversationHandler.END

async def cancel_add(u: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data.clear()
    await u.message.reply_text("❌ בוטל.", reply_markup=MAIN_MENU)
    return ConversationHandler.END

async def cmd_start(u: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await u.message.reply_text(
        "👋 *שלום! בוט ניהול הלידים שלך* 🍸\n\nבחר פעולה:",
        parse_mode="Markdown",
        reply_markup=MAIN_MENU,
    )

async def show_old_leads(u: Update, ctx: ContextTypes.DEFAULT_TYPE):
    old = get_old_leads()
    if not old:
        await u.message.reply_text("📁 אין לידים ישנים.", reply_markup=MAIN_MENU); return
    await u.message.reply_text(
        f"📁 *לידים ישנים ({len(old)}):*\n_מעוניינים אבל לא נסגרו — לחץ לחזרה_",
        parse_mode="Markdown")
    for ld in old:
        ct       = str(ld.get("call_time", "") or "")
        date_str = ct[:10].replace("-", "/") if ct else "—"
        kb = InlineKeyboardMarkup([
            [InlineKeyboardButton("📞 לחזור אליו",  callback_data=f"recall_{ld['id']}"),
             InlineKeyboardButton("✅ נסגר",         callback_data=f"out_{ld['id']}_won")],
            [InlineKeyboardButton("❌ לא רלוונטי",  callback_data=f"out_{ld['id']}_lost")],
        ])
        await u.message.reply_text(
            f"📁 *{ld['name']}*\n📱 {ld.get('phone','—')}\n📅 שיחה אחרונה: {date_str}",
            reply_markup=kb, parse_mode="Markdown")

async def show_reminders(u: Update, ctx: ContextTypes.DEFAULT_TYPE):
    active = get_active_leads()
    if not active:
        await u.message.reply_text("✅ אין תזכורות פעילות!", reply_markup=MAIN_MENU); return
    lines = [f"⏰ *תזכורות פעילות ({len(active)}):*\n"]
    for ld in active:
        ct = str(ld.get("call_time") or ld.get("follow_up_time") or "")
        t  = ""
        if ct:
            try: t = f" | ⏰ {datetime.strptime(ct[:19], '%Y-%m-%dT%H:%M:%S').strftime('%d/%m %H:%M')}"
            except: pass
        lines.append(f"{STATUS_EMOJI.get(ld.get('status'),'❓')} *{ld['name']}* — 📱{ld.get('phone','—')}{t}")
    await u.message.reply_text("\n".join(lines), parse_mode="Markdown", reply_markup=MAIN_MENU)

async def show_summary(u: Update, ctx: ContextTypes.DEFAULT_TYPE):
    now     = datetime.now(TZ)
    weekly  = get_stats(now - timedelta(days=7))
    monthly = get_stats(now.replace(day=1, hour=0, minute=0, second=0))
    rate    = lambda s: f"{s['won_count']/s['total']*100:.0f}%" if s["total"] else "—"
    await u.message.reply_text(
        "📊 *סיכום ביצועים*\n\n"
        f"📅 *שבוע אחרון:*\n"
        f"  לידים: {weekly['total']} | ✅ נסגרו: {weekly['won_count']} | ❌ אבדו: {weekly['lost']}\n"
        f"  💰 הכנסות: ₪{weekly['won_amount']:,.0f} | המרה: {rate(weekly)}\n\n"
        f"📆 *החודש הנוכחי:*\n"
        f"  לידים: {monthly['total']} | ✅ נסגרו: {monthly['won_count']} | ❌ אבדו: {monthly['lost']}\n"
        f"  💰 הכנסות: ₪{monthly['won_amount']:,.0f} | המרה: {rate(monthly)}\n\n"
        f"🟢 פעילים: {len(get_active_leads())} | 📁 ישנים: {len(get_old_leads())}",
        parse_mode="Markdown",
        reply_markup=MAIN_MENU,
    )

async def handle_text(u: Update, ctx: ContextTypes.DEFAULT_TYPE):
    text  = u.message.text.strip()
    state = ctx.user_data.get("state")
    if state == "ask_amount":
        lid = ctx.user_data.get("lid")
        try:
            amount = float(text.replace("₪", "").replace(",", ""))
            ld     = get_lead(lid)
            cal_delete(ld.get("cal_event_id"))
            update_lead(lid, status=ST_WON, sale_amount=amount, cal_event_id="")
            ctx.user_data.clear()
            await u.message.reply_text(
                f"🎉 *כל הכבוד!*\n*{ld['name']}* — עסקה נסגרה\n💰 *₪{amount:,.0f}*",
                parse_mode="Markdown", reply_markup=MAIN_MENU)
        except ValueError:
            await u.message.reply_text("❓ הכנס סכום בשקלים (לדוגמא: `3500`)", parse_mode="Markdown")
        return
    if state in ("ask_snooze", "ask_recall"):
        lid = ctx.user_data.get("lid")
        dt  = parse_time(text)
        if dt:
            ld = get_lead(lid)
            if state == "ask_recall":
                eid = cal_create(lid, ld["name"], ld.get("phone", ""), dt)
                update_lead(lid, status=ST_NEW, call_time=dt.strftime("%Y-%m-%dT%H:%M:%S"),
                            cal_event_id=eid or "")
            else:
                cal_update(ld.get("cal_event_id"), dt)
                update_lead(lid, status=ST_FOLLOW_UP,
                            follow_up_time=dt.strftime("%Y-%m-%dT%H:%M:%S"))
            ctx.user_data.clear()
            await u.message.reply_text(
                f"⏰ תזכורת נקבעה ל-*{dt.strftime('%d/%m/%Y %H:%M')}* ✓",
                parse_mode="Markdown", reply_markup=MAIN_MENU)
        else:
            await u.message.reply_text(
                "❓ נסה: `16:30` | `מחר 10:00` | `25/06 15:00`", parse_mode="Markdown")
        return
    if "לידים ישנים" in text:
        await show_old_leads(u, ctx)
    elif "סיכום" in text:
        await show_summary(u, ctx)
    elif "תזכורות" in text:
        await show_reminders(u, ctx)
    else:
        await u.message.reply_text(
            "💬 לחץ על אחת האפשרויות בתפריט 👇", reply_markup=MAIN_MENU)

async def callback_handler(u: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = u.callback_query
    await q.answer()
    data = q.data
    if data.startswith("out_"):
        _, lid, action = data.split("_", 2)
        ld = get_lead(lid)
        if not ld:
            await q.edit_message_text("❌ ליד לא נמצא"); return
        if action == "won":
            ctx.user_data.update({"state": "ask_amount", "lid": lid})
            await q.edit_message_text(
                f"🎉 *כמה שילם {ld['name']}?*\n_(הכנס סכום בשקלים)_", parse_mode="Markdown")
        elif action == "lost":
            cal_delete(ld.get("cal_event_id"))
            update_lead(lid, status=ST_LOST, cal_event_id="")
            await q.edit_message_text(f"❌ *{ld['name']}* סומן כלא רלוונטי.", parse_mode="Markdown")
        elif action == "snooze":
            await q.edit_message_text(
                f"⏰ *{ld['name']}* — מתי לחזור אליו?",
                reply_markup=snooze_kb(lid), parse_mode="Markdown")
        elif action == "old":
            update_lead(lid, status=ST_OLD)
            await q.edit_message_text(
                f"📁 *{ld['name']}* הועבר ללידים ישנים.\nתוכל לחזור אליו מהתפריט.",
                parse_mode="Markdown")
    elif data.startswith("snz_"):
        parts  = data.split("_")
        lid    = parts[1]
        option = "_".join(parts[2:])
        ld     = get_lead(lid)
        now    = datetime.now(TZ)
        if option == "1h":
            new_dt = now + timedelta(hours=1)
        elif option == "tomorrow":
            new_dt = (now + timedelta(days=1)).replace(hour=9, minute=0, second=0, microsecond=0)
        elif option == "manual":
            ctx.user_data.update({"state": "ask_snooze", "lid": lid})
            await q.edit_message_text(
                f"⏰ מתי לחזור אל *{ld['name']}*?\n`16:30` | `מחר 10:00` | `בעוד שעה`",
                parse_mode="Markdown")
            return
        else:
            return
        cal_update(ld.get("cal_event_id"), new_dt)
        update_lead(lid, status=ST_FOLLOW_UP,
                    follow_up_time=new_dt.strftime("%Y-%m-%dT%H:%M:%S"))
        await q.edit_message_text(
            f"⏰ *{ld['name']}* — אחזור ב-*{new_dt.strftime('%d/%m %H:%M')}* 👍",
            parse_mode="Markdown")
    elif data.startswith("recall_"):
        lid = data.replace("recall_", "")
        ld  = get_lead(lid)
        ctx.user_data.update({"state": "ask_recall", "lid": lid})
        await q.edit_message_text(
            f"📞 מתי השיחה עם *{ld['name']}*?\n`16:30` | `מחר 10:00` | `25/06 15:00`",
            parse_mode="Markdown")

async def job_check_reminders(app):
    for ld in get_leads_pre_reminder():
        try:
            dt = datetime.strptime(str(ld["call_time"])[:19], "%Y-%m-%dT%H:%M:%S").replace(tzinfo=TZ)
            await app.bot.send_message(
                chat_id=OWNER_CHAT_ID,
                text=f"⏰ *שיחה בעוד ~20 דקות!*\n\n"
                     f"👤 *{ld['name']}*\n"
                     f"📱 {ld.get('phone', '—')}\n"
                     f"🕐 שיחה בשעה *{dt.strftime('%H:%M')}*",
                parse_mode="Markdown")
            update_lead(ld["id"], status=ST_PRE_REMINDED)
        except Exception as e:
            logger.error(f"pre_reminder error [{ld.get('id')}]: {e}")
    for ld in get_leads_post_call():
        try:
            await app.bot.send_message(
                chat_id=OWNER_CHAT_ID,
                text=f"📞 *מה קרה עם {ld['name']}?*\n📱 {ld.get('phone', '—')}",
                reply_markup=outcome_kb(ld["id"]),
                parse_mode="Markdown")
            update_lead(ld["id"], status=ST_CALLED)
        except Exception as e:
            logger.error(f"post_call error [{ld.get('id')}]: {e}")
    for ld in get_leads_followup_due():
        try:
            await app.bot.send_message(
                chat_id=OWNER_CHAT_ID,
                text=f"🔄 *זמן לחזור אל {ld['name']}!*\n📱 {ld.get('phone', '—')}",
                reply_markup=outcome_kb(ld["id"]),
                parse_mode="Markdown")
            update_lead(ld["id"], status=ST_CALLED)
        except Exception as e:
            logger.error(f"followup_due error [{ld.get('id')}]: {e}")

async def job_morning_briefing(app):
    active = get_active_leads()
    old    = get_old_leads()
    if not active and not old: return
    now    = datetime.now(TZ)
    lines  = [f"☀️ *בוקר טוב! {now.strftime('%d/%m/%Y')}*\n"]
    if active:
        lines.append(f"📋 *שיחות ותזכורות ({len(active)}):*")
        for ld in active[:10]:
            ct = str(ld.get("call_time") or ld.get("follow_up_time") or "")
            t  = ""
            if ct:
                try: t = f" ⏰{datetime.strptime(ct[:19], '%Y-%m-%dT%H:%M:%S').strftime('%H:%M')}"
                except: pass
            lines.append(f"{STATUS_EMOJI.get(ld.get('status'), '❓')} *{ld['name']}*{t}")
    if old:
        lines.append(f"\n📁 *{len(old)} לידים ישנים* — לחץ 'לידים ישנים'")
    await app.bot.send_message(chat_id=OWNER_CHAT_ID, text="\n".join(lines), parse_mode="Markdown")

async def job_weekly_report(app):
    s    = get_stats(datetime.now(TZ) - timedelta(days=7))
    rate = f"{s['won_count']/s['total']*100:.0f}%" if s["total"] else "—"
    await app.bot.send_message(
        chat_id=OWNER_CHAT_ID,
        text=f"📊 *דוח שבועי*\n\n"
             f"לידים: {s['total']} | ✅ נסגרו: {s['won_count']} | ❌ אבדו: {s['lost']}\n"
             f"💰 הכנסות: ₪{s['won_amount']:,.0f} | המרה: {rate}\n\nשבוע מוצלח! 🍸",
        parse_mode="Markdown")

async def job_monthly_report(app):
    now  = datetime.now(TZ)
    s    = get_stats(now.replace(day=1, hour=0, minute=0, second=0))
    rate = f"{s['won_count']/s['total']*100:.0f}%" if s["total"] else "—"
    await app.bot.send_message(
        chat_id=OWNER_CHAT_ID,
        text=f"📆 *דוח חודשי — {now.strftime('%m/%Y')}*\n\n"
             f"לידים: {s['total']} | ✅ נסגרו: {s['won_count']} | ❌ אבדו: {s['lost']}\n"
             f"💰 הכנסות: ₪{s['won_amount']:,.0f} | המרה: {rate}\n\nחודש מוצלח! 💪",
        parse_mode="Markdown")

def main():
    if not BOT_TOKEN:      raise ValueError("BOT_TOKEN לא הוגדר!")
    if not OWNER_CHAT_ID:  raise ValueError("OWNER_CHAT_ID לא הוגדר!")
    init_google()

    async def post_init(application: Application) -> None:
        s = AsyncIOScheduler(timezone="Asia/Jerusalem")
        s.add_job(job_check_reminders,  "interval", minutes=1,                  args=[application])
        s.add_job(job_morning_briefing, "cron",     hour=9,  minute=0,          args=[application])
        s.add_job(job_weekly_report,    "cron",     day_of_week="sun", hour=10, args=[application])
        s.add_job(job_monthly_report,   "cron",     day=1,   hour=10,           args=[application])
        s.start()
        logger.info("Scheduler ✓")

    app = Application.builder().token(BOT_TOKEN).post_init(post_init).build()

    conv = ConversationHandler(
        entry_points=[MessageHandler(filters.Regex("^➕ הוספת ליד$"), add_start)],
        states={
            ASK_NAME:  [MessageHandler(filters.TEXT & ~filters.COMMAND, got_name)],
            ASK_PHONE: [MessageHandler(filters.TEXT & ~filters.COMMAND, got_phone)],
            ASK_TIME:  [MessageHandler(filters.TEXT & ~filters.COMMAND, got_time)],
        },
        fallbacks=[CommandHandler("cancel", cancel_add)],
        allow_reentry=True,
    )

    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(conv)
    app.add_handler(CallbackQueryHandler(callback_handler))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))

    class _H(BaseHTTPRequestHandler):
        def do_GET(self):
            self.send_response(200); self.end_headers(); self.wfile.write(b"OK")
        def log_message(self, *a): pass

    port = int(os.getenv("PORT", 8080))
    threading.Thread(
        target=lambda: HTTPServer(("0.0.0.0", port), _H).serve_forever(),
        daemon=True).start()
    logger.info(f"Health server ✓ :{port}")

    logger.info("בוט עולה 🚀")
    app.run_polling(drop_pending_updates=True)

if __name__ == "__main__":
    main()
