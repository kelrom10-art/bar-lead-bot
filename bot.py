#!/usr/bin/env python3
"""
בוט ניהול לידים לבר - Telegram Lead Management Bot
"""

import os
import sqlite3
import logging
import threading
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from http.server import HTTPServer, BaseHTTPRequestHandler

from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    CallbackQueryHandler, ContextTypes, filters,
)
from apscheduler.schedulers.asyncio import AsyncIOScheduler

# ── הגדרות ──────────────────────────────────────────────────────────────────
BOT_TOKEN     = os.getenv("BOT_TOKEN", "")
OWNER_CHAT_ID = int(os.getenv("OWNER_CHAT_ID", "0"))
TZ            = ZoneInfo("Asia/Jerusalem")
DB_PATH       = "leads.db"

logging.basicConfig(format="%(asctime)s | %(levelname)s | %(message)s", level=logging.INFO)
logger = logging.getLogger(__name__)

STATE_REMINDER    = "waiting_reminder"
STATE_SALE_AMOUNT = "waiting_sale"
STATE_LOST_REASON = "waiting_lost"
STATE_SNOOZE_TIME = "waiting_snooze"

STATUS_EMOJI = {"new": "🆕", "contacted": "📞", "proposal": "📋", "won": "✅", "lost": "❌"}
STATUS_HEB   = {"new": "חדש", "contacted": "יצרתי קשר", "proposal": "הצעה נשלחה",
                "won": "נסגר ✅", "lost": "אבוד ❌"}

# ── מסד נתונים ──────────────────────────────────────────────────────────────

def init_db():
    with sqlite3.connect(DB_PATH) as conn:
        conn.executescript("""
            CREATE TABLE IF NOT EXISTS leads (
                id             INTEGER PRIMARY KEY AUTOINCREMENT,
                name           TEXT    NOT NULL,
                phone          TEXT,
                event_type     TEXT,
                notes          TEXT,
                status         TEXT    DEFAULT 'new',
                created_at     TEXT    DEFAULT (strftime('%Y-%m-%dT%H:%M:%S','now','localtime')),
                updated_at     TEXT    DEFAULT (strftime('%Y-%m-%dT%H:%M:%S','now','localtime')),
                reminder_at    TEXT,
                reminder_count INTEGER DEFAULT 0,
                sale_amount    REAL
            );
        """)

def _conn():
    c = sqlite3.connect(DB_PATH)
    c.row_factory = sqlite3.Row
    return c

def _now(): return datetime.now(TZ).strftime("%Y-%m-%dT%H:%M:%S")

def add_lead(name):
    with _conn() as c: return c.execute("INSERT INTO leads (name) VALUES (?)", (name,)).lastrowid

def update_lead(lead_id, **kw):
    kw["updated_at"] = _now()
    sets = ", ".join(f"{k}=?" for k in kw)
    with _conn() as c: c.execute(f"UPDATE leads SET {sets} WHERE id=?", [*kw.values(), lead_id])

def get_lead(lid):
    with _conn() as c: return c.execute("SELECT * FROM leads WHERE id=?", (lid,)).fetchone()

def get_active_leads():
    with _conn() as c: return c.execute(
        "SELECT * FROM leads WHERE status NOT IN ('won','lost') ORDER BY created_at DESC").fetchall()

def get_pending_reminders():
    with _conn() as c: return c.execute(
        "SELECT * FROM leads WHERE reminder_at IS NOT NULL AND reminder_at<=? AND status NOT IN ('won','lost')",
        (_now(),)).fetchall()

def get_old_uncontacted(days=3):
    cutoff = (datetime.now(TZ) - timedelta(days=days)).strftime("%Y-%m-%dT%H:%M:%S")
    with _conn() as c: return c.execute(
        "SELECT * FROM leads WHERE status='new' AND reminder_at IS NULL AND created_at<=?", (cutoff,)).fetchall()

def get_stats(since_dt):
    since = since_dt.strftime("%Y-%m-%dT%H:%M:%S")
    with _conn() as c:
        total = c.execute("SELECT COUNT(*) FROM leads WHERE created_at>=?", (since,)).fetchone()[0]
        won   = c.execute("SELECT COUNT(*),COALESCE(SUM(sale_amount),0) FROM leads WHERE status='won' AND updated_at>=?", (since,)).fetchone()
        lost  = c.execute("SELECT COUNT(*) FROM leads WHERE status='lost' AND updated_at>=?", (since,)).fetchone()[0]
    return {"total": total, "won_count": won[0], "won_amount": won[1], "lost": lost}

# ── עזרי עיצוב ──────────────────────────────────────────────────────────────

def parse_time(text):
    now = datetime.now(TZ)
    t = text.strip()
    if "שעות" in t or "שעה" in t:
        for w in t.split():
            if w.isdigit(): return now + timedelta(hours=int(w))
        return now + timedelta(hours=1)
    if "דקות" in t or "דקה" in t:
        for w in t.split():
            if w.isdigit(): return now + timedelta(minutes=int(w))
        return now + timedelta(minutes=30)
    if "מחר" in t:
        base = (now + timedelta(days=1)).replace(second=0, microsecond=0)
        for w in t.split():
            if ":" in w:
                try: h, m = map(int, w.split(":")); return base.replace(hour=h, minute=m)
                except: pass
        return base.replace(hour=9, minute=0)
    if ":" in t:
        try:
            h, m = map(int, t.split(":"))
            dt = now.replace(hour=h, minute=m, second=0, microsecond=0)
            return dt if dt > now else dt + timedelta(days=1)
        except: pass
    return None

def reminder_kb(lead_id):
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("בעוד שעה ⏰", callback_data=f"rem_{lead_id}_1h"),
         InlineKeyboardButton("בעוד 2 שעות", callback_data=f"rem_{lead_id}_2h")],
        [InlineKeyboardButton("מחר 9:00 🌅", callback_data=f"rem_{lead_id}_tomorrow"),
         InlineKeyboardButton("הגדר ידנית ✏️", callback_data=f"rem_{lead_id}_manual")],
    ])

def followup_kb(lead_id):
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("✅ מכרתי!", callback_data=f"status_{lead_id}_won")],
        [InlineKeyboardButton("📞 דחה שעה", callback_data=f"status_{lead_id}_snooze"),
         InlineKeyboardButton("📞 דחה ידנית", callback_data=f"status_{lead_id}_snooze_manual")],
        [InlineKeyboardButton("❌ לא רלוונטי", callback_data=f"status_{lead_id}_lost")],
    ])

# ── פקודות ──────────────────────────────────────────────────────────────────

HELP_TEXT = (
    "👋 *שלום! אני בוט ניהול הלידים שלך* 🍸\n\n"
    "📌 *הוספת ליד:*\n"
    "`ליד חדש [שם]`  ← לדוגמא: `ליד חדש אבי כהן`\n\n"
    "📋 *פקודות:*\n"
    "/list — כל הלידים הפעילים\n"
    "/old — לידים שלא טופלו +3 ימים\n"
    "/summary — דוח שבועי וחודשי\n"
    "/export — קובץ CSV של כל הלידים\n"
    "/help — תפריט זה"
)

async def cmd_start(u, ctx): await u.message.reply_text(HELP_TEXT, parse_mode="Markdown")
async def cmd_help(u, ctx):  await u.message.reply_text(HELP_TEXT, parse_mode="Markdown")

async def cmd_list(u, ctx):
    leads = get_active_leads()
    if not leads:
        await u.message.reply_text("✅ אין לידים פעילים כרגע!"); return
    lines = [f"📋 *לידים פעילים ({len(leads)}):*\n"]
    for ld in leads:
        emoji = STATUS_EMOJI.get(ld["status"], "❓")
        days  = (datetime.now(TZ) - datetime.strptime(ld["created_at"], "%Y-%m-%dT%H:%M:%S").replace(tzinfo=TZ)).days
        line  = f"{emoji} *{ld['name']}*" + (f" ({days}ד')" if days > 0 else "")
        if ld["reminder_at"]:
            rem  = datetime.strptime(ld["reminder_at"], "%Y-%m-%dT%H:%M:%S")
            line += f" | ⏰ {rem.strftime('%d/%m %H:%M')}"
        lines.append(line)
    await u.message.reply_text("\n".join(lines), parse_mode="Markdown")

async def cmd_old(u, ctx):
    leads = get_old_uncontacted(days=3)
    if not leads:
        await u.message.reply_text("👍 אין לידים ישנים שלא טופלו!"); return
    await u.message.reply_text(f"⚠️ *{len(leads)} לידים שלא טופלו מעל 3 ימים:*", parse_mode="Markdown")
    for ld in leads:
        days = (datetime.now(TZ) - datetime.strptime(ld["created_at"], "%Y-%m-%dT%H:%M:%S").replace(tzinfo=TZ)).days
        kb   = InlineKeyboardMarkup([[
            InlineKeyboardButton("📞 יצרתי קשר", callback_data=f"status_{ld['id']}_contacted"),
            InlineKeyboardButton("❌ אבוד",       callback_data=f"status_{ld['id']}_lost"),
        ]])
        await u.message.reply_text(f"⚠️ *{ld['name']}* — לפני {days} ימים", reply_markup=kb, parse_mode="Markdown")

async def cmd_summary(u, ctx):
    now     = datetime.now(TZ)
    weekly  = get_stats(now - timedelta(days=7))
    monthly = get_stats(now.replace(day=1, hour=0, minute=0, second=0))
    active  = get_active_leads()
    rate    = lambda s: f"{s['won_count']/s['total']*100:.0f}%" if s["total"] else "0%"
    await u.message.reply_text(
        "📊 *דוח ביצועים*\n\n"
        f"📅 *השבוע האחרון:*\n"
        f"  לידים: {weekly['total']} | נסגרו: {weekly['won_count']} | אבדו: {weekly['lost']}\n"
        f"  הכנסות: ₪{weekly['won_amount']:,.0f} | המרה: {rate(weekly)}\n\n"
        f"📆 *החודש הנוכחי:*\n"
        f"  לידים: {monthly['total']} | נסגרו: {monthly['won_count']} | אבדו: {monthly['lost']}\n"
        f"  הכנסות: ₪{monthly['won_amount']:,.0f} | המרה: {rate(monthly)}\n\n"
        f"🟢 *פעילים כרגע: {len(active)}*",
        parse_mode="Markdown",
    )

async def cmd_export(u, ctx):
    import io, csv
    with _conn() as c: leads = c.execute("SELECT * FROM leads ORDER BY created_at DESC").fetchall()
    out = io.StringIO()
    w   = csv.writer(out)
    w.writerow(["ID","שם","טלפון","סוג אירוע","סטטוס","סכום","הערות","נוצר","עודכן"])
    for ld in leads:
        w.writerow([ld["id"], ld["name"], ld["phone"] or "", ld["event_type"] or "",
                    STATUS_HEB.get(ld["status"], ld["status"]),
                    f"₪{ld['sale_amount']:,.0f}" if ld["sale_amount"] else "",
                    ld["notes"] or "", ld["created_at"][:16].replace("T"," "), ld["updated_at"][:16].replace("T"," ")])
    out.seek(0)
    await u.message.reply_document(
        document=out.getvalue().encode("utf-8-sig"),
        filename=f"leads_{datetime.now(TZ).strftime('%Y%m%d')}.csv",
        caption=f"📊 ייצוא לידים — {len(leads)} רשומות",
    )

async def handle_text(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    text  = update.message.text.strip()
    state = ctx.user_data.get("state")
    if text.startswith("ליד חדש") or text.replace(" ","").startswith("לידחדש"):
        parts = text.split(maxsplit=2)
        if len(parts) < 3:
            await update.message.reply_text("❓ אנא ציין שם:\n`ליד חדש [שם]`", parse_mode="Markdown"); return
        name    = parts[2].strip()
        lead_id = add_lead(name)
        ctx.user_data.clear()
        await update.message.reply_text(
            f"✅ *{name}* נוסף!\n\n⏰ מתי לתזכר אותך לחזור אליו?",
            reply_markup=reminder_kb(lead_id), parse_mode="Markdown")
        return
    if state == STATE_REMINDER:
        lead_id = ctx.user_data.get("lead_id")
        dt = parse_time(text)
        if dt and lead_id:
            update_lead(lead_id, reminder_at=dt.strftime("%Y-%m-%dT%H:%M:%S"), status="contacted")
            lead = get_lead(lead_id); ctx.user_data.clear()
            await update.message.reply_text(
                f"⏰ תזכורת נקבעה ל-*{dt.strftime('%d/%m %H:%M')}*\nאחזור אחרי שיחה עם *{lead['name']}* 💪",
                parse_mode="Markdown")
        else:
            await update.message.reply_text("❓ נסה: `14:30` | `בעוד 2 שעות` | `מחר 10:00`", parse_mode="Markdown")
        return
    if state == STATE_SALE_AMOUNT:
        lead_id = ctx.user_data.get("lead_id")
        try:
            amount = float(text.replace("₪","").replace(",","").strip())
            update_lead(lead_id, status="won", sale_amount=amount, reminder_at=None)
            lead = get_lead(lead_id); ctx.user_data.clear()
            await update.message.reply_text(
                f"🎉 *כל הכבוד!* עסקה עם *{lead['name']}* נסגרה ב-*₪{amount:,.0f}* 💰", parse_mode="Markdown")
        except ValueError:
            await update.message.reply_text("❓ הכנס סכום בשקלים (לדוגמא: `3500`)", parse_mode="Markdown")
        return
    if state == STATE_LOST_REASON:
        lead_id = ctx.user_data.get("lead_id")
        update_lead(lead_id, status="lost", notes=text, reminder_at=None)
        lead = get_lead(lead_id); ctx.user_data.clear()
        await update.message.reply_text(f"📝 סיבה נרשמה. *{lead['name']}* סומן כאבוד.", parse_mode="Markdown")
        return
    if state == STATE_SNOOZE_TIME:
        lead_id = ctx.user_data.get("lead_id")
        dt = parse_time(text)
        if dt and lead_id:
            update_lead(lead_id, reminder_at=dt.strftime("%Y-%m-%dT%H:%M:%S"))
            ctx.user_data.clear()
            await update.message.reply_text(f"⏰ תזכורת נדחתה ל-*{dt.strftime('%d/%m %H:%M')}*", parse_mode="Markdown")
        else:
            await update.message.reply_text("❓ נסה: `14:30` | `מחר 10:00`", parse_mode="Markdown")
        return
    await update.message.reply_text("💬 לא הבנתי.\n`ליד חדש [שם]` | /help", parse_mode="Markdown")

async def callback_handler(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q    = update.callback_query
    await q.answer()
    data = q.data
    if data.startswith("rem_"):
        _, lid, option = data.split("_", 2)
        lead_id = int(lid)
        lead    = get_lead(lead_id)
        now     = datetime.now(TZ)
        if option == "1h":         dt = now + timedelta(hours=1)
        elif option == "2h":       dt = now + timedelta(hours=2)
        elif option == "tomorrow": dt = (now + timedelta(days=1)).replace(hour=9, minute=0, second=0, microsecond=0)
        elif option == "manual":
            ctx.user_data.update({"state": STATE_REMINDER, "lead_id": lead_id})
            await q.edit_message_text(
                f"⏰ מתי לתזכר אותך לגבי *{lead['name']}*?\n\nנסה: `14:30` | `בעוד 2 שעות` | `מחר 10:00`",
                parse_mode="Markdown"); return
        else: return
        update_lead(lead_id, reminder_at=dt.strftime("%Y-%m-%dT%H:%M:%S"), status="contacted")
        await q.edit_message_text(
            f"⏰ תזכורת נקבעה ל-*{dt.strftime('%d/%m %H:%M')}*\nאחזור אחרי שיחה עם *{lead['name']}* 💪",
            parse_mode="Markdown"); return
    if data.startswith("status_"):
        parts   = data.split("_")
        lead_id = int(parts[1])
        action  = "_".join(parts[2:])
        lead    = get_lead(lead_id)
        if action == "won":
            ctx.user_data.update({"state": STATE_SALE_AMOUNT, "lead_id": lead_id})
            await q.edit_message_text(f"🎉 כמה שילם *{lead['name']}*?\n_(הכנס סכום בשקלים)_", parse_mode="Markdown")
        elif action == "lost":
            ctx.user_data.update({"state": STATE_LOST_REASON, "lead_id": lead_id})
            await q.edit_message_text(f"😞 מה הסיבה שאבדת את *{lead['name']}*?\n_(כתוב בקצרה)_", parse_mode="Markdown")
        elif action == "snooze":
            dt = datetime.now(TZ) + timedelta(hours=1)
            update_lead(lead_id, reminder_at=dt.strftime("%Y-%m-%dT%H:%M:%S"))
            await q.edit_message_text(f"⏰ תזכורת נדחתה — אחזור ב-*{dt.strftime('%H:%M')}* 👍", parse_mode="Markdown")
        elif action == "snooze_manual":
            ctx.user_data.update({"state": STATE_SNOOZE_TIME, "lead_id": lead_id})
            await q.edit_message_text(f"⏰ מתי לתזכר אותך שוב לגבי *{lead['name']}*?\nנסה: `17:00` | `מחר 10:00`", parse_mode="Markdown")
        elif action == "contacted":
            update_lead(lead_id, status="contacted")
            await q.edit_message_text(f"📞 מתי לתזכר אותך לגבי *{lead['name']}*?",
                reply_markup=reminder_kb(lead_id), parse_mode="Markdown")

async def job_check_reminders(app):
    for ld in get_pending_reminders():
        try:
            await app.bot.send_message(chat_id=OWNER_CHAT_ID,
                text=f"⏰ *תזכורת: {ld['name']}*\n\nהאם הצלחת לסגור עסקה?",
                reply_markup=followup_kb(ld["id"]), parse_mode="Markdown")
            update_lead(ld["id"], reminder_at=None, reminder_count=ld["reminder_count"]+1)
        except Exception as e: logger.error(f"reminder error {ld['id']}: {e}")

async def job_morning_briefing(app):
    active = get_active_leads()
    old    = get_old_uncontacted(days=3)
    if not active and not old: return
    now   = datetime.now(TZ)
    lines = [f"☀️ *בוקר טוב! — {now.strftime('%d/%m/%Y')}*\n"]
    if active:
        lines.append(f"📋 *לידים פעילים: {len(active)}*")
        for ld in active[:8]:
            line = f"{STATUS_EMOJI.get(ld['status'],'❓')} {ld['name']}"
            if ld["reminder_at"]:
                rem = datetime.strptime(ld["reminder_at"], "%Y-%m-%dT%H:%M:%S")
                line += f" ⏰ {rem.strftime('%d/%m %H:%M')}"
            lines.append(line)
        if len(active) > 8: lines.append(f"_...ועוד {len(active)-8}_")
    if old: lines.append(f"\n⚠️ *{len(old)} לידים ישנים — שלח /old*")
    await app.bot.send_message(chat_id=OWNER_CHAT_ID, text="\n".join(lines), parse_mode="Markdown")

async def job_old_leads_alert(app):
    old = get_old_uncontacted(days=3)
    if not old: return
    names = "\n".join(f"• {ld['name']}" for ld in old)
    await app.bot.send_message(chat_id=OWNER_CHAT_ID,
        text=f"⚠️ *{len(old)} לידים לא טופלו מעל 3 ימים:*\n\n{names}\n\nשלח /old לטיפול",
        parse_mode="Markdown")

async def job_weekly_report(app):
    s    = get_stats(datetime.now(TZ) - timedelta(days=7))
    rate = f"{s['won_count']/s['total']*100:.0f}%" if s["total"] else "0%"
    await app.bot.send_message(chat_id=OWNER_CHAT_ID,
        text=(f"📊 *דוח שבועי*\n\n  לידים: {s['total']} | נסגרו: {s['won_count']} | אבדו: {s['lost']}\n"
              f"  הכנסות: ₪{s['won_amount']:,.0f} | המרה: {rate}\n\nשבוע מוצלח! 🍸"),
        parse_mode="Markdown")

async def job_monthly_report(app):
    now  = datetime.now(TZ)
    s    = get_stats(now.replace(day=1, hour=0, minute=0, second=0))
    rate = f"{s['won_count']/s['total']*100:.0f}%" if s["total"] else "0%"
    await app.bot.send_message(chat_id=OWNER_CHAT_ID,
        text=(f"📆 *דוח חודשי — {now.strftime('%m/%Y')}*\n\n  לידים: {s['total']} | נסגרו: {s['won_count']} | אבדו: {s['lost']}\n"
              f"  הכנסות: ₪{s['won_amount']:,.0f} | המרה: {rate}\n\nחודש מוצלח! 💪"),
        parse_mode="Markdown")

def main():
    if not BOT_TOKEN:     raise ValueError("BOT_TOKEN לא הוגדר!")
    if not OWNER_CHAT_ID: raise ValueError("OWNER_CHAT_ID לא הוגדר!")
    init_db(); logger.info("DB ✓")

    app = Application.builder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start",   cmd_start))
    app.add_handler(CommandHandler("help",    cmd_help))
    app.add_handler(CommandHandler("list",    cmd_list))
    app.add_handler(CommandHandler("old",     cmd_old))
    app.add_handler(CommandHandler("summary", cmd_summary))
    app.add_handler(CommandHandler("export",  cmd_export))
    app.add_handler(CallbackQueryHandler(callback_handler))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))

    s = AsyncIOScheduler(timezone="Asia/Jerusalem")
    s.add_job(job_check_reminders,  "interval", minutes=1,                  args=[app])
    s.add_job(job_morning_briefing, "cron",     hour=9,  minute=0,          args=[app])
    s.add_job(job_old_leads_alert,  "cron",     hour=18, minute=0,          args=[app])
    s.add_job(job_weekly_report,    "cron",     day_of_week="sun", hour=10, args=[app])
    s.add_job(job_monthly_report,   "cron",     day=1,   hour=10,           args=[app])
    s.start(); logger.info("Scheduler ✓")

    class _Health(BaseHTTPRequestHandler):
        def do_GET(self):
            self.send_response(200); self.end_headers(); self.wfile.write(b"OK")
        def log_message(self, *a): pass

    port = int(os.getenv("PORT", 8080))
    t = threading.Thread(target=lambda: HTTPServer(("0.0.0.0", port), _Health).serve_forever(), daemon=True)
    t.start(); logger.info(f"Health server ✓ :{port}")

    logger.info("בוט עולה 🚀")
    app.run_polling(drop_pending_updates=True)

if __name__ == "__main__":
    main()
