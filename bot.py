"""
Клуб настольного тенниса 🏓 — бот + Mini App в одном процессе.

Переменные окружения:
  BOT_TOKEN        токен от @BotFather                       (Bothost передаёт сам)
  ADMIN_IDS        Telegram ID руководителей через запятую
  PORT             порт веб-сервера                          (Bothost передаёт сам)
  DOMAIN           домен бота на Bothost                     (передаётся при включённом домене)
  WEBAPP_URL       можно задать вручную вместо DOMAIN, например https://club.example.com/
  AUTOMAT_TARGET   посещений для автомата (по умолчанию 3)
  SEMESTER_START   с какой даты считать посещения, ГГГГ-ММ-ДД (пусто = за всё время)
  QR_TTL           сколько секунд живёт один QR (по умолчанию 20)
  DB_PATH          путь к базе (по умолчанию attendance.db)
  BACKUP_CHAT_ID   ID закрытого канала для бэкапов базы — спасает данные на бесплатном тарифе
"""
import asyncio
import base64
import hashlib
import hmac
import html
import io
import json
import logging
import os
import re
import sqlite3
import time
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from urllib.parse import parse_qs, parse_qsl, urlparse

import qrcode
from aiohttp import web
from aiogram import Bot, Dispatcher, F, Router
from aiogram.client.default import DefaultBotProperties
from aiogram.enums import ParseMode
from aiogram.filters import Command, CommandObject, CommandStart
from aiogram.types import (BufferedInputFile, CallbackQuery, InlineKeyboardButton,
                           InlineKeyboardMarkup, MenuButtonWebApp, Message, WebAppInfo)
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ─────────────────────────── НАСТРОЙКИ ───────────────────────────
BOT_TOKEN = os.getenv("BOT_TOKEN", "")
ADMIN_IDS = {int(x) for x in os.getenv("ADMIN_IDS", "").replace(" ", "").split(",") if x}
PORT = int(os.getenv("PORT", "8080"))
_domain = re.sub(r"^https?://", "", os.getenv("DOMAIN", "").strip()).strip("/")
WEBAPP_URL = os.getenv("WEBAPP_URL") or (f"https://{_domain}/" if _domain else "")
AUTOMAT_TARGET = int(os.getenv("AUTOMAT_TARGET", "3"))
SEMESTER_START = os.getenv("SEMESTER_START", "")
QR_TTL = int(os.getenv("QR_TTL", "20"))
DB_PATH = os.getenv("DB_PATH", "attendance.db")
BACKUP_CHAT_ID = os.getenv("BACKUP_CHAT_ID", "").strip()
BACKUP_EVERY = 15

# Согласие на обработку персональных данных. Меняешь текст — подними версию,
# и всех попросит принять согласие заново.
CONSENT_VERSION = "2026-09-26b"
CONTACT = "@armaning"
CONSENT_TEXT = f"""Регистрируясь в клубе настольного тенниса ГАУГН, я даю согласие руководителю клуба (Telegram: {CONTACT}), далее — оператор, на обработку моих персональных данных.

Какие данные: фамилия, имя и отчество, курс и факультет, Telegram ID и имя пользователя в Telegram, даты и время посещения тренировок.

Зачем: учёт посещаемости клуба и передача сведений о моих посещениях в ГАУГН (преподавателю физической культуры) для зачёта.

Что с ними делают: сбор, запись, хранение, использование, передача университету, удаление. Обработка автоматизированная — через Telegram-бота клуба. Данные хранятся на сервере хостинга бота и в резервной копии в закрытом Telegram-канале оператора.

Срок: до окончания текущего учебного года, после чего данные удаляются.

Отзыв: согласие можно отозвать в любой момент, написав {CONTACT}. Данные будут удалены, учёт посещений прекратится.

Согласие действует с момента нажатия «Даю согласие»."""  # секунд между проверками, есть ли что бэкапить
WEB_DIR = Path(__file__).resolve().parent / "webapp"

try:
    from zoneinfo import ZoneInfo
    TZ = ZoneInfo(os.getenv("TZ_NAME", "Asia/Yerevan"))
except Exception:
    TZ = timezone(timedelta(hours=4))

SECRET = hashlib.sha256(("qr-secret:" + BOT_TOKEN).encode()).digest()
BOT: Bot | None = None
BOT_USERNAME = ""
WEEKDAYS = ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс"]


def now() -> datetime:
    return datetime.now(TZ)


def now_iso() -> str:
    return now().isoformat(timespec="seconds")


def is_admin(uid: int) -> bool:
    return uid in ADMIN_IDS


def esc(s) -> str:
    return html.escape(str(s or ""))


def since() -> str:
    return SEMESTER_START or "0"


# ─────────────────────────── БАЗА ───────────────────────────
def pick_db_path(path: str) -> str:
    """Проверяем, что в папку можно писать; если нет — берём папку рядом с bot.py."""
    try:
        folder = os.path.dirname(os.path.abspath(path))
        os.makedirs(folder, exist_ok=True)
        probe = os.path.join(folder, ".write_test")
        with open(probe, "w") as f:
            f.write("ok")
        os.remove(probe)
        return os.path.abspath(path)
    except OSError as e:
        fallback = os.path.join(os.path.dirname(os.path.abspath(__file__)), "attendance.db")
        print(f"[DB] ⚠️ Не могу писать в {path!r} ({e}). Использую {fallback}", flush=True)
        return fallback


DB_FILE = pick_db_path(DB_PATH)
SCHEMA = """
CREATE TABLE IF NOT EXISTS users(
    tg_id INTEGER PRIMARY KEY, full_name TEXT NOT NULL, grp TEXT NOT NULL, username TEXT,
    status TEXT NOT NULL DEFAULT 'pending', created_at TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS sessions(
    id INTEGER PRIMARY KEY AUTOINCREMENT, started_at TEXT NOT NULL, ended_at TEXT,
    qr_chat_id INTEGER, qr_msg_id INTEGER);
CREATE TABLE IF NOT EXISTS attendance(
    session_id INTEGER NOT NULL, tg_id INTEGER NOT NULL, marked_at TEXT NOT NULL,
    method TEXT NOT NULL, PRIMARY KEY(session_id, tg_id));
CREATE TABLE IF NOT EXISTS events(
    id INTEGER PRIMARY KEY AUTOINCREMENT, day TEXT NOT NULL, t_from TEXT, t_to TEXT,
    place TEXT, note TEXT, created_at TEXT NOT NULL);
"""
db: sqlite3.Connection | None = None


def init_db():
    global db
    db = sqlite3.connect(DB_FILE, check_same_thread=False)
    db.row_factory = sqlite3.Row
    db.executescript(SCHEMA)
    cols = {r[1] for r in db.execute("PRAGMA table_info(users)")}
    for c in ("consent_at", "consent_ver"):
        if c not in cols:
            db.execute(f"ALTER TABLE users ADD COLUMN {c} TEXT")
    db.commit()
    print(f"[DB] База: {DB_FILE}", flush=True)


# ─────────────────────────── БЭКАП В TELEGRAM ───────────────────────────
# Бот держит копию базы закреплённым файлом в закрытом канале.
# При старте, если локальной базы нет (бесплатный тариф её стёр), скачивает закреп.
# После изменений каждые ~15 секунд отправляет свежую копию и удаляет старую.
DIRTY = False
BACKUP_BLOCKED = False   # если восстановить не вышло — не затираем хороший бэкап пустой базой
LAST_BACKUP_MSG = None


async def restore_backup():
    global LAST_BACKUP_MSG, BACKUP_BLOCKED
    if not BACKUP_CHAT_ID:
        print("[BACKUP] BACKUP_CHAT_ID не задан — бэкапов нет, на бесплатном тарифе данные сотрутся", flush=True)
        return
    if os.path.exists(DB_FILE) and os.path.getsize(DB_FILE) > 0:
        print("[BACKUP] Локальная база на месте — восстановление не нужно", flush=True)
        return
    for attempt in range(1, 4):
        try:
            chat = await BOT.get_chat(int(BACKUP_CHAT_ID))
            pm = chat.pinned_message
            if not pm or not pm.document:
                print("[BACKUP] Бэкапов в канале пока нет — начинаем с чистой базы", flush=True)
                return
            tmp = DB_FILE + ".restore"
            await BOT.download(pm.document.file_id, destination=tmp)
            sqlite3.connect(tmp).execute("PRAGMA integrity_check").fetchone()
            os.replace(tmp, DB_FILE)
            LAST_BACKUP_MSG = pm.message_id
            print(f"[BACKUP] ✅ База восстановлена из бэкапа ({pm.date:%d.%m %H:%M} UTC)", flush=True)
            return
        except Exception as e:
            print(f"[BACKUP] Попытка {attempt}: не удалось восстановить ({e})", flush=True)
            await asyncio.sleep(3)
    BACKUP_BLOCKED = True
    print("[BACKUP] ⚠️ Восстановить не вышло. Бэкапы ОТКЛЮЧЕНЫ до перезапуска, "
          "чтобы не затереть хорошую копию пустой базой. Проверь, что бот — админ канала.", flush=True)


async def backup_now(force=False) -> bool:
    global DIRTY, LAST_BACKUP_MSG
    if not BACKUP_CHAT_ID or BACKUP_BLOCKED or not (DIRTY or force):
        return False
    DIRTY = False
    chat_id = int(BACKUP_CHAT_ID)
    try:
        msg = await BOT.send_document(
            chat_id, BufferedInputFile(db.serialize(), "attendance.db"),
            caption=f"🗄 Бэкап базы клуба · {now():%d.%m.%Y %H:%M:%S}\nНе удаляй и не открепляй.",
            disable_notification=True)
        await BOT.pin_chat_message(chat_id, msg.message_id, disable_notification=True)
        if LAST_BACKUP_MSG and LAST_BACKUP_MSG != msg.message_id:
            try:
                await BOT.delete_message(chat_id, LAST_BACKUP_MSG)
            except Exception:
                pass
        LAST_BACKUP_MSG = msg.message_id
        return True
    except Exception as e:
        DIRTY = True
        logging.warning("[BACKUP] не удалось отправить бэкап: %s", e)
        return False


async def backup_loop():
    while True:
        await asyncio.sleep(BACKUP_EVERY)
        await backup_now()


def q(sql, *a):
    return db.execute(sql, a).fetchall()


def q1(sql, *a):
    return db.execute(sql, a).fetchone()


def ex(sql, *a):
    global DIRTY
    cur = db.execute(sql, a)
    db.commit()
    DIRTY = True
    return cur


def get_user(uid):
    return q1("SELECT * FROM users WHERE tg_id=?", uid)


def has_consent(u) -> bool:
    return bool(u) and u["consent_ver"] == CONSENT_VERSION


def get_session(sid):
    return q1("SELECT * FROM sessions WHERE id=?", sid)


def active_session():
    return q1("SELECT * FROM sessions WHERE ended_at IS NULL ORDER BY id DESC LIMIT 1")


def last_session():
    return q1("SELECT * FROM sessions ORDER BY id DESC LIMIT 1")


def count_marks(sid) -> int:
    return q1("SELECT COUNT(*) c FROM attendance WHERE session_id=?", sid)["c"]


def fmt_dt(iso: str, pattern="%d.%m.%Y %H:%M") -> str:
    return datetime.fromisoformat(iso).strftime(pattern)


def visit_dates(uid) -> list[str]:
    rows = q("""SELECT DISTINCT substr(s.started_at, 1, 10) d FROM attendance a JOIN sessions s ON s.id=a.session_id
                WHERE a.tg_id=? AND s.started_at>=? ORDER BY d""", uid, since())
    return [f"{r['d'][8:10]}.{r['d'][5:7]}" for r in rows]


def marked_today(uid):
    """Отметка этого человека за сегодня в любом занятии (одно посещение в день)."""
    return q1("""SELECT a.marked_at FROM attendance a JOIN sessions s ON s.id=a.session_id
                 WHERE a.tg_id=? AND substr(s.started_at, 1, 10)=? LIMIT 1""", uid, now().date().isoformat())


def visit_counts() -> dict[int, int]:
    rows = q("""SELECT a.tg_id, COUNT(DISTINCT substr(s.started_at, 1, 10)) c FROM attendance a
                JOIN sessions s ON s.id=a.session_id
                WHERE s.started_at>=? GROUP BY a.tg_id""", since())
    return {r["tg_id"]: r["c"] for r in rows}


def upcoming_events() -> list[dict]:
    rows = q("SELECT * FROM events WHERE day>=? ORDER BY day, t_from", now().date().isoformat())
    return [{k: r[k] for k in ("id", "day", "t_from", "t_to", "place", "note")} for r in rows]


def progress_line(uid) -> str:
    n, t = len(visit_dates(uid)), AUTOMAT_TARGET
    bar = "🟠" * min(n, t) + "⚪️" * max(t - n, 0)
    return f"{bar}  🏆 автомат твой" if n >= t else f"{bar}  {n} из {t}, ещё {t - n} до автомата"


# ─────────────────────────── QR-ТОКЕНЫ ───────────────────────────
def current_window() -> int:
    return int(time.time() // QR_TTL)


def sign(sid: int, w: int) -> str:
    return hmac.new(SECRET, f"{sid}:{w}".encode(), hashlib.sha256).hexdigest()[:12]


def qr_url(sid: int) -> str:
    w = current_window()
    return f"https://t.me/{BOT_USERNAME}?start=c{sid}_{w}_{sign(sid, w)}"


def check_payload(p: str):
    try:
        head, w_s, sig = p.split("_")
        if not head.startswith("c"):
            raise ValueError
        sid, w = int(head[1:]), int(w_s)
    except ValueError:
        return "bad", None
    if not hmac.compare_digest(sig, sign(sid, w)):
        return "bad", None
    if w not in (current_window(), current_window() - 1):
        return "expired", sid
    return "ok", sid


def qr_data_url(url: str) -> str:
    img = qrcode.make(url, box_size=10, border=2)
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode()


def do_checkin(uid: int, code: str) -> dict:
    """code — текст из QR (ссылка t.me/...?start=...) или сам payload."""
    code = (code or "").strip()
    if "start=" in code:
        code = (parse_qs(urlparse(code).query).get("start") or [""])[0]
    status, sid = check_payload(code)
    if status == "bad":
        return {"ok": False, "msg": "Это не QR клуба. Отсканируй код с экрана руководителя."}
    if status == "expired":
        return {"ok": False, "msg": f"Код устарел — отсканируй ещё раз, он меняется каждые {QR_TTL} с."}
    user = get_user(uid)
    if not user:
        return {"ok": False, "msg": "Сначала зарегистрируйся в приложении клуба."}
    if user["status"] == "pending":
        return {"ok": False, "msg": "Заявка ещё не подтверждена руководителем."}
    if user["status"] != "approved":
        return {"ok": False, "msg": "Заявка отклонена. Подойди к руководителю клуба."}
    if not has_consent(user):
        return {"ok": False, "msg": "Сначала прими согласие на обработку данных в приложении клуба."}
    s = get_session(sid)
    if not s or s["ended_at"]:
        return {"ok": False, "msg": "Это занятие уже закончилось."}
    already = q1("SELECT marked_at FROM attendance WHERE session_id=? AND tg_id=?", sid, uid) or marked_today(uid)
    if already:
        return {"ok": False, "msg": f"Ты уже отмечен(а) сегодня в {already['marked_at'][11:16]}."}
    t = now_iso()
    ex("INSERT INTO attendance(session_id, tg_id, marked_at, method) VALUES(?,?,?, 'qr')", sid, uid, t)
    n = len(visit_dates(uid))
    msg = f"Отмечено в {t[11:16]}. "
    msg += "Это посещение закрыло норму — автомат твой! 🏆" if n == AUTOMAT_TARGET else (
        f"Ещё {AUTOMAT_TARGET - n} до автомата." if n < AUTOMAT_TARGET else "Хорошей игры 🏓")
    return {"ok": True, "msg": msg, "visits": n}


# ─────────────────────────── EXCEL ───────────────────────────
def build_excel() -> bytes:
    sessions = q("SELECT * FROM sessions WHERE started_at>=? ORDER BY id", since())
    users = q("SELECT * FROM users WHERE status='approved' ORDER BY grp, full_name")
    marks = {(r["session_id"], r["tg_id"]) for r in q("SELECT session_id, tg_id FROM attendance")}
    labels, seen = [], {}
    for s in sessions:
        d = fmt_dt(s["started_at"], "%d.%m.%y")
        seen[d] = seen.get(d, 0) + 1
        labels.append(d if seen[d] == 1 else f"{d} ({seen[d]})")

    wb = Workbook()
    ws = wb.active
    ws.title = "Посещаемость"
    ws.append(["ФИО", "Курс, факультет", *labels, "Всего", "Автомат", "Согласие на ПДн"])
    for u in users:
        row, days = [u["full_name"], u["grp"]], set()
        for s in sessions:
            hit = (s["id"], u["tg_id"]) in marks
            if hit:
                days.add(s["started_at"][:10])
            row.append("+" if hit else "")
        total = len(days)
        row += [total, "✅" if total >= AUTOMAT_TARGET else f"ещё {AUTOMAT_TARGET - total}",
                fmt_dt(u["consent_at"], "%d.%m.%Y %H:%M") if has_consent(u) else "нет"]
        ws.append(row)
    hf, hfill = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor="2F5597")
    center = Alignment(horizontal="center", vertical="center")
    for c in ws[1]:
        c.font, c.fill, c.alignment = hf, hfill, center
    ws.column_dimensions["A"].width, ws.column_dimensions["B"].width = 36, 24
    for col in range(3, len(labels) + 6):
        ws.column_dimensions[get_column_letter(col)].width = 17 if col == len(labels) + 5 else 11
        for r in range(2, ws.max_row + 1):
            ws.cell(row=r, column=col).alignment = center
    ws.freeze_panes = "C2"

    log = wb.create_sheet("Журнал")
    log.append(["Дата", "Время", "ФИО", "Курс, факультет", "Способ"])
    for r in q("""SELECT s.started_at, a.marked_at, a.method, u.full_name, u.grp FROM attendance a
                  JOIN sessions s ON s.id=a.session_id JOIN users u ON u.tg_id=a.tg_id
                  WHERE s.started_at>=? ORDER BY a.marked_at""", since()):
        log.append([fmt_dt(r["started_at"], "%d.%m.%Y"), r["marked_at"][11:16], r["full_name"], r["grp"],
                    "QR" if r["method"] == "qr" else "вручную"])
    for c in log[1]:
        c.font, c.fill = hf, hfill
    for col, w in zip("ABCDE", (12, 8, 36, 24, 10)):
        log.column_dimensions[col].width = w
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ─────────────────────────── УВЕДОМЛЕНИЯ ───────────────────────────
def app_kb(text="🏓 Открыть клуб"):
    if not WEBAPP_URL:
        return None
    return InlineKeyboardMarkup(inline_keyboard=[[
        InlineKeyboardButton(text=text, web_app=WebAppInfo(url=WEBAPP_URL))]])


async def safe_send(uid, text, **kw):
    try:
        await BOT.send_message(uid, text, **kw)
    except Exception as e:
        logging.info("send %s: %s", uid, e)


async def notify_admins_new(uid):
    u = get_user(uid)
    kb = InlineKeyboardMarkup(inline_keyboard=[[
        InlineKeyboardButton(text="✅ Принять", callback_data=f"ap:{uid}"),
        InlineKeyboardButton(text="🚫 Отклонить", callback_data=f"rj:{uid}")]])
    for a in ADMIN_IDS:
        await safe_send(a, f"🆕 Заявка в клуб: <b>{esc(u['full_name'])}</b> ({esc(u['grp'])})"
                        + (f" @{esc(u['username'])}" if u["username"] else ""), reply_markup=kb)


async def notify_decision(uid):
    u = get_user(uid)
    if not u:
        return
    if u["status"] == "approved":
        await safe_send(uid, "✅ Ты в клубе! Отмечайся по QR на тренировках и следи за прогрессом до автомата.",
                        reply_markup=app_kb())
    elif u["status"] == "rejected":
        await safe_send(uid, "🚫 Заявка в клуб отклонена. Если это ошибка — подойди к руководителю.")


def event_line(e) -> str:
    d = date.fromisoformat(e["day"])
    t = e["t_from"] or ""
    if e["t_from"] and e["t_to"]:
        t = f"{e['t_from']}–{e['t_to']}"
    parts = [f"{WEEKDAYS[d.weekday()]}, {d:%d.%m}", t, e["place"] or ""]
    return " · ".join(p for p in parts if p)


async def broadcast_event(e):
    text = f"📅 <b>Новая встреча клуба</b>\n{esc(event_line(e))}"
    if e["note"]:
        text += f"\n\n{esc(e['note'])}"
    for u in q("SELECT tg_id FROM users WHERE status='approved'"):
        await safe_send(u["tg_id"], text, reply_markup=app_kb("📅 Расписание"))
        await asyncio.sleep(0.05)


# ─────────────────────────── API ДЛЯ MINI APP ───────────────────────────
def check_init_data(init_data: str):
    """Проверка подписи Telegram — без неё никто не выдаст себя за другого."""
    try:
        pairs = dict(parse_qsl(init_data, keep_blank_values=True, strict_parsing=True))
    except ValueError:
        return None
    h = pairs.pop("hash", None)
    if not h:
        return None
    dcs = "\n".join(f"{k}={v}" for k, v in sorted(pairs.items()))
    key = hmac.new(b"WebAppData", BOT_TOKEN.encode(), hashlib.sha256).digest()
    if not hmac.compare_digest(hmac.new(key, dcs.encode(), hashlib.sha256).hexdigest(), h):
        return None
    if time.time() - int(pairs.get("auth_date", "0")) > 86400:
        return None
    try:
        return json.loads(pairs["user"])
    except (KeyError, ValueError):
        return None


def jerr(msg, status=400):
    return web.json_response({"error": msg}, status=status)


@web.middleware
async def api_mw(request, handler):
    if request.path.startswith("/api/"):
        auth = request.headers.get("Authorization", "")
        user = check_init_data(auth[4:]) if auth.startswith("tma ") else None
        if not user:
            return jerr("Открой приложение из Telegram", 401)
        request["uid"] = int(user["id"])
        request["tg"] = user
        if request.path.startswith("/api/admin/") and not is_admin(request["uid"]):
            return jerr("Только для руководителя", 403)
    try:
        return await handler(request)
    except web.HTTPException:
        raise
    except Exception:
        logging.exception("API error")
        return jerr("Что-то сломалось на сервере, попробуй ещё раз", 500)


routes = web.RouteTableDef()


async def body(request) -> dict:
    try:
        data = await request.json()
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


@routes.get("/")
async def index(_):
    return web.FileResponse(WEB_DIR / "index.html", headers={"Cache-Control": "no-store"})


@routes.get("/api/me")
async def api_me(request):
    uid = request["uid"]
    u = get_user(uid)
    return web.json_response({
        "admin": is_admin(uid), "target": AUTOMAT_TARGET, "ttl": QR_TTL,
        "user": {"name": u["full_name"], "grp": u["grp"], "status": u["status"], "consent": has_consent(u)} if u else None,
        "consent_text": CONSENT_TEXT, "consent_version": CONSENT_VERSION, "contact": CONTACT,
        "visits": visit_dates(uid) if u else [],
        "schedule": upcoming_events(),
    })


@routes.post("/api/register")
async def api_register(request):
    uid, d = request["uid"], await body(request)
    name = " ".join(str(d.get("name", "")).split())
    grp = " ".join(str(d.get("grp", "")).split())
    if len(name.split()) < 2 or len(name) > 80:
        return jerr("Напиши фамилию и имя через пробел")
    if not grp or len(grp) > 60:
        return jerr("Укажи курс и факультет, например: 1 курс, Юридический")
    if d.get("consent") is not True:
        return jerr("Без согласия на обработку данных зарегистрироваться нельзя")
    u = get_user(uid)
    if u and u["status"] != "pending":
        return jerr("Ты уже зарегистрирован(а)" if u["status"] == "approved"
                    else "Заявка отклонена — подойди к руководителю")
    status = "approved" if is_admin(uid) else "pending"
    ex("""INSERT OR REPLACE INTO users(tg_id, full_name, grp, username, status, created_at, consent_at, consent_ver)
          VALUES(?,?,?,?,?,?,?,?)""",
       uid, name, grp, request["tg"].get("username"), status, now_iso(), now_iso(), CONSENT_VERSION)
    if status == "pending":
        asyncio.create_task(notify_admins_new(uid))
    return web.json_response({"ok": True})


@routes.post("/api/consent")
async def api_consent(request):
    if (await body(request)).get("consent") is not True:
        return jerr("Нужно отметить галочку согласия")
    if not get_user(request["uid"]):
        return jerr("Сначала зарегистрируйся")
    ex("UPDATE users SET consent_at=?, consent_ver=? WHERE tg_id=?", now_iso(), CONSENT_VERSION, request["uid"])
    return web.json_response({"ok": True})


@routes.post("/api/checkin")
async def api_checkin(request):
    return web.json_response(do_checkin(request["uid"], str((await body(request)).get("code", ""))))


# ── руководитель: занятие ──
@routes.get("/api/admin/session")
async def adm_session(request):
    s = active_session()
    if not s:
        last = last_session()
        return web.json_response({"active": False, "last": last["id"] if last else None})
    w = current_window()
    resp = {"active": True, "sid": s["id"], "started": s["started_at"][11:16], "count": count_marks(s["id"]),
            "window": w, "next_in": round((w + 1) * QR_TTL - time.time(), 2), "ttl": QR_TTL}
    if request.query.get("w") != str(w):
        resp["qr"] = qr_data_url(qr_url(s["id"]))
    return web.json_response(resp)


@routes.post("/api/admin/session/start")
async def adm_start(_):
    s = active_session()
    sid = s["id"] if s else ex("INSERT INTO sessions(started_at) VALUES(?)", now_iso()).lastrowid
    return web.json_response({"ok": True, "sid": sid})


@routes.post("/api/admin/session/stop")
async def adm_stop(_):
    s = active_session()
    if not s:
        return jerr("Занятие не идёт")
    ex("UPDATE sessions SET ended_at=? WHERE id=?", now_iso(), s["id"])
    return web.json_response({"ok": True, "sid": s["id"], "count": count_marks(s["id"])})


@routes.get("/api/admin/marks")
async def adm_marks(request):
    sid = request.query.get("sid")
    s = get_session(int(sid)) if sid and sid.isdigit() else (active_session() or last_session())
    if not s:
        return web.json_response({"sid": None, "items": []})
    rows = q("""SELECT a.tg_id, a.marked_at, a.method, u.full_name, u.grp FROM attendance a
                JOIN users u ON u.tg_id=a.tg_id WHERE a.session_id=? ORDER BY a.marked_at""", s["id"])
    return web.json_response({
        "sid": s["id"], "ended": bool(s["ended_at"]), "date": fmt_dt(s["started_at"], "%d.%m.%Y"),
        "items": [{"id": r["tg_id"], "name": r["full_name"], "grp": r["grp"], "time": r["marked_at"][11:16],
                   "manual": r["method"] == "manual"} for r in rows]})


@routes.post("/api/admin/marks/remove")
async def adm_unmark(request):
    d = await body(request)
    ex("DELETE FROM attendance WHERE session_id=? AND tg_id=?", int(d.get("sid", 0)), int(d.get("id", 0)))
    return web.json_response({"ok": True})


@routes.post("/api/admin/marks/add")
async def adm_mark(request):
    d = await body(request)
    sid, uid = int(d.get("sid", 0)), int(d.get("id", 0))
    u = get_user(uid)
    s = get_session(sid)
    if not s or not u or u["status"] != "approved":
        return jerr("Занятие или участник не найдены")
    if not has_consent(u):
        return jerr("Он ещё не принял согласие на обработку данных — пусть откроет приложение")
    same_day = q1("""SELECT 1 FROM attendance a JOIN sessions x ON x.id=a.session_id
                     WHERE a.tg_id=? AND substr(x.started_at, 1, 10)=?""", uid, s["started_at"][:10])
    if same_day:
        return jerr("Он уже отмечен в этот день")
    ex("INSERT OR IGNORE INTO attendance(session_id, tg_id, marked_at, method) VALUES(?,?,?, 'manual')",
       sid, uid, now_iso())
    return web.json_response({"ok": True})


# ── руководитель: участники ──
@routes.get("/api/admin/users")
async def adm_users(_):
    counts = visit_counts()
    def pack(u):
        return {"id": u["tg_id"], "name": u["full_name"], "grp": u["grp"], "username": u["username"],
                "visits": counts.get(u["tg_id"], 0), "consent": has_consent(u)}
    return web.json_response({
        "pending": [pack(u) for u in q("SELECT * FROM users WHERE status='pending' ORDER BY created_at")],
        "approved": [pack(u) for u in q("SELECT * FROM users WHERE status='approved' ORDER BY full_name")]})


@routes.post("/api/admin/users/action")
async def adm_user_action(request):
    d = await body(request)
    uid, action = int(d.get("id", 0)), d.get("action")
    if not get_user(uid):
        return jerr("Участник не найден")
    if action in ("approve", "reject"):
        ex("UPDATE users SET status=? WHERE tg_id=?", "approved" if action == "approve" else "rejected", uid)
        asyncio.create_task(notify_decision(uid))
    elif action == "rename":
        name = " ".join(str(d.get("name", "")).split())
        grp = " ".join(str(d.get("grp", "")).split())
        if len(name.split()) < 2 or not grp:
            return jerr("Нужны ФИО, курс и факультет")
        ex("UPDATE users SET full_name=?, grp=? WHERE tg_id=?", name, grp, uid)
    elif action == "remove":
        ex("DELETE FROM attendance WHERE tg_id=?", uid)
        ex("DELETE FROM users WHERE tg_id=?", uid)
    else:
        return jerr("Неизвестное действие")
    return web.json_response({"ok": True})


@routes.post("/api/admin/export")
async def adm_export(request):
    await BOT.send_document(request["uid"], BufferedInputFile(build_excel(), f"poseshaemost_{now():%Y-%m-%d}.xlsx"),
                            caption="📊 Лист 1 — таблица для физрука, лист 2 — журнал всех отметок.")
    return web.json_response({"ok": True})


# ── руководитель: расписание ──
TIME_RE = re.compile(r"^([01]\d|2[0-3]):[0-5]\d$")


@routes.post("/api/admin/events")
async def adm_event_add(request):
    d = await body(request)
    try:
        day = date.fromisoformat(str(d.get("day", ""))).isoformat()
    except ValueError:
        return jerr("Выбери дату")
    t_from, t_to = str(d.get("t_from") or ""), str(d.get("t_to") or "")
    if t_from and not TIME_RE.match(t_from) or t_to and not TIME_RE.match(t_to):
        return jerr("Время в формате ЧЧ:ММ")
    place = " ".join(str(d.get("place", "")).split())[:80]
    note = str(d.get("note", "")).strip()[:300]
    eid = ex("INSERT INTO events(day, t_from, t_to, place, note, created_at) VALUES(?,?,?,?,?,?)",
             day, t_from or None, t_to or None, place or None, note or None, now_iso()).lastrowid
    if d.get("notify"):
        asyncio.create_task(broadcast_event(q1("SELECT * FROM events WHERE id=?", eid)))
    return web.json_response({"ok": True, "id": eid})


@routes.post("/api/admin/events/delete")
async def adm_event_del(request):
    ex("DELETE FROM events WHERE id=?", int((await body(request)).get("id", 0)))
    return web.json_response({"ok": True})


def build_web_app() -> web.Application:
    app = web.Application(middlewares=[api_mw])
    app.add_routes(routes)
    return app


# ─────────────────────────── ЧАТ-БОТ ───────────────────────────
router = Router()
IsAdmin = F.from_user.id.in_(ADMIN_IDS)
WELCOME = ("🏓 <b>Клуб настольного тенниса</b>\n\n"
           "Всё внутри приложения: регистрация, отметка по QR, путь к автомату и расписание встреч. "
           "Открой кнопкой ниже или кнопкой «Клуб» слева от поля ввода.")


@router.message(CommandStart())
async def cmd_start(m: Message, command: CommandObject):
    if command.args:  # QR отсканирован обычной камерой телефона
        r = do_checkin(m.from_user.id, command.args)
        text = ("✅ " if r["ok"] else "⚠️ ") + esc(r["msg"])
        if r["ok"]:
            text += "\n\n" + progress_line(m.from_user.id)
        await m.answer(text, reply_markup=app_kb())
        return
    await m.answer(WELCOME if WEBAPP_URL else WELCOME + "\n\n⚠️ Приложение ещё не настроено: включи домен на Bothost.",
                   reply_markup=app_kb())


@router.channel_post(Command("id"))
async def channel_id(m: Message):
    await m.answer(f"Вставь в переменные Bothost:\n<code>BACKUP_CHAT_ID={m.chat.id}</code>\nи перезапусти бота.")


@router.message(IsAdmin, Command("backup"))
async def cmd_backup(m: Message):
    if not BACKUP_CHAT_ID:
        await m.answer("BACKUP_CHAT_ID не задан — бэкапы выключены.")
    elif BACKUP_BLOCKED:
        await m.answer("⚠️ Бэкапы заблокированы: при старте не удалось восстановить базу. Проверь, что бот — админ канала, и перезапусти.")
    else:
        ok = await backup_now(force=True)
        await m.answer("✅ Бэкап отправлен в канал" if ok else "⚠️ Не получилось — бот точно админ канала с правом закреплять?")


@router.message(IsAdmin, Command("excel"))
async def cmd_excel(m: Message):
    await m.answer_document(BufferedInputFile(build_excel(), f"poseshaemost_{now():%Y-%m-%d}.xlsx"))


@router.callback_query(IsAdmin, F.data.startswith(("ap:", "rj:")))
async def cb_decide(c: CallbackQuery):
    action, uid = c.data.split(":")
    uid = int(uid)
    u = get_user(uid)
    if not u:
        await c.answer("Участник уже удалён", show_alert=True)
        return
    ex("UPDATE users SET status=? WHERE tg_id=?", "approved" if action == "ap" else "rejected", uid)
    mark = "✅ Принят" if action == "ap" else "🚫 Отклонён"
    try:
        await c.message.edit_text(f"{mark}: <b>{esc(u['full_name'])}</b> ({esc(u['grp'])})")
    except Exception:
        pass
    await notify_decision(uid)
    await c.answer()


@router.message()
async def fallback(m: Message):
    await m.answer(WELCOME, reply_markup=app_kb())


# ─────────────────────────── ЗАПУСК ───────────────────────────
async def main():
    global BOT, BOT_USERNAME
    logging.basicConfig(level=logging.INFO)
    if not BOT_TOKEN or not ADMIN_IDS:
        raise SystemExit("Укажи BOT_TOKEN и ADMIN_IDS")
    BOT = Bot(BOT_TOKEN, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
    BOT_USERNAME = (await BOT.get_me()).username
    await restore_backup()
    init_db()

    runner = web.AppRunner(build_web_app())
    await runner.setup()
    await web.TCPSite(runner, "0.0.0.0", PORT).start()
    logging.info("Mini App: %s (порт %s)", WEBAPP_URL or "домен не задан", PORT)

    if WEBAPP_URL:
        await BOT.set_chat_menu_button(menu_button=MenuButtonWebApp(text="Клуб", web_app=WebAppInfo(url=WEBAPP_URL)))
    dp = Dispatcher()
    dp.include_router(router)
    await BOT.delete_webhook(drop_pending_updates=False)
    backup_task = asyncio.create_task(backup_loop())
    try:
        await dp.start_polling(BOT)
    finally:
        backup_task.cancel()
        await backup_now()  # финальный бэкап при остановке/перезапуске


if __name__ == "__main__":
    asyncio.run(main())
