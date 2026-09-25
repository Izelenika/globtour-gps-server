import json
import sqlite3
import struct
import time
import os
import hashlib
import hmac
import secrets
from datetime import datetime, timezone, timedelta
from zoneinfo import ZoneInfo
import math
from pathlib import Path
from contextlib import asynccontextmanager

from fastapi import FastAPI, Request, HTTPException
from fastapi.responses import HTMLResponse, JSONResponse, RedirectResponse
from starlette.middleware.sessions import SessionMiddleware
from fastapi.templating import Jinja2Templates
import uvicorn

BASE = Path(__file__).resolve().parent
DATA_DIR = Path(os.getenv("RAILWAY_VOLUME_MOUNT_PATH", "/app/data"))
DATA_DIR.mkdir(parents=True, exist_ok=True)

DB = DATA_DIR / "globtour_gps.db"
TCP_HOST = "0.0.0.0"
TCP_PORT = 9000
WEB_HOST = "0.0.0.0"
WEB_PORT = int(os.getenv("PORT", "8000"))

gps_listener = None


@asynccontextmanager
async def lifespan(app: FastAPI):
    # Web process only. The Teltonika TCP listener runs in gps_tcp.py.
    init_db()
    yield


app = FastAPI(title="Globtour GPS Server", lifespan=lifespan)
templates = Jinja2Templates(directory=str(BASE / "templates"))

SESSION_SECRET_FILE = DATA_DIR / ".session_secret"
SESSION_SECRET = os.getenv("SESSION_SECRET")
if not SESSION_SECRET:
    if SESSION_SECRET_FILE.exists():
        SESSION_SECRET = SESSION_SECRET_FILE.read_text(encoding="utf-8").strip()
    else:
        SESSION_SECRET = secrets.token_urlsafe(48)
        SESSION_SECRET_FILE.write_text(SESSION_SECRET, encoding="utf-8")

app.add_middleware(
    SessionMiddleware,
    secret_key=SESSION_SECRET,
    session_cookie="globtour_gps_session",
    max_age=8 * 60 * 60,
    same_site="lax",
    https_only=os.getenv("COOKIE_SECURE","0").strip().lower() in ("1","true","yes"),
)


def db():
    con = sqlite3.connect(DB)
    con.row_factory = sqlite3.Row
    return con


def init_db():
    con = db()
    con.executescript("""
    CREATE TABLE IF NOT EXISTS devices (
        imei TEXT PRIMARY KEY,
        registration TEXT,
        make_model TEXT,
        vehicle_type TEXT,
        production_year INTEGER,
        note TEXT,
        name TEXT,
        created_at TEXT NOT NULL,
        last_seen TEXT
    );

    CREATE TABLE IF NOT EXISTS positions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        imei TEXT NOT NULL,
        ts_utc TEXT NOT NULL,
        latitude REAL,
        longitude REAL,
        altitude REAL,
        angle REAL,
        satellites INTEGER,
        speed_kmh REAL,
        priority INTEGER,
        ignition INTEGER,
        movement INTEGER,
        gsm_signal INTEGER,
        external_voltage REAL,
        total_odometer REAL,
        raw_json TEXT
    );

    CREATE INDEX IF NOT EXISTS idx_positions_imei_ts
    ON positions(imei, ts_utc);
    """)

    # Safe migration for databases created by earlier versions.
    cols = {row[1] for row in con.execute("PRAGMA table_info(devices)").fetchall()}
    migrations = {
        "make_model": "ALTER TABLE devices ADD COLUMN make_model TEXT",
        "vehicle_type": "ALTER TABLE devices ADD COLUMN vehicle_type TEXT",
        "production_year": "ALTER TABLE devices ADD COLUMN production_year INTEGER",
        "note": "ALTER TABLE devices ADD COLUMN note TEXT",
    }
    for col, sql in migrations.items():
        if col not in cols:
            con.execute(sql)

    con.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            full_name TEXT DEFAULT '',
            role TEXT NOT NULL DEFAULT 'user',
            permissions TEXT DEFAULT '',
            active INTEGER NOT NULL DEFAULT 1,
            created_at TEXT NOT NULL
        )
    """)
    con.commit()

    admin_username = (os.getenv("ADMIN_USERNAME") or "admin").strip() or "admin"
    admin_password = os.getenv("ADMIN_PASSWORD") or "PromijeniMe123!"
    if not os.getenv("ADMIN_PASSWORD"):
        print("[AUTH] UPOZORENJE: ADMIN_PASSWORD nije postavljen; koristi se početna lozinka. Promijenite je odmah.", flush=True)
    existing_admin = con.execute("SELECT id FROM users WHERE username=?", (admin_username,)).fetchone()
    if not existing_admin:
        con.execute(
            "INSERT INTO users(username,password_hash,full_name,role,permissions,active,created_at) VALUES(?,?,?,?,?,?,?)",
            (admin_username, hash_password(admin_password), "Globtour administrator", "admin", "*", 1, now_utc()),
        )
    con.commit()
    con.close()


def now_utc():
    return datetime.now(timezone.utc).isoformat()


def hash_password(password: str) -> str:
    salt = secrets.token_bytes(16)
    digest = hashlib.scrypt(password.encode("utf-8"), salt=salt, n=2**14, r=8, p=1)
    return "scrypt$16384$8$1$" + salt.hex() + "$" + digest.hex()


def verify_password(password: str, stored: str) -> bool:
    try:
        scheme, n, r, p, salt_hex, digest_hex = stored.split("$")
        if scheme != "scrypt":
            return False
        digest = hashlib.scrypt(password.encode("utf-8"), salt=bytes.fromhex(salt_hex), n=int(n), r=int(r), p=int(p))
        return hmac.compare_digest(digest.hex(), digest_hex)
    except Exception:
        return False


def get_current_user(request: Request):
    user_id = request.session.get("user_id")
    if not user_id:
        return None
    con = db()
    row = con.execute("SELECT id, username, full_name, role, permissions, active FROM users WHERE id=?", (user_id,)).fetchone()
    con.close()
    if not row or not int(row["active"] or 0):
        request.session.clear()
        return None
    return dict(row)


def require_user(request: Request):
    user = get_current_user(request)
    if not user:
        raise HTTPException(status_code=401, detail="Niste prijavljeni.")
    return user


def require_admin(request: Request):
    user = require_user(request)
    if str(user.get("role") or "").lower() != "admin":
        raise HTTPException(status_code=403, detail="Potrebna je administratorska ovlast.")
    return user


def crc16_ibm(data: bytes) -> int:
    crc = 0
    for b in data:
        crc ^= b
        for _ in range(8):
            if crc & 1:
                crc = (crc >> 1) ^ 0xA001
            else:
                crc >>= 1
    return crc & 0xFFFF


def read_uint(data, off, n):
    return int.from_bytes(data[off:off+n], "big"), off + n


def read_int(data, off, n):
    return int.from_bytes(data[off:off+n], "big", signed=True), off + n


def parse_io(payload: bytes, off: int, codec: int):
    # Codec 8 Extended uses 2-byte counts and IDs.
    # Codec 8 uses 1-byte counts and IDs.
    wide = codec == 0x8E
    count_size = 2 if wide else 1
    id_size = 2 if wide else 1

    io = {}

    total, off = read_uint(payload, off, count_size)

    n1, off = read_uint(payload, off, count_size)
    for _ in range(n1):
        io_id, off = read_uint(payload, off, id_size)
        val, off = read_uint(payload, off, 1)
        io[io_id] = val

    n2, off = read_uint(payload, off, count_size)
    for _ in range(n2):
        io_id, off = read_uint(payload, off, id_size)
        val, off = read_uint(payload, off, 2)
        io[io_id] = val

    n4, off = read_uint(payload, off, count_size)
    for _ in range(n4):
        io_id, off = read_uint(payload, off, id_size)
        val, off = read_uint(payload, off, 4)
        io[io_id] = val

    n8, off = read_uint(payload, off, count_size)
    for _ in range(n8):
        io_id, off = read_uint(payload, off, id_size)
        val, off = read_uint(payload, off, 8)
        io[io_id] = val

    # Codec 8 Extended has variable-length IO values.
    if wide:
        nx, off = read_uint(payload, off, count_size)
        for _ in range(nx):
            io_id, off = read_uint(payload, off, id_size)
            length, off = read_uint(payload, off, 2)
            io[io_id] = payload[off:off+length].hex()
            off += length

    return io, off, total


def parse_avl_record(payload: bytes, off: int, codec: int):
    if off + 24 > len(payload):
        raise ValueError("Incomplete AVL record")

    ts_ms = int.from_bytes(payload[off:off+8], "big")
    off += 8
    priority = payload[off]
    off += 1

    lon_raw = int.from_bytes(payload[off:off+4], "big", signed=True)
    off += 4
    lat_raw = int.from_bytes(payload[off:off+4], "big", signed=True)
    off += 4
    altitude = int.from_bytes(payload[off:off+2], "big", signed=True)
    off += 2
    angle = int.from_bytes(payload[off:off+2], "big")
    off += 2
    satellites = payload[off]
    off += 1
    speed = int.from_bytes(payload[off:off+2], "big")
    off += 2

    # Event IO ID is 1 byte in Codec 8 and 2 bytes in Codec 8 Extended.
    event_id_size = 2 if codec == 0x8E else 1
    event_io_id, off = read_uint(payload, off, event_id_size)

    io, off, total_io = parse_io(payload, off, codec)

    # Teltonika coordinates are scaled by 10^7.
    record = {
        "timestamp_ms": ts_ms,
        "ts_utc": datetime.fromtimestamp(ts_ms / 1000, tz=timezone.utc).isoformat(),
        "priority": priority,
        "longitude": lon_raw / 10_000_000,
        "latitude": lat_raw / 10_000_000,
        "altitude": altitude,
        "angle": angle,
        "satellites": satellites,
        "speed_kmh": speed,
        "event_io_id": event_io_id,
        "io": io,
        "total_io": total_io,
    }
    return record, off


def parse_tcp_avl(packet: bytes):
    """Parse a Teltonika TCP AVL packet. Returns (codec, records)."""
    if len(packet) < 12 or packet[:4] != b"\x00\x00\x00\x00":
        raise ValueError("Not a Teltonika AVL packet")

    data_len = int.from_bytes(packet[4:8], "big")
    if len(packet) < 8 + data_len + 4:
        raise ValueError("Incomplete AVL packet")

    data = packet[8:8+data_len]
    codec = data[0]
    if codec not in (0x08, 0x8E):
        raise ValueError(f"Unsupported Codec ID: 0x{codec:02X}")

    count1 = data[1]
    off = 2
    records = []
    for _ in range(count1):
        rec, off = parse_avl_record(data, off, codec)
        records.append(rec)

    if off >= len(data):
        raise ValueError("Missing Number of Data 2")

    count2 = data[off]
    off += 1
    if count1 != count2:
        raise ValueError("AVL record count mismatch")

    # CRC is CRC-16/IBM over data from Codec ID through Number of Data 2.
    crc_received = int.from_bytes(packet[8+data_len:8+data_len+4], "big")
    crc_calculated = crc16_ibm(data)
    if crc_received != crc_calculated:
        raise ValueError(
            f"CRC mismatch: received {crc_received:04X}, calculated {crc_calculated:04X}"
        )

    return codec, records


def apply_io(rec):
    io = rec.get("io", {})
    # Common Teltonika IDs used by FMC150.
    # 239 = Ignition, 240 = Movement, 21 = GSM Signal,
    # 66 = External Voltage, 16 = Total Odometer.
    ignition = io.get(239)
    movement = io.get(240)
    gsm = io.get(21)
    ext_v = io.get(66)
    odo = io.get(16)

    # External voltage is typically reported in millivolts.
    if ext_v is not None:
        ext_v = ext_v / 1000.0

    return ignition, movement, gsm, ext_v, odo


def save_records(imei, records):
    con = db()
    con.execute(
        """INSERT INTO devices(imei, created_at, last_seen)
           VALUES(?, ?, ?)
           ON CONFLICT(imei) DO UPDATE SET last_seen=excluded.last_seen""",
        (imei, now_utc(), now_utc()),
    )

    for rec in records:
        ignition, movement, gsm, ext_v, odo = apply_io(rec)
        con.execute(
            """INSERT INTO positions
            (imei, ts_utc, latitude, longitude, altitude, angle, satellites,
             speed_kmh, priority, ignition, movement, gsm_signal,
             external_voltage, total_odometer, raw_json)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                imei, rec["ts_utc"], rec["latitude"], rec["longitude"],
                rec["altitude"], rec["angle"], rec["satellites"],
                rec["speed_kmh"], rec["priority"], ignition, movement,
                gsm, ext_v, odo, json.dumps(rec, ensure_ascii=False),
            ),
        )
    con.commit()
    con.close()


async def read_exact(reader, n):
    return await reader.readexactly(n)


async def handle_tracker(reader, writer):
    peer = writer.get_extra_info("peername")
    print(f"[GPS] Connection from {peer}")

    try:
        # Teltonika TCP handshake: 2-byte IMEI length + ASCII IMEI.
        imei_len = int.from_bytes(await read_exact(reader, 2), "big")
        if imei_len <= 0 or imei_len > 32:
            raise ValueError(f"Invalid IMEI length: {imei_len}")

        imei = (await read_exact(reader, imei_len)).decode("ascii", errors="strict")
        print(f"[GPS] IMEI: {imei}")

        # Accept device.
        writer.write(b"\x01")
        await writer.drain()

        while True:
            header = await reader.readexactly(8)
            if header[:4] != b"\x00\x00\x00\x00":
                raise ValueError("Invalid AVL preamble")

            data_len = int.from_bytes(header[4:8], "big")
            if data_len < 3 or data_len > 1280:
                raise ValueError(f"Invalid AVL data length: {data_len}")

            body_and_crc = await read_exact(reader, data_len + 4)
            packet = header + body_and_crc

            codec, records = parse_tcp_avl(packet)
            save_records(imei, records)

            # Acknowledge number of records as 4-byte integer.
            writer.write(len(records).to_bytes(4, "big"))
            await writer.drain()

            print(
                f"[GPS] {imei}: received {len(records)} record(s), "
                f"codec=0x{codec:02X}"
            )

    except asyncio.IncompleteReadError:
        print(f"[GPS] Connection closed: {peer}")
    except Exception as e:
        print(f"[GPS] Error {peer}: {e}")
    finally:
        writer.close()
        try:
            await writer.wait_closed()
        except Exception:
            pass


LOGIN_HTML = r"""<!doctype html><html lang="hr"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Prijava – Globtour GPS</title><style>*{box-sizing:border-box}body{margin:0;min-height:100vh;display:flex;align-items:center;justify-content:center;background:#f4f6f8;font-family:Inter,Segoe UI,Arial,sans-serif;color:#17202a}.login{width:min(420px,calc(100% - 32px));background:#fff;border:1px solid #e3e7eb;border-radius:18px;padding:30px;box-shadow:0 12px 40px #0002}.logo{font-size:26px;font-weight:850;margin-bottom:5px}.logo span{font-weight:500}.sub{color:#66727e;font-size:13px;margin-bottom:25px}.field{margin-bottom:14px}.field label{display:block;font-size:12px;font-weight:700;margin-bottom:6px}.field input{width:100%;height:44px;border:1px solid #d6dce2;border-radius:9px;padding:0 12px;font-size:14px;outline:none}.field input:focus{border-color:#b6a000;box-shadow:0 0 0 3px #ffd40033}.btn{width:100%;height:44px;border:0;border-radius:9px;background:#ffd400;color:#17202a;font-size:14px;font-weight:800;cursor:pointer}.err{min-height:20px;color:#c73535;font-size:12px;margin:8px 0}.foot{font-size:11px;color:#8a949e;text-align:center;margin-top:18px}</style></head><body><div class="login"><div class="logo">GLOBTOUR <span>GPS</span></div><div class="sub">Prijava u sustav za praćenje vozila</div><form id="f"><div class="field"><label>Korisničko ime</label><input id="u" autocomplete="username" required autofocus></div><div class="field"><label>Lozinka</label><input id="p" type="password" autocomplete="current-password" required></div><div id="e" class="err"></div><button class="btn" type="submit">Prijava</button></form><div class="foot">Pristup je dozvoljen samo ovlaštenim korisnicima.</div></div><script>document.getElementById('f').addEventListener('submit',async e=>{e.preventDefault();const m=document.getElementById('e');m.textContent='Prijava...';try{const r=await fetch('/api/login',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({username:document.getElementById('u').value.trim(),password:document.getElementById('p').value})});const d=await r.json().catch(()=>({}));if(!r.ok){m.textContent=d.detail||'Pogrešno korisničko ime ili lozinka.';return}location.href='/';}catch(x){m.textContent='Greška veze sa serverom.';}});</script></body></html>"""

@app.get("/login", response_class=HTMLResponse)
async def login_page(request: Request):
    if get_current_user(request):
        return RedirectResponse(url="/", status_code=303)
    return HTMLResponse(LOGIN_HTML)

@app.post("/api/login")
async def login(request: Request):
    body = await request.json()
    username = str(body.get("username", "") or "").strip()
    password = str(body.get("password", "") or "")
    con = db()
    row = con.execute("SELECT id, username, full_name, role, permissions, active, password_hash FROM users WHERE username=?", (username,)).fetchone()
    con.close()
    if not row or not int(row["active"] or 0) or not verify_password(password, row["password_hash"]):
        return JSONResponse({"detail":"Pogrešno korisničko ime ili lozinka."}, status_code=401)
    request.session.clear()
    request.session["user_id"] = int(row["id"])
    request.session["username"] = row["username"]
    request.session["role"] = row["role"]
    return {"ok":True,"username":row["username"],"role":row["role"]}

@app.post("/api/logout")
async def logout(request: Request):
    request.session.clear()
    return {"ok":True}

@app.get("/api/me")
async def me(request: Request):
    user = get_current_user(request)
    if not user:
        raise HTTPException(status_code=401, detail="Niste prijavljeni.")
    user.pop("permissions", None)
    return user

@app.get("/", response_class=HTMLResponse)
async def dashboard(request: Request):
    if not get_current_user(request):
        return RedirectResponse(url="/login", status_code=303)
    return templates.TemplateResponse("index.html", {"request": request})


def _vehicle_status(last_seen, speed_kmh, last_motion_ts):
    """
    Status se određuje ISKLJUČIVO na serveru.
    - OFFLINE: server nije primio podatak > 10 min
    - U VOŽNJI: zadnji primljeni GPS zapis ima brzinu > 1 km/h
    - PARKIRAN: online je, ali nema kretanja >= 2 min
    - ZAUSTAVLJEN: online je i zaustavljen kraće od 2 min
    """
    if not last_seen:
        return False, "OFFLINE", None

    try:
        seen = datetime.fromisoformat(str(last_seen).replace("Z", "+00:00"))
        if seen.tzinfo is None:
            seen = seen.replace(tzinfo=timezone.utc)
        age = max(0.0, (datetime.now(timezone.utc) - seen).total_seconds())
    except Exception:
        return False, "OFFLINE", None

    if age > 10 * 60:
        return False, "OFFLINE", round(age, 1)

    speed = float(speed_kmh or 0)
    if speed > 1:
        return True, "U VOŽNJI", round(age, 1)

    motion_age = None
    if last_motion_ts:
        try:
            motion = datetime.fromisoformat(str(last_motion_ts).replace("Z", "+00:00"))
            if motion.tzinfo is None:
                motion = motion.replace(tzinfo=timezone.utc)
            motion_age = max(0.0, (datetime.now(timezone.utc) - motion).total_seconds())
        except Exception:
            motion_age = None

    if motion_age is not None and motion_age >= 2 * 60:
        return True, "PARKIRAN", round(age, 1)

    return True, "ZAUSTAVLJEN", round(age, 1)


@app.get("/api/devices")
async def devices(request: Request):
    require_user(request)
    con = db()
    rows = con.execute("""
        SELECT d.imei, d.registration, d.make_model, d.vehicle_type,
               d.production_year, d.note, d.last_seen,
               p.ts_utc, p.latitude, p.longitude, p.speed_kmh,
               p.angle, p.satellites, p.ignition, p.movement,
               p.gsm_signal, p.external_voltage, p.total_odometer,
               (
                   SELECT p3.ts_utc
                   FROM positions p3
                   WHERE p3.imei=d.imei
                     AND COALESCE(p3.speed_kmh, 0) > 1
                   ORDER BY p3.id DESC
                   LIMIT 1
               ) AS last_motion_ts
        FROM devices d
        LEFT JOIN positions p ON p.id = (
            SELECT id FROM positions p2
            WHERE p2.imei=d.imei
            ORDER BY p2.id DESC LIMIT 1
        )
        ORDER BY d.imei
    """).fetchall()
    con.close()

    result = []
    server_now = datetime.now(timezone.utc).isoformat()

    for row in rows:
        item = dict(row)
        online, status, age = _vehicle_status(
            item.get("last_seen"),
            item.get("speed_kmh"),
            item.get("last_motion_ts"),
        )
        item["online"] = online
        item["status"] = status
        item["last_seen_age_sec"] = age
        item["server_now"] = server_now
        item.pop("last_motion_ts", None)
        result.append(item)

    return JSONResponse(result)


def _local_day_range(date_from: str | None, date_to: str | None):
    """Convert local Bosnia/Croatia calendar dates to UTC ISO boundaries."""
    tz = ZoneInfo("Europe/Sarajevo")
    if not date_from and not date_to:
        return None, None
    if not date_from:
        date_from = date_to
    if not date_to:
        date_to = date_from
    d1 = datetime.strptime(date_from, "%Y-%m-%d").date()
    d2 = datetime.strptime(date_to, "%Y-%m-%d").date()
    if d2 < d1:
        d1, d2 = d2, d1
    start = datetime.combine(d1, datetime.min.time(), tzinfo=tz).astimezone(timezone.utc)
    end = datetime.combine(d2 + timedelta(days=1), datetime.min.time(), tzinfo=tz).astimezone(timezone.utc)
    return start.isoformat(), end.isoformat()


def _distance_km(points):
    total = 0.0
    for a, b in zip(points, points[1:]):
        if a["latitude"] is None or a["longitude"] is None or b["latitude"] is None or b["longitude"] is None:
            continue
        lat1, lon1 = math.radians(a["latitude"]), math.radians(a["longitude"])
        lat2, lon2 = math.radians(b["latitude"]), math.radians(b["longitude"])
        dlat, dlon = lat2 - lat1, lon2 - lon1
        h = math.sin(dlat/2)**2 + math.cos(lat1) * math.cos(lat2) * math.sin(dlon/2)**2
        total += 6371.0088 * 2 * math.asin(math.sqrt(min(1.0, h)))
    return total


@app.get("/api/positions/{imei}")
async def history(imei: str, request: Request, limit: int = 500):
    require_user(request)
    limit = max(1, min(limit, 5000))
    params = request.query_params
    start_utc, end_utc = _local_day_range(params.get("from"), params.get("to"))
    con = db()
    if start_utc and end_utc:
        rows = con.execute("""
            SELECT ts_utc, latitude, longitude, speed_kmh, angle
            FROM positions
            WHERE imei=? AND ts_utc>=? AND ts_utc<?
            ORDER BY id ASC LIMIT ?
        """, (imei, start_utc, end_utc, limit)).fetchall()
    else:
        rows = con.execute("""
            SELECT ts_utc, latitude, longitude, speed_kmh, angle
            FROM positions
            WHERE imei=?
            ORDER BY id DESC LIMIT ?
        """, (imei, limit)).fetchall()
    con.close()
    return JSONResponse([dict(r) for r in rows])


@app.get("/api/history/{imei}")
async def history_summary(imei: str, request: Request, limit: int = 5000):
    require_user(request)
    limit = max(1, min(limit, 20000))
    params = request.query_params
    start_utc, end_utc = _local_day_range(params.get("from"), params.get("to"))
    con = db()
    if start_utc and end_utc:
        rows = con.execute("""
            SELECT ts_utc, latitude, longitude, speed_kmh, angle
            FROM positions
            WHERE imei=? AND ts_utc>=? AND ts_utc<?
            ORDER BY id ASC LIMIT ?
        """, (imei, start_utc, end_utc, limit)).fetchall()
    else:
        rows = con.execute("""
            SELECT ts_utc, latitude, longitude, speed_kmh, angle
            FROM positions
            WHERE imei=?
            ORDER BY id ASC LIMIT ?
        """, (imei, limit)).fetchall()
    con.close()
    points = [dict(r) for r in rows]
    moving_speeds = [float(p["speed_kmh"] or 0) for p in points]
    stops = 0
    was_stopped = False
    for p in points:
        stopped = float(p["speed_kmh"] or 0) <= 2
        if stopped and not was_stopped:
            stops += 1
        was_stopped = stopped
    summary = {
        "points": len(points),
        "distance_km": round(_distance_km(points), 2),
        "max_speed_kmh": round(max(moving_speeds), 1) if moving_speeds else 0,
        "start_time": points[0]["ts_utc"] if points else None,
        "end_time": points[-1]["ts_utc"] if points else None,
        "stops": stops,
    }
    return JSONResponse({"points": points, "summary": summary})


@app.post("/api/devices/{imei}")
async def update_device(imei: str, request: Request):
    require_user(request)
    body = await request.json()
    registration = str(body.get("registration", "") or "").strip()
    make_model = str(body.get("make_model", "") or "").strip()
    vehicle_type = str(body.get("vehicle_type", "") or "").strip()
    note = str(body.get("note", "") or "").strip()

    year_raw = body.get("production_year")
    try:
        production_year = int(year_raw) if year_raw not in (None, "") else None
    except (TypeError, ValueError):
        production_year = None

    con = db()
    exists = con.execute("SELECT 1 FROM devices WHERE imei=?", (imei,)).fetchone()
    if exists:
        con.execute(
            """UPDATE devices
               SET registration=?, make_model=?, vehicle_type=?,
                   production_year=?, note=?
               WHERE imei=?""",
            (registration, make_model, vehicle_type, production_year, note, imei),
        )
    else:
        con.execute(
            """INSERT INTO devices
               (imei, registration, make_model, vehicle_type, production_year, note, created_at, last_seen)
               VALUES (?, ?, ?, ?, ?, ?, ?, NULL)""",
            (imei, registration, make_model, vehicle_type, production_year, note, now_utc()),
        )
    con.commit()
    con.close()
    return {"ok": True}


@app.get("/api/users")
async def list_users(request: Request):
    require_admin(request)
    con = db(); rows = con.execute("SELECT id, username, full_name, role, permissions, active, created_at FROM users ORDER BY username").fetchall(); con.close()
    return JSONResponse([dict(r) for r in rows])

@app.post("/api/users")
async def create_user(request: Request):
    require_admin(request)
    body = await request.json()
    username = str(body.get("username", "") or "").strip()
    password = str(body.get("password", "") or "")
    full_name = str(body.get("full_name", "") or "").strip()
    role = str(body.get("role", "user") or "user").strip().lower()
    if len(username) < 3: raise HTTPException(status_code=400, detail="Korisničko ime mora imati najmanje 3 znaka.")
    if len(password) < 8: raise HTTPException(status_code=400, detail="Lozinka mora imati najmanje 8 znakova.")
    if role not in ("admin", "user"): role = "user"
    con = db()
    try:
        cur = con.execute("INSERT INTO users(username,password_hash,full_name,role,permissions,active,created_at) VALUES(?,?,?,?,?,?,?)", (username,hash_password(password),full_name,role,"*" if role=="admin" else "",1,now_utc()))
        con.commit(); uid=cur.lastrowid
    except sqlite3.IntegrityError:
        con.close(); raise HTTPException(status_code=409, detail="Korisničko ime već postoji.")
    con.close(); return {"ok":True,"id":uid}

@app.patch("/api/users/{user_id}")
async def update_user(user_id: int, request: Request):
    admin = require_admin(request); body = await request.json(); con = db()
    row = con.execute("SELECT id, role, active FROM users WHERE id=?", (user_id,)).fetchone()
    if not row: con.close(); raise HTTPException(status_code=404, detail="Korisnik nije pronađen.")
    role = str(body.get("role", row["role"]) or row["role"]).strip().lower(); active = 1 if bool(body.get("active", bool(row["active"]))) else 0
    if role not in ("admin","user"): role = row["role"]
    if int(row["id"]) == int(admin["id"]) and (role != "admin" or not active):
        n = con.execute("SELECT COUNT(*) AS n FROM users WHERE role='admin' AND active=1 AND id<>?", (user_id,)).fetchone()["n"]
        if int(n)==0: con.close(); raise HTTPException(status_code=400, detail="Ne možete ukloniti ovlasti posljednjem aktivnom administratoru.")
    con.execute("UPDATE users SET role=?, active=? WHERE id=?", (role,active,user_id))
    if body.get("password"):
        password=str(body["password"]);
        if len(password)<8: con.close(); raise HTTPException(status_code=400, detail="Nova lozinka mora imati najmanje 8 znakova.")
        con.execute("UPDATE users SET password_hash=? WHERE id=?", (hash_password(password),user_id))
    con.commit(); con.close(); return {"ok":True}

if __name__ == "__main__":
    uvicorn.run(app, host=WEB_HOST, port=WEB_PORT, reload=False)
    
