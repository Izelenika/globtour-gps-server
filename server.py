import json
import sqlite3
import struct
import time
import os
from datetime import datetime, timezone, timedelta
from zoneinfo import ZoneInfo
import math
from pathlib import Path
from contextlib import asynccontextmanager

from fastapi import FastAPI, Request
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.templating import Jinja2Templates
import uvicorn

BASE = Path(__file__).resolve().parent
DB = BASE / "globtour_gps.db"
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

    con.commit()
    con.close()


def now_utc():
    return datetime.now(timezone.utc).isoformat()


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


@app.get("/", response_class=HTMLResponse)
async def dashboard(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})


@app.get("/api/devices")
async def devices():
    con = db()
    rows = con.execute("""
        SELECT d.imei, d.registration, d.make_model, d.vehicle_type,
               d.production_year, d.note, d.last_seen,
               p.ts_utc, p.latitude, p.longitude, p.speed_kmh,
               p.angle, p.satellites, p.ignition, p.movement,
               p.gsm_signal, p.external_voltage, p.total_odometer
        FROM devices d
        LEFT JOIN positions p ON p.id = (
            SELECT id FROM positions p2
            WHERE p2.imei=d.imei
            ORDER BY p2.id DESC LIMIT 1
        )
        ORDER BY d.imei
    """).fetchall()
    con.close()
    return JSONResponse([dict(r) for r in rows])


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


if __name__ == "__main__":
    uvicorn.run(app, host=WEB_HOST, port=WEB_PORT, reload=False)
    
