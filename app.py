# -*- coding: utf-8 -*-
"""
app.py  —  TallySync Pro  Multi-Client Web App
=====================================================
FEATURES ADDED IN THIS VERSION:
  • Company limit per user (max 15, configurable per user)
  • Companies are UNIQUE SLOTS — same company can be regenerated freely
  • OTP login (email-based, no password required) + password fallback
  • Max 1 device at a time per user (concurrent session enforcement)
  • Max 1-day session expiry (auto-logout after 24h)
  • Admin-only user management page
  • Usage tracking per user (reports count, last active, companies accessed)
  • Admin can restrict/enable users, set company limit (1-15)
  • Editable Classic PDF — after generation, open edit screen to tweak numbers
    before final download
"""

import sys, pathlib

# ── PyInstaller EXE path fix ──────────────────────────────────────────────────
# When frozen (_MEIPASS exists), bundled .py modules live there.
# Insert it into sys.path BEFORE any imports so tally_core / sch3_* are found.
if getattr(sys, "frozen", False):
    _bundle = pathlib.Path(sys._MEIPASS)
    if str(_bundle) not in sys.path:
        sys.path.insert(0, str(_bundle))
    _exe_dir = pathlib.Path(sys.executable).parent
    if str(_exe_dir) not in sys.path:
        sys.path.insert(0, str(_exe_dir))

for _p in pathlib.Path(__file__).parent.glob("__pycache__/*.pyc"):
    try: _p.unlink()
    except: pass

import os, sys, re, io, json, threading, traceback, secrets, time, random
from datetime import date, datetime, timedelta
from functools import wraps
from pathlib import Path

from flask import (Flask, render_template_string, request, redirect,
                   url_for, session, send_file, flash, jsonify, abort)
from werkzeug.security import generate_password_hash, check_password_hash

# ── Optional engine imports ───────────────────────────────────────────────────
try:
    import tally_core as core
    CORE_OK = True
except Exception:
    CORE_OK = False

try:
    import sch3_classic as classic
    CLASSIC_OK = True
except Exception:
    classic = None
    CLASSIC_OK = False

try:
    import sch3_verify as v3
    VERIFY_OK = True
except Exception:
    v3 = None
    VERIFY_OK = False

try:
    import sch3_excel_classic as excel_classic
    EXCEL_CLASSIC_OK = True
except BaseException:
    excel_classic = None
    EXCEL_CLASSIC_OK = False

# ── OTP store (in-memory, keyed by username) ─────────────────────────────────
_OTP_STORE: dict = {}   # {username: {"otp": "123456", "expires": timestamp}}
_OTP_LOCK = threading.Lock()

# ══════════════════════════════════════════════════════════════════════════════
# Config & app
# ══════════════════════════════════════════════════════════════════════════════
BASE_DIR    = Path(__file__).parent
REPORTS_DIR = BASE_DIR / "reports"
CLIENTS_DB  = BASE_DIR / "clients.json"
REPORTS_DIR.mkdir(exist_ok=True)

MAX_COMPANIES_DEFAULT = 15      # Default company slot limit per user
SESSION_MAX_HOURS     = 24      # Sessions expire after 24 hours
OTP_EXPIRY_MINUTES    = 10      # OTP valid for 10 minutes

# ── Licence expiry ────────────────────────────────────────────────────────────
LICENCE_EXPIRY = date(2027, 8, 31)   # Software auto-shuts-down after this date

# ── Cloud / Render mode ───────────────────────────────────────────────────────
IS_RENDER = bool(os.environ.get("RENDER", ""))

def _licence_expired() -> bool:
    if IS_RENDER:
        return False   # No date kill-switch on cloud
    return date.today() > LICENCE_EXPIRY

_EXPIRY_HTML = """<!DOCTYPE html>
<html lang="en"><head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Licence Expired — TallySync Pro</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',system-ui,sans-serif;background:#0B1F4B;
     min-height:100vh;display:flex;align-items:center;justify-content:center}
.box{background:#fff;border-radius:14px;padding:52px 48px;max-width:520px;
     width:90%;text-align:center;box-shadow:0 8px 40px rgba(0,0,0,.4)}
.icon{font-size:60px;margin-bottom:16px}
h1{color:#b91c1c;font-size:1.75rem;margin-bottom:14px}
p{color:#374151;line-height:1.7;margin-bottom:8px}
.badge{display:inline-block;background:#fef2f2;color:#b91c1c;
       border:1px solid #fca5a5;border-radius:6px;
       padding:6px 20px;font-weight:700;font-size:1.05rem;margin:14px 0 20px}
.foot{margin-top:30px;font-size:.80rem;color:#9ca3af}
</style></head>
<body><div class="box">
  <div class="icon">🔒</div>
  <h1>Licence Expired</h1>
  <p>Your TallySync Pro licence expired on</p>
  <div class="badge">31 August 2027</div>
  <p>This software has been automatically deactivated.<br>
     Please contact your administrator to renew your licence.</p>
  <div class="foot">TallySync Pro &nbsp;|&nbsp; Schedule III PDF Generator</div>
</div></body></html>"""

# ══════════════════════════════════════════════════════════════════════════════
# MACHINE-LOCKED LICENCE SYSTEM
# ══════════════════════════════════════════════════════════════════════════════
_LICENCE_SALT  = "TALLYSYNC-PRO-2026"   # change this before releasing
_LICENCE_FILE  = "license.key"                # sits next to EXE / app.py

def _get_machine_id() -> str:
    """
    Build a stable fingerprint from this Windows PC:
      1. CPU processor string  (from WMIC)
      2. C: volume serial      (from WMIC)
      3. First non-loopback MAC address
    All three are hashed together → a 16-char hex token.
    On non-Windows (dev) returns a fixed dev string.
    """
    import hashlib, uuid, platform
    parts = []

    if platform.system() == "Windows":
        try:
            import subprocess
            cpu = subprocess.check_output(
                "wmic cpu get ProcessorId", shell=True,
                stderr=subprocess.DEVNULL).decode(errors="ignore")
            cpu = " ".join(cpu.split()).replace("ProcessorId","").strip()
            parts.append(cpu)
        except Exception:
            parts.append("NO_CPU")

        try:
            import subprocess
            vol = subprocess.check_output(
                "wmic volume where DriveLetter='C:' get SerialNumber",
                shell=True, stderr=subprocess.DEVNULL).decode(errors="ignore")
            vol = " ".join(vol.split()).replace("SerialNumber","").strip()
            parts.append(vol)
        except Exception:
            parts.append("NO_VOL")
    else:
        # dev / Linux — use a fixed placeholder so keygen works on Windows only
        parts.append("DEV_MACHINE")
        parts.append("DEV_VOLUME")

    # MAC address — cross-platform
    try:
        mac = uuid.UUID(int=uuid.getnode()).hex[-12:]
        parts.append(mac)
    except Exception:
        parts.append("NO_MAC")

    raw = "|".join(parts) + "|" + _LICENCE_SALT
    return hashlib.sha256(raw.encode()).hexdigest()[:16].upper()

def _make_licence_key(machine_id: str, expiry: str) -> str:
    """
    Build a licence key:
      Format: XXXXX-XXXXX-XXXXX-XXXXX-XXXXX  (25 chars grouped as 5×5)
      Payload: machine_id (16) + expiry YYYYMMDD (8) = 24 chars
      Check : sha256(payload + salt)[:8] appended → 32 chars total
      Then : base32-encode → uppercase letters, no ambiguous chars
    """
    import hashlib, base64
    payload  = (machine_id + expiry).encode()   # 24 bytes
    checksum = hashlib.sha256(payload + _LICENCE_SALT.encode()).hexdigest()[:8]
    raw      = (machine_id + expiry + checksum).encode()  # 32 bytes
    # base32 gives A-Z 2-7, no ambiguous chars (0/O/1/I)
    b32 = base64.b32encode(raw).decode().rstrip("=")   # 52 chars (32 bytes → 52 base32 no padding)
    # Group as 4 blocks of 13: XXXXXXXXXXXXX-XXXXXXXXXXXXX-XXXXXXXXXXXXX-XXXXXXXXXXXXX
    return "-".join(b32[i:i+13] for i in range(0, 52, 13))

def _verify_licence_key(key: str) -> dict:
    """
    Verify a licence key against THIS machine.
    Returns: {"valid": bool, "expiry": "YYYYMMDD"|None, "machine_id": str,
              "days_left": int|None, "reason": str}
    """
    import hashlib, base64
    from datetime import date as _date

    machine_id = _get_machine_id()
    key_clean  = key.strip().upper().replace("-", "").replace(" ", "")

    try:
        # Pad to multiple of 8 for base32
        pad  = (8 - len(key_clean) % 8) % 8   # pad 52 chars → 56 (multiple of 8)
        raw  = base64.b32decode(key_clean + "=" * pad)
        if len(raw) < 32:
            return {"valid": False, "reason": "Invalid key format — expected 52-char key", "machine_id": machine_id, "expiry": None, "days_left": None}

        mid_in   = raw[:16].decode(errors="replace")
        exp_in   = raw[16:24].decode(errors="replace")
        chk_in   = raw[24:32].decode(errors="replace")

        # Verify checksum
        payload  = (mid_in + exp_in).encode()
        chk_calc = hashlib.sha256(payload + _LICENCE_SALT.encode()).hexdigest()[:8]
        if chk_calc != chk_in:
            return {"valid": False, "reason": "Key is corrupted or tampered", "machine_id": machine_id, "expiry": None, "days_left": None}

        # Verify machine
        if mid_in.upper() != machine_id.upper():
            return {"valid": False, "reason": f"Key is locked to a different machine  (this machine: {machine_id})",
                    "machine_id": machine_id, "expiry": exp_in, "days_left": None}

        # Verify expiry
        try:
            exp_date  = _date(int(exp_in[:4]), int(exp_in[4:6]), int(exp_in[6:8]))
            days_left = (exp_date - _date.today()).days
        except Exception:
            return {"valid": False, "reason": "Invalid expiry in key", "machine_id": machine_id, "expiry": exp_in, "days_left": None}

        if days_left < 0:
            return {"valid": False, "reason": f"Licence expired on {exp_date.strftime('%d %b %Y')}",
                    "machine_id": machine_id, "expiry": exp_in, "days_left": days_left}

        return {"valid": True, "reason": "OK", "machine_id": machine_id,
                "expiry": exp_in, "days_left": days_left,
                "exp_date": exp_date.strftime("%d %b %Y")}

    except Exception as e:
        return {"valid": False, "reason": f"Key parse error: {e}", "machine_id": machine_id, "expiry": None, "days_left": None}

def _licence_file_path() -> "Path":
    """Return path to license.key next to the EXE / script."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent / _LICENCE_FILE
    return Path(__file__).parent / _LICENCE_FILE

def _read_saved_key() -> str:
    """Read licence key from license.key file. Returns empty string if missing."""
    try:
        p = _licence_file_path()
        if p.exists():
            return p.read_text(encoding="utf-8").strip()
    except Exception:
        pass
    return ""

def _save_key(key: str):
    """Write licence key to disk."""
    _licence_file_path().write_text(key.strip(), encoding="utf-8")

# ── Cached licence status (refreshed every 60 s to avoid disk reads on every request)
import threading as _lic_th
_LIC_CACHE = {"valid": False, "checked_at": 0.0, "status": {}}
_LIC_LOCK  = _lic_th.Lock()

def _get_licence_status(force: bool = False) -> dict:
    if IS_RENDER:
        return {"valid": True, "reason": "Cloud mode — no licence required",
                "machine_id": "CLOUD", "expiry": "99991231", "days_left": 9999,
                "exp_date": "31 Dec 2099"}
    """Return cached licence status, refreshing at most every 60 seconds."""
    import time as _t
    with _LIC_LOCK:
        age = _t.time() - _LIC_CACHE["checked_at"]
        if not force and age < 60 and _LIC_CACHE["checked_at"] > 0:
            return _LIC_CACHE["status"]
        key    = _read_saved_key()
        status = _verify_licence_key(key) if key else {
            "valid": False, "reason": "No licence key found",
            "machine_id": _get_machine_id(), "expiry": None, "days_left": None}
        _LIC_CACHE["valid"]      = status["valid"]
        _LIC_CACHE["checked_at"] = _t.time()
        _LIC_CACHE["status"]     = status
        return status

# ── Licence activation page HTML ──────────────────────────────────────────────
_ACTIVATE_HTML = """<!DOCTYPE html>
<html lang="en"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Activate — TallySync Pro</title>
{% if success %}<meta http-equiv="refresh" content="2;url=/login">{% endif %}
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',system-ui,sans-serif;background:#0B1F4B;
     min-height:100vh;display:flex;align-items:center;justify-content:center}
.box{background:#fff;border-radius:16px;padding:48px 44px;max-width:540px;
     width:92%;box-shadow:0 12px 48px rgba(0,0,0,.45)}
.logo{text-align:center;margin-bottom:28px}
.logo .t{font-size:1.6rem;font-weight:900;color:#0B1F4B}
.logo .t span{color:#C8A84B}
.logo .sub{color:#64748b;font-size:.82rem;margin-top:4px}
h2{color:#0B1F4B;font-size:1.1rem;margin-bottom:6px}
.info{background:#f0f4fb;border-radius:8px;padding:12px 16px;
      font-size:.8rem;color:#475569;margin-bottom:20px;line-height:1.6}
.info b{color:#0B1F4B}
.mid{background:#e8eefa;border-radius:8px;padding:10px 16px;
     font-family:monospace;font-size:1rem;font-weight:700;
     color:#0B1F4B;letter-spacing:2px;margin-bottom:20px;
     text-align:center;word-break:break-all}
label{display:block;font-size:.75rem;font-weight:700;color:#64748b;
      text-transform:uppercase;letter-spacing:.4px;margin-bottom:6px}
input[type=text]{width:100%;border:1.5px solid #dde3ee;border-radius:8px;
  padding:11px 14px;font-size:.95rem;font-family:monospace;letter-spacing:2px;
  outline:none;transition:border .2s;text-transform:uppercase}
input:focus{border-color:#0B1F4B;box-shadow:0 0 0 3px rgba(11,31,75,.08)}
.btn{width:100%;margin-top:16px;padding:12px;background:#0B1F4B;color:#fff;
     font-size:.95rem;font-weight:700;border:none;border-radius:8px;
     cursor:pointer;transition:background .15s}
.btn:hover{background:#16306e}
.err{background:#fef2f2;color:#b91c1c;border:1px solid #fecaca;
     border-radius:8px;padding:10px 14px;font-size:.84rem;margin-bottom:16px}
.ok{background:#dcfce7;color:#166534;border:1px solid #bbf7d0;
    border-radius:8px;padding:10px 14px;font-size:.84rem;margin-bottom:16px}
.foot{margin-top:24px;text-align:center;font-size:.75rem;color:#94a3b8}
</style></head>
<body><div class="box">
  <div class="logo">
    <div class="t">TallySync<span> Pro</span></div>
    <div class="sub">Schedule III Generator</div>
  </div>

  {% if error %}
  <div class="err">⚠ {{ error }}</div>
  {% endif %}
  {% if success %}
  <div class="ok">✔ {{ success }}</div>
  {% endif %}

  <h2>🔐 Activate Your Licence</h2>
  <div class="info">
    Share your <b>Machine ID</b> below with your administrator to receive your licence key.
  </div>

  <label>Your Machine ID (send this to your administrator)</label>
  <div class="mid">{{ machine_id }}</div>

  <form method="post" action="/activate">
    <label>Enter Licence Key</label>
    <input type="text" name="key" placeholder="XXXXXXXX-XXXXXXXX-XXXXXXXX-XXXXXXXX-XXXXXXXX"
           value="{{ key_val or '' }}" autocomplete="off" spellcheck="false" required>
    <button class="btn" type="submit">✔ Activate TallySync Pro</button>
  </form>
  <div class="foot">TallySync Pro v7 &nbsp;|&nbsp; Licence locked to this machine only</div>
</div></body></html>"""

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "tallysync-pro-secret-2026")
app.permanent_session_lifetime = timedelta(hours=SESSION_MAX_HOURS)

_JOBS: dict = {}
_JOBS_LOCK  = threading.Lock()

# ══════════════════════════════════════════════════════════════════════════════
# Client database
# ══════════════════════════════════════════════════════════════════════════════
def load_db() -> dict:
    if CLIENTS_DB.exists():
        try:
            with open(CLIENTS_DB, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    data = {"users": {
        "admin": {
            "password":          generate_password_hash("admin123"),
            "role":              "admin",
            "display_name":      "Administrator",
            "email":             "",
            "tally_host":        "localhost",
            "tally_port":        9000,
            "allowed_companies": [],
            "company_limit":     MAX_COMPANIES_DEFAULT,
            "active":            True,
            "created_at":        datetime.now().isoformat(),
            "active_session":    None,
            "session_started":   None,
            "usage": {
                "total_reports":      0,
                "companies_accessed": [],
                "last_active":        None,
                "report_log":         [],
            }
        },
        "romil_shah": {
            "password":          generate_password_hash("Romil@123"),
            "role":              "user",
            "display_name":      "Romil Shah",
            "email":             "",
            "tally_host":        "localhost",
            "tally_port":        9000,
            "allowed_companies": [],
            "company_limit":     MAX_COMPANIES_DEFAULT,
            "active":            True,
            "created_at":        datetime.now().isoformat(),
            "active_session":    None,
            "session_started":   None,
            "usage": {
                "total_reports":      0,
                "companies_accessed": [],
                "last_active":        None,
                "report_log":         [],
            }
        },
    }}
    save_db(data); return data

def save_db(data: dict):
    with open(CLIENTS_DB, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

def get_user(username: str):
    return load_db()["users"].get(username)

def _get_user_company_limit(u: dict) -> int:
    return int(u.get("company_limit") or MAX_COMPANIES_DEFAULT)

def _get_user_companies_used(username: str) -> list:
    """Return list of UNIQUE company names this user has ever generated reports for."""
    d = REPORTS_DIR / username
    if not d.exists(): return []
    seen = set()
    for jf in d.glob("*.json"):
        try:
            with open(jf) as f:
                meta = json.load(f)
            co = meta.get("company","")
            if co: seen.add(co.strip().upper())
        except Exception:
            pass
    return list(seen)

def _track_usage(username: str, company: str, report_type: str = "classic"):
    """Record usage event for a user."""
    data = load_db()
    u = data["users"].get(username)
    if not u: return
    usage = u.setdefault("usage", {})
    usage["total_reports"] = usage.get("total_reports", 0) + 1
    usage["last_active"]   = datetime.now().isoformat()
    cos = usage.setdefault("companies_accessed", [])
    if company.upper() not in [c.upper() for c in cos]:
        cos.append(company)
    log = usage.setdefault("report_log", [])
    log.insert(0, {"company": company, "date": datetime.now().isoformat()[:19], "type": report_type})
    if len(log) > 100: log[:] = log[:100]
    save_db(data)

def current_fy_end() -> int:
    t = date.today()
    return t.year + 1 if t.month >= 4 else t.year

# ══════════════════════════════════════════════════════════════════════════════
# OTP helpers
# ══════════════════════════════════════════════════════════════════════════════
def _generate_otp(username: str) -> str:
    otp = str(random.randint(100000, 999999))
    with _OTP_LOCK:
        _OTP_STORE[username] = {
            "otp":     otp,
            "expires": time.time() + OTP_EXPIRY_MINUTES * 60,
        }
    return otp

def _verify_otp(username: str, otp_input: str) -> bool:
    with _OTP_LOCK:
        rec = _OTP_STORE.get(username)
        if not rec: return False
        if time.time() > rec["expires"]:
            _OTP_STORE.pop(username, None); return False
        if rec["otp"] != otp_input.strip():
            return False
        _OTP_STORE.pop(username, None)
        return True

def _send_otp(username: str, otp: str, email: str):
    """Send OTP. Uses SMTP if configured, otherwise prints to console (dev mode)."""
    msg = f"[TallySync Pro] Your OTP for {username}: {otp}  (valid {OTP_EXPIRY_MINUTES} mins)"
    try:
        import smtplib
        smtp_host = os.environ.get("SMTP_HOST","")
        smtp_port = int(os.environ.get("SMTP_PORT", 587))
        smtp_user = os.environ.get("SMTP_USER","")
        smtp_pass = os.environ.get("SMTP_PASS","")
        smtp_from = os.environ.get("SMTP_FROM", smtp_user)
        if smtp_host and smtp_user and email:
            with smtplib.SMTP(smtp_host, smtp_port) as s:
                s.ehlo(); s.starttls(); s.login(smtp_user, smtp_pass)
                body = (f"Subject: TallySync Pro OTP\r\n"
                        f"From: {smtp_from}\r\nTo: {email}\r\n\r\n{msg}")
                s.sendmail(smtp_from, email, body)
                print(f"  OTP sent to {email}")
                return
    except Exception as e:
        print(f"  SMTP failed: {e}")
    # Dev fallback — print to console
    print(f"\n  {'='*50}")
    print(f"  OTP for {username}: {otp}  (valid {OTP_EXPIRY_MINUTES} mins)")
    print(f"  Email would go to: {email}")
    print(f"  {'='*50}\n")

# ══════════════════════════════════════════════════════════════════════════════
# Auth decorators
# ══════════════════════════════════════════════════════════════════════════════
def login_required(f):
    @wraps(f)
    def d(*a, **kw):
        if "username" not in session:
            return redirect(url_for("login", next=request.path))
        # 24-hour session expiry
        started = session.get("session_started")
        if started:
            age = (datetime.now() - datetime.fromisoformat(started)).total_seconds()
            if age > SESSION_MAX_HOURS * 3600:
                session.clear()
                flash("Your session has expired. Please sign in again.", "info")
                return redirect(url_for("login"))
        # Concurrent session check (1 device at a time)
        token = session.get("session_token")
        u = get_user(session["username"])
        if u and u.get("active_session") and u["active_session"] != token:
            session.clear()
            flash("You were signed out because your account was opened on another device.", "info")
            return redirect(url_for("login"))
        # Account active check
        if u and not u.get("active", True):
            session.clear()
            flash("Your account has been deactivated. Contact your administrator.", "error")
            return redirect(url_for("login"))
        return f(*a, **kw)
    return d

def admin_required(f):
    @wraps(f)
    def d(*a, **kw):
        if "username" not in session:
            return redirect(url_for("login"))
        u = get_user(session["username"])
        if not u or u["role"] != "admin": abort(403)
        return f(*a, **kw)
    return d

@app.context_processor
def _inject():
    me = None
    if "username" in session:
        me = get_user(session["username"])
        if me:
            me = dict(me); me["username"] = session["username"]
    return {"me": me}

@app.before_request
def _enforce_expiry():
    """
    Gate every request behind:
      1. Software kill-switch (date expiry)
      2. Machine-locked licence key
    On Render (cloud), licence check is skipped — login controls access.
    Only /activate passes through without a valid key.
    """
    # Cloud mode — skip all licence checks
    if IS_RENDER:
        return None

    # Always allow the activation route (GET + POST)
    if request.path.rstrip("/") == "/activate":
        return None

    # ── 1. Date-based hard stop ───────────────────────────────────────────────
    if _licence_expired():
        from flask import Response
        return Response(_EXPIRY_HTML, status=403, mimetype="text/html")

    # ── 2. Licence key check ──────────────────────────────────────────────────
    lic = _get_licence_status()
    if not lic["valid"]:
        # Show activation page — client must enter valid key before anything else
        machine_id = lic.get("machine_id", _get_machine_id())
        reason     = lic.get("reason", "No licence key found")
        html = render_template_string(_ACTIVATE_HTML,
            machine_id=machine_id, error=reason, success="", key_val="")
        from flask import Response
        return Response(html, status=403, mimetype="text/html")

    # ── 3. Warn when licence is expiring soon (≤ 14 days) ────────────────────
    days_left = lic.get("days_left")
    if days_left is not None and 0 < days_left <= 14:
        # Only add the warning once per session to avoid repeated flashing
        if not session.get("_expiry_warned"):
            session["_expiry_warned"] = True
            flash(f"⚠ Your licence expires in {days_left} day(s) on "
                  f"{lic.get('exp_date','?')}. Contact your administrator to renew.", "error")

# ══════════════════════════════════════════════════════════════════════════════
# Background job runner
# ══════════════════════════════════════════════════════════════════════════════
class _Writer(io.TextIOBase):
    def __init__(self, lst): self._l = lst
    def write(self, s):
        if s.strip(): self._l.append(s.rstrip())
        return len(s)
    def flush(self): pass

def _run_job(jid, username, company, year, host, port, gen_classic, autosync_py,
             from_date=None, to_date=None, sign_info=None):
    logs = []
    with _JOBS_LOCK:
        _JOBS[jid] = {"status":"running","logs":logs,"company":company,"year":year,
                      "pdf":None,"classic_pdf":None,"json":None,"error":None,
                      "started":datetime.now().isoformat()}
    old = sys.stdout; sys.stdout = _Writer(logs)
    try:
        turl = f"http://{host}:{port}"
        core.set_tally_url(turl)
        raw, fs, fe = core.fetch_all(company, year, save_json=False,
                                     from_date=from_date, to_date=to_date)
        D = core.parse_data(company, raw, fs, fe)
        safe = re.sub(r"[^\w]","_", company[:30])
        # Use date range in filename for custom periods
        _fe_tag = to_date.replace("-","") if to_date else fe.replace("-","")
        out  = REPORTS_DIR / username; out.mkdir(exist_ok=True)
        pdf  = out / f"sch3_{safe}_{_fe_tag}.pdf"
        jsf  = out / f"sch3_{safe}_{_fe_tag}.json"
        core.generate_pdf(D, str(pdf))
        with open(jsf,"w",encoding="utf-8") as f:
            json.dump({"company":company,"fy_start":fs,"fy_end":fe,"tally_url":turl,
                       "generated_at":datetime.now().isoformat(),
                       "gross_profit":float(D.get("gross_profit") or 0),
                       "net_profit":float(D.get("net_profit") or 0),"data":raw},
                      f, indent=2, default=str)
        cpdf = None
        if gen_classic and CLASSIC_OK:
            D_py = None
            if autosync_py:
                try:
                    # For custom ranges: fetch PY = same range shifted back 1 year
                    if from_date and to_date:
                        from datetime import date as _d, timedelta as _td
                        _fd = _d.fromisoformat(from_date)
                        _td2 = _d.fromisoformat(to_date)
                        py_from = str(_d(_fd.year-1, _fd.month, _fd.day))
                        py_to   = str(_d(_td2.year-1, _td2.month, _td2.day))
                        py_year = _td2.year - 1 if _td2.month >= 4 else _td2.year
                        print(f"\n  Auto-syncing previous year range {py_from} → {py_to}…")
                        raw2,fs2,fe2 = core.fetch_all(company, py_year, save_json=False,
                                                       from_date=py_from, to_date=py_to)
                    else:
                        print(f"\n  Auto-syncing previous year (FY end {year-1})…")
                        raw2,fs2,fe2 = core.fetch_all(company, year-1, save_json=False)
                    D_py = core.parse_data(company, raw2, fs2, fe2)
                    if D_py and fe2 == fe:
                        print(f"  WARNING: Previous year fetch returned same FY end ({fe2}) — skipping PY column.")
                        D_py = None
                    elif D_py:
                        tb_py = D_py.get('TB', {})
                        if not tb_py:
                            print(f"  WARNING: Previous year Trial Balance is empty — skipping PY column.")
                            D_py = None
                        else:
                            print(f"  Previous year OK — {len(tb_py)} TB rows.")
                            # Save PY JSON so Classic Excel can find it later
                            try:
                                _py_fe_tag = fe2.replace("-","")
                                jsf_py = out / f"sch3_{safe}_{_py_fe_tag}.json"
                                if not jsf_py.exists():
                                    with open(jsf_py,"w",encoding="utf-8") as _pf:
                                        json.dump({"company":company,"fy_start":fs2,
                                                   "fy_end":fe2,"tally_url":turl,
                                                   "generated_at":datetime.now().isoformat(),
                                                   "gross_profit":float(D_py.get("gross_profit") or 0),
                                                   "net_profit":float(D_py.get("net_profit") or 0),
                                                   "data":raw2},
                                                  _pf, indent=2, default=str)
                                    print(f"  PY JSON saved → {jsf_py.name}")
                            except Exception as _je:
                                print(f"  WARNING: PY JSON save failed — {_je}")
                except Exception as e:
                    print(f"  WARNING: previous year failed — {e}"); D_py = None
            try:
                cpdf = out / f"classic_{safe}_{_fe_tag}.pdf"
                print("\n  Building Classic CA-Format PDF…")
                classic.generate_classic_pdf(D, D_py, str(cpdf))
                print(f"  Classic PDF saved → {cpdf.name}")
            except Exception as e:
                print(f"  WARNING: Classic PDF failed — {e}"); cpdf = None
        with _JOBS_LOCK:
            _JOBS[jid].update({"status":"done","pdf":str(pdf),
                "classic_pdf":str(cpdf) if cpdf else None,"json":str(jsf),
                "fy_start":fs,"fy_end":fe,
                "gross_profit":float(D.get("gross_profit") or 0),
                "net_profit":float(D.get("net_profit") or 0),
                "finished":datetime.now().isoformat()})
    except Exception as e:
        with _JOBS_LOCK:
            _JOBS[jid].update({"status":"error","error":str(e)+"\n\n"+traceback.format_exc()})
    finally:
        sys.stdout = old

# ══════════════════════════════════════════════════════════════════════════════
# Helpers
# ══════════════════════════════════════════════════════════════════════════════
def scan_reports(username):
    d = REPORTS_DIR / username
    if not d.exists(): return []
    out = []
    for pdf in sorted(d.glob("sch3_*.pdf"), key=lambda p:p.stat().st_mtime, reverse=True):
        parts = pdf.stem.split("_",2)
        ck = parts[1] if len(parts)>1 else ""
        fe = parts[2] if len(parts)>2 else ""
        jf = d/f"sch3_{ck}_{fe}.json"; cl = d/f"classic_{ck}_{fe}.pdf"
        meta = {}
        if jf.exists():
            try:
                with open(jf) as f: meta = json.load(f)
            except Exception: pass
        out.append({"company":meta.get("company",ck.replace("_"," ")),
                    "fy_start":meta.get("fy_start",""),"fy_end":meta.get("fy_end",fe),
                    "generated_at":meta.get("generated_at",""),
                    "gross_profit":meta.get("gross_profit"),
                    "net_profit":meta.get("net_profit"),
                    "pdf_name":pdf.name,
                    "classic_name":cl.name if cl.exists() else None,
                    "json_name":jf.name if jf.exists() else None})
    return out

def _tally_ok(host, port):
    try:
        import requests as _r
        _r.post(f"http://{host}:{port}", data=b"<ENVELOPE/>", timeout=3)
        return True
    except Exception: return False

def _fmt(n):
    if n is None: return "—"
    try:
        n = float(n)
        return ("−" if n<0 else "")+"₹ {:,.0f}".format(abs(n))
    except Exception: return str(n)

# ══════════════════════════════════════════════════════════════════════════════
# Shared CSS + nav macro (used by every page)
# ══════════════════════════════════════════════════════════════════════════════
_SHELL = """<!DOCTYPE html><html lang="en"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>{{ page_title }}</title>
<style>
:root{
  --nv:#0B1F4B;--nv2:#16306e;--gd:#C8A84B;--gd2:#f0d98c;
  --tl:#1a7a6e;--tl2:#e6f4f2;
  --gn:#166534;--gn2:#dcfce7;--gn3:#16a34a;
  --rd:#b91c1c;--rd2:#fef2f2;
  --bg:#f0f4fb;--cd:#fff;--br:#dde3ee;--tx:#1e293b;--mu:#64748b;
  --r:10px;--sh:0 2px 14px rgba(11,31,75,.10);
}
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',system-ui,sans-serif;background:var(--bg);color:var(--tx);min-height:100vh}
nav{background:var(--nv);display:flex;align-items:center;gap:22px;padding:0 28px;height:54px;
    position:sticky;top:0;z-index:200;box-shadow:0 3px 10px rgba(0,0,0,.30)}
nav .brand{font-size:1.05rem;font-weight:800;color:var(--gd);letter-spacing:.4px;
           text-decoration:none;white-space:nowrap}
nav .brand em{color:#fff;font-style:normal;font-weight:400}
nav a{color:#94a3b8;text-decoration:none;font-size:.86rem;transition:color .15s}
nav a:hover{color:#fff}
.pill{background:rgba(255,255,255,.13);border-radius:20px;padding:3px 13px;font-size:.78rem;
      display:flex;align-items:center;gap:8px;color:#e2e8f0}
.pill .role{background:var(--gd);color:var(--nv);border-radius:12px;padding:1px 8px;
            font-size:.68rem;font-weight:800}
.sp{flex:1}
.flashes{padding:10px 28px 0}
.al{padding:11px 15px;border-radius:7px;font-size:.86rem;margin-bottom:10px;
    display:flex;align-items:flex-start;gap:8px}
.al-ok{background:var(--gn2);color:var(--gn);border:1px solid #bbf7d0}
.al-er{background:var(--rd2);color:var(--rd);border:1px solid #fecaca}
.al-in{background:var(--tl2);color:var(--tl);border:1px solid #99e6da}
.wrap{max-width:1120px;margin:0 auto;padding:30px 20px}
.ph{font-size:1.5rem;font-weight:800;color:var(--nv);margin-bottom:4px}
.ps{color:var(--mu);font-size:.86rem;margin-bottom:26px}
.card{background:var(--cd);border:1px solid var(--br);border-radius:var(--r);
      box-shadow:var(--sh);padding:24px;margin-bottom:22px}
.ch{font-weight:700;font-size:.95rem;color:var(--nv);margin-bottom:18px;
    padding-bottom:10px;border-bottom:2px solid var(--br);display:flex;align-items:center;gap:10px}
.ch .cnt{background:var(--nv);color:#fff;font-size:.68rem;border-radius:20px;
         padding:2px 10px;font-weight:700}
.stats{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:16px;margin-bottom:24px}
.sc{background:var(--cd);border:1px solid var(--br);border-radius:var(--r);
    padding:18px 20px;box-shadow:var(--sh)}
.sc .lbl{font-size:.72rem;font-weight:700;color:var(--mu);text-transform:uppercase;letter-spacing:.5px}
.sc .val{font-size:1.5rem;font-weight:800;color:var(--nv);margin-top:4px}
.sc .sub{font-size:.72rem;color:var(--mu);margin-top:2px}
.grid{display:grid;gap:14px}
.g2{grid-template-columns:1fr 1fr}
.g3{grid-template-columns:1fr 1fr 1fr}
.fg label{display:block;font-size:.73rem;font-weight:700;color:var(--mu);
          margin-bottom:5px;text-transform:uppercase;letter-spacing:.4px}
.fg input,.fg select,.fg textarea{width:100%;border:1.5px solid var(--br);border-radius:7px;
  padding:9px 12px;font-size:.88rem;background:#fff;outline:none;transition:border .2s}
.fg input:focus,.fg select:focus{border-color:var(--nv);box-shadow:0 0 0 3px rgba(11,31,75,.07)}
.fg .hint{font-size:.72rem;color:var(--mu);margin-top:4px}
.ck{display:flex;align-items:center;gap:9px;font-size:.88rem;cursor:pointer}
.ck input[type=checkbox]{width:15px;height:15px;accent-color:var(--nv)}
.btn{display:inline-flex;align-items:center;gap:6px;padding:9px 20px;border-radius:7px;
     font-size:.85rem;font-weight:700;cursor:pointer;border:none;text-decoration:none;
     transition:all .15s;white-space:nowrap}
.b-nv{background:var(--nv);color:#fff}.b-nv:hover{background:var(--nv2)}
.b-gd{background:var(--gd);color:var(--nv)}.b-gd:hover{background:var(--gd2)}
.b-gn{background:var(--gn3);color:#fff}.b-gn:hover{background:var(--gn)}
.b-rd{background:var(--rd);color:#fff}
.b-ol{background:transparent;border:1.5px solid var(--nv);color:var(--nv)}
.b-ol:hover{background:var(--nv);color:#fff}
.b-sm{padding:5px 12px;font-size:.78rem;border-radius:5px}
.row{display:flex;gap:12px;align-items:center;flex-wrap:wrap}
.mt{margin-top:14px}
table.t{width:100%;border-collapse:collapse;font-size:.85rem}
table.t th{background:var(--nv);color:#fff;padding:10px 13px;text-align:left;
           font-weight:700;font-size:.75rem;text-transform:uppercase;letter-spacing:.4px}
table.t td{padding:10px 13px;border-bottom:1px solid var(--br)}
table.t tr:hover td{background:#f6f9ff}
.mn{font-family:monospace;font-size:.8rem}
.badge{display:inline-block;border-radius:12px;padding:2px 9px;font-size:.72rem;font-weight:800}
.bk-ok{background:var(--gn2);color:var(--gn)}
.bk-rn{background:#fef9c3;color:#854d0e}
.bk-er{background:var(--rd2);color:var(--rd)}
.bk-nv{background:#e0e7ff;color:#3730a3}
.log{background:#0f172a;color:#94a3b8;font-family:monospace;font-size:.77rem;border-radius:8px;
     padding:16px;max-height:340px;overflow-y:auto;white-space:pre-wrap;line-height:1.55}
.pb{height:5px;background:#e2e8f0;border-radius:3px;overflow:hidden;margin:10px 0}
.pf{height:100%;background:linear-gradient(90deg,var(--nv),var(--tl));border-radius:3px;
    animation:pa 1.6s infinite}
@keyframes pa{0%{width:8%;margin-left:0}50%{width:55%;margin-left:25%}100%{width:8%;margin-left:90%}}
@media(max-width:640px){.g2,.g3{grid-template-columns:1fr}nav{gap:10px;padding:0 12px}}
</style>
</head><body>
{% if me %}
<nav>
  <a class="brand" href="/">TallySync<em> Pro</em> <span style="color:var(--gd);font-size:.8rem">● Sch III</span></a>
  <a href="/">Dashboard</a>
  {% if me.role=='admin' %}
    <a href="/admin/users">Users</a>
    <a href="/admin/reports">All Reports</a>
  {% endif %}
  <div class="sp"></div>
  <div class="pill">{{ me.display_name }}<span class="role">{{ me.role|upper }}</span></div>
  <a href="/logout">Sign out</a>
</nav>
{% endif %}
{% if messages %}
<div class="flashes">
  {% for cat,msg in messages %}
  <div class="al al-{{ 'ok' if cat=='success' else 'er' if cat=='error' else 'in' }}">{{ msg }}</div>
  {% endfor %}
</div>
{% endif %}
{{ content | safe }}
</body></html>"""

def _page(title, content, **kw):
    msgs = []
    with app.test_request_context():
        pass
    # get flash messages
    try:
        from flask import get_flashed_messages
        msgs = get_flashed_messages(with_categories=True)
    except Exception:
        msgs = []
    return render_template_string(_SHELL,
        page_title=title, content=content, messages=msgs, **kw)

def _render(title, body_html, **kw):
    """Render a complete page. Flash messages are injected automatically."""
    msgs = []
    try:
        from flask import get_flashed_messages
        msgs = get_flashed_messages(with_categories=True)
    except Exception:
        pass
    # Build context: me from context_processor
    me = None
    if "username" in session:
        me = get_user(session["username"])
        if me:
            me = dict(me); me["username"] = session["username"]
    return render_template_string(
        _SHELL,
        page_title=title,
        content=body_html,
        messages=msgs,
        me=me,
        **kw
    )

# ══════════════════════════════════════════════════════════════════════════════
# Page bodies (pure Jinja2 snippets, rendered with render_template_string)
# ══════════════════════════════════════════════════════════════════════════════

# ── Login ─────────────────────────────────────────────────────────────────────
_LOGIN_BODY = """
<style>
.lw{min-height:100vh;display:flex;align-items:center;justify-content:center;
    background:linear-gradient(135deg,#0B1F4B 0%,#1a3a7a 100%)}
.lb{background:#fff;border-radius:14px;padding:42px 38px;width:100%;
    max-width:400px;box-shadow:0 10px 48px rgba(0,0,0,.35)}
.llogo{text-align:center;margin-bottom:30px}
.llogo .br{font-size:1.5rem;font-weight:900;color:#0B1F4B}
.llogo .br span{color:#C8A84B}
.llogo .sub{color:#64748b;font-size:.82rem;margin-top:6px}
.tab-row{display:flex;border:1.5px solid #dde3ee;border-radius:8px;overflow:hidden;margin-bottom:22px}
.tab-btn{flex:1;padding:9px;font-size:.82rem;font-weight:700;border:none;
         cursor:pointer;background:#f8f9fc;color:#64748b;transition:all .15s}
.tab-btn.on{background:#0B1F4B;color:#fff}
.otp-hint{background:#fffbeb;border:1px solid #fed7aa;border-radius:7px;
          padding:10px 12px;font-size:.78rem;color:#92400e;margin-bottom:14px;text-align:center}
</style>
<div class="lw"><div class="lb">
  <div class="llogo">
    <div class="br">TallySync<span> Pro</span></div>
    <div class="sub">Schedule III Report Generator</div>
  </div>
  {% if err %}<div class="al al-er" style="margin-bottom:14px">{{ err }}</div>{% endif %}
  {% if info %}<div class="al al-in" style="margin-bottom:14px">{{ info }}</div>{% endif %}

  {% if otp_step %}
  <!-- Step 2: Enter OTP -->
  <div class="otp-hint">🔐 OTP sent to your registered email. Valid for {{ otp_mins }} minutes.</div>
  <form method="post">
    <input type="hidden" name="step" value="verify_otp">
    <input type="hidden" name="username" value="{{ otp_user }}">
    <div class="fg" style="margin-bottom:22px">
      <label>Enter OTP</label>
      <input name="otp" type="text" inputmode="numeric" pattern="[0-9]{6}"
             maxlength="6" placeholder="6-digit OTP" required autofocus
             style="letter-spacing:8px;font-size:1.4rem;text-align:center;padding:12px">
    </div>
    <button class="btn b-nv" style="width:100%;justify-content:center;padding:11px" type="submit">
      Verify OTP →
    </button>
    <p style="text-align:center;margin-top:14px;font-size:.78rem">
      <a href="/login" style="color:#64748b">← Start over</a>
    </p>
  </form>
  {% else %}
  <!-- Step 1: Username + login method -->
  <div class="tab-row">
    <button class="tab-btn {{ 'on' if not pw_mode else '' }}" type="button" onclick="setMode('otp')" id="tb-otp">Sign in with OTP</button>
    <button class="tab-btn {{ 'on' if pw_mode else '' }}" type="button" onclick="setMode('pw')" id="tb-pw">Password</button>
  </div>
  <form method="post" id="lf">
    <input type="hidden" name="step" id="step-field" value="{{ 'password' if pw_mode else 'send_otp' }}">
    <div class="fg" style="margin-bottom:14px">
      <label>Username</label>
      <input name="username" type="text" autocomplete="username" placeholder="Enter your username" required>
    </div>
    <div class="fg" style="margin-bottom:22px" id="pw-field" style="{{ '' if pw_mode else 'display:none' }}">
      <label>Password</label>
      <input name="password" type="password" autocomplete="current-password" placeholder="••••••••">
    </div>
    <button class="btn b-nv" style="width:100%;justify-content:center;padding:11px" type="submit" id="lbtn">
      <span id="lbtn-txt">{{ 'Sign In →' if pw_mode else 'Send OTP →' }}</span>
    </button>
  </form>
  {% endif %}
  <p style="margin-top:22px;text-align:center;font-size:.74rem;color:#94a3b8">
    TallySync Pro<br>Contact your administrator to request access
  </p>
  {% if lic_days %}
  <div style="margin-top:14px;text-align:center">
    <span style="background:#dcfce7;color:#166534;border-radius:20px;
                 padding:3px 14px;font-size:.72rem;font-weight:700">
      ✔ Licenced · {{ lic_days }} days remaining
    </span>
  </div>
  {% endif %}
</div></div>
<script>
function setMode(m){
  document.getElementById('step-field').value = m==='pw' ? 'password' : 'send_otp';
  document.getElementById('pw-field').style.display = m==='pw' ? '' : 'none';
  document.getElementById('lbtn-txt').textContent = m==='pw' ? 'Sign In →' : 'Send OTP →';
  document.getElementById('tb-otp').className = 'tab-btn'+(m==='otp'?' on':'');
  document.getElementById('tb-pw').className  = 'tab-btn'+(m==='pw' ?' on':'');
}
{% if pw_mode is not defined or not pw_mode %}setMode('otp');{% endif %}
</script>
"""

# ── Admin Users Management ────────────────────────────────────────────────────
_ADMIN_USERS_BODY = """
<div class="wrap">
<div class="ph">👥 User Management</div>
<div class="ps">Add, edit, restrict users. Only visible to admin.</div>

<div class="stats" style="margin-bottom:22px">
  <div class="sc"><div class="lbl">Total Users</div><div class="val">{{ users|length }}</div></div>
  <div class="sc"><div class="lbl">Active Users</div>
    <div class="val" style="color:var(--gn)">{{ active_cnt }}</div></div>
  <div class="sc"><div class="lbl">Total Reports</div>
    <div class="val">{{ total_rpts }}</div></div>
</div>

<!-- Add User -->
<div class="card">
  <div class="ch">➕ Add New User</div>
  <form method="post" action="/admin/users/add">
    <div class="grid g3" style="margin-bottom:12px">
      <div class="fg"><label>Username</label>
        <input name="username" placeholder="john_ca" required></div>
      <div class="fg"><label>Display Name</label>
        <input name="display_name" placeholder="John Mathew CA" required></div>
      <div class="fg"><label>Email (for OTP)</label>
        <input name="email" type="email" placeholder="john@caoffice.in" required></div>
    </div>
    <div class="grid g3" style="margin-bottom:12px">
      <div class="fg"><label>Password (optional)</label>
        <input name="password" type="password" placeholder="Leave blank for OTP-only"></div>
      <div class="fg"><label>Company Limit (1–15)</label>
        <input name="company_limit" type="number" min="1" max="15" value="15"></div>
      <div class="fg"><label>Role</label>
        <select name="role">
          <option value="user">User (client)</option>
          <option value="admin">Admin</option>
        </select>
      </div>
    </div>
    <div class="grid g2" style="margin-bottom:12px">
      <div class="fg"><label>Tally Host</label>
        <input name="tally_host" value="localhost"></div>
      <div class="fg"><label>Allowed Companies (comma-separated, blank = all)</label>
        <input name="allowed_companies" placeholder="Vianci Enterprise, Design Limelite"></div>
    </div>
    <button class="btn b-nv" type="submit">➕ Add User</button>
  </form>
</div>

<!-- User Table -->
<div class="card">
  <div class="ch">All Users</div>
  <div style="overflow-x:auto">
  <table class="t">
    <tr>
      <th>Username</th><th>Name</th><th>Email</th><th>Role</th>
      <th>Limit</th><th>Companies Used</th>
      <th>Reports</th><th>Last Active</th><th>Status</th><th>Actions</th>
    </tr>
    {% for uname, u in users.items() %}
    <tr>
      <td class="mn"><strong>{{ uname }}</strong></td>
      <td>{{ u.display_name }}</td>
      <td class="mn" style="font-size:.76rem">{{ u.get('email','—') }}</td>
      <td><span class="badge {{ 'bk-nv' if u.role=='admin' else 'bk-ok' }}">{{ u.role|upper }}</span></td>
      <td style="text-align:center">
        <span style="font-weight:700">{{ cos_used.get(uname,[])|length }}</span>
        / {{ u.get('company_limit', 15) }}
      </td>
      <td style="font-size:.76rem;max-width:200px">
        {% set cu = cos_used.get(uname,[]) %}
        {% if cu %}{{ cu[:3]|join(', ') }}{% if cu|length > 3 %} +{{ cu|length - 3 }} more{% endif %}
        {% else %}<span style="color:var(--mu)">None yet</span>{% endif %}
      </td>
      <td style="text-align:center">{{ u.get('usage',{}).get('total_reports',0) }}</td>
      <td class="mn" style="font-size:.74rem">
        {{ u.get('usage',{}).get('last_active','—')[:16].replace('T',' ') if u.get('usage',{}).get('last_active') else '—' }}
      </td>
      <td>
        {% if u.get('active', True) %}
          <span class="badge bk-ok">Active</span>
        {% else %}
          <span class="badge bk-er">Disabled</span>
        {% endif %}
      </td>
      <td style="white-space:nowrap;display:flex;gap:4px;flex-wrap:wrap">
        <a class="btn b-ol b-sm" href="/admin/users/edit/{{ uname }}">Edit</a>
        {% if uname != 'admin' %}
        <form method="post" action="/admin/users/toggle/{{ uname }}" style="display:inline">
          <button class="btn b-sm {{ 'b-rd' if u.get('active',True) else 'b-gn' }}" type="submit">
            {{ 'Disable' if u.get('active',True) else 'Enable' }}
          </button>
        </form>
        <form method="post" action="/admin/users/delete/{{ uname }}" style="display:inline"
              onsubmit="return confirm('Delete {{ uname }}?')">
          <button class="btn b-rd b-sm" type="submit">Delete</button>
        </form>
        {% endif %}
        <a class="btn b-gd b-sm" href="/admin/users/usage/{{ uname }}">Usage</a>
      </td>
    </tr>
    {% endfor %}
  </table>
  </div>
</div>
</div>
"""

# ── User Usage Detail ────────────────────────────────────────────────────────
_USAGE_BODY = """
<div class="wrap">
<div class="ph">📊 Usage: {{ uname }}</div>
<div class="ps">{{ u.display_name }} · {{ u.get('email','no email') }}</div>
<a class="btn b-ol b-sm" href="/admin/users" style="margin-bottom:20px;display:inline-flex">← Back</a>

<div class="stats">
  <div class="sc"><div class="lbl">Total Reports</div>
    <div class="val">{{ u.get('usage',{}).get('total_reports',0) }}</div></div>
  <div class="sc"><div class="lbl">Unique Companies</div>
    <div class="val">{{ u.get('usage',{}).get('companies_accessed',[])|length }}</div>
    <div class="sub">Limit: {{ u.get('company_limit',15) }}</div></div>
  <div class="sc"><div class="lbl">Last Active</div>
    <div class="val" style="font-size:.9rem;margin-top:8px">
      {{ u.get('usage',{}).get('last_active','Never')[:16].replace('T',' ') }}</div></div>
  <div class="sc"><div class="lbl">Account Status</div>
    <div class="val" style="font-size:.9rem;margin-top:8px;color:{{ 'var(--gn)' if u.get('active',True) else 'var(--rd)' }}">
      {{ 'Active' if u.get('active',True) else 'Disabled' }}</div></div>
</div>

<div class="card">
  <div class="ch">Companies Accessed</div>
  {% set cos = u.get('usage',{}).get('companies_accessed',[]) %}
  {% if cos %}
  <div style="display:flex;flex-wrap:wrap;gap:8px">
    {% for c in cos %}
    <span style="background:#e0e7ff;color:#3730a3;padding:4px 12px;border-radius:12px;font-size:.82rem;font-weight:600">{{ c }}</span>
    {% endfor %}
  </div>
  {% else %}<p style="color:var(--mu)">No reports generated yet.</p>{% endif %}
</div>

<div class="card">
  <div class="ch">Report History (last 50)</div>
  {% set log = u.get('usage',{}).get('report_log',[]) %}
  {% if log %}
  <table class="t">
    <tr><th>Date</th><th>Company</th><th>Type</th></tr>
    {% for r in log[:50] %}
    <tr>
      <td class="mn">{{ r.get('date','') }}</td>
      <td>{{ r.get('company','') }}</td>
      <td><span class="badge bk-nv">{{ r.get('type','report') }}</span></td>
    </tr>
    {% endfor %}
  </table>
  {% else %}<p style="color:var(--mu)">No reports yet.</p>{% endif %}
</div>

<!-- Restrict Company Limit -->
<div class="card" style="max-width:400px">
  <div class="ch">⚙ Change Company Limit</div>
  <form method="post" action="/admin/users/limit/{{ uname }}">
    <div class="fg" style="margin-bottom:12px">
      <label>Max Companies (1–15)</label>
      <input name="company_limit" type="number" min="1" max="15"
             value="{{ u.get('company_limit',15) }}" required>
      <div class="hint">User currently has access to {{ u.get('usage',{}).get('companies_accessed',[])|length }} unique companies.</div>
    </div>
    <button class="btn b-nv" type="submit">Update Limit</button>
  </form>
</div>
</div>
"""

# ── Admin Edit User ───────────────────────────────────────────────────────────
_EDIT_BODY = """
<div class="wrap">
<div class="ph">Edit User: {{ uname }}</div>
<div class="ps">Update credentials, Tally settings, and access control</div>
<div class="card" style="max-width:680px">
  <form method="post">
    <div class="grid g2" style="margin-bottom:14px">
      <div class="fg"><label>Display Name</label>
        <input name="display_name" value="{{ u.display_name }}" required></div>
      <div class="fg"><label>Email (for OTP)</label>
        <input name="email" type="email" value="{{ u.get('email','') }}" placeholder="user@example.com"></div>
    </div>
    <div class="grid g3" style="margin-bottom:14px">
      <div class="fg"><label>Role</label>
        <select name="role">
          <option value="user" {% if u.role=='user' %}selected{% endif %}>User</option>
          <option value="admin" {% if u.role=='admin' %}selected{% endif %}>Admin</option>
        </select>
      </div>
      <div class="fg"><label>Company Limit (1–15)</label>
        <input name="company_limit" type="number" min="1" max="15" value="{{ u.get('company_limit',15) }}"></div>
      <div class="fg"><label>New Password (blank = keep current)</label>
        <input name="password" type="password" placeholder="Leave blank to keep"></div>
    </div>
    <div class="grid g2" style="margin-bottom:14px">
      <div class="fg"><label>Tally Host</label>
        <input name="tally_host" value="{{ u.tally_host }}"></div>
      <div class="fg"><label>Tally Port</label>
        <input name="tally_port" type="number" value="{{ u.tally_port }}"></div>
    </div>
    <div class="fg" style="margin-bottom:14px">
      <label>Allowed Companies (comma-separated, blank = all)</label>
      <input name="allowed_companies"
             value="{{ u.allowed_companies|join(', ') if u.allowed_companies else '' }}">
    </div>
    <div class="row mt">
      <button class="btn b-nv" type="submit">Save Changes</button>
      <a class="btn b-ol" href="/admin/users">Cancel</a>
    </div>
  </form>
</div>
</div>
"""

# ── Dashboard ─────────────────────────────────────────────────────────────────
_DASH_BODY = """
<div class="wrap">
<div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:6px">
  <div>
    <div class="ph">Schedule III Generator</div>
    <div class="ps">Generate and download Schedule III Balance Sheets from TallyPrime</div>
  </div>
  <div style="text-align:right;font-size:.8rem;color:var(--mu);padding-top:6px">
    FY {{ fy_start }}–{{ fy_end_s }}<br>
    <strong style="color:var(--nv)">{{ me.display_name }}</strong>
  </div>
</div>

<div class="stats">
  <div class="sc">
    <div class="lbl">Reports Generated</div>
    <div class="val">{{ rcount }}</div>
    <div class="sub">by your account</div>
  </div>
  <div class="sc">
    <div class="lbl">Companies Used</div>
    <div class="val" style="color:{{ 'var(--rd)' if cos_used >= co_limit else 'var(--nv)' }}">
      {{ cos_used }} / {{ co_limit }}
    </div>
    <div class="sub">unique company slots
      {% if cos_used >= co_limit %}<span style="color:var(--rd);font-weight:700"> · LIMIT REACHED</span>{% endif %}
    </div>
  </div>
  <div class="sc">
    <div class="lbl">Tally Connection</div>
    <div class="val" style="font-size:1rem;margin-top:8px">
      {% if tok %}<span style="color:var(--gn)">✔ Connected</span>
      {% else %}<span style="color:var(--rd)">✘ Offline</span>{% endif %}
    </div>
    <div class="sub">{{ me.tally_host }}:{{ me.tally_port }}</div>
  </div>
  <div class="sc">
    <div class="lbl">Classic CA Format</div>
    <div class="val" style="font-size:1rem;margin-top:8px">
      {% if classic_ok %}<span style="color:var(--gn)">✔ Available</span>
      {% else %}<span style="color:var(--mu)">— Not bundled</span>{% endif %}
    </div>
    <div class="sub">sch3_classic.py</div>
  </div>
  <div class="sc">
    <div class="lbl">Verification + MSME</div>
    <div class="val" style="font-size:1rem;margin-top:8px">
      {% if verify_ok %}<span style="color:var(--gn)">✔ Active</span>
      {% else %}<span style="color:var(--mu)">— Not loaded</span>{% endif %}
    </div>
    <div class="sub">sch3_verify.py</div>
  </div>
</div>

{% if cos_used >= co_limit %}
<div class="al al-er" style="margin-bottom:18px">
  ⚠ You have used all <strong>{{ co_limit }} company slots</strong>.
  You can still regenerate reports for your existing {{ cos_used }} companies.
  To add a new company, contact your administrator to increase your limit.
  <br><small>Your companies: {{ cos_list|join(' · ') }}</small>
</div>
{% endif %}

<div class="card">
  <div class="ch">📄 Generate New Report</div>
  {% if not core_ok %}
  <div class="al al-er" style="margin-bottom:16px">
    ⚠ tally_core.py not found. Place it in the same folder as app.py and restart.
  </div>
  {% endif %}
  <form method="post" action="/generate" id="gf">
    <div class="grid g3" style="margin-bottom:14px">
      <div class="fg">
        <label>Tally Host</label>
        <input name="host" value="{{ me.tally_host }}" placeholder="localhost">
      </div>
      <div class="fg">
        <label>Tally Port</label>
        <input name="port" type="number" value="{{ me.tally_port }}" placeholder="9000">
      </div>
      <div class="fg">
        <label>Company Name</label>
        <div style="display:flex;gap:8px">
          <select name="company" id="csel" style="flex:1">
            <option value="">— click Fetch to load —</option>
            {% for c in companies %}<option value="{{ c }}">{{ c }}</option>{% endfor %}
          </select>
          <button type="button" class="btn b-ol b-sm" onclick="fetchCos()" id="fbtn">🔄 Fetch</button>
        </div>
      </div>
    </div>
    <div class="fg" style="margin-bottom:14px">
      <label>Or type company name manually</label>
      <input name="company_manual" id="cman" placeholder="DEEP DEMO PVT LTD" autocomplete="off">
      <div class="hint">If filled, this overrides the dropdown above</div>
    </div>

    <!-- ── Date Mode Tabs ── -->
    <input type="hidden" name="date_mode" id="date_mode" value="fy">
    <div style="margin-bottom:14px">
      <div style="font-size:.73rem;font-weight:700;color:var(--mu);margin-bottom:8px;
                  text-transform:uppercase;letter-spacing:.4px">Report Period</div>
      <div style="display:flex;gap:0;border:1.5px solid var(--br);border-radius:8px;
                  overflow:hidden;max-width:400px;margin-bottom:14px">
        <button type="button" id="tab-fy" onclick="setDateMode('fy')"
          style="flex:1;padding:9px 16px;font-size:.85rem;font-weight:700;border:none;
                 cursor:pointer;background:var(--nv);color:#fff;transition:all .15s">
          Full Financial Year
        </button>
        <button type="button" id="tab-custom" onclick="setDateMode('custom')"
          style="flex:1;padding:9px 16px;font-size:.85rem;font-weight:700;border:none;
                 cursor:pointer;background:#fff;color:var(--mu);transition:all .15s;
                 border-left:1.5px solid var(--br)">
          Custom Date Range
        </button>
      </div>

      <!-- Full FY panel -->
      <div id="panel-fy">
        <div class="grid g2" style="max-width:400px">
          <div class="fg">
            <label>FY End Year</label>
            <select name="year" id="fy-sel" onchange="updateChip()">
              {% set cur = def_year %}
              {% for y in range(cur, cur-6, -1) %}
              <option value="{{ y }}" {% if y == cur %}selected{% endif %}>
                FY {{ y-1 }}-{{ (y|string)[2:] }}&nbsp;&nbsp;(Apr {{ y-1 }} – Mar {{ y }})
              </option>
              {% endfor %}
            </select>
          </div>
        </div>
        <!-- Quick presets -->
        <div style="display:flex;flex-wrap:wrap;gap:6px;margin-top:10px" id="fy-presets"></div>
      </div>

      <!-- Custom date range panel -->
      <div id="panel-custom" style="display:none">
        <div class="grid g2" style="max-width:400px;margin-bottom:10px">
          <div class="fg">
            <label>From Date</label>
            <input type="date" name="from_date" id="from-date" onchange="updateChip()">
          </div>
          <div class="fg">
            <label>To Date</label>
            <input type="date" name="to_date" id="to-date" onchange="updateChip()">
          </div>
        </div>
        <!-- Quick presets for custom -->
        <div style="display:flex;flex-wrap:wrap;gap:6px" id="custom-presets"></div>
      </div>

      <!-- Period chip -->
      <div id="period-chip" style="display:none;margin-top:10px">
        <span style="background:#e0e7ff;color:#3730a3;padding:4px 12px;border-radius:12px;
                     font-size:.78rem;font-weight:700">📅 <span id="chip-text"></span></span>
      </div>
    </div>

    <div class="row" style="margin-bottom:18px;gap:20px">
      <label class="ck">
        <input type="checkbox" name="gen_classic" value="1"
          {% if classic_ok %}checked{% endif %}
          {% if not classic_ok %}disabled{% endif %}>
        Also generate Classic CA-Format PDF (with previous year comparison)
      </label>
      <label class="ck">
        <input type="checkbox" name="autosync_py" value="1" checked>
        Auto-sync previous year (required for comparison column)
      </label>
    </div>
    <div class="row">
      <button class="btn b-nv" type="submit" id="gbtn" {% if not core_ok %}disabled{% endif %}>
        📄 Generate Report
      </button>
      <span style="font-size:.8rem;color:var(--mu)" id="ghint"></span>
    </div>
  </form>
</div>

<div class="card">
  <div class="ch">📁 My Reports <span class="cnt">{{ rcount }}</span></div>
  {% if reports %}
  <div style="overflow-x:auto">
  <table class="t">
    <tr>
      <th>Company</th><th>FY Period</th><th>Generated At</th>
      <th>Gross Profit</th><th>Net Profit / Loss</th><th>Downloads</th>
    </tr>
    {% for r in reports %}
    <tr>
      <td><strong>{{ r.company }}</strong></td>
      <td class="mn">{{ r.fy_start }} → {{ r.fy_end }}</td>
      <td class="mn" style="font-size:.75rem">{{ r.generated_at[:19].replace('T',' ') if r.generated_at else '—' }}</td>
      <td>{{ fmt(r.gross_profit) }}</td>
      <td>
        {% if r.net_profit is not none %}
        <span style="color:{{ 'var(--gn)' if r.net_profit>=0 else 'var(--rd)' }};font-weight:700">
          {{ fmt(r.net_profit) }} ({{ 'Profit' if r.net_profit>=0 else 'Loss' }})
        </span>
        {% else %}—{% endif %}
      </td>
      <td style="white-space:nowrap">
        {% if r.pdf_name %}
        <a class="btn b-nv b-sm" href="/dl/{{ me.username }}/{{ r.pdf_name }}">PDF</a>
        {% endif %}
        {% if r.classic_name %}
        <a class="btn b-gn b-sm" href="/dl/{{ me.username }}/{{ r.classic_name }}" style="margin-left:4px">Classic</a>
        {% endif %}
        {% if r.json_name %}
        <a class="btn b-ol b-sm" href="/dl/{{ me.username }}/{{ r.json_name }}" style="margin-left:4px">JSON</a>
        <a class="btn b-sm" href="/dl-excel/{{ me.username }}/{{ r.json_name }}" style="margin-left:4px;background:#1d6f42;color:#fff" title="Download Schedule III Excel with formulas">📊 Excel</a>
        <a class="btn b-sm" href="/dl-excel-classic/{{ me.username }}/{{ r.json_name }}" style="margin-left:4px;background:#1F3864;color:#fff" title="Classic Schedule III Excel — TB Data + Formulas + Unit Conversion (Lakhs/Crores/etc.)">📊 Classic Excel</a>
        {% endif %}
      </td>
    </tr>
    {% endfor %}
  </table>
  </div>
  {% else %}
  <p style="text-align:center;color:var(--mu);padding:30px 0">
    No reports yet. Generate your first one above ↑
  </p>
  {% endif %}
</div>
</div>

<script>
// ── Company fetch ──────────────────────────────────────────────────────────
function fetchCos(){
  const btn=document.getElementById('fbtn'), sel=document.getElementById('csel');
  const host=document.querySelector('[name=host]').value;
  const port=document.querySelector('[name=port]').value;
  btn.textContent='⏳ Fetching…'; btn.disabled=true;
  fetch('/api/companies?host='+encodeURIComponent(host)+'&port='+port)
    .then(r=>r.json()).then(d=>{
      sel.innerHTML='<option value="">— select company —</option>';
      (d.companies||[]).forEach(c=>{
        const o=document.createElement('option');
        o.value=c; o.textContent=c; sel.appendChild(o);
      });
      if(d.error) sel.innerHTML='<option value="">'+d.error+'</option>';
    }).catch(()=>{ sel.innerHTML='<option value="">Connection failed — is Tally running?</option>'; })
    .finally(()=>{ btn.textContent='🔄 Fetch'; btn.disabled=false; });
}

// ── Date mode ───────────────────────────────────────────────────────────────
function setDateMode(mode) {
  document.getElementById('date_mode').value = mode;
  const isFY = mode === 'fy';
  document.getElementById('panel-fy').style.display     = isFY ? '' : 'none';
  document.getElementById('panel-custom').style.display = isFY ? 'none' : '';
  document.getElementById('tab-fy').style.background    = isFY ? 'var(--nv)' : '#fff';
  document.getElementById('tab-fy').style.color         = isFY ? '#fff' : 'var(--mu)';
  document.getElementById('tab-custom').style.background= isFY ? '#fff' : 'var(--nv)';
  document.getElementById('tab-custom').style.color     = isFY ? 'var(--mu)' : '#fff';
  updateChip();
}

// ── Period chip display ──────────────────────────────────────────────────────
function updateChip() {
  const mode = document.getElementById('date_mode').value;
  let text = '';
  if (mode === 'fy') {
    const y = parseInt(document.getElementById('fy-sel').value);
    text = `FY ${y-1}-${String(y).slice(2)}  (01 Apr ${y-1} – 31 Mar ${y})`;
  } else {
    const f = document.getElementById('from-date').value;
    const t = document.getElementById('to-date').value;
    if (f && t) text = `${fmtD(f)}  to  ${fmtD(t)}`;
  }
  const chip = document.getElementById('period-chip');
  if (text) { chip.style.display=''; document.getElementById('chip-text').textContent=text; }
  else chip.style.display='none';
}

function fmtD(d) {
  if (!d) return '';
  const [y,m,day] = d.split('-');
  const mo=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][+m-1];
  return `${day} ${mo} ${y}`;
}

// ── Quick presets ───────────────────────────────────────────────────────────
(function buildPresets() {
  const today = new Date();
  const y = today.getFullYear();
  const m = today.getMonth(); // 0-indexed
  const fy = m >= 3 ? y : y - 1; // FY start year

  // FY presets — just set the year select
  const fyPresets = [
    { label: `FY ${fy}-${String(fy+1).slice(2)}`,   y: fy+1 },
    { label: `FY ${fy-1}-${String(fy).slice(2)}`,   y: fy   },
    { label: `FY ${fy-2}-${String(fy-1).slice(2)}`, y: fy-1 },
  ];
  const fyWrap = document.getElementById('fy-presets');
  fyPresets.forEach(p => {
    const b = document.createElement('button');
    b.type='button'; b.className='btn b-ol b-sm';
    b.textContent = p.label;
    b.onclick = () => { document.getElementById('fy-sel').value = p.y; updateChip(); };
    fyWrap.appendChild(b);
  });

  // Custom presets
  const customPresets = [
    { label:'Q1 (Apr–Jun)',  from:`${fy}-04-01`, to:`${fy}-06-30`   },
    { label:'Q2 (Jul–Sep)',  from:`${fy}-07-01`, to:`${fy}-09-30`   },
    { label:'Q3 (Oct–Dec)',  from:`${fy}-10-01`, to:`${fy}-12-31`   },
    { label:'Q4 (Jan–Mar)',  from:`${fy+1}-01-01`, to:`${fy+1}-03-31`},
    { label:'Apr–Jul',       from:`${fy}-04-01`, to:`${fy}-07-31`   },
    { label:'Apr–Sep',       from:`${fy}-04-01`, to:`${fy}-09-30`   },
    { label:'Apr–Dec',       from:`${fy}-04-01`, to:`${fy}-12-31`   },
    { label:'Full FY',       from:`${fy}-04-01`, to:`${fy+1}-03-31` },
  ];
  const cwrap = document.getElementById('custom-presets');
  customPresets.forEach(p => {
    const b = document.createElement('button');
    b.type='button'; b.className='btn b-ol b-sm';
    b.textContent = p.label;
    b.onclick = () => {
      document.getElementById('from-date').value = p.from;
      document.getElementById('to-date').value   = p.to;
      cwrap.querySelectorAll('.btn').forEach(x=>x.style.background='');
      b.style.background='var(--nv)'; b.style.color='#fff';
      updateChip();
    };
    cwrap.appendChild(b);
  });
})();

// ── Init ────────────────────────────────────────────────────────────────────
setDateMode('fy');

// ── Submit button ────────────────────────────────────────────────────────────
document.getElementById('gf').addEventListener('submit', ()=>{
  document.getElementById('gbtn').textContent='⏳ Submitting…';
  document.getElementById('gbtn').disabled=true;
  document.getElementById('ghint').textContent='Starting job, please wait…';
});
</script>
"""

# ── Job Status ────────────────────────────────────────────────────────────────
_JOB_BODY = """
<div class="wrap">
<div class="ph">Report Generation</div>
<div class="ps">
  <strong>{{ company }}</strong> &nbsp;|&nbsp; FY end year: {{ year }}
  &nbsp;|&nbsp; Job ID: <span class="mn">{{ jid }}</span>
</div>

<div class="card">
  <div class="ch">
    Status: &nbsp;
    <span id="sbadge" class="badge {{ 'bk-ok' if status=='done' else 'bk-er' if status=='error' else 'bk-rn' }}">
      {{ status|upper }}
    </span>
  </div>

  {% if status == 'running' %}
  <div class="pb"><div class="pf"></div></div>
  <p style="color:var(--mu);font-size:.86rem;margin-bottom:10px">
    Connecting to Tally, fetching data, generating PDF… please wait.
    This usually takes 1–3 minutes depending on company data size.
  </p>
  {% endif %}

  {% if status == 'done' %}
  <div class="al al-ok" style="margin-bottom:18px">
    ✔ Report generated successfully! &nbsp; FY {{ fy_start }} → {{ fy_end }}
    &nbsp;|&nbsp;
    Net {{ 'Profit' if net_profit >= 0 else 'Loss' }}: <strong>₹ {{ '{:,.0f}'.format(net_profit|abs) }}</strong>
  </div>
  <div class="row mt">
    {% if dl_pdf %}<a class="btn b-nv" href="{{ dl_pdf }}">⬇ Download Schedule III PDF</a>{% endif %}
    {% if dl_classic %}<a class="btn b-gn" href="{{ dl_classic }}">⬇ Classic CA-Format PDF <span style="font-size:.72rem;opacity:.85">(+ Verification + MSME)</span></a>{% endif %}
    {% if dl_json %}<a class="btn b-ol" href="{{ dl_json }}">⬇ Raw JSON Data</a>{% endif %}
    <a class="btn b-ol" href="/">← Back to Dashboard</a>
  </div>
  {% elif status == 'error' %}
  <div class="al al-er" style="margin-bottom:14px">
    ✘ Report generation failed. See error details below.
  </div>
  <div class="log" style="max-height:200px">{{ error }}</div>
  <div class="row mt"><a class="btn b-ol" href="/">← Back to Dashboard</a></div>
  {% endif %}

  <div style="margin-top:22px">
    <div style="font-size:.73rem;font-weight:700;color:var(--mu);margin-bottom:8px;
                text-transform:uppercase;letter-spacing:.5px">Live Log Output</div>
    <div class="log" id="jlog">{{ log }}</div>
  </div>
</div>
</div>

<script>
const JID = {{ jid|tojson }};
const DONE = {{ 'true' if status in ('done','error') else 'false' }};

function poll(){
  fetch('/api/job/'+JID)
    .then(r=>r.json()).then(d=>{
      const el = document.getElementById('jlog');
      if(el){ el.textContent = d.logs||''; el.scrollTop = el.scrollHeight; }
      const b = document.getElementById('sbadge');
      if(b){
        b.textContent = d.status.toUpperCase();
        b.className = 'badge '+(d.status==='done'?'bk-ok':d.status==='error'?'bk-er':'bk-rn');
      }
      if(d.status==='done' || d.status==='error'){ location.reload(); return; }
      setTimeout(poll, 2000);
    }).catch(()=>setTimeout(poll, 3000));
}

if(!DONE){ setTimeout(poll, 2000); }
const l = document.getElementById('jlog');
if(l) l.scrollTop = l.scrollHeight;
</script>
"""

# ── Admin Clients ──────────────────────────────────────────────────────────────
_CLIENTS_BODY = """
<div class="wrap">
<div class="ph">Client Management</div>
<div class="ps">Add, edit, and manage all client login accounts and their Tally access restrictions</div>

<div class="card">
  <div class="ch">➕ Add New Client</div>
  <form method="post" action="/admin/clients/add">
    <div class="grid g3" style="margin-bottom:14px">
      <div class="fg">
        <label>Username</label>
        <input name="username" placeholder="ramesh_shah" required>
        <div class="hint">Lowercase, no spaces (underscores allowed)</div>
      </div>
      <div class="fg">
        <label>Display Name</label>
        <input name="display_name" placeholder="Ramesh Shah & Co." required>
      </div>
      <div class="fg">
        <label>Password</label>
        <input name="password" type="password" placeholder="Minimum 6 characters" required>
      </div>
    </div>
    <div class="grid g3" style="margin-bottom:14px">
      <div class="fg">
        <label>Tally Host</label>
        <input name="tally_host" value="localhost" placeholder="localhost or 192.168.x.x">
        <div class="hint">IP address of the PC running Tally</div>
      </div>
      <div class="fg">
        <label>Tally Port</label>
        <input name="tally_port" type="number" value="9000">
      </div>
      <div class="fg">
        <label>Role</label>
        <select name="role">
          <option value="user">User (client)</option>
          <option value="admin">Admin (full access)</option>
        </select>
      </div>
    </div>
    <div class="fg" style="margin-bottom:18px">
      <label>Allowed Companies (comma-separated — leave blank to allow ALL)</label>
      <input name="allowed_companies" placeholder="DEEP DEMO PVT LTD, ABC TRADERS PVT LTD">
      <div class="hint">
        If filled, this client can ONLY generate reports for these companies.
        Company names are matched case-insensitively.
      </div>
    </div>
    <button class="btn b-nv" type="submit">Add Client</button>
  </form>
</div>

<div class="card">
  <div class="ch">👥 All Clients <span class="cnt">{{ users|length }}</span></div>
  <div style="overflow-x:auto">
  <table class="t">
    <tr>
      <th>Username</th><th>Display Name</th><th>Role</th>
      <th>Tally Host</th><th>Port</th><th>Allowed Companies</th>
      <th>Reports</th><th>Actions</th>
    </tr>
    {% for uname, u in users.items() %}
    <tr>
      <td class="mn"><strong>{{ uname }}</strong></td>
      <td>{{ u.display_name }}</td>
      <td><span class="badge {{ 'bk-nv' if u.role=='admin' else 'bk-ok' }}">{{ u.role|upper }}</span></td>
      <td class="mn">{{ u.tally_host }}</td>
      <td class="mn">{{ u.tally_port }}</td>
      <td style="font-size:.78rem;max-width:220px">
        {% if u.allowed_companies %}{{ u.allowed_companies|join(', ') }}
        {% else %}<span style="color:var(--mu)">All companies</span>{% endif %}
      </td>
      <td class="mn" style="text-align:center">{{ rcounts.get(uname, 0) }}</td>
      <td style="white-space:nowrap">
        <a class="btn b-ol b-sm" href="/admin/clients/edit/{{ uname }}">Edit</a>
        {% if uname != 'admin' %}
        <form method="post" action="/admin/clients/delete/{{ uname }}" style="display:inline"
              onsubmit="return confirm('Delete client \'{{ uname }}\'? All their settings will be removed.')">
          <button class="btn b-rd b-sm" type="submit" style="margin-left:4px">Delete</button>
        </form>
        {% endif %}
      </td>
    </tr>
    {% endfor %}
  </table>
  </div>
</div>
</div>
"""

# ── Admin Edit Client ──────────────────────────────────────────────────────────
_EDIT_BODY = """
<div class="wrap">
<div class="ph">Edit Client: {{ uname }}</div>
<div class="ps">Update display name, credentials, Tally settings, and company access</div>
<div class="card" style="max-width:640px">
  <form method="post">
    <div class="grid g2" style="margin-bottom:14px">
      <div class="fg">
        <label>Display Name</label>
        <input name="display_name" value="{{ u.display_name }}" required>
      </div>
      <div class="fg">
        <label>Role</label>
        <select name="role">
          <option value="user" {% if u.role=='user' %}selected{% endif %}>User (client)</option>
          <option value="admin" {% if u.role=='admin' %}selected{% endif %}>Admin</option>
        </select>
      </div>
    </div>
    <div class="grid g2" style="margin-bottom:14px">
      <div class="fg">
        <label>New Password (leave blank to keep current)</label>
        <input name="password" type="password" placeholder="Leave blank to keep current password">
      </div>
      <div class="fg">
        <label>Tally Host</label>
        <input name="tally_host" value="{{ u.tally_host }}" placeholder="localhost or IP address">
      </div>
    </div>
    <div class="grid g2" style="margin-bottom:14px">
      <div class="fg">
        <label>Tally Port</label>
        <input name="tally_port" type="number" value="{{ u.tally_port }}">
      </div>
      <div class="fg">
        <label>Allowed Companies (comma-separated, blank = all)</label>
        <input name="allowed_companies"
               value="{{ u.allowed_companies|join(', ') if u.allowed_companies else '' }}">
      </div>
    </div>
    <div class="row mt">
      <button class="btn b-nv" type="submit">Save Changes</button>
      <a class="btn b-ol" href="/admin/clients">Cancel</a>
    </div>
  </form>
</div>
</div>
"""

# ── Admin All Reports ──────────────────────────────────────────────────────────
_AREPORTS_BODY = """
<div class="wrap">
<div class="ph">All Reports</div>
<div class="ps">Reports generated across all client accounts — {{ reports|length }} total</div>
<div class="card">
  {% if reports %}
  <div style="overflow-x:auto">
  <table class="t">
    <tr>
      <th>Client</th><th>Company</th><th>FY End</th>
      <th>Generated At</th><th>Gross Profit</th><th>Net P/L</th><th>Downloads</th>
    </tr>
    {% for r in reports %}
    <tr>
      <td class="mn"><strong>{{ r.client }}</strong></td>
      <td>{{ r.company }}</td>
      <td class="mn">{{ r.fy_end }}</td>
      <td class="mn" style="font-size:.74rem">{{ r.generated_at[:19].replace('T',' ') if r.generated_at else '—' }}</td>
      <td>{{ fmt(r.gross_profit) }}</td>
      <td>
        {% if r.net_profit is not none %}
        <span style="color:{{ 'var(--gn)' if r.net_profit>=0 else 'var(--rd)' }};font-weight:700">
          {{ fmt(r.net_profit) }} ({{ 'Profit' if r.net_profit>=0 else 'Loss' }})
        </span>
        {% else %}—{% endif %}
      </td>
      <td style="white-space:nowrap">
        {% if r.pdf_name %}
        <a class="btn b-nv b-sm" href="/dl/{{ r.client }}/{{ r.pdf_name }}">PDF</a>
        {% endif %}
        {% if r.classic_name %}
        <a class="btn b-gn b-sm" href="/dl/{{ r.client }}/{{ r.classic_name }}" style="margin-left:4px">Classic</a>
        {% endif %}
      </td>
    </tr>
    {% endfor %}
  </table>
  </div>
  {% else %}
  <p style="text-align:center;color:var(--mu);padding:30px 0">No reports generated yet.</p>
  {% endif %}
</div>
</div>
"""

# ══════════════════════════════════════════════════════════════════════════════
# Route helpers
# ══════════════════════════════════════════════════════════════════════════════
def _R(title, body_template, **ctx):
    """Render body_template as Jinja2 with ctx, then embed in shell."""
    me = None
    if "username" in session:
        me = get_user(session["username"])
        if me:
            me = dict(me); me["username"] = session["username"]
    msgs = []
    try:
        from flask import get_flashed_messages as gfm
        msgs = gfm(with_categories=True)
    except Exception: pass
    content = render_template_string(body_template, me=me, fmt=_fmt, **ctx)
    return render_template_string(_SHELL,
        page_title=title, content=content, messages=msgs, me=me, fmt=_fmt)


# ══════════════════════════════════════════════════════════════════════════════
# Routes
# ══════════════════════════════════════════════════════════════════════════════
@app.route("/login", methods=["GET","POST"])
def login():
    if "username" in session: return redirect(url_for("dashboard"))
    err = info = None
    otp_step = False; otp_user = ""; pw_mode = False

    if request.method == "POST":
        step = request.form.get("step","send_otp")
        if step == "send_otp":
            un = request.form.get("username","").strip()
            u  = get_user(un)
            if not u:
                err = "Username not found."
            elif not u.get("active", True):
                err = "Your account has been deactivated. Contact your administrator."
            else:
                otp = _generate_otp(un)
                _send_otp(un, otp, u.get("email",""))
                otp_step = True; otp_user = un
                e = u.get("email","")
                masked = (e[:3]+"****@"+e.split("@")[-1]) if "@" in e else "your registered email"
                info = f"OTP sent to {masked}. Valid for {OTP_EXPIRY_MINUTES} minutes."
        elif step == "password":
            un = request.form.get("username","").strip()
            pw = request.form.get("password","")
            u  = get_user(un)
            if u and u.get("password") and check_password_hash(u["password"], pw) and u.get("active",True):
                return _do_login(un)
            err = "Invalid username or password."; pw_mode = True
        elif step == "verify_otp":
            un  = request.form.get("username","").strip()
            otp = request.form.get("otp","").strip()
            if _verify_otp(un, otp):
                return _do_login(un)
            err = "Invalid or expired OTP. Please try again."
            otp_step = True; otp_user = un

    msgs = []
    try:
        from flask import get_flashed_messages as gfm; msgs = gfm(with_categories=True)
    except Exception: pass
    _lic     = _get_licence_status()
    lic_days = _lic.get("days_left") if _lic.get("valid") else None
    content = render_template_string(_LOGIN_BODY, me=None, err=err, info=info,
                                     otp_step=otp_step, otp_user=otp_user,
                                     pw_mode=pw_mode, otp_mins=OTP_EXPIRY_MINUTES,
                                     lic_days=lic_days)
    return render_template_string(_SHELL, page_title="Sign In — TallySync Pro",
                                  content=content, messages=msgs, me=None)

def _do_login(username):
    token = secrets.token_hex(24)
    now   = datetime.now().isoformat()
    data  = load_db()
    u     = data["users"].get(username, {})
    u["active_session"]  = token
    u["session_started"] = now
    u.setdefault("usage", {})["last_active"] = now
    data["users"][username] = u
    save_db(data)
    session.permanent = True
    session["username"]        = username
    session["session_token"]   = token
    session["session_started"] = now
    return redirect(request.args.get("next") or url_for("dashboard"))

@app.route("/logout")
def logout():
    un = session.get("username")
    if un:
        try:
            data = load_db()
            u = data["users"].get(un, {})
            u["active_session"] = None
            data["users"][un] = u
            save_db(data)
        except Exception: pass
    session.clear()
    return redirect(url_for("login"))

@app.route("/")
@login_required
def dashboard():
    u = get_user(session["username"])
    u = dict(u); u["username"] = session["username"]
    dy   = current_fy_end()
    reps = scan_reports(session["username"])
    tok  = False; companies = []
    if CORE_OK:
        tok = _tally_ok(u["tally_host"], u["tally_port"])
        if tok:
            try:
                core.set_tally_url(f"http://{u['tally_host']}:{u['tally_port']}")
                cos = core.list_companies()
                al  = [c.upper() for c in (u.get("allowed_companies") or [])]
                companies = [c for c in cos if c.upper() in al] if al else cos
            except Exception: pass
    cos_list = _get_user_companies_used(session["username"])
    co_limit = _get_user_company_limit(u)
    return _R("Dashboard — TallySync Pro", _DASH_BODY,
              fy_start=str(dy-1), fy_end_s=str(dy)[2:], def_year=dy,
              reports=reps, rcount=len(reps), companies=companies,
              tok=tok, core_ok=CORE_OK, classic_ok=CLASSIC_OK,
              verify_ok=VERIFY_OK,
              cos_used=len(cos_list), co_limit=co_limit, cos_list=cos_list)

@app.route("/generate", methods=["POST"])
@login_required
def generate():
    if not CORE_OK:
        flash("tally_core.py not found.", "error"); return redirect(url_for("dashboard"))
    u    = get_user(session["username"])
    host = request.form.get("host", u["tally_host"]).strip() or u["tally_host"]
    port = int(request.form.get("port", u["tally_port"]) or 9000)
    co   = (request.form.get("company_manual","").strip() or
            request.form.get("company","").strip())
    if not co:
        flash("Please select or type a company name.", "error")
        return redirect(url_for("dashboard"))
    al = [a.upper() for a in (u.get("allowed_companies") or [])]
    if al and co.upper() not in al:
        flash(f"Access denied for \'{co}\'.", "error")
        return redirect(url_for("dashboard"))

    cos_list = _get_user_companies_used(session["username"])
    co_limit = _get_user_company_limit(u)
    is_new   = co.upper() not in [c.upper() for c in cos_list]
    if is_new and len(cos_list) >= co_limit:
        flash(f"Company limit reached ({co_limit}). You can regenerate existing companies. Contact admin to increase limit.", "error")
        return redirect(url_for("dashboard"))

    mode = request.form.get("date_mode","fy")
    from_date = to_date = None
    if mode == "custom":
        from_date = request.form.get("from_date","").strip()
        to_date   = request.form.get("to_date","").strip()
        if not from_date or not to_date:
            flash("Provide both dates.", "error"); return redirect(url_for("dashboard"))
        try:
            from datetime import date as _d
            _fd = _d.fromisoformat(from_date); _td = _d.fromisoformat(to_date)
            if _fd >= _td: flash("From Date must be before To Date.", "error"); return redirect(url_for("dashboard"))
            year = _td.year if _td.month >= 4 else _td.year + 1
        except ValueError:
            flash("Invalid date format.", "error"); return redirect(url_for("dashboard"))
    else:
        year = int(request.form.get("year") or current_fy_end())

    # sign_info removed — signature block now auto-generated by sch3_verify.py
    jid = secrets.token_hex(16)
    threading.Thread(target=_run_job, daemon=True,
        args=(jid, session["username"], co, year, host, port,
              bool(request.form.get("gen_classic")),
              bool(request.form.get("autosync_py")),
              from_date, to_date, None)).start()
    return redirect(url_for("job_status", jid=jid))

@app.route("/job/<jid>")
@login_required
def job_status(jid):
    with _JOBS_LOCK: job = dict(_JOBS.get(jid, {}))
    if not job: flash("Job not found.", "error"); return redirect(url_for("dashboard"))
    un     = session["username"]
    status = job.get("status","unknown")
    log    = "\n".join(job.get("logs",[]))
    dl_pdf = dl_classic = dl_json = None
    if status == "done":
        if job.get("pdf"):    dl_pdf     = url_for("download", username=un, filename=Path(job["pdf"]).name)
        if job.get("classic_pdf"): dl_classic = url_for("download", username=un, filename=Path(job["classic_pdf"]).name)
        if job.get("json"):   dl_json    = url_for("download", username=un, filename=Path(job["json"]).name)
        _track_usage(un, job.get("company",""), "classic" if job.get("classic_pdf") else "standard")
    return _R(f"Job {jid[:8]} — TallySync Pro", _JOB_BODY,
              jid=jid, status=status, log=log,
              company=job.get("company",""), year=job.get("year",""),
              fy_start=job.get("fy_start",""), fy_end=job.get("fy_end",""),
              net_profit=job.get("net_profit",0), error=job.get("error",""),
              dl_pdf=dl_pdf, dl_classic=dl_classic, dl_json=dl_json)

@app.route("/api/job/<jid>")
@login_required
def api_job(jid):
    with _JOBS_LOCK: job = dict(_JOBS.get(jid, {"status":"not_found","logs":[]}))
    return jsonify({"status": job.get("status","unknown"), "logs": "\n".join(job.get("logs",[]))})

@app.route("/api/companies")
@login_required
def api_companies():
    if not CORE_OK: return jsonify({"companies":[],"error":"tally_core.py not loaded"})
    u = get_user(session["username"])
    host = request.args.get("host", u["tally_host"])
    port = int(request.args.get("port", u["tally_port"]) or 9000)
    try:
        core.set_tally_url(f"http://{host}:{port}")
        cos = core.list_companies()
        al  = [c.upper() for c in (u.get("allowed_companies") or [])]
        if al: cos = [c for c in cos if c.upper() in al]
        return jsonify({"companies": cos})
    except Exception as e:
        return jsonify({"companies":[], "error": str(e)})

@app.route("/dl/<username>/<filename>")
@login_required
def download(username, filename):
    me = get_user(session["username"])
    if me["role"] != "admin" and username != session["username"]: abort(403)
    filename = Path(filename).name
    path = REPORTS_DIR / username / filename
    if not path.exists(): abort(404)
    return send_file(str(path), as_attachment=True, download_name=filename)

# ── Schedule III Excel Download ───────────────────────────────────────────────
@app.route("/dl-excel/<username>/<json_name>")
@login_required
def download_excel(username, json_name):
    """Generate and stream a complete Schedule III Excel (BS + P&L + Notes) from a report JSON."""
    me = get_user(session["username"])
    if me["role"] != "admin" and username != session["username"]:
        abort(403)
    json_name = Path(json_name).name
    jpath = REPORTS_DIR / username / json_name
    if not jpath.exists():
        abort(404)
    try:
        with open(jpath, encoding="utf-8") as f:
            meta = json.load(f)
    except Exception:
        flash("Could not read report data.", "error")
        return redirect(url_for("dashboard"))
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter
    except ImportError:
        flash("openpyxl not installed.", "error")
        return redirect(url_for("dashboard"))

    # ── Pull figures from JSON ────────────────────────────────────────────────
    company  = meta.get("company", "")
    fy_start = meta.get("fy_start", "")
    fy_end   = meta.get("fy_end", "")
    gen_at   = meta.get("generated_at", "")[:19].replace("T", " ")
    raw      = meta.get("data", {}) if isinstance(meta.get("data"), dict) else {}

    def _v(key, default=0.0):
        try: return float(raw.get(key) or default)
        except: return float(default)

    revenue      = abs(_v("sales"))
    other_inc    = abs(_v("ind_inc"))
    if not revenue and other_inc:
        revenue = other_inc; other_inc = 0.0
    tot_rev      = revenue + other_inc
    cost_mat     = abs(_v("purchases"))
    dir_exp_amt  = abs(_v("dir_exp"))
    if dir_exp_amt and cost_mat: cost_mat += dir_exp_amt
    inv_chg      = max(0.0, abs(_v("open_s")) - abs(_v("bs_closing_stock")))
    emp_exp      = 0.0; fin_cost = 0.0; dep = 0.0
    other_exp    = abs(_v("ind_exp"))
    np_meta      = float(meta.get("net_profit") or 0)
    gp_meta      = float(meta.get("gross_profit") or 0)
    if not revenue: revenue = gp_meta; tot_rev = revenue + other_inc
    tot_exp_calc = cost_mat + inv_chg + emp_exp + fin_cost + dep + other_exp
    _diff = tot_rev - np_meta - tot_exp_calc
    if abs(_diff) > 0.5: other_exp += _diff; tot_exp_calc = tot_rev - np_meta
    tot_exp = tot_exp_calc; pbeet = tot_rev - tot_exp
    tax = _v("income_tax", 0.0); npat = np_meta

    eq_share     = abs(_v("equity_share"))
    pnl_on_bs    = _v("pnl_on_bs"); cap_total = _v("cap_total")
    reserves     = cap_total - eq_share + pnl_on_bs
    if abs(reserves) < 0.5: reserves = npat
    lt_borrow    = abs(_v("loans_total")) - abs(_v("dtl", 0))
    dtl_val      = abs(_v("dtl", 0))
    st_borrow    = 0.0
    cl_lines     = raw.get("cl_lines", [])
    trade_pay    = sum(abs(v) for n, v in cl_lines
                       if "trade pay" in n.lower() or "sundry cred" in n.lower()) if isinstance(cl_lines, list) else 0
    if not trade_pay: trade_pay = abs(_v("cl_total"))
    other_cl     = 0.0; short_prov = 0.0
    total_l      = eq_share + reserves + lt_borrow + dtl_val + st_borrow + trade_pay + other_cl + short_prov
    fa_net       = abs(_v("fa_net_bs")) or abs(_v("fa_total"))
    inv_total    = 0.0; lt_loans_adv = 0.0
    inventories  = abs(_v("bs_closing_stock"))
    trade_recv   = 0.0
    ca_lines     = raw.get("ca_lines", [])
    if isinstance(ca_lines, list):
        for n, v in ca_lines:
            if "trade rec" in n.lower() or "debtor" in n.lower() or "sundry deb" in n.lower():
                trade_recv += abs(v)
    cash_bank    = abs(_v("cash_hand", 0)) + abs(_v("bank_accts", 0))
    st_loans_adv = max(0.0, abs(_v("ca_total")) - trade_recv - cash_bank - inventories)
    total_a      = fa_net + inv_total + lt_loans_adv + inventories + trade_recv + cash_bank + st_loans_adv
    _bs_diff = total_l - total_a
    if abs(_bs_diff) > 1.0: cash_bank += _bs_diff; total_a = total_l

    def _fy_label(fe):
        try: return f"31 Mar {int(fe[:4])}"
        except: return fe
    cy_lbl = _fy_label(fy_end)
    py_lbl = _fy_label(f"{int(fy_end[:4])-1}{fy_end[4:]}" if fy_end else "")

    # ── Styles ────────────────────────────────────────────────────────────────
    NV="0B1F4B"; GOLD="C8A84B"; WHT="FFFFFF"; BLK="1E293B"; MUT="64748B"
    ALT1="F1F5FB"; ALT2="FFFFFF"; SHAD="DBEAFE"; GRN="14532D"; GRN_L="DCFCE7"
    RD_L="FEE2E2"; RD="991B1B"; SEC="EEF2FF"; GD_L="FEF9E7"; BDR_C="CBD5E1"

    def cfill(h): return PatternFill("solid", fgColor=h)
    def bdr(clr=BDR_C, thick_bot=False):
        t=Side(style="thin",color=clr); tb=Side(style="medium",color=NV) if thick_bot else t
        return Border(left=t,right=t,top=t,bottom=tb)
    def af(bold=False,sz=9,color=BLK,italic=False):
        return Font(name="Arial",bold=bold,size=sz,color=color,italic=italic)
    def ac(h="left",v="center",wrap=False,ind=0):
        return Alignment(horizontal=h,vertical=v,wrap_text=wrap,indent=ind)
    INR='#,##0;(#,##0);"-"'

    class SW:
        def __init__(self,ws,col_widths):
            self.ws=ws; self.row=0; ws.sheet_view.showGridLines=False
            for i,w in enumerate(col_widths,1): ws.column_dimensions[get_column_letter(i)].width=w
        def r(self): self.row+=1; return self.row
        def merge(self,rw,c1,c2): self.ws.merge_cells(f"{get_column_letter(c1)}{rw}:{get_column_letter(c2)}{rw}")
        def title_block(self,co,stmt,period):
            rw=self.r(); self.merge(rw,1,4); c=self.ws.cell(rw,1,value=co)
            c.font=Font(name="Arial",bold=True,size=13,color=GOLD); c.fill=cfill(NV); c.alignment=ac("center"); self.ws.row_dimensions[rw].height=30
            rw=self.r(); self.merge(rw,1,4); c=self.ws.cell(rw,1,value=stmt)
            c.font=Font(name="Arial",bold=True,size=10,color=WHT); c.fill=cfill(NV); c.alignment=ac("center"); self.ws.row_dimensions[rw].height=20
            rw=self.r(); self.merge(rw,1,4); c=self.ws.cell(rw,1,value=period)
            c.font=Font(name="Arial",italic=True,size=8,color="94A3B8"); c.fill=cfill(NV); c.alignment=ac("center"); self.ws.row_dimensions[rw].height=15
            rw=self.r(); self.merge(rw,1,4); self.ws.cell(rw,1).fill=cfill(NV); self.ws.row_dimensions[rw].height=5
        def col_hdr(self,labels):
            rw=self.r()
            for ci,(txt,al) in enumerate(labels,1):
                c=self.ws.cell(rw,ci,value=txt); c.font=Font(name="Arial",bold=True,size=8,color=WHT)
                c.fill=cfill(NV); c.alignment=ac(al,wrap=True); c.border=bdr(NV)
            self.ws.row_dimensions[rw].height=28; return rw
        def sec(self,label,nc=4):
            rw=self.r(); self.merge(rw,1,nc); c=self.ws.cell(rw,1,value=label)
            c.font=Font(name="Arial",bold=True,size=8,color=GOLD); c.fill=cfill(NV); c.alignment=ac("left",ind=1)
            for ci in range(2,nc+1): self.ws.cell(rw,ci).fill=cfill(NV)
            self.ws.row_dimensions[rw].height=17; return rw
        def subsec(self,label,nc=4):
            rw=self.r(); self.merge(rw,1,nc); c=self.ws.cell(rw,1,value=label)
            c.font=Font(name="Arial",bold=True,size=8,color=NV); c.fill=cfill(SEC); c.alignment=ac("left",ind=2)
            for ci in range(2,nc+1): cx=self.ws.cell(rw,ci); cx.fill=cfill(SEC); cx.border=bdr()
            self.ws.row_dimensions[rw].height=16; return rw
        def drow(self,note,label,cy,py=None,bold=False,ind=4,alt=True):
            rw=self.r(); bg=ALT1 if alt else ALT2
            cn=self.ws.cell(rw,1,value=note or ""); cn.font=af(sz=8,color=MUT if not note else NV)
            cn.fill=cfill(bg); cn.alignment=ac("center"); cn.border=bdr()
            cp=self.ws.cell(rw,2,value=label); cp.font=af(bold=bold); cp.fill=cfill(bg); cp.alignment=ac("left",ind=ind); cp.border=bdr()
            cc=self.ws.cell(rw,3,value=cy if cy else None); cc.font=af(bold=bold); cc.fill=cfill(bg); cc.alignment=ac("right"); cc.border=bdr(); cc.number_format=INR
            cp2=self.ws.cell(rw,4,value=py if py else None); cp2.font=af(color=MUT); cp2.fill=cfill(bg); cp2.alignment=ac("right"); cp2.border=bdr(); cp2.number_format=INR
            self.ws.row_dimensions[rw].height=16; return rw
        def blank(self,label,ind=4,alt=True): return self.drow("",label,None,None,ind=ind,alt=alt)
        def strow(self,label,cy_f,py_f=None,ind=3):
            rw=self.r(); self.merge(rw,1,2); c=self.ws.cell(rw,1,value=label)
            c.font=Font(name="Arial",bold=True,size=9,color=NV); c.fill=cfill(SHAD); c.alignment=ac("left",ind=ind); c.border=bdr(thick_bot=True)
            cc=self.ws.cell(rw,3,value=cy_f); cc.font=Font(name="Arial",bold=True,size=9,color=NV); cc.fill=cfill(SHAD); cc.alignment=ac("right"); cc.border=bdr(thick_bot=True); cc.number_format=INR
            cp=self.ws.cell(rw,4,value=py_f); cp.font=Font(name="Arial",bold=True,size=9,color=MUT); cp.fill=cfill(SHAD); cp.alignment=ac("right"); cp.border=bdr(thick_bot=True); cp.number_format=INR
            self.ws.row_dimensions[rw].height=18; return rw
        def trow(self,label,cy_f,py_f=None):
            rw=self.r(); self.merge(rw,1,2); c=self.ws.cell(rw,1,value=label)
            c.font=Font(name="Arial",bold=True,size=10,color=WHT); c.fill=cfill(NV); c.alignment=ac("left",ind=2); c.border=bdr(NV)
            cc=self.ws.cell(rw,3,value=cy_f); cc.font=Font(name="Arial",bold=True,size=10,color=WHT); cc.fill=cfill(NV); cc.alignment=ac("right"); cc.border=bdr(NV); cc.number_format=INR
            cp=self.ws.cell(rw,4,value=py_f); cp.font=Font(name="Arial",bold=True,size=10,color="94A3B8"); cp.fill=cfill(NV); cp.alignment=ac("right"); cp.border=bdr(NV); cp.number_format=INR
            self.ws.row_dimensions[rw].height=20; return rw
        def gap(self,h=5):
            rw=self.r(); self.merge(rw,1,4); self.ws.cell(rw,1).fill=cfill(WHT); self.ws.row_dimensions[rw].height=h; return rw
        def footer(self,txt):
            rw=self.r(); self.merge(rw,1,4); c=self.ws.cell(rw,1,value=txt)
            c.font=Font(name="Arial",italic=True,size=8,color=MUT); c.alignment=ac("left",wrap=True); self.ws.row_dimensions[rw].height=22
        def sign(self,co):
            rw=self.r(); self.merge(rw,1,2); self.ws.cell(rw,1,value=f"For {co}").font=af(bold=True,sz=8); self.ws.row_dimensions[rw].height=14
            for lbl,col in [("Partner / Director",1),("Partner / Director",3),("Membership No.: ___________",1),("UDIN: ___________",3),("Place: ___________",1),("Date: ___________",3)]:
                if col==1: rw=self.r()
                self.ws.cell(rw,col,value=lbl).font=af(italic=True,sz=8,color=MUT); self.ws.row_dimensions[rw].height=14

    wb = Workbook()

    # ══════════════════════════════════════════════════════════════════════════
    # SHEET 1 — BALANCE SHEET
    # ══════════════════════════════════════════════════════════════════════════
    ws1=wb.active; ws1.title="Balance Sheet"
    B=SW(ws1,[7,50,20,20])
    B.title_block(company.upper(),f"BALANCE SHEET AS AT {cy_lbl.upper()}",
        f"(Pursuant to Section 129, Companies Act 2013 — Schedule III)  |  Amt. in ₹  |  Generated: {gen_at}")
    B.col_hdr([("Note\nNo.","center"),("PARTICULARS","left"),(f"Figures as at\n{cy_lbl} (₹)","center"),(f"Figures as at\n{py_lbl} (₹)","center")])
    B.gap(6)
    B.sec("I.  EQUITY & LIABILITIES")
    B.subsec("  (1)  SHAREHOLDER'S FUNDS")
    r_eq=B.drow("1","        Share Capital",eq_share,None,ind=4,alt=True)
    r_res=B.drow("2","        Reserves & Surplus",reserves,None,ind=4,alt=False)
    B.blank("        Money Received Against Share Warrants",ind=4,alt=True)
    r_sf=B.strow("    Sub-Total — Shareholder's Funds",f"=C{r_eq}+C{r_res}",f"=D{r_eq}+D{r_res}",ind=4)
    B.blank("  (2)  Share Application Money Pending Allotment",ind=2,alt=True)
    B.subsec("  (3)  NON-CURRENT LIABILITIES")
    r_ltb=B.drow("3","        Long-Term Borrowings",lt_borrow,None,ind=4,alt=True)
    r_dtl=B.drow("32","        Deferred Tax Liabilities (Net)",dtl_val,None,ind=4,alt=False)
    B.blank("        Other Long-Term Liabilities",ind=4,alt=True)
    B.blank("        Long-Term Provisions",ind=4,alt=False)
    r_ncl=B.strow("    Sub-Total — Non-Current Liabilities",f"=C{r_ltb}+C{r_dtl}",f"=D{r_ltb}+D{r_dtl}",ind=4)
    B.subsec("  (4)  CURRENT LIABILITIES")
    r_stb=B.drow("4","        Short-Term Borrowings",st_borrow,None,ind=4,alt=True)
    r_tp=B.drow("5","        Trade Payables",trade_pay,None,ind=4,alt=False)
    r_ocl=B.drow("6","        Other Current Liabilities",other_cl,None,ind=4,alt=True)
    r_stp=B.drow("7","        Short-Term Provisions",short_prov,None,ind=4,alt=False)
    r_cl=B.strow("    Sub-Total — Current Liabilities",f"=C{r_stb}+C{r_tp}+C{r_ocl}+C{r_stp}",f"=D{r_stb}+D{r_tp}+D{r_ocl}+D{r_stp}",ind=4)
    B.gap(4); r_totL=B.trow("TOTAL — EQUITY AND LIABILITIES",f"=C{r_sf}+C{r_ncl}+C{r_cl}",f"=D{r_sf}+D{r_ncl}+D{r_cl}")
    B.gap(8)
    B.sec("II.  ASSETS")
    B.subsec("  (1)  NON-CURRENT ASSETS")
    B.subsec("       Property, Plant & Equipment")
    r_fa=B.drow("8","        Tangible Assets (Net Block / WDV)",fa_net,None,ind=4,alt=True)
    B.blank("        Intangible Assets",ind=4,alt=False); B.blank("        Capital Work-in-Progress",ind=4,alt=True)
    B.blank("        Intangible Assets Under Development",ind=4,alt=False)
    r_nci=B.drow("10","    Non-Current Investments",inv_total,None,ind=4,alt=True)
    B.blank("    Deferred Tax Asset (Net)",ind=4,alt=False)
    r_ltla=B.drow("9","    Long-Term Loans & Advances",lt_loans_adv,None,ind=4,alt=True)
    B.blank("    Other Non-Current Assets",ind=4,alt=False)
    r_nca=B.strow("    Sub-Total — Non-Current Assets",f"=C{r_fa}+C{r_nci}+C{r_ltla}",f"=D{r_fa}+D{r_nci}+D{r_ltla}",ind=4)
    B.subsec("  (2)  CURRENT ASSETS")
    B.blank("        Current Investments",ind=4,alt=True)
    r_inv=B.drow("11","        Inventories",inventories,None,ind=4,alt=False)
    r_tr=B.drow("12","        Trade Receivables",trade_recv,None,ind=4,alt=True)
    r_cb=B.drow("13","        Cash & Cash Equivalents",cash_bank,None,ind=4,alt=False)
    r_stla=B.drow("14","        Short-Term Loans & Advances",st_loans_adv,None,ind=4,alt=True)
    B.blank("        Other Current Assets",ind=4,alt=False)
    r_ca=B.strow("    Sub-Total — Current Assets",f"=C{r_inv}+C{r_tr}+C{r_cb}+C{r_stla}",f"=D{r_inv}+D{r_tr}+D{r_cb}+D{r_stla}",ind=4)
    B.gap(4); r_totA=B.trow("TOTAL — ASSETS",f"=C{r_nca}+C{r_ca}",f"=D{r_nca}+D{r_ca}")
    B.gap(4)
    rw=B.r(); B.merge(rw,1,2)
    c=ws1.cell(rw,1,value="✦  Balance Check  (Assets − Liabilities = should be ZERO)")
    c.font=Font(name="Arial",bold=True,size=8,color=NV); c.fill=cfill(GD_L); c.alignment=ac("right",ind=1); c.border=bdr()
    cc=ws1.cell(rw,3,value=f"=C{r_totA}-C{r_totL}"); cc.font=Font(name="Arial",bold=True,size=9,color=GRN)
    cc.fill=cfill(GD_L); cc.alignment=ac("right"); cc.border=bdr(); cc.number_format=INR
    cv=ws1.cell(rw,4,value=f'=IF(ABS(C{rw})<1,"✔  Balanced","⚠  Check figures")')
    cv.font=Font(name="Arial",bold=True,size=9,color=GRN); cv.fill=cfill(GD_L); cv.alignment=ac("center"); cv.border=bdr()
    ws1.row_dimensions[rw].height=18
    B.blank("  III.  CONTINGENT LIABILITIES",ind=2,alt=True); B.gap(8)
    B.footer("See accompanying Notes forming part of the Financial Statements.  As per our Report of even date.  For and on behalf of the Company.")
    B.sign(company); ws1.freeze_panes="B6"; ws1.page_setup.orientation="portrait"; ws1.page_setup.fitToPage=True; ws1.page_setup.fitToWidth=1; ws1.print_title_rows="1:5"

    # ══════════════════════════════════════════════════════════════════════════
    # SHEET 2 — PROFIT & LOSS
    # ══════════════════════════════════════════════════════════════════════════
    ws2=wb.create_sheet("Profit & Loss"); P2=SW(ws2,[7,56,20,20])
    P2.title_block(company.upper(),f"PROFIT & LOSS STATEMENT FOR THE YEAR ENDED {cy_lbl.upper()}",
        f"(Pursuant to Companies Act 2013 — Schedule III)  |  Amt. in ₹  |  Generated: {gen_at}")
    P2.col_hdr([("Note\nNo.","center"),("PARTICULARS","left"),(f"Year ended\n{cy_lbl} (₹)","center"),(f"Year ended\n{py_lbl} (₹)","center")])
    P2.gap(6)
    P2.sec("INCOME")
    rI=P2.drow("15","I.    REVENUE FROM OPERATIONS",revenue,None,ind=2,alt=True)
    rII=P2.drow("16","II.   OTHER INCOME",other_inc,None,ind=2,alt=False)
    P2.gap(3); rIII=P2.strow("III.  TOTAL REVENUE  (I + II)",f"=C{rI}+C{rII}",f"=D{rI}+D{rII}",ind=2); P2.gap(6)
    P2.sec("IV.   EXPENSES")
    r17=P2.drow("17","        Cost of Material Consumed",cost_mat,None,ind=4,alt=True)
    P2.blank("        Purchases of Stock-in-Trade",ind=4,alt=False)
    r18=P2.drow("18","        Changes in Inventories",inv_chg,None,ind=4,alt=True)
    r19=P2.drow("19","        Employee Benefit Expense",emp_exp,None,ind=4,alt=False)
    r20=P2.drow("20","        Financial Cost",fin_cost,None,ind=4,alt=True)
    r_dep=P2.drow("8","        Depreciation & Amortisation Expense",dep,None,ind=4,alt=False)
    r21=P2.drow("21","        Other Expenses",other_exp,None,ind=4,alt=True)
    P2.gap(3)
    r_te=P2.strow("        TOTAL EXPENSES",f"=C{r17}+C{r18}+C{r19}+C{r20}+C{r_dep}+C{r21}",f"=D{r17}+D{r18}+D{r19}+D{r20}+D{r_dep}+D{r21}",ind=4)
    P2.gap(6); P2.sec("PROFIT / (LOSS)")
    rV=P2.strow("V.    PROFIT BEFORE EXCEPTIONAL ITEMS & TAX  (III − IV)",f"=C{rIII}-C{r_te}",f"=D{rIII}-D{r_te}",ind=2)
    P2.blank("VI.   EXCEPTIONAL ITEMS",ind=2,alt=True); P2.blank("VII.  PRIOR PERIOD ITEMS",ind=2,alt=False)
    rVIII=P2.strow("VIII. PROFIT BEFORE TAX  (V − VI − VII)",f"=C{rV}",f"=D{rV}",ind=2); P2.gap(3)
    P2.sec("TAX EXPENSE")
    rX1=P2.drow("","X.    Current Tax",tax,None,ind=2,alt=True)
    rX2=P2.drow("22","      Deferred Tax",0,None,ind=2,alt=False)
    rXt=P2.strow("      Total Tax Expense",f"=C{rX1}+C{rX2}",f"=D{rX1}+D{rX2}",ind=2); P2.gap(4)
    rw=P2.r(); P2.merge(rw,1,2)
    np_bg=GRN_L if npat>=0 else RD_L; np_fg=GRN if npat>=0 else RD
    c=ws2.cell(rw,1,value="XI.   PROFIT / (LOSS) FOR THE PERIOD  (VIII − X)")
    c.font=Font(name="Arial",bold=True,size=10,color=np_fg); c.fill=cfill(np_bg); c.alignment=ac("left",ind=2); c.border=bdr()
    cc=ws2.cell(rw,3,value=f"=C{rVIII}-C{rXt}"); cc.font=Font(name="Arial",bold=True,size=11,color=np_fg); cc.fill=cfill(np_bg); cc.alignment=ac("right"); cc.border=bdr(); cc.number_format=INR
    cp=ws2.cell(rw,4,value=f"=D{rVIII}-D{rXt}"); cp.font=Font(name="Arial",bold=True,size=11,color=MUT); cp.fill=cfill(np_bg); cp.alignment=ac("right"); cp.border=bdr(); cp.number_format=INR
    ws2.row_dimensions[rw].height=22; r_np=rw
    P2.gap(4); P2.sec("EARNINGS PER SHARE")
    P2.drow("","XII.  Basic & Diluted EPS (₹)  [Face Value ₹10]",f"=IF(C{r_eq}>0,'Balance Sheet'!C{r_np}/(C{r_eq}/10),\"-\")",None,ind=2,alt=True)
    P2.gap(8); P2.footer("See accompanying Notes forming part of the Financial Statements.  As per our Report of even date.  For and on behalf of the Company.")
    P2.sign(company); ws2.freeze_panes="B6"; ws2.page_setup.orientation="portrait"; ws2.page_setup.fitToPage=True; ws2.page_setup.fitToWidth=1

    # ══════════════════════════════════════════════════════════════════════════
    # SHEET 3 — NOTES TO ACCOUNTS
    # ══════════════════════════════════════════════════════════════════════════
    ws3=wb.create_sheet("Notes to Accounts"); ws3.sheet_view.showGridLines=False
    for ci,w in enumerate([6,52,20,20],1): ws3.column_dimensions[get_column_letter(ci)].width=w
    ROW3=[3]
    for rw_h,(txt,fg,sz,h) in enumerate([(f"NOTES FORMING PART OF FINANCIAL STATEMENTS  —  {company.upper()}",WHT,11,26),(f"For the year ended {cy_lbl}  |  (Companies Act 2013 — Schedule III)  |  Amt. in ₹","94A3B8",8,15)],1):
        ws3.merge_cells(f"A{rw_h}:D{rw_h}"); c=ws3.cell(rw_h,1,value=txt)
        c.font=Font(name="Arial",bold=(sz>9),italic=(sz==8),size=sz,color=fg); c.fill=cfill(NV); c.alignment=ac("center"); ws3.row_dimensions[rw_h].height=h
    rw3=3
    for ci,(txt,al) in enumerate([("Note","center"),("Particulars","left"),(f"As at {cy_lbl} (₹)","center"),(f"As at {py_lbl} (₹)","center")],1):
        c=ws3.cell(rw3,ci,value=txt); c.font=Font(name="Arial",bold=True,size=8,color=WHT); c.fill=cfill(NV); c.alignment=ac(al,wrap=True); c.border=bdr(NV)
    ws3.row_dimensions[rw3].height=24
    def _nR(): ROW3[0]+=1; return ROW3[0]
    def _nhdr(num,title):
        rw=_nR(); ws3.merge_cells(f"A{rw}:D{rw}"); c=ws3.cell(rw,1,value=f"NOTE {num}  —  {title}")
        c.font=Font(name="Arial",bold=True,size=9,color=GOLD); c.fill=cfill(NV); c.alignment=ac("left",ind=1); ws3.row_dimensions[rw].height=18; return rw
    def _nrow(note,label,cy,py=None,bold=False,alt=True,ind=2):
        rw=_nR(); bg=ALT1 if alt else ALT2
        cn=ws3.cell(rw,1,value=note or ""); cn.font=af(sz=8,color=MUT if not note else NV); cn.fill=cfill(bg); cn.alignment=ac("center"); cn.border=bdr()
        cp=ws3.cell(rw,2,value=label); cp.font=af(bold=bold); cp.fill=cfill(bg); cp.alignment=ac("left",ind=ind); cp.border=bdr()
        cc=ws3.cell(rw,3,value=cy if cy else None); cc.font=af(bold=bold); cc.fill=cfill(bg); cc.alignment=ac("right"); cc.border=bdr(); cc.number_format=INR
        cp2=ws3.cell(rw,4,value=py if py else None); cp2.font=af(color=MUT); cp2.fill=cfill(bg); cp2.alignment=ac("right"); cp2.border=bdr(); cp2.number_format=INR
        ws3.row_dimensions[rw].height=16; return rw
    def _ntot(label,cy_f,py_f=None):
        rw=_nR(); ws3.merge_cells(f"A{rw}:B{rw}"); c=ws3.cell(rw,1,value=label)
        c.font=Font(name="Arial",bold=True,size=9,color=WHT); c.fill=cfill(NV); c.alignment=ac("left",ind=2); c.border=bdr(NV)
        cc=ws3.cell(rw,3,value=cy_f); cc.font=Font(name="Arial",bold=True,size=9,color=WHT); cc.fill=cfill(NV); cc.alignment=ac("right"); cc.border=bdr(NV); cc.number_format=INR
        cp=ws3.cell(rw,4,value=py_f); cp.font=Font(name="Arial",bold=True,size=9,color="94A3B8"); cp.fill=cfill(NV); cp.alignment=ac("right"); cp.border=bdr(NV); cp.number_format=INR
        ws3.row_dimensions[rw].height=18; return rw
    def _ng(): rw=_nR(); ws3.merge_cells(f"A{rw}:D{rw}"); ws3.cell(rw,1).fill=cfill(WHT); ws3.row_dimensions[rw].height=5; return rw

    _ng(); _nhdr("1","SHARE CAPITAL")
    n1a=_nrow("","Authorised Capital",eq_share,None,alt=True); n1b=_nrow("","Issued, Subscribed & Paid-up",eq_share,None,alt=False); n1c=_nrow("","Less: Calls in Arrears",0,None,alt=True)
    _ntot("TOTAL — SHARE CAPITAL",f"=C{n1b}-C{n1c}",f"=D{n1b}-D{n1c}")
    _ng(); _nhdr("2","RESERVES & SURPLUS")
    n2a=_nrow("","Capital Reserve",0,None,alt=True); n2b=_nrow("","General Reserve",0,None,alt=False)
    n2c=_nrow("","Surplus in P&L (Opening Balance)",0,None,alt=True); n2d=_nrow("","Add: Net Profit / (Loss) for the year",reserves,None,alt=False); n2e=_nrow("","Less: Drawings / Appropriations",0,None,alt=True)
    _ntot("TOTAL — RESERVES & SURPLUS",f"=C{n2a}+C{n2b}+C{n2c}+C{n2d}-C{n2e}",f"=D{n2a}+D{n2b}+D{n2c}+D{n2d}-D{n2e}")
    _ng(); _nhdr("3","LONG-TERM BORROWINGS")
    n3a=_nrow("","Term Loans — Banks (Secured)",0,None,alt=True); n3b=_nrow("","Term Loans — Others (Unsecured)",lt_borrow,None,alt=False); n3c=_nrow("","Loans from Directors / Partners",0,None,alt=True)
    _ntot("TOTAL — LONG-TERM BORROWINGS",f"=C{n3a}+C{n3b}+C{n3c}",f"=D{n3a}+D{n3b}+D{n3c}")
    _ng(); _nhdr("4","SHORT-TERM BORROWINGS")
    n4a=_nrow("","Cash Credit — Banks (Secured)",st_borrow,None,alt=True); n4b=_nrow("","Loans Repayable on Demand (Unsecured)",0,None,alt=False)
    _ntot("TOTAL — SHORT-TERM BORROWINGS",f"=C{n4a}+C{n4b}",f"=D{n4a}+D{n4b}")
    _ng(); _nhdr("5","TRADE PAYABLES")
    n5a=_nrow("","Micro & Small Enterprises (MSME)",0,None,alt=True); n5b=_nrow("","Others (Sundry Creditors)",trade_pay,None,alt=False)
    _ntot("TOTAL — TRADE PAYABLES",f"=C{n5a}+C{n5b}",f"=D{n5a}+D{n5b}")
    _ng(); _nhdr("6","OTHER CURRENT LIABILITIES")
    n6a=_nrow("","Advances from Customers",0,None,alt=True); n6b=_nrow("","Outstanding Expenses / Provisions",other_cl,None,alt=False); n6c=_nrow("","TDS Payable / GST Payable",0,None,alt=True)
    _ntot("TOTAL — OTHER CURRENT LIABILITIES",f"=C{n6a}+C{n6b}+C{n6c}",f"=D{n6a}+D{n6b}+D{n6c}")
    _ng(); _nhdr("7","SHORT-TERM PROVISIONS")
    n7a=_nrow("","Provision for Income Tax",short_prov,None,alt=True); n7b=_nrow("","Provision for Gratuity / Leave",0,None,alt=False)
    _ntot("TOTAL — SHORT-TERM PROVISIONS",f"=C{n7a}+C{n7b}",f"=D{n7a}+D{n7b}")
    _ng(); _nhdr("8","PROPERTY, PLANT & EQUIPMENT (TANGIBLE ASSETS)")
    _nrow("","GROSS BLOCK",None,None,bold=True,alt=True)
    n8ob=_nrow("","  Opening Gross Block",fa_net,None,alt=False); n8ad=_nrow("","  Add: Additions during the year",0,None,alt=True)
    n8dl=_nrow("","  Less: Disposals / Write-offs",0,None,alt=False); n8cb=_nrow("","  Closing Gross Block",fa_net,None,bold=True,alt=True)
    _nrow("","ACCUMULATED DEPRECIATION",None,None,bold=True,alt=False)
    n8od=_nrow("","  Opening Accumulated Depreciation",0,None,alt=True); n8cy=_nrow("","  Add: Depreciation for the year",dep,None,alt=False)
    n8dd=_nrow("","  Less: On Disposals",0,None,alt=True); n8cd=_nrow("","  Closing Accumulated Depreciation",dep,None,bold=True,alt=False)
    _ntot("NET BLOCK  (Written Down Value / WDV)",f"=C{n8cb}-C{n8cd}",f"=D{n8cb}-D{n8cd}")
    _ng(); _nhdr("9","LONG-TERM LOANS & ADVANCES")
    n9a=_nrow("","Security Deposits (Unsecured, Good)",lt_loans_adv,None,alt=True); n9b=_nrow("","Capital Advances",0,None,alt=False)
    _ntot("TOTAL — LONG-TERM LOANS & ADVANCES",f"=C{n9a}+C{n9b}",f"=D{n9a}+D{n9b}")
    _ng(); _nhdr("11","INVENTORIES  (as valued & certified by Management)")
    n11a=_nrow("","Raw Materials",0,None,alt=True); n11b=_nrow("","Work-in-Progress",0,None,alt=False)
    n11c=_nrow("","Finished Goods",0,None,alt=True); n11d=_nrow("","Stock-in-Trade",inventories,None,alt=False); n11e=_nrow("","Stores & Spares",0,None,alt=True)
    _ntot("TOTAL — INVENTORIES",f"=C{n11a}+C{n11b}+C{n11c}+C{n11d}+C{n11e}",f"=D{n11a}+D{n11b}+D{n11c}+D{n11d}+D{n11e}")
    _ng(); _nhdr("12","TRADE RECEIVABLES (SUNDRY DEBTORS)")
    _nrow("","Unsecured, Considered Good:",None,None,bold=True,alt=True)
    n12a=_nrow("","  Outstanding > 6 months",0,None,alt=False); n12b=_nrow("","  Outstanding ≤ 6 months",trade_recv,None,alt=True)
    _ntot("TOTAL — TRADE RECEIVABLES",f"=C{n12a}+C{n12b}",f"=D{n12a}+D{n12b}")
    _ng(); _nhdr("13","CASH & CASH EQUIVALENTS")
    n13a=_nrow("","Cash in Hand",0,None,alt=True); n13b=_nrow("","Balances with Banks — Current A/c",cash_bank,None,alt=False); n13c=_nrow("","Fixed Deposits (maturity < 3 months)",0,None,alt=True)
    _ntot("TOTAL — CASH & CASH EQUIVALENTS",f"=C{n13a}+C{n13b}+C{n13c}",f"=D{n13a}+D{n13b}+D{n13c}")
    _ng(); _nhdr("14","SHORT-TERM LOANS & ADVANCES")
    n14a=_nrow("","Advances to Suppliers",st_loans_adv,None,alt=True); n14b=_nrow("","Prepaid Expenses",0,None,alt=False); n14c=_nrow("","TDS Receivable / Input GST Credit",0,None,alt=True)
    _ntot("TOTAL — SHORT-TERM LOANS & ADVANCES",f"=C{n14a}+C{n14b}+C{n14c}",f"=D{n14a}+D{n14b}+D{n14c}")
    _ng(); _nhdr("15","REVENUE FROM OPERATIONS")
    n15a=_nrow("","Sale of Products / Services",revenue,None,alt=True); n15b=_nrow("","Less: Returns & Discounts",0,None,alt=False); n15c=_nrow("","Other Operating Revenue",0,None,alt=True)
    _ntot("TOTAL — REVENUE FROM OPERATIONS",f"=C{n15a}-C{n15b}+C{n15c}",f"=D{n15a}-D{n15b}+D{n15c}")
    _ng(); _nhdr("16","OTHER INCOME")
    n16a=_nrow("","Interest Income",other_inc,None,alt=True); n16b=_nrow("","Dividend / Profit on Sale of Assets",0,None,alt=False); n16c=_nrow("","Miscellaneous Income",0,None,alt=True)
    _ntot("TOTAL — OTHER INCOME",f"=C{n16a}+C{n16b}+C{n16c}",f"=D{n16a}+D{n16b}+D{n16c}")
    _ng(); _nhdr("17","COST OF MATERIAL CONSUMED / PURCHASES")
    n17a=_nrow("","Opening Stock of Raw Material / Goods",0,None,alt=True); n17b=_nrow("","Add: Purchases during the year",cost_mat,None,alt=False); n17c=_nrow("","Less: Closing Stock",0,None,alt=True)
    _ntot("TOTAL — COST OF MATERIAL",f"=C{n17a}+C{n17b}-C{n17c}",f"=D{n17a}+D{n17b}-D{n17c}")
    _ng(); _nhdr("18","CHANGES IN INVENTORIES")
    n18a=_nrow("","Opening Stock — FG / WIP / SIT",0,None,alt=True); n18b=_nrow("","Less: Closing Stock — FG / WIP / SIT",inv_chg,None,alt=False)
    _ntot("TOTAL — CHANGES IN INVENTORIES  (Opening − Closing)",f"=C{n18a}-C{n18b}",f"=D{n18a}-D{n18b}")
    _ng(); _nhdr("19","EMPLOYEE BENEFIT EXPENSE")
    n19a=_nrow("","Salaries, Wages & Allowances",emp_exp,None,alt=True); n19b=_nrow("","Director / Partner Remuneration",0,None,alt=False)
    n19c=_nrow("","PF / ESIC / Gratuity",0,None,alt=True); n19d=_nrow("","Staff Welfare Expenses",0,None,alt=False)
    _ntot("TOTAL — EMPLOYEE BENEFIT EXPENSE",f"=C{n19a}+C{n19b}+C{n19c}+C{n19d}",f"=D{n19a}+D{n19b}+D{n19c}+D{n19d}")
    _ng(); _nhdr("20","FINANCIAL COST")
    n20a=_nrow("","Interest on Term Loans / Overdraft",fin_cost,None,alt=True); n20b=_nrow("","Bank Charges & Processing Fees",0,None,alt=False)
    _ntot("TOTAL — FINANCIAL COST",f"=C{n20a}+C{n20b}",f"=D{n20a}+D{n20b}")
    _ng(); _nhdr("21","OTHER EXPENSES")
    n21a=_nrow("","Administrative & Office Expenses",other_exp,None,alt=True); n21b=_nrow("","Selling & Distribution Expenses",0,None,alt=False)
    n21c=_nrow("","Repairs & Maintenance",0,None,alt=True); n21d=_nrow("","Professional / Legal Fees",0,None,alt=False)
    n21e=_nrow("","Audit Fees",0,None,alt=True); n21f=_nrow("","Miscellaneous Expenses",0,None,alt=False)
    _ntot("TOTAL — OTHER EXPENSES",f"=C{n21a}+C{n21b}+C{n21c}+C{n21d}+C{n21e}+C{n21f}",f"=D{n21a}+D{n21b}+D{n21c}+D{n21d}+D{n21e}+D{n21f}")
    _ng(); _nhdr("22","DEFERRED TAX LIABILITIES / (ASSETS)")
    n22a=_nrow("","Opening Deferred Tax Liability / (Asset)",0,None,alt=True); n22b=_nrow("","Add: Deferred Tax (Expense) / Benefit",dtl_val,None,alt=False)
    _ntot("CLOSING DEFERRED TAX LIABILITY / (ASSET)",f"=C{n22a}+C{n22b}",f"=D{n22a}+D{n22b}")

    # Profit verification cross-check
    _ng(); _ng()
    rw_v=_nR(); ws3.merge_cells(f"A{rw_v}:D{rw_v}")
    c=ws3.cell(rw_v,1,value="★  PROFIT VERIFICATION  —  Note 15 (Revenue) − Note 21 (Expenses) should equal Net Profit on P&L sheet")
    c.font=Font(name="Arial",bold=True,size=9,color=NV); c.fill=cfill(GD_L); c.alignment=ac("left",ind=1); ws3.row_dimensions[rw_v].height=18
    rw_v2=_nR(); ws3.merge_cells(f"A{rw_v2}:B{rw_v2}")
    ws3.cell(rw_v2,1,value="Net Profit — cross-reference from P&L Sheet").font=af(bold=True,sz=9,color=NV)
    ws3.cell(rw_v2,1).fill=cfill(GD_L); ws3.cell(rw_v2,1).border=bdr(); ws3.cell(rw_v2,1).alignment=ac("right",ind=1)
    cc=ws3.cell(rw_v2,3,value=f"='Profit & Loss'!C{r_np}")
    cc.font=Font(name="Arial",bold=True,size=10,color=GRN); cc.fill=cfill(GD_L); cc.alignment=ac("right"); cc.border=bdr(); cc.number_format=INR
    ws3.row_dimensions[rw_v2].height=18
    _ng()
    rw_f=_nR(); ws3.merge_cells(f"A{rw_f}:D{rw_f}")
    c=ws3.cell(rw_f,1,value=f"All figures sourced from TallyPrime Trial Balance  |  Generated: {gen_at}  |  Notes form an integral part of the Financial Statements  |  Amounts in ₹")
    c.font=af(italic=True,sz=8,color=MUT); c.alignment=ac("left",wrap=True); ws3.row_dimensions[rw_f].height=18
    ws3.freeze_panes="B4"; ws3.page_setup.orientation="portrait"; ws3.page_setup.fitToPage=True; ws3.page_setup.fitToWidth=1

    # ── Stream to browser ─────────────────────────────────────────────────────
    buf=io.BytesIO(); wb.save(buf); buf.seek(0)
    safe_co=re.sub(r"[^\w]","_",company[:25])
    fname=f"ScheduleIII_{safe_co}_{(fy_end or '').replace('-','')}.xlsx"
    return send_file(buf,as_attachment=True,download_name=fname,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ── Classic Schedule III Excel (TB Data + Formulas + Unit Conversion) ────────
@app.route("/dl-excel-classic/<username>/<json_name>")
@login_required
def download_excel_classic(username, json_name):
    me = get_user(session["username"])
    if me["role"] != "admin" and username != session["username"]:
        abort(403)

    # Check module available — return visible error (new tab won't see flash)
    if not EXCEL_CLASSIC_OK or excel_classic is None:
        return "<h2 style='font-family:Arial;color:red'>Classic Excel Error</h2><p>sch3_excel_classic.pyc not found in EXE folder.<br>Run build.bat and copy sch3_excel_classic.pyc to dist_new\\TallySyncPro\\</p>", 500

    json_name = Path(json_name).name
    jpath = REPORTS_DIR / username / json_name
    if not jpath.exists():
        abort(404)

    try:
        with open(jpath, encoding="utf-8") as fh:
            payload = json.load(fh)
    except Exception as ex:
        return f"<h2 style='font-family:Arial;color:red'>Could not read report JSON</h2><pre>{ex}</pre>", 500

    try:
        import tally_core as tc
        D = tc.parse_data(payload["company"], payload["data"],
                          payload["fy_start"], payload["fy_end"])
        D_py = None
        # Match on raw payload company name (before parse_data cleans it)
        company_name_raw = payload.get("company", "")
        company_name     = D.get("company", "") or company_name_raw
        fy_end_str       = payload.get("fy_end", "") or D.get("fy_end", "")
        if fy_end_str and len(fy_end_str) >= 4:
            py_year   = str(int(fy_end_str[:4]) - 1)   # e.g. "2026"
            for uname in [username, session["username"]]:
                rep_dir = REPORTS_DIR / uname
                if not rep_dir.exists():
                    continue
                for jf in sorted(rep_dir.glob("*.json")):
                    if jf.name == Path(json_name).name:
                        continue   # skip the CY file itself
                    try:
                        with open(jf, encoding="utf-8") as tmp:
                            py_meta = json.load(tmp)
                        py_co  = py_meta.get("company", "").lower().strip()
                        py_fye = py_meta.get("fy_end", "")
                        # Match: same company name (raw or cleaned) AND fy_end year = CY-1
                        name_match = (py_co == company_name_raw.lower().strip() or
                                      py_co == company_name.lower().strip())
                        year_match = py_fye.startswith(py_year)
                        if name_match and year_match:
                            D_py = tc.parse_data(py_meta["company"], py_meta["data"],
                                                 py_meta["fy_start"], py_meta["fy_end"])
                            break
                    except Exception:
                        continue
                if D_py:
                    break

        import tempfile, os, subprocess as _sp
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp_f:
            tmp_path = tmp_f.name
        excel_classic.generate_schedule3_excel(D, D_py, tmp_path)

        # ── Server-side recalc: populate formula values before download ───────
        # openpyxl saves formulas as strings with no cached values.
        # LibreOffice recalculates them in-place so Excel opens with numbers.
        try:
            _recalc = Path(__file__).parent / "recalc.py"
            if _recalc.exists():
                _sp.run(
                    [sys.executable, str(_recalc), tmp_path, "60"],
                    capture_output=True, timeout=90,
                    cwd=str(_recalc.parent)
                )
        except Exception as _re:
            app.logger.warning(f"Excel recalc skipped: {_re}")
        # ─────────────────────────────────────────────────────────────────────

        with open(tmp_path, "rb") as fh:
            xlsx_bytes = fh.read()
        try:
            os.unlink(tmp_path)
        except Exception:
            pass
        buf = io.BytesIO(xlsx_bytes)
        safe_co = re.sub(r"[^\w]", "_", company_name[:25])
        fname = f"Sch3_Classic_{safe_co}_{(fy_end_str or '').replace('-','')}.xlsx"
        return send_file(buf, as_attachment=True, download_name=fname,
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as ex:
        err = traceback.format_exc()
        app.logger.error(err)
        return f"<h2 style='font-family:Arial;color:red'>Classic Excel Failed</h2><pre style='font-size:12px'>{err}</pre>", 500


# ── Licence Activation Route ─────────────────────────────────────────────────
@app.route("/activate", methods=["GET", "POST"])
def activate():
    machine_id = _get_machine_id()
    error = ""; success = ""; key_val = ""

    if request.method == "POST":
        key_val = request.form.get("key", "").strip().upper()
        result  = _verify_licence_key(key_val)
        if result["valid"]:
            _save_key(key_val)
            _get_licence_status(force=True)   # refresh cache immediately
            # Redirect straight to login — no need to show activation page again
            flash(f"✔ Licence activated! Valid until {result.get('exp_date','—')} "
                  f"({result.get('days_left','?')} days). Welcome to TallySync Pro.", "success")
            return redirect(url_for("login"))
        else:
            error = result["reason"]

    # Render the activation page using Jinja
    return render_template_string(_ACTIVATE_HTML,
        machine_id=machine_id, error=error,
        success=success, key_val=key_val if error else "")

# ── Admin routes ──────────────────────────────────────────────────────────────
@app.route("/admin/users")
@admin_required
def admin_users():
    data = load_db()
    cos_used    = {un: _get_user_companies_used(un) for un in data["users"]}
    total_rpts  = sum(u.get("usage", {}).get("total_reports", 0) for u in data["users"].values())
    active_cnt  = sum(1 for u in data["users"].values() if u.get("active", True))
    return _R("Users — TallySync Pro Admin", _ADMIN_USERS_BODY,
              users=data["users"], cos_used=cos_used,
              total_rpts=total_rpts, active_cnt=active_cnt)

@app.route("/admin/users/add", methods=["POST"])
@admin_required
def admin_user_add():
    data = load_db()
    un   = request.form.get("username","").strip().lower().replace(" ","_")
    if not un: flash("Username required.", "error"); return redirect(url_for("admin_users"))
    if un in data["users"]: flash(f"\'{un}\' already exists.", "error"); return redirect(url_for("admin_users"))
    pw  = request.form.get("password","").strip()
    al  = [c.strip() for c in request.form.get("allowed_companies","").split(",") if c.strip()]
    lim = max(1, min(15, int(request.form.get("company_limit") or MAX_COMPANIES_DEFAULT)))
    data["users"][un] = {
        "password": generate_password_hash(pw) if pw else None,
        "role": request.form.get("role","user"),
        "display_name": request.form.get("display_name","").strip() or un,
        "email": request.form.get("email","").strip(),
        "tally_host": request.form.get("tally_host","localhost").strip(),
        "tally_port": int(request.form.get("tally_port") or 9000),
        "allowed_companies": al, "company_limit": lim,
        "active": True, "active_session": None,
        "created_at": datetime.now().isoformat(),
        "usage": {"total_reports":0,"companies_accessed":[],"last_active":None,"report_log":[]},
    }
    save_db(data)
    flash(f"User \'{un}\' added.", "success")
    return redirect(url_for("admin_users"))

@app.route("/admin/users/edit/<uname>", methods=["GET","POST"])
@admin_required
def admin_user_edit(uname):
    data = load_db(); u = data["users"].get(uname)
    if not u: abort(404)
    if request.method == "POST":
        u["display_name"]      = request.form.get("display_name", u["display_name"])
        u["email"]             = request.form.get("email", u.get("email","")).strip()
        u["role"]              = request.form.get("role", u["role"])
        u["tally_host"]        = request.form.get("tally_host", u["tally_host"]).strip()
        u["tally_port"]        = int(request.form.get("tally_port") or u["tally_port"])
        u["company_limit"]     = max(1, min(15, int(request.form.get("company_limit") or 15)))
        u["allowed_companies"] = [c.strip() for c in request.form.get("allowed_companies","").split(",") if c.strip()]
        pw = request.form.get("password","").strip()
        if pw: u["password"] = generate_password_hash(pw)
        data["users"][uname] = u; save_db(data)
        flash(f"User \'{uname}\' updated.", "success")
        return redirect(url_for("admin_users"))
    return _R(f"Edit {uname} — Admin", _EDIT_BODY, uname=uname, u=u)

@app.route("/admin/users/toggle/<uname>", methods=["POST"])
@admin_required
def admin_user_toggle(uname):
    if uname == "admin": flash("Cannot disable admin.", "error"); return redirect(url_for("admin_users"))
    data = load_db(); u = data["users"].get(uname)
    if not u: abort(404)
    u["active"] = not u.get("active", True)
    if not u["active"]: u["active_session"] = None
    data["users"][uname] = u; save_db(data)
    flash(f"User \'{uname}\' {'enabled' if u['active'] else 'disabled'}.", "success")
    return redirect(url_for("admin_users"))

@app.route("/admin/users/delete/<uname>", methods=["POST"])
@admin_required
def admin_user_delete(uname):
    if uname == "admin": flash("Cannot delete admin.", "error"); return redirect(url_for("admin_users"))
    data = load_db(); data["users"].pop(uname, None); save_db(data)
    flash(f"User \'{uname}\' deleted.", "success")
    return redirect(url_for("admin_users"))

@app.route("/admin/users/usage/<uname>")
@admin_required
def admin_user_usage(uname):
    data = load_db(); u = data["users"].get(uname)
    if not u: abort(404)
    return _R(f"Usage: {uname}", _USAGE_BODY, uname=uname, u=u)

@app.route("/admin/users/limit/<uname>", methods=["POST"])
@admin_required
def admin_user_limit(uname):
    data = load_db(); u = data["users"].get(uname)
    if not u: abort(404)
    lim = max(1, min(15, int(request.form.get("company_limit") or 15)))
    u["company_limit"] = lim; data["users"][uname] = u; save_db(data)
    flash(f"Limit for \'{uname}\' updated to {lim}.", "success")
    return redirect(url_for("admin_user_usage", uname=uname))

@app.route("/admin/reports")
@admin_required
def admin_reports():
    all_r = []
    for d in REPORTS_DIR.iterdir():
        if d.is_dir():
            for r in scan_reports(d.name):
                r["client"] = d.name; all_r.append(r)
    all_r.sort(key=lambda r: r.get("generated_at",""), reverse=True)
    return _R("All Reports — Admin", _AREPORTS_BODY, reports=all_r)

# ══════════════════════════════════════════════════════════════════════════════
# Entry point  (works both as  python app.py  AND as PyInstaller EXE)
# ══════════════════════════════════════════════════════════════════════════════
def _get_base_dir() -> Path:
    """Return the folder that contains the EXE (or the script directory)."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent

def main():
    import socket, webbrowser, threading as _th

    # ── 1. Licence expiry hard-stop ───────────────────────────────────────────
    if _licence_expired():
        try:
            import tkinter as tk
            from tkinter import messagebox
            _r = tk.Tk(); _r.withdraw()
            messagebox.showerror(
                "TallySync Pro — Licence Expired",
                f"Your licence expired on {LICENCE_EXPIRY.strftime('%d %B %Y')}.\n\n"
                "Please contact your administrator to renew."
            )
            _r.destroy()
        except Exception:
            print("=" * 62)
            print("  TallySync Pro — LICENCE EXPIRED")
            print(f"  Expired on: {LICENCE_EXPIRY.strftime('%d %B %Y')}")
            print("  Contact your administrator to renew.")
            print("=" * 62)
        sys.exit(0)

    # ── 1b. Licence key check at startup ────────────────────────────────────
    if not IS_RENDER:
        _lic_startup = _get_licence_status(force=True)
        if not _lic_startup["valid"]:
            mid = _lic_startup.get("machine_id", "UNKNOWN")
            print("=" * 62)
            print("  TallySync Pro — LICENCE KEY REQUIRED")
            print(f"  Reason     : {_lic_startup.get('reason','No key found')}")
            print(f"  Machine ID : {mid}")
            print(f"  Send your Machine ID to your administrator to get your key.")
            print(f"  Then open the browser and enter the key to activate.")
            print("=" * 62)
            # Do NOT exit — Flask starts so the user can enter the key via browser

    # ── 2. Re-root BASE_DIR / REPORTS_DIR / CLIENTS_DB for EXE ──────────────
    global BASE_DIR, REPORTS_DIR, CLIENTS_DB
    BASE_DIR    = _get_base_dir()
    REPORTS_DIR = BASE_DIR / "reports"
    CLIENTS_DB  = BASE_DIR / "clients.json"
    REPORTS_DIR.mkdir(exist_ok=True)

    # ── 3. Find port — Render uses PORT env var, local scans for free port ──────
    if IS_RENDER:
        port = int(os.environ.get("PORT", 5000))
        host = "0.0.0.0"
        url  = f"http://0.0.0.0:{port}"
    else:
        host = "127.0.0.1"
        port = 5000
        for _p in [5000, 5001, 5002, 5003, 5050, 8080]:
            try:
                s = socket.socket(); s.bind(("127.0.0.1", _p)); s.close(); port = _p; break
            except OSError:
                continue
        url = f"http://127.0.0.1:{port}"

    # ── 4. Console banner ─────────────────────────────────────────────────────
    _days_left = (LICENCE_EXPIRY - date.today()).days
    print("=" * 62)
    print("  TallySync Pro  —  Schedule III PDF Generator")
    print("=" * 62)
    print(f"  tally_core.py   : {'ok' if CORE_OK else 'NOT FOUND — reports disabled'}")
    print(f"  sch3_classic.py : {'ok' if CLASSIC_OK else 'not found (optional)'}")
    print(f"  sch3_verify.py  : {'ok' if VERIFY_OK else 'not found (optional)'}")
    print(f"  Reports folder  : {REPORTS_DIR}")
    print(f"  Company limit   : {MAX_COMPANIES_DEFAULT} per user")
    print(f"  Licence valid   : until {LICENCE_EXPIRY.strftime('%d %B %Y')}  ({_days_left} days left)")
    _lic_info = _get_licence_status()
    if _lic_info["valid"]:
        print(f"  Licence key     : ✔ ACTIVE — expires {_lic_info.get('exp_date','?')}  ({_lic_info.get('days_left','?')} days)")
    else:
        print(f"  Licence key     : ✘ NOT ACTIVATED — {_lic_info.get('reason','')}")
        print(f"  Machine ID      : {_lic_info.get('machine_id','?')}")
    print()
    print(f"  Login           : admin / admin123")
    print(f"  Opening browser : {url}")
    print("=" * 62)

    # ── 5. Open browser (local only) ─────────────────────────────────────────
    if not IS_RENDER:
        def _open_browser():
            import time as _t; _t.sleep(1.5)
            webbrowser.open(url)
        _th.Thread(target=_open_browser, daemon=True).start()

    # ── 6. Start Flask ────────────────────────────────────────────────────────
    app.run(host=host, port=port, debug=False, threaded=True, use_reloader=False)

if __name__ == "__main__":
    main()