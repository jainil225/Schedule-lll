# -*- coding: utf-8 -*-
"""
keygen.py  —  TallySync Pro Licence Key Generator
===================================================
INTERNAL TOOL — DO NOT DISTRIBUTE TO CLIENTS.

Usage:
  python keygen.py                            # interactive (auto-detects THIS machine)
  python keygen.py <MachineID> <YYYY-MM-DD>  # generate for a client
  python keygen.py verify <KEY>              # verify any key
"""

import sys, hashlib, base64, uuid, platform, subprocess, os
from datetime import date, datetime

# ── MUST MATCH THE VALUE IN app.py AND get_machine_id.py ─────────────────────
_LICENCE_SALT = "TALLYSYNC-PRO-2026"
# ─────────────────────────────────────────────────────────────────────────────


def _get_this_machine_id() -> str:
    """Detect the Machine ID of the PC currently running keygen.py."""
    parts = []
    if platform.system() == "Windows":
        try:
            cpu = subprocess.check_output(
                "wmic cpu get ProcessorId", shell=True,
                stderr=subprocess.DEVNULL).decode(errors="ignore")
            cpu = " ".join(cpu.split()).replace("ProcessorId", "").strip()
            parts.append(cpu)
        except: parts.append("NO_CPU")
        try:
            vol = subprocess.check_output(
                "wmic volume where DriveLetter='C:' get SerialNumber",
                shell=True, stderr=subprocess.DEVNULL).decode(errors="ignore")
            vol = " ".join(vol.split()).replace("SerialNumber", "").strip()
            parts.append(vol)
        except: parts.append("NO_VOL")
    else:
        parts.append("DEV_MACHINE"); parts.append("DEV_VOLUME")
    try:
        mac = uuid.UUID(int=uuid.getnode()).hex[-12:]
        parts.append(mac)
    except: parts.append("NO_MAC")
    raw = "|".join(parts) + "|" + _LICENCE_SALT
    return hashlib.sha256(raw.encode()).hexdigest()[:16].upper()


def _make_licence_key(machine_id: str, expiry: str) -> str:
    payload  = (machine_id.upper() + expiry).encode()
    checksum = hashlib.sha256(payload + _LICENCE_SALT.encode()).hexdigest()[:8]
    raw      = (machine_id.upper() + expiry + checksum).encode()
    b32      = base64.b32encode(raw).decode().rstrip("=")   # 52 chars
    return "-".join(b32[i:i+13] for i in range(0, 52, 13))


def _verify_key(key: str, machine_id: str = "") -> dict:
    key_clean = key.strip().upper().replace("-", "").replace(" ", "")
    try:
        pad = (8 - len(key_clean) % 8) % 8
        raw = base64.b32decode(key_clean + "=" * pad)
        if len(raw) < 32:
            return {"valid": False, "reason": "Invalid key format"}
        mid_in  = raw[:16].decode(errors="replace")
        exp_in  = raw[16:24].decode(errors="replace")
        chk_in  = raw[24:32].decode(errors="replace")
        payload = (mid_in + exp_in).encode()
        chk_c   = hashlib.sha256(payload + _LICENCE_SALT.encode()).hexdigest()[:8]
        if chk_in != chk_c:
            return {"valid": False, "reason": "Key is corrupted or tampered"}
        if machine_id and mid_in.upper() != machine_id.upper():
            return {"valid": False, "reason": f"Wrong machine (key: {mid_in})"}
        try:
            exp_date  = date(int(exp_in[:4]), int(exp_in[4:6]), int(exp_in[6:8]))
            days_left = (exp_date - date.today()).days
        except:
            return {"valid": False, "reason": "Invalid expiry in key"}
        return {"valid": True, "machine_id": mid_in, "expiry": exp_in,
                "exp_date": exp_date.strftime("%d %b %Y"),
                "days_left": days_left, "expired": days_left < 0}
    except Exception as e:
        return {"valid": False, "reason": str(e)}


def _banner():
    os.system("cls" if platform.system() == "Windows" else "clear")
    print()
    print("╔══════════════════════════════════════════════════════════╗")
    print("║   TallySync Pro  —  Licence Key Generator               ║")
    print("║                   INTERNAL USE ONLY                     ║")
    print("╚══════════════════════════════════════════════════════════╝")
    print()


def _pick_expiry() -> date:
    today = date.today()
    print("  Licence duration:")
    print("    1  —  1 Year   (recommended — annual renewal)")
    print("    2  —  6 Months")
    print("    3  —  3 Months (trial)")
    print("    4  —  Custom date")
    print()
    choice = input("  Choose [1]: ").strip() or "1"
    if choice == "1":
        return date(today.year + 1, today.month, today.day)
    elif choice == "2":
        m = today.month + 6; y = today.year + (m-1)//12; m = (m-1)%12+1
        return date(y, m, today.day)
    elif choice == "3":
        m = today.month + 3; y = today.year + (m-1)//12; m = (m-1)%12+1
        return date(y, m, today.day)
    else:
        raw = input("  Expiry date (YYYY-MM-DD): ").strip()
        try:
            return datetime.strptime(raw, "%Y-%m-%d").date()
        except ValueError:
            print("  ⚠  Invalid date."); sys.exit(1)


def _print_key_result(machine_id, exp, key):
    today = date.today()
    print()
    print("  ╔════════════════════════════════════════════════════════╗")
    print("  ║              LICENCE KEY GENERATED                     ║")
    print("  ╠════════════════════════════════════════════════════════╣")
    print(f"  ║  Machine ID : {machine_id:<41}║")
    print(f"  ║  Expiry     : {exp.strftime('%d %B %Y'):<41}║")
    print(f"  ║  Days left  : {(exp - today).days:<41}║")
    print("  ╠════════════════════════════════════════════════════════╣")
    print(f"  ║                                                        ║")
    print(f"  ║  KEY :  {key:<47}║")
    print(f"  ║                                                        ║")
    print("  ╚════════════════════════════════════════════════════════╝")
    print()
    v = _verify_key(key, machine_id)
    if v["valid"]:
        print("  ✔  Self-verification passed — key is genuine.")
    else:
        print(f"  ✘  Self-verification FAILED: {v['reason']}")
    print()


def _interactive():
    _banner()

    # Auto-detect THIS machine's ID
    print("  Detecting this machine's ID...", end=" ", flush=True)
    this_mid = _get_this_machine_id()
    print(f"done.\n")

    print("  ┌──────────────────────────────────────────────────────┐")
    print(f"  │  THIS machine ID : {this_mid:<37}│")
    print("  └──────────────────────────────────────────────────────┘")
    print()
    print("  Options:")
    print("    1  —  Generate key for THIS machine (shown above)")
    print("    2  —  Generate key for a CLIENT machine (paste their ID)")
    print()
    opt = input("  Choose [1]: ").strip() or "1"
    print()

    if opt == "1":
        machine_id = this_mid
        print(f"  Using this machine: {machine_id}")
    else:
        print("  Paste the Machine ID from the client's machine_id.txt or activation screen:")
        machine_id = input("  Client Machine ID : ").strip().upper()
        if len(machine_id) != 16:
            print(f"  ⚠  Machine ID must be 16 characters (got {len(machine_id)}).")
            sys.exit(1)

    print()
    exp        = _pick_expiry()
    expiry_str = exp.strftime("%Y%m%d")
    key        = _make_licence_key(machine_id, expiry_str)
    _print_key_result(machine_id, exp, key)

    print("  📋  Copy the KEY above and send it to the client.")
    print("      They enter it on the TallySync Pro activation screen.")
    print()

    # Offer to save to file
    try:
        save_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                 f"licence_{machine_id}.txt")
        with open(save_path, "w", encoding="utf-8") as f:
            f.write(f"TallySync Pro — Licence Key\n")
            f.write(f"===========================\n")
            f.write(f"Machine ID : {machine_id}\n")
            f.write(f"Expiry     : {exp.strftime('%d %B %Y')}\n")
            f.write(f"Generated  : {datetime.now().strftime('%d %b %Y  %H:%M')}\n")
            f.write(f"\nLICENCE KEY:\n{key}\n")
        print(f"  📄  Also saved to: licence_{machine_id}.txt")
    except Exception:
        pass

    print()
    input("  Press Enter to close...")


def _generate_direct(machine_id: str, expiry_raw: str):
    _banner()
    try:
        exp = datetime.strptime(expiry_raw, "%Y-%m-%d").date()
    except ValueError:
        print(f"  ⚠  Invalid date: {expiry_raw!r}  (use YYYY-MM-DD)")
        sys.exit(1)
    machine_id = machine_id.upper()
    if len(machine_id) != 16:
        print(f"  ⚠  Machine ID must be 16 chars (got {len(machine_id)})")
        sys.exit(1)
    key = _make_licence_key(machine_id, exp.strftime("%Y%m%d"))
    _print_key_result(machine_id, exp, key)


def _verify_mode(key: str):
    _banner()
    v = _verify_key(key)
    print(f"  Key : {key}")
    print()
    if v["valid"]:
        status = "⚠  EXPIRED" if v["expired"] else "✔  VALID"
        print(f"  Status     : {status}")
        print(f"  Machine ID : {v['machine_id']}")
        print(f"  Expiry     : {v['exp_date']}  ({v['days_left']} days)")
    else:
        print(f"  Status     : ✘  INVALID")
        print(f"  Reason     : {v['reason']}")
    print()


if __name__ == "__main__":
    args = sys.argv[1:]
    if not args:
        _interactive()
    elif len(args) == 2 and args[0].lower() == "verify":
        _verify_mode(args[1])
    elif len(args) == 2:
        _generate_direct(args[0], args[1])
    else:
        print(__doc__); sys.exit(1)