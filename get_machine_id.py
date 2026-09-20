# -*- coding: utf-8 -*-
"""
get_machine_id.py — TallySync Pro Machine ID Tool
==================================================
Client runs this ONCE to get their Machine ID.
They share the Machine ID with their vendor to receive their licence key.
"""

import sys, hashlib, uuid, platform, subprocess, os
from datetime import datetime

_LICENCE_SALT = "TALLYSYNC-PRO-2026"   # must match app.py and keygen.py

def _get_machine_id() -> str:
    parts = []

    if platform.system() == "Windows":
        try:
            cpu = subprocess.check_output(
                "wmic cpu get ProcessorId", shell=True,
                stderr=subprocess.DEVNULL).decode(errors="ignore")
            cpu = " ".join(cpu.split()).replace("ProcessorId", "").strip()
            parts.append(cpu)
        except Exception:
            parts.append("NO_CPU")

        try:
            vol = subprocess.check_output(
                "wmic volume where DriveLetter='C:' get SerialNumber",
                shell=True, stderr=subprocess.DEVNULL).decode(errors="ignore")
            vol = " ".join(vol.split()).replace("SerialNumber", "").strip()
            parts.append(vol)
        except Exception:
            parts.append("NO_VOL")
    else:
        parts.append("DEV_MACHINE")
        parts.append("DEV_VOLUME")

    try:
        mac = uuid.UUID(int=uuid.getnode()).hex[-12:]
        parts.append(mac)
    except Exception:
        parts.append("NO_MAC")

    raw = "|".join(parts) + "|" + _LICENCE_SALT
    return hashlib.sha256(raw.encode()).hexdigest()[:16].upper()

def _get_pc_info() -> dict:
    info = {}
    try:
        import socket
        info["hostname"] = socket.gethostname()
    except: info["hostname"] = "Unknown"

    try:
        out = subprocess.check_output(
            "wmic computersystem get Name,Manufacturer,Model",
            shell=True, stderr=subprocess.DEVNULL).decode(errors="ignore")
        lines = [l.strip() for l in out.strip().splitlines() if l.strip()]
        if len(lines) >= 2:
            info["pc_model"] = lines[-1]
    except: info["pc_model"] = "Unknown"

    try:
        out = subprocess.check_output(
            "wmic os get Caption,Version",
            shell=True, stderr=subprocess.DEVNULL).decode(errors="ignore")
        lines = [l.strip() for l in out.strip().splitlines() if l.strip()]
        if len(lines) >= 2:
            info["os"] = lines[-1]
    except: info["os"] = platform.system() + " " + platform.release()

    return info

def main():
    os.system("cls" if platform.system() == "Windows" else "clear")

    print()
    print("╔══════════════════════════════════════════════════════════════╗")
    print("║              TallySync Pro  —  Machine ID Tool              ║")
    print("╚══════════════════════════════════════════════════════════════╝")
    print()
    print("  Detecting your machine fingerprint...")
    print()

    machine_id = _get_machine_id()
    pc         = _get_pc_info()

    print("  ┌─────────────────────────────────────────────────────────┐")
    print(f"  │   PC Name   : {pc.get('hostname','?'):<43}│")
    print(f"  │   PC Model  : {pc.get('pc_model','?'):<43}│")
    print(f"  │   OS        : {pc.get('os','?'):<43}│")
    print(f"  │   Detected  : {datetime.now().strftime('%d %b %Y  %H:%M'):<43}│")
    print("  ├─────────────────────────────────────────────────────────┤")
    print(f"  │                                                         │")
    print(f"  │   MACHINE ID :  {machine_id:<41}│")
    print(f"  │                                                         │")
    print("  └─────────────────────────────────────────────────────────┘")
    print()
    print("  ✅  Step 1 complete — Your Machine ID has been detected.")
    print()
    print("  ──────────────────────────────────────────────────────────")
    print("  📋  NEXT STEPS:")
    print()
    print(f"    1. Note down your Machine ID:  {machine_id}")
    print()
    print("    2. Send it to your TallySync Pro provider.")
    print()
    print("    3. You will receive a Licence Key (looks like this):")
    print("         XXXXXXXXXXXXX-XXXXXXXXXXXXX-XXXXXXXXXXXXX-XXXXXXXXXXXXX")
    print()
    print("    4. Open TallySync Pro → Enter the Licence Key → Activate")
    print("  ──────────────────────────────────────────────────────────")
    print()

    # Save to text file for easy sharing
    try:
        out_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                "machine_id.txt")
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(f"TallySync Pro — Machine ID\n")
            f.write(f"==========================\n")
            f.write(f"Machine ID : {machine_id}\n")
            f.write(f"PC Name    : {pc.get('hostname','?')}\n")
            f.write(f"PC Model   : {pc.get('pc_model','?')}\n")
            f.write(f"OS         : {pc.get('os','?')}\n")
            f.write(f"Detected   : {datetime.now().strftime('%d %b %Y  %H:%M')}\n")
            f.write(f"\nSend this file or the Machine ID above to your TallySync Pro provider.\n")
        print(f"  📄  Machine ID also saved to:  machine_id.txt  (in same folder)")
        print(f"      You can send this file directly to your provider.")
    except Exception:
        pass

    print()
    input("  Press Enter to close...")

if __name__ == "__main__":
    main()