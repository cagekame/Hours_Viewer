"""
launcher_hv.py
===========
Mini launcher per Engineering Hours Viewer.

Flusso:
    1. Confronta data modifica exe rete vs locale
    2. Se rete più recente → copia in locale
    3. Aggiorna settings.json con i percorsi
    4. Lancia HoursViewer.exe locale
    5. Si chiude

Utilizzo:
    - Compilare con PyInstaller → launcher.exe
    - Distribuire launcher.exe agli utenti (non cambia mai)
    - Per aggiornare l'app: sovrascrivere HoursViewer.exe sulla rete
"""

import json
import os
import shutil
import subprocess
from pathlib import Path

# ── Percorsi ──────────────────────────────────────────────────────────────────
EXE_NAME    = "HoursViewer.exe"
LOCAL_DIR   = Path(os.environ["USERPROFILE"]) / "Hours_Viewer"
LOCAL_EXE   = LOCAL_DIR / EXE_NAME
SETTINGS    = LOCAL_DIR / "settings.json"
REMOTE_EXE  = Path(r"\\flowserve.net\des\lav\Engineering_utilities\HoursViewer.exe")
REMOTE_XLSX = Path(r"\\flowserve.net\des\lav\00_UTENTI\Caggianelli\Programmi\hours_data.xlsx")
LOCAL_XLSX  = LOCAL_DIR / "hours_data.xlsx"



# ── Helpers ───────────────────────────────────────────────────────────────────

def load_settings() -> dict:
    try:
        with open(SETTINGS, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_settings(s: dict):
    LOCAL_DIR.mkdir(parents=True, exist_ok=True)
    with open(SETTINGS, "w", encoding="utf-8") as f:
        json.dump(s, f, ensure_ascii=False, indent=2)


def get_mtime(path: Path) -> float:
    try:
        return path.stat().st_mtime
    except Exception:
        return 0.0


def copy_exe(remote_path: Path, local_path: Path, status_cb=None):
    local_path.parent.mkdir(parents=True, exist_ok=True)
    tmp = local_path.with_suffix(".tmp")
    try:
        if status_cb: status_cb("Downloading update...")
        shutil.copy2(str(remote_path), str(tmp))
        if local_path.exists():
            local_path.unlink()
        tmp.rename(local_path)
    except Exception as e:
        if tmp.exists():
            tmp.unlink()
        raise RuntimeError(f"Update failed: {e}")


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    import tkinter as tk

    root = tk.Tk()
    root.title("Engineering Hours Viewer")
    root.geometry("360x70")
    root.resizable(False, False)
    root.configure(bg="#9a3412")

    for ico_path in [
        Path(getattr(sys, "_MEIPASS", "")) / "HW.ico",
        LAUNCHER_DIR / "HW.ico",
    ]:
        if ico_path.exists():
            try: root.iconbitmap(str(ico_path))
            except Exception: pass
            break

    lbl = tk.Label(root, text="Starting Engineering Hours Viewer...",
                   font=("Segoe UI", 10), fg="#ffffff", bg="#9a3412")
    lbl.pack(expand=True)
    root.update()

    def status(msg):
        lbl.config(text=msg)
        root.update()

    try:
        status("Checking for updates...")

        # Leggi settings.json se esiste, altrimenti crea con i default
        s = load_settings()
        if not s.get("remote_exe"):
            # Prima esecuzione — crea settings con percorsi default
            s["remote_exe"] = str(REMOTE_EXE)
            s["local_exe"]  = str(LOCAL_EXE)
            s["xlsx_path"]  = str(LOCAL_XLSX)
            save_settings(s)

        remote_exe_path = Path(s["remote_exe"])
        local_exe_path  = Path(s["local_exe"])

        remote_mtime = get_mtime(remote_exe_path)
        local_mtime  = get_mtime(local_exe_path)

        if remote_mtime == 0.0:
            if not local_exe_path.exists():
                status("Network unreachable and no local copy found.")
                root.after(4000, root.destroy)
                root.mainloop()
                return
            status("Network unavailable — using local copy...")

        elif remote_mtime > local_mtime:
            status("New version found — updating...")
            copy_exe(remote_exe_path, local_exe_path, status_cb=status)
            status("Update complete!")
            root.update()

        else:
            status("Starting...")

        # Lancia exe locale
        subprocess.Popen(
            [str(LOCAL_EXE)],
            creationflags=subprocess.DETACHED_PROCESS |
                          subprocess.CREATE_NEW_PROCESS_GROUP
        )

        # Chiudi
        root.after(600, root.destroy)
        root.mainloop()

    except Exception as e:
        lbl.config(text=f"Error: {e}", fg="#fde8d0")
        root.after(5000, root.destroy)
        root.mainloop()


if __name__ == "__main__":
    main()