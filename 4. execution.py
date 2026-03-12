# run_pipeline.py
# Ejecuta en orden los scripts
# Uso: python run_pipeline.py

import sys
import subprocess
import datetime
import time
from pathlib import Path

# === Ajusta los nombres EXACTOS de tus archivos (incluyen espacios) ===
STEPS = [
    ("Extraccion desde Excel",        "0. key_processes_extration.py"),
    ("Detalle de horas desde Excel",  "1. hours_detailed.py"),
    ("Correo SALUD (envio/preview)",  "2. email_salud.py"),
    ("Actualizar Google Sheet",       "3. gsheets_update_siges_salud.py"),
]

# Opcional: tiempo maximo por paso (segundos). None = sin limite.
TIMEOUTS = {
    "Extraccion desde Excel":        None,
    "Correo SALUD (envio/preview)":  180,
    "Actualizar Google Sheet":       180,
}

def run_step(title: str, script_path: Path, timeout=None) -> int:
    """Ejecuta un paso, imprime stdout en tiempo real, devuelve returncode."""
    print(f"\n=== {title} ===")
    print(f"Script: {script_path}")

    proc = subprocess.Popen(
        [sys.executable, str(script_path)],
        cwd=str(script_path.parent),
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
    )

    start = time.time()
    try:
        for line in proc.stdout:
            print(line, end="")

        rc = proc.wait(timeout=timeout)

    except subprocess.TimeoutExpired:
        proc.kill()
        print(f"[ERROR] Timeout en '{title}' ({timeout}s).")
        return 124

    elapsed = time.time() - start
    print(f"[INFO] Fin '{title}' (rc={rc}, {elapsed:.1f}s)")
    return rc

def main():
    base = Path(__file__).resolve().parent
    print(f"[INFO] Pipeline iniciado: {datetime.datetime.now():%Y-%m-%d %H:%M:%S}")

    for title, fname in STEPS:
        script = base / fname
        if not script.exists():
            print(f"[ERROR] No se encontro el script: {script}")
            print("[ERROR] Aborto del pipeline.")
            sys.exit(1)

        rc = run_step(title, script, timeout=TIMEOUTS.get(title))

        if rc != 0:
            print(f"[ERROR] Pipeline abortado en: {title} (rc={rc})")
            sys.exit(rc)

    print("\n[OK] Pipeline completado sin errores.")

if __name__ == "__main__":
    main()
